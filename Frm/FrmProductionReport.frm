VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmProductionreport 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ‹‹ﬁ‹‹«—Ì‹‹‹— «·«‰‰«Ã"
   ClientHeight    =   7170
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10080
   HelpContextID   =   470
   Icon            =   "FrmProductionReport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   7170
   ScaleWidth      =   10080
   Begin C1SizerLibCtl.C1Elastic Ele 
      Height          =   7110
      Index           =   0
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   10020
      _cx             =   17674
      _cy             =   12541
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
      Begin ImpulseButton.ISButton Cmd 
         Height          =   480
         Left            =   0
         TabIndex        =   2
         Top             =   6480
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   847
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
      End
      Begin C1SizerLibCtl.C1Elastic EleMain 
         Height          =   6375
         Left            =   30
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   30
         Width           =   9780
         _cx             =   17251
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
         Appearance      =   5
         MousePointer    =   0
         Version         =   801
         BackColor       =   14871017
         ForeColor       =   -2147483630
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   ""
         Align           =   0
         AutoSizeChildren=   2
         BorderWidth     =   6
         ChildSpacing    =   4
         Splitter        =   -1  'True
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
         Begin C1SizerLibCtl.C1Tab MainTab 
            CausesValidation=   0   'False
            Height          =   6195
            Left            =   90
            TabIndex        =   3
            Top             =   90
            Width           =   9600
            _cx             =   16933
            _cy             =   10927
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
            ForeColor       =   -2147483630
            FrontTabColor   =   14871017
            BackTabColor    =   12648447
            TabOutlineColor =   -2147483632
            FrontTabForeColor=   16711680
            Caption         =   " ﬁ«—Ì— «·«‰ «Ã"
            Align           =   0
            CurrTab         =   0
            FirstTab        =   0
            Style           =   3
            Position        =   1
            AutoSwitch      =   -1  'True
            AutoScroll      =   -1  'True
            TabPreview      =   -1  'True
            ShowFocusRect   =   0   'False
            TabsPerPage     =   0
            BorderWidth     =   0
            BoldCurrent     =   -1  'True
            DogEars         =   -1  'True
            MultiRow        =   0   'False
            MultiRowOffset  =   200
            CaptionStyle    =   0
            TabHeight       =   0
            TabCaptionPos   =   4
            TabPicturePos   =   0
            CaptionEmpty    =   ""
            Separators      =   -1  'True
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   37
            Begin C1SizerLibCtl.C1Elastic ElcContainer 
               Height          =   5820
               Index           =   0
               Left            =   45
               TabIndex        =   4
               TabStop         =   0   'False
               Top             =   45
               Width           =   9510
               _cx             =   16775
               _cy             =   10266
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
               Frame           =   0
               FrameStyle      =   5
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
                  Height          =   5205
                  Index           =   2
                  Left            =   330
                  TabIndex        =   5
                  TabStop         =   0   'False
                  Top             =   510
                  Width           =   9000
                  _cx             =   15875
                  _cy             =   9181
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
                  CaptionPos      =   6
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
                  Frame           =   0
                  FrameStyle      =   5
                  FrameWidth      =   1
                  FrameColor      =   -2147483628
                  FrameShadow     =   -2147483632
                  FloodStyle      =   1
                  _GridInfo       =   ""
                  AccessibleName  =   ""
                  AccessibleDescription=   ""
                  AccessibleValue =   ""
                  AccessibleRole  =   9
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«· ÕÊÌ·« "
                     Height          =   195
                     Index           =   20
                     Left            =   420
                     RightToLeft     =   -1  'True
                     TabIndex        =   45
                     Top             =   2580
                     Width           =   3540
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«Ê«„— «· Ã„Ì⁄"
                     Height          =   195
                     Index           =   19
                     Left            =   5310
                     RightToLeft     =   -1  'True
                     TabIndex        =   44
                     Top             =   2580
                     Width           =   3540
                  End
                  Begin VB.CheckBox chkCust 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "·⁄„Ì· „⁄Ì‰"
                     Height          =   375
                     Left            =   6120
                     RightToLeft     =   -1  'True
                     TabIndex        =   42
                     Top             =   3990
                     Width           =   1335
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "ﬁÌ«” «·«‰Ã«“ «·›⁄·Ì ·ﬂ· ⁄«„·"
                     Height          =   195
                     Index           =   18
                     Left            =   420
                     RightToLeft     =   -1  'True
                     TabIndex        =   41
                     Top             =   2310
                     Width           =   3540
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "ﬁÌ«” «·«‰Ã«“ «·›⁄·Ì ·ﬂ· ¬·…"
                     Height          =   195
                     Index           =   17
                     Left            =   420
                     RightToLeft     =   -1  'True
                     TabIndex        =   40
                     Top             =   2010
                     Width           =   3540
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "„Êﬁ›  ”·Ì„ «·«’‰«› «·„»«⁄Â „⁄ ”‰œ«  «·’—› "
                     Height          =   195
                     Index           =   16
                     Left            =   240
                     RightToLeft     =   -1  'True
                     TabIndex        =   39
                     Top             =   1740
                     Width           =   3720
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   " ﬁ—Ì— „ﬁ«—‰… „ﬂÊ‰«  «·«’‰«› „⁄ «·’—› «·›⁄·Ì"
                     Height          =   195
                     Index           =   15
                     Left            =   420
                     RightToLeft     =   -1  'True
                     TabIndex        =   38
                     Top             =   1440
                     Width           =   3540
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   " ﬁ—Ì— „ﬁ«—‰… «·«‰ «Ã «·›⁄·Ì"
                     Height          =   195
                     Index           =   14
                     Left            =   1260
                     RightToLeft     =   -1  'True
                     TabIndex        =   37
                     Top             =   1170
                     Width           =   2700
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "„—«Ã⁄Â «·«‰ «Ã Ê«·ÕÃ“"
                     CausesValidation=   0   'False
                     Height          =   195
                     Index           =   13
                     Left            =   4140
                     RightToLeft     =   -1  'True
                     TabIndex        =   36
                     Top             =   840
                     Width           =   1860
                  End
                  Begin VB.CommandButton Command1 
                     Caption         =   " ﬁ«—Ì— «· ﬂ«·Ì›  "
                     Height          =   255
                     Left            =   3120
                     RightToLeft     =   -1  'True
                     TabIndex        =   35
                     Top             =   2910
                     Width           =   2655
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "„—«Ã⁄Â «Ê«„— «·«‰ «Ã"
                     CausesValidation=   0   'False
                     Height          =   195
                     Index           =   12
                     Left            =   7260
                     RightToLeft     =   -1  'True
                     TabIndex        =   34
                     Top             =   840
                     Width           =   1620
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   " ﬁ—Ì— ÌÊ÷Õ «Ê«„— «·‘€· «·„› ÊÕ…-·Ì” ·Â« ”‰œ «” ·«„"
                     Height          =   195
                     Index           =   11
                     Left            =   3420
                     RightToLeft     =   -1  'True
                     TabIndex        =   33
                     Top             =   1800
                     Width           =   5460
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   " ﬁ—Ì— ÌÊ÷Õ «·„ Êﬁ⁄ Ê «·›⁄·Ï ··«’‰«›  ·ﬂ· «„— ‘€·"
                     Height          =   195
                     Index           =   10
                     Left            =   4980
                     RightToLeft     =   -1  'True
                     TabIndex        =   32
                     Top             =   1440
                     Width           =   3900
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "„—«Ã⁄Â ”‰œ«  ’—›  ·«Ê«„— «·«‰ «Ã"
                     Height          =   195
                     Index           =   9
                     Left            =   1260
                     RightToLeft     =   -1  'True
                     TabIndex        =   31
                     Top             =   840
                     Width           =   2700
                  End
                  Begin VB.CheckBox chkGroup 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "·„Ã„Ê⁄Â „⁄Ì‰"
                     Height          =   375
                     Left            =   6120
                     RightToLeft     =   -1  'True
                     TabIndex        =   29
                     Top             =   3630
                     Width           =   1335
                  End
                  Begin VB.CheckBox ChkMain 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "⁄—÷ «· ﬁ—Ì— ÿ»ﬁ« ··ÊÕœ… «·ﬂ»—Ï"
                     ForeColor       =   &H000000FF&
                     Height          =   195
                     Left            =   5460
                     RightToLeft     =   -1  'True
                     TabIndex        =   28
                     ToolTipText     =   " ” Œœ„ ··„ƒ””«  «· Ì  »Ì⁄ »ÊÕœ… Ê«Õœ… ›ﬁÿ ··’„› «·Ê«Õœ"
                     Top             =   90
                     Value           =   1  'Checked
                     Width           =   2535
                  End
                  Begin VB.CheckBox Chekopt 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "·„Œ“‰ „⁄Ì‰"
                     Height          =   375
                     Left            =   6120
                     RightToLeft     =   -1  'True
                     TabIndex        =   27
                     Top             =   3270
                     Width           =   1335
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   " ﬁ—Ì—     «·Â«·ﬂ   Œ·«· › —… ÿ»ﬁ« ·”‰œ«  «” ·«„ «·«‰ «Ã «· «„"
                     CausesValidation=   0   'False
                     Height          =   195
                     Index           =   8
                     Left            =   3060
                     RightToLeft     =   -1  'True
                     TabIndex        =   25
                     Top             =   2040
                     Width           =   5820
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   " ﬁ—Ì— «·«‰ «Ã «·‰„ÿÌ „Ã„⁄ Œ·«· ﬁ —… „⁄Ì‰…"
                     CausesValidation=   0   'False
                     Height          =   195
                     Index           =   7
                     Left            =   5580
                     RightToLeft     =   -1  'True
                     TabIndex        =   24
                     Top             =   2280
                     Width           =   3300
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   " ﬁ—Ì— «·«’‰«› «·„ÊÃÊœ… ›Ì ÿ·»Ì… „⁄Ì‰…"
                     Height          =   195
                     Index           =   6
                     Left            =   8520
                     RightToLeft     =   -1  'True
                     TabIndex        =   23
                     Top             =   2640
                     Visible         =   0   'False
                     Width           =   3180
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   " ﬁ—Ì— «·«’‰«› «·„ÊÃÊœ… ›Ì ÿ·»Ì… „⁄Ì‰…"
                     Height          =   195
                     Index           =   5
                     Left            =   8400
                     RightToLeft     =   -1  'True
                     TabIndex        =   22
                     Top             =   3720
                     Visible         =   0   'False
                     Width           =   3180
                  End
                  Begin VB.TextBox Text1 
                     Alignment       =   1  'Right Justify
                     Height          =   330
                     Left            =   3030
                     RightToLeft     =   -1  'True
                     TabIndex        =   20
                     Top             =   4680
                     Visible         =   0   'False
                     Width           =   1665
                  End
                  Begin VB.TextBox Txt_order_no 
                     Alignment       =   1  'Right Justify
                     Height          =   330
                     Left            =   3030
                     RightToLeft     =   -1  'True
                     TabIndex        =   18
                     Top             =   4320
                     Visible         =   0   'False
                     Width           =   1665
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·ÿ·»Ì«  «· Ì ·„  ”·„ Õ Ï «·«‰"
                     Height          =   195
                     Index           =   4
                     Left            =   8160
                     RightToLeft     =   -1  'True
                     TabIndex        =   10
                     Top             =   3000
                     Visible         =   0   'False
                     Width           =   2820
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "„Êﬁ›  ”·Ì„ «·ÿ·»Ì« "
                     Height          =   195
                     Index           =   3
                     Left            =   7920
                     RightToLeft     =   -1  'True
                     TabIndex        =   9
                     Top             =   3345
                     Visible         =   0   'False
                     Width           =   2820
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·«’‰«› «·„‰ Ã…  „‰ ÿ·»Ì… „⁄Ì‰"
                     Height          =   195
                     Index           =   2
                     Left            =   8280
                     RightToLeft     =   -1  'True
                     TabIndex        =   8
                     Top             =   3480
                     Visible         =   0   'False
                     Width           =   2820
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«· ﬂ«·Ì› «·«‰ «ÃÌ… ·√„— «‰ «Ã „⁄Ì‰"
                     Height          =   195
                     Index           =   1
                     Left            =   7680
                     RightToLeft     =   -1  'True
                     TabIndex        =   7
                     Top             =   4680
                     Visible         =   0   'False
                     Width           =   2820
                  End
                  Begin VB.OptionButton OptAccount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   " ﬁ—Ì—    «·«’‰«› «·„ Ã… Œ·«· › —… ÿ»ﬁ« ·”‰œ«  «” ·«„ «·«‰ «Ã «· «„"
                     CausesValidation=   0   'False
                     Height          =   195
                     Index           =   0
                     Left            =   4140
                     RightToLeft     =   -1  'True
                     TabIndex        =   6
                     Top             =   1140
                     Width           =   4740
                  End
                  Begin C1SizerLibCtl.C1Elastic Ele 
                     Height          =   1065
                     Index           =   1
                     Left            =   360
                     TabIndex        =   11
                     TabStop         =   0   'False
                     Top             =   3360
                     Width           =   2355
                     _cx             =   4154
                     _cy             =   1879
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
                     Caption         =   " ÕœÌœ «·› —… «·“„‰Ì…"
                     Align           =   0
                     AutoSizeChildren=   0
                     BorderWidth     =   6
                     ChildSpacing    =   4
                     Splitter        =   0   'False
                     FloodDirection  =   0
                     FloodPercent    =   0
                     CaptionPos      =   7
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
                     Frame           =   0
                     FrameStyle      =   5
                     FrameWidth      =   1
                     FrameColor      =   -2147483628
                     FrameShadow     =   -2147483632
                     FloodStyle      =   1
                     _GridInfo       =   ""
                     AccessibleName  =   ""
                     AccessibleDescription=   ""
                     AccessibleValue =   ""
                     AccessibleRole  =   9
                     Begin MSComCtl2.DTPicker DTPickerAccFrom 
                        BeginProperty DataFormat 
                           Type            =   1
                           Format          =   "dd/MM/yyyy"
                           HaveTrueFalseNull=   0
                           FirstDayOfWeek  =   0
                           FirstWeekOfYear =   0
                           LCID            =   11265
                           SubFormatType   =   3
                        EndProperty
                        Height          =   345
                        Left            =   90
                        TabIndex        =   12
                        ToolTipText     =   "„‰  «—ÌŒ ﬁœÌ„"
                        Top             =   240
                        Width           =   1500
                        _ExtentX        =   2646
                        _ExtentY        =   609
                        _Version        =   393216
                        CalendarBackColor=   -2147483624
                        CalendarTitleBackColor=   10383715
                        CheckBox        =   -1  'True
                        CustomFormat    =   "yyyy/M/d"
                        Format          =   131137539
                        CurrentDate     =   37357
                     End
                     Begin MSComCtl2.DTPicker DTPickerAccTo 
                        BeginProperty DataFormat 
                           Type            =   1
                           Format          =   "dd/MM/yyyy"
                           HaveTrueFalseNull=   0
                           FirstDayOfWeek  =   0
                           FirstWeekOfYear =   0
                           LCID            =   11265
                           SubFormatType   =   3
                        EndProperty
                        Height          =   345
                        Left            =   90
                        TabIndex        =   13
                        ToolTipText     =   " ≈·Ï  «—ÌŒ √ÕœÀ"
                        Top             =   600
                        Width           =   1500
                        _ExtentX        =   2646
                        _ExtentY        =   609
                        _Version        =   393216
                        CalendarBackColor=   -2147483624
                        CalendarTitleBackColor=   10383715
                        CheckBox        =   -1  'True
                        CustomFormat    =   "yyyy/M/d"
                        Format          =   131137539
                        CurrentDate     =   37357
                     End
                     Begin VB.Label lbl 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00E2E9E9&
                        Caption         =   "„‰"
                        Height          =   285
                        Index           =   4
                        Left            =   1590
                        RightToLeft     =   -1  'True
                        TabIndex        =   15
                        Top             =   285
                        Width           =   555
                     End
                     Begin VB.Label lbl 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00E2E9E9&
                        Caption         =   "≈·Ï"
                        Height          =   285
                        Index           =   2
                        Left            =   1590
                        RightToLeft     =   -1  'True
                        TabIndex        =   14
                        Top             =   600
                        Width           =   555
                     End
                  End
                  Begin ImpulseButton.ISButton CmdAccount 
                     Height          =   405
                     Left            =   120
                     TabIndex        =   17
                     Top             =   4440
                     Width           =   1305
                     _ExtentX        =   2302
                     _ExtentY        =   714
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
                     ButtonImage     =   "FrmProductionReport.frx":038A
                     ColorButton     =   14871017
                     ColorHoverText  =   16777215
                     DrawFocusRectangle=   0   'False
                     ColorToggledHoverText=   16777215
                  End
                  Begin MSDataListLib.DataCombo DcboStores 
                     Height          =   315
                     Left            =   3000
                     TabIndex        =   26
                     Top             =   3270
                     Width           =   3255
                     _ExtentX        =   5741
                     _ExtentY        =   556
                     _Version        =   393216
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin MSDataListLib.DataCombo dcGroups 
                     Height          =   315
                     Left            =   3000
                     TabIndex        =   30
                     Top             =   3630
                     Width           =   3255
                     _ExtentX        =   5741
                     _ExtentY        =   556
                     _Version        =   393216
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin MSDataListLib.DataCombo DBCboClientName 
                     Height          =   315
                     Left            =   3000
                     TabIndex        =   43
                     Top             =   3990
                     Width           =   3255
                     _ExtentX        =   5741
                     _ExtentY        =   556
                     _Version        =   393216
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackStyle       =   0  'Transparent
                     Caption         =   "—ﬁ„ √„— «·«‰ «Ã"
                     ForeColor       =   &H00000000&
                     Height          =   285
                     Index           =   0
                     Left            =   4710
                     RightToLeft     =   -1  'True
                     TabIndex        =   21
                     Top             =   4680
                     Visible         =   0   'False
                     Width           =   1335
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackStyle       =   0  'Transparent
                     Caption         =   "—ﬁ„ «·ÿ·»Ì…"
                     ForeColor       =   &H00000000&
                     Height          =   285
                     Index           =   17
                     Left            =   4710
                     RightToLeft     =   -1  'True
                     TabIndex        =   19
                     Top             =   4320
                     Visible         =   0   'False
                     Width           =   1335
                  End
                  Begin VB.Label LblAccountName 
                     Alignment       =   2  'Center
                     BackColor       =   &H00C0C8C0&
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H000000FF&
                     Height          =   405
                     Left            =   930
                     RightToLeft     =   -1  'True
                     TabIndex        =   16
                     Top             =   345
                     Width           =   7230
                  End
               End
            End
         End
      End
   End
End
Attribute VB_Name = "frmProductionreport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StrAccountCode As String
Dim StrAccountName As String

Private Sub ChangeLang()
    'Label1.Caption = "Des"

    Me.Caption = "Accounting Reports"
    Me.MainTab.TabCaption(0) = "Financial Statements"
    OptAccount(0).Caption = "Report Produced Items..."
 OptAccount(13).Caption = "Audituing Production Orders"
 Command1.Caption = "Costing Reports"
    OptAccount(8).Caption = "General  Damaged Items.."
    Chekopt.Caption = "Store"
    chkGroup.Caption = "Group"
    ChkMain.Caption = "Print According To Big Report"
 
    OptAccount(9).Caption = "ˆProduction Auditing"
    OptAccount(10).Caption = " Compare between the expected and actual per order  "
    OptAccount(11).Caption = "job orders that will not close "
 
 OptAccount(12).Caption = "Production Order Audit "
 OptAccount(7).Caption = "Regural Production   "
    Ele(1).Caption = "In"
    lbl(4).Caption = "From"
    lbl(2).Caption = "To"
    CmdAccount.Caption = "&Print"
 
    Cmd.Caption = "Exit"

End Sub

Private Sub Cmd_Click()
    Unload Me
End Sub

Private Sub CmdAccount_Click()
    Dim i As Integer
   ' Dim StrReportTitle1 As String
    
    Dim cAccountReport As ClsAccReports

    For i = 0 To Me.OptAccount.count - 1

        If Me.OptAccount(i).value = True Then Exit For
    Next i
 
    Select Case i
    
    Case 13
    
        On Error GoTo ErrTrap
Dim SaleReport As ClsSaleReport
  
        Set SaleReport = New ClsSaleReport
        SaleReport.ShowPrice 1, 62, , , , , , DTPickerAccFrom.value, DTPickerAccTo.value
 

    Exit Sub
ErrTrap:

    
Case 12

            Screen.MousePointer = 11
            Set cAccountReport = New ClsAccReports
            cAccountReport.BegineDate = Me.DTPickerAccFrom.value
            cAccountReport.EndDate = Me.DTPickerAccTo.value
                         StrReportTitle1 = " ﬁ—Ì— —«Ã⁄Â «Ê«„— «·«‰ «Ã "
            If chkGroup.value = vbChecked Then
                  
                If Me.dcGroups.BoundText = "" Then
                    Msg = "ÌÃ» «Œ Ì«— «”„ «·„Ã„Ê⁄Â...!!" & CHR(13)
                    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                    dcGroups.SetFocus
                    SendKeys "{F4}"
                    Exit Sub
                End If
            
            End If
    
            If Chekopt.value = vbChecked Then
        
                If Me.DcboStores.BoundText = "" Then
                    Msg = "ÌÃ» «Œ Ì«— «”„ «·„Œ“‰...!!" & CHR(13)
                    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                    DcboStores.SetFocus
                    SendKeys "{F4}"
                    Exit Sub
                    
                End If
    
            End If
    
            cAccountReport.ShowProductItems ChkMain.value, val(dcGroups.BoundText), dcGroups.Text, val(DcboStores.BoundText), DcboStores.Text, 1, Me.DTPickerAccFrom.value, Me.DTPickerAccTo.value
            Set cAccountReport = Nothing
            
        Case 0
            Screen.MousePointer = 11
            Set cAccountReport = New ClsAccReports
            cAccountReport.BegineDate = Me.DTPickerAccFrom.value
            cAccountReport.EndDate = Me.DTPickerAccTo.value
                                
            If chkGroup.value = vbChecked Then
                  
                If Me.dcGroups.BoundText = "" Then
                    Msg = "ÌÃ» «Œ Ì«— «”„ «·„Ã„Ê⁄Â...!!" & CHR(13)
                    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                    dcGroups.SetFocus
                    SendKeys "{F4}"
                    Exit Sub
                End If
            
            End If
    
            If Chekopt.value = vbChecked Then
        
                If Me.DcboStores.BoundText = "" Then
                    Msg = "ÌÃ» «Œ Ì«— «”„ «·„Œ“‰...!!" & CHR(13)
                    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                    DcboStores.SetFocus
                    SendKeys "{F4}"
                    Exit Sub
                End If
    
            End If
    
            cAccountReport.ShowProductItems ChkMain.value, val(dcGroups.BoundText), dcGroups.Text, val(DcboStores.BoundText), DcboStores.Text
            Set cAccountReport = Nothing
                
        Case 8
            Screen.MousePointer = 11
            Set cAccountReport = New ClsAccReports
            cAccountReport.BegineDate = Me.DTPickerAccFrom.value
            cAccountReport.EndDate = Me.DTPickerAccTo.value
                
            If Chekopt.value = vbChecked Then
                
                If Me.DcboStores.BoundText = "" Then
                    Msg = "ÌÃ» «Œ Ì«— «”„ «·„Œ“‰...!!" & CHR(13)
                    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                    DcboStores.SetFocus
                    SendKeys "{F4}"
                    Exit Sub
                End If
 
                cAccountReport.ShowDamagesItems val(Me.DcboStores.BoundText)
            Else
                cAccountReport.ShowDamagesItems
            End If
                
            Set cAccountReport = Nothing
            
        Case 9
             Screen.MousePointer = 11
            Set cAccountReport = New ClsAccReports
            cAccountReport.BegineDate = Me.DTPickerAccFrom.value
            cAccountReport.EndDate = Me.DTPickerAccTo.value
    
            cAccountReport.ProductionAuditing
            Set cAccountReport = Nothing
            
        Case 10
            Screen.MousePointer = 11
            Set cAccountReport = New ClsAccReports
            cAccountReport.BegineDate = Me.DTPickerAccFrom.value
            cAccountReport.EndDate = Me.DTPickerAccTo.value
    
            cAccountReport.ProductionAuditing1
            Set cAccountReport = Nothing
                
        Case 11
            Screen.MousePointer = 11
            Set cAccountReport = New ClsAccReports
            cAccountReport.BegineDate = Me.DTPickerAccFrom.value
            cAccountReport.EndDate = Me.DTPickerAccTo.value
    
            cAccountReport.ProductionAuditing1 True
                
            Set cAccountReport = Nothing
                
        Case 1

            If Not IsNumeric(Text1.Text) Then MsgBox "·«»œ „‰ «œŒ«· —ﬁ„ «·«„—": Exit Sub
            Set cAccountReport = New ClsAccReports
            Screen.MousePointer = 11
            cAccountReport.ShowProductOrderExpenses val(Text1.Text)
            Set cAccountReport = Nothing

        Case 2

            If Not IsNumeric(txt_ORDER_NO.Text) Then MsgBox "·«»œ „‰ «œŒ«· —ﬁ„ «·ÿ·»Ì…": Exit Sub
            Set cAccountReport = New ClsAccReports
            Screen.MousePointer = 11
            cAccountReport.ShowOrdersiTEMS txt_ORDER_NO.Text
            Set cAccountReport = Nothing

        Case 3
            Set cAccountReport = New ClsAccReports
            Screen.MousePointer = 11
            cAccountReport.ShowOrdersStatus txt_ORDER_NO.Text
            Set cAccountReport = Nothing

        Case 7

            If checkApility("FrmProductionReport") = False Then
                Exit Sub
            End If
        
            If SystemOptions.TypicalProduction = True Then
                Screen.MousePointer = 11
                Set cAccountReport = New ClsAccReports
        
                cAccountReport.BegineDate = Me.DTPickerAccFrom.value
                cAccountReport.EndDate = Me.DTPickerAccTo.value
                cAccountReport.ShowProductionSummury
                Set cAccountReport = Nothing
            Else

                If IsNull(Me.DTPickerAccFrom.value) Or IsNull(Me.DTPickerAccTo.value) Then MsgBox "Õœœ › —…", vbCritical: Exit Sub
                Screen.MousePointer = 11
                Set cAccountReport = New ClsAccReports
                CreateReportForProduction Me.DTPickerAccFrom.value, Me.DTPickerAccTo.value
                cAccountReport.BegineDate = Me.DTPickerAccFrom.value
                cAccountReport.EndDate = Me.DTPickerAccTo.value
                cAccountReport.ShowProductionSummury2
                Set cAccountReport = Nothing
            End If
         Case 14
         print_reportProductAct
         Case 15
         print_reportProductAct2
        Case 16
            print_reportProductOut
        Case 17
            print_reportProductMach
        Case 18
            print_reportProductEmp
        Case 19
            print_reportItemDef "", 0
        Case 20
            print_reportItemDef "", 1
    End Select

    CuurentLogdata

End Sub



Function print_reportItemDef(Optional NoteSerial As String, Optional indexe As Integer)
    
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
    If indexe = 0 Then
        MySQL = "  SELECT dbo.TblDefComItem.ID,"
        MySQL = MySQL & "         TblDefComItem.RecordDate,"
        MySQL = MySQL & "         dbo.TblDefComItem.RecordDate,"
        MySQL = MySQL & "         dbo.TblDefComItem.StoreID,"
        MySQL = MySQL & "         TblStore_2.StoreName,"
        MySQL = MySQL & "         TblStore_2.StoreNamee,"
        MySQL = MySQL & "         dbo.TblDefComItem.StoreID2,"
        MySQL = MySQL & "         TblStore_1.StoreName   AS StoreNam2,"
        MySQL = MySQL & "         TblStore_1.StoreNamee  AS StoreNamee3,"
        MySQL = MySQL & "         dbo.TblDefComItem.StoreID3,"
        MySQL = MySQL & "         TblStore_2.StoreName   AS StoreName3,"
        MySQL = MySQL & "         TblStore_2.StoreNamee  AS StoreNamee4,"
        MySQL = MySQL & "         dbo.TblDefComItem.CusID,"
        MySQL = MySQL & "         dbo.TblCustemers.CusName,"
        MySQL = MySQL & "         dbo.TblCustemers.CusNamee,"
        MySQL = MySQL & "         dbo.TblDefComItem.MaxNo,"
        MySQL = MySQL & "         dbo.TblDefComItem.MaxName,"
        MySQL = MySQL & "         TblDefComItem.TotalWithVat,"
        MySQL = MySQL & "         dbo.TblDefComItem.Allocated,"
        MySQL = MySQL & "         dbo.TblDefComItem.AlloPay,"
        MySQL = MySQL & "         dbo.TblDefComItem.AlloRecep,"
        MySQL = MySQL & "         dbo.TblDefComItem.ID   AS IDMain,"
        MySQL = MySQL & "         dbo.TblDefComItem.BranchID,"
        MySQL = MySQL & "         dbo.TblBranchesData.branch_name,"
        MySQL = MySQL & "         dbo.TblBranchesData.branch_nameE,"
        MySQL = MySQL & "         STATUS = ISNULL("
        MySQL = MySQL & "             ("
        MySQL = MySQL & "                 SELECT TOP 1 ItemNameID"
        MySQL = MySQL & "                 From TblProductLineDistribution"
        MySQL = MySQL & "                 Where IDDefCIT = TblDefComItem.ID"
        MySQL = MySQL & "             ),"
        MySQL = MySQL & "  0"
        MySQL = MySQL & "         )"
        MySQL = MySQL & "  From dbo.TblBranchesData"
        MySQL = MySQL & "         RIGHT OUTER JOIN dbo.TblDefComItem"
        MySQL = MySQL & "              ON  dbo.TblBranchesData.branch_id = dbo.TblDefComItem.BranchID"
        MySQL = MySQL & "         LEFT OUTER JOIN dbo.TblCustemers"
        MySQL = MySQL & "              ON  dbo.TblDefComItem.CusID = dbo.TblCustemers.CusID"
        MySQL = MySQL & "         LEFT OUTER JOIN dbo.TblStore TblStore_2"
        MySQL = MySQL & "              ON  dbo.TblDefComItem.StoreID3 = TblStore_2.StoreID"
        MySQL = MySQL & "         LEFT OUTER JOIN dbo.TblStore TblStore_3"
        MySQL = MySQL & "              ON  dbo.TblDefComItem.StoreID = TblStore_3.StoreID"
        MySQL = MySQL & "         LEFT OUTER JOIN dbo.TblStore TblStore_1"
        MySQL = MySQL & "              ON  dbo.TblDefComItem.StoreID2 = TblStore_1.StoreID"
                
    MySQL = MySQL & "  Where 1 = 1"
    
    If DBCboClientName.Text <> "" And chkCust.value = vbChecked Then
            MySQL = MySQL & " and  dbo.TblDefComItem.CusID =" & val(DBCboClientName.BoundText) & ""
    End If
    
        If Not IsNull(DTPickerAccFrom.value) Then
            MySQL = MySQL & " and  dbo.TblDefComItem.RecordDate >=" & SQLDate(DTPickerAccFrom.value, True) & ""
        End If
        
        If Not IsNull(DTPickerAccTo.value) Then
            MySQL = MySQL & " and dbo.TblDefComItem.RecordDate <=" & SQLDate(DTPickerAccTo.value, True) & ""
        End If
        
         If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepDefCompItemTotal.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepDefCompItemTotal.rpt"
        End If
   Else
    If SystemOptions.UserInterface = EnglishInterface Then
     MySQL = ""
    
    
    MySQL = " SELECT DISTINCT T.ID,TblDefComItem.RecordDate,TblDefComItem.RecDate, IsPrinted = (           CASE ISNULL(PrintDate, '') WHEN  '' THEN 0 ELSE 1 END),T.IDDefCIT,T2.FormPrint,T.ProductLineID,T.LineID, T2.name   ProductLineName,T.SalesID,case  T2.IsBasicLine When 1 Then 1 Else 2 End as LineType,"
    MySQL = MySQL & "       T.GroupID, g.GroupNamee GroupName,tblItems.ItemNamee ItemName,TblItems.ItemCode,tblItems.lowering,tblItems.increase,"
    MySQL = MySQL & "       T.ItemNameID,T.UnitId,tu.UnitNamee UnitName,T.Qty,T.Qty1,T.PrintDate ,T.PrintTime,"
    MySQL = MySQL & "       t.DateStart ,T2.StoreID,"
    MySQL = MySQL & "       Start =case IsNull(t.DateStart,0) When 0 then 0 else 1 end, "
    MySQL = MySQL & "       [End] =case IsNull(t.DateEnd,0) When 0 then 0 else 1 end "
    MySQL = MySQL & "       ,  t.DateEnd,TimeEnd,TimeStart,T.UserId,Users.UserName,TblCustemers.CusNamee CusName,TblDefComItem.NoteSerial13,"
    MySQL = MySQL & "       LineID22 =T.LineID ,"
    'MySQL = MySQL & "       LineID2 = (select Top 1 LineID FROM TblDefComItemData DD Where DD.IDDefCIT = T.IDDefCIT and  T.ItemNameID= DD.ItemID ),"
    MySQL = MySQL & "       widtj = (select Top 1 widtj FROM TblDefComItemData DD Where DD.IDDefCIT = T.IDDefCIT and  T.ItemNameID= DD.ItemID ),"
    MySQL = MySQL & "       hight = (select Top 1 hight FROM TblDefComItemData DD Where DD.IDDefCIT = T.IDDefCIT and  T.ItemNameID= DD.ItemID ),"
    'MySQL = MySQL & "       lowering = (select Top 1 lowering FROM TblDefComItemData DD Where DD.IDDefCIT = T.IDDefCIT and  T.ItemNameID= DD.ItemID ),"
    'MySQL = MySQL & "       increase = (select Sum( increase) FROM TblDefComItemDet DD Where DD.IDDefCIT = T.IDDefCIT and  T.ItemNameID= DD.ItemID2  and IsNull(DD.IsDeleted,0) <> 1 ),"
    
    MySQL = MySQL & "       BasedLineName = (Select  Name From TblProductLine Where TblProductLine.Id = BaseProductLineID),BaseProductLineID,"
    'MySQL = MySQL & "       BaseProductLineID2 = (Select ProductLineId From TblGroupItemProductLine Where TblGroupItemProductLine.GroupID = T.GroupID),"
    
    MySQL = MySQL & "       BuiltinItemName = (select Top 1 tblItems.ItemNamee ItemName FROM TblDefComItemData DD Inner Join tblItems On DD.BuiltinItemID =tblItems.ItemID  Where DD.IDDefCIT = T.IDDefCIT and  T.ItemNameID= DD.ItemID )"
    
    
Else
     MySQL = ""
    
    
    MySQL = " SELECT DISTINCT T.ID,TblDefComItem.RecordDate,TblDefComItem.RecDate,IsPrinted = (           CASE ISNULL(PrintDate, '') WHEN  '' THEN 0 ELSE 1 END),T.IDDefCIT,T2.FormPrint,T.ProductLineID,T.LineID, T2.name  ProductLineName,T.SalesID,case  T2.IsBasicLine When 1 Then 1 Else 2 End as LineType,"
    MySQL = MySQL & "       T.GroupID, g.GroupName ,tblItems.ItemName,TblItems.ItemCode,tblItems.lowering,tblItems.increase,"
    MySQL = MySQL & "       T.ItemNameID,T.UnitId,tu.UnitName,T.Qty,T.Qty1,T.PrintDate ,T.PrintTime,"
    MySQL = MySQL & "       t.DateStart ,T2.StoreID,"
    MySQL = MySQL & "       Start =case IsNull(t.DateStart,0) When 0 then 0 else 1 end, "
    MySQL = MySQL & "       [End] =case IsNull(t.DateEnd,0) When 0 then 0 else 1 end "
    MySQL = MySQL & "       ,  t.DateEnd,TimeEnd,TimeStart,T.UserId,Users.UserName,TblCustemers.CusName,TblDefComItem.CusID, TblDefComItem.NoteSerial13,"
    MySQL = MySQL & "       LineID22 =T.LineID ,"
    'MySQL = MySQL & "       LineID2 = (select Top 1 LineID FROM TblDefComItemData DD Where DD.IDDefCIT = T.IDDefCIT and  T.ItemNameID= DD.ItemID ),"
    MySQL = MySQL & "       widtj = (select Top 1 widtj FROM TblDefComItemData DD Where DD.IDDefCIT = T.IDDefCIT and  T.ItemNameID= DD.ItemID ),"
    MySQL = MySQL & "       hight = (select Top 1 hight FROM TblDefComItemData DD Where DD.IDDefCIT = T.IDDefCIT and  T.ItemNameID= DD.ItemID ),"
    'MySQL = MySQL & "       lowering = (select Top 1 lowering FROM TblDefComItemData DD Where DD.IDDefCIT = T.IDDefCIT and  T.ItemNameID= DD.ItemID ),"
    'MySQL = MySQL & "       increase = (select Sum( increase) FROM TblDefComItemDet DD Where DD.IDDefCIT = T.IDDefCIT and  T.ItemNameID= DD.ItemID2  and IsNull(DD.IsDeleted,0) <> 1 ),"
    
    MySQL = MySQL & "       BasedLineName = (Select Name From TblProductLine Where TblProductLine.Id = BaseProductLineID),BaseProductLineID,"
    'MySQL = MySQL & "       BaseProductLineID2 = (Select ProductLineId From TblGroupItemProductLine Where TblGroupItemProductLine.GroupID = T.GroupID),"
    
    MySQL = MySQL & "       BuiltinItemName = (select Top 1 tblItems.ItemName FROM TblDefComItemData DD Inner Join tblItems On DD.BuiltinItemID =tblItems.ItemID  Where DD.IDDefCIT = T.IDDefCIT and  T.ItemNameID= DD.ItemID )"
End If
MySQL = MySQL & " FROM   TblProductLineDistribution       AS T "
MySQL = MySQL & "       Inner JOIN TblProductLine  AS T2"
MySQL = MySQL & "            ON  T2.id = T.ProductLineID"

MySQL = MySQL & "       Inner JOIN TblDefComItem "
MySQL = MySQL & "            ON  TblDefComItem.id = T.IDDefCIT"
MySQL = MySQL & "       Left Outer JOIN TblCustemers "
MySQL = MySQL & "            ON  TblDefComItem.CusID= TblCustemers.CusID"

MySQL = MySQL & "       Inner JOIN tblItems"
MySQL = MySQL & "            ON  tblItems.ItemID = T.ItemNameID"
MySQL = MySQL & "       Inner JOIN TblUnites        AS tu"
MySQL = MySQL & "            ON  tu.UnitID = T.UnitID"
MySQL = MySQL & "       LEFT OUTER JOIN Groups           AS g"
MySQL = MySQL & "            ON  G.GroupID = T.GroupID"
MySQL = MySQL & "       LEFT OUTER JOIN TblUsers            AS Users"
MySQL = MySQL & "            ON  Users.UserId = T.UserId"

MySQL = MySQL & "       LEFT OUTER JOIN TblUsersProductLine            "
MySQL = MySQL & "            ON  TblUsersProductLine.ProductLineId = T2.Id"

MySQL = MySQL & "         Where IsNull(T.Qty,0) <> 0   "
'T.ProductLineID In   (SELECT LineID FROM TblProductLineWorker "
'MySQL = MySQL & "         WHERE EmpId IN (SELECT EmpID FROM TblUsers WHERE UserID = " & user_id & ")) and "

   
    
    If DBCboClientName.Text <> "" And chkCust.value = vbChecked Then
            MySQL = MySQL & " and  dbo.TblDefComItem.CusID =" & val(DBCboClientName.BoundText) & ""
    End If
    
        If Not IsNull(DTPickerAccFrom.value) Then
            MySQL = MySQL & " and  dbo.TblDefComItem.RecordDate >=" & SQLDate(DTPickerAccFrom.value, True) & ""
        End If
        
        If Not IsNull(DTPickerAccTo.value) Then
            MySQL = MySQL & " and dbo.TblDefComItem.RecordDate <=" & SQLDate(DTPickerAccTo.value, True) & ""
        End If
    
         MySQL = MySQL & " Order By TblDefComItem.RecordDate Desc"
         
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepDefCompItemTotal2.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepDefCompItemTotal2.rpt"
        End If
   End If
    
       

        ''''''


    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

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
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(val(dcBranch.BoundText)))
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


  Function print_reportProductAct()
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
   If IsNull(Me.DTPickerAccFrom.value) Or IsNull(Me.DTPickerAccTo.value) Then
   If SystemOptions.UserInterface = ArabicInterface Then
   MsgBox "Ì—ÃÏ  ÕœÌœ «·› —…"
   Else
   MsgBox "Please select period"
   End If
   Exit Function
   End If
    MySQL = " SELECT     dbo.TbllProductionPlan.TbllProductionPlanD, dbo.TbllProductionPlan.FromDate, dbo.TbllProductionPlan.Todate, dbo.TbllProductionPlanDetails.ItemID, "
    MySQL = MySQL & "                   dbo.TblItems.ItemName, dbo.TblItems.ItemNamee, dbo.TblItems.Fullcode, dbo.TblItems.GroupID, dbo.TbllProductionPlanDetails.Price as Cunt,"
    MySQL = MySQL & "                  dbo.GetQtyRecevPlanProd(dbo.Transactions.Transaction_Serial, dbo.TbllProductionPlanDetails.ItemID, " & SQLDate(DTPickerAccFrom.value, True) & ","
    MySQL = MySQL & "                  " & SQLDate(DTPickerAccTo.value, True) & ",28) AS ActQty, dbo.TbllProductionPlan.StoreID"
    MySQL = MySQL & "    FROM         dbo.Transactions RIGHT OUTER JOIN"
    MySQL = MySQL & "                  dbo.TbllProductionPlan ON dbo.Transactions.ProductionPlanno = dbo.TbllProductionPlan.TbllProductionPlanD LEFT OUTER JOIN"
    MySQL = MySQL & "                  dbo.TblItems RIGHT OUTER JOIN"
    MySQL = MySQL & "                  dbo.TbllProductionPlanDetails ON dbo.TblItems.ItemID = dbo.TbllProductionPlanDetails.ItemID ON"
    MySQL = MySQL & "                  dbo.TbllProductionPlan.TbllProductionPlanD = dbo.TbllProductionPlanDetails.TbllProductionPlanD"
    MySQL = MySQL & "    Where (1 = 1)"
    If Not IsNull(DTPickerAccFrom.value) Then
        MySQL = MySQL & " and  dbo.TbllProductionPlan.FromDate >=" & SQLDate(DTPickerAccFrom.value, True) & ""
    End If
    
    If Not IsNull(DTPickerAccTo.value) Then
        MySQL = MySQL & " and dbo.TbllProductionPlan.FromDate <=" & SQLDate(DTPickerAccTo.value, True) & ""
    End If
    
    If Not IsNull(DTPickerAccFrom.value) Then
        MySQL = MySQL & " and  dbo.TbllProductionPlan.Todate >=" & SQLDate(DTPickerAccFrom.value, True) & ""
    End If
    
    If Not IsNull(DTPickerAccTo.value) Then
        MySQL = MySQL & " and dbo.TbllProductionPlan.Todate <=" & SQLDate(DTPickerAccTo.value, True) & ""
    End If
    
    If val(Me.DcboStores.BoundText) <> 0 Then
        MySQL = MySQL & " and  dbo.TbllProductionPlan.StoreID =" & val(DcboStores.BoundText) & ""
    End If
    If val(Me.dcGroups.BoundText) <> 0 Then
        MySQL = MySQL & " and  dbo.TblItems.GroupID =" & val(dcGroups.BoundText) & ""
    End If
        MySQL = MySQL & " order by dbo.TbllProductionPlan.TbllProductionPlanD"
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RptProductActQty.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RptProductActQtyE.rpt"
        End If
    
    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        Else
            Msg = "No Data"
        End If
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
        'xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""
    End If
    xReport.ParameterFields(3).AddCurrentValue user_name
    If Not IsNull(DTPickerAccFrom.value) Then
    xReport.ParameterFields(4).AddCurrentValue DTPickerAccFrom.value
    End If
    If Not IsNull(DTPickerAccTo.value) Then
    xReport.ParameterFields(5).AddCurrentValue DTPickerAccTo.value
    End If
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


 Function print_reportProductOut()
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
   If IsNull(Me.DTPickerAccFrom.value) Or IsNull(Me.DTPickerAccTo.value) Then
   If SystemOptions.UserInterface = ArabicInterface Then
   MsgBox "Ì—ÃÏ  ÕœÌœ «·› —…"
   Else
   MsgBox "Please select period"
   End If
   Exit Function
   End If
  
    MySQL = "    SELECT dbo.Transactions.Transaction_Type,"
MySQL = MySQL & "           dbo.Transactions.NoteSerial1,"
MySQL = MySQL & "           dbo.Transactions.NoteSerial2,"
MySQL = MySQL & "           dbo.Transactions.Transaction_ID,"
MySQL = MySQL & "           dbo.Transactions.Transaction_Serial,"
MySQL = MySQL & "           dbo.Transactions.Transaction_Date,"
MySQL = MySQL & "           ti.ItemName CusName,"
MySQL = MySQL & "           dbo.Transactions.BranchId,"
MySQL = MySQL & "           dbo.TblBranchesData.branch_name,"
MySQL = MySQL & "           dbo.TblBranchesData.branch_namee,"
MySQL = MySQL & "           Transactions.remark CashCustomerComment,"
MySQL = MySQL & "           dbo.Transactions.NoteSerial1,"
MySQL = MySQL & "           Transaction_Details.ShowQty Transaction_NetValue,"
MySQL = MySQL & "           ("
MySQL = MySQL & "               SELECT SUM(a2.ShowQty) AS SumValue"
MySQL = MySQL & "               FROM   dbo.Transaction_Details AS A2"
MySQL = MySQL & "                      INNER JOIN Transactions A"
MySQL = MySQL & "                           ON  A.Transaction_ID = A2.Transaction_ID"
MySQL = MySQL & "               Where (a.Transaction_Type = 28)"
MySQL = MySQL & "                      AND a2.Item_ID = Transaction_Details.Item_ID"
MySQL = MySQL & "                      AND ISNULL(a.BillBasedOn, 2) = 2"
MySQL = MySQL & "                            AND ( a.NoteSerial1 = Transactions.Product_Receive_voucher_Serial OR a.order_no = Transactions.NoteSerial1)"
MySQL = MySQL & "           )                         AS RetValue,"
MySQL = MySQL & "           PayedVal = Transaction_Details.ShowQty -IsNull(("
MySQL = MySQL & "               SELECT SUM(a2.ShowQty) AS SumValue"
MySQL = MySQL & "               FROM   dbo.Transaction_Details AS A2"
MySQL = MySQL & "                      INNER JOIN Transactions A"
MySQL = MySQL & "                           ON  A.Transaction_ID = A2.Transaction_ID"
MySQL = MySQL & "               Where (a.Transaction_Type = 28)"
MySQL = MySQL & "                      AND a2.Item_ID = Transaction_Details.Item_ID"
'MySQL = MySQL & "                      AND ISNULL(a.BillBasedOn, 2) = 2"
MySQL = MySQL & "                           AND ( a.NoteSerial1 = Transactions.Product_Receive_voucher_Serial OR a.order_no = Transactions.NoteSerial1)"
MySQL = MySQL & "           ),0)"
MySQL = MySQL & "                   From dbo.transactions"
MySQL = MySQL & "                          LEFT OUTER JOIN dbo.TblBranchesData"
MySQL = MySQL & "                               ON  dbo.Transactions.BranchId = dbo.TblBranchesData.branch_id"
MySQL = MySQL & "                          INNER JOIN dbo.Transaction_Details"
MySQL = MySQL & "                               ON  dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID"
MySQL = MySQL & "                          LEFT OUTER JOIN TblItems  AS ti"
MySQL = MySQL & "                               ON  ti.ItemID = dbo.Transaction_Details.Item_ID"


    MySQL = MySQL & "  WHERE     (dbo.Transactions.Transaction_Type = 21) "
     If BrnchIDes <> "-1" And BrnchIDes <> "" Then
       ' MySQL = MySQL & "   and dbo.Transactions.BranchId in(" & BrnchIDes & ")"
     End If
  
    If Not IsNull(DTPickerAccFrom.value) Then
        MySQL = MySQL & " and  dbo.Transactions.Transaction_Date >=" & SQLDate(DTPickerAccFrom.value, True) & ""
    End If
    
    If Not IsNull(DTPickerAccTo.value) Then
        MySQL = MySQL & " and  dbo.Transactions.Transaction_Date <=" & SQLDate(DTPickerAccTo.value, True) & ""
    End If
    

    If val(Me.DcboStores.BoundText) <> 0 Then
        MySQL = MySQL & " and  dbo.Transactions.StoreID =" & val(DcboStores.BoundText) & ""
    End If
    If val(Me.dcGroups.BoundText) <> 0 Then
        MySQL = MySQL & " and  dbo.TblItems.GroupID =" & val(dcGroups.BoundText) & ""
    End If
       
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rptItemStatusOut.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rptItemStatusOut.rpt"
        End If
    
    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        Else
            Msg = "No Data"
        End If
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
        'xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""
    End If
    xReport.ParameterFields(3).AddCurrentValue user_name
    If Not IsNull(DTPickerAccFrom.value) Then
    'xReport.ParameterFields(4).AddCurrentValue DTPickerAccFrom.value
    End If
    If Not IsNull(DTPickerAccTo.value) Then
  '  xReport.ParameterFields(5).AddCurrentValue DTPickerAccTo.value
    End If
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

 Function print_reportProductMach()
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
   If IsNull(Me.DTPickerAccFrom.value) Or IsNull(Me.DTPickerAccTo.value) Then
   If SystemOptions.UserInterface = ArabicInterface Then
   MsgBox "Ì—ÃÏ  ÕœÌœ «·› —…"
   Else
   MsgBox "Please select period"
   End If
   Exit Function
   End If
  
 
    MySQL = " SELECT DISTINCT tple.Equipmentname CusNamee,"
    MySQL = MySQL & "            tpl.name CusName,"
    MySQL = MySQL & "            TblProductOrderLines.Hour Transaction_NetValue,"
    MySQL = MySQL & "            tblItems.ItemName CashCustomerName,"
    MySQL = MySQL & "            td.Quantity PayedVal"
    MySQL = MySQL & "     From TblProductOrderLines"
    MySQL = MySQL & "            Inner join Transactions"
    MySQL = MySQL & "                ON  Transactions.Transaction_ID = TblProductOrderLines.Transaction_ID"
    MySQL = MySQL & "                AND Transactions.Transaction_Type = 26"
    MySQL = MySQL & "           Inner JOIN Transaction_Details AS td"
    MySQL = MySQL & "                ON  td.Transaction_ID = Transactions.Transaction_ID"
    MySQL = MySQL & "           LEFT OUTER JOIN tblItems"
    MySQL = MySQL & "                ON  tblItems.ItemID = td.Item_ID"
    MySQL = MySQL & "           INNER JOIN TblProductLine  AS tpl"
    MySQL = MySQL & "                ON  TblProductOrderLines.LineID = tpl.id"
    MySQL = MySQL & "           Inner JOIN TblProductLineEquipments AS tple"
    MySQL = MySQL & "                ON  tple.LineID = tpl.id"
    
    
 
     If BrnchIDes <> "-1" And BrnchIDes <> "" Then
       ' MySQL = MySQL & "   and dbo.Transactions.BranchId in(" & BrnchIDes & ")"
     End If
  
    If Not IsNull(DTPickerAccFrom.value) Then
        MySQL = MySQL & " and  dbo.Transactions.Transaction_Date >=" & SQLDate(DTPickerAccFrom.value, True) & ""
    End If
    
    If Not IsNull(DTPickerAccTo.value) Then
        MySQL = MySQL & " and  dbo.Transactions.Transaction_Date <=" & SQLDate(DTPickerAccTo.value, True) & ""
    End If
    

    If val(Me.DcboStores.BoundText) <> 0 Then
        MySQL = MySQL & " and  dbo.Transactions.StoreID =" & val(DcboStores.BoundText) & ""
    End If
    If val(Me.dcGroups.BoundText) <> 0 Then
        MySQL = MySQL & " and  dbo.TblItems.GroupID =" & val(dcGroups.BoundText) & ""
    End If
       
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rpProductionStatusOut.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rpProductionStatusOut.rpt"
        End If
    
    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        Else
            Msg = "No Data"
        End If
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
        'xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""
    End If
    xReport.ParameterFields(3).AddCurrentValue user_name
    If Not IsNull(DTPickerAccFrom.value) Then
    'xReport.ParameterFields(4).AddCurrentValue DTPickerAccFrom.value
    End If
    If Not IsNull(DTPickerAccTo.value) Then
  '  xReport.ParameterFields(5).AddCurrentValue DTPickerAccTo.value
    End If
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


 Function print_reportProductEmp()
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
   If IsNull(Me.DTPickerAccFrom.value) Or IsNull(Me.DTPickerAccTo.value) Then
   If SystemOptions.UserInterface = ArabicInterface Then
   MsgBox "Ì—ÃÏ  ÕœÌœ «·› —…"
   Else
   MsgBox "Please select period"
   End If
   Exit Function
   End If
  
  
   MySQL = " SELECT DISTINCT te.Emp_Name CusNamee,"
   MySQL = MySQL & "           tpl2.name CusName,"
'   MySQL = MySQL & "           TblProductOrderWorker.HourPrice Transaction_NetValue ,"
   MySQL = MySQL & "           tblItems.ItemName CashCustomerName,"
   MySQL = MySQL & "           TblProductOrderWorker.Hour     Transaction_NetValue,TblProductOrderWorker.hourprice * TblProductOrderWorker.Hour AS PayedVal2,"
    MySQL = MySQL & "          td.Quantity PayedVal"
    MySQL = MySQL & "   From TblProductOrderWorker"
    MySQL = MySQL & "    LEFT OUTER JOIN"
    MySQL = MySQL & "    TblProductLineWorker ON TblProductLineWorker.EmpID = TblProductOrderWorker.Emp_id"
    
    MySQL = MySQL & "          inner  JOIN Transactions"
    MySQL = MySQL & "               ON  Transactions.Transaction_ID = TblProductOrderWorker.Transaction_ID"
    MySQL = MySQL & "               AND Transactions.Transaction_Type = 26"
    MySQL = MySQL & "               Inner JOIN TblEmployee AS te ON te.Emp_ID = TblProductOrderWorker.Emp_ID"
    MySQL = MySQL & "          Inner JOIN Transaction_Details AS td"
    MySQL = MySQL & "               ON  td.Transaction_ID = Transactions.Transaction_ID"
    MySQL = MySQL & "                         INNER JOIN TblProductOrderLines"
    MySQL = MySQL & "                         ON TblProductOrderLines.Transaction_ID = Transactions.Transaction_ID"
    MySQL = MySQL & "                            INNER  JOIN TblProductLine      AS tpl2"
    MySQL = MySQL & "                           ON  tpl2.id = TblProductOrderLines.LineID"
    MySQL = MySQL & "          Inner JOIN tblItems"
    MySQL = MySQL & "               ON  tblItems.ItemID = td.Item_ID"
     
             
    
 
 
     If BrnchIDes <> "-1" And BrnchIDes <> "" Then
       ' MySQL = MySQL & "   and dbo.Transactions.BranchId in(" & BrnchIDes & ")"
     End If
  
    If Not IsNull(DTPickerAccFrom.value) Then
        MySQL = MySQL & " and  dbo.Transactions.Transaction_Date >=" & SQLDate(DTPickerAccFrom.value, True) & ""
    End If
    
    If Not IsNull(DTPickerAccTo.value) Then
        MySQL = MySQL & " and  dbo.Transactions.Transaction_Date <=" & SQLDate(DTPickerAccTo.value, True) & ""
    End If
    

    If val(Me.DcboStores.BoundText) <> 0 Then
        MySQL = MySQL & " and  dbo.Transactions.StoreID =" & val(DcboStores.BoundText) & ""
    End If
    If val(Me.dcGroups.BoundText) <> 0 Then
        MySQL = MySQL & " and  dbo.TblItems.GroupID =" & val(dcGroups.BoundText) & ""
    End If
       
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rpProductionStatusEmpOut.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rpProductionStatusEmpOut.rpt"
        End If
    
    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        Else
            Msg = "No Data"
        End If
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
        'xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""
    End If
    xReport.ParameterFields(3).AddCurrentValue user_name
    If Not IsNull(DTPickerAccFrom.value) Then
    'xReport.ParameterFields(4).AddCurrentValue DTPickerAccFrom.value
    End If
    If Not IsNull(DTPickerAccTo.value) Then
  '  xReport.ParameterFields(5).AddCurrentValue DTPickerAccTo.value
    End If
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




  Function print_reportProductAct2()
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
   If IsNull(Me.DTPickerAccFrom.value) Or IsNull(Me.DTPickerAccTo.value) Then
   If SystemOptions.UserInterface = ArabicInterface Then
   MsgBox "Ì—ÃÏ  ÕœÌœ «·› —…"
   Else
   MsgBox "Please select period"
   End If
   Exit Function
   End If
     
   
    MySQL = " SELECT     dbo.Transactions.Transaction_Type,Transaction_Details.ShowPrice , dbo.Transactions.Transaction_Date, dbo.Transactions.NoteSerial1, dbo.Transactions.CusID, dbo.TblCustemers.CusName,IsNull(Transaction_Details.ShowQty,0) * IsNull(TblItemsParts.PartItemQty,0) OldQty ,"
    
    MySQL = MySQL & "                     ActQty = (SELECT SUM(IsNull(TT.ShowQty,0) )"
    MySQL = MySQL & "                          FROM   dbo.Transactions T"
    MySQL = MySQL & "                                 INNER JOIN dbo.Transaction_Details TT"
    MySQL = MySQL & "                                      ON  T.Transaction_ID = TT.Transaction_ID"
    MySQL = MySQL & "                          Where (t.Transaction_Type = 27)"
    MySQL = MySQL & "                                 AND (TT.Item_ID = dbo.TblItemsParts.PartItemID)"
    MySQL = MySQL & "                                 AND (T.WorkOrderNO =  dbo.transactions.NoteSerial1 )"
    MySQL = MySQL & "                  ),"
    
    MySQL = MySQL & "                  ActualCost = (SELECT SUM(IsNull(TT.ShowPrice,0)  * IsNull(TT.ShowQty,0))"
    MySQL = MySQL & "                          FROM   dbo.Transactions T"
    MySQL = MySQL & "                                 INNER JOIN dbo.Transaction_Details TT"
    MySQL = MySQL & "                                      ON  T.Transaction_ID = TT.Transaction_ID"
    MySQL = MySQL & "                          Where (t.Transaction_Type = 27)"
    MySQL = MySQL & "                                 AND (TT.Item_ID = dbo.TblItemsParts.PartItemID)"
    MySQL = MySQL & "                                 AND (T.WorkOrderNO =  dbo.transactions.NoteSerial1 )"
    MySQL = MySQL & "                  ),"
    
    MySQL = MySQL & "                    avCost =("
    MySQL = MySQL & "                        SELECT  CONVERT(Float, Total / TotalQty, 3)"
    MySQL = MySQL & "                     from dbo.QryItemsTransactionsTotals(28, 3, 20, '01/01/1900', ' 01/01/2079 ',  TblItems_1.ItemID, 0))"
    MySQL = MySQL + "       ,TblItemsParts.PartItemQty * ShowQty ActualQty "
    MySQL = MySQL & "  ,                dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode, dbo.Transaction_Details.Item_ID, dbo.Transaction_Details.ShowQty,"
    'MySQL = MySQL & "   dbo.GetQtyRecevPlanProd(dbo.Transactions.NoteSerial1, dbo.TblItemsParts.PartItemID, " & SQLDate(DTPickerAccFrom.value, True) & ", " & SQLDate(DTPickerAccTo.value, True) & ", 27) AS ActQty"
    MySQL = MySQL & "                   dbo.Transactions.StoreID, dbo.TblItems.ItemName, dbo.TblItems.ItemNamee, dbo.TblItems.Fullcode AS Item ,"
    MySQL = MySQL & "                  TblItems_1.ItemName AS PartItemName, TblItems_1.ItemNamee AS PartItemNameE, TblItems_1.Fullcode AS PartFullcode, dbo.TblItemsParts.PartItemID"
    MySQL = MySQL & "     FROM         dbo.TblItemsParts INNER JOIN"
    MySQL = MySQL & "                  dbo.TblItems ON dbo.TblItemsParts.ItemID = dbo.TblItems.ItemID LEFT OUTER JOIN"
    MySQL = MySQL & "                  dbo.TblItems TblItems_1 ON dbo.TblItemsParts.PartItemID = TblItems_1.ItemID RIGHT OUTER JOIN"
    MySQL = MySQL & "                  dbo.Transaction_Details ON dbo.TblItems.ItemID = dbo.Transaction_Details.Item_ID RIGHT OUTER JOIN"
    MySQL = MySQL & "                  dbo.Transactions ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID LEFT OUTER JOIN"
    MySQL = MySQL & "                  dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID"
      MySQL = MySQL & " Where (dbo.transactions.Transaction_Type = 26) and not(dbo.TblItemsParts.PartItemID is null) and dbo.TblItemsParts.PartItemID<>0"
    If Not IsNull(DTPickerAccFrom.value) Then
        MySQL = MySQL & " and  dbo.Transactions.Transaction_Date >=" & SQLDate(DTPickerAccFrom.value, True) & ""
    End If
    
    If Not IsNull(DTPickerAccTo.value) Then
        MySQL = MySQL & " and dbo.Transactions.Transaction_Date <=" & SQLDate(DTPickerAccTo.value, True) & ""
    End If
    
    If val(Me.DcboStores.BoundText) <> 0 Then
        MySQL = MySQL & " and  dbo.Transactions.StoreID =" & val(DcboStores.BoundText) & ""
    End If
    If val(Me.dcGroups.BoundText) <> 0 Then
        MySQL = MySQL & " and  dbo.TblItems.GroupID =" & val(dcGroups.BoundText) & ""
    End If
      MySQL = MySQL & " Order By dbo.transactions.NoteSerial1,Transaction_Details.id "
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RptProductActQty2.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RptProductActQty2E.rpt"
        End If
    
    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        Else
            Msg = "No Data"
        End If
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
        'xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""
    End If
    xReport.ParameterFields(3).AddCurrentValue user_name
    If Not IsNull(DTPickerAccFrom.value) Then
    xReport.ParameterFields(4).AddCurrentValue DTPickerAccFrom.value
    End If
    If Not IsNull(DTPickerAccTo.value) Then
    xReport.ParameterFields(5).AddCurrentValue DTPickerAccTo.value
    End If
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
Private Sub Command1_Click()
FrmAveProdPriceMatrialReports.show
End Sub

Private Sub Form_Load()
    Resize_Form Me, NoChangeInSize
    StrAccountCode = ""
 
    ScreenNameArabic = "  ‹‹ﬁ‹‹«—Ì‹‹‹— «·«‰‰«Ã  "
    ScreenNameEnglish = "  Production Report "
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish, "1"

    Dim StrSQL As String
    Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos
    Dcombos.GetStores DcboStores
 
    Dcombos.GetItemSGroups Me.dcGroups, False
    
    Dcombos.GetCustomersSuppliers 1, Me.DBCboClientName, True
    
    SetDtpickerDate Me.DTPickerAccFrom
    SetDtpickerDate Me.DTPickerAccTo

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
    
End Sub
 
Function CuurentLogdata(Optional Currentmode As String)
    Dim i As Integer
  
    LogTextA = "    ‘«‘… " & ScreenNameArabic & "   ⁄—÷  ﬁ—Ì— "

    For i = 0 To 7

        If OptAccount(i).value = True Then
            LogTextA = LogTextA & OptAccount(i).Caption
        End If
 
    Next i
 
    LogTextA = LogTextA & "    «·› —… „‰  " & DTPickerAccFrom.value & "   «·Ï  " & DTPickerAccTo.value
  
    LogTexte = "    Screen " & ScreenNameEnglish & "   View Report   "

    For i = 0 To 7

        If OptAccount(i).value = True Then
            LogTexte = LogTextA & OptAccount(i).Caption
        End If
 
    Next i
 
    LogTexte = LogTexte & "    From " & DTPickerAccFrom.value & "   To  " & DTPickerAccTo.value
                     
    If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, "V"
    Else
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, "V"
    End If
    
End Function

Private Sub Form_Unload(Cancel As Integer)
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish
End Sub

Private Sub OptAccount_Click(Index As Integer)
DTPickerAccFrom.value = Date
DTPickerAccTo.value = Date
End Sub
