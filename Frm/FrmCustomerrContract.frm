VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form FrmCustomerContract 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "« ›«ﬁÌ«  «·⁄„·«¡"
   ClientHeight    =   8430
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   12135
   BeginProperty Font 
      Name            =   "MS Reference Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   580
   Icon            =   "FrmCustomerrContract.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   8430
   ScaleWidth      =   12135
   Visible         =   0   'False
   Begin C1SizerLibCtl.C1Elastic EleMain 
      Height          =   8415
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   12165
      _cx             =   21458
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
      GridRows        =   5
      GridCols        =   4
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"FrmCustomerrContract.frx":038A
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   2130
         Left            =   30
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   30
         Width           =   12105
         _cx             =   21352
         _cy             =   3757
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
         Begin VB.TextBox TxtCashCustomerName 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   5640
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   84
            Top             =   1800
            Width           =   4890
         End
         Begin VB.TextBox TxtEmployeeID 
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
            Left            =   9360
            TabIndex        =   76
            Top             =   1440
            Width           =   1230
         End
         Begin VB.TextBox TxtSearchCode 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   9360
            RightToLeft     =   -1  'True
            TabIndex        =   72
            Top             =   1140
            Width           =   1215
         End
         Begin VB.TextBox TxtTblCustomerContractD 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   9360
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   38
            Top             =   780
            Width           =   1200
         End
         Begin VB.CheckBox ChkLocked 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«Ìﬁ«› «· ⁄«„·"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   6180
            RightToLeft     =   -1  'True
            TabIndex        =   37
            Top             =   810
            Width           =   2295
         End
         Begin VB.TextBox txtRemarks 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   240
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   36
            Top             =   1275
            Width           =   4410
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   765
            Index           =   5
            Left            =   0
            TabIndex        =   31
            TabStop         =   0   'False
            Top             =   0
            Width           =   12075
            _cx             =   21299
            _cy             =   1349
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   24
               Charset         =   178
               Weight          =   700
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
            Picture         =   "FrmCustomerrContract.frx":040E
            Caption         =   "« ›«ﬁÌ«  «·⁄„·«¡  "
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
               TabIndex        =   32
               Top             =   90
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   661
               ButtonStyle     =   1
               ButtonPositionImage=   4
               Caption         =   ""
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ButtonImage     =   "FrmCustomerrContract.frx":10E8
               ColorButton     =   16777215
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
               TabIndex        =   33
               Top             =   90
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   661
               ButtonStyle     =   1
               ButtonPositionImage=   4
               Caption         =   ""
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ButtonImage     =   "FrmCustomerrContract.frx":1482
               ColorButton     =   16777215
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
               TabIndex        =   34
               Top             =   90
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   661
               ButtonStyle     =   1
               ButtonPositionImage=   4
               Caption         =   ""
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ButtonImage     =   "FrmCustomerrContract.frx":181C
               ColorButton     =   16777215
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
               TabIndex        =   35
               Top             =   90
               Width           =   495
               _ExtentX        =   873
               _ExtentY        =   661
               ButtonStyle     =   1
               ButtonPositionImage=   4
               Caption         =   ""
               BackColor       =   16777215
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ButtonImage     =   "FrmCustomerrContract.frx":1BB6
               ColorButton     =   16777215
               ColorHighlight  =   4194304
               ColorHoverText  =   16777215
               ColorShadow     =   -2147483631
               ColorOutline    =   -2147483631
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
               ColorToggledHoverText=   16777215
               ColorTextShadow =   16777215
            End
            Begin MSComDlg.CommonDialog CD1 
               Left            =   0
               Top             =   0
               _ExtentX        =   847
               _ExtentY        =   847
               _Version        =   393216
            End
         End
         Begin MSComCtl2.DTPicker dbFromDate 
            Height          =   390
            Left            =   2985
            TabIndex        =   39
            Top             =   840
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   688
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Reference Sans Serif"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   98697217
            CurrentDate     =   38784
         End
         Begin MSComCtl2.DTPicker dbTodate 
            Height          =   390
            Left            =   240
            TabIndex        =   40
            Top             =   840
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   688
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Reference Sans Serif"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   98697217
            CurrentDate     =   38784
         End
         Begin MSDataListLib.DataCombo DBCboClientName 
            Height          =   360
            Left            =   5595
            TabIndex        =   41
            Top             =   1140
            Width           =   3735
            _ExtentX        =   6588
            _ExtentY        =   635
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataListLib.DataCombo DcboEmp 
            Height          =   360
            Left            =   5610
            TabIndex        =   77
            Top             =   1440
            Width           =   3720
            _ExtentX        =   6562
            _ExtentY        =   635
            _Version        =   393216
            ListField       =   "7"
            Text            =   ""
            RightToLeft     =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·⁄„Ì· «·‰ﬁœÌ"
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
            Index           =   10
            Left            =   10770
            RightToLeft     =   -1  'True
            TabIndex        =   79
            Top             =   1710
            Width           =   1080
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·»«∆⁄"
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
            Left            =   10770
            RightToLeft     =   -1  'True
            TabIndex        =   78
            Top             =   1440
            Width           =   1080
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "—ﬁ„ «·« ›«ﬁÌ…"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   7
            Left            =   10800
            RightToLeft     =   -1  'True
            TabIndex        =   46
            Top             =   780
            Width           =   1050
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "„œ Â« „‰"
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
            Index           =   5
            Left            =   4575
            RightToLeft     =   -1  'True
            TabIndex        =   45
            Top             =   840
            Width           =   945
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·⁄„Ì·"
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
            Left            =   10770
            RightToLeft     =   -1  'True
            TabIndex        =   44
            Top             =   1095
            Width           =   1080
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·Ï"
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
            Left            =   1785
            RightToLeft     =   -1  'True
            TabIndex        =   43
            Top             =   840
            Width           =   960
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "„·«ÕŸ« "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   3
            Left            =   4650
            RightToLeft     =   -1  'True
            TabIndex        =   42
            Top             =   1155
            Width           =   825
         End
      End
      Begin C1SizerLibCtl.C1Tab C1Tab1 
         Height          =   4800
         Left            =   30
         TabIndex        =   1
         Top             =   2175
         Width           =   12105
         _cx             =   21352
         _cy             =   8467
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
         Caption         =   "«·«’‰«›|«·„Ã„Ê⁄« "
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
         Flags(1)        =   2
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   4380
            Index           =   2
            Left            =   45
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   45
            Width           =   12015
            _cx             =   21193
            _cy             =   7726
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
            Begin VB.CheckBox Check4 
               Alignment       =   1  'Right Justify
               Caption         =   "«œ—«Ã ﬂ· «’‰«› Œÿ… «· ”⁄Ì—"
               Height          =   285
               Left            =   5250
               RightToLeft     =   -1  'True
               TabIndex        =   88
               Top             =   60
               Width           =   2595
            End
            Begin VB.ComboBox CBoBasedON 
               Height          =   345
               ItemData        =   "FrmCustomerrContract.frx":1F50
               Left            =   4290
               List            =   "FrmCustomerrContract.frx":1F52
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   86
               Top             =   345
               Width           =   3540
            End
            Begin VB.TextBox TxtPlanID 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   2850
               TabIndex        =   85
               Top             =   360
               Width           =   1350
            End
            Begin VB.CommandButton Command2 
               Caption         =   " Õ„Ì· «·„·›..."
               BeginProperty Font 
                  Name            =   "MS Reference Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   8580
               RightToLeft     =   -1  'True
               TabIndex        =   83
               Top             =   60
               Width           =   1695
            End
            Begin VB.CommandButton Command1 
               Caption         =   " ÕœÌœ «·„·›..."
               BeginProperty Font 
                  Name            =   "MS Reference Sans Serif"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   10140
               RightToLeft     =   -1  'True
               TabIndex        =   82
               Top             =   60
               Width           =   1815
            End
            Begin VB.TextBox txtFile 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   60
               Locked          =   -1  'True
               TabIndex        =   81
               Top             =   30
               Visible         =   0   'False
               Width           =   975
            End
            Begin VB.TextBox TxtItemCode 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   7080
               TabIndex        =   74
               Top             =   750
               Width           =   1785
            End
            Begin VB.OptionButton Option1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄—÷ ﬂ«›Â «·«’‰«›"
               Height          =   375
               Left            =   10200
               RightToLeft     =   -1  'True
               TabIndex        =   48
               Top             =   750
               Width           =   1680
            End
            Begin VB.OptionButton Option2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«Œ Ì«— ’‰›"
               Height          =   375
               Left            =   9000
               RightToLeft     =   -1  'True
               TabIndex        =   47
               Top             =   750
               Value           =   -1  'True
               Width           =   1095
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   4335
               Index           =   1
               Left            =   0
               TabIndex        =   5
               TabStop         =   0   'False
               Top             =   1140
               Width           =   15345
               _cx             =   27067
               _cy             =   7646
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
               Begin VB.TextBox TxtItemsIDes 
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
                  Left            =   300
                  RightToLeft     =   -1  'True
                  TabIndex        =   80
                  Top             =   2895
                  Visible         =   0   'False
                  Width           =   1395
               End
               Begin VB.CheckBox ChKauto 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·Ì"
                  Enabled         =   0   'False
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
                  Left            =   12480
                  RightToLeft     =   -1  'True
                  TabIndex        =   29
                  Top             =   2430
                  Width           =   1500
               End
               Begin VB.TextBox txtType 
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
                  Height          =   255
                  Left            =   13680
                  RightToLeft     =   -1  'True
                  TabIndex        =   27
                  Text            =   "0"
                  Top             =   2790
                  Visible         =   0   'False
                  Width           =   495
               End
               Begin VB.CheckBox Check1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "⁄—÷ "
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
                  Left            =   12240
                  RightToLeft     =   -1  'True
                  TabIndex        =   18
                  Top             =   2190
                  Width           =   2235
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
                  Left            =   -3960
                  RightToLeft     =   -1  'True
                  TabIndex        =   8
                  Top             =   6810
                  Width           =   2190
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
                  Left            =   13680
                  RightToLeft     =   -1  'True
                  TabIndex        =   6
                  Top             =   2730
                  Visible         =   0   'False
                  Width           =   2175
               End
               Begin MSDataListLib.DataCombo dcopr 
                  Height          =   315
                  Left            =   13200
                  TabIndex        =   10
                  Top             =   1470
                  Width           =   4305
                  _ExtentX        =   7594
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
               Begin MSDataListLib.DataCombo dcproject 
                  Height          =   315
                  Left            =   13200
                  TabIndex        =   11
                  Top             =   1080
                  Width           =   1620
                  _ExtentX        =   2858
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
                  Left            =   13440
                  TabIndex        =   26
                  Top             =   720
                  Width           =   3300
                  _ExtentX        =   5821
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
               Begin VSFlex8Ctl.VSFlexGrid Grid 
                  Height          =   2370
                  Left            =   0
                  TabIndex        =   71
                  Top             =   180
                  Width           =   11940
                  _cx             =   21061
                  _cy             =   4180
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
                  Rows            =   1
                  Cols            =   21
                  FixedRows       =   1
                  FixedCols       =   2
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmCustomerrContract.frx":1F54
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
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   390
                  Index           =   10
                  Left            =   9960
                  TabIndex        =   89
                  Top             =   2640
                  Width           =   1515
                  _ExtentX        =   2672
                  _ExtentY        =   688
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   "Õ–› ﬂ· «·”ÿÊ—"
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
                  ButtonImage     =   "FrmCustomerrContract.frx":2266
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "»œ«Ì… «· Œ’Ì’"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Index           =   8
                  Left            =   13320
                  RightToLeft     =   -1  'True
                  TabIndex        =   9
                  Top             =   1860
                  Width           =   1800
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
                  Height          =   240
                  Left            =   13905
                  RightToLeft     =   -1  'True
                  TabIndex        =   7
                  Top             =   840
                  Width           =   870
               End
            End
            Begin ImpulseButton.ISButton Cmd 
               Height          =   390
               Index           =   20
               Left            =   1320
               TabIndex        =   49
               Top             =   750
               Width           =   960
               _ExtentX        =   1693
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
               ButtonImage     =   "FrmCustomerrContract.frx":2800
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton Cmd 
               Height          =   390
               Index           =   21
               Left            =   120
               TabIndex        =   50
               Top             =   750
               Width           =   915
               _ExtentX        =   1614
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
               ButtonImage     =   "FrmCustomerrContract.frx":2B9A
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin MSDataListLib.DataCombo dcitems 
               Height          =   315
               Left            =   2400
               TabIndex        =   51
               Top             =   750
               Width           =   4680
               _ExtentX        =   8255
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
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "»‰«¡« ⁄·Ï"
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
               Index           =   11
               Left            =   7740
               RightToLeft     =   -1  'True
               TabIndex        =   87
               Top             =   390
               Width           =   960
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
               Left            =   8430
               RightToLeft     =   -1  'True
               TabIndex        =   3
               Top             =   1380
               Width           =   1125
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   4380
            Index           =   0
            Left            =   12750
            TabIndex        =   52
            TabStop         =   0   'False
            Top             =   45
            Width           =   12015
            _cx             =   21193
            _cy             =   7726
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
            Begin VB.OptionButton Option4 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«Œ Ì«— „Ã„Ê⁄…"
               Height          =   375
               Left            =   7560
               RightToLeft     =   -1  'True
               TabIndex        =   54
               Top             =   120
               Value           =   -1  'True
               Width           =   1335
            End
            Begin VB.OptionButton Option3 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄—÷ ﬂ«›Â «·„Ã„Ê⁄« "
               Height          =   375
               Left            =   9480
               RightToLeft     =   -1  'True
               TabIndex        =   53
               Top             =   120
               Width           =   1800
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   4995
               Index           =   3
               Left            =   0
               TabIndex        =   55
               TabStop         =   0   'False
               Top             =   600
               Width           =   15345
               _cx             =   27067
               _cy             =   8811
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
               Begin VB.TextBox Text2 
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
                  Left            =   12480
                  RightToLeft     =   -1  'True
                  TabIndex        =   60
                  Top             =   3060
                  Visible         =   0   'False
                  Width           =   2175
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
                  Index           =   1
                  Left            =   -3960
                  RightToLeft     =   -1  'True
                  TabIndex        =   59
                  Top             =   7800
                  Width           =   2190
               End
               Begin VB.CheckBox Check3 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "⁄—÷ "
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
                  Left            =   12600
                  RightToLeft     =   -1  'True
                  TabIndex        =   58
                  Top             =   2880
                  Width           =   2235
               End
               Begin VB.TextBox Text1 
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
                  Height          =   255
                  Left            =   12360
                  RightToLeft     =   -1  'True
                  TabIndex        =   57
                  Text            =   "0"
                  Top             =   3360
                  Visible         =   0   'False
                  Width           =   495
               End
               Begin VB.CheckBox Check2 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·Ì"
                  Enabled         =   0   'False
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
                  Left            =   12600
                  RightToLeft     =   -1  'True
                  TabIndex        =   56
                  Top             =   2760
                  Width           =   1500
               End
               Begin MSDataListLib.DataCombo DataCombo1 
                  Height          =   315
                  Left            =   12960
                  TabIndex        =   61
                  Top             =   1560
                  Width           =   4305
                  _ExtentX        =   7594
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
               Begin MSDataListLib.DataCombo DataCombo2 
                  Height          =   315
                  Left            =   13080
                  TabIndex        =   62
                  Top             =   1320
                  Width           =   1620
                  _ExtentX        =   2858
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
               Begin MSDataListLib.DataCombo DataCombo3 
                  Height          =   315
                  Left            =   13080
                  TabIndex        =   63
                  Top             =   840
                  Width           =   3300
                  _ExtentX        =   5821
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
               Begin VSFlex8Ctl.VSFlexGrid Grid1 
                  Height          =   2820
                  Left            =   0
                  TabIndex        =   70
                  Top             =   480
                  Width           =   11940
                  _cx             =   21061
                  _cy             =   4974
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
                  Rows            =   1
                  Cols            =   21
                  FixedRows       =   1
                  FixedCols       =   2
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmCustomerrContract.frx":3134
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
               Begin VB.Label Label1 
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
                  Height          =   255
                  Left            =   13905
                  RightToLeft     =   -1  'True
                  TabIndex        =   65
                  Top             =   840
                  Width           =   870
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "»œ«Ì… «· Œ’Ì’"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Index           =   4
                  Left            =   13080
                  RightToLeft     =   -1  'True
                  TabIndex        =   64
                  Top             =   2190
                  Width           =   1800
               End
            End
            Begin ImpulseButton.ISButton Cmd 
               Height          =   390
               Index           =   7
               Left            =   1560
               TabIndex        =   66
               Top             =   120
               Width           =   960
               _ExtentX        =   1693
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
               ButtonImage     =   "FrmCustomerrContract.frx":3450
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton Cmd 
               Height          =   390
               Index           =   8
               Left            =   240
               TabIndex        =   67
               Top             =   120
               Width           =   1035
               _ExtentX        =   1826
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
               ButtonImage     =   "FrmCustomerrContract.frx":37EA
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin MSDataListLib.DataCombo DcGroup 
               Height          =   315
               Left            =   3120
               TabIndex        =   68
               Top             =   120
               Width           =   4080
               _ExtentX        =   7197
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
               Index           =   6
               Left            =   8400
               RightToLeft     =   -1  'True
               TabIndex        =   69
               Top             =   810
               Width           =   1125
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic EltCont 
         Height          =   1080
         Left            =   30
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   7305
         Width           =   12105
         _cx             =   21352
         _cy             =   1905
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
            Left            =   12240
            TabIndex        =   13
            TabStop         =   0   'False
            ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
            Top             =   0
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
            ButtonImage     =   "FrmCustomerrContract.frx":3D84
            ColorButton     =   14737632
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnUpdate 
            Height          =   330
            Left            =   12765
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
            ButtonImage     =   "FrmCustomerrContract.frx":411E
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnPrint 
            Height          =   285
            Left            =   13965
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
            ButtonImage     =   "FrmCustomerrContract.frx":44B8
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   375
            Index           =   0
            Left            =   10680
            TabIndex        =   19
            Top             =   480
            Width           =   1335
            _ExtentX        =   2355
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
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   375
            Index           =   1
            Left            =   9240
            TabIndex        =   20
            Top             =   480
            Width           =   1215
            _ExtentX        =   2143
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
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   375
            Index           =   2
            Left            =   7800
            TabIndex        =   21
            Top             =   480
            Width           =   1245
            _ExtentX        =   2196
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
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton Cmd 
            CausesValidation=   0   'False
            Height          =   375
            Index           =   3
            Left            =   6240
            TabIndex        =   22
            Top             =   480
            Width           =   1365
            _ExtentX        =   2408
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
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   375
            Index           =   4
            Left            =   4680
            TabIndex        =   23
            Top             =   480
            Width           =   1245
            _ExtentX        =   2196
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
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton Cmd 
            CausesValidation=   0   'False
            Height          =   375
            Index           =   6
            Left            =   120
            TabIndex        =   24
            Top             =   480
            Width           =   1005
            _ExtentX        =   1773
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
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   375
            Index           =   5
            Left            =   3240
            TabIndex        =   25
            Top             =   480
            Width           =   1365
            _ExtentX        =   2408
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
            ColorTextShadow =   4210752
         End
         Begin ALLButtonS.ALLButton CmdRemove 
            Height          =   375
            Left            =   11040
            TabIndex        =   28
            Tag             =   "Delete Row"
            Top             =   0
            Width           =   975
            _ExtentX        =   1720
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
            BCOL            =   12632319
            BCOLO           =   12632319
            FCOL            =   255
            FCOLO           =   255
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmCustomerrContract.frx":4852
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ImpulseButton.ISButton ISButton2 
            Height          =   405
            Left            =   2040
            TabIndex        =   73
            TabStop         =   0   'False
            ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
            Top             =   480
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   714
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "ÿ»«⁄… "
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
            ButtonImage     =   "FrmCustomerrContract.frx":486E
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   375
            Index           =   9
            Left            =   840
            TabIndex        =   75
            Top             =   480
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   661
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "‰”Œ… „„«À·Â"
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
            Height          =   210
            Left            =   1560
            RightToLeft     =   -1  'True
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
            RightToLeft     =   -1  'True
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
      ButtonImage     =   "FrmCustomerrContract.frx":B0D0
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
End
Attribute VB_Name = "FrmCustomerContract"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cSearchDCombo As clsDCboSearch
Dim BKGrndPic As ClsBackGroundPic
Dim net_value As Double
Dim isFromExcel As Boolean
Dim net_value1 As Double
Dim My_SQL  As String
Dim StrSQL  As String
Dim RecId As String
Dim rs As ADODB.Recordset
Dim EmpID As Integer
Private Declare Function TextOut _
                Lib "gdi32" _
                Alias "TextOutA" (ByVal hDC As Long, _
                                  ByVal X As Long, _
                                  ByVal Y As Long, _
                                  ByVal lpString As String, _
                                  ByVal nCount As Long) As Long



Private Sub check4_Click()
If Check4.value = vbUnchecked Then
    Exit Sub
Else
    If val(TxtPlanID.Text) <> 0 Then
        TxtPlanID_Validate False
    End If
End If
End Sub

Private Sub Command1_Click()
'        If DCboStoreName.BoundText = "" Then
'            Msg = "ÌÃ» «Œ Ì«— «”„ «·„Œ“‰"
'            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
'            DCboStoreName.SetFocus
'            SendKeys "{F4}"
'            Screen.MousePointer = vbDefault
'            Exit Sub
'        End If
        
CD1.ShowOpen
txtFile.Text = CD1.filename
End Sub

Private Sub Command2_Click()

  FillItem

End Sub


Sub FillItem()
Dim error_string  As String
  error_string = ""
If txtFile.Text = "" Then MsgBox "Õœœ «·„·› «Ê·«": Exit Sub
    Dim ExcelObj As Object
    Dim ExcelBook As Object
    Dim ExcelSheet As Object
    Dim i As Integer
    Dim currentvalue As String, mDesc As String
    Dim Name As String
    Dim itemcode As String
    Dim ITEMPRICE As Double
    Dim itemDisc As Double
    Dim UnitName As String
    Dim mEqu As String
    Dim des As String
    Dim DebitValue As String
    Dim CreditValue As String
   Grid.Rows = 1
    Set ExcelObj = CreateObject("Excel.Application")
'        Set ExcelSheet = Nothing
'    Set ExcelBook = Nothing
'    Set ExcelObj = Nothing
'
    Set ExcelSheet = CreateObject("Excel.Sheet")
    ExcelObj.Workbooks.Open txtFile.Text   ' App.Path & "\TrialBalance.xls"
DoEvents
    Set ExcelBook = ExcelObj.Workbooks(1)
    Set ExcelSheet = ExcelBook.Worksheets(1)
    isFromExcel = True
    With ExcelSheet
    i = 2
    
    Do Until .cells(i, 1) & "" = ""
        itemcode = .cells(i, 1)
        Name = .cells(i, 2)
        UnitName = .cells(i, 3)
        ITEMPRICE = val(.cells(i, 4) & "")
        itemDisc = val(.cells(i, 5) & "")
        
    'mDesc = .cells(i, 5)
 addrow2 itemcode, Name, UnitName, ITEMPRICE, itemDisc
       i = i + 1
     '  NewGrid.CountItems
    Loop
        End With
    ExcelObj.Workbooks.Close

    Set ExcelSheet = Nothing
    Set ExcelBook = Nothing
    Set ExcelObj = Nothing

        If error_string <> "" Then
            CreatLog_File_for_error (error_string)
       End If
       isFromExcel = False
       Me.Grid.Rows = Me.Grid.Rows + 1
'GetNotinGard
'Coloring
End Sub


Function addrow2(Fullcode As String, Name As String, UnitName As String, ITEMPRICE As Double, Optional itemDisc As Double, Optional des As String)

    Dim StrSQL As String
    Dim Rs2 As ADODB.Recordset
    Set Rs2 = New ADODB.Recordset
    Dim UnitID As Double
    Dim LngItemID As Long
    Dim LngUnitID As Long
    Dim ColorID As Integer
    Dim sizeid As Integer
    Dim ClassId As Integer
    Dim ParrtNoCode As String
    Dim ItemDetailedCode As String
 Dim error_string As String
    Dim Price As Double
    Dim i As Long
  '  UnitID = GetUnitID(Name)
    
                       
                'lllllllllllllll
                
                
                
  
  Dim s As String
  Dim rsDummy As New ADODB.Recordset
  Dim RsUnit As New ADODB.Recordset
   If Name <> "" Then
       s = "Select * from tblItems Where Fullcode Like '" & Trim(Fullcode) & "' Or barCodeNO Like '" & Trim(Fullcode) & "'"
       rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly
        If rsDummy.EOF Then
            Exit Function
        Else
            LngItemID = val(rsDummy!ItemID & "")
        End If
        
    If LngItemID <> 0 Then
    Dim mRow As Long
    
    With Me.Grid
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, .ColIndex("ItemId")) = LngItemID
        
            
        .TextMatrix(.Rows - 1, .ColIndex("ItemCode")) = rsDummy!itemcode & ""
        .TextMatrix(.Rows - 1, .ColIndex("ItemName")) = IIf(IsNull(rsDummy.Fields("ItemName").value), "", rsDummy.Fields("ItemName").value)
        
        
            s = "Select UnitID,UnitName From TblUnites Where UnitName Like '" & Trim(UnitName) & "'"
            Set RsUnit = New ADODB.Recordset
            RsUnit.Open s, Cn, adOpenStatic
            If Not RsUnit.EOF Then
                LngUnitID = val(RsUnit!UnitID & "")
            Else
                StrSQL = "SELECT TblItemsUnits.UnitID, TblUnites.UnitName,TblItemsUnits.UnitWholeSalePrice "
                StrSQL = StrSQL + " FROM TblUnites INNER JOIN TblItemsUnits " & "ON TblUnites.UnitID = TblItemsUnits.UnitID "
                StrSQL = StrSQL + " Where TblItemsUnits.DefaultUnit=1 and  TblItemsUnits.ItemID=" & LngItemID
                StrSQL = StrSQL + " Order BY TblItemsUnits.SecOrder "
                Set RsUnit = New ADODB.Recordset
                RsUnit.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
            End If
               

            If Not RsUnit.EOF Then
                
                .TextMatrix(.Rows - 1, .ColIndex("UnitId")) = IIf(IsNull(RsUnit.Fields("UnitId").value), "", RsUnit.Fields("UnitId").value)
                .TextMatrix(.Rows - 1, .ColIndex("UnitName")) = IIf(IsNull(RsUnit.Fields("UnitName").value), "", RsUnit.Fields("UnitName").value)
                .TextMatrix(.Rows - 1, .ColIndex("Price")) = IIf(IsNull(RsUnit.Fields("UnitWholeSalePrice").value), "", RsUnit.Fields("UnitWholeSalePrice").value)
            End If

            RsUnit.Close
        
        
        .TextMatrix(.Rows - 1, .ColIndex("Price")) = ITEMPRICE
        .TextMatrix(.Rows - 1, .ColIndex("Discount")) = itemDisc
        
        
        .Row = .Rows - 1
  
       
        
        

      

     End With
    '      Me.TxtItemCodeB.Text = ""
     
    '\      Unload FrmItemSearch2
     ' Me.TxtItemCodeB.SetFocus
         
    Else
         
    End If
    
    Else
           error_string = error_string & Trim(Fullcode) & "," & ITEMPRICE & "," & Name & vbCrLf

End If
'End If

End Function
Private Sub lbl_MouseMove(Index As Integer, _
                          Button As Integer, _
                          Shift As Integer, _
                          X As Single, _
                          Y As Single)

    If val(lbl(Index).Caption) <> 0 Then
        lbl(Index).ToolTipText = WriteNo(lbl(Index).Caption, 0, True)
    End If

End Sub

Private Sub DcboEmp_Change()
Dim StoreId As Integer
 'If Me.TxtModFlg = "N" Or Me.TxtModFlg = "E" Then
         If val(Me.DcboEmp.BoundText) = 0 Then Exit Sub
           Me.TxtEmployeeID.Text = get_EMPLOYEE_Data(val(Me.DcboEmp.BoundText), "Fullcode")
        'DCEmP.text = DCEmP.text
'End If
 
        

End Sub

Private Sub DcboEmp_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF5 Then
'ReloadCombos
End If
End Sub

Private Sub TxtEmployeeID_Change()

'If Me.TxtModFlg = "N" Or Me.TxtModFlg = "E" Then
'    DcboEmp.BoundText = GeTEmpIDByEmpCode(TxtEmployeeID.text, True)
'End If

End Sub

Private Sub TxtEmployeeID_KeyDown(KeyCode As Integer, Shift As Integer)
   Dim EmpID As Integer

    If KeyCode = vbKeyReturn Then
        GetEmployeeIDFromCode TxtEmployeeID.Text, EmpID
        DcboEmp.BoundText = EmpID
    End If

End Sub

Function CuurentLogdata(Optional Currentmode As String)
   
    LogTextA = "    ‘«‘… " & ScreenNameArabic & CHR(13) & " —ﬁ„ «·« ›«ﬁÌ…    " & TxtTblCustomerContractD.Text & CHR(13) & " «·⁄„»· " & DBCboClientName.Text & CHR(13) & "  „œ Â« „‰  " & dbFromDate & CHR(13) & "  «·Ï " & dbTodate & CHR(13) & "  „·«ÕŸ«  " & txtRemarks

    If ChkLocked.value = Checked Then
        LogTextA = LogTextA & CHR(13) & "   „ «Ìﬁ«› «· ⁄«„· "
    End If
                    
    LogTexte = "    Screen  " & ScreenNameEnglish & CHR(13) & " Contract No    " & TxtTblCustomerContractD.Text & CHR(13) & " Customer " & DBCboClientName.Text & CHR(13) & " From   " & dbFromDate & CHR(13) & "  To  " & dbTodate & CHR(13) & "  Remarks " & txtRemarks

    If ChkLocked.value = Checked Then
        LogTextA = LogTextA & CHR(13) & " Locked "
    End If
                    
    If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, Me.TxtModFlg
    Else
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, "D"
    End If
    
End Function

Private Sub ChkDetails_Click()
    FillGridWithData
End Sub

Private Sub ALLButton1_Click()
    FrmShowCol1.show
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
    Dim i As Long
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
    Dim i As Long
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

Private Sub ALLButton2_Click()
    'Dcemp.text = ""

    dcproject.Text = ""
    FillGridWithData

    DoEvents
    Create_dev
    CmdOk_Click
End Sub

Private Sub ALLButton3_Click()
 
End Sub

Private Sub CboPayMentType_Change()
 
End Sub

Private Sub CboPayMentType_Click()
    CboPayMentType_Change
End Sub

Private Sub CboYear_Click()
    CmdOk_Click
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

Private Sub Combo1_Click()
 
End Sub

Private Sub Del_Trans()
    On Error GoTo ErrTrap
    Dim Msg  As String

    If TxtTblCustomerContractD.Text <> "" Then
        Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & CHR(13)
        Msg = Msg + (TxtTblCustomerContractD.Text) & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            If Not rs.RecordCount < 1 Then
                Cn.Execute "delete TblCustomerContractDetails where TblCustomerContractD=" & val(Me.TxtTblCustomerContractD.Text)
                CuurentLogdata ("D")
                rs.delete
                rs.MoveFirst

                If rs.RecordCount < 1 Then
                    clear_all Me
                    TxtModFlg_Change
                    '   XPTxtCurrent.Caption = 0
                    '   XPTxtCount.Caption = 0
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

Private Sub SaveData()
    Dim Msg As String
    Dim RsTemp As New ADODB.Recordset
    Dim StrSQL As String
    Dim BeginTrans As Boolean
    Dim RsDev As ADODB.Recordset
    Dim LngDevID As Long

    On Error GoTo ErrTrap

    If Me.TxtModFlg.Text <> "R" Then
 
        If Trim(Me.DBCboClientName.BoundText) = "" Then
            Msg = "ÌÃ» ≈Œ Ì«— «·⁄„Ì·..!!"
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            DBCboClientName.SetFocus
            SendKeys "{F4}"
            Exit Sub
        End If
 
    End If

    '-------------------------------------------------------------------------------------------
   
    Cn.BeginTrans
    BeginTrans = True

    If TxtModFlg.Text = "N" Then
        rs.AddNew
    ElseIf Me.TxtModFlg.Text = "E" Then
        Cn.Execute "delete TblCustomerContractDetails where TblCustomerContractD=" & val(Me.TxtTblCustomerContractD.Text)
   
    End If
    
    rs("TblCustomerContractD").value = TxtTblCustomerContractD.Text
    rs("CustomerId").value = IIf(Me.DBCboClientName.BoundText = "", Null, Me.DBCboClientName.BoundText)
    rs("Emp_ID").value = IIf(DcboEmp.BoundText = "", Null, DcboEmp.BoundText)
   ' rs("TxtCustumCode").value = IIf(TxtSearchCode.text <> "", Trim(TxtSearchCode.text), Null)


    If Me.DBCboClientName.BoundText <> 2 Then
    TxtCashCustomerName.Text = DBCboClientName.Text
   
    
    End If
    
    
    
    If Trim$(Me.TxtCashCustomerName.Text) <> "" Then
        rs("CashCustomerName").value = Trim$(Me.TxtCashCustomerName.Text)
    Else
        rs("CashCustomerName").value = Null
    End If
   
    rs("FromDate").value = dbFromDate.value
    rs("Todate").value = dbTodate.value
      
    rs("Remarks").value = IIf(Me.txtRemarks.Text = "", "", Me.txtRemarks.Text)
     rs("PlanID").value = IIf(Me.TxtPlanID.Text = "", "", Me.TxtPlanID.Text)
    
    If ChkLocked.value = vbChecked Then
        rs("Locked").value = 1
    Else
        rs("Locked").value = 0
    End If

    
     If CBoBasedON.ListIndex = -1 Then
        rs("CBoBasedON").value = 0
    Else
        rs("CBoBasedON").value = val(CBoBasedON.ListIndex)
    End If
    

    rs.update
    
    Set RsDev = New ADODB.Recordset
        
    RsDev.Open "TblCustomerContractDetails", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        
    Dim i As Long

    With Me.Grid

        For i = 1 To .Rows - 1

            If .TextMatrix(i, .ColIndex("ItemId")) <> "" Then
         
                RsDev.AddNew
                RsDev("TblCustomerContractD").value = Me.TxtTblCustomerContractD.Text
            
                RsDev("ItemId").value = val(.TextMatrix(i, .ColIndex("ItemId")))
                RsDev("UnitID").value = val(.TextMatrix(i, .ColIndex("UnitID")))
                RsDev("Price").value = val(.TextMatrix(i, .ColIndex("Price")))
                RsDev("Discount").value = val(.TextMatrix(i, .ColIndex("Discount")))
                RsDev.update
                    
            End If
            
            '
        Next i

    End With
 
    RsDev.Close
    'save Groups
    GoTo NotsaveGroups
    Set RsDev = New ADODB.Recordset
        
    RsDev.Open "TblCustomerContractDetailsGroups", Cn, adOpenStatic, adLockOptimistic, adCmdTable
 
    With Me.Grid1

        For i = 1 To .Rows - 1

            If .TextMatrix(i, .ColIndex("GroupID")) <> "" Then
         
                RsDev.AddNew
                RsDev("TblCustomerContractD").value = Me.TxtTblCustomerContractD.Text
            
                RsDev("GroupID").value = val(.TextMatrix(i, .ColIndex("GroupID")))
                RsDev("UnitID").value = val(.TextMatrix(i, .ColIndex("UnitID")))
                RsDev("Price").value = val(.TextMatrix(i, .ColIndex("Price")))
                RsDev("Discount").value = val(.TextMatrix(i, .ColIndex("Discount")))
                RsDev.update
                    
            End If
            
            '
        Next i

    End With
NotsaveGroups:
    Cn.CommitTrans
    BeginTrans = False
    CuurentLogdata

    Select Case Me.TxtModFlg.Text

        Case "N"
            Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & CHR(13)
            Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"

            '    Grid_Journal.Enabled = False
            If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                Cmd_Click (0)
                Exit Sub
            End If

        Case "E"
            MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            '  Grid_Journal.Enabled = False
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

Private Sub Cmd_Click(Index As Integer)
    On Error GoTo ErrTrap

    Select Case Index
    Case 10
    
            Grid.Clear flexClearScrollable, flexClearEverything
            Grid.Rows = 1
            
Case 9
  TxtModFlg.Text = "N"
      Me.TxtTblCustomerContractD.Text = CStr(new_id("TblCustomerContract", "TblCustomerContractD", "", True))
       
        Case 0

            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.Text = "N"
            clear_all Me
            Me.TxtTblCustomerContractD.Text = CStr(new_id("TblCustomerContract", "TblCustomerContractD", "", True))
       
            Me.dbFromDate.value = Date
            Me.dbTodate.value = Date
       
            'XPDtbTrans.SetFocus
            Grid.Clear flexClearScrollable, flexClearEverything
            Grid.Rows = 1
             Grid1.Clear flexClearScrollable, flexClearEverything
            Grid1.Rows = 1
            Grid.Enabled = True
            Option2.value = True
            Option4.value = True
            
        Case 1

            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.Text = "E"
            '         Grid.Rows = Grid.Rows + 1
            Grid.Enabled = True
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

        Case 5

            If DoPremis(Do_Search, Me.Name, True) = False Then
                Exit Sub
            End If
'wael
    '          Load FrmCustomerContractSearch
    '          FrmCustomerContractSearch.show vbModal

        Case 6
            Unload Me

        Case 7
            '   ViewDataList
            addrowGroups
    
        Case 8
            RemoveGridRowGroup
    
        Case 20
            addrow

        Case 21
            RemoveGridRow
    End Select

    Exit Sub
ErrTrap:

End Sub

Private Sub RemoveGridRowGroup()

    With Me.Grid1

        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With

    ReLineGrid
End Sub

Private Sub RemoveGridRow()

    With Me.Grid

        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With

    ReLineGrid
End Sub

Function addrow(Optional ByVal lastrow2 As Long = 0)
    If Check4.value = vbChecked And val(TxtPlanID) <> 0 Then
        TxtPlanID_Validate False
        Exit Function
    End If
    Dim wherestr As String

    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim Rs2 As ADODB.Recordset
    Set Rs2 = New ADODB.Recordset
    Dim RsUnit As ADODB.Recordset
    Set RsUnit = New ADODB.Recordset

    Dim j As Integer

    Dim sql As String
    Dim i As Long
    Dim Msg  As String
    Dim lastrow As Integer
    Dim LngItemID As Integer
Dim mUnitId As Long
    If Option2.value = True Then
        If dcitems.BoundText = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "ÌÃ»       «Œ Ì«— «·’‰›  ...!!!"
            Else
                Msg = "must Specify item Name ...!!!"
            End If

            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Exit Function
        End If

        wherestr = "  TblItems.ItemID= " & val(dcitems.BoundText)
    End If

    'sql = "Select * from TblItems "
    sql = "SELECT  TblItems.* ,TblItemsUnits.UnitID, TblUnites.UnitName,TblItemsUnits.UnitWholeSalePrice "
    sql = sql + " FROM TblUnites INNER JOIN TblItemsUnits " & "ON TblUnites.UnitID = TblItemsUnits.UnitID "
    sql = sql + " Inner join TblItems On TblItems.ItemID = TblItemsUnits.ItemID"
                
                
                If wherestr <> "" Then
                    sql = sql + " Where 1 = 1"
                    sql = sql + " and " & wherestr
                End If
                
                sql = sql + " Order BY TblItemsUnits.SecOrder "
    
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then Exit Function

    With Grid
        If lastrow2 <> 0 Then lastrow = lastrow2 Else lastrow = .Rows
        
    
        If Rs3.RecordCount > 0 Then
             If lastrow2 = 0 Then
                .Rows = Rs3.RecordCount + lastrow
            End If
            Rs3.MoveFirst
         
            For i = lastrow To Rs3.RecordCount + lastrow - 1
                .TextMatrix(i, .ColIndex("ItemId")) = IIf(IsNull(Rs3.Fields("ItemId").value), "", Rs3.Fields("ItemId").value)
                LngItemID = IIf(IsNull(Rs3.Fields("ItemId").value), "", Rs3.Fields("ItemId").value)
                  mUnitId = IIf(IsNull(Rs3.Fields("UnitId").value), "", Rs3.Fields("UnitId").value)
                .TextMatrix(i, .ColIndex("ItemCode")) = IIf(IsNull(Rs3.Fields("ItemCode").value), "", Rs3.Fields("ItemCode").value)
                .TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(Rs3.Fields("ItemName").value), "", Rs3.Fields("ItemName").value)
                       
                        .TextMatrix(i, .ColIndex("UnitId")) = IIf(IsNull(Rs3.Fields("UnitId").value), "", Rs3.Fields("UnitId").value)
                    .TextMatrix(i, .ColIndex("UnitName")) = IIf(IsNull(Rs3.Fields("UnitName").value), "", Rs3.Fields("UnitName").value)
                    .TextMatrix(i, .ColIndex("Price")) = IIf(IsNull(Rs3.Fields("UnitWholeSalePrice").value), "", Rs3.Fields("UnitWholeSalePrice").value)
                     
                'lllllllllllllll
    
                       
                Rs3.MoveNext
            Next i
 
            '    .AutoSize 0, .Cols - 1, False
        End If

    End With
 
    Rs3.Close

    ReLineGrid

End Function

Function addrowGroups()

    Dim wherestr As String

    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim Rs2 As ADODB.Recordset
    Set Rs2 = New ADODB.Recordset
    Dim RsUnit As ADODB.Recordset
    Set RsUnit = New ADODB.Recordset

    Dim j As Integer

    Dim sql As String
    Dim i As Long
    Dim Msg  As String
    Dim lastrow As Integer
    Dim LngItemID As Integer

    If Option4.value = True Then
        If DcGroup.BoundText = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "ÌÃ»       «Œ Ì«— „Ã„Ê⁄Â  ...!!!"
            Else
                Msg = "must Specify item Name ...!!!"
            End If

            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Exit Function
        End If

        wherestr = "  where GroupID= " & val(DcGroup.BoundText)
    End If

    sql = "Select * from Groups "

    If wherestr <> "" Then
        sql = sql & wherestr
    End If
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then Exit Function

    With Grid1
 
        lastrow = .Rows
    
        If Rs3.RecordCount > 0 Then
            .Rows = Rs3.RecordCount + lastrow
            Rs3.MoveFirst
         
            For i = lastrow To Rs3.RecordCount + lastrow - 1
                .TextMatrix(i, .ColIndex("GroupID")) = IIf(IsNull(Rs3.Fields("GroupID").value), "", Rs3.Fields("GroupID").value)
                LngItemID = IIf(IsNull(Rs3.Fields("GroupID").value), "", Rs3.Fields("GroupID").value)
                       
                .TextMatrix(i, .ColIndex("Fullcode")) = IIf(IsNull(Rs3.Fields("Fullcode").value), "", Rs3.Fields("Fullcode").value)
                .TextMatrix(i, .ColIndex("GroupName")) = IIf(IsNull(Rs3.Fields("GroupName").value), "", Rs3.Fields("GroupName").value)
                       
                'lllllllllllllll
                '     StrSQL = "SELECT TblItemsUnits.UnitID, TblUnites.UnitName "
                '   StrSQL = StrSQL + " FROM TblUnites INNER JOIN TblItemsUnits " & _
                '   "ON TblUnites.UnitID = TblItemsUnits.UnitID "
                '   StrSQL = StrSQL + " Where TblItemsUnits.DefaultUnit=1 and  TblItemsUnits.ItemID=" & LngItemID
                '   StrSQL = StrSQL + " Order BY TblItemsUnits.SecOrder "
                 
                StrSQL = "SELECT TblUnites.UnitID, TblUnites.UnitName "
                StrSQL = StrSQL + " FROM TblUnites  "
                
                RsUnit.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                If RsUnit.RecordCount > 0 Then
                    RsUnit.MoveFirst
                    .TextMatrix(i, .ColIndex("UnitId")) = IIf(IsNull(RsUnit.Fields("UnitId").value), "", RsUnit.Fields("UnitId").value)
                    .TextMatrix(i, .ColIndex("UnitName")) = IIf(IsNull(RsUnit.Fields("UnitName").value), "", RsUnit.Fields("UnitName").value)
                     '.TextMatrix(i, .ColIndex("Price")) = IIf(IsNull(RsUnit.Fields("UnitWholeSalePrice").value), "", RsUnit.Fields("UnitWholeSalePrice").value)
               
                End If

                RsUnit.Close
                       
                Rs3.MoveNext
            Next i
 
            '    .AutoSize 0, .Cols - 1, False
        End If

    End With
 
    Rs3.Close

    ReLineGrid

End Function

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

Private Sub DBCboClientName_Change()
  On Error Resume Next
  If Me.TxtModFlg = "R" Then Exit Sub
    If val(DBCboClientName.BoundText) = 0 Then Exit Sub
    Dim Fullcode  As String, DefaultSalesPersonId As Integer
    GetCustomersDetail val(DBCboClientName.BoundText), DefaultSalesPersonId, Fullcode, 1
    TxtSearchCode.Text = Fullcode
    If Not DefaultSalesPersonId = 0 Then

        Me.DcboEmp.BoundText = DefaultSalesPersonId
    End If
    If DefaultSalesPersonId = 0 Then
       Me.DcboEmp.BoundText = EmpID
    End If
    
    If Me.DBCboClientName.BoundText <> 2 Then
    TxtCashCustomerName.Text = DBCboClientName.Text
    TxtCashCustomerName.Visible = False
    lbl(10).Visible = False
    Else
    TxtCashCustomerName.Visible = True
        lbl(10).Visible = True
    End If
    
End Sub

Private Sub DBCboClientName_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        FrmCustemerSearch.SearchType = 20915
        FrmCustemerSearch.show vbModal

    End If
End Sub

Private Sub dcitems_KeyUp(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyF3 Then
        Load FrmItemSearch
        FrmItemSearch.RetrunType = 21915
        FrmItemSearch.show vbModal
    End If
End Sub

Private Sub dcproject_Click(Area As Integer)

    If dcproject.BoundText = "" Then Exit Sub
    My_SQL = " select  fullcode,des from projects_des where project_id=" & val(dcproject.BoundText)
    fill_combo Dcterm, My_SQL

End Sub

Private Sub Dcterm_Click(Area As Integer)

    If Dcterm.BoundText = "" Then Exit Sub

    My_SQL = " select  fullcode,name from terms_operations where term_fullcode='" & Dcterm.BoundText & "'"
    fill_combo dcopr, My_SQL
End Sub

Private Sub Form_Load()
 
    Me.Left = (mdifrmmain.Width - Me.Width) / 2
    Me.Top = (mdifrmmain.Height - Me.Height) / 2 - 500

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

    Dim GrdBack As ClsBackGroundPic
    Set GrdBack = New ClsBackGroundPic

    With Me.Grid
        Set .WallPaper = GrdBack.Picture
     
    End With

    With Me.Grid1
        Set .WallPaper = GrdBack.Picture
     
    End With

    'My_SQL = " select id,Project_name from projects"
    'fill_combo dcproject, My_SQL
    '
    'My_SQL = " select  fullcode,des from projects_des"
    'fill_combo Dcterm, My_SQL

    'My_SQL = " select  fullcode,name from terms_operations"
    'fill_combo dcopr, My_SQL

    Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos
    Set cSearchDCombo = New clsDCboSearch
 
    Set BKGrndPic = New ClsBackGroundPic

    Dcombos.GetCustomersSuppliers 1, Me.DBCboClientName, True
    Dcombos.GetItemsNames dcitems
    Dcombos.GetItemSGroups DcGroup
    Dcombos.GetSalesRepData Me.DcboEmp
    
    With Me.Grid
        .Rows = 1
        .ExplorerBar = flexExSortShowAndMove
        .RowHeightMin = 300
        .ExtendLastCol = True
    End With
      
      
      

    With Me.CBoBasedON
        .Clear
        '.AddItem "»·«"
        .AddItem "»·« "
        .AddItem "»‰«¡« ⁄·Ï Œÿ…  ”⁄Ì— ”⁄— «·Ã„·…"
        .AddItem "»‰«¡« ⁄·Ï Œÿ…  ”⁄Ì— ”⁄— «·»Ì⁄ «›—«œ"


    End With
      
      
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    Set rs = New ADODB.Recordset
    StrSQL = "select * From TblCustomerContract  "
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    XPBtnMove_Click 2
    Me.TxtModFlg.Text = "R"

    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If

End Sub

Private Sub ChangeLang()
    ChKauto.Caption = "Auto"
    Cmd(0).Caption = "New"
    Cmd(1).Caption = "Edit"
    Cmd(2).Caption = "Save"
    Cmd(3).Caption = "Undo"
    Cmd(4).Caption = "Delete"
    Cmd(5).Caption = "Search"
    ISButton2.Caption = "Print"
    Cmd(6).Caption = "Exit"
    'CmdHelp.Caption = "Help"

    Dim XPic As IPictureDisp
    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic
    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic

    Me.Caption = "Customer Contarcts"
    Ele(5).Caption = Me.Caption
    lbl(7).Caption = "ID"
    lbl(5).Caption = "Start Date"
    lbl(2).Caption = "End Date"
    lbl(0).Caption = "Customer"
    lbl(3).Caption = "Remarks"
  
    ChkLocked.Caption = "Locked"
    Option3.Caption = "All Groups"
    Option4.Caption = "Select Group"
    Cmd(7).Caption = "Add"
    Cmd(8).Caption = "Remove"

    Option1.Caption = "All Items"
    Option2.Caption = "Select Item"
    Cmd(20).Caption = "Add"
    Cmd(21).Caption = "Remove"

    CmdRemove.Caption = "Remove Line"
 
    With Me.Grid1
        .TextMatrix(0, .ColIndex("ser")) = "I"
        .TextMatrix(0, .ColIndex("fullcode")) = "Code"
        .TextMatrix(0, .ColIndex("GroupName")) = "Group Name"
        .TextMatrix(0, .ColIndex("UnitName")) = "UnitName"
        .TextMatrix(0, .ColIndex("Price")) = "Price"
        .TextMatrix(0, .ColIndex("discount")) = "Discount"
    End With

    With Me.Grid
        .TextMatrix(0, .ColIndex("ser")) = "I"
        .TextMatrix(0, .ColIndex("ItemCode")) = "ItemCode"
        .TextMatrix(0, .ColIndex("ItemName")) = "Item Name"
        .TextMatrix(0, .ColIndex("UnitName")) = "UnitName"
        .TextMatrix(0, .ColIndex("Price")) = "Price"
        .TextMatrix(0, .ColIndex("discount")) = "Discount"
    End With

    Me.C1Tab1.TabCaption(1) = "Groups"
    Me.C1Tab1.TabCaption(0) = "Items"
End Sub

Public Sub get_all_employee()
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim Rs2 As ADODB.Recordset
    Set Rs2 = New ADODB.Recordset
    Dim j As Integer

    Dim sql As String
    Dim i As Long

    sql = "Select * from emp_all_details "
 
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
                .TextMatrix(i, .ColIndex("DepartmentName")) = IIf(IsNull(Rs3.Fields("DepartmentName").value), "", Rs3.Fields("DepartmentName").value)
                .TextMatrix(i, .ColIndex("JobTypeName")) = IIf(IsNull(Rs3.Fields("JobTypeName").value), "", Rs3.Fields("JobTypeName").value)
                .TextMatrix(i, .ColIndex("work_status")) = IIf(IsNull(Rs3.Fields("name").value), "", Rs3.Fields("name").value)
                       
                Rs3.MoveNext
            Next i
 
            .AutoSize 0, .Cols - 1, False
        End If

    End With
 
    Rs3.Close

End Sub

Public Sub FillGridWithData()
    Exit Sub

    Dim i As Long
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

Private Sub Grid_AfterEdit(ByVal Row As Long, _
                           ByVal Col As Long)
    On Error Resume Next
    Dim StrAccountCode As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long
    Dim code  As String

    With Grid

        Select Case .ColKey(Col)
 
            Case "UnitName"
                code = .ComboData
           
                '   LngRow = .FindRow(Code, .FixedRows, .ColIndex("UnitID"), False, True)
                .TextMatrix(Row, .ColIndex("UnitID")) = code
                .TextMatrix(Row, .ColIndex("UnitName")) = .ComboItem
 
        End Select
   
        If Row = .Rows - 1 Then
    
            '.Rows = .Rows + 1
        End If

        ReLineGrid
    End With
 
    If Me.TxtModFlg <> "E" Then Exit Sub

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\LogFile Saving//////////////////////////////////////////////////////////////////////////
    'Grid.TextMatrix(Row, Grid.ColIndex("Code"))
    'Grid.TextMatrix(Row, Grid.ColIndex("Name"))
    If Col = Grid.ColIndex("ItemCode") Or Col = Grid.ColIndex("ItemName") Then
        RegisterItemData Me.Name, Me.TxtModFlg, Grid.TextMatrix(Row, Grid.ColIndex("ItemCode")), Grid.TextMatrix(Row, Grid.ColIndex("ItemName")), , , , , , , , , , , Me.TxtTblCustomerContractD
    ElseIf Col = Grid.ColIndex("UnitName") Then
        RegisterItemData Me.Name, Me.TxtModFlg, Grid.TextMatrix(Row, Grid.ColIndex("ItemCode")), Grid.TextMatrix(Row, Grid.ColIndex("ItemName")), Grid.TextMatrix(Row, Grid.ColIndex("UnitName")), , , , , , , , , , Me.TxtTblCustomerContractD
    ElseIf Col = Grid.ColIndex("Price") Then
        RegisterItemData Me.Name, Me.TxtModFlg, Grid.TextMatrix(Row, Grid.ColIndex("ItemCode")), Grid.TextMatrix(Row, Grid.ColIndex("ItemName")), , , (Grid.TextMatrix(Row, Grid.ColIndex("Price"))), , , , , , , , Me.TxtTblCustomerContractD
    ElseIf Col = Grid.ColIndex("Discount") Then
        RegisterItemData Me.Name, Me.TxtModFlg, Grid.TextMatrix(Row, Grid.ColIndex("ItemCode")), Grid.TextMatrix(Row, Grid.ColIndex("ItemName")), , , , , , , , , Grid.TextMatrix(Row, Grid.ColIndex("Discount")), , Me.TxtTblCustomerContractD

    End If

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\LogFile Saving//////////////////////////////////////////////////////////////////////////

End Sub

Private Sub ReLineGrid()
    Dim IntCounter As Integer
    IntCounter = 0
    Dim i As Long

    With Me.Grid

        For i = .FixedRows To .Rows - 1
    
            If .TextMatrix(i, .ColIndex("ItemId")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("Ser")) = IntCounter
  
            End If

        Next i
   
    End With

    With Me.Grid1

        For i = .FixedRows To .Rows - 1
    
            If .TextMatrix(i, .ColIndex("GroupID")) <> "" Then
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

        If .ColKey(Col) <> "UnitName" Then
       
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
    Dim LngItemID As Integer
    Dim MyStrList As String

    With Me.Grid

        Select Case .ColKey(Col)

            Case "UnitName"

                LngItemID = val(.TextMatrix(.Row, .ColIndex("ItemId")))

                'LngItemID = 1
                If LngItemID = 0 Then
                    Cancel = True
                Else
            
                    StrSQL = "SELECT TblItemsUnits.UnitID, TblUnites.UnitName "
                    StrSQL = StrSQL + " FROM TblUnites INNER JOIN TblItemsUnits " & "ON TblUnites.UnitID = TblItemsUnits.UnitID "
                    StrSQL = StrSQL + " Where TblItemsUnits.ItemID=" & LngItemID
                    StrSQL = StrSQL + " Order BY TblItemsUnits.SecOrder "
                    Set rs = New ADODB.Recordset
                    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                    If Not (rs.BOF Or rs.EOF) Then
                        MyStrList = .BuildComboList(rs, "UnitName", "UnitID")
                        '                    Grid.ColComboList = MyStrList
                        Grid.ColComboList(.ColIndex("UnitName")) = "|" & MyStrList
                    Else
                        Cancel = True
                    End If
                End If
            
        End Select

    End With

End Sub

Public Sub Retrive(Optional Lngid As Long = 0)
    Dim RsDev As ADODB.Recordset
    Dim StrSQL As String
    Dim i As Long

 

    Check4.value = vbUnchecked
    
    'On Error GoTo ErrTrap
    Grid.Clear flexClearScrollable, flexClearEverything
    Grid.Rows = 1
          
    Grid1.Clear flexClearScrollable, flexClearEverything
    Grid1.Rows = 1

    If rs.RecordCount < 1 Then
        Exit Sub
    End If

    If rs.EOF Or rs.BOF Then
        Exit Sub
    Else

    End If
 
    Me.TxtTblCustomerContractD.Text = IIf(IsNull(rs("TblCustomerContractD").value), "", rs("TblCustomerContractD").value)
 
    dbFromDate.value = IIf(IsNull(rs("FromDate").value), Date, rs("FromDate").value)
    dbTodate.value = IIf(IsNull(rs("Todate").value), Date, rs("Todate").value)

    DBCboClientName.BoundText = IIf(IsNull(rs("CustomerId").value), "", rs("CustomerId").value)
    Me.DcboEmp.BoundText = IIf(IsNull(rs("Emp_ID").value), "", rs("Emp_ID").value)
    
    Me.TxtPlanID.Text = IIf(IsNull(rs("PlanID").value), "", rs("PlanID").value)
    CBoBasedON.ListIndex = IIf(IsNull(rs("CBoBasedON").value), 0, (rs("CBoBasedON").value))
    If Not (IsNull(rs("CashCustomerName").value)) Then
        Me.TxtCashCustomerName.Text = rs("CashCustomerName").value
    Else
        Me.TxtCashCustomerName.Text = ""
    End If
    
   ' TxtSearchCode.text = IIf(IsNull(rs("TxtCustumCode").value), "", rs("TxtCustumCode").value)
    txtRemarks.Text = IIf(IsNull(rs("Remarks").value), 0, rs("Remarks").value)

    If (rs("Locked").value) = True Then
     ChkLocked.value = vbChecked
     Else
     ChkLocked.value = vbUnchecked
     End If

    StrSQL = " SELECT     dbo.TblCustomerContractDetails.TblCustomerContractD, dbo.TblCustomerContractDetails.UnitID, dbo.TblCustomerContractDetails.ItemID, dbo.TblCustomerContractDetails.Discount, "
    StrSQL = StrSQL & "     dbo.TblCustomerContractDetails.Price , dbo.TblUnites.unitname, dbo.TblItems.ItemName, dbo.TblItems.ItemCode"
    StrSQL = StrSQL & " FROM         dbo.TblCustomerContractDetails INNER JOIN"
    StrSQL = StrSQL & " dbo.TblItems ON dbo.TblCustomerContractDetails.ItemID = dbo.TblItems.ItemID LEFT OUTER JOIN"
    StrSQL = StrSQL & " dbo.TblUnites ON dbo.TblCustomerContractDetails.UnitID = dbo.TblUnites.UnitID"
    StrSQL = StrSQL & "  where TblCustomerContractD=" & val(Me.TxtTblCustomerContractD.Text)
    
    Set RsDev = New ADODB.Recordset
    RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsDev.BOF Or rs.EOF) Then
        RsDev.MoveFirst
    
        With Me.Grid
    
            .Rows = .FixedRows + RsDev.RecordCount

            For i = .FixedRows To .Rows - 1
 
                .TextMatrix(i, .ColIndex("ItemId")) = IIf(IsNull(RsDev("ItemId").value), "", RsDev("ItemId").value)
            
                .TextMatrix(i, .ColIndex("ItemCode")) = IIf(IsNull(RsDev("ItemCode").value), "", RsDev("ItemCode").value)
            
                .TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(RsDev("ItemName").value), "", RsDev("ItemName").value)
                .TextMatrix(i, .ColIndex("UnitId")) = IIf(IsNull(RsDev("UnitId").value), "", RsDev("UnitId").value)
                .TextMatrix(i, .ColIndex("UnitName")) = IIf(IsNull(RsDev("UnitName").value), "", RsDev("UnitName").value)
            
                .TextMatrix(i, .ColIndex("Price")) = IIf(IsNull(RsDev("Price").value), 0, val(RsDev("Price").value))
            
                .TextMatrix(i, .ColIndex("Discount")) = IIf(IsNull(RsDev("Discount").value), 0, val(RsDev("Discount").value))
            
                RsDev.MoveNext
            Next i
 
        End With

    End If

    RsDev.Close
    'fill Group grid
   GoTo not_fill_group
    StrSQL = " SELECT     dbo.TblCustomerContractDetailsGroups.TblCustomerContractD, dbo.TblCustomerContractDetailsGroups.UnitID, dbo.TblCustomerContractDetailsGroups.GroupID, dbo.TblCustomerContractDetailsGroups.Discount, "
    StrSQL = StrSQL & "     dbo.TblCustomerContractDetailsGroups.Price , dbo.TblUnites.unitname, dbo.Groups.GroupName, dbo.Groups.Fullcode"
    StrSQL = StrSQL & " FROM         dbo.TblCustomerContractDetailsGroups INNER JOIN"
    StrSQL = StrSQL & " dbo.Groups ON dbo.TblCustomerContractDetailsGroups.GroupID = dbo.Groups.GroupID LEFT OUTER JOIN"
    StrSQL = StrSQL & " dbo.TblUnites ON dbo.TblCustomerContractDetailsGroups.UnitID = dbo.TblUnites.UnitID"
    StrSQL = StrSQL & "  where TblCustomerContractD=" & val(Me.TxtTblCustomerContractD.Text)
 
    RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsDev.BOF Or rs.EOF) Then
        RsDev.MoveFirst
    
        With Me.Grid1
    
            .Rows = .FixedRows + RsDev.RecordCount

            For i = .FixedRows To .Rows - 1
 
                .TextMatrix(i, .ColIndex("GroupID")) = IIf(IsNull(RsDev("GroupID").value), "", RsDev("GroupID").value)
            
                .TextMatrix(i, .ColIndex("Fullcode")) = IIf(IsNull(RsDev("Fullcode").value), "", RsDev("Fullcode").value)
            
                .TextMatrix(i, .ColIndex("GroupName")) = IIf(IsNull(RsDev("GroupName").value), "", RsDev("GroupName").value)
                .TextMatrix(i, .ColIndex("UnitId")) = IIf(IsNull(RsDev("UnitId").value), "", RsDev("UnitId").value)
                .TextMatrix(i, .ColIndex("UnitName")) = IIf(IsNull(RsDev("UnitName").value), "", RsDev("UnitName").value)
            
                .TextMatrix(i, .ColIndex("Price")) = IIf(IsNull(RsDev("Price").value), 0, val(RsDev("Price").value))
            
                .TextMatrix(i, .ColIndex("Discount")) = IIf(IsNull(RsDev("Discount").value), 0, val(RsDev("Discount").value))
            
                RsDev.MoveNext
            Next i
 
        End With

    End If
not_fill_group:
    ReLineGrid
    Exit Sub
ErrTrap:
End Sub
 
Private Sub Grid1_AfterEdit(ByVal Row As Long, _
                            ByVal Col As Long)
    On Error Resume Next
    Dim StrAccountCode As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long
    Dim code  As String

    With Grid1

        Select Case .ColKey(Col)
 
            Case "UnitName"
                code = .ComboData
         
                .TextMatrix(Row, .ColIndex("UnitID")) = code
                .TextMatrix(Row, .ColIndex("UnitName")) = .ComboItem
 
        End Select
   
        If Row = .Rows - 1 Then
    
            '.Rows = .Rows + 1
        End If

        ReLineGrid
    End With

End Sub

Private Sub Grid1_BeforeEdit(ByVal Row As Long, _
                             ByVal Col As Long, _
                             Cancel As Boolean)

    With Grid1

        If .ColKey(Col) <> "UnitName" Then
       
            .ComboList = ""
        End If

    End With

End Sub

Private Sub Grid1_StartEdit(ByVal Row As Long, _
                            ByVal Col As Long, _
                            Cancel As Boolean)
    Dim rs As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim Msg As String
    Dim LngItemID As Integer
    Dim MyStrList As String

    With Me.Grid

        Select Case .ColKey(Col)

            Case "UnitName"

                LngItemID = val(.TextMatrix(.Row, .ColIndex("ItemId")))

                'LngItemID = 1
                If LngItemID = 0 Then
                    Cancel = True
                Else
            
                    '        StrSQL = "SELECT TblItemsUnits.UnitID, TblUnites.UnitName "
                    '        StrSQL = StrSQL + " FROM TblUnites INNER JOIN TblItemsUnits " & _
                    '        "ON TblUnites.UnitID = TblItemsUnits.UnitID "
                    '        StrSQL = StrSQL + " Where TblItemsUnits.ItemID=" & LngItemID
                    '        StrSQL = StrSQL + " Order BY TblItemsUnits.SecOrder "
                    StrSQL = "SELECT TblUnites.UnitID, TblUnites.UnitName "
                    StrSQL = StrSQL + " FROM TblUnites   "
                
                    Set rs = New ADODB.Recordset
                    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                    If Not (rs.BOF Or rs.EOF) Then
                        MyStrList = .BuildComboList(rs, "UnitName", "UnitID")
                        '                    Grid.ColComboList = MyStrList
                        Grid.ColComboList(.ColIndex("UnitName")) = "|" & MyStrList
                    Else
                        Cancel = True
                    End If
                End If
            
        End Select

    End With

End Sub

Private Sub ISButton2_Click()
On Error GoTo ErrTrap
   If val(Me.TxtTblCustomerContractD.Text) <> 0 Then
       print_report
   End If
ErrTrap:
End Sub
Function print_report(Optional NoteSerial As String)
On Error GoTo ErrTrap
    Dim sql As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
    sql = "SELECT     dbo.TblCustomerContract.TblCustomerContractD, dbo.TblCustomerContract.CustomerId, dbo.TblCustomerContract.FromDate, dbo.TblCustomerContract.Todate, dbo.TblCustomerContract.Remarks,"
    sql = sql & "      dbo.TblCustomerContract.Locked, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.Type, dbo.TblCustemers.Fullcode, dbo.TblCustemers.CustomerandVendor,"
    sql = sql & "      dbo.TblCustomerContractDetails.Discount, dbo.TblCustomerContractDetails.Price, dbo.TblCustomerContractDetails.ItemID, dbo.TblCustomerContractDetails.UnitID,"
    sql = sql & "      dbo.TblCustomerContractDetails.TblCustomerContractD AS TblCustomerContractDETAILS, dbo.TblCustomerContractDetails.Id AS IDDetails, dbo.TblItems.ItemName, dbo.TblItems.ItemNamee,"
    sql = sql & "      dbo.TblItems.Fullcode AS FullcodeTblItems, dbo.TblItems.code, dbo.TblItems.PurchasePrice, dbo.TblUnites.UnitName, dbo.TblUnites.UnitNamee, dbo.TblItems.ItemCode"
    sql = sql & "      FROM         dbo.TblItems RIGHT OUTER JOIN"
    sql = sql & "      dbo.TblCustomerContractDetails ON dbo.TblItems.ItemID = dbo.TblCustomerContractDetails.ItemID RIGHT OUTER JOIN"
    sql = sql & "       dbo.TblCustomerContract ON dbo.TblCustomerContractDetails.TblCustomerContractD = dbo.TblCustomerContract.TblCustomerContractD LEFT OUTER JOIN"
    sql = sql & "     dbo.TblUnites ON dbo.TblCustomerContractDetails.UnitID = dbo.TblUnites.UnitID LEFT OUTER JOIN"
    sql = sql & "     dbo.TblCustemers ON dbo.TblCustomerContract.CustomerId = dbo.TblCustemers.CusID"
    sql = sql & " Where (dbo.TblCustomerContract.TblCustomerContractD = " & val(TxtTblCustomerContractD.Text) & ")"
    
     
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "CustomerContractRPT.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "CustomerContractRPTEE.rpt"
        End If
   If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If
   Set RsData = New ADODB.Recordset
    RsData.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
     If RsData.BOF Or RsData.EOF Then
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
        StrReportTitle = "" '& StrAccountName
        Else
         xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
          xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""
    End If
    xReport.ParameterFields(3).AddCurrentValue user_name
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.title
    xReport.ReportAuthor = App.title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName
    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
ErrTrap:
  End Function
Private Sub TxtModFlg_Change()

    If Me.TxtModFlg.Text = "N" Then
        CmdRemove.Enabled = True
        Ele(1).Enabled = True
        Cmd(0).Enabled = False
        Cmd(1).Enabled = False
        Cmd(4).Enabled = False
        Cmd(5).Enabled = False
        ISButton2.Enabled = False
        Cmd(2).Enabled = True
        Cmd(3).Enabled = True

    ElseIf Me.TxtModFlg.Text = "E" Then
        CmdRemove.Enabled = True
        Ele(1).Enabled = True
        Cmd(2).Enabled = True
        Cmd(3).Enabled = True
        ISButton2.Enabled = False
        Cmd(0).Enabled = False
        Cmd(1).Enabled = False
        Cmd(4).Enabled = False

        Cmd(5).Enabled = False

    Else
        Ele(1).Enabled = False
       ISButton2.Enabled = True
        CmdRemove.Enabled = False
        Cmd(2).Enabled = False
        Cmd(3).Enabled = False
        Cmd(0).Enabled = True
        Cmd(1).Enabled = True
        Cmd(4).Enabled = True

        Cmd(5).Enabled = True

    End If

End Sub


Private Sub TxtPlanID_Validate(Cancel As Boolean)
    
    If Check4.value = vbUnchecked Then Exit Sub
    If Me.TxtModFlg = "R" Then Exit Sub
        Dim s As String
           Dim Rs2 As New ADODB.Recordset
           ' If CBoBasedON.ListIndex = 0 And val(TxtPlanID.Text) <> 0 Then Exit Sub
              '  Dim s As String
            
            If CBoBasedON.ListIndex = -1 Then Exit Sub
            'Else
               
            'End If
            
            Dim StrSQL As String
            Dim orderStatus As Integer
     
     

    
            
            Set Rs2 = New ADODB.Recordset
            If CBoBasedON.ListIndex > 0 Then
                StrSQL = "select TblSalesPricesPlanDetails2.*,TblItems.ItemName,TblItems.itemcode,  TblUnites.UnitName from TblSalesPricesPlanDetails2 Inner join TblSalesPricesPlan On "
                StrSQL = StrSQL & " TblSalesPricesPlan.PlanID = TblSalesPricesPlanDetails2.PlanID"
                StrSQL = StrSQL & " Inner join TblItems On TblItems.ItemId = TblSalesPricesPlanDetails2.ItemID "
                StrSQL = StrSQL & " Inner join TblUnites On TblUnites.UnitID = TblSalesPricesPlanDetails2.UnitID "
                StrSQL = StrSQL & " where TblSalesPricesPlanDetails2.PlanId= " & val(TxtPlanID.Text)
                
                Rs2.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                      
 
            
                      
                Grid.Rows = 1
                Do While Not Rs2.EOF
                    Grid.Rows = Grid.Rows + 1
                    Grid.TextMatrix(Grid.Rows - 1, Grid.ColIndex("ItemId")) = Rs2!ItemID & ""
                    Grid.TextMatrix(Grid.Rows - 1, Grid.ColIndex("ItemCode")) = Rs2!itemcode & ""
                    Grid.TextMatrix(Grid.Rows - 1, Grid.ColIndex("ItemName")) = Rs2!ItemName & ""
                    Grid.TextMatrix(Grid.Rows - 1, Grid.ColIndex("UnitId")) = Rs2!UnitID & ""
                    Grid.TextMatrix(Grid.Rows - 1, Grid.ColIndex("UnitName")) = Rs2!UnitName & ""
  
                    If CBoBasedON.ListIndex = 1 Then
                        Grid.TextMatrix(Grid.Rows - 1, Grid.ColIndex("Price")) = Rs2!UnitWholeSalePriceNew & ""
                    ElseIf CBoBasedON.ListIndex = 2 Then
                        Grid.TextMatrix(Grid.Rows - 1, Grid.ColIndex("Price")) = Rs2!SalePriceNew & ""
                    End If
                    
                        
                        Rs2.MoveNext
                Loop
                    
            
            End If
         

'End If
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
' aladein add
'''''''''''''''''''''''''''''''''''''''''
Public Function FindRec(ByVal RecId As Long)
    On Error GoTo ErrTrap
    rs.find "TblCustomerContractD=" & RecId, , adSearchForward, 1
    If Not (rs.EOF) Then
        Retrive
        End If
    Exit Function
ErrTrap:
    If rs.EditMode <> adEditNone Then
        rs.CancelUpdate
     '   BtnUndo_Click
    End If
  End Function
Private Sub DBCboClientName_Click(Area As Integer)
  On Error Resume Next
    If val(DBCboClientName.BoundText) = 0 Then Exit Sub
    Dim Fullcode  As String
    GetCustomersDetail val(DBCboClientName.BoundText), , Fullcode, 1
    TxtSearchCode.Text = Fullcode
End Sub
Private Sub TxtSearchCode_KeyPress(KeyAscii As Integer)
   Dim CUSTID As Integer
    If KeyAscii = vbKeyReturn Then
        GetCustomersDetail CUSTID, , TxtSearchCode.Text, 1
        DBCboClientName.BoundText = CUSTID
    End If
 End Sub
'TxtItemCode
Private Sub dcitems_Click(Area As Integer)
  On Error Resume Next
    If val(dcitems.BoundText) = 0 Then Exit Sub
     Me.TxtItemCode.Text = GetItemCode(val(Me.dcitems.BoundText))
   End Sub
Private Sub TxtItemCode_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then
        If TxtItemCode.Text = "" Then
            Me.dcitems.BoundText = ""
        Else
            Me.dcitems.BoundText = GetItemID(Trim$(Me.TxtItemCode.Text))
        End If
    End If
End Sub



Public Function Retrive_Items_data()
    Dim StrSQL  As String
    Dim row_count As Long
    Dim Num As Long
    Dim i As Long
    Dim Rs2 As ADODB.Recordset
    Set Rs2 = New ADODB.Recordset
    
    If CBoBasedON.ListIndex > 0 Then
        StrSQL = "select TblItems.ItemName,TblItems.itemcode,  TblUnites.UnitName,TblItems.ItemID,TblUnites.UnitID, "
        
        If CBoBasedON.ListIndex = 1 Then
            StrSQL = StrSQL & " TblSalesPricesPlanDetails2.UnitWholeSalePriceNew UnitWholeSalePrice"
        Else
            StrSQL = StrSQL & " TblSalesPricesPlanDetails2.SalePriceNew as UnitWholeSalePrice"
        End If
        StrSQL = StrSQL & " from TblSalesPricesPlanDetails2 Inner join TblSalesPricesPlan On TblSalesPricesPlan.PlanID = TblSalesPricesPlanDetails2.PlanID"
        StrSQL = StrSQL & " Inner join TblItems On TblItems.ItemId = TblSalesPricesPlanDetails2.ItemID "
        StrSQL = StrSQL & " Inner join TblUnites On TblUnites.UnitID = TblSalesPricesPlanDetails2.UnitID "
        StrSQL = StrSQL & " where TblSalesPricesPlanDetails2.PlanId= " & val(TxtPlanID.Text)
        StrSQL = StrSQL + " and TblItems.ItemID in(" & TxtItemsIDes.Text & ")"
        StrSQL = StrSQL + " Order By   TblItems.ItemID   "
        Rs2.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
        If Not Rs2.EOF Then GoTo InsertData
        
    End If
      StrSQL = "SELECT  TblItems.* ,TblItemsUnits.UnitID, TblUnites.UnitName,TblItemsUnits.UnitWholeSalePrice "
    StrSQL = StrSQL + " FROM TblUnites INNER JOIN TblItemsUnits " & "ON TblUnites.UnitID = TblItemsUnits.UnitID "
    StrSQL = StrSQL + " Inner join TblItems On TblItems.ItemID = TblItemsUnits.ItemID"
    StrSQL = StrSQL + " Where TblItems.ItemID in(" & TxtItemsIDes.Text & ")"
                
    StrSQL = StrSQL + " Order By   TblItems.ItemID   "
    
  '  StrSQL = "select * from TblItems where ItemID in(" & TxtItemsIDes.Text & ")"
  Set Rs2 = New ADODB.Recordset
    Rs2.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
InsertData:
   If Rs2.RecordCount > 0 Then
        row_count = Grid.Rows
       ' If Grid.TextMatrix(row_count - 1, Grid.ColIndex("Code")) = "" Then
        '    row_count = row_count - 1
       ' End If
     With Grid
       Rs2.MoveFirst
       .Rows = Rs2.RecordCount + .Rows
        For Num = row_count To .Rows - 1 'RsDetails.RecordCount
            dcitems.BoundText = val(Rs2("ItemID").value & "")
        '.TextMatrix(Num, .ColIndex("ItemId")) = IIf(IsNull(Rs2("ItemID").value), 0, Rs2("ItemID").value)
            
            
            .TextMatrix(Num, .ColIndex("ItemId")) = IIf(IsNull(Rs2.Fields("ItemId").value), "", Rs2.Fields("ItemId").value)
            
            .TextMatrix(Num, .ColIndex("ItemCode")) = IIf(IsNull(Rs2.Fields("ItemCode").value), "", Rs2.Fields("ItemCode").value)
            .TextMatrix(Num, .ColIndex("ItemName")) = IIf(IsNull(Rs2.Fields("ItemName").value), "", Rs2.Fields("ItemName").value)
            
            .TextMatrix(Num, .ColIndex("UnitId")) = IIf(IsNull(Rs2.Fields("UnitId").value), "", Rs2.Fields("UnitId").value)
            .TextMatrix(Num, .ColIndex("UnitName")) = IIf(IsNull(Rs2.Fields("UnitName").value), "", Rs2.Fields("UnitName").value)
            .TextMatrix(Num, .ColIndex("Price")) = IIf(IsNull(Rs2.Fields("UnitWholeSalePrice").value), "", Rs2.Fields("UnitWholeSalePrice").value)
            
                'lllllllllllllll
     
          '  addrow Num
            Rs2.MoveNext
        Next Num
'        For i = row_count To .Rows - 1 'RsDetails.RecordCount
'          Grid_AfterEdit i, .ColIndex("ItemCode")
'        Next i
'            Grid_AfterEdit row_count, .ColIndex("ItemCode")
    End With
    End If


End Function

Public Sub CreatLog_File_for_error(str As String)
    Dim StrLogFileName As String
    Dim IntFreeFile As Integer
    Dim ss As String

    StrLogFileName = App.path & "\Gard.txt"

    If Dir(StrLogFileName) <> "" Then
        Kill StrLogFileName
    End If

    ss = "»Ì«‰ »«”„«¡  «·«’‰«› €Ì— «·„ÊÃÊœ… "
    ss = ss & vbCrLf & "Byte Informations Systems "
    ss = ss & vbCrLf & "BYTE "
    ss = ss & vbCrLf & "Create Date:- " & Now
    ss = ss & vbCrLf & str & vbCrLf
    IntFreeFile = FreeFile

    Open StrLogFileName For Output As #IntFreeFile
    Print #IntFreeFile, ss
    Close #IntFreeFile
End Sub
 
