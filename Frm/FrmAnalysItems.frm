VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmAnalysItems 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   8415
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10350
   Icon            =   "FrmAnalysItems.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8415
   ScaleWidth      =   10350
   ShowInTaskbar   =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   495
      Left            =   5880
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   8880
      Width           =   1095
      _cx             =   1931
      _cy             =   873
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
   End
   Begin C1SizerLibCtl.C1Tab C1Tab1 
      Height          =   7680
      Left            =   0
      TabIndex        =   2
      Top             =   720
      Width           =   10335
      _cx             =   18230
      _cy             =   13547
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
      Caption         =   " Õ·Ì·Ì «·«’‰«›| ﬁ«—Ì— «·‘»ﬂ…|ÿ—ﬁ «·œ›⁄"
      Align           =   0
      CurrTab         =   2
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
      Begin VB.Frame Fra 
         BackColor       =   &H00E2E9E9&
         Height          =   7260
         Index           =   2
         Left            =   45
         TabIndex        =   84
         Top             =   45
         Width           =   10245
         Begin VB.OptionButton Optrans 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "›Ê« Ì— «·„»Ì⁄«  "
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   2
            Left            =   600
            RightToLeft     =   -1  'True
            TabIndex        =   106
            Top             =   360
            Width           =   1455
         End
         Begin VB.CommandButton btnClear 
            Caption         =   "„”Õ"
            Height          =   495
            Index           =   2
            Left            =   2640
            TabIndex        =   94
            Top             =   6360
            Width           =   1335
         End
         Begin VB.Frame Frame7 
            Height          =   7095
            Left            =   5880
            TabIndex        =   92
            Top             =   120
            Width           =   4455
            Begin VB.Label Label7 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "«·”« —Ì…"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   27.75
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00008000&
               Height          =   615
               Left            =   240
               TabIndex        =   93
               Top             =   6360
               Visible         =   0   'False
               Width           =   3975
            End
            Begin VB.Image Image3 
               Height          =   5715
               Left            =   0
               Picture         =   "FrmAnalysItems.frx":038A
               Stretch         =   -1  'True
               Top             =   120
               Width           =   4395
            End
         End
         Begin VB.Frame Frame6 
            BackColor       =   &H00E2E9E9&
            Caption         =   "„‰ «·› —Â"
            Height          =   735
            Left            =   240
            RightToLeft     =   -1  'True
            TabIndex        =   87
            Top             =   1920
            Width           =   4455
            Begin MSComCtl2.DTPicker DTPicker1 
               Height          =   330
               Left            =   2280
               TabIndex        =   88
               Top             =   240
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   582
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   131137539
               CurrentDate     =   41640
            End
            Begin MSComCtl2.DTPicker DTPicker2 
               Height          =   330
               Left            =   120
               TabIndex        =   89
               Top             =   240
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   582
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   131137539
               CurrentDate     =   41640
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„‰"
               Height          =   195
               Index           =   15
               Left            =   3690
               RightToLeft     =   -1  'True
               TabIndex        =   91
               Top             =   240
               Width           =   540
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "≈·Ï"
               Height          =   195
               Index           =   1
               Left            =   1710
               RightToLeft     =   -1  'True
               TabIndex        =   90
               Top             =   240
               Width           =   480
            End
         End
         Begin VB.OptionButton Optrans 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "›Ê« Ì— "
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   0
            Left            =   4680
            RightToLeft     =   -1  'True
            TabIndex        =   86
            Top             =   360
            Width           =   975
         End
         Begin VB.OptionButton Optrans 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„ﬁ»Ê÷« "
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   1
            Left            =   2760
            RightToLeft     =   -1  'True
            TabIndex        =   85
            Top             =   360
            Width           =   1335
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   495
            Index           =   4
            Left            =   120
            TabIndex        =   95
            Top             =   6360
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   873
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
            BackStyle       =   0
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
         Begin ImpulseButton.ISButton Cmd 
            Height          =   495
            Index           =   5
            Left            =   1320
            TabIndex        =   96
            Top             =   6360
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   873
            ButtonPositionImage=   1
            Caption         =   "⁄—÷ «· ﬁ—Ì—"
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
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            ColorToggledHoverText=   16711680
            LowerToggledContent=   0   'False
            ColorTextShadow =   -2147483637
         End
         Begin MSDataListLib.DataCombo DcbBranch2 
            Height          =   315
            Left            =   120
            TabIndex        =   97
            Top             =   1440
            Width           =   4560
            _ExtentX        =   8043
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "7"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCRegionID2 
            Height          =   315
            Left            =   120
            TabIndex        =   102
            Top             =   1080
            Width           =   4560
            _ExtentX        =   8043
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "7"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCActivity 
            Bindings        =   "FrmAnalysItems.frx":28E2
            Height          =   315
            Left            =   120
            TabIndex        =   107
            Top             =   720
            Width           =   4560
            _ExtentX        =   8043
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ListField       =   "account_name"
            BoundColumn     =   "code"
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
         Begin VB.Label Label9 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "«·‰‘«ÿ"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   4680
            RightToLeft     =   -1  'True
            TabIndex        =   108
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„‰ÿﬁ…"
            Height          =   255
            Index           =   17
            Left            =   4485
            RightToLeft     =   -1  'True
            TabIndex        =   103
            Top             =   1080
            Width           =   1740
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬂÊœ «·’‰›"
            Height          =   255
            Index           =   26
            Left            =   6525
            RightToLeft     =   -1  'True
            TabIndex        =   101
            Top             =   960
            Width           =   2925
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "Õœœ «· ﬁ—Ì— «·„ÿ·Ê»"
            Height          =   255
            Left            =   3840
            TabIndex        =   100
            Top             =   120
            Width           =   1935
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Height          =   255
            Index           =   20
            Left            =   6120
            RightToLeft     =   -1  'True
            TabIndex        =   99
            Top             =   4560
            Width           =   1740
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·›—⁄"
            Height          =   255
            Index           =   18
            Left            =   4485
            RightToLeft     =   -1  'True
            TabIndex        =   98
            Top             =   1440
            Width           =   1740
         End
      End
      Begin VB.Frame Fra 
         BackColor       =   &H00E2E9E9&
         Height          =   7260
         Index           =   0
         Left            =   -10890
         TabIndex        =   58
         Top             =   45
         Width           =   10245
         Begin VB.OptionButton optNetWork 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«Ã„«·Ì  «·ÌÊ„Ì… «Ã·"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   9
            Left            =   3960
            RightToLeft     =   -1  'True
            TabIndex        =   122
            Top             =   1200
            Width           =   1815
         End
         Begin VB.OptionButton optNetWork 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«Ã„«·Ì  «·ÌÊ„Ì… ›Ì“«"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   8
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   121
            Top             =   960
            Width           =   1815
         End
         Begin VB.OptionButton optNetWork 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«Ã„«·Ì  «·ÌÊ„Ì… ‘»ﬂ…"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   7
            Left            =   2160
            RightToLeft     =   -1  'True
            TabIndex        =   120
            Top             =   960
            Width           =   1815
         End
         Begin VB.OptionButton optNetWork 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«Ã„«·Ì  «·ÌÊ„Ì… ‰ﬁœÌ"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   6
            Left            =   3960
            RightToLeft     =   -1  'True
            TabIndex        =   119
            Top             =   960
            Width           =   1815
         End
         Begin VB.OptionButton optNetWork 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«Ã„«·Ì  «·ÌÊ„Ì… ﬂ«„· "
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   5
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   118
            Top             =   720
            Width           =   1815
         End
         Begin VB.OptionButton optNetWork 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„ƒ‘—«  «·»Ì⁄  ›’Ì·Ì "
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   4
            Left            =   1800
            RightToLeft     =   -1  'True
            TabIndex        =   117
            Top             =   720
            Width           =   2175
         End
         Begin VB.OptionButton optNetWork 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„ƒ‘—«  «·»Ì⁄ «Ã„«·Ì"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   3
            Left            =   1800
            RightToLeft     =   -1  'True
            TabIndex        =   116
            Top             =   360
            Width           =   2175
         End
         Begin VB.Frame Frame8 
            Caption         =   "Õœœ"
            Height          =   615
            Left            =   240
            RightToLeft     =   -1  'True
            TabIndex        =   112
            Top             =   3840
            Width           =   4815
            Begin VB.OptionButton optPos 
               Alignment       =   1  'Right Justify
               Caption         =   "«·ﬂ·"
               Height          =   255
               Index           =   2
               Left            =   720
               TabIndex        =   115
               Top             =   240
               Width           =   1095
            End
            Begin VB.OptionButton optPos 
               Alignment       =   1  'Right Justify
               Caption         =   "ÿ··»Ì« "
               Height          =   255
               Index           =   1
               Left            =   2280
               TabIndex        =   114
               Top             =   240
               Width           =   1095
            End
            Begin VB.OptionButton optPos 
               Alignment       =   1  'Right Justify
               Caption         =   "‰ﬁÿÂ ›ﬁÿ"
               Height          =   255
               Index           =   0
               Left            =   3480
               TabIndex        =   113
               Top             =   240
               Width           =   1095
            End
         End
         Begin VB.OptionButton optNetWork 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«· Õ·Ì· »«·›Ê« Ì— "
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   2
            Left            =   360
            RightToLeft     =   -1  'True
            TabIndex        =   111
            Top             =   360
            Width           =   1455
         End
         Begin VB.ComboBox CboPayMentType 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   82
            Top             =   3240
            Width           =   4560
         End
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   3630
            TabIndex        =   79
            Top             =   2880
            Width           =   1050
         End
         Begin VB.OptionButton optNetWork 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " Õ·Ì·Ì «·‘»ﬂ…"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   1
            Left            =   4200
            RightToLeft     =   -1  'True
            TabIndex        =   76
            Top             =   720
            Width           =   1575
         End
         Begin VB.OptionButton optNetWork 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«Ã„«·Ì «·‘»ﬂ…"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   0
            Left            =   4200
            RightToLeft     =   -1  'True
            TabIndex        =   67
            Top             =   360
            Width           =   1575
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00E2E9E9&
            Caption         =   "„‰ «·› —Â"
            Height          =   1395
            Left            =   240
            RightToLeft     =   -1  'True
            TabIndex        =   62
            Top             =   4560
            Width           =   4455
            Begin MSComCtl2.DTPicker DtpDateFrom2 
               Height          =   330
               Left            =   2280
               TabIndex        =   63
               Top             =   270
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   582
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   141230083
               CurrentDate     =   41640
            End
            Begin MSComCtl2.DTPicker DtpDateTo2 
               Height          =   330
               Left            =   120
               TabIndex        =   64
               Top             =   240
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   582
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   141230083
               CurrentDate     =   41640
            End
            Begin MSComCtl2.DTPicker XPDtbTransTimeFrom 
               Height          =   285
               Left            =   2340
               TabIndex        =   123
               Top             =   810
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   503
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   141230082
               CurrentDate     =   38784
            End
            Begin MSComCtl2.DTPicker XPDtbTransTimeTo 
               Height          =   285
               Left            =   120
               TabIndex        =   125
               Top             =   780
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   503
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   141230082
               CurrentDate     =   38784
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "≈·Ï"
               Height          =   270
               Index           =   27
               Left            =   1575
               TabIndex        =   126
               Top             =   810
               Width           =   630
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "„‰"
               Height          =   270
               Index           =   21
               Left            =   3795
               TabIndex        =   124
               Top             =   840
               Width           =   630
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "≈·Ï"
               Height          =   195
               Index           =   12
               Left            =   1710
               RightToLeft     =   -1  'True
               TabIndex        =   66
               Top             =   240
               Width           =   480
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„‰"
               Height          =   195
               Index           =   11
               Left            =   3690
               RightToLeft     =   -1  'True
               TabIndex        =   65
               Top             =   240
               Width           =   540
            End
         End
         Begin VB.Frame Frame4 
            Height          =   7095
            Left            =   5880
            TabIndex        =   60
            Top             =   120
            Width           =   4455
            Begin VB.Image Image2 
               Height          =   5715
               Left            =   0
               Picture         =   "FrmAnalysItems.frx":28F7
               Stretch         =   -1  'True
               Top             =   120
               Width           =   4395
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "«·”« —Ì…"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   27.75
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00008000&
               Height          =   615
               Left            =   240
               TabIndex        =   61
               Top             =   6360
               Visible         =   0   'False
               Width           =   3975
            End
         End
         Begin VB.CommandButton btnClear 
            Caption         =   "„”Õ"
            Height          =   495
            Index           =   1
            Left            =   2640
            TabIndex        =   59
            Top             =   6360
            Width           =   1335
         End
         Begin MSDataListLib.DataCombo DCboStoreName2 
            Height          =   315
            Left            =   120
            TabIndex        =   68
            Top             =   2520
            Width           =   4560
            _ExtentX        =   8043
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "7"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   495
            Index           =   1
            Left            =   120
            TabIndex        =   69
            Top             =   6360
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   873
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
            BackStyle       =   0
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
         Begin ImpulseButton.ISButton Cmd 
            Height          =   495
            Index           =   3
            Left            =   1320
            TabIndex        =   70
            Top             =   6360
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   873
            ButtonPositionImage=   1
            Caption         =   "⁄—÷ «· ﬁ—Ì—"
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
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            ColorToggledHoverText=   16711680
            LowerToggledContent=   0   'False
            ColorTextShadow =   -2147483637
         End
         Begin MSDataListLib.DataCombo DcbBranch 
            Height          =   315
            Left            =   120
            TabIndex        =   77
            Top             =   2160
            Width           =   4560
            _ExtentX        =   8043
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "7"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcbEmp 
            Height          =   315
            Left            =   120
            TabIndex        =   80
            Top             =   2880
            Width           =   3525
            _ExtentX        =   6218
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "6"
            BoundColumn     =   ""
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCRegionID 
            Height          =   315
            Left            =   120
            TabIndex        =   104
            Top             =   1800
            Width           =   4560
            _ExtentX        =   8043
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "7"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCActivity2 
            Bindings        =   "FrmAnalysItems.frx":4E4F
            Height          =   315
            Left            =   120
            TabIndex        =   109
            Top             =   1440
            Width           =   4560
            _ExtentX        =   8043
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ListField       =   "account_name"
            BoundColumn     =   "code"
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
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "«·‰‘«ÿ"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   4680
            RightToLeft     =   -1  'True
            TabIndex        =   110
            Top             =   1440
            Width           =   1455
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„‰ÿﬁ…"
            Height          =   255
            Index           =   19
            Left            =   4485
            RightToLeft     =   -1  'True
            TabIndex        =   105
            Top             =   1800
            Width           =   1740
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÿ—Ìﬁ… «·œ›⁄"
            Height          =   255
            Index           =   13
            Left            =   4440
            RightToLeft     =   -1  'True
            TabIndex        =   81
            Top             =   3240
            Width           =   1740
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "·›—⁄ „⁄Ì‰"
            Height          =   255
            Index           =   24
            Left            =   4485
            RightToLeft     =   -1  'True
            TabIndex        =   78
            Top             =   2160
            Width           =   1740
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·»«∆⁄"
            Height          =   285
            Index           =   23
            Left            =   4275
            TabIndex        =   75
            Top             =   2940
            Width           =   1365
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Height          =   255
            Index           =   22
            Left            =   6120
            RightToLeft     =   -1  'True
            TabIndex        =   74
            Top             =   4560
            Width           =   1740
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "Õœœ «· ﬁ—Ì— «·„ÿ·Ê»"
            Height          =   255
            Left            =   3840
            TabIndex        =   73
            Top             =   120
            Width           =   1935
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "·„Œ“‰ „⁄Ì‰"
            Height          =   255
            Index           =   16
            Left            =   4440
            RightToLeft     =   -1  'True
            TabIndex        =   72
            Top             =   2520
            Width           =   1740
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬂÊœ «·’‰›"
            Height          =   255
            Index           =   14
            Left            =   6525
            RightToLeft     =   -1  'True
            TabIndex        =   71
            Top             =   960
            Width           =   2925
         End
      End
      Begin VB.Frame Fra 
         BackColor       =   &H00E2E9E9&
         Height          =   7260
         Index           =   1
         Left            =   -11190
         TabIndex        =   3
         Top             =   45
         Width           =   10245
         Begin VB.TextBox TxtItemCode 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   1
            Left            =   3120
            RightToLeft     =   -1  'True
            TabIndex        =   83
            Top             =   2580
            Width           =   1575
         End
         Begin VB.CommandButton btnClear 
            Caption         =   "„”Õ"
            Height          =   495
            Index           =   0
            Left            =   2640
            TabIndex        =   55
            Top             =   6720
            Width           =   1335
         End
         Begin VB.Frame Frame3 
            Height          =   4575
            Left            =   5880
            TabIndex        =   35
            Top             =   120
            Width           =   4455
            Begin VB.Label lblCompanyname 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "«·”« —Ì…"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   27.75
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00008000&
               Height          =   615
               Left            =   240
               TabIndex        =   36
               Top             =   3840
               Visible         =   0   'False
               Width           =   3975
            End
            Begin VB.Image Image1 
               Height          =   3675
               Left            =   0
               Picture         =   "FrmAnalysItems.frx":4E64
               Stretch         =   -1  'True
               Top             =   120
               Width           =   4395
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "„‰ «·› —Â"
            Height          =   735
            Left            =   240
            RightToLeft     =   -1  'True
            TabIndex        =   30
            Top             =   5880
            Width           =   4455
            Begin MSComCtl2.DTPicker DtpDateFrom 
               Height          =   330
               Left            =   2280
               TabIndex        =   31
               Top             =   270
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   582
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   111935491
               CurrentDate     =   41640
            End
            Begin MSComCtl2.DTPicker DtpDateTo 
               Height          =   330
               Left            =   120
               TabIndex        =   32
               Top             =   240
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   582
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   111935491
               CurrentDate     =   41640
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„‰"
               Height          =   195
               Index           =   4
               Left            =   3690
               RightToLeft     =   -1  'True
               TabIndex        =   34
               Top             =   240
               Width           =   540
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "≈·Ï"
               Height          =   195
               Index           =   3
               Left            =   1710
               RightToLeft     =   -1  'True
               TabIndex        =   33
               Top             =   240
               Width           =   480
            End
         End
         Begin VB.TextBox ItemDetailedCode 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   75
            RightToLeft     =   -1  'True
            TabIndex        =   29
            Top             =   3000
            Width           =   4560
         End
         Begin VB.TextBox ParrtNoCode 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   75
            RightToLeft     =   -1  'True
            TabIndex        =   28
            Top             =   3360
            Width           =   4560
         End
         Begin VB.OptionButton opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„Êﬁ› „Œ“Ê‰"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   0
            Left            =   4200
            RightToLeft     =   -1  'True
            TabIndex        =   27
            Top             =   360
            Width           =   1575
         End
         Begin VB.OptionButton opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«Ã„«·Ì „»Ì⁄« "
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   1
            Left            =   4200
            RightToLeft     =   -1  'True
            TabIndex        =   26
            Top             =   600
            Width           =   1575
         End
         Begin VB.OptionButton opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«Ã„«·Ì „‘ —Ì« "
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   2
            Left            =   4200
            RightToLeft     =   -1  'True
            TabIndex        =   25
            Top             =   840
            Width           =   1575
         End
         Begin VB.OptionButton opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " Õ·Ì·Ì „»Ì⁄« "
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   3
            Left            =   4200
            RightToLeft     =   -1  'True
            TabIndex        =   24
            Top             =   1080
            Width           =   1575
         End
         Begin VB.OptionButton opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " Õ·Ì·Ì „‘ —Ì« "
            ForeColor       =   &H00FF0000&
            Height          =   375
            Index           =   4
            Left            =   4200
            RightToLeft     =   -1  'True
            TabIndex        =   23
            Top             =   1320
            Width           =   1575
         End
         Begin VB.OptionButton opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«Ã„«·Ì „—œÊœ«  „»Ì⁄« "
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   5
            Left            =   2280
            RightToLeft     =   -1  'True
            TabIndex        =   22
            Top             =   600
            Width           =   1935
         End
         Begin VB.OptionButton opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«Ã„«·Ì  „—œÊœ«  „‘ —Ì« "
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   6
            Left            =   2160
            RightToLeft     =   -1  'True
            TabIndex        =   21
            Top             =   840
            Width           =   2055
         End
         Begin VB.OptionButton opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " Õ·Ì·Ì „—œÊœ«  „»Ì⁄« "
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   7
            Left            =   2160
            RightToLeft     =   -1  'True
            TabIndex        =   20
            Top             =   1080
            Width           =   2055
         End
         Begin VB.OptionButton opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " Õ·Ì·Ì „—œÊœ«  „‘ —Ì« "
            ForeColor       =   &H00FF0000&
            Height          =   375
            Index           =   8
            Left            =   2160
            RightToLeft     =   -1  'True
            TabIndex        =   19
            Top             =   1320
            Width           =   2055
         End
         Begin VB.OptionButton opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "’«›Ì «·„»Ì⁄« "
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   9
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   18
            Top             =   600
            Width           =   2055
         End
         Begin VB.OptionButton opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "’«›Ì «·„‘ —Ì« "
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   10
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   17
            Top             =   840
            Width           =   2055
         End
         Begin VB.OptionButton opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " Õ·Ì·Ì „»Ì⁄«  Ê„—œÊœ« "
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   11
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   16
            Top             =   1080
            Width           =   2055
         End
         Begin VB.OptionButton opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " Õ·Ì·Ì „‘ —Ì«  Ê„—œÊœ« "
            ForeColor       =   &H00FF0000&
            Height          =   375
            Index           =   12
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   15
            Top             =   1320
            Width           =   2175
         End
         Begin VB.OptionButton opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—’Ìœ «›  «ÕÌ «Ã„«·Ì"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   13
            Left            =   2400
            RightToLeft     =   -1  'True
            TabIndex        =   14
            Top             =   360
            Width           =   1815
         End
         Begin VB.OptionButton opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—’Ìœ «›  «ÕÌ  Õ·Ì·Ì"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   14
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   13
            Top             =   360
            Value           =   -1  'True
            Width           =   2055
         End
         Begin VB.CheckBox ChsERIAL 
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÿ»ﬁ« ··”Ì—Ì«·/«·»«—ﬂÊœ"
            Height          =   195
            Left            =   7800
            TabIndex        =   12
            Top             =   5160
            Width           =   2055
         End
         Begin VB.TextBox TxtSearchCode 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   3600
            TabIndex        =   11
            Top             =   4440
            Width           =   1050
         End
         Begin VB.Frame Frame2 
            Height          =   1095
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   5
            Top             =   4800
            Width           =   4575
            Begin VB.TextBox percent1 
               Alignment       =   1  'Right Justify
               Height          =   375
               Left            =   1920
               TabIndex        =   7
               Text            =   "0"
               Top             =   240
               Width           =   615
            End
            Begin VB.TextBox percent2 
               Alignment       =   1  'Right Justify
               Height          =   375
               Left            =   240
               TabIndex        =   6
               Text            =   "0"
               Top             =   240
               Width           =   615
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               Caption         =   "‰”»… «·‰„Ê ›Ï «·„»Ì⁄«  "
               Height          =   375
               Left            =   2520
               TabIndex        =   10
               Top             =   360
               Width           =   1815
            End
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               Caption         =   "‰”»… «·«„«‰"
               Height          =   375
               Left            =   960
               TabIndex        =   9
               Top             =   360
               Width           =   855
            End
            Begin VB.Label Label6 
               Alignment       =   1  'Right Justify
               Caption         =   "„ﬁ«—‰Â „⁄ „»Ì⁄« "
               ForeColor       =   &H00C00000&
               Height          =   495
               Left            =   1560
               TabIndex        =   8
               Top             =   720
               Width           =   2655
            End
         End
         Begin VB.OptionButton opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " ﬁ—Ì— „ Œ’’ ··ÿ·»Ì« "
            ForeColor       =   &H00FF0000&
            Height          =   375
            Index           =   15
            Left            =   5040
            RightToLeft     =   -1  'True
            TabIndex        =   4
            Top             =   5040
            Width           =   2055
         End
         Begin MSDataListLib.DataCombo DCboItemsName 
            Height          =   315
            Left            =   75
            TabIndex        =   37
            Top             =   2580
            Width           =   3000
            _ExtentX        =   5292
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCboStoreName 
            Height          =   315
            Left            =   75
            TabIndex        =   38
            Top             =   1800
            Width           =   4560
            _ExtentX        =   8043
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "7"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcbColor 
            Height          =   315
            Left            =   75
            TabIndex        =   39
            Top             =   3720
            Width           =   4560
            _ExtentX        =   8043
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcbSize 
            Height          =   315
            Left            =   75
            TabIndex        =   40
            Top             =   4080
            Width           =   4560
            _ExtentX        =   8043
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCboGroup1 
            Height          =   315
            Left            =   75
            TabIndex        =   41
            Top             =   2160
            Width           =   4560
            _ExtentX        =   8043
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "7"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DBCboClientName 
            Height          =   315
            Left            =   75
            TabIndex        =   42
            Top             =   4440
            Width           =   3525
            _ExtentX        =   6218
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "6"
            BoundColumn     =   ""
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   495
            Index           =   2
            Left            =   120
            TabIndex        =   56
            Top             =   6720
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   873
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
            BackStyle       =   0
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
         Begin ImpulseButton.ISButton Cmd 
            Height          =   495
            Index           =   0
            Left            =   1320
            TabIndex        =   57
            Top             =   6720
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   873
            ButtonPositionImage=   1
            Caption         =   "⁄—÷ «· ﬁ—Ì—"
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
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            ColorToggledHoverText=   16711680
            LowerToggledContent=   0   'False
            ColorTextShadow =   -2147483637
         End
         Begin VB.Shape Shape1 
            BorderWidth     =   2
            Height          =   975
            Left            =   6120
            Top             =   5520
            Width           =   3615
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Caption         =   "‘«‘…  ﬁ«—Ì—   Õ·Ì·Ì ··«’‰«›"
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
            Height          =   900
            Index           =   25
            Left            =   6120
            RightToLeft     =   -1  'True
            TabIndex        =   54
            Top             =   5520
            Width           =   3495
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬂÊœ «·’‰›"
            Height          =   255
            Index           =   31
            Left            =   6525
            RightToLeft     =   -1  'True
            TabIndex        =   53
            Top             =   960
            Width           =   2925
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "’‰› —∆Ì”Ì"
            Height          =   255
            Index           =   30
            Left            =   4440
            RightToLeft     =   -1  'True
            TabIndex        =   52
            Top             =   2640
            Width           =   1740
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "·„Œ“‰ „⁄Ì‰"
            Height          =   255
            Index           =   8
            Left            =   4440
            RightToLeft     =   -1  'True
            TabIndex        =   51
            Top             =   1800
            Width           =   1740
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "·ﬂÊœ  Õ·Ì·Ì „⁄Ì‰"
            Height          =   255
            Index           =   0
            Left            =   4440
            RightToLeft     =   -1  'True
            TabIndex        =   50
            Top             =   3000
            Width           =   1740
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "··Ê‰ „⁄Ì‰"
            Height          =   255
            Index           =   2
            Left            =   4440
            RightToLeft     =   -1  'True
            TabIndex        =   49
            Top             =   3720
            Width           =   1740
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "·„ﬁ«” „⁄Ì‰"
            Height          =   255
            Index           =   5
            Left            =   4440
            RightToLeft     =   -1  'True
            TabIndex        =   48
            Top             =   4080
            Width           =   1740
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "·„Ã„Ê⁄Â „⁄Ì‰…"
            Height          =   255
            Index           =   6
            Left            =   4440
            RightToLeft     =   -1  'True
            TabIndex        =   47
            Top             =   2160
            Width           =   1740
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "»«—ﬂÊœ Œ«—ÃÌ"
            Height          =   255
            Index           =   7
            Left            =   4440
            RightToLeft     =   -1  'True
            TabIndex        =   46
            Top             =   3360
            Width           =   1740
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "Õœœ «· ﬁ—Ì— «·„ÿ·Ê»"
            Height          =   255
            Left            =   3840
            TabIndex        =   45
            Top             =   120
            Width           =   1935
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Height          =   255
            Index           =   9
            Left            =   6120
            RightToLeft     =   -1  'True
            TabIndex        =   44
            Top             =   4560
            Width           =   1740
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·⁄„Ì·"
            Height          =   285
            Index           =   10
            Left            =   4275
            TabIndex        =   43
            Top             =   4500
            Width           =   1365
         End
      End
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   " Õ·Ì·Ì ‰ﬁ«ÿ «·»Ì⁄"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   15
      TabIndex        =   0
      Top             =   0
      Width           =   10320
   End
End
Attribute VB_Name = "FrmAnalysItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsSavRec As ADODB.Recordset
Dim BKGrndPic As ClsBackGroundPic
Dim RecId As String
Dim II As Long
Dim cSearch  As clsDCboSearch
Dim rs As ADODB.Recordset
Dim DCboSearch As clsDCboSearch
 Dim reportid As Integer


Private Sub ChangeLang()
On Error GoTo ErrTrap

Label9.Caption = "Activity"
Label11.Caption = "Activity"
Label5.Caption = "Items Analysis Report"
lbl(25).Caption = Label5.Caption
Label2.Caption = "Select Reports"
lblCompanyname.Caption = "AL SATTARYAH"
lbl(2).Caption = "Color"
lbl(5).Caption = "Size"
ChsERIAL.Caption = "Serials/barcode"
lbl(19).Caption = "Region"
lbl(17).Caption = "Region"
lbl(4).Caption = "From"
lbl(3).Caption = "To"
Frame1.Caption = "Period"
lbl(8).Caption = "Store"
lbl(6).Caption = "Group"
lbl(30).Caption = "Item"
lbl(0).Caption = "Code"
lbl(7).Caption = "BarCode"
Opt(0).RightToLeft = False
Opt(1).RightToLeft = False
Opt(2).RightToLeft = False
Opt(3).RightToLeft = False
Opt(4).RightToLeft = False
Opt(5).RightToLeft = False
Opt(6).RightToLeft = False
Opt(7).RightToLeft = False
Opt(8).RightToLeft = False
Opt(9).RightToLeft = False
Opt(10).RightToLeft = False
Opt(11).RightToLeft = False
Opt(12).RightToLeft = False
Opt(13).RightToLeft = False
Opt(14).RightToLeft = False
'opt(15).RightToLeft = False
BtnClear(0).Caption = "Clear"
BtnClear(1).Caption = "Clear"
Cmd(0).Caption = "Show"
Cmd(2).Caption = "Exit"
Opt(0).Caption = "Inventory Stock"
Opt(1).Caption = "Total Sales"
Opt(2).Caption = "Total Purchases"
Opt(3).Caption = "Analytical Sales"
Opt(4).Caption = "Analytical Purchases"
Opt(13).Caption = "Total Op Balance "
Opt(14).Caption = "Analytical Op Balance "
Opt(5).Caption = "Total Sales Returns "
Opt(6).Caption = "Total Purchases Returns "
Opt(7).Caption = "Anal. Sales Returns "
Opt(8).Caption = "Anal. Purchases Returns "
Opt(9).Caption = "Net Sales "
Opt(10).Caption = "Net Purchases "
Opt(11).Caption = "Analy.Sales and Returns "
Opt(12).Caption = "Analy.Purchases and Returns "
ErrTrap:
End Sub

Private Sub btnClear_Click(Index As Integer)
clear_all Me
DtpDateFrom.value = ""
DtpDateTo.value = ""
DtpDateFrom2.value = ""
DtpDateTo2.value = ""
DTPicker2.value = ""
DTPicker1.value = ""
optNetWork(0).value = True
End Sub

Private Sub Cmd_Click(Index As Integer)

    Select Case Index
        Case 0
        GetData
        Case 1
             Unload Me
        Case 2, 4
            Unload Me
        Case 3
            GetDataNetwork
        Case 5
        If Optrans(1).value = True Then
        GetDataTrans2
        Else
        GetDataTrans
        End If
    End Select

End Sub
Private Sub DBCboClientName_Click(Area As Integer)
Dim Fullcode As String

 GetCustomersDetail val(DBCboClientName.BoundText), , Fullcode, 1
    TxtSearchCode.Text = Fullcode
    
End Sub

Private Sub Form_Activate()
   PutFormOnTop Me.hwnd
End Sub

Sub FillPayment()
Dim I As Integer
Dim sql As String
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
With CboPayMentType
.Clear
.AddItem "‰ﬁœÌ"
End With
sql = "SELECT        PaymentID, PaymentName, PaymentNamee"
sql = sql & " From dbo.TblPaymentType"
sql = sql & " order by PaymentID "
Rs3.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
Rs3.MoveFirst
For I = 1 To Rs3.RecordCount
With CboPayMentType
.AddItem IIf(IsNull(Rs3("PaymentName").value), "", Rs3("PaymentName").value)
End With
Rs3.MoveNext
Next I
End If
Rs3.Close
End Sub
Private Sub TxtItemCode_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim Msg As String
    Dim StrSQL As String
    Dim rs As ADODB.Recordset

            If KeyCode = vbKeyReturn Then
                If Trim(Me.txtItemCode(Index).Text) = "" Then Exit Sub
                StrSQL = "Select ItemID From TblItems Where ItemCode='" & Trim(Me.txtItemCode(Index).Text) & "'"
                Set rs = New ADODB.Recordset
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If Not (rs.BOF Or rs.EOF) Then
                    DCboItemsName.BoundText = rs("ItemID").value
                Else
                    Msg = "·«ÌÊÃœ ’‰› „”Ã· »Â–« «·ﬂÊœ..!"
                    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                End If
            End If



 If KeyCode = vbKeyF3 Then
            Load FrmItemSearch
            FrmItemSearch.RetrunType = 1
            Set FrmItemSearch.DcboItems = Me.DCboItemsName
            FrmItemSearch.show vbModal

End If

End Sub

Private Sub TxtSearchCode_KeyPress(KeyAscii As Integer)
  Dim CUSTID As Integer

    If KeyAscii = vbKeyReturn Then
        GetCustomersDetail CUSTID, , TxtSearchCode.Text, 1
        DBCboClientName.BoundText = CUSTID
    End If

End Sub



Private Sub Form_Load()
   
 'On Error GoTo ErrTrap
    Dim I As Integer
    Dim My_SQL As String
    Dim Dcombos As ClsDataCombos
    Dim StrSQL As String
    FillPayment
  
    Set Dcombos = New ClsDataCombos
      Dcombos.GetStores Me.DCboStoreName2
      Dcombos.GetBranches Me.DcbBranch
      Dcombos.GetCustomersSuppliers 1, Me.DBCboClientName
      Dcombos.GetItemsSizes Me.DcbSize
      Dcombos.GetItemsColors Me.DcbColor
      Dcombos.GetItemsNames Me.DCboItemsName
      Dcombos.GetStores Me.DCboStoreName
      Dcombos.GetItemSGroups Me.DCboGroup1, False
      Dcombos.GetSalesRepData Me.DcbEmp
      Dcombos.GetBranches Me.DcbBranch2
      Dcombos.GetSection Me.DCRegionID
      Dcombos.GetSection Me.DCRegionID2
     If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "  select id,name from tblActivitesType   "
    Else
        StrSQL = "  select id,namee from tblActivitesType   "
    End If
    fill_combo DCActivity, StrSQL
    fill_combo DCActivity2, StrSQL

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
        SwitchKeyboardLang LANG_ENGLISH
        Else
        SwitchKeyboardLang LANG_ARABIC
    End If
    DTPicker2.value = Date
    DTPicker1.value = Date
    DTPicker2.value = ""
    DTPicker1.value = ""
DtpDateFrom.value = ""
DtpDateTo.value = ""
DtpDateFrom2.value = Date
DtpDateTo2.value = Date
    Resize_Form Me

End Sub


Private Sub Form_Unload(Cancel As Integer)

    FormPostion Me, SavePostion
    Set DCboSearch = Nothing
End Sub
Sub GetDataTrans2()
Dim sql As String
Dim BrnchesReg As String
Dim BrnchAct As String
    BrnchesReg = BranchRegion(val(DCRegionID2.BoundText))
    BrnchAct = BrcnhActivityType(val(DCActivity.BoundText))
Dim Rs2 As ADODB.Recordset
Set Rs2 = New ADODB.Recordset
sql = " SELECT     TradingContractID, dbo.Notes.NoteID, dbo.Notes.NoteDate, dbo.Notes.NoteType, dbo.Notes.NoteSerial1, dbo.Notes.Note_Value, dbo.Notes.CashingType, "
sql = sql & "                     dbo.Notes.CusID, dbo.TblCustemers.Fullcode, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.Notes.project_id, dbo.projects.Project_name,"
sql = sql & "                      dbo.projects.Project_account, dbo.projects.opening_balance_voucher_id, isnull(dbo.TblMultuPayment.PaymentID,0)as PaymentID, dbo.TblMultuPayment.[Value],"
sql = sql & "                      dbo.TblMultuPayment.CardNo, dbo.TblPaymentType.PaymentName, dbo.TblPaymentType.PaymentNamee, dbo.Notes.branch_no, dbo.TblBranchesData.branch_name,"
sql = sql & "                      dbo.TblBranchesData.branch_namee, dbo.Notes.NoteCashingType, dbo.Notes.EmployeeID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode AS Expr1,"
sql = sql & "                      dbo.TblEmployee.Emp_Namee , dbo.Notes.AccountsCode, dbo.Accounts.account_name, dbo.Accounts.account_serial, dbo.Accounts.Account_NameEng"
sql = sql & " FROM         dbo.ACCOUNTS RIGHT OUTER JOIN"
sql = sql & "                      dbo.Notes ON dbo.ACCOUNTS.Account_Code = dbo.Notes.AccountsCode LEFT OUTER JOIN"
sql = sql & "                      dbo.TblEmployee ON dbo.Notes.EmployeeID = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
sql = sql & "                      dbo.TblBranchesData ON dbo.Notes.branch_no = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
sql = sql & "                      dbo.TblPaymentType RIGHT OUTER JOIN"
sql = sql & "                      dbo.TblMultuPayment ON dbo.TblPaymentType.PaymentID = dbo.TblMultuPayment.PaymentID ON dbo.Notes.NoteID = dbo.TblMultuPayment.NoteID AND ISNULL(TblMultuPayment.[Value],0) <> 0 LEFT OUTER JOIN"
sql = sql & "                      dbo.projects ON dbo.Notes.project_id = dbo.projects.id LEFT OUTER JOIN"
sql = sql & "                      dbo.TblCustemers ON dbo.Notes.CusID = dbo.TblCustemers.CusID"
sql = sql & " Where (dbo.Notes.notetype = 4) "
If BrnchesReg <> "-1" Then
        sql = sql & " AND dbo.Notes.branch_no in( " & BrnchesReg & " )"
End If
If BrnchAct <> "-1" Then
        sql = sql & " AND dbo.Notes.branch_no in( " & BrnchAct & " )"
End If
If val(DcbBranch2.BoundText) <> 0 Then
sql = sql & " and dbo.Notes.branch_no =" & val(DcbBranch2.BoundText) & ""
Else
sql = sql & " AND      dbo.Notes.branch_no  in(" & Current_branchSql & ")"
End If
      If Not IsNull(Me.DTPicker1.value) Then
                   sql = sql & " AND dbo.Notes.NoteDate >=" & SQLDate(Me.DTPicker1.value, True) & ""
      End If
       If Not IsNull(Me.DTPicker2.value) Then
                   sql = sql & " AND dbo.Notes.NoteDate <=" & SQLDate(Me.DTPicker2.value, True) & ""
      End If
    sql = sql & " order by dbo.TblMultuPayment.PaymentID"
print_report2 sql, 2
End Sub
Sub GetDataTrans()
Dim sql As String
Dim BrnchesReg As String
Dim BrnchAct As String

       BrnchesReg = BranchRegion(val(DCRegionID2.BoundText))
      BrnchAct = BrcnhActivityType(val(DCActivity.BoundText))

Dim Rs2 As ADODB.Recordset
Set Rs2 = New ADODB.Recordset
sql = " SELECT     dbo.Transactions.Transaction_ID, dbo.Transactions.Transaction_Date, dbo.Transactions.Transaction_Type, dbo.TransactionTypes.TransactionTypeName, "
sql = sql & "                      dbo.TransactionTypes.TransactionEnglishName, dbo.Transactions.Transaction_HijriDate, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee,"
sql = sql & "                      dbo.TblCustemers.Fullcode, dbo.Transactions.CusID, dbo.Transactions.NoteSerial1, dbo.TblCustemers.Cus_Phone, dbo.TblCustemers.Cus_mobile,"
sql = sql & "                      dbo.Transactions.CashCustomerName, dbo.Transactions.CashCustomerPhone, dbo.Transactions.CashCustomerMobile, dbo.TblSalesPayment.[Value],"
sql = sql & "                      dbo.TblSalesPayment.CardNo,"
sql = sql & "      PaymentName=Case "
sql = sql & "     When  Transactions.PaymentType=1   Then '«Ã·' "
 sql = sql & "    Else  "
 
sql = sql & "   ISNULL(dbo.TblPaymentType.PaymentName, N'‰ﬁœÌ')  "
sql = sql & "           END,"
sql = sql & "       dbo.TblPaymentType.PaymentNamee,"
'ISNULL(dbo.TblPaymentType.PaymentName, '????') AS PaymentName, dbo.TblPaymentType.PaymentNamee,"
sql = sql & "                      dbo.Transactions.PaymentType, dbo.Transactions.BranchId, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee,"
sql = sql & "                      dbo.Transactions.Transaction_NetValue, dbo.Transactions.VAT, ISNULL(dbo.TblSalesPayment.PaymentID, 0) AS PaymentID"
sql = sql & "  FROM         dbo.TblBranchesData RIGHT OUTER JOIN"
sql = sql & "                       dbo.Transactions ON dbo.TblBranchesData.branch_id = dbo.Transactions.BranchId LEFT OUTER JOIN"
sql = sql & "                       dbo.TblPaymentType RIGHT OUTER JOIN"
sql = sql & "                       dbo.TblSalesPayment ON dbo.TblPaymentType.PaymentID = dbo.TblSalesPayment.PaymentID ON"
sql = sql & "                       dbo.Transactions.Transaction_ID = dbo.TblSalesPayment.TransID LEFT OUTER JOIN"
sql = sql & "                       dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
sql = sql & "                       dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type"
If Optrans(2).value = True Then
sql = sql & "  WHERE  (dbo.Transactions.Transaction_Type = 21) AND (dbo.TblSalesPayment.[Value] <> 0 OR"
sql = sql & "                      dbo.TblSalesPayment.[Value] IS NULL)"
Else
sql = sql & "  WHERE  (dbo.Transactions.Transaction_Type = 21 OR"
sql = sql & "                      dbo.Transactions.Transaction_Type = 22) AND (dbo.TblSalesPayment.[Value] <> 0 OR"
sql = sql & "                      dbo.TblSalesPayment.[Value] IS NULL)"
End If
    If BrnchesReg <> "-1" Then
        sql = sql & " AND dbo.Transactions.BranchId in( " & BrnchesReg & " )"
    End If
        If BrnchAct <> "-1" Then
        sql = sql & " AND dbo.Transactions.BranchId in( " & BrnchAct & " )"
    End If
If val(DcbBranch2.BoundText) <> 0 Then
sql = sql & " AND dbo.Transactions.BranchId =" & val(DcbBranch2.BoundText) & ""
Else
sql = sql & " AND      dbo.Transactions.BranchId   in(" & Current_branchSql & ")"
End If
      If Not IsNull(Me.DTPicker1.value) Then
                   sql = sql & " AND dbo.Transactions.Transaction_Date >=" & SQLDate(Me.DTPicker1.value, True) & ""
      End If
       If Not IsNull(Me.DTPicker2.value) Then
                   sql = sql & " AND dbo.Transactions.Transaction_Date<=" & SQLDate(Me.DTPicker2.value, True) & ""
      End If
      
      
    
      
            
      
    sql = sql & " order by dbo.Transactions.Transaction_Type, dbo.TblSalesPayment.PaymentID"
print_report2 sql, 1
End Sub
Public Sub GetData()
    Dim StrSQL As String
      Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim I As Integer
If Opt(0).value = True Or Opt(1).value = True Or Opt(2).value = True Or Opt(5).value = True Or Opt(6).value = True Or Opt(9).value = True Or Opt(10).value = True Or Opt(13).value = True Then
reportid = 0
ElseIf Opt(15).value = True Then
reportid = 15
Else
reportid = 1
End If

'StrSQL = " SELECT     TOP 100 PERCENT dbo.ItemsDetails.ItemDetailedCode, dbo.ItemsDetails.ParrtNoCode, SUM(dbo.ItemsDetails.[Count] * dbo.TransactionTypes.StockEffect) "
'StrSQL = StrSQL & "                       AS countsactual, dbo.TblStore.StoreName, dbo.TblStore.StoreNamee, dbo.TblItemsclasses.SizeName AS [cclASS NAME], dbo.TblItemsColors.ColorName,"
'StrSQL = StrSQL & "                       dbo.TblItemsSizes.SizeName, dbo.TblUnites.UnitName, dbo.TblUnites.UnitNamee, dbo.Transactions.StoreID, dbo.Transactions.Transaction_Date,"
'StrSQL = StrSQL & "                       dbo.TblItems.ItemCode , dbo.TblItems.itemname, dbo.TblItems.ItemNamee, dbo.ItemsDetails.ItemID"
'StrSQL = StrSQL & "  FROM         dbo.ItemsDetails INNER JOIN"
'StrSQL = StrSQL & "                       dbo.Transactions ON dbo.ItemsDetails.Transaction_ID = dbo.Transactions.Transaction_ID INNER JOIN"
'StrSQL = StrSQL & "                       dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type INNER JOIN"
'StrSQL = StrSQL & "                       dbo.TblStore ON dbo.Transactions.StoreID = dbo.TblStore.StoreID LEFT OUTER JOIN"
'StrSQL = StrSQL & "                       dbo.TblItems ON dbo.ItemsDetails.ItemId = dbo.TblItems.ItemID LEFT OUTER JOIN"
'StrSQL = StrSQL & "                       dbo.TblUnites ON dbo.ItemsDetails.UnitID = dbo.TblUnites.UnitID LEFT OUTER JOIN"
'StrSQL = StrSQL & "                       dbo.TblItemsColors ON dbo.ItemsDetails.ColorID = dbo.TblItemsColors.ColorID LEFT OUTER JOIN"
'StrSQL = StrSQL & "                       dbo.TblItemsSizes ON dbo.ItemsDetails.SizeID = dbo.TblItemsSizes.SizeId LEFT OUTER JOIN"
'StrSQL = StrSQL & "                       dbo.TblItemsclasses ON dbo.ItemsDetails.ClassId = dbo.TblItemsclasses.SizeId"
'StrSQL = StrSQL & " where 1=1"


'StrSQL = "SELECT     dbo.ItemsDetails.ItemDetailedCode, dbo.ItemsDetails.ParrtNoCode, SUM(dbo.ItemsDetails.[Count] * dbo.TransactionTypes.StockEffect) "
'StrSQL = StrSQL & " AS countsactual, dbo.TblStore.StoreName, dbo.TblStore.StoreNamee, dbo.TblItemsclasses.SizeName AS [cclASS NAME], dbo.TblItemsColors.ColorName,"
'StrSQL = StrSQL & " dbo.TblItemsSizes.SizeName, dbo.TblUnites.UnitName, dbo.TblUnites.UnitNamee, dbo.Transactions.StoreID, dbo.TblItems.ItemCode, dbo.TblItems.ItemName,"
'StrSQL = StrSQL & " dbo.TblItems.ItemNamee , dbo.ItemsDetails.ItemID"
'StrSQL = StrSQL & " FROM         dbo.ItemsDetails INNER JOIN"
'StrSQL = StrSQL & " dbo.Transactions ON dbo.ItemsDetails.Transaction_ID = dbo.Transactions.Transaction_ID INNER JOIN"
'StrSQL = StrSQL & "   dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type INNER JOIN"
'StrSQL = StrSQL & "   dbo.TblStore ON dbo.Transactions.StoreID = dbo.TblStore.StoreID LEFT OUTER JOIN"
'StrSQL = StrSQL & "   dbo.TblItems ON dbo.ItemsDetails.ItemId = dbo.TblItems.ItemID LEFT OUTER JOIN"
'StrSQL = StrSQL & "   dbo.TblUnites ON dbo.ItemsDetails.UnitID = dbo.TblUnites.UnitID LEFT OUTER JOIN"
'StrSQL = StrSQL & "   dbo.TblItemsColors ON dbo.ItemsDetails.ColorID = dbo.TblItemsColors.ColorID LEFT OUTER JOIN"
'StrSQL = StrSQL & "   dbo.TblItemsSizes ON dbo.ItemsDetails.SizeID = dbo.TblItemsSizes.SizeId LEFT OUTER JOIN"
'StrSQL = StrSQL & "    dbo.TblItemsclasses ON dbo.ItemsDetails.ClassId = dbo.TblItemsclasses.SizeId"
 
 StrSQL = "  SELECT     SUM(dbo.ItemsDetails.[Count] * dbo.ItemsDetails.EffectN) AS countsactual, "
StrSQL = StrSQL & " dbo.TblStore.StoreName, dbo.TblStore.StoreNamee, dbo.TblItemsclasses.SizeName AS [cclASS NAME], dbo.TblItemsColors.ColorName, dbo.TblItemsSizes.SizeName,"
StrSQL = StrSQL & " dbo.TblUnites.UnitName, dbo.TblUnites.UnitNamee, dbo.Transactions.StoreID, dbo.TblItems.ItemCode, dbo.TblItems.ItemName, dbo.TblItems.ItemNamee,"
StrSQL = StrSQL & " dbo.ItemsDetails.ItemID , dbo.Groups.GroupName, dbo.Groups.GroupNamee"
StrSQL = StrSQL & "  FROM         dbo.Groups INNER JOIN"
StrSQL = StrSQL & "  dbo.TblItems ON dbo.Groups.GroupID = dbo.TblItems.GroupID RIGHT OUTER JOIN"
StrSQL = StrSQL & "  dbo.ItemsDetails INNER JOIN"
StrSQL = StrSQL & "  dbo.Transactions ON dbo.ItemsDetails.Transaction_ID = dbo.Transactions.Transaction_ID INNER JOIN"
StrSQL = StrSQL & "  dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type INNER JOIN"
StrSQL = StrSQL & "  dbo.TblStore ON dbo.Transactions.StoreID = dbo.TblStore.StoreID ON dbo.TblItems.ItemID = dbo.ItemsDetails.ItemId LEFT OUTER JOIN"
StrSQL = StrSQL & "  dbo.TblUnites ON dbo.ItemsDetails.UnitID = dbo.TblUnites.UnitID LEFT OUTER JOIN"
StrSQL = StrSQL & "  dbo.TblItemsColors ON dbo.ItemsDetails.ColorID = dbo.TblItemsColors.ColorID LEFT OUTER JOIN"
StrSQL = StrSQL & "  dbo.TblItemsSizes ON dbo.ItemsDetails.SizeID = dbo.TblItemsSizes.SizeId LEFT OUTER JOIN"
StrSQL = StrSQL & "  dbo.TblItemsclasses ON dbo.ItemsDetails.ClassId = dbo.TblItemsclasses.SizeId"
 

StrSQL = StrSQL & "   Where (1 = 1)"

 
 If reportid = 1 Then
 StrSQL = "SELECT       dbo.Transactions.NoteSerial1, dbo.TransactionTypes.Transaction_Type, dbo.TransactionTypes.TransactionTypeName, dbo.TransactionTypes.TransactionEnglishName, "
StrSQL = StrSQL & "   dbo.ItemsDetails.ParrtNoCode, SUM(dbo.ItemsDetails.[Count] * dbo.ItemsDetails.EffectN) AS countsactual,"
StrSQL = StrSQL & " dbo.TblStore.StoreName, dbo.TblStore.StoreNamee, dbo.TblItemsclasses.SizeName AS [cclASS NAME], dbo.TblItemsColors.ColorName, dbo.TblItemsSizes.SizeName,"
StrSQL = StrSQL & " dbo.TblUnites.UnitName, dbo.TblUnites.UnitNamee, dbo.Transactions.StoreID, dbo.TblItems.ItemCode, dbo.TblItems.ItemName, dbo.TblItems.ItemNamee,"
StrSQL = StrSQL & " dbo.ItemsDetails.ItemID , dbo.Groups.GroupName, dbo.Groups.GroupNamee, dbo.Transactions.Transaction_Date, dbo.Transactions.Transaction_ID"
StrSQL = StrSQL & " FROM         dbo.Groups INNER JOIN"
StrSQL = StrSQL & "                      dbo.TblItems ON dbo.Groups.GroupID = dbo.TblItems.GroupID RIGHT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TransactionTypes INNER JOIN"
StrSQL = StrSQL & "                      dbo.Transactions ON dbo.TransactionTypes.Transaction_Type = dbo.Transactions.Transaction_Type INNER JOIN"
StrSQL = StrSQL & "                      dbo.TblStore ON dbo.Transactions.StoreID = dbo.TblStore.StoreID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.ItemsDetails ON dbo.Transactions.Transaction_ID = dbo.ItemsDetails.Transaction_ID ON dbo.TblItems.ItemID = dbo.ItemsDetails.ItemId LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblUnites ON dbo.ItemsDetails.UnitID = dbo.TblUnites.UnitID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblItemsColors ON dbo.ItemsDetails.ColorID = dbo.TblItemsColors.ColorID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblItemsSizes ON dbo.ItemsDetails.SizeID = dbo.TblItemsSizes.SizeId LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblItemsclasses ON dbo.ItemsDetails.ClassId = dbo.TblItemsclasses.SizeId"
StrSQL = StrSQL & "  Where (1 = 1)  "

 End If
 '''''''''''''''''''
 
  If ItemDetailedCode.Text <> "" Then
     StrSQL = StrSQL & " AND dbo.ItemsDetails.ItemDetailedCode like '%" & Me.ItemDetailedCode.Text & "%'"
    End If
    
     If ParrtNoCode.Text <> "" Then
     StrSQL = StrSQL & " AND dbo.ItemsDetails.ParrtNoCode like '%" & Me.ParrtNoCode.Text & "%'"
    End If
    
    
    
    
    
If Me.DCboStoreName.Text <> "" And val(DCboStoreName.BoundText) <> 0 Then
    StrSQL = StrSQL & " AND   dbo.Transactions.StoreID = " & val(Me.DCboStoreName.BoundText)

End If

If Me.DCboItemsName.Text <> "" And val(DCboItemsName.BoundText) <> 0 Then
    StrSQL = StrSQL & " AND   dbo.ItemsDetails.ItemId = " & val(Me.DCboItemsName.BoundText)

End If
If Me.DcbColor.Text <> "" And val(DcbColor.BoundText) <> 0 Then
    StrSQL = StrSQL & " AND   dbo.ItemsDetails.ColorID = " & val(Me.DcbColor.BoundText)

End If
If Me.DcbSize.Text <> "" And val(DcbSize.BoundText) <> 0 Then
    StrSQL = StrSQL & " AND   dbo.ItemsDetails.SizeID = " & val(Me.DcbSize.BoundText)

End If

If Me.DCboGroup1.Text <> "" And val(DCboGroup1.BoundText) <> 0 Then
    StrSQL = StrSQL & " AND    dbo.TblItems.GroupID= " & val(Me.DCboGroup1.BoundText)

End If

If Opt(0).value = True Then

ElseIf Opt(1).value = True Or Opt(3).value = True Then
 StrSQL = StrSQL & " AND dbo.Transactions.Transaction_Type =21"
 
 ElseIf Opt(1).value = True Or Opt(3).value = True Then
 StrSQL = StrSQL & " AND dbo.Transactions.Transaction_Type =21"
 
 ElseIf Opt(2).value = True Or Opt(4).value = True Then
 StrSQL = StrSQL & " AND dbo.Transactions.Transaction_Type =22"
 ElseIf Opt(5).value = True Or Opt(7).value = True Then
 StrSQL = StrSQL & " AND dbo.Transactions.Transaction_Type =9"
 ElseIf Opt(6).value = True Or Opt(8).value = True Then
 StrSQL = StrSQL & " AND dbo.Transactions.Transaction_Type =5"
 
 ElseIf Opt(9).value = True Or Opt(11).value = True Then
 StrSQL = StrSQL & " AND  ( dbo.Transactions.Transaction_Type =9  or dbo.Transactions.Transaction_Type =21 )"
 
 
 ElseIf Opt(10).value = True Or Opt(12).value = True Then
 StrSQL = StrSQL & " AND  ( dbo.Transactions.Transaction_Type =5 or  dbo.Transactions.Transaction_Type =22 )"
  ElseIf Opt(13).value = True Or Opt(14).value = True Then
  StrSQL = StrSQL & " AND dbo.Transactions.Transaction_Type =3"
End If






 If DBCboClientName.BoundText <> "" And DBCboClientName.Text <> "" Then
                   StrSQL = StrSQL & " AND dbo.Transactions.CusID =" & val(DBCboClientName.BoundText)
      End If
 If Not IsNull(Me.DtpDateFrom.value) Then
                   StrSQL = StrSQL & " AND dbo.Transactions.Transaction_Date >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
      End If
       If Not IsNull(Me.DtpDateTo.value) Then
                   StrSQL = StrSQL & " AND dbo.Transactions.Transaction_Date<=" & SQLDate(Me.DtpDateTo.value, True) & ""
      End If
  If reportid = 0 Then
 StrSQL = StrSQL & "  GROUP BY  dbo.ItemsDetails.ColorID, dbo.ItemsDetails.ClassId, dbo.Transactions.StoreID, "
StrSQL = StrSQL & "  dbo.TblStore.StoreName, dbo.TblStore.StoreNamee, dbo.TblItemsclasses.SizeName, dbo.TblItemsColors.ColorName, dbo.TblItemsSizes.SizeName,"
StrSQL = StrSQL & "  dbo.TblUnites.UnitName, dbo.TblUnites.UnitNamee, dbo.TblItems.ItemCode, dbo.TblItems.ItemName, dbo.TblItems.ItemNamee, dbo.ItemsDetails.ItemId,"
StrSQL = StrSQL & "  dbo.Groups.GroupName , dbo.Groups.GroupNamee"
StrSQL = StrSQL & "  ORDER BY dbo.TblItems.ItemCode"
Else
StrSQL = StrSQL & "   GROUP BY dbo.ItemsDetails.ParrtNoCode, dbo.ItemsDetails.ColorID, dbo.ItemsDetails.ClassId, dbo.Transactions.StoreID, "
StrSQL = StrSQL & "  dbo.TblStore.StoreName, dbo.TblStore.StoreNamee, dbo.TblItemsclasses.SizeName, dbo.TblItemsColors.ColorName, dbo.TblItemsSizes.SizeName,"
StrSQL = StrSQL & " dbo.TblUnites.UnitName, dbo.TblUnites.UnitNamee, dbo.TblItems.ItemCode, dbo.TblItems.ItemName, dbo.TblItems.ItemNamee, dbo.ItemsDetails.ItemId,"
StrSQL = StrSQL & " dbo.Groups.GroupName, dbo.Groups.GroupNamee, dbo.TblItems.GroupID, dbo.TransactionTypes.TransactionTypeName,"
StrSQL = StrSQL & " dbo.TransactionTypes.TransactionEnglishName, dbo.Transactions.Transaction_Date, dbo.Transactions.Transaction_ID,"
StrSQL = StrSQL & " dbo.TransactionTypes.Transaction_Type,  dbo.Transactions.NoteSerial1"
 
StrSQL = StrSQL & "  ORDER BY dbo.Transactions.Transaction_Date, dbo.Transactions.Transaction_ID"

End If

If reportid = 15 Then
StrSQL = "SELECT       SUM(dbo.ItemsDetails.[Count] * dbo.ItemsDetails.EffectN) AS countsactual, dbo.ItemsDetails.ColorID, dbo.ItemsDetails.SizeID, "

If val(DCboStoreName.BoundText) <> 0 Then
'strSQL = strSQL & "   dbo.ItemsDetails.ItemId, dbo.GardTransactionDetails1(" & SQLDate(Me.DtpDateFrom.value, True) & ", dbo.ItemsDetails.ColorID, dbo.ItemsDetails.SizeID, dbo.ItemsDetails.ItemId,  " & val(DCboStoreName.BoundText) & ") AS QtyAvilable, "

 StrSQL = StrSQL & "   dbo.ItemsDetails.ItemId, dbo.GardTransactionDetails1('" & SQLDate(Date, False) & "', dbo.ItemsDetails.ColorID, dbo.ItemsDetails.SizeID, dbo.ItemsDetails.ItemId, " & val(DCboStoreName.BoundText) & ") AS QtyAvilable, "
Else
 
StrSQL = StrSQL & "   dbo.ItemsDetails.ItemId, dbo.GardTransactionDetails1('" & SQLDate(Date, False) & "', dbo.ItemsDetails.ColorID, dbo.ItemsDetails.SizeID, dbo.ItemsDetails.ItemId, Null) AS QtyAvilable, "
End If

StrSQL = StrSQL & "                        dbo.TblItems.ItemNamee , dbo.TblItems.ItemName, dbo.TblItemsSizes.sizename, dbo.TblItemsColors.colorname, dbo.TblItems.fullcode"
StrSQL = StrSQL & "  FROM         dbo.Groups INNER JOIN"
StrSQL = StrSQL & "                        dbo.TblItems ON dbo.Groups.GroupID = dbo.TblItems.GroupID RIGHT OUTER JOIN"
StrSQL = StrSQL & "                        dbo.ItemsDetails INNER JOIN"
StrSQL = StrSQL & "                        dbo.Transactions ON dbo.ItemsDetails.Transaction_ID = dbo.Transactions.Transaction_ID INNER JOIN"
 StrSQL = StrSQL & "  dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type INNER JOIN"
StrSQL = StrSQL & "                        dbo.TblStore ON dbo.Transactions.StoreID = dbo.TblStore.StoreID ON dbo.TblItems.ItemID = dbo.ItemsDetails.ItemId LEFT OUTER JOIN"
StrSQL = StrSQL & "                        dbo.TblUnites ON dbo.ItemsDetails.UnitID = dbo.TblUnites.UnitID LEFT OUTER JOIN"
StrSQL = StrSQL & "                        dbo.TblItemsColors ON dbo.ItemsDetails.ColorID = dbo.TblItemsColors.ColorID LEFT OUTER JOIN"
StrSQL = StrSQL & "                        dbo.TblItemsSizes ON dbo.ItemsDetails.SizeID = dbo.TblItemsSizes.SizeId LEFT OUTER JOIN"
StrSQL = StrSQL & "                        dbo.TblItemsclasses ON dbo.ItemsDetails.ClassId = dbo.TblItemsclasses.SizeId"
StrSQL = StrSQL & "  WHERE     (1 = 1)   "
'StrSQL = StrSQL & "   (dbo.Transactions.StoreID = 2) "
'StrSQL = StrSQL & "   AND (dbo.ItemsDetails.ItemId = 70)"
'StrSQL = StrSQL & "   AND (dbo.Transactions.Transaction_Date >= '01-Oct-2016')  "
'StrSQL = StrSQL & "    And                    (dbo.Transactions.Transaction_Date <= '01-Oct-2016') "
StrSQL = StrSQL & "   AND (dbo.Transactions.Transaction_Type = 21)"

  If ItemDetailedCode.Text <> "" Then
     StrSQL = StrSQL & " AND dbo.ItemsDetails.ItemDetailedCode like '%" & Me.ItemDetailedCode.Text & "%'"
    End If
    
     If ParrtNoCode.Text <> "" Then
     StrSQL = StrSQL & " AND dbo.ItemsDetails.ParrtNoCode like '%" & Me.ParrtNoCode.Text & "%'"
    End If
    
    
    
    
    
If Me.DCboStoreName.Text <> "" And val(DCboStoreName.BoundText) <> 0 Then
    StrSQL = StrSQL & " AND   dbo.Transactions.StoreID = " & val(Me.DCboStoreName.BoundText)

End If

If Me.DCboItemsName.Text <> "" And val(DCboItemsName.BoundText) <> 0 Then
    StrSQL = StrSQL & " AND   dbo.ItemsDetails.ItemId = " & val(Me.DCboItemsName.BoundText)

End If
If Me.DcbColor.Text <> "" And val(DcbColor.BoundText) <> 0 Then
    StrSQL = StrSQL & " AND   dbo.ItemsDetails.ColorID = " & val(Me.DcbColor.BoundText)

End If
If Me.DcbSize.Text <> "" And val(DcbSize.BoundText) <> 0 Then
    StrSQL = StrSQL & " AND   dbo.ItemsDetails.SizeID = " & val(Me.DcbSize.BoundText)

End If

If Me.DCboGroup1.Text <> "" And val(DCboGroup1.BoundText) <> 0 Then
    StrSQL = StrSQL & " AND    dbo.TblItems.GroupID= " & val(Me.DCboGroup1.BoundText)

End If


 If DBCboClientName.BoundText <> "" And DBCboClientName.Text <> "" Then
                   StrSQL = StrSQL & " AND dbo.Transactions.CusID =" & val(DBCboClientName.BoundText)
      End If
 If Not IsNull(Me.DtpDateFrom.value) Then
                   StrSQL = StrSQL & " AND dbo.Transactions.Transaction_Date >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
      End If
       If Not IsNull(Me.DtpDateTo.value) Then
                   StrSQL = StrSQL & " AND dbo.Transactions.Transaction_Date<=" & SQLDate(Me.DtpDateTo.value, True) & ""
      End If
      
      
StrSQL = StrSQL & "  GROUP BY dbo.ItemsDetails.ColorID, dbo.ItemsDetails.SizeID, dbo.ItemsDetails.ItemId, dbo.TblItems.ItemNamee, dbo.TblItems.ItemName, dbo.TblItemsSizes.SizeName,"
StrSQL = StrSQL & "                        dbo.TblItemsColors.colorname , dbo.TblItems.fullcode"
StrSQL = StrSQL & "   ORDER BY dbo.TblItemsColors.ColorName"


End If

    
    
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
           ' Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
          '  Me.lbl(10).Caption = "Search Results=0"
        End If

        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «· ﬁ—Ì—"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    Else
 rs.MoveFirst
 print_report StrSQL
            If SystemOptions.UserInterface = ArabicInterface Then
             '   Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
               ' Me.lbl(10).Caption = "Search Results=" & rs.RecordCount
            End If
    End If

End Sub
Public Sub GetDataNetwork()
    Dim StrSQL As String
      Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim I As Integer
    Dim BrnchesReg As String
    Dim BrnchAct As String
        BrnchesReg = BranchRegion(val(DCRegionID.BoundText))
        BrnchAct = BrcnhActivityType(val(DCActivity2.BoundText))
        
        If optNetWork(3).value = True Or optNetWork(4).value = True Then
 '       StrSQL = "  SELECT     dbo.Transactions.last_changed,   dbo.Transactions.Transaction_Type, dbo.Transactions.Transaction_Date, dbo.Transactions.Transaction_NetValue, dbo.Transactions.Transaction_ID, dbo.Transaction_Details.Item_ID, dbo.Transaction_Details.ShowQty, "
 '       StrSQL = StrSQL & "                                       dbo.Transactions.noteserial1, { fn HOUR(dbo.Transactions.last_changed) } AS hourx"
 '       StrSQL = StrSQL & "                 FROM            dbo.Transactions INNER JOIN"
 '       StrSQL = StrSQL & "                                          dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID"
 '       StrSQL = StrSQL & "                 Where (dbo.transactions.Transaction_Type = 21)"

StrSQL = "SELECT     CAST(last_changed AS TIME) Time2,dbo.Transactions.last_changed, dbo.Transactions.Transaction_Type, dbo.Transactions.Transaction_Date, dbo.Transactions.Transaction_NetValue, "
     StrSQL = StrSQL & "                               dbo.Transactions.Transaction_ID, dbo.Transactions.NoteSerial1, { fn HOUR(dbo.Transactions.last_changed) } AS hourx, dbo.TblBranchesData.branch_id,"
     StrSQL = StrSQL & "                               dbo.TblEmployee.Emp_Name, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblEmployee.Emp_Namee, dbo.TblStore.StoreName,"
     StrSQL = StrSQL & "                               dbo.TblStore.storenamee"
     StrSQL = StrSQL & "         FROM         dbo.Transactions INNER JOIN"
     StrSQL = StrSQL & "                               dbo.TblStore ON dbo.Transactions.StoreID = dbo.TblStore.StoreID INNER JOIN"
     StrSQL = StrSQL & "                               dbo.TblEmployee ON dbo.Transactions.Emp_ID = dbo.TblEmployee.Emp_ID INNER JOIN"
     StrSQL = StrSQL & "                               dbo.TblBranchesData ON dbo.Transactions.BranchId = dbo.TblBranchesData.branch_id"
     StrSQL = StrSQL & "         WHERE     (dbo.Transactions.Transaction_Type = 21) "
                            
    If Me.DcbEmp.Text <> "" And val(DcbEmp.BoundText) <> 0 Then
        StrSQL = StrSQL & " AND    Emp_ID= " & val(Me.DcbEmp.BoundText)
    End If
    
                            If Me.DCboStoreName2.Text <> "" And val(DCboStoreName2.BoundText) <> 0 Then
        StrSQL = StrSQL & " AND   Transactions.StoreID = " & val(Me.DCboStoreName2.BoundText)
    End If
     If BrnchesReg <> "-1" Then
        StrSQL = StrSQL & " AND Transactions.BranchId in( " & BrnchesReg & " )"
       End If
      If BrnchAct <> "-1" Then
        StrSQL = StrSQL & " AND Transactions.BranchId in( " & BrnchAct & " )"
       End If
    If Me.DcbBranch.Text <> "" And val(DcbBranch.BoundText) <> 0 Then
        StrSQL = StrSQL & " AND   Transactions.BranchId = " & val(Me.DcbBranch.BoundText)
    End If
    
    
    If Not IsNull(Me.DtpDateFrom2.value) Then
        StrSQL = StrSQL & " AND Transactions.Transaction_Date >=" & SQLDate(Me.DtpDateFrom2.value, True) & ""
    End If
    If Not IsNull(Me.DtpDateTo2.value) Then
        StrSQL = StrSQL & " AND Transactions.Transaction_Date<=" & SQLDate(Me.DtpDateTo2.value, True) & ""
    End If
    
      If Not IsNull(Me.XPDtbTransTimeFrom.value) Then
       
                   StrSQL = StrSQL & " AND CAST(last_changed as time) >='" & FormatDateTime(Me.XPDtbTransTimeFrom.value, vbShortTime) & "'"
      End If
       If Not IsNull(Me.XPDtbTransTimeFrom.value) Then
                   StrSQL = StrSQL & " AND CAST(last_changed as time)<='" & FormatDateTime(Me.XPDtbTransTimeTo.value, vbShortTime) & "'"
      End If
      
If optPos(0).value = True Then
StrSQL = StrSQL & " AND   Transactions.POSBillType =1 "

ElseIf optPos(1).value = True Then
StrSQL = StrSQL & " AND    isnull(Transactions.POSBillType,0) =0 "
 
End If
  '  StrSQL = StrSQL & " ORDER BY last_changed,CAST(last_changed AS TIME)"
    GoTo xl:
        End If
        
    StrSQL = "select * from (( SELECT     POSBillType,dbo.Transactions.CashCustomerName,dbo.Transactions.PaymentType, dbo.TblTransactionPayments.id, dbo.TblTransactionPayments.Transaction_ID, dbo.TblTransactionPayments.PaymentID, ISNULL(dbo.TblTransactionPayments.[value], "
    StrSQL = StrSQL & "                   dbo.Transactions.Transaction_NetValue) AS Value, dbo.TblTransactionPayments.CardNo, dbo.Transactions.Transaction_Date, dbo.Transactions.Transaction_NetValue,"
    StrSQL = StrSQL & "                    dbo.Transactions.Transaction_HijriDate, dbo.Transactions.Transaction_Serial, dbo.Transactions.NoteSerial, dbo.Transactions.NoteSerial1,"
    StrSQL = StrSQL & "                   dbo.TblPaymentType.PaymentName, dbo.TblPaymentType.PaymentNamee, dbo.Transactions.Emp_ID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Name1,"
    StrSQL = StrSQL & "                   dbo.TblEmployee.Emp_Name2, dbo.TblEmployee.Emp_Name3, dbo.TblEmployee.Emp_Name4, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee4,"
    StrSQL = StrSQL & "                   dbo.TblEmployee.Emp_Namee3, dbo.TblEmployee.Emp_Namee2, dbo.TblEmployee.Emp_Namee1, dbo.TblEmployee.Emp_Namee,"
    StrSQL = StrSQL & "                   dbo.TransactionTypes.TransactionTypeName, dbo.TransactionTypes.TransactionEnglishName, dbo.Transactions.Transaction_Type, dbo.Transactions.StoreID,"
    StrSQL = StrSQL & "                   dbo.TblStore.StoreName, dbo.TblStore.StoreNamee, dbo.TblStore.Code, dbo.Transactions.BranchId, dbo.TblBranchesData.branch_name,"
    StrSQL = StrSQL & "                   dbo.TblBranchesData.branch_namee"
    StrSQL = StrSQL & "        FROM         dbo.TblEmployee INNER JOIN"
    StrSQL = StrSQL & "                   dbo.Transactions ON dbo.TblEmployee.Emp_ID = dbo.Transactions.Emp_ID INNER JOIN"
    StrSQL = StrSQL & "                   dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type INNER JOIN"
    StrSQL = StrSQL & "                   dbo.TblStore ON dbo.Transactions.StoreID = dbo.TblStore.StoreID INNER JOIN"
    StrSQL = StrSQL & "                  dbo.TblBranchesData ON dbo.Transactions.BranchId = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
    StrSQL = StrSQL & "                    dbo.TblTransactionPayments ON dbo.Transactions.Transaction_ID = dbo.TblTransactionPayments.Transaction_ID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.TblPaymentType ON dbo.TblTransactionPayments.PaymentID = dbo.TblPaymentType.PaymentID)"
    'where  "
    'StrSQL = StrSQL & "                  not(id is null)  and "
  '  StrSQL = StrSQL & "                  value>0)"
    StrSQL = StrSQL & " Union (SELECT   POSBillType,dbo.Transactions.CashCustomerName,dbo.Transactions.PaymentType,  dbo.TblSalesPayment.ID AS id, dbo.TblSalesPayment.TransID AS Transaction_ID, dbo.TblSalesPayment.PaymentID, ISNULL(dbo.TblSalesPayment.[Value], "
    StrSQL = StrSQL & "                  dbo.Transactions.Transaction_NetValue) AS value, dbo.TblSalesPayment.CardNo, dbo.Transactions.Transaction_Date, dbo.Transactions.Transaction_NetValue,"
    StrSQL = StrSQL & "                  dbo.Transactions.Transaction_HijriDate, dbo.Transactions.Transaction_Serial, dbo.Transactions.NoteSerial, dbo.Transactions.NoteSerial1,"
    StrSQL = StrSQL & "                  dbo.TblPaymentType.PaymentName, dbo.TblPaymentType.PaymentNamee, dbo.Transactions.Emp_ID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Name1,"
    StrSQL = StrSQL & "                  dbo.TblEmployee.Emp_Name2, dbo.TblEmployee.Emp_Name3, dbo.TblEmployee.Emp_Name4, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee4,"
    StrSQL = StrSQL & "                  dbo.TblEmployee.Emp_Namee3, dbo.TblEmployee.Emp_Namee2, dbo.TblEmployee.Emp_Namee1, dbo.TblEmployee.Emp_Namee,"
    StrSQL = StrSQL & "                  dbo.TransactionTypes.TransactionTypeName, dbo.TransactionTypes.TransactionEnglishName, dbo.Transactions.Transaction_Type, dbo.Transactions.StoreID,"
    StrSQL = StrSQL & "                  dbo.TblStore.StoreName, dbo.TblStore.StoreNamee, dbo.TblStore.Code, dbo.Transactions.BranchId, dbo.TblBranchesData.branch_name,"
    StrSQL = StrSQL & "                  dbo.TblBranchesData.branch_namee"
    StrSQL = StrSQL & "         FROM         dbo.TblSalesPayment RIGHT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.Transactions INNER JOIN"
    StrSQL = StrSQL & "                  dbo.TblEmployee ON dbo.Transactions.Emp_ID = dbo.TblEmployee.Emp_ID INNER JOIN"
    StrSQL = StrSQL & "                  dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type INNER JOIN"
    StrSQL = StrSQL & "                  dbo.TblStore ON dbo.Transactions.StoreID = dbo.TblStore.StoreID INNER JOIN"
    StrSQL = StrSQL & "                  dbo.TblBranchesData ON dbo.Transactions.BranchId = dbo.TblBranchesData.branch_id ON"
    StrSQL = StrSQL & "                  dbo.TblSalesPayment.TransID = dbo.Transactions.Transaction_ID LEFT OUTER JOIN"
'StrSQL = StrSQL & "                  dbo.TblPaymentType ON dbo.TblSalesPayment.PaymentID = dbo.TblPaymentType.PaymentID where dbo.TblSalesPayment.PaymentID =0  )) as x where (Transaction_Type = 9 OR Transaction_Type = 21)"
    StrSQL = StrSQL & "                  dbo.TblPaymentType ON dbo.TblSalesPayment.PaymentID = dbo.TblPaymentType.PaymentID where    value>0  )"
   
StrSQL = StrSQL & "          Union ( "
StrSQL = StrSQL & "   SELECT     dbo.Transactions.POSBillType, dbo.Transactions.CashCustomerName, dbo.Transactions.PaymentType, dbo.TblTransactionPayments.id, "
StrSQL = StrSQL & "                        dbo.TblTransactionPayments.Transaction_ID, dbo.TblTransactionPayments.PaymentID, ISNULL(dbo.TblTransactionPayments.[value],"
StrSQL = StrSQL & "                        dbo.Transactions.Transaction_NetValue) AS Value, dbo.TblTransactionPayments.CardNo, dbo.Transactions.Transaction_Date, dbo.Transactions.Transaction_NetValue,"
StrSQL = StrSQL & "                         dbo.Transactions.Transaction_HijriDate, dbo.Transactions.Transaction_Serial, dbo.Transactions.NoteSerial, dbo.Transactions.NoteSerial1,"
StrSQL = StrSQL & "                        dbo.TblPaymentType.PaymentName, dbo.TblPaymentType.PaymentNamee, dbo.Transactions.Emp_ID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Name1,"
StrSQL = StrSQL & "                        dbo.TblEmployee.Emp_Name2, dbo.TblEmployee.Emp_Name3, dbo.TblEmployee.Emp_Name4, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee4,"
StrSQL = StrSQL & "                        dbo.TblEmployee.Emp_Namee3, dbo.TblEmployee.Emp_Namee2, dbo.TblEmployee.Emp_Namee1, dbo.TblEmployee.Emp_Namee,"
StrSQL = StrSQL & "                        dbo.TransactionTypes.TransactionTypeName, dbo.TransactionTypes.TransactionEnglishName, dbo.Transactions.Transaction_Type, dbo.Transactions.StoreID,"
StrSQL = StrSQL & "                        dbo.TblStore.StoreName, dbo.TblStore.StoreNamee, dbo.TblStore.Code, dbo.Transactions.BranchId, dbo.TblBranchesData.branch_name,"
StrSQL = StrSQL & "                        dbo.TblBranchesData.branch_namee"
StrSQL = StrSQL & "  FROM         dbo.TblEmployee INNER JOIN"
StrSQL = StrSQL & "                        dbo.Transactions ON dbo.TblEmployee.Emp_ID = dbo.Transactions.Emp_ID INNER JOIN"
StrSQL = StrSQL & "                        dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type INNER JOIN"
StrSQL = StrSQL & "                        dbo.TblStore ON dbo.Transactions.StoreID = dbo.TblStore.StoreID INNER JOIN"
StrSQL = StrSQL & "  dbo.TblBranchesData ON dbo.Transactions.BranchId = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
StrSQL = StrSQL & "                        dbo.TblTransactionPayments ON dbo.Transactions.Transaction_ID = dbo.TblTransactionPayments.Transaction_ID LEFT OUTER JOIN"
StrSQL = StrSQL & "                        dbo.TblPaymentType ON dbo.TblTransactionPayments.PaymentID = dbo.TblPaymentType.PaymentID"
StrSQL = StrSQL & "   where (dbo.transactions.POSBillType Is Null) And (dbo.transactions.Transaction_Type = 9)"
StrSQL = StrSQL & "  )"
StrSQL = StrSQL & "  ) as x where (Transaction_Type = 9 OR Transaction_Type = 21)"
' StrSQL = StrSQL & " AND  not( Transaction_ID  is null)"
  
    If Me.DCboStoreName2.Text <> "" And val(DCboStoreName2.BoundText) <> 0 Then
        StrSQL = StrSQL & " AND   StoreID = " & val(Me.DCboStoreName2.BoundText)
    End If
     If BrnchesReg <> "-1" Then
        StrSQL = StrSQL & " AND BranchId in( " & BrnchesReg & " )"
       End If
      If BrnchAct <> "-1" Then
        StrSQL = StrSQL & " AND BranchId in( " & BrnchAct & " )"
       End If
    If Me.DcbBranch.Text <> "" And val(DcbBranch.BoundText) <> 0 Then
        StrSQL = StrSQL & " AND   BranchId = " & val(Me.DcbBranch.BoundText)
    End If
    If Me.CboPayMentType.Text <> "" And val(CboPayMentType.ListIndex) <> -1 Then
        StrSQL = StrSQL & " AND    PaymentID= " & val(Me.CboPayMentType.ListIndex)
    End If
    If Me.DcbEmp.Text <> "" And val(DcbEmp.BoundText) <> 0 Then
        StrSQL = StrSQL & " AND    Emp_ID= " & val(Me.DcbEmp.BoundText)
    End If


    If Not IsNull(Me.DtpDateFrom2.value) Then
        StrSQL = StrSQL & " AND Transaction_Date >=" & SQLDate(Me.DtpDateFrom2.value, True) & ""
    End If
    If Not IsNull(Me.DtpDateTo2.value) Then
        StrSQL = StrSQL & " AND Transaction_Date<=" & SQLDate(Me.DtpDateTo2.value, True) & ""
    End If
    
If optPos(0).value = True Then
StrSQL = StrSQL & " AND   POSBillType =1 "

ElseIf optPos(1).value = True Then
StrSQL = StrSQL & " AND    isnull(POSBillType,0) =0 "
 
End If
    Set rs = New ADODB.Recordset
    
    If optNetWork(6).value = True Then
 StrSQL = StrSQL & " AND         PaymentType<>1"
StrSQL = StrSQL & " AND    PaymentID = 0"
ElseIf optNetWork(7).value = True Then
 StrSQL = StrSQL & " AND         PaymentType<>1"
StrSQL = StrSQL & " AND    PaymentID in(2,5,7)"

 
ElseIf optNetWork(8).value = True Then
 StrSQL = StrSQL & " AND         PaymentType<>1"
StrSQL = StrSQL & " AND    PaymentID in(4,6,8)"

ElseIf optNetWork(9).value = True Then
 StrSQL = StrSQL & " AND         PaymentType=1"
  
    End If
    
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «· ﬁ—Ì—"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    Else
 rs.MoveFirst
xl:
 If optNetWork(0).value = True Then
 print_report StrSQL, 1
 ElseIf optNetWork(1).value = True Then
 print_report StrSQL, 2
  ElseIf optNetWork(2).value = True Then
 print_report StrSQL, 3
 
   ElseIf optNetWork(5).value = True Then
 print_report StrSQL, 6
 
 
    ElseIf optNetWork(6).value = True Then
 print_report StrSQL, 7
 
 
    ElseIf optNetWork(7).value = True Then
 print_report StrSQL, 8
 
 
    ElseIf optNetWork(8).value = True Then
 print_report StrSQL, 9
 
 
    ElseIf optNetWork(9).value = True Then
 print_report StrSQL, 10
  
  
 
   ElseIf optNetWork(3).value = True Then
   StrSQL = StrSQL & " ORDER BY Transaction_Date,CAST(last_changed AS TIME) ASC"
 print_report StrSQL, 4
   ElseIf optNetWork(4).value = True Then
 print_report StrSQL, 5
 
End If
    End If
End Sub
Function print_report2(Optional NoteSerial As String, Optional Ind As Integer = 0)
     
    Set rs = New ADODB.Recordset
    rs.Open NoteSerial, Cn, adOpenStatic, adLockReadOnly, adCmdText
     Debug.Print NoteSerial
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
    If Ind = 1 Then
    If Optrans(2).value = True Then
         If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepTransNetWorkSales.rpt"
            Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepTransNetWorkSales.rpt"
         End If
   Else
          If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepTransNetWork.rpt"
            Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepTransNetWork.rpt"
          End If
   End If
   ElseIf Ind = 2 Then
         If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepTransNetWork2.rpt"
            Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepTransNetWork2.rpt"
            
       End If
 End If
    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open NoteSerial, Cn, adOpenStatic, adLockReadOnly, adCmdText

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

    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
  
    End If

   If Ind = 2 Or Ind = 1 Then
  If Not IsNull(DTPicker1.value) And Not IsNull(DTPicker2.value) Then
   xReport.ParameterFields(8).AddCurrentValue DTPicker1.value
    xReport.ParameterFields(10).AddCurrentValue DTPicker2.value
    End If
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
Function print_report(Optional NoteSerial As String, Optional Ind As Integer = 0)
     
    Set rs = New ADODB.Recordset
    rs.Open NoteSerial, Cn, adOpenStatic, adLockReadOnly, adCmdText
     Debug.Print NoteSerial
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
    If Ind = 1 Then
         If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepSlesPointNewtWorkTotal.rpt"
            Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepSlesPointNewtWorkTotalE.rpt"
            
       End If
     ElseIf Ind = 2 Then
         If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepSlesPointNewtWorkAnaly.rpt"
            Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepSlesPointNewtWorkAnalyE.rpt"
            
       End If
         ElseIf Ind = 3 Then
         If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepSlesPointNewtWorkAnalysis2.rpt"
            Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepSlesPointNewtWorkAnalysis2E.rpt"
            
       End If
       
       
             ElseIf Ind = 4 Then
         If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "kpi1.rpt"
            Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "kpi1.rpt"
            
       End If
       
            ElseIf Ind = 5 Then
         If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "kpi2.rpt"
            Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "kpi2.rpt"
            
       End If
       
       
                   ElseIf Ind = 6 Then
         If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepSlesPointNewtWorkAnalysis2days.rpt"
            Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepSlesPointNewtWorkAnalysis2days.rpt"
            
       End If
       
       
       
       
'********************************************

                   ElseIf Ind = 7 Then
         If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepSlesPointNewtWorkAnalysis2dayscash.rpt"
            Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepSlesPointNewtWorkAnalysis2dayscash.rpt"
            
       End If
      
                   ElseIf Ind = 8 Then
         If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepSlesPointNewtWorkAnalysis2daysMada.rpt"
            Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepSlesPointNewtWorkAnalysis2daysMada.rpt"
            
       End If
       
                          ElseIf Ind = 9 Then
         If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepSlesPointNewtWorkAnalysis2daysvisa.rpt"
            Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepSlesPointNewtWorkAnalysis2daysvisa.rpt"
            
       End If
       
      
                         ElseIf Ind = 10 Then
         If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepSlesPointNewtWorkAnalysis2daysCredit.rpt"
            Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepSlesPointNewtWorkAnalysis2daysvisaCredit.rpt"
            
       End If
        
        
  '********************************************
       
       
    Else
    
   If reportid = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAnalysisItems.rpt"
            Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAnalysisItems.rpt"
            
       End If
       
       If ChsERIAL.value = vbChecked Then
       StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAnalysisItemsSerials.rpt"
       End If
       
      ElseIf reportid = 15 Then '«·ÿ·»Ì«  «·ÃœÌœ…
      
      
      
            If SystemOptions.UserInterface = ArabicInterface Then
        
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "DetailsOrderNew.rpt"
            Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "DetailsOrderNew.rpt"
            
       End If
       
       
      
Else


       If SystemOptions.UserInterface = ArabicInterface Then
        
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAnalysisItems1.rpt"
            Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAnalysisItems1.rpt"
            
       End If
       
       If ChsERIAL.value = vbChecked Then
       StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAnalysisItems1Serials.rpt"
       End If
              
              
End If
End If

    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open NoteSerial, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        'GetMsgs 138, vbExclamation
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

Dim I As Integer
For I = 0 To 14
 If Opt(I).value = True Then
 StrReportTitle = Opt(I).Caption
 
 If DBCboClientName.Text <> "" Then
 StrReportTitle = StrReportTitle & CHR(13) & "  ··⁄„Ì· : " & DBCboClientName.Text
 End If
 
 
 If DCboStoreName.Text <> "" Then
 StrReportTitle = StrReportTitle & CHR(13) & "  ··„Œ“‰ : " & DCboStoreName.Text
 End If
 
  If DCboGroup1.Text <> "" Then
StrReportTitle = StrReportTitle & CHR(13) & "  ··„Ã„Ê⁄Â : " & DCboGroup1.Text
 End If
 
 If ItemDetailedCode.Text <> "" Then
 StrReportTitle = StrReportTitle & CHR(13) & "  ··ﬂÊœ : " & ItemDetailedCode.Text
 End If
 
  If ParrtNoCode.Text <> "" Then
 StrReportTitle = StrReportTitle & CHR(13) & "  ··»«—ﬂÊœ : " & ParrtNoCode.Text
 End If
 
   If DcbColor.Text <> "" Then
 StrReportTitle = StrReportTitle & CHR(13) & "  ··Ê‰ : " & DcbColor.Text
 End If
 
 
   If DcbSize.Text <> "" Then
 StrReportTitle = StrReportTitle & CHR(13) & "  ··„ﬁ«” : " & DcbSize.Text
 End If
 
 
 
 
 
 End If
 
 
  
Next I


    If SystemOptions.UserInterface = ArabicInterface Then
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
        If reportid = 15 Then
        xReport.ParameterFields(12).AddCurrentValue val(percent1.Text)
        xReport.ParameterFields(13).AddCurrentValue val(percent2.Text)
        
        End If

    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        'xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
 '       StrReportTitle = ""
  
    End If

   If Ind = 0 Then
  If Not IsNull(DtpDateFrom.value) And Not IsNull(DtpDateTo.value) Then
   xReport.ParameterFields(8).AddCurrentValue DtpDateFrom.value
    xReport.ParameterFields(10).AddCurrentValue DtpDateTo.value
    End If
Else
  If Not IsNull(DtpDateFrom2.value) And Not IsNull(DtpDateTo2.value) Then
   xReport.ParameterFields(8).AddCurrentValue DtpDateFrom2.value
    xReport.ParameterFields(10).AddCurrentValue DtpDateTo2.value
    End If
End If
  Dim Total As String
  Dim totl As Double


    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.title
    xReport.ReportAuthor = App.title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName, , NoteSerial

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault

End Function




