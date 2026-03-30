VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SUITEC~1.OCX"
Begin VB.Form FrmHajjReports 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   6990
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10365
   Icon            =   "FrmHajjReports.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6990
   ScaleWidth      =   10365
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton btnClear 
      BackColor       =   &H00E2E9E9&
      Caption         =   "„”Õ"
      Height          =   495
      Left            =   1440
      TabIndex        =   40
      Top             =   6480
      Width           =   1335
   End
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   495
      Left            =   5880
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   7080
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
   Begin C1SizerLibCtl.C1Tab XPTab301 
      Height          =   5775
      Left            =   0
      TabIndex        =   2
      Top             =   600
      Width           =   10440
      _cx             =   18415
      _cy             =   10186
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
      Caption         =   " ﬁ«—Ì— | ﬁ«—Ì— «·⁄„—…  «Õ’«∆Ì…| ﬁ«—Ì—  «·⁄„—… «Ê«„— «· ‘€Ì·| ﬁ«—Ì— «·⁄„—… Ãœ«Ê· «· —ÕÌ· |  ﬁ«—Ì— √Œ—Ï| ﬁ«—Ì— «·ÕÃ"
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
      Picture(0)      =   "FrmHajjReports.frx":038A
      Begin VB.Frame Fra 
         BackColor       =   &H00E2E9E9&
         Height          =   5310
         Index           =   5
         Left            =   11685
         TabIndex        =   179
         Top             =   45
         Width           =   10350
         Begin VB.Frame Frame18 
            BackColor       =   &H00E2E9E9&
            Height          =   735
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   212
            Top             =   480
            Width           =   5775
            Begin XtremeSuiteControls.RadioButton RdHajEtmad 
               Height          =   255
               Index           =   0
               Left            =   3600
               TabIndex        =   213
               Top             =   120
               Width           =   1935
               _Version        =   786432
               _ExtentX        =   3413
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "«⁄ „«œ«  ·Â« „ÿ«·»…"
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton RdHajEtmad 
               Height          =   255
               Index           =   1
               Left            =   840
               TabIndex        =   214
               Top             =   120
               Width           =   1935
               _Version        =   786432
               _ExtentX        =   3413
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "«⁄ „«œ«  ·Ì” ·Â« „ÿ«·»…"
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton RdHajEtmad 
               Height          =   255
               Index           =   2
               Left            =   2040
               TabIndex        =   215
               Top             =   360
               Width           =   1935
               _Version        =   786432
               _ExtentX        =   3413
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "ﬂ· «·«⁄ „«œ« "
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
         End
         Begin VB.Frame Frame22 
            BackColor       =   &H00E2E9E9&
            Caption         =   "„‰ «·› —Â"
            Height          =   1095
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   191
            Top             =   4080
            Width           =   4575
            Begin MSComCtl2.DTPicker DtpDateFrom5 
               Height          =   330
               Left            =   1800
               TabIndex        =   192
               Top             =   240
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   582
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   94175235
               CurrentDate     =   41640
            End
            Begin MSComCtl2.DTPicker DtpDateTo5 
               Height          =   330
               Left            =   1800
               TabIndex        =   193
               Top             =   600
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   582
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   94175235
               CurrentDate     =   41640
            End
            Begin Dynamic_Byte.NourHijriCal DtpDateFromH5 
               Height          =   315
               Left            =   120
               TabIndex        =   194
               Top             =   240
               Width           =   1635
               _ExtentX        =   2884
               _ExtentY        =   556
            End
            Begin Dynamic_Byte.NourHijriCal DtpDateToH5 
               Height          =   315
               Left            =   120
               TabIndex        =   195
               Top             =   630
               Width           =   1635
               _ExtentX        =   2884
               _ExtentY        =   556
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„‰"
               Height          =   195
               Index           =   65
               Left            =   3450
               RightToLeft     =   -1  'True
               TabIndex        =   197
               Top             =   360
               Width           =   540
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "≈·Ï"
               Height          =   195
               Index           =   64
               Left            =   3510
               RightToLeft     =   -1  'True
               TabIndex        =   196
               Top             =   720
               Width           =   480
            End
         End
         Begin VB.Frame Frame21 
            BackColor       =   &H00E2E9E9&
            Height          =   495
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   188
            Top             =   0
            Width           =   5775
            Begin XtremeSuiteControls.RadioButton RdHaj 
               Height          =   255
               Index           =   0
               Left            =   3240
               TabIndex        =   189
               Top             =   120
               Width           =   2415
               _Version        =   786432
               _ExtentX        =   4260
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "  ﬁ—Ì— «⁄ „«œ «·ÕÃ"
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton RdHaj 
               Height          =   255
               Index           =   1
               Left            =   960
               TabIndex        =   190
               Top             =   120
               Width           =   1815
               _Version        =   786432
               _ExtentX        =   3201
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "  ﬁ—Ì— «⁄ „«œ «·„‘«⁄—"
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
         End
         Begin VB.Frame Frame20 
            Height          =   5175
            Left            =   5880
            TabIndex        =   186
            Top             =   120
            Width           =   4455
            Begin VB.Label Label6 
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
               Height          =   1095
               Left            =   960
               TabIndex        =   187
               Top             =   4080
               Width           =   2895
            End
            Begin VB.Image Image6 
               Height          =   3675
               Left            =   120
               Picture         =   "FrmHajjReports.frx":0724
               Stretch         =   -1  'True
               Top             =   120
               Width           =   4395
            End
         End
         Begin VB.Frame Frame19 
            BackColor       =   &H00E2E9E9&
            Height          =   735
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   180
            Top             =   1080
            Width           =   5775
            Begin VB.TextBox Text6 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   182
               Top             =   240
               Width           =   1155
            End
            Begin VB.TextBox Text5 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   2400
               RightToLeft     =   -1  'True
               TabIndex        =   181
               Top             =   240
               Width           =   1155
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„  «·«⁄ „«œ"
               Height          =   195
               Index           =   63
               Left            =   4320
               RightToLeft     =   -1  'True
               TabIndex        =   185
               Top             =   240
               Width           =   1185
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "≈·Ï"
               Height          =   315
               Index           =   62
               Left            =   1560
               RightToLeft     =   -1  'True
               TabIndex        =   184
               Top             =   240
               Width           =   645
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "„‰"
               Height          =   315
               Index           =   61
               Left            =   3600
               RightToLeft     =   -1  'True
               TabIndex        =   183
               Top             =   240
               Width           =   660
            End
         End
         Begin MSDataListLib.DataCombo DcbPath5 
            Height          =   315
            Left            =   120
            TabIndex        =   198
            Top             =   3000
            Width           =   4290
            _ExtentX        =   7567
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin ImpulseButton.ISButton ISButton4 
            Height          =   975
            Left            =   4800
            TabIndex        =   199
            Top             =   4200
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   1720
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
         Begin MSDataListLib.DataCombo DcbVehicleType5 
            Height          =   315
            Left            =   120
            TabIndex        =   200
            Top             =   2640
            Width           =   4290
            _ExtentX        =   7567
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcbSeasonsID5 
            Height          =   315
            Left            =   120
            TabIndex        =   201
            Top             =   1920
            Width           =   4290
            _ExtentX        =   7567
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo CompanyID5 
            Height          =   315
            Left            =   120
            TabIndex        =   206
            Top             =   2310
            Width           =   4290
            _ExtentX        =   7567
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcbDepandID5 
            Height          =   315
            Left            =   120
            TabIndex        =   208
            Top             =   3360
            Width           =   4290
            _ExtentX        =   7567
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo Nationality5 
            Height          =   315
            Left            =   120
            TabIndex        =   210
            Top             =   3720
            Width           =   4290
            _ExtentX        =   7567
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·Ã‰”Ì…"
            Height          =   210
            Index           =   59
            Left            =   4500
            RightToLeft     =   -1  'True
            TabIndex        =   211
            Top             =   3720
            Width           =   1320
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·«⁄ „«œ"
            Height          =   210
            Index           =   58
            Left            =   4500
            RightToLeft     =   -1  'True
            TabIndex        =   209
            Top             =   3360
            Width           =   1320
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„ƒ””…"
            Height          =   210
            Index           =   71
            Left            =   4500
            RightToLeft     =   -1  'True
            TabIndex        =   207
            Top             =   2280
            Width           =   1320
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·›∆…"
            Height          =   210
            Index           =   69
            Left            =   4440
            RightToLeft     =   -1  'True
            TabIndex        =   205
            Top             =   2640
            Width           =   1320
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "Œÿ «·”Ì—"
            Height          =   210
            Index           =   70
            Left            =   4500
            RightToLeft     =   -1  'True
            TabIndex        =   204
            Top             =   3000
            Width           =   1320
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            ForeColor       =   &H00000080&
            Height          =   285
            Index           =   67
            Left            =   240
            TabIndex        =   203
            Top             =   3720
            Width           =   1785
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„Ê”„"
            Height          =   210
            Index           =   66
            Left            =   4500
            RightToLeft     =   -1  'True
            TabIndex        =   202
            Top             =   1920
            Width           =   1320
         End
      End
      Begin VB.Frame Fra 
         BackColor       =   &H00E2E9E9&
         Height          =   5310
         Index           =   4
         Left            =   11385
         TabIndex        =   132
         Top             =   45
         Width           =   10350
         Begin VB.Frame Frame17 
            BackColor       =   &H00E2E9E9&
            Height          =   735
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   165
            Top             =   1440
            Visible         =   0   'False
            Width           =   5775
            Begin VB.TextBox TxtRunTo 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   167
               Top             =   240
               Width           =   1155
            End
            Begin VB.TextBox TxtRunFrom 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   2400
               RightToLeft     =   -1  'True
               TabIndex        =   166
               Top             =   240
               Width           =   1155
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„  «· ‘€Ì·"
               Height          =   195
               Index           =   39
               Left            =   4320
               RightToLeft     =   -1  'True
               TabIndex        =   170
               Top             =   240
               Width           =   1185
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "≈·Ï"
               Height          =   315
               Index           =   43
               Left            =   1560
               RightToLeft     =   -1  'True
               TabIndex        =   169
               Top             =   240
               Width           =   645
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "„‰"
               Height          =   315
               Index           =   44
               Left            =   3600
               RightToLeft     =   -1  'True
               TabIndex        =   168
               Top             =   240
               Width           =   660
            End
         End
         Begin VB.Frame Frame13 
            BackColor       =   &H00E2E9E9&
            Height          =   615
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   159
            Top             =   840
            Width           =   5775
            Begin VB.TextBox TxtConfFrom 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   2400
               RightToLeft     =   -1  'True
               TabIndex        =   161
               Top             =   240
               Width           =   1155
            End
            Begin VB.TextBox TxtConfTo 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   160
               Top             =   240
               Width           =   1155
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "„‰"
               Height          =   315
               Index           =   46
               Left            =   3600
               RightToLeft     =   -1  'True
               TabIndex        =   164
               Top             =   240
               Width           =   660
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "≈·Ï"
               Height          =   315
               Index           =   47
               Left            =   1560
               RightToLeft     =   -1  'True
               TabIndex        =   163
               Top             =   240
               Width           =   645
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„  «· «ﬂÌœ"
               Height          =   195
               Index           =   48
               Left            =   4320
               RightToLeft     =   -1  'True
               TabIndex        =   162
               Top             =   240
               Width           =   1185
            End
         End
         Begin VB.TextBox TxtVichNo 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3240
            RightToLeft     =   -1  'True
            TabIndex        =   155
            Top             =   3720
            Width           =   1155
         End
         Begin VB.TextBox myOutClientCode 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   3405
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   154
            Top             =   2670
            Width           =   1005
         End
         Begin VB.Frame Frame16 
            Height          =   5175
            Left            =   5880
            TabIndex        =   142
            Top             =   120
            Width           =   4455
            Begin VB.Image Image5 
               Height          =   3675
               Left            =   120
               Picture         =   "FrmHajjReports.frx":2C7C
               Stretch         =   -1  'True
               Top             =   120
               Width           =   4395
            End
            Begin VB.Label Label4 
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
               Height          =   1095
               Left            =   960
               TabIndex        =   143
               Top             =   4080
               Width           =   2895
            End
         End
         Begin VB.Frame Frame15 
            BackColor       =   &H00E2E9E9&
            Height          =   1215
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   140
            Top             =   120
            Width           =   5775
            Begin XtremeSuiteControls.RadioButton Rdam2 
               Height          =   255
               Index           =   6
               Left            =   840
               TabIndex        =   141
               Top             =   120
               Width           =   4695
               _Version        =   786432
               _ExtentX        =   8281
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   " ﬁ—Ì— «·»—«„Ã ( «·ÕÃÊ“«  ) «·„‰ ÂÌ… «· ‰›Ì– Ê·„ Ì’œ— ·Â« «„—  ‘€Ì·"
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton Rdam2 
               Height          =   255
               Index           =   7
               Left            =   1800
               TabIndex        =   151
               Top             =   360
               Width           =   3735
               _Version        =   786432
               _ExtentX        =   6588
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   " ﬁ—Ì— «·»—«„Ã ( «·ÕÃÊ“«  ) «·’«œ— ·Â« «„—  ‘€Ì·"
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
         End
         Begin VB.Frame Frame14 
            BackColor       =   &H00E2E9E9&
            Caption         =   "„‰ «·› —Â"
            Height          =   1095
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   133
            Top             =   4080
            Width           =   4575
            Begin MSComCtl2.DTPicker DTPicker1 
               Height          =   330
               Left            =   1800
               TabIndex        =   134
               Top             =   240
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   582
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   94175235
               CurrentDate     =   41640
            End
            Begin MSComCtl2.DTPicker DTPicker2 
               Height          =   330
               Left            =   1800
               TabIndex        =   135
               Top             =   600
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   582
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   94175235
               CurrentDate     =   41640
            End
            Begin Dynamic_Byte.NourHijriCal NourHijriCal1 
               Height          =   315
               Left            =   120
               TabIndex        =   136
               Top             =   240
               Width           =   1635
               _ExtentX        =   2884
               _ExtentY        =   556
            End
            Begin Dynamic_Byte.NourHijriCal NourHijriCal2 
               Height          =   315
               Left            =   120
               TabIndex        =   137
               Top             =   630
               Width           =   1635
               _ExtentX        =   2884
               _ExtentY        =   556
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "≈·Ï"
               Height          =   195
               Index           =   50
               Left            =   3510
               RightToLeft     =   -1  'True
               TabIndex        =   139
               Top             =   720
               Width           =   480
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„‰"
               Height          =   195
               Index           =   49
               Left            =   3450
               RightToLeft     =   -1  'True
               TabIndex        =   138
               Top             =   360
               Width           =   540
            End
         End
         Begin MSDataListLib.DataCombo myDCGroup 
            Height          =   315
            Left            =   120
            TabIndex        =   144
            Top             =   3000
            Width           =   4290
            _ExtentX        =   7567
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin ImpulseButton.ISButton ISButton3 
            Height          =   975
            Left            =   4800
            TabIndex        =   145
            Top             =   4200
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   1720
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
         Begin MSDataListLib.DataCombo myOutClientID3 
            Height          =   315
            Left            =   120
            TabIndex        =   146
            Top             =   2670
            Width           =   3300
            _ExtentX        =   5821
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo myProgrammID3 
            Height          =   315
            Left            =   120
            TabIndex        =   147
            Top             =   3360
            Width           =   4290
            _ExtentX        =   7567
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcbSeasonsID4 
            Height          =   315
            Left            =   120
            TabIndex        =   177
            Top             =   2280
            Width           =   4290
            _ExtentX        =   7567
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„Ê”„"
            Height          =   210
            Index           =   57
            Left            =   4500
            RightToLeft     =   -1  'True
            TabIndex        =   178
            Top             =   2280
            Width           =   1320
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            ForeColor       =   &H00000080&
            Height          =   285
            Index           =   1
            Left            =   240
            TabIndex        =   153
            Top             =   3720
            Width           =   1785
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «·Õ«›·« "
            Height          =   210
            Index           =   45
            Left            =   4560
            RightToLeft     =   -1  'True
            TabIndex        =   152
            Top             =   3720
            Width           =   1200
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·⁄„Ì·"
            Height          =   210
            Index           =   54
            Left            =   4500
            RightToLeft     =   -1  'True
            TabIndex        =   150
            Top             =   2640
            Width           =   1320
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·Õ«›·…"
            Height          =   210
            Index           =   52
            Left            =   4500
            RightToLeft     =   -1  'True
            TabIndex        =   149
            Top             =   3000
            Width           =   1320
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·»—‰«„Ã"
            Height          =   210
            Index           =   51
            Left            =   4500
            RightToLeft     =   -1  'True
            TabIndex        =   148
            Top             =   3360
            Width           =   1320
         End
      End
      Begin VB.Frame Fra 
         BackColor       =   &H00E2E9E9&
         Height          =   5310
         Index           =   3
         Left            =   11085
         TabIndex        =   94
         Top             =   45
         Width           =   10350
         Begin VB.TextBox TxtSearchCode 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   3585
            RightToLeft     =   -1  'True
            TabIndex        =   127
            Top             =   2640
            Width           =   825
         End
         Begin VB.Frame Frame12 
            Height          =   5175
            Left            =   5880
            TabIndex        =   118
            Top             =   120
            Width           =   4455
            Begin VB.Image Image4 
               Height          =   3675
               Left            =   0
               Picture         =   "FrmHajjReports.frx":51D4
               Stretch         =   -1  'True
               Top             =   120
               Width           =   4395
            End
            Begin VB.Label Label3 
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
               Height          =   1095
               Left            =   960
               TabIndex        =   119
               Top             =   4080
               Width           =   2895
            End
         End
         Begin VB.Frame Frame11 
            BackColor       =   &H00E2E9E9&
            Height          =   1215
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   113
            Top             =   0
            Width           =   5895
            Begin XtremeSuiteControls.RadioButton Rdam3 
               Height          =   255
               Index           =   2
               Left            =   1200
               TabIndex        =   114
               Top             =   240
               Width           =   2175
               _Version        =   786432
               _ExtentX        =   3836
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   " ﬁ—Ì— »„Êﬁ⁄ «·Õ«›·… «·Õ«·Ì"
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton Rdam3 
               Height          =   255
               Index           =   3
               Left            =   120
               TabIndex        =   115
               Top             =   600
               Width           =   3255
               _Version        =   786432
               _ExtentX        =   5741
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   " ﬁ—Ì—  Ã„Ì⁄Ì »‰Ê⁄ «·Õ«›·«  Ê«„«ﬂ‰  Ê«ÃœÂ«"
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton Rdam3 
               Height          =   255
               Index           =   1
               Left            =   3720
               TabIndex        =   116
               Top             =   480
               Width           =   2055
               _Version        =   786432
               _ExtentX        =   3625
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   " ﬁ—Ì— Õ—ﬂ… «·Õ«›·« "
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton Rdam3 
               Height          =   255
               Index           =   0
               Left            =   3480
               TabIndex        =   117
               Top             =   240
               Width           =   2295
               _Version        =   786432
               _ExtentX        =   4048
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   " ﬁ—Ì— »Ãœ«Ê· «· —ÕÌ· «·’«œ—…"
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton Rdam3 
               Height          =   255
               Index           =   4
               Left            =   3720
               TabIndex        =   126
               Top             =   840
               Width           =   2055
               _Version        =   786432
               _ExtentX        =   3625
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   " ﬁ—Ì— »Ì«‰«  «·”«∆ﬁÌ‰"
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
         End
         Begin VB.Frame Frame10 
            BackColor       =   &H00E2E9E9&
            Caption         =   "„‰ «·› —Â"
            Height          =   975
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   106
            Top             =   4080
            Width           =   4575
            Begin MSComCtl2.DTPicker DtpDateFrom4 
               Height          =   330
               Left            =   1800
               TabIndex        =   107
               Top             =   150
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   582
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   94175235
               CurrentDate     =   41640
            End
            Begin MSComCtl2.DTPicker DtpDateTo4 
               Height          =   330
               Left            =   1800
               TabIndex        =   108
               Top             =   480
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   582
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   94175235
               CurrentDate     =   41640
            End
            Begin Dynamic_Byte.NourHijriCal DtpDateFromH4 
               Height          =   315
               Left            =   120
               TabIndex        =   109
               Top             =   150
               Width           =   1635
               _ExtentX        =   2884
               _ExtentY        =   556
            End
            Begin Dynamic_Byte.NourHijriCal DtpDateToH4 
               Height          =   315
               Left            =   120
               TabIndex        =   110
               Top             =   510
               Width           =   1635
               _ExtentX        =   2884
               _ExtentY        =   556
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "≈·Ï"
               Height          =   195
               Index           =   38
               Left            =   3510
               RightToLeft     =   -1  'True
               TabIndex        =   112
               Top             =   600
               Width           =   480
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„‰"
               Height          =   195
               Index           =   37
               Left            =   3450
               RightToLeft     =   -1  'True
               TabIndex        =   111
               Top             =   240
               Width           =   540
            End
         End
         Begin VB.Frame Frame9 
            BackColor       =   &H00E2E9E9&
            Height          =   975
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   95
            Top             =   1200
            Width           =   5715
            Begin VB.TextBox TxtToID 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   99
               Top             =   240
               Width           =   1155
            End
            Begin VB.TextBox TxtFromID 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   2340
               RightToLeft     =   -1  'True
               TabIndex        =   98
               Top             =   240
               Width           =   1155
            End
            Begin VB.TextBox TxtToOrderID 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   97
               Top             =   600
               Width           =   1155
            End
            Begin VB.TextBox TxtFromOrderID 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   2340
               RightToLeft     =   -1  'True
               TabIndex        =   96
               Top             =   600
               Width           =   1155
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„  «·ÃœÊ·"
               Height          =   195
               Index           =   36
               Left            =   4320
               RightToLeft     =   -1  'True
               TabIndex        =   105
               Top             =   240
               Width           =   1185
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "≈·Ï"
               Height          =   315
               Index           =   35
               Left            =   1560
               RightToLeft     =   -1  'True
               TabIndex        =   104
               Top             =   240
               Width           =   645
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "„‰"
               Height          =   315
               Index           =   34
               Left            =   3600
               RightToLeft     =   -1  'True
               TabIndex        =   103
               Top             =   240
               Width           =   660
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„  «ﬂÌœ «·ÕÃ“"
               Height          =   195
               Index           =   33
               Left            =   4320
               RightToLeft     =   -1  'True
               TabIndex        =   102
               Top             =   600
               Width           =   1185
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "≈·Ï"
               Height          =   315
               Index           =   32
               Left            =   1560
               RightToLeft     =   -1  'True
               TabIndex        =   101
               Top             =   600
               Width           =   645
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "„‰"
               Height          =   315
               Index           =   31
               Left            =   3600
               RightToLeft     =   -1  'True
               TabIndex        =   100
               Top             =   600
               Width           =   660
            End
         End
         Begin MSDataListLib.DataCombo DCGroup4 
            Height          =   315
            Left            =   120
            TabIndex        =   120
            Top             =   3000
            Width           =   4290
            _ExtentX        =   7567
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcbPath4 
            Height          =   315
            Left            =   120
            TabIndex        =   121
            Top             =   3720
            Width           =   4290
            _ExtentX        =   7567
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin ImpulseButton.ISButton ISButton2 
            Height          =   855
            Left            =   4800
            TabIndex        =   122
            Top             =   4200
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   1508
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
         Begin MSDataListLib.DataCombo DcbDriverID 
            Bindings        =   "FrmHajjReports.frx":772C
            Height          =   315
            Left            =   120
            TabIndex        =   128
            Top             =   2640
            Width           =   3375
            _ExtentX        =   5953
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
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
         Begin MSDataListLib.DataCombo DcbEqupID 
            Bindings        =   "FrmHajjReports.frx":7741
            Height          =   315
            Left            =   120
            TabIndex        =   130
            Top             =   3360
            Width           =   4290
            _ExtentX        =   7567
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
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
         Begin MSDataListLib.DataCombo DcbSeasonsID3 
            Height          =   315
            Left            =   120
            TabIndex        =   175
            Top             =   2280
            Width           =   4290
            _ExtentX        =   7567
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„Ê”„"
            Height          =   210
            Index           =   56
            Left            =   4500
            RightToLeft     =   -1  'True
            TabIndex        =   176
            Top             =   2280
            Width           =   1320
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·Õ«›·…"
            Height          =   210
            Index           =   30
            Left            =   4500
            RightToLeft     =   -1  'True
            TabIndex        =   129
            Top             =   3360
            Width           =   1320
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·”«∆ﬁ"
            Height          =   210
            Index           =   42
            Left            =   4500
            RightToLeft     =   -1  'True
            TabIndex        =   125
            Top             =   2640
            Width           =   1320
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„”«—"
            Height          =   210
            Index           =   41
            Left            =   4500
            RightToLeft     =   -1  'True
            TabIndex        =   124
            Top             =   3720
            Width           =   1320
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·Õ«›·…"
            Height          =   210
            Index           =   40
            Left            =   4500
            RightToLeft     =   -1  'True
            TabIndex        =   123
            Top             =   3000
            Width           =   1320
         End
      End
      Begin VB.Frame Fra 
         BackColor       =   &H00E2E9E9&
         Height          =   5310
         Index           =   2
         Left            =   45
         TabIndex        =   57
         Top             =   45
         Width           =   10350
         Begin VB.Frame lbprocess 
            BackColor       =   &H00E2E9E9&
            Height          =   1335
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   78
            Top             =   1080
            Width           =   5715
            Begin VB.TextBox TxtIDFrom 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   2400
               RightToLeft     =   -1  'True
               TabIndex        =   93
               Top             =   600
               Width           =   1155
            End
            Begin VB.TextBox TxtIDTO 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   92
               Top             =   600
               Width           =   1155
            End
            Begin VB.TextBox TxtFromOrder 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   2400
               RightToLeft     =   -1  'True
               TabIndex        =   91
               Top             =   240
               Width           =   1155
            End
            Begin VB.TextBox TxToOrder 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   90
               Top             =   240
               Width           =   1155
            End
            Begin VB.TextBox TxtCutNo 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   85
               Top             =   960
               Width           =   3435
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„  √ﬂÌœ «·⁄„Ì·"
               Height          =   195
               Index           =   28
               Left            =   4320
               RightToLeft     =   -1  'True
               TabIndex        =   86
               Top             =   960
               Width           =   1185
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "„‰"
               Height          =   315
               Index           =   27
               Left            =   3600
               RightToLeft     =   -1  'True
               TabIndex        =   84
               Top             =   600
               Width           =   660
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "≈·Ï"
               Height          =   315
               Index           =   24
               Left            =   1560
               RightToLeft     =   -1  'True
               TabIndex        =   83
               Top             =   600
               Width           =   645
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ «„— «· ‘€Ì·"
               Height          =   195
               Index           =   23
               Left            =   4320
               RightToLeft     =   -1  'True
               TabIndex        =   82
               Top             =   600
               Width           =   1185
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "„‰"
               Height          =   315
               Index           =   22
               Left            =   3600
               RightToLeft     =   -1  'True
               TabIndex        =   81
               Top             =   240
               Width           =   660
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "≈·Ï"
               Height          =   315
               Index           =   20
               Left            =   1560
               RightToLeft     =   -1  'True
               TabIndex        =   80
               Top             =   240
               Width           =   645
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„  «· «ﬂÌœ"
               Height          =   195
               Index           =   19
               Left            =   4320
               RightToLeft     =   -1  'True
               TabIndex        =   79
               Top             =   240
               Width           =   1185
            End
         End
         Begin VB.Frame Frame8 
            Caption         =   "„‰ «·› —Â"
            Height          =   975
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   64
            Top             =   4320
            Width           =   4575
            Begin MSComCtl2.DTPicker DtpDateFrom3 
               Height          =   330
               Left            =   1800
               TabIndex        =   65
               Top             =   150
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   582
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   94175235
               CurrentDate     =   41640
            End
            Begin MSComCtl2.DTPicker DtpDateTo3 
               Height          =   330
               Left            =   1800
               TabIndex        =   66
               Top             =   480
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   582
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   94175235
               CurrentDate     =   41640
            End
            Begin Dynamic_Byte.NourHijriCal DtpDateFromH3 
               Height          =   315
               Left            =   120
               TabIndex        =   67
               Top             =   120
               Width           =   1635
               _ExtentX        =   2884
               _ExtentY        =   556
            End
            Begin Dynamic_Byte.NourHijriCal DtpDateToH3 
               Height          =   315
               Left            =   120
               TabIndex        =   68
               Top             =   510
               Width           =   1635
               _ExtentX        =   2884
               _ExtentY        =   556
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "„‰"
               Height          =   195
               Index           =   15
               Left            =   3450
               RightToLeft     =   -1  'True
               TabIndex        =   70
               Top             =   240
               Width           =   540
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "≈·Ï"
               Height          =   195
               Index           =   13
               Left            =   3510
               RightToLeft     =   -1  'True
               TabIndex        =   69
               Top             =   600
               Width           =   480
            End
         End
         Begin VB.Frame Frame7 
            BackColor       =   &H00E2E9E9&
            Height          =   1095
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   61
            Top             =   0
            Width           =   5775
            Begin XtremeSuiteControls.RadioButton Rdam2 
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   62
               Top             =   240
               Width           =   3135
               _Version        =   786432
               _ExtentX        =   5530
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   " ﬁ—Ì— √Ê„— «· ‘€»· «·’«œ—… „⁄ ﬁÌ„ Â«"
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton Rdam2 
               Height          =   255
               Index           =   0
               Left            =   3120
               TabIndex        =   63
               Top             =   240
               Width           =   2415
               _Version        =   786432
               _ExtentX        =   4260
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "  ﬁ—Ì— «·Õ—ﬂ… «·ÌÊ„Ì…"
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton Rdam2 
               Height          =   255
               Index           =   5
               Left            =   360
               TabIndex        =   89
               Top             =   720
               Width           =   2895
               _Version        =   786432
               _ExtentX        =   5106
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   " ﬁ—Ì— ÿ·»«  «·„‰ ÂÌ… «· ‰›Ì–"
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton Rdam2 
               Height          =   255
               Index           =   2
               Left            =   3360
               TabIndex        =   156
               Top             =   480
               Width           =   2175
               _Version        =   786432
               _ExtentX        =   3836
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   " ﬁ—Ì— «·„”«—«  «· Ì ·„  ‰›–"
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton Rdam2 
               Height          =   255
               Index           =   3
               Left            =   360
               TabIndex        =   157
               Top             =   480
               Width           =   2895
               _Version        =   786432
               _ExtentX        =   5106
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   " ﬁ—Ì— «·„”«—«  «· Ì  „  ‰›Ì–Â«"
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton Rdam2 
               Height          =   255
               Index           =   4
               Left            =   3480
               TabIndex        =   158
               Top             =   720
               Width           =   2055
               _Version        =   786432
               _ExtentX        =   3625
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   " ﬁ—Ì— »ﬂ· «·„”«—« "
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
         End
         Begin VB.Frame Frame6 
            Height          =   5175
            Left            =   5880
            TabIndex        =   59
            Top             =   120
            Width           =   4455
            Begin VB.Label Label2 
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
               Height          =   1095
               Left            =   960
               TabIndex        =   60
               Top             =   4080
               Width           =   2895
            End
            Begin VB.Image Image3 
               Height          =   3675
               Left            =   120
               Picture         =   "FrmHajjReports.frx":7756
               Stretch         =   -1  'True
               Top             =   120
               Width           =   4395
            End
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   3645
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   58
            Top             =   2910
            Width           =   765
         End
         Begin MSDataListLib.DataCombo DCGroup3 
            Height          =   315
            Left            =   120
            TabIndex        =   71
            Top             =   3240
            Width           =   4290
            _ExtentX        =   7567
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcbPath 
            Height          =   315
            Left            =   120
            TabIndex        =   72
            Top             =   3600
            Width           =   4290
            _ExtentX        =   7567
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin ImpulseButton.ISButton ISButton1 
            Height          =   855
            Left            =   4800
            TabIndex        =   73
            Top             =   4440
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   1508
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
         Begin MSDataListLib.DataCombo OutClientID3 
            Height          =   315
            Left            =   120
            TabIndex        =   74
            Top             =   2910
            Width           =   3420
            _ExtentX        =   6033
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo ProgrammID3 
            Height          =   315
            Left            =   120
            TabIndex        =   87
            Top             =   3960
            Width           =   4290
            _ExtentX        =   7567
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcbSeasonsID2 
            Height          =   315
            Left            =   120
            TabIndex        =   173
            Top             =   2520
            Width           =   4290
            _ExtentX        =   7567
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„Ê”„"
            Height          =   210
            Index           =   55
            Left            =   4500
            RightToLeft     =   -1  'True
            TabIndex        =   174
            Top             =   2520
            Width           =   1320
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·»—‰«„Ã"
            Height          =   210
            Index           =   29
            Left            =   4500
            RightToLeft     =   -1  'True
            TabIndex        =   88
            Top             =   3960
            Width           =   1320
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·Õ«›·…"
            Height          =   210
            Index           =   18
            Left            =   4500
            RightToLeft     =   -1  'True
            TabIndex        =   77
            Top             =   3240
            Width           =   1320
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„”«—"
            Height          =   210
            Index           =   17
            Left            =   4500
            RightToLeft     =   -1  'True
            TabIndex        =   76
            Top             =   3600
            Width           =   1320
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·⁄„Ì·"
            Height          =   210
            Index           =   16
            Left            =   4500
            RightToLeft     =   -1  'True
            TabIndex        =   75
            Top             =   2880
            Width           =   1320
         End
      End
      Begin VB.Frame Fra 
         BackColor       =   &H00E2E9E9&
         Height          =   5310
         Index           =   0
         Left            =   -10995
         TabIndex        =   28
         Top             =   45
         Width           =   10350
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   3645
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   54
            Top             =   2790
            Width           =   765
         End
         Begin VB.Frame Frame5 
            Height          =   5175
            Left            =   5880
            TabIndex        =   51
            Top             =   120
            Width           =   4455
            Begin VB.Image Image2 
               Height          =   3675
               Left            =   0
               Picture         =   "FrmHajjReports.frx":9CAE
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
               Height          =   1095
               Left            =   960
               TabIndex        =   52
               Top             =   4080
               Width           =   2895
            End
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00E2E9E9&
            Caption         =   " «Õ’«∆Ì… »⁄œœ «·œÊ—«  »Õ”»"
            Height          =   1095
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   46
            Top             =   120
            Width           =   5775
            Begin XtremeSuiteControls.RadioButton Rdam 
               Height          =   255
               Index           =   1
               Left            =   2520
               TabIndex        =   47
               Top             =   480
               Width           =   3015
               _Version        =   786432
               _ExtentX        =   5318
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "«·ÕÃÊ“«  «·„ƒﬂœ… Õ”» ‰Ê⁄ «·»—‰«„Ã"
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton Rdam 
               Height          =   255
               Index           =   2
               Left            =   1800
               TabIndex        =   48
               Top             =   720
               Width           =   3735
               _Version        =   786432
               _ExtentX        =   6588
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "«·ÕÃÊ“«  «·„ƒﬂœ… Õ”» ‰Ê⁄ «·»—‰«„Ã Ê‰Ê⁄ «·Õ«›·…"
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton Rdam 
               Height          =   255
               Index           =   3
               Left            =   480
               TabIndex        =   49
               Top             =   240
               Width           =   1215
               _Version        =   786432
               _ExtentX        =   2143
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "«·⁄„Ì·"
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton Rdam 
               Height          =   255
               Index           =   0
               Left            =   3120
               TabIndex        =   50
               Top             =   240
               Width           =   2415
               _Version        =   786432
               _ExtentX        =   4260
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "Õ«·… «·ÕÃ“"
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton Rdam 
               Height          =   255
               Index           =   4
               Left            =   480
               TabIndex        =   131
               Top             =   720
               Width           =   1215
               _Version        =   786432
               _ExtentX        =   2143
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "«·ÕÃÊ“« "
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
         End
         Begin VB.ComboBox DcbStus 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "FrmHajjReports.frx":C206
            Left            =   120
            List            =   "FrmHajjReports.frx":C210
            RightToLeft     =   -1  'True
            TabIndex        =   36
            Top             =   1680
            Width           =   4290
         End
         Begin VB.Frame Frame4 
            Caption         =   "„‰ «·› —Â"
            Height          =   975
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   29
            Top             =   3480
            Width           =   5655
            Begin MSComCtl2.DTPicker DtpDateFrom2 
               Height          =   330
               Left            =   1800
               TabIndex        =   30
               Top             =   150
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   582
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   94175235
               CurrentDate     =   41640
            End
            Begin MSComCtl2.DTPicker DtpDateTo2 
               Height          =   330
               Left            =   1800
               TabIndex        =   31
               Top             =   480
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   582
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   94175235
               CurrentDate     =   41640
            End
            Begin Dynamic_Byte.NourHijriCal DtpDateFromh2 
               Height          =   315
               Left            =   120
               TabIndex        =   32
               Top             =   150
               Width           =   1635
               _ExtentX        =   2884
               _ExtentY        =   556
            End
            Begin Dynamic_Byte.NourHijriCal DtpDateToh2 
               Height          =   315
               Left            =   120
               TabIndex        =   33
               Top             =   510
               Width           =   1635
               _ExtentX        =   2884
               _ExtentY        =   556
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "≈·Ï"
               Height          =   195
               Index           =   6
               Left            =   3630
               RightToLeft     =   -1  'True
               TabIndex        =   35
               Top             =   600
               Width           =   480
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "„‰"
               Height          =   195
               Index           =   5
               Left            =   3570
               RightToLeft     =   -1  'True
               TabIndex        =   34
               Top             =   120
               Width           =   540
            End
         End
         Begin MSDataListLib.DataCombo DCGroup2 
            Height          =   315
            Left            =   120
            TabIndex        =   42
            Top             =   2040
            Width           =   4290
            _ExtentX        =   7567
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo ProgrammID 
            Height          =   315
            Left            =   120
            TabIndex        =   44
            Top             =   2400
            Width           =   4290
            _ExtentX        =   7567
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin ImpulseButton.ISButton Print 
            Height          =   495
            Left            =   2880
            TabIndex        =   53
            Top             =   4680
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   873
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
         Begin MSDataListLib.DataCombo OutClientID 
            Height          =   315
            Left            =   120
            TabIndex        =   55
            Top             =   2790
            Width           =   3420
            _ExtentX        =   6033
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcbSeasonsID 
            Height          =   315
            Left            =   120
            TabIndex        =   172
            Top             =   1320
            Width           =   4290
            _ExtentX        =   7567
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„Ê”„"
            Height          =   210
            Index           =   53
            Left            =   4500
            RightToLeft     =   -1  'True
            TabIndex        =   171
            Top             =   1320
            Width           =   1320
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·⁄„Ì·"
            Height          =   210
            Index           =   26
            Left            =   4500
            RightToLeft     =   -1  'True
            TabIndex        =   56
            Top             =   2760
            Width           =   1320
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·»—‰«„Ã"
            Height          =   210
            Index           =   12
            Left            =   4500
            RightToLeft     =   -1  'True
            TabIndex        =   45
            Top             =   2400
            Width           =   1320
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·Õ«›·…"
            Height          =   210
            Index           =   11
            Left            =   4500
            RightToLeft     =   -1  'True
            TabIndex        =   43
            Top             =   2040
            Width           =   1320
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "Õ«·… «·ÕÃ“"
            Height          =   210
            Index           =   9
            Left            =   4500
            RightToLeft     =   -1  'True
            TabIndex        =   37
            Top             =   1680
            Width           =   1320
         End
      End
      Begin VB.Frame Fra 
         BackColor       =   &H00E2E9E9&
         Height          =   5310
         Index           =   1
         Left            =   -11295
         TabIndex        =   3
         Top             =   45
         Width           =   10350
         Begin VB.Frame Frame3 
            Height          =   4575
            Left            =   5880
            TabIndex        =   13
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
               Height          =   1095
               Left            =   240
               TabIndex        =   14
               Top             =   3840
               Width           =   2895
            End
            Begin VB.Image Image1 
               Height          =   3675
               Left            =   0
               Picture         =   "FrmHajjReports.frx":C220
               Stretch         =   -1  'True
               Top             =   120
               Width           =   4395
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "„‰ «·› —Â"
            Height          =   975
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   6
            Top             =   2880
            Width           =   5535
            Begin MSComCtl2.DTPicker DtpDateFrom 
               Height          =   330
               Left            =   1800
               TabIndex        =   7
               Top             =   150
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   582
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   94175235
               CurrentDate     =   41640
            End
            Begin MSComCtl2.DTPicker DtpDateTo 
               Height          =   330
               Left            =   1800
               TabIndex        =   8
               Top             =   480
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   582
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   94175235
               CurrentDate     =   41640
            End
            Begin Dynamic_Byte.NourHijriCal DtpDateFromH 
               Height          =   315
               Left            =   120
               TabIndex        =   9
               Top             =   150
               Width           =   1635
               _ExtentX        =   2884
               _ExtentY        =   556
            End
            Begin Dynamic_Byte.NourHijriCal DtpDateToH 
               Height          =   315
               Left            =   120
               TabIndex        =   10
               Top             =   510
               Width           =   1635
               _ExtentX        =   2884
               _ExtentY        =   556
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "„‰"
               Height          =   195
               Index           =   4
               Left            =   3570
               RightToLeft     =   -1  'True
               TabIndex        =   12
               Top             =   120
               Width           =   540
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "≈·Ï"
               Height          =   195
               Index           =   3
               Left            =   3630
               RightToLeft     =   -1  'True
               TabIndex        =   11
               Top             =   600
               Width           =   480
            End
         End
         Begin VB.TextBox TxtReceptOffice 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   240
            RightToLeft     =   -1  'True
            TabIndex        =   5
            Top             =   1320
            Width           =   4170
         End
         Begin VB.TextBox TxtOrder 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   240
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   4
            Top             =   240
            Width           =   4170
         End
         Begin MSDataListLib.DataCombo SeasonsID 
            Height          =   315
            Left            =   240
            TabIndex        =   15
            Top             =   600
            Width           =   4170
            _ExtentX        =   7355
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo CompanyID 
            Height          =   315
            Left            =   240
            TabIndex        =   16
            Top             =   960
            Width           =   4170
            _ExtentX        =   7355
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcbDriver 
            Height          =   315
            Left            =   240
            TabIndex        =   17
            Top             =   1680
            Width           =   4170
            _ExtentX        =   7355
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCGroup 
            Height          =   315
            Left            =   240
            TabIndex        =   18
            Top             =   2040
            Width           =   4170
            _ExtentX        =   7355
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcbVehicleType 
            Height          =   315
            Left            =   240
            TabIndex        =   19
            Top             =   2400
            Width           =   4170
            _ExtentX        =   7355
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   495
            Index           =   0
            Left            =   2760
            TabIndex        =   38
            Top             =   4800
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   873
            ButtonPositionImage=   1
            Caption         =   "ﬂ‘›  Ê“Ì⁄ «·«⁄ „«œ"
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
         Begin ImpulseButton.ISButton Cmd 
            Height          =   495
            Index           =   1
            Left            =   4440
            TabIndex        =   39
            Top             =   4800
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   873
            ButtonPositionImage=   1
            Caption         =   "ÿ»«⁄… «· Ê“Ì⁄"
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
            Height          =   735
            Left            =   120
            Top             =   3960
            Width           =   5775
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Caption         =   "Â–… «·‘«‘…  ﬁÊ„ »≈ŸÂ«—  »Ì«‰«  «·ÕÃ Ê«·⁄„—… ÿ»ﬁ« ·· √—ÌŒ"
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
            Height          =   780
            Index           =   25
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   27
            Top             =   3960
            Width           =   5775
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„Ê”„"
            Height          =   330
            Index           =   10
            Left            =   4395
            RightToLeft     =   -1  'True
            TabIndex        =   26
            Top             =   600
            Width           =   1185
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„ﬂ »"
            Height          =   330
            Index           =   21
            Left            =   4395
            RightToLeft     =   -1  'True
            TabIndex        =   25
            Top             =   1320
            Width           =   1185
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„ƒ””…"
            Height          =   330
            Index           =   7
            Left            =   4395
            RightToLeft     =   -1  'True
            TabIndex        =   24
            Top             =   960
            Width           =   1185
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·«⁄ „«œ"
            Height          =   330
            Index           =   8
            Left            =   4395
            RightToLeft     =   -1  'True
            TabIndex        =   23
            Top             =   240
            Width           =   1185
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·”«∆ﬁ"
            Height          =   330
            Index           =   14
            Left            =   4470
            RightToLeft     =   -1  'True
            TabIndex        =   22
            Top             =   1680
            Width           =   1185
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·«—ﬂ«»"
            Height          =   330
            Index           =   0
            Left            =   4440
            RightToLeft     =   -1  'True
            TabIndex        =   21
            Top             =   2400
            Width           =   1185
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·Õ«›·…"
            Height          =   330
            Index           =   2
            Left            =   4440
            RightToLeft     =   -1  'True
            TabIndex        =   20
            Top             =   2040
            Width           =   1185
         End
      End
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   495
      Index           =   2
      Left            =   240
      TabIndex        =   41
      Top             =   6480
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
   Begin VB.Image ImgFavorites 
      Height          =   390
      Left            =   360
      Picture         =   "FrmHajjReports.frx":E778
      Stretch         =   -1  'True
      Top             =   0
      Width           =   525
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "‘«‘…  ﬁ«—Ì—  «·ÕÃ  Ê«·⁄„—…"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   -75
      TabIndex        =   0
      Top             =   0
      Width           =   10410
   End
End
Attribute VB_Name = "FrmHajjReports"
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

Public Function ReloadCombos()
 Dim Dcombos As ClsDataCombos
  Dim str As String
   Set Dcombos = New ClsDataCombos
   Dcombos.GetTblCarsDataGroup DCGroup, 1, True
   Dcombos.GetTblCarsDataGroup DCGroup2, 1, True
   Dcombos.GetTblCarsDataGroup DCGroup3, 1, True
   Dcombos.GetTblCarsDataGroup DCGroup4, 1, True
   Dcombos.GetTblCarsDataGroup myDCGroup, 1, True
   Dcombos.GetCompany OutClientID, 2, 0
   Dcombos.GetCompany OutClientID3, 2, 0
   Dcombos.GetCompany myOutClientID3, 2, 0
   Dcombos.GetTblShrines Me.DcbPath
   Dcombos.GetTblShrines Me.DcbPath4
   ladData
      If SystemOptions.UserInterface = ArabicInterface Then
   str = "select id ,name from TblProgrammTypes "
   Else
   str = "select id ,nameE from TblProgrammTypes "
   End If
   fill_combo ProgrammID, str
   fill_combo ProgrammID3, str
   fill_combo myProgrammID3, str
  If SystemOptions.UserInterface = ArabicInterface Then
   str = " select id , name from TblCompaniesGroup  "
   Else
   str = " select id , nameE from TblCompaniesGroup  "
 End If
 '1
   fill_combo SeasonsID, str
 If SystemOptions.UserInterface = ArabicInterface Then
    str = " select  ID , Name   from TblTourismCompanies "
Else
    str = " Select ID , NameE  from TblTourismCompanies "
End If
fill_combo CompanyID, str
  str = "  select   e.Emp_ID Emp_ID , e.Emp_Name,e.Emp_NameE   from TblEmployee e, TblEmpJobsTypes  j"
  str = str & "   Where e.JobTypeID = j.JobTypeID"
  str = str & "     and  ( j.JobTypeName like '%”«∆ﬁ%'  or j.JobTypeNamee like '%driver%')"
  fill_combo DcbDriver, str
    If SystemOptions.UserInterface = ArabicInterface Then
   str = " select id , name from TblvehicleType  "
   Else
   str = " select id , nameE from TblvehicleType  "
 End If
fill_combo DcbVehicleType, str
  If SystemOptions.UserInterface = ArabicInterface Then
   str = " select id , name from TblCompaniesGroup  "
   Else
   str = " select id , nameE from TblCompaniesGroup  "
  End If
  str = str & " where Omra_Hajj=0"
   fill_combo DcbSeasonsID, str
   fill_combo DcbSeasonsID2, str
   fill_combo DcbSeasonsID3, str
   fill_combo DcbSeasonsID4, str
   
    If SystemOptions.UserInterface = ArabicInterface Then
    With DcbStus
    .Clear
    .AddItem "«·ÃœÌœ ›ﬁÿ"
    .AddItem "«·„ƒﬂœ ›ﬁÿ"
    .AddItem "«·€Ì— „ƒﬂœ ›ﬁÿ"
    .AddItem "«·ﬂ·"
    End With
    Else
   With DcbStus
    .Clear
    .AddItem "New "
    .AddItem "Confirmed Reservation"
    .AddItem "No Confirmed Reservation"
    .AddItem "ALL "
    End With
    End If
   End Function

Private Sub DtpDateFrom2_Change()
If Not IsNull(DtpDateFrom2.value) Then
   DtpDateFromh2.value = ToHijriDate(DtpDateFrom2.value)
   End If
End Sub

Private Sub DtpDateFrom3_Change()
If Not IsNull(DtpDateFrom3.value) Then
   DtpDateFromH3.value = ToHijriDate(DtpDateFrom3.value)
   End If
End Sub

Private Sub DtpDateFrom4_Change()
If Not IsNull(DtpDateFrom4.value) Then
   DtpDateFromH4.value = ToHijriDate(DtpDateFrom4.value)
   End If
End Sub

Private Sub DtpDateFrom5_Change()
If Not IsNull(DtpDateFrom5.value) Then
   DtpDateFromH5.value = ToHijriDate(DtpDateFrom5.value)
   End If
End Sub

Private Sub DtpDateFromh2_LostFocus()
 VBA.Calendar = vbCalGreg
            DtpDateFrom2.value = ToGregorianDate(DtpDateFromh2.value)
End Sub

Private Sub DtpDateFromH3_LostFocus()
 VBA.Calendar = vbCalGreg
            DtpDateFrom3.value = ToGregorianDate(DtpDateFromH3.value)
End Sub

Private Sub DtpDateFromH4_LostFocus()
 VBA.Calendar = vbCalGreg
            DtpDateFrom4.value = ToGregorianDate(DtpDateFromH4.value)
End Sub

Private Sub DtpDateFromH5_LostFocus()
 VBA.Calendar = vbCalGreg
            DtpDateFrom5.value = ToGregorianDate(DtpDateFromH5.value)
End Sub

Private Sub DtpDateTo2_Change()
If Not IsNull(DtpDateTo2.value) Then
   DtpDateToh2.value = ToHijriDate(DtpDateTo2.value)
   End If
End Sub

Private Sub DtpDateTo3_Change()
If Not IsNull(DtpDateTo3.value) Then
   DtpDateToH3.value = ToHijriDate(DtpDateTo3.value)
   End If
End Sub

Private Sub DtpDateTo4_Change()
If Not IsNull(DtpDateTo4.value) Then
   DtpDateToH4.value = ToHijriDate(DtpDateTo4.value)
   End If
End Sub



Private Sub DtpDateTo5_Change()
If Not IsNull(DtpDateTo5.value) Then
   DtpDateToH5.value = ToHijriDate(DtpDateTo5.value)
   End If
End Sub

Private Sub DtpDateToh2_LostFocus()
 VBA.Calendar = vbCalGreg
            DtpDateTo2.value = ToGregorianDate(DtpDateToh2.value)
End Sub



Private Sub DtpDateToH3_LostFocus()
 VBA.Calendar = vbCalGreg
            DtpDateTo3.value = ToGregorianDate(DtpDateToH3.value)
End Sub



Private Sub DtpDateToH4_LostFocus()
 VBA.Calendar = vbCalGreg
            DtpDateTo4.value = ToGregorianDate(DtpDateToH4.value)
End Sub

Private Sub Fill_CombosHajj()
 Dim Dcombos As ClsDataCombos
  Dim str As String
   Set Dcombos = New ClsDataCombos
   Dcombos.GetTblShrines DcbPath5
   Dcombos.GETNationality Nationality5
   Dcombos.GetTypeDependence Me.DcbDepandID5
  If SystemOptions.UserInterface = ArabicInterface Then
   str = " select id , name from TblCompaniesGroup  "
   Else
   str = " select id , nameE from TblCompaniesGroup  "
 End If
 str = str & " where Omra_Hajj=1"
   fill_combo DcbSeasonsID5, str
   
    If SystemOptions.UserInterface = ArabicInterface Then
   str = " select id , name from TblvehicleType  "
   Else
   str = " select id , nameE from TblvehicleType  "
 End If
fill_combo DcbVehicleType5, str
 If SystemOptions.UserInterface = ArabicInterface Then
    str = " select  ID , Name   from TblTourismCompanies "
Else
    str = " Select ID , NameE from  TblTourismCompanies "
End If
fill_combo CompanyID5, str
End Sub

Private Sub DtpDateToH5_LostFocus()
 VBA.Calendar = vbCalGreg
            DtpDateTo5.value = ToGregorianDate(DtpDateToH5.value)
End Sub

Private Sub DTPicker1_Change()
If Not IsNull(DtpDateFrom3.value) Then
   NourHijriCal1.value = ToHijriDate(DTPicker1.value)
   End If
End Sub

Private Sub DTPicker2_Change()
If Not IsNull(DtpDateFrom3.value) Then
   NourHijriCal2.value = ToHijriDate(DTPicker2.value)
   End If
End Sub

Private Sub ImgFavorites_Click()
AddTofaforites Me.Name, Me.Caption, Me.Caption
End Sub
Private Sub btnClear_Click()
clear_all Me
DtpDateFrom.value = ""
DtpDateTo.value = ""
DtpDateFrom2.value = ""
DtpDateTo2.value = ""
DtpDateFrom3.value = ""
DtpDateTo3.value = ""
DtpDateFrom4.value = ""
DtpDateTo4.value = ""
DTPicker1.value = ""
DTPicker2.value = ""
DtpDateTo5.value = ""
DtpDateFrom5.value = ""
End Sub

Private Sub Cmd_Click(Index As Integer)

    Select Case Index

        Case 0
GetData
        Case 1
GetData 1
    
        Case 2
            Unload Me
            Case 3
'print_report
    End Select

End Sub



Private Sub DtpDateFrom_Change()
If Not IsNull(DtpDateFrom.value) Then
   DtpDateFromH.value = ToHijriDate(DtpDateFrom.value)
   End If
End Sub


Private Sub DtpDateFromH_LostFocus()
 VBA.Calendar = vbCalGreg
            DtpDateFrom.value = ToGregorianDate(DtpDateFromH.value)
End Sub



Private Sub DtpDateTo_Change()
If Not IsNull(DtpDateTo.value) Then
   DtpDateToH.value = ToHijriDate(DtpDateTo.value)
   End If
End Sub


Private Sub DtpDateToH_LostFocus()
 VBA.Calendar = vbCalGreg
            DtpDateTo.value = ToGregorianDate(DtpDateToH.value)
End Sub

Private Sub Form_Activate()
   PutFormOnTop Me.hWnd
End Sub
Sub ladData()
     Dim str As String
If SystemOptions.UserInterface = ArabicInterface Then
    str = " SELECT     dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.FlagDriver, dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmpJobsTypes.JobTypeNamee, "
    str = str & "                   dbo.TblEmployee.Emp_Namee ,dbo.TblEmployee.BranchId"
   Else
   str = " SELECT     dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Namee, dbo.TblEmployee.FlagDriver, dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmpJobsTypes.JobTypeNamee, "
    str = str & "                   dbo.TblEmployee.Emp_Name ,dbo.TblEmployee.BranchId "
   End If
    ' If Me.TxtModFlg.Text <> "R" Then
    str = str & "    FROM         dbo.TblEmployee LEFT OUTER JOIN"
    str = str & "                    dbo.TblEmpJobsTypes ON dbo.TblEmployee.JobTypeID = dbo.TblEmpJobsTypes.JobTypeID"
    str = str & "     where  (( JobTypeName like '%”«∆ﬁ%'  or JobTypeNamee like '%driver%' )or (FlagDriver=1)) "
    str = str & "  and  (dbo.TblEmployee.BranchId=0 or dbo.TblEmployee.BranchId is null or         dbo.TblEmployee.BranchId in(" & Current_branchSql & "))"
    fill_combo DcbDriverID, str
     str = "  select   id, OperatorN from TblCarsData WHERE     ( NOT (OperatorN IS NULL)  AND OperatorN <>  '')  "
     str = str & "  and  (TblCarsData.Branch_NO=0 or TblCarsData.Branch_NO is null or    TblCarsData.Branch_NO in(" & Current_branchSql & "))"
    fill_combo DcbEqupID, str
End Sub
Private Sub Form_Load()
   
 'On Error GoTo ErrTrap
    Dim I As Integer
    Dim My_SQL As String
ReloadCombos
Fill_CombosHajj
DtpDateFrom.value = Date
DtpDateTo.value = Date
DtpDateFrom2.value = Date
DtpDateTo2.value = Date
DtpDateFrom3.value = Date
DtpDateTo3.value = Date
DtpDateFrom4.value = Date
DtpDateTo4.value = Date
DTPicker1.value = Date
DTPicker2.value = Date
DtpDateFrom5.value = Date
DtpDateTo5.value = Date

DTPicker1.value = ""
DTPicker2.value = ""
DtpDateFrom.value = ""
DtpDateTo.value = ""
DtpDateFrom2.value = ""
DtpDateTo2.value = ""
DtpDateFrom3.value = ""
DtpDateTo3.value = ""
DtpDateFrom4.value = ""
DtpDateTo4.value = ""
DtpDateFrom5.value = ""
DtpDateTo5.value = ""
    Resize_Form Me
End Sub
Private Sub Form_Unload(Cancel As Integer)

    FormPostion Me, SavePostion
    Set DCboSearch = Nothing
End Sub
Public Sub GetDataDeported(Optional Index As Integer = 0)
    Dim StrSQL As String
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim I As Integer
 If Index = 0 Or Index = 1 Or Index = 2 Then
StrSQL = " SELECT     dbo.TblDeported.ID, dbo.TblDeported.BranchID, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblDeported.RecordDate, "
StrSQL = StrSQL & "                       dbo.TblDeported.RecordDateH, dbo.TblDeported.CurrDate, dbo.TblDeported.CurrDateH, dbo.TblDeported.DriverName, dbo.TblDeported.Remarks,"
StrSQL = StrSQL & "                       dbo.TblDeported.Address, dbo.TblDeported.CurrDate2, dbo.TblDeported.CurrDateH2, dbo.TblDeported.DayName1, dbo.TblDeported.DayName2,"
StrSQL = StrSQL & "                       dbo.TblDeported.Phone, dbo.TblDeported.SuperVisName, dbo.TblDeported.ISArrived, dbo.TblDeported.TypeDrive, dbo.TblDeported.TypeTrip,"
StrSQL = StrSQL & "                       dbo.TblDeported.TimeOut, dbo.TblDeported.TimeIn, dbo.TblDeported.OrderID, dbo.TblDeported.HajzNo, dbo.TblDeported.EqupID, dbo.TblCarsData.Fullcode,"
StrSQL = StrSQL & "                       dbo.TblCarsData.Name, dbo.TblCarsData.BoardNO, dbo.TblCarsData.OperatorN, dbo.TblDeported.LocatioID, TblLocations_1.Name AS LocatinName,"
StrSQL = StrSQL & "                       TblLocations_1.NameE AS LocatinNameE, dbo.TblDeported.LocatioID2, TblLocations_1.Name AS LocatinName2, TblLocations_1.NameE AS LocatinName2E,"
StrSQL = StrSQL & "                       dbo.TblDeported.PathID, dbo.TblShrines.Name AS PathName, dbo.TblShrines.NameE AS PathNameE, dbo.TblDeported.DriverID, dbo.TblEmployee.Emp_Name,"
StrSQL = StrSQL & "                       dbo.TblEmployee.Emp_Name1, dbo.TblEmployee.Emp_Name2, dbo.TblEmployee.Emp_Name3, dbo.TblEmployee.Emp_Name4,"
StrSQL = StrSQL & "                       dbo.TblEmployee.Fullcode AS EmpFullcode, dbo.TblEmployee.Emp_Namee4, dbo.TblEmployee.Emp_Namee3, dbo.TblEmployee.Emp_Namee2,"
StrSQL = StrSQL & "                       dbo.TblEmployee.Emp_Namee1, dbo.TblEmployee.Emp_Namee, dbo.TblDeported.DiffTime, dbo.TblDeported.DiffDate, dbo.TblDeported.TypePath,"
StrSQL = StrSQL & "                       dbo.TblHotels.Name AS TypePathName, dbo.TblHotels.NameE AS TypePathNameE, dbo.TblDeported.NoteSerialOrder, dbo.TblDeported.SeasonsID,"
StrSQL = StrSQL & "                       dbo.TblCompaniesGroup.Name AS SeasonsName, dbo.TblCompaniesGroup.NameE AS SeasonsNameE, dbo.TblDeported.NoteSerial1"
StrSQL = StrSQL & "  FROM         dbo.TblDeported LEFT OUTER JOIN"
StrSQL = StrSQL & "                       dbo.TblCompaniesGroup ON dbo.TblDeported.SeasonsID = dbo.TblCompaniesGroup.ID LEFT OUTER JOIN"
StrSQL = StrSQL & "                       dbo.TblHotels ON dbo.TblDeported.TypePath = dbo.TblHotels.ID LEFT OUTER JOIN"
StrSQL = StrSQL & "                       dbo.TblEmployee ON dbo.TblDeported.DriverID = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
StrSQL = StrSQL & "                       dbo.TblShrines ON dbo.TblDeported.PathID = dbo.TblShrines.ID LEFT OUTER JOIN"
StrSQL = StrSQL & "                       dbo.TblLocations TblLocations_1 ON dbo.TblDeported.LocatioID2 = TblLocations_1.ID LEFT OUTER JOIN"
StrSQL = StrSQL & "                       dbo.TblLocations TblLocations_2 ON dbo.TblDeported.LocatioID = TblLocations_2.ID LEFT OUTER JOIN"
StrSQL = StrSQL & "                       dbo.TblCarsData ON dbo.TblDeported.EqupID = dbo.TblCarsData.id LEFT OUTER JOIN"
StrSQL = StrSQL & "                       dbo.TblBranchesData ON dbo.TblDeported.BranchID = dbo.TblBranchesData.branch_id"
StrSQL = StrSQL & "  where 1=1"
If Index = 2 Then
StrSQL = StrSQL & " and dbo.TblDeported.ID in (SELECT     MAX(ID) AS MaxID"
StrSQL = StrSQL & " From dbo.TblDeported"
StrSQL = StrSQL & " GROUP BY EqupID)"
End If
If val(DcbPath4.BoundText) <> 0 And DcbPath4.Text <> "" Then
StrSQL = StrSQL & "  and dbo.TblDeported.PathID =" & val(DcbPath4.BoundText) & ""
End If
If val(DcbDriverID.BoundText) <> 0 And DcbDriverID.Text <> "" Then
StrSQL = StrSQL & "  and dbo.TblDeported.DriverID =" & val(DcbDriverID.BoundText) & ""
End If
If val(DcbEqupID.BoundText) <> 0 And DcbEqupID.Text <> "" Then
StrSQL = StrSQL & "  and dbo.TblDeported.EqupID =" & val(DcbEqupID.BoundText) & ""
End If
If val(Me.DcbSeasonsID3.BoundText) <> 0 And DcbSeasonsID3.Text <> "" Then
StrSQL = StrSQL & "  and dbo.TblDeported.SeasonsID =" & val(DcbSeasonsID3.BoundText) & ""
End If

If Index = 0 Or Index = 2 Then
If val(TxtFromID.Text) <> 0 Then
StrSQL = StrSQL & "  and dbo.TblDeported.NoteSerial1 >=" & val(TxtFromID.Text) & ""
End If
If val(TxtToID.Text) <> 0 Then
StrSQL = StrSQL & "  and dbo.TblDeported.NoteSerial1 <=" & val(TxtToID.Text) & ""
End If
If val(TxtFromOrderID.Text) <> 0 Then
StrSQL = StrSQL & "  and dbo.TblDeported.NoteSerialOrder >=" & val(TxtFromOrderID.Text) & ""
End If
If val(TxtToOrderID.Text) <> 0 Then
StrSQL = StrSQL & "  and dbo.TblDeported.NoteSerialOrder <=" & val(TxtToOrderID.Text) & ""
End If
End If

End If
 If Not IsNull(Me.DtpDateFrom4.value) Then
                   StrSQL = StrSQL & " AND dbo.TblDeported.RecordDate >=" & SQLDate(Me.DtpDateFrom4.value, True) & ""
   End If
  If Not IsNull(Me.DtpDateTo4.value) Then
                   StrSQL = StrSQL & " AND dbo.TblDeported.RecordDate<=" & SQLDate(Me.DtpDateTo4.value, True) & ""
   End If
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «· ﬁ—Ì—"
  Else
  Msg = "Not Found Data"
  
  End If
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    Else
 rs.MoveFirst
 print_reportAmara StrSQL, Index + 10
    End If
End Sub
Public Sub GetDataDeporTypCars(Optional Index As Integer = 0)
    Dim StrSQL As String
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim I As Integer
 If Index = 3 Then
StrSQL = "SELECT     COUNT(dbo.TblCarsData.CarsTypeId) AS CountCarsTypeId, dbo.TBLCarTypes.name, dbo.TBLCarTypes.namee, dbo.TblDeported.LocatioID2, "
StrSQL = StrSQL & "                      dbo.TblLocations.Name AS LocationName, dbo.TblLocations.NameE AS LocationNameE, dbo.TblDeported.NoteSerialOrder, dbo.TblDeported.SeasonsID,"
StrSQL = StrSQL & "                      dbo.TblCompaniesGroup.Name AS SeasonsName, dbo.TblCompaniesGroup.NameE AS SeasonsNameE"
StrSQL = StrSQL & " FROM         dbo.TblCompaniesGroup RIGHT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblDeported ON dbo.TblCompaniesGroup.ID = dbo.TblDeported.SeasonsID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblLocations ON dbo.TblDeported.LocatioID2 = dbo.TblLocations.ID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TBLCarTypes RIGHT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblCarsData ON dbo.TBLCarTypes.id = dbo.TblCarsData.CarsTypeId ON dbo.TblDeported.EqupID = dbo.TblCarsData.id"
StrSQL = StrSQL & "  Where (dbo.TblDeported.ID in (SELECT     MAX(ID) AS MaxID"
StrSQL = StrSQL & " From dbo.TblDeported"
StrSQL = StrSQL & " GROUP BY EqupID))"
If val(DCGroup4.BoundText) <> 0 And DCGroup4.Text <> "" Then
StrSQL = StrSQL & "  and dbo.TblCarsData.CarsTypeId =" & val(DCGroup4.BoundText) & ""
End If
If val(Me.DcbSeasonsID3.BoundText) <> 0 And DcbSeasonsID3.Text <> "" Then
StrSQL = StrSQL & "  and dbo.TblDeported.SeasonsID =" & val(DcbSeasonsID3.BoundText) & ""
End If
StrSQL = StrSQL & " GROUP BY dbo.TBLCarTypes.name, dbo.TBLCarTypes.namee, dbo.TblDeported.LocatioID2, dbo.TblLocations.Name, dbo.TblLocations.NameE, "
StrSQL = StrSQL & "                       dbo.TblDeported.NoteSerialOrder , dbo.TblDeported.SeasonsID, dbo.TblCompaniesGroup.Name, dbo.TblCompaniesGroup.NameE"
End If

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «· ﬁ—Ì—"
  Else
  Msg = "Not Found Data"
  
  End If
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    Else
 rs.MoveFirst
 print_reportAmara StrSQL, Index + 10
    End If
End Sub
Public Sub GetBaiscDtaData(Optional Index As Integer = 0)
    Dim StrSQL As String
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim I As Integer
 If Index = 4 Then
StrSQL = " SELECT     dbo.TblCarsData.Branch_NO ,dbo.TblCarsData.id, dbo.TblCarsData.Fullcode, dbo.TblCarsData.BoardNO, dbo.TblCarsData.Name, dbo.TblCarsData.Emp_id, dbo.TblEmployee.Emp_Name, "
StrSQL = StrSQL & "                      dbo.TblEmployee.Emp_Name1, dbo.TblEmployee.Emp_Name2, dbo.TblEmployee.Emp_Name3, dbo.TblEmployee.Emp_Name4, dbo.TblEmployee.Nationality,"
StrSQL = StrSQL & "                      dbo.TblEmployee.dean, dbo.TblEmployee.NumEkama, dbo.TblEmployee.NumPoket, dbo.TblEmployee.Fullcode AS EmpFullcode, dbo.TblEmployee.Emp_Namee4,"
StrSQL = StrSQL & "                      dbo.TblEmployee.Emp_Namee3, dbo.TblEmployee.Emp_Namee2, dbo.TblEmployee.Emp_Namee1, dbo.TblEmployee.Emp_Namee, dbo.TblCarsData.OperatorN,"
StrSQL = StrSQL & "                      dbo.TblCarsData.CarsTypeId, dbo.TBLCarTypes.name AS TypeCarname, dbo.TBLCarTypes.namee AS TypeCarnameE"
StrSQL = StrSQL & " FROM         dbo.TblCarsData LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TBLCarTypes ON dbo.TblCarsData.CarsTypeId = dbo.TBLCarTypes.id LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblEmployee ON dbo.TblCarsData.Emp_id = dbo.TblEmployee.Emp_ID"
StrSQL = StrSQL & "  where 1=1"
StrSQL = StrSQL & " and  (dbo.TblCarsData.Branch_NO in(" & Current_branchSql & "))"
StrSQL = StrSQL & " and (NOT (dbo.TblCarsData.Emp_id IS NULL))"
If val(DcbDriverID.BoundText) <> 0 And DcbDriverID.Text <> "" Then
StrSQL = StrSQL & "  and dbo.TblCarsData.Emp_id =" & val(DcbDriverID.BoundText) & ""
End If
If val(DCGroup4.BoundText) <> 0 And DCGroup4.Text <> "" Then
StrSQL = StrSQL & "  and dbo.TblCarsData.CarsTypeId =" & val(DCGroup4.BoundText) & ""
End If
If val(DcbEqupID.BoundText) <> 0 And DcbEqupID.Text <> "" Then
StrSQL = StrSQL & "  and dbo.TblCarsData.ID =" & val(DcbEqupID.BoundText) & ""
End If

End If

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «· ﬁ—Ì—"
  Else
  Msg = "Not Found Data"
  
  End If
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    Else
 rs.MoveFirst
 print_reportAmara StrSQL, Index + 10
    End If
End Sub
Public Sub GetDataAmraORDER(Optional Index As Integer = 0)
    Dim StrSQL As String
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim I As Integer
 If Index = 0 Or Index = 2 Or Index = 3 Or Index = 4 Then
   If Index = 0 Then
    StrSQL = " SELECT     dbo.TblBookingRequest.ID, dbo.TblBookingRequest.SDate, dbo.TblBookingRequest.BranchID, dbo.TblBranchesData.branch_name, "
    StrSQL = StrSQL & "                   dbo.TblBranchesData.branch_namee, dbo.TblBookingRequest.FlightNo, dbo.TblBookingRequest.emp, dbo.TblBookingRequest.other,"
    StrSQL = StrSQL & "                   dbo.TblBookingRequest.EmpName, dbo.TblBookingRequest.EmpCode, dbo.TblBookingRequest.EmpMbile, dbo.TblBookingRequest.ArriveDate,"
    StrSQL = StrSQL & "                   dbo.TblBookingRequest.ArriveTime, dbo.TblBookingRequest.VehicleNo, dbo.TblBookingRequest.GroupName, dbo.TblBookingRequest.ApproveTime,"
    StrSQL = StrSQL & "                   dbo.TblBookingRequest.ApproveDate, dbo.TblBookingRequest.ApproveFlag, dbo.TblBookingRequest.ReservNo, dbo.TblBookingRequest.RemarkApprove,"
    StrSQL = StrSQL & "                   dbo.TblBookingRequest.HotelMakh, dbo.TblBookingRequest.HotelMadinh, dbo.TblBookingRequest.HotelJaddah, dbo.TblBookingRequest.CusNo,"
    StrSQL = StrSQL & "                   dbo.TblBookingRequest.CompnyIn, dbo.TblBookingRequest.CompnyOut, dbo.TblBookingRequest.OutClientID, dbo.TblCustemers.CusName,"
    StrSQL = StrSQL & "                   dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode, dbo.TblBookingRequest.VehicleType, dbo.TBLCarTypes.name, dbo.TBLCarTypes.namee,"
    StrSQL = StrSQL & "                   dbo.TblFlightDetails.[Date], dbo.TblFlightDetails.[Time], dbo.TblFlightDetails.Remarks, dbo.TblFlightDetails.PathID, dbo.TblShrines.Name AS PathName,"
    StrSQL = StrSQL & "                   dbo.TblShrines.NameE AS PathNameE, dbo.TblBookingRequest.NoteSerial1, dbo.TblBookingRequest.SeasonsID, dbo.TblCompaniesGroup.Name AS SeasonsName,"
    StrSQL = StrSQL & "                   dbo.TblCompaniesGroup.NameE AS SeasonsNameE, dbo.TblBookingRequest.Prefix"
    StrSQL = StrSQL & "        FROM         dbo.TblCompaniesGroup RIGHT OUTER JOIN"
    StrSQL = StrSQL & "                   dbo.TblBookingRequest ON dbo.TblCompaniesGroup.ID = dbo.TblBookingRequest.SeasonsID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                   dbo.TblShrines RIGHT OUTER JOIN"
    StrSQL = StrSQL & "                   dbo.TblFlightDetails ON dbo.TblShrines.ID = dbo.TblFlightDetails.PathID ON dbo.TblBookingRequest.ID = dbo.TblFlightDetails.HID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                   dbo.TBLCarTypes ON dbo.TblBookingRequest.VehicleType = dbo.TBLCarTypes.id LEFT OUTER JOIN"
    StrSQL = StrSQL & "                   dbo.TblCustemers ON dbo.TblBookingRequest.OutClientID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                   dbo.TblBranchesData ON dbo.TblBookingRequest.BranchID = dbo.TblBranchesData.ActivityTypeId"
    StrSQL = StrSQL & "  Where (dbo.TblBookingRequest.StusID = 1)"

   ElseIf Index = 2 Or Index = 3 Or Index = 4 Then
   StrSQL = " SELECT     dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode, "
   StrSQL = StrSQL & "                   dbo.TblBookingRequest.CusNo, dbo.TblShrines.Name AS PathName, dbo.TblShrines.NameE AS PathNameE, dbo.TBLCarTypes.name, dbo.TBLCarTypes.namee,"
   StrSQL = StrSQL & "                   dbo.TblBookingRequest.OutClientID, dbo.TblBookingRequest.VehicleType, dbo.TblFlightDetails.Remarks, dbo.TblFlightDetails.[Date], dbo.TblFlightDetails.[Time],"
   StrSQL = StrSQL & "                   dbo.TblFlightDetails.PathID, dbo.TblBookingRequest.BranchID, dbo.TblBookingRequest.SDate, dbo.TblBookingRequest.ID, dbo.TblBookingRequest.VehicleNo,"
   StrSQL = StrSQL & "                   dbo.TblBookingRequest.NoteSerial1, dbo.TblBookingRequest.Prefix, dbo.TblBookingRequest.SeasonsID, dbo.TblCompaniesGroup.Name AS SeasonsName,"
   StrSQL = StrSQL & "                   dbo.TblCompaniesGroup.NameE AS SeasonsNameE"
   StrSQL = StrSQL & "     FROM         dbo.TBLCarTypes RIGHT OUTER JOIN"
   StrSQL = StrSQL & "                   dbo.TblCompaniesGroup RIGHT OUTER JOIN"
   StrSQL = StrSQL & "                   dbo.TblBookingRequest ON dbo.TblCompaniesGroup.ID = dbo.TblBookingRequest.SeasonsID LEFT OUTER JOIN"
   StrSQL = StrSQL & "                   dbo.TblFlightDetails LEFT OUTER JOIN"
   StrSQL = StrSQL & "                   dbo.TblShrines ON dbo.TblFlightDetails.PathID = dbo.TblShrines.ID ON dbo.TblBookingRequest.ID = dbo.TblFlightDetails.HID ON"
   StrSQL = StrSQL & "                   dbo.TBLCarTypes.id = dbo.TblBookingRequest.VehicleType LEFT OUTER JOIN"
   StrSQL = StrSQL & "                   dbo.TblCustemers ON dbo.TblBookingRequest.OutClientID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
   StrSQL = StrSQL & "                   dbo.TblBranchesData ON dbo.TblBookingRequest.BranchID = dbo.TblBranchesData.branch_id"
   StrSQL = StrSQL & "  Where (dbo.TblBookingRequest.StusID = 1)"
  End If
If Index = 2 Then
StrSQL = StrSQL & " and  (dbo.tblbookingrequest.VehicleNo - dbo.GetNoPath(dbo.tblbookingrequest.ID, dbo.TblFlightDetails.PathID) > 0)"
End If
If Index = 3 Then
StrSQL = StrSQL & " and  (dbo.tblbookingrequest.VehicleNo - dbo.GetNoPath(dbo.tblbookingrequest.ID, dbo.TblFlightDetails.PathID) = 0)"
End If

If val(DcbPath.BoundText) <> 0 And DcbPath.Text <> "" Then
StrSQL = StrSQL & "  and dbo.TblFlightDetails.PathID =" & val(DcbPath.BoundText) & ""
End If
If val(OutClientID3.BoundText) <> 0 And OutClientID3.Text <> "" Then
StrSQL = StrSQL & "  and dbo.TblBookingRequest.OutClientID =" & val(OutClientID3.BoundText) & ""
End If
If val(DCGroup3.BoundText) <> 0 And DCGroup3.Text <> "" Then
StrSQL = StrSQL & "  and dbo.TblBookingRequest.VehicleType =" & val(DCGroup3.BoundText) & ""
End If
If val(TxtFromOrder.Text) <> 0 Then
StrSQL = StrSQL & "  and dbo.TblBookingRequest.NoteSerial1 >=" & val(TxtFromOrder.Text) & ""
End If
If val(TxToOrder.Text) <> 0 Then
StrSQL = StrSQL & "  and dbo.TblBookingRequest.NoteSerial1 <=" & val(TxToOrder.Text) & ""
End If
If TxtCutNo.Text <> "" Then
StrSQL = StrSQL & "  and dbo.TblBookingRequest.CusNo Like N'%" & TxtCutNo.Text & "%'"
End If
If Me.DcbSeasonsID2.Text <> "" And val(DcbSeasonsID2.BoundText) <> 0 Then
StrSQL = StrSQL & "  and dbo.TblBookingRequest.SeasonsID =" & val(DcbSeasonsID2.BoundText) & ""
End If

 If Not IsNull(Me.DtpDateFrom3.value) Then
                   StrSQL = StrSQL & " AND dbo.TblFlightDetails.Date >=" & SQLDate(Me.DtpDateFrom3.value, True) & ""
   End If
  If Not IsNull(Me.DtpDateTo3.value) Then
                   StrSQL = StrSQL & " AND dbo.TblFlightDetails.Date<=" & SQLDate(Me.DtpDateTo3.value, True) & ""
   End If

  'End If
ElseIf Index = 1 Then

StrSQL = " SELECT                     dbo.tblbookingrequest2.NoteSerial ,    dbo.tblbookingrequest2.ID, dbo.tblbookingrequest2.SDate, dbo.tblbookingrequest2.BranchID, dbo.TblBranchesData.branch_name, "
StrSQL = StrSQL & "                      dbo.TblBranchesData.branch_namee, dbo.tblbookingrequest2.OrdeNo, dbo.tblbookingrequest2.OutClientID, dbo.TblCustemers.CusName,"
StrSQL = StrSQL & "                      dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode, dbo.tblbookingrequest2.VehicleNo, dbo.tblbookingrequest2.VehicleType, dbo.TBLCarTypes.name,"
StrSQL = StrSQL & "                      dbo.TBLCarTypes.namee, dbo.tblbookingrequest2.TypeDiscount, dbo.tblbookingrequest2.ProAdd, dbo.tblbookingrequest2.NetDis, dbo.tblbookingrequest2.Discount,"
StrSQL = StrSQL & "                      dbo.tblbookingrequest2.Total, dbo.tblbookingrequest2.PathAddValue AS HPathAddValue, dbo.tblbookingrequest2.ProgValue,"
StrSQL = StrSQL & "                      dbo.tblbookingrequest2.Remarks AS HRemarks, dbo.tblbookingrequest2.Mobile2, dbo.tblbookingrequest2.ProgrammID, dbo.TblProgrammTypes.Name AS PrgName,"
StrSQL = StrSQL & "                      dbo.TblProgrammTypes.NameE AS PrgNameE, dbo.TblBookingRequest.CusNo, dbo.tblbookingrequest2.NoteSerialOrder, dbo.tblbookingrequest2.NoteSerial1,"
StrSQL = StrSQL & "                      dbo.tblbookingrequest2.SeasonsID, dbo.TblCompaniesGroup.Name AS SeasonsName, dbo.TblCompaniesGroup.NameE AS SeasonsNameE,"
StrSQL = StrSQL & "                      dbo.tblbookingrequest2.CusNo AS Expr1, dbo.tblbookingrequest2.TotalValue, dbo.tblbookingrequest2.FATValue, dbo.tblbookingrequest2.FATYou"
StrSQL = StrSQL & " FROM         dbo.TblCompaniesGroup RIGHT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.tblbookingrequest2 ON dbo.TblCompaniesGroup.ID = dbo.tblbookingrequest2.SeasonsID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblBookingRequest ON dbo.tblbookingrequest2.OrdeNo = dbo.TblBookingRequest.ID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblProgrammTypes ON dbo.tblbookingrequest2.ProgrammID = dbo.TblProgrammTypes.ID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TBLCarTypes ON dbo.tblbookingrequest2.VehicleType = dbo.TBLCarTypes.id LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblCustemers ON dbo.tblbookingrequest2.OutClientID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblBranchesData ON dbo.tblbookingrequest2.BranchID = dbo.TblBranchesData.branch_id"
StrSQL = StrSQL & "  Where (1= 1)"


ElseIf Index = 5 Then
StrSQL = " SELECT     dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode, "
StrSQL = StrSQL & "                       dbo.TblBookingRequest.CusNo, dbo.TBLCarTypes.name, dbo.TBLCarTypes.namee, dbo.TblCompaniesGroup.Name AS SeasonsName,"
StrSQL = StrSQL & "                       dbo.TblCompaniesGroup.NameE AS SeasonsNameE, dbo.TblBookingRequest.ID, dbo.TblBookingRequest.SDate, dbo.TblBookingRequest.BranchID,"
StrSQL = StrSQL & "                       dbo.TblBookingRequest.OutClientID, dbo.TblBookingRequest.VehicleNo, dbo.TblBookingRequest.VehicleType, dbo.TblBookingRequest.NoteSerial1,"
StrSQL = StrSQL & "                       dbo.GetNoAllPath(dbo.TblBookingRequest.ID) AS Allpath, dbo.GetNoPathInReq(dbo.TblBookingRequest.ID) AS PathInOrder, dbo.TblBookingRequest.SeasonsID,"
StrSQL = StrSQL & "                       dbo.TblBookingRequest.ProgrammID, dbo.TblProgrammTypes.Name AS ProjramName, dbo.TblProgrammTypes.NameE AS ProjramNameE"
StrSQL = StrSQL & "  FROM         dbo.TblCustemers RIGHT OUTER JOIN"
StrSQL = StrSQL & "                       dbo.TBLCarTypes RIGHT OUTER JOIN"
StrSQL = StrSQL & "                       dbo.TblProgrammTypes RIGHT OUTER JOIN"
StrSQL = StrSQL & "                       dbo.TblBookingRequest ON dbo.TblProgrammTypes.ID = dbo.TblBookingRequest.ProgrammID LEFT OUTER JOIN"
StrSQL = StrSQL & "                       dbo.TblCompaniesGroup ON dbo.TblBookingRequest.SeasonsID = dbo.TblCompaniesGroup.ID ON dbo.TBLCarTypes.id = dbo.TblBookingRequest.VehicleType ON"
StrSQL = StrSQL & "                       dbo.TblCustemers.CusID = dbo.TblBookingRequest.OutClientID LEFT OUTER JOIN"
StrSQL = StrSQL & "                       dbo.TblBranchesData ON dbo.TblBookingRequest.BranchID = dbo.TblBranchesData.branch_id"
StrSQL = StrSQL & "  Where (1= 1)"
StrSQL = StrSQL & " and  (dbo.GetNoPathInReq(dbo.tblbookingrequest.ID) * dbo.tblbookingrequest.VehicleNo - dbo.GetNoAllPath(dbo.tblbookingrequest.ID) = 0)"
If Me.DcbSeasonsID2.Text <> "" And val(DcbSeasonsID2.BoundText) <> 0 Then
StrSQL = StrSQL & "  and dbo.TblBookingRequest.SeasonsID =" & val(DcbSeasonsID2.BoundText) & ""
End If
If val(OutClientID3.BoundText) <> 0 And OutClientID3.Text <> "" Then
StrSQL = StrSQL & "  and dbo.tblbookingrequest.OutClientID =" & val(OutClientID3.BoundText) & ""
End If
If val(ProgrammID3.BoundText) <> 0 And ProgrammID3.Text <> "" And Index <> 5 Then
StrSQL = StrSQL & "  and dbo.tblbookingrequest.ProgrammID =" & val(ProgrammID3.BoundText) & ""
End If
End If

If Index = 1 Then
If Me.DcbSeasonsID2.Text <> "" And val(DcbSeasonsID2.BoundText) <> 0 Then
StrSQL = StrSQL & "  and dbo.TblBookingRequest2.SeasonsID =" & val(DcbSeasonsID2.BoundText) & ""
End If
If val(DcbPath.BoundText) <> 0 And DcbPath.Text <> "" And Index <> 5 And Index <> 1 And Index <> 0 Then
StrSQL = StrSQL & "  and dbo.TblFlightDetails2.PathID =" & val(DcbPath.BoundText) & ""
End If
If val(OutClientID3.BoundText) <> 0 And OutClientID3.Text <> "" Then
StrSQL = StrSQL & "  and dbo.tblbookingrequest2.OutClientID =" & val(OutClientID3.BoundText) & ""
End If
If val(ProgrammID3.BoundText) <> 0 And ProgrammID3.Text <> "" And Index <> 5 Then
StrSQL = StrSQL & "  and dbo.tblbookingrequest2.ProgrammID =" & val(ProgrammID3.BoundText) & ""
End If
If val(DCGroup3.BoundText) <> 0 And DCGroup3.Text <> "" And Index <> 5 Then
StrSQL = StrSQL & "  and dbo.tblbookingrequest2.VehicleType =" & val(DCGroup3.BoundText) & ""
End If
If val(TxtFromOrder.Text) <> 0 Then
StrSQL = StrSQL & "  and dbo.tblbookingrequest2.OrdeNo >=" & val(TxtFromOrder.Text) & ""
End If
If val(TxToOrder.Text) <> 0 Then
StrSQL = StrSQL & "  and dbo.tblbookingrequest2.OrdeNo <=" & val(TxToOrder.Text) & ""
End If
If val(TxtIDFrom.Text) <> 0 Then
StrSQL = StrSQL & "  and dbo.tblbookingrequest2.NoteSerial1 >=" & val(TxtIDFrom.Text) & ""
End If
If val(TxtIDTO.Text) <> 0 Then
StrSQL = StrSQL & "  and dbo.tblbookingrequest2.NoteSerial1 <=" & val(TxtIDTO.Text) & ""
End If
If TxtCutNo.Text <> "" Then
StrSQL = StrSQL & "  and dbo.TblBookingRequest.CusNo Like N'%" & TxtCutNo.Text & "%'"
End If
 If Not IsNull(Me.DtpDateFrom3.value) Then
                   StrSQL = StrSQL & " AND dbo.tblbookingrequest2.SDate >=" & SQLDate(Me.DtpDateFrom3.value, True) & ""
   End If
  If Not IsNull(Me.DtpDateTo3.value) Then
                   StrSQL = StrSQL & " AND dbo.tblbookingrequest2.SDate<=" & SQLDate(Me.DtpDateTo3.value, True) & ""
   End If
Else
''//////////////
If val(DCGroup3.BoundText) <> 0 And DCGroup3.Text <> "" And Index <> 5 Then
StrSQL = StrSQL & "  and dbo.tblbookingrequest.VehicleType =" & val(DCGroup3.BoundText) & ""
End If

If val(TxtIDFrom.Text) <> 0 Then
StrSQL = StrSQL & "  and dbo.tblbookingrequest.NoteSerial1 >=" & val(TxtIDFrom.Text) & ""
End If
If val(TxtIDTO.Text) <> 0 Then
StrSQL = StrSQL & "  and dbo.tblbookingrequest.NoteSerial1 <=" & val(TxtIDTO.Text) & ""
End If
If TxtCutNo.Text <> "" Then
StrSQL = StrSQL & "  and dbo.TblBookingRequest.CusNo Like N'%" & TxtCutNo.Text & "%'"
End If
If Index <> 0 Then
 If Not IsNull(Me.DtpDateFrom3.value) Then
                   StrSQL = StrSQL & " AND dbo.tblbookingrequest.SDate >=" & SQLDate(Me.DtpDateFrom3.value, True) & ""
   End If
  If Not IsNull(Me.DtpDateTo3.value) Then
                   StrSQL = StrSQL & " AND dbo.tblbookingrequest.SDate<=" & SQLDate(Me.DtpDateTo3.value, True) & ""
   End If
 End If
End If

'''''''''

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «· ﬁ—Ì—"
  Else
  Msg = "Not Found Data"
  
  End If
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    Else
 rs.MoveFirst
 print_reportAmara StrSQL, Index + 4
    End If
End Sub
Public Sub GetDataHajjORDER1()
    Dim StrSQL As String
    Dim rs As ADODB.Recordset
    Dim Msg As String
    StrSQL = " SELECT     dbo.TblEndorseTrans.ID, dbo.TblEndorseTrans.SDate, dbo.TblEndorseTrans.SeasonsID, dbo.TblCompaniesGroup.Name, dbo.TblCompaniesGroup.NameE, "
    StrSQL = StrSQL & "                   dbo.TblEndorseTrans.Nationality, dbo.Nationality.name AS NationalityName, dbo.Nationality.namee AS NationalityNameE, dbo.TblEndorseTrans.PathID,"
    StrSQL = StrSQL & "                   dbo.TblShrines.Name AS PathName, dbo.TblShrines.NameE AS PathNameE, dbo.TblEndorseTrans.VehicleType, dbo.TblVehicleType.Name AS VehicleTypeName,"
    StrSQL = StrSQL & "                   dbo.TblVehicleType.NameE AS VehicleTypeNameE, dbo.TblEndorseTrans.CompanyID, dbo.TblTourismCompanies.Name AS CompanyName,"
    StrSQL = StrSQL & "                   dbo.TblTourismCompanies.NameE AS CompanyNameE, dbo.TblEndorseTrans.DepandID, dbo.TblTypeDependence.Name AS DepandName,"
    StrSQL = StrSQL & "                   dbo.TblTypeDependence.NameE AS DepandNameE, dbo.TblEndorseTrans.FlagDepand, dbo.TblEndorseTrans.TypeCont, dbo.TblEndorseTrans.RecordDateH,"
    StrSQL = StrSQL & "                   dbo.TblEndorseTrans.ReceptTime, dbo.TblEndorseTrans.Phone, dbo.TblEndorseTrans.NoVehicle, dbo.TblEndorseTrans.Capacity, dbo.TblEndorseTrans.TotalPrice,"
    StrSQL = StrSQL & "                   dbo.TblEndorseTrans.SmalPrice, dbo.TblEndorseTrans.LargPrice, dbo.TblEndorseTrans.Password, dbo.TblEndorseTrans.[Session],"
    StrSQL = StrSQL & "                   dbo.TblEndorseTrans.ReceptOffice, dbo.TblEndorseTrans.ReceptName, dbo.TblEndorseTrans.GroupName, dbo.TblEndorseTrans.CreationDate,"
    StrSQL = StrSQL & "                   dbo.TblEndorseTrans.Remark , dbo.TblEndorseTrans.Total, dbo.TblEndorseTrans.TotYoungs, dbo.TblEndorseTrans.TotOlds, dbo.TblEndorseTrans.ApproveID"
    StrSQL = StrSQL & "       FROM         dbo.TblEndorseTrans LEFT OUTER JOIN"
    StrSQL = StrSQL & "                   dbo.TblTypeDependence ON dbo.TblEndorseTrans.DepandID = dbo.TblTypeDependence.ID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                   dbo.TblTourismCompanies ON dbo.TblEndorseTrans.CompanyID = dbo.TblTourismCompanies.ID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                   dbo.TblVehicleType ON dbo.TblEndorseTrans.VehicleType = dbo.TblVehicleType.ID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                   dbo.TblShrines ON dbo.TblEndorseTrans.PathID = dbo.TblShrines.ID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                   dbo.Nationality ON dbo.TblEndorseTrans.Nationality = dbo.Nationality.id LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.TblCompaniesGroup ON dbo.TblEndorseTrans.SeasonsID = dbo.TblCompaniesGroup.ID"
    StrSQL = StrSQL & "  Where (1 = 1)"
      If RdHajEtmad(0).value = True Then
      StrSQL = StrSQL & "  and dbo.TblEndorseTrans.FlagDepand  =1"
      End If
      If RdHajEtmad(1).value = True Then
      StrSQL = StrSQL & "  and ( dbo.TblEndorseTrans.FlagDepand  is null or dbo.TblEndorseTrans.FlagDepand =0) "
      End If
        If val(Text5.Text) <> 0 Then
     StrSQL = StrSQL & "  and dbo.TblEndorseTrans.ID >=" & val(Text5.Text) & ""
     End If

  If val(Text5.Text) <> 0 Then
     StrSQL = StrSQL & "  and dbo.TblEndorseTrans.ID >=" & val(Text5.Text) & ""
  End If
  If val(Text6.Text) <> 0 Then
     StrSQL = StrSQL & "  and dbo.TblEndorseTrans.ID <=" & val(Text6.Text) & ""
  End If
  If val(DcbSeasonsID5.BoundText) <> 0 And DcbSeasonsID5.Text <> "" Then
     StrSQL = StrSQL & "  and dbo.TblEndorseTrans.SeasonsID =" & val(DcbSeasonsID5.BoundText) & ""
  End If
  If val(CompanyID5.BoundText) <> 0 And CompanyID5.Text <> "" Then
     StrSQL = StrSQL & "  and dbo.TblEndorseTrans.CompanyID =" & val(CompanyID5.BoundText) & ""
  End If
  If val(DcbVehicleType5.BoundText) <> 0 And DcbVehicleType5.Text <> "" Then
     StrSQL = StrSQL & "  and dbo.TblEndorseTrans.VehicleType =" & val(DcbVehicleType5.BoundText) & ""
  End If
  If val(DcbPath5.BoundText) <> 0 And DcbPath5.Text <> "" Then
     StrSQL = StrSQL & "  and dbo.TblEndorseTrans.PathID =" & val(DcbPath5.BoundText) & ""
  End If
   If val(DcbDepandID5.BoundText) <> 0 And DcbDepandID5.Text <> "" Then
     StrSQL = StrSQL & "  and dbo.TblEndorseTrans.DepandID =" & val(DcbDepandID5.BoundText) & ""
   End If
    If val(Nationality5.BoundText) <> 0 And Nationality5.Text <> "" Then
     StrSQL = StrSQL & "  and dbo.TblEndorseTrans.Nationality =" & val(Nationality5.BoundText) & ""
   End If
   If Not IsNull(Me.DtpDateFrom5.value) Then
                   StrSQL = StrSQL & " AND dbo.TblEndorseTrans.SDate >=" & SQLDate(Me.DtpDateFrom5.value, True) & ""
   End If
   If Not IsNull(Me.DtpDateTo5.value) Then
                   StrSQL = StrSQL & " AND dbo.TblEndorseTrans.SDate <=" & SQLDate(Me.DtpDateTo5.value, True) & ""
   End If
  'End If



'''''''''

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «· ﬁ—Ì—"
  Else
  Msg = "Not Found Data"
  
  End If
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    Else
 rs.MoveFirst
 print_reportHajj StrSQL, 0
    End If
End Sub
Public Sub GetDataHajjORDER2()
    Dim StrSQL As String
    Dim rs As ADODB.Recordset
    Dim Msg As String
    StrSQL = " SELECT     dbo.TblEndorseTransMashar.ID, dbo.TblEndorseTransMashar.ApproveID, dbo.TblEndorseTransMashar.SeasonsID, dbo.TblCompaniesGroup.Name, "
    StrSQL = StrSQL & "                  dbo.TblCompaniesGroup.NameE, dbo.TblEndorseTransMashar.Nationality, dbo.Nationality.name AS NationalityName, dbo.Nationality.namee AS NationalityNameE,"
    StrSQL = StrSQL & "                  dbo.TblEndorseTransMashar.VehicleType, dbo.TblVehicleType.Name AS VehicleTypeName, dbo.TblVehicleType.NameE AS VehicleTypeNameE,"
    StrSQL = StrSQL & "                  dbo.TblEndorseTransMashar.DepandID, dbo.TblTypeDependence.Name AS DepandName, dbo.TblTypeDependence.NameE AS DepandNameE,"
    StrSQL = StrSQL & "                  dbo.TblEndorseTransMashar.PathID, dbo.TblShrines.Name AS PathName, dbo.TblShrines.NameE AS PathNameE, dbo.TblEndorseTransMashar.CompanyID,"
    StrSQL = StrSQL & "                  dbo.TblTourismCompanies.Name AS CompanyName, dbo.TblTourismCompanies.NameE AS CompanyNameE, dbo.TblEndorseTransMashar.SDate,"
    StrSQL = StrSQL & "                  dbo.TblEndorseTransMashar.FlagDepand, dbo.TblEndorseTransMashar.RecordDateH, dbo.TblEndorseTransMashar.ReceptTime, dbo.TblEndorseTransMashar.Phone,"
    StrSQL = StrSQL & "                  dbo.TblEndorseTransMashar.NoVehicle, dbo.TblEndorseTransMashar.Capacity, dbo.TblEndorseTransMashar.TotalPrice, dbo.TblEndorseTransMashar.SmalPrice,"
    StrSQL = StrSQL & "                  dbo.TblEndorseTransMashar.LargPrice, dbo.TblEndorseTransMashar.[Session], dbo.TblEndorseTransMashar.Password, dbo.TblEndorseTransMashar.PArCode,"
    StrSQL = StrSQL & "                  dbo.TblEndorseTransMashar.ReceptOffice, dbo.TblEndorseTransMashar.ReceptID, dbo.TblEndorseTransMashar.ReceptName,"
    StrSQL = StrSQL & "                  dbo.TblEndorseTransMashar.GroupName, dbo.TblEndorseTransMashar.CreationDate, dbo.TblEndorseTransMashar.Remark, dbo.TblEndorseTransMashar.Total,"
    StrSQL = StrSQL & "                  dbo.TblEndorseTransMashar.TotYoungs , dbo.TblEndorseTransMashar.TotOlds"
    StrSQL = StrSQL & "   FROM         dbo.TblEndorseTransMashar LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.TblTourismCompanies ON dbo.TblEndorseTransMashar.CompanyID = dbo.TblTourismCompanies.ID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.TblShrines ON dbo.TblEndorseTransMashar.PathID = dbo.TblShrines.ID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.TblTypeDependence ON dbo.TblEndorseTransMashar.DepandID = dbo.TblTypeDependence.ID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.TblVehicleType ON dbo.TblEndorseTransMashar.VehicleType = dbo.TblVehicleType.ID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.Nationality ON dbo.TblEndorseTransMashar.Nationality = dbo.Nationality.id LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.TblCompaniesGroup ON dbo.TblEndorseTransMashar.SeasonsID = dbo.TblCompaniesGroup.ID"
    StrSQL = StrSQL & "  Where (1 = 1)"
      If RdHajEtmad(0).value = True Then
      StrSQL = StrSQL & "  and dbo.TblEndorseTransMashar.FlagDepand  =1"
      End If
      If RdHajEtmad(1).value = True Then
      StrSQL = StrSQL & "  and ( dbo.TblEndorseTransMashar.FlagDepand  is null or dbo.TblEndorseTransMashar.FlagDepand =0) "
      End If
        If val(Text5.Text) <> 0 Then
     StrSQL = StrSQL & "  and dbo.TblEndorseTransMashar.ID >=" & val(Text5.Text) & ""
     End If

  If val(Text5.Text) <> 0 Then
     StrSQL = StrSQL & "  and dbo.TblEndorseTransMashar.ID >=" & val(Text5.Text) & ""
  End If
  If val(Text6.Text) <> 0 Then
     StrSQL = StrSQL & "  and dbo.TblEndorseTransMashar.ID <=" & val(Text6.Text) & ""
  End If
  If val(DcbSeasonsID5.BoundText) <> 0 And DcbSeasonsID5.Text <> "" Then
     StrSQL = StrSQL & "  and dbo.TblEndorseTransMashar.SeasonsID =" & val(DcbSeasonsID5.BoundText) & ""
  End If
  If val(CompanyID5.BoundText) <> 0 And CompanyID5.Text <> "" Then
     StrSQL = StrSQL & "  and dbo.TblEndorseTransMashar.CompanyID =" & val(CompanyID5.BoundText) & ""
  End If
  If val(DcbVehicleType5.BoundText) <> 0 And DcbVehicleType5.Text <> "" Then
     StrSQL = StrSQL & "  and dbo.TblEndorseTransMashar.VehicleType =" & val(DcbVehicleType5.BoundText) & ""
  End If
  If val(DcbPath5.BoundText) <> 0 And DcbPath5.Text <> "" Then
     StrSQL = StrSQL & "  and dbo.TblEndorseTransMashar.PathID =" & val(DcbPath5.BoundText) & ""
  End If
   If val(DcbDepandID5.BoundText) <> 0 And DcbDepandID5.Text <> "" Then
     StrSQL = StrSQL & "  and dbo.TblEndorseTransMashar.DepandID =" & val(DcbDepandID5.BoundText) & ""
   End If
    If val(Nationality5.BoundText) <> 0 And Nationality5.Text <> "" Then
     StrSQL = StrSQL & "  and dbo.TblEndorseTransMashar.Nationality =" & val(Nationality5.BoundText) & ""
   End If
   If Not IsNull(Me.DtpDateFrom5.value) Then
                   StrSQL = StrSQL & " AND dbo.TblEndorseTransMashar.SDate >=" & SQLDate(Me.DtpDateFrom5.value, True) & ""
   End If
   If Not IsNull(Me.DtpDateTo5.value) Then
                   StrSQL = StrSQL & " AND dbo.TblEndorseTransMashar.SDate <=" & SQLDate(Me.DtpDateTo5.value, True) & ""
   End If

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «· ﬁ—Ì—"
  Else
  Msg = "Not Found Data"
  
  End If
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    Else
 rs.MoveFirst
 print_reportHajj StrSQL, 1
    End If
End Sub

Public Sub GetDataAmraStus(Optional Index As Integer = 0)
    Dim StrSQL As String
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim I As Integer
 If Index = 0 Then
StrSQL = "SELECT     COUNT(ID) AS CuntID, StusID, SUM(VehicleNo) AS SumVehicleNo"
StrSQL = StrSQL & " From dbo.TblBookingRequest"
StrSQL = StrSQL & " Where (1 = 1)"
If val(Me.DcbSeasonsID.BoundText) <> 0 And Me.DcbSeasonsID.Text <> "" Then
StrSQL = StrSQL & "  and SeasonsID =" & val(Me.DcbSeasonsID.BoundText) & ""
End If
If val(DcbStus.ListIndex) = 0 Then
StrSQL = StrSQL & "  and (StusID is null or   StusID=3)"
ElseIf val(DcbStus.ListIndex) = 1 Then
StrSQL = StrSQL & "  and StusID =1"
ElseIf val(DcbStus.ListIndex) = 2 Then
StrSQL = StrSQL & "  and StusID =2"
End If
ElseIf Index = 1 Then
StrSQL = " SELECT     COUNT(dbo.TblBookingRequest.ID) AS CuntID, dbo.TblBookingRequest.ProgrammID, dbo.TblProgrammTypes.Name, dbo.TblProgrammTypes.NameE,"
StrSQL = StrSQL & "                      SUM(dbo.TblBookingRequest.VehicleNo) As SumVehicleNo"
StrSQL = StrSQL & " FROM         dbo.TblBookingRequest LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblProgrammTypes ON dbo.TblBookingRequest.ProgrammID = dbo.TblProgrammTypes.ID"
StrSQL = StrSQL & " Where  (dbo.TblBookingRequest.StusID = 1)"
If val(Me.DcbSeasonsID.BoundText) <> 0 And Me.DcbSeasonsID.Text <> "" Then
StrSQL = StrSQL & "  and dbo.TblBookingRequest.SeasonsID =" & val(Me.DcbSeasonsID.BoundText) & ""
End If
If val(ProgrammID.BoundText) <> 0 And ProgrammID.Text <> "" Then
StrSQL = StrSQL & "  and dbo.TblBookingRequest.ProgrammID =" & val(ProgrammID.BoundText) & ""
End If
ElseIf Index = 2 Then
StrSQL = " SELECT     dbo.TblBookingRequest.ProgrammID, dbo.TblProgrammTypes.Name, dbo.TblProgrammTypes.NameE, SUM(dbo.TblBookingRequest.VehicleNo) AS SumVehicleNo, "
StrSQL = StrSQL & "                      dbo.TblBookingRequest.VehicleType, dbo.TBLCarTypes.name AS CarTypname, dbo.TBLCarTypes.namee AS CarTypnameE"
StrSQL = StrSQL & " FROM         dbo.TblBookingRequest LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TBLCarTypes ON dbo.TblBookingRequest.VehicleType = dbo.TBLCarTypes.id LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblProgrammTypes ON dbo.TblBookingRequest.ProgrammID = dbo.TblProgrammTypes.ID"
StrSQL = StrSQL & "  Where (dbo.TblBookingRequest.StusID = 1)"
If val(Me.DcbSeasonsID.BoundText) <> 0 And Me.DcbSeasonsID.Text <> "" Then
StrSQL = StrSQL & "  and dbo.TblBookingRequest.SeasonsID =" & val(Me.DcbSeasonsID.BoundText) & ""
End If
If val(ProgrammID.BoundText) <> 0 And ProgrammID.Text <> "" Then
StrSQL = StrSQL & "  and dbo.TblBookingRequest.ProgrammID =" & val(ProgrammID.BoundText) & ""
End If
If val(DCGroup2.BoundText) <> 0 And DCGroup2.Text <> "" Then
StrSQL = StrSQL & "  and  dbo.TblBookingRequest.VehicleType =" & val(DCGroup2.BoundText) & ""
End If
ElseIf Index = 3 Then
StrSQL = " SELECT     dbo.TblBookingRequest.ProgrammID, dbo.TblProgrammTypes.Name, dbo.TblProgrammTypes.NameE, SUM(dbo.TblBookingRequest.VehicleNo) AS SumVehicleNo,"
StrSQL = StrSQL & "                       dbo.TblBookingRequest.OutClientID , dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.fullcode"
StrSQL = StrSQL & " FROM         dbo.TblBookingRequest LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblCustemers ON dbo.TblBookingRequest.OutClientID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblProgrammTypes ON dbo.TblBookingRequest.ProgrammID = dbo.TblProgrammTypes.ID"
StrSQL = StrSQL & " where 1=1"
If val(Me.DcbSeasonsID.BoundText) <> 0 And Me.DcbSeasonsID.Text <> "" Then
StrSQL = StrSQL & "  and dbo.TblBookingRequest.SeasonsID =" & val(Me.DcbSeasonsID.BoundText) & ""
End If
If val(ProgrammID.BoundText) <> 0 And ProgrammID.Text <> "" Then
StrSQL = StrSQL & "  and dbo.TblBookingRequest.ProgrammID =" & val(ProgrammID.BoundText) & ""
End If
If val(OutClientID.BoundText) <> 0 And OutClientID.Text <> "" Then
StrSQL = StrSQL & "  and dbo.TblBookingRequest.OutClientID =" & val(OutClientID.BoundText) & ""
End If

ElseIf Index = 4 Then
StrSQL = "SELECT     dbo.TblBookingRequest.ID, dbo.TblBookingRequest.SDate, dbo.TblBookingRequest.BranchID, dbo.TblBranchesData.branch_name, "
StrSQL = StrSQL & "                      dbo.TblBranchesData.branch_namee, dbo.TblBookingRequest.OutClientID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode,"
StrSQL = StrSQL & "                      dbo.TblBookingRequest.ProgrammID, dbo.TblProgrammTypes.Name, dbo.TblProgrammTypes.NameE, dbo.TblBookingRequest.StusID,"
StrSQL = StrSQL & "                      dbo.TblBookingRequest.VehicleNo, dbo.TblBookingRequest.ArriveDate, dbo.TblBookingRequest.ArriveTime, dbo.TblBookingRequest.VehicleType,"
StrSQL = StrSQL & "                      dbo.TBLCarTypes.name AS CarsType, dbo.TBLCarTypes.namee AS CarsTypeE, dbo.TblBookingRequest.CompnyIn, dbo.TblBookingRequest.NoteSerial1,"
StrSQL = StrSQL & "                      dbo.TblBookingRequest.SeasonsID, dbo.TblCompaniesGroup.Name AS SeasonsName, dbo.TblCompaniesGroup.NameE AS SeasonsNameE, dbo.TblBookingRequest.CusNo"
StrSQL = StrSQL & " FROM         dbo.TblBookingRequest LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblCompaniesGroup ON dbo.TblBookingRequest.SeasonsID = dbo.TblCompaniesGroup.ID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TBLCarTypes ON dbo.TblBookingRequest.VehicleType = dbo.TBLCarTypes.id LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblProgrammTypes ON dbo.TblBookingRequest.ProgrammID = dbo.TblProgrammTypes.ID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblCustemers ON dbo.TblBookingRequest.OutClientID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblBranchesData ON dbo.TblBookingRequest.BranchID = dbo.TblBranchesData.branch_id"
StrSQL = StrSQL & " where 1=1"
If val(Me.DcbSeasonsID.BoundText) <> 0 And Me.DcbSeasonsID.Text <> "" Then
StrSQL = StrSQL & "  and dbo.TblBookingRequest.SeasonsID =" & val(Me.DcbSeasonsID.BoundText) & ""
End If
If val(ProgrammID.BoundText) <> 0 And ProgrammID.Text <> "" Then
StrSQL = StrSQL & "  and dbo.TblBookingRequest.ProgrammID =" & val(ProgrammID.BoundText) & ""
End If
If val(OutClientID.BoundText) <> 0 And OutClientID.Text <> "" Then
StrSQL = StrSQL & "  and dbo.TblBookingRequest.OutClientID =" & val(OutClientID.BoundText) & ""
End If
If val(DCGroup2.BoundText) <> 0 And DCGroup2.Text <> "" Then
StrSQL = StrSQL & "  and  dbo.TblBookingRequest.VehicleType =" & val(DCGroup2.BoundText) & ""
End If
If val(DcbStus.ListIndex) = 0 Then
StrSQL = StrSQL & "  and (dbo.TblBookingRequest.StusID is null or   dbo.TblBookingRequest.StusID=3)"
ElseIf val(DcbStus.ListIndex) = 1 Then
StrSQL = StrSQL & "  and dbo.TblBookingRequest.StusID =1"
ElseIf val(DcbStus.ListIndex) = 2 Then
StrSQL = StrSQL & "  and dbo.TblBookingRequest.StusID =2"
End If
End If
'''''''''
 If Not IsNull(Me.DtpDateFrom2.value) Then
                   StrSQL = StrSQL & " AND dbo.TblBookingRequest.SDate >=" & SQLDate(Me.DtpDateFrom2.value, True) & ""
   End If
  If Not IsNull(Me.DtpDateTo2.value) Then
                   StrSQL = StrSQL & " AND dbo.TblBookingRequest.SDate<=" & SQLDate(Me.DtpDateTo2.value, True) & ""
      End If
If Index = 0 Then
StrSQL = StrSQL & " GROUP BY StusID"
ElseIf Index = 1 Then
StrSQL = StrSQL & " GROUP BY dbo.TblBookingRequest.ProgrammID, dbo.TblProgrammTypes.Name, dbo.TblProgrammTypes.NameE"
ElseIf Index = 2 Then
StrSQL = StrSQL & " GROUP BY dbo.TblBookingRequest.ProgrammID, dbo.TblProgrammTypes.Name, dbo.TblProgrammTypes.NameE, dbo.TblBookingRequest.VehicleType,"
StrSQL = StrSQL & "                      dbo.TBLCarTypes.Name , dbo.TBLCarTypes.NameE"
ElseIf Index = 3 Then
StrSQL = StrSQL & " GROUP BY dbo.TblBookingRequest.ProgrammID, dbo.TblProgrammTypes.Name, dbo.TblProgrammTypes.NameE, dbo.TblBookingRequest.OutClientID,"
StrSQL = StrSQL & "                      dbo.TblCustemers.CusName , dbo.TblCustemers.CusNamee, dbo.TblCustemers.fullcode"
End If
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «· ﬁ—Ì—"
  Else
  Msg = "Not Found Data"
  
  End If
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    Else
 rs.MoveFirst
 If Index = 4 Then
 Index = 15
 End If
 print_reportAmara StrSQL, Index
    End If
End Sub
Public Sub GetData(Optional Index As Integer = 0)
    Dim StrSQL As String
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim I As Integer
    
StrSQL = "SELECT     dbo.TblEndorseTransMashar.ID, dbo.TblEndorseTransMashar.SDate, dbo.TblEndorseTransMashar.BranchID, dbo.TblBranchesData.branch_name, "
StrSQL = StrSQL & "                      dbo.TblBranchesData.branch_namee, dbo.TblEndorseTransMashar.ApproveID, dbo.TblEndorseTransMashar.Nationality, dbo.Nationality.name, dbo.Nationality.namee,"
StrSQL = StrSQL & "                       dbo.TblEndorseTransMashar.TotOlds, dbo.TblEndorseTransMashar.TotYoungs, dbo.TblEndorseTransMashar.Total, dbo.TblEndorseTransMashar.Remark,"
StrSQL = StrSQL & "                      dbo.TblEndorseTransMashar.CreationDate, dbo.TblEndorseTransMashar.GroupName, dbo.TblEndorseTransMashar.ReceptName,"
StrSQL = StrSQL & "                      dbo.TblEndorseTransMashar.ReceptOffice, dbo.TblEndorseTransMashar.PArCode, dbo.TblEndorseTransMashar.Password, dbo.TblEndorseTransMashar.[Session],"
StrSQL = StrSQL & "                      dbo.TblEndorseTransMashar.LargPrice, dbo.TblEndorseTransMashar.SmalPrice, dbo.TblEndorseTransMashar.TotalPrice, dbo.TblEndorseTransMashar.Capacity,"
StrSQL = StrSQL & "                      dbo.TblEndorseTransMashar.NoVehicle, dbo.TblEndorseTransMashar.Phone, dbo.TblEndorseTransMashar.ReceptTime, dbo.TblEndorseTransMashar.RecordDateH,"
StrSQL = StrSQL & "                      dbo.TblEndorseTransMashar.NoTrip, dbo.TblEndorseTransMashar.PathID, dbo.TblShrines.Name AS PathName, dbo.TblShrines.NameE AS PathNameE,"
StrSQL = StrSQL & "                      dbo.TblEndorseTransMashar.SeasonsID, dbo.TblCompaniesGroup.Name AS SeasonName, dbo.TblCompaniesGroup.NameE AS SeasonNameE,"
StrSQL = StrSQL & "                      dbo.TblEndorseTransMashar.CompanyID, dbo.TblTourismCompanies.Name AS CompnyName, dbo.TblTourismCompanies.NameE AS CompnyNameE,"
StrSQL = StrSQL & "                      dbo.TblEndorseTransMashar.LocationID, dbo.TblLocations.Name AS LocationName, dbo.TblLocations.NameE AS LocationNameE,"
StrSQL = StrSQL & "                      dbo.TblEndorseTransMashar.VehicleType, dbo.TblBusesDistribution.ID AS DistrID, dbo.TblBusesDistribution.SDate AS DistSDate,"
StrSQL = StrSQL & "                      dbo.TblBusesDistribution.RecordDateH AS DistRecordDateH, dbo.TblBusesDistribution.NoVehicle AS DistNoVehicle,"
StrSQL = StrSQL & "                      dbo.TblBusesDistribution.Capacity AS DistCapacity, dbo.TblBusesDistribution.OrderNo, dbo.TblBusesDistribution.BranchID AS DIstBranchID,"
StrSQL = StrSQL & "                      TblBranchesData_1.branch_name AS Distbranch_name, TblBranchesData_1.branch_namee AS Distbranch_nameE,"
StrSQL = StrSQL & "                      dbo.TblBusesDistributionDet.Capacity AS DistCapacityDet, dbo.TblBusesDistributionDet.CarID, dbo.TblCarsData.BoardNO, dbo.TblCarsData.CarsTypeId,"
StrSQL = StrSQL & "                      dbo.TblCarsData.Model, dbo.TblBusesDistributionDet.EmpID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Name1, dbo.TblEmployee.Emp_Name2,"
StrSQL = StrSQL & "                      dbo.TblEmployee.Emp_Name3, dbo.TblEmployee.Emp_Name4, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee4, dbo.TblEmployee.Emp_Namee3,"
StrSQL = StrSQL & "                      dbo.TblEmployee.Emp_Namee2, dbo.TblEmployee.Emp_Namee1, dbo.TblEmployee.Emp_Namee, dbo.TBLCarTypes.name AS CarType,"
StrSQL = StrSQL & "                      dbo.TBLCarTypes.namee AS CarTypeE, dbo.TblVehicleType.Name AS VehicleName, dbo.TblVehicleType.NameE AS VehicleNameE"
StrSQL = StrSQL & " FROM         dbo.TblVehicleType RIGHT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblEndorseTransMashar ON dbo.TblVehicleType.ID = dbo.TblEndorseTransMashar.VehicleType RIGHT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TBLCarTypes RIGHT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblCarsData ON dbo.TBLCarTypes.id = dbo.TblCarsData.CarsTypeId RIGHT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblEmployee RIGHT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblBusesDistributionDet ON dbo.TblEmployee.Emp_ID = dbo.TblBusesDistributionDet.EmpID ON"
StrSQL = StrSQL & "                      dbo.TblCarsData.id = dbo.TblBusesDistributionDet.CarID RIGHT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblBranchesData TblBranchesData_1 RIGHT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblBusesDistribution ON TblBranchesData_1.branch_id = dbo.TblBusesDistribution.BranchID ON"
StrSQL = StrSQL & "                      dbo.TblBusesDistributionDet.BusDistID = dbo.TblBusesDistribution.ID ON dbo.TblEndorseTransMashar.ID = dbo.TblBusesDistribution.OrderNo LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblLocations ON dbo.TblEndorseTransMashar.LocationID = dbo.TblLocations.ID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblTourismCompanies ON dbo.TblEndorseTransMashar.CompanyID = dbo.TblTourismCompanies.ID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblCompaniesGroup ON dbo.TblEndorseTransMashar.SeasonsID = dbo.TblCompaniesGroup.ID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblShrines ON dbo.TblEndorseTransMashar.PathID = dbo.TblShrines.ID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.Nationality ON dbo.TblEndorseTransMashar.Nationality = dbo.Nationality.id LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblBranchesData ON dbo.TblEndorseTransMashar.BranchID = dbo.TblBranchesData.branch_id"
StrSQL = StrSQL & " Where (dbo.TblBusesDistribution.orderNo <> 0 And Not (dbo.TblBusesDistribution.orderNo Is Null))"

If val(TxtOrder.Text) <> 0 Then
    StrSQL = StrSQL & " AND   dbo.TblBusesDistribution.orderNo = " & val(Me.TxtOrder.Text)

End If

If val(SeasonsID.BoundColumn) <> 0 And (SeasonsID.Text) <> "" Then
    StrSQL = StrSQL & " AND   dbo.TblEndorseTransMashar.SeasonsID = " & val(Me.SeasonsID.BoundText) & " "

End If

If val(CompanyID.BoundText) <> 0 And (Me.CompanyID.Text) <> "" Then
    StrSQL = StrSQL & " AND   dbo.TblEndorseTransMashar.CompanyID = " & val(Me.CompanyID.BoundText)

End If
If Me.TxtReceptOffice.Text <> "" Then
    StrSQL = StrSQL & " AND   dbo.TblEndorseTransMashar.ReceptOffice = '" & Me.TxtReceptOffice.Text & " '"

End If
If val(DcbDriver.BoundText) <> 0 And (Me.DcbDriver.Text) <> "" Then
    StrSQL = StrSQL & " AND   dbo.TblBusesDistributionDet.EmpID = " & val(Me.DcbDriver.BoundText)
End If
If val(DCGroup.BoundText) <> 0 And (Me.DCGroup.Text) <> "" Then
    StrSQL = StrSQL & " AND   dbo.TblCarsData.CarsTypeId = " & val(Me.DCGroup.BoundText)
End If
If val(DcbVehicleType.BoundText) <> 0 And (Me.DcbVehicleType.Text) <> "" Then
    StrSQL = StrSQL & " AND   dbo.TblEndorseTransMashar.VehicleType = " & val(Me.DcbVehicleType.BoundText)
End If

'''''''''
 If Not IsNull(Me.DtpDateFrom.value) Then
                   StrSQL = StrSQL & " AND dbo.TblBusesDistribution.SDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
      End If
       If Not IsNull(Me.DtpDateTo.value) Then
                   StrSQL = StrSQL & " AND dbo.TblBusesDistribution.SDate<=" & SQLDate(Me.DtpDateTo.value, True) & ""
      End If

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «· ﬁ—Ì—"
  Else
  Msg = "Not Found Data"
  End If
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    Else
 rs.MoveFirst
  print_report StrSQL, 0
    End If
End Sub
Function print_reportAmara(Optional NoteSerial As String, Optional Index As Integer = 0)
     
    Set rs = New ADODB.Recordset
    rs.Open NoteSerial, Cn, adOpenStatic, adLockReadOnly, adCmdText
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
If Index = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepHajAmraStus.rpt"
            Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepHajAmraStus.rpt"
       End If
 ElseIf Index = 1 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepHajAmrVehProg.rpt"
            Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepHajAmrVehProg.rpt"
       End If
  ElseIf Index = 2 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepHajAmrVehProgCarTy.rpt"
            Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepHajAmrVehProgCarTy.rpt"
       End If
    ElseIf Index = 3 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepHajAmrVehProgCustomer.rpt"
            Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepHajAmrVehProgCustomer.rpt"
       End If
       ElseIf Index = 4 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepHajAmrTransection.rpt"
            Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepHajAmrTransection.rpt"
       End If
     ElseIf Index = 5 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepHajAmrOrder.rpt"
            Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepHajAmrOrder.rpt"
       End If
    ElseIf Index = 6 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepHajAmrOrderPathNotExe.rpt"
            Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepHajAmrOrderPathNotExe.rpt"
       End If
ElseIf Index = 7 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepHajAmrOrderPathExe.rpt"
            Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepHajAmrOrderPathExe.rpt"
       End If
  ElseIf Index = 8 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepHajAmrOrderAllPath.rpt"
            Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepHajAmrOrderAllPath.rpt"
       End If
  ElseIf Index = 9 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepHajAmrOrderEnd.rpt"
            Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepHajAmrOrderEnd.rpt"
       End If
    ElseIf Index = 10 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepHajAmrDeported.rpt"
            Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepHajAmrDeported.rpt"
       End If
        ElseIf Index = 11 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepHajAmrTransCar.rpt"
            Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepHajAmrTransCar.rpt"
       End If
     ElseIf Index = 12 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepHajAmrCarCuurlocation.rpt"
            Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepHajAmrCarCuurlocation.rpt"
       End If
       ElseIf Index = 13 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepHajAmrTypeCarsDat.rpt"
            Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepHajAmrTypeCarsDat.rpt"
       End If
       
         ElseIf Index = 14 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepHajAmrDriverData.rpt"
            Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepHajAmrDriverData.rpt"
       End If
      ElseIf Index = 15 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepHajAmrBooking.rpt"
            Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepHajAmrBooking.rpt"
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
        If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
       Else
       Msg = "Not Found Data"
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
        ' xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
        StrReportTitle = "" '& StrAccountName
    Else
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        'xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        StrReportTitle = ""
    End If
    If Index <= 3 Or Index = 15 Then
    If Not IsNull(DtpDateFrom2.value) Then
    xReport.ParameterFields(4).AddCurrentValue DtpDateFrom2.value
    xReport.ParameterFields(5).AddCurrentValue DtpDateFromh2.value
    End If
    If Not IsNull(DtpDateTo2.value) Then
    xReport.ParameterFields(6).AddCurrentValue DtpDateTo2.value
    xReport.ParameterFields(7).AddCurrentValue DtpDateToh2.value
    End If
  xReport.ParameterFields(8).AddCurrentValue Me.DcbSeasonsID.Text
    ElseIf Index <= 9 Then
        If Not IsNull(DtpDateFrom3.value) Then
    xReport.ParameterFields(4).AddCurrentValue DtpDateFrom3.value
    xReport.ParameterFields(5).AddCurrentValue DtpDateFromH3.value
    End If
    If Not IsNull(DtpDateTo3.value) Then
    xReport.ParameterFields(6).AddCurrentValue DtpDateTo3.value
    xReport.ParameterFields(7).AddCurrentValue DtpDateToH3.value
    End If
    Else
    If Not IsNull(DtpDateFrom4.value) Then
    xReport.ParameterFields(4).AddCurrentValue DtpDateFrom4.value
    xReport.ParameterFields(5).AddCurrentValue DtpDateFromH4.value
    End If
    If Not IsNull(DtpDateTo4.value) Then
    xReport.ParameterFields(6).AddCurrentValue DtpDateTo4.value
    xReport.ParameterFields(7).AddCurrentValue DtpDateToH4.value
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
Function print_reportHajj(Optional NoteSerial As String, Optional Index As Integer = 0)
     
    Set rs = New ADODB.Recordset
    rs.Open NoteSerial, Cn, adOpenStatic, adLockReadOnly, adCmdText
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
If Index = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepHajjrptDetectingHajj.rpt"
            Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepHajjrptDetectingHajj.rpt"
       End If
 Else
     If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepHajjrptDetectingHajj2.rpt"
            Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepHajjrptDetectingHajj2.rpt"
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
        If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
       Else
       Msg = "Not Found Data"
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
        ' xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
        StrReportTitle = "" '& StrAccountName
    Else
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        'xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        StrReportTitle = ""
    End If
   If Not IsNull(DtpDateFrom5.value) Then
    xReport.ParameterFields(4).AddCurrentValue DtpDateFrom5.value
    xReport.ParameterFields(5).AddCurrentValue DtpDateFromH5.value
    End If
    If Not IsNull(DtpDateTo5.value) Then
    xReport.ParameterFields(6).AddCurrentValue DtpDateTo5.value
    xReport.ParameterFields(7).AddCurrentValue DtpDateToH5.value
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
Function print_report(Optional NoteSerial As String, Optional Index As Integer = 0)
     
    Set rs = New ADODB.Recordset
    rs.Open NoteSerial, Cn, adOpenStatic, adLockReadOnly, adCmdText
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
If Index = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepHajjrptDetecting.rpt"
            Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepHajjrptDetecting.rpt"
       End If
 Else
       If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepHajjrptDetecting2.rpt"
            Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepHajjrptDetecting2.rpt"
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
        If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
       Else
       Msg = "Not Found Data"
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
        ' xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
        StrReportTitle = "" '& StrAccountName
    Else
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        'xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        StrReportTitle = ""
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


Private Sub ISButton1_Click()
If Me.Rdam2(0).value = True Then
GetDataAmraORDER 0
ElseIf Me.Rdam2(1).value = True Then
GetDataAmraORDER 1
ElseIf Me.Rdam2(2).value = True Then
GetDataAmraORDER 2
ElseIf Me.Rdam2(3).value = True Then
GetDataAmraORDER 3
ElseIf Me.Rdam2(4).value = True Then
GetDataAmraORDER 4
ElseIf Me.Rdam2(5).value = True Then
GetDataAmraORDER 5
Else
MsgBox "Ì—ÃÏ «Œ Ì«— ‰Ê⁄ «· ﬁ—Ì—"
Exit Sub
End If
End Sub

Private Sub ISButton2_Click()
If Me.Rdam3(0).value = True Then
GetDataDeported 0
ElseIf Me.Rdam3(1).value = True Then
GetDataDeported 1
ElseIf Me.Rdam3(2).value = True Then
GetDataDeported 2
ElseIf Me.Rdam3(3).value = True Then
GetDataDeporTypCars 3
ElseIf Me.Rdam3(4).value = True Then
GetBaiscDtaData 4
Else
MsgBox "Ì—ÃÏ «Œ Ì«— ‰Ê⁄ «· ﬁ—Ì—"
Exit Sub
End If
End Sub


Private Sub ISButton4_Click()
If RdHaj(0).value = True Or RdHaj(1).value = True Then
If RdHaj(0).value Then
GetDataHajjORDER1
Else
GetDataHajjORDER2
End If
Else
MsgBox "Ì—ÃÏ «Œ Ì«— ‰Ê⁄ «· ﬁ—Ì—"
Exit Sub
End If
End Sub

Private Sub NourHijriCal1_LostFocus()
 VBA.Calendar = vbCalGreg
            DTPicker1.value = ToGregorianDate(NourHijriCal1.value)
End Sub

Private Sub NourHijriCal2_LostFocus()
 VBA.Calendar = vbCalGreg
            DTPicker2.value = ToGregorianDate(NourHijriCal2.value)
End Sub

Private Sub OutClientID3_Change()
OutClientID3_Click (0)
End Sub

Private Sub OutClientID3_Click(Area As Integer)
   Dim Fullcode As String
    GetCustomersDetail val(OutClientID3.BoundText), , Fullcode, 1
    Text1.Text = Fullcode
End Sub

Private Sub Print_Click()
If Me.Rdam(0).value = True Then
GetDataAmraStus 0
ElseIf Me.Rdam(1).value = True Then
GetDataAmraStus 1
ElseIf Me.Rdam(2).value = True Then
GetDataAmraStus 2
ElseIf Me.Rdam(3).value = True Then
GetDataAmraStus 3
ElseIf Me.Rdam(4).value = True Then
GetDataAmraStus 4
Else
MsgBox "Ì—ÃÏ «Œ Ì«— ‰Ê⁄ «· ﬁ—Ì—"
Exit Sub
End If
End Sub



Private Sub Text1_KeyPress(KeyAscii As Integer)
Dim CUSTID As Integer
 If KeyAscii = vbKeyReturn Then
        GetCustomersDetail CUSTID, , Text1.Text
        OutClientID3.BoundText = CUSTID
        OutClientID3.SetFocus
    End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
Dim CUSTID As Integer
 If KeyAscii = vbKeyReturn Then
        GetCustomersDetail CUSTID, , Text2.Text
        OutClientID.BoundText = CUSTID
        OutClientID.SetFocus
    End If
End Sub
Private Sub OutClientID_Change()
OutClientID_Click (0)
End Sub

Private Sub OutClientID_Click(Area As Integer)
   Dim Fullcode As String
    GetCustomersDetail val(OutClientID.BoundText), , Fullcode, 1
    Text2.Text = Fullcode
End Sub


Private Sub TxtFromOrder_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, TxtFromOrder.Text, 0)
End Sub

Private Sub TxtIDFrom_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, TxtIDFrom.Text, 0)
End Sub

Private Sub TxtIDTO_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, TxtIDFrom.Text, 0)
End Sub

Private Sub TxToOrder_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, TxToOrder.Text, 0)
End Sub
'########################## Khaled part ##################################
Function my_print_report(Optional NoteSerial As String, Optional Index As Integer = 0)
     
    Set rs = New ADODB.Recordset
    rs.Open NoteSerial, Cn, adOpenStatic, adLockReadOnly, adCmdText
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
    If Rdam2(6).value Then
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "BookingReqRepNonOp.rpt"
            Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "BookingReqRepNonOp.rpt"
       End If
    ElseIf Rdam2(7).value Then
       If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "Booking ReqRepOp.rpt"
            Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "Booking ReqRepOp.rpt"
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
        If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
       Else
       Msg = "Not Found Data"
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
        ' xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
        StrReportTitle = "" '& StrAccountName
    Else
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        'xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        StrReportTitle = ""
    End If
    If Not IsNull(DTPicker1.value) Then
    xReport.ParameterFields(4).AddCurrentValue DTPicker1.value
    End If
    If Not IsNull(DTPicker2.value) Then
    xReport.ParameterFields(5).AddCurrentValue DTPicker2.value
    End If
    If Not IsNull(DTPicker1.value) Then
    xReport.ParameterFields(6).AddCurrentValue ToHijriDate(DTPicker1.value)
    End If
    If Not IsNull(DTPicker2.value) Then
    xReport.ParameterFields(7).AddCurrentValue ToHijriDate(DTPicker2.value)
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
Private Sub myOutClientID3_Change()
myOutClientID3_Click (0)
End Sub

Private Sub myOutClientID3_Click(Area As Integer)
   Dim Fullcode As String
    GetCustomersDetail val(myOutClientID3.BoundText), , Fullcode, 1
    myOutClientCode.Text = Fullcode
End Sub
Private Sub myOutClientCode_KeyPress(KeyAscii As Integer)
Dim CUSTID As Integer
 If KeyAscii = vbKeyReturn Then
        GetCustomersDetail CUSTID, , myOutClientCode.Text
        myOutClientID3.BoundText = CUSTID
        myOutClientID3.SetFocus
    End If
End Sub
Private Sub ISButton3_Click()
    Dim SqlQu As String
    If Rdam2(6).value = False And Rdam2(7).value = False Then Exit Sub
    SqlQu = " SELECT     dbo.TBLCarTypes.name, dbo.TblCustemers.CusName, dbo.TblBookingRequest.ID, dbo.TblBookingRequest.SDate, dbo.TblBookingRequest.BranchID, "
    SqlQu = SqlQu & "                   dbo.TblBookingRequest.InClientID, dbo.TblBookingRequest.OutClientID, dbo.TblBookingRequest.GroupID, dbo.TblBookingRequest.CompanyID,"
    SqlQu = SqlQu & "                  dbo.TblBookingRequest.AirPortID, dbo.TblBookingRequest.AirLineID, dbo.TblBookingRequest.FlightNo, dbo.TblBookingRequest.emp, dbo.TblBookingRequest.other,"
    SqlQu = SqlQu & "                  dbo.TblBookingRequest.EmpID, dbo.TblBookingRequest.EmpName, dbo.TblBookingRequest.EmpCode, dbo.TblBookingRequest.EmpMbile,"
    SqlQu = SqlQu & "                  dbo.TblBookingRequest.ArriveDate, dbo.TblBookingRequest.ArriveTime, dbo.TblBookingRequest.VehicleNo, dbo.TblBookingRequest.ProgrammID,"
    SqlQu = SqlQu & "                  dbo.TblBookingRequest.Model, dbo.TblBookingRequest.MekkaHotelID, dbo.TblBookingRequest.MadinaHotelID, dbo.TblBookingRequest.JeddahHotelID,"
    SqlQu = SqlQu & "                  dbo.TblBookingRequest.VehicleType, dbo.TblBookingRequest.CreationUserID, dbo.TblBookingRequest.CreationDate, dbo.TblBookingRequest.ModelID,"
    SqlQu = SqlQu & "                  dbo.TblBookingRequest.GroupName, dbo.TblBookingRequest.UserID, dbo.TblBookingRequest.UserID2, dbo.TblBookingRequest.ApproveTime,"
    SqlQu = SqlQu & "                  dbo.TblBookingRequest.ApproveDate, dbo.TblBookingRequest.ApproveFlag, dbo.TblBookingRequest.ReservNo, dbo.TblBookingRequest.UseFlag,"
    SqlQu = SqlQu & "                  dbo.TblBookingRequest.StusID, dbo.TblBookingRequest.RemarkApprove, dbo.TblBookingRequest.HotelMakh, dbo.TblBookingRequest.HotelMadinh,"
    SqlQu = SqlQu & "                  dbo.TblBookingRequest.HotelJaddah, dbo.TblBookingRequest.CusNo, dbo.TblBookingRequest.CompnyIn, dbo.TblBookingRequest.CompnyOut,"
    SqlQu = SqlQu & "                  dbo.TblProgrammTypes.Name AS ProgName, dbo.tblbookingrequest2.ID AS ActivOrderID, dbo.tblbookingrequest2.Total, dbo.TblBookingRequest.NoteSerial1,"
    SqlQu = SqlQu & "                  dbo.TblBookingRequest.SeasonsID, dbo.TblCompaniesGroup.Name AS SeasonsName, dbo.TblCompaniesGroup.NameE AS SeasonsNameE,"
    SqlQu = SqlQu & "                  dbo.tblbookingrequest2.NoteSerial1 AS OderNoteSerial1"
    SqlQu = SqlQu & "      FROM         dbo.TblBookingRequest LEFT OUTER JOIN"
    SqlQu = SqlQu & "                  dbo.TblCompaniesGroup ON dbo.TblBookingRequest.SeasonsID = dbo.TblCompaniesGroup.ID LEFT OUTER JOIN"
    SqlQu = SqlQu & "                  dbo.TblProgrammTypes ON dbo.TblBookingRequest.ProgrammID = dbo.TblProgrammTypes.ID LEFT OUTER JOIN"
    SqlQu = SqlQu & "                  dbo.TBLCarTypes ON dbo.TblBookingRequest.VehicleType = dbo.TBLCarTypes.id LEFT OUTER JOIN"
    SqlQu = SqlQu & "                  dbo.tblbookingrequest2 ON dbo.TblBookingRequest.ID = dbo.tblbookingrequest2.OrdeNo LEFT OUTER JOIN"
    SqlQu = SqlQu & "                  dbo.TblCustemers ON dbo.TblBookingRequest.OutClientID = dbo.TblCustemers.CusID"
    SqlQu = SqlQu & " Where StusID = 1 And ((dbo.GetTripNo(dbo.tblbookingrequest.ID) * tblbookingrequest.VehicleNo) = dbo.GetTotalTripNo(dbo.tblbookingrequest.ID)) "
    
    If Rdam2(6).value = True Then
        SqlQu = SqlQu & " and (select count(*) from tblbookingrequest2 where OrdeNo = dbo.TblBookingRequest.ID) = 0 "
    ElseIf Rdam2(7).value = True Then
        SqlQu = SqlQu & " and (select count(*) from tblbookingrequest2 where OrdeNo = dbo.TblBookingRequest.ID) <> 0 "
    End If
    
    If TxtConfFrom.Text <> "" And TxtConfTo.Text <> "" Then
         SqlQu = SqlQu & " and TblBookingRequest.NoteSerial1 >= " & TxtConfFrom.Text & "  and  TblBookingRequest.NoteSerial1 <= " & TxtConfTo.Text & " "
    End If
    
    If myOutClientID3.BoundText <> "" Then
        SqlQu = SqlQu & " and TblBookingRequest.OutClientID =  " & myOutClientID3.BoundText & " "
    End If
      If DcbSeasonsID4.Text <> "" And val(Me.DcbSeasonsID4.BoundText) <> 0 Then
        SqlQu = SqlQu & " and TblBookingRequest.SeasonsID =  " & val(DcbSeasonsID4.BoundText) & " "
    End If
    
    If myDCGroup.BoundText <> "" Then
        SqlQu = SqlQu & " and TblBookingRequest.VehicleType = " & myDCGroup.BoundText & " "
    End If
       If myDCGroup.BoundText <> "" Then
        SqlQu = SqlQu & " and TblBookingRequest.VehicleType = " & myDCGroup.BoundText & " "
    End If
    
    If myProgrammID3.BoundText <> "" Then
        SqlQu = SqlQu & " and TblBookingRequest.ProgrammID = " & myProgrammID3.BoundText & " "
    End If
    
    If TxtVichNo.Text <> "" Then
        SqlQu = SqlQu & " and TblBookingRequest.VehicleNo = " & TxtVichNo.Text & " "
    End If
    
    If TxtRunFrom.Text <> "" And TxtRunTo.Text <> "" And Rdam2(7).value = True Then
        SqlQu = SqlQu & " and TblBookingRequest2.NoteSerial1 >= " & val(TxtRunFrom.Text) & "  and  TblBookingRequest2.NoteSerial1 <= " & val(TxtRunTo.Text) & " "
    End If
    
    If Not IsNull(Me.DTPicker1.value) Then
        SqlQu = SqlQu & " AND dbo.TblBookingRequest.SDate >=" & SQLDate(Me.DTPicker1.value, True) & ""
    End If
    
    If Not IsNull(Me.DTPicker2.value) Then
        SqlQu = SqlQu & " AND dbo.TblBookingRequest.SDate <=" & SQLDate(Me.DTPicker2.value, True) & ""
    End If
    
    my_print_report SqlQu
End Sub

Private Sub Rdam2_Click(Index As Integer)
    If Rdam2(6).value = True Then
        Frame17.Visible = False
    ElseIf Rdam2(7).value = True Then
        Frame17.Visible = True
    End If
End Sub
