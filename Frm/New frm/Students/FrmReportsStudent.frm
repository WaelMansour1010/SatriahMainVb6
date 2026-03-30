VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SUITEC~1.OCX"
Begin VB.Form FrmReportsStudent 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   7320
   ClientLeft      =   3060
   ClientTop       =   1890
   ClientWidth     =   10200
   Icon            =   "FrmReportsStudent.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   10200
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton btnClear 
      BackColor       =   &H00E2E9E9&
      Caption         =   "„”Õ"
      Height          =   495
      Left            =   1320
      TabIndex        =   8
      Top             =   6750
      Width           =   1335
   End
   Begin VB.Frame XPPnlTime 
      BackColor       =   &H00E2E9E9&
      Caption         =   "›Ï «·› —…"
      Height          =   1185
      Left            =   4320
      TabIndex        =   3
      Top             =   7320
      Visible         =   0   'False
      Width           =   2415
      Begin MSComCtl2.DTPicker XPDtbFrom 
         Height          =   345
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   609
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   93454337
         CurrentDate     =   38784
      End
      Begin MSComCtl2.DTPicker XPDtpTo 
         Height          =   345
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   609
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   93454337
         CurrentDate     =   38784
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   285
         Index           =   2
         Left            =   1560
         TabIndex        =   7
         Top             =   360
         Width           =   465
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï"
         Height          =   285
         Index           =   0
         Left            =   1560
         TabIndex        =   6
         Top             =   720
         Width           =   465
      End
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   495
      Index           =   2
      Left            =   120
      TabIndex        =   0
      Top             =   6750
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
   Begin C1SizerLibCtl.C1Tab XPTab301 
      Height          =   6045
      Left            =   0
      TabIndex        =   9
      Top             =   600
      Width           =   12720
      _cx             =   22437
      _cy             =   10663
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
      Caption         =   " ﬁ«—Ì— | ﬁ—Ì— „⁄·Ê„«  «·Õ÷Ê— ›Ì «·„Ã„Ê⁄« | ﬁ«—Ì— «·„ﬁ«”Ì« | ﬁ—Ì— «·« ›«ﬁÌ« |«Ã„«·Ì «·« ›«ﬁÌ« "
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
      Picture(0)      =   "FrmReportsStudent.frx":038A
      Begin VB.Frame Fra 
         BackColor       =   &H00E2E9E9&
         Caption         =   "{"
         Height          =   5580
         Index           =   4
         Left            =   14265
         RightToLeft     =   -1  'True
         TabIndex        =   141
         Top             =   45
         Width           =   12630
         Begin VB.Frame Frame2 
            BackColor       =   &H00E2E9E9&
            Caption         =   "Õœœ «·› —…"
            Height          =   1260
            Index           =   3
            Left            =   150
            RightToLeft     =   -1  'True
            TabIndex        =   146
            Top             =   2790
            Width           =   5832
            Begin VB.OptionButton optStatus 
               Caption         =   "«·ﬂ·"
               Height          =   195
               Index           =   5
               Left            =   180
               RightToLeft     =   -1  'True
               TabIndex        =   165
               Top             =   840
               Width           =   765
            End
            Begin VB.OptionButton optStatus 
               Caption         =   "«·« ›«ﬁÌ«  «·„·€«… ›ﬁÿ"
               Height          =   195
               Index           =   4
               Left            =   1530
               RightToLeft     =   -1  'True
               TabIndex        =   164
               Top             =   840
               Width           =   1785
            End
            Begin VB.OptionButton optStatus 
               Caption         =   "«·« ›«ﬁÌ«  «·„›⁄·… ›ﬁÿ"
               Height          =   195
               Index           =   3
               Left            =   3450
               RightToLeft     =   -1  'True
               TabIndex        =   163
               Top             =   840
               Value           =   -1  'True
               Width           =   1785
            End
            Begin MSComCtl2.DTPicker toDate5 
               Height          =   330
               Left            =   330
               TabIndex        =   147
               Top             =   240
               Width           =   1605
               _ExtentX        =   2831
               _ExtentY        =   582
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   93454337
               CurrentDate     =   41640
            End
            Begin MSComCtl2.DTPicker Fromdate5 
               Height          =   330
               Left            =   2730
               TabIndex        =   148
               Top             =   270
               Width           =   1665
               _ExtentX        =   2937
               _ExtentY        =   582
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   93454337
               CurrentDate     =   41640
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "„‰"
               Height          =   435
               Index           =   31
               Left            =   4320
               RightToLeft     =   -1  'True
               TabIndex        =   150
               Top             =   270
               Width           =   540
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "≈·Ï"
               Height          =   435
               Index           =   27
               Left            =   1980
               RightToLeft     =   -1  'True
               TabIndex        =   149
               Top             =   270
               Width           =   540
            End
         End
         Begin VB.Frame Frame4 
            Height          =   4572
            Index           =   3
            Left            =   6000
            TabIndex        =   144
            Top             =   120
            Width           =   4095
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "«·”« —Ì…"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   15.75
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00008000&
               Height          =   375
               Index           =   3
               Left            =   120
               TabIndex        =   145
               Top             =   3480
               Width           =   3855
            End
         End
         Begin VB.TextBox Text13 
            Height          =   285
            Left            =   6360
            TabIndex        =   143
            Top             =   5760
            Width           =   855
         End
         Begin VB.TextBox Text9 
            Height          =   285
            Left            =   5400
            TabIndex        =   142
            Top             =   6000
            Width           =   855
         End
         Begin ImpulseButton.ISButton ISButton3 
            Height          =   492
            Left            =   120
            TabIndex        =   151
            Top             =   4200
            Width           =   5808
            _ExtentX        =   10239
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
         Begin MSDataListLib.DataCombo DcboUsers 
            Height          =   315
            Index           =   2
            Left            =   600
            TabIndex        =   158
            Top             =   2310
            Width           =   3795
            _ExtentX        =   6694
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„” Œœ„"
            Height          =   375
            Index           =   35
            Left            =   4350
            RightToLeft     =   -1  'True
            TabIndex        =   159
            Top             =   2340
            Width           =   1275
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
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
            Height          =   795
            Index           =   33
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   153
            Top             =   4740
            Width           =   5760
         End
         Begin VB.Shape Shape5 
            BorderWidth     =   2
            Height          =   495
            Left            =   0
            Top             =   6000
            Width           =   6975
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Caption         =   "Ì—ÃÏ «Œ Ì«— «·›—⁄ «Ê «· «—ÌŒ «Ê ”Ê› ÌﬂÊ‰ «· ﬁ—Ì— «Ã„«·Ì ·ﬂ· «·›—Ê⁄  Ê«·„œ…"
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
            Height          =   450
            Index           =   32
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   152
            Top             =   6240
            Width           =   6975
         End
      End
      Begin VB.Frame Fra 
         BackColor       =   &H00E2E9E9&
         Caption         =   "{"
         Height          =   5580
         Index           =   3
         Left            =   13965
         RightToLeft     =   -1  'True
         TabIndex        =   118
         Top             =   45
         Width           =   12630
         Begin VB.TextBox Text12 
            Height          =   285
            Left            =   5400
            TabIndex        =   130
            Top             =   6000
            Width           =   855
         End
         Begin VB.TextBox Text11 
            Height          =   285
            Left            =   6360
            TabIndex        =   129
            Top             =   5760
            Width           =   855
         End
         Begin VB.Frame Frame4 
            Height          =   4572
            Index           =   2
            Left            =   6000
            TabIndex        =   127
            Top             =   120
            Width           =   4095
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "«·”« —Ì…"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   15.75
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00008000&
               Height          =   375
               Index           =   2
               Left            =   120
               TabIndex        =   128
               Top             =   3480
               Width           =   3855
            End
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00E2E9E9&
            Caption         =   "Õœœ «·› —…"
            Height          =   1080
            Index           =   2
            Left            =   150
            RightToLeft     =   -1  'True
            TabIndex        =   122
            Top             =   3060
            Width           =   5832
            Begin VB.OptionButton optStatus 
               Caption         =   "«·ﬂ·"
               Height          =   195
               Index           =   2
               Left            =   480
               RightToLeft     =   -1  'True
               TabIndex        =   162
               Top             =   750
               Width           =   765
            End
            Begin VB.OptionButton optStatus 
               Caption         =   "«·« ›«ﬁÌ«  «·„·€«… ›ﬁÿ"
               Height          =   195
               Index           =   1
               Left            =   1830
               RightToLeft     =   -1  'True
               TabIndex        =   161
               Top             =   750
               Width           =   1785
            End
            Begin VB.OptionButton optStatus 
               Caption         =   "«·« ›«ﬁÌ«  «·„›⁄·… ›ﬁÿ"
               Height          =   195
               Index           =   0
               Left            =   3750
               RightToLeft     =   -1  'True
               TabIndex        =   160
               Top             =   750
               Value           =   -1  'True
               Width           =   1785
            End
            Begin MSComCtl2.DTPicker toDate4 
               Height          =   330
               Left            =   330
               TabIndex        =   123
               Top             =   210
               Width           =   1605
               _ExtentX        =   2831
               _ExtentY        =   582
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   93454337
               CurrentDate     =   41640
            End
            Begin MSComCtl2.DTPicker Fromdate4 
               Height          =   330
               Left            =   2730
               TabIndex        =   124
               Top             =   210
               Width           =   1665
               _ExtentX        =   2937
               _ExtentY        =   582
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   93454337
               CurrentDate     =   41640
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "≈·Ï"
               Height          =   435
               Index           =   23
               Left            =   1980
               RightToLeft     =   -1  'True
               TabIndex        =   126
               Top             =   270
               Width           =   540
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "„‰"
               Height          =   435
               Index           =   22
               Left            =   4320
               RightToLeft     =   -1  'True
               TabIndex        =   125
               Top             =   240
               Width           =   540
            End
         End
         Begin VB.TextBox Text10 
            Alignment       =   2  'Center
            Height          =   312
            Left            =   240
            TabIndex        =   121
            Top             =   840
            Visible         =   0   'False
            Width           =   3804
         End
         Begin VB.TextBox TxtSearchCode 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   1
            Left            =   3312
            RightToLeft     =   -1  'True
            TabIndex        =   120
            Top             =   1200
            Width           =   732
         End
         Begin VB.TextBox Text4 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   3312
            RightToLeft     =   -1  'True
            TabIndex        =   119
            Top             =   1200
            Visible         =   0   'False
            Width           =   732
         End
         Begin ImpulseButton.ISButton ISButton2 
            Height          =   492
            Left            =   120
            TabIndex        =   131
            Top             =   4200
            Width           =   5808
            _ExtentX        =   10239
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
         Begin MSDataListLib.DataCombo DcCustmer 
            Bindings        =   "FrmReportsStudent.frx":0724
            Height          =   315
            Index           =   1
            Left            =   240
            TabIndex        =   132
            Top             =   1200
            Width           =   3060
            _ExtentX        =   5398
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
         Begin MSDataListLib.DataCombo DcbEmployee22 
            Bindings        =   "FrmReportsStudent.frx":0739
            Height          =   315
            Index           =   1
            Left            =   240
            TabIndex        =   133
            Top             =   1200
            Visible         =   0   'False
            Width           =   3060
            _ExtentX        =   5398
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
         Begin MSDataListLib.DataCombo DcboUsers 
            Height          =   315
            Index           =   1
            Left            =   210
            TabIndex        =   156
            Top             =   1590
            Width           =   3795
            _ExtentX        =   6694
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„” Œœ„"
            Height          =   375
            Index           =   34
            Left            =   3960
            RightToLeft     =   -1  'True
            TabIndex        =   157
            Top             =   1620
            Width           =   1275
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Caption         =   "Ì—ÃÏ «Œ Ì«— «·›—⁄ «Ê «· «—ÌŒ «Ê ”Ê› ÌﬂÊ‰ «· ﬁ—Ì— «Ã„«·Ì ·ﬂ· «·›—Ê⁄  Ê«·„œ…"
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
            Height          =   450
            Index           =   30
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   138
            Top             =   6240
            Width           =   6975
         End
         Begin VB.Shape Shape4 
            BorderWidth     =   2
            Height          =   495
            Left            =   0
            Top             =   6000
            Width           =   6975
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
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
            Height          =   795
            Index           =   29
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   137
            Top             =   4740
            Width           =   5760
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "»‰«∆« ⁄·Ï « ›«ﬁÌ…"
            Height          =   276
            Index           =   28
            Left            =   4032
            TabIndex        =   136
            Top             =   1200
            Width           =   1608
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·ÿ·»"
            Height          =   276
            Index           =   25
            Left            =   4440
            TabIndex        =   135
            Top             =   840
            Visible         =   0   'False
            Width           =   888
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„ÊŸ›"
            Height          =   276
            Index           =   24
            Left            =   4320
            TabIndex        =   134
            Top             =   1200
            Visible         =   0   'False
            Width           =   1128
         End
      End
      Begin VB.Frame Fra 
         BackColor       =   &H00E2E9E9&
         Caption         =   "{"
         Height          =   5580
         Index           =   2
         Left            =   13665
         RightToLeft     =   -1  'True
         TabIndex        =   85
         Top             =   45
         Width           =   12630
         Begin VB.TextBox Txt_OrderNumber2 
            Alignment       =   2  'Center
            Height          =   312
            Left            =   240
            TabIndex        =   139
            Top             =   750
            Visible         =   0   'False
            Width           =   3804
         End
         Begin VB.TextBox TxtSearchCode 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   0
            Left            =   3312
            RightToLeft     =   -1  'True
            TabIndex        =   116
            Top             =   1470
            Visible         =   0   'False
            Width           =   732
         End
         Begin VB.TextBox TxtEmpCode 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   3312
            RightToLeft     =   -1  'True
            TabIndex        =   115
            Top             =   1470
            Visible         =   0   'False
            Width           =   732
         End
         Begin VB.TextBox Txt_OrderNumber 
            Alignment       =   2  'Center
            Height          =   312
            Left            =   240
            TabIndex        =   112
            Top             =   1110
            Visible         =   0   'False
            Width           =   3804
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00E2E9E9&
            Caption         =   "Õœœ «·› —…"
            Height          =   1710
            Index           =   1
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   90
            Top             =   3150
            Width           =   5832
            Begin VB.OptionButton optStatus 
               Caption         =   "«·« ›«ﬁÌ«  «·„›⁄·… ›ﬁÿ"
               Height          =   195
               Index           =   8
               Left            =   3660
               RightToLeft     =   -1  'True
               TabIndex        =   168
               Top             =   1260
               Value           =   -1  'True
               Width           =   1785
            End
            Begin VB.OptionButton optStatus 
               Caption         =   "«·« ›«ﬁÌ«  «·„·€«… ›ﬁÿ"
               Height          =   195
               Index           =   7
               Left            =   1740
               RightToLeft     =   -1  'True
               TabIndex        =   167
               Top             =   1260
               Width           =   1785
            End
            Begin VB.OptionButton optStatus 
               Caption         =   "«·ﬂ·"
               Height          =   195
               Index           =   6
               Left            =   390
               RightToLeft     =   -1  'True
               TabIndex        =   166
               Top             =   1260
               Width           =   765
            End
            Begin MSComCtl2.DTPicker Fromdate3 
               Height          =   330
               Left            =   2640
               TabIndex        =   91
               Top             =   240
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   582
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   93454337
               CurrentDate     =   41640
            End
            Begin MSComCtl2.DTPicker toDate3 
               Height          =   336
               Left            =   240
               TabIndex        =   92
               Top             =   240
               Width           =   1752
               _ExtentX        =   3096
               _ExtentY        =   582
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   93454337
               CurrentDate     =   41640
            End
            Begin Dynamic_Byte.NourHijriCal toDateH3 
               Height          =   330
               Left            =   240
               TabIndex        =   93
               Top             =   600
               Visible         =   0   'False
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   582
            End
            Begin Dynamic_Byte.NourHijriCal FromdateH3 
               Height          =   330
               Left            =   2655
               TabIndex        =   94
               Top             =   600
               Visible         =   0   'False
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   582
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "„‰"
               Height          =   435
               Index           =   8
               Left            =   4320
               RightToLeft     =   -1  'True
               TabIndex        =   96
               Top             =   270
               Width           =   540
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "≈·Ï"
               Height          =   435
               Index           =   7
               Left            =   1980
               RightToLeft     =   -1  'True
               TabIndex        =   95
               Top             =   270
               Width           =   540
            End
         End
         Begin VB.Frame Frame4 
            Height          =   5415
            Index           =   1
            Left            =   6000
            TabIndex        =   88
            Top             =   120
            Width           =   4095
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "«·”« —Ì…"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   15.75
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00008000&
               Height          =   375
               Index           =   1
               Left            =   120
               TabIndex        =   89
               Top             =   3480
               Width           =   3855
            End
         End
         Begin VB.TextBox Text8 
            Height          =   285
            Left            =   6360
            TabIndex        =   87
            Top             =   5760
            Width           =   855
         End
         Begin VB.TextBox Text7 
            Height          =   285
            Left            =   5400
            TabIndex        =   86
            Top             =   6000
            Width           =   855
         End
         Begin ImpulseButton.ISButton ISButton1 
            Height          =   495
            Left            =   120
            TabIndex        =   97
            Top             =   4980
            Width           =   5805
            _ExtentX        =   10239
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
         Begin XtremeSuiteControls.RadioButton Rd 
            Height          =   372
            Index           =   4
            Left            =   3000
            TabIndex        =   100
            Top             =   240
            Width           =   1572
            _Version        =   786432
            _ExtentX        =   2773
            _ExtentY        =   656
            _StockProps     =   79
            Caption         =   " ﬁ«—Ì— «·«⁄„«· «·ÌÊ„Ì…"
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton Rd 
            Height          =   372
            Index           =   3
            Left            =   4560
            TabIndex        =   101
            Top             =   240
            Width           =   1212
            _Version        =   786432
            _ExtentX        =   2143
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "ÿ»«⁄… «·„⁄«ÌÌ—"
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcCustmer 
            Bindings        =   "FrmReportsStudent.frx":074E
            Height          =   315
            Index           =   0
            Left            =   240
            TabIndex        =   102
            Top             =   1470
            Visible         =   0   'False
            Width           =   3060
            _ExtentX        =   5398
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
         Begin MSDataListLib.DataCombo DCombo1 
            Bindings        =   "FrmReportsStudent.frx":0763
            Height          =   315
            Left            =   240
            TabIndex        =   104
            Top             =   1800
            Visible         =   0   'False
            Width           =   3840
            _ExtentX        =   6773
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
         Begin MSDataListLib.DataCombo DCombo2 
            Bindings        =   "FrmReportsStudent.frx":0778
            Height          =   315
            Left            =   240
            TabIndex        =   106
            Top             =   2130
            Visible         =   0   'False
            Width           =   3840
            _ExtentX        =   6773
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
         Begin MSDataListLib.DataCombo DCombo3 
            Bindings        =   "FrmReportsStudent.frx":078D
            Height          =   315
            Left            =   240
            TabIndex        =   108
            Top             =   2490
            Visible         =   0   'False
            Width           =   3840
            _ExtentX        =   6773
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
         Begin XtremeSuiteControls.RadioButton Rd 
            Height          =   372
            Index           =   5
            Left            =   1560
            TabIndex        =   110
            Top             =   240
            Width           =   1452
            _Version        =   786432
            _ExtentX        =   2561
            _ExtentY        =   656
            _StockProps     =   79
            Caption         =   " ﬁ—Ì— Õ—ﬂ… «·ÿ·»« "
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcbEmployee2 
            Bindings        =   "FrmReportsStudent.frx":07A2
            Height          =   315
            Left            =   240
            TabIndex        =   113
            Top             =   1470
            Visible         =   0   'False
            Width           =   3060
            _ExtentX        =   5398
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
         Begin XtremeSuiteControls.RadioButton Rd 
            Height          =   372
            Index           =   6
            Left            =   120
            TabIndex        =   117
            Top             =   240
            Width           =   1332
            _Version        =   786432
            _ExtentX        =   2350
            _ExtentY        =   656
            _StockProps     =   79
            Caption         =   " ﬁ—Ì— —›⁄ «·ﬁÌ«”"
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcboUsers 
            Height          =   315
            Index           =   0
            Left            =   240
            TabIndex        =   154
            Top             =   2850
            Visible         =   0   'False
            Width           =   3825
            _ExtentX        =   6747
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„” Œœ„"
            Height          =   375
            Index           =   65
            Left            =   4020
            RightToLeft     =   -1  'True
            TabIndex        =   155
            Top             =   2880
            Visible         =   0   'False
            Width           =   1275
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·Õ—ﬂ…"
            Height          =   270
            Index           =   26
            Left            =   4440
            TabIndex        =   140
            Top             =   750
            Visible         =   0   'False
            Width           =   885
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„ÊŸ›"
            Height          =   270
            Index           =   21
            Left            =   4320
            TabIndex        =   114
            Top             =   1470
            Visible         =   0   'False
            Width           =   1125
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·ÿ·»"
            Height          =   270
            Index           =   20
            Left            =   4440
            TabIndex        =   111
            Top             =   1110
            Visible         =   0   'False
            Width           =   885
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·„⁄·„"
            Height          =   255
            Index           =   6
            Left            =   4080
            RightToLeft     =   -1  'True
            TabIndex        =   109
            Top             =   2520
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·›Ê—„«‰"
            Height          =   375
            Index           =   19
            Left            =   3960
            RightToLeft     =   -1  'True
            TabIndex        =   107
            Top             =   2160
            Visible         =   0   'False
            Width           =   1845
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·⁄«„·"
            Height          =   255
            Index           =   18
            Left            =   4320
            RightToLeft     =   -1  'True
            TabIndex        =   105
            Top             =   1800
            Visible         =   0   'False
            Width           =   1050
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "»‰«∆« ⁄·Ï « ›«ﬁÌ…"
            Height          =   270
            Index           =   17
            Left            =   4245
            TabIndex        =   103
            Top             =   1470
            Visible         =   0   'False
            Width           =   1605
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
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
            Height          =   795
            Index           =   16
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   99
            Top             =   4740
            Width           =   5760
         End
         Begin VB.Shape Shape3 
            BorderWidth     =   2
            Height          =   495
            Left            =   0
            Top             =   6000
            Width           =   6975
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Caption         =   "Ì—ÃÏ «Œ Ì«— «·›—⁄ «Ê «· «—ÌŒ «Ê ”Ê› ÌﬂÊ‰ «· ﬁ—Ì— «Ã„«·Ì ·ﬂ· «·›—Ê⁄  Ê«·„œ…"
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
            Height          =   450
            Index           =   15
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   98
            Top             =   6240
            Width           =   6975
         End
      End
      Begin VB.Frame Fra 
         BackColor       =   &H00E2E9E9&
         Height          =   5580
         Index           =   1
         Left            =   13365
         RightToLeft     =   -1  'True
         TabIndex        =   50
         Top             =   45
         Width           =   12630
         Begin VB.Frame Frame1 
            BackColor       =   &H00E2E9E9&
            Height          =   600
            Index           =   0
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   80
            Top             =   2400
            Width           =   5235
            Begin XtremeSuiteControls.RadioButton Rd 
               Height          =   375
               Index           =   0
               Left            =   3840
               TabIndex        =   81
               Top             =   120
               Width           =   1215
               _Version        =   786432
               _ExtentX        =   2143
               _ExtentY        =   661
               _StockProps     =   79
               Caption         =   "«·Õ÷Ê—"
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton Rd 
               Height          =   375
               Index           =   1
               Left            =   1800
               TabIndex        =   82
               Top             =   120
               Width           =   1215
               _Version        =   786432
               _ExtentX        =   2143
               _ExtentY        =   661
               _StockProps     =   79
               Caption         =   "«·€Ì«»"
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton Rd 
               Height          =   375
               Index           =   2
               Left            =   120
               TabIndex        =   83
               Top             =   120
               Width           =   1215
               _Version        =   786432
               _ExtentX        =   2143
               _ExtentY        =   661
               _StockProps     =   79
               Caption         =   "«·ﬂ·"
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
         End
         Begin VB.TextBox TxtSudCode2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3720
            TabIndex        =   63
            Top             =   2040
            Width           =   1080
         End
         Begin VB.TextBox Text6 
            Height          =   285
            Left            =   5400
            TabIndex        =   62
            Top             =   6000
            Width           =   855
         End
         Begin VB.TextBox Text5 
            Height          =   285
            Left            =   6360
            TabIndex        =   61
            Top             =   5760
            Width           =   855
         End
         Begin VB.Frame Frame4 
            Height          =   4215
            Index           =   0
            Left            =   6000
            TabIndex        =   59
            Top             =   120
            Width           =   4095
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "«·”« —Ì…"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   15.75
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00008000&
               Height          =   375
               Index           =   0
               Left            =   120
               TabIndex        =   60
               Top             =   3480
               Width           =   3855
            End
            Begin VB.Image Image2 
               Height          =   2790
               Left            =   0
               Picture         =   "FrmReportsStudent.frx":07B7
               Stretch         =   -1  'True
               Top             =   120
               Width           =   4020
            End
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00E2E9E9&
            Caption         =   "Õœœ «·› —…"
            Height          =   1080
            Index           =   0
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   52
            Top             =   3000
            Width           =   5235
            Begin MSComCtl2.DTPicker Fromdate2 
               Height          =   330
               Left            =   2655
               TabIndex        =   53
               Top             =   240
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   582
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   93454337
               CurrentDate     =   41640
            End
            Begin MSComCtl2.DTPicker toDate2 
               Height          =   330
               Left            =   240
               TabIndex        =   54
               Top             =   240
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   582
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   93454337
               CurrentDate     =   41640
            End
            Begin Dynamic_Byte.NourHijriCal toDateH2 
               Height          =   330
               Left            =   240
               TabIndex        =   55
               Top             =   600
               Visible         =   0   'False
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   582
            End
            Begin Dynamic_Byte.NourHijriCal FromdateH2 
               Height          =   330
               Left            =   2655
               TabIndex        =   56
               Top             =   600
               Visible         =   0   'False
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   582
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "≈·Ï"
               Height          =   435
               Index           =   10
               Left            =   1980
               RightToLeft     =   -1  'True
               TabIndex        =   58
               Top             =   480
               Width           =   540
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "„‰"
               Height          =   435
               Index           =   9
               Left            =   4320
               RightToLeft     =   -1  'True
               TabIndex        =   57
               Top             =   480
               Width           =   540
            End
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3720
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   51
            Top             =   600
            Width           =   1080
         End
         Begin MSDataListLib.DataCombo groupDBox2 
            Height          =   315
            Left            =   120
            TabIndex        =   64
            Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«— √”„ «·„” «Ã—"
            Top             =   960
            Width           =   4680
            _ExtentX        =   8255
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo cursBox2 
            Height          =   315
            Left            =   120
            TabIndex        =   65
            Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«— √”„ «·„” «Ã—"
            Top             =   1320
            Width           =   4680
            _ExtentX        =   8255
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcbCompany2 
            Height          =   315
            Left            =   120
            TabIndex        =   66
            Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«— √”„ «·„” «Ã—"
            Top             =   600
            Width           =   3480
            _ExtentX        =   6138
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo instruDBox2 
            Height          =   315
            Left            =   120
            TabIndex        =   67
            Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«— √”„ «·„” «Ã—"
            Top             =   1680
            Width           =   4680
            _ExtentX        =   8255
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcbStudent2 
            Height          =   315
            Left            =   120
            TabIndex        =   68
            Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«— √”„ «·„” «Ã—"
            Top             =   2040
            Width           =   3480
            _ExtentX        =   6138
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcbBranch2 
            Height          =   315
            Left            =   120
            TabIndex        =   69
            Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«— √”„ «·„” «Ã—"
            Top             =   240
            Width           =   4680
            _ExtentX        =   8255
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin ImpulseButton.ISButton CmdPrint 
            Height          =   495
            Left            =   360
            TabIndex        =   79
            Top             =   4200
            Width           =   4245
            _ExtentX        =   7488
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
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·„ œ—»"
            Height          =   300
            Index           =   11
            Left            =   5040
            TabIndex        =   77
            Top             =   2040
            Width           =   630
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·„œ—»"
            Height          =   300
            Index           =   10
            Left            =   5040
            RightToLeft     =   -1  'True
            TabIndex        =   76
            Top             =   1680
            Width           =   630
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·„«œ…"
            Height          =   300
            Index           =   9
            Left            =   5040
            RightToLeft     =   -1  'True
            TabIndex        =   75
            Top             =   1320
            Width           =   630
         End
         Begin VB.Label gDBox 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·„Ã„Ê⁄…"
            Height          =   300
            Index           =   0
            Left            =   5040
            RightToLeft     =   -1  'True
            TabIndex        =   74
            Top             =   960
            Width           =   630
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Caption         =   "Ì—ÃÏ «Œ Ì«— «·›—⁄ «Ê «· «—ÌŒ «Ê ”Ê› ÌﬂÊ‰ «· ﬁ—Ì— «Ã„«·Ì ·ﬂ· «·›—Ê⁄  Ê«·„œ…"
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
            Height          =   450
            Index           =   13
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   73
            Top             =   6240
            Width           =   6975
         End
         Begin VB.Shape Shape1 
            BorderWidth     =   2
            Height          =   495
            Left            =   0
            Top             =   6000
            Width           =   6975
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·‘—ﬂ…"
            Height          =   300
            Index           =   8
            Left            =   5040
            RightToLeft     =   -1  'True
            TabIndex        =   72
            Top             =   600
            Width           =   750
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
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
            Height          =   795
            Index           =   12
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   71
            Top             =   4740
            Width           =   5760
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·›—⁄"
            Height          =   300
            Index           =   7
            Left            =   5160
            RightToLeft     =   -1  'True
            TabIndex        =   70
            Top             =   240
            Width           =   630
         End
      End
      Begin VB.Frame Fra 
         BackColor       =   &H00E2E9E9&
         Height          =   5580
         Index           =   0
         Left            =   45
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   45
         Width           =   12630
         Begin VB.Frame AttFrame 
            BackColor       =   &H00E2E9E9&
            Caption         =   "Õœœ «·› —…"
            Height          =   1080
            Left            =   240
            RightToLeft     =   -1  'True
            TabIndex        =   44
            Top             =   2520
            Width           =   5235
            Begin VB.ComboBox CmbMonth 
               Height          =   315
               Left            =   2760
               RightToLeft     =   -1  'True
               TabIndex        =   46
               Top             =   480
               Width           =   1695
            End
            Begin VB.ComboBox CboYear 
               Height          =   315
               Left            =   240
               RightToLeft     =   -1  'True
               TabIndex        =   45
               Top             =   480
               Width           =   1695
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "”‰…"
               Height          =   315
               Index           =   5
               Left            =   1860
               RightToLeft     =   -1  'True
               TabIndex        =   48
               Top             =   480
               Width           =   540
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "‘Â—"
               Height          =   315
               Index           =   6
               Left            =   4440
               RightToLeft     =   -1  'True
               TabIndex        =   47
               Top             =   480
               Width           =   540
            End
         End
         Begin VB.TextBox Text15 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3720
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   30
            Top             =   960
            Width           =   1080
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3720
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   29
            Top             =   1320
            Width           =   1080
         End
         Begin VB.Frame Frame8 
            BackColor       =   &H00E2E9E9&
            Caption         =   "Õœœ «·› —…"
            Height          =   1080
            Left            =   240
            RightToLeft     =   -1  'True
            TabIndex        =   22
            Top             =   2520
            Width           =   5235
            Begin MSComCtl2.DTPicker Fromdate 
               Height          =   330
               Left            =   2655
               TabIndex        =   23
               Top             =   240
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   582
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   93454337
               CurrentDate     =   41640
            End
            Begin MSComCtl2.DTPicker toDate 
               Height          =   330
               Left            =   240
               TabIndex        =   24
               Top             =   240
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   582
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   93454337
               CurrentDate     =   41640
            End
            Begin Dynamic_Byte.NourHijriCal toDateH 
               Height          =   330
               Left            =   240
               TabIndex        =   25
               Top             =   600
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   582
            End
            Begin Dynamic_Byte.NourHijriCal FromdateH 
               Height          =   330
               Left            =   2655
               TabIndex        =   26
               Top             =   600
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   582
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "„‰"
               Height          =   435
               Index           =   3
               Left            =   4320
               RightToLeft     =   -1  'True
               TabIndex        =   28
               Top             =   480
               Width           =   540
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "≈·Ï"
               Height          =   435
               Index           =   14
               Left            =   1980
               RightToLeft     =   -1  'True
               TabIndex        =   27
               Top             =   480
               Width           =   540
            End
         End
         Begin VB.Frame Frame3 
            Height          =   4215
            Left            =   6000
            TabIndex        =   20
            Top             =   120
            Width           =   4095
            Begin VB.Image Image1 
               Height          =   2790
               Left            =   0
               Picture         =   "FrmReportsStudent.frx":21715
               Stretch         =   -1  'True
               Top             =   120
               Visible         =   0   'False
               Width           =   4020
            End
            Begin VB.Label lblCompanyname 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "«·”« —Ì…"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   15.75
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00008000&
               Height          =   375
               Left            =   120
               TabIndex        =   21
               Top             =   3480
               Width           =   3855
            End
         End
         Begin VB.TextBox txtCodeBranch 
            Height          =   285
            Left            =   6360
            TabIndex        =   19
            Top             =   5760
            Width           =   855
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   5400
            TabIndex        =   18
            Top             =   6000
            Width           =   855
         End
         Begin VB.OptionButton callsRB 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "« ’«·« "
            Height          =   255
            Left            =   4680
            MaskColor       =   &H80000006&
            RightToLeft     =   -1  'True
            TabIndex        =   17
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton AttRB 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " ﬁ—Ì— «·„ƒ””Â"
            Height          =   255
            Left            =   3120
            RightToLeft     =   -1  'True
            TabIndex        =   16
            Top             =   240
            Width           =   1335
         End
         Begin VB.OptionButton StuInfoRB 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„⁄·Ê„«  «·ÿ·«»"
            Height          =   255
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   13
            Top             =   240
            Width           =   1335
         End
         Begin VB.OptionButton ComRep 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " ﬁ—Ì— «·‘—ﬂ«  "
            Height          =   255
            Left            =   1560
            RightToLeft     =   -1  'True
            TabIndex        =   12
            Top             =   240
            Width           =   1335
         End
         Begin VB.TextBox TxtSudCode 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3720
            TabIndex        =   11
            Top             =   1680
            Width           =   1080
         End
         Begin MSDataListLib.DataCombo cursBox 
            Height          =   315
            Left            =   120
            TabIndex        =   14
            Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«— √”„ «·„” «Ã—"
            Top             =   960
            Width           =   4680
            _ExtentX        =   8255
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo groupDBox 
            Height          =   315
            Left            =   120
            TabIndex        =   15
            Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«— √”„ «·„” «Ã—"
            Top             =   600
            Width           =   4680
            _ExtentX        =   8255
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcbCompany 
            Height          =   315
            Left            =   120
            TabIndex        =   31
            Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«— √”„ «·„” «Ã—"
            Top             =   960
            Width           =   3480
            _ExtentX        =   6138
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcbEmployee 
            Height          =   315
            Left            =   120
            TabIndex        =   32
            Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«— √”„ «·„” «Ã—"
            Top             =   1320
            Width           =   3480
            _ExtentX        =   6138
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo instruDBox 
            Height          =   315
            Left            =   120
            TabIndex        =   33
            Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«— √”„ «·„” «Ã—"
            Top             =   2040
            Width           =   4680
            _ExtentX        =   8255
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcbStudent 
            Height          =   315
            Left            =   120
            TabIndex        =   34
            Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«— √”„ «·„” «Ã—"
            Top             =   1680
            Width           =   3480
            _ExtentX        =   6138
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcbBranch 
            Height          =   315
            Left            =   120
            TabIndex        =   49
            Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«— √”„ «·„” «Ã—"
            Top             =   600
            Width           =   4680
            _ExtentX        =   8255
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   732
            Index           =   1
            Left            =   600
            TabIndex        =   78
            Top             =   3600
            Width           =   4248
            _ExtentX        =   7488
            _ExtentY        =   1296
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
         Begin MSDataListLib.DataCombo DcbUserName 
            Height          =   315
            Left            =   0
            TabIndex        =   84
            Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«— √”„ «·„” «Ã—"
            Top             =   4320
            Visible         =   0   'False
            Width           =   4680
            _ExtentX        =   8255
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·„ ’·"
            Height          =   300
            Index           =   3
            Left            =   5040
            RightToLeft     =   -1  'True
            TabIndex        =   43
            Top             =   1320
            Width           =   630
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·›—⁄"
            Height          =   300
            Index           =   1
            Left            =   5160
            RightToLeft     =   -1  'True
            TabIndex        =   42
            Top             =   600
            Width           =   636
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
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
            Height          =   795
            Index           =   11
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   41
            Top             =   4740
            Width           =   5760
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·‘—ﬂ…"
            Height          =   300
            Index           =   5
            Left            =   4920
            RightToLeft     =   -1  'True
            TabIndex        =   40
            Top             =   960
            Width           =   750
         End
         Begin VB.Shape Shape2 
            BorderWidth     =   2
            Height          =   495
            Left            =   0
            Top             =   6000
            Width           =   6975
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Caption         =   "Ì—ÃÏ «Œ Ì«— «·›—⁄ «Ê «· «—ÌŒ «Ê ”Ê› ÌﬂÊ‰ «· ﬁ—Ì— «Ã„«·Ì ·ﬂ· «·›—Ê⁄  Ê«·„œ…"
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
            Height          =   450
            Index           =   4
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   39
            Top             =   6240
            Width           =   6975
         End
         Begin VB.Label gDBox 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·„Ã„Ê⁄…"
            Height          =   300
            Index           =   2
            Left            =   5040
            RightToLeft     =   -1  'True
            TabIndex        =   38
            Top             =   600
            Width           =   630
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·„«œ…"
            Height          =   300
            Index           =   0
            Left            =   5040
            RightToLeft     =   -1  'True
            TabIndex        =   37
            Top             =   960
            Width           =   630
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·„œ—»"
            Height          =   300
            Index           =   2
            Left            =   5040
            RightToLeft     =   -1  'True
            TabIndex        =   36
            Top             =   2040
            Width           =   630
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·„ œ—»"
            Height          =   300
            Index           =   4
            Left            =   5040
            TabIndex        =   35
            Top             =   1680
            Width           =   630
         End
      End
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "‘«‘…  ﬁ«—Ì— "
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
      Left            =   -360
      TabIndex        =   2
      Top             =   0
      Width           =   18735
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      ForeColor       =   &H00000080&
      Height          =   285
      Index           =   1
      Left            =   60
      TabIndex        =   1
      Top             =   3060
      Width           =   1785
   End
End
Attribute VB_Name = "FrmReportsStudent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim amoutId As Integer
Dim RsSavRec As ADODB.Recordset
Dim BKGrndPic As ClsBackGroundPic
Dim RecId As String
Dim II As Long
Dim cSearch  As clsDCboSearch
Dim rs As ADODB.Recordset
Dim DCboSearch As clsDCboSearch
Public Indx As Integer
Dim Dcombos As New ClsDataCombos


Private Sub lbl_Click(Index As Integer)
Select Case Index
Case 139

FrmSelectEmployee.show
FrmSelectEmployee.lblflag = 0

Case 140

FrmSelectEmployee.show
FrmSelectEmployee.lblflag = 0

Case 141

FrmSelectEmployee.show
FrmSelectEmployee.lblflag = 0
Case 142
 
FrmSelectEmployee.show
FrmSelectEmployee.lblflag = 0
Case 149
'
FrmSelectEmployee.show
FrmSelectEmployee.lblflag = 0
Case 150
'
FrmSelectEmployee.show
FrmSelectEmployee.lblflag = 0
Case 151
 FrmSelectEmployee.supplierVendor = 1
 FrmSelectEmployee.show
 FrmSelectEmployee.lblflag = 0


End Select
End Sub

Sub Relaod2(Optional GroupID As Double = 0)
Dim StrSQL As String
   If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT ID, Name From TblStudentCurs "
    Else
        StrSQL = "SELECT ID, NameE From TblStudentCurs "
    End If
    StrSQL = StrSQL & "   where id in(SELECT CursID from TblStuGroupDet where StudGrouID=" & GroupID & ")"
    If SystemOptions.UserInterface = ArabicInterface Then
    StrSQL = StrSQL & " order by Name "
    Else
    StrSQL = StrSQL & " order By NameE "
    End If
     fill_combo cursBox2, StrSQL
 End Sub
Sub Relaod(Optional GroupID As Double = 0)
Dim StrSQL As String
   If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "SELECT ID, Name From TblStudentCurs "
    Else
        StrSQL = "SELECT ID, NameE From TblStudentCurs "
    End If
    StrSQL = StrSQL & "   where id in(SELECT CursID from TblStuGroupDet where StudGrouID=" & GroupID & ")"
    If SystemOptions.UserInterface = ArabicInterface Then
    StrSQL = StrSQL & " order by Name "
    Else
    StrSQL = StrSQL & " order By NameE "
    End If
     fill_combo cursBox, StrSQL
 End Sub

Private Sub AttRB_Click()
'***************************************
TxtSudCode.Visible = False
DcbStudent.Visible = False
Label1(4).Visible = False
groupDBox.Visible = True
cursBox.Visible = True
instruDBox.Visible = True
Label1(0).Visible = True
Label1(2).Visible = True
gDBox(2).Visible = True
Frame8.Visible = True
DcbBranch.Visible = True
Label1(1).Visible = True
Label1(5).Visible = True
Text15.Visible = True
DcbCompany.Visible = True
AttFrame.Visible = True
instruDBox.Visible = False
Label1(2).Visible = False
DcbCompany.Visible = False
Label1(5).Visible = False
DcbBranch.Visible = False
Label1(1).Visible = False
DcbEmployee.Visible = False
Text1.Visible = False
'***************************************
Label1(3).Visible = False
Frame8.Visible = False
Text15.Visible = False

End Sub

Private Sub btnClear_Click()
Cmd_Click (7)
End Sub
Private Sub callsRB_Click()
'***************************************
groupDBox.Visible = False
cursBox.Visible = False
instruDBox.Visible = False
Label1(0).Visible = False
Label1(2).Visible = False
gDBox(2).Visible = False
AttFrame.Visible = False
'***************************************
TxtSudCode.Visible = True
DcbStudent.Visible = True
Label1(4).Visible = True
Label1(3).Visible = True
DcbBranch.Visible = True
Text15.Visible = True
DcbCompany.Visible = True
Text1.Visible = True
DcbEmployee.Visible = True
Label1(1).Visible = True
Label1(5).Visible = True
Frame8.Visible = True

End Sub

Public Function MonthLastDay(ByVal dCurrDate As Date) As Date
    Dim dFirstDayNextMonth As Date
  
    MonthLastDay = Empty
    dCurrDate = Format(dCurrDate, "DD/MM/YYYY")
    dFirstDayNextMonth = DateSerial(CInt(Format(dCurrDate, "yyyy")), CInt(Format(dCurrDate, "mm")) + 1, 1)
    MonthLastDay = DateAdd("d", -1, dFirstDayNextMonth)
  
    Exit Function
 
End Function
Private Sub Cmd_Click(Index As Integer)

    Select Case Index
        Case 1
        
        If callsRB.value = True Then
          GetCallsData
        ElseIf AttRB.value = True Then
        '"""""""""""""""""""""" CHECK """""""""""""""""""""""
          If val(cursBox.BoundText) = 0 Or (cursBox.Text = "") Then
            MsgBox "«·—Ã«¡  ÕœÌœ «·„«œ…", vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            cursBox.SetFocus
            Exit Sub
            ElseIf (val(CmbMonth.ListIndex) = -1) Or (CmbMonth.Text = "") Or (val(CboYear.ListIndex) = -1) Or (CboYear.Text = "") Then
            MsgBox "«·—Ã«¡  ÕœÌœ «·”‰… Ê«·‘Â—", vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            cursBox.SetFocus
            Exit Sub
          End If
        '""""""""""""""""""""""""""""""""""""""""""""""""""
          GetAttData
        ElseIf ComRep.value = True Then
        '"""""""""""""""""""""" CHECK """""""""""""""""""""""
            If val(DcbCompany.BoundText) = 0 Or (DcbCompany.Text = "") Then
                MsgBox "«·—Ã«¡  ÕœÌœ «·‘—ﬂ…", vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                DcbCompany.SetFocus
                Exit Sub

            End If
          Dim DiFFNo As Integer
          If IsNull(FromDate.value) Or IsNull(ToDate.value) Then
          If SystemOptions.UserInterface = ArabicInterface Then
          MsgBox "Ì—ÃÏ  ÕœÌœ «· «—ÌŒ"
          Else
          MsgBox "Please Select Date"
          End If
          Exit Sub
          End If
            DiFFNo = DateDiff("d", FromDate.value, ToDate.value) + 1
      If DiFFNo > 31 Or DiFFNo <= 0 Then
      MsgBox "Ì—ÃÏ «· «ﬂœ „‰ «· «—ÌŒ »ÌÕÀ ·«Ì“Ìœ ⁄‰ 31 ÌÊ„"
      Exit Sub
      End If
            FillDatConp
            GetAttData22
        '""""""""""""""""""""""""""""""""""""""""""""""""""
          'GetAttCompData
        Else
          GetStuInfoData
        End If
          
        Case 7
          clear_all Me
         
          Rd(2).value = True
          FromDate.value = Date
          ToDate.value = Date
          FromDate2.value = Date
          ToDate2.value = Date

        Case 2
          Unload Me
        Case 3
          'If callsRB.value = True Then
            'print_report_Calls
          'ElseIf AttRB.value = True Then
            'print_report_Att
          'Else
            'print_report_StuInfo
          'End If
        End Select
End Sub

Private Sub CmdPrint_Click()
StudentInGroup
End Sub

Private Sub ComRep_Click()
'***************************************
DcbEmployee.Visible = False
Text15.Visible = True
TxtSudCode.Visible = False
DcbStudent.Visible = False
Label1(4).Visible = False
groupDBox.Visible = True
cursBox.Visible = False
instruDBox.Visible = True
Label1(0).Visible = False
Label1(2).Visible = True
gDBox(2).Visible = False
groupDBox.Visible = False
Frame8.Visible = True
DcbBranch.Visible = True
Label1(1).Visible = True
Label1(5).Visible = True
Text15.Visible = True
DcbCompany.Visible = True
AttFrame.Visible = False
Text1.Visible = False
TxtSudCode.Visible = False
DcbStudent.Visible = False
Label1(4).Visible = False
Label1(2).Visible = False
instruDBox.Visible = False
'***************************************
Label1(3).Visible = False
Frame8.Visible = True
End Sub

Private Sub DcbCompany_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
   FrmCustemerSearch.SearchType = 26
        FrmCustemerSearch.show vbModal
  End If
End Sub

Private Sub DcbCompany2_Change()
DcbCompany2_Click (0)
End Sub

Private Sub DcbCompany2_Click(Area As Integer)
  If val(DcbCompany2.BoundText) = 0 Then Exit Sub
    Dim EmpCode  As String
    GetTblCustemersCode , , DcbCompany2.BoundText, EmpCode
    Me.Text3.Text = EmpCode
End Sub

Private Sub DcbEmployee2_Change()

 If val(DcbEmployee2.BoundText) = 0 Then Exit Sub


    Dim EmpCode  As String
 
    GetEmployeeIDFromCode , , DcbEmployee2.BoundText, EmpCode
    TxtEmpCode.Text = EmpCode


End Sub

Private Sub DcbStudent_Change()
DcbStudent_Click (0)
End Sub

Private Sub DcbStudent_Click(Area As Integer)
Dim UQama As String
  If val(DcbStudent.BoundText) = 0 Then Exit Sub
    Dim EmpCode  As String
    GetStudentCode val(DcbStudent.BoundText), EmpCode, 0
    Me.TxtSudCode.Text = EmpCode
End Sub

Private Sub DcbStudent_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
FrmSearStudent.inde = 103
Load FrmSearStudent
FrmSearStudent.show vbModal
End If
End Sub

Private Sub DcbStudent2_Change()
DcbStudent2_Click (0)
End Sub

Private Sub DcbStudent2_Click(Area As Integer)
  If val(DcbStudent2.BoundText) = 0 Then Exit Sub
    Dim EmpCode  As String
    GetStudentCode val(DcbStudent2.BoundText), EmpCode, 0
    Me.TxtSudCode2.Text = EmpCode
End Sub

Private Sub Form_Activate()
   PutFormOnTop Me.hwnd
End Sub
Private Sub DcbCompany_Change()
DcbCompany_Click (0)
End Sub

Private Sub DcbCompany_Click(Area As Integer)
  If val(DcbCompany.BoundText) = 0 Then Exit Sub
    Dim EmpCode  As String
    GetTblCustemersCode , , DcbCompany.BoundText, EmpCode
    Me.Text15.Text = EmpCode
 Dim Dcombos As New ClsDataCombos
   Dcombos.GetStudent Me.DcbStudent, 0, val(DcbCompany.BoundText)
End Sub
Private Sub Form_Load()
    Dim i As Integer
    Dim My_SQL As String
    Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos
    Rd(2).value = True

    
    
    XPTab301.TabVisible(0) = False
    XPTab301.TabVisible(1) = False
    XPTab301.TabVisible(2) = False

If Indx = 0 Then
XPTab301.TabVisible(0) = True
XPTab301.TabVisible(1) = True

    XPTab301.TabVisible(2) = True

XPTab301.CurrTab = 0
ElseIf Indx = 1 Then
 
XPTab301.TabVisible(2) = True
XPTab301.CurrTab = 2
End If


   Dcombos.GetCustomersSuppliers 55, Me.DcbCompany
   
   Dcombos.GetEmployees Me.DcbEmployee
   Dcombos.GetBranches Me.DcbBranch
  ' Dcombos.GetStudentCurs Me.cursBox
   Dcombos.GeInstructor Me.instruDBox
   Dcombos.GetStudentGroup Me.groupDBox
   Dcombos.GetUsers DcbUserName
   Dcombos.GetCustomersSuppliers 55, Me.DcbCompany2
   Dcombos.GetBranches Me.DcbBranch2
   Dcombos.GetStudentCurs Me.cursBox2
   Dcombos.GeInstructor Me.instruDBox2
   Dcombos.GetStudentGroup Me.groupDBox2
   Dcombos.GetStudent Me.DcbStudent2
   
    Dcombos.GetCustomersSuppliers 1, Me.DcCustmer(0)
    Dcombos.GetCustomersSuppliers 1, Me.DcCustmer(1)
    
    Dcombos.GetEmployees Me.DcbEmployee2
    Dcombos.GetEmployees Me.DCombo1
    Dcombos.GetEmployees Me.DCombo2
    Dcombos.GetEmployees Me.DCombo3
     Dcombos.GetUsers Me.DcboUsers(0)
    Dcombos.GetUsers Me.DcboUsers(1)
    Dcombos.GetUsers Me.DcboUsers(2)
   Fromdate3.value = ""
   ToDate3.value = ""
   
   
Relaod
DcbUserName.BoundText = user_id
Dim mFromDate  As String
Dim mToDate  As String
mFromDate = "1-1-" & year(Date)
mToDate = "31-12-" & year(Date)

FromDate.value = mFromDate
          ToDate.value = mToDate
       FromDate2.value = mFromDate
          ToDate2.value = mToDate
          
       Fromdate3.value = mToDate
          ToDate3.value = mToDate
          
    Fromdate4.value = mFromDate
    ToDate4.value = mToDate
     Fromdate4.value = Null
    ToDate4.value = Null
     
     Fromdate5.value = mFromDate
     toDate5.value = mToDate
    Set cSearch = New clsDCboSearch
    Set BKGrndPic = New ClsBackGroundPic
    Set RsSavRec = New ADODB.Recordset
    Resize_Form Me
      If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
    callsRB.value = True
    callsRB_Click
    YearMonth
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FormPostion Me, SavePostion
    Set DCboSearch = Nothing
End Sub

Public Sub GetCallsData()
    Dim StrSQL As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
 StrSQL = " SELECT     dbo.TblStudCalling.ID, dbo.TblStudCalling.BranchID, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblStudCalling.RecordDateH, "
 StrSQL = StrSQL & "                      dbo.TblStudCalling.RecordDate, dbo.TblStudCalling.Remarks, dbo.TblStudCalling.EnterDateH, dbo.TblStudCalling.EnterDate, dbo.TblStudCalling.EnterTime,"
 StrSQL = StrSQL & "                     dbo.TblStudCalling.EmpID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Name1, dbo.TblEmployee.Emp_Name2, dbo.TblEmployee.Emp_Name3,"
 StrSQL = StrSQL & "                     dbo.TblEmployee.Emp_Name4, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee4, dbo.TblEmployee.Emp_Namee3, dbo.TblEmployee.Emp_Namee,"
 StrSQL = StrSQL & "                     dbo.TblEmployee.Emp_Namee1, dbo.TblEmployee.Emp_Namee2, dbo.TblStudCalling.CompID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee,"
 StrSQL = StrSQL & "                     dbo.TblCustemers.Fullcode AS CusFullcode, dbo.TblStudCalling.Mobile, dbo.TblStudCalling.Phone, dbo.TblStudCalling.Email, dbo.TblStudCalling.StudID,"
 StrSQL = StrSQL & "                     dbo.TblStudent.Name, dbo.TblStudent.NameE, dbo.TblStudent.FullCode AS StudFullCode, dbo.TblStudent.UQama"
 StrSQL = StrSQL & "  FROM         dbo.TblStudCalling LEFT OUTER JOIN"
 StrSQL = StrSQL & "                     dbo.TblStudent ON dbo.TblStudCalling.StudID = dbo.TblStudent.ID LEFT OUTER JOIN"
 StrSQL = StrSQL & "                     dbo.TblCustemers ON dbo.TblStudCalling.CompID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
 StrSQL = StrSQL & "                     dbo.TblEmployee ON dbo.TblStudCalling.EmpID = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
 StrSQL = StrSQL & "                     dbo.TblBranchesData ON dbo.TblStudCalling.BranchID = dbo.TblBranchesData.branch_id"
 StrSQL = StrSQL & "  where  (dbo.TblStudCalling.BranchID=0 or dbo.TblStudCalling.BranchID is null or dbo.TblStudCalling.BranchID in(" & Current_branchSql & "))"
 'StrSQL = StrSQL & "  where 1=1"
  
    BolBegine = False
    StrWhere = ""
    

If val(Me.DcbBranch.BoundText) <> 0 Or Me.DcbBranch.Text <> "" Then
StrWhere = StrWhere & " AND dbo.TblStudCalling.BranchID = " & val(Me.DcbBranch.BoundText)
End If
If val(Me.DcbEmployee.BoundText) <> 0 Or Me.DcbEmployee.Text <> "" Then
StrWhere = StrWhere & " AND dbo.TblStudCalling.EmpID = " & val(Me.DcbEmployee.BoundText)
End If

If val(Me.DcbStudent.BoundText) <> 0 Or Me.DcbStudent.Text <> "" Then
StrWhere = StrWhere & " AND dbo.TblStudCalling.StudID   = " & val(DcbStudent.BoundText)
End If

If val(Me.DcbCompany.BoundText) <> 0 Or Me.DcbCompany.Text <> "" Then
StrWhere = StrWhere & " AND dbo.TblStudCalling.CompID   = " & val(DcbCompany.BoundText)
End If
   If Not IsNull(Me.FromDate.value) Then
                   StrWhere = StrWhere & " AND dbo.TblStudCalling.RecordDate >=" & SQLDate(Me.FromDate.value, True) & ""
      End If

    If Not IsNull(Me.ToDate.value) Then
            StrWhere = StrWhere & " AND  dbo.TblStudCalling.RecordDate <=" & SQLDate(Me.ToDate.value, True) & ""
    End If
    '---------------------------------
    
    StrSQL = StrSQL & StrWhere
    'StrSQL = StrSQL & " Order By EmpAsID"
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
           ' Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
          '  Me.lbl(10).Caption = "Search Results=0"
        End If
If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «· ﬁ—Ì—"
      Else
     Msg = "No Data"
  End If
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    Else
 rs.MoveFirst

 print_report_Calls StrSQL

    End If

End Sub

Private Sub FromDate_Change()
If Not IsNull(FromDate.value) Then
FromDateH.value = ToHijriDate(FromDate.value)
End If
End Sub

Private Sub Fromdate2_Change()
If Not IsNull(FromDate2.value) Then
FromdateH2.value = ToHijriDate(FromDate2.value)
End If
End Sub

Private Sub Fromdateh_LostFocus()
FromDate.value = ToGregorianDate(FromDateH.value)
End Sub

Private Sub FromdateH2_LostFocus()
FromDate2.value = ToGregorianDate(FromdateH2.value)
End Sub

Private Sub groupDBox_Change()
Relaod val(groupDBox.BoundText)
End Sub

Private Sub groupDBox_Click(Area As Integer)
groupDBox_Change
End Sub

Private Sub groupDBox_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
FrmSearStudent.inde = 702
Load FrmSearStudent
FrmSearStudent.show vbModal
End If
End Sub

Private Sub groupDBox2_Change()
groupDBox2_Click (0)
End Sub

Private Sub groupDBox2_Click(Area As Integer)
Relaod2 val(groupDBox2.BoundText)
End Sub

Private Sub instruDBox_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
FrmSearStudent.inde = 203
Load FrmSearStudent
FrmSearStudent.show vbModal
End If
End Sub

Private Sub ISButton1_Click()
If Rd(3).value = True Then
GetReportMeasurement
ElseIf Rd(4).value = True Then
GetReportBusinessDialy
ElseIf Rd(5).value = True Then
GetReportTransOrder
ElseIf Rd(6).value = True Then
GetReportMeasurementOne
End If

End Sub

Sub GetReportTransOrder()
Dim My_SQL1 As String
Dim StrWhere As String

My_SQL1 = "SELECT dbo.Tbl_TransOrder.ID, dbo.Tbl_TransOrder.TOrder_OrderNum, dbo.Tbl_TransOrder.TOrder_Status, dbo.TblBranchesData.branch_name,"
My_SQL1 = My_SQL1 & "               dbo.TblBranchesData.branch_namee, dbo.Tbl_TransOrder.TOrder_DateOrder,User2.UserId Emp_ID, User2.UserName Emp_Name,"
My_SQL1 = My_SQL1 & "               User2.UserName Emp_Namee, dbo.Tbl_TransOrder.TOrder_Notes, dbo.Tbl_TransOrder.TOrder_DateNote, dbo.Tbl_TransOrder.TOrder_Time,"
My_SQL1 = My_SQL1 & "               dbo.TblUsers.UserID, dbo.TblUsers.UserName, dbo.Tbl_TransOrder.TOrder_StatusDept, dbo.Tbl_TransOrder.TOrder_Days,"
My_SQL1 = My_SQL1 & "               dbo.Tbl_TransOrder.TOrder_BranchID"
My_SQL1 = My_SQL1 & " FROM  dbo.Tbl_TransOrder INNER JOIN"
My_SQL1 = My_SQL1 & "               dbo.TblBranchesData ON dbo.Tbl_TransOrder.TOrder_BranchID = dbo.TblBranchesData.branch_id INNER JOIN"
My_SQL1 = My_SQL1 & "               dbo.TblUsers User2 ON dbo.Tbl_TransOrder.TOrder_EmpID = User2.UserID INNER JOIN"
My_SQL1 = My_SQL1 & "               dbo.TblUsers ON dbo.Tbl_TransOrder.UserID = dbo.TblUsers.UserID"


My_SQL1 = My_SQL1 & "  WHERE (dbo.Tbl_TransOrder.ID is not null OR Tbl_TransOrder.TOrder_BranchID in(" & Current_branchSql & "))"
  
 StrWhere = ""


If val(Me.Txt_OrderNumber2.Text) <> 0 Then
  StrWhere = StrWhere & " And dbo.Tbl_TransOrder.ID = " & val(Txt_OrderNumber2.Text)
End If


If val(Me.Txt_OrderNumber.Text) <> 0 Then
  StrWhere = StrWhere & " And dbo.Tbl_TransOrder.TOrder_OrderNum = " & val(Txt_OrderNumber.Text)
End If

If val(Me.DcbEmployee2.BoundText) <> 0 Then
  StrWhere = StrWhere & " And User2.UserID = " & val(DcbEmployee2.BoundText)
End If

If Not IsNull(Fromdate3.value) Then
  StrWhere = StrWhere & " And dbo.Tbl_TransOrder.TOrder_DateOrder  >= " & SQLDate(Fromdate3.value, True)
End If
If Not IsNull(ToDate3.value) Then
  StrWhere = StrWhere & " AND dbo.Tbl_TransOrder.TOrder_DateOrder  <= " & SQLDate(ToDate3.value, True)
End If


If Me.DcboUsers(0).BoundText <> "" Then
    StrWhere = StrWhere + " and Tbl_TransOrder.UserID= " & Me.DcboUsers(0).BoundText & ""
End If


'StrWhere = StrWhere & " order by TblStuFingerprint.StudID"
  My_SQL1 = My_SQL1 & StrWhere
  print_report_MeasureInfo My_SQL1, 3
End Sub

Sub GetReportMeasurement()
Dim My_SQL1 As String
Dim StrWhere As String


 My_SQL1 = " SELECT dbo.Tbl_TradingContract.ID, dbo.TblCustemers.CusName, dbo.Tbl_TradingContract.TContract_DateH, dbo.Tbl_TradingContract.TContract_Date,"
 My_SQL1 = My_SQL1 & "               dbo.Tbl_TradingContract.TOrder_Address, dbo.Tbl_TradingContract.TOrder_Phone, dbo.Tbl_TradingContractDet.TContractDet_Qun,"
 My_SQL1 = My_SQL1 & "               dbo.Tbl_TradingContractDet.TContractDet_SalPrice, dbo.Tbl_TradingContractDet.TContractDet_InstallPrice, dbo.Tbl_TradingContractDet.TContractDet_Value,"
 My_SQL1 = My_SQL1 & "               dbo.Tbl_TradingContractDet.TContractDet_TotalSalPrice, dbo.Tbl_TradingContractDet.TContractDet_TotalInstallPrice,"
 My_SQL1 = My_SQL1 & "              dbo.Tbl_TradingContractDet.TContractDet_DayMeter, dbo.TblProcessUnites.UnitID, dbo.TblProcessUnites.UnitName, dbo.TblProcessUnites.UnitNamee,"
 My_SQL1 = My_SQL1 & "               dbo.Tbl_TradingContractDet.ProcessDEFID ,dbo.TblProcessDEF.ProcessName TContractDet_specification, dbo.TblProcessDEF.ProcessNameE TContractDet_specificationEn"
 My_SQL1 = My_SQL1 & "  FROM  dbo.Tbl_TradingContract INNER JOIN"
 My_SQL1 = My_SQL1 & "               dbo.Tbl_TradingContractDet ON dbo.Tbl_TradingContract.ID = dbo.Tbl_TradingContractDet.TContractDet_TContractID INNER JOIN"
 My_SQL1 = My_SQL1 & "               dbo.TblCustemers ON dbo.Tbl_TradingContract.TContract_CustID = dbo.TblCustemers.CusID INNER JOIN"
 My_SQL1 = My_SQL1 & "              dbo.TblProcessUnites ON dbo.Tbl_TradingContractDet.TContractDet_UnitID = dbo.TblProcessUnites.UnitID INNER JOIN"
 My_SQL1 = My_SQL1 & "              dbo.TblProcessDEF ON dbo.Tbl_TradingContractDet.ProcessDEFID = dbo.TblProcessDEF.TblProcessDEFID"
 My_SQL1 = My_SQL1 & "  WHERE (dbo.Tbl_TradingContract.ID is not null )"
  
 StrWhere = " Where 1 = 1 "

My_SQL1 = " Select TblProcessDEFID,       dbo.TblProcessDEF.ProcessName      TContractDet_specification,TblProcessDEF.Interval TContractDet_Qun,TblProcessUnites.UnitName,"
My_SQL1 = My_SQL1 & "  dbo.TblProcessDEF.ProcessNameE     TContractDet_specificationEn"
My_SQL1 = My_SQL1 & "  FROM    dbo.TblProcessDEF "
My_SQL1 = My_SQL1 & "  LEFT OUTER JOIN"
My_SQL1 = My_SQL1 & "  TblProcessUnites ON dbo.TblProcessDEF.UnitID = TblProcessUnites.UnitID"

'If Me.DcboUsers(0).BoundText <> "" Then
'    StrWhere = StrWhere + " and Tbl_TradingContract.UserID= " & Me.DcboUsers(0).BoundText & ""
'End If

'StrWhere = StrWhere & " order by TblStuFingerprint.StudID"
  My_SQL1 = My_SQL1 & StrWhere
  print_report_MeasureInfo My_SQL1, 1
End Sub

Sub GetReportBusinessDialy()
Dim My_SQL1 As String




Dim StrWhere As String
  
  
My_SQL1 = " SELECT dbo.Tbl_BusinessDialy.ID, dbo.Tbl_BusinessDialy.BD_Date, dbo.Tbl_BusinessDialy.BD_BranchID, dbo.TblBranchesData.branch_name,"
My_SQL1 = My_SQL1 & "             dbo.TblBranchesData.branch_namee, dbo.Tbl_BusinessDialy.BD_Notes, dbo.Tbl_BusinessDialyDet.BDet_BD_ID, dbo.Tbl_BusinessDialyDet.BDet_BandNo,"
My_SQL1 = My_SQL1 & "   dbo.Tbl_TradingContractDet.TContractDet_TotalSalPrice, dbo.Tbl_TradingContractDet.TContractDet_TotalInstallPrice,"
My_SQL1 = My_SQL1 & "            dbo.Tbl_BusinessDialyDet.BDet_Qun, dbo.Tbl_BusinessDialyDet.BDet_Name, dbo.Tbl_BusinessDialyDet.BDet_NameE, TblEmployee_1.Emp_ID AS EmpID,"
My_SQL1 = My_SQL1 & "            TblEmployee_1.Emp_Name AS EmpName, TblEmployee_1.Emp_Namee AS EmpNameE, dbo.TblEmployee.Emp_ID AS ForemanID,"
My_SQL1 = My_SQL1 & "            dbo.TblEmployee.Emp_Name AS ForemanName, dbo.TblEmployee.Emp_Namee AS ForemanNameE, TblEmployee_2.Emp_ID AS TeacherID,"
My_SQL1 = My_SQL1 & "            TblEmployee_2.Emp_Name AS TeacherName, TblEmployee_2.Emp_Namee AS TeacherNameE, dbo.Tbl_TradingContract.TContract_CustID,"
My_SQL1 = My_SQL1 & "           dbo.TblCustemers.CusName , dbo.TblCustemers.Fullcode, dbo.Tbl_BusinessDialy.TradingContractID "
',dbo.Tbl_BusinessDialyDet.BDet_DayMeter"
My_SQL1 = My_SQL1 & " FROM  dbo.Tbl_BusinessDialy INNER JOIN"
My_SQL1 = My_SQL1 & "             dbo.Tbl_BusinessDialyDet ON dbo.Tbl_BusinessDialy.ID = dbo.Tbl_BusinessDialyDet.BDet_BD_ID INNER JOIN"
My_SQL1 = My_SQL1 & "             dbo.TblBranchesData ON dbo.Tbl_BusinessDialy.BD_BranchID = dbo.TblBranchesData.branch_id INNER JOIN"
My_SQL1 = My_SQL1 & "            dbo.TblEmployee AS TblEmployee_1 ON dbo.Tbl_BusinessDialyDet.BDet_EmpID = TblEmployee_1.Emp_ID INNER JOIN"
My_SQL1 = My_SQL1 & "            dbo.Tbl_TradingContract ON dbo.Tbl_BusinessDialy.TradingContractID = dbo.Tbl_TradingContract.ID INNER JOIN"
My_SQL1 = My_SQL1 & "           dbo.TblCustemers ON dbo.Tbl_TradingContract.TContract_CustID = dbo.TblCustemers.CusID RIGHT OUTER JOIN"
My_SQL1 = My_SQL1 & "           dbo.TblEmployee ON dbo.Tbl_BusinessDialyDet.BDet_EmpFormanID = dbo.TblEmployee.Emp_ID RIGHT OUTER JOIN"
My_SQL1 = My_SQL1 & "           dbo.TblEmployee AS TblEmployee_2 ON dbo.Tbl_BusinessDialyDet.BDet_EmpTecherID = TblEmployee_2.Emp_ID"
My_SQL1 = My_SQL1 & "  WHERE   (dbo.Tbl_BusinessDialy.ID is not null OR dbo.Tbl_BusinessDialy.BD_BranchID in(" & Current_branchSql & "))"
  
 StrWhere = ""

My_SQL1 = " SELECT dbo.Tbl_BusinessDialy.ID, dbo.Tbl_BusinessDialy.BD_Date, dbo.Tbl_BusinessDialy.BD_BranchID, dbo.TblBranchesData.branch_name,"
My_SQL1 = My_SQL1 & "             dbo.TblBranchesData.branch_namee, dbo.Tbl_BusinessDialy.BD_Notes, dbo.Tbl_BusinessDialyDet.BDet_BD_ID, dbo.Tbl_BusinessDialyDet.BDet_BandNo,"
My_SQL1 = My_SQL1 & "            dbo.Tbl_BusinessDialyDet.BDet_Qun, dbo.Tbl_BusinessDialyDet.BDet_Name, dbo.Tbl_BusinessDialyDet.BDet_NameE, TblEmployee_1.Emp_ID AS EmpID,"
My_SQL1 = My_SQL1 & "            TblEmployee_1.Emp_Name AS EmpName, TblEmployee_1.Emp_Namee AS EmpNameE, dbo.TblEmployee.Emp_ID AS ForemanID,"
My_SQL1 = My_SQL1 & "            dbo.TblEmployee.Emp_Name AS ForemanName, dbo.TblEmployee.Emp_Namee AS ForemanNameE, TblEmployee_2.Emp_ID AS TeacherID,"
My_SQL1 = My_SQL1 & "            TblEmployee_2.Emp_Name AS TeacherName, TblEmployee_2.Emp_Namee AS TeacherNameE, dbo.Tbl_TradingContract.TContract_CustID,"
My_SQL1 = My_SQL1 & "           dbo.TblCustemers.CusName , dbo.TblCustemers.Fullcode, dbo.Tbl_BusinessDialy.TradingContractID, dbo.Tbl_BusinessDialyDet.BDet_DayMeter"
My_SQL1 = My_SQL1 & "           ,Users.UserName"
My_SQL1 = My_SQL1 & " FROM  dbo.Tbl_BusinessDialy INNER JOIN"
My_SQL1 = My_SQL1 & "             dbo.Tbl_BusinessDialyDet ON dbo.Tbl_BusinessDialy.ID = dbo.Tbl_BusinessDialyDet.BDet_BD_ID INNER JOIN"
My_SQL1 = My_SQL1 & "             dbo.TblBranchesData ON dbo.Tbl_BusinessDialy.BD_BranchID = dbo.TblBranchesData.branch_id Left Outer JOIN"
My_SQL1 = My_SQL1 & "            dbo.TblEmployee AS TblEmployee_1 ON dbo.Tbl_BusinessDialyDet.BDet_EmpID = TblEmployee_1.Emp_ID INNER JOIN"
My_SQL1 = My_SQL1 & "            dbo.Tbl_TradingContract ON dbo.Tbl_BusinessDialy.TradingContractID = dbo.Tbl_TradingContract.ID INNER JOIN"
My_SQL1 = My_SQL1 & "           dbo.TblCustemers ON dbo.Tbl_TradingContract.TContract_CustID = dbo.TblCustemers.CusID Left OUTER JOIN"
My_SQL1 = My_SQL1 & "           dbo.TblEmployee ON dbo.Tbl_BusinessDialyDet.BDet_EmpFormanID = dbo.TblEmployee.Emp_ID Left OUTER JOIN"
My_SQL1 = My_SQL1 & "           dbo.TblEmployee AS TblEmployee_2 ON dbo.Tbl_BusinessDialyDet.BDet_EmpTecherID = TblEmployee_2.Emp_ID"
My_SQL1 = My_SQL1 & "           Left OUTER JOIN dbo.TblUsers AS Users ON dbo.Tbl_BusinessDialy.UserID = Users.UserId"

My_SQL1 = My_SQL1 & "  WHERE   (dbo.Tbl_BusinessDialy.ID is not null OR dbo.Tbl_BusinessDialy.BD_BranchID in(" & Current_branchSql & "))"
  
 StrWhere = ""

 If val(Me.TxtSearchCode(0).Text) <> 0 Then
  StrWhere = StrWhere & " And dbo.Tbl_BusinessDialy.TradingContractID = " & val(TxtSearchCode(0).Text)
End If

 If val(Me.DCombo1.BoundText) <> 0 Or Me.DCombo1.Text <> "" Then
  StrWhere = StrWhere & " AND dbo.Tbl_BusinessDialyDet.BDet_EmpID = " & val(DCombo1.BoundText)
End If

 If val(Me.DCombo2.BoundText) <> 0 Or Me.DCombo2.Text <> "" Then
  StrWhere = StrWhere & " AND Tbl_BusinessDialyDet.BDet_EmpFormanID = " & val(DCombo2.BoundText)
End If

 If val(Me.DCombo3.BoundText) <> 0 Or Me.DCombo3.Text <> "" Then
  StrWhere = StrWhere & " AND dbo.Tbl_BusinessDialyDet.BDet_EmpTecherID = " & val(DCombo3.BoundText)
 End If

If optStatus(8) Then
    StrWhere = StrWhere & " AND IsNull(dbo.Tbl_TradingContract.IsCanceld,0) =0"
ElseIf optStatus(7) Then
    StrWhere = StrWhere & " AND IsNull(dbo.Tbl_TradingContract.IsCanceld,0) =1"
End If

If Not IsNull(Fromdate3.value) Then
  StrWhere = StrWhere & " And dbo.Tbl_BusinessDialy.BD_Date  >= " & SQLDate(Fromdate3.value, True)
End If
If Not IsNull(ToDate3.value) Then
  StrWhere = StrWhere & " AND dbo.Tbl_BusinessDialy.BD_Date  <= " & SQLDate(ToDate3.value, True)
End If

If Me.DcboUsers(0).BoundText <> "" Then
    StrWhere = StrWhere + " and Tbl_BusinessDialy.UserID= " & Me.DcboUsers(0).BoundText & ""
End If


'StrWhere = StrWhere & " order by TblStuFingerprint.StudID"
  My_SQL1 = My_SQL1 & StrWhere
  print_report_MeasureInfo My_SQL1, 2
End Sub

Sub GetReportMeasurementOne()
Dim My_SQL1 As String
Dim MySQL As String
Dim StrWhere As String

 My_SQL1 = "SELECT    dbo.TBL_measureMent.ID,dbo.TblCustemers.CusName, dbo.TBL_measureMent.Cust_Mobile, dbo.TBL_measureMent.Cust_City, dbo.TBL_measureMent.Cust_Time,"
    My_SQL1 = My_SQL1 & "   dbo.TBL_measureMent.Cust_District, dbo.TBL_measureMent.Date_Order, dbo.TBL_measureMent.Date_measureMent, dbo.TBL_measureMent.level1, "
    My_SQL1 = My_SQL1 & "  dbo.TBL_measureMent.WCMen1, dbo.TBL_measureMent.WCWomen1, dbo.TBL_measureMent.WCChildren1, dbo.TBL_measureMent.WCGirls1,"
    My_SQL1 = My_SQL1 & "  dbo.TBL_measureMent.WCCount1, dbo.TBL_measureMent.WCNote1, dbo.TBL_measureMent.laundryMen1, dbo.TBL_measureMent.laundryWomen1,"
    My_SQL1 = My_SQL1 & "  dbo.TBL_measureMent.laundryChildren1, dbo.TBL_measureMent.laundryGirls1, dbo.TBL_measureMent.laundryCount1, dbo.TBL_measureMent.laundryNote1,"
    My_SQL1 = My_SQL1 & "  dbo.TBL_measureMent.laundryareaMen1, dbo.TBL_measureMent.laundryareaWomen1, dbo.TBL_measureMent.laundryareaChildren1, "
    My_SQL1 = My_SQL1 & " dbo.TBL_measureMent.laundryareaGirls1, dbo.TBL_measureMent.laundryareaCount1, dbo.TBL_measureMent.laundryareaNote1, "

    My_SQL1 = My_SQL1 & " dbo.TBL_measureMent.MainHall1, dbo.TBL_measureMent.MainHallCount1, dbo.TBL_measureMent.MainHallNote1, dbo.TBL_measureMent.kitchen1, "

    My_SQL1 = My_SQL1 & "  dbo.TBL_measureMent.kitchenCount1, dbo.TBL_measureMent.kitchenNote1, dbo.TBL_measureMent.BoardMen1, dbo.TBL_measureMent.BoardWomen1, "
    My_SQL1 = My_SQL1 & "   dbo.TBL_measureMent.BoardCount1, dbo.TBL_measureMent.BoardNote1, dbo.TBL_measureMent.MaklatMen1, dbo.TBL_measureMent.MaklatWomen1, "

    My_SQL1 = My_SQL1 & "  dbo.TBL_measureMent.MaklatCount1, dbo.TBL_measureMent.MaklatNote1, dbo.TBL_measureMent.EntranceMen1, dbo.TBL_measureMent.EntranceWomen1, "
    My_SQL1 = My_SQL1 & " dbo.TBL_measureMent.EntranceCount1, dbo.TBL_measureMent.EntranceNote1, dbo.TBL_measureMent.Dorginside1, dbo.TBL_measureMent.DorgOutside1, "
    My_SQL1 = My_SQL1 & "  dbo.TBL_measureMent.DorgTwacheh1, dbo.TBL_measureMent.DorgCount1, dbo.TBL_measureMent.DorgNote1, dbo.TBL_measureMent.ElevatorInside1, "

    My_SQL1 = My_SQL1 & "  dbo.TBL_measureMent.ElevatorOutSide1, dbo.TBL_measureMent.ElevatorCount1, dbo.TBL_measureMent.ElevatorNote1, dbo.TBL_measureMent.HoshInside1, "

    My_SQL1 = My_SQL1 & "     dbo.TBL_measureMent.HoshOutSide1, dbo.TBL_measureMent.HoshCount1, dbo.TBL_measureMent.HoshNote1, dbo.TBL_measureMent.MainRoom1, "
    My_SQL1 = My_SQL1 & "   dbo.TBL_measureMent.MRoom1, dbo.TBL_measureMent.LRoom1, dbo.TBL_measureMent.MainRoomCount1, dbo.TBL_measureMent.MainRoomNote1, "

   My_SQL1 = My_SQL1 & " dbo.TBL_measureMent.Na3laNormal1, dbo.TBL_measureMent.Na3laDorg1, dbo.TBL_measureMent.Na3laCount1, dbo.TBL_measureMent.Na3laNote1,"
   My_SQL1 = My_SQL1 & "          dbo.TBL_measureMent.ClothesInside1, dbo.TBL_measureMent.ClothesOutInside1, dbo.TBL_measureMent.ClothesCount1, dbo.TBL_measureMent.ClothesNote1,"
   My_SQL1 = My_SQL1 & "       dbo.TBL_measureMent.ParkingGround1, dbo.TBL_measureMent.ParkingGdar1, dbo.TBL_measureMent.ParkingCount1, dbo.TBL_measureMent.ParkingNote1,"
        My_SQL1 = My_SQL1 & "     dbo.TBL_measureMent.Office1, dbo.TBL_measureMent.OfficeCount1, dbo.TBL_measureMent.OfficeNote1, dbo.TBL_measureMent.Cust_Mobile2,"
        My_SQL1 = My_SQL1 & "     dbo.TBL_measureMent.Cust_City2, dbo.TBL_measureMent.Cust_Time2, dbo.TBL_measureMent.Cust_District2, dbo.TBL_measureMent.Date_Order2,"
        My_SQL1 = My_SQL1 & "    dbo.TBL_measureMent.Date_measureMent2, dbo.TBL_measureMent.level2, dbo.TBL_measureMent.WCMen2, dbo.TBL_measureMent.WCWomen2,"
        My_SQL1 = My_SQL1 & " dbo.TBL_measureMent.WCChildren2, dbo.TBL_measureMent.WCGirls2, dbo.TBL_measureMent.WCCount2, dbo.TBL_measureMent.WCNote2,"
        My_SQL1 = My_SQL1 & "  dbo.TBL_measureMent.laundryMen2, dbo.TBL_measureMent.laundryWomen2, dbo.TBL_measureMent.laundryChildren2, dbo.TBL_measureMent.laundryGirls2,"
        My_SQL1 = My_SQL1 & "  dbo.TBL_measureMent.laundryCount2, dbo.TBL_measureMent.laundryNote2, dbo.TBL_measureMent.laundryareaMen2,"
        My_SQL1 = My_SQL1 & "    dbo.TBL_measureMent.laundryareaWomen2, dbo.TBL_measureMent.laundryareaChildren2, dbo.TBL_measureMent.laundryareaGirls2,"
        My_SQL1 = My_SQL1 & "  dbo.TBL_measureMent.laundryareaCount2, dbo.TBL_measureMent.laundryareaNote2, dbo.TBL_measureMent.MainHall2, dbo.TBL_measureMent.MainHallCount2,"
        My_SQL1 = My_SQL1 & "    dbo.TBL_measureMent.MainHallNote2, dbo.TBL_measureMent.kitchen2, dbo.TBL_measureMent.kitchenCount2, dbo.TBL_measureMent.kitchenNote2,"
        My_SQL1 = My_SQL1 & " dbo.TBL_measureMent.BoardMen2, dbo.TBL_measureMent.BoardWomen2, dbo.TBL_measureMent.BoardCount2, dbo.TBL_measureMent.BoardNote2,"
        My_SQL1 = My_SQL1 & "   dbo.TBL_measureMent.MaklatMen2, dbo.TBL_measureMent.MaklatWomen2, dbo.TBL_measureMent.MaklatCount2, dbo.TBL_measureMent.MaklatNote2,"
        My_SQL1 = My_SQL1 & "   dbo.TBL_measureMent.EntranceMen2, dbo.TBL_measureMent.EntranceWomen2, dbo.TBL_measureMent.EntranceCount2,"
        My_SQL1 = My_SQL1 & "   dbo.TBL_measureMent.EntranceNote2, dbo.TBL_measureMent.Dorginside2, dbo.TBL_measureMent.DorgOutside2, dbo.TBL_measureMent.DorgTwacheh2,"
        My_SQL1 = My_SQL1 & "     dbo.TBL_measureMent.DorgCount2, dbo.TBL_measureMent.DorgNote2, dbo.TBL_measureMent.ElevatorInside2, dbo.TBL_measureMent.ElevatorOutSide2,"
        My_SQL1 = My_SQL1 & "    dbo.TBL_measureMent.ElevatorCount2, dbo.TBL_measureMent.ElevatorNote2, dbo.TBL_measureMent.HoshInside2, dbo.TBL_measureMent.HoshOutSide2,"
        My_SQL1 = My_SQL1 & "  dbo.TBL_measureMent.HoshCount2, dbo.TBL_measureMent.HoshNote2, dbo.TBL_measureMent.MainRoom2, dbo.TBL_measureMent.MRoom2,"
        My_SQL1 = My_SQL1 & "   dbo.TBL_measureMent.LRoom2, dbo.TBL_measureMent.MainRoomCount2, dbo.TBL_measureMent.MainRoomNote2, dbo.TBL_measureMent.Na3laNormal2,"
        My_SQL1 = My_SQL1 & "    dbo.TBL_measureMent.Na3laDorg2, dbo.TBL_measureMent.Na3laCount2, dbo.TBL_measureMent.Na3laNote2, dbo.TBL_measureMent.ClothesInside2,"
        My_SQL1 = My_SQL1 & "    dbo.TBL_measureMent.ClothesOutInside2, dbo.TBL_measureMent.ClothesCount2, dbo.TBL_measureMent.ClothesNote2,"
        My_SQL1 = My_SQL1 & "  dbo.TBL_measureMent.ParkingGround2, dbo.TBL_measureMent.ParkingGdar2, dbo.TBL_measureMent.ParkingCount2, dbo.TBL_measureMent.ParkingNote2,"
        My_SQL1 = My_SQL1 & "  dbo.TBL_measureMent.Office2 , dbo.TBL_measureMent.OfficeCount2, dbo.TBL_measureMent.OfficeNote2"

        My_SQL1 = My_SQL1 & " From dbo.TBL_measureMent INNER JOIN dbo.TblCustemers ON dbo.TBL_measureMent.Cust_name_ID = dbo.TblCustemers.CusID"
        My_SQL1 = My_SQL1 & "  Where (TBL_measureMent.ID is not null)"
              
              
              
              
        MySQL = " SELECT Users.UserName,TBL_measureMent.ID,IsNull(T2.Level,0) as level1,T2.LevelName ,"
       MySQL = MySQL & "  TBL_measureMent.CustomerName as CusName,TBL_measureMent.Cust_Mobile,TBL_measureMent.Cust_City,"
       MySQL = MySQL & " TBL_measureMent.Cust_Time,TBL_measureMent.Cust_District,TBL_measureMent.Date_Order,TBL_measureMent.Date_measureMent,T2.WCMen WCMen1,"
       MySQL = MySQL & " T2.WCWomen WCWomen1,T2.WCChildren WCChildren1,T2.WCGirls WCGirls1,"
       MySQL = MySQL & " T2.WCCount WCCount1,T2.WCNote WCNote1,T2.laundryMen laundryMen1,T2.laundryWomen laundryWomen1,"
       MySQL = MySQL & " T2.laundryChildren laundryChildren1,T2.laundryGirls laundryGirls1,T2.laundryCount laundryCount1,"
       MySQL = MySQL & " T2.laundryNote laundryNote1,T2.laundryareaMen laundryareaMen1,T2.laundryareaWomen laundryareaWomen1,"
       MySQL = MySQL & " T2.laundryareaChildren laundryareaChildren1,T2.laundryareaGirls laundryareaGirls1,"
       MySQL = MySQL & " T2.laundryareaCount laundryareaCount1,T2.laundryareaNote laundryareaNote1,T2.MainHall MainHall1,"
       MySQL = MySQL & " T2.MainHallCount MainHallCount1,T2.MainHallNote MainHallNote1,T2.kitchen kitchen1,T2.kitchenCount kitchenCount1,"
       MySQL = MySQL & " T2.kitchenNote kitchenNote1,T2.BoardMen BoardMen1,T2.BoardWomen BoardWomen1,T2.BoardCount BoardCount1,"
       MySQL = MySQL & " T2.BoardNote BoardNote1,T2.MaklatMen MaklatMen1,T2.MaklatWomen MaklatWomen1,T2.MaklatCount MaklatCount1,"
       MySQL = MySQL & " T2.MaklatNote MaklatNote1,T2.EntranceMen EntranceMen1,T2.EntranceWomen EntranceWomen1,"
       MySQL = MySQL & " T2.EntranceCount EntranceCount1,T2.EntranceNote EntranceNote1,T2.Dorginside Dorginside1,"
       MySQL = MySQL & " T2.DorgOutside DorgOutside1,T2.DorgTwacheh DorgTwacheh1,T2.DorgCount DorgCount1,"
       MySQL = MySQL & " T2.DorgNote DorgNote1,T2.ElevatorInside ElevatorInside1,T2.ElevatorOutSide ElevatorOutSide1,"
       MySQL = MySQL & " T2.ElevatorCount ElevatorCount1,T2.ElevatorNote ElevatorNote1,T2.HoshInside HoshInside1,"
       MySQL = MySQL & " T2.HoshOutSide HoshOutSide1,T2.HoshCount HoshCount1,T2.HoshNote HoshNote1,"
       MySQL = MySQL & " T2.MainRoom MainRoom1,T2.MRoom MRoom1,T2.LRoom LRoom1,T2.MainRoomCount MainRoomCount1,"
       MySQL = MySQL & " T2.MainRoomNote MainRoomNote1,T2.Na3laNormal Na3laNormal1,T2.Na3laDorg Na3laDorg1,"
       MySQL = MySQL & " T2.Na3laCount Na3laCount1,T2.Na3laNote Na3laNote1,T2.ClothesInside ClothesInside1,"
       MySQL = MySQL & " T2.ClothesOutInside ClothesOutInside1,T2.ClothesCount ClothesCount1,"
       MySQL = MySQL & " T2.ClothesNote ClothesNote1,T2.ParkingGround ParkingGround1,T2.ParkingGdar ParkingGdar1,"
       MySQL = MySQL & " T2.ParkingCount ParkingCount1,T2.ParkingNote ParkingNote1,T2.Office Office1,"
       MySQL = MySQL & " T2.OfficeCount OfficeCount1,T2.OfficeNote OfficeNote1"
       MySQL = MySQL & " From TBL_measureMent"
       MySQL = MySQL & "        LEFT Outer JOIN TBL_measureMent2  T2"
       MySQL = MySQL & "       ON  TBL_measureMent.ID= T2.BDet_BD_ID"
       MySQL = MySQL & "        LEFT Outer JOIN TblUsers  Users"
       MySQL = MySQL & "       ON  TBL_measureMent.UserID= Users.UserID"
  StrWhere = " Where 1 = 1 "


If val(Me.Txt_OrderNumber.Text) <> 0 Then
  StrWhere = StrWhere & " And dbo.TBL_measureMent.ID = " & val(Txt_OrderNumber.Text)
End If



If Me.DcboUsers(0).BoundText <> "" Then
    StrWhere = StrWhere + " and TBL_measureMent.UserID= " & Me.DcboUsers(0).BoundText & ""
End If

If Not IsNull(Fromdate3.value) Then
  StrWhere = StrWhere & " And dbo.TBL_measureMent.Date_Order  >= " & SQLDate(Fromdate3.value, True)
End If
If Not IsNull(ToDate3.value) Then
  StrWhere = StrWhere & " AND dbo.TBL_measureMent.Date_Order  <= " & SQLDate(ToDate3.value, True)
End If

'StrWhere = StrWhere & " order by TblStuFingerprint.StudID"
  My_SQL1 = MySQL & StrWhere
  print_report_MeasureInfo My_SQL1, 4
End Sub

Function print_report_MeasureInfo(Optional NoteSerial As String, Optional Ind As Integer = 0)
     
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
    Dim FromDate As Date
    Dim ToDate As Date
    Dim IsToDate As Boolean
    Dim IsFromDate As Boolean
If Ind = 1 Then
 If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "Rpt_ProductionS.rpt"
            Else
             StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "Rpt_ProductionS.rpt"
       End If
       
ElseIf Ind = 2 Then
     If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "Rpt_AllBusinessDialy.rpt"
            Else
             StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "Rpt_AllBusinessDialy.rpt"
       End If
       If Not IsNull(Fromdate3.value) Then
            IsFromDate = True
            FromDate = Fromdate3.value
        End If
        If Not IsNull(ToDate3.value) Then
            ToDate = ToDate3.value
            IsToDate = True
        End If
ElseIf Ind = 3 Then
     If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "Rpt_TransOrder.rpt"
            Else
             StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "Rpt_TransOrder.rpt"
       End If
        If Not IsNull(Fromdate3.value) Then
            IsFromDate = True
            FromDate = Fromdate3.value
        End If
        If Not IsNull(ToDate3.value) Then
            ToDate = ToDate3.value
            IsToDate = True
        End If

       
ElseIf Ind = 4 Then
     If SystemOptions.UserInterface = ArabicInterface Then
     
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "Rpt_measureMent.rpt"
            Else
             StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "Rpt_measureMent.rpt"
       End If
        If Not IsNull(Fromdate3.value) Then
            IsFromDate = True
            FromDate = Fromdate3.value
        End If
        If Not IsNull(ToDate3.value) Then
            ToDate = ToDate3.value
            IsToDate = True
        End If

   
ElseIf Ind = 5 Then
     If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "Rpt_BusinessDialy2.rpt"
            Else
             StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "Rpt_BusinessDialy2.rpt"
       End If
    
        If Not IsNull(Fromdate4.value) Then
            IsFromDate = True
            FromDate = Fromdate4.value
        End If
        If Not IsNull(ToDate4.value) Then
            ToDate = ToDate4.value
            IsToDate = True
        End If

ElseIf Ind = 6 Then
     If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "Rpt_TradingContractTotal.rpt"
            Else
             StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "Rpt_TradingContractTotal.rpt"
       End If

        If Not IsNull(Fromdate5.value) Then
            IsFromDate = True
            FromDate = Fromdate5.value
        End If
        If Not IsNull(toDate5.value) Then
            ToDate = toDate5.value
            IsToDate = True
        End If
        
End If
    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
        Cn.CommandTimeout = 10000
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
        StrReportTitle = "" '& StrAccountName
    Else
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        StrReportTitle = ""
    End If
    Dim i As Long
    
        xReport.EnableParameterPrompting = False
    For i = 1 To xReport.ParameterFields.count
        Select Case xReport.ParameterFields.Item(i).ParameterFieldName
        Case "FrmDate"
            If IsFromDate Then
                xReport.ParameterFields.Item(i).AddCurrentValue FromDate
            End If
        Case "ToDate"
            If IsToDate Then
                xReport.ParameterFields.Item(i).AddCurrentValue ToDate
            End If
        Case "ParPrintUser"
            xReport.ParameterFields.Item(i).AddCurrentValue user_name
        End Select
    Next i
    
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




Private Sub ISButton2_Click()

Dim My_SQL1 As String
Dim StrWhere As String


  
  

My_SQL1 = " SELECT Tu1.UserName,dbo.Tbl_TradingContract.ID,TotalDisc,CAST(Tbl_TradingContract.Vat2 AS NVARCHAR(10)) Vat2,Tbl_TradingContract.VAt2 +Tbl_TradingContract.ProjectTotal - TotalDisc as  TotalWithVat, dbo.Tbl_TradingContract.TContract_Date BD_Date, dbo.Tbl_BusinessDialy.BD_BranchID, dbo.TblBranchesData.branch_name,"
My_SQL1 = My_SQL1 & "             (case When IsNull(NetBVat,0) = 0 Then Tbl_TradingContract.ProjectTotal + TotalDisc else NetBVat End) as NetBVat,"
My_SQL1 = My_SQL1 & "             dbo.Tbl_BusinessDialyDet.ID , TblProcessUnites.UnitName,dbo.TblProcessDEF.Interval,"
My_SQL1 = My_SQL1 & " Tbl_TradingContractDet.Periods Period,Tbl_TradingContract.ProjectTotal,Tbl_TradingContract.Location,"
My_SQL1 = My_SQL1 & " Tbl_TradingContract.DtProjStart,"
My_SQL1 = My_SQL1 & "         dbo.Tbl_TradingContractDet.TContractDet_SalPrice ,dbo.Tbl_TradingContractDet.TContractDet_Value,"
My_SQL1 = My_SQL1 & "         dbo.Tbl_TradingContractDet.TContractDet_InstallPrice ,"
My_SQL1 = My_SQL1 & " Tbl_TradingContract.Resp,Tbl_TradingContractDet.TContractDet_specificationEn Discr, Tbl_TradingContract.TOrder_Address,Tbl_TradingContract.TOrder_Phone,Tbl_TradingContract.RespWorker,"

My_SQL1 = My_SQL1 & "             dbo.TblBranchesData.branch_namee, dbo.Tbl_BusinessDialy.BD_Notes, dbo.Tbl_BusinessDialyDet.BDet_BD_ID, dbo.Tbl_BusinessDialyDet.BDet_BandNo,"
My_SQL1 = My_SQL1 & "                dbo.Tbl_TradingContractDet.TContractDet_TotalSalPrice, dbo.Tbl_TradingContractDet.TContractDet_TotalInstallPrice,"
My_SQL1 = My_SQL1 & "            Tbl_TradingContractDet.TContractDet_Qun BDet_Qun ,dbo.Tbl_BusinessDialyDet.BDet_Qun  BdQun, TblProcessDEF.ProcessName BDet_Name, dbo.Tbl_BusinessDialyDet.BDet_NameE, TblEmployee_1.Emp_ID AS EmpID,"
My_SQL1 = My_SQL1 & "            TblEmployee_1.Emp_Name AS EmpName, TblEmployee_1.Emp_Namee AS EmpNameE, dbo.TblEmployee.Emp_ID AS ForemanID,"
My_SQL1 = My_SQL1 & "            dbo.TblEmployee.Emp_Name AS ForemanName, dbo.TblEmployee.Emp_Namee AS ForemanNameE, TblEmployee_2.Emp_ID AS TeacherID,"
My_SQL1 = My_SQL1 & "            TblEmployee_2.Emp_Name AS TeacherName, TblEmployee_2.Emp_Namee AS TeacherNameE, dbo.Tbl_TradingContract.TContract_CustID,"
My_SQL1 = My_SQL1 & "           dbo.TblCustemers.CusName , dbo.TblCustemers.Fullcode, dbo.Tbl_TradingContract.ID TradingContractID ,"
My_SQL1 = My_SQL1 & "           dbo.Tbl_BusinessDialyDet.BDet_DayMeter,"

My_SQL1 = My_SQL1 & "                       ExpensesTotal = (SELECT Sum(n.Note_Value)"
My_SQL1 = My_SQL1 & "                    FROM Notes AS n WHERE IsNull(n.TradingContractID,0)  = dbo.Tbl_TradingContract.ID"
My_SQL1 = My_SQL1 & "                    AND (n.NoteType = 5 OR n.NoteType = 3)),"

My_SQL1 = My_SQL1 & "                       TContractDet_TotalInstallPrice = (SELECT Sum(n.Note_ValueSales)"
My_SQL1 = My_SQL1 & "                    FROM Notes AS n WHERE IsNull(n.TradingContractID,0)  = dbo.Tbl_TradingContract.ID"
My_SQL1 = My_SQL1 & "                    AND (n.NoteType = 180)),"

My_SQL1 = My_SQL1 & "                  ReciptTotal = (SELECT Sum(n.Note_Value)"
My_SQL1 = My_SQL1 & "                    FROM Notes AS n WHERE IsNull(n.TradingContractID,0)  = dbo.Tbl_TradingContract.ID"
My_SQL1 = My_SQL1 & "                    AND ( n.NoteType = 4))"


My_SQL1 = My_SQL1 & "           From TblProcessUnites"
My_SQL1 = My_SQL1 & "                  RIGHT OUTER JOIN TblProcessDEF"
My_SQL1 = My_SQL1 & "                  RIGHT OUTER JOIN TblEmployee   AS TblEmployee_2"
My_SQL1 = My_SQL1 & "                  RIGHT OUTER JOIN Tbl_TradingContractDet"
My_SQL1 = My_SQL1 & "                  RIGHT OUTER JOIN Tbl_TradingContract"
My_SQL1 = My_SQL1 & "                       ON  Tbl_TradingContractDet.TContractDet_TContractID = Tbl_TradingContract.ID"
My_SQL1 = My_SQL1 & "                  LEFT OUTER JOIN TblEmployee"
My_SQL1 = My_SQL1 & "                  RIGHT OUTER JOIN Tbl_BusinessDialyDet"
My_SQL1 = My_SQL1 & "                       ON  TblEmployee.Emp_ID = Tbl_BusinessDialyDet.BDet_EmpFormanID"
My_SQL1 = My_SQL1 & "                  RIGHT OUTER JOIN Tbl_BusinessDialy"
My_SQL1 = My_SQL1 & "                       ON  Tbl_BusinessDialyDet.BDet_BD_ID = Tbl_BusinessDialy.ID"
If Not IsNull(Fromdate4.value) Then
  My_SQL1 = My_SQL1 & " And dbo.Tbl_BusinessDialy.BD_Date  >= " & SQLDate(Fromdate4.value, True)
End If
If Not IsNull(ToDate4.value) Then
  My_SQL1 = My_SQL1 & " AND dbo.Tbl_BusinessDialy.BD_Date  <= " & SQLDate(ToDate4.value, True)
End If
My_SQL1 = My_SQL1 & "                  LEFT OUTER JOIN TblBranchesData"
My_SQL1 = My_SQL1 & "                       ON  Tbl_BusinessDialy.BD_BranchID = TblBranchesData.branch_id"
My_SQL1 = My_SQL1 & "                  LEFT OUTER JOIN TblEmployee    AS TblEmployee_1"
My_SQL1 = My_SQL1 & "                       ON  Tbl_BusinessDialyDet.BDet_EmpID = TblEmployee_1.Emp_ID"
My_SQL1 = My_SQL1 & "                       ON  Tbl_TradingContract.ID = Tbl_BusinessDialy.TradingContractID"
My_SQL1 = My_SQL1 & "                       AND Tbl_TradingContractDet.ProcessDEFID = Tbl_BusinessDialyDet.TConID"
My_SQL1 = My_SQL1 & "                       AND Tbl_TradingContractDet.TContractDet_TContractID = Tbl_BusinessDialy.TradingContractID"
My_SQL1 = My_SQL1 & "                  LEFT OUTER JOIN TblCustemers"
My_SQL1 = My_SQL1 & "                       ON  Tbl_TradingContract.TContract_CustID = TblCustemers.CusID"
My_SQL1 = My_SQL1 & "                       ON  TblEmployee_2.Emp_ID = Tbl_BusinessDialyDet.BDet_EmpTecherID"
My_SQL1 = My_SQL1 & "                       ON  TblProcessDEF.TblProcessDEFID = Tbl_TradingContractDet.ProcessDEFID"
My_SQL1 = My_SQL1 & "                       ON  TblProcessUnites.UnitID = Tbl_TradingContractDet.TContractDet_UnitID"
My_SQL1 = My_SQL1 & "                  LEFT OUTER JOIN TblUsers TU1"
My_SQL1 = My_SQL1 & "                       ON  Tbl_TradingContract.UserID = TU1.UserID"

'My_SQL1 = My_SQL1 & " FROM  dbo.Tbl_BusinessDialy INNER JOIN"
'My_SQL1 = My_SQL1 & "             dbo.Tbl_BusinessDialyDet ON dbo.Tbl_BusinessDialy.ID = dbo.Tbl_BusinessDialyDet.BDet_BD_ID INNER JOIN"
'My_SQL1 = My_SQL1 & "             dbo.TblBranchesData ON dbo.Tbl_BusinessDialy.BD_BranchID = dbo.TblBranchesData.branch_id INNER JOIN"
'My_SQL1 = My_SQL1 & "            dbo.TblEmployee AS TblEmployee_1 ON dbo.Tbl_BusinessDialyDet.BDet_EmpID = TblEmployee_1.Emp_ID INNER JOIN"
'My_SQL1 = My_SQL1 & "            dbo.Tbl_TradingContract ON dbo.Tbl_BusinessDialy.TradingContractID = dbo.Tbl_TradingContract.ID INNER JOIN"
'My_SQL1 = My_SQL1 & "           dbo.TblCustemers ON dbo.Tbl_TradingContract.TContract_CustID = dbo.TblCustemers.CusID RIGHT OUTER JOIN"
'My_SQL1 = My_SQL1 & "           dbo.TblEmployee ON dbo.Tbl_BusinessDialyDet.BDet_EmpFormanID = dbo.TblEmployee.Emp_ID RIGHT OUTER JOIN"
'My_SQL1 = My_SQL1 & "           dbo.TblEmployee AS TblEmployee_2 ON dbo.Tbl_BusinessDialyDet.BDet_EmpTecherID = TblEmployee_2.Emp_ID"
'
'
'
'My_SQL1 = My_SQL1 & "           LEFT OUTER JOIN Tbl_TradingContractDet"
'My_SQL1 = My_SQL1 & "           ON  Tbl_BusinessDialyDet.TConID = dbo.Tbl_TradingContractDet.ProcessDEFID"
'My_SQL1 = My_SQL1 & "                       AND Tbl_TradingContractDet.TContractDet_TContractID =Tbl_BusinessDialy.TradingContractID"

'My_SQL1 = My_SQL1 & "           AND dbo.Tbl_TradingContractDet.ProcessDEFID = dbo.Tbl_TradingContractDet.ProcessDEFID"
'
'My_SQL1 = My_SQL1 & "           Left OUTER JOIN TblProcessDEF"
'My_SQL1 = My_SQL1 & "           On dbo.Tbl_TradingContractDet.ProcessDEFID = TblProcessDEF.TblProcessDEFID"
'
'My_SQL1 = My_SQL1 & "           Left OUTER JOIN TblProcessUnites"
'My_SQL1 = My_SQL1 & "           On dbo.Tbl_TradingContractDet.TContractDet_UnitID = TblProcessUnites.UnitID"


'My_SQL1 = My_SQL1 & "  WHERE   (dbo.Tbl_BusinessDialy.ID is not null OR dbo.Tbl_BusinessDialy.BD_BranchID in(" & Current_branchSql & "))"
  
 StrWhere = " Where 1 = 1 "

Dim Msg As String
If val(Me.TxtSearchCode(1).Text) <> 0 Then
    StrWhere = StrWhere & " And dbo.Tbl_TradingContract.ID = " & val(TxtSearchCode(1).Text)
Else
  '  Msg = "«œŒ· —ﬁ„ «·« ›«ﬁÌ…"
  '  MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
End If

If val(Me.DcCustmer(1).BoundText) <> 0 Or Me.DcCustmer(1).Text <> "" Then
  StrWhere = StrWhere & " AND dbo.Tbl_TradingContract.TContract_CustID = " & val(DcCustmer(1).BoundText)
 End If
If optStatus(0) Then
    StrWhere = StrWhere & " AND IsNull(dbo.Tbl_TradingContract.IsCanceld,0) =0"
ElseIf optStatus(1) Then
    StrWhere = StrWhere & " AND IsNull(dbo.Tbl_TradingContract.IsCanceld,0) =1"
End If
If Me.DcboUsers(1).BoundText <> "" Then
    StrWhere = StrWhere + " and Tbl_TradingContract.UserID= " & Me.DcboUsers(1).BoundText & ""
End If

If Not IsNull(Fromdate4.value) Then
  StrWhere = StrWhere & " And dbo.Tbl_TradingContract.TContract_Date  >= " & SQLDate(Fromdate4.value, True)
End If
If Not IsNull(ToDate4.value) Then
  StrWhere = StrWhere & " AND dbo.Tbl_TradingContract.TContract_Date  <= " & SQLDate(ToDate4.value, True)
End If





'StrWhere = StrWhere & " order by TblStuFingerprint.StudID"
  My_SQL1 = My_SQL1 & StrWhere
  print_report_MeasureInfo My_SQL1, 5

End Sub

Private Sub ISButton3_Click()



Dim MySQL As String
Dim StrWhere As String


  
 MySQL = ""


MySQL = MySQL & "  SELECT ProjectTotal , TContractDet_InstallPrice = ExpensesTotal2,Period TContractDet_UnitID,      cast(Vat2            AS NVARCHAR(10)) as  TContractDet_SalPrice, NetBVat TContractDet_DayMeter,TotalDisc TContractDet_Qun,TotalWithVat TContractDet_Value, BDet_Qun, BdQun,BDet_Qun -BdQun     Diff,TContract_CustID,"
MySQL = MySQL & "         CusName , TradingContractID, ExpensesTotal TContractDet_TotalSalPrice , ReciptTotal TContractDet_TotalInstallPrice,UserName,EmpId"
MySQL = MySQL & "  FROM   ("
MySQL = MySQL & "             SELECT Tbl_TradingContract.ProjectTotal,Tbl_TradingContract.Vat2,tbl_TradingContract.Period,"
MySQL = MySQL & "             case When IsNull(NetBVat,0) = 0 Then Tbl_TradingContract.ProjectTotal + TotalDisc else NetBVat End NetBVat"
MySQL = MySQL & "             ,TotalDisc,"
MySQL = MySQL & "                               Tbl_TradingContract.VAt2 + Tbl_TradingContract.ProjectTotal - TotalDisc AS TotalWithVat,"
MySQL = MySQL & "             tu.UserName,Tbl_TradingContract.UserID AS EmpId,"
MySQL = MySQL & "                    SUM(ISNULL(Tbl_TradingContractDet.TContractDet_Qun, 0)) AS BDet_Qun,"
MySQL = MySQL & "                    BdQun = ISNULL("
MySQL = MySQL & "                        ("
MySQL = MySQL & "                            SELECT SUM(ISNULL(BDet_Qun, 0))"
MySQL = MySQL & "                            From Tbl_BusinessDialyDet"
MySQL = MySQL & "                                   RIGHT OUTER JOIN Tbl_BusinessDialy"
MySQL = MySQL & "                                        ON  Tbl_BusinessDialyDet.BDet_BD_ID = Tbl_BusinessDialy.ID"
MySQL = MySQL & "                            Where Tbl_TradingContract.ID = Tbl_BusinessDialy.TradingContractID"
MySQL = MySQL & "                                   AND Tbl_TradingContract.ID = Tbl_BusinessDialy.TradingContractID"
MySQL = MySQL & "                        ),"
MySQL = MySQL & "  0"
MySQL = MySQL & "                    ),"
MySQL = MySQL & "                    Tbl_TradingContract.TContract_CustID,"
MySQL = MySQL & "                    TblCustemers.CusName,"
MySQL = MySQL & "                    Tbl_TradingContract.ID  AS TradingContractID,"
MySQL = MySQL & "                    ("
MySQL = MySQL & "                        SELECT SUM(Note_Value) AS Expr1"
MySQL = MySQL & "                        FROM   Notes AS n"
MySQL = MySQL & "                        Where (IsNull(TradingContractID, 0) = Tbl_TradingContract.ID)"
MySQL = MySQL & "                               AND (NoteType = 5 OR NoteType = 3 )"
MySQL = MySQL & "                    )                       AS ExpensesTotal,"

MySQL = MySQL & "                    ("
MySQL = MySQL & "                        SELECT SUM(Note_ValueSales) AS Expr1"
MySQL = MySQL & "                        FROM   Notes AS n"
MySQL = MySQL & "                        Where (IsNull(TradingContractID, 0) = Tbl_TradingContract.ID)"
MySQL = MySQL & "                               AND (NoteType = 180)"
MySQL = MySQL & "                    )                       AS ExpensesTotal2,"


MySQL = MySQL & "                    ("
MySQL = MySQL & "                        SELECT SUM(Note_Value) AS Expr1"
MySQL = MySQL & "                        FROM   Notes AS n"
MySQL = MySQL & "                        Where (IsNull(TradingContractID, 0) = Tbl_TradingContract.ID)"
MySQL = MySQL & "                               AND (NoteType = 4)"
MySQL = MySQL & "                    )                       AS ReciptTotal"
MySQL = MySQL & "             From Tbl_TradingContractDet"
MySQL = MySQL & "                    RIGHT OUTER JOIN Tbl_TradingContract"
MySQL = MySQL & "                         ON  Tbl_TradingContractDet.TContractDet_TContractID = Tbl_TradingContract.ID"
MySQL = MySQL & "                    LEFT OUTER JOIN TblCustemers"
MySQL = MySQL & "                         ON  Tbl_TradingContract.TContract_CustID = TblCustemers.CusID"
MySQL = MySQL & "                         LEFT OUTER JOIN TblUsers AS tu"
MySQL = MySQL & "                                                ON  Tbl_TradingContract.uSERid = tu.uSERid"
MySQL = MySQL & "             Where (1 = 1)"



If Not IsNull(Fromdate5.value) Then
    MySQL = MySQL & " And dbo.Tbl_TradingContract.TContract_Date  >= " & SQLDate(Fromdate5.value, True)
End If
If Not IsNull(toDate5.value) Then
    MySQL = MySQL & " AND dbo.Tbl_TradingContract.TContract_Date  <= " & SQLDate(toDate5.value, True)
End If
If Me.DcboUsers(2).BoundText <> "" Then
    MySQL = MySQL + " and Tbl_TradingContract.UserID= " & Me.DcboUsers(2).BoundText & ""
End If
If optStatus(3) Then
    MySQL = MySQL & " AND IsNull(dbo.Tbl_TradingContract.IsCanceld,0) =0"
ElseIf optStatus(4) Then
    MySQL = MySQL & " AND IsNull(dbo.Tbl_TradingContract.IsCanceld,0) =1"
End If


MySQL = MySQL & "             Group By"
MySQL = MySQL & "                    Tbl_TradingContract.ProjectTotal,"
MySQL = MySQL & "                    Tbl_TradingContract.TContract_CustID,"
MySQL = MySQL & "                    TblCustemers.CusName,Tbl_TradingContract.Vat2,NetBVat,TotalDisc , "
MySQL = MySQL & "                    Tbl_TradingContract.ID,Tbl_TradingContract.Period,"
MySQL = MySQL & "                    tu.UserName,Tbl_TradingContract.UserID"
MySQL = MySQL & "                                         "
MySQL = MySQL & "         )                   T "
'MySQL = MySQL & "         )                   WHERE IsNull(T.ProjectTotal,0)- IsNull(T.ReciptTotal,0) <> 0"
  
  
  
 'StrWhere = " Where 1 = 1 "

Dim Msg As String








'StrWhere = StrWhere & " order by TblStuFingerprint.StudID"
 ' My_SQL1 = My_SQL1 & StrWhere
  print_report_MeasureInfo MySQL, 6



End Sub

Private Sub Rd_Click(Index As Integer)
Txt_OrderNumber2.Visible = False
lbl(26).Visible = False
Txt_OrderNumber = ""
If Rd(3).value = True Then
    DcboUsers(0).Visible = False
    lbl(65).Visible = False
    
    TxtSearchCode(0).Visible = False
    'TxtSearchCode(1).Visible = False
    DcCustmer(0).Visible = False
    'DcCustmer(1).Visible = False
    lbl(17).Visible = False
    lbl(18).Visible = False
    lbl(19).Visible = False
    Label1(6).Visible = False
    DCombo1.Visible = False
    DCombo2.Visible = False
    DCombo3.Visible = False
    lbl(20).Visible = False
    lbl(21).Visible = False
    TxtEmpCode.Visible = False
    DcbEmployee2.Visible = False
    Txt_OrderNumber.Visible = False
    
    
ElseIf Rd(4).value = True Then
    DcboUsers(0).Visible = True
    lbl(65).Visible = True
    TxtSearchCode(0).Visible = True
    'TxtSearchCode(1).Visible = False
    
    DcCustmer(0).Visible = True
    DcCustmer(1).Visible = True
    lbl(17).Visible = True
    lbl(18).Visible = True
    lbl(19).Visible = True
    Label1(6).Visible = True
    DCombo1.Visible = True
    DCombo2.Visible = True
    DCombo3.Visible = True
    lbl(20).Visible = False
    lbl(21).Visible = False
    TxtEmpCode.Visible = False
    DcbEmployee2.Visible = False
    Txt_OrderNumber.Visible = False

ElseIf Rd(5).value = True Then
    DcboUsers(0).Visible = True
    lbl(65).Visible = True
    Txt_OrderNumber2.Visible = True
    lbl(26).Visible = True
    TxtSearchCode(0).Visible = False
    'TxtSearchCode(1).Visible = False
    
    DcCustmer(0).Visible = False
    'DcCustmer(1).Visible = False
    
    lbl(17).Visible = False
    lbl(18).Visible = False
    lbl(19).Visible = False
    Label1(6).Visible = False
    DCombo1.Visible = False
    DCombo2.Visible = False
    DCombo3.Visible = False
    lbl(20).Visible = True
    Txt_OrderNumber.Visible = True
    lbl(21).Visible = True
    TxtEmpCode.Visible = True
    DcbEmployee2.Visible = True
    
    
    
    Dcombos.GetUsers Me.DcbEmployee2


ElseIf Rd(6).value = True Then
    DcboUsers(0).Visible = True
    lbl(65).Visible = True

    TxtSearchCode(0).Visible = False
    'TxtSearchCode(1).Visible = False
    
    DcCustmer(0).Visible = False
    'DcCustmer(1).Visible = False
    lbl(17).Visible = False
    lbl(18).Visible = False
    lbl(19).Visible = False
    Label1(6).Visible = False
    DCombo1.Visible = False
    DCombo2.Visible = False
    DCombo3.Visible = False
    lbl(20).Visible = True
    lbl(21).Visible = False
    TxtEmpCode.Visible = False
    DcbEmployee2.Visible = False
    Txt_OrderNumber.Visible = True

End If


End Sub

Private Sub StuInfoRB_Click()
'***********************************************
TxtSudCode.Visible = False
DcbStudent.Visible = False
Label1(4).Visible = False
groupDBox.Visible = False
cursBox.Visible = False
instruDBox.Visible = False
Label1(0).Visible = False
Label1(2).Visible = False
gDBox(2).Visible = False
'***********************************************
Label1(3).Visible = False
'***********************************************
Frame8.Visible = False
Label1(5).Visible = False
Label1(1).Visible = False
DcbBranch.Visible = False
Text15.Visible = False
DcbCompany.Visible = False
Text1.Visible = False
DcbEmployee.Visible = False
AttFrame.Visible = False
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
'If Me.TxtModFlg = "N" Or Me.TxtModFlg = "E" Then
'    Me.DcbEmployee.BoundText = GeTEmpIDByEmpCode(Text1.text, True)
'End If
End Sub

Private Sub Text15_KeyPress(KeyAscii As Integer)
 Dim EmpID As Integer
    If KeyAscii = vbKeyReturn Then
        GetTblCustemersCode Text15.Text, EmpID
        DcbCompany.BoundText = EmpID
    End If
End Sub
Private Sub DcbEmployee_Change()
DcbEmployee_Click (0)
End Sub

Private Sub DcbEmployee_Click(Area As Integer)
If val(Me.DcbEmployee.BoundText) = 0 Then Exit Sub
           Me.Text1.Text = get_EMPLOYEE_Data(val(Me.DcbEmployee.BoundText), "Fullcode")
End Sub
Private Sub ChangeLang()
On Error GoTo ErrTrap
Label5.Caption = "Report Calling"
Label1(1).Caption = "Branch"
gDBox(2).Caption = "Group"
Label1(5).Caption = "Company"
Frame8.Caption = "Period"
lbl(3).Caption = "From"
lbl(2).Caption = "From"
lbl(14).Caption = "To"
btnClear.Caption = "Clear"
Cmd(1).Caption = "Show"
Cmd(2).Caption = "Exit"
Label1(3).Caption = "Caller"
Label1(0).Caption = "Subject"
Label1(2).Caption = "Instructor "
Label1(7).Caption = "Branch"
Label1(8).Caption = "Company"
Label1(9).Caption = "Subject"
Label1(8).Caption = "Company"
CmdPrint.Caption = "Print"
callsRB.Caption = "Calls"
AttRB.Caption = "Attendance"
StuInfoRB.Caption = "Students Info"
lbl(6) = "Month"
lbl(4) = "Please select the branch or date or the total report will be for each branch and duration"
lbl(13) = "Please select the branch or date or the total report will be for each branch and duration"
lbl(15) = "Please select the branch or date or the total report will be for each branch and duration"
lbl(30) = "Please select the branch or date or the total report will be for each branch and duration"
lbl(32) = "Please select the branch or date or the total report will be for each branch and duration"
lbl(5) = "Year"
lbl(7) = "To"
lbl(23) = "To"
lbl(14) = "To"
lbl(10) = "To"
lbl(27) = "To"
lbl(8) = "From"
lbl(9) = "From"
lbl(31) = "From"
lbl(17) = "Building an agreement"
lbl(18) = "Building an agreement"
lbl(28) = "Building an agreement"
'lbl(17) = "Worker"
lbl(19) = "Forman"
lbl(20) = "Request No"
lbl(25) = "Request No"
lbl(21) = "Employee"
lbl(24) = "Employee"
lbl(22) = "From"
lbl(19) = "Forman"


lblCompanyname = "AlSatria"

lblCompanyname.Caption = "AL SATTARYAH GROUP"
XPTab301.TabCaption(0) = "Reports"
XPTab301.TabCaption(1) = "Attendance report in groups"
XPTab301.TabCaption(2) = "Reports of measurements"
XPTab301.TabCaption(3) = "Report of Conventions"
XPTab301.TabCaption(4) = "Total agreements"
lbl(26) = "Order No"
lbl(26) = "Order No"
Rd(3).Caption = "Print standards"
Rd(4).Caption = "Daily Business Reports"
Rd(5).Caption = "Motion report requests"
Rd(6).Caption = "Measurement lift report"
ISButton3.Caption = "View the report"
Frame2(3).Caption = "Select the period"
Frame2(2).Caption = "Select the period"
Frame2(1).Caption = "Select the period"
ISButton2.Caption = "View the report"
ISButton1.Caption = "View the report"
Label1(6) = "Teacher"
'Label5.Caption "Reports"
ErrTrap:
End Sub
Function print_report_Calls(Optional NoteSerial As String)
     
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

     If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepStudentCall.rpt"
            Else
             StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepStudentCallE.rpt"
       End If
    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
        Cn.CommandTimeout = 10000
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
        StrReportTitle = "" '& StrAccountName
    Else
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
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


Private Sub Text3_KeyPress(KeyAscii As Integer)
 Dim EmpID As Integer
    If KeyAscii = vbKeyReturn Then
        GetTblCustemersCode Text3.Text, EmpID
        DcbCompany2.BoundText = EmpID
    End If
End Sub

Private Sub ToDate_Change()
If Not IsNull(ToDate.value) Then
todateH.value = ToHijriDate(ToDate.value)
End If
End Sub

Private Sub toDate2_Change()
If Not IsNull(ToDate2.value) Then
toDateH2.value = ToHijriDate(ToDate2.value)
End If
End Sub

Private Sub ToDateH_LostFocus()
ToDate.value = ToGregorianDate(todateH.value)
End Sub
Public Sub GetAttData22()
    Dim StrSQL As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
 '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
 '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'FillStuRepTable
FillStuRepTable22
 '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
 '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    StrSQL = "select * from TblStuRepTab2"
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
           ' Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
          '  Me.lbl(10).Caption = "Search Results=0"
        End If
If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «· ﬁ—Ì—"
      Else
     Msg = "No Data"
  End If
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    Else
 rs.MoveFirst

 print_report_Att StrSQL, 1

    End If

End Sub
'********************* khaleds reports for Sub****************************
Public Sub GetAttData()
    Dim StrSQL As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
 '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
 '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'FillStuRepTable
FillStuRepTable1111
 '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
 '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    StrSQL = "select * from TblStuRepTab"
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
           ' Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
          '  Me.lbl(10).Caption = "Search Results=0"
        End If
If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «· ﬁ—Ì—"
      Else
     Msg = "No Data"
  End If
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    Else
 rs.MoveFirst

 print_report_Att StrSQL

    End If

End Sub
'********************* khaleds reports for comp****************************
Sub FillDatConp()
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
Dim sql As String
Dim i As Integer
Dim DiFFNo As Integer
Dim TemDate As Date
DiFFNo = DateDiff("d", FromDate.value, ToDate.value) + 1
Cn.Execute "Delete from TblStuNoDay"
sql = "Select * from TblStuNoDay where 1=-1"
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
For i = 1 To DiFFNo
Rs3.AddNew
If i = 1 Then
Rs3("RecordDate").value = FromDate.value
Else
Rs3("RecordDate").value = DateAdd("d", i - 1, FromDate.value)
End If
Rs3("DayID").value = i
Rs3.update
Next i

End Sub
'Public Sub GetAttCompData()
'    Dim StrSQL As String
'    Dim StrWhere As String
'    Dim BolBegine As Boolean
'    Dim rs As ADODB.Recordset
'    Dim Msg As String
'    Dim i As Integer
'    FillStuRepTable22
' 'StrSQL = "SELECT TblStudent.Name, TblStudent.ID, TblAttendance.RecordDate, TblAttendanceDet.IsAttend, TblStuGroup.Name AS GName, TblStudentCurs.Name AS CName, TblInstructors.Name AS IName "
' 'StrSQL = StrSQL & "FROM TblAttendanceDet INNER JOIN "
' 'StrSQL = StrSQL & "TblAttendance ON TblAttendanceDet.AttenID = TblAttendance.ID INNER JOIN "
 'StrSQL = StrSQL & "TblStudent ON TblAttendanceDet.StudID = TblStudent.ID INNER JOIN "
 'StrSQL = StrSQL & "TblCustemers ON TblStudent.CompID = TblCustemers.CusID INNER JOIN "
 'StrSQL = StrSQL & "TblInstructors ON TblAttendance.InstrcID = TblInstructors.ID INNER JOIN "
 'StrSQL = StrSQL & "TblStuGroup ON TblAttendance.GroupID = TblStuGroup.ID INNER JOIN "
 'StrSQL = StrSQL & "TblStudentCurs ON TblAttendance.CursID = TblStudentCurs.ID "
 'StrSQL = StrSQL & "where 1=1"
  
 '   BolBegine = False
 '   StrWhere = ""
    

'If val(Me.DcbBranch.BoundText) <> 0 Or Me.DcbBranch.text <> "" Then
'  StrWhere = StrWhere & " AND dbo.TblStudCalling.BranchID = " & val(Me.DcbBranch.BoundText)
'End If

'If val(Me.DcbCompany.BoundText) <> 0 Or Me.DcbCompany.Text <> "" Then
'  StrWhere = StrWhere & " AND dbo.TblCustemers.CusID  = " & val(DcbCompany.BoundText)
'End If
'
'If val(Me.groupDBox.BoundText) <> 0 Or Me.groupDBox.Text <> "" Then
'  StrWhere = StrWhere & " AND dbo.TblAttendance.GroupID = " & val(groupDBox.BoundText)
'End If
'
'If val(Me.cursBox.BoundText) <> 0 Or Me.cursBox.Text <> "" Then
'  StrWhere = StrWhere & " AND dbo.TblAttendance.CursID = " & val(cursBox.BoundText)
'End If
'
'If val(Me.instruDBox.BoundText) <> 0 Or Me.instruDBox.Text <> "" Then
'  StrWhere = StrWhere & " AND dbo.TblAttendance.InstrcID = " & val(instruDBox.BoundText)
'End If
'

'If Not IsNull(Me.FromDate.value) Then
'  StrWhere = StrWhere & " AND dbo.TblAttendance.RecordDate >=" & SQLDate(Me.FromDate.value, True) & ""
'End If
'
'If Not IsNull(Me.ToDate.value) Then
'  StrWhere = StrWhere & " AND  dbo.TblAttendance.RecordDate <=" & SQLDate(Me.ToDate.value, True) & ""
'End If
'    '---------------------------------
'    StrSQL = StrSQL & StrWhere
'    Set rs = New ADODB.Recordset
'    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    If rs.BOF Or rs.EOF Then
'        If SystemOptions.UserInterface = ArabicInterface Then
'           ' Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=’›—"
'        ElseIf SystemOptions.UserInterface = EnglishInterface Then
'          '  Me.lbl(10).Caption = "Search Results=0"
'        End If
'If SystemOptions.UserInterface = ArabicInterface Then
'        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «· ﬁ—Ì—"
'      Else
'     Msg = "No Data"
'  End If
'        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
'        Exit Sub
'    Else
' rs.MoveFirst
'
' print_report_Att_Comp StrSQL
'
'    End If

'End Sub
Function print_report_Att_Comp(Optional NoteSerial As String)
     
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

     If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "CrossStuAtt.rpt"
            Else
             StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "CrossStuAtt.rpt"
       End If
    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
        Cn.CommandTimeout = 10000
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
        'xReport.ParameterFields(1).AddCurrentValue Fromdate.value 'AddCurrentValue Fromdate 'RPTCompany_Name_Arabic
        'xReport.ParameterFields(2).AddCurrentValue toDate.value
        
        xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(CDate("01/" & (CmbMonth.ListIndex + 1) & "/" & (CboYear.ListIndex + 2006))), CDate("01/" & (CmbMonth.ListIndex + 1) & "/" & (CboYear.ListIndex + 2006)), CDate("01/" & (CmbMonth.ListIndex + 1) & "/" & (CboYear.ListIndex + 2006)))   'AddCurrentValue Fromdate 'RPTCompany_Name_Arabic
        xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(MonthLastDay(CDate("01/" & (CmbMonth.ListIndex + 1) & "/" & (CboYear.ListIndex + 2006)))), MonthLastDay(CDate("01/" & (CmbMonth.ListIndex + 1) & "/" & (CboYear.ListIndex + 2006))), MonthLastDay(CDate("01/" & (CmbMonth.ListIndex + 1) & "/" & (CboYear.ListIndex + 2006)))) ' toDate.value
        
        xReport.ParameterFields(3).AddCurrentValue cCompanyInfo.ArabCompanyName
        
        StrReportTitle = "" '& StrAccountName
    Else
        'xReport.ParameterFields(1).AddCurrentValue  ' RPTCompany_Name_Eng
        'StrReportTitle = ""
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
Public Sub GetStuInfoData()
    Dim StrSQL As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
 StrSQL = "SELECT Name, FullCode, UQama FROM dbo.TblStudent"
 StrSQL = StrSQL & "  where  (BranchID=0 or BranchID is null or         BranchID in(" & Current_branchSql & "))"
    
    '---------------------------------
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
           ' Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
          '  Me.lbl(10).Caption = "Search Results=0"
        End If
If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «· ﬁ—Ì—"
      Else
     Msg = "No Data"
  End If
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    Else
 rs.MoveFirst

 print_report_StuInfo StrSQL
    End If

End Sub


Function print_report_Att(Optional NoteSerial As String, Optional Ty As Integer = 0)
     
    Set rs = New ADODB.Recordset
    rs.Open NoteSerial, Cn, adOpenStatic, adLockReadOnly, adCmdText
        Dim monthdays As Integer
        Dim strmonthdays As String
        Dim datemonthdays As Date
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
If Ty = 0 Then
     If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "StuAtt.rpt"
            Else
             StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "StuAtt.rpt"
       End If
   Else
      If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "StuAtt2.rpt"
            Else
             StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "StuAtt2.rpt"
       End If
   End If
    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If
Dim DiFFNo As Double
Dim i As Integer
    Set RsData = New ADODB.Recordset
        Cn.CommandTimeout = 10000
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
    If Ty = 0 Then
        xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(CDate("01/" & (CmbMonth.ListIndex + 1) & "/" & (CboYear.ListIndex + 2006))), CDate("01/" & (CmbMonth.ListIndex + 1) & "/" & (CboYear.ListIndex + 2006)), CDate("01/" & (CmbMonth.ListIndex + 1) & "/" & (CboYear.ListIndex + 2006)))   'AddCurrentValue Fromdate 'RPTCompany_Name_Arabic
        xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(MonthLastDay(CDate("01/" & (CmbMonth.ListIndex + 1) & "/" & (CboYear.ListIndex + 2006)))), MonthLastDay(CDate("01/" & (CmbMonth.ListIndex + 1) & "/" & (CboYear.ListIndex + 2006))), MonthLastDay(CDate("01/" & (CmbMonth.ListIndex + 1) & "/" & (CboYear.ListIndex + 2006)))) ' toDate.value
         monthdays = day(MonthLastDay(CDate("01/" & (CmbMonth.ListIndex + 1) & "/" & (CboYear.ListIndex + 2006))))
        
        xReport.ParameterFields(4).AddCurrentValue monthdays
        Else
                xReport.ParameterFields(1).AddCurrentValue FromDate.value
        xReport.ParameterFields(2).AddCurrentValue ToDate.value
     End If
        xReport.ParameterFields(3).AddCurrentValue cCompanyInfo.ArabCompanyName
        
     

        StrReportTitle = "" '& StrAccountName
If Ty = 1 Then
DiFFNo = DateDiff("d", FromDate.value, ToDate.value) + 1
For i = 1 To DiFFNo
If i = 1 Then
xReport.ParameterFields(i + 4).AddCurrentValue FromDate.value
Else
xReport.ParameterFields(i + 4).AddCurrentValue DateAdd("d", i - 1, FromDate.value)
End If
Next i
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


Function print_report_StuInfo(Optional NoteSerial As String, Optional Ind As Integer = 0)
     
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
If Ind = 1 Then
 If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepStuInGrop.rpt"
            Else
             StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepStuInGrop.rpt"
       End If
Else
     If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "StuInfo.rpt"
            Else
             StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "StuInfo.rpt"
       End If
End If
    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
        Cn.CommandTimeout = 10000
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
        StrReportTitle = "" '& StrAccountName
    Else
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
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
Sub StudentInGroup()
Dim My_SQL1 As String
Dim StrWhere As String
 My_SQL1 = " SELECT     dbo.TblStuFingerprint.ID, dbo.TblStuGroup.Name, dbo.TblStuGroup.NameE, dbo.TblStuFingerprint.GroupID, dbo.TblStuFingerprint.StudID,"
 My_SQL1 = My_SQL1 & "                     dbo.TblStudent.Name AS StudName, dbo.TblStudent.NameE AS StudNameE, dbo.TblStudent.FullCode AS StudFullCode, dbo.TblStuFingerprint.CompID,"
 My_SQL1 = My_SQL1 & "                      dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode AS CusFullcode, dbo.TblStuFingerprint.CursID,"
 My_SQL1 = My_SQL1 & "                      dbo.TblStudentCurs.Name AS CursName, dbo.TblStudentCurs.NameE AS CursNameE, dbo.TblStuFingerprint.GDateH, dbo.TblStuFingerprint.GDate,"
 My_SQL1 = My_SQL1 & "                      dbo.TblStuFingerprint.FrmTime, dbo.TblStuFingerprint.ToTime, dbo.TblStuFingerprint.Fingerprint, dbo.TblStuFingerprint.Fingerprint2, dbo.TblStuFingerprint.DiffTime,"
 My_SQL1 = My_SQL1 & "                      dbo.TblStuFingerprint.ActTime, dbo.TblStuFingerprint.HallID, dbo.TblStudentClassRooms.Name AS HallName, dbo.TblStudentClassRooms.NameE AS HallNameE,"
 My_SQL1 = My_SQL1 & "                      dbo.TblStuFingerprint.DoplomID, dbo.TblStudentTypeCurs.Name AS DeplomName, dbo.TblStudentTypeCurs.NameE AS DeplomNameE,"
 My_SQL1 = My_SQL1 & "                      dbo.TblStuFingerprint.InstructID, dbo.TblInstructors.Name AS InstrName, dbo.TblInstructors.NameE AS InstrNameE, dbo.TblInstructors.FullCode AS InstrFullCode,"
 My_SQL1 = My_SQL1 & "                      dbo.TblStudent.StutsID , dbo.TblStuFingerprint.brnchid, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee"
 My_SQL1 = My_SQL1 & "  FROM         dbo.TblStuFingerprint LEFT OUTER JOIN"
 My_SQL1 = My_SQL1 & "                      dbo.TblBranchesData ON dbo.TblStuFingerprint.BrnchID = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
 My_SQL1 = My_SQL1 & "                      dbo.TblInstructors ON dbo.TblStuFingerprint.InstructID = dbo.TblInstructors.ID LEFT OUTER JOIN"
 My_SQL1 = My_SQL1 & "                      dbo.TblStudentTypeCurs ON dbo.TblStuFingerprint.DoplomID = dbo.TblStudentTypeCurs.ID LEFT OUTER JOIN"
 My_SQL1 = My_SQL1 & "                      dbo.TblStudentClassRooms ON dbo.TblStuFingerprint.HallID = dbo.TblStudentClassRooms.ID LEFT OUTER JOIN"
 My_SQL1 = My_SQL1 & "                      dbo.TblStudentCurs ON dbo.TblStuFingerprint.CursID = dbo.TblStudentCurs.ID LEFT OUTER JOIN"
 My_SQL1 = My_SQL1 & "                      dbo.TblCustemers ON dbo.TblStuFingerprint.CompID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
 My_SQL1 = My_SQL1 & "                      dbo.TblStudent ON dbo.TblStuFingerprint.StudID = dbo.TblStudent.ID LEFT OUTER JOIN"
 My_SQL1 = My_SQL1 & "                      dbo.TblStuGroup ON dbo.TblStuFingerprint.GroupID = dbo.TblStuGroup.ID"
' My_SQL1 = My_SQL1 & " where 1=1"
'My_SQL1 = My_SQL1 & " WHERE (dbo.TblStudent.EndDate >=" & SQLDate(startingDate, True) & " or dbo.TblStudent.EndDate is null)"
My_SQL1 = My_SQL1 & "  WHERE  (dbo.TblStuFingerprint.BrnchID=0 or dbo.TblStuFingerprint.BrnchID is null or         dbo.TblStuFingerprint.BrnchID in(" & Current_branchSql & "))"
StrWhere = ""
  If val(Me.DcbBranch2.BoundText) <> 0 Or Me.DcbBranch2.Text <> "" Then
  StrWhere = StrWhere & " AND dbo.TblStuFingerprint.BrnchID = " & val(DcbBranch2.BoundText)
End If

 If val(Me.DcbCompany2.BoundText) <> 0 Or Me.DcbCompany2.Text <> "" Then
  StrWhere = StrWhere & " AND dbo.TblStuFingerprint.CompID = " & val(DcbCompany2.BoundText)
End If
If val(Me.groupDBox2.BoundText) <> 0 Or Me.groupDBox2.Text <> "" Then
  StrWhere = StrWhere & " AND dbo.TblStuFingerprint.GroupID = " & val(groupDBox2.BoundText)
End If

If val(Me.cursBox2.BoundText) <> 0 Or Me.cursBox2.Text <> "" Then
  StrWhere = StrWhere & " AND dbo.TblStuFingerprint.CursID = " & val(cursBox2.BoundText)
End If
If val(Me.cursBox2.BoundText) <> 0 Or Me.cursBox2.Text <> "" Then
  StrWhere = StrWhere & " AND dbo.TblStuFingerprint.CursID = " & val(cursBox2.BoundText)
End If
If val(Me.DcbStudent2.BoundText) <> 0 Or Me.DcbStudent2.Text <> "" Then
  StrWhere = StrWhere & " AND dbo.TblStuFingerprint.StudID = " & val(DcbStudent2.BoundText)
End If
If Not IsNull(FromDate2.value) Then
  StrWhere = StrWhere & " AND dbo.TblStuFingerprint.GDate  >= " & SQLDate(FromDate2.value, True)
End If
If Not IsNull(ToDate2.value) Then
  StrWhere = StrWhere & " AND dbo.TblStuFingerprint.GDate  <= " & SQLDate(ToDate2.value, True)
End If
If Rd(0).value = True Then
StrWhere = StrWhere & " AND dbo.TblStuFingerprint.Fingerprint2  =1"
End If
If Rd(1).value = True Then
StrWhere = StrWhere & " AND dbo.TblStuFingerprint.Fingerprint2 is null"
End If
 'StrWhere = StrWhere & " AND (dbo.TblStuFingerprint.FlgGrpuoUpdae  is null or dbo.TblStuFingerprint.FlgGrpuoUpdae=1 or dbo.TblStuFingerprint.FlgGrpuoUpdae=0 )"
'StrWhere = StrWhere & " order by TblStuFingerprint.StudID"
My_SQL1 = My_SQL1 & StrWhere
  print_report_StuInfo My_SQL1, 1
End Sub
Private Sub FillStuRepTable1111()
Dim My_SQL1 As String
Dim My_SQL2 As String

Dim StrWhere As String
Dim j As Integer
Dim StartDate As Date
Dim tempHDay As Integer
Dim Moonth As Integer
Dim i As Integer
Dim TempDate As Date
Dim dayNumber As Integer
Dim isFirstTime As Boolean
Dim tempStuId As Integer
Dim numOfAttdays As Integer
Dim numOfVecDays As Integer
Dim checkRibDate As Date
Dim DyStus As Integer
Dim Abscen As Integer

Dim startingDate As Date
Dim EndingDate As Date
Dim m As Integer
Dim DIFFdAY As Double
startingDate = CDate("01/" & (CmbMonth.ListIndex + 1) & "/" & (CboYear.ListIndex + 2006))
EndingDate = MonthLastDay(startingDate)
DIFFdAY = Abs(DateDiff("d", startingDate, EndingDate)) + 1

Dim StudID As Double
Dim Rs1 As ADODB.Recordset
Set Rs1 = New ADODB.Recordset

Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
Moonth = Month(startingDate)
Dim Yeear As Integer
Yeear = year(startingDate)
Dim startDateStr As String

'######################################## setup the table first #########################################
 Cn.Execute "delete from TblStuRepTab"
'Rs3.Open My_SQL3, Cn, adOpenStatic, adLockOptimistic, adCmdText
'########################################################################################################

 My_SQL1 = " SELECT     dbo.TblStuFingerprint.ID, dbo.TblStuGroup.Name, dbo.TblStuGroup.NameE, dbo.TblStuFingerprint.GroupID, dbo.TblStuFingerprint.StudID,"
 My_SQL1 = My_SQL1 & "                     dbo.TblStudent.Name AS StudName, dbo.TblStudent.NameE AS StudNameE, dbo.TblStudent.FullCode AS StudFullCode, dbo.TblStuFingerprint.CompID,"
 My_SQL1 = My_SQL1 & "                      dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode AS CusFullcode, dbo.TblStuFingerprint.CursID,"
 My_SQL1 = My_SQL1 & "                      dbo.TblStudentCurs.Name AS CursName, dbo.TblStudentCurs.NameE AS CursNameE, dbo.TblStuFingerprint.GDateH, dbo.TblStuFingerprint.GDate,"
 My_SQL1 = My_SQL1 & "                      dbo.TblStuFingerprint.FrmTime, dbo.TblStuFingerprint.ToTime, dbo.TblStuFingerprint.Fingerprint, dbo.TblStuFingerprint.Fingerprint2, dbo.TblStuFingerprint.DiffTime,"
 My_SQL1 = My_SQL1 & "                      dbo.TblStuFingerprint.ActTime, dbo.TblStuFingerprint.HallID, dbo.TblStudentClassRooms.Name AS HallName, dbo.TblStudentClassRooms.NameE AS HallNameE,"
 My_SQL1 = My_SQL1 & "                      dbo.TblStuFingerprint.DoplomID, dbo.TblStudentTypeCurs.Name AS DeplomName, dbo.TblStudentTypeCurs.NameE AS DeplomNameE,"
 My_SQL1 = My_SQL1 & "                      dbo.TblStuFingerprint.InstructID, dbo.TblInstructors.Name AS InstrName, dbo.TblInstructors.NameE AS InstrNameE, dbo.TblInstructors.FullCode AS InstrFullCode,"
 My_SQL1 = My_SQL1 & "                      dbo.TblStudent.StutsID , dbo.TblStuFingerprint.brnchid, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee"
 My_SQL1 = My_SQL1 & "  FROM         dbo.TblStuFingerprint LEFT OUTER JOIN"
 My_SQL1 = My_SQL1 & "                      dbo.TblBranchesData ON dbo.TblStuFingerprint.BrnchID = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
 My_SQL1 = My_SQL1 & "                      dbo.TblInstructors ON dbo.TblStuFingerprint.InstructID = dbo.TblInstructors.ID LEFT OUTER JOIN"
 My_SQL1 = My_SQL1 & "                      dbo.TblStudentTypeCurs ON dbo.TblStuFingerprint.DoplomID = dbo.TblStudentTypeCurs.ID LEFT OUTER JOIN"
 My_SQL1 = My_SQL1 & "                      dbo.TblStudentClassRooms ON dbo.TblStuFingerprint.HallID = dbo.TblStudentClassRooms.ID LEFT OUTER JOIN"
 My_SQL1 = My_SQL1 & "                      dbo.TblStudentCurs ON dbo.TblStuFingerprint.CursID = dbo.TblStudentCurs.ID LEFT OUTER JOIN"
 My_SQL1 = My_SQL1 & "                      dbo.TblCustemers ON dbo.TblStuFingerprint.CompID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
 My_SQL1 = My_SQL1 & "                      dbo.TblStudent ON dbo.TblStuFingerprint.StudID = dbo.TblStudent.ID LEFT OUTER JOIN"
 My_SQL1 = My_SQL1 & "                      dbo.TblStuGroup ON dbo.TblStuFingerprint.GroupID = dbo.TblStuGroup.ID"
' My_SQL1 = My_SQL1 & " where 1=1"
My_SQL1 = My_SQL1 & " WHERE (dbo.TblStudent.EndDate >=" & SQLDate(startingDate, True) & " or dbo.TblStudent.EndDate is null)"
My_SQL1 = My_SQL1 & "  and  (dbo.TblStuFingerprint.BrnchID=0 or dbo.TblStuFingerprint.BrnchID is null or         dbo.TblStuFingerprint.BrnchID in(" & Current_branchSql & "))"
 StrWhere = ""

If val(Me.groupDBox.BoundText) <> 0 Or Me.groupDBox.Text <> "" Then
  StrWhere = StrWhere & " AND dbo.TblStuFingerprint.GroupID = " & val(groupDBox.BoundText)
End If

If val(Me.cursBox.BoundText) <> 0 Or Me.cursBox.Text <> "" Then
  StrWhere = StrWhere & " AND dbo.TblStuFingerprint.CursID = " & val(cursBox.BoundText)
End If
If val(CmbMonth.ListIndex) <> -1 And CmbMonth.Text <> "" Then
StrWhere = StrWhere & " AND month(dbo.TblStuFingerprint.GDate)= " & val(CmbMonth.ListIndex + 1)
End If
If val(CboYear.ListIndex) <> -1 And CboYear.Text <> "" Then
StrWhere = StrWhere & " AND year(dbo.TblStuFingerprint.GDate)= " & val(CboYear.ListIndex) + 2006
End If

 StrWhere = StrWhere & " AND (dbo.TblStuFingerprint.FlgGrpuoUpdae  is null or dbo.TblStuFingerprint.FlgGrpuoUpdae=1 or dbo.TblStuFingerprint.FlgGrpuoUpdae=0 )"
'  StrWhere = StrWhere & " AND TblStuFingerprint.StudID = 86"
 
StrWhere = StrWhere & " order by TblStuFingerprint.StudID"
My_SQL1 = My_SQL1 & StrWhere
    
My_SQL2 = "Select * from TblStuRepTab where 1 = -1"

Rs1.Open My_SQL1, Cn, adOpenStatic, adLockReadOnly, adCmdText
rs2.Open My_SQL2, Cn, adOpenStatic, adLockOptimistic, adCmdText

If Rs1.RecordCount > 0 Then
Rs1.MoveFirst
End If
Dim IsAttend As Boolean

For i = 1 To Rs1.RecordCount
IsAttend = IIf(IsNull(Rs1.Fields("Fingerprint2").value), False, Rs1.Fields("Fingerprint2").value)
 tempStuId = IIf(IsNull(Rs1.Fields("StutsID").value), 0, Rs1.Fields("StutsID").value)
TempDate = IIf(IsNull(Rs1.Fields("GDate").value), "", Rs1.Fields("GDate").value)
If i = 1 Then
StudID = IIf(IsNull(Rs1.Fields("StudID").value), 0, Rs1.Fields("StudID").value)
startDateStr = "01/" & Moonth & "/" & Yeear
StartDate = CDate(startDateStr)
m = 0
Abscen = 0
rs2.AddNew
End If
If StudID <> IIf(IsNull(Rs1.Fields("StudID").value), 0, Rs1.Fields("StudID").value) Then
StudID = IIf(IsNull(Rs1.Fields("StudID").value), 0, Rs1.Fields("StudID").value)
startDateStr = "01/" & Moonth & "/" & Yeear
StartDate = CDate(startDateStr)
m = 0
Abscen = 0
rs2.AddNew
End If

rs2.Fields("stuId").value = IIf(IsNull(Rs1.Fields("StudID").value), 0, Rs1.Fields("StudID").value)

rs2.Fields("UserName").value = DcbUserName.Text
If SystemOptions.UserInterface = ArabicInterface Then
rs2.Fields("StuName").value = IIf(IsNull(Rs1.Fields("StudName").value), "", Rs1.Fields("StudName").value)
rs2.Fields("GroupName").value = IIf(IsNull(Rs1.Fields("Name").value), "", Rs1.Fields("Name").value)
rs2.Fields("CursName").value = IIf(IsNull(Rs1.Fields("CursName").value), "", Rs1.Fields("CursName").value)
rs2.Fields("InstName").value = IIf(IsNull(Rs1.Fields("InstrName").value), "", Rs1.Fields("InstrName").value)
rs2.Fields("BranchName").value = IIf(IsNull(Rs1.Fields("branch_name").value), "", Rs1.Fields("branch_name").value)
rs2.Fields("DeplomName").value = IIf(IsNull(Rs1.Fields("DeplomName").value), "", Rs1.Fields("DeplomName").value)
Else
rs2.Fields("BranchName").value = IIf(IsNull(Rs1.Fields("branch_namee").value), "", Rs1.Fields("branch_namee").value)
rs2.Fields("InstName").value = IIf(IsNull(Rs1.Fields("InstrNameE").value), "", Rs1.Fields("InstrNameE").value)
rs2.Fields("CursName").value = IIf(IsNull(Rs1.Fields("CursNameE").value), "", Rs1.Fields("CursNameE").value)
rs2.Fields("StuName").value = IIf(IsNull(Rs1.Fields("StudNameE").value), "", Rs1.Fields("StudNameE").value)
rs2.Fields("GroupName").value = IIf(IsNull(Rs1.Fields("NameE").value), "", Rs1.Fields("NameE").value)
rs2.Fields("DeplomName").value = IIf(IsNull(Rs1.Fields("DeplomNameE").value), "", Rs1.Fields("DeplomNameE").value)
End If

'######################################## get Vec days ###################################
If m = 0 Then
For j = 1 To DIFFdAY

If tempStuId <> 0 And CheckEndStudent(StartDate, StudID) = 1 Then
DyStus = -3
ElseIf CheckNoInGroup(StartDate) = 0 And CheckHolidaies(StartDate) <> 1 Then
DyStus = -1 '''€Ì— „ÊÃ„Êœ
ElseIf CheckHolidaies(StartDate) = 1 Then
DyStus = -2
Else
If IsAttend = True And j = 1 Then
DyStus = 1
Else
DyStus = 0
If TempDate = StartDate Then
Abscen = Abscen + 1
End If
End If
End If
tempHDay = day(StartDate)
If tempHDay <> 0 Then
rs2.Fields("D" & tempHDay).value = DyStus
End If
rs2.Fields("NumOfAttDays").value = Abscen
rs2.update
StartDate = DateAdd("d", 1, StartDate)
Next j
m = 1
Else

If tempStuId <> 0 And CheckEndStudent(TempDate, StudID) = 1 Then
DyStus = -3

Else
If IsAttend = True Then
DyStus = 1
Else
DyStus = 0
Abscen = Abscen + 1
End If
End If
tempHDay = day(TempDate)

If tempHDay <> 0 Then
rs2.Fields("D" & tempHDay).value = DyStus
End If
rs2.Fields("NumOfAttDays").value = Abscen
rs2.update

End If
Rs1.MoveNext
Next i
End Sub
Function CheckEndStudent(Optional TempDate As Date, Optional StudID As Double) As Integer
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = " SELECT     ID, EndDate"
sql = sql & " FROM         dbo.TblStudent"
sql = sql & " where EndDate >" & SQLDate(TempDate, True) & " and id=" & StudID & ""
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
CheckEndStudent = 0
Else
CheckEndStudent = 1
End If
End Function
Private Sub FillStuRepTable22()
Dim My_SQL1 As String
Dim My_SQL2 As String

Dim StrWhere As String
Dim j As Integer
Dim StartDate As Date
Dim tempHDay As Integer
Dim Moonth As Integer
Dim i As Integer
Dim TempDate As Date
Dim dayNumber As Integer
Dim isFirstTime As Boolean
Dim tempStuId As Integer
Dim numOfAttdays As Integer
Dim numOfVecDays As Integer
Dim checkRibDate As Date
Dim DyStus As Integer
Dim Abscen As Integer

Dim m As Integer

Dim StudID As Double
Dim Rs1 As ADODB.Recordset
Set Rs1 = New ADODB.Recordset

Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset


'######################################## setup the table first #########################################
 Cn.Execute "delete from TblStuRepTab2"
'Rs3.Open My_SQL3, Cn, adOpenStatic, adLockOptimistic, adCmdText
'########################################################################################################

  My_SQL1 = " SELECT     dbo.TblStuFingerprint.ID, dbo.TblStuGroup.Name, dbo.TblStuGroup.NameE, dbo.TblStuFingerprint.GroupID, dbo.TblStuFingerprint.StudID,"
  My_SQL1 = My_SQL1 & "                      dbo.TblStudent.Name AS StudName, dbo.TblStudent.NameE AS StudNameE, dbo.TblStudent.FullCode AS StudFullCode, dbo.TblStuFingerprint.CompID,"
  My_SQL1 = My_SQL1 & "                     dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode AS CusFullcode, dbo.TblStuFingerprint.CursID,"
  My_SQL1 = My_SQL1 & "                     dbo.TblStudentCurs.Name AS CursName, dbo.TblStudentCurs.NameE AS CursNameE, dbo.TblStuFingerprint.GDateH, dbo.TblStuFingerprint.GDate,"
  My_SQL1 = My_SQL1 & "                     dbo.TblStuFingerprint.FrmTime, dbo.TblStuFingerprint.ToTime, dbo.TblStuFingerprint.Fingerprint, dbo.TblStuFingerprint.Fingerprint2, dbo.TblStuFingerprint.DiffTime,"
  My_SQL1 = My_SQL1 & "                     dbo.TblStuFingerprint.ActTime, dbo.TblStuFingerprint.HallID, dbo.TblStudentClassRooms.Name AS HallName, dbo.TblStudentClassRooms.NameE AS HallNameE,"
  My_SQL1 = My_SQL1 & "                     dbo.TblStuFingerprint.DoplomID, dbo.TblStudentTypeCurs.Name AS DeplomName, dbo.TblStudentTypeCurs.NameE AS DeplomNameE,"
  My_SQL1 = My_SQL1 & "                     dbo.TblStuFingerprint.InstructID, dbo.TblInstructors.Name AS InstrName, dbo.TblInstructors.NameE AS InstrNameE, dbo.TblInstructors.FullCode AS InstrFullCode,"
  My_SQL1 = My_SQL1 & "                     dbo.TblStudent.StutsID , dbo.TblStuFingerprint.brnchid, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee"
  My_SQL1 = My_SQL1 & " FROM         dbo.TblStuFingerprint LEFT OUTER JOIN"
  My_SQL1 = My_SQL1 & "                     dbo.TblBranchesData ON dbo.TblStuFingerprint.BrnchID = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
  My_SQL1 = My_SQL1 & "                     dbo.TblInstructors ON dbo.TblStuFingerprint.InstructID = dbo.TblInstructors.ID LEFT OUTER JOIN"
  My_SQL1 = My_SQL1 & "                     dbo.TblStudentTypeCurs ON dbo.TblStuFingerprint.DoplomID = dbo.TblStudentTypeCurs.ID LEFT OUTER JOIN"
  My_SQL1 = My_SQL1 & "                      dbo.TblStudentClassRooms ON dbo.TblStuFingerprint.HallID = dbo.TblStudentClassRooms.ID LEFT OUTER JOIN"
  My_SQL1 = My_SQL1 & "                     dbo.TblStudentCurs ON dbo.TblStuFingerprint.CursID = dbo.TblStudentCurs.ID LEFT OUTER JOIN"
  My_SQL1 = My_SQL1 & "                      dbo.TblCustemers ON dbo.TblStuFingerprint.CompID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
  My_SQL1 = My_SQL1 & "                     dbo.TblStudent ON dbo.TblStuFingerprint.StudID = dbo.TblStudent.ID LEFT OUTER JOIN"
  My_SQL1 = My_SQL1 & "                     dbo.TblStuGroup ON dbo.TblStuFingerprint.GroupID = dbo.TblStuGroup.ID"
  'My_SQL1 = My_SQL1 & " WHERE ((dbo.TblStudent.EndDate >=" & SQLDate(FromDate.value, True) & ") or dbo.TblStudent.EndDate is null)"
  My_SQL1 = My_SQL1 & "  WHERE   (dbo.TblStuFingerprint.BrnchID=0 or dbo.TblStuFingerprint.BrnchID is null or         dbo.TblStuFingerprint.BrnchID in(" & Current_branchSql & "))"
    
 StrWhere = ""
 If val(Me.DcbBranch.BoundText) <> 0 Or Me.DcbBranch.Text <> "" Then
  StrWhere = StrWhere & " AND dbo.TblStuFingerprint.BrnchID = " & val(DcbBranch.BoundText)
End If

If val(Me.DcbCompany.BoundText) <> 0 Or Me.DcbCompany.Text <> "" Then
  StrWhere = StrWhere & " AND dbo.TblStuFingerprint.CompID  = " & val(DcbCompany.BoundText)
End If
If Not IsNull(FromDate.value) Then
  StrWhere = StrWhere & " AND dbo.TblStuFingerprint.GDate  >= " & SQLDate(FromDate.value, True)
End If
If Not IsNull(ToDate.value) Then
  StrWhere = StrWhere & " AND dbo.TblStuFingerprint.GDate  <= " & SQLDate(ToDate.value, True)
End If

 StrWhere = StrWhere & " AND (dbo.TblStuFingerprint.FlgGrpuoUpdae  is null or dbo.TblStuFingerprint.FlgGrpuoUpdae=1 or dbo.TblStuFingerprint.FlgGrpuoUpdae=0 )"
'StrWhere = StrWhere & " AND  TblStuFingerprint.StudID = 137"
StrWhere = StrWhere & " order by TblStuFingerprint.StudID ,TblStuFingerprint.GDate"
My_SQL1 = My_SQL1 & StrWhere
    
My_SQL2 = "Select * from TblStuRepTab2 where 1 = -1"

Rs1.Open My_SQL1, Cn, adOpenStatic, adLockReadOnly, adCmdText
rs2.Open My_SQL2, Cn, adOpenStatic, adLockOptimistic, adCmdText

If Rs1.RecordCount > 0 Then
Rs1.MoveFirst
End If
Dim IsAttend As Boolean

For i = 1 To Rs1.RecordCount
IsAttend = IIf(IsNull(Rs1.Fields("Fingerprint2").value), False, Rs1.Fields("Fingerprint2").value)
 tempStuId = IIf(IsNull(Rs1.Fields("StutsID").value), 0, Rs1.Fields("StutsID").value)
TempDate = IIf(IsNull(Rs1.Fields("GDate").value), "", Rs1.Fields("GDate").value)
If i = 1 Then
StudID = IIf(IsNull(Rs1.Fields("StudID").value), 0, Rs1.Fields("StudID").value)
m = 0
Abscen = 0
rs2.AddNew
End If
If StudID <> IIf(IsNull(Rs1.Fields("StudID").value), 0, Rs1.Fields("StudID").value) Then
StudID = IIf(IsNull(Rs1.Fields("StudID").value), 0, Rs1.Fields("StudID").value)
m = 0
Abscen = 0
rs2.AddNew
End If
rs2.Fields("TypeTrans").value = 1
rs2.Fields("stuId").value = IIf(IsNull(Rs1.Fields("StudID").value), 0, Rs1.Fields("StudID").value)
rs2.Fields("UserName").value = DcbUserName.Text
If SystemOptions.UserInterface = ArabicInterface Then
rs2.Fields("company").value = IIf(IsNull(Rs1.Fields("CusName").value), "", Rs1.Fields("CusName").value)
rs2.Fields("StuName").value = IIf(IsNull(Rs1.Fields("StudName").value), "", Rs1.Fields("StudName").value)
rs2.Fields("GroupName").value = IIf(IsNull(Rs1.Fields("Name").value), "", Rs1.Fields("Name").value)
rs2.Fields("CursName").value = IIf(IsNull(Rs1.Fields("CursName").value), "", Rs1.Fields("CursName").value)
rs2.Fields("InstName").value = IIf(IsNull(Rs1.Fields("InstrName").value), "", Rs1.Fields("InstrName").value)
rs2.Fields("BranchName").value = IIf(IsNull(Rs1.Fields("branch_name").value), "", Rs1.Fields("branch_name").value)
rs2.Fields("DeplomName").value = IIf(IsNull(Rs1.Fields("DeplomName").value), "", Rs1.Fields("DeplomName").value)
Else
rs2.Fields("DeplomName").value = IIf(IsNull(Rs1.Fields("DeplomNameE").value), "", Rs1.Fields("DeplomNameE").value)
rs2.Fields("company").value = IIf(IsNull(Rs1.Fields("CusNamee").value), "", Rs1.Fields("CusNamee").value)
rs2.Fields("BranchName").value = IIf(IsNull(Rs1.Fields("branch_namee").value), "", Rs1.Fields("branch_namee").value)
rs2.Fields("InstName").value = IIf(IsNull(Rs1.Fields("InstrNameE").value), "", Rs1.Fields("InstrNameE").value)
rs2.Fields("CursName").value = IIf(IsNull(Rs1.Fields("CursNameE").value), "", Rs1.Fields("CursNameE").value)
rs2.Fields("StuName").value = IIf(IsNull(Rs1.Fields("StudNameE").value), "", Rs1.Fields("StudNameE").value)
rs2.Fields("GroupName").value = IIf(IsNull(Rs1.Fields("NameE").value), "", Rs1.Fields("NameE").value)
End If

'######################################## get Vec days ###################################


Dim DiFFNo As Integer
DiFFNo = DateDiff("d", FromDate.value, ToDate.value) + 1
StartDate = FromDate.value
If m = 0 And StartDate <= ToDate.value Then
For j = 1 To DiFFNo
If tempStuId <> 0 And CheckEndStudent(StartDate, StudID) = 1 Then
DyStus = -3
ElseIf CheckNoInGroup(StartDate) = 0 And CheckHolidaies(StartDate) <> 1 Then
DyStus = 0 '''€Ì— „ÊÃ„Êœ
'Abscen = Abscen + 1
ElseIf CheckHolidaies(StartDate) = 1 Then
DyStus = -2
Else
'If IsAttend = True And TempDate = StartDate Then
If IsAttend = True And j = 1 Then
DyStus = 1
Else
DyStus = 0

If TempDate = StartDate Then
Abscen = Abscen + 1
End If
End If
End If
If StartDate <= ToDate.value Then
tempHDay = GetNoDay(StartDate)
Else
tempHDay = day(StartDate)
End If
If tempHDay <> 0 Then
If IsNull(rs2.Fields("D" & tempHDay).value) Then
rs2.Fields("D" & tempHDay).value = DyStus
ElseIf rs2.Fields("D" & tempHDay).value <> 1 Then
rs2.Fields("D" & tempHDay).value = DyStus
End If
End If
rs2.Fields("NumOfAttDays").value = Abscen

rs2.update
StartDate = DateAdd("d", 1, StartDate)
Next j
m = 1
Else

If tempStuId <> 0 And CheckEndStudent(TempDate, StudID) = 1 And IsAttend = False Then
DyStus = -3

ElseIf CheckHolidaies(TempDate) = 1 Then
DyStus = -2

Else
If IsAttend = True Then
DyStus = 1

Else
DyStus = 0
Abscen = Abscen + 1
End If
End If
tempHDay = GetNoDay(TempDate)
If tempHDay <> 0 Then
If IsNull(rs2.Fields("D" & tempHDay).value) Then
rs2.Fields("D" & tempHDay).value = DyStus
ElseIf rs2.Fields("D" & tempHDay).value <> 1 Then
rs2.Fields("D" & tempHDay).value = DyStus
End If
'Rs2.Fields("D" & tempHDay).value = DyStus

rs2.Fields("NumOfAttDays").value = Abscen
End If
rs2.update

End If
Rs1.MoveNext
Next i
'''///////////
Dim Sql2 As String
Dim Rs7 As ADODB.Recordset
Set Rs7 = New ADODB.Recordset
Dim Rs8 As ADODB.Recordset
Set Rs8 = New ADODB.Recordset
Dim k As Integer
Dim d As Integer
My_SQL2 = "Select * from TblStuRepTab2 where 1 = -1"
Rs7.Open My_SQL2, Cn, adOpenStatic, adLockOptimistic, adCmdText
Sql2 = "SELECT DISTINCT ID,Name,NameE, CompID, BranchID, Mobile, StutsID"
Sql2 = Sql2 & " From dbo.TblStudent"
Sql2 = Sql2 & " WHERE        (NOT (ID IN"
Sql2 = Sql2 & "                             (SELECT        StudID"
Sql2 = Sql2 & "                                From dbo.TblStuFingerprint"
Sql2 = Sql2 & "                                WHERE        (CompID = " & DcbCompany.BoundText & ") AND (GDate <= " & SQLDate(ToDate.value, True) & ") AND (GDate >= " & SQLDate(FromDate.value, True) & ")))) AND (CompID = " & val(DcbCompany.BoundText) & ") AND (StutsID = 0)"
Rs8.Open Sql2, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs8.RecordCount > 0 Then
Rs8.MoveFirst
For k = 1 To Rs8.RecordCount
Rs7.AddNew
Rs7.Fields("TypeTrans").value = 2
DiFFNo = DateDiff("d", FromDate.value, ToDate.value) + 1
For d = 1 To DiFFNo
Rs7.Fields("D" & d).value = 0
Next d
Rs7.Fields("stuId").value = IIf(IsNull(Rs8.Fields("ID").value), 0, Rs8.Fields("ID").value)
If SystemOptions.UserInterface = ArabicInterface Then
Rs7.Fields("StuName").value = IIf(IsNull(Rs8.Fields("Name").value), "", Rs8.Fields("Name").value)
Else
Rs7.Fields("StuName").value = IIf(IsNull(Rs8.Fields("NameE").value), "", Rs8.Fields("NameE").value)
End If
Rs7.update
Rs8.MoveNext
Next k
End If
End Sub

Function GetNoDay(Optional TempDate As Date) As Integer
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = "Select * from TblStuNoDay where RecordDate=" & SQLDate(TempDate, True) & ""
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
GetNoDay = IIf(IsNull(rs2("DayID").value), 0, rs2("DayID").value)
Else
GetNoDay = 0
End If
End Function
'Private Sub FillStuRepTable()
'Dim My_SQL1 As String
'Dim My_SQL2 As String
'Dim My_SQL3 As String
'Dim My_SQL4 As String
'
'Dim StrWhere As String
'
'Dim I As Integer
'Dim TempDate As Date
'Dim dayNumber As Integer
'Dim isFirstTime As Boolean
'Dim tempStuId As Integer
'Dim numOfAttdays As Integer
'Dim numOfVecDays As Integer
'Dim checkRibDate As Date
'
'Dim startingDate As Date
'Dim EndingDate As Date
'startingDate = CDate("01/" & (CmbMonth.ListIndex + 1) & "/" & (CboYear.ListIndex + 2006))
'EndingDate = MonthLastDay(startingDate)
'
'Dim rs1 As ADODB.Recordset
'Set rs1 = New ADODB.Recordset
'
'Dim Rs2 As ADODB.Recordset
'Set Rs2 = New ADODB.Recordset
'
'Dim Rs3 As ADODB.Recordset
'Set Rs3 = New ADODB.Recordset
'
'Dim Rs4 As ADODB.Recordset
'Set Rs4 = New ADODB.Recordset
''######################################## setup the table first #########################################
'My_SQL3 = "delete from TblStuRepTab"
'Rs3.Open My_SQL3, Cn, adOpenStatic, adLockOptimistic, adCmdText
''########################################################################################################
'
' 'My_SQL1 = "SELECT TblStudent.Name, TblStudent.ID, TblAttendance.RecordDate, TblAttendanceDet.IsAttend, TblStuGroup.Name AS GName, TblStudentCurs.Name AS CName, TblInstructors.Name AS IName ,month(dbo.TblAttendance.RecordDate) as monthID "
' 'My_SQL1 = My_SQL1 & "FROM TblAttendanceDet INNER JOIN "
' 'My_SQL1 = My_SQL1 & "TblAttendance ON TblAttendanceDet.AttenID = TblAttendance.ID INNER JOIN "
' 'My_SQL1 = My_SQL1 & "TblStudent ON TblAttendanceDet.StudID = TblStudent.ID INNER JOIN "
' 'My_SQL1 = My_SQL1 & "TblCustemers ON TblStudent.CompID = TblCustemers.CusID INNER JOIN "
' 'My_SQL1 = My_SQL1 & "TblInstructors ON TblAttendance.InstrcID = TblInstructors.ID INNER JOIN "
' 'My_SQL1 = My_SQL1 & "TblStuGroup ON TblAttendance.GroupID = TblStuGroup.ID INNER JOIN "
' 'My_SQL1 = My_SQL1 & "TblStudentCurs ON TblAttendance.CursID = TblStudentCurs.ID "
' 'My_SQL1 = My_SQL1 & " where 1=1"
' My_SQL1 = My_SQL1 & " SELECT     dbo.TblStuFingerprint.ID, dbo.TblStuGroup.Name, dbo.TblStuGroup.NameE, dbo.TblStuFingerprint.GroupID, dbo.TblStuFingerprint.StudID,"
' My_SQL1 = My_SQL1 & "                     dbo.TblStudent.Name AS StudName, dbo.TblStudent.NameE AS StudNameE, dbo.TblStudent.FullCode AS StudFullCode, dbo.TblStuFingerprint.CompID,"
' My_SQL1 = My_SQL1 & "                     dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode AS CusFullcode, dbo.TblStuFingerprint.CursID,"
' My_SQL1 = My_SQL1 & "                     dbo.TblStudentCurs.Name AS CursName, dbo.TblStudentCurs.NameE AS CursNameE, dbo.TblStuFingerprint.GDateH, dbo.TblStuFingerprint.GDate,"
' My_SQL1 = My_SQL1 & "                     dbo.TblStuFingerprint.FrmTime, dbo.TblStuFingerprint.ToTime, dbo.TblStuFingerprint.Fingerprint, dbo.TblStuFingerprint.Fingerprint2, dbo.TblStuFingerprint.DiffTime,"
' My_SQL1 = My_SQL1 & "                     dbo.TblStuFingerprint.ActTime, dbo.TblStuFingerprint.HallID, dbo.TblStudentClassRooms.Name AS HallName, dbo.TblStudentClassRooms.NameE AS HallNameE,"
' My_SQL1 = My_SQL1 & "                     dbo.TblStuFingerprint.DoplomID, dbo.TblStudentTypeCurs.Name AS DeplomName, dbo.TblStudentTypeCurs.NameE AS DeplomNameE,"
' My_SQL1 = My_SQL1 & "                     dbo.TblStuFingerprint.InstructID, dbo.TblInstructors.Name AS InstrName, dbo.TblInstructors.NameE AS InstrNameE, dbo.TblInstructors.FullCode AS InstrFullCode"
' My_SQL1 = My_SQL1 & "   FROM         dbo.TblStuFingerprint LEFT OUTER JOIN"
' My_SQL1 = My_SQL1 & "                     dbo.TblInstructors ON dbo.TblStuFingerprint.InstructID = dbo.TblInstructors.ID LEFT OUTER JOIN"
' My_SQL1 = My_SQL1 & "                     dbo.TblStudentTypeCurs ON dbo.TblStuFingerprint.DoplomID = dbo.TblStudentTypeCurs.ID LEFT OUTER JOIN"
' My_SQL1 = My_SQL1 & "                     dbo.TblStudentClassRooms ON dbo.TblStuFingerprint.HallID = dbo.TblStudentClassRooms.ID LEFT OUTER JOIN"
' My_SQL1 = My_SQL1 & "                     dbo.TblStudentCurs ON dbo.TblStuFingerprint.CursID = dbo.TblStudentCurs.ID LEFT OUTER JOIN"
' My_SQL1 = My_SQL1 & "                     dbo.TblCustemers ON dbo.TblStuFingerprint.CompID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
' My_SQL1 = My_SQL1 & "                     dbo.TblStudent ON dbo.TblStuFingerprint.StudID = dbo.TblStudent.ID LEFT OUTER JOIN"
' My_SQL1 = My_SQL1 & "                     dbo.TblStuGroup ON dbo.TblStuFingerprint.GroupID = dbo.TblStuGroup.ID"
' My_SQL1 = My_SQL1 & " where 1=1"
' StrWhere = ""
'
''If val(Me.DcbCompany.BoundText) <> 0 Or Me.DcbCompany.Text <> "" Then
''  StrWhere = StrWhere & " AND dbo.TblStuFingerprint.CompID  = " & val(DcbCompany.BoundText)
''End If
'If val(Me.groupDBox.BoundText) <> 0 Or Me.groupDBox.Text <> "" Then
'  StrWhere = StrWhere & " AND dbo.TblStuFingerprint.GroupID = " & val(groupDBox.BoundText)
'End If
'If val(Me.cursBox.BoundText) <> 0 Or Me.cursBox.Text <> "" Then
'  StrWhere = StrWhere & " AND dbo.TblStuFingerprint.CursID = " & val(cursBox.BoundText)
'End If
''If val(Me.instruDBox.BoundText) <> 0 Or Me.instruDBox.Text <> "" Then
''  StrWhere = StrWhere & " AND dbo.TblStuFingerprint.InstructID = " & val(instruDBox.BoundText)
''End If
''If Not IsNull(Me.FromDate.value) Then
''  StrWhere = StrWhere & " AND dbo.TblAttendance.RecordDate >=" & SQLDate(startingDate, True) & ""
'End If
'If Not IsNull(Me.ToDate.value) Then
'  StrWhere = StrWhere & " AND  dbo.TblAttendance.RecordDate <=" & SQLDate(EndingDate, True) & ""
'End If
'If val(CmbMonth.ListIndex) <> -1 And CmbMonth.Text <> "" Then
'StrWhere = StrWhere & " AND month(dbo.TblAttendance.RecordDate)= " & val(CmbMonth.ListIndex + 1)
'End If
'If val(CboYear.ListIndex) <> -1 And CboYear.Text <> "" Then
'StrWhere = StrWhere & " AND month(dbo.TblAttendance.RecordDate)= " & val(CboYear.ListIndex) + 2006
'End If
''StrWhere = StrWhere & "AND TblAttendanceDet.IsAttend = 1 order by TblStudent.id"
'StrWhere = StrWhere & " order by TblStuFingerprint.StudID"
'My_SQL1 = My_SQL1 & StrWhere
'
'My_SQL2 = "Select * from TblStuRepTab where 1 = -1"
'
'rs1.Open My_SQL1, Cn, adOpenStatic, adLockReadOnly, adCmdText
'Rs2.Open My_SQL2, Cn, adOpenStatic, adLockOptimistic, adCmdText
'
'If rs1.RecordCount > 0 Then
'rs1.MoveFirst
'End If
'Dim IsAttend As Boolean

'isFirstTime = True
'
'For I = 1 To rs1.RecordCount
'IsAttend = IIf(IsNull(rs1.Fields("Fingerprint2").value), False, rs1.Fields("Fingerprint2").value)
'If tempStuId = IIf(IsNull(rs1.Fields("StudID").value), "", rs1.Fields("StudID").value) Then

'TempDate = IIf(IsNull(rs1.Fields("RecordDate").value), "", rs1.Fields("RecordDate").value)


'If TempDate <> checkRibDate Then
'
'checkRibDate = TempDate
'
'dayNumber = Day(TempDate)
'
'numOfAttdays = numOfAttdays + 1
'
'Select Case dayNumber
'Case 1
'Rs2.Fields("D1").value = 1
'Case 2
'Rs2.Fields("D2").value = 1
'Case 3
'Rs2.Fields("D3").value = 1
'Case 4
'Rs2.Fields("D4").value = 1
'Case 5
'Rs2.Fields("D5").value = 1
'Case 6
'Rs2.Fields("D6").value = 1
'Case 7
'Rs2.Fields("D7").value = 1
'Case 8
'Rs2.Fields("D8").value = 1
'Case 9
'Rs2.Fields("D9").value = 1
'Case 10
'Rs2.Fields("D10").value = 1
'Case 11
'Rs2.Fields("D11").value = 1
'Case 12
'Rs2.Fields("D12").value = 1
'Case 13
'Rs2.Fields("D13").value = 1
'Case 14
'Rs2.Fields("D14").value = 1
'Case 15
'Rs2.Fields("D15").value = 1
'Case 16
'Rs2.Fields("D16").value = 1
'Case 17
'Rs2.Fields("D17").value = 1
'Case 18
'Rs2.Fields("D18").value = 1
'Case 19
'Rs2.Fields("D19").value = 1
'Case 20
'Rs2.Fields("D20").value = 1
'Case 21
'Rs2.Fields("D21").value = 1
'Case 22
'Rs2.Fields("D22").value = 1
'Case 23
'Rs2.Fields("D23").value = 1
'Case 24
'Rs2.Fields("D24").value = 1
'Case 25
'Rs2.Fields("D25").value = 1
'Case 26
'Rs2.Fields("D26").value = 1
'Case 27
'Rs2.Fields("D27").value = 1
'Case 28
'Rs2.Fields("D28").value = 1
'Case 29
'Rs2.Fields("D29").value = 1
'Case 30
'Rs2.Fields("D30").value = 1
'Case 31
'Rs2.Fields("D31").value = 1
'End Select
'End If
'Rs2.Update
'rs1.MoveNext
'
''Else
'
'
'If isFirstTime Then
'Rs2.AddNew
'isFirstTime = False
'Else
'Rs2.Fields("NumOfAttDays").value = Day(EndingDate) - (numOfAttdays + numOfVecDays)
'numOfAttdays = 0
'numOfVecDays = 0
'Rs2.AddNew
'End If
'
'
'tempStuId = IIf(IsNull(rs1.Fields("ID").value), "", rs1.Fields("ID").value)
'Rs2.Fields("stuId").value = tempStuId
'Rs2.Fields("StuName").value = IIf(IsNull(rs1.Fields("Name").value), "", rs1.Fields("Name").value)
'Rs2.Fields("GroupName").value = IIf(IsNull(rs1.Fields("GName").value), "", rs1.Fields("GName").value)
'Rs2.Fields("CursName").value = IIf(IsNull(rs1.Fields("CName").value), "", rs1.Fields("CName").value)
'Rs2.Fields("InstName").value = IIf(IsNull(rs1.Fields("IName").value), "", rs1.Fields("IName").value)
'
''######################################## get Vec days ###################################
'Dim j As Integer
'Dim StartDate As Date
'Dim HolidaiesCheck As Integer
'Dim tempHDay As Integer
'
'Dim Moonth As Integer
'Moonth = Month(startingDate)
'
'Dim Yeear As Integer
'Yeear = year(startingDate)
'
'Dim startDateStr As String
'startDateStr = "01/" & Moonth & "/" & Yeear
'
'StartDate = CDate(startDateStr)
'
'For j = 1 To 31
'
'
'HolidaiesCheck = CheckHolidaies(StartDate)
'
'If (HolidaiesCheck = 1) Then
'
'tempHDay = Day(StartDate)
'
'numOfVecDays = numOfVecDays + 1
''«·«Ã«“«  «·—”„Ì…
'Select Case tempHDay
'Case 1
'Rs2.Fields("D1").value = -1
'Case 2
'Rs2.Fields("D2").value = -1
'Case 3
'Rs2.Fields("D3").value = -1
'Case 4
'Rs2.Fields("D4").value = -1
'Case 5
''Rs2.Fields("D5").value = -1
'Case 6
'Rs2.Fields("D6").value = -1
'Case 7
'Rs2.Fields("D7").value = -1
'Case 8
'Rs2.Fields("D8").value = -1
'Case 9
'Rs2.Fields("D9").value = -1
'Case 10
'Rs2.Fields("D10").value = -1
'Case 11
'Rs2.Fields("D11").value = -1
'Case 12
'Rs2.Fields("D12").value = -1
'Case 13
'Rs2.Fields("D13").value = -1
'Case 14
'Rs2.Fields("D14").value = -1
'Case 15
'Rs2.Fields("D15").value = -1
'Case 16
'Rs2.Fields("D16").value = -1
'Case 17
'Rs2.Fields("D17").value = -1
'Case 18
'Rs2.Fields("D18").value = -1
'Case 19
'Rs2.Fields("D19").value = -1
'Case 20
'Rs2.Fields("D20").value = -1
'Case 21
'Rs2.Fields("D21").value = -1
'Case 22
'Rs2.Fields("D22").value = -1
'Case 23
'Rs2.Fields("D23").value = -1
'Case 24
'Rs2.Fields("D24").value = -1
'Case 25
'Rs2.Fields("D25").value = -1
'Case 26
'Rs2.Fields("D26").value = -1
'Case 27
'Rs2.Fields("D27").value = -1
'Case 28
'Rs2.Fields("D28").value = -1
'Case 29
'Rs2.Fields("D29").value = -1
'Case 30
'Rs2.Fields("D30").value = -1
'Case 31
'Rs2.Fields("D31").value = -1
'End Select
'
'Else
''not vac
'End If
'
'StartDate = DateAdd("d", 1, StartDate)
'
'Next
''###########################################################################################
'
'TempDate = IIf(IsNull(rs1.Fields("RecordDate").value), "", rs1.Fields("RecordDate").value)
'checkRibDate = TempDate
'dayNumber = Day(TempDate)
'numOfAttdays = numOfAttdays + 1
''Õ÷Ê—
'If IsAttend = True Then
'Select Case dayNumber
'Case 1
'
'Rs2.Fields("D1").value = 1
'Case 2
'Rs2.Fields("D2").value = 1
'Case 3
'Rs2.Fields("D3").value = 1
'Case 4
'Rs2.Fields("D4").value = 1
'Case 5
'Rs2.Fields("D5").value = 1
'Case 6
'Rs2.Fields("D6").value = 1
'Case 7
'Rs2.Fields("D7").value = 1
'Case 8
'Rs2.Fields("D8").value = 1
'Case 9
'Rs2.Fields("D9").value = 1
'Case 10
'Rs2.Fields("D10").value = 1
'Case 11
'Rs2.Fields("D11").value = 1
'Case 12
'Rs2.Fields("D12").value = 1
'Case 13
'Rs2.Fields("D13").value = 1
'Case 14
'Rs2.Fields("D14").value = 1
'Case 15
'Rs2.Fields("D15").value = 1
'Case 16
'Rs2.Fields("D16").value = 1
'Case 17
'Rs2.Fields("D17").value = 1
'Case 18
'Rs2.Fields("D18").value = 1
'Case 19
'Rs2.Fields("D19").value = 1
'Case 20
'Rs2.Fields("D20").value = 1
'Case 21
'Rs2.Fields("D21").value = 1
'Case 22
'Rs2.Fields("D22").value = 1
'Case 23
'Rs2.Fields("D23").value = 1
'Case 24
'Rs2.Fields("D24").value = 1
'Case 25
'Rs2.Fields("D25").value = 1
'Case 26
'Rs2.Fields("D26").value = 1
'Case 27
'Rs2.Fields("D27").value = 1
'Case 28
'Rs2.Fields("D28").value = 1
'Case 29
'Rs2.Fields("D29").value = 1
'Case 30
'Rs2.Fields("D30").value = 1
'Case 31
'Rs2.Fields("D31").value = 1
'End Select
'
'Else '€Ì«»
'
'Select Case dayNumber
'Case 1
'
'Rs2.Fields("D1").value = 0
'Case 2
'Rs2.Fields("D2").value = 0
'Case 3
'Rs2.Fields("D3").value = 0
'Case 4
'Rs2.Fields("D4").value = 0
'Case 5
'Rs2.Fields("D5").value = 0
'Case 6
'Rs2.Fields("D6").value = 0
'Case 7
'Rs2.Fields("D7").value = 0
'Case 8
'Rs2.Fields("D8").value = 0
'Case 9
'Rs2.Fields("D9").value = 0
'Case 10
'Rs2.Fields("D10").value = 0
'Case 11
'Rs2.Fields("D11").value = 0
'Case 12
'Rs2.Fields("D12").value = 0
'Case 13
'Rs2.Fields("D13").value = 0
'Case 14
'Rs2.Fields("D14").value = 0
'Case 15
'Rs2.Fields("D15").value = 0
'Case 16
'Rs2.Fields("D16").value = 0
'Case 17
'Rs2.Fields("D17").value = 0
'Case 18
'Rs2.Fields("D18").value = 0
'Case 19
'Rs2.Fields("D19").value = 0
'Case 20
'Rs2.Fields("D20").value = 0
'Case 21
'Rs2.Fields("D21").value = 0
'Case 22
'Rs2.Fields("D22").value = 0
'Case 23
'Rs2.Fields("D23").value = 0
'Case 24
'Rs2.Fields("D24").value = 0
'Case 25
'Rs2.Fields("D25").value = 0
'Case 26
'Rs2.Fields("D26").value = 0
'Case 27
'Rs2.Fields("D27").value = 0
'Case 28
'Rs2.Fields("D28").value = 0
'Case 29
'Rs2.Fields("D29").value = 0
'Case 30
'Rs2.Fields("D30").value = 0
'Case 31
'Rs2.Fields("D31").value = 0
'End Select


'End If
'
'Rs2.Update
'rs1.MoveNext
'
'End If
'
'Next

'End Sub
Public Sub YearMonth()

    Dim i As Integer
    Dim IntDefIndex As Integer

    CmbMonth.Clear

    For i = 1 To 12
        CmbMonth.AddItem MonthName(i)
    Next

    CmbMonth.ListIndex = Month(Date) - 1
    CboYear.Clear

    For i = 2006 To 3000
        CboYear.AddItem i

        If i = year(Date) Then
            IntDefIndex = CboYear.NewIndex
        End If
    Next
    
    CboYear.ListIndex = IntDefIndex

End Sub

Function CheckNoInGroup(Optional RecDate As Date) As Integer
 Dim sql As String
 Dim Rs7 As ADODB.Recordset
 Set Rs7 = New ADODB.Recordset
 sql = "Select * from dbo.TblStuFingerprint Where 1=1"
 If val(groupDBox.BoundText) <> 0 Then
 sql = sql & " and  GroupID = " & val(groupDBox.BoundText) & " "
  sql = sql & " And CursID = " & val(cursBox.BoundText) & " "
  Else
  sql = sql & " AND dbo.TblStuFingerprint.CompID  = " & val(DcbCompany.BoundText)
 End If

 sql = sql & " And GDate = " & SQLDate(RecDate, True) & ""
 Rs7.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
 If Rs7.RecordCount > 0 Then
 CheckNoInGroup = 1
 Else
 CheckNoInGroup = 0
 End If
 End Function
 
Function CheckHolidaies(Optional RecDate As Date) As Integer
 Dim sql As String
 Dim Rs7 As ADODB.Recordset
 Set Rs7 = New ADODB.Recordset
 sql = "Select * from dbo.TblVacationschedule22 where ISVac = 1 and Date =" & SQLDate(RecDate, True) & ""
 Rs7.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
 If Rs7.RecordCount > 0 Then
 CheckHolidaies = 1
 Else
 CheckHolidaies = 0
 End If
 End Function

Private Sub toDateH2_LostFocus()
ToDate2.value = ToGregorianDate(toDateH2.value)
End Sub

Private Sub Txt_OrderNumber_KeyUp(KeyCode As Integer, Shift As Integer)


 If KeyCode = vbKeyF3 Then
       Dim mIndex As Long
       If Rd(6) Then
        mIndex = 1
        FrmProjectSearch.Label11.Caption = 5
         FrmProjectSearch.Indx2 = 0
       ElseIf Rd(5) Then
        mIndex = 2
        FrmProjectSearch.Indx2 = 0
        FrmProjectSearch.Caption = "»ÕÀ »Õ—ﬂ… «·ÿ·»« "
       End If
       FrmProjectSearch.C1Tab1 = mIndex
       
        FrmProjectSearch.show vbModal
   End If


End Sub

Private Sub Txt_OrderNumber2_KeyUp(KeyCode As Integer, Shift As Integer)

 If KeyCode = vbKeyF3 Then
       Dim mIndex As Long
       If Rd(6) Then
        mIndex = 1
        FrmProjectSearch.Label11.Caption = 5
        FrmProjectSearch.Indx2 = 0
       ElseIf Rd(5) Then
        mIndex = 2
        FrmProjectSearch.Indx2 = 2
        FrmProjectSearch.Caption = "»ÕÀ »Õ—ﬂ… «·ÿ·»« "
       End If
       FrmProjectSearch.C1Tab1 = mIndex
       
        FrmProjectSearch.show vbModal
   End If

End Sub

Private Sub txtEmpCode_KeyPress(KeyAscii As Integer)
   Dim EmpID As Integer

       If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode TxtEmpCode.Text, EmpID
        DcbEmployee2.BoundText = EmpID
    End If
    
End Sub

Private Sub TxtEmpCode_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF3 Then
     FrmEmployeeSearch.lbltype = 47
        FrmEmployeeSearch.show
    End If

End Sub

Private Sub TxtSearchCode_KeyPress(Index As Integer, KeyAscii As Integer)
Dim TContractCustID As Double

If TxtSearchCode(Index).Text = "" Then Exit Sub

If KeyAscii = vbKeyReturn Then
Get_TradingContractinfo TxtSearchCode(Index).Text, TContractCustID, 0

DcCustmer(Index).BoundText = TContractCustID
End If
End Sub

Private Sub TxtSearchCode_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
Dim TContractCustID As Double
    
 FrmProjectSearch.C1Tab1 = 4
 FrmProjectSearch.Label11.Caption = IIf(Index = 0, 3, 4)
 FrmProjectSearch.Caption = "»ÕÀ «·« ›«ﬁÌ«  "
 FrmProjectSearch.show vbModal
 
 Get_TradingContractinfo val(TxtSearchCode(Index).Text), TContractCustID, 0

DcCustmer(Index).BoundText = TContractCustID
    End If

End Sub

Private Sub TxtSudCode_KeyPress(KeyAscii As Integer)
 Dim EmpID As Integer
 Dim UQama As String
    If KeyAscii = vbKeyReturn Then
        GetStudentCode EmpID, TxtSudCode.Text, 1, UQama
        DcbStudent.BoundText = EmpID
    End If
End Sub

Private Sub TxtSudCode2_KeyPress(KeyAscii As Integer)
 Dim EmpID As Integer
 Dim UQama As String
    If KeyAscii = vbKeyReturn Then
        GetStudentCode EmpID, TxtSudCode2.Text, 1, UQama
        DcbStudent2.BoundText = EmpID
    End If
End Sub

