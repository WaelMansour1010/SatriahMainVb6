VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form dean 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11370
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   19605
   Icon            =   "dean.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   11370
   ScaleWidth      =   19605
   WindowState     =   2  'Maximized
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin C1SizerLibCtl.C1Tab TabMain 
      Height          =   11370
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   19605
      _cx             =   34581
      _cy             =   20055
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
      Caption         =   ""
      Align           =   5
      CurrTab         =   11
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
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   11025
         Index           =   1
         Left            =   30
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   30
         Width           =   19530
         _cx             =   34449
         _cy             =   19447
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
         Begin VB.Frame Frm2 
            BackColor       =   &H00E2E9E9&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   2445
            Left            =   3900
            RightToLeft     =   -1  'True
            TabIndex        =   63
            Top             =   2325
            Width           =   14010
            Begin VB.ComboBox CmbType 
               BackColor       =   &H80000018&
               Height          =   315
               ItemData        =   "dean.frx":058A
               Left            =   2280
               List            =   "dean.frx":059A
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   75
               Top             =   2670
               Visible         =   0   'False
               Width           =   1005
            End
            Begin VB.TextBox TxtSerial 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Enabled         =   0   'False
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
               Left            =   13785
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   74
               Top             =   990
               Visible         =   0   'False
               Width           =   1410
            End
            Begin VB.TextBox TxtPassWord 
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
               IMEMode         =   3  'DISABLE
               Left            =   9720
               MaxLength       =   50
               PasswordChar    =   "#"
               RightToLeft     =   -1  'True
               TabIndex        =   73
               Tag             =   "⁄›Ê« Ì—ÃÏ   ‰”»… «·Œ’„"
               Top             =   750
               Width           =   2370
            End
            Begin VB.TextBox TXTCode 
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
               Left            =   9720
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   72
               Top             =   0
               Width           =   2370
            End
            Begin VB.ComboBox CboPriv 
               BackColor       =   &H00C0FFFF&
               Height          =   315
               ItemData        =   "dean.frx":05B3
               Left            =   360
               List            =   "dean.frx":05BD
               RightToLeft     =   -1  'True
               TabIndex        =   71
               Top             =   0
               Width           =   2775
            End
            Begin VB.TextBox XPTxtPassConfirm 
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
               IMEMode         =   3  'DISABLE
               Left            =   5160
               MaxLength       =   50
               PasswordChar    =   "#"
               RightToLeft     =   -1  'True
               TabIndex        =   70
               Tag             =   "⁄›Ê« Ì—ÃÏ   ‰”»… «·Œ’„"
               Top             =   750
               Width           =   3090
            End
            Begin VB.TextBox XPTxtUserName 
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
               Left            =   9720
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   69
               Top             =   360
               Width           =   2370
            End
            Begin VB.CheckBox chkNextLogin 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " €ÌÌ— ﬂ·„… «·„—Ê— ⁄‰œ «·œŒÊ·"
               Height          =   195
               Left            =   2400
               RightToLeft     =   -1  'True
               TabIndex        =   68
               Top             =   720
               Width           =   2355
            End
            Begin VB.TextBox TxtSearchCode 
               Alignment       =   2  'Center
               Height          =   345
               Left            =   7560
               TabIndex        =   67
               Top             =   1080
               Width           =   690
            End
            Begin VB.TextBox TxtSearchCode1 
               Alignment       =   2  'Center
               Height          =   345
               Left            =   7560
               TabIndex        =   66
               Top             =   1440
               Width           =   690
            End
            Begin VB.CheckBox isDeactivatedchk 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "≈Ìﬁ«› «·„” Œœ„"
               Height          =   195
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   65
               Top             =   720
               Width           =   2115
            End
            Begin VB.CheckBox chkHidLowering 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«ŸÂ«— «·«⁄„œ… ﬂ«„·… ›Ì   ‰»ÌÂ«  «·«‰ «Ã"
               Height          =   195
               Left            =   990
               RightToLeft     =   -1  'True
               TabIndex        =   64
               Top             =   1830
               Width           =   3675
            End
            Begin MSComDlg.CommonDialog cdg 
               Left            =   4800
               Top             =   240
               _ExtentX        =   847
               _ExtentY        =   847
               _Version        =   393216
            End
            Begin MSDataListLib.DataCombo DCEmP 
               Height          =   315
               Left            =   5160
               TabIndex        =   76
               Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· «Œ Ì«— «·„‰œÊ»"
               Top             =   0
               Width           =   3090
               _ExtentX        =   5450
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DCSalesRepGroups 
               Height          =   315
               Left            =   120
               TabIndex        =   77
               Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·„Ã„Ê⁄Â"
               Top             =   -360
               Width           =   3075
               _ExtentX        =   5424
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcBranches 
               Height          =   315
               Left            =   5160
               TabIndex        =   78
               Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·›—⁄"
               Top             =   360
               Width           =   3090
               _ExtentX        =   5450
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DCJob 
               Height          =   315
               Left            =   120
               TabIndex        =   79
               Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«·  «·ÊŸÌ€…"
               Top             =   -600
               Visible         =   0   'False
               Width           =   3075
               _ExtentX        =   5424
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DCStore 
               Height          =   315
               Left            =   9720
               TabIndex        =   80
               Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·›—⁄"
               Top             =   1080
               Width           =   2370
               _ExtentX        =   4180
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DCBoxes 
               Height          =   315
               Left            =   240
               TabIndex        =   81
               Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·›—⁄"
               Top             =   1080
               Width           =   2595
               _ExtentX        =   4577
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo Dbanks 
               Height          =   315
               Left            =   360
               TabIndex        =   82
               Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·›—⁄"
               Top             =   360
               Width           =   2775
               _ExtentX        =   4895
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DBCboClientName 
               Height          =   315
               Left            =   5160
               TabIndex        =   83
               Top             =   1080
               Width           =   2325
               _ExtentX        =   4101
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "6"
               BoundColumn     =   ""
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DCStore1 
               Height          =   315
               Left            =   9720
               TabIndex        =   84
               Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·›—⁄"
               Top             =   1440
               Width           =   2370
               _ExtentX        =   4180
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DBCboClientName1 
               Height          =   315
               Left            =   5160
               TabIndex        =   85
               Top             =   1440
               Width           =   2325
               _ExtentX        =   4101
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "6"
               BoundColumn     =   ""
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DCBoxes1 
               Height          =   315
               Left            =   240
               TabIndex        =   86
               Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·›—⁄"
               Top             =   1440
               Width           =   2595
               _ExtentX        =   4577
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DCStore2 
               Height          =   315
               Left            =   9720
               TabIndex        =   87
               Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·›—⁄"
               Top             =   1770
               Width           =   2370
               _ExtentX        =   4180
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DCStore3 
               Height          =   315
               Left            =   5160
               TabIndex        =   88
               Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·›—⁄"
               Top             =   1770
               Width           =   3090
               _ExtentX        =   5450
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·„ÊŸ›"
               Height          =   285
               Index           =   1
               Left            =   8760
               RightToLeft     =   -1  'True
               TabIndex        =   104
               Top             =   30
               Width           =   930
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬂÊœ «·„” Œœ„"
               Height          =   195
               Index           =   3
               Left            =   12165
               RightToLeft     =   -1  'True
               TabIndex        =   103
               Top             =   30
               Width           =   1470
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬂ·„… «·”—"
               Height          =   195
               Index           =   0
               Left            =   12165
               RightToLeft     =   -1  'True
               TabIndex        =   102
               Top             =   720
               Width           =   1470
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·Œ“Ì‰… «·«› —«÷Ì… ··Ì⁄"
               Height          =   285
               Index           =   4
               Left            =   2970
               RightToLeft     =   -1  'True
               TabIndex        =   101
               Top             =   1080
               Width           =   1650
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·›—⁄"
               Height          =   285
               Index           =   5
               Left            =   8490
               RightToLeft     =   -1  'True
               TabIndex        =   100
               Top             =   360
               Width           =   930
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·»‰ﬂ «·«› —«÷Ì"
               Height          =   285
               Index           =   6
               Left            =   3330
               RightToLeft     =   -1  'True
               TabIndex        =   99
               Top             =   360
               Width           =   1290
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„Œ“‰ «·»Ì⁄ «·«› —«÷Ì"
               Height          =   195
               Index           =   7
               Left            =   12165
               RightToLeft     =   -1  'True
               TabIndex        =   98
               Top             =   960
               Width           =   1470
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "’·«ÕÌ« "
               Height          =   330
               Index           =   4
               Left            =   3840
               RightToLeft     =   -1  'True
               TabIndex        =   97
               Top             =   30
               Width           =   735
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " √ﬂÌœ ﬂ·„… «·”—"
               Height          =   285
               Index           =   8
               Left            =   8400
               RightToLeft     =   -1  'True
               TabIndex        =   96
               Top             =   720
               Width           =   1170
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·„” Œœ„"
               Height          =   195
               Index           =   9
               Left            =   12165
               RightToLeft     =   -1  'True
               TabIndex        =   95
               Top             =   360
               Width           =   1470
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·⁄„Ì· «·«› —«÷Ì"
               Height          =   285
               Index           =   10
               Left            =   8400
               RightToLeft     =   -1  'True
               TabIndex        =   94
               Top             =   1080
               Width           =   1290
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„Ê—œ «·«› —«÷Ì"
               Height          =   285
               Index           =   11
               Left            =   8400
               RightToLeft     =   -1  'True
               TabIndex        =   93
               Top             =   1440
               Width           =   1290
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„Œ“‰ «·‘—«¡ «·«› —«÷Ì"
               Height          =   195
               Index           =   12
               Left            =   12165
               RightToLeft     =   -1  'True
               TabIndex        =   92
               Top             =   1440
               Width           =   1470
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·Œ“Ì‰… «·«› —«÷Ì… ··‘—«¡"
               Height          =   285
               Index           =   14
               Left            =   2850
               RightToLeft     =   -1  'True
               TabIndex        =   91
               Top             =   1440
               Width           =   1770
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„Œ“‰ ’—› «·„Ê«œ «·Œ«„"
               Height          =   195
               Index           =   15
               Left            =   12165
               RightToLeft     =   -1  'True
               TabIndex        =   90
               Top             =   1770
               Width           =   1530
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„Œ“‰ «” ·«„ «·„Ê«œ «·Œ«„"
               Height          =   195
               Index           =   16
               Left            =   8325
               RightToLeft     =   -1  'True
               TabIndex        =   89
               Top             =   1770
               Width           =   1350
            End
         End
         Begin VB.Frame FraHeader 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   555
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   52
            Top             =   0
            Width           =   18030
            Begin VB.Frame Frmo2 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   375
               Left            =   540
               RightToLeft     =   -1  'True
               TabIndex        =   55
               Top             =   450
               Visible         =   0   'False
               Width           =   3105
               Begin MSDataListLib.DataCombo DCUser 
                  CausesValidation=   0   'False
                  Height          =   315
                  Left            =   -255
                  TabIndex        =   56
                  Tag             =   "„‰ ›÷·ﬂ √œŒ· —ﬁ„ «·ﬁ÷Ì…"
                  Top             =   15
                  Width           =   2340
                  _ExtentX        =   4128
                  _ExtentY        =   556
                  _Version        =   393216
                  Enabled         =   0   'False
                  BackColor       =   -2147483624
                  Text            =   ""
                  RightToLeft     =   -1  'True
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E0E0E0&
                  Caption         =   "«·„” Œœ„"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   12
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   270
                  Index           =   13
                  Left            =   2160
                  RightToLeft     =   -1  'True
                  TabIndex        =   57
                  Top             =   45
                  Width           =   855
               End
            End
            Begin VB.TextBox TxtModFlg 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H0000FF00&
               Enabled         =   0   'False
               Height          =   285
               Left            =   2580
               RightToLeft     =   -1  'True
               TabIndex        =   54
               Text            =   "modflag"
               Top             =   90
               Visible         =   0   'False
               Width           =   465
            End
            Begin VB.TextBox TxtVac_ID 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H000000FF&
               Height          =   240
               Left            =   3030
               RightToLeft     =   -1  'True
               TabIndex        =   53
               Top             =   510
               Visible         =   0   'False
               Width           =   945
            End
            Begin MSComctlLib.ImageList GrdImageList 
               Left            =   3120
               Top             =   0
               _ExtentX        =   1005
               _ExtentY        =   1005
               BackColor       =   -2147483643
               ImageWidth      =   16
               ImageHeight     =   16
               MaskColor       =   12632256
               _Version        =   393216
               BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
                  NumListImages   =   8
                  BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "dean.frx":05D9
                     Key             =   "CompanyName"
                  EndProperty
                  BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "dean.frx":0973
                     Key             =   "Ser"
                  EndProperty
                  BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "dean.frx":0D0D
                     Key             =   "Vac_Name"
                  EndProperty
                  BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "dean.frx":10A7
                     Key             =   "ShareCount"
                  EndProperty
                  BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "dean.frx":1441
                     Key             =   "Dis_Count"
                  EndProperty
                  BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "dean.frx":17DB
                     Key             =   "Bouns"
                  EndProperty
                  BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "dean.frx":1B75
                     Key             =   "SharesValue"
                  EndProperty
                  BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "dean.frx":210F
                     Key             =   "BuyValue"
                  EndProperty
               EndProperty
            End
            Begin ImpulseButton.ISButton btnLast 
               Height          =   315
               Left            =   90
               TabIndex        =   58
               Top             =   30
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   556
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   ""
               BackColor       =   14871017
               FontSize        =   12
               FontName        =   "Arial"
               FontBold        =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ButtonImage     =   "dean.frx":24A9
               ColorButton     =   14871017
               AcclimateGrayTones=   -1  'True
               DrawFocusRectangle=   0   'False
               DisabledImageExtraction=   0
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnNext 
               Height          =   315
               Left            =   555
               TabIndex        =   59
               Top             =   30
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   556
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   ""
               BackColor       =   14871017
               FontSize        =   12
               FontName        =   "Arial"
               FontBold        =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ButtonImage     =   "dean.frx":2843
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnPrevious 
               Height          =   315
               Left            =   1155
               TabIndex        =   60
               Top             =   30
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   556
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   ""
               BackColor       =   14871017
               FontSize        =   12
               FontName        =   "Arial"
               FontBold        =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ButtonImage     =   "dean.frx":2BDD
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnFirst 
               Height          =   315
               Left            =   1620
               TabIndex        =   61
               Top             =   30
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   556
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   ""
               BackColor       =   14871017
               FontSize        =   12
               FontName        =   "Arial"
               FontBold        =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ButtonImage     =   "dean.frx":2F77
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "»Ì«‰«  «·„” Œœ„Ì‰"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   14.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   375
               Index           =   2
               Left            =   10935
               RightToLeft     =   -1  'True
               TabIndex        =   62
               Top             =   90
               Width           =   2790
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "«·Œ“‰ «·„— »ÿ »Â«"
            Height          =   2550
            Left            =   7035
            RightToLeft     =   -1  'True
            TabIndex        =   41
            Top             =   7515
            Width           =   7110
            Begin VB.ListBox ListBoxesSelected 
               BackColor       =   &H0080FFFF&
               Height          =   1815
               ItemData        =   "dean.frx":3311
               Left            =   240
               List            =   "dean.frx":3318
               RightToLeft     =   -1  'True
               TabIndex        =   45
               Top             =   420
               Width           =   3015
            End
            Begin VB.CommandButton Command2 
               Caption         =   "Command1"
               Height          =   435
               Left            =   360
               RightToLeft     =   -1  'True
               TabIndex        =   44
               Top             =   3000
               Visible         =   0   'False
               Width           =   495
            End
            Begin VB.ListBox ListBoxesAll 
               Height          =   1815
               ItemData        =   "dean.frx":332F
               Left            =   3630
               List            =   "dean.frx":3336
               MultiSelect     =   1  'Simple
               RightToLeft     =   -1  'True
               TabIndex        =   43
               Top             =   420
               Width           =   3135
            End
            Begin VB.CommandButton cmdReloadList 
               Caption         =   "«·€«¡ «·„Õœœ"
               Height          =   225
               Index           =   0
               Left            =   2610
               RightToLeft     =   -1  'True
               TabIndex        =   42
               Top             =   120
               Width           =   1980
            End
            Begin VB.Label Label13 
               Alignment       =   2  'Center
               Caption         =   "«·Œ“‰ «·„Õœœ…"
               Height          =   255
               Left            =   1080
               RightToLeft     =   -1  'True
               TabIndex        =   51
               Top             =   120
               Width           =   1335
            End
            Begin VB.Label Label14 
               Alignment       =   2  'Center
               Caption         =   "ﬂ· «·Œ“‰"
               Height          =   255
               Left            =   4680
               RightToLeft     =   -1  'True
               TabIndex        =   50
               Top             =   120
               Width           =   1095
            End
            Begin VB.Label Label15 
               Alignment       =   2  'Center
               Caption         =   "<"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   255
               Left            =   3210
               RightToLeft     =   -1  'True
               TabIndex        =   49
               Top             =   1380
               Width           =   495
            End
            Begin VB.Label Label16 
               Alignment       =   2  'Center
               Caption         =   "<<"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   255
               Left            =   3210
               RightToLeft     =   -1  'True
               TabIndex        =   48
               Top             =   1140
               Width           =   495
            End
            Begin VB.Label Label17 
               Alignment       =   2  'Center
               Caption         =   ">>"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Left            =   3210
               RightToLeft     =   -1  'True
               TabIndex        =   47
               Top             =   780
               Width           =   495
            End
            Begin VB.Label Label18 
               Alignment       =   2  'Center
               Caption         =   ">"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Left            =   3210
               RightToLeft     =   -1  'True
               TabIndex        =   46
               Top             =   540
               Width           =   495
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "«·Õ”«»«  «·„— »ÿ »Â«"
            Height          =   2085
            Left            =   30
            RightToLeft     =   -1  'True
            TabIndex        =   31
            Top             =   7425
            Width           =   7020
            Begin VB.ListBox ListAllAccount 
               Height          =   1425
               ItemData        =   "dean.frx":3348
               Left            =   4440
               List            =   "dean.frx":334F
               MultiSelect     =   1  'Simple
               RightToLeft     =   -1  'True
               TabIndex        =   34
               Top             =   390
               Width           =   2445
            End
            Begin VB.ListBox ListAccountSelect 
               BackColor       =   &H0080FFFF&
               Height          =   1425
               ItemData        =   "dean.frx":3361
               Left            =   60
               List            =   "dean.frx":3368
               MultiSelect     =   1  'Simple
               RightToLeft     =   -1  'True
               TabIndex        =   33
               Top             =   420
               Width           =   3315
            End
            Begin VB.CommandButton cmdReloadList 
               Caption         =   "«·€«¡ «·„Õœœ"
               Height          =   225
               Index           =   2
               Left            =   2460
               RightToLeft     =   -1  'True
               TabIndex        =   32
               Top             =   150
               Width           =   1980
            End
            Begin VB.Label Label19 
               Alignment       =   2  'Center
               Caption         =   "«·Õ”«»«  «·„Õœœ…"
               Height          =   255
               Left            =   360
               RightToLeft     =   -1  'True
               TabIndex        =   40
               Top             =   120
               Width           =   2295
            End
            Begin VB.Label Label20 
               Alignment       =   2  'Center
               Caption         =   "ﬂ· «·Õ”«»« "
               Height          =   255
               Left            =   4200
               RightToLeft     =   -1  'True
               TabIndex        =   39
               Top             =   120
               Width           =   1815
            End
            Begin VB.Label Label21 
               Alignment       =   2  'Center
               Caption         =   ">"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Left            =   3540
               RightToLeft     =   -1  'True
               TabIndex        =   38
               Top             =   420
               Width           =   495
            End
            Begin VB.Label Label23 
               Alignment       =   2  'Center
               Caption         =   ">>"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Left            =   3540
               RightToLeft     =   -1  'True
               TabIndex        =   37
               Top             =   660
               Width           =   495
            End
            Begin VB.Label Label24 
               Alignment       =   2  'Center
               Caption         =   "<<"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   255
               Left            =   3540
               RightToLeft     =   -1  'True
               TabIndex        =   36
               Top             =   1020
               Width           =   495
            End
            Begin VB.Label Label25 
               Alignment       =   2  'Center
               Caption         =   "<"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   255
               Left            =   3540
               RightToLeft     =   -1  'True
               TabIndex        =   35
               Top             =   1260
               Width           =   495
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "«·ŒÿÊÿ «·„— »ÿ »Â«"
            Height          =   2055
            Left            =   14310
            RightToLeft     =   -1  'True
            TabIndex        =   22
            Top             =   7680
            Width           =   5040
            Begin VB.ListBox ListProductLineSelected 
               BackColor       =   &H0080FFFF&
               Height          =   1230
               ItemData        =   "dean.frx":337F
               Left            =   150
               List            =   "dean.frx":3386
               RightToLeft     =   -1  'True
               TabIndex        =   24
               Top             =   720
               Width           =   1515
            End
            Begin VB.ListBox ListProductLineAll 
               Height          =   1230
               ItemData        =   "dean.frx":339D
               Left            =   3270
               List            =   "dean.frx":33A4
               RightToLeft     =   -1  'True
               TabIndex        =   23
               Top             =   570
               Width           =   1545
            End
            Begin VB.Label Label31 
               Alignment       =   2  'Center
               Caption         =   "<"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   255
               Left            =   2400
               RightToLeft     =   -1  'True
               TabIndex        =   30
               Top             =   1530
               Width           =   495
            End
            Begin VB.Label Label30 
               Alignment       =   2  'Center
               Caption         =   "<<"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   255
               Left            =   2400
               RightToLeft     =   -1  'True
               TabIndex        =   29
               Top             =   1290
               Width           =   495
            End
            Begin VB.Label Label29 
               Alignment       =   2  'Center
               Caption         =   ">>"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Left            =   2400
               RightToLeft     =   -1  'True
               TabIndex        =   28
               Top             =   930
               Width           =   495
            End
            Begin VB.Label Label28 
               Alignment       =   2  'Center
               Caption         =   ">"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Left            =   2400
               RightToLeft     =   -1  'True
               TabIndex        =   27
               Top             =   690
               Width           =   495
            End
            Begin VB.Label Label27 
               Alignment       =   2  'Center
               Caption         =   "ﬂ· ŒÿÊÿ «·«‰ «Ã"
               Height          =   255
               Left            =   2310
               RightToLeft     =   -1  'True
               TabIndex        =   26
               Top             =   240
               Width           =   1335
            End
            Begin VB.Label Label26 
               Alignment       =   2  'Center
               Caption         =   "«·ŒÿÊÿ «·„Õœœ…"
               Height          =   255
               Left            =   300
               RightToLeft     =   -1  'True
               TabIndex        =   25
               Top             =   240
               Width           =   1335
            End
         End
         Begin VB.Frame Frame10 
            Caption         =   "«·„Œ«“‰ «·„— »ÿ »Â«"
            Height          =   1740
            Left            =   30
            RightToLeft     =   -1  'True
            TabIndex        =   12
            Top             =   5640
            Width           =   8580
            Begin VB.ListBox ListStoreall 
               Height          =   1035
               ItemData        =   "dean.frx":33B6
               Left            =   4650
               List            =   "dean.frx":33BD
               MultiSelect     =   1  'Simple
               RightToLeft     =   -1  'True
               TabIndex        =   15
               Top             =   510
               Width           =   3825
            End
            Begin VB.ListBox ListStoreSelected 
               BackColor       =   &H0080FFFF&
               Height          =   1035
               ItemData        =   "dean.frx":33CF
               Left            =   30
               List            =   "dean.frx":33D6
               MultiSelect     =   1  'Simple
               RightToLeft     =   -1  'True
               TabIndex        =   14
               Top             =   540
               Width           =   3765
            End
            Begin VB.CommandButton cmdReloadList 
               Caption         =   "«·€«¡ «·„Õœœ"
               Height          =   225
               Index           =   1
               Left            =   3510
               RightToLeft     =   -1  'True
               TabIndex        =   13
               Top             =   210
               Width           =   1980
            End
            Begin VB.Label Label10 
               Alignment       =   2  'Center
               Caption         =   "«·„Œ«“‰ «·„Õœœ…"
               Height          =   255
               Left            =   540
               RightToLeft     =   -1  'True
               TabIndex        =   21
               Top             =   180
               Width           =   1335
            End
            Begin VB.Label Label9 
               Alignment       =   2  'Center
               Caption         =   "ﬂ· «·„Œ«“‰"
               Height          =   255
               Left            =   6510
               RightToLeft     =   -1  'True
               TabIndex        =   20
               Top             =   240
               Width           =   1335
            End
            Begin VB.Label LblSelect 
               Alignment       =   2  'Center
               Caption         =   ">"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Left            =   4080
               RightToLeft     =   -1  'True
               TabIndex        =   19
               Top             =   360
               Width           =   495
            End
            Begin VB.Label Label22 
               Alignment       =   2  'Center
               Caption         =   ">>"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Left            =   4080
               RightToLeft     =   -1  'True
               TabIndex        =   18
               Top             =   600
               Width           =   495
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               Caption         =   "<<"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   255
               Left            =   4080
               RightToLeft     =   -1  'True
               TabIndex        =   17
               Top             =   960
               Width           =   495
            End
            Begin VB.Label Label4 
               Alignment       =   2  'Center
               Caption         =   "<"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   255
               Left            =   4080
               RightToLeft     =   -1  'True
               TabIndex        =   16
               Top             =   1200
               Width           =   495
            End
         End
         Begin VB.Frame Frame11 
            Caption         =   "«·›—Ê⁄ «·„— »ÿ »Â«"
            Height          =   1800
            Left            =   9420
            RightToLeft     =   -1  'True
            TabIndex        =   2
            Top             =   5610
            Width           =   10005
            Begin VB.ListBox ListGroupSelected 
               BackColor       =   &H0080FFFF&
               Height          =   1425
               ItemData        =   "dean.frx":33ED
               Left            =   240
               List            =   "dean.frx":33F4
               RightToLeft     =   -1  'True
               TabIndex        =   5
               Top             =   240
               Width           =   4545
            End
            Begin VB.ListBox ListGroupAll 
               Height          =   1425
               ItemData        =   "dean.frx":340B
               Left            =   5520
               List            =   "dean.frx":3412
               RightToLeft     =   -1  'True
               TabIndex        =   4
               Top             =   240
               Width           =   3495
            End
            Begin VB.CommandButton Command1 
               Caption         =   "Command1"
               Height          =   435
               Left            =   360
               RightToLeft     =   -1  'True
               TabIndex        =   3
               Top             =   3000
               Visible         =   0   'False
               Width           =   495
            End
            Begin VB.Label Label12 
               Alignment       =   2  'Center
               Caption         =   "«·›—Ê⁄ «·„Õœœ…"
               Height          =   255
               Left            =   2550
               RightToLeft     =   -1  'True
               TabIndex        =   11
               Top             =   0
               Width           =   1335
            End
            Begin VB.Label Label11 
               Alignment       =   2  'Center
               Caption         =   "ﬂ· «·›—Ê⁄"
               Height          =   255
               Left            =   6270
               RightToLeft     =   -1  'True
               TabIndex        =   10
               Top             =   0
               Width           =   1335
            End
            Begin VB.Label Label5 
               Alignment       =   2  'Center
               Caption         =   "<"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   375
               Left            =   4920
               RightToLeft     =   -1  'True
               TabIndex        =   9
               Top             =   1140
               Width           =   495
            End
            Begin VB.Label Label6 
               Alignment       =   2  'Center
               Caption         =   "<<"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   255
               Left            =   4920
               RightToLeft     =   -1  'True
               TabIndex        =   8
               Top             =   840
               Width           =   495
            End
            Begin VB.Label Label7 
               Alignment       =   2  'Center
               Caption         =   ">>"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Left            =   4920
               RightToLeft     =   -1  'True
               TabIndex        =   7
               Top             =   480
               Width           =   495
            End
            Begin VB.Label Label8 
               Alignment       =   2  'Center
               Caption         =   ">"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Left            =   4920
               RightToLeft     =   -1  'True
               TabIndex        =   6
               Top             =   240
               Width           =   495
            End
         End
         Begin C1SizerLibCtl.C1Elastic EltCont 
            Height          =   885
            Left            =   7005
            TabIndex        =   105
            TabStop         =   0   'False
            Top             =   10140
            Width           =   7125
            _cx             =   12568
            _cy             =   1561
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
            Begin ImpulseButton.ISButton btnNew 
               Height          =   330
               Left            =   5295
               TabIndex        =   106
               Top             =   435
               Width           =   750
               _ExtentX        =   1323
               _ExtentY        =   582
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
               ButtonImage     =   "dean.frx":3424
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnSave 
               Height          =   330
               Left            =   3510
               TabIndex        =   107
               Top             =   435
               Width           =   750
               _ExtentX        =   1323
               _ExtentY        =   582
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
               ButtonImage     =   "dean.frx":37BE
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnModify 
               Height          =   330
               Left            =   4395
               TabIndex        =   108
               Top             =   435
               Width           =   750
               _ExtentX        =   1323
               _ExtentY        =   582
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
               ButtonImage     =   "dean.frx":3B58
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton BtnUndo 
               Height          =   330
               Left            =   2745
               TabIndex        =   109
               Top             =   435
               Width           =   750
               _ExtentX        =   1323
               _ExtentY        =   582
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
               ButtonImage     =   "dean.frx":3EF2
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnDelete 
               Height          =   330
               Left            =   1020
               TabIndex        =   110
               Top             =   435
               Width           =   750
               _ExtentX        =   1323
               _ExtentY        =   582
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
               ButtonImage     =   "dean.frx":428C
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnQuery 
               Height          =   330
               Left            =   1800
               TabIndex        =   111
               TabStop         =   0   'False
               ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
               Top             =   450
               Width           =   840
               _ExtentX        =   1482
               _ExtentY        =   582
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
               ButtonImage     =   "dean.frx":4826
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton BtnUpdate 
               Height          =   330
               Left            =   5010
               TabIndex        =   112
               TabStop         =   0   'False
               ToolTipText     =   " ÕœÌÀ ﬁ«⁄œ… «·»Ì«‰« "
               Top             =   60
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
               ButtonImage     =   "dean.frx":4BC0
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton BtnPrint 
               Height          =   285
               Left            =   3960
               TabIndex        =   113
               TabStop         =   0   'False
               Top             =   90
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
               ButtonImage     =   "dean.frx":4F5A
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnCancel 
               Height          =   330
               Left            =   225
               TabIndex        =   114
               Top             =   435
               Width           =   750
               _ExtentX        =   1323
               _ExtentY        =   582
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
               ButtonImage     =   "dean.frx":52F4
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin VB.Label LabCountRec 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Height          =   210
               Left            =   720
               RightToLeft     =   -1  'True
               TabIndex        =   118
               Top             =   45
               Width           =   540
            End
            Begin VB.Label LabCurrRec 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Height          =   210
               Left            =   2190
               RightToLeft     =   -1  'True
               TabIndex        =   117
               Top             =   90
               Width           =   675
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄œœ «·”Ã·« :"
               Height          =   210
               Index           =   1
               Left            =   1290
               RightToLeft     =   -1  'True
               TabIndex        =   116
               Top             =   45
               Width           =   975
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·”Ã· «·Õ«·Ì:"
               Height          =   210
               Index           =   0
               Left            =   2895
               RightToLeft     =   -1  'True
               TabIndex        =   115
               Top             =   75
               Width           =   975
            End
         End
         Begin VSFlex8UCtl.VSFlexGrid Grid 
            Height          =   1665
            Left            =   30
            TabIndex        =   119
            Top             =   585
            Width           =   18000
            _cx             =   31750
            _cy             =   2937
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
            AllowBigSelection=   0   'False
            AllowUserResizing=   0
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   50
            Cols            =   13
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"dean.frx":568E
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
            Editable        =   0
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
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   840
            Index           =   6
            Left            =   240
            TabIndex        =   120
            TabStop         =   0   'False
            Top             =   4680
            Width           =   14025
            _cx             =   24739
            _cy             =   1482
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
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
            ForeColor       =   192
            FloodColor      =   6553600
            ForeColorDisabled=   -2147483631
            Caption         =   "«· ÊﬁÌ⁄ «·«·ﬂ —Ê‰Ì"
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
            Begin Dynamic_Byte.NewViewBox ImgPic 
               Height          =   270
               Left            =   120
               TabIndex        =   121
               ToolTipText     =   "≈÷€ÿ ⁄·Ï «·’Ê—… „— Ì‰ ·· ﬂ»Ì—"
               Top             =   210
               Width           =   5070
               _ExtentX        =   8943
               _ExtentY        =   476
            End
            Begin ImpulseButton.ISButton CmdPic 
               Height          =   240
               Index           =   0
               Left            =   10680
               TabIndex        =   122
               Top             =   210
               Width           =   1500
               _ExtentX        =   2646
               _ExtentY        =   423
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   "≈÷«›… ’Ê—…"
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
               ButtonImage     =   "dean.frx":5869
               ColorButton     =   14871017
               Alignment       =   1
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton CmdPic 
               Height          =   225
               Index           =   1
               Left            =   9000
               TabIndex        =   123
               Top             =   210
               Width           =   1380
               _ExtentX        =   2434
               _ExtentY        =   397
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   "Õ–› «·’Ê—…"
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
               ButtonImage     =   "dean.frx":5C03
               ColorButton     =   14871017
               Alignment       =   1
               DrawFocusRectangle=   0   'False
            End
         End
         Begin VSFlex8Ctl.VSFlexGrid FG 
            Height          =   1530
            Left            =   0
            TabIndex        =   124
            Top             =   9555
            Width           =   6930
            _cx             =   12224
            _cy             =   2699
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
            BackColorFixed  =   -2147483633
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
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   1
            Cols            =   4
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"dean.frx":5F9D
            ScrollTrack     =   0   'False
            ScrollBars      =   2
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
      End
   End
End
Attribute VB_Name = "dean"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsSavRec As ADODB.Recordset
Dim BKGrndPic As ClsBackGroundPic
Dim RecId As String
Dim II As Long
Public mIndex As Long
Dim cSearch  As clsDCboSearch
Dim RsTemp As New ADODB.Recordset
Private Sub BtnCancel_Click()
    Unload Me
End Sub
Function CuurentLogdata(Optional Currentmode As String)
     LogTextA = "    ‘«‘… " & ScreenNameArabic & CHR(13) & "ﬂÊœ  «·„” Œœ„ " & TXTCode.Text & CHR(13) & "   «”„ «·„” Œœ„  " & XPTxtUserName.Text & CHR(13) & "   «·›—⁄ " & DcBranches.Text
        LogTexte = "  Screen  " & ScreenNameEnglish & CHR(13) & " User Code " & TXTCode.Text & CHR(13) & "   User Name  " & XPTxtUserName.Text & CHR(13) & "   Branch " & DcBranches
       If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), , Date, Time, LogTextA, LogTexte, Me.Name, Me.TxtModFlg, , , TXTCode, TXTCode
    Else
        AddToLogFile CInt(user_id), , Date, Time, LogTextA, LogTexte, Me.Name, "D", , , TXTCode, TXTCode
    End If
End Function
Private Sub btnDelete_Click()

    Dim MSGType As Integer
    Dim BlnRecExist As Boolean
    Dim StrMSG  As String
    Dim Msg As String
    
    On Error GoTo ErrTrap

    If DoPremis(Do_Delete, Me.Name, True) = False Then
        Exit Sub
    End If

    If TxtVac_ID.Text <> "" Then
    If SystemOptions.UserInterface = ArabicInterface Then
        MSGType = MsgBox("Â·  —€» ›Ì Õ–› Â–« «·”Ã·", vbQuestion + vbMsgBoxRtlReading + vbYesNo + vbMsgBoxRight, App.title)
     Else
        MSGType = MsgBox("Confirm Delete", vbQuestion + vbMsgBoxRtlReading + vbYesNo + vbMsgBoxRight, App.title)
   End If
        If MSGType = vbYes Then
           Cn.Execute "Delete from TblUsersStores where userid = " & val(TxtVac_ID.Text) & ""
           Cn.Execute "Delete from TblUsersBranches where userid = " & val(TxtVac_ID.Text) & ""
           Cn.Execute "Delete from TblUsersBoxes where userid = " & val(TxtVac_ID.Text) & ""
           Cn.Execute "Delete from TblUserAccount where UserID = " & val(TxtVac_ID.Text) & ""
           Cn.Execute "Delete from TblUsersProductLine where UserID = " & val(TxtVac_ID.Text) & ""
            RsSavRec.Find "userid=" & val(TxtVac_ID.Text), , adSearchForward, 1
            CuurentLogdata ("D")
            Dim StrSQL As String
            StrSQL = "Delete From TblUsersStores Where userid=" & RsSavRec("UserID").value
            Cn.Execute StrSQL, , adExecuteNoRecords
            
            
            StrSQL = "Delete From TblUsersProductLine Where userid=" & RsSavRec("UserID").value
            Cn.Execute StrSQL, , adExecuteNoRecords
                       
            
            StrSQL = "Delete From TblUsersBranches Where userid=" & RsSavRec("UserID").value
            Cn.Execute StrSQL, , adExecuteNoRecords
            
            Set Me.ImgPic.Picture = Nothing
            RsSavRec.delete
            If SystemOptions.UserInterface = ArabicInterface Then
               MsgBox " „  ⁄„·Ì… «·Õ–› »‰Ã«Õ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
            Else
               MsgBox "Delete Successfully", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
            End If
            '------------------------------ Move Next ---------------------------.
            FillGridWithData
            BtnNext_Click
        End If
    End If

    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217873, -2147467259
            'StrMSG = "⁄›Ê« ·« ÌÃÊ“ Õ–› «·”Ã· ·«— »«ÿÂ »»Ì«‰«  √Œ—Ì"
           If SystemOptions.UserInterface = ArabicInterface Then
            StrMSG = "⁄›Ê« ·« ÌÃÊ“ Õ–› «·”Ã· ·«— »«ÿÂ »»Ì«‰«  √Œ—Ì"
            Else
            StrMSG = "Sorry ... This record can not be deleted because it is linked to other data"
            End If
            RsSavRec.CancelUpdate
            MsgBox StrMSG, vbOKOnly + vbMsgBoxRight, App.title
            'clear the ConnectiOn Errors
            Cn.Errors.Clear
    End Select
End Sub
Private Sub BtnFirst_Click()

    On Error GoTo ErrTrap

    Dim Msg As String

    If Me.TxtModFlg.Text = "N" Then
        FindRec val(TxtVac_ID.Text)
        Me.TxtModFlg.Text = "R"
    End If

    TxtModFlg = "R"

    If RsSavRec.RecordCount = 0 Then
        clear_all Me
        Exit Sub
    End If

BegnieWork:
    RsSavRec.MoveFirst
    FiLLTXT

    Exit Sub

ErrTrap:

    Select Case Err.Number

        Case -2147217885
           ' Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & Chr(13)
           ' Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & Chr(13)
           ' Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
      If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
     Else
           Msg = "Sorry..The following record has been deleted" & CHR(13)
           Msg = Msg & "By another user on the network " & CHR(13)
           Msg = Msg & "The data will be updated " & CHR(13)
       End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select

End Sub

Private Sub BtnLast_Click()
    On Error GoTo ErrTrap

    Dim Msg As String

    If Me.TxtModFlg.Text = "N" Then
        FindRec val(TxtVac_ID.Text)
        Me.TxtModFlg.Text = "R"
    End If

    TxtModFlg = "R"

    If RsSavRec.RecordCount = 0 Then
        clear_all Me
        Exit Sub
    End If

BegnieWork:

    RsSavRec.MoveLast
    FiLLTXT
    Exit Sub

ErrTrap:

    Select Case Err.Number

        Case -2147217885
          '  Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & Chr(13)
          '  Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & Chr(13)
          '  Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
    If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
     Else
           Msg = "Sorry..The following record has been deleted" & CHR(13)
           Msg = Msg & "By another user on the network " & CHR(13)
           Msg = Msg & "The data will be updated " & CHR(13)
       End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select

End Sub

Private Sub btnModify_Click()
    Dim Msg As String

    If DoPremis(Do_Edit, Me.Name, True) = False Then
        Exit Sub
    End If

    On Error GoTo ErrTrap

    If TxtVac_ID.Text <> "" Then
        TxtModFlg = "E"
        Frm2.Enabled = True
        'Me.TXTDiscounts.SetFocus
        CuurentLogdata
    End If

    Exit Sub
ErrTrap:

    Select Case Err.Number

        Case -2147467259
            'Could not update; currently locked.
          '  Msg = "⁄›Ê«" & Chr(13)
          '  Msg = Msg & " ·«Ì„ﬂ‰  ⁄œÌ· Â–« «·”Ã· ›Ï «·Êﬁ  «·Õ«·Ï" & Chr(13)
          '  Msg = Msg & "ÕÌÀ «‰Â ﬁÌœ «· ⁄œÌ· „‰ ﬁ»· „” Œœ„ «Œ— ⁄·Ï «·‘»ﬂ…"
            If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê«" & CHR(13)
            Msg = Msg & " ·«Ì„ﬂ‰  ⁄œÌ· Â–« «·”Ã· ›Ï «·Êﬁ  «·Õ«·Ï" & CHR(13)
            Msg = Msg & "ÕÌÀ «‰Â ﬁÌœ «· ⁄œÌ· „‰ ﬁ»· „” Œœ„ «Œ— ⁄·Ï «·‘»ﬂ…"
           Else
            Msg = "Sorry..." & CHR(13)
            Msg = Msg & " This record can not be edited at this time" & CHR(13)
            Msg = Msg & "Because it was modified by another user on the network"
         
           End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
        
            If RsSavRec.EditMode <> adEditNone Then
                RsSavRec.CancelUpdate
                'RsSavRec.Requery
            End If

    End Select

End Sub

Private Sub btnNew_Click()
    Dim My_SQL As String
    Dim rs As ADODB.Recordset

    If DoPremis(Do_New, Me.Name, True) = False Then
        Exit Sub
    End If

    On Error GoTo ErrTrap
    Set rs = New ADODB.Recordset
    Frm2.Enabled = True
    '-----------------------------------
    Me.TxtVac_ID.Text = ""
 
    Me.DcBranches.BoundText = ""
    Me.DCEmP.BoundText = ""
    Me.DCJob.BoundText = ""
    Me.DCSalesRepGroups.BoundText = ""
    
    clear_all Me
    FillGridWithData
    CboPriv.ListIndex = 0
    '-----------------------------------
    TxtModFlg.Text = "N"

    My_SQL = "TBLSalesRepData"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable

    If rs.RecordCount > 0 Then
        TxtSerial.Text = rs.RecordCount + 1
    Else
        TxtSerial.Text = 1
    End If

    rs.Close
    CmbType.ListIndex = 0

    ListGroupSelected.Clear
    ListBoxesSelected.Clear
    ListStoreSelected.Clear
    
 
ErrTrap:
End Sub

Private Sub BtnNext_Click()
    On Error GoTo ErrTrap
    Dim Msg As String

    If Me.TxtModFlg.Text = "N" Then
        FindRec val(TxtVac_ID.Text)
        Me.TxtModFlg.Text = "R"
    End If

    TxtModFlg = "R"

    If RsSavRec.RecordCount = 0 Then
        clear_all Me
        Exit Sub
    End If

BegnieWork:

    If RsSavRec.EOF Then
        RsSavRec.MoveLast
    Else
        RsSavRec.MoveNext

        If RsSavRec.EOF Then
            RsSavRec.MoveLast
        End If
    End If

    FiLLTXT
    Exit Sub
ErrTrap:

    Select Case Err.Number

        Case -2147217885
            'Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & Chr(13)
            'Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & Chr(13)
            'Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
    If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
           Else
           Msg = "Sorry..The following record has been deleted" & CHR(13)
           Msg = Msg & "By another user on the network " & CHR(13)
           Msg = Msg & "The data will be updated " & CHR(13)
       End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select

End Sub

Private Sub BtnPrevious_Click()
    On Error GoTo ErrTrap
    Dim Msg As String

    If Me.TxtModFlg.Text = "N" Then
        FindRec val(TxtVac_ID.Text)
        Me.TxtModFlg.Text = "R"
    End If

    TxtModFlg = "R"

    If RsSavRec.RecordCount = 0 Then
        clear_all Me
        Exit Sub
    End If

BegnieWork:

    RsSavRec.MovePrevious

    If RsSavRec.BOF Then
        RsSavRec.MoveFirst
    End If

    FiLLTXT
    Exit Sub
ErrTrap:

    Select Case Err.Number

        Case -2147217885
          '  Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & Chr(13)
          '  Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & Chr(13)
          '  Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
         If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
           Else
           Msg = "Sorry..The following record has been deleted" & CHR(13)
           Msg = Msg & "By another user on the network " & CHR(13)
           Msg = Msg & "The data will be updated " & CHR(13)
       End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select

End Sub

Private Sub btnQuery_Click()
FrmUserSearch.show
FrmUserSearch.lblSearchtype = 0

End Sub

Private Sub btnSave_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
    Dim StrVacCode As String
    Dim StrVacName As String
    Dim CtrlTxt As Control
    '---------------------- check if data Vaclete -----------------------
    If CboPriv.ListIndex = -1 Then
    
        If SystemOptions.UserInterface = EnglishInterface Then
            Msg = "Specify Privligies"
        Else
              Msg = "Õœœ «·’·«ÕÌ« "
        End If
        
        
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        CboPriv.SetFocus
        Exit Sub
    End If
 
    If Trim(DcBranches.BoundText) = "" Then
        If SystemOptions.UserInterface = EnglishInterface Then
            Msg = "Specify Branch"
        Else
            Msg = "Õœœ «·›—⁄ «·«› —«÷Ì  "
        End If

        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        DcBranches.SetFocus
        SendKeys "{F4}"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
 
    If Trim(Me.DCEmP.BoundText) = "" Then
        If SystemOptions.UserInterface = EnglishInterface Then
            Msg = "Specify Employee"
        Else
            Msg = "Õœœ «·„ÊŸ›    "
        End If

        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        DCEmP.SetFocus
        SendKeys "{F4}"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
 
    If XPTxtUserName.Text = "" Then
    
         If SystemOptions.UserInterface = EnglishInterface Then
            Msg = "Specify User"
        Else
           Msg = "√œŒ· «”„ «·„” Œœ„"
        End If
        
        
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        XPTxtUserName.SetFocus
        Exit Sub
    End If

 '   If TxtPassWord.text = "" Then
 '       Msg = "√œŒ· ﬂ·„… «·„—Ê—"
 '       MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
 '       TxtPassWord.SetFocus
 '       Exit Sub
 '   End If
 '
 '   If XPTxtPassConfirm.text = "" Then
 '       Msg = "√œŒ·  √ﬂÌœ ﬂ·„… «·„—Ê—"
 '       MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
 '       XPTxtPassConfirm.SetFocus
 '       Exit Sub
 '   End If
Dim StrSQL As String
    If StrComp(TxtPassWord.Text, XPTxtPassConfirm.Text, vbTextCompare) <> 0 Then
            If SystemOptions.UserInterface = EnglishInterface Then
                Msg = "Passwords not matched"
            Else
                Msg = "ﬂ·„… «·„—Ê— Ê √ﬂÌœ ﬂ·„… «·„—Ê— " & CHR(13)
                Msg = Msg + "€Ì— „ ÿ«»ﬁ Ì‰"
             End If
        
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        XPTxtPassConfirm.SetFocus
        Exit Sub
    End If

    StrSQL = "select * From TblUsers where UserName='" & Trim(XPTxtUserName.Text) & "'" & " and UserID<>" & val(TxtVac_ID.Text)
    Set RsTemp = New ADODB.Recordset
    RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Not (RsTemp.EOF Or RsTemp.BOF) Then
    If SystemOptions.UserInterface = EnglishInterface Then
    Msg = "Another user already Exist with the same name"
    Else
        Msg = "ÌÊÃœ „” Œœ„ „”Ã· „”»ﬁ« »Â–« «·«”„" & CHR(13)
        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·«”„ «·’ÕÌÕ" & CHR(13)
        Msg = Msg + "√Ê  €ÌÌ— √Ê  „ÌÌ“ «”„ «·„” Œœ„"
    End If
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        XPTxtUserName.SetFocus
        RsTemp.Close
        Exit Sub
    End If

 
    '------------------------------ check if Empcode exist ----------------------
 
    ' -------------------------------------- txtmodflg type -------------------
    Select Case Me.TxtModFlg.Text

            '------------------------------ new record ----------------------------
        Case "N"
            '------------------------- save record -----------------------------
            AddNewRec
            BtnLast_Click

        Case "E"
            '----------------------------- save edit -------------------------------
            'RsEmployee("userid").value = RsSavRec("UserID").value
            StrSQL = "Delete From TblUsersStores Where userid=" & RsSavRec("UserID").value
            Cn.Execute StrSQL, , adExecuteNoRecords
            
            StrSQL = "Delete From TblUsersProductLine Where userid=" & RsSavRec("UserID").value
            Cn.Execute StrSQL, , adExecuteNoRecords
            
            StrSQL = "Delete From TblUsersBranches Where userid=" & RsSavRec("UserID").value
            Cn.Execute StrSQL, , adExecuteNoRecords
            
            
            FiLLRec
    End Select

    Exit Sub
ErrTrap:
If SystemOptions.UserInterface = EnglishInterface Then
MsgBox "error during saving", vbOKOnly + vbMsgBoxRight, App.title
Else
    MsgBox "Â‰«ﬂ Œÿ√ „« ›Ì ≈œŒ«· «·»Ì«‰« ", vbOKOnly + vbMsgBoxRight, App.title
End If
End Sub
 
Private Sub BtnUndo_Click()
    FindRec val(TxtVac_ID.Text)
    Me.TxtModFlg.Text = "R"
End Sub

Private Sub BtnUpdate_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
    Dim FristCount As Long
    Dim LastCount As Long
    FristCount = RsSavRec.RecordCount
    RsSavRec.Requery
    LastCount = RsSavRec.RecordCount
    BtnUndo_Click
If SystemOptions.UserInterface = ArabicInterface Then
    If FristCount = LastCount Then
        Msg = "·«  ÊÃœ »Ì«‰«  ÃœÌœ…"
    Else
        Msg = "⁄œœ «·”Ã·«  ﬁ»· «· ÕœÌÀ" & vbCrLf & FristCount & vbCrLf & "⁄œœ «·”Ã·«  »⁄œ «· ÕœÌÀ" & vbCrLf & LastCount
        
        If LastCount > FristCount Then
            Msg = Msg + vbCrLf & "⁄œœ «·”Ã·«  «·ÃœÌœ…" & vbCrLf & LastCount - FristCount
        Else
            Msg = Msg + vbCrLf & "⁄œœ «·”Ã·«  «·„Õ–Ê›…" & vbCrLf & FristCount - LastCount
        End If
    End If
Else
    If FristCount = LastCount Then
        Msg = "No new data"
    Else
        Msg = "Number of records before update" & vbCrLf & FristCount & vbCrLf & "Number of records after  update" & vbCrLf & LastCount
        
        If LastCount > FristCount Then
            Msg = Msg + vbCrLf & "Number of new records" & vbCrLf & LastCount - FristCount
        Else
            Msg = Msg + vbCrLf & "Number of records deleted" & vbCrLf & FristCount - LastCount
        End If
    End If
End If
    MsgBox Msg, vbOKOnly + vbMsgBoxRight + vbInformation, App.title
ErrTrap:
End Sub

Private Sub CmdPic_Click(Index As Integer)
On Error GoTo ErrTrap
    Select Case Index

        Case 0

            With cdg
               
                .CancelError = False
                .DialogTitle = " ≈Œ Ì«— ’Ê—…"
                'Set The Filter to show pictures only
                .filter = "Bitmap (*.bmp)|*.bmp|JPEG(*.JPG,*.JPEG,*.JPE,*.JFIF)|*.jpg;*.jpeg;*.jpe;*.jfif|"  ' choose formats to include
          
                .ShowOpen

                If .filename <> "" Then
                    Set Me.ImgPic.Picture = LoadPicture(.filename)
                End If

            End With

        Case 1
            Set Me.ImgPic.Picture = Nothing
    End Select
    Exit Sub
ErrTrap:
If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox " ÕÃ„ «·’Ê—… €Ì— „œ⁄Ê„", vbCritical
Else
MsgBox " image Size Not Siutable, vbCritical"
End If


End Sub

Private Sub DBCboClientName_Change()
Dim Fullcode As String
    GetCustomersDetail val(DBCboClientName.BoundText), , Fullcode ', 1, DepitIntervalID, DepitInterval, , creditlocked
    TxtSearchCode.Text = Fullcode

End Sub

Private Sub ImgPic_DblClick()
  Load FrmViewPic
    Set FrmViewPic.MainView.Picture = ImgPic.Picture
    FrmViewPic.show vbModal
End Sub

Private Sub Label15_Click()
'    If ListBoxesSelected.ListIndex > -1 Then
'        ListBoxesSelected.RemoveItem ListBoxesSelected.ListIndex
'    End If
'
'
Dim i As Long

For i = 0 To ListBoxesSelected.ListCount - 1
    If i > ListBoxesSelected.ListCount - 1 Then Exit For
    If ListBoxesSelected.Selected(i) Then
        ListBoxesSelected.RemoveItem i
        'ListStoreSelected.ListIndex
        i = i - 1
    End If
Next

End Sub

Private Sub Label16_Click()
    ListBoxesSelected.Clear
End Sub
Private Sub Label17_Click()
    Dim i As Integer
    
    ListBoxesSelected.Clear

    For i = 0 To ListBoxesAll.ListCount - 1
        ListBoxesSelected.AddItem ListBoxesAll.List(i)
        ListBoxesSelected.ItemData(i) = ListBoxesAll.ItemData(i)
    Next i

End Sub

Private Sub Label18_Click()
'    If ListBoxesAll.ListIndex = -1 Then Exit Sub
'    ListBoxesSelected.AddItem ListBoxesAll.List(ListBoxesAll.ListIndex)
'    ListBoxesSelected.ItemData(ListBoxesSelected.NewIndex) = ListBoxesAll.ItemData(ListBoxesAll.ListIndex)

    Dim i As Long
    
    For i = 0 To ListBoxesAll.ListCount - 1
        If ListBoxesAll.Selected(i) Then
            ListBoxesSelected.AddItem ListBoxesAll.List(i)
            ListBoxesSelected.ItemData(ListBoxesSelected.NewIndex) = ListBoxesAll.ItemData(i)
            
        End If
    Next
            

End Sub

Private Sub Label21_Click()
'    If ListAllAccount.ListIndex = -1 Then Exit Sub
'    ListAccountSelect.AddItem ListAllAccount.List(ListAllAccount.ListIndex)
'    ListAccountSelect.ItemData(ListAccountSelect.NewIndex) = ListAllAccount.ItemData(ListAllAccount.ListIndex)
'
    
            If ListStoreall.ListIndex = -1 Then Exit Sub
'    ListStoreSelected.AddItem ListStoreall.List(ListStoreall.ListIndex)
'    ListStoreSelected.ItemData(ListStoreSelected.NewIndex) = ListStoreall.ItemData(ListStoreall.ListIndex)
'
    
Dim i As Long

For i = 0 To ListAllAccount.ListCount - 1
    If ListAllAccount.Selected(i) Then
        ListAccountSelect.AddItem ListAllAccount.List(i)
        ListAccountSelect.ItemData(ListAccountSelect.NewIndex) = ListAllAccount.ItemData(i)
        
    End If
Next

'ItemData (i)

End Sub

Private Sub Label23_Click()
    Dim i As Integer
    ListAccountSelect.Clear
    For i = 0 To ListAllAccount.ListCount - 1
        ListAccountSelect.AddItem ListAllAccount.List(i)
        ListAccountSelect.ItemData(i) = ListAllAccount.ItemData(i)
    Next i
End Sub

Private Sub Label24_Click()
 ListAccountSelect.Clear
End Sub

Private Sub Label25_Click()
 
        

Dim i As Long

For i = 0 To ListAccountSelect.ListCount - 1
    If i > ListAccountSelect.ListCount - 1 Then Exit For
    If ListAccountSelect.Selected(i) Then
        ListAccountSelect.RemoveItem i
        'ListStoreSelected.ListIndex
        i = i - 1
    End If
Next

End Sub

Private Sub ListAllAccount_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
                      Account_search.show
                     Account_search.case_id = 78912

                   End If
End Sub

Private Sub ListBoxesAll_KeyUp(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyF3 Then
        FrmExpensesSearch.Indx = 2
        FrmExpensesSearch.RetrunType = 986
        FrmExpensesSearch.show
    End If
End Sub

Private Sub ListStoreall_KeyUp(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyF3 Then
        FrmStoreSearch.mIndex = 1
        Set FrmStoreSearch.RetrunFrm = Me
        FrmStoreSearch.show
    End If
End Sub

Private Sub TxtSearchCode_KeyPress(KeyAscii As Integer)
    Dim CUSTID As Integer
    
    If KeyAscii = vbKeyReturn Then
        GetCustomersDetail CUSTID, , TxtSearchCode.Text, 1
        DBCboClientName.BoundText = CUSTID
    End If
End Sub
Private Sub dcEmp_Change()
    If Me.TxtModFlg = "N" Or Me.TxtModFlg = "E" Then
        If (Me.DCEmP.BoundText) = "" Then Exit Sub
        Me.TXTCode.Text = get_EMPLOYEE_Data(val(Me.DCEmP.BoundText), "Fullcode")
        'DCEmp.text = DCEmp.text
    End If
End Sub
Private Sub Dcemp_Click(Area As Integer)
    dcEmp_Change
End Sub
Private Sub DCEmP_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        FrmEmployeeSearch.lbltype = 2911
        Set FrmEmployeeSearch.RetrunFrm = Me
        FrmEmployeeSearch.show
    End If
End Sub
Private Sub Form_Load()

    On Error GoTo ErrTrap
    
    Dim i As Integer
    Dim My_SQL As String
    Dim Dcombos As ClsDataCombos

    My_SQL = "TblUsers"
    Set BKGrndPic = New ClsBackGroundPic
    Set RsSavRec = New ADODB.Recordset
    RsSavRec.CursorLocation = adUseClient

    RsSavRec.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdTableDirect
    
    
    Me.TxtModFlg.Text = "R"
    Resize_Form Me
    'load tblUsers -----------------------------------------------
    My_SQL = "select UserID,UserName From tblUsers "
    fill_combo DCUser, My_SQL
    Set Dcombos = New ClsDataCombos
    Dcombos.GetBranches Me.DcBranches
    Dcombos.GetEmployees Me.DCEmP
    Dcombos.GetCustomersSuppliers 1, Me.DBCboClientName
    Dcombos.GetCustomersSuppliers 2, Me.DBCboClientName1
    Dcombos.GetStores Me.DCStore
    Dcombos.GetStores Me.DCStore1
    Dcombos.GetStores Me.DCStore3
    Dcombos.GetStores Me.DCStore2
    Dcombos.GetBoxes Me.DCBoxes
    Dcombos.GetBoxes Me.DCBoxes1
    Dcombos.GetBanks Me.Dbanks

    Set cSearch = New clsDCboSearch
    Set cSearch.Client = Me.DCEmP
    Set cSearch.Client = Me.DcBranches
    Set cSearch.Client = Me.DCStore
    Set cSearch.Client = Me.DCBoxes
    Set cSearch.Client = Me.DCBoxes1
    Set cSearch.Client = Me.Dbanks


    ModFgLib.LinkFgColWithDataCombo Grid, Grid.ColIndex("EmpName"), Me.DCEmP
    ModFgLib.LinkFgColWithDataCombo Grid, Grid.ColIndex("BranchId"), Me.DcBranches
    
    ModFgLib.LinkFgColWithDataCombo Grid, Grid.ColIndex("StoreID"), Me.DCStore
    ModFgLib.LinkFgColWithDataCombo Grid, Grid.ColIndex("BoxID"), Me.DCBoxes
    'ModFgLib.LinkFgColWithDataCombo Grid, Grid.ColIndex("BoxID"), Me.DCBoxes
    ModFgLib.LinkFgColWithDataCombo Grid, Grid.ColIndex("BankID"), Me.Dbanks

    FillGridWithData

    With Me.Grid
        .Cell(flexcpPicture, 0, .ColIndex("DiscountValue")) = Me.GrdImageList.ListImages("Vac_Name").ExtractIcon
        .Cell(flexcpPicture, 0, .ColIndex("Ser")) = Me.GrdImageList.ListImages("Ser").ExtractIcon
        For i = 0 To .Cols - 1
            .Cell(flexcpPictureAlignment, 0, i) = flexPicAlignRightCenter
        Next
   
        .ExtendLastCol = True
        .WallPaper = BKGrndPic.Picture
        .RowHeight(-1) = 300
    End With

    BtnFirst_Click
    ShowTip

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    FillMylist

    If OPEN_NEW_SCREEN = True Then
        btnNew_Click
    End If

ErrTrap:
End Sub
Private Sub ChangeLang()
    Dim XPic As IPictureDisp
    Set XPic = Me.btnFirst.ButtonImage
    Set Me.btnFirst.ButtonImage = Me.btnLast.ButtonImage
    Set Me.btnLast.ButtonImage = XPic
    Set XPic = Me.btnPrevious.ButtonImage
    Set Me.btnPrevious.ButtonImage = Me.btnNext.ButtonImage
    Set Me.btnNext.ButtonImage = XPic
    Label20.Caption = "All Accounts"
    Label19.Caption = "Selected Accounts"
    btnQuery.Caption = "Search"
    Ele(6).Caption = "Electronic Signature"
    Me.Caption = "Users Data"
    Me.Label1(2).Caption = Me.Caption
    Label1(3).Caption = "Code"
    Label1(1).Caption = "Name"
    lbl(4).Caption = "Privligies"
    Label1(9).Caption = "User Name"
    Label1(5).Caption = "Branch"
    Label1(4).Caption = "Box for Sale"
    Label1(14).Caption = "Box for Purchase"
    Label1(5).Caption = "Branch"
    Label1(0).Caption = "Password"
    Label1(8).Caption = "Re. password"
    Label1(7).Caption = "Sale Store"
    Label1(12).Caption = "Store Purchase"
    Label1(10).Caption = "Default Client"
    Label1(11).Caption = "Default Supplier"
    CmdPic(0).Caption = "Add Picture"
    CmdPic(1).Caption = "Delete Picture"
    Label1(6).Caption = "Bank"
    chkNextLogin.Caption = "Change password at login"
    
    chkHidLowering.Caption = "Hide the subtraction of output alerts"
Frame1.Caption = "Selected Boxes"
Frame2.Caption = "Selected Accounts"
    Frame11.Caption = "Selected Branch"
    Frame10.Caption = "Selected Stores"
    Label11.Caption = "All Branch"
    Label12.Caption = "Selected Branch"

    Label9.Caption = "All Stores"
    Label10.Caption = "Selected Stores"

    Label2(0).Caption = "Current Record"
    Label2(1).Caption = "NO. Recordes"

    btnNew.Caption = "New"
    btnModify.Caption = "Modify"
    btnSave.Caption = "Save"
    BtnUndo.Caption = "Undo"
    btnDelete.Caption = "Delete"
    btnCancel.Caption = "Exit"
 
    With Me.Grid
        .TextMatrix(0, .ColIndex("ser")) = "Ser"
        .TextMatrix(0, .ColIndex("EmpCode")) = "Code"
        .TextMatrix(0, .ColIndex("EmpName")) = "Emp Name"
        .TextMatrix(0, .ColIndex("JobID")) = "Job"
        .TextMatrix(0, .ColIndex("groupid")) = "Group"
        .TextMatrix(0, .ColIndex("BranchId")) = "Branch"
        .TextMatrix(0, .ColIndex("discountvalue")) = "Discount%"
        .TextMatrix(0, .ColIndex("UserName")) = "UserName"
        .TextMatrix(0, .ColIndex("StoreId")) = "Store Name"
        .TextMatrix(0, .ColIndex("boxId")) = "Box Name"
        .TextMatrix(0, .ColIndex("BankID")) = "Bank"
    End With
    
    '######### khaled was here ############
    isDeactivatedchk.Caption = "Deactivate User"
    Label14.Caption = "All Boxes"
    Label13.Caption = "Selected Boxes"

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
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
                btnSave_Click
            Case vbCancel
                Cancel = True
        End Select
    End If

    Exit Sub
ErrTrap:
End Sub
Private Sub Form_Terminate()
    'Set cSearchDCombo = Nothing
    'Set BKGrndPic = Nothing
    Set FrmVacancy = Nothing
End Sub
Private Sub Form_Unload(Cancel As Integer)

    On Error GoTo ErrTrap

    If RsSavRec.State = adStateOpen Then
        If Not (RsSavRec.EOF Or RsSavRec.BOF) Then
            If RsSavRec.EditMode <> adEditNone Then
                RsSavRec.CancelUpdate
            End If
        End If

        RsSavRec.Close
        Set RsSavRec = Nothing
    End If

    Set cSearch = Nothing
ErrTrap:
End Sub
Private Sub Label5_Click()
    If ListGroupSelected.ListIndex > -1 Then
        ListGroupSelected.RemoveItem ListGroupSelected.ListIndex
    End If
End Sub
Private Sub Label6_Click()
    ListGroupSelected.Clear
End Sub
Private Sub Label7_Click()
    Dim i As Integer
    ListGroupSelected.Clear

    For i = 0 To ListGroupAll.ListCount - 1
        ListGroupSelected.AddItem ListGroupAll.List(i)
        ListGroupSelected.ItemData(i) = ListGroupAll.ItemData(i)
    Next i
End Sub
Private Sub Label8_Click()
    If ListGroupAll.ListIndex = -1 Then Exit Sub
    ListGroupSelected.AddItem ListGroupAll.List(ListGroupAll.ListIndex)
    ListGroupSelected.ItemData(ListGroupSelected.NewIndex) = ListGroupAll.ItemData(ListGroupAll.ListIndex)
End Sub
Private Sub LblSelect_Click()
'    If ListStoreall.ListIndex = -1 Then Exit Sub
'    ListStoreSelected.AddItem ListStoreall.List(ListStoreall.ListIndex)
'    ListStoreSelected.ItemData(ListStoreSelected.NewIndex) = ListStoreall.ItemData(ListStoreall.ListIndex)
'
    
        If ListStoreall.ListIndex = -1 Then Exit Sub
'    ListStoreSelected.AddItem ListStoreall.List(ListStoreall.ListIndex)
'    ListStoreSelected.ItemData(ListStoreSelected.NewIndex) = ListStoreall.ItemData(ListStoreall.ListIndex)
'
    
Dim i As Long

For i = 0 To ListStoreall.ListCount - 1
    If ListStoreall.Selected(i) Then
        ListStoreSelected.AddItem ListStoreall.List(i)
        ListStoreSelected.ItemData(ListStoreSelected.NewIndex) = ListStoreall.ItemData(i)
        
    End If
Next

'ItemData (i)
End Sub
Private Sub Label22_Click()
    Dim i As Integer
    ListStoreSelected.Clear
    For i = 0 To ListStoreall.ListCount - 1
        ListStoreSelected.AddItem ListStoreall.List(i)
        ListStoreSelected.ItemData(i) = ListStoreall.ItemData(i)
    Next i
End Sub
Private Sub Label3_Click()
    ListStoreSelected.Clear
End Sub
Private Sub Label4_Click()
'    If ListStoreSelected.ListIndex > -1 Then
'        ListStoreSelected.RemoveItem ListStoreSelected.ListIndex
'    End If
    

Dim i As Long

For i = 0 To ListStoreSelected.ListCount - 1
    If i > ListStoreSelected.ListCount - 1 Then Exit For
    If ListStoreSelected.Selected(i) Then
        ListStoreSelected.RemoveItem i
        'ListStoreSelected.ListIndex
        i = i - 1
    End If
Next
    
    
End Sub
Function createlistString(mylist As ListBox, Optional ByRef Listitems As String)
    Dim i As Integer
    Dim str As String
    str = "0"
    Listitems = ""
    For i = 0 To mylist.ListCount - 1
        str = str & "," & mylist.ItemData(i)
        Listitems = Listitems & "," & mylist.List(i)
    Next i
    createlistString = str
End Function


Function FillMylist(Optional ByVal mIndexd As Long = 0)

    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim Balance As Double
    Dim i As Integer
    
    If mIndexd = 1 Or mIndexd = 0 Then
        sql = " SELECT * from  TblStore "
    
        If SystemOptions.UserInterface = ArabicInterface Then
            sql = sql & " order by  StoreName"
        Else
            sql = sql & " order by  StoreNamee"
        End If
        rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    
        ListStoreall.Clear
        'ListStoreSelected.Clear
    
        If rs.RecordCount > 0 Then
            For i = 1 To rs.RecordCount
                If SystemOptions.UserInterface = ArabicInterface Then
                    ListStoreall.AddItem IIf(IsNull(rs("StoreName").value), "", rs("StoreName").value)
                Else
                    ListStoreall.AddItem IIf(IsNull(rs("StoreNamee").value), "", rs("StoreNamee").value)
                End If
    
                ListStoreall.ItemData(ListStoreall.NewIndex) = rs("StoreID").value
                rs.MoveNext
            Next i
        End If
    
        rs.Close
    End If
    If mIndexd = 0 Then
        sql = " SELECT * from  TblBranchesData "
     
        If SystemOptions.UserInterface = ArabicInterface Then
            sql = sql & " order by  branch_name"
        Else
            sql = sql & " order by  branch_namee"
        End If
     
        rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    
        ListGroupAll.Clear
        'ListGroupSelected.Clear
    
        If rs.RecordCount > 0 Then
            For i = 1 To rs.RecordCount
                If SystemOptions.UserInterface = ArabicInterface Then
                    ListGroupAll.AddItem IIf(IsNull(rs("branch_name").value), "", rs("branch_name").value)
                Else
                    ListGroupAll.AddItem IIf(IsNull(rs("branch_namee").value), "", rs("branch_namee").value)
                End If
    
                ListGroupAll.ItemData(ListGroupAll.NewIndex) = rs("branch_id").value
                rs.MoveNext
            Next i
        End If
        rs.Close
    End If
'    sql = "select* from TblBoxesData where Type = 0 "
If mIndexd = 2 Or mIndexd = 0 Then
        sql = "select* from TblBoxesData    "
        ' sql = "select* from TblBoxesData where  "
        If SystemOptions.UserInterface = ArabicInterface Then
            sql = sql & " order by  BoxName"
        Else
            sql = sql & " order by  BoxNameE"
        End If
     
        rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    
        ListBoxesAll.Clear
        
        If rs.RecordCount > 0 Then
            For i = 1 To rs.RecordCount
                If SystemOptions.UserInterface = ArabicInterface Then
                    ListBoxesAll.AddItem IIf(IsNull(rs("BoxName").value), "", rs("BoxName").value)
                Else
                    ListBoxesAll.AddItem IIf(IsNull(rs("BoxNameE").value), "", rs("BoxNameE").value)
                End If
    
                ListBoxesAll.ItemData(ListBoxesAll.NewIndex) = rs("BoxID").value
                rs.MoveNext
            Next i
        End If
        rs.Close
      ''/////Account
   End If
   If mIndexd = 3 Or mIndexd = 0 Then
        sql = " SELECT * from  ACCOUNTS "
        sql = sql & " where   last_account=0"
        If SystemOptions.UserInterface = ArabicInterface Then
            sql = sql & " order by  Account_Name"
        Else
            sql = sql & " order by  Account_NameEng"
        End If
     
        rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    
        ListAllAccount.Clear
        'ListGroupSelected.Clear
    
        If rs.RecordCount > 0 Then
            For i = 1 To rs.RecordCount
                If SystemOptions.UserInterface = ArabicInterface Then
                    ListAllAccount.AddItem IIf(IsNull(rs("Account_Name").value), "", rs("Account_Name").value)
                Else
                    ListAllAccount.AddItem IIf(IsNull(rs("Account_NameEng").value), "", rs("Account_NameEng").value)
                End If
    
                ListAllAccount.ItemData(ListAllAccount.NewIndex) = rs("Account_ID").value
                rs.MoveNext
            Next i
        End If
        rs.Close
        
    End If
    If mIndexd = 0 Then
      sql = "select * from TblProductLine "
        ' sql = "select* from TblBoxesData where  "
       
        sql = sql & " order by  Name"
        
     
        rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    
        ListProductLineAll.Clear
        
        If rs.RecordCount > 0 Then
            For i = 1 To rs.RecordCount
                ListProductLineAll.AddItem IIf(IsNull(rs("Name").value), "", rs("Name").value)
    
                ListProductLineAll.ItemData(ListProductLineAll.NewIndex) = rs("ID").value
                rs.MoveNext
            Next i
        End If
    End If
End Function
Private Sub Form_Activate()
    Me.ZOrder 0
End Sub
Public Sub AddNewRec()

    On Error GoTo ErrTrap
    
    Dim StrRecID As String
    
    'StrRecID = new_id("TBLSalesRepData", "id", "")
    
    RsSavRec.AddNew
    RsSavRec("UserID").value = CStr(new_id("TblUsers", "UserID", "", True))
    'RsSavRec.Fields("id").value = IIf(StrRecID <> "", StrRecID, Null)
    FiLLRec
ErrTrap:
End Sub
Public Sub FiLLRec()

    On Error GoTo ErrTrap

    RsSavRec.Fields("PassWord").value = IIf((TxtPassWord.Text) <> "", (TxtPassWord.Text), "")
    RsSavRec.Fields("EmpID").value = IIf(val(Me.DCEmP.BoundText) <> 0, val(Me.DCEmP.BoundText), Null)
    RsSavRec.Fields("BranchId").value = IIf(val(Me.DcBranches.BoundText) <> 0, val(Me.DcBranches.BoundText), Null)
    RsSavRec.Fields("BoxID").value = IIf(val(Me.DCBoxes.BoundText) <> 0, val(Me.DCBoxes.BoundText), Null)
    RsSavRec.Fields("BoxID1").value = IIf(val(Me.DCBoxes1.BoundText) <> 0, val(Me.DCBoxes1.BoundText), Null)
    
    RsSavRec.Fields("BankID").value = IIf(val(Me.Dbanks.BoundText) <> 0, val(Me.Dbanks.BoundText), Null)
    RsSavRec.Fields("StoreID").value = IIf(val(Me.DCStore.BoundText) <> 0, val(Me.DCStore.BoundText), Null)
    RsSavRec.Fields("Custid").value = IIf(val(Me.DBCboClientName.BoundText) <> 0, val(Me.DBCboClientName.BoundText), Null)
    
    RsSavRec.Fields("StoreID1").value = IIf(val(Me.DCStore1.BoundText) <> 0, val(Me.DCStore1.BoundText), Null)
    RsSavRec.Fields("StoreID3").value = IIf(val(Me.DCStore3.BoundText) <> 0, val(Me.DCStore3.BoundText), Null)
    RsSavRec.Fields("StoreID2").value = IIf(val(Me.DCStore2.BoundText) <> 0, val(Me.DCStore2.BoundText), Null)
    RsSavRec.Fields("Custid1").value = IIf(val(Me.DBCboClientName1.BoundText) <> 0, val(Me.DBCboClientName1.BoundText), Null)
    
    RsSavRec("UserName").value = Trim(XPTxtUserName.Text)
    If ImgPic.Picture = 0 Then
        RsSavRec("UserSign").value = Null
    Else
        If SavePictureToDB(ImgPic, RsSavRec, "UserSign") = False Then
            GoTo ErrTrap
        End If
    End If

    If Me.CboPriv.ListIndex = 0 Then
        RsSavRec("UserType").value = 2
    Else
        RsSavRec("InvPrices").value = 1
        RsSavRec("InvPrices1").value = 1
        RsSavRec("InvPrices2").value = 1
        
        RsSavRec("ShowInvProfit").value = 1
        RsSavRec("AllowOverMax").value = 1

        RsSavRec("FullPremis").value = 1
        RsSavRec("UserType").value = 0
    End If
 
    RsSavRec("PassConfirm").value = Trim(XPTxtPassConfirm.Text)
   
    RsSavRec("IsActive").value = 1
    
 
    If chkNextLogin.value = vbChecked Then
        RsSavRec("ChangePW").value = 1
    Else
        RsSavRec("ChangePW").value = 0
    End If
    
    If chkHidLowering.value = vbChecked Then
        RsSavRec("HidLowering").value = 1
    Else
        RsSavRec("HidLowering").value = 0
    End If
        
    
    
    '########## Khaled's was here #################
    If isDeactivatedchk.value = vbChecked Then
        RsSavRec("isDeactivated").value = 1
    Else
        RsSavRec("isDeactivated").value = 0
    End If
    '###############################################
 
    
    'RsSavRec.Fields("JobID").value = IIf(Me.DCJob.BoundText <> 0, Val(Me.DCJob.BoundText), Null)

    RsSavRec.update
    Dim UsrID As Double
   UsrID = IIf(IsNull(RsSavRec("UserID").value), 0, RsSavRec("UserID").value)
    If Me.TxtModFlg.Text = "E" Then
    Cn.Execute "Delete from TblUsersStores where userid = " & UsrID & ""
    Cn.Execute "Delete from TblUsersBranches where userid = " & UsrID & ""
    Cn.Execute "Delete from TblUsersBoxes where userid = " & UsrID & ""
    Cn.Execute "Delete from TblUserAccount where UserID = " & UsrID & ""
    Cn.Execute "Delete from TblUsersProductLine where UserID = " & UsrID & ""
    
    End If
    Dim i As Integer
    Dim RsEmployee As New ADODB.Recordset
    
        If ListStoreSelected.ListCount <> 0 Then
        Set RsEmployee = New ADODB.Recordset
            RsEmployee.Open "TblUsersStores", Cn, adOpenKeyset, adLockOptimistic, adCmdTable
            For i = 0 To ListStoreSelected.ListCount - 1
                RsEmployee.AddNew
                RsEmployee("storeId").value = ListStoreSelected.ItemData(i)
                RsEmployee("userid").value = UsrID
                RsEmployee.update
            Next i
            RsEmployee.Close
            Set RsEmployee = Nothing
        End If

        If ListGroupSelected.ListCount <> 0 Then
        Set RsEmployee = New ADODB.Recordset
            RsEmployee.Open "TblUsersBranches", Cn, adOpenKeyset, adLockOptimistic, adCmdTable
                For i = 0 To ListGroupSelected.ListCount - 1
                    RsEmployee.AddNew
                    RsEmployee("BranchID").value = ListGroupSelected.ItemData(i)
                    RsEmployee("userid").value = UsrID
                    RsEmployee.update
            Next i
            RsEmployee.Close
            Set RsEmployee = Nothing
        End If
        
        If ListBoxesSelected.ListCount <> 0 Then
        Set RsEmployee = New ADODB.Recordset
            RsEmployee.Open "TblUsersBoxes", Cn, adOpenKeyset, adLockOptimistic, adCmdTable
                For i = 0 To ListBoxesSelected.ListCount - 1
                    RsEmployee.AddNew
                    RsEmployee("BoxId").value = ListBoxesSelected.ItemData(i)
                    RsEmployee("userid").value = UsrID
                    RsEmployee.update
            Next i
            RsEmployee.Close
            Set RsEmployee = Nothing
        End If
        
    If ListProductLineSelected.ListCount <> 0 Then
        Set RsEmployee = New ADODB.Recordset
            RsEmployee.Open "TblUsersProductLine", Cn, adOpenKeyset, adLockOptimistic, adCmdTable
                For i = 0 To ListProductLineSelected.ListCount - 1
                    RsEmployee.AddNew
                    RsEmployee("ProductLineId").value = ListProductLineSelected.ItemData(i)
                    RsEmployee("userid").value = UsrID
                    'RsEmployee("ShowAlarm").value = FG.ValueMatrix(i, FG.ColIndex("ShowAlarm"))
                    RsEmployee("TypeLine").value = 0
                    
                    RsEmployee.update
            Next i
            RsEmployee.Close
            Set RsEmployee = Nothing
        End If
        Dim sql As String
        
        sql = "Select * from TblUsersProductLine Where  TypeLine = 1 "
        
        saveGrid sql, FG, "ShowAlarm", "", "userId", UsrID, "TypeLine", 1
        
        
         If ListAccountSelect.ListCount <> 0 Then
         sql = "select * from TblUserAccount   where 1=-1"
        Set RsEmployee = New ADODB.Recordset
            RsEmployee.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
                For i = 0 To ListAccountSelect.ListCount - 1
                    RsEmployee.AddNew
                    RsEmployee("Account_ID").value = ListAccountSelect.ItemData(i)
                    RsEmployee("UserID").value = UsrID
                    RsEmployee.update
            Next i
            RsEmployee.Close
            Set RsEmployee = Nothing
        End If
        
        CuurentLogdata
        
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox " „  ⁄„·Ì… «·Õ›Ÿ »‰Ã«Õ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Else
            MsgBox "Saved Successfully", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        End If
    
        FillGridWithData
        TxtModFlg = "R"

    Exit Sub
ErrTrap:

    If RsSavRec.EditMode <> adEditNone Then
        RsSavRec.CancelUpdate
    End If

End Sub
Public Sub FiLLTXT()

    On Error GoTo ErrTrap
    
    Dim i As Integer
    Frm2.Enabled = False
    TxtVac_ID.Text = IIf(IsNull(RsSavRec.Fields("UserID").value), "", RsSavRec.Fields("UserID").value)
    TxtPassWord.Text = IIf(IsNull(RsSavRec.Fields("PassWord").value), "", RsSavRec.Fields("PassWord").value)
    XPTxtPassConfirm.Text = IIf(IsNull(RsSavRec.Fields("PassConfirm").value), "", RsSavRec.Fields("PassConfirm").value)
    XPTxtUserName.Text = IIf(IsNull(RsSavRec.Fields("UserName").value), 0, RsSavRec.Fields("UserName").value)
    If Not IsNull(RsSavRec("UserType").value) Then
        If RsSavRec("UserType").value = 2 Then
            CboPriv.ListIndex = 0
        Else
            CboPriv.ListIndex = 1
        End If
    End If
      
    If Not IsNull(RsSavRec("ChangePW").value) Then
        If RsSavRec("ChangePW").value = 0 Then
            chkNextLogin.value = vbUnchecked
        Else
            chkNextLogin.value = vbChecked
        End If
    Else
        chkNextLogin.value = vbUnchecked
    End If
    
   
    If Not IsNull(RsSavRec("HidLowering").value) Then
        If RsSavRec("HidLowering").value = 0 Then
            chkHidLowering.value = vbUnchecked
        Else
            chkHidLowering.value = vbChecked
        End If
    Else
        chkHidLowering.value = vbUnchecked
    End If
    
     
    
    
    '################# khaled was here #####################
    If Not IsNull(RsSavRec("isDeactivated").value) Then
        If RsSavRec("isDeactivated").value = 0 Then
            isDeactivatedchk.value = vbUnchecked
        Else
            isDeactivatedchk.value = vbChecked
        End If
    Else
        isDeactivatedchk.value = vbUnchecked
    End If
    '#######################################################
       
    Me.DCEmP.BoundText = IIf(IsNull(RsSavRec.Fields("EmpID").value), "", RsSavRec.Fields("EmpID").value)
    TXTCode.Text = get_EMPLOYEE_Data(val(Me.DCEmP.BoundText), "fullcode")
    Me.DcBranches.BoundText = IIf(IsNull(RsSavRec.Fields("BranchId").value), "", RsSavRec.Fields("BranchId").value)
    Me.DCStore.BoundText = IIf(IsNull(RsSavRec.Fields("StoreID").value), "", RsSavRec.Fields("StoreID").value)
    Me.DCStore1.BoundText = IIf(IsNull(RsSavRec.Fields("StoreID1").value), "", RsSavRec.Fields("StoreID1").value)
    Me.DCStore3.BoundText = IIf(IsNull(RsSavRec.Fields("StoreID3").value), "", RsSavRec.Fields("StoreID3").value)
        Me.DCStore2.BoundText = IIf(IsNull(RsSavRec.Fields("StoreID2").value), "", RsSavRec.Fields("StoreID2").value)
    Me.DCBoxes1.BoundText = IIf(IsNull(RsSavRec.Fields("BoxID1").value), "", RsSavRec.Fields("BoxID1").value)
    Me.DCBoxes.BoundText = IIf(IsNull(RsSavRec.Fields("BoxID").value), "", RsSavRec.Fields("BoxID").value)
    Me.Dbanks.BoundText = IIf(IsNull(RsSavRec.Fields("BankID").value), "", RsSavRec.Fields("BankID").value)
    Me.DBCboClientName.BoundText = IIf(IsNull(RsSavRec.Fields("Custid").value), "", RsSavRec.Fields("Custid").value)
    Me.DBCboClientName1.BoundText = IIf(IsNull(RsSavRec.Fields("Custid1").value), "", RsSavRec.Fields("Custid1").value)
    If Not IsNull(RsSavRec("UserSign").value) Then
        If LenB(RsSavRec("UserSign")) Then
            LoadPictureFromDB ImgPic, RsSavRec, "UserSign"
        Else
            Set ImgPic.Picture = Nothing
        End If
    Else
        Set ImgPic.Picture = Nothing
    End If
    

'********************************************************************
     
    ListStoreSelected.Clear

    Dim RsEmployee As ADODB.Recordset
    Dim StrSQL As String
 
    StrSQL = " SELECT     TOP 100 PERCENT dbo.TblStore.StoreName, dbo.TblStore.StoreNamee, dbo.TblStore.StoreID"
    StrSQL = StrSQL & "  FROM         dbo.TblUsersStores INNER JOIN"
    StrSQL = StrSQL & "   dbo.TblStore ON dbo.TblUsersStores.StoreID = dbo.TblStore.StoreID"
    StrSQL = StrSQL & "  Where (dbo.TblUsersStores.UserID = " & val(TxtVac_ID.Text) & ")"
    StrSQL = StrSQL & "  ORDER BY dbo.TblUsersStores.id"

    Set RsEmployee = New ADODB.Recordset

    RsEmployee.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsEmployee.BOF Or RsEmployee.EOF) Then
        For i = 0 To RsEmployee.RecordCount - 1
            If SystemOptions.UserInterface = ArabicInterface Then
                ListStoreSelected.AddItem IIf(IsNull(RsEmployee("StoreName").value), "", RsEmployee("StoreName").value)
            Else
                ListStoreSelected.AddItem IIf(IsNull(RsEmployee("StoreNameE").value), "", RsEmployee("StoreNameE").value)
            End If
            ListStoreSelected.ItemData(i) = val(IIf(IsNull(RsEmployee("StoreID").value), 0, (RsEmployee("StoreID").value)))
            RsEmployee.MoveNext
        Next i

        RsEmployee.Close
        Set RsEmployee = Nothing
    End If


'*********************************************************************************
     
    ListGroupSelected.Clear

    StrSQL = " SELECT     TOP 100 PERCENT dbo.TblBranchesData.branch_id, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee"
    StrSQL = StrSQL & " FROM         dbo.TblUsersBranches INNER JOIN"
    StrSQL = StrSQL & " dbo.TblBranchesData ON dbo.TblUsersBranches.BranchID = dbo.TblBranchesData.branch_id"
    StrSQL = StrSQL & " Where (dbo.TblUsersBranches.UserID = " & val(TxtVac_ID.Text) & ")"
    StrSQL = StrSQL & " ORDER BY dbo.TblUsersBranches.id"

    Set RsEmployee = New ADODB.Recordset

    RsEmployee.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsEmployee.BOF Or RsEmployee.EOF) Then
        For i = 0 To RsEmployee.RecordCount - 1
            If SystemOptions.UserInterface = ArabicInterface Then
                ListGroupSelected.AddItem IIf(IsNull(RsEmployee("branch_name").value), "", RsEmployee("branch_name").value)
            Else
                ListGroupSelected.AddItem IIf(IsNull(RsEmployee("branch_nameE").value), "", RsEmployee("branch_nameE").value)
            End If
            ListGroupSelected.ItemData(i) = val(IIf(IsNull(RsEmployee("branch_id").value), 0, (RsEmployee("branch_id").value)))
                
            RsEmployee.MoveNext
        Next i

        RsEmployee.Close
        Set RsEmployee = Nothing
    End If
'*********************************************************************************
    ListBoxesSelected.Clear
    
    StrSQL = "SELECT TblUsersBoxes.id, TblBoxesData.BoxName, TblUsersBoxes.BoxId, TblUsersBoxes.userid, TblBoxesData.BoxNameE"
    StrSQL = StrSQL & " FROM TblUsersBoxes INNER JOIN"
    StrSQL = StrSQL & " TblBoxesData ON TblUsersBoxes.BoxId = TblBoxesData.BoxID"
    StrSQL = StrSQL & " Where (TblUsersBoxes.UserID = " & val(TxtVac_ID.Text) & ")"

    Set RsEmployee = New ADODB.Recordset

    RsEmployee.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsEmployee.BOF Or RsEmployee.EOF) Then
        For i = 0 To RsEmployee.RecordCount - 1
            If SystemOptions.UserInterface = ArabicInterface Then
                ListBoxesSelected.AddItem IIf(IsNull(RsEmployee("BoxName").value), "", RsEmployee("BoxName").value)
            Else
                ListBoxesSelected.AddItem IIf(IsNull(RsEmployee("BoxNameE").value), "", RsEmployee("BoxNameE").value)
            End If
            ListBoxesSelected.ItemData(i) = val(IIf(IsNull(RsEmployee("BoxId").value), 0, (RsEmployee("BoxId").value)))
                
            RsEmployee.MoveNext
        Next i

        RsEmployee.Close
        Set RsEmployee = Nothing
    End If
    
  ListProductLineSelected.Clear
    
    StrSQL = "SELECT TblUsersProductLine.id,TblUsersProductLine.ShowAlarm, TblProductLine.Name, TblUsersProductLine.ProductLineId, TblUsersProductLine.userid "
    StrSQL = StrSQL & " FROM TblUsersProductLine INNER JOIN"
    StrSQL = StrSQL & " TblProductLine ON TblUsersProductLine.ProductLineId = TblProductLine.ID"
    StrSQL = StrSQL & " Where (TblUsersProductLine.UserID = " & val(TxtVac_ID.Text) & ") and IsNull( TypeLine,0) = 0"

    Set RsEmployee = New ADODB.Recordset

    RsEmployee.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    
    StrSQL = "SELECT TblProductLine.id,TblProductLine.id as ProductLineID ,TblUsersProductLine.ShowAlarm, TblProductLine.Name,  TblUsersProductLine.userid "
    StrSQL = StrSQL & " FROM TblUsersProductLine RIGHT outer JOIN "
    StrSQL = StrSQL & " TblProductLine ON TblUsersProductLine.ProductLineId = TblProductLine.ID and (TblUsersProductLine.UserID = " & val(TxtVac_ID.Text) & ") and TypeLine = 1"
    
    loadgrid StrSQL, FG, True, False
    
    If Not (RsEmployee.BOF Or RsEmployee.EOF) Then
        For i = 0 To RsEmployee.RecordCount - 1
            
                ListProductLineSelected.AddItem IIf(IsNull(RsEmployee("Name").value), "", RsEmployee("Name").value)
            
            ListProductLineSelected.ItemData(i) = val(IIf(IsNull(RsEmployee("ProductLineId").value), 0, (RsEmployee("ProductLineId").value)))
                
            RsEmployee.MoveNext
        Next i

        RsEmployee.Close
        Set RsEmployee = Nothing
    End If
        
'''//////////////
    ListAccountSelect.Clear
    
    StrSQL = " SELECT     dbo.TblUserAccount.UserID, dbo.TblUserAccount.Account_ID, dbo.ACCOUNTS.Account_Name, dbo.ACCOUNTS.Account_NameEng"
    StrSQL = StrSQL & " FROM         dbo.TblUserAccount LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.ACCOUNTS ON dbo.TblUserAccount.Account_ID = dbo.ACCOUNTS.Account_ID"
    StrSQL = StrSQL & "     Where (dbo.TblUserAccount.UserID = " & val(TxtVac_ID.Text) & ")"
    Set RsEmployee = New ADODB.Recordset

    RsEmployee.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsEmployee.BOF Or RsEmployee.EOF) Then
        For i = 0 To RsEmployee.RecordCount - 1
            If SystemOptions.UserInterface = ArabicInterface Then
                ListAccountSelect.AddItem IIf(IsNull(RsEmployee("Account_Name").value), "", RsEmployee("Account_Name").value)
            Else
                ListAccountSelect.AddItem IIf(IsNull(RsEmployee("Account_NameEng").value), "", RsEmployee("Account_NameEng").value)
            End If
            ListAccountSelect.ItemData(i) = val(IIf(IsNull(RsEmployee("Account_ID").value), 0, (RsEmployee("Account_ID").value)))
                
            RsEmployee.MoveNext
        Next i

        RsEmployee.Close
        Set RsEmployee = Nothing
    End If
'*********************************************************************************
    'Me.DCJob.BoundText = IIf(IsNull(RsSavRec.Fields("JobID").value), "", RsSavRec.Fields("JobID").value)

    LabCurrRec.Caption = RsSavRec.AbsolutePosition
    LabCountRec.Caption = RsSavRec.RecordCount
    With Grid
        For i = 1 To .Rows - 1
            If Trim(TxtVac_ID.Text) = .TextMatrix(i, .ColIndex("UserID")) Then
                TxtSerial.Text = .TextMatrix(i, .ColIndex("Ser"))
                .Row = i
                Exit Sub
            End If
        Next
    End With
ErrTrap:

End Sub
Public Sub EditRec(StrTable As String, RecId As String)
    'My_SQL = "select * From " & StrTable & " where "
    'RsSavRec.Open My_SQL, cn, adOpenStatic, adLockReadOnly, adCmdTableDirect
    FiLLRec
End Sub

Private Sub Grid_EnterCell()
    On Error GoTo ErrTrap
    FindRec val(Me.Grid.TextMatrix(Me.Grid.Row, Me.Grid.ColIndex("UserID")))
ErrTrap:
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
If Me.TxtModFlg = "N" Or Me.TxtModFlg = "E" Then
    DCEmP.BoundText = GeTEmpIDByEmpCode(TXTCode.Text)
End If
End Sub
Private Sub TxtPassWord_DblClick()
'     If user_id = 1 Then
'     MsgBox txtPassword
'     End If
End Sub
Private Sub TxtSearchCode1_KeyPress(KeyAscii As Integer)
    Dim CUSTID As Integer
    If KeyAscii = vbKeyReturn Then
        GetCustomersDetail CUSTID, , TxtSearchCode.Text, 1
        DBCboClientName.BoundText = CUSTID
    End If
End Sub
Private Sub TxtVac_ID_Change()
    Dim TxtMod As String
    TxtMod = TxtModFlg.Text
    TxtModFlg.Text = ""
    TxtModFlg = TxtMod
End Sub
Public Function FindRec(ByVal RecId As Long)
    On Error GoTo ErrTrap
    RsSavRec.Find "UserID=" & RecId, , adSearchForward, 1

    If Not (RsSavRec.EOF) Then
        FiLLTXT
    End If

    Exit Function
ErrTrap:

    If RsSavRec.EditMode <> adEditNone Then
        RsSavRec.CancelUpdate
        BtnUndo_Click
    End If

    'RsSavRec.Filter = adFilterNone
End Function

'Private Sub TxtVacCode_KeyPress(KeyAscii As Integer)
'KeyAscii = DataFormat(ChrOnly, KeyAscii)
'End Sub

Private Sub TxtModFlg_Change()

    If TxtModFlg.Text = "N" Then
        Frm2.Enabled = True
        Me.btnNew.Enabled = False
        btnModify.Enabled = False
        btnDelete.Enabled = False
        Me.btnQuery.Enabled = False
        Grid.Enabled = False
        BtnUndo.Enabled = True
        Me.btnSave.Enabled = True
        BtnUpdate.Enabled = False
        'btnNext.Enabled = False
        'btnPrevious.Enabled = False
        'btnFirst.Enabled = False
        'btnLast.Enabled = False
        ListGroupAll.Enabled = True
        ListStoreall.Enabled = True
        ListBoxesAll.Enabled = True
        ListAllAccount.Enabled = True
        ListProductLineAll.Enabled = True
    ElseIf TxtModFlg.Text = "R" Then
        ListAllAccount.Enabled = False
        ListProductLineAll.Enabled = False
        Frm2.Enabled = False
        Grid.Enabled = True
        btnModify.Enabled = False
        btnDelete.Enabled = False
        If TxtVac_ID.Text <> "" Then
            btnModify.Enabled = True
            btnDelete.Enabled = True
        End If
        BtnUpdate.Enabled = True
        Me.btnQuery.Enabled = True
        Me.btnNew.Enabled = True
        BtnUndo.Enabled = False
        Me.btnSave.Enabled = False
        btnNext.Enabled = True
        btnPrevious.Enabled = True
        btnFirst.Enabled = True
        btnLast.Enabled = True
        ListGroupAll.Enabled = False
        ListStoreall.Enabled = False
        ListBoxesAll.Enabled = False
    ElseIf TxtModFlg.Text = "E" Then
        ListAllAccount.Enabled = True
                ListProductLineAll.Enabled = True
        Frm2.Enabled = True
        Me.btnNew.Enabled = False
        btnModify.Enabled = False
        btnDelete.Enabled = False
        Me.btnQuery.Enabled = False
        BtnUpdate.Enabled = False
        BtnUndo.Enabled = True
        Me.btnSave.Enabled = True
        Grid.Enabled = False
        btnNext.Enabled = False
        btnPrevious.Enabled = False
        btnFirst.Enabled = False
        btnLast.Enabled = False
        ListGroupAll.Enabled = True
        ListStoreall.Enabled = True
        ListBoxesAll.Enabled = True
    End If

End Sub
Public Sub FillGridWithData()

    On Error GoTo ErrTrap

    Dim i As Integer
    Dim rs As ADODB.Recordset
    Dim My_SQL As String

    Set rs = New ADODB.Recordset
    My_SQL = "select * From TblUsers order by userid"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText

    With Me.Grid
        .Rows = 2
        .Clear flexClearScrollable
        If rs.RecordCount > 0 Then
            .Rows = rs.RecordCount + 1
            rs.MoveFirst
            For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("Ser")) = i
                .TextMatrix(i, .ColIndex("UserID")) = IIf(IsNull(rs.Fields("UserID").value), "", rs.Fields("UserID").value)
                .TextMatrix(i, .ColIndex("EmpID")) = IIf(IsNull(rs.Fields("EmpID").value), "", rs.Fields("EmpID").value)
                .TextMatrix(i, .ColIndex("EmpCode")) = get_EMPLOYEE_Data(val(.TextMatrix(i, .ColIndex("EmpID"))), "fullcode")
                .TextMatrix(i, .ColIndex("UserName")) = IIf(IsNull(rs.Fields("UserName").value), "", rs.Fields("UserName").value)
                .TextMatrix(i, .ColIndex("EmpName")) = IIf(IsNull(rs.Fields("EmpID").value), "", rs.Fields("EmpID").value)
                .TextMatrix(i, .ColIndex("BranchId")) = IIf(IsNull(rs.Fields("BranchId").value), "", rs.Fields("BranchId").value)
                .TextMatrix(i, .ColIndex("StoreID")) = IIf(IsNull(rs.Fields("StoreID").value), "", rs.Fields("StoreID").value)
                .TextMatrix(i, .ColIndex("BoxID")) = IIf(IsNull(rs.Fields("BoxID").value), "", rs.Fields("BoxID").value)
                .TextMatrix(i, .ColIndex("BankID")) = IIf(IsNull(rs.Fields("BankID").value), "", rs.Fields("BankID").value)
                rs.MoveNext
            Next i
            rs.Close
        End If
        .AutoSize 0, .Cols - 1, False
        .RowHeight(-1) = 300
    End With

ErrTrap:
End Sub
Private Sub ShowTip()
    On Error GoTo ErrTrap
    Dim TTP As New clstooltip
    Dim Wrap As String
    Dim Msg As String
    Wrap = CHR(13) + CHR(10)

    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "ÃœÌœ" & Wrap & "·› Õ ”Ã· ÃœÌœ " & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F12 √Ê Enter"
            
        .AddControl btnNew, Msg, True
    End With

    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = " ⁄œÌ·" & Wrap & "· ⁄œÌ·  ”Ã· «·Õ«·Ï " & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F11"
        .AddControl btnModify, Msg, True
    End With

    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Õ›Ÿ" & Wrap & "· ”ÃÌ· «·»Ì«‰«  œ«Œ· ﬁ«⁄œ… " & Wrap & "«·»Ì«‰«  ≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F10"
        .AddControl btnSave, Msg, True
    End With

    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = " —«Ã⁄" & Wrap & "·· —«Ã⁄ ⁄‰ «·⁄„·Ì… «·Õ«·Ì…" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F9"
        .AddControl BtnUndo, Msg, True
    End With

    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Õ–› «·”Ã·" & Wrap & "·Õ–› «·”Ã· «·Õ«·Ï" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F18"
        .AddControl btnDelete, Msg, True
    End With

    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Œ—ÊÃ" & Wrap & "·≈€·«ﬁ Â–Â «·‰«›–…" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " Ctrl+x"
        .AddControl btnCancel, Msg, True
    End With

    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«·«Ê·" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·«Ê·" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " Home √Ê UpArrow"
        .AddControl btnFirst, Msg, True
    End With

    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«·”«»ﬁ" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " PageUp √Ê LeftArrow"
        .AddControl btnPrevious, Msg, True
    End With

    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«· «·Ï" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ï" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " PageDown √Ê RightArrow"
        .AddControl btnNext, Msg, True
    End With

    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«·«ŒÌ—" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·«ŒÌ—" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " End √Ê DownArrow"
        .AddControl btnLast, Msg, True
    End With

ErrTrap:
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        If Me.TxtModFlg.Text = "R" Then
            btnNew_Click
        Else
            SendKeys "{TAB}"
        End If
    End If
    'New -------------------------
    If KeyCode = vbKeyF12 Then
        If btnNew.Enabled = False Then Exit Sub
        btnNew_Click
    End If
    'Edit ------------------------
    If KeyCode = vbKeyF11 Then
        If btnModify.Enabled = False Then Exit Sub
        btnModify_Click
    End If
    'save ------------------------
    If KeyCode = vbKeyF10 Then
        If btnSave.Enabled = False Then Exit Sub
        btnSave_Click
    End If
    'undo ------------------------------------------------------------------------------
    If KeyCode = vbKeyF9 Then
        If BtnUndo.Enabled = False Then Exit Sub
        BtnUndo_Click
    End If
    'Delete ---------------------------------------------------------------------------
    If KeyCode = vbKeyF8 Then
        If btnDelete.Enabled = False Then Exit Sub
        btnDelete_Click
    End If
    'Exit ----------------------------------------------------------------------
    If Shift = 2 Then
        If KeyCode = vbKeyX Then
            If btnCancel.Enabled = False Then Exit Sub
            BtnCancel_Click
        End If
    End If
    'Moveing through Records ---------------------------------------------------------------------------
    'If TxtModFlg.Text = "R" Then
    'Move first --------------------------------------------
    If KeyCode = vbKeyUp Or KeyCode = vbKeyHome Then
        If btnFirst.Enabled = False Then Exit Sub
        BtnFirst_Click
    End If
    'Move Previous---------------------------------------------------------
    If KeyCode = vbKeyLeft Or KeyCode = vbKeyPageUp Then
        If btnPrevious.Enabled = False Then Exit Sub
        BtnPrevious_Click
    End If
    'Move Next---------------------------------------------------------
    If KeyCode = vbKeyRight Or KeyCode = vbKeyPageDown Then
        If btnNext.Enabled = False Then Exit Sub
        BtnNext_Click
    End If
    'Move Last---------------------------------------------------------
    If KeyCode = vbKeyDown Or KeyCode = vbKeyEnd Then
        If btnLast.Enabled = False Then Exit Sub
        BtnLast_Click
    End If

    Exit Sub
ErrTrap:
End Sub
Private Function CheckDelCountry(Lngid As Long) As Boolean
    'Dim Rs As ADODB.Recordset
    'Dim StrSQL As String
    'StrSQL = "Select * From TblEmployee Where GovernmentID=" & Lngid & ""
    'Set Rs = New ADODB.Recordset
    'Rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    'If Not (Rs.BOF Or Rs.EOF) Then
    '    CheckDelCountry = False
    'Else
    '    CheckDelCountry = True
    'End If
    'Rs.Close
    'Set Rs = Nothing
End Function

Private Sub Label28_Click()
    If ListProductLineAll.ListIndex = -1 Then Exit Sub
    ListProductLineSelected.AddItem ListProductLineAll.List(ListProductLineAll.ListIndex)
    ListProductLineSelected.ItemData(ListProductLineSelected.NewIndex) = ListProductLineAll.ItemData(ListProductLineAll.ListIndex)
'    FG.Rows = ListProductLineSelected.ListCount + 1
'    FG.TextMatrix(FG.Rows - 1, FG.ColIndex("Name")) = ListProductLineAll.List(ListProductLineAll.ListIndex)
'    FG.TextMatrix(FG.Rows - 1, FG.ColIndex("ProductLineID")) = ListProductLineAll.ItemData(ListProductLineAll.ListIndex)
End Sub

Private Sub Label29_Click()
    Dim i As Integer
    ListProductLineSelected.Clear
'    FG.Rows = 1
'    FG.Rows = ListProductLineSelected.ListCount + 1
    For i = 0 To ListProductLineAll.ListCount - 1
        ListProductLineSelected.AddItem ListProductLineAll.List(i)
        ListProductLineSelected.ItemData(i) = ListProductLineAll.ItemData(i)
'        FG.TextMatrix(i + 1, FG.ColIndex("Name")) = ListProductLineAll.List(ListProductLineAll.ListIndex)
'        FG.TextMatrix(i + 1, FG.ColIndex("ProductLineID")) = ListProductLineAll.ItemData(ListProductLineAll.ListIndex)
        
    Next i

End Sub

Private Sub Label30_Click()
 ListProductLineSelected.Clear
' FG.Rows = 1
End Sub

Private Sub Label31_Click()
    If ListProductLineSelected.ListIndex > -1 Then
      ListProductLineSelected.RemoveItem ListProductLineSelected.ListIndex
        'FG.RemoveItem
    End If

End Sub





