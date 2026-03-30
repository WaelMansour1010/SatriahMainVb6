VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmExchangeRequest 
   BackColor       =   &H00E2E9E9&
   Caption         =   "   ÿ·» ’—› „ ⁄ÂœÌ‰   "
   ClientHeight    =   9285
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   17565
   Icon            =   "FrmExchangeRequest.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   9285
   ScaleWidth      =   17565
   WindowState     =   2  'Maximized
   Begin C1SizerLibCtl.C1Elastic Main_CLE 
      Height          =   9285
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   17565
      _cx             =   30983
      _cy             =   16378
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
      Begin C1SizerLibCtl.C1Elastic C1Elastic3 
         Height          =   1410
         Left            =   0
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   7785
         Width           =   17550
         _cx             =   30956
         _cy             =   2487
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
         Begin MSComDlg.CommonDialog cd 
            Left            =   480
            Top             =   480
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.CommandButton Command4 
            Caption         =   "«ŸÂ«— »«ﬁÏ «·œ›⁄« "
            CausesValidation=   0   'False
            Enabled         =   0   'False
            Height          =   285
            Left            =   4200
            RightToLeft     =   -1  'True
            TabIndex        =   65
            Top             =   285
            Visible         =   0   'False
            Width           =   1530
         End
         Begin VB.CommandButton Command3 
            Caption         =   " ’œÌ—«·Ï «·«ﬂ”Ì·"
            Height          =   285
            Left            =   5772
            RightToLeft     =   -1  'True
            TabIndex        =   51
            Top             =   285
            Width           =   1425
         End
         Begin VB.Frame Frame9 
            Caption         =   "»Ì«‰«  „Õ«”»Ì…"
            Height          =   690
            Left            =   7224
            RightToLeft     =   -1  'True
            TabIndex        =   42
            Top             =   0
            Width           =   8568
            Begin VB.CommandButton Command5 
               Caption         =   "Õ–› «·ﬁÌœ"
               Height          =   375
               Left            =   1560
               RightToLeft     =   -1  'True
               TabIndex        =   66
               Top             =   240
               Width           =   1335
            End
            Begin VB.CheckBox chkEntryCreated 
               Alignment       =   1  'Right Justify
               Caption         =   " „ «‰‘«¡ «·ﬁÌœ"
               Enabled         =   0   'False
               Height          =   252
               Left            =   3600
               RightToLeft     =   -1  'True
               TabIndex        =   64
               Top             =   120
               Visible         =   0   'False
               Width           =   1452
            End
            Begin VB.CommandButton Command8 
               Caption         =   "«‰‘«¡ «·ﬁÌœ"
               Enabled         =   0   'False
               Height          =   375
               Left            =   3960
               RightToLeft     =   -1  'True
               TabIndex        =   46
               Top             =   240
               Width           =   1335
            End
            Begin VB.CommandButton Command9 
               Caption         =   "ÿ»«⁄Â «·ﬁÌœ"
               Height          =   375
               Left            =   2880
               RightToLeft     =   -1  'True
               TabIndex        =   45
               Top             =   240
               Width           =   1095
            End
            Begin VB.TextBox TxtNoteSerial 
               Alignment       =   1  'Right Justify
               Height          =   405
               Left            =   5280
               RightToLeft     =   -1  'True
               TabIndex        =   44
               Top             =   240
               Width           =   2415
            End
            Begin VB.TextBox TxtNoteID 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   2760
               RightToLeft     =   -1  'True
               TabIndex        =   43
               Top             =   120
               Visible         =   0   'False
               Width           =   855
            End
            Begin MSComCtl2.DTPicker EntryDate 
               Height          =   300
               Left            =   120
               TabIndex        =   71
               TabStop         =   0   'False
               Top             =   360
               Width           =   1416
               _ExtentX        =   2487
               _ExtentY        =   529
               _Version        =   393216
               CalendarBackColor=   12648447
               CalendarTitleBackColor=   10383715
               CustomFormat    =   "yyyy/M/d"
               Format          =   114688003
               CurrentDate     =   37140
            End
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               Caption         =   " «—ÌŒ «‰‘«¡ «·ﬁÌœ"
               Height          =   252
               Left            =   360
               RightToLeft     =   -1  'True
               TabIndex        =   72
               Top             =   120
               Width           =   1092
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "—ﬁ„ «·ﬁÌœ"
               Height          =   195
               Index           =   35
               Left            =   7440
               RightToLeft     =   -1  'True
               TabIndex        =   47
               Top             =   240
               Width           =   990
            End
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   300
            Index           =   0
            Left            =   13608
            TabIndex        =   8
            Top             =   780
            Width           =   1392
            _ExtentX        =   2461
            _ExtentY        =   529
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
            Height          =   300
            Index           =   1
            Left            =   11940
            TabIndex        =   9
            Top             =   780
            Width           =   1632
            _ExtentX        =   2884
            _ExtentY        =   529
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
            Height          =   300
            Index           =   2
            Left            =   10740
            TabIndex        =   10
            Top             =   780
            Width           =   1176
            _ExtentX        =   2064
            _ExtentY        =   529
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
            Height          =   300
            Index           =   3
            Left            =   9696
            TabIndex        =   11
            Top             =   780
            Width           =   1116
            _ExtentX        =   1958
            _ExtentY        =   529
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
            Height          =   300
            Index           =   4
            Left            =   8328
            TabIndex        =   12
            Top             =   780
            Width           =   1308
            _ExtentX        =   2302
            _ExtentY        =   529
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
            Height          =   300
            Index           =   6
            Left            =   1785
            TabIndex        =   14
            Top             =   780
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   529
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
         Begin ImpulseButton.ISButton CmdAttach 
            Height          =   300
            Left            =   600
            TabIndex        =   15
            Top             =   780
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   529
            ButtonPositionImage=   1
            Caption         =   "«·„—›ﬁ« "
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
            Height          =   300
            Index           =   7
            Left            =   7080
            TabIndex        =   13
            Top             =   780
            Width           =   1212
            _ExtentX        =   2143
            _ExtentY        =   529
            ButtonPositionImage=   1
            Caption         =   "ÿ»«⁄Â"
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
            Height          =   300
            Index           =   5
            Left            =   3090
            TabIndex        =   48
            Top             =   780
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   529
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
            Height          =   372
            Index           =   8
            Left            =   0
            TabIndex        =   49
            Top             =   0
            Visible         =   0   'False
            Width           =   612
            _ExtentX        =   1085
            _ExtentY        =   661
            ButtonPositionImage=   1
            Caption         =   "≈·€«¡ «·œ›⁄« "
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
            Height          =   300
            Index           =   10
            Left            =   4230
            TabIndex        =   53
            Top             =   780
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   529
            ButtonPositionImage=   1
            Caption         =   "ÿ»«⁄Â ‰„Ê–Ã «·»‰ﬂ"
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
         Begin MSDataListLib.DataCombo DCboUserName 
            Height          =   312
            Left            =   0
            TabIndex        =   69
            Top             =   1068
            Width           =   2856
            _ExtentX        =   5027
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   300
            Index           =   11
            Left            =   5640
            TabIndex        =   74
            Top             =   780
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   529
            ButtonPositionImage=   1
            Caption         =   "ÿ»«⁄Â „Êﬁ› «·”œ«œ"
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
            Height          =   300
            Index           =   12
            Left            =   7080
            TabIndex        =   75
            Top             =   1080
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   529
            ButtonPositionImage=   1
            Caption         =   "ÿ»«⁄Â 2"
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
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "Õ—— »Ê«”ÿ… : "
            Height          =   288
            Index           =   12
            Left            =   3120
            RightToLeft     =   -1  'True
            TabIndex        =   70
            Top             =   1080
            Width           =   900
         End
         Begin VB.Label XPTxtCurrent 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   255
            Left            =   2145
            RightToLeft     =   -1  'True
            TabIndex        =   30
            Top             =   195
            Width           =   705
         End
         Begin VB.Label XPTxtCount 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   255
            Left            =   165
            RightToLeft     =   -1  'True
            TabIndex        =   29
            Top             =   210
            Width           =   765
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «·”Ã· «·Õ«·Ì:"
            Height          =   252
            Index           =   2
            Left            =   2892
            RightToLeft     =   -1  'True
            TabIndex        =   28
            Top             =   216
            Width           =   1224
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " ⁄œœ «·”Ã·« :"
            Height          =   255
            Index           =   4
            Left            =   975
            RightToLeft     =   -1  'True
            TabIndex        =   27
            Top             =   210
            Width           =   1095
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic2 
         Height          =   6270
         Left            =   0
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   1545
         Width           =   17550
         _cx             =   30956
         _cy             =   11060
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
         Begin VB.TextBox txtRemarks 
            Alignment       =   1  'Right Justify
            Height          =   492
            Left            =   840
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   67
            Top             =   120
            Width           =   15384
         End
         Begin VB.CheckBox chkChooseAll 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«Œ Ì«— «·ﬂ·"
            Height          =   420
            Left            =   16272
            RightToLeft     =   -1  'True
            TabIndex        =   40
            Top             =   720
            Width           =   1125
         End
         Begin VSFlex8Ctl.VSFlexGrid Grid 
            Height          =   5085
            Left            =   0
            TabIndex        =   6
            Top             =   1080
            Width           =   17505
            _cx             =   30877
            _cy             =   8969
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
            Cols            =   56
            FixedRows       =   2
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmExchangeRequest.frx":038A
            ScrollTrack     =   0   'False
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   0   'False
            AutoSizeMode    =   0
            AutoSearch      =   2
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
            WallPaperAlignment=   9
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   24
         End
         Begin MSComctlLib.ProgressBar ProgressBar1 
            Height          =   456
            Left            =   3288
            TabIndex        =   41
            Top             =   600
            Visible         =   0   'False
            Width           =   10908
            _ExtentX        =   19235
            _ExtentY        =   794
            _Version        =   393216
            Appearance      =   0
         End
         Begin VB.Label lblRow 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   372
            Left            =   720
            RightToLeft     =   -1  'True
            TabIndex        =   73
            Top             =   720
            Width           =   2052
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„·«ÕŸ« "
            Height          =   288
            Index           =   11
            Left            =   16572
            RightToLeft     =   -1  'True
            TabIndex        =   68
            Top             =   120
            Width           =   732
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   864
         Left            =   0
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   636
         Width           =   17544
         _cx             =   30956
         _cy             =   1535
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
         Begin VB.TextBox txtRecordno 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   5904
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   61
            Top             =   528
            Width           =   1680
         End
         Begin VB.TextBox txtfullcode 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   8832
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   60
            Top             =   528
            Width           =   1815
         End
         Begin VB.CommandButton Command1 
            Caption         =   "⁄—÷"
            Height          =   165
            Left            =   48
            RightToLeft     =   -1  'True
            TabIndex        =   59
            Top             =   -108
            Visible         =   0   'False
            Width           =   1035
         End
         Begin VB.CommandButton Command2 
            Caption         =   "⁄—÷"
            Enabled         =   0   'False
            Height          =   285
            Left            =   14808
            RightToLeft     =   -1  'True
            TabIndex        =   50
            Top             =   528
            Width           =   585
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   255
            Left            =   15444
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   39
            Top             =   528
            Width           =   735
         End
         Begin VB.ComboBox cbType 
            Height          =   288
            ItemData        =   "FrmExchangeRequest.frx":0B69
            Left            =   10356
            List            =   "FrmExchangeRequest.frx":0B6B
            RightToLeft     =   -1  'True
            TabIndex        =   3
            Top             =   -165
            Visible         =   0   'False
            Width           =   1635
         End
         Begin VB.TextBox txtCode 
            Alignment       =   1  'Right Justify
            Height          =   276
            Left            =   12060
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   2
            Top             =   240
            Width           =   1515
         End
         Begin VB.TextBox txtID 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   252
            Left            =   14808
            Locked          =   -1  'True
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   1
            Top             =   210
            Width           =   1395
         End
         Begin MSDataListLib.DataCombo DcDur 
            Height          =   288
            Left            =   3312
            TabIndex        =   4
            Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
            Top             =   276
            Width           =   1572
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo dcMontth 
            Height          =   312
            Left            =   3312
            TabIndex        =   5
            Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
            Top             =   528
            Width           =   1572
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSComCtl2.DTPicker Date 
            Height          =   300
            Left            =   240
            TabIndex        =   33
            TabStop         =   0   'False
            Top             =   216
            Width           =   1776
            _ExtentX        =   3122
            _ExtentY        =   529
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CustomFormat    =   "yyyy/M/d"
            Format          =   114688003
            CurrentDate     =   37140
         End
         Begin Dynamic_Byte.NourHijriCal DateH 
            Height          =   300
            Left            =   240
            TabIndex        =   34
            Top             =   528
            Width           =   1776
            _ExtentX        =   3122
            _ExtentY        =   529
         End
         Begin MSDataListLib.DataCombo dcBranch 
            Height          =   288
            Left            =   8832
            TabIndex        =   36
            Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
            Top             =   216
            Width           =   1812
            _ExtentX        =   3201
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo dcMangerialAreaID 
            Height          =   288
            Left            =   5904
            TabIndex        =   54
            Top             =   216
            Width           =   1680
            _ExtentX        =   2963
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo dcCustomer 
            Height          =   312
            Left            =   12060
            TabIndex        =   55
            Top             =   528
            Width           =   1512
            _ExtentX        =   2672
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·”Ã· "
            Height          =   288
            Index           =   21
            Left            =   7800
            RightToLeft     =   -1  'True
            TabIndex        =   63
            Top             =   528
            Width           =   888
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·ﬂÊœ"
            Height          =   288
            Index           =   22
            Left            =   10968
            RightToLeft     =   -1  'True
            TabIndex        =   62
            Top             =   528
            Width           =   612
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   " «—ÌŒ «·”‰œ Â‹ "
            ForeColor       =   &H00000000&
            Height          =   228
            Left            =   2136
            TabIndex        =   58
            Top             =   528
            Width           =   972
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·«œ«—… «· ⁄·Ì„Ì…"
            Height          =   288
            Index           =   10
            Left            =   7668
            RightToLeft     =   -1  'True
            TabIndex        =   57
            Top             =   216
            Width           =   1080
            WordWrap        =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„ ⁄Âœ"
            Height          =   288
            Index           =   7
            Left            =   13956
            RightToLeft     =   -1  'True
            TabIndex        =   56
            Top             =   528
            Width           =   528
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "»‰«¡ ⁄·Ï «” Õﬁ«ﬁ"
            Height          =   276
            Index           =   6
            Left            =   16188
            RightToLeft     =   -1  'True
            TabIndex        =   38
            Top             =   528
            Width           =   1272
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·›—⁄"
            Height          =   240
            Index           =   5
            Left            =   10992
            RightToLeft     =   -1  'True
            TabIndex        =   37
            Top             =   216
            Width           =   600
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   " «—ÌŒ «·”‰œ „"
            ForeColor       =   &H00000000&
            Height          =   228
            Left            =   2136
            TabIndex        =   35
            Top             =   216
            Width           =   972
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·› —…"
            Height          =   240
            Index           =   1
            Left            =   5148
            RightToLeft     =   -1  'True
            TabIndex        =   32
            Top             =   528
            Width           =   696
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”‰… «·œ—«”Ì…"
            Height          =   228
            Index           =   3
            Left            =   4188
            RightToLeft     =   -1  'True
            TabIndex        =   31
            Top             =   276
            Width           =   1620
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·—ﬁ„ «·ÌœÊÏ"
            Height          =   288
            Index           =   9
            Left            =   12840
            RightToLeft     =   -1  'True
            TabIndex        =   25
            Top             =   240
            Width           =   1608
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·”‰œ"
            Height          =   288
            Index           =   8
            Left            =   16020
            RightToLeft     =   -1  'True
            TabIndex        =   24
            Top             =   240
            Width           =   1272
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·’—›"
            Height          =   228
            Index           =   0
            Left            =   11916
            RightToLeft     =   -1  'True
            TabIndex        =   23
            Top             =   -168
            Visible         =   0   'False
            Width           =   1332
         End
      End
      Begin C1SizerLibCtl.C1Elastic EleHeader 
         Height          =   735
         Left            =   -120
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   0
         Width           =   17595
         _cx             =   31036
         _cy             =   1296
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
         Appearance      =   0
         MousePointer    =   0
         Version         =   801
         BackColor       =   16777215
         ForeColor       =   4210688
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   "     ÿ·» ’—› „ ⁄ÂœÌ‰   "
         Align           =   0
         AutoSizeChildren=   0
         BorderWidth     =   2
         ChildSpacing    =   1
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
         CaptionStyle    =   1
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
            Left            =   2250
            RightToLeft     =   -1  'True
            TabIndex        =   17
            Top             =   180
            Visible         =   0   'False
            Width           =   855
         End
         Begin ImpulseButton.ISButton XPBtnMove 
            Height          =   345
            Index           =   0
            Left            =   1155
            TabIndex        =   18
            Top             =   120
            Width           =   495
            _ExtentX        =   873
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
            ButtonImage     =   "FrmExchangeRequest.frx":0B6D
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
            TabIndex        =   19
            Top             =   120
            Width           =   495
            _ExtentX        =   873
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
            ButtonImage     =   "FrmExchangeRequest.frx":0F07
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
            Left            =   1680
            TabIndex        =   20
            Top             =   120
            Width           =   495
            _ExtentX        =   873
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
            ButtonImage     =   "FrmExchangeRequest.frx":12A1
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
            Left            =   615
            TabIndex        =   21
            Top             =   120
            Width           =   495
            _ExtentX        =   873
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
            ButtonImage     =   "FrmExchangeRequest.frx":163B
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
   Begin ImpulseButton.ISButton Cmd 
      Height          =   396
      Index           =   9
      Left            =   0
      TabIndex        =   52
      Top             =   0
      Width           =   1332
      _ExtentX        =   2355
      _ExtentY        =   688
      ButtonPositionImage=   1
      Caption         =   "ÿ»«⁄Â"
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
Attribute VB_Name = "FrmExchangeRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim RsTemp As ADODB.Recordset
Dim RsTemp2 As ADODB.Recordset
Dim Rs_Temp As ADODB.Recordset
Dim RsTemp3 As ADODB.Recordset
Dim RsTemp4 As ADODB.Recordset
Dim RsExcep As ADODB.Recordset
Dim Account_Code_dynamic As String
Dim Account_Code_dynamic1 As String
Dim RsT As ADODB.Recordset
Dim TTP As clstooltip



Public Function CREATE_VOUCHER_GE(general_noteid As Long, BranchID As Integer, UserID As Long _
, NoteDate As Date)
Dim BasicSalaryAccount As String
Dim StrSQL As String
         StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & general_noteid
        Cn.Execute StrSQL, , adExecuteNoRecords

    Dim i As Integer
    Dim LngDevID As Long
    Dim Msg As String
   
    Dim Employee_account As String
    Dim StrAccountCode As String
    Dim X As Integer
    Dim rs As New ADODB.Recordset
    Dim notes_serial As String
    Dim notes_id As String
    Dim j As Integer
    Dim ColumnName As String
    Dim SalaryAccount As String
    Dim BonusAccount As String
    Dim DiscountAccount As String
    
    ' Msg = EleHeader.Caption & " ??? " & txtID & " ??????" & Date
    'Dim msg As String
    
    Msg = EleHeader.Caption & CHR(13) & " »—ﬁ„ : " & txtid.Text & CHR(13) & "  ··„‰ÿﬁ…  " & dcBranch.Text & CHR(13) & " ··”‰… " & DcDur.Text & CHR(13) & " ··› —…  " & dcMontth.Text
    notes_id = general_noteid

  
   
    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")

    Dim line_no As Integer
    line_no = 1
                
    'C???? C??I?? C?C?C?CE
     
    Dim CValue As Double
    Dim Branch As Integer
    Dim ProjectID As Integer
    
    BranchID = 1
    
    With Grid


line_no = 1
        For i = .FixedRows To .Rows - 1
    BranchID = val(dcBranch.BoundText)
       If .TextMatrix(i, .ColIndex("net")) >= 0 Then
            If .TextMatrix(i, .ColIndex("Value")) > 0 And Account_Code_dynamic1 <> "" And .Cell(flexcpChecked, i, .ColIndex("Status")) = flexChecked Then     'C?C??? C???E??E IC??
                'Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code1") 'C?C??? C???E??E
         '       StrAccountCode = .TextMatrix(i, .ColIndex("Account_Code"))
        
                If ModAccounts.AddNewDev(LngDevID, line_no, Account_Code_dynamic1, .TextMatrix(i, .ColIndex("Value")), 0, Msg & CHR(13) & "  ﬁÌ„… «·œ›⁄Â  «·„” Õﬁ…  ··„ ⁄Âœ " & .TextMatrix(i, .ColIndex("cusname")) & CHR(13) & " ··”Ì«—…" & .TextMatrix(i, .ColIndex("Car")), val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                    GoTo ErrTrap
                End If

                line_no = line_no + 1
                  
            End If
    PercentgValueAddedAccount_Transec Me.Date.value, 47, 0, StrAccountCode
            
            If .TextMatrix(i, .ColIndex("Vat2")) > 0 And StrAccountCode <> "" And .Cell(flexcpChecked, i, .ColIndex("Status")) = flexChecked Then   'C?C??? C???E??E IC??
                'Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code1") 'C?C??? C???E??E
         '       StrAccountCode = .TextMatrix(i, .ColIndex("Account_Code"))
        
                If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCode, .TextMatrix(i, .ColIndex("Vat2")), 0, Msg & "  Õ”„Ì«      ··„ ⁄Âœ " & .TextMatrix(i, .ColIndex("cusname")) & " ··”Ì«—…" & .TextMatrix(i, .ColIndex("Car")), val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                    GoTo ErrTrap
                End If

                line_no = line_no + 1
                  
            End If
            
       If .TextMatrix(i, .ColIndex("net")) >= 0 And .TextMatrix(i, .ColIndex("Account_Code")) <> "" And .Cell(flexcpChecked, i, .ColIndex("Status")) = flexChecked Then   'C?C??? C???E??E IC??
                'Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code1") 'C?C??? C???E??E
             StrAccountCode = .TextMatrix(i, .ColIndex("Account_Code"))
        
                If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCode, .TextMatrix(i, .ColIndex("net")), 1, Msg & CHR(13) & "     ’«›Ì   ··„ ⁄Âœ " & .TextMatrix(i, .ColIndex("cusname")) & CHR(13) & " ··”Ì«—…" & .TextMatrix(i, .ColIndex("Car")), val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                    GoTo ErrTrap
                End If

                line_no = line_no + 1
                  
            End If
           
                        
  
            
       If .TextMatrix(i, .ColIndex("Total")) > 0 And Account_Code_dynamic <> "" And .Cell(flexcpChecked, i, .ColIndex("Status")) = flexChecked Then   'C?C??? C???E??E IC??
                'Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code1") 'C?C??? C???E??E
         '       StrAccountCode = .TextMatrix(i, .ColIndex("Account_Code"))
        
                If ModAccounts.AddNewDev(LngDevID, line_no, Account_Code_dynamic, .TextMatrix(i, .ColIndex("Total")), 1, Msg & "  Õ”„Ì«      ··„ ⁄Âœ " & .TextMatrix(i, .ColIndex("cusname")) & " ··”Ì«—…" & .TextMatrix(i, .ColIndex("Car")), val(notes_id), , , , NoteDate, user_id, , , , , , , , , setfoxy_Line, , , , , , , , , BranchID) = False Then
                    GoTo ErrTrap
                End If

                line_no = line_no + 1
                  
            End If
                        
     End If
     Next i
     
     End With
           
   updateNotesValueAndNobytext (val(notes_id))
    Exit Function
ErrTrap:
   End Function


   Function createVoucher()
Dim NoteID As Long
Dim NoteDate As Date
Dim NoteSerial As String
Dim Notevalue As Double
Dim notytype As Integer
Dim des As String
des = EleHeader.Caption & " ··”‰… " & DcDur.Text & " ··› —…  " & dcMontth.Text
Dim tablename As String
Dim Filedname As String
Dim NoteSerial1 As Long
Dim BranchID As Integer
 

Dim sql As String
tablename = "TblExchangeRequest"
Filedname = "ID"
NoteSerial1 = val(txtid)
Notevalue = 0

 notytype = 8069
'Notevalue = val(total)
 

 BranchID = val(dcBranch.BoundText)

'NoteDate = Me.Date.value
NoteDate = Me.EntryDate.value

 
'If Notevalue > 0 Then
                                If Me.TxtModFlg = "N" Then
                                      CreateNotes NoteID, NoteDate, BranchID, notytype, Notevalue, NoteSerial, (NoteSerial1), tablename, Filedname, NoteSerial1, des         ', recordDateH.value
                                              TxtNoteID.Text = NoteID
                                                     TxtNoteSerial.Text = NoteSerial
                                     Else
                                                 If TxtNoteID.Text = "" Or TxtNoteSerial.Text = "" Then
                                            CreateNotes NoteID, NoteDate, BranchID, notytype, Notevalue, NoteSerial, (NoteSerial1), tablename, Filedname, NoteSerial1, des   ', recordDateH.value
                                                                 TxtNoteID.Text = NoteID
                                                                TxtNoteSerial.Text = NoteSerial
                                                   Else
                                                                 sql = "update notes  set Note_Value=" & Notevalue & ",note_value_by_characters='" & WriteNo(val(Notevalue), 0, True) & "'"
                                                                sql = sql & ",NoteSerial1='" & (NoteSerial1) & "'"
                                                                   sql = sql & " where NoteID=" & val(TxtNoteID.Text)
                                                                   Cn.Execute sql
                                                               
                                                 End If
                                       
                                End If

CREATE_VOUCHER_GE val(TxtNoteID.Text), BranchID, user_id, NoteDate
Dim str As String
 str = " update TblExchangeRequest set EntryCreated =1 "
 Cn.Execute str
 chkEntryCreated.value = vbChecked
rs.Resync adAffectCurrent
 

'     End If

End Function

Private Sub chkChooseAll_Click()

  If Me.TxtModFlg.Text = "E" Or Me.TxtModFlg.Text = "N" Then
            Else
                  Exit Sub
  End If
            
Dim i As Integer

For i = 1 To Grid.Rows - 1
    If Grid.TextMatrix(i, Grid.ColIndex("fullcode")) <> "" Then
            If chkChooseAll.value = 1 Then
                    Grid.TextMatrix(i, Grid.ColIndex("Status")) = 1
            Else
                    Grid.TextMatrix(i, Grid.ColIndex("Status")) = 0
            End If
    End If
Next
End Sub

Private Sub Cmd_Click(Index As Integer)
 '    On Error GoTo ErrTrap
  Me.DCboUserName.BoundText = user_id
  
    Select Case Index
        Case 0
            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If
            TxtModFlg.Text = "N"
            clear_all Me
            txtid.Text = CStr(new_id("TblExchangeRequest", "ID", "", True))
          '  TXTid.SetFocus
             Grid.Rows = Grid.FixedRows
             Command2.Enabled = True
        Case 1
        
        
                If ChekClodePeriod(Me.Date.value) = True Then
                If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
                Else
                MsgBox "Please Change Date Becouse This is Period is Closed"
                End If
                Exit Sub
                End If
              
              
               If Me.chkEntryCreated = vbChecked Then
                 MsgBox ("»—Ã«¡ «·€«¡ «·ﬁÌœ  «·„”Ã· «Ê·« ⁄·Ï Â–« «·ÿ·»")
                    Exit Sub
            End If

            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If
                        
            If IsRecordLocked(val(txtid.Text)) = True Then
                    MsgBox ("»—Ã«¡ «·€«¡ «·”œ«œ «·„”Ã· «Ê·« ⁄·Ï Â–« «·ÿ·»")
                    Exit Sub
            Else
                   ' Check Current User
                     Dim StrSQL As String
                     Dim mss As String
                     StrSQL = "SELECT  *  From TblExchangeRequest  where  ID = " & val(txtid.Text)
                     Set RsT = New ADODB.Recordset
                     RsT.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                     If RsT.RecordCount > 0 Then
                            If Not IsNull(RsT("CurrentUser").value) Then
                                        mss = " Â‰«ﬂ „” Œœ„  "
                                        mss = mss + user_name + CHR(13) + " ›—⁄ "
                                        mss = mss + CurrentBranchName + CHR(13)
                                        mss = mss + " Ì⁄œ· ⁄·Ï ‰›” «·”Ã· "
                                        MsgBox (mss)
                                        Exit Sub
                            Else
                                RsT("CurrentUser") = user_id
                                RsT.update
                            End If
                     End If
                     rs.Resync
                     '/////////////////
                                  
                     TxtModFlg.Text = "E"
                     C1Elastic1.Enabled = False
                     Command4.Enabled = True
                     Command4_Click
                     
            End If
                       
            
        Case 2
 
                                                       If ChekClodePeriod(Me.Date.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
              
              
            SaveData
 Me.Retrive val(txtid)
 
        Case 3
            Undo
'C1Elastic2.Enabled = False
                LogTextA = "  —«Ã⁄ "
                LogTexte = " Undo "
                AddToLogFile CInt(user_id), 8069, Now, Time, LogTextA, LogTexte, Me.Name, "E", "", "", val(TxtNoteSerial), txtid
                
                Clear_Current_User
                
        Case 4
                                             If ChekClodePeriod(Me.Date.value) = True Then
               If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
               Else
               MsgBox "Please Change Date Becouse This is Period is Closed"
              End If
              Exit Sub
              End If
              
              
            If DoPremis(Do_Delete, Me.Name, True) = False Then
                Exit Sub
            End If
                
                 If IsRecordLocked(val(txtid.Text)) = True Then
                    MsgBox ("»—Ã«¡ «·€«¡ «·”œ«œ «·„”Ã· «Ê·« ⁄·Ï Â–« «·ÿ·»")
                       
                Else
                    Del_Company
                End If
                
        Case 5
        
        
                       Unload FrmSearch_Request
                       FrmSearch_Request.SendForm = "ER_ER"
                       FrmSearch_Request.show

        Case 6
            Unload Me
         Case 12
            print_report2
         Case 7
         If chkEntryCreated.value = vbChecked Then
                    print_report
          Else
          MsgBox "·« Ì„ﬂ‰ «·ÿ»«⁄Â «·« »⁄œ «‰‘«¡ «·ﬁÌœ Ì„ﬂ‰  ’œÌ— «·»Ì«‰« ", vbInformation
          End If
                    LogTextA = " ÿ»«⁄… ÿ·» «·’—› "
                    LogTexte = " Print Exchange Request"
                    AddToLogFile CInt(user_id), 8069, Now, Time, LogTextA, LogTexte, Me.Name, "E", "", "", val(TxtNoteSerial), txtid
                    
            Case 8
               Dim Msg As String
               If SystemOptions.UserInterface = ArabicInterface Then
                        Msg = " ”Ì „ Õ–› «·„œ›Ê⁄«  ··”‰œ ° Â·  —Ìœ «” ﬂ„«· ⁄„·Ì… «·Õ–› ø "
               Else
                        Msg = " This action will delete Paid for this receipt "
               End If
               If MsgBox(Msg, vbOKCancel) = vbOK Then
                    ' Cancel_Paid
                    ' Command1_Click
                    
                    Command4_Click
                    UnCheckALL
                    C1Elastic1.Enabled = True
                    C1Elastic2.Enabled = True
                    
                    LogTextA = "  ≈·€«¡ «·œ›⁄«  "
                    LogTexte = " Cancel Paid"
                    AddToLogFile CInt(user_id), 8069, Now, Time, LogTextA, LogTexte, Me.Name, "E", "", "", val(TxtNoteSerial), txtid
                    
               End If
                  Command2.Enabled = True
                  
                  Case 10
                         If chkEntryCreated.value = vbChecked Then
                    print_form
          Else
          MsgBox "·« Ì„ﬂ‰ «·ÿ»«⁄Â «·« »⁄œ «‰‘«¡ «·ﬁÌœ Ì„ﬂ‰  ’œÌ— «·»Ì«‰« ", vbInformation
          End If
          
                   
                     
                    LogTextA = "  ÿ»«⁄… ‰„Ê–Ã «·»‰ﬂ  "
                    LogTexte = " Print Bank Form"
                    AddToLogFile CInt(user_id), 8069, Now, Time, LogTextA, LogTexte, Me.Name, "E", "", "", val(TxtNoteSerial), txtid
                    
                   Case 11
                   print_form1
                   
                             
                    LogTextA = "  ÿ»«⁄… „Êﬁ› «·’—›  "
                    LogTexte = " Print Payed Status"
                    AddToLogFile CInt(user_id), 8069, Now, Time, LogTextA, LogTexte, Me.Name, "E", "", "", val(TxtNoteSerial), txtid
                    
                    
    End Select

    Exit Sub
ErrTrap:
End Sub


Private Sub UnCheckALL()
Dim i As Integer
With Grid
For i = Grid.FixedRows To Grid.Rows - 1
.TextMatrix(i, .ColIndex("Status")) = 0
Next
End With
End Sub

Private Function IsRecordLocked(ID As Integer) As Boolean
Dim str As String, Check   As Boolean

Check = False
Set Rs_Temp = New ADODB.Recordset
str = " select * from TblAttributionInstallmentDivided where reid = " & ID & " and paymentpayed = 1 "
Rs_Temp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText

If Rs_Temp.RecordCount > 0 Then
        Check = True
        IsRecordLocked = True
        Exit Function
End If
IsRecordLocked = False

End Function



Private Sub Cancel_Paid()

Dim str As String, AllID As String, i As Integer
If rs.RecordCount > 0 Then
        AllID = IIf(IsNull(rs("AllID").value), "", rs("AllID").value)
        rs("AllID") = Null
        rs.update
End If

If AllID = "" Then
    Exit Sub
End If

str = " select * from TblAttributionInstallmentDivided where   ID in (  " & AllID & "  )"
Set Rs_Temp = New ADODB.Recordset
Rs_Temp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs_Temp.RecordCount > 0 Then
        For i = 0 To Rs_Temp.RecordCount - 1
                Rs_Temp("RE_Paid") = Null
                Rs_Temp("REID") = Null
                Rs_Temp.update
                Rs_Temp.MoveNext
        Next
End If



End Sub




Public Function ISPaid() As Boolean

Dim str As String
str = "   select * from TblExchangeReques_Detailst where   HID =  " & val(txtid.Text) & " and  TblExchangeReques_Detailst.paid  =1    "
Set Rs_Temp = New ADODB.Recordset
Rs_Temp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText

If Rs_Temp.RecordCount > 0 Then
        ISPaid = True
        Exit Function
End If
ISPaid = False

End Function


Private Sub CmdAttach_Click()
            On Error Resume Next
'ShowAttachments XPTxtBoxID, "0701201405"
 

End Sub

Private Sub CmdHelp_Click()
    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
    SystemOptions.SysHelp.HHDisplayTopicID Me.hWnd
End Sub

Private Sub Command2_Click()
Retrive_Data
End Sub

Private Sub Command3_Click()
  On Error Resume Next
    Dim StrFileName As String
    'StrFileName = CurDir & "\" & "\Report1.xls"

    If Dir(StrFileName) <> "" Then
        Kill StrFileName
    End If
'Grid.RightToLeft = True
 
    On Error Resume Next
      cd.CancelError = True 'allow escape key/cancel
     cd.filename = "Report1"
    cd.ShowSave     'show the dialog screen
    If Err <> 32755 Then    ' User didn't chose Cancel.
   Else
       Exit Sub
    End If
 StrFileName = cd.filename & ".xls"
Me.Grid.saveGrid StrFileName, flexFileCustomText, True
   
    OpenFile StrFileName
    
End Sub

Private Sub Command4_Click()
ProgressBar1.Visible = True
Fill_Grid
ProgressBar1.Visible = False
ProgressBar1.value = 0
 Command4.Enabled = False
End Sub

Private Sub Command5_Click()

        If DoPremis(Do_Edit, Me.Name, True) = False Then
            Exit Sub
        End If
        
        
        If IsRecordLocked(val(txtid.Text)) = True Then
                MsgBox ("»—Ã«¡ «·€«¡ «·”œ«œ «·„”Ã· «Ê·« ⁄·Ï Â–« «·ÿ·»")
                Exit Sub
        Else
              '   TxtModFlg.text = "E"
                 C1Elastic1.Enabled = False
                 Command4.Enabled = True
        End If
                       
                       
        
        If ChekClodePeriod(Me.Date.value) = True Then
            If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
            Else
            MsgBox "Please Change Date Becouse This is Period is Closed"
            End If
            Exit Sub
        End If
        
        
        
        
        Dim StrSQL As String
        
        StrSQL = "Delete From notes Where NoteID=" & val(Me.TxtNoteID.Text)
        Cn.Execute StrSQL, , adExecuteNoRecords
        
        
        StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(Me.TxtNoteID.Text)
        Cn.Execute StrSQL, , adExecuteNoRecords
                           
                          
                          
        StrSQL = "update TblExchangeRequest set EntryCreated =null,    NoteID=null,NoteSerial=null where  id =  " & val(txtid.Text)
        Cn.Execute StrSQL, , adExecuteNoRecords
        
        
        rs.Resync adAffectCurrent
        chkEntryCreated.value = vbUnchecked
        TxtNoteSerial.Text = ""
        Me.TxtNoteID.Text = ""
        
        MsgBox " „ Õ–› «·ﬁÌœ"
        LogTextA = "  Õ–› «·ﬁÌœ "
        LogTexte = " Delete Entry "
         AddToLogFile CInt(user_id), 8069, Now, Time, LogTextA, LogTexte, Me.Name, "E", "", "", val(TxtNoteSerial), txtid
    
    
End Sub

Private Sub Command8_Click()
           If ChekClodePeriod(Me.Date.value) = True Then
                If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ  €ÌÌ—  «—ÌŒ «·Õ—ﬂ… ·«‰ Â–Â «·› —… „€·ﬁ…"
                Else
                MsgBox "Please Change Date Becouse This is Period is Closed"
                End If
                Exit Sub
                End If
                
 Me.Retrive val(txtid)
                  Account_Code_dynamic = get_account_code_branch(106, my_branch)
        
        If Account_Code_dynamic = "NO branch" Then
            MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
            Exit Sub
        ElseIf Account_Code_dynamic = "NO account" Then
            MsgBox "·„ Ì „  ÕœÌœ  ﬂ·›… «·‰ﬁ· ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
            Exit Sub
                
        End If
     
     
                       Account_Code_dynamic1 = get_account_code_branch(107, my_branch)
        
        If Account_Code_dynamic = "NO branch" Then
            MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
            Exit Sub
        ElseIf Account_Code_dynamic = "NO account" Then
            MsgBox "·„ Ì „  ÕœÌœ  œ›⁄«  „ ⁄ÂœÌ‰ „” Õﬁ…   ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
            Exit Sub
                
        End If
        
        



If TxtNoteSerial <> "" Then
MsgBox "«·ﬁÌœ „‰‘√ »«·›⁄· ﬁ„ »Õ–›… «Ê·«"
Exit Sub
End If

Cmd_Click (1)
Cmd_Click (2)

createVoucher

    If SystemOptions.UserInterface = ArabicInterface Then

        MsgBox " „ «‰‘«¡ «·ﬁÌœ   "
chkEntryCreated.value = vbChecked

    Else
        MsgBox "Ge Created"
        chkEntryCreated.value = vbChecked
        
    End If
    
      LogTextA = "  √‰‘«¡  «·ﬁÌœ "
    LogTexte = " Create  GE "
      AddToLogFile CInt(user_id), 8069, Now, Time, LogTextA, LogTexte, Me.Name, "E", "", "", val(TxtNoteSerial), txtid
    
End Sub

Private Sub Command9_Click()
ShowGL_cc Me.TxtNoteSerial.Text, , 200

     LogTextA = "  ÿ»«⁄…  «·ﬁÌœ "
    LogTexte = " Create  GE "
    AddToLogFile CInt(user_id), 8069, Now, Time, LogTextA, LogTexte, Me.Name, "E", "", "", val(TxtNoteSerial), txtid
End Sub

Private Sub Date_Change()
If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
TxtNoteSerial.Text = ""
calculation False
End If
End Sub

Private Sub Dcbranch_Change()
If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
TxtNoteSerial.Text = ""
End If


End Sub

Private Sub Dcbranch_Click(Area As Integer)
Dcbranch_Change
End Sub

Private Sub dcBranch_KeyUp(KeyCode As Integer, _
                           Shift As Integer)

 
End Sub

Private Sub DCEmP_KeyUp(KeyCode As Integer, _
                        Shift As Integer)

 

End Sub

Private Sub Option1_Click()
 
End Sub

Private Sub Option2_Click()
 
End Sub

 
Private Sub Command1_Click()
Grid.Rows = Grid.FixedRows
ProgressBar1.Visible = True
Fill_Grid
ProgressBar1.Visible = False
ProgressBar1.value = 0

End Sub

Private Function check_reg() As Boolean

Dim str As String
Grid.Rows = Grid.FixedRows

str = " select * from tblexchangerequest where durationid = " & val(DcDur.BoundText) & "  and Month =   " & val(dcMontth.BoundText) & "  and BranchID = " & val(dcBranch.BoundText)
Set RsTemp = New ADODB.Recordset
RsTemp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
If RsTemp.RecordCount > 0 Then
        MsgBox (" „  ”ÃÌ· ÿ·» ’—› ·Â–Â «·› —… „‰ ﬁ»· ")
        check_reg = True
Else

        check_reg = False
End If
End Function


Private Sub dcCustomer_Change()
Dim val1, val2, recordno As String, Fullcode As String
txtRecordno.Text = ""
txtfullcode.Text = ""

If dcCustomer.BoundText = "" Then Exit Sub
Dim str As String
    str = " select * From TblCustemers where Type=2  and cusid = " & val(dcCustomer.BoundText)
    Set Rs_Temp = New ADODB.Recordset
    Rs_Temp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If Rs_Temp.RecordCount > 0 Then
        Rs_Temp.MoveFirst '
        recordno = IIf(IsNull(Rs_Temp("recordno").value), "", Rs_Temp("recordno").value)
        Fullcode = IIf(IsNull(Rs_Temp("fullcode").value), "", Rs_Temp("fullcode").value)
     End If
     txtRecordno.Text = recordno
     txtfullcode.Text = Fullcode
    

End Sub

Private Sub dcCustomer_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
        Unload FrmCompanySearch
        FrmCompanySearch.lblSearchtype = "20160211"
        FrmCompanySearch.show vbModal
End If

End Sub

Private Sub DcDur_Change()
Dim i As Integer, j As Integer, str As String
    i = val(DcDur.BoundText)
    
    If i > 0 Then
        str = "  select id , Name  from TblDurations_Details where did =   " & i
        fill_combo dcMontth, str
    Else
        str = "  select id , Name  from TblDurations_Details where did =   " & -1
        fill_combo dcMontth, str
    End If

End Sub

Private Sub Fill_Grid()
On Error Resume Next
Dim i As Integer, j As Integer, str As String
    i = val(DcDur.BoundText)
   
  '  Grid.Rows = Grid.FixedRows
 
'str = str & "   SELECT TblAttributionInstallmentDivided.DDEmbarkDate , TblAttributionInstallmentDivided.DDEmbarkDateH   , dbo.TblCustemers.RecordNo, dbo.TblAttributionContract.IDAC, dbo.TblAttributionContract.DurationID, dbo.TblDurations.Name AS DurationName,"
'str = str & "   dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode, dbo.TblCustemers.CusID, dbo.TblAttributionContract.StartContractDate,"
'str = str & "   dbo.TblAttributionContract.EndContractDate, dbo.TblAttributionContract.FromDate, dbo.TblDurations.FromDate AS DurFromDate, dbo.TblDurations.FromDateH AS DurFromDateH,  dbo.TblDurations.ToDate AS DurToDate,"
'str = str & "   dbo.TblVehicleAllocation_Details.Type, dbo.TblVehicleAllocation_Details.StudentCount, dbo.TblVehicleAllocation_Details.BoardNo, dbo.TblVehicleAllocation_Details.rate,"
'str = str & "   dbo.TblVehicleAllocation_Details.Custom, dbo.TblVehicleAllocation_Details.DayRate, dbo.TblVehicleAllocation_Details.StudentCustom,"
'str = str & "   dbo.TblVehicleAllocation_Details.CarID, dbo.TblDurations_Details.Name, dbo.TblDurations_Details.ID AS MonthID, dbo.TblVehicleAllocation_Details.ID,"
'str = str & "   dbo.TblVehicleAllocation_Details.SchoolFileID, dbo.TblVendorCars.StopDeal, dbo.TblVendorCars.StopDate, dbo.TblVendorCars.StopDateH,"
'str = str & "   dbo.TblAttributionContract.BranchID, dbo.TblCustemers.Account_Code, dbo.ACCOUNTS.Account_Serial, dbo.TblCustemers.IBAN, dbo.TblCustemers.BankAccount,"
'str = str & "   dbo.TblDurations.Type AS DurationType, dbo.TblAttributionInstallmentDivided.ID AS DividID, dbo.TblAttributionContract.MangerialAreaID,"
'str = str & "   dbo.TblManagerialArea.Name AS MAName , TblAttributionContract.cityID , TblAttributionContract.vendorid , dbo.TblVehicleAllocation_Details.schoolfileID "
'str = str & "   FROM     dbo.TblVendorCars RIGHT OUTER JOIN"
'str = str & "   dbo.TblDurations_Details INNER JOIN"
'str = str & "   dbo.TblAttributionContract INNER JOIN"
'str = str & "   dbo.TblVehicleAllocation_Details ON dbo.TblAttributionContract.IDAC = dbo.TblVehicleAllocation_Details.IDVA INNER JOIN"
'str = str & "   dbo.TblAttributionInstallmentDivided ON dbo.TblVehicleAllocation_Details.ID = dbo.TblAttributionInstallmentDivided.DetailsID ON"
'str = str & "   dbo.TblDurations_Details.ID = dbo.TblAttributionInstallmentDivided.MonthID LEFT OUTER JOIN"
'str = str & "   dbo.TblManagerialArea ON dbo.TblAttributionContract.MangerialAreaID = dbo.TblManagerialArea.ID RIGHT OUTER JOIN"
'str = str & "   dbo.TblCustemers ON dbo.TblAttributionContract.VendorID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
'str = str & "   dbo.ACCOUNTS ON dbo.TblCustemers.Account_Code = dbo.ACCOUNTS.Account_Code ON dbo.TblVendorCars.ID = dbo.TblVehicleAllocation_Details.CarID LEFT OUTER JOIN"
'str = str & "   dbo.TblDurations ON dbo.TblAttributionContract.DurationID = dbo.TblDurations.ID"
'***************************************************************************
str = "SELECT distinct      dbo.TblAttributionInstallmentDivided.DDEmbarkDate, dbo.TblAttributionInstallmentDivided.DDEmbarkDateH, dbo.TblCustemers.RecordNo, "
 str = str & "    dbo.TblAttributionContract.IDAC, dbo.TblAttributionContract.DurationID, dbo.TblDurations.Name AS DurationName, dbo.TblCustemers.CusName,"
str = str & "   dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode, dbo.TblCustemers.CusID, dbo.TblAttributionContract.StartContractDate,"
 str = str & "                        dbo.TblAttributionContract.EndContractDate, dbo.TblAttributionContract.FromDate, dbo.TblDurations.FromDate AS DurFromDate,"
str = str & "                         dbo.TblDurations.FromDateH AS DurFromDateH, dbo.TblDurations.ToDate AS DurToDate, dbo.TblVehicleAllocation_Details.Type,"
str = str & "                         dbo.TblVehicleAllocation_Details.StudentCount, dbo.TblVehicleAllocation_Details.BoardNo, dbo.TblVehicleAllocation_Details.rate,"
str = str & "                         dbo.TblVehicleAllocation_Details.Custom, dbo.TblVehicleAllocation_Details.DayRate, dbo.TblVehicleAllocation_Details.StudentCustom,"
str = str & "                         dbo.TblVehicleAllocation_Details.CarID, dbo.TblDurations_Details.Name, dbo.TblDurations_Details.ID AS MonthID, dbo.TblVehicleAllocation_Details.ID,"
str = str & "                         dbo.TblVehicleAllocation_Details.SchoolFileID, dbo.TblAttributionContract.StopDeal, dbo.TblAttributionContract.StopDate, dbo.TblAttributionContract.StopDateH,"
str = str & "                         dbo.TblAttributionContract.BranchID, dbo.TblCustemers.Account_Code, dbo.ACCOUNTS.Account_Serial, dbo.TblCustemers.IBAN, dbo.TblCustemers.BankAccount,"
str = str & "                         dbo.TblDurations.Type AS DurationType, dbo.TblAttributionInstallmentDivided.ID AS DividID, dbo.TblAttributionContract.MangerialAreaID,"
str = str & "                         dbo.TblManagerialArea.Name AS MAName, dbo.TblAttributionContract.CityID, dbo.TblAttributionContract.VendorID, dbo.TblSchooleFile.Name AS schoolName,"
str = str & "                         dbo.TblSchooleFile.Namee AS schoolNamee"
str = str & "   FROM         dbo.ACCOUNTS RIGHT OUTER JOIN"
str = str & "                         dbo.TblDurations_Details INNER JOIN"
str = str & "                         dbo.TblAttributionContract INNER JOIN"
str = str & "                         dbo.TblVehicleAllocation_Details ON dbo.TblAttributionContract.IDAC = dbo.TblVehicleAllocation_Details.IDVA INNER JOIN"
str = str & "                         dbo.TblAttributionInstallmentDivided ON dbo.TblVehicleAllocation_Details.ID = dbo.TblAttributionInstallmentDivided.DetailsID ON"
str = str & "                         dbo.TblDurations_Details.ID = dbo.TblAttributionInstallmentDivided.MonthID INNER JOIN"
str = str & "                         dbo.TblSchooleFile ON dbo.TblVehicleAllocation_Details.SchoolFileID = dbo.TblSchooleFile.ID LEFT OUTER JOIN"
str = str & "                         dbo.TblManagerialArea ON dbo.TblAttributionContract.MangerialAreaID = dbo.TblManagerialArea.ID RIGHT OUTER JOIN"
str = str & "                         dbo.TblCustemers ON dbo.TblAttributionContract.VendorID = dbo.TblCustemers.CusID ON"
str = str & "                         dbo.ACCOUNTS.Account_Code = dbo.TblCustemers.Account_Code LEFT OUTER JOIN"
str = str & "                         dbo.TblVendorCars ON dbo.TblVehicleAllocation_Details.CarID = dbo.TblVendorCars.ID LEFT OUTER JOIN"
str = str & "                         dbo.TblDurations ON dbo.TblAttributionContract.DurationID = dbo.TblDurations.ID"
                      
'***************************************************************************
str = str & "       WHERE  (     dbo.TblVehicleAllocation_Details.Type = 3) AND (dbo.TblAttributionInstallmentDivided.RE_Paid IS NULL OR  dbo.TblAttributionInstallmentDivided.RE_Paid = 0 ) "


  If DcDur.BoundText <> "" Then
          str = str & "  and  TblAttributionContract.DurationID  = " & val(DcDur.BoundText)
  End If
    
  If dcMontth.BoundText <> "" Then
          str = str & "      and  TblAttributionInstallmentDivided.MonthID = " & val(dcMontth.BoundText)
  End If
     
  If dcBranch.BoundText <> "" Then
          str = str & "      and  TblAttributionContract.BranchID  = " & val(dcBranch.BoundText)
  End If
  
 If dcMangerialAreaID.BoundText <> "" Then
        str = str & "  and dbo.TblAttributionContract.MangerialAreaID = " & val(dcMangerialAreaID.BoundText)
 End If
  
  If dcCustomer.BoundText <> "" Then
        str = str & " and dbo.TblCustemers.CusID =  " & val(dcCustomer.BoundText)
  End If
  
  str = str & "   and TblAttributionContract.IDAC  in (  SELECT TblAttributionContract.IDAC   from    TblAttributionContract  , TblMinistryContract_Installment    "
  str = str & "   Where TblAttributionContract.IDAC = TblMinistryContract_Installment.IDMC "
  str = str & "   and TblMinistryContract_Installment.type = 2 and  TblMinistryContract_Installment.VRID = " & val(Me.Text1.Text) & " ) "
  
      
     str = str & "      order by dbo.TblAttributionContract.IDAC , dbo.TblVehicleAllocation_Details.BoardNo  "
     
     
    Set RsTemp = New ADODB.Recordset
    RsTemp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
    Dim V As Integer, H As Integer, WD
    Dim tot As Integer, daycount As Integer, SchoolFileID As Integer
    
    Dim mm As Integer
    mm = Grid.Rows
    
     If RsTemp.RecordCount > 0 Then
            RsTemp.MoveFirst
            With Grid
             ProgressBar1.Max = RsTemp.RecordCount + mm
            ' Grid.Rows = Grid.FixedRows + RsTemp.RecordCount
           '  For j = Grid.FixedRows To Grid.Rows - 1
           Grid.Rows = mm + RsTemp.RecordCount
           
             For j = mm To Grid.Rows - 1
                    ProgressBar1.value = j - 2
                    .TextMatrix(j, .ColIndex("Serial")) = j - 1
               
      '              If IsNull(RsTemp("PayMentPayed").value) Then
      '                        TextMatrix(j, .ColIndex("PayMentPayed")) = ""
      '              Else
      '                      If RsTemp("PayMentPayed").value = 0 Then
      '                        TextMatrix(j, .ColIndex("PayMentPayed")) = ""
      '                      Else
      '                        TextMatrix(j, .ColIndex("PayMentPayed")) = "‰⁄„"
      '                      End If
      '
      '              End If
      '
                    
                    .TextMatrix(j, .ColIndex("fullcode")) = IIf(IsNull(RsTemp("Fullcode").value), "", RsTemp("Fullcode").value)
                    .TextMatrix(j, .ColIndex("cusname")) = IIf(IsNull(RsTemp("CusName").value), "", RsTemp("CusName").value)
                    If SystemOptions.UserInterface = ArabicInterface Then
                     .TextMatrix(j, .ColIndex("schoolName")) = IIf(IsNull(RsTemp("schoolName").value), "", RsTemp("schoolName").value)
                     Else
                     .TextMatrix(j, .ColIndex("schoolName")) = IIf(IsNull(RsTemp("schoolNamee").value), "", RsTemp("schoolNamee").value)
                     End If
                    
                    '.TextMatrix(j, .ColIndex("InstallmentNo")) = IIf(IsNull(RsTemp("InstallmentNo").value), "", RsTemp("InstallmentNo").value)
                    .TextMatrix(j, .ColIndex("ID")) = IIf(IsNull(RsTemp("ID").value), "", RsTemp("ID").value)
                    .TextMatrix(j, .ColIndex("recordno")) = IIf(IsNull(RsTemp("recordno").value), "", RsTemp("recordno").value)
                    .TextMatrix(j, .ColIndex("DividID")) = IIf(IsNull(RsTemp("DividID").value), "", RsTemp("DividID").value)
                    .TextMatrix(j, .ColIndex("IDAC")) = IIf(IsNull(RsTemp("IDAC").value), "", RsTemp("IDAC").value)
                    .TextMatrix(j, .ColIndex("CusID")) = IIf(IsNull(RsTemp("CusID").value), "", RsTemp("CusID").value)
                    ' TextMatrix(mm +j, .ColIndex("MonthID")) = IIf(IsNull(RsTemp("MonthID").value), "", RsTemp("MonthID").value)
                    .TextMatrix(j, .ColIndex("StartContractDate")) = IIf(IsNull(RsTemp("StartContractDate").value), "", RsTemp("StartContractDate").value)
                    .TextMatrix(j, .ColIndex("EndContractDate")) = IIf(IsNull(RsTemp("EndContractDate").value), "", RsTemp("EndContractDate").value)
                    .TextMatrix(j, .ColIndex("FromDate")) = IIf(IsNull(RsTemp("FromDate").value), "", RsTemp("FromDate").value)
                    .TextMatrix(j, .ColIndex("Car")) = IIf(IsNull(RsTemp("BoardNo").value), "", RsTemp("BoardNo").value)
                    'MsgBox
                     '.TextMatrix(j, .ColIndex("CarID")) = IIf(IsNull(RsTemp("CarID").value), "", RsTemp("CarID").value)
                    .TextMatrix(j, .ColIndex("Account_Code")) = IIf(IsNull(RsTemp("Account_Code").value), "", RsTemp("Account_Code").value)
                    .TextMatrix(j, .ColIndex("Account_Serial")) = IIf(IsNull(RsTemp("Account_Serial").value), "", RsTemp("Account_Serial").value)
                    .TextMatrix(j, .ColIndex("IBAN")) = IIf(IsNull(RsTemp("IBAN").value), "", RsTemp("IBAN").value)
                    .TextMatrix(j, .ColIndex("BankAccount")) = IIf(IsNull(RsTemp("BankAccount").value), "", RsTemp("BankAccount").value)
                    Dim stopDeal  As Boolean, DateStop As Date, DateStopH As String, Typ As Integer, sf As Integer
                    
                    stopDeal = IIf(IsNull(RsTemp("StopDeal").value), False, RsTemp("StopDeal").value)
                    DateStop = IIf(IsNull(RsTemp("StopDate").value), 0, RsTemp("StopDate").value)
                    DateStopH = IIf(IsNull(RsTemp("StopDateH").value), 0, RsTemp("StopDateH").value)
                    Typ = IIf(IsNull(RsTemp("DurationType").value), 0, RsTemp("DurationType").value)
                    sf = IIf(IsNull(RsTemp("SchoolFileID").value), 0, RsTemp("SchoolFileID").value)
                    
                    Dim DDEmbarkDate As Date
                    Dim DDEmbarkDateH As String
                    DDEmbarkDate = IIf(IsNull(RsTemp("DDEmbarkDate").value), RsTemp("DurFromDate").value, RsTemp("DDEmbarkDate").value)
                    DDEmbarkDateH = IIf(IsNull(RsTemp("DDEmbarkDateH").value), RsTemp("DurFromDateH").value, RsTemp("DDEmbarkDateH").value)
                    
                    Dim exc As Integer
                    If Not (IsNull(RsTemp("IDAC").value) Or IsNull(RsTemp("MonthID").value)) Then
                            GetVac (RsTemp("IDAC").value), (RsTemp("MonthID").value), RsTemp("DurationID").value, sf, tot, daycount, stopDeal, DateStop, DateStopH, Typ, DDEmbarkDate, DDEmbarkDateH
                            H = GetHold(RsTemp("ID").value, (RsTemp("MonthID").value), DDEmbarkDate, DDEmbarkDateH, stopDeal, DateStop, DateStopH, Typ)
                            
                            GetDeducts RsTemp("IDAC").value, val(DcDur.BoundText), RsTemp("MonthID").value, IIf(IsNull(RsTemp("CarID").value), 0, RsTemp("CarID").value), j, _
                            IIf(IsNull(RsTemp("schoolfileID").value), 0, RsTemp("schoolfileID").value)
                            
                            exc = calExcepDays(val(DcDur.BoundText), RsTemp("MonthID").value, _
                            IIf(IsNull(RsTemp("vendorid").value), -1, RsTemp("vendorid").value), _
                            IIf(IsNull(RsTemp("CityID").value), -1, RsTemp("CityID").value), _
                            IIf(IsNull(RsTemp("MangerialAreaID").value), -1, RsTemp("MangerialAreaID").value))
                    End If
                    
                     .TextMatrix(j, .ColIndex("adddays")) = exc
                     If Not IsNull(RsTemp("MonthID").value) Then
                            ' WD = GetMonthDays(RsTemp("MonthID").value, stopDeal, DateStop, DateStopH, Typ)
                            WD = GetMonthDays(RsTemp("ID").value, RsTemp("MonthID").value, DDEmbarkDate, _
                            DDEmbarkDateH, _
                            stopDeal, DateStop, DateStopH, Typ)
                     End If
                 '    If H > 10 Then MsgBox .TextMatrix(j, .ColIndex("Car"))
                    .TextMatrix(j, .ColIndex("VacDay")) = daycount
                    .TextMatrix(j, .ColIndex("WorkDay")) = WD - H
                     
                     Dim dayrate As Double
                     dayrate = IIf(IsNull(RsTemp("DayRate").value), 0, RsTemp("DayRate").value)
                    
                     If Not (IsNull(RsTemp("IDAC").value)) Then
                           .TextMatrix(j, .ColIndex("VacValue")) = daycount * dayrate
                     End If
                     .TextMatrix(j, .ColIndex("DayRate")) = dayrate
                    .TextMatrix(j, .ColIndex("Value")) = dayrate * (WD + exc - H)
                    
                    .TextMatrix(j, .ColIndex("MA")) = IIf(IsNull(RsTemp("MaName").value), "", RsTemp("MaName").value)
                    
                     If j > 15 Then
                      '      .TopRow = j - 10
                    End If
                    lblRow.Caption = " Row No.  " & j
                    DoEvents
                    RsTemp.MoveNext
             Next
            End With
    End If
calculation
RemoveStopedCar
lblRow.Caption = ""

End Sub

Public Sub CalculteValueAdded(LongRow As Long, Optional ByVal mIsDisplay As Boolean = False)

'If SystemOptions.PriceWithVAT = True Then Exit Sub
If (TxtModFlg.Text = "R" Or TxtModFlg.Text = "" Or TxtModFlg.Text = "E") Then Exit Sub
 Dim Percentg As Double
Dim LngItemID As Double
Dim AccountVATCreit As String
Dim cCompanyInfo As New ClsCompanyInfo

If mdifrmmain.taxes.Visible = True Then
'If TransType = 9 And ReturnSales = True Then
    If mIsDisplay And val(Grid.TextMatrix(LongRow, Grid.ColIndex("Vatyo"))) <> 0 Then Exit Sub
  
    Dim mDate As Date
    
    If SystemOptions.AllItemInVAT = True Then
        Percentg = val(cCompanyInfo.VATItems)
    Else
      
      PercentgValueAddedAccount_Transec Me.Date.value, 47, 1, AccountVATCreit, Percentg
        
    End If
    If Percentg = -1 Then
        Percentg = 0
        If SystemOptions.UserInterface = ArabicInterface Then
            Grid.TextMatrix(LongRow, Grid.ColIndex("TypeVAT")) = "„⁄›Ì"
        Else
            Grid.TextMatrix(LongRow, Grid.ColIndex("TypeVAT")) = "Exempt"
        End If
    Else
        If Grid.ColIndex("TypeVAT") <> -1 Then 'salim1503
            Grid.TextMatrix(LongRow, Grid.ColIndex("TypeVAT")) = Percentg
        End If
    
    End If
    If Grid.ColIndex("Vatyo") <> -1 Then
    '    If val(Grid.TextMatrix(LongRow, Grid.ColIndex("Vatyo"))) = 0 Then
    '    Grid.TextMatrix(LongRow, Grid.ColIndex("Vatyo")) = Percentg
    '    Else
    '    Percentg = val(Grid.TextMatrix(LongRow, Grid.ColIndex("Vatyo")))
    '
    '    End If
    Grid.TextMatrix(LongRow, Grid.ColIndex("Vatyo")) = Percentg
    
    End If
     Grid.TextMatrix(LongRow, Grid.ColIndex("Vat2")) = val(Grid.TextMatrix(LongRow, Grid.ColIndex("TotalNet"))) * Percentg / 100
     Grid.TextMatrix(LongRow, Grid.ColIndex("Net")) = val(Grid.TextMatrix(LongRow, Grid.ColIndex("TotalNet"))) + val(Grid.TextMatrix(LongRow, Grid.ColIndex("Vat2")))
    

End If

End Sub

Private Function calExcepDays(DurID As Integer, MonthID As Integer, Optional vendorID As Integer, Optional CityID As Integer, Optional MA As Integer) As Integer
Dim str As String, cnt As Integer
cnt = 0


str = " SELECT sum(days) days from TblAddExceptionDays where durationID = " & DurID & " and monthid = " & MonthID & " and "
'str = str & "  (  ( ( vendorid = " & vendorID & "  or  ( cityid = " & CityID & "  and  managerialareaid =   " & MA & " )  or ( cityid = " & CityID & " and   managerialareaid is null )   ) and spes = 1  )  or ( alls = 1 )  "
str = str & " ("
str = str & "  (Alls = 1 ) or"
str = str & "   (vendor = 1 and vendorid = " & vendorID & ") or"
str = str & "   (city = 1 and cityID = " & CityID & ") or"
str = str & "   (Ma = 1 and CityID = " & CityID & " and VendorID = " & vendorID & " )"
str = str & " ) "


Set RsExcep = New ADODB.Recordset
RsExcep.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText

If RsExcep.RecordCount > 0 Then
        cnt = IIf(IsNull(RsExcep("Days").value), 0, RsExcep("Days").value)
End If
calExcepDays = cnt
End Function

Private Sub RemoveStopedCar()
On Error Resume Next
Dim i As Integer, count As Integer, j As Integer
With Grid
count = .Rows - 1
j = 1
For i = count To .FixedRows Step -1
                    If val(.TextMatrix(i, .ColIndex("WorkDay"))) <= 0 Then
                                .RemoveItem (i)
                                
                                count = count - 1
                    Else
                       '         .TextMatrix(i, .ColIndex("Serial")) = j
                       '          j = j + 1
                    End If
Next

End With

End Sub


Private Sub GetVac(IDMC As Integer, MonthID As Integer, DurationID As Integer, SchoolFileID As Integer, ByRef Total As Integer, _
ByRef daycount As Integer, stopDeal As Boolean, DateStop As Date, DateStopH As String, Typ As Integer, EmbarkDate As Date, EmbarkDateH As String)

        Dim str As String, cunt As Integer, CityID As Integer, DurID As Integer, DayDiff As Integer, j As Integer, i As Integer
        Total = 0
        daycount = 0
        str = " select h.DurationID , h.MonthID ,d.SchoolFileID ,sum (d.daycount) daycount , sum (d.dayvalue)  dayvalue , sum ( (d.daycount * d.dayvalue )) Total  , D.FromDateH  , D.FromDate "
        str = str & " from TblconfirmVacation  h, TblConfirmVacation_Details d "
        str = str & " where      h.ID = d.HID and  DurationID = " & DurationID & "  and  MonthID = " & MonthID & " and  SchoolFileID = " & SchoolFileID
        
        If Typ = 0 Then
                str = str & " and D.FromDate >= " & EmbarkDate & "     "
        ElseIf Typ = 1 Then
                str = str & "  and  D.FromDateH  >= '" & EmbarkDateH & "'"
        End If
        
        If stopDeal = True Then
                If Typ = 0 Then
                    str = str & "    and  h.FromDate <   " & DateStop & "   "
                ElseIf Typ = 1 Then
                    str = str & "   and  h.FromDateH <   '" & DateStopH & "'  "
                End If
        End If
        str = str & " group by  h.DurationID , h.MonthID ,d.SchoolFileID  ,D.FromDateH  ,D.FromDate  "
        Set RsTemp3 = New ADODB.Recordset
        RsTemp3.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
        If RsTemp3.RecordCount > 0 Then
            For i = 0 To RsTemp3.RecordCount - 1
                 Total = Total + IIf(IsNull(RsTemp3("Total").value), 0, RsTemp3("Total").value)
                 daycount = daycount + IIf(IsNull(RsTemp3("daycount").value), 0, RsTemp3("daycount").value)
                 RsTemp3.MoveNext
           Next
        End If
     
End Sub

Private Function GetDayRate(IDMC As Integer, FromDate As String, ToDate As String)

 Dim str As String, days As Integer, net As Double, Operation As String
 str = " select * from TblAttributionContract  where idac = " & IDMC
 Set Rs_Temp = New ADODB.Recordset
 Rs_Temp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
 If Rs_Temp.RecordCount > 0 Then
        days = DateDiff("d", FromDate, ToDate)
        If days > 0 Then
                Operation = IIf(IsNull(Rs_Temp("AdditionalType").value), "", Rs_Temp("AdditionalType").value)
                If Operation = "add" Then
                       net = val(Rs_Temp("studentcount").value) * val(Rs_Temp("StudentCustom").value) + val(Rs_Temp("StudentCustom").value)
                ElseIf Operation = "sub" Then
                       net = val(Rs_Temp("studentcount").value) * val(Rs_Temp("StudentCustom").value) - val(Rs_Temp("StudentCustom").value)
                Else
                        net = val(Rs_Temp("studentcount").value) * val(Rs_Temp("StudentCustom").value)
                End If
                GetDayRate = net / days
        End If
 End If
 
End Function

Private Sub GetDeducts(IDMC As Integer, dur As Integer, MonthID As Integer, CarID As Integer, Row As Integer, SchoolFileID As Integer)

         Dim str As String, i As Integer, j As Integer, absc As Boolean
         str = "SELECT   dbo.TblConfirmViolation.ID, dbo.TblConfirmViolation.DurationID, dbo.TblConfirmViolation.ViolationID, dbo.TblConfirmViolation.MinistryContractID ,TblConfirmViolation.AbsenceCount,"
         str = str & " dbo.TblConfirmViolation.Date , dbo.TblConfirmViolation.value, dbo.TblConfirmViolation.monthid, dbo.TblViolationTypes.name ,TblViolationTypes.absence"
         str = str & " FROM     dbo.TblConfirmViolation INNER JOIN   dbo.TblViolationTypes ON dbo.TblConfirmViolation.ViolationID = dbo.TblViolationTypes.ID"
         str = str & " where DurationID = " & dur & " and MonthID = " & MonthID & "  and MinistryContractID =  " & IDMC & " and TblConfirmViolation.CarID =  " & CarID & " and TblConfirmViolation.schoolID = " & SchoolFileID
         
         Set RsTemp4 = New ADODB.Recordset
         RsTemp4.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
                        
         With Grid
         If RsTemp4.RecordCount > 0 Then
         
                For i = 1 To RsTemp4.RecordCount
                
                         absc = IIf(IsNull(RsTemp4("absence").value), False, RsTemp4("absence").value)
                         
                         If absc = True Then
                                .TextMatrix(Row, .ColIndex("AbsenceCount")) = val(.TextMatrix(Row, .ColIndex("AbsenceCount"))) + IIf(IsNull(RsTemp4("AbsenceCount").value), 0, RsTemp4("AbsenceCount").value)
                                .TextMatrix(Row, .ColIndex("Avalue")) = val(.TextMatrix(Row, .ColIndex("Avalue"))) + IIf(IsNull(RsTemp4("Value").value), 0, RsTemp4("Value").value)
                         End If
                
                        For j = 1 To 20
                                If .TextMatrix(1, .ColIndex("d" & j)) = RsTemp4("Name").value Then
                                        .TextMatrix(Row, .ColIndex("d" & j)) = val(.TextMatrix(Row, .ColIndex("d" & j))) + IIf(IsNull(RsTemp4("Value").value), 0, RsTemp4("Value").value)
                                End If
                        Next
                        RsTemp4.MoveNext
                        
                Next
         End If
         End With
         
End Sub

Private Function GetHold(ID As Integer, MonthID As Integer, DDEmbarkDate As Date, DDEmbarkDateH As String, Optional stopDeal As Boolean, Optional StopDate As Date, Optional StopDateH As String, Optional Typ As Integer)
    Dim str As String, cunt As Integer, mm As String, _
      stoped As Boolean, StartStopDealing As String, StartStopDealingH As String, EndStopDealing As String, EndStopDealingH As String
    'salimHolding
    Set Rs_Temp = New ADODB.Recordset
    mm = " select * from tblvehicleallocation_details   where id =  " & ID
    Rs_Temp.Open mm, Cn, adOpenStatic, adLockOptimistic, adCmdText
   If Rs_Temp.RecordCount > 0 Then
        stoped = IIf(IsNull(Rs_Temp("stoped").value), False, Rs_Temp("stoped").value)
        StartStopDealing = IIf(IsNull(Rs_Temp("StartStopDealing").value), "", Rs_Temp("StartStopDealing").value)
        StartStopDealingH = IIf(IsNull(Rs_Temp("StartStopDealingH").value), "", Rs_Temp("StartStopDealingH").value)
        EndStopDealing = IIf(IsNull(Rs_Temp("EndStopDealing").value), "", Rs_Temp("EndStopDealing").value)
        EndStopDealingH = IIf(IsNull(Rs_Temp("EndStopDealingH").value), "", Rs_Temp("EndStopDealingH").value)
   End If
      
      
   If stopDeal = False Then
            If Typ = 0 Then
                         If stoped = True Then
                                     str = " select count (*)  cunt, DDID  from TblVacationSchedule where  ISVac = 1 and  ddid =   " & MonthID & "  and date >= " & DDEmbarkDate & "   and  date <   " & CDate(StartStopDealing) & "  group by DDID"
                         Else
                                    If StartStopDealing <> "" And EndStopDealing <> "" Then
                                             str = " select count (*)  cunt, DDID  from TblVacationSchedule where ISVac = 1 and    ddid =   " & MonthID & "  and date >= " & DDEmbarkDate & "   and  date between   " & CDate(StartStopDealing) & " and  " & CDate(EndStopDealing) & "  group by DDID"
                                    ElseIf StartStopDealing <> "" And EndStopDealing = "" Then
                                             str = " select count (*)  cunt, DDID  from TblVacationSchedule where  ISVac = 1 and   ddid =   " & MonthID & "  and date >= " & DDEmbarkDate & "   and  date <   " & CDate(StartStopDealing) & "  group by DDID"
                                    Else
                                             str = " select count (*)  cunt, DDID  from TblVacationSchedule where  ISVac = 1 and   ddid =   " & MonthID & "  and date >= " & DDEmbarkDate & "     group by DDID "
                                    End If
                         End If
            ElseIf Typ = 1 Then
                         If stoped = True Then
                                     str = " select count (*)  cunt, DDID  from TblVacationSchedule where   ddid =   " & MonthID & "  and DateH >= '" & DDEmbarkDateH & "'   and  dateH <   '" & StartStopDealingH & "'  group by DDID"
                         Else
                                    If StartStopDealingH <> "" And EndStopDealingH <> "" Then
                                             str = " select count (*)  cunt, DDID  from TblVacationSchedule where  ISVac = 1 and  ddid =   " & MonthID & "  and dateH >= '" & DDEmbarkDateH & "'   and  dateH between   '" & StartStopDealingH & "' and  '" & EndStopDealingH & "'  group by DDID"
                                    ElseIf StartStopDealing <> "" And EndStopDealing = "" Then
                                             str = " select count (*)  cunt, DDID  from TblVacationSchedule where ISVac = 1 and   ddid =   " & MonthID & "  and dateH >= '" & DDEmbarkDateH & "'   and  dateH <   '" & StartStopDealingH & "'  group by DDID"
                                     Else
                                             str = " select count (*)  cunt, DDID  from TblVacationSchedule where ISVac = 1 and   ddid =   " & MonthID & "  and dateH >= '" & DDEmbarkDateH & "'   group by DDID "
                                    End If
                         End If
            
            End If
    End If
        
    If stopDeal = True Then
            If Typ = 0 Then
                str = " select count (*)  cunt, DDID  from TblVacationSchedule where ISVac = 1 and  ddid =   " & MonthID & "  and date >= " & DDEmbarkDate & "  and  date <   " & StopDate & "  group by DDID"
            ElseIf Typ = 1 Then
                If DDEmbarkDateH = "" Then
                            str = " select count (*)  cunt, DDID  from TblVacationSchedule where ISVac = 1 and  ddid =   " & MonthID & "  and  dateH <   '" & StopDateH & "'  group by DDID"
                Else
                            str = " select count (*)  cunt, DDID  from TblVacationSchedule where ISVac = 1 and  ddid =   " & MonthID & "  and dateH >= '" & DDEmbarkDateH & "'  and dateH <   '" & StopDateH & "'  group by DDID"
                End If
            End If

    End If
          
             Set RsTemp2 = New ADODB.Recordset
             RsTemp2.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
             If RsTemp2.RecordCount > 0 Then
                    cunt = IIf(IsNull(RsTemp2("cunt").value), 0, RsTemp2("cunt").value)
             End If
    GetHold = cunt
End Function

Private Function GetVac1(IDMC As Integer, MonthID As Integer)

        Dim str As String, cunt As Integer, CityID As Integer, DurID As Integer, DayDiff As Integer, j As Integer
        str = " select CityID , DurationID  from  TblAttributionContract where IDAC = " & IDMC
        Set RsTemp2 = New ADODB.Recordset
        RsTemp2.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
        If RsTemp2.RecordCount > 0 Then
               CityID = IIf(IsNull(RsTemp2("CityID").value), 0, RsTemp2("CityID").value)
               DurID = IIf(IsNull(RsTemp2("DurationID").value), 0, RsTemp2("DurationID").value)
        End If
        str = "select * from TblConfirmVacation  where DurationID = " & DurID & "and CityID = " & CityID & " and MonthID = " & MonthID
        Set RsTemp2 = New ADODB.Recordset
        RsTemp2.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
        If RsTemp2.RecordCount > 0 Then
            For j = 0 To RsTemp2.RecordCount - 1
                 DayDiff = DayDiff + DateDiff("d", RsTemp2("FromDate").value, RsTemp2("ToDate").value, vbSaturday)
           Next
        End If
        GetVac1 = DayDiff
End Function

Private Function GetMonthDays(ID As Integer, MonthID As Integer, Embark As Date, EmbarkH As String, Optional stopDeal As Boolean, Optional StopDate As Date, Optional StopDateH As String, Optional Typ As Integer)
    
    Dim str As String, cunt As Integer, mm As String, _
    stoped As Boolean, StartStopDealing As String, StartStopDealingH As String, EndStopDealing As String, EndStopDealingH As String
    
    Set Rs_Temp = New ADODB.Recordset
    mm = " select * from tblvehicleallocation_details   where id =  " & ID
    Rs_Temp.Open mm, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If Rs_Temp.RecordCount > 0 Then
        stoped = IIf(IsNull(Rs_Temp("stoped").value), False, Rs_Temp("stoped").value)
        StartStopDealing = IIf(IsNull(Rs_Temp("StartStopDealing").value), "", Rs_Temp("StartStopDealing").value)
        StartStopDealingH = IIf(IsNull(Rs_Temp("StartStopDealingH").value), "", Rs_Temp("StartStopDealingH").value)
        EndStopDealing = IIf(IsNull(Rs_Temp("EndStopDealing").value), "", Rs_Temp("EndStopDealing").value)
        EndStopDealingH = IIf(IsNull(Rs_Temp("EndStopDealingH").value), "", Rs_Temp("EndStopDealingH").value)
    End If
   
    If stopDeal = False Then
           If Typ = 0 Then
                         If stoped = True Then
                                     str = " select count (*)  cunt, DDID  from TblVacationSchedule where   ddid =   " & MonthID & "  and date >= " & Embark & "   and  date <   " & CDate(StartStopDealing) & "  group by DDID"
                         Else
                                    If StartStopDealing <> "" And EndStopDealing <> "" Then
                                             str = " select count (*)  cunt, DDID  from TblVacationSchedule where   ddid =   " & MonthID & "  and date >= " & Embark & "   and  date between   " & CDate(StartStopDealing) & " and  " & CDate(EndStopDealing) & "  group by DDID"
                                    ElseIf StartStopDealing <> "" And EndStopDealing = "" Then
                                             str = " select count (*)  cunt, DDID  from TblVacationSchedule where   ddid =   " & MonthID & "  and date >= " & Embark & "   and  date <   " & CDate(StartStopDealing) & "  group by DDID"
                                    Else
                                             str = " select count (*)  cunt, DDID  from TblVacationSchedule where   ddid =   " & MonthID & "  and date >= " & Embark & "     group by DDID"
                                    End If
                         End If
           ElseIf Typ = 1 Then
                         If stoped = True Then
                                     str = " select count (*)  cunt, DDID  from TblVacationSchedule where   ddid =   " & MonthID & "  and DateH >= '" & EmbarkH & "'   and  dateH <   '" & StartStopDealingH & "'  group by DDID"
                         Else
                                    If StartStopDealingH <> "" And EndStopDealingH <> "" Then
                                             str = " select count (*)  cunt, DDID  from TblVacationSchedule where   ddid =   " & MonthID & "  and dateH >= '" & EmbarkH & "'   and  dateH between   '" & StartStopDealingH & "' and  '" & EndStopDealingH & "'  group by DDID"
                                    ElseIf StartStopDealing <> "" And EndStopDealing = "" Then
                                             str = " select count (*)  cunt, DDID  from TblVacationSchedule where   ddid =   " & MonthID & "  and dateH >= '" & EmbarkH & "'   and  dateH <   '" & StartStopDealingH & "'  group by DDID"
                                     Else
                                             str = " select count (*)  cunt, DDID  from TblVacationSchedule where   ddid =   " & MonthID & "  and dateH >= '" & EmbarkH & "'     group by DDID"
                                    End If
                         End If
            
            End If
   End If
    
   If stopDeal = True Then
            If Typ = 0 Then
                str = " select count (*)  cunt, DDID  from TblVacationSchedule where   ddid =   " & MonthID & "  and date >= " & Embark & "   and  date <   " & StopDate & "  group by DDID"
            ElseIf Typ = 1 Then
                If EmbarkH = "" Then
                        str = " select count (*)  cunt, DDID  from TblVacationSchedule where   ddid =   " & MonthID & "  and  dateH <=  '" & StopDateH & "'  group by DDID"
                Else
                        str = " select count (*)  cunt, DDID  from TblVacationSchedule where   ddid =   " & MonthID & " and dateh >= '" & EmbarkH & "'  and  dateH <   '" & StopDateH & "'  group by DDID"
                End If
            End If

    End If
   
    
     Set RsTemp2 = New ADODB.Recordset
        RsTemp2.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
        If RsTemp2.RecordCount > 0 Then
               cunt = IIf(IsNull(RsTemp2("cunt").value), 0, RsTemp2("cunt").value)
        End If
    GetMonthDays = cunt

End Function


Private Sub dcMangerialAreaID_KeyUp(KeyCode As Integer, Shift As Integer)
        
        If KeyCode = vbKeyF3 Then
            Unload FrmSearch_BasicData
            FrmSearch_BasicData.SendForm = "ExchangeRequest"
            FrmSearch_BasicData.show
        End If
        
End Sub

Private Sub dcMontth_Click(Area As Integer)
'Fill_Grid
End Sub

Private Sub Form_Activate()
'    XPTxtBoxID.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        If Me.TxtModFlg.Text = "R" Then
            Cmd_Click (0)
        Else
            SendKeys "{TAB}"
        End If
    End If

    If Me.TxtModFlg.Text = "R" Then
        If KeyCode = vbKeyDown Or KeyCode = vbKeyEnd Then
            XPBtnMove_Click (2)
        ElseIf KeyCode = vbKeyUp Or KeyCode = vbKeyHome Then
            XPBtnMove_Click (1)
        ElseIf KeyCode = vbKeyRight Or KeyCode = vbKeyPageDown Then
            XPBtnMove_Click (3)
        ElseIf KeyCode = vbKeyLeft Or KeyCode = vbKeyPageUp Then
            XPBtnMove_Click (0)
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

    If Shift = 2 Then
        If KeyCode = vbKeyX Then
            If Cmd(6).Enabled = False Then Exit Sub
            Cmd_Click (6)
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim Dcombos As ClsDataCombos, str As String

    Set Dcombos = New ClsDataCombos
    'Dcombos.GetEmployees DcboGovernmentID
   ' Dcombos.getCountriesGovernments Me.DcboGovernmentID
    Dcombos.GetBranches dcBranch
    Dcombos.GetCustomersSuppliers 2, dcCustomer
    Dcombos.GetUsers Me.DCboUserName
    EntryDate.value = Now


    If SystemOptions.UserInterface = ArabicInterface Then
    str = " Select ID , Name   from TblManagerialArea "
    Else
    str = " Select ID , NameE   from TblManagerialArea "
    End If
    fill_combo dcMangerialAreaID, str

 If SystemOptions.AllowRequestgl = True Then
    Command8.Enabled = True
   End If
 
 
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    LogTextA = "   «·œŒÊ· «·Ì ‘«‘… " & " ÿ·» ’—› „ ⁄ÂœÌ‰  "
    LogTexte = " Open Window " & "  Boxes Data "
    AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, "O", "", ""

    Dim My_SQL As String
    My_SQL = " Select id , name from  TblDurations "
    fill_combo DcDur, My_SQL
  
   

    Resize_Form Me
    AddTip
    Set rs = New ADODB.Recordset
  
   Dim StrSQL As String
   StrSQL = "SELECT  *  From TblExchangeRequest order by ID"
   rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
      
 '  With cbType
 '       If SystemOptions.UserInterface = ArabicInterface Then
 '               .Clear
 '               .AddItem ("‰ﬁœÏ")
 '               .AddItem ("‘Ìﬂ")
 '       Else
 '               .Clear
 '               .AddItem ("Cash")
 ''               .AddItem ("Cheque")
 ''       End If
 '   End With
    Intialize_Deducts
    Inatial_Grid
    
    Me.TxtModFlg.Text = "R"
    XPBtnMove_Click 2
    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If

    Exit Sub

ErrTrap:
End Sub

Private Sub Inatial_Grid()

 With Grid

        .MergeCells = flexMergeFixedOnly
        .MergeRow(0) = True
        
       ' .MergeCol(.ColIndex("No")) = True
'        .Cell(flexcpText, 0, .ColIndex("No"), 1, .ColIndex("No")) = "—ﬁ„ «·”ÿ—"

        .MergeCol(.ColIndex("cusname")) = True
        .Cell(flexcpText, 0, .ColIndex("cusname"), 1, .ColIndex("cusname")) = "«·«”„"

      '  .MergeCol(.ColIndex("PayNo")) = True
      '  .Cell(flexcpText, 0, .ColIndex("PayNo"), 1, .ColIndex("PayNo")) = "—ﬁ„ «·œ›⁄…"

        .MergeCol(.ColIndex("Value")) = True
        .Cell(flexcpText, 0, .ColIndex("Value"), 1, .ColIndex("Value")) = "«·ﬁÌ„…"

        .MergeCol(.ColIndex("Total")) = True
        .Cell(flexcpText, 0, .ColIndex("total"), 1, .ColIndex("total")) = "«Ã„«·Ï «·Õ”„Ì« "
        
        .MergeCol(.ColIndex("Net")) = True
        .Cell(flexcpText, 0, .ColIndex("Net"), 1, .ColIndex("Net")) = "«·’«›Ï «·„” Õﬁ"
        .Cell(flexcpText, 0, .ColIndex("d1"), 0, .ColIndex("d20")) = "Õ”„Ì« "
 
    End With



End Sub



Private Sub Intialize_Deducts()
Dim str As String, i As Integer
Set Rs_Temp = New ADODB.Recordset
str = " select * from TblViolationTypes  where absence = 0 or absence is null"
Rs_Temp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
Rs_Temp.MoveFirst

If Rs_Temp.RecordCount > 0 Then
    For i = 1 To Rs_Temp.RecordCount
        Grid.TextMatrix(1, Grid.ColIndex("d" & i)) = IIf(IsNull(Rs_Temp("Name").value), "", Rs_Temp("Name").value)
        Rs_Temp.MoveNext
    Next
End If


For i = 1 To 20
    If Grid.TextMatrix(1, Grid.ColIndex("d" & i)) = "" Then
         Grid.ColWidth(Grid.ColIndex("d" & i)) = 0
    End If
Next


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

Private Sub ChangeLang()
    Dim XPic As IPictureDisp

    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic

 
   lbl(0).Caption = "No."
   lbl(3).Caption = " Name Ar"
   lbl(7).Caption = " Name En"
   'Label3.Caption = "City"
   
  lbl(2).Caption = "Current Record"
  lbl(4).Caption = "Recors Count"
   
    Me.Caption = "Managerial Area"
    EleHeader.Caption = Me.Caption
   
    Me.Cmd(0).Caption = "New"
    Me.Cmd(1).Caption = "Edit"
    Me.Cmd(2).Caption = "Save"
    Me.Cmd(3).Caption = "Undo"
    Me.Cmd(4).Caption = "Delete"
    'Me.Cmd(5).Caption = "Search"
    Me.Cmd(6).Caption = "Exit"
    Me.Cmd(7).Caption = "Print"
    CmdAttach.Caption = "Attachment"

End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap
    LogTextA = "     «·Œ—ÊÃ „‰ ‘«‘… " & "  ÿ·» ’—› „ ⁄ÂœÌ‰   "
    LogTexte = " Exit Window " & "  Boxes Data "
    AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, "O", "", ""

    If rs.State = adStateOpen Then
        If Not (rs.EOF Or rs.BOF) Then
            If rs.EditMode <> adEditNone Then
                rs.CancelUpdate
            End If
        End If

        rs.Close
    End If
     Clear_Current_User
    Set rs = Nothing
    Set TTP = Nothing
    Exit Sub
ErrTrap:
End Sub



Private Sub Grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)

     With Grid
            Select Case .ColKey(Col)
                Case "d1", "d2", "d3", "d4", "d5", "d6", "d7", "d8", "d9", "d10", "d11", "d12", "d13", "d14", "d15", "d16", "d17", "d18", "d19", "d20", "Value"
                        .TextMatrix(Row, .ColIndex("Total")) = val(.TextMatrix(Row, .ColIndex("Avalue"))) + val(.TextMatrix(Row, .ColIndex("d1"))) + val(.TextMatrix(Row, .ColIndex("d2"))) + val(.TextMatrix(Row, .ColIndex("d3"))) + val(.TextMatrix(Row, .ColIndex("d4"))) + val(.TextMatrix(Row, .ColIndex("d5"))) + val(.TextMatrix(Row, .ColIndex("d6"))) + val(.TextMatrix(Row, .ColIndex("d7"))) + val(.TextMatrix(Row, .ColIndex("d8"))) + val(.TextMatrix(Row, .ColIndex("d9"))) + val(.TextMatrix(Row, .ColIndex("d10"))) + val(.TextMatrix(Row, .ColIndex("d11"))) + val(.TextMatrix(Row, .ColIndex("d12"))) + val(.TextMatrix(Row, .ColIndex("d13"))) + val(.TextMatrix(Row, .ColIndex("d14"))) + val(.TextMatrix(Row, .ColIndex("d15"))) + val(.TextMatrix(Row, .ColIndex("d16"))) + val(.TextMatrix(Row, .ColIndex("d17"))) + val(.TextMatrix(Row, .ColIndex("d18"))) + val(.TextMatrix(Row, .ColIndex("d19"))) + val(.TextMatrix(Row, .ColIndex("d20")))
                        .TextMatrix(Row, .ColIndex("TotalNet")) = val(.TextMatrix(Row, .ColIndex("Value"))) - val(.TextMatrix(Row, .ColIndex("Total")))
                Case "Vatyo"
                    
                     Grid.TextMatrix(Row, Grid.ColIndex("Vat2")) = val(Grid.TextMatrix(Row, Grid.ColIndex("TotalNet"))) * val(Grid.TextMatrix(Row, Grid.ColIndex("Vatyo"))) / 100
                    Grid.TextMatrix(Row, Grid.ColIndex("Net")) = val(Grid.TextMatrix(Row, Grid.ColIndex("TotalNet"))) + val(Grid.TextMatrix(Row, Grid.ColIndex("Vat2")))

            End Select
       End With


End Sub

Private Sub calculation(Optional ByVal IsDisplay As Boolean = False)
    Dim i As Long
     With Grid
            For i = 2 To .Rows - 1
                        .TextMatrix(i, .ColIndex("Total")) = val(.TextMatrix(i, .ColIndex("Avalue"))) + val(.TextMatrix(i, .ColIndex("d1"))) + val(.TextMatrix(i, .ColIndex("d2"))) + val(.TextMatrix(i, .ColIndex("d3"))) + val(.TextMatrix(i, .ColIndex("d4"))) + val(.TextMatrix(i, .ColIndex("d5"))) + val(.TextMatrix(i, .ColIndex("d6"))) + val(.TextMatrix(i, .ColIndex("d7"))) + val(.TextMatrix(i, .ColIndex("d8"))) + val(.TextMatrix(i, .ColIndex("d9"))) + val(.TextMatrix(i, .ColIndex("d10"))) + val(.TextMatrix(i, .ColIndex("VacValue"))) + val(.TextMatrix(i, .ColIndex("d11"))) + val(.TextMatrix(i, .ColIndex("d12"))) + val(.TextMatrix(i, .ColIndex("d13"))) + val(.TextMatrix(i, .ColIndex("d14"))) + val(.TextMatrix(i, .ColIndex("d15"))) + val(.TextMatrix(i, .ColIndex("d16"))) + val(.TextMatrix(i, .ColIndex("d17"))) + val(.TextMatrix(i, .ColIndex("d18"))) + val(.TextMatrix(i, .ColIndex("d19"))) + val(.TextMatrix(i, .ColIndex("d20")))
                        .TextMatrix(i, .ColIndex("TotalNet")) = val(.TextMatrix(i, .ColIndex("Value"))) - val(.TextMatrix(i, .ColIndex("Total")))
                        CalculteValueAdded i, IsDisplay
            Next
       End With



End Sub


Public Sub Retrive_Data()

Dim str As String
str = " select * from TblExchangeRequest2 where id =  " & val(Text1.Text)
Set Rs_Temp = New ADODB.Recordset
Rs_Temp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs_Temp.RecordCount > 0 Then
        dcBranch.BoundText = IIf(IsNull(Rs_Temp("BranchID").value), "", Rs_Temp("BranchID").value)
        DcDur.BoundText = IIf(IsNull(Rs_Temp("DurationID").value), "", Rs_Temp("DurationID").value)
        dcMontth.BoundText = IIf(IsNull(Rs_Temp("Month").value), "", Rs_Temp("Month").value)
        Command1_Click
Else
         Grid.Rows = Grid.FixedRows
End If

End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With Grid
Select Case .ColKey(Col)

Case "Status"
            If Me.TxtModFlg.Text = "E" Or Me.TxtModFlg.Text = "N" Then
            Else
                    Cancel = True
            End If

Case "remark"
         If Me.TxtModFlg.Text = "E" Or Me.TxtModFlg.Text = "N" Then
            Else
                    Cancel = True
            End If
Case "IDAC"
       Cancel = True
Case "MA"
         Cancel = True
Case "fullcode"
          Cancel = True

Case "fullcode"
          Cancel = True
          
Case "cusname"
          Cancel = True
          
Case "Account_Serial"
          Cancel = True
          
          
Case "BankAccount"
          Cancel = True

Case "IBAN"
          Cancel = True

Case "recordno"
          Cancel = True


Case "Car"
          Cancel = True


Case "DayRate"
          Cancel = True


Case "WorkDay"
          Cancel = True


Case "Value"
          Cancel = True


Case "VacDay"
          Cancel = True


Case "VacValue"
          Cancel = True


Case "AbsenceCount"
          Cancel = True

Case "Avalue"
          Cancel = True


Case "StartContractDate"
          Cancel = True


Case "EndContractDate"
          Cancel = True


Case "FromDate"
          Cancel = True

Case "d1"
          Cancel = True


Case "d2"
          Cancel = True


Case "d3"
          Cancel = True


Case "d4"
          Cancel = True


Case "d5"
          Cancel = True


Case "d6"
          Cancel = True


Case "d7"
          Cancel = True


Case "d8"
          Cancel = True


Case "d9"
          Cancel = True


Case "d10"
          Cancel = True


Case "d11"
          Cancel = True


Case "d12"
          Cancel = True


Case "d13"
          Cancel = True


Case "d14"
          Cancel = True


Case "d15"
          Cancel = True


Case "d16"
          Cancel = True



Case "d17"
          Cancel = True
          
Case "d18"
          Cancel = True
          
          
Case "d19"
          Cancel = True
          
Case "d20"
          Cancel = True
          
          
Case "Total"
          Cancel = True
          
          
Case "Net"
          Cancel = True
          
        
          
End Select
End With
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
            Retrive_Data
    End If

End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
        If KeyCode = vbKeyF3 Then
        
                    Unload FrmSearch_Request
                    FrmSearch_Request.SendForm = "MR_VR"
                   FrmSearch_Request.show
        End If
        
        
End Sub

Private Sub txtfullcode_Change()
Dim val1, val2
If txtfullcode.Text = "" Then Exit Sub
Dim str As String, recordno As String, CusID As String
recordno = ""
CusID = ""

    str = " select * From TblCustemers where Type=2  and fullcode = '" & txtfullcode & "'"
    Set Rs_Temp = New ADODB.Recordset
    Rs_Temp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If Rs_Temp.RecordCount > 0 Then
        Rs_Temp.MoveFirst '
        recordno = IIf(IsNull(Rs_Temp("recordno").value), "", Rs_Temp("recordno").value)
        CusID = IIf(IsNull(Rs_Temp("cusID").value), "", Rs_Temp("cusID").value)
     Else
        txtRecordno.Text = ""
        dcCustomer.BoundText = ""
    End If
    
    txtRecordno.Text = recordno
    dcCustomer.BoundText = CusID

End Sub

Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap

    Select Case Me.TxtModFlg.Text

        Case "R"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "ÿ·» ’—› „ ⁄ÂœÌ‰ "
            Else
                Me.Caption = "Boxes Data"
            End If
             Me.Cmd(8).Enabled = False
            Me.Cmd(2).Enabled = False
            Me.Cmd(3).Enabled = False
        
            Me.Cmd(0).Enabled = True
            Me.Cmd(1).Enabled = True
            Me.Cmd(4).Enabled = True
            Me.Cmd(5).Enabled = True
        
            Me.XPBtnMove(0).Enabled = True
            Me.XPBtnMove(1).Enabled = True
            Me.XPBtnMove(2).Enabled = True
            Me.XPBtnMove(3).Enabled = True
        
            

            If rs.RecordCount < 1 Then
                Me.XPBtnMove(0).Enabled = False
                Me.XPBtnMove(1).Enabled = False
                Me.XPBtnMove(2).Enabled = False
                Me.XPBtnMove(3).Enabled = False
                Me.Cmd(1).Enabled = False
                Me.Cmd(4).Enabled = False
            End If
            C1Elastic1.Enabled = False
      
      
      
        Case "N"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "ÿ·» ’—› „ ⁄ÂœÌ‰ ( ÃœÌœ )"
            Else
                Me.Caption = "Exchange Request (New)"
            End If
            Me.Cmd(8).Enabled = True
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "ÿ·» ’—› „ ⁄ÂœÌ‰( ÃœÌœ )"
            Else
                Me.Caption = "Exchange Request  (New)"
            End If
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(5).Enabled = False
            
            C1Elastic1.Enabled = True
            C1Elastic2.Enabled = True
            
        Case "E"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "ÿ·» ’—› „ ⁄ÂœÌ‰ (  ⁄œÌ· )"
            Else
                Me.Caption = "Exchange Request (Edit)"
            End If
            Me.Cmd(8).Enabled = True
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(5).Enabled = False
            
            Me.XPBtnMove(0).Enabled = False
            Me.XPBtnMove(1).Enabled = False
            Me.XPBtnMove(2).Enabled = False
            Me.XPBtnMove(3).Enabled = False
            
            C1Elastic1.Enabled = False
          '  C1Elastic2.Enabled = True
            
    End Select

    Exit Sub
ErrTrap:
End Sub


Public Sub Retrive(Optional Lngid As Long = 0, Optional NoteID As Long = 0)

    'On Error GoTo ErrTrap
    If rs.RecordCount < 1 Then
        XPTxtCurrent.Caption = 0
        XPTxtCount.Caption = 0
        Exit Sub
    End If
    
      If rs.EOF Or rs.BOF Then
        Exit Sub
    Else
        If Lngid <> 0 Then
            rs.Find "ID =" & Lngid, , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Exit Sub
            End If
        End If
    
           If NoteID <> 0 Then
            rs.Find "NoteID =" & NoteID, , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Exit Sub
            End If
        End If
        
    End If
    
    
    Grid.Rows = Grid.FixedRows
    
    Me.TxtNoteID.Text = IIf(IsNull(rs.Fields("NoteID").value), "", rs.Fields("NoteID").value)
    Me.TxtNoteSerial.Text = IIf(IsNull(rs("NoteSerial").value), "", rs("NoteSerial").value)
     
   TxtRemarks.Text = IIf(IsNull(rs("Remarks").value), "", (rs("Remarks").value))
   Me.DCboUserName.BoundText = IIf(IsNull(rs("UserID").value), "", rs("UserID").value)



    txtid.Text = IIf(IsNull(rs("ID").value), "", val(rs("ID").value))
    TXTCode.Text = IIf(IsNull(rs("code").value), "", Trim(rs("code").value))
    DcDur.BoundText = IIf(IsNull(rs("DurationID").value), "", Trim(rs("DurationID").value))
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    Me.Date.value = IIf(IsNull(rs("Date").value), Date, rs("Date").value)
    Me.DateH.value = IIf(IsNull(rs("DateH").value), Date, rs("DateH").value)
    dcBranch.BoundText = IIf(IsNull(rs("BranchID").value), "", rs("BranchID").value)
    dcMontth.BoundText = IIf(IsNull(rs("Month").value), "", Trim(rs("Month").value))
    Text1.Text = IIf(IsNull(rs("DependID").value), "", rs("DependID").value)
     
     dcCustomer.BoundText = IIf(IsNull(rs("VendorID").value), "", Trim(rs("VendorID").value))
     dcMangerialAreaID.BoundText = IIf(IsNull(rs("MangerialAreaID").value), "", Trim(rs("MangerialAreaID").value))
     EntryDate.value = IIf(IsNull(rs("EntryDate").value), Date, rs("EntryDate").value)
     
       Dim ec As Boolean
       ec = IIf(IsNull(rs("EntryCreated")), 0, rs("EntryCreated"))
       
       If ec = True Then
                chkEntryCreated.value = 1
                
       Else
                chkEntryCreated.value = 0
       End If
       
       If TxtNoteSerial.Text <> "" Then
                       chkEntryCreated.value = 1
       Else
                       chkEntryCreated.value = 0
       End If
     Dim AllID As String
     
    AllID = IIf(IsNull(rs("AllID").value), "", rs("AllID").value)
    If AllID = "" Then
            Exit Sub
    End If
     
     
Dim str As String, j As Integer

'str = str & "   SELECT   TblAttributionInstallmentDivided.DDEmbarkDate , TblAttributionInstallmentDivided.DDEmbarkDateH  ,  TblAttributionInstallmentDivided.remark ,dbo.TblCustemers.RecordNo, dbo.TblAttributionContract.IDAC, dbo.TblAttributionContract.DurationID, dbo.TblDurations.Name AS DurationName,"
'str = str & "                     dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode, dbo.TblCustemers.CusID, dbo.TblAttributionContract.StartContractDate,"
'str = str & "                     dbo.TblAttributionContract.EndContractDate, dbo.TblAttributionContract.FromDate, dbo.TblDurations.FromDate AS DurFromDate, dbo.TblDurations.ToDate AS DurToDate,   dbo.TblDurations.FromDateH AS DurFromDateH ,"
'str = str & "                     dbo.TblVehicleAllocation_Details.Type, dbo.TblVehicleAllocation_Details.StudentCount, dbo.TblVehicleAllocation_Details.BoardNo, dbo.TblVehicleAllocation_Details.rate,"
'str = str & "                    dbo.TblVehicleAllocation_Details.Custom, dbo.TblVehicleAllocation_Details.DayRate, dbo.TblVehicleAllocation_Details.StudentCustom,"
'str = str & "                     dbo.TblVehicleAllocation_Details.CarID, dbo.TblDurations_Details.Name, dbo.TblVehicleAllocation_Details.ID, dbo.TblVehicleAllocation_Details.SchoolFileID,"
'str = str & "                    dbo.TblVendorCars.StopDeal, dbo.TblVendorCars.StopDate, dbo.TblVendorCars.StopDateH, dbo.TblAttributionContract.BranchID, dbo.TblCustemers.Account_Code,"
'str = str & "                    dbo.ACCOUNTS.Account_Serial, dbo.TblCustemers.IBAN, dbo.TblCustemers.BankAccount, dbo.TblDurations.Type AS DurationType,"
'str = str & "                    dbo.TblAttributionInstallmentDivided.MonthID, dbo.TblAttributionInstallmentDivided.ID AS DividID, dbo.TblAttributionContract.MangerialAreaID,"
'str = str & "                     dbo.TblManagerialArea.Name AS MAName   , TblAttributionContract.cityID , TblAttributionContract.vendorid , dbo.TblVehicleAllocation_Details.schoolfileID "
'str = str & "   FROM     dbo.TblVendorCars RIGHT OUTER JOIN"
'str = str & "                    dbo.ACCOUNTS INNER JOIN"
'str = str & "                    dbo.TblAttributionContract INNER JOIN"
'str = str & "                   dbo.TblCustemers ON dbo.TblAttributionContract.VendorID = dbo.TblCustemers.CusID INNER JOIN"
'str = str & "                     dbo.TblVehicleAllocation_Details ON dbo.TblAttributionContract.IDAC = dbo.TblVehicleAllocation_Details.IDVA ON"
'str = str & "                    dbo.ACCOUNTS.Account_Code = dbo.TblCustemers.Account_Code INNER JOIN"
'str = str & "                   dbo.TblAttributionInstallmentDivided ON dbo.TblVehicleAllocation_Details.ID = dbo.TblAttributionInstallmentDivided.DetailsID INNER JOIN"
'str = str & "                 dbo.TblDurations_Details ON dbo.TblAttributionInstallmentDivided.MonthID = dbo.TblDurations_Details.ID LEFT OUTER JOIN"
'str = str & "                   dbo.TblManagerialArea ON dbo.TblAttributionContract.MangerialAreaID = dbo.TblManagerialArea.ID ON"
'str = str & "                     dbo.TblVendorCars.ID = dbo.TblVehicleAllocation_Details.CarID LEFT OUTER JOIN"
'str = str & "                   dbo.TblDurations ON dbo.TblAttributionContract.DurationID = dbo.TblDurations.ID"
'str = str & "   Where (dbo.TblVehicleAllocation_Details.Type = 3)"

str = "SELECT  DISTINCT dbo.TblAttributionInstallmentDivided.PayMentPayed ,   dbo.TblAttributionInstallmentDivided.DDEmbarkDate, dbo.TblAttributionInstallmentDivided.DDEmbarkDateH, dbo.TblAttributionInstallmentDivided.remark, "
str = str & "                        dbo.TblCustemers.RecordNo, dbo.TblAttributionContract.IDAC, dbo.TblAttributionContract.DurationID, dbo.TblDurations.Name AS DurationName,"
str = str & "                  TblAttributionInstallmentDivided.TypeVAT,TblAttributionInstallmentDivided.Vatyo,TblAttributionInstallmentDivided.TotalNet,TblAttributionInstallmentDivided.Vat2,"
str = str & "                        dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode, dbo.TblCustemers.CusID, dbo.TblAttributionContract.StartContractDate,"
str = str & "                        dbo.TblAttributionContract.EndContractDate, dbo.TblAttributionContract.FromDate, dbo.TblDurations.FromDate AS DurFromDate, dbo.TblDurations.ToDate AS DurToDate,"
str = str & "                         dbo.TblDurations.FromDateH AS DurFromDateH, dbo.TblVehicleAllocation_Details.Type, dbo.TblVehicleAllocation_Details.StudentCount,"
str = str & "                        dbo.TblVehicleAllocation_Details.BoardNo, dbo.TblVehicleAllocation_Details.rate, dbo.TblVehicleAllocation_Details.Custom, dbo.TblVehicleAllocation_Details.DayRate,"
str = str & "                        dbo.TblVehicleAllocation_Details.StudentCustom, dbo.TblVehicleAllocation_Details.CarID, dbo.TblDurations_Details.Name, dbo.TblVehicleAllocation_Details.ID,"
str = str & "                        dbo.TblVehicleAllocation_Details.SchoolFileID, dbo.TblAttributionContract.StopDeal, dbo.TblAttributionContract.StopDate, dbo.TblAttributionContract.StopDateH,"
str = str & "                        dbo.TblAttributionContract.BranchID, dbo.TblCustemers.Account_Code, dbo.ACCOUNTS.Account_Serial, dbo.TblCustemers.IBAN, dbo.TblCustemers.BankAccount,"
str = str & "                        dbo.TblDurations.Type AS DurationType, dbo.TblAttributionInstallmentDivided.MonthID, dbo.TblAttributionInstallmentDivided.ID AS DividID,"
str = str & "                        dbo.TblAttributionContract.MangerialAreaID, dbo.TblManagerialArea.Name AS MAName, dbo.TblAttributionContract.CityID, dbo.TblAttributionContract.VendorID,"
str = str & "                        dbo.TblSchooleFile.Name AS SchoolnMae, dbo.TblSchooleFile.Namee AS SchoolnMaeE"
str = str & "  FROM         dbo.TblManagerialArea RIGHT OUTER JOIN"
                      str = str & "  dbo.ACCOUNTS INNER JOIN"
str = str & "                        dbo.TblAttributionContract INNER JOIN"
str = str & "                        dbo.TblCustemers ON dbo.TblAttributionContract.VendorID = dbo.TblCustemers.CusID INNER JOIN"
str = str & "                        dbo.TblVehicleAllocation_Details ON dbo.TblAttributionContract.IDAC = dbo.TblVehicleAllocation_Details.IDVA ON"
str = str & "                        dbo.ACCOUNTS.Account_Code = dbo.TblCustemers.Account_Code INNER JOIN"
str = str & "                        dbo.TblAttributionInstallmentDivided ON dbo.TblVehicleAllocation_Details.ID = dbo.TblAttributionInstallmentDivided.DetailsID INNER JOIN"
str = str & "                        dbo.TblDurations_Details ON dbo.TblAttributionInstallmentDivided.MonthID = dbo.TblDurations_Details.ID LEFT OUTER JOIN"
str = str & "                        dbo.TblSchooleFile ON dbo.TblVehicleAllocation_Details.SchoolFileID = dbo.TblSchooleFile.ID ON"
str = str & "                        dbo.TblManagerialArea.ID = dbo.TblAttributionContract.MangerialAreaID LEFT OUTER JOIN"
str = str & "                        dbo.TblVendorCars ON dbo.TblVehicleAllocation_Details.CarID = dbo.TblVendorCars.ID LEFT OUTER JOIN"
str = str & "                        dbo.TblDurations ON dbo.TblAttributionContract.DurationID = dbo.TblDurations.ID"
   str = str & "   Where (dbo.TblVehicleAllocation_Details.Type = 3)"
    str = str & "  and  TblAttributionInstallmentDivided.id   in  ( " & AllID & "  )"
     
     

    If DcDur.BoundText <> "" Then
            str = str & "  and  TblAttributionContract.DurationID  = " & val(DcDur.BoundText)
    End If
      
    If dcMontth.BoundText <> "" Then
            str = str & "      and  TblDurations_Details.ID = " & val(dcMontth.BoundText)
    End If
       
    If dcBranch.BoundText <> "" Then
            str = str & "      and  TblAttributionContract.BranchID  = " & val(dcBranch.BoundText)
    End If
     
      str = str & "      order by dbo.TblAttributionContract.IDAC , dbo.TblVehicleAllocation_Details.BoardNo  "
      
    Set RsTemp = New ADODB.Recordset
    RsTemp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
    Dim V As Integer, H As Integer, WD
    Dim tot As Integer, daycount As Integer, SchoolFileID As Integer
           
    If RsTemp.RecordCount > 0 Then
            RsTemp.MoveFirst
            With Grid
             ProgressBar1.Max = RsTemp.RecordCount
             Grid.Rows = Grid.FixedRows + RsTemp.RecordCount
             For j = Grid.FixedRows To Grid.Rows - 1
                    ProgressBar1.value = j - 2
                    .TextMatrix(j, .ColIndex("Serial")) = j - 1
                    .TextMatrix(j, .ColIndex("Status")) = 1
                    
                                        If IsNull(RsTemp("PayMentPayed").value) Then
                              .TextMatrix(j, .ColIndex("PayMentPayed")) = ""
                    Else
                            If RsTemp("PayMentPayed").value = 0 Then
                              .TextMatrix(j, .ColIndex("PayMentPayed")) = ""
                            Else
                             .TextMatrix(j, .ColIndex("PayMentPayed")) = "‰⁄„"
                            End If
                    
                    End If

                    .TextMatrix(j, .ColIndex("fullcode")) = IIf(IsNull(RsTemp("Fullcode").value), "", RsTemp("Fullcode").value)
                    .TextMatrix(j, .ColIndex("cusname")) = IIf(IsNull(RsTemp("CusName").value), "", RsTemp("CusName").value)
                    '.TextMatrix(j, .ColIndex("InstallmentNo")) = IIf(IsNull(RsTemp("InstallmentNo").value), "", RsTemp("InstallmentNo").value)
                    .TextMatrix(j, .ColIndex("ID")) = IIf(IsNull(RsTemp("ID").value), "", RsTemp("ID").value)
                    .TextMatrix(j, .ColIndex("DividID")) = IIf(IsNull(RsTemp("DividID").value), "", RsTemp("DividID").value)
                    .TextMatrix(j, .ColIndex("recordno")) = IIf(IsNull(RsTemp("recordno").value), "", RsTemp("recordno").value)
                    .TextMatrix(j, .ColIndex("IDAC")) = IIf(IsNull(RsTemp("IDAC").value), "", RsTemp("IDAC").value)
                    .TextMatrix(j, .ColIndex("CusID")) = IIf(IsNull(RsTemp("CusID").value), "", RsTemp("CusID").value)
                    '   .TextMatrix(j, .ColIndex("MonthID")) = IIf(IsNull(RsTemp("MonthID").value), "", RsTemp("MonthID").value)
                    .TextMatrix(j, .ColIndex("StartContractDate")) = IIf(IsNull(RsTemp("StartContractDate").value), "", RsTemp("StartContractDate").value)
                    .TextMatrix(j, .ColIndex("EndContractDate")) = IIf(IsNull(RsTemp("EndContractDate").value), "", RsTemp("EndContractDate").value)
                    .TextMatrix(j, .ColIndex("FromDate")) = IIf(IsNull(RsTemp("FromDate").value), "", RsTemp("FromDate").value)
                    .TextMatrix(j, .ColIndex("Car")) = IIf(IsNull(RsTemp("BoardNo").value), "", RsTemp("BoardNo").value)
                    '   .TextMatrix(j, .ColIndex("CarID")) = IIf(IsNull(RsTemp("CarID").value), "", RsTemp("CarID").value)
                    .TextMatrix(j, .ColIndex("Account_Code")) = IIf(IsNull(RsTemp("Account_Code").value), "", RsTemp("Account_Code").value)
                    .TextMatrix(j, .ColIndex("Account_Serial")) = IIf(IsNull(RsTemp("Account_Serial").value), "", RsTemp("Account_Serial").value)
                    .TextMatrix(j, .ColIndex("IBAN")) = IIf(IsNull(RsTemp("IBAN").value), "", RsTemp("IBAN").value)
                    .TextMatrix(j, .ColIndex("BankAccount")) = IIf(IsNull(RsTemp("BankAccount").value), "", RsTemp("BankAccount").value)
                    Dim stopDeal  As Boolean, DateStop As Date, DateStopH As String, Typ As Integer, sf As Integer
                    stopDeal = IIf(IsNull(RsTemp("StopDeal").value), False, RsTemp("StopDeal").value)
                    DateStop = IIf(IsNull(RsTemp("StopDate").value), 0, RsTemp("StopDate").value)
                    DateStopH = IIf(IsNull(RsTemp("StopDateH").value), 0, RsTemp("StopDateH").value)
                    Typ = IIf(IsNull(RsTemp("DurationType").value), 0, RsTemp("DurationType").value)
                    sf = IIf(IsNull(RsTemp("SchoolFileID").value), 0, RsTemp("SchoolFileID").value)
                    
                    Dim DDEmbarkDate As Date
                    Dim DDEmbarkDateH As String
                    DDEmbarkDate = IIf(IsNull(RsTemp("DDEmbarkDate").value), RsTemp("DurFromDate").value, RsTemp("DDEmbarkDate").value)
                    DDEmbarkDateH = IIf(IsNull(RsTemp("DDEmbarkDateH").value), RsTemp("DurFromDateH").value, RsTemp("DDEmbarkDateH").value)
                                        
                    
                    
                     
                    Dim exc As Integer
    
                    If Not (IsNull(RsTemp("IDAC").value) Or IsNull(RsTemp("MonthID").value)) Then
                            GetVac (RsTemp("IDAC").value), (RsTemp("MonthID").value), RsTemp("DurationID").value, sf, tot, daycount, stopDeal, DateStop, DateStopH, Typ, DDEmbarkDate, DDEmbarkDateH
                            H = GetHold(RsTemp("ID").value, (RsTemp("MonthID").value), DDEmbarkDate, DDEmbarkDateH, stopDeal, DateStop, DateStopH, Typ)
                            
                            GetDeducts RsTemp("IDAC").value, val(DcDur.BoundText), RsTemp("MonthID").value, IIf(IsNull(RsTemp("CarID").value), 0, RsTemp("CarID").value), j, _
                             IIf(IsNull(RsTemp("schoolfileID").value), 0, RsTemp("schoolfileID").value)


                                exc = calExcepDays(val(DcDur.BoundText), RsTemp("MonthID").value, _
                            IIf(IsNull(RsTemp("vendorid").value), -1, RsTemp("vendorid").value), _
                            IIf(IsNull(RsTemp("CityID").value), -1, RsTemp("CityID").value), _
                            IIf(IsNull(RsTemp("MangerialAreaID").value), -1, RsTemp("MangerialAreaID").value))
                    End If
                    
                    '////////////////
                    .TextMatrix(j, .ColIndex("adddays")) = exc
                     If Not IsNull(RsTemp("MonthID").value) Then
                            WD = GetMonthDays(RsTemp("ID").value, RsTemp("MonthID").value, DDEmbarkDate, _
                            DDEmbarkDateH, _
                            stopDeal, DateStop, DateStopH, Typ)
                     End If
                     
                    .TextMatrix(j, .ColIndex("VacDay")) = daycount
                    .TextMatrix(j, .ColIndex("WorkDay")) = WD - H
                     
                     Dim dayrate As Double
                     dayrate = IIf(IsNull(RsTemp("DayRate").value), 0, RsTemp("DayRate").value)
                      
                    
                     If Not (IsNull(RsTemp("IDAC").value)) Then
                           .TextMatrix(j, .ColIndex("VacValue")) = daycount * dayrate
                     End If
                                       
                     .TextMatrix(j, .ColIndex("DayRate")) = dayrate
                   
                   ' .TextMatrix(j, .ColIndex("Value")) = dayrate * (WD - H)
                   
                   .TextMatrix(j, .ColIndex("schoolName")) = IIf(IsNull(RsTemp("SchoolnMae").value), "", RsTemp("SchoolnMae").value)
                   
                    .TextMatrix(j, .ColIndex("Value")) = dayrate * (WD + exc - H)
                    .TextMatrix(j, .ColIndex("MA")) = IIf(IsNull(RsTemp("MAName").value), "", RsTemp("MAName").value)
                    .TextMatrix(j, .ColIndex("remark")) = IIf(IsNull(RsTemp("remark").value), "", RsTemp("remark").value)
                    
                    .TextMatrix(j, .ColIndex("TypeVAT")) = IIf(IsNull(RsTemp("TypeVAT").value), "", RsTemp("TypeVAT").value)
                    .TextMatrix(j, .ColIndex("Vatyo")) = IIf(IsNull(RsTemp("Vatyo").value), "", RsTemp("Vatyo").value)
                    .TextMatrix(j, .ColIndex("Net")) = val(RsTemp!TotalNet & "") + val(RsTemp!Vat2 & "")
                    
                    .TextMatrix(j, .ColIndex("Vat2")) = IIf(IsNull(RsTemp("Vat2").value), "", RsTemp("Vat2").value)
                    

                    
                            .Row = j
                            .Col = .ColIndex("IDAC")
                             .ShowCell j, .ColIndex("IDAC")
                            
'                             .SetFocus


                    RsTemp.MoveNext
             Next
            End With
    End If
calculation True
RemoveStopedCar
         
    
    Exit Sub
ErrTrap:
End Sub

Private Sub txtRecordNo_Change()
Dim val1, val2, CusID As String, Fullcode As String
If txtRecordno.Text = "" Then Exit Sub
Dim str As String
    str = " select * From TblCustemers where Type=2  and recordno = '" & txtRecordno.Text & "'"
    Set Rs_Temp = New ADODB.Recordset
    Rs_Temp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If Rs_Temp.RecordCount > 0 Then
        Rs_Temp.MoveFirst '
         CusID = IIf(IsNull(Rs_Temp("cusID").value), "", Rs_Temp("cusID").value)
        Fullcode = IIf(IsNull(Rs_Temp("fullcode").value), "", Rs_Temp("fullcode").value)
     Else
        dcCustomer.BoundText = ""
        txtfullcode.Text = ""
    End If
    
   dcCustomer.BoundText = CusID
   txtfullcode.Text = Fullcode
End Sub

Private Sub XPBtnMove_Click(Index As Integer)

    'On Error GoTo ErrTrap
    If Me.TxtModFlg.Text = "N" Then
        clear_all Me
        Me.TxtModFlg.Text = "R"
        XPBtnMove_Click (1)
    End If

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

 
Private Sub SaveData()
    Dim Msg As String
    Dim StrSQL As String
    Dim RsTemp As New ADODB.Recordset
    Dim RsTempM As New ADODB.Recordset
    Dim BeginTrans As Boolean
 On Error GoTo ErrTrap


    If Me.TxtModFlg.Text <> "R" Then
    
        If DcDur.BoundText = "" Then
            MsgBox "„‰ ›÷·ﬂ √œŒ· «”„ «·› —… ", vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
'            DcDur.SetFocus
            Exit Sub
        End If
    
    
          If dcBranch.BoundText = "" Then
            MsgBox "„‰ ›÷·ﬂ «·›—⁄ ", vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            dcBranch.SetFocus
            Exit Sub
        End If
        
        If checkedRow = False Then
                MsgBox ("«Œ — «·œ›⁄«  «Ê·«")
                Exit Sub
        End If
        
        
        Select Case Me.TxtModFlg.Text
            Case "N"
                 rs.AddNew
                 txtid.Text = CStr(new_id("TblExchangeRequest", "ID", "", True))
            Case "E"
               

        End Select

        Cn.BeginTrans
        BeginTrans = True
          
         If TxtModFlg.Text = "E" Then
                 Cancel_Paid
                 StrSQL = "Delete From TblExchangeReques_Detailst Where HID =" & val(Me.txtid.Text)
                 Cn.Execute StrSQL, , adExecuteNoRecords
                 StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(Me.TxtNoteID.Text)
                 Cn.Execute StrSQL, , adExecuteNoRecords
         End If
          
        rs("ID").value = val(txtid.Text)
        rs("Code").value = Trim(TXTCode.Text)
       ' rs("ExchangeType").value = IIf(cbType.ListIndex = -1, Null, cbType.ListIndex)
        rs("Remarks").value = TxtRemarks.Text
        
        rs("DurationID").value = val(DcDur.BoundText)
        rs("DurationName").value = DcDur.Text
        
        rs("Month").value = dcMontth.BoundText
        rs("Date").value = Me.Date.value
        rs("DateH").value = Me.DateH.value
        rs("BranchID").value = dcBranch.BoundText
        rs("DependID").value = IIf(Text1.Text = "", Null, val(Text1.Text))
         'Me.DCboUserName.BoundText = IIf(IsNull(rs("UserID").value), "", rs("UserID").value)
        rs("UserID").value = IIf(DCboUserName.BoundText = "", Null, val(DCboUserName.BoundText))

        rs("VendorID").value = IIf(dcCustomer.BoundText = "", Null, val(dcCustomer.BoundText))
        rs("MangerialAreaID").value = IIf(dcMangerialAreaID.BoundText = "", Null, val(dcMangerialAreaID.BoundText))
        
       If chkEntryCreated.value = 1 Then
                rs("EntryCreated") = True
       Else
                rs("EntryCreated") = False
       End If
       
        rs("CurrentUser").value = Null
        rs("EntryDate").value = Me.EntryDate.value
        rs.update
        
        
       
       
 Dim i As Integer, AllID As String
       With Grid
            For i = .FixedRows To .Rows - 1
               If .Cell(flexcpChecked, i, .ColIndex("Status")) = flexChecked And .TextMatrix(i, .ColIndex("ID")) <> "" Then
                     
                                AllID = AllID & IIf(.TextMatrix(i, .ColIndex("DividID")) = "", " ", ",  " & .TextMatrix(i, .ColIndex("DividID")))
             
                End If
            Next
        End With
          
        rs("AllID").value = mId$(AllID, 2)
        rs.update
        
        
        
       With Grid
            For i = .FixedRows To .Rows - 1
               If .Cell(flexcpChecked, i, .ColIndex("Status")) = flexChecked And val(.TextMatrix(i, .ColIndex("DividID"))) <> 0 Then
                        Set RsTemp = New ADODB.Recordset
                        Dim m As String
                        m = "  select * from TblAttributionInstallmentDivided  where MonthID = " & dcMontth.BoundText & " and  id =  " & val(.TextMatrix(i, .ColIndex("DividID")))
                        RsTemp.Open m, Cn, adOpenStatic, adLockOptimistic, adCmdText
                        If RsTemp.RecordCount > 0 Then
                        RsTemp.MoveFirst
                                RsTemp("TotalValue").value = val(.TextMatrix(i, .ColIndex("Net")))
                                RsTemp("remark").value = (.TextMatrix(i, .ColIndex("remark"))) & "1"
                                RsTemp("RE_Paid").value = 1
                                RsTemp("REID").value = val(txtid.Text)
                                RsTemp("TypeVAT").value = (.TextMatrix(i, .ColIndex("TypeVAT")))
                                RsTemp("Vatyo").value = val(.TextMatrix(i, .ColIndex("Vatyo")))
                                RsTemp("TotalNet").value = val(.TextMatrix(i, .ColIndex("TotalNet")))
                                RsTemp("Vat2").value = val(.TextMatrix(i, .ColIndex("Vat2")))
                                





                                RsTemp.update
                                RsTemp.Close
                                
                        End If
                End If
            Next
        End With
        
       
       
       Set RsTemp = New ADODB.Recordset
       RsTemp.Open "TblExchangeReques_Detailst", Cn, adOpenStatic, adLockOptimistic, adCmdTable
       
      With Grid
   
            For i = .FixedRows To .Rows - 1
               If .Cell(flexcpChecked, i, .ColIndex("Status")) = flexChecked And .TextMatrix(i, .ColIndex("ID")) <> "" Then
                        RsTemp.AddNew
                        RsTemp("ID").value = CStr(new_id("TblExchangeReques_Detailst", "ID", "", True))
                        RsTemp("HID").value = val(txtid.Text)
                        RsTemp("CusID").value = .TextMatrix(i, .ColIndex("CusID"))
                        RsTemp("IDAC").value = IIf(.TextMatrix(i, .ColIndex("IDAC")) = "", Null, .TextMatrix(i, .ColIndex("IDAC")))
                        RsTemp("InsID").value = .TextMatrix(i, .ColIndex("ID"))
                        RsTemp("fullcode").value = .TextMatrix(i, .ColIndex("fullcode"))
                        RsTemp("cusname").value = .TextMatrix(i, .ColIndex("cusname"))
                      ' RsTemp("InsNo").value = .TextMatrix(i, .ColIndex("InstallmentNo"))
                        RsTemp("Value").value = .TextMatrix(i, .ColIndex("Value"))
                        'RsTemp("carid").value = IIf(.TextMatrix(i, .ColIndex("CarID")) = "", Null, .TextMatrix(i, .ColIndex("CarID")))
                        RsTemp("boardno").value = .TextMatrix(i, .ColIndex("Car"))
                        RsTemp("dayvalue").value = val(.TextMatrix(i, .ColIndex("DayRate")))
                        RsTemp("absenceDays").value = val(.TextMatrix(i, .ColIndex("AbsenceCount")))
                        RsTemp("absenceValue").value = val(.TextMatrix(i, .ColIndex("Avalue")))
                        RsTemp("ContractBDate").value = .TextMatrix(i, .ColIndex("StartContractDate"))
                        RsTemp("ContractEDate").value = .TextMatrix(i, .ColIndex("EndContractDate"))
                        RsTemp("ContractDate").value = .TextMatrix(i, .ColIndex("FromDate"))
                         RsTemp("d1").value = IIf(.TextMatrix(i, .ColIndex("d1")) = "", 0, .TextMatrix(i, .ColIndex("d1")))
                                       RsTemp("DividID").value = IIf(.TextMatrix(i, .ColIndex("DividID")) = "", 0, .TextMatrix(i, .ColIndex("DividID")))
                                       
                         RsTemp("d2").value = IIf(.TextMatrix(i, .ColIndex("d2")) = "", 0, .TextMatrix(i, .ColIndex("d2")))
                         RsTemp("d3").value = IIf(.TextMatrix(i, .ColIndex("d3")) = "", 0, .TextMatrix(i, .ColIndex("d3")))
                         RsTemp("d4").value = IIf(.TextMatrix(i, .ColIndex("d4")) = "", 0, .TextMatrix(i, .ColIndex("d4")))
                         RsTemp("d5").value = IIf(.TextMatrix(i, .ColIndex("d5")) = "", 0, .TextMatrix(i, .ColIndex("d5")))
                         RsTemp("d6").value = IIf(.TextMatrix(i, .ColIndex("d6")) = "", 0, .TextMatrix(i, .ColIndex("d6")))
                         RsTemp("d7").value = IIf(.TextMatrix(i, .ColIndex("d7")) = "", 0, .TextMatrix(i, .ColIndex("d7")))
                         RsTemp("d8").value = IIf(.TextMatrix(i, .ColIndex("d8")) = "", 0, .TextMatrix(i, .ColIndex("d8")))
                         RsTemp("d9").value = IIf(.TextMatrix(i, .ColIndex("d9")) = "", 0, .TextMatrix(i, .ColIndex("d9")))
                         RsTemp("d10").value = IIf(.TextMatrix(i, .ColIndex("d10")) = "", 0, .TextMatrix(i, .ColIndex("d10")))
                         RsTemp("d11").value = IIf(.TextMatrix(i, .ColIndex("d11")) = "", 0, .TextMatrix(i, .ColIndex("d11")))
                         RsTemp("d12").value = IIf(.TextMatrix(i, .ColIndex("d12")) = "", 0, .TextMatrix(i, .ColIndex("d12")))
                         RsTemp("d13").value = IIf(.TextMatrix(i, .ColIndex("d13")) = "", 0, .TextMatrix(i, .ColIndex("d13")))
                         RsTemp("d14").value = IIf(.TextMatrix(i, .ColIndex("d14")) = "", 0, .TextMatrix(i, .ColIndex("d14")))
                         RsTemp("d15").value = IIf(.TextMatrix(i, .ColIndex("d15")) = "", 0, .TextMatrix(i, .ColIndex("d15")))
                         RsTemp("d16").value = IIf(.TextMatrix(i, .ColIndex("d16")) = "", 0, .TextMatrix(i, .ColIndex("d16")))
                         RsTemp("d17").value = IIf(.TextMatrix(i, .ColIndex("d17")) = "", 0, .TextMatrix(i, .ColIndex("d17")))
                         RsTemp("d18").value = IIf(.TextMatrix(i, .ColIndex("d18")) = "", 0, .TextMatrix(i, .ColIndex("d18")))
                         RsTemp("d19").value = IIf(.TextMatrix(i, .ColIndex("d19")) = "", 0, .TextMatrix(i, .ColIndex("d19")))
                         RsTemp("d20").value = IIf(.TextMatrix(i, .ColIndex("d20")) = "", 0, .TextMatrix(i, .ColIndex("d20")))
  
                        RsTemp("DB1").value = .TextMatrix(1, .ColIndex("d1"))
                        RsTemp("DB2").value = .TextMatrix(1, .ColIndex("d2"))
                        RsTemp("DB3").value = .TextMatrix(1, .ColIndex("d3"))
                        RsTemp("DB4").value = .TextMatrix(1, .ColIndex("d4"))
                        RsTemp("DB5").value = .TextMatrix(1, .ColIndex("d5"))
                        RsTemp("DB6").value = .TextMatrix(1, .ColIndex("d6"))
                        RsTemp("DB7").value = .TextMatrix(1, .ColIndex("d7"))
                        RsTemp("DB8").value = .TextMatrix(1, .ColIndex("d8"))
                        RsTemp("DB9").value = .TextMatrix(1, .ColIndex("d9"))
                        RsTemp("DB10").value = .TextMatrix(1, .ColIndex("d10"))
                        RsTemp("DB11").value = .TextMatrix(1, .ColIndex("d11"))
                        RsTemp("DB12").value = .TextMatrix(1, .ColIndex("d12"))
                        RsTemp("DB13").value = .TextMatrix(1, .ColIndex("d13"))
                        RsTemp("DB14").value = .TextMatrix(1, .ColIndex("d14"))
                        RsTemp("DB15").value = .TextMatrix(1, .ColIndex("d15"))
                       RsTemp("DB16").value = .TextMatrix(1, .ColIndex("d16"))
                       RsTemp("DB17").value = .TextMatrix(1, .ColIndex("d17"))
                       RsTemp("DB18").value = .TextMatrix(1, .ColIndex("d18"))
                       RsTemp("DB19").value = .TextMatrix(1, .ColIndex("d19"))
                       RsTemp("DB20").value = .TextMatrix(1, .ColIndex("d20"))
                        
                        
                        RsTemp("TypeVAT").value = .TextMatrix(i, .ColIndex("TypeVAT"))
                        RsTemp("Vatyo").value = .TextMatrix(i, .ColIndex("Vatyo"))
                        RsTemp("TotalNet").value = .TextMatrix(i, .ColIndex("TotalNet"))
                        RsTemp("Vat2").value = .TextMatrix(i, .ColIndex("Vat2"))
  
                         RsTemp("Total_deduct").value = .TextMatrix(i, .ColIndex("Total"))
                         RsTemp("Net").value = .TextMatrix(i, .ColIndex("Net"))
  
                         RsTemp("wokdays").value = IIf(.TextMatrix(i, .ColIndex("WorkDay")) = "", 0, .TextMatrix(i, .ColIndex("WorkDay")))
                         RsTemp("stopdays").value = IIf(.TextMatrix(i, .ColIndex("VacDay")) = "", 0, .TextMatrix(i, .ColIndex("VacDay")))
                         RsTemp("stopvalue").value = IIf(.TextMatrix(i, .ColIndex("VacValue")) = "", 0, .TextMatrix(i, .ColIndex("VacValue")))
                         RsTemp("BankAccount").value = IIf(.TextMatrix(i, .ColIndex("BankAccount")) = "", "", .TextMatrix(i, .ColIndex("BankAccount")))
                         RsTemp("Account_Code").value = IIf(.TextMatrix(i, .ColIndex("Account_Code")) = "", "", .TextMatrix(i, .ColIndex("Account_Code")))
                         RsTemp("Account_Serial").value = IIf(.TextMatrix(i, .ColIndex("Account_Serial")) = "", "", .TextMatrix(i, .ColIndex("Account_Serial")))
   
                        RsTemp("remark").value = IIf(.TextMatrix(i, .ColIndex("remark")) = "", "", .TextMatrix(i, .ColIndex("remark")))
                        RsTemp.update
                End If
            Next
        End With
'



        Cn.CommitTrans
        BeginTrans = False
        XPTxtCurrent.Caption = rs.AbsolutePosition
        XPTxtCount.Caption = rs.RecordCount
        
        
       CuurentLogdata

        Select Case Me.TxtModFlg.Text

            Case "N"
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "  „ Õ›Ÿ «·»Ì«‰«    " & CHR(13)
                    Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
                Else
                    Msg = "Saved" & CHR(13)
                    Msg = Msg + "Do you want enter another One"
                End If

                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                    Cmd_Click (0)
                    Exit Sub
                End If
            
            Case "E"
        
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Else
                    MsgBox "Saved", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                End If

        End Select

        TxtModFlg.Text = "R"
    End If

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

Private Sub Undo()
    On Error GoTo ErrTrap

    Select Case TxtModFlg.Text

        Case "N"
            clear_all Me
            Me.TxtModFlg.Text = "R"
            XPBtnMove_Click (1)

        Case "E"
            rs.Find "ID='" & val(txtid.Text) & "'", , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Me.TxtModFlg.Text = "R"
                Exit Sub
            End If

            Retrive
            Me.TxtModFlg.Text = "R"
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Function checkedRow() As Boolean
Dim i As Integer, Check As Boolean
For i = 1 To Grid.Rows - 1
        If Grid.TextMatrix(i, Grid.ColIndex("status")) <> "" Then
                If Grid.TextMatrix(i, Grid.ColIndex("status")) <> 0 Then
                        checkedRow = True
                        Exit Function
                End If
        End If
Next
checkedRow = False
End Function


Private Sub Del_Company()
    Dim Msg As String
    Dim StrSQL As String
    Dim RsTemp As New ADODB.Recordset
    Dim StrAccountCode As String
    Dim StrAccountCode1 As String
    Dim ParentAccount As String
    '
 
    On Error GoTo ErrTrap
            
        If txtid.Text <> "" Then

    
        Msg = "”Ì „ Õ–› »Ì«‰«   —ﬁ„ " & CHR(13)
        Msg = Msg + (txtid.Text) & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            
    
            If Not rs.RecordCount < 1 Then
                 
                StrSQL = "delete From TblExchangeReques_Detailst where  HID =" & val(txtid.Text)
                Cn.Execute StrSQL, , adExecuteNoRecords
                Cancel_Paid
            
                StrSQL = "delete From TblExchangeRequest where  ID =" & val(txtid.Text)
                Cn.Execute StrSQL, , adExecuteNoRecords
                
                          StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(Me.TxtNoteID.Text)
                          Cn.Execute StrSQL, , adExecuteNoRecords
                           Grid.Rows = Grid.FixedRows

                  
                  CuurentLogdata ("D")
                   StrSQL = "SELECT  *  From TblExchangeRequest "
                   rs.Close
                   rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                

                If rs.RecordCount < 1 Then
                    clear_all Me
                    TxtModFlg_Change
                    XPTxtCurrent.Caption = 0
                    XPTxtCount.Caption = 0
                    Grid.Rows = Grid.FixedRows
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
    'If Err.Number = -2147217887 Then
    Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & CHR(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «·Œ“‰… "
    Msg = Msg & CHR(13) & Err.description
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
    rs.CancelUpdate
    'End If
End Sub


Private Sub AddTip()
    Dim Wrap As String
    On Error GoTo ErrTrap
    Set TTP = New clstooltip
    Wrap = CHR(13) + CHR(10)

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«  Œ“‰… ÃœÌœ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  «·Œ“‰…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  «·Œ“‰… «·ÃœÌœ" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  Â–« «·Œ“‰…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(5), "»ÕÀ ..." & Wrap & "···»ÕÀ ⁄‰ Œ“‰…" & Wrap & "Ì‰ÿ»ﬁ ⁄·ÌÂ ‘—Êÿ „⁄Ì‰…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  «·Œ“‰", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
    '    .AddControl CmdHelp, "„”«⁄œ… ..." & Wrap & "·· ⁄—› ⁄·Ï ÊŸÌ›… Â–Â «·‰«›–…" & Wrap & "ÊﬂÌ›Ì… «· ⁄«„· „⁄Â«" & Wrap & "≈÷€ÿ Â‰«" & Wrap, True
    End With

    Exit Sub
ErrTrap:
End Sub

Private Sub XPTxtBoxName_GotFocus()

    SwitchKeyboardLang LANG_ARABIC
End Sub



Function print_report(Optional NoteSerial As Integer)
    
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String



MySQL = MySQL & "  SELECT tb1.remark ,  tb1.remarks , TB1.BankAccount ,TB1.DurName, TB1.HID, TB1.CusID, TB1.InsID, TB1.InsNo, TB1.CusName, TB1.FullCode, TB1.Value, TB1.Total_Deduct, TB1.net, TB1.wokdays, TB1.stopdays, TB1.stopvalue,"
MySQL = MySQL & "                    TB1.Account_Serial, TB1.Account_Code, TB1.carid, TB1.boardno, TB1.dayvalue, TB1.absenceDays, TB1.absenceValue, TB1.ContractBDate, TB1.ContractEDate,"
MySQL = MySQL & "                    TB1.ContractDate, TB1.DurationID, TB1.Month, TB1.MonthName, TB1.Date, TB1.DateH, TB1.branch_name, TB1.branch_namee, TB1.MainTblID, TB1.IDAC, TB1.DID, db.ID,"
MySQL = MySQL & "                    db.VName, db.value AS VValue"
MySQL = MySQL & "  FROM     (SELECT  TblExchangeRequest.remarks , TblExchangeReques_detailst.remark, TblExchangeReques_Detailst.BankAccount , dbo.TblDurations.Name AS DurName, dbo.TblExchangeReques_Detailst.HID, dbo.TblExchangeReques_Detailst.CusID, dbo.TblExchangeReques_Detailst.InsID,"
MySQL = MySQL & "                                      dbo.TblExchangeReques_Detailst.InsNo, dbo.TblExchangeReques_Detailst.CusName, dbo.TblExchangeReques_Detailst.FullCode,"
MySQL = MySQL & "                                      dbo.TblExchangeReques_Detailst.Value, dbo.TblExchangeReques_Detailst.Total_Deduct, dbo.TblExchangeReques_Detailst.net,"
MySQL = MySQL & "                                      dbo.TblExchangeReques_Detailst.wokdays, dbo.TblExchangeReques_Detailst.stopdays, dbo.TblExchangeReques_Detailst.stopvalue,"
MySQL = MySQL & "                                      dbo.TblExchangeReques_Detailst.Account_Serial, dbo.TblExchangeReques_Detailst.Account_Code, dbo.TblExchangeReques_Detailst.carid,"
 MySQL = MySQL & "                                     dbo.TblExchangeReques_Detailst.boardno, dbo.TblExchangeReques_Detailst.dayvalue, dbo.TblExchangeReques_Detailst.absenceDays,"
MySQL = MySQL & "                                      dbo.TblExchangeReques_Detailst.absenceValue, dbo.TblExchangeReques_Detailst.ContractBDate, dbo.TblExchangeReques_Detailst.ContractEDate,"
MySQL = MySQL & "                                      dbo.TblExchangeReques_Detailst.ContractDate, dbo.TblExchangeRequest.DurationID, dbo.TblExchangeRequest.Month,"
MySQL = MySQL & "                                      dbo.TblDurations_Details.Name AS MonthName, dbo.TblExchangeRequest.Date, dbo.TblExchangeRequest.DateH, dbo.TblBranchesData.branch_name,"
MySQL = MySQL & "                                      dbo.TblBranchesData.branch_namee, dbo.TblExchangeRequest.ID AS MainTblID, dbo.TblExchangeReques_Detailst.IDAC,"
MySQL = MySQL & "                                      dbo.TblExchangeReques_Detailst.ID AS DID"
MySQL = MySQL & "                    FROM      dbo.TblDurations_Details INNER JOIN"
MySQL = MySQL & "                                      dbo.TblDurations INNER JOIN"
MySQL = MySQL & "                                      dbo.TblExchangeRequest INNER JOIN"
MySQL = MySQL & "                                      dbo.TblExchangeReques_Detailst ON dbo.TblExchangeRequest.ID = dbo.TblExchangeReques_Detailst.HID ON"
MySQL = MySQL & "                                      dbo.TblDurations.ID = dbo.TblExchangeRequest.DurationID ON dbo.TblDurations_Details.ID = dbo.TblExchangeRequest.Month INNER JOIN"
MySQL = MySQL & "                                      dbo.TblBranchesData ON dbo.TblExchangeRequest.BranchID = dbo.TblBranchesData.branch_id) AS TB1 LEFT OUTER JOIN"
 MySQL = MySQL & "                       (SELECT ID, Db1 AS VName, d1 AS value"
MySQL = MySQL & "                         FROM      dbo.TblExchangeReques_Detailst AS TblExchangeReques_Detailst_20"
MySQL = MySQL & "                         Union"
MySQL = MySQL & "                         SELECT ID, Db2 AS VName, d2 AS value"
MySQL = MySQL & "                         FROM     dbo.TblExchangeReques_Detailst AS TblExchangeReques_Detailst_19"
    MySQL = MySQL & "                     Union"
 MySQL = MySQL & "                        SELECT ID, Db3 AS VName, d3 AS value"
     MySQL = MySQL & "                    FROM     dbo.TblExchangeReques_Detailst AS TblExchangeReques_Detailst_18"
  MySQL = MySQL & "                       Union"
  MySQL = MySQL & "                       SELECT ID, Db4 AS VName, d4 AS value"
  MySQL = MySQL & "                       FROM     dbo.TblExchangeReques_Detailst AS TblExchangeReques_Detailst_17"
   MySQL = MySQL & "                      Union"
 MySQL = MySQL & "                        SELECT ID, Db5 AS VName, d5 AS value"
   MySQL = MySQL & "                      FROM     dbo.TblExchangeReques_Detailst AS TblExchangeReques_Detailst_16"
   MySQL = MySQL & "                      Union"
MySQL = MySQL & "                        SELECT ID, Db6 AS VName, d6 AS value"
 MySQL = MySQL & "                        FROM     dbo.TblExchangeReques_Detailst AS TblExchangeReques_Detailst_15"
 MySQL = MySQL & "                        Union"
  MySQL = MySQL & "                       SELECT ID, Db7 AS VName, d7 AS value"
 MySQL = MySQL & "                        FROM     dbo.TblExchangeReques_Detailst AS TblExchangeReques_Detailst_14"
   MySQL = MySQL & "                      Union"
   MySQL = MySQL & "                      SELECT ID, Db8 AS VName, d8 AS value"
  MySQL = MySQL & "                       FROM     dbo.TblExchangeReques_Detailst AS TblExchangeReques_Detailst_13"
      MySQL = MySQL & "                   Union"
   MySQL = MySQL & "                      SELECT ID, Db9 AS VName, d9 AS value"
   MySQL = MySQL & "                      FROM     dbo.TblExchangeReques_Detailst AS TblExchangeReques_Detailst_12"
MySQL = MySQL & "                         Union"
  MySQL = MySQL & "                       SELECT ID, Db10 AS VName, d10 AS value"
 MySQL = MySQL & "                        FROM     dbo.TblExchangeReques_Detailst AS TblExchangeReques_Detailst_11"
  MySQL = MySQL & "                       Union"
  MySQL = MySQL & "                       SELECT ID, Db11 AS VName, d11 AS value"
  MySQL = MySQL & "                       FROM     dbo.TblExchangeReques_Detailst AS TblExchangeReques_Detailst_10"
   MySQL = MySQL & "                      Union"
  MySQL = MySQL & "                       SELECT ID, Db12 AS VName, d12 AS value"
   MySQL = MySQL & "                      FROM     dbo.TblExchangeReques_Detailst AS TblExchangeReques_Detailst_9"
    MySQL = MySQL & "                     Union"
   MySQL = MySQL & "                      SELECT ID, Db13 AS VName, d13 AS value"
   MySQL = MySQL & "                      FROM     dbo.TblExchangeReques_Detailst AS TblExchangeReques_Detailst_8"
  MySQL = MySQL & "                       Union"
 MySQL = MySQL & "                        SELECT ID, Db14 AS VName, d14 AS value"
 MySQL = MySQL & "                        FROM     dbo.TblExchangeReques_Detailst AS TblExchangeReques_Detailst_7"
MySQL = MySQL & "                         Union"
    MySQL = MySQL & "                     SELECT ID, Db15 AS VName, d15 AS value"
  MySQL = MySQL & "                       FROM     dbo.TblExchangeReques_Detailst AS TblExchangeReques_Detailst_6"
    MySQL = MySQL & "                     Union"
   MySQL = MySQL & "                      SELECT ID, Db16 AS VName, d16 AS value"
  MySQL = MySQL & "                       FROM     dbo.TblExchangeReques_Detailst AS TblExchangeReques_Detailst_5"
MySQL = MySQL & "                         Union"
MySQL = MySQL & "                        SELECT ID, Db17 AS VName, d17 AS value"
  MySQL = MySQL & "                       FROM     dbo.TblExchangeReques_Detailst AS TblExchangeReques_Detailst_4"
   MySQL = MySQL & "                      Union"
  MySQL = MySQL & "                       SELECT ID, Db18 AS VName, d18 AS value"
  MySQL = MySQL & "                       FROM     dbo.TblExchangeReques_Detailst AS TblExchangeReques_Detailst_3"
 MySQL = MySQL & "                        Union"
       MySQL = MySQL & "                 SELECT ID, Db19 AS VName, d19 AS value"
   MySQL = MySQL & "                      FROM     dbo.TblExchangeReques_Detailst AS TblExchangeReques_Detailst_2"
    MySQL = MySQL & "                     Union"
       MySQL = MySQL & "                  SELECT ID, Db20 AS VName, d20 AS value"
     MySQL = MySQL & "                    FROM     dbo.TblExchangeReques_Detailst AS TblExchangeReques_Detailst_1) AS db ON TB1.DID = db.ID"





  MySQL = MySQL & "   where  MainTblID = " & val(txtid.Text)
  
  
  
  
  
     MySQL = MySQL & "  and   TB1.HID =  " & val(txtid.Text)
    
MySQL = "SELECT     TblExchangeRequest.remarks, TblExchangeReques_detailst.remark, TblExchangeReques_Detailst.BankAccount, dbo.TblDurations.Name AS DurName, "
MySQL = MySQL & "                                               dbo.TblExchangeReques_Detailst.HID, dbo.TblExchangeReques_Detailst.CusID, dbo.TblExchangeReques_Detailst.InsID,"
MySQL = MySQL & "                                               dbo.TblExchangeReques_Detailst.InsNo, dbo.TblExchangeReques_Detailst.CusName, dbo.TblExchangeReques_Detailst.FullCode,"
MySQL = MySQL & "                                               dbo.TblExchangeReques_Detailst.Value, dbo.TblExchangeReques_Detailst.Total_Deduct, dbo.TblExchangeReques_Detailst.net,"
MySQL = MySQL & "                                               dbo.TblExchangeReques_Detailst.wokdays, dbo.TblExchangeReques_Detailst.stopdays, dbo.TblExchangeReques_Detailst.stopvalue,"
MySQL = MySQL & "                                               dbo.TblExchangeReques_Detailst.Account_Serial, dbo.TblExchangeReques_Detailst.Account_Code, dbo.TblExchangeReques_Detailst.carid,"
MySQL = MySQL & "                                               dbo.TblExchangeReques_Detailst.boardno, dbo.TblExchangeReques_Detailst.dayvalue, dbo.TblExchangeReques_Detailst.absenceDays,"
MySQL = MySQL & "  dbo.TblExchangeReques_Detailst.absenceValue, dbo.TblExchangeReques_Detailst.ContractBDate, dbo.TblExchangeReques_Detailst.ContractEDate,"
MySQL = MySQL & "  dbo.TblExchangeReques_Detailst.ContractDate, dbo.TblExchangeRequest.DurationID, dbo.TblExchangeRequest.Month,"
MySQL = MySQL & " dbo.TblDurations_Details.Name AS MonthName, dbo.TblExchangeRequest.Date, dbo.TblExchangeRequest.DateH, dbo.TblBranchesData.branch_name,"
MySQL = MySQL & " dbo.TblBranchesData.branch_namee, dbo.TblExchangeRequest.ID AS MainTblID, dbo.TblExchangeReques_Detailst.IDAC,"
MySQL = MySQL & " dbo.TblExchangeReques_Detailst.ID AS DID"
MySQL = MySQL & "  FROM         dbo.TblDurations_Details INNER JOIN"
MySQL = MySQL & " dbo.TblDurations INNER JOIN"
MySQL = MySQL & "                                               dbo.TblExchangeRequest INNER JOIN"
MySQL = MySQL & " dbo.TblExchangeReques_Detailst ON dbo.TblExchangeRequest.ID = dbo.TblExchangeReques_Detailst.HID ON"
MySQL = MySQL & " dbo.TblDurations.ID = dbo.TblExchangeRequest.DurationID ON dbo.TblDurations_Details.ID = dbo.TblExchangeRequest.Month INNER JOIN"
MySQL = MySQL & "                                               dbo.TblBranchesData ON dbo.TblExchangeRequest.BranchID = dbo.TblBranchesData.branch_id"
MySQL = MySQL & "   where  HID = " & val(txtid.Text)
 

    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\Reports\REPORTS NEW\" & "rpt_ExchangeRequestReceipt.rpt"
    Else
        StrFileName = App.path & "\Reports\REPORTS NEW\" & "rpt_ExchangeRequestReceipt.rpt"
    End If
    

    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

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

          If SystemOptions.UserInterface = ArabicInterface Then
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
    
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
   
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



Function print_report2(Optional NoteSerial As Integer)
    
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String




MySQL = " SELECT TblExchangeRequest.Remarks,"
MySQL = MySQL & "        SchoolnMae,"
MySQL = MySQL & "        t.Vatyo,"
MySQL = MySQL & "        t.Vat2,"
MySQL = MySQL & "        TblExchangeReques_Detailst.remark,"
MySQL = MySQL & "        TblExchangeReques_Detailst.BankAccount,"
MySQL = MySQL & "        TblDurations.Name              AS DurName,"
MySQL = MySQL & "        TblExchangeReques_Detailst.HID,t.totalnet,"
MySQL = MySQL & "        TblExchangeReques_Detailst.CusID,"
MySQL = MySQL & "        TblExchangeReques_Detailst.InsID,"
MySQL = MySQL & "        TblExchangeReques_Detailst.InsNo,"
MySQL = MySQL & "        TblExchangeReques_Detailst.CusName,"
MySQL = MySQL & "        TblExchangeReques_Detailst.FullCode,"
MySQL = MySQL & "        TblExchangeReques_Detailst.Value,"
MySQL = MySQL & "        TblExchangeReques_Detailst.Total_Deduct,"
MySQL = MySQL & "        TblExchangeReques_Detailst.net,"
MySQL = MySQL & "        TblExchangeReques_Detailst.wokdays,"
MySQL = MySQL & "        TblExchangeReques_Detailst.stopdays,"
MySQL = MySQL & "        TblExchangeReques_Detailst.stopvalue,"
MySQL = MySQL & "        TblExchangeReques_Detailst.Account_Serial,"
MySQL = MySQL & "        TblExchangeReques_Detailst.carid,"
MySQL = MySQL & "        TblExchangeReques_Detailst.boardno,"
MySQL = MySQL & "        TblExchangeReques_Detailst.dayvalue,"
MySQL = MySQL & "        TblExchangeReques_Detailst.absenceDays,"
MySQL = MySQL & "        TblExchangeReques_Detailst.absenceValue,"
MySQL = MySQL & "        TblExchangeReques_Detailst.ContractBDate,"
MySQL = MySQL & "        TblExchangeReques_Detailst.ContractEDate,"
MySQL = MySQL & "        TblExchangeReques_Detailst.ContractDate,"
MySQL = MySQL & "        TblExchangeRequest.DurationID,"
MySQL = MySQL & "        TblExchangeRequest.Month,"
MySQL = MySQL & "        TblDurations_Details.Name      AS MonthName,"
MySQL = MySQL & "        TblExchangeRequest.Date,"
MySQL = MySQL & "        TblExchangeRequest.DateH,"
MySQL = MySQL & "        TblBranchesData.branch_name,"
MySQL = MySQL & "        TblBranchesData.branch_namee,"
MySQL = MySQL & "        TblExchangeRequest.ID          AS MainTblID,"
MySQL = MySQL & "        TblExchangeReques_Detailst.IDAC,"
MySQL = MySQL & "        TblExchangeReques_Detailst.ID  AS DID"
MySQL = MySQL & " From TblDurations_Details"
MySQL = MySQL & "        INNER JOIN TblDurations"
MySQL = MySQL & "        INNER JOIN TblExchangeRequest"
MySQL = MySQL & "        INNER JOIN TblExchangeReques_Detailst"
MySQL = MySQL & "             ON  TblExchangeRequest.ID = TblExchangeReques_Detailst.HID"
MySQL = MySQL & "             ON  TblDurations.ID = TblExchangeRequest.DurationID"
MySQL = MySQL & "             ON  TblDurations_Details.ID = TblExchangeRequest.Month"
MySQL = MySQL & "        INNER JOIN TblBranchesData"
MySQL = MySQL & "             ON  TblExchangeRequest.BranchID = TblBranchesData.branch_id"
MySQL = MySQL & "        LEFT OUTER JOIN ("
MySQL = MySQL & "                 SELECT DISTINCT TblAttributionInstallmentDivided.ID DividID,"
MySQL = MySQL & "                        TblAttributionInstallmentDivided.REID,"
MySQL = MySQL & "                        TblAttributionInstallmentDivided.Vatyo,"
MySQL = MySQL & "                        TblAttributionInstallmentDivided.Vat2,"
MySQL = MySQL & "                        dbo.TblAttributionInstallmentDivided.PayMentPayed,"
MySQL = MySQL & "                        dbo.TblAttributionInstallmentDivided.DDEmbarkDate,"
MySQL = MySQL & "                        dbo.TblAttributionInstallmentDivided.DDEmbarkDateH,"
MySQL = MySQL & "                        dbo.TblAttributionInstallmentDivided.remark,"
MySQL = MySQL & "                        dbo.TblCustemers.RecordNo,"
MySQL = MySQL & "                        dbo.TblAttributionContract.IDAC,"
MySQL = MySQL & "                        dbo.TblAttributionContract.DurationID,"
MySQL = MySQL & "                        dbo.TblDurations.Name AS DurationName,"
MySQL = MySQL & "                        TblAttributionInstallmentDivided.TypeVAT,TblAttributionInstallmentDivided.TotalNet,dbo.TblCustemers.CusName,"
MySQL = MySQL & "                        dbo.TblCustemers.CusNamee,dbo.TblCustemers.Fullcode,dbo.TblCustemers.CusID,dbo.TblAttributionContract.StartContractDate,"
MySQL = MySQL & "                        dbo.TblAttributionContract.EndContractDate,dbo.TblAttributionContract.FromDate,dbo.TblDurations.FromDate AS DurFromDate,"
MySQL = MySQL & "                        dbo.TblDurations.ToDate AS DurToDate,dbo.TblDurations.FromDateH AS DurFromDateH,dbo.TblVehicleAllocation_Details.Type,"
MySQL = MySQL & "                        dbo.TblVehicleAllocation_Details.StudentCount,dbo.TblVehicleAllocation_Details.BoardNo,dbo.TblVehicleAllocation_Details.rate,"
MySQL = MySQL & "                        dbo.TblVehicleAllocation_Details.Custom,dbo.TblVehicleAllocation_Details.DayRate,dbo.TblVehicleAllocation_Details.StudentCustom,"
MySQL = MySQL & "                        dbo.TblVehicleAllocation_Details.CarID,dbo.TblDurations_Details.Name,dbo.TblVehicleAllocation_Details.ID,"
MySQL = MySQL & "                        dbo.TblVehicleAllocation_Details.SchoolFileID,dbo.TblAttributionContract.StopDeal,dbo.TblAttributionContract.StopDate,"
MySQL = MySQL & "                        dbo.TblAttributionContract.StopDateH,dbo.TblAttributionContract.BranchID,dbo.TblCustemers.Account_Code,"
MySQL = MySQL & "                        dbo.ACCOUNTS.Account_Serial,dbo.TblCustemers.IBAN,dbo.TblCustemers.BankAccount,"
MySQL = MySQL & "                        dbo.TblDurations.Type AS DurationType,dbo.TblAttributionInstallmentDivided.MonthID,"
MySQL = MySQL & "                        dbo.TblAttributionContract.MangerialAreaID,dbo.TblManagerialArea.Name AS MAName,"
MySQL = MySQL & "                        dbo.TblAttributionContract.CityID,dbo.TblAttributionContract.VendorID,dbo.TblSchooleFile.Name AS SchoolnMae,"
MySQL = MySQL & "                        dbo.TblSchooleFile.Namee AS SchoolnMaeE                "
MySQL = MySQL & " From dbo.TblManagerialArea"
MySQL = MySQL & "                        RIGHT OUTER JOIN dbo.ACCOUNTS"
MySQL = MySQL & "                        INNER JOIN dbo.TblAttributionContract"
MySQL = MySQL & "                        INNER JOIN dbo.TblCustemers"
MySQL = MySQL & "                             ON  dbo.TblAttributionContract.VendorID = dbo.TblCustemers.CusID"
MySQL = MySQL & "                        INNER JOIN dbo.TblVehicleAllocation_Details"
MySQL = MySQL & "                             ON  dbo.TblAttributionContract.IDAC = dbo.TblVehicleAllocation_Details.IDVA"
MySQL = MySQL & "                             ON  dbo.ACCOUNTS.Account_Code = dbo.TblCustemers.Account_Code"
MySQL = MySQL & "                        INNER JOIN dbo.TblAttributionInstallmentDivided"
MySQL = MySQL & "                             ON  dbo.TblVehicleAllocation_Details.ID = dbo.TblAttributionInstallmentDivided.DetailsID"
MySQL = MySQL & "                        INNER JOIN dbo.TblDurations_Details"
MySQL = MySQL & "                             ON  dbo.TblAttributionInstallmentDivided.MonthID = dbo.TblDurations_Details.ID"
MySQL = MySQL & "                        LEFT OUTER JOIN dbo.TblSchooleFile"
MySQL = MySQL & "                             ON  dbo.TblVehicleAllocation_Details.SchoolFileID = dbo.TblSchooleFile.ID"
MySQL = MySQL & "                             ON  dbo.TblManagerialArea.ID = dbo.TblAttributionContract.MangerialAreaID"
MySQL = MySQL & "                        LEFT OUTER JOIN dbo.TblVendorCars"
MySQL = MySQL & "                             ON  dbo.TblVehicleAllocation_Details.CarID = dbo.TblVendorCars.ID"
MySQL = MySQL & "                        LEFT OUTER JOIN dbo.TblDurations"
MySQL = MySQL & "                             ON  dbo.TblAttributionContract.DurationID = dbo.TblDurations.ID"
                 
MySQL = MySQL & "             ) T"
MySQL = MySQL & "             ON  T.DividID = TblExchangeReques_Detailst.DividID"
MySQL = MySQL & "             AND T.REID = TblExchangeReques_Detailst.HID"
MySQL = MySQL & " Where TblExchangeReques_Detailst.Hid = " & val(txtid.Text)
       
       

 

    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\Reports\REPORTS NEW\" & "rpt_ExchangeRequestReceipt2.rpt"
    Else
        StrFileName = App.path & "\Reports\REPORTS NEW\" & "rpt_ExchangeRequestReceipt2.rpt"
    End If
    

    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

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

          If SystemOptions.UserInterface = ArabicInterface Then
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
    
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
   
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




Function print_form1(Optional NoteSerial As Integer)
    
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String



 

MySQL = "  SELECT     dbo.TblAttributionInstallmentDivided.REID, dbo.TblExchangeRequest.Remarks, dbo.TblExchangeReques_Detailst.remark, "
  MySQL = MySQL & "                                     dbo.TblExchangeReques_Detailst.BankAccount, dbo.TblDurations.Name AS DurName, dbo.TblExchangeReques_Detailst.HID,"
  MySQL = MySQL & "                                     dbo.TblExchangeReques_Detailst.CusID, dbo.TblExchangeReques_Detailst.InsID, dbo.TblExchangeReques_Detailst.InsNo,"
  MySQL = MySQL & "                                     dbo.TblExchangeReques_Detailst.CusName, dbo.TblExchangeReques_Detailst.FullCode, dbo.TblExchangeReques_Detailst.[Value],"
  MySQL = MySQL & "                                     dbo.TblExchangeReques_Detailst.Total_Deduct, dbo.TblExchangeReques_Detailst.net, dbo.TblExchangeReques_Detailst.wokdays,"
  MySQL = MySQL & "                                     dbo.TblExchangeReques_Detailst.stopdays, dbo.TblExchangeReques_Detailst.stopvalue, dbo.TblExchangeReques_Detailst.Account_Serial,"
  MySQL = MySQL & "                                     dbo.TblExchangeReques_Detailst.Account_Code, dbo.TblExchangeReques_Detailst.carid, dbo.TblExchangeReques_Detailst.boardno,"
  MySQL = MySQL & "                                     dbo.TblExchangeReques_Detailst.dayvalue, dbo.TblExchangeReques_Detailst.absenceDays, dbo.TblExchangeReques_Detailst.absenceValue,"
  MySQL = MySQL & "                                     dbo.TblExchangeReques_Detailst.ContractBDate, dbo.TblExchangeReques_Detailst.ContractEDate, dbo.TblExchangeReques_Detailst.ContractDate,"
  MySQL = MySQL & "                                     dbo.TblExchangeRequest.DurationID, dbo.TblExchangeRequest.[Month], dbo.TblDurations_Details.Name AS MonthName, dbo.TblExchangeRequest.[Date],"
  MySQL = MySQL & "                                     dbo.TblExchangeRequest.DateH, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblExchangeRequest.ID AS MainTblID,"
  MySQL = MySQL & "                                     dbo.TblExchangeReques_Detailst.IDAC, dbo.TblExchangeReques_Detailst.ID AS DID, dbo.TblAttributionInstallmentDivided.RE_Paid,"
  MySQL = MySQL & "                                     dbo.TblExchangeRequest.EntryCreated, dbo.TblAttributionInstallmentDivided.TotalValue, dbo.TblAttributionInstallmentDivided.PayMentPayed,"
  MySQL = MySQL & "                                     dbo.TblAttributionInstallmentDivided.VReceipt_Paid , dbo.TblAttributionInstallmentDivided.VReceiptID"
  MySQL = MySQL & "               FROM         dbo.TblDurations_Details INNER JOIN"
  MySQL = MySQL & "                                     dbo.TblDurations INNER JOIN"
  MySQL = MySQL & "                                     dbo.TblExchangeRequest INNER JOIN"
  MySQL = MySQL & "                                     dbo.TblExchangeReques_Detailst ON dbo.TblExchangeRequest.ID = dbo.TblExchangeReques_Detailst.HID ON"
  MySQL = MySQL & "                                     dbo.TblDurations.ID = dbo.TblExchangeRequest.DurationID ON dbo.TblDurations_Details.ID = dbo.TblExchangeRequest.[Month] INNER JOIN"
  MySQL = MySQL & "                                     dbo.TblBranchesData ON dbo.TblExchangeRequest.BranchID = dbo.TblBranchesData.branch_id INNER JOIN"
  MySQL = MySQL & "                                     dbo.TblAttributionInstallmentDivided ON dbo.TblExchangeReques_Detailst.InsID = dbo.TblAttributionInstallmentDivided.DetailsID AND"
  MySQL = MySQL & "                                     dbo.TblExchangeReques_Detailst.HID = dbo.TblAttributionInstallmentDivided.REID"
  MySQL = MySQL & "                Where (1 = 1) And (dbo.TblAttributionInstallmentDivided.REID = " & val(txtid.Text) & ")"

 
   
    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\Reports\REPORTS NEW\" & "rpt_AttributionContractHeader3Specific.rpt"
    Else
        StrFileName = App.path & "\Reports\REPORTS NEW\" & "rpt_AttributionContractHeader3Specific.rpt"
    End If
 

    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

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

      If SystemOptions.UserInterface = ArabicInterface Then
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
    
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
   
    End If

xReport.ParameterFields(5).AddCurrentValue dcBranch.Text
    
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
Function print_form(Optional NoteSerial As Integer)
    
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String


 

MySQL = MySQL & "   SELECT SUM(dbo.TblExchangeReques_Detailst.net) AS net, dbo.TblExchangeRequest.ID, dbo.TblExchangeReques_Detailst.IDAC, dbo.TblExchangeReques_Detailst.CusID,"
  MySQL = MySQL & "                   dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.RecordNo, dbo.TblCustemers.IBAN, dbo.TblCustemers.BankIBAN,"
 MySQL = MySQL & "                    dbo.TblExchangeReques_Detailst.BankAccount , dbo.TblCustemers.BankCode, dbo.TblCustemers.BankName"
MySQL = MySQL & "   FROM     dbo.TblExchangeRequest INNER JOIN"
 MySQL = MySQL & "                    dbo.TblExchangeReques_Detailst ON dbo.TblExchangeRequest.ID = dbo.TblExchangeReques_Detailst.HID INNER JOIN"
  MySQL = MySQL & "                   dbo.TblCustemers ON dbo.TblExchangeReques_Detailst.CusID = dbo.TblCustemers.CusID"




  MySQL = MySQL & "   where  TblExchangeRequest.id = " & val(txtid.Text)
  
   
    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\Reports\REPORTS NEW\" & "rpt_VendorBankReceipt.rpt"
    Else
        StrFileName = App.path & "\Reports\REPORTS NEW\" & "rpt_VendorBankReceipt.rpt"
    End If
    
    
  '  MySQL = MySQL & "  GROUP BY dbo.TblExchangeRequest.ID, dbo.TblExchangeReques_Detailst.IDAC, dbo.TblExchangeReques_Detailst.CusID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee,"
  ' MySQL = MySQL & "                 dbo.TblCustemers.recordno , dbo.TblCustemers.IBAN, dbo.TblCustemers.BankIBAN, dbo.TblExchangeReques_Detailst.BankAccount"
MySQL = MySQL & "   GROUP BY dbo.TblExchangeRequest.ID, dbo.TblExchangeReques_Detailst.IDAC, dbo.TblExchangeReques_Detailst.CusID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee,"
   MySQL = MySQL & "                  dbo.TblCustemers.RecordNo, dbo.TblCustemers.IBAN, dbo.TblCustemers.BankIBAN, dbo.TblExchangeReques_Detailst.BankAccount, dbo.TblCustemers.BankCode,"
    MySQL = MySQL & "                 dbo.TblCustemers.BankName"



MySQL = "SELECT     SUM(dbo.TblExchangeReques_Detailst.net) AS NET, dbo.TblExchangeRequest.ID, dbo.TblExchangeReques_Detailst.IDAC, dbo.TblExchangeReques_Detailst.CusID, "
    MySQL = MySQL & "                        dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.RecordNo, dbo.TblCustemers.IBAN, dbo.TblCustemers.BankIBAN,"
    MySQL = MySQL & "                        dbo.TblExchangeReques_Detailst.BankAccount , dbo.TblCustemers.BankCode, dbo.TblCustemers.BankName, Max(dbo.TblAttributionInstallmentDivided.NoteID)"
    MySQL = MySQL & "                        AS noteID, MAX(dbo.TblAttributionInstallmentDivided.noteserial1) AS noteserial1"
    MySQL = MySQL & "  FROM         dbo.TblExchangeRequest INNER JOIN"
    MySQL = MySQL & "                        dbo.TblExchangeReques_Detailst ON dbo.TblExchangeRequest.ID = dbo.TblExchangeReques_Detailst.HID INNER JOIN"
    MySQL = MySQL & "                        dbo.TblCustemers ON dbo.TblExchangeReques_Detailst.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
    MySQL = MySQL & "                        dbo.TblAttributionInstallmentDivided ON dbo.TblExchangeReques_Detailst.InsID = dbo.TblAttributionInstallmentDivided.DetailsID"
    MySQL = MySQL & "  Where (dbo.TblAttributionInstallmentDivided.MonthID = " & val(Me.dcMontth.BoundText) & ") And (dbo.TblAttributionInstallmentDivided.DurationID = " & val(Me.DcDur.BoundText) & ")"
    MySQL = MySQL & " AND                       (dbo.TblAttributionInstallmentDivided.REID = " & val(txtid.Text) & ") "
    
    MySQL = MySQL & "  GROUP BY dbo.TblExchangeRequest.ID, dbo.TblExchangeReques_Detailst.IDAC, dbo.TblExchangeReques_Detailst.CusID, dbo.TblCustemers.CusName,"
    MySQL = MySQL & "                        dbo.TblCustemers.CusNamee, dbo.TblCustemers.RecordNo, dbo.TblCustemers.IBAN, dbo.TblCustemers.BankIBAN, dbo.TblExchangeReques_Detailst.BankAccount,"
    MySQL = MySQL & "                        dbo.TblCustemers.BankCode , dbo.TblCustemers.BankName"
    MySQL = MySQL & "  Having (dbo.TblExchangeRequest.ID =" & val(txtid.Text) & ")"
   
'01052015salime
 


    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

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

      If SystemOptions.UserInterface = ArabicInterface Then
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
    
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
   
    End If

xReport.ParameterFields(5).AddCurrentValue dcBranch.Text
    
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

Private Sub XPTxtBoxNamee_GotFocus()
    SwitchKeyboardLang LANG_ENGLISH
End Sub

Function CuurentLogdata(Optional Currentmode As String)
    
       LogTextA = "  Õ›Ÿ ‘«‘… " & "   ÿ·» ’—› „ ⁄ÂœÌ‰   " _
       & CHR(13) & " —ﬁ„ «·”‰œ  " & txtid.Text _
       & CHR(13) & "»‰«¡« ⁄·Ï «” Õﬁ«ﬁ  " & Text1.Text _
       & CHR(13) & "   «·—ﬁ„ «·ÌœÊÏ   " & TXTCode.Text _
       & CHR(13) & " «·„ ⁄Âœ     " & dcCustomer.Text _
       & CHR(13) & " «·›—⁄     " & dcBranch.Text _
       & CHR(13) & " «·ﬂÊœ     " & txtfullcode.Text _
       & CHR(13) & "  «·«œ«—… «· ⁄·Ì„Ì…       " & dcMangerialAreaID.Text _
       & CHR(13) & " —ﬁ„ «·”Ã·   " & txtRecordno.Text _
       & CHR(13) & " «·”‰… «·œ—«”Ì…    " & DcDur.Text _
       & CHR(13) & "  «·› —…  " & dcMontth.Text _
       & CHR(13) & "   «—ÌŒ «·”‰œ „Ì·«œÏ  " & Me.Date.value _
       & CHR(13) & "  «—ÌŒ «·”‰œ ÂÃ—Ï" & DateH.value _
       & CHR(13) & "   „·«ÕŸ«   " & TxtRemarks.Text _
        & CHR(13) & "  «—ÌŒ «·”‰œ ÂÃ—Ï" & TxtNoteSerial.Text

    If Currentmode <> "D" Then
       ' AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTextE, Me.name, Me.TxtModFlg, "", ""
        AddToLogFile CInt(user_id), 8069, Now, Time, LogTextA, LogTexte, Me.Name, "E", "", "", val(TxtNoteSerial), txtid
    Else
     AddToLogFile CInt(user_id), 8069, Now, Time, LogTextA, LogTexte, Me.Name, "D", "", "", val(TxtNoteSerial), txtid
       ' AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTextE, Me.name, "D", "", ""
    End If
End Function


Private Sub Clear_Current_User()
Dim StrSQL As String
        StrSQL = "SELECT  *  From TblExchangeRequest  where  ID = " & val(txtid.Text)
        Set RsT = New ADODB.Recordset
        RsT.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
        If RsT.RecordCount > 0 Then
                   RsT("CurrentUser") = Null
                   RsT.update
        End If
        rs.Resync
End Sub

















