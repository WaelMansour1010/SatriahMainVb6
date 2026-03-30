VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form FrmPermission 
   BackColor       =   &H009E7163&
   Caption         =   "’·«ÕÌ«  «·„” Œœ„Ì‰"
   ClientHeight    =   8760
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13995
   HelpContextID   =   711
   Icon            =   "FrmPermission.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   8760
   ScaleWidth      =   13995
   WindowState     =   2  'Maximized
   Begin C1SizerLibCtl.C1Elastic Ele 
      Height          =   8760
      Index           =   2
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   13995
      _cx             =   24686
      _cy             =   15452
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
      GridRows        =   3
      GridCols        =   4
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"FrmPermission.frx":038A
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin C1SizerLibCtl.C1Tab TabMain 
         Height          =   7335
         Left            =   15
         TabIndex        =   16
         Top             =   690
         Width           =   13965
         _cx             =   24633
         _cy             =   12938
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
         Caption         =   "‘«‘«  «·»—‰«„Ã|’·«ÕÌ«  „Œ’’…|’·«ÕÌ«  „Œ’’… 2"
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
         Begin VB.Frame Frame1 
            Height          =   6960
            Left            =   45
            RightToLeft     =   -1  'True
            TabIndex        =   95
            Top             =   45
            Width           =   13875
            Begin VB.CheckBox Check 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "’·«ÕÌ… «⁄ „«œ ›Ê« Ì— «·„÷Œ« "
               Height          =   300
               Index           =   54
               Left            =   9360
               RightToLeft     =   -1  'True
               TabIndex        =   108
               Top             =   4140
               Width           =   4260
            End
            Begin VB.CheckBox Check 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "’·«ÕÌ… ›Ê« Ì—"
               Height          =   300
               Index           =   53
               Left            =   9360
               RightToLeft     =   -1  'True
               TabIndex        =   107
               Top             =   3720
               Width           =   4260
            End
            Begin VB.CheckBox Check 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "’·«ÕÌ… —›⁄ ›Ê« Ì— «·«„·«ﬂ ÊÌ»"
               Height          =   300
               Index           =   52
               Left            =   9360
               RightToLeft     =   -1  'True
               TabIndex        =   106
               Top             =   3270
               Width           =   4260
            End
            Begin VB.CheckBox Check 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "’·«ÕÌ…  ﬁ—Ì— «·„‘«—Ì⁄ ›Ï «·Õ”«»«  ›ﬁÿ"
               Height          =   300
               Index           =   51
               Left            =   9375
               RightToLeft     =   -1  'True
               TabIndex        =   105
               Top             =   2700
               Width           =   4260
            End
            Begin VB.CheckBox Check 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "’·«ÕÌÂ “Ì«œ… «·”⁄— ›ﬁÿ"
               Height          =   300
               Index           =   50
               Left            =   9375
               RightToLeft     =   -1  'True
               TabIndex        =   104
               Top             =   2190
               Width           =   4260
            End
            Begin VB.CheckBox Check 
               Alignment       =   1  'Right Justify
               BackColor       =   &H000000FF&
               Caption         =   "Ì»œ√ » ›« Ê—… «·„Õÿ« "
               ForeColor       =   &H0080FFFF&
               Height          =   225
               Index           =   49
               Left            =   180
               RightToLeft     =   -1  'True
               TabIndex        =   103
               Top             =   1800
               Width           =   2640
            End
            Begin VB.CheckBox Check 
               Alignment       =   1  'Right Justify
               BackColor       =   &H000000FF&
               Caption         =   "Ì»œ√ » ›« Ê—… «·„÷Œ« "
               ForeColor       =   &H0080FFFF&
               Height          =   225
               Index           =   48
               Left            =   180
               RightToLeft     =   -1  'True
               TabIndex        =   102
               Top             =   1320
               Width           =   2640
            End
            Begin VB.CheckBox Check 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "’·«ÕÌÂ › Õ «„— «·«’·«Õ"
               Height          =   300
               Index           =   35
               Left            =   9375
               RightToLeft     =   -1  'True
               TabIndex        =   101
               Top             =   1710
               Width           =   4260
            End
            Begin VB.CheckBox Check 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "Ì› Õ ⁄·Ì  ‰»ÌÂ«  «·«‰ «Ã"
               Height          =   360
               Index           =   17
               Left            =   7920
               RightToLeft     =   -1  'True
               TabIndex        =   100
               Top             =   300
               Width           =   5715
            End
            Begin VB.CheckBox Check 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«Œ›«¡ «·„⁄·Ê„«  «·„«·Ì… ›Ì ‰ﬁÿ… «·»Ì⁄"
               Height          =   360
               Index           =   18
               Left            =   7920
               RightToLeft     =   -1  'True
               TabIndex        =   99
               Top             =   840
               Width           =   5715
            End
            Begin VB.CheckBox Check 
               Alignment       =   1  'Right Justify
               BackColor       =   &H000000FF&
               Caption         =   "«ŸÂ«— «·ﬁÊ«∆„ «·„«·ÌÂ ﬂ«„·…"
               ForeColor       =   &H0080FFFF&
               Height          =   225
               Index           =   34
               Left            =   240
               RightToLeft     =   -1  'True
               TabIndex        =   98
               Top             =   450
               Width           =   2640
            End
            Begin VB.CheckBox Check 
               Alignment       =   1  'Right Justify
               BackColor       =   &H000000FF&
               Caption         =   "Ì»œ√ » ›« Ê—… „Œ‰’—Â"
               ForeColor       =   &H0080FFFF&
               Height          =   225
               Index           =   37
               Left            =   210
               RightToLeft     =   -1  'True
               TabIndex        =   97
               Top             =   870
               Width           =   2640
            End
            Begin VB.CheckBox Check 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«Œ›«¡   «·Õ–› ›Ì ‰ﬁÿ… «·»Ì⁄"
               Height          =   270
               Index           =   45
               Left            =   9930
               RightToLeft     =   -1  'True
               TabIndex        =   96
               Top             =   1290
               Width           =   3705
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   6960
            Index           =   5
            Left            =   -14520
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   45
            Width           =   13875
            _cx             =   24474
            _cy             =   12277
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
            Begin VB.CheckBox Check 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "’·«ÕÌ… «· ÕÊÌ· «·Ï «·‘∆Ê‰ «·ﬁ«‰Ê‰Ì…"
               Height          =   195
               Index           =   47
               Left            =   -60
               RightToLeft     =   -1  'True
               TabIndex        =   94
               Top             =   300
               Width           =   4350
            End
            Begin VB.CheckBox Check 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "’·«ÕÌ… ‘«„· ›Ï «·«„·«ﬂ"
               Height          =   300
               Index           =   46
               Left            =   6720
               RightToLeft     =   -1  'True
               TabIndex        =   93
               Top             =   5760
               Width           =   2835
            End
            Begin VB.CheckBox Check 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "’·«ÕÌ… «· ⁄œÌ· ›Ï «ﬁ· ﬁÌ„… «ÌÃ«—Ì…"
               Height          =   300
               Index           =   44
               Left            =   4680
               RightToLeft     =   -1  'True
               TabIndex        =   92
               Top             =   5400
               Width           =   3435
            End
            Begin VB.CheckBox Check 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄œ„ «·”„«Õ »⁄„· Œ’Ê„«  ⁄·Ï «·”ÿ—"
               Height          =   300
               Index           =   43
               Left            =   4410
               RightToLeft     =   -1  'True
               TabIndex        =   91
               Top             =   5130
               Width           =   3675
            End
            Begin VB.CheckBox Check 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄œ„ «· ⁄œÌ· ›Ï «·«”⁄«— ›Ï «·ÿ·»«  «·œ«Œ·Ì…"
               Height          =   300
               Index           =   42
               Left            =   4410
               RightToLeft     =   -1  'True
               TabIndex        =   90
               Top             =   4860
               Width           =   3675
            End
            Begin VB.CheckBox Check 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄œ„ «· ⁄œÌ· ›Ï «·«”⁄«— ›Ï «·„—œÊœ« "
               Height          =   300
               Index           =   41
               Left            =   4530
               RightToLeft     =   -1  'True
               TabIndex        =   89
               Top             =   4500
               Width           =   3555
            End
            Begin VB.CheckBox Check 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "Õ—ﬂ«  «·»Ì⁄ Ê«·„—œÊœ ··„” Œœ„ ·«  ‰‘√ ”‰œ«  „Œ“‰Ì…"
               ForeColor       =   &H000000FF&
               Height          =   420
               Index           =   40
               Left            =   4440
               RightToLeft     =   -1  'True
               TabIndex        =   88
               Top             =   4110
               Width           =   3660
            End
            Begin VB.CheckBox Check 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " »‰«¡ ⁄·Ì «·“«„Ì ›Ì ÿ·» «„—  «·‘—«¡ Ê ›« Ê—… «·‘—«¡"
               Height          =   420
               Index           =   39
               Left            =   4560
               RightToLeft     =   -1  'True
               TabIndex        =   87
               Top             =   3720
               Width           =   3555
            End
            Begin VB.CheckBox Check 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ ÿ·» «·’—› «·“«„Ì ›Ì ’—› «·„œ›Ê⁄«  Ê«·„’—Ê›« "
               Height          =   420
               Index           =   38
               Left            =   4560
               RightToLeft     =   -1  'True
               TabIndex        =   86
               Top             =   3420
               Width           =   3555
            End
            Begin VB.CheckBox Check 
               Alignment       =   1  'Right Justify
               BackColor       =   &H000000FF&
               Caption         =   "’·«ÕÌ… ‰ÊÀÌﬁ ⁄ﬁÊœ «·⁄ﬁ«—"
               ForeColor       =   &H0080FFFF&
               Height          =   225
               Index           =   36
               Left            =   1560
               RightToLeft     =   -1  'True
               TabIndex        =   85
               Top             =   3000
               Width           =   2640
            End
            Begin VB.CheckBox Check 
               Alignment       =   1  'Right Justify
               BackColor       =   &H000000FF&
               Caption         =   "«„ﬂ«‰Ì…  €ÌÌ— «·÷—Ì»Â ÌœÊÌ"
               ForeColor       =   &H0080FFFF&
               Height          =   300
               Index           =   33
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   84
               Top             =   2160
               Width           =   4260
            End
            Begin VB.CheckBox Check 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "’·«ÕÌÂ  ⁄œÌ· ”‰œ «” ·«„ «‰ «Ã  «„ „—»Êÿ »«„— «‰ «Ã"
               ForeColor       =   &H000000FF&
               Height          =   300
               Index           =   32
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   83
               Top             =   1800
               Width           =   4260
            End
            Begin VB.CheckBox Check 
               Alignment       =   1  'Right Justify
               BackColor       =   &H000000FF&
               Caption         =   "’·«ÕÌ… «⁄ÿ«¡ «·’·«ÕÌ«  ··„” Œœ„Ì‰"
               DataField       =   " "
               ForeColor       =   &H0080FFFF&
               Height          =   300
               Index           =   31
               Left            =   -1320
               RightToLeft     =   -1  'True
               TabIndex        =   82
               Top             =   1440
               Width           =   5580
            End
            Begin VB.CheckBox Check 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "’·«ÕÌ… «·”œ«œ ›Ì ‰ﬁ«ÿ «·»Ì⁄ »œÊ‰ «·ÿ»«⁄Â"
               DataField       =   " "
               Height          =   300
               Index           =   30
               Left            =   -1320
               RightToLeft     =   -1  'True
               TabIndex        =   77
               Top             =   1200
               Width           =   5580
            End
            Begin VB.CheckBox Check 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "’·«ÕÌ… ÿ»«⁄… ›« Ê—… «·„»Ì⁄«  ⁄œ… „—« "
               Height          =   300
               Index           =   29
               Left            =   -1290
               RightToLeft     =   -1  'True
               TabIndex        =   76
               Top             =   840
               Width           =   5580
            End
            Begin VB.CheckBox Check 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "’·«ÕÌ… «· ÕÊÌ· „‰ «„— «·«‰ «Ã"
               Height          =   300
               Index           =   28
               Left            =   -1290
               RightToLeft     =   -1  'True
               TabIndex        =   75
               Top             =   480
               Width           =   5580
            End
            Begin VB.CheckBox Check 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "  «· ⁄œÌ· ›ﬁÿ  ›Ì ÿ—ﬁ «·œ›⁄ ›ﬁÿ  ··„»Ì⁄«  Ê «·„ﬁ»Ê÷« "
               Height          =   195
               Index           =   27
               Left            =   -60
               RightToLeft     =   -1  'True
               TabIndex        =   74
               Top             =   60
               Width           =   4350
            End
            Begin VB.CheckBox Check 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«„ﬂ«‰Ì…  ⁄œÌ· «·„—ﬂ»… Õ«· «— »«ÿÂ« »√„— ‘€· ’Ì«‰…"
               ForeColor       =   &H000000FF&
               Height          =   300
               Index           =   26
               Left            =   -1440
               RightToLeft     =   -1  'True
               TabIndex        =   73
               Top             =   5370
               Width           =   5700
            End
            Begin VB.CheckBox Check 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·€«¡ ’·«ÕÌ… ⁄„Ì· Ê„Ê—œ „‰ ‘«‘… «·⁄„·«¡«"
               ForeColor       =   &H000000FF&
               Height          =   300
               Index           =   25
               Left            =   -1440
               RightToLeft     =   -1  'True
               TabIndex        =   72
               Top             =   5100
               Width           =   5700
            End
            Begin VB.CheckBox Check 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "’·«ÕÌ… «·€«¡ «·« ›«ﬁÌÂ ··„ﬁ«Ì”« "
               ForeColor       =   &H000000FF&
               Height          =   300
               Index           =   24
               Left            =   -1440
               RightToLeft     =   -1  'True
               TabIndex        =   71
               Top             =   4830
               Width           =   5700
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic1 
               Height          =   1200
               Left            =   7485
               TabIndex        =   69
               TabStop         =   0   'False
               Top             =   6300
               Visible         =   0   'False
               Width           =   6510
               _cx             =   11483
               _cy             =   2117
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
               Begin VSFlex8UCtl.VSFlexGrid Grid 
                  Height          =   1080
                  Left            =   600
                  TabIndex        =   70
                  Top             =   120
                  Width           =   7995
                  _cx             =   14102
                  _cy             =   1905
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
                  FormatString    =   $"FrmPermission.frx":03FE
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
            End
            Begin VB.CheckBox Check 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "’·«ÕÌ… «· ⁄œÌ· ⁄·Ì ”‰œ «·’—› »ı‰«¡« ⁄·Ì «·ÿ·» «·œ«Œ·Ì"
               ForeColor       =   &H000000FF&
               Height          =   300
               Index           =   23
               Left            =   -1440
               RightToLeft     =   -1  'True
               TabIndex        =   66
               Top             =   4590
               Width           =   5700
            End
            Begin VB.CheckBox Check 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "’·«ÕÌ…  ⁄œÌ· «·—Õ·«  »⁄œ ⁄„· ›Ê« Ì— ·Â«"
               ForeColor       =   &H000000FF&
               Height          =   300
               Index           =   22
               Left            =   -1440
               RightToLeft     =   -1  'True
               TabIndex        =   65
               Top             =   4320
               Width           =   5700
            End
            Begin VB.CheckBox Check 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "’·«ÕÌ…  ⁄œÌ· Õ«·… Ê «—ÌŒ «·ÿ·» ›Ì ÿ·»«  «·’Ì«‰…"
               ForeColor       =   &H000000FF&
               Height          =   300
               Index           =   21
               Left            =   -1440
               RightToLeft     =   -1  'True
               TabIndex        =   64
               Top             =   4080
               Width           =   5700
            End
            Begin VB.CheckBox Check 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "’·«ÕÌ…  ⁄œÌ·   ›Ê« Ì— «·„»Ì⁄«  «·„‰‘√Â «·Ì« „‰ ‘«‘Â «· Ã„Ì⁄"
               ForeColor       =   &H000000FF&
               Height          =   300
               Index           =   20
               Left            =   -1440
               RightToLeft     =   -1  'True
               TabIndex        =   63
               Top             =   3720
               Width           =   5700
            End
            Begin VB.CheckBox Check 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "’·«ÕÌ…  ⁄œÌ· «·„” ‰œ«  ﬁÌœ «·«⁄ „«œ"
               ForeColor       =   &H000000FF&
               Height          =   300
               Index           =   19
               Left            =   -1440
               RightToLeft     =   -1  'True
               TabIndex        =   62
               Top             =   3480
               Width           =   5700
            End
            Begin VB.CheckBox Check 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "’·«ÕÌÂ  ŒÿÏ Œ’„ «·„Ã„Ê⁄« "
               Height          =   300
               Index           =   16
               Left            =   2400
               RightToLeft     =   -1  'True
               TabIndex        =   61
               Top             =   2520
               Width           =   5715
            End
            Begin VB.CheckBox Check 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "’·«ÕÌÂ  ÕÊÌ·  ‰»ÌÂ «·ŒÿÂ  ·«„— ‘€·"
               Height          =   300
               Index           =   15
               Left            =   2400
               RightToLeft     =   -1  'True
               TabIndex        =   60
               Top             =   2190
               Width           =   5715
            End
            Begin VB.CheckBox Check 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "’·«ÕÌÂ  ⁄œÌ·   «·—’Ìœ «·«›  «ÕÌ ›Ì ‘«‘Â «·⁄„·«¡"
               Height          =   300
               Index           =   14
               Left            =   2400
               RightToLeft     =   -1  'True
               TabIndex        =   59
               Top             =   1860
               Width           =   5715
            End
            Begin VB.CheckBox Check 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "’·«ÕÌÂ  ⁄œÌ· «·Õœ «·«∆ „«‰Ì ›Ì ‘«‘Â «·⁄„·«¡"
               Height          =   300
               Index           =   13
               Left            =   2400
               RightToLeft     =   -1  'True
               TabIndex        =   58
               Top             =   1590
               Width           =   5715
            End
            Begin VB.CheckBox Check 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄œ„ «· ⁄«„· „⁄ „›—œ«  «·—« » „‰ ‘«‘… «·„ÊŸ›Ì‰"
               Height          =   300
               Index           =   12
               Left            =   2400
               RightToLeft     =   -1  'True
               TabIndex        =   57
               Top             =   1320
               Width           =   5715
            End
            Begin VB.CheckBox Check 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "’·«ÕÌ… «‰‘«¡ ”‰œ «·«” ·«„ Ê«·’—› ›Ì «·ÃÊœ…"
               Height          =   300
               Index           =   11
               Left            =   2400
               RightToLeft     =   -1  'True
               TabIndex        =   56
               Top             =   1080
               Width           =   5715
            End
            Begin VB.CheckBox Check 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "’·«ÕÌ…  ŒÿÌ «⁄ „«œ  ŒÿÌ Õœ «·«∆ „«‰ ›Ì ›Ê« Ì— «·„»Ì⁄« "
               Height          =   300
               Index           =   10
               Left            =   8055
               RightToLeft     =   -1  'True
               TabIndex        =   55
               Top             =   6000
               Width           =   5700
            End
            Begin VB.CheckBox Check 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "’·«ÕÌ… «·œ›⁄ ›Ì ‰ﬁ«ÿ «·»Ì⁄"
               Height          =   300
               Index           =   9
               Left            =   2400
               RightToLeft     =   -1  'True
               TabIndex        =   54
               Top             =   840
               Width           =   5715
            End
            Begin VB.CheckBox Check 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄œ„ «„ﬂ«‰Ì…  ⁄œÌ· «·Ã“¡ «·À«» "
               Height          =   300
               Index           =   8
               Left            =   2400
               RightToLeft     =   -1  'True
               TabIndex        =   53
               Top             =   600
               Width           =   5715
            End
            Begin VB.CheckBox Check 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄œ„ «„ﬂ«‰Ì…  ⁄œÌ· «·›—⁄"
               Height          =   300
               Index           =   7
               Left            =   2400
               RightToLeft     =   -1  'True
               TabIndex        =   52
               Top             =   360
               Width           =   5715
            End
            Begin VB.CheckBox Check 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄œ„ «„ﬂ«‰Ì…  ⁄œÌ· «· «—ÌŒ"
               Height          =   300
               Index           =   6
               Left            =   2400
               RightToLeft     =   -1  'True
               TabIndex        =   51
               Top             =   90
               Width           =   5715
            End
            Begin VB.CheckBox Check 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "’·«ÕÌ… ≈ŸÂ«— ﬂ· «·„ÊŸ›Ì‰"
               Height          =   300
               Index           =   5
               Left            =   8055
               RightToLeft     =   -1  'True
               TabIndex        =   50
               Top             =   5760
               Width           =   5700
            End
            Begin VB.CheckBox Check 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "’·«ÕÌ…  ŒÿÌ «·Õœ «·«∆ „«‰Ì"
               Height          =   300
               Index           =   4
               Left            =   8055
               RightToLeft     =   -1  'True
               TabIndex        =   49
               Top             =   5520
               Width           =   5700
            End
            Begin VB.CheckBox Check 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "’·«ÕÌ…  ⁄œÌ· Õ«·… «·ÊÕœ… ›Ì „·› «·⁄ﬁ«—« "
               Height          =   300
               Index           =   3
               Left            =   8055
               RightToLeft     =   -1  'True
               TabIndex        =   48
               Top             =   5280
               Width           =   5700
            End
            Begin VB.CheckBox Check 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "’·«ÕÌ…  ⁄œÌ· ›« Ê—… «·„»Ì⁄«  ›Ì Õ«·Â «· ÕÊÌ· «·„Œ“‰Ì"
               Height          =   300
               Index           =   2
               Left            =   8055
               RightToLeft     =   -1  'True
               TabIndex        =   47
               Top             =   5040
               Width           =   5700
            End
            Begin VB.CheckBox Check 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "’·«ÕÌ… ⁄„· ›« Ê—… „»Ì⁄«   ›Ì Õ«·Â ⁄œ„ ÊÃÊœ  ﬂ·›… ··’‰›"
               Height          =   300
               Index           =   1
               Left            =   8055
               RightToLeft     =   -1  'True
               TabIndex        =   46
               Top             =   4800
               Width           =   5700
            End
            Begin VB.CheckBox Check 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "’·«ÕÌ…  ⁄œÌ·  ”⁄— «·»‰œ ›Ì ›Ê« Ì— «·„‘«—Ì⁄"
               Height          =   300
               Index           =   0
               Left            =   9015
               RightToLeft     =   -1  'True
               TabIndex        =   45
               Top             =   4560
               Width           =   4740
            End
            Begin VB.CheckBox ChAllowCompChanPrice 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "’·«ÕÌ…  ⁄œÌ· «”⁄«— « ›«ﬁÌ… «·‘—ﬂ« "
               Height          =   300
               Left            =   9015
               RightToLeft     =   -1  'True
               TabIndex        =   44
               Top             =   4320
               Width           =   4740
            End
            Begin VB.CheckBox AllowCreateHajomraVoucher 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "’·«ÕÌ… «‰‘«¡ ﬁÌÊœ «·⁄„—…"
               Height          =   300
               Left            =   9015
               RightToLeft     =   -1  'True
               TabIndex        =   43
               Top             =   4080
               Width           =   4740
            End
            Begin VB.CheckBox ChkAllowOrbonDate 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·”„«Õ » ŒÿÌ „œ… «·⁄—»Ê‰"
               Height          =   300
               Left            =   9015
               RightToLeft     =   -1  'True
               TabIndex        =   42
               Top             =   3840
               Width           =   4740
            End
            Begin VB.CheckBox chkDev 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "’·«ÕÌ…   ﬁÌÌ„ «·«œ«¡ Ê«·„Â«„"
               Height          =   300
               Left            =   8970
               RightToLeft     =   -1  'True
               TabIndex        =   41
               Top             =   3600
               Width           =   4740
            End
            Begin VB.CheckBox AllowRequestgl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "’·«ÕÌ…  ‰›Ì– ﬁÌœ  ÿ·» «·’—› ··„ ⁄ÂœÌ‰"
               Height          =   300
               Left            =   9015
               RightToLeft     =   -1  'True
               TabIndex        =   40
               Top             =   3360
               Width           =   4740
            End
            Begin VB.CheckBox AllowBigAccount 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "’·«ÕÌ… ⁄—÷ «·ﬁÊ«∆„ «·„«·Ì…"
               Height          =   300
               Left            =   9015
               RightToLeft     =   -1  'True
               TabIndex        =   39
               Top             =   3120
               Width           =   4740
            End
            Begin VB.CheckBox Allowpayroll 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "’·«ÕÌ…  ‰›Ì– ﬁÌœ «·«” Õﬁ«ﬁ ··—Ê« »"
               Height          =   300
               Left            =   9015
               RightToLeft     =   -1  'True
               TabIndex        =   38
               Top             =   2880
               Width           =   4740
            End
            Begin VB.CheckBox AllowSett1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "’·«ÕÌ…  ‰›Ì– «· ”ÊÌ«  «·Ã—œÌ…"
               Height          =   300
               Left            =   9015
               RightToLeft     =   -1  'True
               TabIndex        =   37
               Top             =   2640
               Width           =   4740
            End
            Begin VB.CheckBox AllowSett 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "’·«ÕÌ…  ‰›Ì– «·Ã—œ"
               Height          =   300
               Left            =   9015
               RightToLeft     =   -1  'True
               TabIndex        =   36
               Top             =   2400
               Width           =   4740
            End
            Begin VB.CheckBox chkExceedShipment 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "’·«ÕÌ… «·‘Õ‰ «ﬂ — „‰ «·ﬂ„Ì… «·„ÿ·Ê»…"
               Height          =   300
               Left            =   9015
               RightToLeft     =   -1  'True
               TabIndex        =   35
               Top             =   2160
               Width           =   4740
            End
            Begin VB.CheckBox chkhideColumn 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄œ„ «„ﬂ«‰Ì…  ⁄œÌ· «·«⁄„œ… ›Ì «·”‰œ«  «·„Œ“‰Ì…"
               Height          =   300
               Left            =   9015
               RightToLeft     =   -1  'True
               TabIndex        =   31
               Top             =   1920
               Width           =   4740
            End
            Begin VB.CheckBox chkHideCost 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«Œ›«¡ «· ﬂ·›…"
               Height          =   300
               Left            =   9015
               RightToLeft     =   -1  'True
               TabIndex        =   30
               Top             =   1680
               Width           =   4740
            End
            Begin VB.CheckBox ChkShowCommisions 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«ŸÂ«— «·⁄„Ê·« "
               Height          =   300
               Left            =   9015
               RightToLeft     =   -1  'True
               TabIndex        =   29
               Top             =   1440
               Width           =   4740
            End
            Begin VB.CheckBox ChkFixedCustomer 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—»ÿ «·„” Œœ„ »⁄„·«∆Â ÊÕ—ﬂ« Â  ›ﬁÿ"
               Height          =   300
               Left            =   9015
               RightToLeft     =   -1  'True
               TabIndex        =   28
               Top             =   1200
               Width           =   4740
            End
            Begin VB.CheckBox ChkInvAbility2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "·Â «·ﬁœ—…  ⁄·Ï   ⁄œÌ· «·«”⁄«— ›Ì ”‰œ«  «·«” ·«„ «·„Œ“‰Ì"
               Height          =   345
               Left            =   9375
               RightToLeft     =   -1  'True
               TabIndex        =   25
               Top             =   600
               Width           =   4380
            End
            Begin VB.CheckBox ChkInvAbility1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "·Â «·ﬁœ—…  ⁄·Ï   ⁄œÌ· «·«”⁄«— ›Ì ”‰œ«  «·’—› «·„Œ“‰Ì"
               Height          =   345
               Left            =   9375
               RightToLeft     =   -1  'True
               TabIndex        =   24
               Top             =   360
               Width           =   4380
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   690
               Index           =   6
               Left            =   8205
               TabIndex        =   22
               TabStop         =   0   'False
               Top             =   6270
               Visible         =   0   'False
               Width           =   5610
               _cx             =   9895
               _cy             =   1217
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
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
               BackColor       =   14871017
               ForeColor       =   192
               FloodColor      =   6553600
               ForeColorDisabled=   -2147483631
               Caption         =   "’·«ÕÌ«  Œ«’… »≈–‰ ’—› «·»÷«⁄…"
               Align           =   0
               AutoSizeChildren=   7
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
               Begin VB.CheckBox Chk 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·”„«Õ ··„” Œœ„ »’—› ﬂ„Ì«  √⁄·Ï „‰ ﬂ„Ì… «·’‰› «·„Ã„⁄"
                  Height          =   315
                  Left            =   135
                  RightToLeft     =   -1  'True
                  TabIndex        =   23
                  Top             =   450
                  Visible         =   0   'False
                  Width           =   1170
               End
            End
            Begin VB.CheckBox ChkInvAbility 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "·Â «·ﬁœ—… ⁄·Ï  ⁄œÌ· ›Ì ›Ê« Ì— «·»Ì⁄"
               Height          =   300
               Left            =   9015
               RightToLeft     =   -1  'True
               TabIndex        =   21
               Top             =   120
               Width           =   4740
            End
            Begin VB.CheckBox ChkInvProfit 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "·Â Õﬁ „‘«Âœ… ’«›Ï «·—»Õ ›Ï «·›« Ê—…"
               Height          =   300
               Left            =   9015
               RightToLeft     =   -1  'True
               TabIndex        =   20
               Top             =   885
               Width           =   4740
            End
            Begin VSFlex8UCtl.VSFlexGrid fg3 
               Height          =   750
               Left            =   120
               TabIndex        =   78
               Top             =   6180
               Width           =   6300
               _cx             =   11112
               _cy             =   1323
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
               SelectionMode   =   0
               GridLines       =   1
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   12
               Cols            =   4
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   320
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmPermission.frx":061B
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
            Begin MSDataListLib.DataCombo cmbAccounts 
               Height          =   315
               Left            =   1410
               TabIndex        =   79
               Top             =   5790
               Width           =   3630
               _ExtentX        =   6403
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin ImpulseButton.ISButton ISButton3 
               Height          =   375
               Left            =   60
               TabIndex        =   80
               ToolTipText     =   "«÷«›… «·»Ì«‰«  «·Ï «·œ« «"
               Top             =   5730
               Width           =   1230
               _ExtentX        =   2170
               _ExtentY        =   661
               Caption         =   "«÷«›…  ”ÿ—"
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
               ButtonImage     =   "FrmPermission.frx":06D9
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageExtraction=   0
               LowerToggledContent=   0   'False
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               Caption         =   "«Œ — «·Õ”«»"
               Height          =   255
               Left            =   4920
               RightToLeft     =   -1  'True
               TabIndex        =   81
               Top             =   5850
               Width           =   1200
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   6960
            Index           =   4
            Left            =   -14820
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   45
            Width           =   13875
            _cx             =   24474
            _cy             =   12277
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
            Begin VSFlex8Ctl.VSFlexGrid Fg 
               Height          =   6885
               Left            =   30
               TabIndex        =   19
               Top             =   30
               Width           =   13815
               _cx             =   24368
               _cy             =   12144
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
               BackColorBkg    =   -2147483643
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
               Rows            =   3
               Cols            =   12
               FixedRows       =   1
               FixedCols       =   0
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"FrmPermission.frx":6F3B
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
               OutlineBar      =   1
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
               WordWrap        =   -1  'True
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
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   660
         Index           =   3
         Left            =   15
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   15
         Width           =   13965
         _cx             =   24633
         _cy             =   1164
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
         Begin VB.TextBox TxtCode 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   360
            IMEMode         =   3  'DISABLE
            Left            =   11880
            MaxLength       =   20
            RightToLeft     =   -1  'True
            TabIndex        =   26
            Top             =   120
            Width           =   945
         End
         Begin VB.TextBox TxtModFlg 
            Alignment       =   1  'Right Justify
            Height          =   330
            Left            =   1665
            RightToLeft     =   -1  'True
            TabIndex        =   7
            Top             =   -210
            Visible         =   0   'False
            Width           =   810
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   555
            Index           =   0
            Left            =   120
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   0
            Width           =   8880
            _cx             =   15663
            _cy             =   979
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
            Begin VB.TextBox TxtCode1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   360
               IMEMode         =   3  'DISABLE
               Left            =   2370
               MaxLength       =   20
               RightToLeft     =   -1  'True
               TabIndex        =   33
               Top             =   120
               Width           =   900
            End
            Begin VB.OptionButton Opt 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "’·«ÕÌ… „À·"
               Height          =   315
               Index           =   3
               Left            =   3240
               RightToLeft     =   -1  'True
               TabIndex        =   32
               Top             =   180
               Value           =   -1  'True
               Width           =   1050
            End
            Begin VB.OptionButton Opt 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "’·«ÕÌ…  «„… "
               Height          =   315
               Index           =   0
               Left            =   7380
               RightToLeft     =   -1  'True
               TabIndex        =   6
               Top             =   180
               Width           =   1230
            End
            Begin VB.OptionButton Opt 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "≈·€«¡ Ã„Ì⁄ «·’·«ÕÌ« "
               Height          =   315
               Index           =   1
               Left            =   5610
               RightToLeft     =   -1  'True
               TabIndex        =   5
               Top             =   180
               Width           =   1560
            End
            Begin VB.OptionButton Opt 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " Œ’Ì’ ...."
               Height          =   315
               Index           =   2
               Left            =   4185
               RightToLeft     =   -1  'True
               TabIndex        =   4
               Top             =   180
               Width           =   1035
            End
            Begin MSDataListLib.DataCombo DcboUsers1 
               Height          =   315
               Left            =   0
               TabIndex        =   34
               Top             =   120
               Width           =   2325
               _ExtentX        =   4101
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               BackColor       =   -2147483624
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Image Img 
               Height          =   240
               Index           =   0
               Left            =   8610
               Picture         =   "FrmPermission.frx":7103
               Top             =   180
               Width           =   240
            End
            Begin VB.Image Img 
               Height          =   240
               Index           =   1
               Left            =   7140
               Picture         =   "FrmPermission.frx":748D
               Top             =   180
               Width           =   240
            End
            Begin VB.Image Img 
               Height          =   240
               Index           =   2
               Left            =   5385
               Picture         =   "FrmPermission.frx":7A17
               Top             =   180
               Width           =   240
            End
         End
         Begin MSDataListLib.DataCombo DcboUsers 
            Height          =   315
            Left            =   8970
            TabIndex        =   8
            Top             =   120
            Width           =   2115
            _ExtentX        =   3731
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   -2147483624
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSComctlLib.ImageList ImgLstScreens 
            Left            =   60
            Top             =   30
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   3
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmPermission.frx":7DA1
                  Key             =   "GroupImg"
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmPermission.frx":813B
                  Key             =   "GroupImgOpen"
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmPermission.frx":84D5
                  Key             =   "ScreenImg"
               EndProperty
            EndProperty
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬂÊœ «·„” Œœ„"
            Height          =   330
            Left            =   12960
            RightToLeft     =   -1  'True
            TabIndex        =   27
            Top             =   120
            Width           =   975
         End
         Begin VB.Label LblScreenName 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   270
            Left            =   855
            RightToLeft     =   -1  'True
            TabIndex        =   10
            Top             =   150
            Width           =   675
         End
         Begin VB.Label LblUsers 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·„” Œœ„"
            Height          =   330
            Left            =   10845
            RightToLeft     =   -1  'True
            TabIndex        =   9
            Top             =   150
            Width           =   975
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   705
         Index           =   1
         Left            =   15
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   8040
         Width           =   13965
         _cx             =   24633
         _cy             =   1244
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
         BackColor       =   14737632
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
         Begin VB.CommandButton CmdImport 
            Caption         =   "Õœœ «·„·›"
            Height          =   375
            Left            =   7080
            TabIndex        =   68
            Top             =   120
            Width           =   1335
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   375
            Index           =   0
            Left            =   4320
            TabIndex        =   12
            Top             =   150
            Visible         =   0   'False
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   661
            Caption         =   "⁄—÷"
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
         Begin MSComctlLib.ProgressBar ProgBar 
            Height          =   240
            Left            =   8265
            TabIndex        =   11
            Top             =   210
            Visible         =   0   'False
            Width           =   5535
            _ExtentX        =   9763
            _ExtentY        =   423
            _Version        =   393216
            Appearance      =   0
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   375
            Index           =   1
            Left            =   60
            TabIndex        =   13
            Top             =   150
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   661
            Caption         =   "Œ—ÊÃ"
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
         Begin ImpulseButton.ISButton Cmd 
            Height          =   375
            Index           =   2
            Left            =   2895
            TabIndex        =   14
            Top             =   150
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   661
            Caption         =   " ⁄œÌ·"
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
         Begin ImpulseButton.ISButton Cmd 
            Height          =   375
            Index           =   3
            Left            =   1335
            TabIndex        =   15
            Top             =   150
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   661
            Caption         =   " —«Ã⁄"
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
         Begin ImpulseButton.ISButton Loadmex 
            Height          =   375
            Index           =   4
            Left            =   5640
            TabIndex        =   67
            Top             =   120
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   661
            Caption         =   " ÕœÌÀ «·’·«ÕÌ« "
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
         Begin MSComDlg.CommonDialog CD1 
            Left            =   4200
            Top             =   0
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
      End
   End
End
Attribute VB_Name = "FrmPermission"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Myfile As String


Private Sub cmbAccounts_KeyUp(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyF3 Then
        Account_search.show
        'Account_search.mIndex = Index
        Account_search.case_id = 7897286
    End If
End Sub

Private Sub fg3_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
  
    Dim rs As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim Msg As String
    Dim StrAccountCode As String
    Dim StrAccountCode1 As String

    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long

    With FG

        Select Case .ColKey(Col)
 
            Case "TasksName"
                
                .TextMatrix(Row, .ColIndex("TasksName")) = ""
                StrSQL = "select * from TblTasks "
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = FG.BuildComboList(rs, "Name", "ID")
                Else
                    StrComboList = FG.BuildComboList(rs, "Namee", "ID")
                End If
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If
                 .ComboList = StrComboList
            End Select
        End With
End Sub

Private Sub ISButton3_Click()
  Dim k As Long, LngNewRow As Long
  If Trim(FG3.TextMatrix(FG3.rows - 1, FG3.ColIndex("AccountName"))) = "" Then
        FG3.rows = FG3.rows - 1
    End If
    If FG3.rows = 1 Then FG3.rows = 2 Else FG3.rows = FG3.rows + 1
    
    
    k = FG3.rows
   
    If FG3.rows <= 1 Then
        FG3.rows = FG3.rows + 1
    End If
    LngNewRow = FG3.rows - 1
     'mNewId = LngNewRow
     
    
       
        
    
        
    FG3.TextMatrix(LngNewRow, FG3.ColIndex("AccountName")) = cmbAccounts.text
    FG3.TextMatrix(LngNewRow, FG3.ColIndex("AccountCode")) = cmbAccounts.BoundText
    
     
   
    
'    Fg_AfterEdit LngNewRow, fg3.ColIndex("TasksName")
End Sub


Private Sub Cmd_Click(Index As Integer)
    Dim i As Integer
    Dim Msg As String
    Dim StrSQL As String
    Dim rs As New ADODB.Recordset
    Dim IntRes As Integer

    Select Case Index

        Case 0

            If val(Me.DcboUsers.BoundText) = 0 Then
            '    GetMsgs 204, vbExclamation
                Exit Sub
            End If

            StrSQL = "SELECT TblUsers.* From TblUsers WHERE TblUsers.UserID= " & Me.DcboUsers.BoundText & ""
            Set rs = Nothing
            rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

            If Not IsNull(rs("FullPremis").value) Then
                If rs("FullPremis").value = 0 Then
                    'NO Permis
                
                   If opt(3).value = False Then
                          opt(1).value = True
                    Opt_Click (1)
                      End If
                      
                    Me.ChkInvAbility.value = vbUnchecked
                    Me.ChkInvProfit.value = vbUnchecked
                ElseIf rs("FullPremis").value = 1 Then
                    'Full
                    If opt(3).value = False Then
                    opt(0).value = True
                    
                    Opt_Click (0)
                    End If
                    Me.ChkInvAbility.value = vbChecked
                    Me.ChkInvProfit.value = vbChecked
                ElseIf rs("FullPremis").value = 2 Then
                    'Custome
                   ' Opt(2).value = True
                    If opt(3).value = False Then
                      opt(2).value = True
                      End If
                '    Opt_Click (2)
                    LoadPremis Me.DcboUsers.BoundText
                End If

            Else
                opt(2).value = True
                Opt_Click (2)
       
            End If
         LoadPremis Me.DcboUsers.BoundText
            rs.Close
            Set rs = Nothing

        Case 1
            Unload Me

        Case 2

            If Me.TxtModFlg.text = "N" Then
                SavePremis
                TxtModFlg.text = "R"
            ElseIf Me.TxtModFlg.text = "R" Then

                If val(Me.DcboUsers.BoundText) = 0 Then
                    'GetMsgs 204, vbExclamation
                    Msg = "ÌÃ» «‰  ﬁÊ„ »≈Œ Ì«— «·„” Œœ„ «·„—«œ  ÕœÌœ ’·«ÕÌ« Â...!!!"
                    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    Me.DcboUsers.SetFocus
                    'SendKeys "{TAB}"
                    Exit Sub
                End If

                TxtModFlg.text = "N"
            End If

        Case 3
            IntRes = GetMsgs(205, vbQuestion + vbYesNo)

            If IntRes = vbYes Then
                Cmd_Click (0)
                TxtModFlg.text = "R"
            End If
    
    End Select

End Sub

Private Sub CmdImport_Click()
CD1.ShowOpen
Myfile = CD1.FileName

End Sub

Private Sub DcboUsers_Change()
On Error Resume Next
    If val(DcboUsers.BoundText) = 0 Then Exit Sub

    
    FG3.rows = 1
    
    TxtCode.text = GetusercodeByid(DcboUsers.BoundText)
 
 
    Cmd_Click (0)
End Sub

Private Sub DcboUsers1_Change()
    If val(DcboUsers1.BoundText) = 0 Then Exit Sub

    
 
    
    txtCode1.text = GetusercodeByid(DcboUsers1.BoundText)
 
  LoadPremis val(Me.DcboUsers1.BoundText)
    'Cmd_Click (0)
End Sub

Private Sub DcboUsers_Click(Area As Integer)

'    If val(Me.DcboUsers.BoundText) <> 0 Then
'        DcboUsers_Change
'    End If

End Sub

Private Sub DcboUsers1_Click(Area As Integer)
DcboUsers1_Change
End Sub

Private Sub FG_AfterEdit(ByVal Row As Long, _
                         ByVal Col As Long)
    Dim i As Integer
    Dim IntScreenType As Integer

    With FG

        Select Case .ColKey(Col)

            Case "FullAccess"

                If Row = .FixedRows - 1 Then

                    For i = Row To .rows - 1
                        IntScreenType = val((.TextMatrix(i, .ColIndex("ScreenType"))))

                        If Not .IsSubtotal(i) And IntScreenType < 50 Then
                            .cell(flexcpChecked, i, .ColIndex("AddNew"), i, .ColIndex("FullAccess")) = .cell(flexcpChecked, Row, Col)
                        ElseIf Not .IsSubtotal(i) And IntScreenType >= 5 Then
                            .cell(flexcpChecked, i, .ColIndex("FullAccess")) = .cell(flexcpChecked, Row, Col)
                        End If

                    Next i

                End If

                IntScreenType = val((.TextMatrix(Row, .ColIndex("ScreenType"))))

                If IntScreenType < 50 Then
                    .cell(flexcpChecked, Row, .ColIndex("AddNew"), Row, .ColIndex("Atta")) = .cell(flexcpChecked, Row, Col)
                Else
                    .cell(flexcpChecked, Row, .ColIndex("FullAccess")) = .cell(flexcpChecked, Row, Col)
                End If

            Case "NoAccess"
            Case "CanShow"

           If Row = .FixedRows - 1 Then

                    For i = Row To .rows - 1
                        IntScreenType = val((.TextMatrix(i, .ColIndex("ScreenType"))))

                        If Not .IsSubtotal(i) And IntScreenType < 50 Then
                            .cell(flexcpChecked, i, .ColIndex("CanShow"), i, .ColIndex("CanShow")) = .cell(flexcpChecked, Row, Col)
                        End If

                    Next i

                End If
                
            Case "AddNew"

                If Row = .FixedRows - 1 Then

                    For i = Row To .rows - 1
                        IntScreenType = val((.TextMatrix(i, .ColIndex("ScreenType"))))

                        If Not .IsSubtotal(i) And IntScreenType < 50 Then
                            .cell(flexcpChecked, i, .ColIndex("AddNew"), i, .ColIndex("AddNew")) = .cell(flexcpChecked, Row, Col)
                        End If

                    Next i

                End If

            Case "Edit"

                If Row = .FixedRows - 1 Then

                    For i = Row To .rows - 1
                        IntScreenType = val((.TextMatrix(i, .ColIndex("ScreenType"))))

                        If Not .IsSubtotal(i) And IntScreenType < 50 Then
                            .cell(flexcpChecked, i, .ColIndex("Edit"), i, .ColIndex("Edit")) = .cell(flexcpChecked, Row, Col)
                        End If

                    Next i

                End If

            Case "Print"

                If Row = .FixedRows - 1 Then

                    For i = Row To .rows - 1
                        IntScreenType = val((.TextMatrix(i, .ColIndex("ScreenType"))))

                        If Not .IsSubtotal(i) And IntScreenType < 50 Then
                            .cell(flexcpChecked, i, .ColIndex("Print"), i, .ColIndex("Print")) = .cell(flexcpChecked, Row, Col)
                        End If

                    Next i

                End If

            Case "Delete"

                If Row = .FixedRows - 1 Then

                    For i = Row To .rows - 1
                        IntScreenType = val((.TextMatrix(i, .ColIndex("ScreenType"))))

                        If Not .IsSubtotal(i) And IntScreenType < 50 Then
                            .cell(flexcpChecked, i, .ColIndex("Delete"), i, .ColIndex("Delete")) = .cell(flexcpChecked, Row, Col)
                        End If

                    Next i

                End If

            Case "Search"

                If Row = .FixedRows - 1 Then

                    For i = Row To .rows - 1
                        IntScreenType = val((.TextMatrix(i, .ColIndex("ScreenType"))))

                        If Not .IsSubtotal(i) And IntScreenType < 50 Then
                            .cell(flexcpChecked, i, .ColIndex("Search"), i, .ColIndex("Search")) = .cell(flexcpChecked, Row, Col)
                        End If

                    Next i

                End If
                '############################################ Khaled Was Here ###########################################################
                Case "Atta"

                  If Row = .FixedRows - 1 Then

                    For i = Row To .rows - 1
                        IntScreenType = val((.TextMatrix(i, .ColIndex("ScreenType"))))

                        If Not .IsSubtotal(i) And IntScreenType < 50 Then
                            .cell(flexcpChecked, i, .ColIndex("Atta"), i, .ColIndex("Atta")) = .cell(flexcpChecked, Row, Col)
                        End If
                        
                    Next i

                End If
                '#########################################################################################################################

        End Select

        CellCheck CInt(Row)
        relign (.TextMatrix(Row, .ColIndex("Frm_Name"))), Row, Col
    End With

End Sub
Function relign(formname As String, Row As Long, Col As Long)
Dim Frm_Name As String
   Dim i As Integer
    With FG
                    For i = 1 To .rows - 1
                        Frm_Name = ((.TextMatrix(i, .ColIndex("Frm_Name"))))

                        If Frm_Name = formname And Frm_Name <> "" Then
                         .cell(flexcpChecked, i, .ColIndex("FullAccess")) = .cell(flexcpChecked, Row, .ColIndex("FullAccess"))
                            .cell(flexcpChecked, i, .ColIndex("AddNew"), i, .ColIndex("AddNew")) = .cell(flexcpChecked, Row, .ColIndex("AddNew"))
                            .cell(flexcpChecked, i, .ColIndex("Edit"), i, .ColIndex("Edit")) = .cell(flexcpChecked, Row, .ColIndex("Edit"))
                            .cell(flexcpChecked, i, .ColIndex("Delete"), i, .ColIndex("Delete")) = .cell(flexcpChecked, Row, .ColIndex("Delete"))
                            .cell(flexcpChecked, i, .ColIndex("Search"), i, .ColIndex("Search")) = .cell(flexcpChecked, Row, .ColIndex("Search"))
                            '############################################ Khaled Was Here ###########################################################
                            .cell(flexcpChecked, i, .ColIndex("Atta"), i, .ColIndex("Atta")) = .cell(flexcpChecked, Row, .ColIndex("Atta"))
                            '########################################################################################################################
                            .cell(flexcpChecked, i, .ColIndex("Print"), i, .ColIndex("Print")) = .cell(flexcpChecked, Row, .ColIndex("Print"))
                            .cell(flexcpChecked, i, .ColIndex("CanShow"), i, .ColIndex("CanShow")) = .cell(flexcpChecked, Row, .ColIndex("CanShow"))
                            
                            '
                        End If

                    Next i

               
          End With
End Function

Private Sub fg_Click()

    With FG
    
        If .Row <= 0 Then Exit Sub
        If .Col <> 1 Then Exit Sub
        If Not .IsSubtotal(.Row) Then Exit Sub
        If .IsCollapsed(.Row) = flexOutlineCollapsed Then
            .IsCollapsed(.Row) = flexOutlineExpanded
        Else
            .IsCollapsed(.Row) = flexOutlineCollapsed
        End If

    End With

End Sub

Private Sub fg_StartEdit(ByVal Row As Long, _
                         ByVal Col As Long, _
                         Cancel As Boolean)
    Dim Msg As String
    Dim IntRes As Integer

    If Me.opt(0).value = True Then
        
        
   If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "·ﬁœ ﬁ„  »≈Œ Ì«— ≈⁄ÿ«¡ ’·«ÕÌ…  «„…" & CHR(13)
        Msg = Msg & "··„” Œœ„ ,,," & CHR(13)
        Msg = Msg & "Â·  —Ìœ  Œ’Ì’ ’·«ÕÌ«  «·„” Œœ„..øø"
    Else
      Msg = "this User have full Permissions" & CHR(13)
      '  Msg = Msg & "··„” Œœ„ ,,," & Chr(13)
        Msg = Msg & "You need Custom permissions..øø"
    
    End If
    
        
        
        IntRes = MsgBox(Msg, vbQuestion + vbYesNo + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title)

        Select Case IntRes

            Case vbYes
                Cancel = False
                opt(2).value = True

            Case vbNo
                Cancel = True
        End Select
    
    ElseIf Me.opt(1).value = True Then
 If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "·ﬁœ ﬁ„  »≈·€«¡ Ã„Ì⁄ ’·«ÕÌ«  «·„” Œœ„" & CHR(13)
        Msg = Msg & "Â·  —Ìœ  Œ’Ì’ ’·«ÕÌ«  «·„” Œœ„..øø"
  Else
   
      Msg = "this User have No Permissions" & CHR(13)
      '  Msg = Msg & "··„” Œœ„ ,,," & Chr(13)
        Msg = Msg & "You need Custom permissions..øø"
   
   
  End If
        
        IntRes = MsgBox(Msg, vbQuestion + vbYesNo + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title)

        Select Case IntRes

            Case vbYes
                Cancel = False
                opt(2).value = True

            Case vbNo
                Cancel = True
        End Select

    End If

End Sub

Private Sub Form_Load()
On Error Resume Next
    Dim StrSQL As String
    Dim rs As New ADODB.Recordset
    Dim GrdPic          As New ClsBackGroundPic
    Dim i               As Integer
    Dim RowOutLever     As Integer
    Dim IntOldType      As Integer
    Dim RowCounter      As Integer
    Dim BolRtl          As Boolean
    Dim Dcombos As ClsDataCombos
    BolRtl = True

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
        BolRtl = False
    End If
    FG3.rows = 1
    'If SystemOptions.UserType = UserAdminAll Then
    '    StrSQL = "SELECT  USERS.User_ID,USERS.UserName  FROM USERS where " & _
    '        "(USERS.Password_Type =2 OR  USERS.Password_Type =3);"
    'ElseIf SystemOptions.UserType = UserAdmin Then
    '    StrSQL = "SELECT  USERS.User_ID,USERS.UserName  FROM USERS where " & _
    '        "(USERS.Password_Type =3);"
    'ElseIf SystemOptions.UserType = UserNourCo Then
    '    StrSQL = "SELECT  USERS.User_ID,USERS.UserName  FROM USERS"
    'End If
    
    Set Dcombos = New ClsDataCombos
                          If user_id <> 1 Then
    Dcombos.GetUsers Me.DcboUsers, False, True
    
    Dcombos.GetUsers Me.DcboUsers1, False, True
    
    Else
        Dcombos.GetUsers Me.DcboUsers, True
    Dcombos.GetUsers Me.DcboUsers1, True
    
    End If
    
    
  
   

    Dcombos.GetAccountingCodes cmbAccounts
    
    If SystemOptions.SysDataBaseType = AccessDataBase Then
        StrSQL = "SELECT Screens.ScreenName, Screens.ScreenCaption,Screens.ScreenTitleEng," & "Screens.ScreenType," & "S-creens.ScreenOder,ScreenImgKey  From Screens " & " Where Screens.ScreenVisible=True "
    ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
        StrSQL = "SELECT Screens.ScreenName, Screens.ScreenCaption,Screens.ScreenTitleEng," & "Screens.ScreenType," & "Screens.ScreenOrder,ScreenImgKey  From Screens " & " Where Screens.ScreenVisible=1 " & " AND ScreenType IS NOT NULL " & " AND ScreenOrder IS NOT NULL "
    End If
    StrSQL = StrSQL & " and ScreenType<>11"  '«·«”Â„
    
    If mdifrmmain.MnuProjects.Visible = False Then
StrSQL = StrSQL & " and ScreenType<>2"
End If

    If mdifrmmain.prdo.Visible = False Then
StrSQL = StrSQL & " and ScreenType<>3"
End If


    If mdifrmmain.StockControl.Visible = False Then
StrSQL = StrSQL & " and ScreenType<>4"
End If



    If mdifrmmain.Purchase.Visible = False Then
StrSQL = StrSQL & " and ScreenType<>5"
End If


    If mdifrmmain.Sales.Visible = False Then
'StrSQL = StrSQL & " and ScreenType<>6"
End If



    If mdifrmmain.Currency.Visible = False Then
StrSQL = StrSQL & " and ScreenType<>7"
End If



    If mdifrmmain.mnuEmployee.Visible = False Then
StrSQL = StrSQL & " and ScreenType<>8"
End If

' If mdifrmmain.MnuAccounts.Visible = False Then
'  StrSQL = StrSQL & " and ScreenType<>9"
'  End If
'
    If mdifrmmain.MNUFixedAssets.Visible = False Then
StrSQL = StrSQL & " and ScreenType<>10"
End If


  '  If mdifrmmain.MNUFixedAssets.Visible = false Then
'StrSQL = StrSQL & " and ScreenType<>11"«·«”Â„
'End If


    If mdifrmmain.AssetsMngBase.Visible = False Then
StrSQL = StrSQL & " and ScreenType<>12"
End If

'If mdifrmmain.Reports.Visible = False Then
'StrSQL = StrSQL & " and ScreenType<>13"
'End If


    If mdifrmmain.Tools.Visible = False Then
StrSQL = StrSQL & " and ScreenType<>14"
End If


    If mdifrmmain.POSTRansactiosG.Visible = False Then
'StrSQL = StrSQL & " and ScreenType<>15"
End If


    If mdifrmmain.TransporterMain.Visible = False Then '
StrSQL = StrSQL & " and ScreenType<>16"
End If



    If mdifrmmain.FinAnalysis.Visible = False Then '
StrSQL = StrSQL & " and ScreenType<>17"
End If


    If mdifrmmain.tech.Visible = False Then '
StrSQL = StrSQL & " and ScreenType<>18"
End If


    If mdifrmmain.tech.Visible = False Then '
StrSQL = StrSQL & " and ScreenType<>18"
End If


    'If mdifrmmain.tech.Visible = false Then '
'StrSQL = StrSQL & " and ScreenType<>19"
'End If

    If mdifrmmain.CarMaintenance.Visible = False Then '
StrSQL = StrSQL & " and ScreenType<>20"
End If


    If mdifrmmain.MarketingMnu.Visible = False Then '
StrSQL = StrSQL & " and ScreenType<>21"
End If


    If mdifrmmain.shipmentMnu.Visible = False Then '
StrSQL = StrSQL & " and ScreenType<>22"
End If


    If mdifrmmain.shipmentMnu.Visible = False Then '
StrSQL = StrSQL & " and ScreenType<>22"
End If


    'If mdifrmmain.shipmentMnu.Visible = false Then ' ‰»ÌÂ« 
'StrSQL = StrSQL & " and ScreenType<>23"
'End If

    If mdifrmmain.Strategy.Visible = False Then '
StrSQL = StrSQL & " and ScreenType<>24"
End If



    If mdifrmmain.dev.Visible = False Then '
StrSQL = StrSQL & " and ScreenType<>25"
End If

    If mdifrmmain.rsInvestment.Visible = False Then '
StrSQL = StrSQL & " and ScreenType<>26"
End If


    If mdifrmmain.MnuMaintnance.Visible = False Then '
StrSQL = StrSQL & " and ScreenType<>27"
End If


    If mdifrmmain.SalesIns.Visible = False Then ' ﬁ”Ìÿ
StrSQL = StrSQL & " and ScreenType<>28"
End If


    If mdifrmmain.MnuElevators.Visible = False Then '„’«⁄œ
StrSQL = StrSQL & " and ScreenType<>29"
End If


    If mdifrmmain.hajMnu.Visible = False Then '
StrSQL = StrSQL & " and ScreenType<>30"
End If


    If mdifrmmain.StudentMenue.Visible = False Then '
StrSQL = StrSQL & " and ScreenType<>31"
End If



'     If mdifrmmain.AgeingMAster.Visible = False Then '
'      StrSQL = StrSQL & " and ScreenType<>32"
'  End If



    If mdifrmmain.Archiving.Visible = False Then '
StrSQL = StrSQL & " and ScreenType<>33"
End If


    If mdifrmmain.taxes.Visible = False Then '
StrSQL = StrSQL & " and ScreenType<>34"
End If

    If mdifrmmain.ProductionPlan.Visible = False Then '
StrSQL = StrSQL & " and ScreenType<>35"
End If












StrSQL = StrSQL & "  ORDER BY Screens.ScreenType, Screens.ScreenOrder "
 

 

    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    With FG
        .rows = .FixedRows + 1
        .ExtendLastCol = True
        .GridLines = flexGridNone
        .RowHeightMin = 300
        .FixedCols = 1
        .FixedRows = 2
        .ExplorerBar = flexExMove
        .MergeCells = flexMergeSpill
        .OutlineCol = .ColIndex("TreeCol")
        .OutlineBar = flexOutlineBarComplete
        .NodeClosedPicture = Me.ImgLstScreens.ListImages("GroupImg").ExtractIcon
        .NodeOpenPicture = Me.ImgLstScreens.ListImages("GroupImgOpen").ExtractIcon
        .MergeCol(.ColIndex("UserName")) = True
        .MergeCol(.ColIndex("AddNew")) = True
        .ColWidth(.ColIndex("AddNew")) = 1240
        .MergeCol(.ColIndex("Edit")) = True
        .ColWidth(.ColIndex("Edit")) = 1240
        .MergeCol(.ColIndex("Delete")) = True
        .ColWidth(.ColIndex("Delete")) = 1240
        .MergeCol(.ColIndex("Print")) = True
        .ColWidth(.ColIndex("Print")) = 1240
        .MergeCol(.ColIndex("Search")) = True
        .ColWidth(.ColIndex("Search")) = 1240
        '############# Khaled was here ##############
        .MergeCol(.ColIndex("Atta")) = True
        .ColWidth(.ColIndex("Atta")) = 1240
        '############################################
        .MergeCol(.ColIndex("FullAccess")) = True
        .ColWidth(.ColIndex("FullAccess")) = 1240
        .MergeCol(.ColIndex("NoAccess")) = True
        .MergeCol(.ColIndex("UserName")) = True

        If BolRtl = True Then   ' add arabic caption
            .cell(flexcpText, 0, .ColIndex("AddNew"), 0) = "≈÷«›… ”Ã·"
            .cell(flexcpText, 0, .ColIndex("Edit"), 0) = " ⁄œÌ· ”Ã·"
            .cell(flexcpText, 0, .ColIndex("Delete"), 0) = "Õ–› ”Ã·"
            .cell(flexcpText, 0, .ColIndex("Print"), 0) = "ÿ»«⁄…"
            .cell(flexcpText, 0, .ColIndex("Search"), 0) = "»ÕÀ"
            '########################## KHALED WAS HERE ##################################
            .cell(flexcpText, 0, .ColIndex("Atta"), 0) = "«·„—›ﬁ« "
            '#############################################################################
            .cell(flexcpText, 0, .ColIndex("FullAccess"), 0) = "’·«ÕÌ…  «„…"
            .cell(flexcpText, 0, .ColIndex("NoAccess"), 0) = "≈·€«¡ Ã„Ì⁄ «·’·«ÕÌ« "
        Else                    'add english captions
            .cell(flexcpText, 0, .ColIndex("AddNew"), 0) = "Add New Record"
            .cell(flexcpText, 0, .ColIndex("Edit"), 0) = "Edit Record"
            .cell(flexcpText, 0, .ColIndex("Delete"), 0) = "Delete Record"
            .cell(flexcpText, 0, .ColIndex("Print"), 0) = "Print"
            .cell(flexcpText, 0, .ColIndex("Search"), 0) = "Search"
            '########################## KHALED WAS HERE ##################################
            .cell(flexcpText, 0, .ColIndex("Atta"), 0) = "Attachments"
            '#############################################################################
            .cell(flexcpText, 0, .ColIndex("CanShow"), 0) = "Show only"
            
            .cell(flexcpText, 0, .ColIndex("FullAccess"), 0) = "Full Access"
            .cell(flexcpText, 0, .ColIndex("NoAccess"), 0) = "Deny Access"
        End If

        'Add Fixed Cols Icons
        .cell(flexcpPicture, 0, .ColIndex("AddNew"), 0) = mdifrmmain.ImgLstMenuIcons.ListImages("New").ExtractIcon
        .cell(flexcpPicture, 0, .ColIndex("Edit"), 0) = mdifrmmain.ImgLstMenuIcons.ListImages("Edit").ExtractIcon
        .cell(flexcpPicture, 0, .ColIndex("Delete"), 0) = mdifrmmain.ImgLstMenuIcons.ListImages("Delete").ExtractIcon
        .cell(flexcpPicture, 0, .ColIndex("Print"), 0) = mdifrmmain.ImgLstMenuIcons.ListImages("Print").ExtractIcon
        .cell(flexcpPicture, 0, .ColIndex("Search"), 0) = mdifrmmain.ImgLstMenuIcons.ListImages("Find").ExtractIcon
        '############################################## Khaled Was Here #################################################
        .cell(flexcpPicture, 0, .ColIndex("Atta"), 0) = mdifrmmain.ImgLstMenuIcons.ListImages("New").ExtractIcon
        '################################################################################################################
        .cell(flexcpPicture, 0, .ColIndex("FullAccess"), 0) = mdifrmmain.ImgLstMenuIcons.ListImages("Tick").ExtractIcon
        .cell(flexcpPicture, 1, .ColIndex("NoAccess"), 1) = mdifrmmain.ImgLstMenuIcons.ListImages("Stop").ExtractIcon

        If BolRtl = True Then
            .cell(flexcpPictureAlignment, 0, .ColIndex("AddNew"), 0, .ColIndex("FullAccess")) = flexAlignRightCenter
        Else
            .cell(flexcpPictureAlignment, 0, .ColIndex("AddNew"), 0, .ColIndex("FullAccess")) = flexAlignLeftCenter
        End If

        .cell(flexcpPictureAlignment, 1, .ColIndex("AddNew"), 1, .ColIndex("FullAccess")) = flexAlignCenterCenter

        .cell(flexcpChecked, 1, .ColIndex("AddNew"), 1, .ColIndex("FullAccess")) = flexUnchecked
        .ColWidth(.ColIndex("TreeCol")) = 2800

        If BolRtl = True Then
            .AddItem "‘«‘«  «·»—‰«„Ã"
        Else
            .AddItem "Programe Windows"
        End If

        .IsSubtotal(.rows - 1) = True
        .cell(flexcpFontBold, .rows - 1, 1) = True
        .RowOutlineLevel(.rows - 1) = 0
        rs.MoveFirst

        Do While Not rs.EOF
            RowCounter = RowCounter + 1

            If IntOldType <> rs("ScreenType").value Then
                .AddItem ScreenType(rs("ScreenType").value, BolRtl)
                .IsSubtotal(.rows - 1) = True
                .RowOutlineLevel(.rows - 1) = 1
                .cell(flexcpForeColor, .rows - 1, .ColIndex("TreeCol")) = vbBlue
                IntOldType = rs("ScreenType").value
            End If

            If BolRtl = True Then
                .AddItem IIf(IsNull(rs("ScreenCaption").value), "", rs("ScreenCaption").value) ' rs("ScreenCaption").value
            Else
                .AddItem IIf(IsNull(rs("ScreenTitleEng").value), rs("ScreenCaption").value, rs("ScreenTitleEng").value)
            End If

            .TextMatrix(.rows - 1, .ColIndex("Frm_Name")) = IIf(IsNull(rs("ScreenName").value), "", rs("ScreenName").value)
            .TextMatrix(.rows - 1, .ColIndex("ScreenType")) = rs("ScreenType").value

            If Not IsNull(rs("ScreenImgKey").value) Then
                If rs("ScreenImgKey").value <> "" Then
                    .cell(flexcpPicture, .rows - 1, .ColIndex("TreeCol")) = mdifrmmain.ImgLstMenuIcons.ListImages(Trim(CStr(rs("ScreenImgKey").value))).ExtractIcon
                
                End If

            Else
                .cell(flexcpPicture, .rows - 1, .ColIndex("TreeCol")) = Me.ImgLstScreens.ListImages("ScreenImg").ExtractIcon
            End If

            If BolRtl = True Then
                .cell(flexcpPictureAlignment, .rows - 1, .ColIndex("TreeCol")) = flexAlignRightCenter
            Else
                .cell(flexcpPictureAlignment, .rows - 1, .ColIndex("TreeCol")) = flexAlignLeftCenter
            End If

            If rs("ScreenType").value = 50 Or rs("ScreenType").value = 60 Or rs("ScreenType").value = 70 Then
                .cell(flexcpChecked, .rows - 1, .ColIndex("AddNew")) = flexNoCheckbox
                .cell(flexcpChecked, .rows - 1, .ColIndex("Edit")) = flexNoCheckbox
                .cell(flexcpChecked, .rows - 1, .ColIndex("Delete")) = flexNoCheckbox
                .cell(flexcpChecked, .rows - 1, .ColIndex("Search")) = flexNoCheckbox
                '########################### Khaled Was Here ##################################
                .cell(flexcpChecked, .rows - 1, .ColIndex("Atta")) = flexNoCheckbox
                '##############################################################################
                .cell(flexcpChecked, .rows - 1, .ColIndex("FullAccess")) = flexUnchecked
                .cell(flexcpPictureAlignment, .rows - 1, .ColIndex("AddNew"), .rows - 1, .ColIndex("FullAccess")) = flexAlignCenterCenter
            Else
                .cell(flexcpChecked, .rows - 1, .ColIndex("AddNew"), .rows - 1, .ColIndex("FullAccess")) = flexUnchecked
                .cell(flexcpPictureAlignment, .rows - 1, .ColIndex("AddNew"), .rows - 1, .ColIndex("FullAccess")) = flexAlignCenterCenter
            End If

            rs.MoveNext
        Loop

        FormateGrid
         .WallPaper = GrdPic.Picture
          
    End With

    'Resize_Form Me, ReportSize
    TxtModFlg.text = "R"
    Me.TabMain.CurrTab = 0
End Sub

Private Function ScreenType(IntScreenType As Integer, _
                            Optional RTL As Boolean = True) As String
On Error Resume Next
    If SystemOptions.SysAppAccoutingType = SimpleAccoutning Then
        If RTL = True Then
            ScreenType = Choose(IntScreenType, " »Ì«‰«  √”«”Ì…", "„⁄«„·«   Ã«—Ì…", "„⁄«„·«  „«·Ì…", "‘∆Ê‰ «·„ÊŸ›Ì‰", " ﬁ«—Ì—", "≈” ⁄·«„« ", "√œÊ«  «·‰Ÿ«„")
        Else
            ScreenType = Choose(IntScreenType, "Basic Data", "Inventory", "Financial transactions", "  HR", "Reports", "Inforamtion", "System Tools", "POS")
        End If

    Else
 
        If RTL = True Then
            ScreenType = Choose(IntScreenType, " »Ì«‰«  √”«”Ì…", "«œ«—… «·„‘«—Ì⁄", "«·«‰ «Ã Ê«Ê«„— «·‘€·  ", " „—«ﬁ»Â «·„Œ“Ê‰", "«·„‘ —Ì« ", "«·„»Ì⁄« ", "«·„⁄«„·«  «·„«·Ì…  ", "‘∆Ê‰ «·„ÊŸ›Ì‰", "«·Õ”«»« ", "«·«’Ê· «·À«» Â", "„ «»⁄Â «·«”Â„", "«œ«—… «·«„·«ﬂ", "«· ﬁ«—Ì— «·⁄«„Â", "„œÌ— «·‰Ÿ«„", "‰ﬁ«ÿ «·»Ì⁄", "ﬁÿ«⁄    «·‰ﬁ·Ì«  ", "  «· Õ·Ì· «·„«·Ì ", "«·«œÊ«  «·›‰Ì…", "≈⁄œ«œ«  «·‰Ÿ«„ ", "’Ì«‰… «·„⁄œ« /«·”Ì«—« ", "«· ”ÊÌﬁ", "«·‘Õ‰", "«· ‰»ÌÂ« ", "«·‰ﬁ· «·„œ—”Ì", "«· ÿÊÌ—", "«·„”«Â„« ", "«·’Ì«‰…", "«· ﬁ”Ìÿ", "«œ«—… «·„’«⁄œ", "«·ÕÃ Ê «·⁄„—…", "«·„⁄«Âœ «· ⁄·Ì„Ì…", "«⁄„«— «·œÌÊ‰", "«·«—‘Ì›", " «·ﬁÌ„Â «·„÷«›… ", "«„—«ﬁ»… ·ÃÊœ…", "«·„ﬁ«Ì”« ", "«· Õ’Ì·« ", "«·Õ«ÊÌ« ", "«·‘∆Ê‰ «·ﬁ«‰Ê‰Ì…", "≈œ«—… «·„‘«€·", "  «·„⁄œ« /«·”Ì«—« ", "«· Ã„Ì·", "«·»’—Ì« ", "«·„“—⁄Â")
        Else
            ScreenType = Choose(IntScreenType, "Basic Data", "Project Mangements", "Production", "   Stock Control", "Purchase", "Sales", "Financial Transactions", "Hr", "Accounting", "Fixed Assets", " Arrows Follow", " ÒRealState Mangement", "Reports", "  System Manger", "pos", "Transportation", "Financial Analysis", "Technical. Tools", "Settings", "Cars Maintenance ", "Marketing", "Shipment", "Warnings", "Bus Transportation", "Task Mangement", "Real State Investment", "Maintenance", "Installments", "Elevators", "Hajj and Omra ", "Training Course", "Agening", "Archiving", "Taxes", "Quality", "Measurment", "Collections", "Container", " Legal Issue", "Salons", "Car Rent", "Beauty", "Optics", "Farm")
            
            
        End If
    End If

End Function

Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)
    'Dim Msg As String
    'Dim IntRes As Integer
    'If Me.TxtModFlg.text = "N" Then
    '
    '    IntRes = QueryCloseMsg("E", Me.Caption)
    '    Select Case IntRes
    '        Case vbYes
    '            Cancel = True
    '            Cmd_Click (2)
    '        Case vbNo
    '            Cancel = False
    '        Case vbCancel
    '            Cancel = True
    '    End Select
    '
    'End If
End Sub

Private Sub Loadmex_Click(Index As Integer)
On Error Resume Next
C1Elastic1.Visible = True
Dim sql As String
If Myfile = "" Then MsgBox "Õœœ «·„·› «Ê·«": Exit Sub
Dim ExcelObj As Object
    Dim ExcelBook As Object
    Dim ExcelSheet As Object
    Dim i As Integer
Dim currentvalue As String

Dim BranchID As String
Dim account_serial As String
Dim des As String
Dim DebitValue As String
Dim CreditValue As String
  

    Set ExcelObj = CreateObject("Excel.Application")
    Set ExcelSheet = CreateObject("Excel.Sheet")

    ExcelObj.Workbooks.Open Myfile   ' App.Path & "\TrialBalance.xls"
DoEvents
Cn.Execute " delete Screens"
    Set ExcelBook = ExcelObj.Workbooks(1)
    Set ExcelSheet = ExcelBook.Worksheets(1)
 
 Dim ScreenName As String
 
  Dim ScreenCaption As String
   Dim ScreenTitleEng As String
    Dim ScreenType As String
     Dim ScreenOrder As String
      Dim ScreenVisible As String
      
      Grid.rows = 1
    With ExcelSheet
    i = 1
    Do Until .cells(i, 2) & "" = ""
 '       Set l = lvwList.ListItems.Add(, , .Cells(i, 1))
' If i = 209 Then
' MsgBox ""
' End If
 
   ScreenName = .cells(i, 2)
    ScreenCaption = .cells(i, 3)
         ScreenTitleEng = .cells(i, 4)
        ScreenType = .cells(i, 5)
         ScreenOrder = .cells(i, 6)
         ScreenVisible = .cells(i, 8)
         
        Grid.rows = Grid.rows + 1
 With Grid

      .TextMatrix(i, .ColIndex("Ser")) = i
  .TextMatrix(i, .ColIndex("ScreenName")) = (ScreenName)
  .TextMatrix(i, .ColIndex("ScreenCaption")) = (ScreenCaption)
  .TextMatrix(i, .ColIndex("ScreenTitleEng")) = (ScreenTitleEng)
  .TextMatrix(i, .ColIndex("ScreenType")) = (ScreenType)
  .TextMatrix(i, .ColIndex("ScreenOrder")) = (ScreenOrder)
  .TextMatrix(i, .ColIndex("ScreenVisible")) = (ScreenVisible)
  
          Grid.Row = i
                            Grid.Col = Grid.ColIndex("ScreenName")
                            Grid.ShowCell i, Grid.ColIndex("ScreenName")
                            
                            Grid.SetFocus


 End With
 If .cells(i, 2) & "" = "" Then Exit Sub
        i = i + 1
    Loop

    End With

       ExcelObj.Workbooks.Close

    Set ExcelSheet = Nothing
    Set ExcelBook = Nothing
    Set ExcelObj = Nothing

 With Grid

For i = 1 To Grid.rows

 'If i = 209 Then
 'MsgBox ""
 'End If
  
  ScreenName = .TextMatrix(i, .ColIndex("ScreenName"))
  ScreenCaption = .TextMatrix(i, .ColIndex("ScreenCaption"))
  ScreenTitleEng = .TextMatrix(i, .ColIndex("ScreenTitleEng"))
  ScreenType = .TextMatrix(i, .ColIndex("ScreenType"))
  ScreenOrder = .TextMatrix(i, .ColIndex("ScreenOrder"))
  ScreenVisible = .TextMatrix(i, .ColIndex("ScreenVisible"))



sql = " insert into   Screens (ScreenName ,ScreenCaption , ScreenTitleEng , ScreenType , ScreenOrder , ScreenVisible )"
      sql = sql & "               Values ('" & ScreenName & "','" & ScreenCaption & "','" & ScreenTitleEng & "'," & ScreenType & "," & ScreenOrder & "," & ScreenVisible & ")"
      
      Cn.Execute sql
      
Next i
End With

End Sub

Private Sub Opt_Click(Index As Integer)

    Select Case Index

        Case 0
            FG.cell(flexcpChecked, 1, FG.ColIndex("FullAccess")) = flexChecked
            FG_AfterEdit 1, FG.ColIndex("FullAccess")
            Me.ChkInvAbility.value = vbChecked
            Me.ChkInvProfit.value = vbChecked

        Case 1
            FG.cell(flexcpChecked, 1, FG.ColIndex("FullAccess")) = flexUnchecked
            FG_AfterEdit 1, FG.ColIndex("FullAccess")
            Me.ChkInvAbility.value = vbUnchecked
            Me.ChkInvProfit.value = vbUnchecked

        Case 2
        
        Case 3
        DcboUsers1_Change
    End Select

End Sub

Private Sub SavePremis()
    Dim IntRes As Integer
    Dim i  As Integer
    Dim StrSQL As String
    Dim TransBegine As Boolean
    Dim BolAdd As Boolean, BolEdit As Boolean, BolDelete As Boolean, BolShow As Boolean
    '#################################### Khaled was here ######################################
    Dim BolPrint As Boolean, BolSearch  As Boolean, BolFullAccess As Boolean, BolAtta As Boolean
    '###########################################################################################
    Dim Msg As String
    Dim IntFullPre As Integer
    Dim IntScreenType As Integer
    Dim rs As ADODB.Recordset

    On Error GoTo ErrTrap
    Cn.BeginTrans
    TransBegine = True
    StrSQL = "Delete  From ScreenJuncUser Where User_ID=" & Me.DcboUsers.BoundText & ""
    Cn.Execute StrSQL, adExecuteNoRecords

    If Me.opt(1).value = True Then
        'IntRes = GetMsgs(206, vbQuestion + vbYesNo)
        'If IntRes = vbYes Then
        'StrSQL = "Update USERS Set IsActive = False  Where USERS.USER_ID=" & CLng(Trim(Me.DcboUsers.BoundText)) & ""
        'Cn.Execute StrSQL, adExecuteNoRecords
        'GoTo Exit_Sub
        'End If
    End If

    Set rs = New ADODB.Recordset
    StrSQL = "Select * From TblUsers Where UserID=" & Me.DcboUsers.BoundText
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If rs.BOF Or rs.EOF Then
        Exit Sub
    End If

    rs("InvPrices").value = IIf(Me.ChkInvAbility.value = vbChecked, 1, 0)
    rs("InvPrices1").value = IIf(Me.ChkInvAbility1.value = vbChecked, 1, 0)
    rs("InvPrices2").value = IIf(Me.ChkInvAbility2.value = vbChecked, 1, 0)
    rs("ShowInvProfit").value = IIf(Me.ChkInvProfit.value = vbChecked, 1, 0)
    rs("FixedCustomer").value = IIf(Me.ChkFixedCustomer.value = vbChecked, 1, 0)
    rs("ShowBillCommisions").value = IIf(Me.ChkShowCommisions.value = vbChecked, 1, 0)
    rs("hideCost").value = IIf(Me.chkHideCost.value = vbChecked, 1, 0)
    rs("hideColumn").value = IIf(Me.chkhideColumn.value = vbChecked, 1, 0)
    rs("ExceedShipment").value = IIf(Me.chkExceedShipment.value = vbChecked, 1, 0)
    rs("AllowSett").value = IIf(Me.AllowSett.value = vbChecked, 1, 0)
    rs("AllowSett1").value = IIf(Me.AllowSett1.value = vbChecked, 1, 0)
    rs("Allowpayroll").value = IIf(Me.Allowpayroll.value = vbChecked, 1, 0)
    rs("AllowCreateHajomraVoucher").value = IIf(Me.AllowCreateHajomraVoucher.value = vbChecked, 1, 0)
    rs("AllowBigAccount").value = IIf(Me.AllowBigAccount.value = vbChecked, 1, 0)
    rs("AllowRequestgl").value = IIf(Me.AllowRequestgl.value = vbChecked, 1, 0)
    rs("Allowrank").value = IIf(Me.chkDev.value = vbChecked, 1, 0)
    rs("AllowOrbonDate").value = IIf(Me.ChkAllowOrbonDate.value = vbChecked, 1, 0)
    rs("AllowCompChanPrice").value = IIf(Me.ChAllowCompChanPrice.value = vbChecked, 1, 0)
   '31032017egypt
    rs("AllowSalesSaveWithoutCostPrice").value = IIf(Me.Check(1).value = vbChecked, 1, 0)
    rs("AllowChanProjectBillPrice").value = IIf(Me.Check(0).value = vbChecked, 1, 0)
    rs("AllowChangeSalesAtTransfer").value = IIf(Me.Check(2).value = vbChecked, 1, 0)
    rs("AllowChangeUnitIqar").value = IIf(Me.Check(3).value = vbChecked, 1, 0)
    rs("AllowCreditPass").value = IIf(Me.Check(4).value = vbChecked, 1, 0)
    rs("AllowShowAllEmployee").value = IIf(Me.Check(5).value = vbChecked, 1, 0)
    rs("DateCanNotEdit").value = IIf(Me.Check(6).value = vbChecked, 1, 0)
    rs("BranchCanNotEdit").value = IIf(Me.Check(7).value = vbChecked, 1, 0)
    rs("PreFixCanNotEdit").value = IIf(Me.Check(8).value = vbChecked, 1, 0)
    rs("AllowPOSPAy").value = IIf(Me.Check(9).value = vbChecked, 1, 0)
    rs("AllowAprovedSalesBill").value = IIf(Me.Check(10).value = vbChecked, 1, 0)
    rs("AllowCraeJLQuality").value = IIf(Me.Check(11).value = vbChecked, 1, 0)
    rs("CantWorkwithComponenetinEmpScr").value = IIf(Me.Check(12).value = vbChecked, 1, 0)
    
    rs("AllowEditCreditLimit").value = IIf(Me.Check(13).value = vbChecked, 1, 0)
    rs("AllowEditCreditBalance").value = IIf(Me.Check(14).value = vbChecked, 1, 0)
    
    
    rs("USERautoIssueVoucher").value = IIf(Me.Check(40).value = vbChecked, 1, 0)
    rs("HideTbarInPos").value = IIf(Me.Check(45).value = vbChecked, 1, 0)
    
    
    rs("AllowConvertAlertToJob").value = IIf(Me.Check(15).value = vbChecked, 1, 0)
    rs("AllowSkipDiscountGroup").value = IIf(Me.Check(16).value = vbChecked, 1, 0)
    rs("OpenAtProduction").value = IIf(Me.Check(17).value = vbChecked, 1, 0)
    rs("HideInfroCasher").value = IIf(Me.Check(18).value = vbChecked, 1, 0)
    rs("CaNUpdateApprovedDoc").value = IIf(Me.Check(19).value = vbChecked, 1, 0)
    rs("CaNUpdateAutoSalesInvoice").value = IIf(Me.Check(20).value = vbChecked, 1, 0)
    rs("CanChangeStatusDateRequest").value = IIf(Me.Check(21).value = vbChecked, 1, 0)
    rs("CanChangeTripAfterInvoiceing").value = IIf(Me.Check(22).value = vbChecked, 1, 0)
    rs("CanChangeOut").value = IIf(Me.Check(23).value = vbChecked, 1, 0)
    rs("CanCancelContract").value = IIf(Me.Check(24).value = vbChecked, 1, 0)
    
    rs("CanCustomerandVendor").value = IIf(Me.Check(25).value = vbChecked, 1, 0)
    rs("CanEditCars").value = IIf(Me.Check(26).value = vbChecked, 1, 0)
    
    rs("CanEditOnlyPayMethod").value = IIf(Me.Check(27).value = vbChecked, 1, 0)
    
    rs("CanTransferItemDef").value = IIf(Me.Check(28).value = vbChecked, 1, 0)
    rs("CanPrintMultiSales").value = IIf(Me.Check(29).value = vbChecked, 1, 0)
    
    rs("CanPayWithoutPrint").value = IIf(Me.Check(30).value = vbChecked, 1, 0)
    rs("PlaywithAuthorityMatrix").value = IIf(Me.Check(31).value = vbChecked, 1, 0)
    rs("AllowEditProductionOutManulay").value = IIf(Me.Check(32).value = vbChecked, 1, 0)
    rs("AllowEditVaTManulay").value = IIf(Me.Check(33).value = vbChecked, 1, 0)
    
    
    rs("ShowOldAccountReports").value = IIf(Me.Check(34).value = vbChecked, 1, 0)
    rs("CanOpenWorkOrder").value = IIf(Me.Check(35).value = vbChecked, 1, 0)
    
    rs("CanChangePriceUpOnly").value = IIf(Me.Check(50).value = vbChecked, 1, 0)
    
    rs("CanProjectAccountOnly").value = IIf(Me.Check(51).value = vbChecked, 1, 0)
     rs("CanUploadZakat").value = IIf(Me.Check(52).value = vbChecked, 1, 0)
     rs("IsHiddenUser").value = IIf(Me.Check(53).value = vbChecked, 1, 0)
     rs("CanPostPumpInv").value = IIf(Me.Check(54).value = vbChecked, 1, 0)
    
    rs("CanAcreditRsContract").value = IIf(Me.Check(36).value = vbChecked, 1, 0)
    rs("CanIsShamel").value = IIf(Me.Check(46).value = vbChecked, 1, 0)
    rs("CanEditLegalAffairs").value = IIf(Me.Check(47).value = vbChecked, 1, 0)
    
    
    rs("OPenShortInvoice").value = IIf(Me.Check(37).value = vbChecked, 1, 0)
    
    rs("OPenShortInvoicePump").value = IIf(Me.Check(48).value = vbChecked, 1, 0)
    rs("OPenShortInvoicePetrol").value = IIf(Me.Check(49).value = vbChecked, 1, 0)
    
    rs("MonyeIssueVchrNoMust").value = IIf(Me.Check(38).value = vbChecked, 1, 0)
    rs("POMustentryAndBillMustEntry").value = IIf(Me.Check(39).value = vbChecked, 1, 0)
  
  
      rs("NotEditSalesRetPrice").value = IIf(Me.Check(41).value = vbChecked, 1, 0)
      rs("NotEditInternalPrice").value = IIf(Me.Check(42).value = vbChecked, 1, 0)
      rs("NotEditDiscountLine").value = IIf(Me.Check(43).value = vbChecked, 1, 0)
      
        rs("CanEditMinRentValue").value = IIf(Me.Check(44).value = vbChecked, 1, 0)
       

    
   '31032017egypt
   
   'ChkAllowOrbonDate
   
      
     
    rs.update
    Set rs = New ADODB.Recordset
    rs.Open "ScreenJuncUser", Cn, adOpenStatic, adLockOptimistic, adCmdTable

    With FG

        For i = 4 To .rows - 1

            DoEvents
            Me.ProgBar.Visible = True
            Me.ProgBar.Max = .rows - 1
            Me.ProgBar.value = i

            If Not .IsSubtotal(i) Then
                IntScreenType = val(.TextMatrix(i, .ColIndex("ScreenType")))

                If IntScreenType <> 50 And IntScreenType <> 60 And IntScreenType <> 70 Then
                    BolAdd = IIf(.cell(flexcpChecked, i, .ColIndex("AddNew")) = flexChecked, True, False)
                    BolEdit = IIf(.cell(flexcpChecked, i, .ColIndex("Edit")) = flexChecked, True, False)
                    BolDelete = IIf(.cell(flexcpChecked, i, .ColIndex("Delete")) = flexChecked, True, False)
                    BolPrint = IIf(.cell(flexcpChecked, i, .ColIndex("Print")) = flexChecked, True, False)
                    BolSearch = IIf(.cell(flexcpChecked, i, .ColIndex("Search")) = flexChecked, True, False)
                    '################################## Khaled Was Here ####################################
                    BolAtta = IIf(.cell(flexcpChecked, i, .ColIndex("Atta")) = flexChecked, True, False)
                    '#######################################################################################
                    BolShow = IIf(.cell(flexcpChecked, i, .ColIndex("CanShow")) = flexChecked, True, False)
                ElseIf IntScreenType = 50 Or IntScreenType = 60 Or IntScreenType = 70 Then
                    BolAdd = IIf(.cell(flexcpChecked, i, .ColIndex("FullAccess")) = flexChecked, True, False)
                    BolEdit = IIf(.cell(flexcpChecked, i, .ColIndex("FullAccess")) = flexChecked, True, False)
                    BolDelete = IIf(.cell(flexcpChecked, i, .ColIndex("FullAccess")) = flexChecked, True, False)
                    BolPrint = IIf(.cell(flexcpChecked, i, .ColIndex("FullAccess")) = flexChecked, True, False)
                    BolSearch = IIf(.cell(flexcpChecked, i, .ColIndex("FullAccess")) = flexChecked, True, False)
                    '################################## Khaled Was Here ####################################
                    BolAtta = IIf(.cell(flexcpChecked, i, .ColIndex("FullAccess")) = flexChecked, True, False)
                    '#######################################################################################
                    BolShow = IIf(.cell(flexcpChecked, i, .ColIndex("FullAccess")) = flexChecked, True, False)
                End If

                BolFullAccess = IIf(.cell(flexcpChecked, i, .ColIndex("FullAccess")) = flexChecked, True, False)
                rs.AddNew
                rs("ScreenName").value = .TextMatrix(i, .ColIndex("Frm_Name"))
                rs("User_ID").value = val(Me.DcboUsers.BoundText)

                If SystemOptions.SysDataBaseType = AccessDataBase Then
                    rs("CanAdd").value = BolAdd
                    rs("CanEdit").value = BolEdit
                    rs("CanDelete").value = BolDelete
                    rs("CanPrint").value = BolPrint
                    rs("CanSearch").value = BolSearch
                    '############# Khaled Was Here ###############
                    rs("Attachments").value = BolAtta
                    '#############################################
                    rs("FullAccess").value = BolFullAccess
                ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
                    rs("CanAdd").value = IIf(BolAdd = True, 1, 0)
                    rs("CanEdit").value = IIf(BolEdit = True, 1, 0)
                    rs("CanDelete").value = IIf(BolDelete = True, 1, 0)
                    rs("CanPrint").value = IIf(BolPrint = True, 1, 0)
                    rs("CanSearch").value = IIf(BolSearch = True, 1, 0)
                    '############# Khaled Was Here ###############
                    rs("Attachments").value = IIf(BolAtta = True, 1, 0)
                    '#############################################
                    rs("FullAccess").value = IIf(BolFullAccess = True, 1, 0)
                    rs("CanShow").value = IIf(BolShow = True, 1, 0)
                    
                End If

                rs.update
            End If

            DoEvents
            Me.ProgBar.value = i

            DoEvents
        Next i

        If opt(0).value = True Then
            IntFullPre = 1
        ElseIf opt(1).value = True Then
            IntFullPre = 0
        ElseIf opt(2).value = True Then
            IntFullPre = 2
          ElseIf opt(3).value = True Then
            IntFullPre = 2
            
        End If

        If SystemOptions.SysDataBaseType = AccessDataBase Then
            StrSQL = "Update TblUsers Set IsActive=True,FullPremis=" & IntFullPre & " Where TblUsers.UserID=" & Me.DcboUsers.BoundText & ""
        Else
            StrSQL = "Update TblUsers Set IsActive=1,FullPremis=" & IntFullPre & " Where TblUsers.UserID=" & Me.DcboUsers.BoundText & ""
        End If

        Cn.Execute StrSQL, , adExecuteNoRecords
        
        
        StrSQL = "Delete tblUserPermAccounts Where UserId =  " & val(Me.DcboUsers.BoundText) & ""
        Cn.Execute StrSQL
        StrSQL = " Select * from tblUserPermAccounts Where UserId = -1 "
        
        saveGrid StrSQL, FG3, "AccountCode", "", "UserId", val(Me.DcboUsers.BoundText)
    End With

Exit_Sub:

    'GetMsgs 70, vbInformation
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = " „  ⁄„·Ì… «·Õ›Ÿ...!!!"
    Else
        Msg = "     Saved ...!!!"
    End If

    MsgBox Msg, vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    Me.ProgBar.value = 0
    Me.ProgBar.Visible = False
    Cn.CommitTrans
    TransBegine = False
    Exit Sub
ErrTrap:
    'GetMsgs 71, vbInformation
  If SystemOptions.UserInterface = ArabicInterface Then
    Msg = "⁄›Ê«... ÕœÀ Œÿ« √À‰«¡ Õ›Ÿ «·’·«ÕÌ« ...!!!"
  Else
  Msg = "Sorry..... Error during Saving data"
  End If
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title

    'Resume
    If TransBegine Then
        Cn.RollbackTrans
    End If

End Sub

Private Sub TxtCode_Change()
    DcboUsers.BoundText = GeTuserIDByEmpCode(TxtCode.text)

End Sub


Private Sub TXTCode1_Change()
    DcboUsers1.BoundText = GeTuserIDByEmpCode(txtCode1.text)

End Sub

Private Sub TXTCode_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
 FrmUserSearch.show
 FrmUserSearch.lblSearchtype = 1

End If
End Sub

Private Sub TxtModFlg_Change()

    Select Case Me.TxtModFlg.text

        Case "R"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Cmd(2).Caption = " ⁄œÌ·"
            Else
                Me.Cmd(2).Caption = "&Edit"
            End If

            Me.FG.Editable = flexEDNone
            Me.Ele(0).Enabled = False
            Me.Cmd(0).Enabled = True
            Me.Cmd(1).Enabled = True
            Me.DcboUsers.Enabled = True
            Me.Cmd(3).Enabled = False
            opt(2).value = True
        
            'Me.ChkInvAbility.Enabled = False
            'Me.ChkInvProfit.Enabled = False

        Case "N"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Cmd(2).Caption = "Õ›Ÿ"
            Else
                Me.Cmd(2).Caption = "&Save"
            End If

            Me.FG.Editable = flexEDKbdMouse
            Me.Ele(0).Enabled = True
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.DcboUsers.Enabled = False
            Me.Cmd(3).Enabled = True
            Me.ChkInvAbility.Enabled = True
            Me.ChkInvProfit.Enabled = True
    End Select

End Sub

Private Sub ClearChecks()
    Dim i  As Integer
    Dim IntScreenType As String

    With FG

        For i = 4 To FG.rows - 1

            If Not FG.IsSubtotal(i) Then
                IntScreenType = val(.TextMatrix(i, .ColIndex("ScreenType")))

                If (IntScreenType < 50) Then
                    FG.cell(flexcpChecked, i, FG.ColIndex("AddNew"), i, FG.ColIndex("FullAccess")) = flexUnchecked
                Else
                    FG.cell(flexcpChecked, i, FG.ColIndex("FullAccess"), i, FG.ColIndex("FullAccess")) = flexUnchecked
                End If
            End If

        Next i

    End With

End Sub

Private Sub CellCheck(LngRow As Integer)
    Dim IntScreenType As Integer

    With FG
        IntScreenType = val(.TextMatrix(LngRow, .ColIndex("ScreenType")))
        
        If IntScreenType <> 50 And IntScreenType <> 60 And IntScreenType <> 70 Then
            '############################################### Khaled Was here #########################################################
            If .cell(flexcpChecked, LngRow, .ColIndex("AddNew")) = flexChecked And .cell(flexcpChecked, LngRow, .ColIndex("Edit")) = flexChecked And .cell(flexcpChecked, LngRow, .ColIndex("Delete")) = flexChecked And .cell(flexcpChecked, LngRow, .ColIndex("Print")) = flexChecked And .cell(flexcpChecked, LngRow, .ColIndex("Search")) = flexChecked And .cell(flexcpChecked, LngRow, .ColIndex("Atta")) = flexChecked Then
            '#########################################################################################################################
                .cell(flexcpChecked, LngRow, .ColIndex("FullAccess")) = flexChecked
            Else
                .cell(flexcpChecked, LngRow, .ColIndex("FullAccess")) = flexUnchecked
            End If
        End If

    End With

End Sub

Private Sub ChangeLang()
    Me.Caption = "Authority Matrix"
    LblUsers.Caption = "Users"
    opt(0).Caption = "Full Perm."
    opt(1).Caption = "Deny Perm."
    opt(2).Caption = "Custom Perm."
    opt(3).Caption = "As Perm."
    Check(4).Caption = "Allow Over Limit"
    Check(10).Caption = "Allow Approved Sales Invoice"
    Check(5).Caption = "Allow Show All Employees"
    Check(6).Caption = "Date Can Not Edit"
    Check(7).Caption = "Branch Can Not Edit"
    Check(8).Caption = "PreFix Can Not Edit"
    Check(9).Caption = "Allow POS PAy"
    Check(9).Caption = "Allow POS PAy"
    Label1.Caption = "Code"

    'Cmd(0).Caption = ""
    Cmd(1).Caption = "E&xit"
    Cmd(2).Caption = "&Edit"
    Cmd(3).Caption = "&Undo"

    TabMain.TabCaption(0) = "Screens Permissions"
    TabMain.TabCaption(1) = "Specific Permissions"
    ChkInvAbility.Caption = "Can Change Price In Sales Invoices"
    ChkInvProfit.Caption = "Can show Profit Sales Invoices"
    
   '############################### Khaled Was Here ######################################
    ChkInvAbility1.Caption = "He has the ability to adjust prices in stock exchange bonds"
    'ChkInvAbility2.Caption = ""
    ChkFixedCustomer.Caption = "Connect the user to his customers and his processes only"
    ChkShowCommisions.Caption = "Show commissions"
    chkHideCost.Caption = "Show cost"
    chkhideColumn.Caption = "Columns can not be modified in treasury bonds"
    chkExceedShipment.Caption = "The Authority of shipping more than the quantity required"
    AllowSett.Caption = "The Authority to conduct an inventory"
    AllowSett1.Caption = "The Authority to execute inventory settlement"
    Allowpayroll.Caption = "The Authority to execute payroll enrollment"
    AllowBigAccount.Caption = "The Authority to show the financial statements"
    AllowRequestgl.Caption = "The Authority to execute Exchange request for contractors"
    chkDev.Caption = "The Authority for performance evaluation"
    ChkAllowOrbonDate.Caption = "Allowance to exceed retainer duration"
    AllowCreateHajomraVoucher.Caption = "The validity to establish Hijj and Omra J L"
    ChAllowCompChanPrice.Caption = "The Authority to amend the prices of corporate agreements"
    Check(0).Caption = "The Authority of item price adjustment in project invoices"
    Check(1).Caption = "The Authorit to creat sales invoice in case there is no cost for the item"
    Check(2).Caption = "The validity of the sales invoice modification in the case of store transformation"
    Check(3).Caption = "Allow Change Unit in Real Estate"
    Check(11).Caption = "Allow Create Voucher Of Quality"
    Check(12).Caption = "Can't work with componenet in Employee Screen"
    Check(13).Caption = "The validity of the credit limit adjustment"
    Check(14).Caption = "The validity of the credit balance adjustment"
    
    
   '######################################################################################
    

End Sub

Private Sub LoadPremis(Lngid As Long)
    Dim i As Integer
    Dim LngGrdRow As Long
    Dim StrSQL As String
    Dim rs As ADODB.Recordset
    Dim IntScreenType As Integer

    If Lngid = 0 Then
        Exit Sub
    End If
If opt(3).Enabled = True And val(DcboUsers1.BoundText) <> 0 Then Lngid = val(DcboUsers1.BoundText)
    StrSQL = "Select * From TblUsers Where UserID=" & Lngid
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        Exit Sub
    End If

    Me.ChkInvAbility.value = IIf(rs("InvPrices").value = 1, vbChecked, vbUnchecked)
    Me.ChkInvAbility1.value = IIf(rs("InvPrices1").value = 1, vbChecked, vbUnchecked)
    Me.ChkInvAbility2.value = IIf(rs("InvPrices2").value = 1, vbChecked, vbUnchecked)
    
    Me.ChkInvProfit.value = IIf(rs("ShowInvProfit").value = 1, vbChecked, vbUnchecked)
    Me.ChkFixedCustomer.value = IIf(rs("FixedCustomer").value = 1, vbChecked, vbUnchecked)
     Me.ChkShowCommisions.value = IIf(rs("ShowBillCommisions").value = 1, vbChecked, vbUnchecked)
   chkHideCost.value = IIf(rs("HideCost").value = True, vbChecked, vbUnchecked)
   chkhideColumn.value = IIf(rs("hideColumn").value = True, vbChecked, vbUnchecked)
   chkExceedShipment.value = IIf(rs("ExceedShipment").value = True, vbChecked, vbUnchecked)
   
   AllowSett.value = IIf(rs("AllowSett").value = True, vbChecked, vbUnchecked)
   AllowSett1.value = IIf(rs("AllowSett1").value = True, vbChecked, vbUnchecked)
   Allowpayroll.value = IIf(rs("Allowpayroll").value = True, vbChecked, vbUnchecked)
   AllowCreateHajomraVoucher.value = IIf(rs("AllowCreateHajomraVoucher").value = True, vbChecked, vbUnchecked)
   
   AllowRequestgl.value = IIf(rs("AllowRequestgl").value = True, vbChecked, vbUnchecked)
   chkDev.value = IIf(rs("Allowrank").value = True, vbChecked, vbUnchecked)
   ChkAllowOrbonDate.value = IIf(rs("AllowOrbonDate").value = True, vbChecked, vbUnchecked)
   ChAllowCompChanPrice.value = IIf(rs("AllowCompChanPrice").value = True, vbChecked, vbUnchecked)
  '31032017egypt
Check(40).value = IIf(rs("USERautoIssueVoucher").value = True, vbChecked, vbUnchecked)
Check(45).value = IIf(rs("HideTbarInPos").value = True, vbChecked, vbUnchecked)

   Check(1).value = IIf(rs("AllowSalesSaveWithoutCostPrice").value = True, vbChecked, vbUnchecked)
   Check(0).value = IIf(rs("AllowChanProjectBillPrice").value = True, vbChecked, vbUnchecked)
   Check(2).value = IIf(rs("AllowChangeSalesAtTransfer").value = True, vbChecked, vbUnchecked)
   Check(3).value = IIf(rs("AllowChangeUnitIqar").value = True, vbChecked, vbUnchecked)
   Check(4).value = IIf(rs("AllowCreditPass").value = True, vbChecked, vbUnchecked)
   Check(5).value = IIf(rs("AllowShowAllEmployee").value = True, vbChecked, vbUnchecked)
   Check(6).value = IIf(rs("DateCanNotEdit").value = True, vbChecked, vbUnchecked)
   Check(7).value = IIf(rs("BranchCanNotEdit").value = True, vbChecked, vbUnchecked)
   Check(8).value = IIf(rs("PreFixCanNotEdit").value = True, vbChecked, vbUnchecked)
   Check(9).value = IIf(rs("AllowPOSPAy").value = True, vbChecked, vbUnchecked)
   Check(10).value = IIf(rs("AllowAprovedSalesBill").value = True, vbChecked, vbUnchecked)
   Check(11).value = IIf(rs("AllowCraeJLQuality").value = True, vbChecked, vbUnchecked)
   Check(12).value = IIf(rs("CantWorkwithComponenetinEmpScr").value = True, vbChecked, vbUnchecked)
   
   Check(13).value = IIf(rs("AllowEditCreditLimit").value = True, vbChecked, vbUnchecked)
   Check(14).value = IIf(rs("AllowEditCreditBalance").value = True, vbChecked, vbUnchecked)
   Check(15).value = IIf(rs("AllowConvertAlertToJob").value = True, vbChecked, vbUnchecked)
   Check(16).value = IIf(rs("AllowSkipDiscountGroup").value = True, vbChecked, vbUnchecked)
   Check(17).value = IIf(rs("OpenAtProduction").value = True, vbChecked, vbUnchecked)
   Check(18).value = IIf(rs("HideInfroCasher").value = True, vbChecked, vbUnchecked)
   Check(19).value = IIf(rs("CaNUpdateApprovedDoc").value = True, vbChecked, vbUnchecked)
      Check(20).value = IIf(rs("CaNUpdateAutoSalesInvoice").value = True, vbChecked, vbUnchecked)
      Check(21).value = IIf(rs("CanChangeStatusDateRequest").value = True, vbChecked, vbUnchecked)
      Check(22).value = IIf(rs("CanChangeTripAfterInvoiceing").value = True, vbChecked, vbUnchecked)
      Check(23).value = IIf(rs("CanChangeOut").value = True, vbChecked, vbUnchecked)
     Check(24).value = IIf(rs("CanCancelContract").value = True, vbChecked, vbUnchecked)
     Check(25).value = IIf(rs("CanCustomerandVendor").value = True, vbChecked, vbUnchecked)
     Check(26).value = IIf(rs("CanEditCars").value = True, vbChecked, vbUnchecked)
     Check(27).value = IIf(rs("CanEditOnlyPayMethod").value = True, vbChecked, vbUnchecked)
     
     Check(28).value = IIf(rs("CanTransferItemDef").value = True, vbChecked, vbUnchecked)
     Check(29).value = IIf(rs("CanPrintMultiSales").value = True, vbChecked, vbUnchecked)
     
     
     Check(30).value = IIf(rs("CanPayWithoutPrint").value = True, vbChecked, vbUnchecked)
     Check(31).value = IIf(rs("PlaywithAuthorityMatrix").value = True, vbChecked, vbUnchecked)
     Check(32).value = IIf(rs("AllowEditProductionOutManulay").value = True, vbChecked, vbUnchecked)
     Check(33).value = IIf(rs("AllowEditVaTManulay").value = True, vbChecked, vbUnchecked)
     Check(34).value = IIf(rs("ShowOldAccountReports").value = True, vbChecked, vbUnchecked)
     Check(35).value = IIf(rs("CanOpenWorkOrder").value = True, vbChecked, vbUnchecked)
     Check(50).value = IIf(rs("CanChangePriceUpOnly").value = True, vbChecked, vbUnchecked)
     Check(51).value = IIf(rs("CanProjectAccountOnly").value = True, vbChecked, vbUnchecked)
     Check(52).value = IIf(rs("CanUploadZakat").value = True, vbChecked, vbUnchecked)
     Check(53).value = IIf(rs("IsHiddenUser").value = True, vbChecked, vbUnchecked)
     Check(54).value = IIf(rs("CanPostPumpInv").value = True, vbChecked, vbUnchecked)
     
     
     
     
     Check(36).value = IIf(rs("CanAcreditRsContract").value = True, vbChecked, vbUnchecked)
     Check(46).value = IIf(rs("CanIsShamel").value = True, vbChecked, vbUnchecked)
     Check(47).value = IIf(rs("CanEditLegalAffairs").value = True, vbChecked, vbUnchecked)
     
     
     Check(37).value = IIf(rs("OPenShortInvoice").value = True, vbChecked, vbUnchecked)
    Check(48).value = IIf(rs("OPenShortInvoicePump").value = True, vbChecked, vbUnchecked)
    Check(49).value = IIf(rs("OPenShortInvoicePetrol").value = True, vbChecked, vbUnchecked)
    
 
     
     Check(38).value = IIf(rs("MonyeIssueVchrNoMust").value = True, vbChecked, vbUnchecked)
     Check(39).value = IIf(rs("POMustentryAndBillMustEntry").value = True, vbChecked, vbUnchecked)
     
     Check(41).value = IIf(rs("NotEditSalesRetPrice").value = True, vbChecked, vbUnchecked)
     Check(42).value = IIf(rs("NotEditInternalPrice").value = True, vbChecked, vbUnchecked)
     Check(43).value = IIf(rs("NotEditDiscountLine").value = True, vbChecked, vbUnchecked)
     Check(44).value = IIf(rs("CanEditMinRentValue").value = True, vbChecked, vbUnchecked)
     
     
      
      StrSQL = " Select *,AccountName = (Select Account_Name From Accounts Where Account_Code =tblUserPermAccounts.AccountCode  ) from tblUserPermAccounts Where UserId = " & val(DcboUsers.BoundText)
    
    loadgrid StrSQL, FG3, True, True
      
      
'   31032017egypt

   
   '
   AllowBigAccount.value = IIf(rs("AllowBigAccount").value = True, vbChecked, vbUnchecked)
   'AllowBigAccount
    StrSQL = "SELECT ScreenJuncUser.ScreenName, ScreenJuncUser.CanAdd,ScreenJuncUser.CanEdit,"
    StrSQL = StrSQL + "  ScreenJuncUser.CanShow,  ScreenJuncUser.CanDelete, ScreenJuncUser.CanPrint, ScreenJuncUser.CanSearch,"
    StrSQL = StrSQL + " ScreenJuncUser.FullAccess,TBLUSERS.[UserID] "
    StrSQL = StrSQL + " FROM TblUsers INNER JOIN ScreenJuncUser ON "
    StrSQL = StrSQL + " TblUsers.[UserID] =ScreenJuncUser.[User_ID] " & " Where ScreenJuncUser.[User_ID]=" & Lngid & ""
  '  StrSQL = StrSQL & " AND (dbo.ScreenJuncUser.CanAdd = 1) "
    StrSQL = StrSQL + " order by JuncID"
    
    
    
 '********************
'################################$###################### Khaled Was Here #################################################################
StrSQL = "SELECT     TOP 100 PERCENT dbo.ScreenJuncUser.ScreenName, dbo.ScreenJuncUser.CanAdd, dbo.ScreenJuncUser.CanEdit, dbo.ScreenJuncUser.CanShow, "
StrSQL = StrSQL + "  dbo.ScreenJuncUser.CanDelete , dbo.ScreenJuncUser.CanPrint, dbo.ScreenJuncUser.CanSearch, dbo.ScreenJuncUser.Attachments, dbo.ScreenJuncUser.FullAccess"
StrSQL = StrSQL + " FROM         dbo.TblUsers INNER JOIN"
StrSQL = StrSQL + "  dbo.ScreenJuncUser ON dbo.TblUsers.UserID = dbo.ScreenJuncUser.User_ID"
StrSQL = StrSQL + "   Where (dbo.TblUsers.UserID = " & Lngid & ")"
StrSQL = StrSQL + "   GROUP BY dbo.ScreenJuncUser.ScreenName, dbo.ScreenJuncUser.CanAdd, dbo.ScreenJuncUser.CanEdit, dbo.ScreenJuncUser.CanShow, dbo.ScreenJuncUser.CanDelete,"
StrSQL = StrSQL + "  dbo.ScreenJuncUser.CanPrint , dbo.ScreenJuncUser.CanSearch, dbo.ScreenJuncUser.Attachments, dbo.ScreenJuncUser.FullAccess"
StrSQL = StrSQL + "   HAVING      (dbo.ScreenJuncUser.CanAdd = 1) OR"
StrSQL = StrSQL + "   (dbo.ScreenJuncUser.CanEdit = 1) OR"
StrSQL = StrSQL + "    (dbo.ScreenJuncUser.CanShow = 1) OR"
StrSQL = StrSQL + "    (dbo.ScreenJuncUser.CanDelete = 1) OR"
StrSQL = StrSQL + "    (dbo.ScreenJuncUser.CanPrint = 1) OR"
StrSQL = StrSQL + "  (dbo.ScreenJuncUser.CanSearch = 1) OR"
StrSQL = StrSQL + "  (dbo.ScreenJuncUser.Attachments = 1) OR"
'############################################################################################################################################
StrSQL = StrSQL + "      (dbo.ScreenJuncUser.FullAccess = 1)"
 
 '/*********************
    
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        'Msg = "Â–« «·„” Œœ„ ·„  Õœœ ·Â ’·«ÕÌ« Â ·≈” Œœ«„ «·»—‰«„Ã"
        'MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        ClearChecks
        Exit Sub
    End If

    rs.MoveFirst

    With Me.FG
        ClearChecks

        For i = 0 To rs.RecordCount - 1
            LngGrdRow = .FindRow(rs("ScreenName").value, .FixedRows, .ColIndex("Frm_Name"), False, True)

            If LngGrdRow <> -1 Then
                IntScreenType = val(.TextMatrix(LngGrdRow, .ColIndex("ScreenType")))

                If IntScreenType < 50 Then
                    .cell(flexcpChecked, LngGrdRow, .ColIndex("CanShow")) = IIf(rs("CanShow").value = False, flexUnchecked, IIf(IsNull(rs("CanShow").value), flexUnchecked, flexChecked))
                
                    .cell(flexcpChecked, LngGrdRow, .ColIndex("AddNew")) = IIf(rs("CanAdd").value = False, flexUnchecked, flexChecked)
                    .cell(flexcpChecked, LngGrdRow, .ColIndex("Edit")) = IIf(rs("CanEdit").value = False, flexUnchecked, flexChecked)
                    .cell(flexcpChecked, LngGrdRow, .ColIndex("Delete")) = IIf(rs("CanDelete").value = False, flexUnchecked, flexChecked)
                    .cell(flexcpChecked, LngGrdRow, .ColIndex("Print")) = IIf(rs("CanPrint").value = False, flexUnchecked, flexChecked)
                    '############################################# Khaled Was Here #############################################################
                    .cell(flexcpChecked, LngGrdRow, .ColIndex("Atta")) = IIf(rs("Attachments").value = False, flexUnchecked, flexChecked)
                    '###########################################################################################################################
                    .cell(flexcpChecked, LngGrdRow, .ColIndex("Search")) = IIf(rs("CanSearch").value = False, flexUnchecked, flexChecked)
                    CellCheck CInt(LngGrdRow)
                    relign (.TextMatrix(LngGrdRow, .ColIndex("Frm_Name"))), LngGrdRow, 0
            
                Else
                    .cell(flexcpChecked, LngGrdRow, .ColIndex("FullAccess")) = IIf(rs("CanAdd").value = False, flexUnchecked, flexChecked)
                End If
            End If

            rs.MoveNext
        Next i

    End With

    rs.Close
    Set rs = Nothing
    
   
End Sub

Private Sub FormateGrid()
    Dim IntScreenType As Integer
    Dim i As Integer

    With Me.FG

        For i = .FixedRows To .rows - 1

            If Not .IsSubtotal(i) Then
                IntScreenType = val(.TextMatrix(i, .ColIndex("ScreenType")))

                If i Mod 2 = 0 Then
                    .cell(flexcpBackColor, i, .ColIndex("AddNew"), i, .ColIndex("FullAccess")) = &HE2E9E9
                    '                If IntScreenType < 5 Then
                    '                    .Cell(flexcpBackColor, I, .ColIndex("AddNew"), I, .ColIndex("FullAccess")) = &HE2E9E9
                    '                Else
                    '                    .Cell(flexcpBackColor, I, .ColIndex("FullAccess")) = &HE2E9E9
                    '                End If
                Else
                    .cell(flexcpBackColor, i, .ColIndex("AddNew"), i, .ColIndex("FullAccess")) = vbWhite
                End If
            End If

        Next i

        .cell(flexcpBackColor, .FixedRows, .ColIndex("TreeCol"), .rows - 1) = vbWhite
        .cell(flexcpBackColor, 4, .ColIndex("NoAccess"), .rows - 1) = vbRed
        
    End With

End Sub
