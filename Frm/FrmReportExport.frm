VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FrmReportExport 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ’œÌ— «· ﬁ«—Ì—"
   ClientHeight    =   6630
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8490
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6630
   ScaleWidth      =   8490
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog Cdg 
      Left            =   180
      Top             =   510
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CheckBox ChkRunFile 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "ﬁ„ » ‘€Ì· «·„·› »⁄œ  „«„ ⁄„·Ì… «· ’œÌ—"
      Height          =   315
      Left            =   5160
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   6000
      Width           =   3225
   End
   Begin VB.Frame FrmPages 
      BackColor       =   &H00E2E9E9&
      Caption         =   "‰ÿ«ﬁ «·’›Õ« («·’›Õ«  «·„—«œ  ’œÌ—Â«)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1125
      Left            =   4620
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   990
      Width           =   3825
      Begin VB.TextBox TxtTo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   270
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   585
         Width           =   555
      End
      Begin VB.TextBox TxtFrom 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   1230
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   585
         Width           =   555
      End
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "’›Õ«  „Õœœ…"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   2220
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   600
         Width           =   1395
      End
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ﬂ· «·’›Õ« "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   2280
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   270
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   2
         Left            =   840
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   615
         Width           =   345
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   345
         Index           =   1
         Left            =   1770
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   585
         Width           =   345
      End
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   345
      Index           =   0
      Left            =   780
      TabIndex        =   11
      Top             =   570
      Width           =   585
      _ExtentX        =   1032
      _ExtentY        =   609
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
      ColorButton     =   14871017
   End
   Begin VB.TextBox TxtFilePath 
      Height          =   345
      Left            =   1380
      TabIndex        =   9
      Top             =   570
      Width           =   5835
   End
   Begin C1SizerLibCtl.C1Tab TabMain 
      Height          =   3765
      Left            =   30
      TabIndex        =   2
      Top             =   2160
      Width           =   8385
      _cx             =   14790
      _cy             =   6641
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
      FrontTabForeColor=   -2147483630
      Caption         =   "PDF|EXCEL|Word|HTML|Email|FAX|Text"
      Align           =   0
      CurrTab         =   4
      FirstTab        =   0
      Style           =   3
      Position        =   0
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
         Height          =   3390
         Index           =   6
         Left            =   9330
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   330
         Width           =   8295
         _cx             =   14631
         _cy             =   5980
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
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   3390
         Index           =   5
         Left            =   9030
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   330
         Width           =   8295
         _cx             =   14631
         _cy             =   5980
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
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   3390
         Index           =   4
         Left            =   45
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   330
         Width           =   8295
         _cx             =   14631
         _cy             =   5980
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
         Begin VB.Frame Fra 
            BackColor       =   &H00E2E9E9&
            Caption         =   "ŒÌ«—«  «Œ—Ï"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   1065
            Index           =   6
            Left            =   3780
            RightToLeft     =   -1  'True
            TabIndex        =   52
            Top             =   2250
            Width           =   4455
            Begin VB.OptionButton OptSend 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "√—”· «·»—Ìœ »’Ì€… ‰’Ì… "
               Height          =   405
               Index           =   1
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   54
               Top             =   570
               Width           =   4155
            End
            Begin VB.OptionButton OptSend 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "√—”· «·»—Ìœ ﬂ’Ì€… HTML"
               Height          =   405
               Index           =   0
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   53
               Top             =   240
               Value           =   -1  'True
               Width           =   4155
            End
         End
         Begin VB.Frame Fra 
            BackColor       =   &H00E2E9E9&
            Caption         =   "»Ì«‰«  Œ«’… »«·„—”· ≈·ÌÂ„"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   3285
            Index           =   5
            Left            =   60
            RightToLeft     =   -1  'True
            TabIndex        =   49
            Top             =   60
            Width           =   3675
            Begin VB.TextBox TxtAddMail 
               Alignment       =   1  'Right Justify
               Height          =   345
               Left            =   480
               RightToLeft     =   -1  'True
               TabIndex        =   51
               Top             =   630
               Width           =   3105
            End
            Begin VSFlex8UCtl.VSFlexGrid FgMails 
               Height          =   2205
               Left            =   60
               TabIndex        =   50
               Top             =   1020
               Width           =   3555
               _cx             =   6271
               _cy             =   3889
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
               Rows            =   50
               Cols            =   2
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmReportExport.frx":0000
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
            Begin ImpulseButton.ISButton Cmd 
               Height          =   345
               Index           =   3
               Left            =   2970
               TabIndex        =   55
               Top             =   240
               Width           =   585
               _ExtentX        =   1032
               _ExtentY        =   609
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
               ColorButton     =   14871017
            End
            Begin ImpulseButton.ISButton Cmd 
               Height          =   345
               Index           =   4
               Left            =   2370
               TabIndex        =   56
               Top             =   240
               Width           =   585
               _ExtentX        =   1032
               _ExtentY        =   609
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
               ColorButton     =   14871017
            End
            Begin ImpulseButton.ISButton Cmd 
               Height          =   345
               Index           =   5
               Left            =   30
               TabIndex        =   57
               Top             =   630
               Width           =   435
               _ExtentX        =   767
               _ExtentY        =   609
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
               ColorButton     =   14871017
            End
            Begin ImpulseButton.ISButton Cmd 
               Height          =   345
               Index           =   7
               Left            =   1770
               TabIndex        =   61
               Top             =   240
               Width           =   585
               _ExtentX        =   1032
               _ExtentY        =   609
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
               ColorButton     =   14871017
            End
         End
         Begin VB.Frame Fra 
            BackColor       =   &H00E2E9E9&
            Caption         =   "»Ì«‰«  Œ«’… »«·„—”·"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   2175
            Index           =   4
            Left            =   3750
            RightToLeft     =   -1  'True
            TabIndex        =   40
            Top             =   60
            Width           =   4455
            Begin VB.CheckBox Chk 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "≈” Œœ«„ «·≈⁄œ«œ«  «·Œ«’… »‘—ﬂ… ‰Ê— ··»—„ÃÌ« "
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   315
               Index           =   5
               Left            =   660
               RightToLeft     =   -1  'True
               TabIndex        =   59
               Top             =   1740
               Width           =   3675
            End
            Begin VB.TextBox TxtMailFrom 
               Alignment       =   1  'Right Justify
               Height          =   345
               Left            =   90
               RightToLeft     =   -1  'True
               TabIndex        =   48
               Top             =   1350
               Width           =   3075
            End
            Begin VB.TextBox TxtUserPass 
               Alignment       =   1  'Right Justify
               Height          =   345
               IMEMode         =   3  'DISABLE
               Left            =   90
               PasswordChar    =   "*"
               RightToLeft     =   -1  'True
               TabIndex        =   46
               Top             =   990
               Width           =   3075
            End
            Begin VB.TextBox TxtUserName 
               Alignment       =   1  'Right Justify
               Height          =   345
               Left            =   90
               RightToLeft     =   -1  'True
               TabIndex        =   44
               Top             =   630
               Width           =   3075
            End
            Begin VB.TextBox TxtMailServer 
               Alignment       =   1  'Right Justify
               Height          =   345
               Left            =   90
               RightToLeft     =   -1  'True
               TabIndex        =   41
               Top             =   270
               Width           =   3075
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·»—Ìœ «·„—”· „‰Â"
               Height          =   345
               Index           =   11
               Left            =   3210
               RightToLeft     =   -1  'True
               TabIndex        =   47
               Top             =   1350
               Width           =   1155
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬂ·„… «·„—Ê—"
               Height          =   345
               Index           =   10
               Left            =   3210
               RightToLeft     =   -1  'True
               TabIndex        =   45
               Top             =   990
               Width           =   1155
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·„” Œœ„"
               Height          =   345
               Index           =   9
               Left            =   3210
               RightToLeft     =   -1  'True
               TabIndex        =   43
               Top             =   630
               Width           =   1155
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·„Ì· ”Ì—›—"
               Height          =   345
               Index           =   8
               Left            =   3210
               RightToLeft     =   -1  'True
               TabIndex        =   42
               Top             =   270
               Width           =   1155
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   3390
         Index           =   3
         Left            =   -8940
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   330
         Width           =   8295
         _cx             =   14631
         _cy             =   5980
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
         Begin VB.CheckBox Chk 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈Ã⁄· «·’›Õ«  «·‰« Ã… ⁄»«—… ⁄‰ ’›Õ«  „‰›’·…"
            Height          =   345
            Index           =   4
            Left            =   2280
            RightToLeft     =   -1  'True
            TabIndex        =   39
            Top             =   600
            Width           =   5865
         End
         Begin VB.CheckBox Chk 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "÷⁄ „ ’›Õ ›Ï «”›· «·’›Õ«  ·Ì”«⁄œ «·„” Œœ„ ⁄·Ï «· ‰ﬁ· »”—⁄… »Ì‰ «·’›Õ« "
            Height          =   345
            Index           =   3
            Left            =   2280
            RightToLeft     =   -1  'True
            TabIndex        =   38
            Top             =   300
            Value           =   1  'Checked
            Width           =   5865
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   3390
         Index           =   2
         Left            =   -9240
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   330
         Width           =   8295
         _cx             =   14631
         _cy             =   5980
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
         Begin VB.Frame Fra 
            BackColor       =   &H00E2E9E9&
            Caption         =   "ŒÌ«—«  «· ’œÌ— ≈·Ï «·Ê—Êœ"
            Height          =   3315
            Index           =   3
            Left            =   60
            RightToLeft     =   -1  'True
            TabIndex        =   32
            Top             =   30
            Width           =   8175
            Begin VB.OptionButton Opt 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " ’œÌ— ≈·Ï „·› *.rft"
               Height          =   375
               Index           =   5
               Left            =   6270
               RightToLeft     =   -1  'True
               TabIndex        =   34
               Top             =   750
               Width           =   1755
            End
            Begin VB.OptionButton Opt 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " ’œÌ— ≈·Ï „·› .Doc"
               Height          =   375
               Index           =   4
               Left            =   6270
               RightToLeft     =   -1  'True
               TabIndex        =   33
               Top             =   360
               Value           =   -1  'True
               Width           =   1755
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "√Ê›Ì” 97 Ê„«ﬁ»·Â« „‰ ≈’œ«—« "
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   345
               Index           =   5
               Left            =   1440
               RightToLeft     =   -1  'True
               TabIndex        =   36
               Top             =   810
               Width           =   4725
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "√Ê›Ì” 2000 ,√ﬂ” »Ï,2003 ,2007 ,,,,,Ê„«»⁄œÂ« „‰ ≈’œ«—« "
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   345
               Index           =   4
               Left            =   1440
               RightToLeft     =   -1  'True
               TabIndex        =   35
               Top             =   420
               Width           =   4725
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   3390
         Index           =   1
         Left            =   -9540
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   330
         Width           =   8295
         _cx             =   14631
         _cy             =   5980
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
         Begin VB.Frame Fra 
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·Ã“¡ «·„—«œ  ’œÌ—Â „‰ «· ﬁ—Ì—"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   1155
            Index           =   0
            Left            =   60
            RightToLeft     =   -1  'True
            TabIndex        =   28
            Top             =   90
            Width           =   3465
         End
         Begin VB.Frame Fra 
            BackColor       =   &H00E2E9E9&
            Caption         =   "ŒÌ«—«  ≈Œ—Ï „ ﬁœ„…"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   2025
            Index           =   2
            Left            =   90
            RightToLeft     =   -1  'True
            TabIndex        =   27
            Top             =   1290
            Width           =   8145
            Begin VB.CheckBox Chk 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "≈⁄—÷ ÕœÊœ «·Œ·«Ì« ›Ï ’›Õ… «·√ﬂ”·"
               Height          =   255
               Index           =   2
               Left            =   3720
               RightToLeft     =   -1  'True
               TabIndex        =   31
               Top             =   840
               Width           =   4335
            End
            Begin VB.CheckBox Chk 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ÕÊ· «· Ê«—ÌŒ ›Ï «· ﬁ—Ì— ≈·Ï ‰’Ê’ ›Ï «·≈ﬂ”·"
               Height          =   255
               Index           =   1
               Left            =   3720
               RightToLeft     =   -1  'True
               TabIndex        =   30
               Top             =   540
               Width           =   4335
            End
            Begin VB.CheckBox Chk 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "≈⁄„· Õœ ›«’· »Ì‰ ﬂ· ’›Õ… Ê«Œ—Ï ›Ï «· ﬁ—Ì—"
               Height          =   255
               Index           =   0
               Left            =   3720
               RightToLeft     =   -1  'True
               TabIndex        =   29
               Top             =   270
               Width           =   4335
            End
         End
         Begin VB.Frame Fra 
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄—÷ «·√⁄„œ… ›Ï ’›Õ… «·√ﬂ”·"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   1155
            Index           =   1
            Left            =   3540
            RightToLeft     =   -1  'True
            TabIndex        =   22
            Top             =   90
            Width           =   4695
            Begin VB.TextBox TxtColWidth 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               Height          =   345
               Left            =   1110
               RightToLeft     =   -1  'True
               TabIndex        =   26
               Top             =   600
               Width           =   915
            End
            Begin VB.OptionButton Opt 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄—÷ À«»  ·ﬂ· «·√⁄„œ…(In Points)"
               Height          =   285
               Index           =   3
               Left            =   1800
               RightToLeft     =   -1  'True
               TabIndex        =   24
               Top             =   630
               Width           =   2805
            End
            Begin VB.OptionButton Opt 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "»‰«¡ ⁄·Ï «·»Ì«‰«  «·„ÊÃÊœ… ›Ï ﬂ· ⁄„Êœ"
               Height          =   285
               Index           =   2
               Left            =   1530
               RightToLeft     =   -1  'True
               TabIndex        =   23
               Top             =   300
               Value           =   -1  'True
               Width           =   3075
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "(Ì›÷· Â–« «·ŒÌ«—)"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   285
               Index           =   3
               Left            =   240
               RightToLeft     =   -1  'True
               TabIndex        =   25
               Top             =   300
               Width           =   1215
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   3390
         Index           =   0
         Left            =   -9840
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   330
         Width           =   8295
         _cx             =   14631
         _cy             =   5980
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
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "·« ÊÃœ Õ Ï «·¬‰ «Ï ŒÌ«—«  Œ«’… „⁄ «· ’œÌ— ≈·Ï «·√ﬂ—Ê»«  —ÌœÌ—"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   14.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   945
            Index           =   6
            Left            =   930
            RightToLeft     =   -1  'True
            TabIndex        =   37
            Top             =   630
            Width           =   6675
         End
      End
   End
   Begin VB.ComboBox CboExportType 
      Height          =   315
      Left            =   3930
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   210
      Width           =   3285
   End
   Begin ImpulseButton.ISButton Cmd 
      Cancel          =   -1  'True
      Height          =   405
      Index           =   1
      Left            =   60
      TabIndex        =   12
      Top             =   6030
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   714
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
      Height          =   405
      Index           =   2
      Left            =   1020
      TabIndex        =   21
      Top             =   6030
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   714
      Caption         =   " ’œÌ—"
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
      Height          =   405
      Index           =   6
      Left            =   2160
      TabIndex        =   60
      Top             =   6030
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   714
      Caption         =   "„”«⁄œ…"
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
      Caption         =   "„”«— Õ›Ÿ «·„·›"
      Height          =   315
      Index           =   7
      Left            =   7260
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   600
      Width           =   1155
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " ’œÌ— «· ﬁ«—Ì—"
      Height          =   315
      Index           =   0
      Left            =   7260
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   210
      Width           =   1155
   End
End
Attribute VB_Name = "FrmReportExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public xReport As CRAXDRT.Report
Dim TTP As clstooltipdemand

Private Sub CboExportType_Change()
    Dim i As Integer

    For i = 0 To Me.TabMain.NumTabs - 1
        Me.TabMain.TabVisible(i) = False
    Next i

    Me.lbl(7).Enabled = True
    Me.TxtFilePath.Enabled = False
    Me.Cmd(0).Enabled = True

    If Me.CboExportType.ListIndex <= -1 Then
        Me.TabMain.Visible = False
    ElseIf Me.CboExportType.ListIndex = 0 Then
        Me.TabMain.TabCaption(0) = "ŒÌ«—«  «· ’œÌ— ≈·Ï „·› √ﬂ—Ê»«  —ÌœÌ—"
        Me.TabMain.TabVisible(0) = True
        Me.TabMain.CurrTab = 0
    ElseIf Me.CboExportType.ListIndex = 1 Then
        Me.TabMain.TabCaption(1) = "ŒÌ«—«  «· ’œÌ— ≈·Ï „·› „Ìﬂ—Ê”Ê›  ≈ﬂ”·"
        Me.TabMain.TabVisible(1) = True
        Me.TabMain.CurrTab = 1
    ElseIf Me.CboExportType.ListIndex = 2 Then
        Me.TabMain.TabCaption(2) = "ŒÌ«—«  «· ’œÌ— ≈·Ï „·› „Ìﬂ—Ê”Ê›  Ê—œ"
        Me.TabMain.TabVisible(2) = True
        Me.TabMain.CurrTab = 2
    ElseIf Me.CboExportType.ListIndex = 3 Then
        Me.TabMain.TabCaption(3) = "ŒÌ«—«  «· ’œÌ— ≈·Ï „·› HTML"
        Me.TabMain.TabVisible(3) = True
        Me.TabMain.CurrTab = 3
    ElseIf Me.CboExportType.ListIndex = 4 Then
        Me.TabMain.TabCaption(4) = "ŒÌ«—«  «·≈—”«· ⁄»— «·»—Ìœ «·≈·ﬂ —Ê‰Ï"
        Me.TabMain.TabVisible(4) = True
        Me.TabMain.CurrTab = 4
    
        Me.lbl(7).Enabled = False
        Me.TxtFilePath.Enabled = False
        Me.Cmd(0).Enabled = False

    ElseIf Me.CboExportType.ListIndex = 5 Then
        Me.TabMain.TabCaption(5) = "ŒÌ«—«  «·≈—”«· ⁄»— «·›«ﬂ”"
        Me.TabMain.TabVisible(5) = True
        Me.TabMain.CurrTab = 5
    ElseIf Me.CboExportType.ListIndex = 6 Then
        Me.TabMain.TabCaption(6) = "ŒÌ«—«  «· ’œÌ— ≈·Ï „·› ‰’Ï"
        Me.TabMain.TabVisible(6) = True
        Me.TabMain.CurrTab = 6
    End If

End Sub

Private Sub CboExportType_Click()
    CboExportType_Change
End Sub

Private Sub Chk_Click(Index As Integer)

    Select Case Index

        Case 5

            If Me.Chk(5).value = vbChecked Then
                Me.TxtMailServer.Enabled = False
                Me.txtUserName.Enabled = False
                Me.TxtUserPass.Enabled = False
                Me.lbl(8).Enabled = False
                Me.lbl(9).Enabled = False
                Me.lbl(10).Enabled = False
                Me.TxtMailServer.text = "mail.¡¡¡.com"
                Me.txtUserName.text = "nourhost"
                Me.TxtUserPass.text = "nour1234nour"
            ElseIf Me.Chk(5).value = vbUnchecked Then
                Me.TxtMailServer.Enabled = True
                Me.txtUserName.Enabled = True
                Me.TxtUserPass.Enabled = True
                Me.lbl(8).Enabled = True
                Me.lbl(9).Enabled = True
                Me.lbl(10).Enabled = True
                Me.TxtMailServer.text = ""
                Me.txtUserName.text = ""
                Me.TxtUserPass.text = ""
            End If
        
    End Select

End Sub

Private Sub Cmd_Click(Index As Integer)

    Select Case Index

        Case 0
            ShowFileSave

        Case 1
            Unload Me

        Case 2
            StartExport

        Case 5
            AddNewEmail
    End Select

End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim GrdBack As ClsBackGroundPic

    CenterForm Me

    FormPostion Me, GetPostion

    For i = Me.Cmd.LBound To Me.Cmd.UBound
        Me.Cmd(i).ButtonStyle = impActive
        Me.Cmd(i).ButtonPositionImage = impRightOfText
        Me.Cmd(i).BackColor = Me.BackColor
    Next i

    Me.Icon = mdifrmmain.ImgLstMenuIcons.ListImages("Export").Picture

    Set Me.Cmd(0).ButtonImage = mdifrmmain.ImgLstMenuIcons.ListImages("BrowseFile").Picture
    Set Me.Cmd(1).ButtonImage = mdifrmmain.ImgLstMenuIcons.ListImages("Exit").Picture
    Set Me.Cmd(2).ButtonImage = mdifrmmain.ImgLstMenuIcons.ListImages("Export").Picture
    Set Me.Cmd(3).ButtonImage = mdifrmmain.ImgLstMenuIcons.ListImages("OpenFolder").Picture
    Set Me.Cmd(4).ButtonImage = mdifrmmain.ImgLstMenuIcons.ListImages("Save").Picture
    Set Me.Cmd(5).ButtonImage = mdifrmmain.ImgLstMenuIcons.ListImages("Plus").Picture
    Set Me.Cmd(6).ButtonImage = mdifrmmain.ImgLstMenuIcons.ListImages("Help2").Picture
    Set Me.Cmd(7).ButtonImage = mdifrmmain.ImgLstMenuIcons.ListImages("Mins").Picture

    '
    Set Me.TabMain.TabPicture(0) = mdifrmmain.ImgLstMenuIcons.ListImages("ExportPDF").Picture
    Set Me.TabMain.TabPicture(1) = mdifrmmain.ImgLstMenuIcons.ListImages("ExportExcel").Picture
    Set Me.TabMain.TabPicture(2) = mdifrmmain.ImgLstMenuIcons.ListImages("ExportWord").Picture
    Set Me.TabMain.TabPicture(3) = mdifrmmain.ImgLstMenuIcons.ListImages("ExportHTML").Picture
    Set Me.TabMain.TabPicture(4) = mdifrmmain.ImgLstMenuIcons.ListImages("ExportMail").Picture
    Set Me.TabMain.TabPicture(5) = mdifrmmain.ImgLstMenuIcons.ListImages("Report").Picture

    With Me.CboExportType
        .Clear
        .AddItem "„·› √ﬂ—Ê»«  —ÌœÌ— PDF"
        .AddItem "„·› ≈ﬂ”·"
        .AddItem "„·› Ê—œ"
        .AddItem "„·› HTML"
        .AddItem "≈—”«· «· ﬁ—Ì— ⁄»— «·»—Ìœ «·≈·ﬂ —Ê‰Ï"
        .AddItem "≈—”«· «· ﬁ—Ì— ⁄»— «·›«ﬂ”"
        .AddItem "„·› ‰’"
    End With

    '------------------------------------------------------------------------------
    Me.TxtAddMail.RightToLeft = False
    Me.TxtAddMail.Alignment = vbLeftJustify

    With Me.FgMails
        .ExplorerBar = flexExSortShowAndMove
        .ExtendLastCol = True
        Set GrdBack = New ClsBackGroundPic
        Set .WallPaper = GrdBack.Picture
        .Rows = .FixedRows
        .RowHeightMin = 300
        .AutoSize 0, .Cols - 1, False
    End With

    '------------------------------------------------------------------------------
    Me.Opt(0).value = True
    Opt_Click 0
    CboExportType.ListIndex = 0

End Sub

Private Sub Form_Unload(Cancel As Integer)

    FormPostion Me, SavePostion
End Sub

Private Sub StartExport()
    Dim x As CRAXDRT.ExportOptions
    Dim Y  As CRAXDRT.ExportOptions
    Dim Msg As String
    Dim StrFileName As String
    On Error GoTo ErrTrap

    If Trim$(Me.TxtFilePath.text) = "" Then
        Msg = "ÌÃ»  ÕœÌœ „”«— Õ›Ÿ «·„·› «·„—«œ  ’œÌ—Â...!!!"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    Else
        StrFileName = Trim$(Me.TxtFilePath.text)
    End If

    If Me.Opt(1).value = True Then
        If val(Me.txtfrom.text) = 0 Then
            Me.txtfrom.text = 1
        End If

        If val(Me.txtto.text) > 0 Then
            If val(Me.txtto.text) < val(Me.txtfrom.text) Then
                Msg = "ÌÃ»  ÕœÌœ —ﬁ„ ’›Õ… «·»œ«Ì…"
            End If
        End If
    End If

    If Not xReport Is Nothing Then
        xReport.ReportAuthor = "‘—ﬂ… ” «—“  ﬂ‰Ê·ÊÃÌ ··»—„ÃÌ« "
        Msg = "‘—ﬂ…  ” «—“  ﬂ‰Ê·ÊÃÌ ··»—„ÃÌ« "
        Msg = Msg & Chr(13) & " "
        Msg = Msg & Chr(13) & "BYTE"
        Msg = Msg & Chr(13) & "info@.com"
        xReport.ReportComments = Msg

        If Me.CboExportType.ListIndex = 0 Then
        
            If Dir(StrFileName) <> "" Then
                Kill StrFileName
            End If

            Set x = xReport.ExportOptions
            x.DestinationType = crEDTDiskFile
            x.FormatType = 31
            x.DiskFileName = StrFileName

            If Me.Opt(0).value = True Then
                x.PDFExportAllPages = True
            Else
                x.PDFExportAllPages = False
                x.PDFFirstPageNumber = val(Me.txtfrom.text)
                x.PDFLastPageNumber = val(Me.txtto.text)
            End If

            xReport.Export False

            If Me.ChkRunFile.value = vbChecked Then
                OpenFile StrFileName
            End If

        ElseIf Me.CboExportType.ListIndex = 1 Then
            Set x = xReport.ExportOptions
            x.DestinationType = crEDTDiskFile
            x.ExcelMaintainColumnAlignment = True
            x.ExcelMaintainRelativeObjectPosition = True
        
            If Me.Chk(0).value = vbChecked Then
                x.ExcelPageBreaks = True
            Else
                x.ExcelPageBreaks = False
            End If

            If Me.Chk(1).value = vbChecked Then
                x.ExcelConvertDateToString = True
            Else
                x.ExcelConvertDateToString = False
            End If

            If Me.Chk(2).value = vbChecked Then
                x.ExcelShowGridlines = True
            Else
                x.ExcelShowGridlines = False
            End If

            x.FormatType = crEFTExcel97
            x.DiskFileName = StrFileName

            If Me.Opt(0).value = True Then
                x.ExcelExportAllPages = True
            Else
                x.ExcelExportAllPages = False
                x.ExcelFirstPageNumber = val(Me.txtfrom.text)
                x.ExcelLastPageNumber = val(Me.txtto.text)
            End If

            If Dir(StrFileName) <> "" Then
                Kill StrFileName
            End If

            xReport.Export False

            If Me.ChkRunFile.value = vbChecked Then
                OpenFile StrFileName
            End If

        ElseIf Me.CboExportType.ListIndex = 2 Then
            Set x = xReport.ExportOptions
            x.DestinationType = crEDTDiskFile

            If Me.Opt(4).value = True Then
                x.FormatType = crEFTWordForWindows

                If Me.Opt(0).value = True Then
                    x.WORDWExportAllPages = True
                Else
                    x.WORDWExportAllPages = False
                    x.WORDWFirstPageNumber = val(Me.txtfrom.text)
                    x.WORDWLastPageNumber = val(Me.txtto.text)
                End If

            ElseIf Me.Opt(5).value = True Then
                x.FormatType = crEFTExactRichText

                If Me.Opt(0).value = True Then
                    x.RTFExportAllPages = True
                Else
                    x.RTFExportAllPages = False
                    x.RTFFirstPageNumber = val(Me.txtfrom.text)
                    x.RTFLastPageNumber = val(Me.txtto.text)
                End If
            End If

            x.DiskFileName = StrFileName

            If Dir(StrFileName) <> "" Then
                Kill StrFileName
            End If

            xReport.Export False

            If Me.ChkRunFile.value = vbChecked Then
                OpenFile StrFileName
            End If

        ElseIf Me.CboExportType.ListIndex = 3 Then
            Set x = xReport.ExportOptions
            x.DestinationType = crEDTDiskFile

            If Dir(StrFileName) <> "" Then
                Kill StrFileName
            End If

            x.FormatType = crEFTHTML40
            x.DiskFileName = StrFileName

            If Me.Chk(4).value = vbChecked Then
                x.HTMLEnableSeparatedPages = True
            Else
                x.HTMLEnableSeparatedPages = False
            End If

            If Me.Chk(3).value = vbChecked Then
                x.HTMLHasPageNavigator = True
            Else
                x.HTMLHasPageNavigator = False
            End If

            x.ApplicationFileName = "bisegypt"
            x.HTMLFileName = StrFileName
            xReport.Export False

            If Me.ChkRunFile.value = vbChecked Then
                OpenFile StrFileName
            End If

        ElseIf Me.CboExportType.ListIndex = 4 Then
            Set x = xReport.ExportOptions
            x.FormatType = crEFTPortableDocFormat
            x.DestinationType = crEDTEMailMAPI
        
            'crxReport.ExportOptions.PDFExportAllPages = True
            x.MailSubject = "Testing Report Export via Email"
            x.MailMessage = "Here is the latest Micro Report"
            x.MailToList = "info@.com"
            xReport.Export False

        End If
    End If

    Exit Sub
ErrTrap:
    Msg = "⁄›Ê« ÕœÀ Œÿ« √À‰«¡ „Õ«Ê·…  ’œÌ— «· ﬁ—Ì—...!!!"
    Msg = Msg & Chr(13) & Err.description
    Msg = Msg & Chr(13) & Err.Number
    Msg = Msg & Chr(13) & Err.Source
    Msg = Msg & Chr(13) & Err.LastDllError
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
End Sub

Private Sub Opt_Click(Index As Integer)
    Me.lbl(1).Enabled = Me.Opt(1).value
    Me.lbl(2).Enabled = Me.Opt(1).value
    Me.txtfrom.Enabled = Me.Opt(1).value
    Me.txtto.Enabled = Me.Opt(1).value
    Me.txtfrom.text = "1"
    '----------------------------------
    Me.TxtColWidth.Enabled = Me.Opt(3).value
    '----------------------------------
End Sub

Private Sub TxtColWidth_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtColWidth.text, 0)
End Sub

Private Sub TxtFrom_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.txtfrom.text, 1)
End Sub

Private Sub TxtTo_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.txtto.text, 1)
End Sub

Private Sub ShowFileSave()
    Dim Msg As String
    Dim StrFileName As String

    If Me.CboExportType.ListIndex = -1 Then
        Msg = "ÌÃ»  ÕœÌœ ‰Ê⁄ «·„·› «·„—«œ «· ’œÌ— ·Â...!!!"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Me.CboExportType.SetFocus
        Exit Sub
    End If

    Me.Cdg.Flags = cdlOFNOverwritePrompt

    If Not Me.xReport Is Nothing Then
        StrFileName = xReport.reporttitle
    End If

    If Me.CboExportType.ListIndex = 0 Then
        If SystemOptions.UserInterface = EnglishInterface Then
            Me.Cdg.Filter = "Acrobar Reader Files|*.pdf"
        
        ElseIf SystemOptions.UserInterface = ArabicInterface Then
            Me.Cdg.Filter = "„·›«  √ﬂ—Ê»«  —Ìœ—|*.pdf"
        End If

        StrFileName = StrFileName & ".pdf"
    ElseIf Me.CboExportType.ListIndex = 1 Then

        If SystemOptions.UserInterface = EnglishInterface Then
            Me.Cdg.Filter = "Microsoft Excel|*.XLS"
        ElseIf SystemOptions.UserInterface = ArabicInterface Then
            Me.Cdg.Filter = "„·›«  √ﬂ”·|*.XLS"
        End If

        StrFileName = StrFileName & ".xls"
    ElseIf Me.CboExportType.ListIndex = 2 Then

        If SystemOptions.UserInterface = EnglishInterface Then
            Me.Cdg.Filter = "Microsoft Word|*.doc"
        ElseIf SystemOptions.UserInterface = ArabicInterface Then
            Me.Cdg.Filter = "„·›«  „Ìﬂ—Ê”Êﬁ  Ê—Êœ|*.doc"
        End If

        StrFileName = StrFileName & ".doc"
    ElseIf Me.CboExportType.ListIndex = 3 Then

        If SystemOptions.UserInterface = EnglishInterface Then
            Me.Cdg.Filter = "HTML Files|*.HTML"
        ElseIf SystemOptions.UserInterface = ArabicInterface Then
            Me.Cdg.Filter = "„·›«  ÊÌ»|*.HTML"
        End If

        StrFileName = StrFileName & ".html"
    End If

    StrFileName = Replace(StrFileName, ":", "_", 1, -1, vbTextCompare)
    Me.Cdg.FileName = StrFileName
    Me.Cdg.ShowSave
    Me.TxtFilePath.text = Me.Cdg.FileName
End Sub

Private Sub AddNewEmail()
    Dim StrMSG As String
    Dim StrNewEmail  As String
    Dim LngFindRow As Long

    If Trim$(Me.TxtAddMail.text) = "" Then
        Set TTP = New clstooltipdemand
        Set TTP.m_From = Me
        TTP.Style = TTBalloon
        TTP.Icon = TTIconError
        TTP.Centered = True
        TTP.RightToLeft = True
        TTP.CreateToolTip TxtAddMail.hWnd
        TTP.DelayTime = 250
        TTP.VisibleTime = 5000
        StrMSG = "ÌÃ» ﬂ «»… «·√Ì„Ì· ...!!!"
        TTP.Title = StrMSG
        StrMSG = "ÌÃ» «‰  ﬁÊ„ »ﬂ «»… «·√Ì„Ì· «·–Ï  —Ìœ ≈÷«› Â"
        TTP.TipText = StrMSG
        TTP.PopupOnDemand = True
        TTP.Show (TxtAddMail.Width / Screen.TwipsPerPixelY), (TxtAddMail.Height / Screen.TwipsPerPixelX - 1)     '//In Pixel only
        Me.TxtAddMail.SetFocus
        Exit Sub
    Else

        If Not TTP Is Nothing Then
            TTP.Destroy
        End If
    End If

    StrNewEmail = Trim$(Me.TxtAddMail.text)

    With Me.FgMails
        LngFindRow = .FindRow(StrNewEmail, .FixedRows, .ColIndex("Email"), False, True)

        If LngFindRow <> -1 Then
            StrMSG = "Â–« «·√Ì„Ì· „ÊÃÊœ ›Ï «·ﬁ«∆„…...!!!"
            MsgBox StrMSG, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            Exit Sub
        Else
            .AddItem "" & vbTab & StrNewEmail
            Me.TxtAddMail.text = ""
            Me.TxtAddMail.SetFocus
        End If

    End With

End Sub
