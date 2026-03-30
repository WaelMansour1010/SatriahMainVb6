VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form RSRentAlarm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " ‰»ÌÂ«  «·«ÌÃ«—«  «·„” ÕﬁÂ Œ·«· › —…"
   ClientHeight    =   9465
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   16380
   Icon            =   "ReRentAlarm.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   9465
   ScaleWidth      =   16380
   ShowInTaskbar   =   0   'False
   Begin C1SizerLibCtl.C1Tab TabMain 
      Height          =   9495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   17835
      _cx             =   31459
      _cy             =   16748
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
      Caption         =   "«·ÿ·»«  «·œ«Œ·Ì…|«·«ÌÃ«—«  «·„” Õﬁ…"
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
      Begin C1SizerLibCtl.C1Elastic ELe 
         Height          =   9120
         Index           =   1
         Left            =   45
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   45
         Width           =   17745
         _cx             =   31300
         _cy             =   16087
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
         Begin VB.TextBox TXTTransactionID4 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   630
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   62
            Top             =   450
            Visible         =   0   'False
            Width           =   1260
         End
         Begin VB.TextBox TxtNoteSerial14 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   14505
            RightToLeft     =   -1  'True
            TabIndex        =   61
            Top             =   1995
            Visible         =   0   'False
            Width           =   1830
         End
         Begin VB.CommandButton cmdCreateProduction 
            Caption         =   "«‰‘«¡ «„— «‰ «Ã"
            Enabled         =   0   'False
            Height          =   330
            Left            =   9855
            RightToLeft     =   -1  'True
            TabIndex        =   54
            Top             =   1635
            Width           =   6480
         End
         Begin VB.Frame Frame8 
            Caption         =   "Õœœ «·› —…"
            Height          =   1470
            Index           =   1
            Left            =   9855
            RightToLeft     =   -1  'True
            TabIndex        =   45
            Top             =   105
            Width           =   6510
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               Caption         =   " ÕœÌœ «·ﬂ·"
               Height          =   195
               Left            =   4440
               RightToLeft     =   -1  'True
               TabIndex        =   46
               Top             =   840
               Visible         =   0   'False
               Width           =   1695
            End
            Begin MSComCtl2.DTPicker txtFromDate 
               Height          =   330
               Left            =   3135
               TabIndex        =   47
               Top             =   240
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   582
               _Version        =   393216
               Format          =   95944705
               CurrentDate     =   41640
            End
            Begin Dynamic_Byte.NourHijriCal NourHijriCal1 
               Height          =   255
               Left            =   3120
               TabIndex        =   48
               Top             =   600
               Visible         =   0   'False
               Width           =   1815
               _extentx        =   3201
               _extenty        =   450
            End
            Begin ImpulseButton.ISButton Cmd 
               Height          =   510
               Index           =   1
               Left            =   120
               TabIndex        =   49
               Top             =   240
               Width           =   720
               _ExtentX        =   1270
               _ExtentY        =   900
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   "≈÷«›…"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ButtonImage     =   "ReRentAlarm.frx":058A
               DrawFocusRectangle=   0   'False
            End
            Begin Dynamic_Byte.NourHijriCal NourHijriCal2 
               Height          =   255
               Left            =   840
               TabIndex        =   50
               Top             =   600
               Visible         =   0   'False
               Width           =   1755
               _extentx        =   3201
               _extenty        =   450
            End
            Begin MSComCtl2.DTPicker txtToDate 
               Height          =   330
               Left            =   840
               TabIndex        =   51
               Top             =   240
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   582
               _Version        =   393216
               Format          =   95944705
               CurrentDate     =   41640
            End
            Begin MSDataListLib.DataCombo Dcbranch 
               Height          =   315
               Left            =   780
               TabIndex        =   63
               Top             =   1020
               Width           =   1875
               _ExtentX        =   3307
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·›—⁄"
               ForeColor       =   &H00000000&
               Height          =   285
               Index           =   36
               Left            =   2805
               RightToLeft     =   -1  'True
               TabIndex        =   64
               Top             =   1020
               Width           =   660
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«·› —… „‰"
               Height          =   315
               Index           =   2
               Left            =   4980
               RightToLeft     =   -1  'True
               TabIndex        =   53
               Top             =   240
               Width           =   1035
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«·Ï"
               Height          =   435
               Index           =   1
               Left            =   2460
               RightToLeft     =   -1  'True
               TabIndex        =   52
               Top             =   240
               Width           =   570
            End
         End
         Begin VB.Timer Timer2 
            Left            =   0
            Top             =   0
         End
         Begin VSFlex8UCtl.VSFlexGrid Fg 
            Height          =   6630
            Left            =   150
            TabIndex        =   2
            Top             =   2475
            Width           =   16155
            _cx             =   28496
            _cy             =   11695
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
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"ReRentAlarm.frx":0924
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
         Begin MSDataListLib.DataCombo DCboStoreName2 
            Height          =   315
            Left            =   2640
            TabIndex        =   55
            Top             =   495
            Width           =   4530
            _ExtentX        =   7990
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCboStoreName 
            Height          =   315
            Left            =   2640
            TabIndex        =   56
            Top             =   795
            Width           =   4530
            _ExtentX        =   7990
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DBCboClientName 
            Height          =   315
            Left            =   2640
            TabIndex        =   57
            Top             =   180
            Visible         =   0   'False
            Width           =   4500
            _ExtentX        =   7938
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "6"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "„Œ“‰  «·«‰ «Ã «· «„"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   34
            Left            =   7440
            RightToLeft     =   -1  'True
            TabIndex        =   60
            Top             =   795
            Width           =   2325
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "„Œ“‰ «·„Ê«œ «·Œ«„"
            ForeColor       =   &H00000000&
            Height          =   210
            Index           =   33
            Left            =   7440
            RightToLeft     =   -1  'True
            TabIndex        =   59
            Top             =   495
            Width           =   2325
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·⁄„Ì·"
            ForeColor       =   &H00000000&
            Height          =   225
            Index           =   42
            Left            =   7440
            RightToLeft     =   -1  'True
            TabIndex        =   58
            Top             =   180
            Visible         =   0   'False
            Width           =   2055
         End
      End
      Begin C1SizerLibCtl.C1Elastic ELe 
         Height          =   9120
         Index           =   0
         Left            =   18480
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   45
         Width           =   17745
         _cx             =   31300
         _cy             =   16087
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
         Begin VB.CommandButton Command1 
            Caption         =   "«—”«· "
            Height          =   420
            Left            =   2280
            RightToLeft     =   -1  'True
            TabIndex        =   38
            Top             =   6165
            Width           =   1470
         End
         Begin VB.Frame Frame9 
            Caption         =   "«Ã„«·Ì« "
            Height          =   765
            Left            =   4800
            RightToLeft     =   -1  'True
            TabIndex        =   25
            Top             =   6075
            Visible         =   0   'False
            Width           =   13035
            Begin VB.TextBox TxtTotalContract 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFC0&
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
               Left            =   10320
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   31
               Top             =   360
               Width           =   1065
            End
            Begin VB.TextBox TxtInsuranceValue 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFC0&
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
               Left            =   6240
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   30
               Top             =   360
               Width           =   1065
            End
            Begin VB.TextBox TxtWater 
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
               Left            =   4080
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   29
               Top             =   360
               Width           =   1065
            End
            Begin VB.TextBox TxtElectricity 
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
               Left            =   2160
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   28
               Top             =   360
               Width           =   945
            End
            Begin VB.TextBox TxtCommiValue 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFC0&
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
               Left            =   8280
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   27
               Top             =   360
               Width           =   1065
            End
            Begin VB.TextBox TxtPhone 
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
               Left            =   120
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   26
               Top             =   360
               Width           =   945
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "ﬁÌ„… «·«ÌÃ«—"
               Height          =   195
               Index           =   6
               Left            =   11505
               RightToLeft     =   -1  'True
               TabIndex        =   37
               Top             =   480
               Width           =   870
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   " √„Ì‰"
               Height          =   195
               Index           =   19
               Left            =   7560
               RightToLeft     =   -1  'True
               TabIndex        =   36
               Top             =   360
               Width           =   510
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "„Ì«Â"
               Height          =   195
               Index           =   20
               Left            =   5385
               RightToLeft     =   -1  'True
               TabIndex        =   35
               Top             =   480
               Width           =   750
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "ﬂÂ—»«¡"
               Height          =   195
               Index           =   21
               Left            =   2985
               RightToLeft     =   -1  'True
               TabIndex        =   34
               Top             =   480
               Width           =   750
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "”⁄Ì/—”Ê„"
               Height          =   405
               Index           =   25
               Left            =   9360
               RightToLeft     =   -1  'True
               TabIndex        =   33
               Top             =   360
               Width           =   810
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "Â« › Ê«‰ —‰ "
               Height          =   195
               Index           =   27
               Left            =   1035
               RightToLeft     =   -1  'True
               TabIndex        =   32
               Top             =   480
               Width           =   990
            End
         End
         Begin VB.Frame Frame8 
            Caption         =   "Õœœ «·› —…"
            Height          =   990
            Index           =   0
            Left            =   11235
            RightToLeft     =   -1  'True
            TabIndex        =   16
            Top             =   480
            Width           =   6510
            Begin VB.CheckBox Check17 
               Alignment       =   1  'Right Justify
               Caption         =   " ÕœÌœ «·ﬂ·"
               Height          =   195
               Left            =   4440
               RightToLeft     =   -1  'True
               TabIndex        =   17
               Top             =   840
               Width           =   1695
            End
            Begin MSComCtl2.DTPicker Fromdate 
               Height          =   330
               Left            =   3135
               TabIndex        =   18
               Top             =   240
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   582
               _Version        =   393216
               Format          =   95944705
               CurrentDate     =   41640
            End
            Begin Dynamic_Byte.NourHijriCal Fromdateh 
               Height          =   255
               Left            =   3120
               TabIndex        =   19
               Top             =   600
               Width           =   1815
               _extentx        =   3201
               _extenty        =   450
            End
            Begin ImpulseButton.ISButton Cmd 
               Height          =   510
               Index           =   9
               Left            =   120
               TabIndex        =   20
               Top             =   240
               Width           =   720
               _ExtentX        =   1270
               _ExtentY        =   900
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   "≈÷«›…"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ButtonImage     =   "ReRentAlarm.frx":0A99
               DrawFocusRectangle=   0   'False
            End
            Begin Dynamic_Byte.NourHijriCal todateH 
               Height          =   255
               Left            =   840
               TabIndex        =   21
               Top             =   600
               Width           =   1755
               _extentx        =   3201
               _extenty        =   450
            End
            Begin MSComCtl2.DTPicker toDate 
               Height          =   330
               Left            =   840
               TabIndex        =   22
               Top             =   240
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   582
               _Version        =   393216
               Format          =   95944705
               CurrentDate     =   41640
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "≈«·Ï"
               Height          =   435
               Index           =   14
               Left            =   2460
               RightToLeft     =   -1  'True
               TabIndex        =   24
               Top             =   240
               Width           =   540
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«·› —… „‰"
               Height          =   315
               Index           =   0
               Left            =   4980
               RightToLeft     =   -1  'True
               TabIndex        =   23
               Top             =   240
               Width           =   945
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "œ·«·«  «·«·Ê«‰"
            Height          =   405
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   11
            Top             =   480
            Width           =   4590
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
               Caption         =   "€Ì— „”œœ ﬂ«„·«"
               Height          =   255
               Left            =   360
               RightToLeft     =   -1  'True
               TabIndex        =   15
               Top             =   120
               Width           =   1095
            End
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               BackColor       =   &H008080FF&
               Height          =   255
               Left            =   1320
               RightToLeft     =   -1  'True
               TabIndex        =   14
               Top             =   120
               Width           =   375
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               Caption         =   "„”œœ Ã“∆Ì"
               Height          =   255
               Left            =   1800
               RightToLeft     =   -1  'True
               TabIndex        =   13
               Top             =   120
               Width           =   1215
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H0000FFFF&
               Height          =   255
               Left            =   3240
               RightToLeft     =   -1  'True
               TabIndex        =   12
               Top             =   120
               Width           =   375
            End
         End
         Begin VB.Frame FraHeader 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   450
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   4
            Top             =   0
            Width           =   17745
            Begin VB.TextBox TxtVac_ID 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H000000FF&
               Height          =   240
               Left            =   3030
               RightToLeft     =   -1  'True
               TabIndex        =   9
               Top             =   510
               Visible         =   0   'False
               Width           =   945
            End
            Begin VB.TextBox TxtModFlg 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H0000FF00&
               Enabled         =   0   'False
               Height          =   285
               Left            =   1920
               RightToLeft     =   -1  'True
               TabIndex        =   8
               Text            =   "modflag"
               Top             =   120
               Visible         =   0   'False
               Width           =   465
            End
            Begin VB.Frame Frmo2 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   375
               Left            =   540
               RightToLeft     =   -1  'True
               TabIndex        =   5
               Top             =   450
               Visible         =   0   'False
               Width           =   3105
               Begin MSDataListLib.DataCombo DCUser 
                  CausesValidation=   0   'False
                  Height          =   315
                  Left            =   -255
                  TabIndex        =   6
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
                  TabIndex        =   7
                  Top             =   45
                  Width           =   855
               End
            End
            Begin MSComctlLib.ImageList GrdImageList 
               Left            =   2760
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
                     Picture         =   "ReRentAlarm.frx":0E33
                     Key             =   "CompanyName"
                  EndProperty
                  BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "ReRentAlarm.frx":11CD
                     Key             =   "Ser"
                  EndProperty
                  BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "ReRentAlarm.frx":1567
                     Key             =   "Vac_Name"
                  EndProperty
                  BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "ReRentAlarm.frx":1901
                     Key             =   "ShareCount"
                  EndProperty
                  BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "ReRentAlarm.frx":1C9B
                     Key             =   "Dis_Count"
                  EndProperty
                  BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "ReRentAlarm.frx":2035
                     Key             =   "Bouns"
                  EndProperty
                  BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "ReRentAlarm.frx":23CF
                     Key             =   "SharesValue"
                  EndProperty
                  BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "ReRentAlarm.frx":2969
                     Key             =   "BuyValue"
                  EndProperty
               EndProperty
            End
            Begin VB.Image ImgFavorites 
               Height          =   390
               Left            =   5760
               Picture         =   "ReRentAlarm.frx":2D03
               Stretch         =   -1  'True
               Top             =   0
               Width           =   525
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   " ‰»ÌÂ«  «·«ÌÃ«—«  «·„” ÕﬁÂ Œ·«· › —…"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   14.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   375
               Index           =   2
               Left            =   9600
               RightToLeft     =   -1  'True
               TabIndex        =   10
               Top             =   120
               Width           =   5880
            End
         End
         Begin VSFlex8UCtl.VSFlexGrid FgItems 
            Height          =   8985
            Index           =   1
            Left            =   24705
            TabIndex        =   39
            Top             =   780
            Width           =   17505
            _cx             =   30877
            _cy             =   15849
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
            Cols            =   5
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"ReRentAlarm.frx":696B
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
         Begin ImpulseButton.ISButton btnCancel 
            Height          =   270
            Left            =   120
            TabIndex        =   40
            Top             =   6270
            Width           =   750
            _ExtentX        =   1323
            _ExtentY        =   476
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
            ButtonImage     =   "ReRentAlarm.frx":6A2B
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton CmdPrint 
            Height          =   330
            Left            =   1080
            TabIndex        =   41
            Top             =   6270
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   582
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
            ButtonImage     =   "ReRentAlarm.frx":6DC5
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin VSFlex8Ctl.VSFlexGrid Grid 
            Height          =   3360
            Left            =   1980
            TabIndex        =   42
            Top             =   9480
            Width           =   14025
            _cx             =   24739
            _cy             =   5927
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
            Rows            =   2
            Cols            =   14
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"ReRentAlarm.frx":715F
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
         Begin VSFlex8UCtl.VSFlexGrid GridInstallments 
            Height          =   4545
            Left            =   0
            TabIndex        =   43
            Top             =   1470
            Width           =   17805
            _cx             =   31406
            _cy             =   8017
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
            Rows            =   12
            Cols            =   43
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"ReRentAlarm.frx":7398
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
            WallPaperAlignment=   9
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   24
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   270
            Left            =   0
            TabIndex        =   44
            Top             =   975
            Visible         =   0   'False
            Width           =   1770
            _ExtentX        =   3122
            _ExtentY        =   476
            _Version        =   393216
            Format          =   95944705
            CurrentDate     =   41640
         End
      End
   End
End
Attribute VB_Name = "RSRentAlarm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

     Dim My_SQL As String

Public mIndex As Integer

Private Sub cmdCreateProduction_Click()

    
'    If DBCboClientName.BoundText = "" Then
'        MsgBox ("·« Ì„ﬂ‰ «‰‘«¡ «„— «·«‰ «Ã »œÊ‰ «œŒ«· «·⁄„Ì·")
'        DBCboClientName.SetFocus
'        Exit Sub
'    End If
                    
    If DCboStoreName2.BoundText = "" Then
If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox ("·« Ì„ﬂ‰ «‰‘«¡ «„— «·«‰ «Ã »œÊ‰ «œŒ«· «·„Œ“‰")
Else
        MsgBox (" Must Select Store")
        
End If
        DCboStoreName2.SetFocus
        Exit Sub
    End If
                    
                    
    Dim Transaction_ID As Long
Dim Transaction_serial As String
    Dim RsTemp As New ADODB.Recordset
    Dim RsTest As New ADODB.Recordset
    Dim RowNum As Long
    Dim Transaction_Date As Date
    Transaction_Date = Date
Transaction_Type = 26
Dim mBranchID As Integer
Dim rsBranchDummy As New ADODB.Recordset

Dim i As Long
Dim s As String
s = "Select BranchId FROM TblStore Where StoreId = " & val(DCboStoreName2.BoundText)
rsBranchDummy.Open s, Cn, adOpenKeyset, adLockReadOnly
If Not rsBranchDummy.EOF Then
    mBranchID = val(rsBranchDummy!BranchID & "")
End If
    
Dim NoteSerial1 As String

            Dim Current_case As Integer, mBoxID As Long
        
       
          
    Dim rs As ADODB.Recordset

    Set rs = New ADODB.Recordset
  Dim rsDummy2 As New ADODB.Recordset
Dim notpayed As Double
notpayed = 0
 
 

My_SQL = " SELECT "
My_SQL = My_SQL & "       tblItems.GroupID"
My_SQL = My_SQL & " FROM   Transaction_Details      AS td"
My_SQL = My_SQL & "       INNER JOIN Transactions  AS t"
My_SQL = My_SQL & "            ON  t.Transaction_ID = td.Transaction_ID"
My_SQL = My_SQL & "            left Outer join tblItems on  td.Item_ID= tblItems.ItemID"
My_SQL = My_SQL & " Where IsNull(td.TransactionID4, 0) = 0"

My_SQL = My_SQL + " and (t.Transaction_Date >='" & SQLDate(txtFromDate.value) & "'"
My_SQL = My_SQL + " and  t.Transaction_Date <=" & SQLDate(txtToDate, True) & ")"
     
If dcBranch.Text <> "" Then
    My_SQL = My_SQL + "   AND (t.BranchID = " & val(dcBranch.BoundText) & ")"
End If
     

My_SQL = My_SQL + "   AND  t.Transaction_Type = 38"



My_SQL = My_SQL & " Group by"
My_SQL = My_SQL & "       tblItems.GroupID"
Set rs = New ADODB.Recordset
rs.Open My_SQL, Cn, adOpenStatic, adLockReadOnly
Dim RSTransDetails As New ADODB.Recordset
Do While Not rs.EOF

         TXTTransactionID4 = 0
         TxtNoteSerial14 = 0

         
   
            TxtNoteSerial14.Text = Voucher_coding(val(mBranchID), Date, 49, 0, , 26, , val(DCboStoreName2.BoundText))
         '  TxtNoteSerial14.Text = Voucher_coding(val(mBranchID), date, 49, 0, , 26, , val(DCboStoreName2.BoundText))
  
      '  TxtNoteSerial14.Text = Voucher_coding(val(mBranchID), Date, 49, 0, , 26, , val(DCboStoreName2.BoundText))
               
        Dim BranchID  As Double, StoreId As Double, StoreId2 As Double
               
         BranchID = val(mBranchID)
         StoreId = val(DCboStoreName2.BoundText)
         StoreId2 = val(DCboStoreName.BoundText)
                     
                     
         CostTOTAL = 0

         StoreId = val(DCboStoreName2.BoundText)
          
        If DCboStoreName2.Text = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "ÌÃ»  ÕœÌœ «·„Œ“‰"
            Else
                Msg = "Select Inventory First"
            End If
    
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
          If DCboStoreName2.Enabled = True Then
            DCboStoreName2.SetFocus
          SendKeys "{F4}"
            End If
          ' Cmd(2).Enabled = True
            Screen.MousePointer = vbDefault
          '  Cmd(2).Enabled = True
            Exit Sub
        End If
            
        
        
         If TxtNoteSerial14 = "" Then
            NoteSerial1 = Voucher_coding(val(BranchID), Transaction_Date, 49, 0, , 26)
            TxtNoteSerial14 = CStr(NoteSerial1)
         End If
        Transaction_ID = CStr(new_id("Transactions", "Transaction_ID", "", True))
         
          
        NoteSerial1 = Voucher_coding(val(BranchID), Transaction_Date, 49, 0, , 26)     '„»Ì⁄« 
            
        If Voucher_coding(val(mBranchID), Date, 49, 0, , 26, , val(DCboStoreName2.BoundText)) = "error" Then
                        MsgBox " ·« Ì„ﬂ‰ «÷«›…   Â–« «·”‰œ ·«‰ﬂ  ⁄œÌ  «·Õœ «·„”„ÊÕ »… „‰ «·”‰œ«   ": Exit Sub
                    Else
                               
                        If Voucher_coding(val(mBranchID), Date, 49, 0, , 26, , val(DCboStoreName2.BoundText)) = "" Then
                            NoteSerial1 = CStr(new_id("Transactions", "Transaction_Serial", "", True, "Transaction_Type=75"))
                        Else
                            NoteSerial1 = Voucher_coding(val(mBranchID), Date, 49, 0, , 26, , val(DCboStoreName2.BoundText))
                        End If
                    End If
               
        
        NoteSerial = Notes_coding(val(BranchID), Transaction_Date)
         If NoteSerial = "" Then
                    If NoteSerial = "error" Then
                        MsgBox " ·« Ì„ﬂ‰ «÷«›… ﬁÌÊœ ·Â–… «·⁄„·Ì… ·«‰ﬂ  ⁄œÌ  «·Õœ «·«ﬁ’Ì ··ﬁÌÊœ ﬂ„« Õœœ  ›Ì  —ﬁÌ„ «·”‰œ«  ": Exit Sub
                    ElseIf NoteSerial = "" Then
                            MsgBox " ·«Ì„ﬂ‰ «‰‘«¡ «·ﬁÌœ ·Â–« «·„” ‰œ ·«‰ﬂ Õœœ   —ﬁÌ„ ﬁÌÊœ ÌœÊÌ  ": Exit Sub
                         
                    End If
        End If
                   
            
        
          
                      
        
            StoreAccount = get_store_Account(CInt(StoreId), "Account_Code")
        
                TXTTransactionID4.Text = Transaction_ID
                TxtNoteSerial14.Text = NoteSerial1
                Transaction_serial = NoteSerial1
             Dim rsOut As New ADODB.Recordset
        
                    mBoxID = 0
         sql = "INSERT INTO  Transactions (  "
        sql = sql & " Transaction_ID ,"
        sql = sql & " BranchID ,"
        sql = sql & " NoteSerial ,"
        sql = sql & " NoteSerial1 ,"
        sql = sql & " boxId ,"
        sql = sql & " Transaction_serial ,"
        sql = sql & " Transaction_Date ,"
        sql = sql & " Transaction_Type ,"
        sql = sql & " BillBasedOn ,"
        sql = sql & " UserID ,"
        sql = sql & " Trans_DiscountType ,"
        sql = sql & " CusID ,"
        sql = sql & " StoreId ,"
        sql = sql & " StoreId1 ,"
        
        sql = sql & " PaymentType ,"
        sql = sql & " Emp_id ,"
        sql = sql & " Transaction_NetValue ,"
        sql = sql & " Vat, netvalue, PayedValue, "
        sql = sql & " Currency_rate, Currency_id,sumVatLine,DueDate,"
         sql = sql & " TransactionComment,MIxCode,MixID,CBoBasedON,OrderType )"
         
            
         sql = sql & " VALUES("
        sql = sql & " " & Transaction_ID & " ,"
        sql = sql & " " & BranchID & " ,"
        sql = sql & "'" & NoteSerial & "' ,"
        sql = sql & "'" & NoteSerial1 & "' ,"
        sql = sql & " " & val(BoxID) & " ,"
        sql = sql & "'" & Transaction_serial & "',"
        sql = sql & " " & SQLDate(Transaction_Date, True) & " ,"
        sql = sql & " " & 26 & " ,"
        sql = sql & " 0 ,"
        sql = sql & " " & user_id & " ,"
        sql = sql & " 0 ,"
        sql = sql & "   2 ,"
        sql = sql & " " & StoreId2 & " ,"
        sql = sql & " " & StoreId & " ,"
        sql = sql & " " & 0 & " ,"
        sql = sql & " " & val(Emp_id) & " ,"
        sql = sql & " " & val(txtTotalWithVat2) & " ,"
        sql = sql & " " & val(TxtVAt22) & " ,"
        sql = sql & " " & val(txtNet2) & " ,"
        sql = sql & " " & val(txtNet2) & " ,"
        sql = sql & " " & 1 & " ,"
        sql = sql & " " & 1 & " ,0,"
        sql = sql & " " & SQLDate(Transaction_Date, True) & " ,"
        sql = sql & "'" & TransactionComment & "',"
        sql = sql & "" & val(TxtMaxNo) & "," & val(TxtMaxNo) & ",0,0)"
        
         
        Cn.Execute sql
        
       
        My_SQL = " SELECT td.Item_ID,td.UnitID,"
        My_SQL = My_SQL & "       tblItems.ItemName,"
        
        My_SQL = My_SQL & "       TblItems.GroupID,"
        My_SQL = My_SQL & "       SUM(td.ShowQty)             Qty,"
        My_SQL = My_SQL & "       SUM(td.showPrice)           Price,"
        My_SQL = My_SQL & "       total = SUM(td.ShowQty) * SUM(td.ShowPrice)"
        My_SQL = My_SQL & "FROM   Transaction_Details      AS td"
        My_SQL = My_SQL & "       INNER JOIN Transactions  AS t"
        My_SQL = My_SQL & "            ON  t.Transaction_ID = td.Transaction_ID"
        My_SQL = My_SQL & "       LEFT OUTER JOIN TblItems"
        My_SQL = My_SQL & "            ON  TblItems.ItemID = td.Item_ID"
        My_SQL = My_SQL & "       LEFT OUTER JOIN Groups   AS g"
        My_SQL = My_SQL & "            ON  g.GroupID = TblItems.GroupID"
            
        My_SQL = My_SQL & " Where IsNull(td.TransactionID4, 0) = 0"
        
        My_SQL = My_SQL + " and (t.Transaction_Date >='" & SQLDate(txtFromDate.value) & "'"
        My_SQL = My_SQL + " and  t.Transaction_Date <=" & SQLDate(txtToDate, True) & ")"
             
        If dcBranch.Text <> "" Then
            My_SQL = My_SQL + "   AND (t.BranchID = " & val(dcBranch.BoundText) & ")"
        End If

        My_SQL = My_SQL + "   AND  t.Transaction_Type = 38"
        My_SQL = My_SQL + "   AND  IsNull(tblItems.GroupID,0)  =" & val(rs!GroupID & "")
        
        
        My_SQL = My_SQL & "Group by"
        My_SQL = My_SQL & "       td.Item_ID,td.UnitID,"
        My_SQL = My_SQL & "       tblItems.ItemName,"
        
        My_SQL = My_SQL & "       tblItems.GroupID"
        Set rsDummy2 = New ADODB.Recordset
        rsDummy2.Open My_SQL, Cn, adOpenStatic, adLockReadOnly
        Do While Not rsDummy2.EOF
                   
                    'For i = 1 To Fg.Rows - 1
        
          
                          '  CreateProduction BranchID, 0, Date, 26, 0, val(user_id), 0, DBCboClientName.BoundText, StoreId, 0, val(DcboEmp.BoundText), "«„— «‰ «Ã", val(rsDummy2!ID & ""), val(TXTTransactionID4)
                       ' End If
                        Set RSTransDetails = New ADODB.Recordset
             
                        StrSQL = "SELECT     dbo.Transaction_Details.* from dbo.Transaction_Details Where (Transaction_ID = -1)"
                        Set RSTransDetails = New ADODB.Recordset
                        RSTransDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
           
            
                        Dim mItemNo As Long, mUnitNo As Long, mQty As Long, mVAt2 As Double, mTotal As Double
                        Dim mwidtj As Double, mhight As Double, mTotalAdd As Double, mTotalDisc As Double, mNet As Double, mTotalWithVat As Double, mLength As Double
                        Dim mItemName2 As String, mCostPercent As Double
                        Dim mRemark As String
                        mItemNo = val(rsDummy2!Item_ID & "")
                        If mItemNo = 0 Then GoTo NextRow
                
                       
                    
                    mUnitNo = val(rsDummy2!UnitID & "")
                   
                    mQty = val(rsDummy2!Qty & "")
                    
                    mPrice = val(rsDummy2!Price & "")
        
                    'mCostPercent = val(Fg.TextMatrix(i, Fg.ColIndex("PercentCost")))
                    
                RSTransDetails.AddNew
                RSTransDetails("Transaction_ID").value = Transaction_ID
                RSTransDetails("ColorID").value = 1
                RSTransDetails("ItemSize").value = 1
                RSTransDetails("ClassId").value = 1
                RSTransDetails("Item_ID").value = mItemNo
                RSTransDetails("UnitID").value = mUnitNo
                RSTransDetails("SHOWQTY").value = mQty
                RSTransDetails("PercentCost").value = mCostPercent
                RSTransDetails("showPrice").value = mPrice
                RSTransDetails("Lineexpenses").value = mPrice
                
                RSTransDetails("ItemDiscountType").value = 2
                
                If SystemOptions.TypicalProduction = False Then
        
                    RSTransDetails("CostPrice").value = ModItemCostPrice.GetCostItemPrice(mItemNo, 0, , , SystemOptions.SysMainStockCostMethod, , , Date, 0, RSTransDetails("UnitID").value, StoreId)
        
                    If RSTransDetails("CostPrice").value = 0 Then
                        RSTransDetails("CostPrice").value = ModItemCostPrice.GetCostItemPrice(mItemNo, 0, , , LastPurPriceType, , , Date, 0, RSTransDetails("UnitID").value, val(Me.DCboStoreName.BoundText))
                        
                    End If
                      RSTransDetails("CostPrice").value = mPrice
                Else
                    RSTransDetails("CostPrice").value = 0
                
                End If
                              
                  
                              '«·ÊÕœ« 
               
                Dim RsUnitData As ADODB.Recordset
                Dim LngCurItemID As Long
                Dim LngUnitID As Long
                Dim DblQty As Double
            
                LngCurItemID = val(mItemNo)
                LngUnitID = val(mUnitNo)
                DblQty = val(mQty)
        
                StrSQL = "Select * From TblItemsUnits Where ItemID=" & LngCurItemID
                StrSQL = StrSQL + " AND UnitID=" & LngUnitID
                Set RsUnitData = New ADODB.Recordset
                RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        
                If Not (RsUnitData.BOF Or RsUnitData.EOF) Then
                    RSTransDetails("QtyBySmalltUnit").value = RsUnitData("UnitFactor").value
                    RSTransDetails("Quantity").value = RSTransDetails("QtyBySmalltUnit").value * RSTransDetails("showqty").value
                    RSTransDetails("OpeningSalesQty").value = RSTransDetails("Quantity").value
                    RSTransDetails("OpeningSalesValue").value = RSTransDetails("CostPrice").value * val(mQty)
                    If val(RSTransDetails("QtyBySmalltUnit").value & "") <> 0 Then
                        RSTransDetails("Price").value = val(IIf((mPrice = 0), 0, val(mPrice))) / RSTransDetails("QtyBySmalltUnit").value
                    Else
                        RSTransDetails("Price").value = val(IIf((mPrice = 0), 0, val(mPrice)))
                    End If
                
                End If
        
            
                 UpdateTransactionsCost CStr(Transaction_ID)
                 RSTransDetails.update
            
              '  Dim i As Integer
                'Dim sql As String
          '  End If
NextRow:
        
        
        NoteSerial = Notes_coding(val(BranchID), Transaction_Date)
        
        
        
        
        
        '38
        '
        '
'                        rsDummy2!TransactionID4 = val(TXTTransactionID4)
'                        rsDummy2!NoteSerial14 = Trim(TxtNoteSerial14)
'                        rsDummy2.update
                        rsDummy2.MoveNext
                        
                    Loop
                    
        My_SQL = " Select Transaction_Details.* "
        
        My_SQL = My_SQL & "FROM   Transaction_Details      "
        My_SQL = My_SQL & "       INNER JOIN Transactions  AS t"
        My_SQL = My_SQL & "            ON  t.Transaction_ID = Transaction_Details.Transaction_ID"
        My_SQL = My_SQL & "            Left Outer join  TblItems On  TblItems.ItemID= Transaction_Details.Item_ID"
        
        My_SQL = My_SQL & " Where IsNull(Transaction_Details.TransactionID4, 0) = 0"
        
        My_SQL = My_SQL + " and (t.Transaction_Date >='" & SQLDate(txtFromDate.value) & "'"
        My_SQL = My_SQL + " and  t.Transaction_Date <=" & SQLDate(txtToDate, True) & ")"
             
        If dcBranch.Text <> "" Then
            My_SQL = My_SQL + "   AND (t.BranchID = " & val(dcBranch.BoundText) & ")"
        End If

        My_SQL = My_SQL + "   AND  t.Transaction_Type = 38"
        My_SQL = My_SQL + "   AND  IsNull(TblItems.GroupID,0) = " & val(rs!GroupID & "")
        Dim rsDummy As ADODB.Recordset
        Set rsDummy = New ADODB.Recordset
        
        rsDummy.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic
        Do While Not rsDummy.EOF
            rsDummy!TransactionID4 = val(TXTTransactionID4)
            rsDummy!NoteSerial14 = Trim(TxtNoteSerial14)
            rsDummy.update
            rsDummy.MoveNext
        Loop
rs.MoveNext
TXTTransactionID4 = 0
TxtNoteSerial14 = 0
Loop
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox " „ «‰‘«¡ «„— «·«‰ «Ã"
Else
MsgBox "Prody=uction order Created"
End If

cmdCreateProduction.Enabled = False
End Sub
Private Sub BtnCancel_Click()
    Me.Hide
End Sub

Private Sub Check17_Click()
ReLineGrid
End Sub

Private Sub Cmd_Click(Index As Integer)
If Index = 9 Then Index = 0
Select Case Index
Case 0
    FillGrid
Case 1
    FillGrid2
    CalcCostPercent
End Select
End Sub

       

  Public Sub FillGrid2(Optional str As String)

  '  On Error GoTo ErrTrap
On Error Resume Next
    Dim i As Integer
    Dim rs As ADODB.Recordset

    Set rs = New ADODB.Recordset
  
Dim notpayed As Double
notpayed = 0
 
 

My_SQL = " SELECT td.Item_ID,td.UnitID,"

If SystemOptions.UserInterface = EnglishInterface Then
    My_SQL = My_SQL & "       tblItems.ItemNamee ItemName ,"
    My_SQL = My_SQL & "       G.GroupNamee GroupName,"
Else
    My_SQL = My_SQL & "       tblItems.ItemName,"
    My_SQL = My_SQL & "       G.GroupName,"

End If
My_SQL = My_SQL & "       g.GroupID,"
My_SQL = My_SQL & "       SUM(td.ShowQty)             Qty,"
My_SQL = My_SQL & "       SUM(td.showPrice)           Price,"
My_SQL = My_SQL & "       total = SUM(td.ShowQty) * SUM(td.ShowPrice)"
My_SQL = My_SQL & "FROM   Transaction_Details      AS td"
My_SQL = My_SQL & "       INNER JOIN Transactions  AS t"
My_SQL = My_SQL & "            ON  t.Transaction_ID = td.Transaction_ID"
My_SQL = My_SQL & "       LEFT OUTER JOIN TblItems"
My_SQL = My_SQL & "            ON  TblItems.ItemID = td.Item_ID"
My_SQL = My_SQL & "       LEFT OUTER JOIN Groups   AS g"
My_SQL = My_SQL & "            ON  g.GroupID = TblItems.GroupID"
       
My_SQL = My_SQL & " Where IsNull(td.TransactionID4, 0) = 0"

My_SQL = My_SQL + " and (t.Transaction_Date >='" & SQLDate(txtFromDate.value) & "'"
My_SQL = My_SQL + " and  t.Transaction_Date <=" & SQLDate(txtToDate, True) & ")"
If dcBranch.Text <> "" Then
    My_SQL = My_SQL + "   AND (t.BranchID = " & val(dcBranch.BoundText) & ")"
End If

My_SQL = My_SQL + "   AND  t.Transaction_Type = 38"



My_SQL = My_SQL & "Group by"
My_SQL = My_SQL & "       td.Item_ID,td.UnitID,"

If SystemOptions.UserInterface = EnglishInterface Then
    My_SQL = My_SQL & "       tblItems.ItemNamee ,"
    My_SQL = My_SQL & "       G.GroupNamee ,"
Else
    My_SQL = My_SQL & "       tblItems.ItemName,"
    My_SQL = My_SQL & "       G.GroupName,"

End If


My_SQL = My_SQL & "       G.GroupID"


Fg.Rows = 1
loadgrid My_SQL, Fg, True, False
 cmdCreateProduction.Enabled = True
'ReLineGrid
End Sub


Private Sub CreateProduction(BranchID As Double, _
BoxID As Double, _
Transaction_Date As Date, _
Transaction_Type As Double, _
CBoBasedON As Double, _
UserID As Double, _
Trans_DiscountType As Double, _
CusID As Double, _
StoreId As Double, _
PaymentType As Double, _
Emp_id As Double, _
TransactionComment As String, ByVal mmID As Long, Transaction_ID As Long)

Dim BolTemp As Boolean
Dim sql As String
Dim Msg As String
Dim NoteID As Long

Dim Transaction_ID1 As Long
Dim Transaction_serial As String
Dim NoteSerial As String
Dim NoteSerial1 As String
Dim StrSQL As String
Dim Percetage As Double
Dim AccountVATCreit As String
Dim mPrice As Double
Dim rsDummy As New ADODB.Recordset
' «·”⁄— Â‰« ÂÊ ’«›Ï «·”⁄— »⁄œ Œ’„ «·«÷«›Ï Ê«·Œ’Ê„« 
'
'PercentgValueAddedAccount_Transec date, 21, 0, AccountVATCreit, Percetage
'PercetageVat = Percetage

'BillTOTAL = 0


  

             
        StrSQL = "SELECT     dbo.Transaction_Details.* from dbo.Transaction_Details Where (Transaction_ID = -1)"
        RSTransDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
           
        StrSQL = "Select Item_ID,UnitID,Sum(ShowQty) Qty,Sum(ShowPrice) Price,Sum(PercentCost) PercentCost from Transaction_Details Where ID = " & mmID
        StrSQL = StrSQL & " Group By Item_ID,UnitID "
        Set rsDummy = New ADODB.Recordset
        rsDummy.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic
            If Not rsDummy.EOF Then
            
                Dim mItemNo As Long, mUnitNo As Long, mQty As Long, mVAt2 As Double, mTotal As Double
                Dim mwidtj As Double, mhight As Double, mTotalAdd As Double, mTotalDisc As Double, mNet As Double, mTotalWithVat As Double, mLength As Double
                Dim mItemName2 As String, mCostPercent As Double
                Dim mRemark As String
                mItemNo = val(rsDummy!Item_ID & "")
                If mItemNo = 0 Then GoTo NextRow
                
                       
                    mItemNo = val(rsDummy!Item_ID & "")
                   
                    mUnitNo = val(rsDummy!UnitID & "")
                    mQty = val(rsDummy!Qty & "")
                    mPrice = val(rsDummy!Price & "")
        '            mwidtj = val(rsDummy!widtj & "")
        '            mhight = val(rsDummy!hight & "")
        '            mLength = val(rsDummy!Length & "")
                   ' mTotal = val(rsDummy!Total & "")
                '    mRemark = Trim(rsDummy!Remark & "")
                '    mTotalDisc = val(rsDummy!TotalDisc & "")
                '    mTotalAdd = val(rsDummy!TotalAdd & "")
                '    mNet = val(rsDummy!net & "")
                '    mVAt2 = val(rsDummy!Vat2 & "")
                   ' mTotalWithVat = val(rsDummy!TotalWithVat & "")
                  '  mPrice = (val(mTotal) + val(mTotalAdd)) / val(mQty)
                    mCostPercent = val(rsDummy!PercentCost & "")
                    
                RSTransDetails.AddNew
                RSTransDetails("Transaction_ID").value = Transaction_ID
                RSTransDetails("ColorID").value = 1
                RSTransDetails("ItemSize").value = 1
                RSTransDetails("ClassId").value = 1
                RSTransDetails("Item_ID").value = mItemNo
                RSTransDetails("UnitID").value = mUnitNo
                RSTransDetails("SHOWQTY").value = mQty
                RSTransDetails("PercentCost").value = mCostPercent
                RSTransDetails("showPrice").value = mPrice
                RSTransDetails("Lineexpenses").value = mPrice
                
                RSTransDetails("ItemDiscountType").value = 2
                
                If SystemOptions.TypicalProduction = False Then
        
                    RSTransDetails("CostPrice").value = ModItemCostPrice.GetCostItemPrice(mItemNo, 0, , , SystemOptions.SysMainStockCostMethod, , , Date, 0, RSTransDetails("UnitID").value, StoreId)
        
                    If RSTransDetails("CostPrice").value = 0 Then
                        RSTransDetails("CostPrice").value = ModItemCostPrice.GetCostItemPrice(mItemNo, 0, , , LastPurPriceType, , , Date, 0, RSTransDetails("UnitID").value, val(Me.DCboStoreName.BoundText))
                        
                    End If
                      RSTransDetails("CostPrice").value = mPrice
                Else
                    RSTransDetails("CostPrice").value = 0
                
                End If
                              
                  
                              '«·ÊÕœ« 
               
                Dim RsUnitData As ADODB.Recordset
                Dim LngCurItemID As Long
                Dim LngUnitID As Long
                Dim DblQty As Double
            
                LngCurItemID = val(mItemNo)
                LngUnitID = val(mUnitNo)
                DblQty = val(mQty)
        
                StrSQL = "Select * From TblItemsUnits Where ItemID=" & LngCurItemID
                StrSQL = StrSQL + " AND UnitID=" & LngUnitID
                Set RsUnitData = New ADODB.Recordset
                RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        
                If Not (RsUnitData.BOF Or RsUnitData.EOF) Then
                    RSTransDetails("QtyBySmalltUnit").value = RsUnitData("UnitFactor").value
                    RSTransDetails("Quantity").value = RSTransDetails("QtyBySmalltUnit").value * RSTransDetails("showqty").value
                    RSTransDetails("OpeningSalesQty").value = RSTransDetails("Quantity").value
                    RSTransDetails("OpeningSalesValue").value = RSTransDetails("CostPrice").value * val(mQty)
                    RSTransDetails("Price").value = val(IIf((mPrice = 0), 0, val(mPrice))) / RSTransDetails("QtyBySmalltUnit").value
                
                End If
        
            
                 UpdateTransactionsCost CStr(Transaction_ID)
                 RSTransDetails.update
            
              '  Dim i As Integer
                'Dim sql As String
            End If
NextRow:
        
        
        NoteSerial = Notes_coding(val(BranchID), Transaction_Date)





'***********************
'End If
  
'******************************************************issueVoucher








     
 
    '
 
ErrTrap:



 

End Sub
     
Public Sub FillGrid(Optional str As String)

  '  On Error GoTo ErrTrap
On Error Resume Next
    Dim i As Integer
    Dim rs As ADODB.Recordset

    Set rs = New ADODB.Recordset
  
Dim notpayed As Double
notpayed = 0
 
 If str = "" Then

My_SQL = " SELECT     DISTINCT dbo.TblCustemers.CusName AS CusName, dbo.TblCustemers.Cus_mobile AS Cus_mobile, dbo.TblCustemers.CusNamee , "
My_SQL = My_SQL & "                      dbo.TblContract.NoteSerial1 AS NoteSerial11, dbo.TblContract.CusID , dbo.TblContract.StrDate , dbo.TblContractInstallments.Installdate ,"
My_SQL = My_SQL & "                      dbo.TblContractInstallments.InstalldateH , "
My_SQL = My_SQL & "                      dbo.TblContractInstallments.RentValue AS RentValue_1, dbo.TblContractInstallments.Insurance AS Insurance_1, dbo.TblContract.ContNo ,"
My_SQL = My_SQL & "                      TblContractInstallments.ID , TblContractInstallments.InstallNo, TblContractInstallments.Installdateh, TblContractInstallments.Installdate, TblContractInstallments.installValue,  TblContractInstallments.hijri, TblContractInstallments.RentValue, TblContractInstallments.Commissions, TblContractInstallments.Insurance, TblContractInstallments.Water, TblContractInstallments.Electric, TblContractInstallments.allocations, TblContractInstallments.Countsofall, TblContractInstallments.Doneofall,"
'My_SQL = My_SQL & "                      { fn IFNULL(dbo.ContracttBillInstallmentsDone.[Value], 0) } AS Allpayed, { fn IFNULL(dbo.TblContractInstallments.installValue, 0)"
'My_SQL = My_SQL & "                      } - { fn IFNULL(dbo.ContracttBillInstallmentsDone.[Value], 0) } AS newremains"
My_SQL = My_SQL & "                       dbo.TblAqar.aqarNo AS IaqarNo, dbo.TblAqar.aqarname AS Iaqarname,"
My_SQL = My_SQL & "                      dbo.TblAkarUnit.name AS Unitname, dbo.TblAkarUnit.namee AS Unitnamee, dbo.TblAqarDetai.unitno AS unitnoNam, dbo.TblContract.Phone AS Phone"
My_SQL = My_SQL & " FROM         dbo.TblContractInstallments INNER JOIN"
My_SQL = My_SQL & "                      dbo.TblContract ON dbo.TblContractInstallments.ContNo = dbo.TblContract.ContNo INNER JOIN"
My_SQL = My_SQL & "                      dbo.TblCustemers ON dbo.TblContract.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
My_SQL = My_SQL & "                      dbo.TblAqarDetai ON dbo.TblContract.UnitNo = dbo.TblAqarDetai.Id LEFT OUTER JOIN"
My_SQL = My_SQL & "                      dbo.TblAkarUnit ON dbo.TblContract.UnitType = dbo.TblAkarUnit.id LEFT OUTER JOIN"
My_SQL = My_SQL & "                      dbo.TblAqar ON dbo.TblContract.Iqar = dbo.TblAqar.Aqarid LEFT OUTER JOIN"
My_SQL = My_SQL & "                      dbo.ContracttBillInstallmentsDone ON dbo.TblContract.ContNo = dbo.ContracttBillInstallmentsDone.istallid"
My_SQL = My_SQL & " WHERE   (dbo.TblContractInstallments.Status = 0 OR dbo.TblContractInstallments.Status IS NULL)"

        My_SQL = My_SQL + " and (Installdate >='" & SQLDate(Fromdate.value) & "'"
    My_SQL = My_SQL + " and  Installdate <=" & SQLDate(ToDate, True) & ")"
     
 
    'My_SQL = My_SQL + "   AND (dbo.TblContract.Branch_NO = " & Current_branch & ")"
 

My_SQL = My_SQL + "   order by TblContractInstallments.Installdate "

Else

My_SQL = str
 
 
         

End If
   
'  My_SQL = "SELECT DISTINCT "
'   My_SQL = My_SQL & "                    dbo.TblCustemers.CusName AS CusName, dbo.TblCustemers.Cus_mobile AS Cus_mobile, dbo.TblCustemers.CusNamee,"
'  My_SQL = My_SQL & "                     dbo.TblContract.NoteSerial1 AS NoteSerial11, dbo.TblContract.CusID, dbo.TblContract.StrDate, dbo.TblContractInstallments.Installdate,"
'My_SQL = My_SQL & "                       dbo.TblContractInstallments.InstalldateH, dbo.TblContractInstallments.RentValue AS RentValue_1, dbo.TblContractInstallments.Insurance AS Insurance_1,"
'My_SQL = My_SQL & "                       dbo.TblContract.ContNo, dbo.TblContractInstallments.id, dbo.TblContractInstallments.InstallNo, dbo.TblContractInstallments.InstalldateH AS Expr1,"
'My_SQL = My_SQL & "                       dbo.TblContractInstallments.Installdate AS Expr2, dbo.TblContractInstallments.installValue, dbo.TblContractInstallments.hijri, dbo.TblContractInstallments.RentValue,"
'My_SQL = My_SQL & "                       dbo.TblContractInstallments.Commissions, dbo.TblContractInstallments.Insurance, dbo.TblContractInstallments.Water, dbo.TblContractInstallments.Electric,"
'My_SQL = My_SQL & "                       dbo.TblContractInstallments.allocations, dbo.TblContractInstallments.Countsofall, dbo.TblContractInstallments.Doneofall, dbo.TblAqar.aqarNo AS IaqarNo,"
'My_SQL = My_SQL & "                       dbo.TblAqar.aqarname AS Iaqarname, dbo.TblAkarUnit.name AS Unitname, dbo.TblAkarUnit.namee AS Unitnamee, dbo.TblAqarDetai.unitno AS unitnoNam,"
'My_SQL = My_SQL & "                       dbo.TblContract.Phone AS Phone"
'My_SQL = My_SQL & " FROM         dbo.TblContractInstallments INNER JOIN"
'My_SQL = My_SQL & "                       dbo.TblContract ON dbo.TblContractInstallments.ContNo = dbo.TblContract.ContNo INNER JOIN"
'My_SQL = My_SQL & "                       dbo.TblCustemers ON dbo.TblContract.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
'My_SQL = My_SQL & "                       dbo.TblAqarDetai ON dbo.TblContract.UnitNo = dbo.TblAqarDetai.Id LEFT OUTER JOIN"
'My_SQL = My_SQL & "                       dbo.TblAkarUnit ON dbo.TblContract.UnitType = dbo.TblAkarUnit.id LEFT OUTER JOIN"
'My_SQL = My_SQL & "                       dbo.TblAqar ON dbo.TblContract.Iqar = dbo.TblAqar.Aqarid LEFT OUTER JOIN"
'My_SQL = My_SQL & "                       dbo.ContracttBillInstallmentsDone ON dbo.TblContract.ContNo = dbo.ContracttBillInstallmentsDone.istallid"
'My_SQL = My_SQL & " Where (dbo.TblContract.ContNo = 605)"

 Dim ActualTotal As Double
rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText
'    rs1.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
      With Me.GridInstallments
       .Rows = 1
        .Clear flexClearScrollable

        If rs.RecordCount > 0 Then
           .Rows = rs.RecordCount + 1
           rs.MoveFirst
'
            For i = 1 To .Rows - 1
              .TextMatrix(i, .ColIndex("Installid")) = (IIf(IsNull(rs.Fields("id").value), 0, rs.Fields("id").value))
               .TextMatrix(i, .ColIndex("InstallNo")) = (IIf(IsNull(rs.Fields("InstallNo").value), 0, rs.Fields("InstallNo").value))
.TextMatrix(i, .ColIndex("NoteSerial1")) = (IIf(IsNull(rs.Fields("NoteSerial11").value), "", rs.Fields("NoteSerial11").value))
.TextMatrix(i, .ColIndex("Cus_mobile")) = (IIf(IsNull(rs.Fields("Cus_mobile").value), "", rs.Fields("Cus_mobile").value))
.TextMatrix(i, .ColIndex("Iaqarname")) = (IIf(IsNull(rs.Fields("Iaqarname").value), "", rs.Fields("Iaqarname").value))
.TextMatrix(i, .ColIndex("unitnoNam")) = (IIf(IsNull(rs.Fields("unitnoNam").value), "", rs.Fields("unitnoNam").value))
.TextMatrix(i, .ColIndex("unitnoNam")) = (IIf(IsNull(rs.Fields("unitnoNam").value), "", rs.Fields("unitnoNam").value))

                
 .TextMatrix(i, .ColIndex("Due_DateH")) = (IIf(IsNull(rs.Fields("Installdateh").value), ToHijriDate(Date), rs.Fields("Installdateh").value))
  .TextMatrix(i, .ColIndex("Due_Date")) = IIf(IsNull(rs.Fields("Installdate").value), Date, rs.Fields("Installdate").value)
  DTPicker1.value = IIf(IsNull(rs.Fields("Installdate").value), Date, rs.Fields("Installdate").value)
 .TextMatrix(i, .ColIndex("DelayDay")) = DateDiff("d", DTPicker1.value, Date)
    .TextMatrix(i, .ColIndex("Value")) = (IIf(IsNull(rs.Fields("installValue").value), 0, rs.Fields("installValue").value))
     
                          ActualTotal = getinsttPayedTocontract(val(rs.Fields("id").value))
 .TextMatrix(i, .ColIndex("payed")) = ActualTotal
  .TextMatrix(i, .ColIndex("Remains")) = Round(val(.TextMatrix(i, .ColIndex("Value"))), 2) - Round(val(.TextMatrix(i, .ColIndex("payed"))), 2)
  
  
If val(.TextMatrix(i, .ColIndex("Remains"))) < 1 Then
'salim salim salah mno llah
'Cn.Execute "Update TblContractInstallments set Status=1 where id=" & rs.Fields("id").value & ""
'FillGrid
Else
'Cn.Execute "Update TblContractInstallments set Status=null where id=" & rs.Fields("id").value & ""
End If


If ActualTotal = 0 Then
          .Cell(flexcpBackColor, i, 1, i, 37) = &H8080FF
Else
          .Cell(flexcpBackColor, i, 1, i, 37) = vbYellow
End If
     
     
     .TextMatrix(i, .ColIndex("CusID")) = (IIf(IsNull(rs.Fields("CusID").value), "", rs.Fields("CusID").value))
   
   If SystemOptions.UserInterface = ArabicInterface Then
   .TextMatrix(i, .ColIndex("Unitname")) = (IIf(IsNull(rs.Fields("Unitname").value), "", rs.Fields("Unitname").value))
   .TextMatrix(i, .ColIndex("CusName")) = (IIf(IsNull(rs.Fields("CusName").value), "", rs.Fields("CusName").value))
   Else
   .TextMatrix(i, .ColIndex("Unitname")) = (IIf(IsNull(rs.Fields("Unitnamee").value), "", rs.Fields("Unitnamee").value))
   .TextMatrix(i, .ColIndex("CusName")) = (IIf(IsNull(rs.Fields("CusNamee").value), "", rs.Fields("CusNamee").value))
   End If
 .TextMatrix(i, .ColIndex("hijri")) = (IIf(IsNull(rs.Fields("hijri").value), 0, rs.Fields("hijri").value))   '
   '.Cell(flexcpChecked, i, .ColIndex("Select")) = flexChecked
 '
    .TextMatrix(i, .ColIndex("RentValue")) = (IIf(IsNull(rs.Fields("RentValue").value), 0, rs.Fields("RentValue").value))
    .TextMatrix(i, .ColIndex("Commissions")) = (IIf(IsNull(rs.Fields("Commissions").value), 0, rs.Fields("Commissions").value))
    .TextMatrix(i, .ColIndex("Insurance")) = (IIf(IsNull(rs.Fields("Insurance").value), 0, rs.Fields("Insurance").value))
    .TextMatrix(i, .ColIndex("Water")) = (IIf(IsNull(rs.Fields("Water").value), 0, rs.Fields("Water").value))
    .TextMatrix(i, .ColIndex("Electric")) = (IIf(IsNull(rs.Fields("Electric").value), 0, rs.Fields("Electric").value))
    .TextMatrix(i, .ColIndex("TelandNet")) = (IIf(IsNull(rs.Fields("Phone").value), 0, rs.Fields("Phone").value))
 
    
       .TextMatrix(i, .ColIndex("allocations")) = (IIf(IsNull(rs.Fields("allocations").value), 0, rs.Fields("allocations").value))
.TextMatrix(i, .ColIndex("Countsofall")) = (IIf(IsNull(rs.Fields("Countsofall").value), 0, rs.Fields("Countsofall").value))
.TextMatrix(i, .ColIndex("Doneofall")) = (IIf(IsNull(rs.Fields("Doneofall").value), 0, rs.Fields("Doneofall").value))

        rs.MoveNext
            Next i
 
            rs.Close
        End If
  ' .AutoSize 1, .Cols - 1, False

        .RowHeight(-1) = 300
    End With
ReLineGrid
End Sub



Private Sub ReLineGrid()
    Dim IntCounter As Integer
    IntCounter = 0
    Dim i As Integer
 
    Dim Percenrage As Double
 
 
    IntCounter = 0
  Me.TxtTotalContract.Text = 0
  Me.TxtCommiValue.Text = 0
    Me.TxtInsuranceValue.Text = 0
      Me.TxtWater.Text = 0
      Me.TxtElectricity.Text = 0
        Me.TxtPhone.Text = 0
     
    With Me.GridInstallments

        For i = .FixedRows To .Rows - 1
                                   If Check17.value = vbChecked Then
                .TextMatrix(i, .ColIndex("Send")) = -1
                Else
                .TextMatrix(i, .ColIndex("Send")) = 0
               
      End If
    
            If .TextMatrix(i, .ColIndex("Send")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("Ser")) = IntCounter
              
                '''////

                '''//
                
                
                     If .Cell(flexcpChecked, i, .ColIndex("Send")) = flexChecked Then
  Me.TxtTotalContract.Text = val(Me.TxtTotalContract.Text) + val(.TextMatrix(i, .ColIndex("RentValue")))
  Me.TxtCommiValue.Text = val(Me.TxtCommiValue.Text) + val(.TextMatrix(i, .ColIndex("Commissions")))
  Me.TxtInsuranceValue.Text = val(Me.TxtInsuranceValue.Text) + val(.TextMatrix(i, .ColIndex("Insurance")))
  Me.TxtWater.Text = val(Me.TxtWater.Text) + val(.TextMatrix(i, .ColIndex("Water")))
  Me.TxtElectricity.Text = val(Me.TxtElectricity.Text) + val(.TextMatrix(i, .ColIndex("Electric")))
  Me.TxtPhone.Text = val(Me.TxtPhone.Text) + val(.TextMatrix(i, .ColIndex("TelandNet")))
  
  End If
  
     
         
            End If

        Next i
   
    End With

End Sub


Private Sub CmdPrint_Click()
    On Error Resume Next
   ' Dim GrdBack As ClsBackGroundPic
   ' 'Grid.ExtendLastCol = True
   ' Grid.WallPaper = Nothing
   ' 'Grid.AutoSize  0, Grid.Cols - 1, False
   ' Printer.Orientation = VBRUN.PrinterObjectConstants.vbPRORLandscape
 
   ' 'Printer.RightToLeft = True
   ' 'Printer.Print ("Employee Salary Report")

   ' Me.Grid.PrintGrid " ‰»Ì…    „” Œ·’«  ·„  ”œœ »«·ﬂ«„·", True, 2, 1, 1500
   print_report My_SQL
End Sub

Function print_report(Optional NoteSerial As String)
    
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String

        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rpt_RentAlert.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rpt_RentAlert.rpt"
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

    If SystemOptions.UserInterface = ArabicInterface Then
        xReport.ParameterFields(2).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
        
        ' xReport.ParameterFields(2).AddCurrentValue "test"
        ' xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
        StrReportTitle = "" '& StrAccountName
        'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        '    StrReportTitle = StrReportTitle + " »œ«Ì… „‰ " & Format(Me.DTPickerAccFrom.value, "yyyy/M/d") & ""
        'End If
        'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        '    StrReportTitle = StrReportTitle + " ≈·Ï " & Format(Me.DTPickerAccTo.value, "yyyy/M/d") & " "
        'End If
      '    xReport.ParameterFields(14).AddCurrentValue Format(Me.Fromdate.value, "yyyy/M/d")
      '  xReport.ParameterFields(15).AddCurrentValue Format(Me.todate.value, "yyyy/M/d")
      ' xReport.ParameterFields(16).AddCurrentValue Me.Fromdate√H.value
      '  xReport.ParameterFields(17).AddCurrentValue Me.todateH.value
    Else
 
        xReport.ParameterFields(2).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        'xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
       ' xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
       ' StrReportTitle = ""
        'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        '    StrReportTitle = StrReportTitle + " From Date " & (Me.DTPickerAccFrom.value) & ""
        'End If
        'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        '    StrReportTitle = StrReportTitle + " To Date :  " & (Me.DTPickerAccTo.value) & ""
        'End If
    End If

   ' xReport.ParameterFields(3).AddCurrentValue user_name
       ' xReport.ParameterFields(4).AddCurrentValue WriteNo(Format(val(TxtAdvanceValue.text), "0.00"), 0, True, ".")
       ' xReport.ParameterFields(6).AddCurrentValue val(lbl(23).Caption)
         xReport.ParameterFields(3).AddCurrentValue Fromdate.value
          xReport.ParameterFields(4).AddCurrentValue FromDateH.value
          xReport.ParameterFields(5).AddCurrentValue ToDate.value
          xReport.ParameterFields(6).AddCurrentValue todateH.value
   ' xReport.ParameterFields(8).AddCurrentValue IIf(IsNumeric(Fg.TextMatrix(Me.Fg.FixedRows, Fg.ColIndex("PartValue"))), val(Fg.TextMatrix(Me.Fg.FixedRows, Fg.ColIndex("PartValue"))), 0)
' xReport.ParameterFields(9).AddCurrentValue val(lbl(22).Caption)
 ' xReport.ParameterFields(10).AddCurrentValue val(TxtDiscount.text)
'
   
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



Private Sub Command1_Click()
    Dim Numbers As String
    Dim RowNum As Integer
    Dim Opt As Integer
    Dim CurrentMessage As String
    Numbers = ""

    With GridInstallments

        For RowNum = .FixedRows To .Rows - 1
    
            If .Cell(flexcpChecked, RowNum, .ColIndex("Send")) = flexChecked Then

                '  MsgBox (.TextMatrix(RowNum, .ColIndex("Numbers")))
                If (.TextMatrix(RowNum, .ColIndex("Cus_mobile"))) <> "" Then
                    If Numbers = "" Then
                        Numbers = (.TextMatrix(RowNum, .ColIndex("Cus_mobile")))
                    Else
                        Numbers = Numbers & "," & (.TextMatrix(RowNum, .ColIndex("Cus_mobile")))
                    End If
             
                End If
            End If
          
        Next RowNum
      
        CurrentMessage = ComposMessage(Me.Name)  ', 0, "", Me.TXTMessageDES.text, Opt)

        If Numbers = "" Then Exit Sub
        SMSSeTTings.SendMessage CurrentMessage, Numbers
        SMSSeTTings.Hide
                                    
    End With

End Sub

Private Sub Form_Load()


 
    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500
     Fromdate.value = Date
      ToDate.value = Date
      txtFromDate = Date
      txtToDate = Date
    Dim Dcombos As ClsDataCombos
        Set Dcombos = New ClsDataCombos
    Dcombos.GetCustomersSuppliers 0, Me.DBCboClientName, True
    Dcombos.GetStores Me.DCboStoreName
    Dcombos.GetStores Me.DCboStoreName2
    Dcombos.GetBranches Me.dcBranch
    

  s = "Select StoreID,StoreID1,StoreID2,StoreID3 from tblUsers Where UserID = " & user_id
  Set rsDummy = New ADODB.Recordset

rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly, adCmdText
If Not rsDummy.EOF Then
    DCboStoreName2.BoundText = val(rsDummy!StoreId2 & "")
    DCboStoreName.BoundText = val(rsDummy!StoreID3 & "")
End If

    cmdCreateProduction.Enabled = False
    If mIndex = 1 Then
        If SystemOptions.UserInterface = EnglishInterface Then
            Me.Caption = "Internal request alerts"
        Else
            Me.Caption = " ‰»ÌÂ«  «·ÿ·»«  «·œ«Œ·Ì…"
        End If
        TabMain.TabVisible(0) = False
        TabMain.TabVisible(1) = True
        TabMain.CurrTab = 0
    Else
         TabMain.TabVisible(0) = True
        TabMain.TabVisible(1) = False
        TabMain.CurrTab = 1
    End If
    If SystemOptions.UserInterface = EnglishInterface Then

        SetInterface Me
        cahngelang
    End If
Cmd_Click 0
End Sub

Function cahngelang()
    Label1(2).Caption = "Project Invoices Not Payed"
    Me.Caption = Label1(2).Caption
    Frame1.Caption = "Color Map"
    Label3.Caption = "Fully"
    Label5.Caption = "Partial"

    Me.Caption = Label1(2).Caption
    CmdPrint.Caption = "Print"
    btnCancel.Caption = "Cancel"

    With Me.Grid
        .TextMatrix(0, .ColIndex("Ser")) = "Ser"
        .TextMatrix(0, .ColIndex("id")) = " Bill ID"
        .TextMatrix(0, .ColIndex("bill_date")) = "Bill Date  "
        .TextMatrix(0, .ColIndex("Project_name")) = "Project Name"
        .TextMatrix(0, .ColIndex("End_user_name")) = "End_user_name"
        .TextMatrix(0, .ColIndex("Sub_user_name")) = "Sub_user_name"
        .TextMatrix(0, .ColIndex("total")) = "Bill Total"
        .TextMatrix(0, .ColIndex("ActualTotal")) = "Payed"
        .TextMatrix(0, .ColIndex("result")) = "Variance"
        .TextMatrix(0, .ColIndex("resultpercentage")) = "Variance%"

    End With
Frame8(1).Caption = "Select the perid"
lbl(2).Caption = "Period of"
lbl(1).Caption = "To"
lbl(36).Caption = "Branch"
lbl(42).Caption = "Customer"
lbl(33).Caption = "Store raw materials"
lbl(34).Caption = "Store full production"
Cmd(1).Caption = "Add"
cmdCreateProduction.Caption = "Create an output command"

    With Me.Fg
        .TextMatrix(0, .ColIndex("Ser")) = "Ser"
        .TextMatrix(0, .ColIndex("GroupID")) = "Group ID"
        .TextMatrix(0, .ColIndex("GroupName")) = "Group Name"
        .TextMatrix(0, .ColIndex("Item_ID")) = "Item Code"
        .TextMatrix(0, .ColIndex("ItemName")) = "Item Name"
        .TextMatrix(0, .ColIndex("Qty")) = "Qty"
        .TextMatrix(0, .ColIndex("Price")) = "Price"
        .TextMatrix(0, .ColIndex("Total")) = "Total"


    End With
End Function

Private Sub FromDate_Change()
'If Fromdate.value <> Null Then
 FromDateH.value = ToHijriDate(Fromdate.value)
 'End If
End Sub

Private Sub Fromdateh_LostFocus()
Fromdate.value = ToGregorianDate(FromDateH.value)
End Sub



Private Sub ImgFavorites_Click()
AddTofaforites Me.Name, Me.Caption, Me.Caption
End Sub

Private Sub ToDate_Change()
   ' If toDate.value <> Null Then
    todateH.value = ToHijriDate(ToDate.value)
   ' End If
End Sub

Private Sub ToDateH_LostFocus()
 ToDate.value = ToGregorianDate(todateH.value)
End Sub


 Private Sub CalcCostPercent()
    Dim i As Long
    Dim mCostPercent As Double
    Dim mCostTotal As Double
    If Fg.Rows = 1 Then Exit Sub
    mCostTotal = Fg.Aggregate(flexSTSum, Fg.FixedRows, Fg.ColIndex("Total"), Fg.Rows - 1, Fg.ColIndex("Total"))
    For i = 1 To Fg.Rows - 1
        If mCostTotal <> 0 Then
            Fg.TextMatrix(i, Fg.ColIndex("PercentCost")) = val(Fg.TextMatrix(i, Fg.ColIndex("Total"))) / mCostTotal * 100
        End If
        
    Next
    
 End Sub


