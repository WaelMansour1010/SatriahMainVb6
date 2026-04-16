VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmSearch_Hajj 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   9525
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13200
   Icon            =   "FrmSearch_Hajj.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   9525
   ScaleWidth      =   13200
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
   Begin VB.Frame frm_VehicleOperatorOrder 
      BackColor       =   &H00E2E9E9&
      Height          =   8775
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   153
      Top             =   -120
      Visible         =   0   'False
      Width           =   13212
      Begin VB.Frame Frame4 
         BackColor       =   &H00E2E9E9&
         Height          =   1812
         Left            =   -1080
         RightToLeft     =   -1  'True
         TabIndex        =   154
         Top             =   6720
         Width           =   14292
         Begin VB.TextBox ID2 
            Alignment       =   1  'Right Justify
            Height          =   288
            Left            =   10560
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   158
            Top             =   240
            Width           =   1980
         End
         Begin VB.TextBox EmpName2 
            Alignment       =   1  'Right Justify
            Height          =   288
            Left            =   1200
            RightToLeft     =   -1  'True
            TabIndex        =   157
            Top             =   240
            Width           =   1776
         End
         Begin VB.TextBox EmpCode2 
            Alignment       =   1  'Right Justify
            Height          =   288
            Left            =   4200
            RightToLeft     =   -1  'True
            TabIndex        =   156
            Top             =   240
            Width           =   1812
         End
         Begin VB.TextBox FlightNo2 
            Alignment       =   1  'Right Justify
            Height          =   288
            Left            =   1200
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   155
            Top             =   600
            Width           =   1776
         End
         Begin MSDataListLib.DataCombo BranchID2 
            Height          =   288
            Left            =   7080
            TabIndex        =   159
            Top             =   240
            Width           =   2004
            _ExtentX        =   3545
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo InClientID2 
            Height          =   288
            Left            =   10560
            TabIndex        =   160
            Top             =   600
            Width           =   2016
            _ExtentX        =   3545
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo GroupID2 
            Height          =   288
            Left            =   10560
            TabIndex        =   161
            Top             =   960
            Width           =   2016
            _ExtentX        =   3545
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo CompanyID2 
            Height          =   288
            Left            =   7080
            TabIndex        =   162
            Top             =   960
            Width           =   2016
            _ExtentX        =   3545
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo AirPortID2 
            Height          =   288
            Left            =   4200
            TabIndex        =   163
            Top             =   600
            Width           =   1776
            _ExtentX        =   3122
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo AirLineID2 
            Height          =   288
            Left            =   4200
            TabIndex        =   164
            Top             =   960
            Width           =   1776
            _ExtentX        =   3122
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo OutClientID2 
            Height          =   288
            Left            =   7080
            TabIndex        =   165
            Top             =   600
            Width           =   2016
            _ExtentX        =   3545
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo ProgrammID2 
            Height          =   288
            Left            =   1200
            TabIndex        =   166
            Top             =   1320
            Width           =   1776
            _ExtentX        =   3122
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo MekkaHotelID2 
            Height          =   288
            Left            =   10560
            TabIndex        =   167
            Top             =   1320
            Width           =   2016
            _ExtentX        =   3545
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo MadinaHotelID2 
            Height          =   288
            Left            =   4200
            TabIndex        =   168
            Top             =   1320
            Width           =   1776
            _ExtentX        =   3122
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo JeddahHotelID2 
            Height          =   288
            Left            =   7080
            TabIndex        =   169
            Top             =   1320
            Width           =   2016
            _ExtentX        =   3545
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo VehicleType2 
            Height          =   288
            Left            =   1200
            TabIndex        =   170
            Top             =   960
            Width           =   1776
            _ExtentX        =   3122
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·ŒÿÊÿ «·ÃÊÌ…"
            Height          =   312
            Index           =   27
            Left            =   5832
            RightToLeft     =   -1  'True
            TabIndex        =   186
            Top             =   960
            Width           =   1068
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„ÿ«—"
            Height          =   312
            Index           =   26
            Left            =   5712
            RightToLeft     =   -1  'True
            TabIndex        =   185
            Top             =   600
            Width           =   1188
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·‘—ﬂ…"
            Height          =   312
            Index           =   25
            Left            =   9312
            RightToLeft     =   -1  'True
            TabIndex        =   184
            Top             =   960
            Width           =   1068
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·„Ã„Ê⁄…"
            Height          =   312
            Index           =   23
            Left            =   12912
            RightToLeft     =   -1  'True
            TabIndex        =   183
            Top             =   960
            Width           =   1068
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·‘—ﬂ… „‰ «·Œ«—Ã"
            Height          =   312
            Index           =   22
            Left            =   9192
            RightToLeft     =   -1  'True
            TabIndex        =   182
            Top             =   600
            Width           =   1188
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·‘—ﬂ… «·”⁄ÊœÌ…"
            Height          =   312
            Index           =   21
            Left            =   12792
            RightToLeft     =   -1  'True
            TabIndex        =   181
            Top             =   600
            Width           =   1188
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·›—⁄"
            Height          =   312
            Index           =   20
            Left            =   9876
            RightToLeft     =   -1  'True
            TabIndex        =   180
            Top             =   240
            Width           =   528
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„”·”·"
            Height          =   336
            Index           =   19
            Left            =   12780
            RightToLeft     =   -1  'True
            TabIndex        =   179
            Top             =   240
            Width           =   1188
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·„‘—›"
            Height          =   252
            Left            =   3000
            RightToLeft     =   -1  'True
            TabIndex        =   178
            Top             =   240
            Width           =   972
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬂÊœ «·„‘—›"
            Height          =   252
            Index           =   0
            Left            =   6120
            RightToLeft     =   -1  'True
            TabIndex        =   177
            Top             =   240
            Width           =   732
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·Õ«›·« "
            Height          =   312
            Index           =   14
            Left            =   2832
            RightToLeft     =   -1  'True
            TabIndex        =   176
            Top             =   960
            Width           =   1188
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "›‰œﬁ «·„œÌ‰…"
            Height          =   312
            Index           =   13
            Left            =   5712
            RightToLeft     =   -1  'True
            TabIndex        =   175
            Top             =   1320
            Width           =   1188
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "›‰œﬁ Ãœ…"
            Height          =   312
            Index           =   12
            Left            =   9192
            RightToLeft     =   -1  'True
            TabIndex        =   174
            Top             =   1320
            Width           =   1188
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "›‰œﬁ ›Ï „ﬂ…"
            Height          =   312
            Index           =   10
            Left            =   12792
            RightToLeft     =   -1  'True
            TabIndex        =   173
            Top             =   1320
            Width           =   1188
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·»—‰«„Ã"
            Height          =   312
            Index           =   8
            Left            =   2952
            RightToLeft     =   -1  'True
            TabIndex        =   172
            Top             =   1320
            Width           =   1068
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·—Õ·…"
            Height          =   336
            Index           =   5
            Left            =   2820
            RightToLeft     =   -1  'True
            TabIndex        =   171
            Top             =   600
            Width           =   1188
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   588
         Left            =   0
         TabIndex        =   187
         TabStop         =   0   'False
         Top             =   120
         Width           =   13176
         _cx             =   23230
         _cy             =   1032
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
         Caption         =   "       √„—  ‘€Ì· Õ«›·…   "
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
         Begin VB.TextBox Text5 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   12960
            RightToLeft     =   -1  'True
            TabIndex        =   188
            Top             =   1320
            Visible         =   0   'False
            Width           =   855
         End
      End
      Begin VSFlex8UCtl.VSFlexGrid fg_VehicleOperatorOrder 
         Height          =   5760
         Left            =   120
         TabIndex        =   189
         Top             =   720
         Width           =   12975
         _cx             =   22886
         _cy             =   10160
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
         Rows            =   50
         Cols            =   19
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FrmSearch_Hajj.frx":038A
         ScrollTrack     =   -1  'True
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
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00E2E9E9&
      Height          =   8415
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   123
      Top             =   0
      Visible         =   0   'False
      Width           =   13212
      Begin VB.Frame Frame10 
         BackColor       =   &H00E2E9E9&
         Height          =   2295
         Left            =   -1080
         RightToLeft     =   -1  'True
         TabIndex        =   124
         Top             =   6240
         Width           =   14292
         Begin VB.Frame Frame14 
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·Õ—ﬂÂ"
            Height          =   645
            Left            =   10080
            RightToLeft     =   -1  'True
            TabIndex        =   129
            Top             =   120
            Width           =   4155
            Begin VB.TextBox TxtFromID 
               Alignment       =   1  'Right Justify
               Height          =   345
               Left            =   1920
               RightToLeft     =   -1  'True
               TabIndex        =   131
               Top             =   240
               Width           =   915
            End
            Begin VB.TextBox TxtToID 
               Alignment       =   1  'Right Justify
               Height          =   345
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   130
               Top             =   240
               Width           =   915
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„‰"
               Height          =   195
               Index           =   67
               Left            =   2775
               RightToLeft     =   -1  'True
               TabIndex        =   133
               Top             =   240
               Width           =   540
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "≈·Ï"
               Height          =   195
               Index           =   66
               Left            =   1260
               RightToLeft     =   -1  'True
               TabIndex        =   132
               Top             =   240
               Width           =   525
            End
         End
         Begin VB.Frame Frame13 
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «·Õ—ﬂ…"
            Height          =   1395
            Left            =   1080
            RightToLeft     =   -1  'True
            TabIndex        =   125
            Top             =   120
            Width           =   4695
            Begin MSComCtl2.DTPicker ToSDate 
               Height          =   330
               Left            =   1920
               TabIndex        =   126
               Top             =   720
               Width           =   1590
               _ExtentX        =   2805
               _ExtentY        =   582
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   66256899
               CurrentDate     =   38887
            End
            Begin MSComCtl2.DTPicker SDate 
               Height          =   330
               Left            =   1920
               TabIndex        =   145
               TabStop         =   0   'False
               Top             =   360
               Width           =   1590
               _ExtentX        =   2805
               _ExtentY        =   582
               _Version        =   393216
               CalendarBackColor=   12648447
               CalendarTitleBackColor=   10383715
               CheckBox        =   -1  'True
               Format          =   66256899
               CurrentDate     =   37140
            End
            Begin Dynamic_Byte.NourHijriCal SDateH 
               Height          =   315
               Left            =   240
               TabIndex        =   146
               Top             =   360
               Width           =   1515
               _ExtentX        =   2778
               _ExtentY        =   556
            End
            Begin Dynamic_Byte.NourHijriCal ToSDateH 
               Height          =   315
               Left            =   240
               TabIndex        =   147
               Top             =   720
               Width           =   1515
               _ExtentX        =   2778
               _ExtentY        =   556
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "≈·Ï"
               Height          =   195
               Index           =   65
               Left            =   3735
               RightToLeft     =   -1  'True
               TabIndex        =   128
               Top             =   660
               Width           =   480
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„‰"
               Height          =   195
               Index           =   64
               Left            =   3720
               RightToLeft     =   -1  'True
               TabIndex        =   127
               Top             =   330
               Width           =   540
            End
         End
         Begin MSDataListLib.DataCombo DcbBranch4 
            Height          =   315
            Left            =   5880
            TabIndex        =   134
            Top             =   360
            Width           =   3075
            _ExtentX        =   5424
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcbCars 
            Height          =   315
            Left            =   5880
            TabIndex        =   135
            Top             =   960
            Width           =   7275
            _ExtentX        =   12832
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcbTypeCar 
            Height          =   315
            Left            =   5880
            TabIndex        =   136
            Top             =   1320
            Width           =   7275
            _ExtentX        =   12832
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcbDriver 
            Height          =   315
            Left            =   5880
            TabIndex        =   137
            Top             =   1680
            Width           =   7275
            _ExtentX        =   12832
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·›—⁄"
            Height          =   195
            Index           =   76
            Left            =   9240
            RightToLeft     =   -1  'True
            TabIndex        =   141
            Top             =   360
            Width           =   525
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·”«∆ﬁ"
            Height          =   195
            Index           =   74
            Left            =   13200
            RightToLeft     =   -1  'True
            TabIndex        =   140
            Top             =   1680
            Width           =   1005
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·Õ«›·…"
            Height          =   195
            Index           =   73
            Left            =   13200
            RightToLeft     =   -1  'True
            TabIndex        =   139
            Top             =   960
            Width           =   1005
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·Õ«›·…"
            Height          =   195
            Index           =   71
            Left            =   13200
            RightToLeft     =   -1  'True
            TabIndex        =   138
            Top             =   1320
            Width           =   1005
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic5 
         Height          =   585
         Left            =   0
         TabIndex        =   142
         TabStop         =   0   'False
         Top             =   0
         Width           =   13170
         _cx             =   23230
         _cy             =   1032
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
         Caption         =   "       Ê“Ì⁄ Õ«›·«  «·„‘«⁄—"
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
         Begin VB.TextBox Text9 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   12960
            RightToLeft     =   -1  'True
            TabIndex        =   143
            Top             =   1320
            Visible         =   0   'False
            Width           =   855
         End
      End
      Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
         Height          =   5400
         Left            =   120
         TabIndex        =   144
         Top             =   720
         Width           =   12975
         _cx             =   22886
         _cy             =   9525
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
         Rows            =   50
         Cols            =   9
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FrmSearch_Hajj.frx":0683
         ScrollTrack     =   -1  'True
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
   End
   Begin VB.Frame Fram_Deported 
      BackColor       =   &H00E2E9E9&
      Height          =   8775
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   78
      Top             =   -120
      Visible         =   0   'False
      Width           =   13212
      Begin VB.Frame Frame7 
         BackColor       =   &H00E2E9E9&
         Height          =   2295
         Left            =   -1080
         RightToLeft     =   -1  'True
         TabIndex        =   79
         Top             =   6120
         Width           =   14292
         Begin VB.TextBox TxtLeader 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   10080
            RightToLeft     =   -1  'True
            TabIndex        =   106
            Top             =   840
            Width           =   3075
         End
         Begin VB.Frame Frame8 
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «·Ê’Ê·"
            Height          =   675
            Left            =   1080
            RightToLeft     =   -1  'True
            TabIndex        =   99
            Top             =   1560
            Width           =   4695
            Begin MSComCtl2.DTPicker FrmDateArrive 
               Height          =   330
               Left            =   2400
               TabIndex        =   100
               Top             =   270
               Width           =   1590
               _ExtentX        =   2805
               _ExtentY        =   582
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   66256899
               CurrentDate     =   38887
            End
            Begin MSComCtl2.DTPicker TODateArrive 
               Height          =   330
               Left            =   90
               TabIndex        =   101
               Top             =   270
               Width           =   1590
               _ExtentX        =   2805
               _ExtentY        =   582
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   66256899
               CurrentDate     =   38887
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "≈·Ï"
               Height          =   195
               Index           =   49
               Left            =   1695
               RightToLeft     =   -1  'True
               TabIndex        =   103
               Top             =   300
               Width           =   480
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„‰"
               Height          =   195
               Index           =   47
               Left            =   3960
               RightToLeft     =   -1  'True
               TabIndex        =   102
               Top             =   330
               Width           =   540
            End
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «·„€«œ—…"
            Height          =   675
            Left            =   1080
            RightToLeft     =   -1  'True
            TabIndex        =   94
            Top             =   840
            Width           =   4695
            Begin MSComCtl2.DTPicker FrmDateGO 
               Height          =   330
               Left            =   2400
               TabIndex        =   95
               Top             =   270
               Width           =   1590
               _ExtentX        =   2805
               _ExtentY        =   582
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   66256899
               CurrentDate     =   38887
            End
            Begin MSComCtl2.DTPicker ToDateGO 
               Height          =   330
               Left            =   90
               TabIndex        =   96
               Top             =   270
               Width           =   1590
               _ExtentX        =   2805
               _ExtentY        =   582
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   66256899
               CurrentDate     =   38887
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "≈·Ï"
               Height          =   195
               Index           =   46
               Left            =   1695
               RightToLeft     =   -1  'True
               TabIndex        =   98
               Top             =   300
               Width           =   480
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„‰"
               Height          =   195
               Index           =   45
               Left            =   3960
               RightToLeft     =   -1  'True
               TabIndex        =   97
               Top             =   330
               Width           =   540
            End
         End
         Begin VB.Frame lbreg 
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «·Õ—ﬂ…"
            Height          =   675
            Left            =   1080
            RightToLeft     =   -1  'True
            TabIndex        =   89
            Top             =   120
            Width           =   4695
            Begin MSComCtl2.DTPicker DtpDateFrom 
               Height          =   330
               Left            =   2400
               TabIndex        =   90
               Top             =   270
               Width           =   1590
               _ExtentX        =   2805
               _ExtentY        =   582
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   66256899
               CurrentDate     =   38887
            End
            Begin MSComCtl2.DTPicker DtpDateTo 
               Height          =   330
               Left            =   90
               TabIndex        =   91
               Top             =   240
               Width           =   1590
               _ExtentX        =   2805
               _ExtentY        =   582
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   66256899
               CurrentDate     =   38887
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„‰"
               Height          =   195
               Index           =   57
               Left            =   3960
               RightToLeft     =   -1  'True
               TabIndex        =   93
               Top             =   330
               Width           =   540
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "≈·Ï"
               Height          =   195
               Index           =   56
               Left            =   1695
               RightToLeft     =   -1  'True
               TabIndex        =   92
               Top             =   300
               Width           =   480
            End
         End
         Begin VB.Frame lbprocess 
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·Õ—ﬂÂ"
            Height          =   645
            Left            =   10080
            RightToLeft     =   -1  'True
            TabIndex        =   84
            Top             =   120
            Width           =   4155
            Begin VB.TextBox TxtIDTO 
               Alignment       =   1  'Right Justify
               Height          =   345
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   86
               Top             =   240
               Width           =   915
            End
            Begin VB.TextBox TxtIDFrom 
               Alignment       =   1  'Right Justify
               Height          =   345
               Left            =   1920
               RightToLeft     =   -1  'True
               TabIndex        =   85
               Top             =   240
               Width           =   915
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "≈·Ï"
               Height          =   195
               Index           =   55
               Left            =   1260
               RightToLeft     =   -1  'True
               TabIndex        =   88
               Top             =   240
               Width           =   525
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„‰"
               Height          =   195
               Index           =   48
               Left            =   2775
               RightToLeft     =   -1  'True
               TabIndex        =   87
               Top             =   240
               Width           =   540
            End
         End
         Begin VB.TextBox TxtNoFrom 
            Alignment       =   1  'Right Justify
            Height          =   288
            Left            =   10560
            RightToLeft     =   -1  'True
            TabIndex        =   80
            Top             =   240
            Width           =   1980
         End
         Begin MSDataListLib.DataCombo DcbBrnch 
            Height          =   315
            Left            =   5880
            TabIndex        =   104
            Top             =   360
            Width           =   3075
            _ExtentX        =   5424
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcbEqupID 
            Height          =   315
            Left            =   10080
            TabIndex        =   109
            Top             =   1200
            Width           =   3075
            _ExtentX        =   5424
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcbTypeTrip 
            Height          =   315
            Left            =   5880
            TabIndex        =   111
            Top             =   1200
            Width           =   3075
            _ExtentX        =   5424
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcbProgramID 
            Height          =   315
            Left            =   10080
            TabIndex        =   113
            Top             =   1560
            Width           =   3075
            _ExtentX        =   5424
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcbPath 
            Height          =   315
            Left            =   5880
            TabIndex        =   115
            Top             =   1560
            Width           =   3075
            _ExtentX        =   5424
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcbLocatioID 
            Height          =   315
            Left            =   10080
            TabIndex        =   117
            Top             =   1920
            Width           =   3075
            _ExtentX        =   5424
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcbLocatioID2 
            Height          =   315
            Left            =   5880
            TabIndex        =   119
            Top             =   1920
            Width           =   3075
            _ExtentX        =   5424
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcbEmp 
            Height          =   315
            Left            =   5880
            TabIndex        =   121
            Top             =   840
            Width           =   3075
            _ExtentX        =   5424
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„Êﬁ⁄ «·Ê’Ê·"
            Height          =   195
            Index           =   61
            Left            =   9000
            RightToLeft     =   -1  'True
            TabIndex        =   120
            Top             =   1920
            Width           =   1005
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„Êﬁ⁄ «·„€«œ—…"
            Height          =   195
            Index           =   60
            Left            =   13200
            RightToLeft     =   -1  'True
            TabIndex        =   118
            Top             =   1920
            Width           =   1005
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "« Ã«Â «·—Õ·…"
            Height          =   195
            Index           =   59
            Left            =   9000
            RightToLeft     =   -1  'True
            TabIndex        =   116
            Top             =   1560
            Width           =   1005
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·»—‰«„Ã"
            Height          =   195
            Index           =   58
            Left            =   13200
            RightToLeft     =   -1  'True
            TabIndex        =   114
            Top             =   1560
            Width           =   1005
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·—Õ·…"
            Height          =   195
            Index           =   54
            Left            =   9000
            RightToLeft     =   -1  'True
            TabIndex        =   112
            Top             =   1200
            Width           =   1005
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·Õ«›·…"
            Height          =   195
            Index           =   53
            Left            =   13200
            RightToLeft     =   -1  'True
            TabIndex        =   110
            Top             =   1200
            Width           =   1005
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·„‘—›"
            Height          =   195
            Index           =   52
            Left            =   9000
            RightToLeft     =   -1  'True
            TabIndex        =   108
            Top             =   840
            Width           =   1005
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬁ«∆œ «·Õ«ﬁ·…"
            Height          =   195
            Index           =   51
            Left            =   13200
            RightToLeft     =   -1  'True
            TabIndex        =   107
            Top             =   840
            Width           =   1005
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·›—⁄"
            Height          =   195
            Index           =   50
            Left            =   9240
            RightToLeft     =   -1  'True
            TabIndex        =   105
            Top             =   360
            Width           =   525
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic4 
         Height          =   555
         Left            =   0
         TabIndex        =   81
         TabStop         =   0   'False
         Top             =   120
         Width           =   13170
         _cx             =   23230
         _cy             =   979
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
         Caption         =   "       »ÕÀ ÃœÊ· «· —ÕÌ·"
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
         Begin VB.TextBox Text4 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   12960
            RightToLeft     =   -1  'True
            TabIndex        =   82
            Top             =   1320
            Visible         =   0   'False
            Width           =   855
         End
      End
      Begin VSFlex8UCtl.VSFlexGrid FgDeported 
         Height          =   5520
         Left            =   0
         TabIndex        =   83
         Top             =   720
         Width           =   13095
         _cx             =   23098
         _cy             =   9737
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
         Rows            =   50
         Cols            =   15
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FrmSearch_Hajj.frx":07C9
         ScrollTrack     =   -1  'True
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
   End
   Begin VB.Frame frm_EndorseTrans 
      BackColor       =   &H00E2E9E9&
      Height          =   8655
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   57
      Top             =   0
      Visible         =   0   'False
      Width           =   13095
      Begin C1SizerLibCtl.C1Elastic C1Elastic3 
         Height          =   585
         Left            =   0
         TabIndex        =   68
         TabStop         =   0   'False
         Top             =   0
         Width           =   13170
         _cx             =   23230
         _cy             =   1032
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
         Caption         =   "       √⁄ „«œ «—ﬂ«» «·ÕÃ«Ã"
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
         Begin VB.TextBox Text8 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   12960
            RightToLeft     =   -1  'True
            TabIndex        =   69
            Top             =   1320
            Visible         =   0   'False
            Width           =   855
         End
      End
      Begin VSFlex8UCtl.VSFlexGrid fg_EndorseTrans 
         Height          =   5760
         Left            =   120
         TabIndex        =   70
         Top             =   600
         Width           =   12975
         _cx             =   22886
         _cy             =   10160
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
         Rows            =   50
         Cols            =   15
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FrmSearch_Hajj.frx":0A16
         ScrollTrack     =   -1  'True
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
      Begin VB.Frame Frame6 
         BackColor       =   &H00E2E9E9&
         Height          =   1812
         Left            =   -1080
         RightToLeft     =   -1  'True
         TabIndex        =   58
         Top             =   6600
         Width           =   14292
         Begin VB.TextBox ApproveID 
            Alignment       =   1  'Right Justify
            Height          =   288
            Left            =   1200
            RightToLeft     =   -1  'True
            TabIndex        =   71
            Top             =   240
            Width           =   1764
         End
         Begin VB.TextBox ID4 
            Alignment       =   1  'Right Justify
            Height          =   288
            Left            =   10560
            RightToLeft     =   -1  'True
            TabIndex        =   60
            Top             =   240
            Width           =   1980
         End
         Begin VB.TextBox Remark4 
            Alignment       =   1  'Right Justify
            Height          =   525
            Left            =   1200
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   59
            Top             =   960
            Width           =   11376
         End
         Begin MSDataListLib.DataCombo BranchID4 
            Height          =   315
            Left            =   4200
            TabIndex        =   61
            Top             =   240
            Width           =   5370
            _ExtentX        =   9472
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DataCombo2 
            Height          =   315
            Left            =   1200
            TabIndex        =   62
            Top             =   600
            Width           =   4770
            _ExtentX        =   8414
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo CompanyID4 
            Height          =   315
            Left            =   1200
            TabIndex        =   73
            Top             =   960
            Width           =   4770
            _ExtentX        =   8414
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo Nationality 
            Height          =   315
            Left            =   7080
            TabIndex        =   74
            Top             =   960
            Width           =   5490
            _ExtentX        =   9684
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcbPath2 
            Height          =   315
            Left            =   7080
            TabIndex        =   122
            Top             =   600
            Width           =   5490
            _ExtentX        =   9684
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„ƒ””…"
            Height          =   330
            Index           =   43
            Left            =   5850
            RightToLeft     =   -1  'True
            TabIndex        =   76
            Top             =   960
            Width           =   1080
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·Ã‰”Ì…"
            Height          =   408
            Index           =   36
            Left            =   12912
            RightToLeft     =   -1  'True
            TabIndex        =   75
            Top             =   960
            Width           =   1068
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·«⁄ „«œ"
            Height          =   324
            Index           =   42
            Left            =   2664
            RightToLeft     =   -1  'True
            TabIndex        =   72
            Top             =   240
            Width           =   1320
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„”«—"
            Height          =   312
            Index           =   40
            Left            =   12792
            RightToLeft     =   -1  'True
            TabIndex        =   67
            Top             =   600
            Width           =   1188
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„Ê”„"
            Height          =   315
            Index           =   39
            Left            =   5715
            RightToLeft     =   -1  'True
            TabIndex        =   66
            Top             =   600
            Width           =   1185
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·›—⁄"
            Height          =   312
            Index           =   38
            Left            =   9756
            RightToLeft     =   -1  'True
            TabIndex        =   65
            Top             =   240
            Width           =   528
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„”·”·"
            Height          =   336
            Index           =   37
            Left            =   12780
            RightToLeft     =   -1  'True
            TabIndex        =   64
            Top             =   240
            Width           =   1188
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„·«ÕŸ« "
            Height          =   336
            Index           =   30
            Left            =   12780
            RightToLeft     =   -1  'True
            TabIndex        =   63
            Top             =   1320
            Width           =   1188
         End
      End
   End
   Begin VB.Frame frm_Evacation 
      BackColor       =   &H00E2E9E9&
      Height          =   8655
      Left            =   -120
      RightToLeft     =   -1  'True
      TabIndex        =   34
      Top             =   0
      Visible         =   0   'False
      Width           =   13335
      Begin VB.Frame Frame5 
         BackColor       =   &H00E2E9E9&
         Height          =   1812
         Left            =   -1080
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   6720
         Width           =   14292
         Begin VB.TextBox Remark3 
            Alignment       =   1  'Right Justify
            Height          =   648
            Left            =   1200
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   55
            Top             =   960
            Width           =   11376
         End
         Begin VB.TextBox Behavior3 
            Alignment       =   1  'Right Justify
            Height          =   288
            Left            =   4200
            RightToLeft     =   -1  'True
            TabIndex        =   53
            Top             =   600
            Width           =   1776
         End
         Begin VB.TextBox Discipline3 
            Alignment       =   1  'Right Justify
            Height          =   288
            Left            =   1200
            RightToLeft     =   -1  'True
            TabIndex        =   38
            Top             =   600
            Width           =   1776
         End
         Begin VB.TextBox EmployeeCode3 
            Alignment       =   1  'Right Justify
            Height          =   288
            Left            =   4200
            RightToLeft     =   -1  'True
            TabIndex        =   37
            Top             =   240
            Width           =   1812
         End
         Begin VB.TextBox ID3 
            Alignment       =   1  'Right Justify
            Height          =   288
            Left            =   10560
            RightToLeft     =   -1  'True
            TabIndex        =   36
            Top             =   240
            Width           =   1980
         End
         Begin MSDataListLib.DataCombo BranchID3 
            Height          =   288
            Left            =   7080
            TabIndex        =   39
            Top             =   240
            Width           =   2004
            _ExtentX        =   3545
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo SeasonsID3 
            Height          =   288
            Left            =   10560
            TabIndex        =   40
            Top             =   600
            Width           =   2016
            _ExtentX        =   3545
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo CarID3 
            Height          =   288
            Left            =   7080
            TabIndex        =   41
            Top             =   600
            Width           =   2016
            _ExtentX        =   3545
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo EmployeeID3 
            Height          =   288
            Left            =   1200
            TabIndex        =   42
            Top             =   240
            Width           =   1776
            _ExtentX        =   3122
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„·«ÕŸ« "
            Height          =   336
            Index           =   29
            Left            =   12780
            RightToLeft     =   -1  'True
            TabIndex        =   56
            Top             =   960
            Width           =   1188
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·«Œ·«ﬁ Ê«·”·Êﬂ"
            Height          =   336
            Index           =   28
            Left            =   5820
            RightToLeft     =   -1  'True
            TabIndex        =   54
            Top             =   600
            Width           =   1188
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„Ê«Ÿ»…"
            Height          =   336
            Index           =   41
            Left            =   2820
            RightToLeft     =   -1  'True
            TabIndex        =   49
            Top             =   600
            Width           =   1188
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬂÊœ «·„ÊŸ›"
            Height          =   252
            Index           =   2
            Left            =   6240
            RightToLeft     =   -1  'True
            TabIndex        =   48
            Top             =   240
            Width           =   732
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·„ÊŸ›"
            Height          =   252
            Left            =   3000
            RightToLeft     =   -1  'True
            TabIndex        =   47
            Top             =   240
            Width           =   972
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„”·”·"
            Height          =   336
            Index           =   35
            Left            =   12780
            RightToLeft     =   -1  'True
            TabIndex        =   46
            Top             =   240
            Width           =   1188
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·›—⁄"
            Height          =   312
            Index           =   34
            Left            =   9876
            RightToLeft     =   -1  'True
            TabIndex        =   45
            Top             =   240
            Width           =   528
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„Ê”„"
            Height          =   312
            Index           =   33
            Left            =   12792
            RightToLeft     =   -1  'True
            TabIndex        =   44
            Top             =   600
            Width           =   1188
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·Õ«›·…"
            Height          =   312
            Index           =   32
            Left            =   9192
            RightToLeft     =   -1  'True
            TabIndex        =   43
            Top             =   600
            Width           =   1188
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic2 
         Height          =   705
         Left            =   120
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   0
         Width           =   13170
         _cx             =   23230
         _cy             =   1244
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
         Caption         =   "       √Œ·«¡ ÿ—›   "
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
         Begin VB.TextBox Text6 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   12960
            RightToLeft     =   -1  'True
            TabIndex        =   51
            Top             =   1320
            Visible         =   0   'False
            Width           =   855
         End
      End
      Begin VSFlex8UCtl.VSFlexGrid fg_Evacation 
         Height          =   5760
         Left            =   120
         TabIndex        =   52
         Top             =   720
         Width           =   13095
         _cx             =   23098
         _cy             =   10160
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
         Rows            =   50
         Cols            =   14
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FrmSearch_Hajj.frx":0C36
         ScrollTrack     =   -1  'True
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
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E2E9E9&
      Height          =   852
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   8640
      Width           =   13212
      Begin ImpulseButton.ISButton Cmd 
         Height          =   432
         Index           =   0
         Left            =   7020
         TabIndex        =   6
         Top             =   240
         Width           =   996
         _ExtentX        =   1746
         _ExtentY        =   767
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
         Height          =   432
         Index           =   1
         Left            =   5928
         TabIndex        =   7
         Top             =   240
         Width           =   1032
         _ExtentX        =   1826
         _ExtentY        =   767
         ButtonPositionImage=   1
         Caption         =   "„”Õ"
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
         Cancel          =   -1  'True
         Height          =   432
         Index           =   2
         Left            =   4920
         TabIndex        =   8
         Top             =   240
         Width           =   972
         _ExtentX        =   1720
         _ExtentY        =   767
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
   End
   Begin VB.Frame frm_BookingRequest 
      BackColor       =   &H00E2E9E9&
      Height          =   8772
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   -120
      Visible         =   0   'False
      Width           =   13215
      Begin VB.Frame Frame2 
         BackColor       =   &H00E2E9E9&
         Height          =   2175
         Left            =   -1080
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   6360
         Width           =   14292
         Begin VB.TextBox TxtCompnyOut 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   7080
            RightToLeft     =   -1  'True
            TabIndex        =   150
            Top             =   960
            Width           =   5490
         End
         Begin VB.TextBox TxtCompnyIn 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   7080
            RightToLeft     =   -1  'True
            TabIndex        =   149
            Top             =   600
            Width           =   5490
         End
         Begin VB.TextBox TxtGroupName 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1200
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   77
            Top             =   1320
            Width           =   4770
         End
         Begin VB.TextBox FlightNo 
            Alignment       =   1  'Right Justify
            Height          =   288
            Left            =   1200
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   28
            Top             =   600
            Width           =   1776
         End
         Begin VB.TextBox EmpCode 
            Alignment       =   1  'Right Justify
            Height          =   288
            Left            =   4200
            RightToLeft     =   -1  'True
            TabIndex        =   25
            Top             =   240
            Visible         =   0   'False
            Width           =   1812
         End
         Begin VB.TextBox EmpName 
            Alignment       =   1  'Right Justify
            Height          =   288
            Left            =   1200
            RightToLeft     =   -1  'True
            TabIndex        =   24
            Top             =   240
            Width           =   1776
         End
         Begin VB.TextBox ID 
            Alignment       =   1  'Right Justify
            Height          =   288
            Left            =   10596
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   9
            Top             =   240
            Width           =   1980
         End
         Begin MSDataListLib.DataCombo BranchID 
            Height          =   315
            Left            =   4200
            TabIndex        =   10
            Top             =   240
            Width           =   4890
            _ExtentX        =   8625
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo InClientID 
            Height          =   285
            Left            =   3840
            TabIndex        =   11
            Top             =   1320
            Width           =   2010
            _ExtentX        =   3545
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo GroupID 
            Height          =   315
            Left            =   3960
            TabIndex        =   12
            Top             =   1320
            Width           =   2010
            _ExtentX        =   3545
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo CompanyID 
            Height          =   285
            Left            =   3600
            TabIndex        =   13
            Top             =   1320
            Visible         =   0   'False
            Width           =   2010
            _ExtentX        =   3545
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo AirPortID 
            Height          =   288
            Left            =   4200
            TabIndex        =   14
            Top             =   600
            Width           =   1776
            _ExtentX        =   3122
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo AirLineID 
            Height          =   285
            Left            =   4200
            TabIndex        =   15
            Top             =   960
            Width           =   1770
            _ExtentX        =   3122
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo OutClientID 
            Height          =   315
            Left            =   7080
            TabIndex        =   16
            Top             =   1320
            Width           =   5490
            _ExtentX        =   9684
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo ProgrammID 
            Height          =   315
            Left            =   1200
            TabIndex        =   29
            Top             =   1680
            Width           =   4770
            _ExtentX        =   8414
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcbSeasons 
            Height          =   315
            Left            =   7080
            TabIndex        =   30
            Top             =   1680
            Width           =   5490
            _ExtentX        =   9684
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo VehicleType 
            Height          =   288
            Left            =   1200
            TabIndex        =   31
            Top             =   960
            Width           =   1776
            _ExtentX        =   3122
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·„Ã„Ê⁄…"
            Height          =   315
            Index           =   2
            Left            =   5955
            RightToLeft     =   -1  'True
            TabIndex        =   152
            Top             =   1320
            Width           =   1065
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·»—‰«„Ã"
            Height          =   315
            Index           =   11
            Left            =   5955
            RightToLeft     =   -1  'True
            TabIndex        =   151
            Top             =   1680
            Width           =   1065
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„Ê”„"
            Height          =   315
            Index           =   15
            Left            =   12903
            RightToLeft     =   -1  'True
            TabIndex        =   148
            Top             =   1680
            Width           =   1065
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·—Õ·…"
            Height          =   336
            Index           =   1
            Left            =   2820
            RightToLeft     =   -1  'True
            TabIndex        =   33
            Top             =   600
            Width           =   1188
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·Õ«›·« "
            Height          =   312
            Index           =   18
            Left            =   2832
            RightToLeft     =   -1  'True
            TabIndex        =   32
            Top             =   960
            Width           =   1188
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬂÊœ «·„‘—›"
            Height          =   255
            Index           =   1
            Left            =   6000
            RightToLeft     =   -1  'True
            TabIndex        =   27
            Top             =   240
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·„‘—›"
            Height          =   252
            Left            =   3000
            RightToLeft     =   -1  'True
            TabIndex        =   26
            Top             =   240
            Width           =   972
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„”·”·"
            Height          =   336
            Index           =   9
            Left            =   12780
            RightToLeft     =   -1  'True
            TabIndex        =   23
            Top             =   240
            Width           =   1188
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·›—⁄"
            Height          =   312
            Index           =   24
            Left            =   9876
            RightToLeft     =   -1  'True
            TabIndex        =   22
            Top             =   240
            Width           =   528
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·‘—ﬂ… «·”⁄ÊœÌ…"
            Height          =   312
            Index           =   6
            Left            =   12780
            RightToLeft     =   -1  'True
            TabIndex        =   21
            Top             =   600
            Width           =   1188
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·‘—ﬂ… „‰ «·Œ«—Ã"
            Height          =   315
            Index           =   4
            Left            =   12663
            RightToLeft     =   -1  'True
            TabIndex        =   20
            Top             =   960
            Width           =   1305
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·⁄„Ì·"
            Height          =   315
            Index           =   3
            Left            =   12063
            RightToLeft     =   -1  'True
            TabIndex        =   19
            Top             =   1320
            Width           =   1905
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„ÿ«—"
            Height          =   315
            Index           =   0
            Left            =   5835
            RightToLeft     =   -1  'True
            TabIndex        =   18
            Top             =   600
            Width           =   1185
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·ŒÿÊÿ «·ÃÊÌ…"
            Height          =   315
            Index           =   7
            Left            =   5955
            RightToLeft     =   -1  'True
            TabIndex        =   17
            Top             =   960
            Width           =   1065
         End
      End
      Begin C1SizerLibCtl.C1Elastic EleHeader 
         Height          =   585
         Left            =   0
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   0
         Width           =   13170
         _cx             =   23230
         _cy             =   1032
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
         Caption         =   "       »ÕÀ ÿ·» ÕÃ“   "
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
            Left            =   12960
            RightToLeft     =   -1  'True
            TabIndex        =   2
            Top             =   1320
            Visible         =   0   'False
            Width           =   855
         End
      End
      Begin VSFlex8UCtl.VSFlexGrid Fg_BookingRequest 
         Height          =   5640
         Left            =   -120
         TabIndex        =   3
         Top             =   600
         Width           =   13215
         _cx             =   23310
         _cy             =   9948
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
         Rows            =   50
         Cols            =   22
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FrmSearch_Hajj.frx":0E50
         ScrollTrack     =   -1  'True
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
   End
End
Attribute VB_Name = "FrmSearch_Hajj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit
Dim rs As ADODB.Recordset
Dim Rs_Temp As ADODB.Recordset
Dim DCboSearch As clsDCboSearch
Public calltype As Integer
Public SendForm As String

Private Sub Cmd_Click(Index As Integer)
On Error Resume Next
    Select Case Index

        Case 0
                If frm_BookingRequest.Visible = True Then
                If SendForm = "BookingRequest" Or SendForm = "BookingRequest11" Or SendForm = "BookingRequest12" Then
                    GetData_BookingRequest
                 ElseIf SendForm = "BookingRequest2" Or SendForm = "BookingRequest21" Then
                 GetData_BookingRequest2
                End If
                ElseIf frm_VehicleOperatorOrder.Visible = True Then
                    GetData_VehicleOperatorOrder
                ElseIf frm_Evacation.Visible = True Then
                    GetData_Evacation
                ElseIf frm_EndorseTrans.Visible = True Then
                If SendForm = "EndorseTransMash" Then
                GetData_EndorseTransMashar
                Else
                    GetData_EndorseTrans
                End If
                 ElseIf Fram_Deported.Visible = True Then
                    GetData_Deported
                  ElseIf Frame9.Visible = True Then
                    GetData_Distribution
                End If
        Case 1
            clear_all Me
            ResetDate
        Case 2
            Unload Me
    End Select

End Sub


Private Sub EmpCode_Change()

        Dim val1, val2
        If EmpCode.Text = "" Then Exit Sub
        Dim str As String, Name As String, Mobile As String
        Name = ""
        Mobile = ""
        
            str = " select  * from TblEmployee  where  fullcode = '" & EmpCode.Text & "'"
            Set Rs_Temp = New ADODB.Recordset
            Rs_Temp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
            If Rs_Temp.RecordCount > 0 Then
                Rs_Temp.MoveFirst '
                Name = IIf(IsNull(Rs_Temp("emp_name").value), "", Rs_Temp("emp_name").value)
                Mobile = IIf(IsNull(Rs_Temp("emp_Mobile").value), "", Rs_Temp("emp_Mobile").value)
             Else
                EmpName.Text = ""
              
            End If
            
            EmpName.Text = Name

End Sub

Private Sub EmpCode2_Change()

        Dim val1, val2
        If EmpCode2.Text = "" Then Exit Sub
        Dim str As String, Name As String, Mobile As String
        Name = ""
        Mobile = ""
        
            str = " select  * from TblEmployee  where  fullcode = '" & EmpCode2.Text & "'"
            Set Rs_Temp = New ADODB.Recordset
            Rs_Temp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
            If Rs_Temp.RecordCount > 0 Then
                Rs_Temp.MoveFirst '
                Name = IIf(IsNull(Rs_Temp("emp_name").value), "", Rs_Temp("emp_name").value)
                Mobile = IIf(IsNull(Rs_Temp("emp_Mobile").value), "", Rs_Temp("emp_Mobile").value)
             Else
                EmpName2.Text = ""
              
            End If
            
            EmpName2.Text = Name
End Sub

Private Sub EmployeeCode3_Change()
Dim val1, val2, recordno As String, fullcode As String, Emp_id As Integer
If EmployeeCode3.Text = "" Then Exit Sub
Dim str As String
    str = " select   Emp_ID , Fullcode , JobTypeID , BignDateWork  from tblemployee  where Fullcode = '" & EmployeeCode3.Text & "'"
    Set Rs_Temp = New ADODB.Recordset
    Rs_Temp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If Rs_Temp.RecordCount > 0 Then
        Rs_Temp.MoveFirst '
        Emp_id = IIf(IsNull(Rs_Temp("Emp_ID").value), 0, Rs_Temp("Emp_ID").value)
     End If
     Me.EmployeeID3.BoundText = Emp_id
End Sub

Private Sub EmployeeID3_Change()
Dim val1, val2, recordno As String, fullcode As String, Emp_id As Integer, JobTypeID   As Integer, BignDateWork As Date
If EmployeeID3.BoundText = "" Then Exit Sub
Dim str As String
    str = " select   Emp_ID , Fullcode , JobTypeID , BignDateWork  from tblemployee  where Emp_ID = " & EmployeeID3.BoundText
    Set Rs_Temp = New ADODB.Recordset
    Rs_Temp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If Rs_Temp.RecordCount > 0 Then
        Rs_Temp.MoveFirst '
        Emp_id = IIf(IsNull(Rs_Temp("Emp_ID").value), 0, Rs_Temp("Emp_ID").value)
        fullcode = IIf(IsNull(Rs_Temp("fullcode").value), "", Rs_Temp("fullcode").value)
        JobTypeID = IIf(IsNull(Rs_Temp("JobTypeID").value), 0, Rs_Temp("JobTypeID").value)
        BignDateWork = IIf(IsNull(Rs_Temp("BignDateWork").value), Date, Rs_Temp("BignDateWork").value)
     End If
      Me.EmployeeID3.BoundText = Emp_id
      Me.EmployeeCode3.Text = fullcode
   
End Sub

Private Sub fg_EndorseTrans_Click()
Dim I As Integer
   I = val(fg_EndorseTrans.TextMatrix(fg_EndorseTrans.Row, fg_EndorseTrans.ColIndex("ID")))
   If I > 0 Then
        If SendForm = "EndorseTransMash" Then
              FrmEndorseTransMashar.Retrive (I)
          Else
          FrmEndorseTrans.Retrive (I)
         End If
   End If

'Unload Me
ErrTrap:
End Sub

Private Sub fg_Evacation_Click()
Dim I As Integer
   I = val(fg_Evacation.TextMatrix(fg_Evacation.Row, fg_Evacation.ColIndex("ID")))
   If I > 0 Then
        If SendForm = "Evacation" Then
              FrmEvacation.Retrive (I)
         End If
   End If

'Unload Me
ErrTrap:
End Sub

Private Sub Fg_BookingRequest_Click()
 Dim I As Double
   I = val(Fg_BookingRequest.TextMatrix(Fg_BookingRequest.Row, Fg_BookingRequest.ColIndex("ID")))
   If I > 0 Then
        If SendForm = "BookingRequest" Then
              FrmBookingRequest.Retrive (I)
          ElseIf SendForm = "BookingRequest2" Then
              FrmBookingRequest2.Retrive (I)
           ElseIf SendForm = "BookingRequest11" Then
              FrmBookingRequest2.TxtOreder.Text = (I)
              FrmBookingRequest2.Booking
            ElseIf SendForm = "BookingRequest12" Then
            
              FrmDeported.TxtOrderID.Text = I
              FrmDeported.Booking
         End If
   End If
'Unload Me
ErrTrap:
End Sub


Private Sub FgDeported_Click()
FrmDeported.FindRec val(FgDeported.TextMatrix(FgDeported.Row, FgDeported.ColIndex("ID")))
End Sub

Private Sub Form_Activate()
'   PutFormOnTop Me.hWnd
   If SendForm = "EndorseTrans" Or SendForm = "EndorseTransMash" Then
    If SendForm = "EndorseTransMash" Then
    C1Elastic3.Caption = "»ÕÀ «—ﬂ«» «·„‘«⁄—"
    Else
    C1Elastic3.Caption = "»ÕÀ «—ﬂ«» «·ÕÃ«Ã "
    End If
  End If
'PutFormOnTop Me.hWnd, True
'mdifrmmain.Enabled = False
End Sub
Private Sub Fill_Deported()
 Dim Dcombos As ClsDataCombos
 Dim str As String
  Set Dcombos = New ClsDataCombos
    Dcombos.GetBranches DcbBrnch
    Dcombos.GetTblLocations Me.DcbLocatioID
    Dcombos.GetTblLocations Me.DcbLocatioID2
    Dcombos.GetTblTrips Me.DcbTypeTrip
    Dcombos.GetTblProgrammTypes Me.DcbProgramID
    Dcombos.GetTblShrines Me.DcbPath
    
    Dcombos.GetEmployees Me.DcbEmp
   str = "select ID, OperatorN from TblCarsData"
   fill_combo DcbEqupID, str
    SetDtpickerDate Me.DtpDateFrom
    SetDtpickerDate Me.DtpDateTo
    SetDtpickerDate Me.FrmDateGO
    SetDtpickerDate Me.ToDateGO
    SetDtpickerDate Me.FrmDateArrive
    SetDtpickerDate Me.TODateArrive
End Sub

Private Sub Fill_Evacation()
 Dim Dcombos As ClsDataCombos
  Dim str As String
   Set Dcombos = New ClsDataCombos
   Dcombos.GetBranches BranchID3
   'Dcombos.GetEmpJobsTypes Me.JobID
   Dcombos.GetEmployees EmployeeID3
   str = "select ID, OperatorN from TblCarsData"
   fill_combo CarID3, str
   str = " select id , name from TblSeasons  "
   fill_combo SeasonsID3, str
   ' Dcombos.getCountriesGovernments Me.inCity
End Sub

Private Sub Fill_EndorseTrans()
 Dim Dcombos As ClsDataCombos
  Dim str As String
   Set Dcombos = New ClsDataCombos
   
   Dcombos.GetBranches BranchID4
   Dcombos.GETNationality Nationality
   Dcombos.GetTblShrines Me.DcbPath2
   str = " select id , name from TblSeasons  "
   fill_combo SeasonsID3, str
End Sub



Private Sub Fill_Combos_BookingRequest()
  Dim Dcombos As ClsDataCombos
  Dim str As String
   Set Dcombos = New ClsDataCombos
   Dcombos.GetBranches BranchID
  ' Dcombos.GetCompany InClientID, 0, 0
  ' Dcombos.GetCompany OutClientID, 1, 0
   Dcombos.GetCompany OutClientID, 2, 0
   If SystemOptions.UserInterface = ArabicInterface Then
   str = "select ID, Name from tblcompaniesgroup"
   Else
   str = "select ID, Namee from tblcompaniesgroup"
   End If
   fill_combo GroupID, str
    
   If SystemOptions.UserInterface = ArabicInterface Then
   str = "select Id , Name from TblTourismCompanies "
   Else
   str = "select Id , NameE from TblTourismCompanies "
   End If
   fill_combo CompanyID, str
If SystemOptions.UserInterface = ArabicInterface Then
    str = "select Id , name  from tblairlines"
 Else
 str = "select Id , namee  from tblairlines"
 End If
   fill_combo AirLineID, str
 If SystemOptions.UserInterface = ArabicInterface Then
    str = "select id , name from TblAirport "
 Else
 str = "select id , namee from TblAirport "
 End If
   fill_combo AirPortID, str
  If SystemOptions.UserInterface = ArabicInterface Then
    str = "select id ,name from TblProgrammTypes "
 Else
    str = "select id ,nameE from TblProgrammTypes "
 End If
   fill_combo ProgrammID, str
  If SystemOptions.UserInterface = ArabicInterface Then
   str = " select id , name from TblCompaniesGroup  "
   Else
   str = " select id , nameE from TblCompaniesGroup  "
  End If
  str = str & " where Omra_Hajj=0"
   fill_combo Me.DcbSeasons, str
   Dcombos.GetTblCarsDataGroup VehicleType
   
   ' Dcombos.getCountriesGovernments Me.inCity
End Sub

Private Sub Fill_Combos_VehicleOperatorOrder()
   
   Dim Dcombos As ClsDataCombos
   Dim str As String
   Set Dcombos = New ClsDataCombos
   Dcombos.GetBranches BranchID2
   Dcombos.GetCompany InClientID2, 0, 0
   Dcombos.GetCompany OutClientID2, 2, 0
    
   str = "select ID, Name from tblcompaniesgroup"
   fill_combo GroupID2, str
   str = "select Id , Name from TblTourismCompanies "
   fill_combo CompanyID2, str
   
    str = "select Id , name  from tblairlines"
   fill_combo AirLineID2, str
   
    str = "select id , name from TblAirport "
   fill_combo AirPortID2, str
   
    str = "select id ,name from TblProgrammTypes "
   fill_combo ProgrammID2, str
   
   
   Dcombos.GetTblCarsDataGroup VehicleType2
     If SystemOptions.UserInterface = ArabicInterface Then
   str = " select id , name from TblCompaniesGroup  "
   Else
   str = " select id , nameE from TblCompaniesGroup  "
  End If
  str = str & " where Omra_Hajj=0"
   fill_combo Me.DcbSeasons, str
 '    If SystemOptions.UserInterface = ArabicInterface Then
 '  str = " select id , name from TblCompaniesGroup  "
 '  Else
 '  str = " select id , nameE from TblCompaniesGroup  "
 ' End If
 ' str = str & " where Omra_Hajj=0"
 '  fill_combo DcbSeasons2, str
   
   ' Dcombos.getCountriesGovernments Me.inCity
End Sub





Private Sub Form_Load()
    Dim GrdBack As ClsBackGroundPic
    Set DCboSearch = New clsDCboSearch
    Set GrdBack = New ClsBackGroundPic
    With Me.Fg_BookingRequest
        Set .WallPaper = GrdBack.Picture
        .AutoSize 0, .Cols - 1, False
    End With
 
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
    
  ResetDate
    
    
    If SendForm = "BookingRequest" Or SendForm = "BookingRequest2" Or SendForm = "BookingRequest11" Or SendForm = "BookingRequest12" Or SendForm = "BookingRequest21" Then
    If SendForm = "BookingRequest2" Or SendForm = "BookingRequest21" Then
    EleHeader.Caption = "»ÕÀ «„—  ‘€Ì·"
    Else
    EleHeader.Caption = " »ÕÀ ÿ·» ÕÃ“"
    End If
            frm_BookingRequest.Visible = True
           Fill_Combos_BookingRequest
    ElseIf SendForm = "VehicleOperatorOrder" Then
            frm_VehicleOperatorOrder.Visible = True
            Fill_Combos_VehicleOperatorOrder
    ElseIf SendForm = "Evacation" Then
            frm_Evacation.Visible = True
            Fill_Evacation
    ElseIf SendForm = "EndorseTrans" Or SendForm = "EndorseTransMash" Then
    If SendForm = "EndorseTransMash" Then
    C1Elastic3.Caption = "»ÕÀ «—ﬂ«» «·„‘«⁄—"
    Else
    C1Elastic3.Caption = "»ÕÀ «—ﬂ«» «·ÕÃ«Ã "
    End If
            frm_EndorseTrans.Visible = True
            Fill_EndorseTrans
   ElseIf SendForm = "Deported" Then
            Fram_Deported.Visible = True
            Fill_Deported
     ElseIf SendForm = "Distribution" Then
            Frame9.Visible = True
            ReloadCombos
    End If
    
 
End Sub

Private Sub ResetDate()
DtpDateFrom.value = ""
DtpDateTo.value = ""
FrmDateGO.value = ""
ToDateGO.value = ""
FrmDateArrive.value = ""
TODateArrive.value = ""
SDate.value = ""
ToSDate.value = ""
End Sub


Private Sub Form_Unload(Cancel As Integer)
mdifrmmain.Enabled = True
   ' FormPostion Me, SavePostion
   ' Set DCboSearch = Nothing
End Sub
Public Sub GetData_BookingRequest()
    
    Dim StrSQL As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim I As Integer
 StrSQL = " SELECT     dbo.TblBookingRequest.ID, dbo.TblBookingRequest.SDate, dbo.TblBookingRequest.BranchID, dbo.TblBranchesData.branch_name, "
 StrSQL = StrSQL & "                      dbo.TblBranchesData.branch_namee, dbo.TblBookingRequest.FlightNo, dbo.TblBookingRequest.EmpName, dbo.TblBookingRequest.EmpMbile,"
 StrSQL = StrSQL & "                     dbo.TblBookingRequest.ArriveDate, dbo.TblBookingRequest.ArriveTime, dbo.TblBookingRequest.VehicleNo, dbo.TblBookingRequest.ApproveFlag,"
 StrSQL = StrSQL & "                     dbo.TblBookingRequest.ApproveDate, dbo.TblBookingRequest.ApproveTime, dbo.TblBookingRequest.GroupName, dbo.TblBookingRequest.ModelID,"
 StrSQL = StrSQL & "                     dbo.TblBookingRequest.AirPortID, dbo.TblAirport.Name AS AirPortName, dbo.TblAirport.NameE AS AirPortNameE, dbo.TblBookingRequest.AirLineID,"
 StrSQL = StrSQL & "                     dbo.TblAirlines.Name AS AirLinName, dbo.TblAirlines.NameE AS AirLinNameE, dbo.TblBookingRequest.MekkaHotelID, dbo.TblBookingRequest.MadinaHotelID,"
 StrSQL = StrSQL & "                     dbo.TblBookingRequest.JeddahHotelID, dbo.TblBookingRequest.ProgrammID, dbo.TblProgrammTypes.Name, dbo.TblProgrammTypes.NameE,"
 StrSQL = StrSQL & "                     dbo.TblBookingRequest.VehicleType, dbo.TBLCarTypes.name AS VehicleTypeName, dbo.TBLCarTypes.namee AS VehicleTypeNameE,"
 StrSQL = StrSQL & "                     dbo.TblBookingRequest.NoteSerial1, dbo.TblBookingRequest.CompnyOut, dbo.TblBookingRequest.CompnyIn, dbo.TblBookingRequest.HotelJaddah,"
 StrSQL = StrSQL & "                     dbo.TblBookingRequest.HotelMadinh, dbo.TblBookingRequest.HotelMakh, dbo.TblBookingRequest.SeasonsID, dbo.TblCompaniesGroup.Name AS SeasonsName,"
 StrSQL = StrSQL & "                     dbo.TblCompaniesGroup.NameE AS SeasonsNameE, dbo.TblBookingRequest.ReservNo, dbo.TblBookingRequest.OutClientID, dbo.TblCustemers.CusName,"
 StrSQL = StrSQL & "                     dbo.TblCustemers.CusNamee , dbo.TblCustemers.fullcode"
 StrSQL = StrSQL & "  FROM         dbo.TBLCarTypes RIGHT OUTER JOIN"
 StrSQL = StrSQL & "                     dbo.TblCustemers RIGHT OUTER JOIN"
 StrSQL = StrSQL & "                     dbo.TblBookingRequest ON dbo.TblCustemers.CusID = dbo.TblBookingRequest.OutClientID LEFT OUTER JOIN"
 StrSQL = StrSQL & "                     dbo.TblCompaniesGroup ON dbo.TblBookingRequest.SeasonsID = dbo.TblCompaniesGroup.ID ON"
 StrSQL = StrSQL & "                     dbo.TBLCarTypes.id = dbo.TblBookingRequest.VehicleType LEFT OUTER JOIN"
 StrSQL = StrSQL & "                     dbo.TblProgrammTypes ON dbo.TblBookingRequest.ProgrammID = dbo.TblProgrammTypes.ID LEFT OUTER JOIN"
 StrSQL = StrSQL & "                     dbo.TblAirlines ON dbo.TblBookingRequest.AirLineID = dbo.TblAirlines.ID LEFT OUTER JOIN"
 StrSQL = StrSQL & "                     dbo.TblAirport ON dbo.TblBookingRequest.AirPortID = dbo.TblAirport.ID LEFT OUTER JOIN"
  StrSQL = StrSQL & "                    dbo.TblBranchesData ON dbo.TblBookingRequest.BranchID = dbo.TblBranchesData.branch_id"
 StrSQL = StrSQL & "   WHERE  (1 = 1) "
   
   If SendForm = "BookingRequest11" Or SendForm = "BookingRequest12" Then
   StrSQL = StrSQL & "   and  TblBookingRequest.StusID =  1"
   
   ' StrSQL = StrSQL & "   and  TblBookingRequest.ApproveFlag =  1"
   End If
    If Me.ID.Text <> "" Then
            StrSQL = StrSQL & "   and  TblBookingRequest.NoteSerial1 =  " & val(ID.Text) & ""
    End If
    
    If Me.BranchID.Text <> "" And val(Me.BranchID.BoundText) <> 0 Then
            StrSQL = StrSQL & "   and  TblBookingRequest.BranchID =  " & val(BranchID.BoundText)
    End If



    If TxtGroupName.Text <> "" Then
            StrSQL = StrSQL & "   and  TblBookingRequest.GroupName like '%" & TxtGroupName.Text & "%'"
    End If
    If TxtCompnyIn.Text <> "" Then
            StrSQL = StrSQL & "   and  TblBookingRequest.CompnyIn like '%" & TxtCompnyIn.Text & "%'"
    End If
    If TxtCompnyOut.Text <> "" Then
            StrSQL = StrSQL & "   and  TblBookingRequest.CompnyOut like '%" & TxtCompnyOut.Text & "%'"
    End If
    If Me.OutClientID.Text <> "" And val(OutClientID.BoundText) <> 0 Then
            StrSQL = StrSQL & "   and  TblBookingRequest.OutClientID =  " & val(OutClientID.BoundText)
    End If
    If val(Me.ProgrammID.BoundText) <> 0 And ProgrammID.Text <> "" Then
            StrSQL = StrSQL & "   and  TblBookingRequest.ProgrammID =  " & val(ProgrammID.BoundText)
    End If
    If val(Me.DcbSeasons.BoundText) <> 0 And DcbSeasons.Text <> "" Then
            StrSQL = StrSQL & "   and  TblBookingRequest.SeasonsID =  " & val(DcbSeasons.BoundText)
    End If
    If Me.VehicleType.BoundText <> "" Then
            StrSQL = StrSQL & "   and  TblBookingRequest.VehicleType =  " & val(VehicleType.BoundText)
    End If
    If Me.EmpName.Text <> "" Then
            StrSQL = StrSQL & "   and  TblBookingRequest.EmpName  like   '%" & EmpName.Text & "%'"
    End If
    

    
    If Me.FlightNo.Text <> "" Then
            StrSQL = StrSQL & "   and  TblBookingRequest.FlightNo  like  '%" & FlightNo.Text & "%' "
    End If
    
    If Me.AirPortID.Text <> "" And val(AirPortID.BoundText) <> 0 Then
            StrSQL = StrSQL & "   and  TblBookingRequest.AirPortID =  " & val(AirPortID.BoundText)
    End If
   
    
       If Me.AirLineID.BoundText <> "" Then
            StrSQL = StrSQL & "   and  TblBookingRequest.AirLineID =  " & val(AirLineID.BoundText)
    End If
    
    StrSQL = StrSQL
    StrSQL = StrSQL & " Order By TblBookingRequest.ID "
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
         Fg_BookingRequest.Rows = Fg_BookingRequest.FixedRows
         Exit Sub
    Else

        With Me.Fg_BookingRequest
            .Clear flexClearScrollable, flexClearEverything
            .Rows = .FixedRows
            .Rows = rs.RecordCount + .FixedRows

            rs.MoveFirst
        
            For I = .FixedRows To .Rows - 1
                .TextMatrix(I, .ColIndex("Serial")) = I
                .TextMatrix(I, .ColIndex("ID")) = IIf(IsNull(rs("ID").value), "", rs("ID").value)
                .TextMatrix(I, .ColIndex("EMPName")) = IIf(IsNull(rs("EMPName").value), "", rs("EMPName").value)
                .TextMatrix(I, .ColIndex("SDate")) = IIf(IsNull(rs("SDate").value), "", rs("SDate").value)
                .TextMatrix(I, .ColIndex("GroupName")) = IIf(IsNull(rs("GroupName").value), "", rs("GroupName").value)
                .TextMatrix(I, .ColIndex("NoteSerial1")) = IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value)
                .TextMatrix(I, .ColIndex("FlieghtNo")) = IIf(IsNull(rs("FlightNo").value), "", rs("FlightNo").value)
                .TextMatrix(I, .ColIndex("ArriveDate")) = IIf(IsNull(rs("ArriveDate").value), "", rs("ArriveDate").value)
                .TextMatrix(I, .ColIndex("ArriveTime")) = IIf(IsNull(rs("ArriveTime").value), "", rs("ArriveTime").value)
                .TextMatrix(I, .ColIndex("Model")) = IIf(IsNull(rs("ModelID").value), "", rs("ModelID").value)
                .TextMatrix(I, .ColIndex("Model")) = val(.TextMatrix(I, .ColIndex("Model"))) + 2015
              '  .TextMatrix(I, .ColIndex("JeddahHotelName")) = IIf(IsNull(rs("HotelJaddah").value), "", rs("HotelJaddah").value)
                '.TextMatrix(I, .ColIndex("HotelMadinh")) = IIf(IsNull(rs("HotelMadinh").value), "", rs("HotelMadinh").value)
              '  .TextMatrix(I, .ColIndex("HotelMakh")) = IIf(IsNull(rs("HotelMakh").value), "", rs("HotelMakh").value)
                .TextMatrix(I, .ColIndex("OutClientName")) = IIf(IsNull(rs("CompnyOut").value), "", rs("CompnyOut").value)
                .TextMatrix(I, .ColIndex("InClientName")) = IIf(IsNull(rs("CompnyIn").value), "", rs("CompnyIn").value)
             If SystemOptions.UserInterface = ArabicInterface Then
             .TextMatrix(I, .ColIndex("CusName")) = IIf(IsNull(rs("CusName").value), "", rs("CusName").value)
             .TextMatrix(I, .ColIndex("VehicleType")) = IIf(IsNull(rs("VehicleTypeName").value), "", rs("VehicleTypeName").value)
             .TextMatrix(I, .ColIndex("ProgrammName")) = IIf(IsNull(rs("Name").value), "", rs("Name").value)
             .TextMatrix(I, .ColIndex("SeasonsName")) = IIf(IsNull(rs("SeasonsName").value), "", rs("SeasonsName").value)
             Else
             .TextMatrix(I, .ColIndex("CusName")) = IIf(IsNull(rs("CusNamee").value), "", rs("CusNamee").value)
             .TextMatrix(I, .ColIndex("SeasonsName")) = IIf(IsNull(rs("SeasonsNameE").value), "", rs("SeasonsNameE").value)
             .TextMatrix(I, .ColIndex("VehicleType")) = IIf(IsNull(rs("VehicleTypeNameE").value), "", rs("VehicleTypeNameE").value)
             .TextMatrix(I, .ColIndex("ProgrammName")) = IIf(IsNull(rs("NameE").value), "", rs("NameE").value)
             End If
                 rs.MoveNext
            Next I

            .AutoSize 0, .Cols - 1, False
        
        End With

    End If

End Sub
Public Sub GetData_BookingRequest2()
    
    Dim StrSQL As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim I As Integer
StrSQL = " SELECT     TOP 100 PERCENT dbo.tblbookingrequest2.ID, dbo.tblbookingrequest2.SDate, dbo.tblbookingrequest2.BranchID, dbo.TblBranchesData.branch_name, "
StrSQL = StrSQL & "                      dbo.TblBranchesData.branch_namee, dbo.tblbookingrequest2.FlightNo, dbo.tblbookingrequest2.EmpName, dbo.tblbookingrequest2.EmpMbile,"
StrSQL = StrSQL & "                      dbo.tblbookingrequest2.ArriveDate, dbo.tblbookingrequest2.ArriveTime, dbo.tblbookingrequest2.VehicleNo, dbo.tblbookingrequest2.ApproveFlag,"
StrSQL = StrSQL & "                      dbo.tblbookingrequest2.ApproveDate, dbo.tblbookingrequest2.ApproveTime, dbo.tblbookingrequest2.GroupName, dbo.tblbookingrequest2.ModelID,"
StrSQL = StrSQL & "                      dbo.tblbookingrequest2.AirPortID, dbo.TblAirport.Name AS AirPortName, dbo.TblAirport.NameE AS AirPortNameE, dbo.tblbookingrequest2.AirLineID,"
StrSQL = StrSQL & "                      dbo.TblAirlines.Name AS AirLinName, dbo.TblAirlines.NameE AS AirLinNameE, dbo.tblbookingrequest2.MekkaHotelID, dbo.tblbookingrequest2.MadinaHotelID,"
StrSQL = StrSQL & "                      dbo.tblbookingrequest2.JeddahHotelID, dbo.tblbookingrequest2.ProgrammID, dbo.TblProgrammTypes.Name, dbo.TblProgrammTypes.NameE,"
StrSQL = StrSQL & "                      dbo.tblbookingrequest2.VehicleType, dbo.TBLCarTypes.name AS VehicleTypeName, dbo.TBLCarTypes.namee AS VehicleTypeNameE,"
StrSQL = StrSQL & "                      dbo.tblbookingrequest2.HotelMadinh, dbo.tblbookingrequest2.HotelJaddah, dbo.tblbookingrequest2.HotelMakh, dbo.tblbookingrequest2.NoteSerialOrder,"
StrSQL = StrSQL & "                      dbo.tblbookingrequest2.NoteSerial1, dbo.tblbookingrequest2.SeasonsID, dbo.tblbookingrequest2.NoteSerial, dbo.tblbookingrequest2.CompnyOut,"
StrSQL = StrSQL & "                      dbo.tblbookingrequest2.CompnyIn, dbo.TblCompaniesGroup.Name AS SeasonsName, dbo.TblCompaniesGroup.NameE AS SeasonsNameE,"
StrSQL = StrSQL & "                      dbo.tblbookingrequest2.OutClientID , dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.fullcode"
StrSQL = StrSQL & " FROM         dbo.TblCustemers RIGHT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.tblbookingrequest2 ON dbo.TblCustemers.CusID = dbo.tblbookingrequest2.OutClientID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblCompaniesGroup ON dbo.tblbookingrequest2.SeasonsID = dbo.TblCompaniesGroup.ID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TBLCarTypes ON dbo.tblbookingrequest2.VehicleType = dbo.TBLCarTypes.id LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblProgrammTypes ON dbo.tblbookingrequest2.ProgrammID = dbo.TblProgrammTypes.ID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblAirlines ON dbo.tblbookingrequest2.AirLineID = dbo.TblAirlines.ID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblAirport ON dbo.tblbookingrequest2.AirPortID = dbo.TblAirport.ID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblBranchesData ON dbo.tblbookingrequest2.BranchID = dbo.TblBranchesData.branch_id"

        
        StrSQL = StrSQL & "   WHERE  (1 = 1) "
   
   
    If Me.ID.Text <> "" Then
            StrSQL = StrSQL & "   and  TblBookingRequest2.NoteSerial1 =  " & val(ID.Text) & ""
    End If
    
    If Me.BranchID.Text <> "" And val(Me.BranchID.BoundText) <> 0 Then
            StrSQL = StrSQL & "   and  TblBookingRequest2.BranchID =  " & val(BranchID.BoundText)
    End If
    If val(Me.DcbSeasons.BoundText) <> 0 And DcbSeasons.Text <> "" Then
            StrSQL = StrSQL & "   and  TblBookingRequest2.SeasonsID =  " & val(DcbSeasons.BoundText)
    End If
  '      If Me.InClientID.BoundText <> "" Then
  '          StrSQL = StrSQL & "   and  TblBookingRequest2.InClientID =  " & val(InClientID.BoundText)
  '  End If
   If TxtGroupName.Text <> "" Then
            StrSQL = StrSQL & "   and  TblBookingRequest2.GroupName like '%" & TxtGroupName.Text & "%'"
    End If
    If TxtCompnyIn.Text <> "" Then
            StrSQL = StrSQL & "   and  TblBookingRequest2.CompnyIn like '%" & TxtCompnyIn.Text & "%'"
    End If
    If TxtCompnyOut.Text <> "" Then
            StrSQL = StrSQL & "   and  TblBookingRequest2.CompnyOut like '%" & TxtCompnyOut.Text & "%'"
    End If
    If Me.OutClientID.Text <> "" And val(OutClientID.BoundText) <> 0 Then
            StrSQL = StrSQL & "   and  TblBookingRequest2.OutClientID =  " & val(OutClientID.BoundText)
    End If
    
    'If Me.OutClientID.BoundText <> "" Then
    '        StrSQL = StrSQL & "   and  TblBookingRequest2.OutClientID =  " & val(OutClientID.BoundText)
    'End If

    If TxtGroupName.Text <> "" Then
            StrSQL = StrSQL & "   and  TblBookingRequest2.GroupName = '" & TxtGroupName.Text & "'"
    End If

  

   ' If Me.MekkaHotelID.BoundText <> "" Then
   '         StrSQL = StrSQL & "   and  TblBookingRequest2.MekkaHotelID =  " & val(MekkaHotelID.BoundText)
   ' End If

   ' If Me.JeddahHotelID.BoundText <> "" Then
   '         StrSQL = StrSQL & "   and  TblBookingRequest2.JeddahHotelID =  " & val(JeddahHotelID.BoundText)
   ' End If

   '     If Me.MadinaHotelID.BoundText <> "" Then
   '         StrSQL = StrSQL & "   and  TblBookingRequest2.MadinaHotelID =  " & val(MadinaHotelID.BoundText)
   ' End If

    If Me.ProgrammID.BoundText <> "" Then
            StrSQL = StrSQL & "   and  TblBookingRequest2.ProgrammID =  " & val(ProgrammID.BoundText)
    End If

    If Me.VehicleType.BoundText <> "" Then
            StrSQL = StrSQL & "   and  TblBookingRequest2.VehicleType =  " & val(VehicleType.BoundText)
    End If

    
    If Me.EmpName.Text <> "" Then
            StrSQL = StrSQL & "   and  TblBookingRequest2.EmpName  like   '%" & EmpName.Text & "%'"
    End If
    
    If Me.FlightNo.Text <> "" Then
            StrSQL = StrSQL & "   and  TblBookingRequest2.FlightNo  like  '%" & FlightNo.Text & "%' "
    End If
    
    If Me.AirPortID.BoundText <> "" Then
            StrSQL = StrSQL & "   and  TblBookingRequest2.AirPortID =  " & val(AirPortID.BoundText)
    End If
    
       If Me.AirLineID.BoundText <> "" Then
            StrSQL = StrSQL & "   and  TblBookingRequest2.AirLineID =  " & val(AirLineID.BoundText)
    End If
    
    StrSQL = StrSQL & " Order By TblBookingRequest2.ID "
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
         Fg_BookingRequest.Rows = Fg_BookingRequest.FixedRows
         Exit Sub
    Else

        With Me.Fg_BookingRequest
            .Clear flexClearScrollable, flexClearEverything
            .Rows = .FixedRows
            .Rows = rs.RecordCount + .FixedRows

            rs.MoveFirst
        
            For I = .FixedRows To .Rows - 1
                .TextMatrix(I, .ColIndex("Serial")) = I
                .TextMatrix(I, .ColIndex("ID")) = IIf(IsNull(rs("ID").value), "", rs("ID").value)
                .TextMatrix(I, .ColIndex("EMPName")) = IIf(IsNull(rs("EMPName").value), "", rs("EMPName").value)
                .TextMatrix(I, .ColIndex("SDate")) = IIf(IsNull(rs("SDate").value), "", rs("SDate").value)
                .TextMatrix(I, .ColIndex("GroupName")) = IIf(IsNull(rs("GroupName").value), "", rs("GroupName").value)
                .TextMatrix(I, .ColIndex("NoteSerial1")) = IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value)
                .TextMatrix(I, .ColIndex("FlieghtNo")) = IIf(IsNull(rs("FlightNo").value), "", rs("FlightNo").value)
                .TextMatrix(I, .ColIndex("ArriveDate")) = IIf(IsNull(rs("ArriveDate").value), "", rs("ArriveDate").value)
                .TextMatrix(I, .ColIndex("ArriveTime")) = IIf(IsNull(rs("ArriveTime").value), "", rs("ArriveTime").value)
                .TextMatrix(I, .ColIndex("Model")) = IIf(IsNull(rs("ModelID").value), "", rs("ModelID").value)
                .TextMatrix(I, .ColIndex("Model")) = val(.TextMatrix(I, .ColIndex("Model"))) + 2015
                .TextMatrix(I, .ColIndex("InClientName")) = IIf(IsNull(rs("CompnyIn").value), "", rs("CompnyIn").value)
                .TextMatrix(I, .ColIndex("OutClientName")) = IIf(IsNull(rs("CompnyOut").value), "", rs("CompnyOut").value)
             If SystemOptions.UserInterface = ArabicInterface Then
             .TextMatrix(I, .ColIndex("VehicleType")) = IIf(IsNull(rs("VehicleTypeName").value), "", rs("VehicleTypeName").value)
             .TextMatrix(I, .ColIndex("ProgrammName")) = IIf(IsNull(rs("Name").value), "", rs("Name").value)
             .TextMatrix(I, .ColIndex("SeasonsName")) = IIf(IsNull(rs("SeasonsName").value), "", rs("SeasonsName").value)
             .TextMatrix(I, .ColIndex("CusName")) = IIf(IsNull(rs("CusName").value), "", rs("CusName").value)
             Else
             .TextMatrix(I, .ColIndex("VehicleType")) = IIf(IsNull(rs("VehicleTypeNameE").value), "", rs("VehicleTypeNameE").value)
             .TextMatrix(I, .ColIndex("ProgrammName")) = IIf(IsNull(rs("NameE").value), "", rs("NameE").value)
             .TextMatrix(I, .ColIndex("SeasonsName")) = IIf(IsNull(rs("SeasonsNameE").value), "", rs("SeasonsNameE").value)
             .TextMatrix(I, .ColIndex("CusName")) = IIf(IsNull(rs("CusNamee").value), "", rs("CusNamee").value)
             End If
                 rs.MoveNext
            Next I

            .AutoSize 0, .Cols - 1, False
        
        End With

    End If

End Sub

Public Sub GetData_VehicleOperatorOrder()
    
    Dim StrSQL As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim I As Integer

  
StrSQL = StrSQL & "   SELECT dbo.tblvehicleoperatorOrder.ID, dbo.tblvehicleoperatorOrder.ProgrammID, dbo.tblvehicleoperatorOrder.AirLineID, dbo.tblvehicleoperatorOrder.AirPortID, dbo.tblvehicleoperatorOrder.CompanyID,"
StrSQL = StrSQL & "   dbo.tblvehicleoperatorOrder.MekkaHotelID, TblHotels_2.Name AS MekkaHotelName, dbo.tblvehicleoperatorOrder.MadinaHotelID, TblHotels_1.Name AS MadinaHotelName,"
StrSQL = StrSQL & "   dbo.tblvehicleoperatorOrder.JeddahHotelID, dbo.TblHotels.Name AS JeddahHotelName, dbo.tblvehicleoperatorOrder.InClientID, TblCustemers_1.CusName AS InClientName,"
StrSQL = StrSQL & "   dbo.tblvehicleoperatorOrder.OutClientID, dbo.TblCustemers.CusName AS OutClientName, dbo.TblAirport.Name AS AirPortName, dbo.TblAirlines.Name AS AirLineName,"
StrSQL = StrSQL & "   dbo.TblTourismCompanies.Name AS CompanyName, dbo.TblBranchesData.branch_name AS BranchName, dbo.TblCompaniesGroup.Name AS GroupName,"
StrSQL = StrSQL & "   dbo.TblProgrammTypes.Name AS ProgrammName, dbo.tblvehicleoperatorOrder.SDate, dbo.tblvehicleoperatorOrder.BranchID, dbo.tblvehicleoperatorOrder.FlightNo,"
StrSQL = StrSQL & "   dbo.tblvehicleoperatorOrder.emp, dbo.tblvehicleoperatorOrder.GroupID, dbo.tblvehicleoperatorOrder.other, dbo.tblvehicleoperatorOrder.EmpID, dbo.tblvehicleoperatorOrder.EmpName,"
StrSQL = StrSQL & "   dbo.tblvehicleoperatorOrder.EmpCode, dbo.tblvehicleoperatorOrder.EmpMbile, CONVERT(char(10), dbo.tblvehicleoperatorOrder.ArriveTime, 108) AS ArriveTime,"
StrSQL = StrSQL & "   dbo.tblvehicleoperatorOrder.ArriveDate, dbo.tblvehicleoperatorOrder.VehicleNo, dbo.tblvehicleoperatorOrder.Model, dbo.tblvehicleoperatorOrder.VehicleType,"
StrSQL = StrSQL & "   dbo.tblvehicleoperatorOrder.CreationUserID , dbo.tblvehicleoperatorOrder.CreationDate"
StrSQL = StrSQL & "   FROM     dbo.TblCustemers INNER JOIN"
StrSQL = StrSQL & "   dbo.tblvehicleoperatorOrder INNER JOIN"
StrSQL = StrSQL & "   dbo.TblProgrammTypes ON dbo.tblvehicleoperatorOrder.ProgrammID = dbo.TblProgrammTypes.ID INNER JOIN"
StrSQL = StrSQL & "   dbo.TblBranchesData ON dbo.tblvehicleoperatorOrder.BranchID = dbo.TblBranchesData.branch_id INNER JOIN"
StrSQL = StrSQL & "   dbo.TblCompaniesGroup ON dbo.tblvehicleoperatorOrder.GroupID = dbo.TblCompaniesGroup.ID INNER JOIN"
StrSQL = StrSQL & "   dbo.TblAirlines ON dbo.tblvehicleoperatorOrder.AirLineID = dbo.TblAirlines.ID INNER JOIN"
StrSQL = StrSQL & "   dbo.TblAirport ON dbo.tblvehicleoperatorOrder.AirPortID = dbo.TblAirport.ID INNER JOIN"
StrSQL = StrSQL & "   dbo.TblTourismCompanies ON dbo.tblvehicleoperatorOrder.CompanyID = dbo.TblTourismCompanies.ID INNER JOIN"
StrSQL = StrSQL & "   dbo.TblHotels AS TblHotels_2 ON dbo.tblvehicleoperatorOrder.MekkaHotelID = TblHotels_2.ID INNER JOIN"
StrSQL = StrSQL & "   dbo.TblHotels AS TblHotels_1 ON dbo.tblvehicleoperatorOrder.MadinaHotelID = TblHotels_1.ID INNER JOIN"
StrSQL = StrSQL & "   dbo.TblHotels ON dbo.tblvehicleoperatorOrder.JeddahHotelID = dbo.TblHotels.ID INNER JOIN"
StrSQL = StrSQL & "   dbo.TblCustemers AS TblCustemers_1 ON dbo.tblvehicleoperatorOrder.InClientID = TblCustemers_1.CusID ON dbo.TblCustemers.CusID = dbo.tblvehicleoperatorOrder.OutClientID"

        
        StrSQL = StrSQL & "   WHERE  (1 = 1) "
   
   
    If Me.ID2.Text <> "" Then
            StrSQL = StrSQL & "   and  tblvehicleoperatorOrder.id =  '" & ID2.Text & "'"
    End If
    
    If Me.BranchID2.BoundText <> "" Then
            StrSQL = StrSQL & "   and  tblvehicleoperatorOrder.BranchID =  " & val(BranchID2.BoundText)
    End If

        If Me.InClientID2.BoundText <> "" Then
            StrSQL = StrSQL & "   and  tblvehicleoperatorOrder.InClientID =  " & val(InClientID2.BoundText)
    End If

    If Me.OutClientID2.BoundText <> "" Then
            StrSQL = StrSQL & "   and  tblvehicleoperatorOrder.OutClientID =  " & val(OutClientID2.BoundText)
    End If

    If Me.GroupID2.BoundText <> "" Then
            StrSQL = StrSQL & "   and  tblvehicleoperatorOrder.GroupID =  " & val(GroupID2.BoundText)
    End If

    If Me.CompanyID2.BoundText <> "" Then
            StrSQL = StrSQL & "   and  tblvehicleoperatorOrder.CompanyID =  " & val(CompanyID2.BoundText)
    End If

    If Me.MekkaHotelID2.BoundText <> "" Then
            StrSQL = StrSQL & "   and  tblvehicleoperatorOrder.MekkaHotelID =  " & val(MekkaHotelID2.BoundText)
    End If

    If Me.JeddahHotelID2.BoundText <> "" Then
            StrSQL = StrSQL & "   and  tblvehicleoperatorOrder.JeddahHotelID =  " & val(JeddahHotelID2.BoundText)
    End If

        If Me.MadinaHotelID2.BoundText <> "" Then
            StrSQL = StrSQL & "   and  tblvehicleoperatorOrder.MadinaHotelID =  " & val(MadinaHotelID2.BoundText)
    End If

    If Me.ProgrammID2.BoundText <> "" Then
            StrSQL = StrSQL & "   and  tblvehicleoperatorOrder.ProgrammID =  " & val(ProgrammID2.BoundText)
    End If

    If Me.VehicleType2.BoundText <> "" Then
            StrSQL = StrSQL & "   and  tblvehicleoperatorOrder.VehicleType =  " & val(VehicleType2.BoundText)
    End If

    
    If Me.EmpName2.Text <> "" Then
            StrSQL = StrSQL & "   and  tblvehicleoperatorOrder.EmpName  like   '%" & EmpName2.Text & "%'"
    End If
    
    If Me.FlightNo2.Text <> "" Then
            StrSQL = StrSQL & "   and  tblvehicleoperatorOrder.FlightNo  like  '%" & FlightNo2.Text & "%' "
    End If
    
    If Me.AirPortID2.BoundText <> "" Then
            StrSQL = StrSQL & "   and  tblvehicleoperatorOrder.AirPortID =  " & val(AirPortID2.BoundText)
    End If
    
       If Me.AirLineID2.BoundText <> "" Then
            StrSQL = StrSQL & "   and  tblvehicleoperatorOrder.AirLineID =  " & val(AirLineID2.BoundText)
    End If
    
    StrSQL = StrSQL
    StrSQL = StrSQL & " Order By tblvehicleoperatorOrder.ID "
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
         fg_VehicleOperatorOrder.Rows = fg_VehicleOperatorOrder.FixedRows
         Exit Sub
    Else

        With Me.fg_VehicleOperatorOrder
            .Clear flexClearScrollable, flexClearEverything
            .Rows = .FixedRows
            .Rows = rs.RecordCount + .FixedRows

            rs.MoveFirst
        
            For I = .FixedRows To .Rows - 1
            
                .TextMatrix(I, .ColIndex("Serial")) = I
                .TextMatrix(I, .ColIndex("ID")) = IIf(IsNull(rs("ID").value), "", rs("ID").value)
                .TextMatrix(I, .ColIndex("EMPName")) = IIf(IsNull(rs("EMPName").value), "", rs("EMPName").value)
                .TextMatrix(I, .ColIndex("SDate")) = IIf(IsNull(rs("SDate").value), "", rs("SDate").value)
                .TextMatrix(I, .ColIndex("InClientName")) = IIf(IsNull(rs("InClientName").value), "", rs("InClientName").value)
                .TextMatrix(I, .ColIndex("OutClientName")) = IIf(IsNull(rs("OutClientName").value), "", rs("OutClientName").value)
                .TextMatrix(I, .ColIndex("GroupName")) = IIf(IsNull(rs("GroupName").value), "", rs("GroupName").value)
                .TextMatrix(I, .ColIndex("ComapanyName")) = IIf(IsNull(rs("CompanyName").value), "", rs("CompanyName").value)
                .TextMatrix(I, .ColIndex("FlieghtNo")) = IIf(IsNull(rs("FlightNo").value), "", rs("FlightNo").value)
                .TextMatrix(I, .ColIndex("ArriveDate")) = IIf(IsNull(rs("ArriveDate").value), "", rs("ArriveDate").value)
                .TextMatrix(I, .ColIndex("ArriveTime")) = IIf(IsNull(rs("ArriveTime").value), "", rs("ArriveTime").value)
                .TextMatrix(I, .ColIndex("ProgrammName")) = IIf(IsNull(rs("ProgrammName").value), "", rs("ProgrammName").value)
                .TextMatrix(I, .ColIndex("Model")) = IIf(IsNull(rs("Model").value), "", rs("Model").value)
                .TextMatrix(I, .ColIndex("MekkaHotelName")) = IIf(IsNull(rs("MekkaHotelName").value), "", rs("MekkaHotelName").value)
                .TextMatrix(I, .ColIndex("MadinaHotelName")) = IIf(IsNull(rs("MadinaHotelName").value), "", rs("MadinaHotelName").value)
                .TextMatrix(I, .ColIndex("JeddahHotelName")) = IIf(IsNull(rs("JeddahHotelName").value), "", rs("JeddahHotelName").value)
                .TextMatrix(I, .ColIndex("VehicleType")) = IIf(IsNull(rs("VehicleType").value), "", rs("VehicleType").value)
             
                
                 rs.MoveNext
            Next I

            .AutoSize 0, .Cols - 1, False
        
        End With

    End If

End Sub




Private Sub ChangeLang()

End Sub





Public Sub GetData_Evacation()
    
    Dim StrSQL As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim I As Integer

        
    StrSQL = StrSQL & "  SELECT dbo.TblSeasons.Name AS SeasonName, dbo.TblSeasons.NameE AS SeasonNameE, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode,"
    StrSQL = StrSQL & "  dbo.TblEmployee.BignDateWork, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblCarsData.id AS Expr1, dbo.TblCarsData.Branch_NO,"
    StrSQL = StrSQL & "  dbo.TblEvacation.ID, dbo.TblEvacation.SDate, dbo.TblEvacation.BranchID, dbo.TblEvacation.EmployeeID, dbo.TblEvacation.CarID, dbo.TblEvacation.LeaveDate,"
    StrSQL = StrSQL & "  dbo.TblEvacation.Trips, dbo.TblEvacation.Behavior, dbo.TblEvacation.Discipline, dbo.TblEvacation.Remark, dbo.TblEvacation.SeasonsID, dbo.TblEvacation.CreationUserID,"
    StrSQL = StrSQL & "  dbo.TblEvacation.CreationDate"
    StrSQL = StrSQL & "  FROM     dbo.TblEvacation LEFT OUTER JOIN"
    StrSQL = StrSQL & "  dbo.TblBranchesData ON dbo.TblEvacation.BranchID = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
    StrSQL = StrSQL & "  dbo.TblCarsData ON dbo.TblEvacation.CarID = dbo.TblCarsData.id LEFT OUTER JOIN"
    StrSQL = StrSQL & "  dbo.TblEmployee ON dbo.TblEvacation.EmployeeID = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
    StrSQL = StrSQL & "  dbo.TblSeasons ON dbo.TblEvacation.SeasonsID = dbo.TblSeasons.ID"
        
                  
        StrSQL = StrSQL & "   WHERE  (1 = 1) "
   
   
    If Me.ID3.Text <> "" Then
            StrSQL = StrSQL & "   and  TblEvacation.id =  '" & ID3.Text & "'"
    End If
    
    If Me.BranchID3.BoundText <> "" Then
            StrSQL = StrSQL & "   and  TblEvacation.BranchID =  " & val(BranchID3.BoundText)
    End If

     

    If Me.EmployeeID3.BoundText <> "" Then
            StrSQL = StrSQL & "   and  TblEvacation.EmployeeID =  " & val(EmployeeID3.BoundText)
    End If

    If Me.SeasonsID3.BoundText <> "" Then
            StrSQL = StrSQL & "   and  TblEvacation.SeasonsID =  " & val(SeasonsID3.BoundText)
    End If

    If Me.CarID3.BoundText <> "" Then
            StrSQL = StrSQL & "   and  TblEvacation.CarID =  " & val(CarID3.BoundText)
    End If

    If Me.Discipline3.Text <> "" Then
            StrSQL = StrSQL & "   and  TblEvacation.Discipline  like  '%" & Discipline3.Text & "%' "
    End If
    
   If Me.Behavior3.Text <> "" Then
            StrSQL = StrSQL & "   and  TblEvacation.Behavior  like  '%" & Behavior3.Text & "%' "
    End If
    
     If Me.Remark3.Text <> "" Then
            StrSQL = StrSQL & "   and  TblEvacation.Remark  like  '%" & Remark3.Text & "%' "
    End If
    
    
    StrSQL = StrSQL
    StrSQL = StrSQL & " Order By TblEvacation.ID "
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
         fg_Evacation.Rows = fg_Evacation.FixedRows
         Exit Sub
    Else

        With Me.fg_Evacation
            .Clear flexClearScrollable, flexClearEverything
            .Rows = .FixedRows
            .Rows = rs.RecordCount + .FixedRows

            rs.MoveFirst
        
            For I = .FixedRows To .Rows - 1
            
                .TextMatrix(I, .ColIndex("Serial")) = I
                .TextMatrix(I, .ColIndex("ID")) = IIf(IsNull(rs("ID").value), "", rs("ID").value)
                .TextMatrix(I, .ColIndex("branch_name")) = IIf(IsNull(rs("branch_name").value), "", rs("branch_name").value)
                .TextMatrix(I, .ColIndex("SDate")) = IIf(IsNull(rs("SDate").value), "", rs("SDate").value)
                .TextMatrix(I, .ColIndex("Emp_Name")) = IIf(IsNull(rs("Emp_Name").value), "", rs("Emp_Name").value)
                .TextMatrix(I, .ColIndex("SeasonName")) = IIf(IsNull(rs("SeasonName").value), "", rs("SeasonName").value)
                
                .TextMatrix(I, .ColIndex("Fullcode")) = IIf(IsNull(rs("Fullcode").value), "", rs("Fullcode").value)
                .TextMatrix(I, .ColIndex("Emp_Name")) = IIf(IsNull(rs("Emp_Name").value), "", rs("Emp_Name").value)
                .TextMatrix(I, .ColIndex("BignDateWork")) = IIf(IsNull(rs("BignDateWork").value), "", rs("BignDateWork").value)
                .TextMatrix(I, .ColIndex("LeaveDate")) = IIf(IsNull(rs("LeaveDate").value), "", rs("LeaveDate").value)
                .TextMatrix(I, .ColIndex("Trips")) = IIf(IsNull(rs("Trips").value), "", rs("Trips").value)
                .TextMatrix(I, .ColIndex("Behavior")) = IIf(IsNull(rs("Behavior").value), "", rs("Behavior").value)
                .TextMatrix(I, .ColIndex("Discipline")) = IIf(IsNull(rs("Discipline").value), "", rs("Discipline").value)
                .TextMatrix(I, .ColIndex("Remark")) = IIf(IsNull(rs("Remark").value), "", rs("Remark").value)
            
                
                 rs.MoveNext
            Next I

            .AutoSize 0, .Cols - 1, False
        
        End With

    End If

End Sub
Public Sub GetData_EndorseTransMashar()
    
    Dim StrSQL As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim I As Integer

StrSQL = " SELECT     TOP 100 PERCENT dbo.TblEndorseTransMashar.ID, dbo.TblEndorseTransMashar.Nationality, dbo.Nationality.name AS NationalityName, dbo.TblEndorseTransMashar.BranchID, "
StrSQL = StrSQL & "                      dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblEndorseTransMashar.CompanyID, dbo.TblTourismCompanies.Name AS CampanyName,"
StrSQL = StrSQL & "                      dbo.TblEndorseTransMashar.ApproveID, dbo.TblEndorseTransMashar.SDate, dbo.TblEndorseTransMashar.TotOlds, dbo.TblEndorseTransMashar.TotYoungs, dbo.TblEndorseTransMashar.Total,"
StrSQL = StrSQL & "                      dbo.TblEndorseTransMashar.Remark, dbo.TblEndorseTransMashar.CreationUserID, dbo.TblEndorseTransMashar.CreationDate, dbo.TblEndorseTransMashar.RecordDateH,"
StrSQL = StrSQL & "                      dbo.TblEndorseTransMashar.ReceptTime, dbo.TblEndorseTransMashar.Phone, dbo.TblEndorseTransMashar.NoVehicle, dbo.TblEndorseTransMashar.Capacity, dbo.TblEndorseTransMashar.TotalPrice,"
StrSQL = StrSQL & "                      dbo.TblEndorseTransMashar.SmalPrice, dbo.TblEndorseTransMashar.LargPrice, dbo.TblEndorseTransMashar.PathID, dbo.TblShrines.Name, dbo.TblShrines.NameE,"
StrSQL = StrSQL & "                      dbo.TblEndorseTransMashar.GroupName, dbo.Nationality.namee AS NationalityNameE, dbo.TblTourismCompanies.NameE AS CampanyNameE,"
StrSQL = StrSQL & "                      dbo.TblEndorseTransMashar.SeasonsID, dbo.TblCompaniesGroup.Name AS SeasonsName, dbo.TblCompaniesGroup.NameE AS SeasonsNameE"
StrSQL = StrSQL & " FROM         dbo.TblEndorseTransMashar LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblCompaniesGroup ON dbo.TblEndorseTransMashar.SeasonsID = dbo.TblCompaniesGroup.ID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblShrines ON dbo.TblEndorseTransMashar.PathID = dbo.TblShrines.ID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.Nationality ON dbo.TblEndorseTransMashar.Nationality = dbo.Nationality.id LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblTourismCompanies ON dbo.TblEndorseTransMashar.CompanyID = dbo.TblTourismCompanies.ID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblBranchesData ON dbo.TblEndorseTransMashar.BranchID = dbo.TblBranchesData.branch_id"
StrSQL = StrSQL & "   WHERE  (1 = 1) "
      
    If val(Me.ID4.Text) <> 0 Then
            StrSQL = StrSQL & "   and  TblEndorseTransMashar.id =  " & val(ID4.Text) & ""
    End If
    
    If val(Me.BranchID4.BoundText) <> 0 And Me.BranchID4.Text <> "" Then
            StrSQL = StrSQL & "   and  TblEndorseTransMashar.BranchID =  " & val(BranchID4.BoundText)
    End If

    If Me.SeasonsID3.BoundText <> "" Then
            StrSQL = StrSQL & "   and  TblEndorseTransMashar.SeasonsID =  " & val(SeasonsID3.BoundText)
    End If
    If val(Me.ApproveID.Text) <> 0 Then
            StrSQL = StrSQL & "   and  TblEndorseTransMashar.ApproveID =  " & val(ApproveID.Text)
    End If
    If Me.CompanyID4.BoundText <> "" Then
            StrSQL = StrSQL & "   and  TblEndorseTransMashar.CompanyID =  " & val(CompanyID4.BoundText)
    End If
    If Me.Nationality.BoundText <> "" Then
            StrSQL = StrSQL & "   and  TblEndorseTransMashar.Nationality =  " & val(Nationality.BoundText)
    End If
     If Me.Remark4.Text <> "" Then
            StrSQL = StrSQL & "   and  TblEndorseTransMashar.Remark  like  '%" & Remark4.Text & "%' "
    End If
     If val(Me.DcbPath2.BoundText) <> 0 And Me.DcbPath2.Text <> "" Then
            StrSQL = StrSQL & "   and  TblEndorseTransMashar.PathID =  " & val(DcbPath2.BoundText)
    End If
    
    StrSQL = StrSQL
    StrSQL = StrSQL & " Order By TblEndorseTransMashar.ID "
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
         fg_EndorseTrans.Rows = fg_EndorseTrans.FixedRows
         Exit Sub
    Else

        With Me.fg_EndorseTrans
            .Clear flexClearScrollable, flexClearEverything
            .Rows = .FixedRows
            .Rows = rs.RecordCount + .FixedRows

            rs.MoveFirst
        
            For I = .FixedRows To .Rows - 1
            
                .TextMatrix(I, .ColIndex("Serial")) = I
                .TextMatrix(I, .ColIndex("ID")) = IIf(IsNull(rs("ID").value), "", rs("ID").value)
                .TextMatrix(I, .ColIndex("branch_name")) = IIf(IsNull(rs("branch_name").value), "", rs("branch_name").value)
                .TextMatrix(I, .ColIndex("SDate")) = IIf(IsNull(rs("SDate").value), "", rs("SDate").value)
                If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(I, .ColIndex("Name")) = IIf(IsNull(rs("Name").value), "", rs("Name").value)
                .TextMatrix(I, .ColIndex("SeasonName")) = IIf(IsNull(rs("SeasonsName").value), "", rs("SeasonsName").value)
                .TextMatrix(I, .ColIndex("CompanyName")) = IIf(IsNull(rs("CampanyName").value), "", rs("CampanyName").value)
                .TextMatrix(I, .ColIndex("NationalityName")) = IIf(IsNull(rs("NationalityName").value), "", rs("NationalityName").value)
                Else
                .TextMatrix(I, .ColIndex("NationalityName")) = IIf(IsNull(rs("NationalityNameE").value), "", rs("NationalityNameE").value)
                .TextMatrix(I, .ColIndex("CompanyName")) = IIf(IsNull(rs("CampanyNameE").value), "", rs("CampanyNameE").value)
                .TextMatrix(I, .ColIndex("Name")) = IIf(IsNull(rs("NameE").value), "", rs("NameE").value)
                .TextMatrix(I, .ColIndex("SeasonName")) = IIf(IsNull(rs("SeasonsNameE").value), "", rs("SeasonsNameE").value)
                End If
                .TextMatrix(I, .ColIndex("GroupName")) = IIf(IsNull(rs("GroupName").value), "", rs("GroupName").value)
                .TextMatrix(I, .ColIndex("TotOlds")) = IIf(IsNull(rs("TotOlds").value), "", rs("TotOlds").value)
                .TextMatrix(I, .ColIndex("TotYoungs")) = IIf(IsNull(rs("TotYoungs").value), "", rs("TotYoungs").value)
                .TextMatrix(I, .ColIndex("Total")) = IIf(IsNull(rs("Total").value), "", rs("Total").value)
                .TextMatrix(I, .ColIndex("Remark")) = IIf(IsNull(rs("Remark").value), "", rs("Remark").value)
                 rs.MoveNext
            Next I

            .AutoSize 0, .Cols - 1, False
        
        End With

    End If

End Sub
Public Sub GetData_EndorseTrans()
    
    Dim StrSQL As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim I As Integer

StrSQL = " SELECT     TOP 100 PERCENT dbo.TblEndorseTrans.ID, dbo.TblEndorseTrans.Nationality, dbo.Nationality.name AS NationalityName, dbo.TblEndorseTrans.BranchID, "
StrSQL = StrSQL & "                      dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblEndorseTrans.CompanyID, dbo.TblTourismCompanies.Name AS CampanyName,"
StrSQL = StrSQL & "                      dbo.TblEndorseTrans.ApproveID, dbo.TblEndorseTrans.SDate, dbo.TblEndorseTrans.TotOlds, dbo.TblEndorseTrans.TotYoungs, dbo.TblEndorseTrans.Total,"
StrSQL = StrSQL & "                      dbo.TblEndorseTrans.Remark, dbo.TblEndorseTrans.CreationUserID, dbo.TblEndorseTrans.CreationDate, dbo.TblEndorseTrans.RecordDateH,"
StrSQL = StrSQL & "                      dbo.TblEndorseTrans.ReceptTime, dbo.TblEndorseTrans.Phone, dbo.TblEndorseTrans.NoVehicle, dbo.TblEndorseTrans.Capacity, dbo.TblEndorseTrans.TotalPrice,"
StrSQL = StrSQL & "                      dbo.TblEndorseTrans.SmalPrice, dbo.TblEndorseTrans.LargPrice, dbo.TblEndorseTrans.PathID, dbo.TblShrines.Name, dbo.TblShrines.NameE,"
StrSQL = StrSQL & "                      dbo.TblEndorseTrans.GroupName, dbo.Nationality.namee AS NationalityNameE, dbo.TblTourismCompanies.NameE AS CampanyNameE,"
StrSQL = StrSQL & "                      dbo.TblEndorseTrans.SeasonsID, dbo.TblCompaniesGroup.Name AS SeasonsName, dbo.TblCompaniesGroup.NameE AS SeasonsNameE"
StrSQL = StrSQL & " FROM         dbo.TblEndorseTrans LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblCompaniesGroup ON dbo.TblEndorseTrans.SeasonsID = dbo.TblCompaniesGroup.ID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblShrines ON dbo.TblEndorseTrans.PathID = dbo.TblShrines.ID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.Nationality ON dbo.TblEndorseTrans.Nationality = dbo.Nationality.id LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblTourismCompanies ON dbo.TblEndorseTrans.CompanyID = dbo.TblTourismCompanies.ID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblBranchesData ON dbo.TblEndorseTrans.BranchID = dbo.TblBranchesData.branch_id"
StrSQL = StrSQL & "   WHERE  (1 = 1) "
      
    If val(Me.ID4.Text) <> 0 Then
            StrSQL = StrSQL & "   and  TblEndorseTrans.id =  " & val(ID4.Text) & ""
    End If
    
    If val(Me.BranchID4.BoundText) <> 0 And Me.BranchID4.Text <> "" Then
            StrSQL = StrSQL & "   and  TblEndorseTrans.BranchID =  " & val(BranchID4.BoundText)
    End If

    If Me.SeasonsID3.BoundText <> "" Then
            StrSQL = StrSQL & "   and  TblEndorseTrans.SeasonsID =  " & val(SeasonsID3.BoundText)
    End If
    If val(Me.ApproveID.Text) <> 0 Then
            StrSQL = StrSQL & "   and  TblEndorseTrans.ApproveID =  " & val(ApproveID.Text)
    End If
    If Me.CompanyID4.BoundText <> "" Then
            StrSQL = StrSQL & "   and  TblEndorseTrans.CompanyID =  " & val(CompanyID4.BoundText)
    End If
    If Me.Nationality.BoundText <> "" Then
            StrSQL = StrSQL & "   and  TblEndorseTrans.Nationality =  " & val(Nationality.BoundText)
    End If
     If Me.Remark4.Text <> "" Then
            StrSQL = StrSQL & "   and  TblEndorseTrans.Remark  like  '%" & Remark4.Text & "%' "
    End If
     If val(Me.DcbPath2.BoundText) <> 0 And Me.DcbPath2.Text <> "" Then
            StrSQL = StrSQL & "   and  TblEndorseTrans.PathID =  " & val(DcbPath2.BoundText)
    End If
    
    StrSQL = StrSQL
    StrSQL = StrSQL & " Order By TblEndorseTrans.ID "
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
         fg_EndorseTrans.Rows = fg_EndorseTrans.FixedRows
         Exit Sub
    Else

        With Me.fg_EndorseTrans
            .Clear flexClearScrollable, flexClearEverything
            .Rows = .FixedRows
            .Rows = rs.RecordCount + .FixedRows

            rs.MoveFirst
        
            For I = .FixedRows To .Rows - 1
            
                .TextMatrix(I, .ColIndex("Serial")) = I
                .TextMatrix(I, .ColIndex("ID")) = IIf(IsNull(rs("ID").value), "", rs("ID").value)
                .TextMatrix(I, .ColIndex("branch_name")) = IIf(IsNull(rs("branch_name").value), "", rs("branch_name").value)
                .TextMatrix(I, .ColIndex("SDate")) = IIf(IsNull(rs("SDate").value), "", rs("SDate").value)
                If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(I, .ColIndex("Name")) = IIf(IsNull(rs("Name").value), "", rs("Name").value)
                .TextMatrix(I, .ColIndex("SeasonName")) = IIf(IsNull(rs("SeasonsName").value), "", rs("SeasonsName").value)
                .TextMatrix(I, .ColIndex("CompanyName")) = IIf(IsNull(rs("CampanyName").value), "", rs("CampanyName").value)
                .TextMatrix(I, .ColIndex("NationalityName")) = IIf(IsNull(rs("NationalityName").value), "", rs("NationalityName").value)
                Else
                .TextMatrix(I, .ColIndex("NationalityName")) = IIf(IsNull(rs("NationalityNameE").value), "", rs("NationalityNameE").value)
                .TextMatrix(I, .ColIndex("CompanyName")) = IIf(IsNull(rs("CampanyNameE").value), "", rs("CampanyNameE").value)
                .TextMatrix(I, .ColIndex("Name")) = IIf(IsNull(rs("NameE").value), "", rs("NameE").value)
                .TextMatrix(I, .ColIndex("SeasonName")) = IIf(IsNull(rs("SeasonsNameE").value), "", rs("SeasonsNameE").value)
                End If
                .TextMatrix(I, .ColIndex("GroupName")) = IIf(IsNull(rs("GroupName").value), "", rs("GroupName").value)
                .TextMatrix(I, .ColIndex("TotOlds")) = IIf(IsNull(rs("TotOlds").value), "", rs("TotOlds").value)
                .TextMatrix(I, .ColIndex("TotYoungs")) = IIf(IsNull(rs("TotYoungs").value), "", rs("TotYoungs").value)
                .TextMatrix(I, .ColIndex("Total")) = IIf(IsNull(rs("Total").value), "", rs("Total").value)
                .TextMatrix(I, .ColIndex("Remark")) = IIf(IsNull(rs("Remark").value), "", rs("Remark").value)
                 rs.MoveNext
            Next I

            .AutoSize 0, .Cols - 1, False
        
        End With

    End If

End Sub
Public Sub GetData_Deported()
    
    Dim StrSQL As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim I As Integer
StrSQL = " SELECT     dbo.TblDeported.ID, dbo.TblDeported.BranchID, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblDeported.RecordDate,"
StrSQL = StrSQL & "                      dbo.TblDeported.RecordDateH, dbo.TblDeported.CurrDate, dbo.TblDeported.CurrDateH, dbo.TblDeported.LocatioID, dbo.TblLocations.Name, dbo.TblLocations.NameE,"
StrSQL = StrSQL & "                      dbo.TblDeported.DayName1, dbo.TblDeported.DayName2, dbo.TblDeported.OrderID, dbo.TblDeported.Phone, dbo.TblDeported.LocatioID2,"
StrSQL = StrSQL & "                      TblLocations_1.Name AS LocationName, TblLocations_1.NameE AS LocationNameE, dbo.TblDeported.DriverName, dbo.TblDeported.TimeOut, dbo.TblDeported.TimeIn,"
StrSQL = StrSQL & "                       dbo.TblDeported.Address, dbo.TblDeported.CurrDate2, dbo.TblDeported.CurrDateH2, dbo.TblDeported.SuperID, dbo.TblEmployee.Emp_Name,"
StrSQL = StrSQL & "                      dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee, dbo.TblDeported.PathID, dbo.TblShrines.Name AS PathName, dbo.TblShrines.NameE AS PathNameE,"
StrSQL = StrSQL & "                      dbo.TblDeported.ProgramID, dbo.TblProgrammTypes.Name AS ProgName, dbo.TblProgrammTypes.NameE AS ProgNameE, dbo.TblDeported.TypeTrip,"
StrSQL = StrSQL & "                      dbo.TblTrips.Name AS TripName, dbo.TblTrips.NameE AS TripNameE, dbo.TblDeported.EqupID, dbo.TblCarsData.OperatorN, dbo.TblDeported.DriverID,"
StrSQL = StrSQL & "                      TblEmployee_1.Emp_Name AS DriveEmp_Name, TblEmployee_1.Fullcode AS DrivFullcode, TblEmployee_1.Emp_Namee AS DriveEmp_NameE"
StrSQL = StrSQL & " FROM         dbo.TblDeported LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblEmployee TblEmployee_1 ON dbo.TblDeported.DriverID = TblEmployee_1.Emp_ID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblCarsData ON dbo.TblDeported.EqupID = dbo.TblCarsData.id LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblTrips ON dbo.TblDeported.TypeTrip = dbo.TblTrips.ID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblProgrammTypes ON dbo.TblDeported.ProgramID = dbo.TblProgrammTypes.ID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblShrines ON dbo.TblDeported.PathID = dbo.TblShrines.ID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblEmployee ON dbo.TblDeported.SuperID = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblLocations TblLocations_1 ON dbo.TblDeported.LocatioID2 = TblLocations_1.ID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblLocations ON dbo.TblDeported.LocatioID = dbo.TblLocations.ID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblBranchesData ON dbo.TblDeported.BranchID = dbo.TblBranchesData.branch_id"
StrSQL = StrSQL & "   WHERE  (1 = 1) "
      
    If val(TxtIDFrom.Text) <> 0 Then
            StrSQL = StrSQL & "   and  TblDeported.id >=  " & val(TxtIDFrom.Text) & ""
    End If
      If val(Me.TxtIDTO.Text) <> 0 Then
            StrSQL = StrSQL & "   and  TblDeported.id <=  " & val(TxtIDTO.Text) & ""
    End If
    
    If Me.DcbBrnch.Text <> "" And val(DcbBrnch.BoundText) <> 0 Then
            StrSQL = StrSQL & "   and  TblDeported.BranchID =  " & val(DcbBrnch.BoundText)
    End If
    If Me.DcbEqupID.Text <> "" And val(DcbEqupID.BoundText) <> 0 Then
            StrSQL = StrSQL & "   and  TblDeported.EqupID =  " & val(DcbEqupID.BoundText)
    End If
      If Me.DcbProgramID.Text <> "" And val(DcbProgramID.BoundText) <> 0 Then
            StrSQL = StrSQL & "   and  TblDeported.ProgramID =  " & val(DcbProgramID.BoundText)
    End If
        If Me.DcbLocatioID.Text <> "" And val(DcbLocatioID.BoundText) <> 0 Then
            StrSQL = StrSQL & "   and  TblDeported.LocatioID =  " & val(DcbLocatioID.BoundText)
    End If
         If Me.DcbLocatioID2.Text <> "" And val(DcbLocatioID2.BoundText) <> 0 Then
            StrSQL = StrSQL & "   and  TblDeported.LocatioID2 =  " & val(DcbLocatioID2.BoundText)
    End If
     If Me.DcbPath.Text <> "" And val(DcbPath.BoundText) <> 0 Then
            StrSQL = StrSQL & "   and  TblDeported.PathID =  " & val(DcbPath.BoundText)
    End If
       If Me.DcbTypeTrip.Text <> "" And val(DcbTypeTrip.BoundText) <> 0 Then
            StrSQL = StrSQL & "   and  TblDeported.TypeTrip =  " & val(DcbTypeTrip.BoundText)
    End If
       If Me.DcbEmp.Text <> "" And val(DcbEmp.BoundText) <> 0 Then
            StrSQL = StrSQL & "   and  TblDeported.SuperID =  " & val(DcbEmp.BoundText)
    End If
       If Me.TxtLeader.Text <> "" Then
            StrSQL = StrSQL & "   and  (TblDeported.DriverName  like  '%" & TxtLeader.Text & "%' "
            If SystemOptions.UserInterface = ArabicInterface Then
            StrSQL = StrSQL & "    or TblEmployee_1.Emp_Name  like  '%" & TxtLeader.Text & "%' )"
            Else
            StrSQL = StrSQL & "    or TblEmployee_1.Emp_Namee  like  '%" & TxtLeader.Text & "%' )"
            End If
    End If
        If Not IsNull(DtpDateFrom.value) Then
            StrSQL = StrSQL & "   and  TblDeported.RecordDate >=  " & SQLDate(DtpDateFrom.value, True)
    End If
         If Not IsNull(DtpDateTo.value) Then
            StrSQL = StrSQL & "   and  TblDeported.RecordDate <=  " & SQLDate(DtpDateTo.value, True)
    End If
            If Not IsNull(FrmDateGO.value) Then
            StrSQL = StrSQL & "   and  TblDeported.CurrDate >=  " & SQLDate(FrmDateGO.value, True)
    End If
         If Not IsNull(ToDateGO.value) Then
            StrSQL = StrSQL & "   and  TblDeported.CurrDate <=  " & SQLDate(ToDateGO.value, True)
    End If
               If Not IsNull(FrmDateArrive.value) Then
            StrSQL = StrSQL & "   and  TblDeported.CurrDate2 >=  " & SQLDate(FrmDateArrive.value, True)
    End If
         If Not IsNull(TODateArrive.value) Then
            StrSQL = StrSQL & "   and  TblDeported.CurrDate2 <=  " & SQLDate(TODateArrive.value, True)
    End If
    
    StrSQL = StrSQL
    StrSQL = StrSQL & " Order By TblDeported.ID "
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
         FgDeported.Rows = FgDeported.FixedRows
         Exit Sub
    Else

        With Me.FgDeported
            .Clear flexClearScrollable, flexClearEverything
            .Rows = .FixedRows
            .Rows = rs.RecordCount + .FixedRows

            rs.MoveFirst
        
            For I = .FixedRows To .Rows - 1
            
                .TextMatrix(I, .ColIndex("Serial")) = I
                .TextMatrix(I, .ColIndex("ID")) = IIf(IsNull(rs("ID").value), "", rs("ID").value)
                .TextMatrix(I, .ColIndex("SDate")) = IIf(IsNull(rs("RecordDate").value), "", rs("RecordDate").value)
                .TextMatrix(I, .ColIndex("CurrDate")) = IIf(IsNull(rs("CurrDate").value), "", rs("CurrDate").value)
                .TextMatrix(I, .ColIndex("CurrDate2")) = IIf(IsNull(rs("CurrDate2").value), "", rs("CurrDate2").value)
                .TextMatrix(I, .ColIndex("Branch_NO")) = IIf(IsNull(rs("OperatorN").value), "", rs("OperatorN").value)
                If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(I, .ColIndex("Name")) = IIf(IsNull(rs("Name").value), "", rs("Name").value)
                .TextMatrix(I, .ColIndex("Emp_Name")) = IIf(IsNull(rs("Emp_Name").value), "", rs("Emp_Name").value)
                .TextMatrix(I, .ColIndex("LocationName")) = IIf(IsNull(rs("LocationName").value), "", rs("LocationName").value)
                .TextMatrix(I, .ColIndex("PathName")) = IIf(IsNull(rs("PathName").value), "", rs("PathName").value)
                .TextMatrix(I, .ColIndex("ProgName")) = IIf(IsNull(rs("ProgName").value), "", rs("ProgName").value)
                 .TextMatrix(I, .ColIndex("TripName")) = IIf(IsNull(rs("TripName").value), "", rs("TripName").value)
                 .TextMatrix(I, .ColIndex("DriveEmp_Name")) = IIf(IsNull(rs("DriveEmp_Name").value), "", rs("DriveEmp_Name").value)
                .TextMatrix(I, .ColIndex("branch_name")) = IIf(IsNull(rs("branch_name").value), "", rs("branch_name").value)
                Else
                .TextMatrix(I, .ColIndex("Name")) = IIf(IsNull(rs("NameE").value), "", rs("NameE").value)
                .TextMatrix(I, .ColIndex("Emp_Name")) = IIf(IsNull(rs("Emp_Namee").value), "", rs("Emp_Namee").value)
                .TextMatrix(I, .ColIndex("LocationName")) = IIf(IsNull(rs("LocationNameE").value), "", rs("LocationNameE").value)
                .TextMatrix(I, .ColIndex("PathName")) = IIf(IsNull(rs("PathNameE").value), "", rs("PathNameE").value)
                .TextMatrix(I, .ColIndex("ProgName")) = IIf(IsNull(rs("ProgNameE").value), "", rs("ProgNameE").value)
                 .TextMatrix(I, .ColIndex("TripName")) = IIf(IsNull(rs("TripNameE").value), "", rs("TripNameE").value)
                 .TextMatrix(I, .ColIndex("DriveEmp_Name")) = IIf(IsNull(rs("DriveEmp_NameE").value), "", rs("DriveEmp_NameE").value)
                .TextMatrix(I, .ColIndex("branch_name")) = IIf(IsNull(rs("branch_namee").value), "", rs("branch_namee").value)
                End If
                If .TextMatrix(I, .ColIndex("DriveEmp_Name")) = "" Then
                .TextMatrix(I, .ColIndex("DriveEmp_Name")) = IIf(IsNull(rs("DriverName").value), "", rs("DriverName").value)
                End If
                 rs.MoveNext
            Next I

            .AutoSize 0, .Cols - 1, False
        
        End With

    End If

End Sub
Public Sub GetData_Distribution()
    
    Dim StrSQL As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim I As Integer
StrSQL = " SELECT     dbo.TblBusesDistribution.ID, dbo.TblBusesDistribution.SDate, dbo.TblBusesDistribution.BranchID, dbo.TblBranchesData.branch_name, "
StrSQL = StrSQL & "                      dbo.TblBranchesData.branch_namee, dbo.TblBusesDistribution.RecordDateH, dbo.TblBusesDistribution.NoVehicle, dbo.TblBusesDistribution.Capacity,"
StrSQL = StrSQL & "                      dbo.TblBusesDistribution.OrderNo, dbo.TblBusesDistributionDet.Capacity AS CapacityDet, dbo.TblBusesDistributionDet.EmpID, dbo.TblEmployee.Emp_Name,"
StrSQL = StrSQL & "                      dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee, dbo.TblBusesDistributionDet.CarID, dbo.TblCarsData.Branch_NO, dbo.TblCarsData.OperatorN,"
StrSQL = StrSQL & "                      dbo.TblCarsData.CarsTypeId , dbo.TBLCarTypes.name, dbo.TBLCarTypes.NameE"
StrSQL = StrSQL & " FROM         dbo.TblCarsData LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TBLCarTypes ON dbo.TblCarsData.CarsTypeId = dbo.TBLCarTypes.id RIGHT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblBusesDistributionDet ON dbo.TblCarsData.id = dbo.TblBusesDistributionDet.CarID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblEmployee ON dbo.TblBusesDistributionDet.EmpID = dbo.TblEmployee.Emp_ID RIGHT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblBusesDistribution ON dbo.TblBusesDistributionDet.BusDistID = dbo.TblBusesDistribution.ID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblBranchesData ON dbo.TblBusesDistribution.BranchID = dbo.TblBranchesData.branch_id"

StrSQL = StrSQL & "   WHERE  (1 = 1) "
      
    If val(TxtFromID.Text) <> 0 Then
            StrSQL = StrSQL & "   and  dbo.TblBusesDistribution.ID>=  " & val(TxtFromID.Text) & ""
    End If
      If val(Me.TxtToID.Text) <> 0 Then
            StrSQL = StrSQL & "   and  dbo.TblBusesDistribution.ID<=  " & val(TxtToID.Text) & ""
    End If
    
    If Me.DcbBranch4.Text <> "" And val(DcbBranch4.BoundText) <> 0 Then
            StrSQL = StrSQL & "   and  dbo.TblBusesDistribution.BranchID =  " & val(DcbBranch4.BoundText)
    End If
    If Me.DcbCars.Text <> "" And val(DcbCars.BoundText) <> 0 Then
            StrSQL = StrSQL & "   and  dbo.TblBusesDistributionDet.CarID =  " & val(DcbCars.BoundText)
    End If
      If Me.DcbTypeCar.Text <> "" And val(DcbTypeCar.BoundText) <> 0 Then
            StrSQL = StrSQL & "   and  dbo.TblCarsData.CarsTypeId =  " & val(DcbTypeCar.BoundText)
    End If
        If Me.DcbDriver.Text <> "" And val(DcbDriver.BoundText) <> 0 Then
            StrSQL = StrSQL & "   and  dbo.TblBusesDistributionDet.EmpID=  " & val(DcbDriver.BoundText)
    End If

        If Not IsNull(SDate.value) Then
            StrSQL = StrSQL & "   and  TblBusesDistribution.SDate >=  " & SQLDate(SDate.value, True)
    End If
         If Not IsNull(ToSDate.value) Then
            StrSQL = StrSQL & "   and  TblBusesDistribution.SDate <=  " & SQLDate(ToSDate.value, True)
    End If
 
    
    StrSQL = StrSQL & " GROUP BY dbo.TblBusesDistribution.ID, dbo.TblBusesDistribution.SDate, dbo.TblBusesDistribution.BranchID, dbo.TblBranchesData.branch_name,"
StrSQL = StrSQL & "                      dbo.TblBranchesData.branch_namee, dbo.TblBusesDistribution.RecordDateH, dbo.TblBusesDistribution.NoVehicle, dbo.TblBusesDistribution.Capacity,"
StrSQL = StrSQL & "                      dbo.TblBusesDistribution.OrderNo, dbo.TblBusesDistributionDet.Capacity, dbo.TblBusesDistributionDet.EmpID, dbo.TblEmployee.Emp_Name,"
StrSQL = StrSQL & "                      dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee, dbo.TblBusesDistributionDet.CarID, dbo.TblCarsData.Branch_NO, dbo.TblCarsData.OperatorN,"
StrSQL = StrSQL & "                      dbo.TblCarsData.CarsTypeId , dbo.TBLCarTypes.name, dbo.TBLCarTypes.NameE"
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
         VSFlexGrid1.Rows = VSFlexGrid1.FixedRows
         Exit Sub
    Else

        With Me.VSFlexGrid1
            .Clear flexClearScrollable, flexClearEverything
            .Rows = .FixedRows
            .Rows = rs.RecordCount + .FixedRows

            rs.MoveFirst
        
            For I = .FixedRows To .Rows - 1
            
                .TextMatrix(I, .ColIndex("Serial")) = I
                .TextMatrix(I, .ColIndex("ID")) = IIf(IsNull(rs("ID").value), "", rs("ID").value)
                .TextMatrix(I, .ColIndex("SDate")) = IIf(IsNull(rs("SDate").value), "", rs("SDate").value)
                .TextMatrix(I, .ColIndex("RecordDateH")) = IIf(IsNull(rs("RecordDateH").value), "", rs("RecordDateH").value)
                
                .TextMatrix(I, .ColIndex("OperatorN")) = IIf(IsNull(rs("OperatorN").value), "", rs("OperatorN").value)
                If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(I, .ColIndex("name")) = IIf(IsNull(rs("Name").value), "", rs("Name").value)
                .TextMatrix(I, .ColIndex("Emp_Name")) = IIf(IsNull(rs("Emp_Name").value), "", rs("Emp_Name").value)
                .TextMatrix(I, .ColIndex("branch_name")) = IIf(IsNull(rs("branch_name").value), "", rs("branch_name").value)
                Else
                .TextMatrix(I, .ColIndex("name")) = IIf(IsNull(rs("NameE").value), "", rs("NameE").value)
                .TextMatrix(I, .ColIndex("branch_name")) = IIf(IsNull(rs("branch_namee").value), "", rs("branch_namee").value)
                End If
            
                 rs.MoveNext
            Next I

            .AutoSize 0, .Cols - 1, False
        
        End With

    End If

End Sub

Private Sub GroupID4_Change()
        Dim str As String
        Set Rs_Temp = New ADODB.Recordset
        Set CompanyID4.RowSource = Rs_Temp
        If SystemOptions.UserInterface = ArabicInterface Then
      '  str = " select  ID , Name   from TblTourismCompanies where GroupID   = " & val(GroupID4.BoundText)
        Else
      '  str = " Select ID , NameE   TblTourismCompanies where GroupID  = " & val(GroupID4.BoundText)
        End If
        fill_combo CompanyID4, str
        
        CompanyID4.Refresh
End Sub
Public Function ReloadCombos()
 Dim Dcombos As ClsDataCombos
  Dim str As String
   Set Dcombos = New ClsDataCombos
   Dcombos.GetTblCarsDataGroup DcbTypeCar
  str = "  select   e.Emp_ID Emp_ID , e.Emp_Name,e.Emp_NameE   from TblEmployee e, TblEmpJobsTypes  j"
  str = str & "   Where e.JobTypeID = j.JobTypeID"
  str = str & "     and  ( j.JobTypeName like '%”«∆ﬁ%'  or j.JobTypeNamee like '%driver%')"
  fill_combo DcbDriver, str
   str = " select   id, OperatorN from TblCarsData "
     fill_combo DcbCars, str

   End Function


Private Sub SDate_Change()
If Not IsNull(SDate.value) Then
SDateH.value = ToHijriDate(SDate.value)
End If
End Sub

Private Sub SDateH_LostFocus()
      VBA.Calendar = vbCalGreg
    SDate.value = ToGregorianDate(SDateH.value)
End Sub

Private Sub ToSDate_Change()
If Not IsNull(ToSDate.value) Then
ToSDateH.value = ToHijriDate(ToSDate.value)
End If
End Sub

Private Sub ToSDateH_LostFocus()
  VBA.Calendar = vbCalGreg
    ToSDate.value = ToGregorianDate(ToSDateH.value)
End Sub

Private Sub TxtFromID_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtFromID.Text, 0)
End Sub


Private Sub TxtToID_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtToID.Text, 0)
End Sub

Private Sub VSFlexGrid1_Click()
FrmBusesDistribution.Retrive val(VSFlexGrid1.TextMatrix(VSFlexGrid1.Row, VSFlexGrid1.ColIndex("ID")))
End Sub
