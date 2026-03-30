VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form FrmEmpSalarynew 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ﬁ—Ì— —Ê« » «·„ÊŸ›Ì‰ ÿ»⁄ » «—ÌŒ"
   ClientHeight    =   7455
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   15375
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   580
   Icon            =   "FrmEmpSalary.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   7455
   ScaleWidth      =   15375
   Begin C1SizerLibCtl.C1Elastic EleMain 
      Height          =   7455
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   15375
      _cx             =   27120
      _cy             =   13150
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
      Align           =   5
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
      GridRows        =   4
      GridCols        =   4
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"FrmEmpSalary.frx":038A
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin C1SizerLibCtl.C1Tab C1Tab1 
         Height          =   6420
         Left            =   30
         TabIndex        =   4
         Top             =   30
         Width           =   15315
         _cx             =   27014
         _cy             =   11324
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arabic Transparent"
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
         ForeColor       =   -2147483630
         FrontTabColor   =   14871017
         BackTabColor    =   12648447
         TabOutlineColor =   -2147483632
         FrontTabForeColor=   -2147483630
         Caption         =   " Œ’Ì’ «·—« »| ›«’Ì· „— » „ÊŸ›|”œ«œ «·—Ê« »"
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
         Picture(0)      =   "FrmEmpSalary.frx":0410
         Picture(1)      =   "FrmEmpSalary.frx":07AA
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   5955
            Index           =   2
            Left            =   15960
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   45
            Width           =   15225
            _cx             =   26855
            _cy             =   10504
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
            Begin MSDataListLib.DataCombo DCmboEmp 
               Height          =   315
               Left            =   6060
               TabIndex        =   8
               Top             =   90
               Width           =   2295
               _ExtentX        =   4048
               _ExtentY        =   556
               _Version        =   393216
               Text            =   "DataCombo1"
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
               Index           =   1
               Left            =   8400
               RightToLeft     =   -1  'True
               TabIndex        =   9
               Top             =   90
               Width           =   1125
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   5955
            Index           =   1
            Left            =   45
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   45
            Width           =   15225
            _cx             =   26855
            _cy             =   10504
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
            Begin VB.TextBox TxtModFlg 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   3000
               RightToLeft     =   -1  'True
               TabIndex        =   60
               Text            =   "Text1"
               Top             =   120
               Visible         =   0   'False
               Width           =   375
            End
            Begin VB.CheckBox Check16 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«· ÊﬁÌ⁄"
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
               Left            =   -135
               RightToLeft     =   -1  'True
               TabIndex        =   42
               Top             =   240
               Value           =   1  'Checked
               Visible         =   0   'False
               Width           =   990
            End
            Begin VB.CheckBox Check15 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·’«›Ì"
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
               Left            =   1215
               RightToLeft     =   -1  'True
               TabIndex        =   41
               Top             =   240
               Value           =   1  'Checked
               Visible         =   0   'False
               Width           =   855
            End
            Begin VB.CheckBox Check14 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«Ã„«·Ì 2"
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
               Left            =   2025
               RightToLeft     =   -1  'True
               TabIndex        =   40
               Top             =   240
               Value           =   1  'Checked
               Visible         =   0   'False
               Width           =   990
            End
            Begin VB.CheckBox Check13 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "Ã“«¡« "
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
               Left            =   2880
               RightToLeft     =   -1  'True
               TabIndex        =   39
               Top             =   240
               Value           =   1  'Checked
               Visible         =   0   'False
               Width           =   990
            End
            Begin VB.CheckBox Check12 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "”·›"
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
               Left            =   3825
               RightToLeft     =   -1  'True
               TabIndex        =   38
               Top             =   240
               Value           =   1  'Checked
               Visible         =   0   'False
               Width           =   630
            End
            Begin VB.CheckBox Check11 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«Ã„«·Ì1"
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
               Left            =   4455
               RightToLeft     =   -1  'True
               TabIndex        =   37
               Top             =   240
               Value           =   1  'Checked
               Visible         =   0   'False
               Width           =   855
            End
            Begin VB.CheckBox Check10 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄„Ê·« "
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
               Left            =   5175
               RightToLeft     =   -1  'True
               TabIndex        =   36
               Top             =   240
               Value           =   1  'Checked
               Visible         =   0   'False
               Width           =   855
            End
            Begin VB.CheckBox Check9 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„ﬂ«›√ "
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
               Left            =   5985
               RightToLeft     =   -1  'True
               TabIndex        =   35
               Top             =   240
               Value           =   1  'Checked
               Visible         =   0   'False
               Width           =   900
            End
            Begin VB.CheckBox Check8 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«÷«›Ì"
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
               Left            =   6705
               RightToLeft     =   -1  'True
               TabIndex        =   34
               Top             =   240
               Value           =   1  'Checked
               Visible         =   0   'False
               Width           =   915
            End
            Begin VB.CheckBox Check7 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«Œ—Ï"
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
               Left            =   7425
               RightToLeft     =   -1  'True
               TabIndex        =   33
               Top             =   240
               Value           =   1  'Checked
               Visible         =   0   'False
               Width           =   780
            End
            Begin VB.CheckBox Check6 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ÿ⁄«„"
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
               Left            =   8430
               RightToLeft     =   -1  'True
               TabIndex        =   32
               Top             =   240
               Value           =   1  'Checked
               Visible         =   0   'False
               Width           =   630
            End
            Begin VB.CheckBox Check5 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " „Ê«’·« "
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
               Left            =   9015
               RightToLeft     =   -1  'True
               TabIndex        =   31
               Top             =   240
               Value           =   1  'Checked
               Visible         =   0   'False
               Width           =   990
            End
            Begin VB.CheckBox Check4 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "»œ· ”ﬂ‰"
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
               Left            =   9870
               RightToLeft     =   -1  'True
               TabIndex        =   30
               Top             =   240
               Value           =   1  'Checked
               Visible         =   0   'False
               Width           =   1080
            End
            Begin VB.CheckBox Check3 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—« » «”«”Ì"
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
               Left            =   10860
               RightToLeft     =   -1  'True
               TabIndex        =   26
               Top             =   240
               Value           =   1  'Checked
               Visible         =   0   'False
               Width           =   1260
            End
            Begin VB.CheckBox Check2 
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
               Height          =   240
               Left            =   12120
               RightToLeft     =   -1  'True
               TabIndex        =   25
               Top             =   240
               Value           =   1  'Checked
               Visible         =   0   'False
               Width           =   1395
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬂÊœ «·„ÊŸ›"
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
               Left            =   13470
               RightToLeft     =   -1  'True
               TabIndex        =   24
               Top             =   240
               Value           =   1  'Checked
               Visible         =   0   'False
               Width           =   1260
            End
            Begin VSFlex8Ctl.VSFlexGrid Grid 
               Height          =   5235
               Left            =   -225
               TabIndex        =   7
               Top             =   570
               Width           =   15180
               _cx             =   26776
               _cy             =   9234
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
               Rows            =   50
               Cols            =   33
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"FrmEmpSalary.frx":0B44
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
            Begin ALLButtonS.ALLButton ALLButton2 
               Height          =   495
               Left            =   225
               TabIndex        =   58
               Top             =   0
               Width           =   2205
               _ExtentX        =   3889
               _ExtentY        =   873
               BTYPE           =   2
               TX              =   "«‰‘«¡ ﬁÌœ «·«” Õﬁ«ﬁ"
               ENAB            =   -1  'True
               BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               COLTYPE         =   1
               FOCUSR          =   -1  'True
               BCOL            =   15790320
               BCOLO           =   15790320
               FCOL            =   0
               FCOLO           =   0
               MCOL            =   12632256
               MPTR            =   1
               MICON           =   "FrmEmpSalary.frx":1028
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   5955
            Index           =   4
            Left            =   16260
            TabIndex        =   44
            TabStop         =   0   'False
            Top             =   45
            Width           =   15225
            _cx             =   26855
            _cy             =   10504
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
            Begin VB.TextBox TxtChequeNumber 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   5175
               RightToLeft     =   -1  'True
               TabIndex        =   53
               Top             =   90
               Width           =   1665
            End
            Begin VB.ComboBox CboPaymentType 
               Height          =   330
               ItemData        =   "FrmEmpSalary.frx":1044
               Left            =   11265
               List            =   "FrmEmpSalary.frx":104E
               RightToLeft     =   -1  'True
               TabIndex        =   50
               Top             =   120
               Width           =   1125
            End
            Begin VB.CheckBox Check17 
               Alignment       =   1  'Right Justify
               Caption         =   " ÕœÌœ «·ﬂ·"
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
               Left            =   13290
               RightToLeft     =   -1  'True
               TabIndex        =   46
               Top             =   120
               Width           =   1125
            End
            Begin VSFlex8Ctl.VSFlexGrid Grid1 
               Height          =   5220
               Left            =   -225
               TabIndex        =   45
               Top             =   585
               Width           =   15180
               _cx             =   26776
               _cy             =   9208
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
               Rows            =   50
               Cols            =   29
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"FrmEmpSalary.frx":105D
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
            Begin MSDataListLib.DataCombo DcboBankName 
               Height          =   330
               Left            =   7800
               TabIndex        =   47
               Top             =   90
               Width           =   2220
               _ExtentX        =   3916
               _ExtentY        =   582
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcboBox 
               Height          =   330
               Left            =   7800
               TabIndex        =   48
               Top             =   120
               Width           =   2220
               _ExtentX        =   3916
               _ExtentY        =   582
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSComCtl2.DTPicker DtpChequeDueDate 
               Height          =   315
               Left            =   2070
               TabIndex        =   54
               Top             =   90
               Width           =   1710
               _ExtentX        =   3016
               _ExtentY        =   556
               _Version        =   393216
               Format          =   91422721
               CurrentDate     =   39614
            End
            Begin ALLButtonS.ALLButton ALLButton3 
               Height          =   435
               Left            =   135
               TabIndex        =   59
               Top             =   0
               Width           =   1755
               _ExtentX        =   3096
               _ExtentY        =   767
               BTYPE           =   2
               TX              =   "”œ«œ «·—« »"
               ENAB            =   -1  'True
               BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               COLTYPE         =   1
               FOCUSR          =   -1  'True
               BCOL            =   15790320
               BCOLO           =   15790320
               FCOL            =   0
               FCOLO           =   0
               MCOL            =   12632256
               MPTR            =   1
               MICON           =   "FrmEmpSalary.frx":14C0
               UMCOL           =   -1  'True
               SOFT            =   0   'False
               PICPOS          =   0
               NGREY           =   0   'False
               FX              =   0
               HAND            =   0   'False
               CHECK           =   0   'False
               VALUE           =   0   'False
            End
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   " «—ÌŒ «·«” Õﬁ«ﬁ"
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
               Left            =   3870
               RightToLeft     =   -1  'True
               TabIndex        =   55
               Top             =   120
               Width           =   1170
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "—ﬁ„ «·‘Ìﬂ"
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
               Left            =   6975
               RightToLeft     =   -1  'True
               TabIndex        =   52
               Top             =   120
               Width           =   690
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«·Œ“Ì‰…"
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
               Left            =   10080
               RightToLeft     =   -1  'True
               TabIndex        =   51
               Top             =   120
               Width           =   975
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "ÿ—Ìﬁ… «·œ›⁄"
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
               Left            =   12120
               RightToLeft     =   -1  'True
               TabIndex        =   49
               Top             =   120
               Width           =   1080
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   960
         Index           =   0
         Left            =   30
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   6465
         Width           =   15315
         _cx             =   27014
         _cy             =   1693
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
         Begin VB.ComboBox Combo1 
            Height          =   330
            ItemData        =   "FrmEmpSalary.frx":14DC
            Left            =   3480
            List            =   "FrmEmpSalary.frx":14E6
            RightToLeft     =   -1  'True
            TabIndex        =   56
            Top             =   480
            Width           =   2655
         End
         Begin MSDataListLib.DataCombo Dcemp 
            Height          =   315
            Left            =   7695
            TabIndex        =   18
            Top             =   120
            Width           =   2160
            _ExtentX        =   3810
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
         Begin VB.Frame Fra 
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ ”«⁄«  «·‘Â—"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   885
            Left            =   11490
            RightToLeft     =   -1  'True
            TabIndex        =   16
            Top             =   30
            Width           =   1500
            Begin VB.TextBox TxtMonthHours 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   180
               RightToLeft     =   -1  'True
               TabIndex        =   17
               Text            =   "176"
               Top             =   330
               Width           =   705
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   915
            Index           =   3
            Left            =   13035
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   0
            Width           =   2265
            _cx             =   3995
            _cy             =   1614
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
            Caption         =   "≈Œ Ì«— «· «—ÌŒ"
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
            Begin VB.ComboBox CboYear 
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
               Left            =   270
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   13
               Top             =   540
               Width           =   1305
            End
            Begin VB.ComboBox CmbMonth 
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
               Left            =   270
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   11
               Top             =   210
               Width           =   1305
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "”‰…"
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
               Left            =   1590
               RightToLeft     =   -1  'True
               TabIndex        =   14
               Top             =   570
               Width           =   555
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "‘Â—"
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
               Index           =   0
               Left            =   1665
               RightToLeft     =   -1  'True
               TabIndex        =   12
               Top             =   270
               Width           =   375
            End
         End
         Begin ImpulseButton.ISButton CmdOk 
            Height          =   345
            Left            =   2325
            TabIndex        =   2
            Top             =   495
            Visible         =   0   'False
            Width           =   930
            _ExtentX        =   1640
            _ExtentY        =   609
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "⁄—÷  "
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
            ButtonImage     =   "FrmEmpSalary.frx":14FC
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton CmdExit 
            Height          =   345
            Left            =   0
            TabIndex        =   3
            Top             =   135
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   609
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
            ButtonImage     =   "FrmEmpSalary.frx":1896
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton CmdPrint 
            Height          =   390
            Left            =   780
            TabIndex        =   15
            Top             =   120
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   688
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
            ButtonImage     =   "FrmEmpSalary.frx":1C30
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin MSDataListLib.DataCombo Dcdep 
            Height          =   315
            Left            =   7725
            TabIndex        =   20
            Top             =   480
            Width           =   2160
            _ExtentX        =   3810
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
            Left            =   3510
            TabIndex        =   22
            Top             =   120
            Width           =   2760
            _ExtentX        =   4868
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
         Begin ImpulseButton.ISButton ISButton2 
            Height          =   465
            Left            =   9885
            TabIndex        =   28
            Top             =   600
            Visible         =   0   'False
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   820
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "⁄—÷ 2"
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
            ButtonImage     =   "FrmEmpSalary.frx":1FCA
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton ISButton3 
            Height          =   345
            Left            =   7650
            TabIndex        =   29
            Top             =   735
            Visible         =   0   'False
            Width           =   810
            _ExtentX        =   1429
            _ExtentY        =   609
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "⁄—÷ 3 "
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
            ButtonImage     =   "FrmEmpSalary.frx":2364
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ALLButtonS.ALLButton ALLButton1 
            Height          =   375
            Left            =   1800
            TabIndex        =   43
            Top             =   120
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            BTYPE           =   2
            TX              =   " ⁄œÌ· «·‘«‘…"
            ENAB            =   -1  'True
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   1
            FOCUSR          =   -1  'True
            BCOL            =   15790320
            BCOLO           =   15790320
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmEmpSalary.frx":26FE
            PICN            =   "FrmEmpSalary.frx":271A
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‘ﬂ· «·ÿ»«⁄…"
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
            Index           =   6
            Left            =   6240
            RightToLeft     =   -1  'True
            TabIndex        =   57
            Top             =   480
            Width           =   1170
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·„‘—Ê⁄"
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
            Index           =   5
            Left            =   6225
            RightToLeft     =   -1  'True
            TabIndex        =   23
            Top             =   120
            Width           =   1170
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ ««·ﬁ”„"
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
            Index           =   4
            Left            =   7050
            RightToLeft     =   -1  'True
            TabIndex        =   21
            Top             =   480
            Width           =   4020
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
            Index           =   3
            Left            =   9360
            RightToLeft     =   -1  'True
            TabIndex        =   19
            Top             =   120
            Width           =   1710
         End
      End
   End
   Begin ImpulseButton.ISButton ISButton1 
      Height          =   345
      Left            =   3360
      TabIndex        =   27
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
      ButtonImage     =   "FrmEmpSalary.frx":2BC6
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
End
Attribute VB_Name = "FrmEmpSalarynew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cSearchDCombo As clsDCboSearch
Dim BKGrndPic As ClsBackGroundPic
Dim net_value As Double
Dim net_value1 As Double

Private Declare Function TextOut _
                Lib "gdi32" _
                Alias "TextOutA" (ByVal hDC As Long, _
                                  ByVal X As Long, _
                                  ByVal Y As Long, _
                                  ByVal lpString As String, _
                                  ByVal nCount As Long) As Long

Public Sub YearMonth()

    Dim i As Integer
    Dim IntDefIndex As Integer

    CmbMonth.Clear

    For i = 1 To 12
        CmbMonth.AddItem MonthName(i)
    Next

    CmbMonth.ListIndex = Month(Date) - 1
    CboYear.Clear

    For i = 2000 To 2050
        CboYear.AddItem i

        If i = year(Date) Then
            IntDefIndex = CboYear.NewIndex
        End If

    Next

    CboYear.ListIndex = IntDefIndex

End Sub

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
    Dim i As Integer
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

    If check_previous_dev(CboYear.text, CmbMonth.ListIndex) Then
        MsgBox " „ «‰‘«¡ ﬁÌœ „”»ﬁ« ·Â–« «·‘Â—", vbCritical
        Exit Function
    End If
        
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
        
    Msg = "ﬁÌœ «” Õﬁ«ﬁ —Ê« » «·„ÊŸ›Ì‰ ⁄‰ ‘Â— " & CmbMonth.text & "   ”‰… " & CboYear.text

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
    
    rs("salary").value = CboYear.text & CmbMonth.ListIndex

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
    Dim i As Integer
    Dim LngDevID As Long
    Dim Msg As String
    Dim Account_Code_dynamic As String
    Dim Account_Code_dynamic1 As String
        
    Dim Employee_account As String
    Dim StrAccountCode As String
    Dim X As Integer
    Dim rs As ADODB.Recordset
        
    If check_previous_dev(CboYear.text, CmbMonth.text) Then
        MsgBox " „ «‰‘«¡ ﬁÌœ „”»ﬁ« ·Â–« «·‘Â—", vbCritical
        Exit Function
    End If
        
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
    Msg = "ﬁÌœ «” Õﬁ«ﬁ —Ê« » «·„ÊŸ›Ì‰ ⁄‰ ‘Â— " & CmbMonth.text & "   ”‰… " & CboYear.text
        
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
    rs("m_year").value = CboYear.text
    rs("m_month").value = CmbMonth.text
  
    rs.update
 
    MsgBox " „ «‰‘«¡ «·ﬁÌœ", vbInformation
    create_report_data

    DoEvents

    Exit Function
ErrTrap:
    MsgBox "ÕœÀ Œÿ√ «À‰«¡ Õ›Ÿ «·»Ì«‰« ", vbExclamation
  
End Function

Private Sub ALLButton2_Click()
    DCEmP.text = ""
    dcdep.text = ""
    dcproject.text = ""
    FillGridWithData

    DoEvents
    Create_dev
    CMDOK_Click
End Sub

Private Sub ALLButton3_Click()

    With Grid1

        If .Rows = 3 And Not IsNumeric(.TextMatrix(1, .ColIndex("Emp_code"))) Then
            Exit Sub
        End If

    End With

    Dim i As Integer
    Dim LngDevID As Long
    Dim Msg As String
    Dim depit_side As String
    Dim credit_side As String
    Dim total_value As Single
    On Error Resume Next

    If Me.CboPaymentType.ListIndex = -1 Then
        Msg = "ÌÃ» ≈Œ Ì«— ÿ—Ìﬁ… «·œ›⁄ ...!!!"
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        CboPaymentType.SetFocus
        Exit Sub
    End If

    If Me.CboPaymentType.ListIndex = 0 Then
        If Trim(Me.DcboBox.BoundText) = "" Then
            Msg = "ÌÃ» ≈Œ Ì«— «·Œ“‰…..!!"
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            DcboBox.SetFocus
            SendKeys "{F4}"
            Exit Sub
        End If

    ElseIf Me.CboPaymentType.ListIndex = 1 Then

        If Me.DcboBankName.BoundText = "" Then
            Msg = "ÌÃ» ≈Œ Ì«— «·»‰ﬂ...!!"
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            DcboBankName.SetFocus
            SendKeys "{F4}"
            Exit Sub
        End If

        If Trim$(Me.TxtChequeNumber.text) = "" Then
            Msg = "ÌÃ» ﬂ «»… —ﬁ„ «·‘Ìﬂ...!!"
            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            TxtChequeNumber.SetFocus
            Exit Sub
        End If

        '       If DateDiff("d", Me.DtpChequeDueDate.value, Date) >= 0 Then
        '           Msg = " «—ÌŒ ≈” Õﬁ«ﬁ «·‘Ìﬂ €Ì— ’ÕÌÕ...!!"
        '           MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        '           DtpChequeDueDate.SetFocus
        '           SendKeys "{F4}"
        '           Exit Sub
        '       End If
    End If

    credit_side = ""
    depit_side = ""
    total_value = 0

    If Me.CboPaymentType.ListIndex = 0 Then

        credit_side = ModAccounts.GetMyAccountCode("TblBoxesData", "BoxID", val(Me.DcboBox.BoundText))

        If credit_side = "" Then
            MsgBox "Â‰«ﬂ Œÿ√ ›Ì —ﬁ„ Õ”«» «·Œ“Ì‰…": Exit Sub
        End If
                 
    Else

        If Me.CboPaymentType.ListIndex = 1 Then
    
            Dim rsbank As New ADODB.Recordset
            Set rsbank = New ADODB.Recordset
            rsbank.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
       
            If Not (rsbank.EOF Or rsbank.BOF) Then
                If rsbank!banks_Accounts = True Then
                  
                    credit_side = get_bank_Account(val(Me.DcboBankName.BoundText), "Account_Code1")
                Else
                    credit_side = get_bank_Account(val(Me.DcboBankName.BoundText), "Account_Code")
                End If
                 
                If credit_side = "" Then
                    MsgBox "Â‰«ﬂ Œÿ√ ›Ì —ﬁ„ Õ”«» «·»‰ﬂ": Exit Sub
                End If
            End If
        
        End If
    End If

    Dim StrSQL As String
    Dim notes_id As String
    Dim notes_serial As String
    Dim rs As New ADODB.Recordset
 
    StrSQL = "select * From Notes where NoteType=5 order by NoteID"

    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    notes_id = CStr(new_id("Notes", "NoteID", "", True))
    notes_serial = CStr(new_id("Notes", "NoteSerial", "", True, "NoteType=66"))
    Msg = "ﬁÌœ ”œ«œ —« » ⁄‰ ‘Â—   " & CmbMonth.text & "     ·”‰… " & CboYear.text
 
    With Grid1

        For i = .FixedRows To .Rows - 2
 
            If .TextMatrix(i, .ColIndex("payed")) <> "" Then
                total_value = total_value + .TextMatrix(i, .ColIndex("EmpTotalNet"))
            End If

        Next i

    End With
 
    rs.AddNew
    rs("NoteID").value = notes_id
    rs("NoteSerial").value = notes_serial '
    rs("Note_Value").value = total_value
    rs("Remark").value = Msg
    rs("salary").value = CboYear.text & CmbMonth.ListIndex
    rs("NoteType").value = 5
    rs("NoteDate").value = Date
    rs("UserID").value = user_id

    '
    If Me.CboPaymentType.ListIndex = 0 Then
        rs("BoxID").value = val(DcboBox.BoundText)
        rs("BankID").value = Null
        rs("ChqueNum").value = Null
        rs("DueDate").value = Null
        rs("NoteCashingType").value = 0
    ElseIf Me.CboPaymentType.ListIndex = 1 Then
        rs("BoxID").value = Null
        rs("BankID").value = val(Me.DcboBankName.BoundText)
        rs("ChqueNum").value = Trim$(Me.TxtChequeNumber.text)
        rs("DueDate").value = Me.DtpChequeDueDate.value
        rs("NoteCashingType").value = 1
    End If

    rs.update
    
    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
    total_value = 0
                
    With Grid1

        For i = .FixedRows To .Rows - 2
 
            If .TextMatrix(i, .ColIndex("payed")) <> "" Then
                total_value = total_value + .TextMatrix(i, .ColIndex("EmpTotalNet"))
                depit_side = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_code"))), "Account_Code1")
            
                If ModAccounts.AddNewDev(LngDevID, i + 1, depit_side, .TextMatrix(i, .ColIndex("EmpTotalNet")), 0, Msg, val(notes_id), , , , Date, user_id) = False Then
                    GoTo ErrTrap
                End If
                   
            End If

        Next i

    End With

    If total_value > 0 Then
        If ModAccounts.AddNewDev(LngDevID, 1, credit_side, total_value, 1, Msg, val(notes_id), , , , Date, user_id, 200) = False Then
            GoTo ErrTrap
        End If
    End If

    MsgBox " „ «‰‘«¡ ”‰œ «·”œ«œ", vbInformation

    With Grid1

        For i = .FixedRows To .Rows - 2
         
            If .TextMatrix(i, .ColIndex("payed")) <> "" Then
                If Change_filed_value(val(.TextMatrix(i, .ColIndex("id"))), "id", "Payed", "emp_salary", 1) Then
                End If
            End If

        Next i

    End With
        
    FillGridWithData2
    Exit Sub
ErrTrap:
    'Dim StrSQL As String
    StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Double_Entry_Vouchers_ID=" & LngDevID
    Cn.Execute StrSQL, , adExecuteNoRecords
    MsgBox "ÕœÀ Œÿ√ «À‰«¡ Õ›Ÿ «·ﬁÌœ ", vbCritical

End Sub

Private Sub CboPayMentType_Change()

    If Me.CboPaymentType.ListIndex = 0 Then
     
        Me.DcboBox.Visible = True
        Me.DcboBankName.Visible = False
        Label2.Caption = " «·Œ“Ì‰…"
        Me.TxtChequeNumber.Enabled = False
      
        Me.DtpChequeDueDate.Enabled = False
    ElseIf Me.CboPaymentType.ListIndex = 1 Then
        Me.DcboBox.Visible = False
        Me.DcboBankName.Visible = True
        Me.TxtChequeNumber.Enabled = True
        Me.DtpChequeDueDate.Enabled = True
        Label2.Caption = " «·»‰ﬂ"
    Else
        Me.DcboBankName.Visible = False
        Me.DcboBox.Visible = False
        Me.DcboBankName.Enabled = False
        Me.TxtChequeNumber.Enabled = False
        Me.DtpChequeDueDate.Enabled = False
    End If

End Sub

Private Sub CboPayMentType_Click()
    CboPayMentType_Change
End Sub

Private Sub CboYear_Click()
    CMDOK_Click
End Sub

Private Sub Check1_Click()

    If Check1.value = vbChecked Then
        Grid.ColHidden(Grid.ColIndex("Emp_Code")) = False
    Else
        Grid.ColHidden(Grid.ColIndex("Emp_Code")) = True

    End If

End Sub

Private Sub Check10_Click()

    If Check10.value = vbChecked Then
        Grid.ColHidden(Grid.ColIndex("SalesCom")) = False
    Else
        Grid.ColHidden(Grid.ColIndex("SalesCom")) = True

    End If

End Sub

Private Sub Check11_Click()

    If Check11.value = vbChecked Then
        Grid.ColHidden(Grid.ColIndex("total1")) = False
    Else
        Grid.ColHidden(Grid.ColIndex("total1")) = True

    End If

End Sub

Private Sub Check12_Click()

    If Check12.value = vbChecked Then
        Grid.ColHidden(Grid.ColIndex("TotalAdvance")) = False
    Else
        Grid.ColHidden(Grid.ColIndex("TotalAdvance")) = True

    End If

End Sub

Private Sub Check13_Click()

    If Check13.value = vbChecked Then
        Grid.ColHidden(Grid.ColIndex("TotalDiscount")) = False
    Else
        Grid.ColHidden(Grid.ColIndex("TotalDiscount")) = True

    End If

End Sub

Private Sub Check14_Click()

    If Check14.value = vbChecked Then
        Grid.ColHidden(Grid.ColIndex("total2")) = False
    Else
        Grid.ColHidden(Grid.ColIndex("total2")) = True

    End If

End Sub

Private Sub Check15_Click()

    If Check15.value = vbChecked Then
        Grid.ColHidden(Grid.ColIndex("EmpTotalNet")) = False
    Else
        Grid.ColHidden(Grid.ColIndex("EmpTotalNet")) = True

    End If

End Sub

Private Sub Check16_Click()

    If Check16.value = vbChecked Then
        Grid.ColHidden(Grid.ColIndex("sgn")) = False
    Else
        Grid.ColHidden(Grid.ColIndex("sgn")) = True

    End If

End Sub

Private Sub Check17_Click()
    Dim i As Integer

    If Check17.value = vbChecked Then

        With Me.Grid1
 
            For i = 1 To .Rows - 2
        
                .TextMatrix(i, .ColIndex("payed")) = True
            Next i

        End With

    Else

        With Me.Grid1

            For i = 1 To .Rows - 2
        
                .TextMatrix(i, .ColIndex("payed")) = False
            Next i

        End With

    End If

End Sub

Private Sub Check2_Click()

    If Check2.value = vbChecked Then
        Grid.ColHidden(Grid.ColIndex("Emp_Name")) = False
    Else
        Grid.ColHidden(Grid.ColIndex("Emp_Name")) = True

    End If

End Sub

Private Sub Check3_Click()

    If Check3.value = vbChecked Then
        Grid.ColHidden(Grid.ColIndex("Emp_Salary")) = False
    Else
        Grid.ColHidden(Grid.ColIndex("Emp_Salary")) = True

    End If

End Sub

Private Sub Check4_Click()

    If Check4.value = vbChecked Then
        Grid.ColHidden(Grid.ColIndex("Emp_Salary_sakn")) = False
    Else
        Grid.ColHidden(Grid.ColIndex("Emp_Salary_sakn")) = True

    End If

End Sub

Private Sub Check5_Click()

    If Check5.value = vbChecked Then
        Grid.ColHidden(Grid.ColIndex("Emp_Salary_bus")) = False
    Else
        Grid.ColHidden(Grid.ColIndex("Emp_Salary_bus")) = True

    End If

End Sub

Private Sub Check6_Click()

    If Check6.value = vbChecked Then
        Grid.ColHidden(Grid.ColIndex("Emp_Salary_food")) = False
    Else
        Grid.ColHidden(Grid.ColIndex("Emp_Salary_food")) = True

    End If

End Sub

Private Sub Check7_Click()

    If Check7.value = vbChecked Then
        Grid.ColHidden(Grid.ColIndex("Emp_Salary_others")) = False
    Else
        Grid.ColHidden(Grid.ColIndex("Emp_Salary_others")) = True

    End If

End Sub

Private Sub Check8_Click()

    If Check8.value = vbChecked Then
        Grid.ColHidden(Grid.ColIndex("OverTimePrice")) = False
    Else
        Grid.ColHidden(Grid.ColIndex("OverTimePrice")) = True

    End If

End Sub

Private Sub Check9_Click()

    If Check9.value = vbChecked Then
        Grid.ColHidden(Grid.ColIndex("Mokafea")) = False
    Else
        Grid.ColHidden(Grid.ColIndex("Mokafea")) = True

    End If

End Sub

Private Sub CmbMonth_Click()
    CMDOK_Click
    'FillGridWithData
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CMDOK_Click()
    FillGridWithData2
    FillGridWithData

End Sub

Function create_report_data()
    On Error Resume Next
    Dim StrSQL As String
    Dim i As Integer
    StrSQL = "Delete   emp_salary where m_year='" & CboYear.text & "' and m_month='" & CmbMonth.text & "'"
    Cn.Execute StrSQL, , adExecuteNoRecords
    
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.Open "emp_salary", Cn, adOpenStatic, adLockOptimistic, adCmdTable

    With Grid

        For i = .FixedRows To .Rows - 2
   
            rs.AddNew
 
            rs("Emp_Code").value = .TextMatrix(i, .ColIndex("Emp_Code"))
            rs("Emp_Name").value = .TextMatrix(i, .ColIndex("Emp_Name"))
            rs("Emp_Salary").value = .TextMatrix(i, .ColIndex("Emp_Salary"))
            rs("Emp_Salary_sakn").value = .TextMatrix(i, .ColIndex("Emp_Salary_sakn"))
            rs("Emp_Salary_bus").value = .TextMatrix(i, .ColIndex("Emp_Salary_bus"))
            rs("Emp_Salary_food").value = .TextMatrix(i, .ColIndex("Emp_Salary_food"))
            rs("Emp_Salary_mob").value = .TextMatrix(i, .ColIndex("Emp_Salary_mob"))
            rs("Emp_Salary_mang").value = .TextMatrix(i, .ColIndex("Emp_Salary_mang"))
            rs("Emp_Salary_others").value = .TextMatrix(i, .ColIndex("Emp_Salary_others"))
            rs("OverTimePrice").value = .TextMatrix(i, .ColIndex("OverTimePrice"))
            rs("Mokafea").value = .TextMatrix(i, .ColIndex("Mokafea"))
            rs("SalesCom").value = .TextMatrix(i, .ColIndex("SalesCom"))
            rs("total1").value = .TextMatrix(i, .ColIndex("total1"))
            rs("TotalAdvance").value = .TextMatrix(i, .ColIndex("TotalAdvance"))
            rs("TotalDiscount").value = .TextMatrix(i, .ColIndex("TotalDiscount"))
            rs("total2").value = .TextMatrix(i, .ColIndex("total2"))
            rs("EmpTotalNet").value = .TextMatrix(i, .ColIndex("EmpTotalNet"))
            rs("m_year").value = CboYear.text
            rs("m_month").value = CmbMonth.text
            rs("DepartmentID").value = .TextMatrix(i, .ColIndex("dep"))
            rs("project_id").value = .TextMatrix(i, .ColIndex("project"))
 
            ',,
    
            rs.update
   
        Next i

    End With

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

    If Combo1.ListIndex > -1 Then
        If Combo1.ListIndex = 0 Then
            ISButton2_Click
        Else

            If Combo1.ListIndex = 1 Then
                ISButton3_Click
            End If
        End If

    End If

End Sub

Private Sub Dcdep_Click(Area As Integer)
    CMDOK_Click
End Sub

Private Sub Dcedara_Click(Area As Integer)
    CMDOK_Click
End Sub

Private Sub Dcemp_Click(Area As Integer)
    CMDOK_Click
End Sub

Private Sub DCmboEmp_Click(Area As Integer)
    FillGridWithData
End Sub

Function SHow_grig_col()
    Dim rs2 As ADODB.Recordset
    Set rs2 = New ADODB.Recordset
    rs2.Open "Employee_salary_col", Cn, adOpenStatic, adLockOptimistic, adCmdTable

    With Grid
     
        If rs2("s1").value = True Then
            .ColHidden(.ColIndex("Emp_Code")) = False
        Else
            .ColHidden(.ColIndex("Emp_Code")) = True
        End If
    
        If rs2("s2").value = True Then
            .ColHidden(.ColIndex("Emp_Name")) = False
        Else
            .ColHidden(.ColIndex("Emp_Name")) = True
        End If
   
        If rs2("s3").value = True Then
            .ColHidden(.ColIndex("Emp_Salary")) = False
        Else
            .ColHidden(.ColIndex("Emp_Salary")) = True
        End If
        
        If rs2("s4").value = True Then
            .ColHidden(.ColIndex("Emp_Salary_sakn")) = False
        Else
            .ColHidden(.ColIndex("Emp_Salary_sakn")) = True
        End If
       
        If rs2("s5").value = True Then
            .ColHidden(.ColIndex("Emp_Salary_bus")) = False
        Else
            .ColHidden(.ColIndex("Emp_Salary_bus")) = True
        End If
        
        If rs2("s6").value = True Then
            .ColHidden(.ColIndex("Emp_Salary_food")) = False
        Else
            .ColHidden(.ColIndex("Emp_Salary_food")) = True
        End If
    
        If rs2("s7").value = True Then
            .ColHidden(.ColIndex("Emp_Salary_mob")) = False
        Else
            .ColHidden(.ColIndex("Emp_Salary_mob")) = True
        End If
        
        If rs2("s8").value = True Then
            .ColHidden(.ColIndex("Emp_Salary_mang")) = False
        Else
            .ColHidden(.ColIndex("Emp_Salary_mang")) = True
        End If
              
        If rs2("s9").value = True Then
            .ColHidden(.ColIndex("Emp_Salary_others")) = False
        Else
            .ColHidden(.ColIndex("Emp_Salary_others")) = True
        End If
                  
        If rs2("s10").value = True Then
            .ColHidden(.ColIndex("OverTimePrice")) = False
        Else
            .ColHidden(.ColIndex("OverTimePrice")) = True
        End If
                  
        If rs2("s11").value = True Then
            .ColHidden(.ColIndex("Mokafea")) = False
        Else
            .ColHidden(.ColIndex("Mokafea")) = True
        End If
                 
        If rs2("s12").value = True Then
            .ColHidden(.ColIndex("SalesCom")) = False
        Else
            .ColHidden(.ColIndex("SalesCom")) = True
        End If
                 
        If rs2("s13").value = True Then
            .ColHidden(.ColIndex("total1")) = False
        Else
            .ColHidden(.ColIndex("total1")) = True
        End If
                
        If rs2("s14").value = True Then
            .ColHidden(.ColIndex("TotalAdvance")) = False
        Else
            .ColHidden(.ColIndex("TotalAdvance")) = True
        End If
                
        If rs2("s15").value = True Then
            .ColHidden(.ColIndex("TotalDiscount")) = False
        Else
            .ColHidden(.ColIndex("TotalDiscount")) = True
        End If
                  
        If rs2("s16").value = True Then
            .ColHidden(.ColIndex("total2")) = False
        Else
            .ColHidden(.ColIndex("total2")) = True
        End If
                 
        If rs2("s17").value = True Then
            .ColHidden(.ColIndex("EmpTotalNet")) = False
        Else
            .ColHidden(.ColIndex("EmpTotalNet")) = True
        End If
                  
        If rs2("s18").value = True Then
            .ColHidden(.ColIndex("sgn")) = False
        Else
            .ColHidden(.ColIndex("sgn")) = True
        End If
     
    End With

End Function

Private Sub dcproject_Click(Area As Integer)
    CMDOK_Click
End Sub

Private Sub Form_Load()
    Dim My_SQL As String
 
    My_SQL = "select Emp_Code,Emp_Name From TblEmployee "
    fill_combo DCEmP, My_SQL

    My_SQL = "select DeparmentID,DepartmentName From TblEmpDepartments "
    fill_combo dcdep, My_SQL

    My_SQL = " select id,Project_name from projects"
    fill_combo dcproject, My_SQL

    Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos
    Dcombos.GetEmployees Me.DCmboEmp, True
    Set cSearchDCombo = New clsDCboSearch
    Set cSearchDCombo.Client = DCmboEmp

    Dcombos.GetBoxes Me.DcboBox
    Dcombos.GetBanks Me.DcboBankName

    Set BKGrndPic = New ClsBackGroundPic

    With Me.Grid
        .Rows = 1
        .ExplorerBar = flexExSortShowAndMove
        .RowHeightMin = 300
        .ExtendLastCol = True
        .WallPaper = BKGrndPic.Picture
        .AutoSize 0, .Cols - 1, False
    End With

    Me.C1Tab1.TabVisible(1) = False
    'SetDtpickerDate Me.DtpFrom
    'SetDtpickerDate Me.DtpTO

    YearMonth
    'SHow_grig_col
    'Resize_Form Me, True
End Sub

Public Sub FillGridWithData()
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim rs2 As ADODB.Recordset
    Set rs2 = New ADODB.Recordset
    Dim j As Integer

    Dim sql As String
    Dim i As Integer
    sql = "Select * from mofrdat "
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then Exit Sub
  
    For i = 1 To Rs3.RecordCount
 
        With FrmEmpSalary.Grid
    
            .ColHidden(.ColIndex("A" & Rs3("mofrad_code").value)) = False
            .TextMatrix(0, .ColIndex("A" & Rs3("mofrad_code").value)) = IIf(IsNull(Rs3.Fields("mofrad_name").value), "", Rs3.Fields("mofrad_name").value)
    
            .AutoSize 0, .Cols - 1, False
        End With

        Rs3.MoveNext
    Next i

    Rs3.Close

    sql = "Select * from TblEmployee "
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then Exit Sub
 
    With FrmEmpSalary.Grid

        '.Rows = 2
        ' .Clear flexClearScrollable
        If Rs3.RecordCount > 0 Then
            .Rows = Rs3.RecordCount + 1
            Rs3.MoveFirst
         
            For i = 1 To Rs3.RecordCount
                .TextMatrix(i, .ColIndex("Emp_Code")) = IIf(IsNull(Rs3.Fields("Emp_Code").value), "", Rs3.Fields("Emp_Code").value)
                .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(Rs3.Fields("Emp_Name").value), "", Rs3.Fields("Emp_Name").value)
                       
                sql = "Select * from emp_mofrad where Emp_ID=" & Rs3.Fields("Emp_ID").value
 
                rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
                        
                If rs2.RecordCount > 0 Then

                    For j = 1 To rs2.RecordCount
                        .TextMatrix(i, .ColIndex("A" & rs2("mofrad_code").value)) = IIf(IsNull(rs2.Fields("value").value), 0, rs2.Fields("value").value)
                        rs2.MoveNext
                    Next j

                    rs2.Close
                        
                End If
                       
                Rs3.MoveNext
            Next i
            
            '.Cell(flexcpBackColor, .Rows - 1, 1, .Rows - 1, .Cols - 1) = vbYellow
            ' .Cell(flexcpFontBold, .Rows - 1, 1, .Rows - 1, .Cols - 1) = True
            ' .Cell(flexcpFontSize, .Rows - 1, 1, .Rows - 1, .Cols - 1) = 10
            ' .Cell(flexcpFontName, .Rows - 1, 1, .Rows - 1, .Cols - 1) = "Tahoma"
            .AutoSize 0, .Cols - 1, False
        End If

    End With
 
    Rs3.Close

    Exit Sub
 
    Dim rs As ADODB.Recordset
    'Dim Rs2 As ADODB.Recordset
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
    'If DateDiff("d", Me.DtpFrom.Value, Me.DtpTO.Value, vbSaturday) < 0 Then
    '    Exit Sub
    'End If
    Set rs = New ADODB.Recordset
    Set rs2 = New ADODB.Recordset

    If Me.CmbMonth.ListIndex = -1 Then Exit Sub
    If Me.CboYear.ListIndex = -1 Then Exit Sub

    If val(Me.TxtMonthHours.text) = 0 Then
        Msg = "ÌÃ» ≈œŒ«· ⁄œœ ”«⁄«  «·⁄„· ·Â–« «·‘Â—"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

    IntYear = val(Me.CboYear.text)
    IntMonth = Me.CmbMonth.ListIndex + 1

    If SystemOptions.SysDataBaseType = SQLServerDataBase Then
        Dim id As String
        My_SQL = " Select Emp_ID,Emp_Code,Emp_Name,DepartmentID,project_id "
        My_SQL = My_SQL + ",IsNUll(Emp_Salary,0)as Emp_Salary,IsNUll(Emp_Salary_sakn,0)as Emp_Salary_sakn,IsNUll(Emp_Salary_bus,0)as Emp_Salary_bus,IsNUll(Emp_Salary_food,0)as Emp_Salary_food,IsNUll(Emp_Salary_others,0)as Emp_Salary_others,IsNUll(Emp_Salary_mob,0)as Emp_Salary_mob,IsNUll(Emp_Salary_mang,0)as Emp_Salary_mang,  "
        My_SQL = My_SQL + "IsNUll( TotalDiscount,0)as TotalDiscount,"
        My_SQL = My_SQL + "IsNUll(TotalMokafea, 0) As TotalMokafea"
        My_SQL = My_SQL + ""
        My_SQL = My_SQL + ",(IsNUll(Emp_Salary,0)+IsNUll( TotalMokafea,0))-"
        My_SQL = My_SQL + "(IsNUll(TotalDiscount,0)) as EmpTotalNet "
    
        My_SQL = My_SQL + " From "
        My_SQL = My_SQL + "("
        My_SQL = My_SQL + "SELECT TOP 100 PERCENT  dbo.TblEmployee.project_id, dbo.TblEmployee.DepartmentID , dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Salary_sakn, dbo.TblEmployee.Emp_Salary_bus, dbo.TblEmployee.Emp_Salary_food, dbo.TblEmployee.Emp_Salary_others, dbo.TblEmployee.Emp_Salary_mob, dbo.TblEmployee.Emp_Salary_mang,"
        My_SQL = My_SQL + "dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Salary,"
        My_SQL = My_SQL + "SUM(QryAllDiscountWithMkafea.TotalDiscount) AS TotalDiscount,"
        My_SQL = My_SQL + "SUM(QryAllDiscountWithMkafea.Mokafea) AS TotalMokafea"
        My_SQL = My_SQL + ""
    
        My_SQL = My_SQL + " From dbo.QryAllDiscountWithMkafea(" & IntMonth & "," & IntYear & ")"
        My_SQL = My_SQL + " QryAllDiscountWithMkafea RIGHT OUTER JOIN"
        My_SQL = My_SQL + " dbo.TblEmployee ON QryAllDiscountWithMkafea.Emp_ID = dbo.TblEmployee.Emp_ID"
    
        If DCEmP.text <> "" Then
            My_SQL = My_SQL + " Where dbo.TblEmployee.workstate=1 and dbo.TblEmployee.emp_code='" & DCEmP.BoundText & "'"
        Else

            If dcdep.text <> "" Then
    
                If dcproject.BoundText = "" Then
                    My_SQL = My_SQL + " Where dbo.TblEmployee.workstate=1 and dbo.TblEmployee.DepartmentID='" & dcdep.BoundText & "'"
                Else
                    My_SQL = My_SQL + " Where dbo.TblEmployee.workstate=1 and dbo.TblEmployee.DepartmentID='" & dcdep.BoundText & "' and dbo.TblEmployee.project_id='" & Me.dcproject.BoundText & "'"
                End If

            Else

                If dcdep.text = "" Then
    
                    If dcproject.BoundText <> "" Then
        
                        My_SQL = My_SQL + " Where dbo.TblEmployee.workstate=1 and  dbo.TblEmployee.project_id='" & Me.dcproject.BoundText & "'"
                    Else
                        My_SQL = My_SQL + " Where dbo.TblEmployee.workstate=1"
                    End If
    
                Else
    
                    My_SQL = My_SQL + " Where dbo.TblEmployee.workstate=1"
                End If
            End If
        End If
    
        My_SQL = My_SQL + " GROUP BY dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Code,dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Salary_sakn, dbo.TblEmployee.Emp_Salary_bus, dbo.TblEmployee.Emp_Salary_food, dbo.TblEmployee.Emp_Salary_others,dbo.TblEmployee.Emp_Salary_mob, dbo.TblEmployee.Emp_Salary_mang,"
        My_SQL = My_SQL + " dbo.TblEmployee.Emp_Salary,dbo.TblEmployee.DepartmentID ,dbo.TblEmployee.project_id"
        My_SQL = My_SQL + " ORDER BY dbo.TblEmployee.Emp_ID"
    
        My_SQL = My_SQL + ")XTable"
    Else
        FrstDay = "1-" & CmbMonth.ListIndex + 1 & "-" & year(Date)
        LstDay = DateAdd("d", -1, "1-" & CmbMonth.ListIndex + 2 & "-" & year(Date))

        My_SQL = "select Emp_ID,Emp_Name,Emp_Salary ,sum(TotalDiscount) as TotalDiscount," & "sum(Mokafea) as Mokafea  From QryEmpAllValues where TransDate >=#" & Format(FrstDay, "mm/dd/yyyy") & "# and TransDate<=#" & Format(LstDay, "mm/dd/yyyy") & "# " & StrWhere & " GROUP BY Emp_ID, Emp_Name, " & "Emp_Salary  "
    End If

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

        GetAdvanceValues IntMonth, IntYear
        GetWorkHours
        CalculateNets
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

Public Sub FillGridWithData2()
    Dim i As Integer
    Dim rs As ADODB.Recordset
    Dim rs2 As ADODB.Recordset
    Dim LstDay As Date
    Dim FrstDay As Date
    Dim StrTxt As String
    Dim My_SQL As String
    Dim StrWhere As String
    Dim StrGrp As String
    Dim IntMonth As Integer
    Dim IntYear As Integer
    Dim Msg As String

    'On Error GoTo ErrTrap
    'If DateDiff("d", Me.DtpFrom.Value, Me.DtpTO.Value, vbSaturday) < 0 Then
    '    Exit Sub
    'End If
    Set rs = New ADODB.Recordset
    Set rs2 = New ADODB.Recordset

    If Me.CmbMonth.ListIndex = -1 Then Exit Sub
    If Me.CboYear.ListIndex = -1 Then Exit Sub

    IntYear = val(Me.CboYear.text)
    IntMonth = Me.CmbMonth.ListIndex + 1

    If SystemOptions.SysDataBaseType = SQLServerDataBase Then
        Dim id As String
    
        My_SQL = "SELECT    id,project_id, DepartmentID,id, Emp_Code, Emp_Name, Emp_Salary, Emp_Salary_sakn, Emp_Salary_bus, Emp_Salary_food, Emp_Salary_mob, Emp_Salary_mang, Emp_Salary_others,"
        My_SQL = My_SQL + "OverTimePrice, Mokafea, SalesCom, total1, TotalAdvance, TotalDiscount, total2, EmpTotalNet, sgn, m_year, m_month, payed"
        My_SQL = My_SQL + " from dbo.emp_salary WHERE     (m_year = '" & Me.CboYear.text & "') AND (m_month = '" & Me.CmbMonth.text & "') AND (payed =0) "

        If DCEmP.text <> "" Then
            My_SQL = My_SQL + "  and  emp_code='" & DCEmP.BoundText & "'"
        Else

            If dcdep.text <> "" Then
    
                If dcproject.BoundText = "" Then
                    My_SQL = My_SQL + "  and  DepartmentID='" & dcdep.BoundText & "'"
                Else
                    My_SQL = My_SQL + "   and  DepartmentID='" & dcdep.BoundText & "' and  project_id='" & Me.dcproject.BoundText & "'"
                End If

            Else

                If dcdep.text = "" Then
    
                    If dcproject.BoundText <> "" Then
        
                        My_SQL = My_SQL + "  and  project_id='" & Me.dcproject.BoundText & "'"
                    End If
    
                End If
            End If
        End If
    
        rs.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        With Me.Grid1
            .Rows = 2
            .Clear flexClearScrollable

            If rs.RecordCount > 0 Then
                .Rows = rs.RecordCount + 1
                rs.MoveFirst

                For i = 1 To .Rows - 1
        
                    .TextMatrix(i, .ColIndex("Ser")) = i
            
                    '  .TextMatrix(i, .ColIndex("Emp_ID")) = IIf(IsNull(Rs.Fields("ID").value), _
                       "", Rs.Fields("ID").value)
            
                    .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs.Fields("id").value), "", rs.Fields("id").value)
            
                    .TextMatrix(i, .ColIndex("Emp_Code")) = IIf(IsNull(rs.Fields("Emp_Code").value), "", rs.Fields("Emp_Code").value)
            
                    .TextMatrix(i, .ColIndex("dep")) = IIf(IsNull(rs.Fields("DepartmentID").value), "", rs.Fields("DepartmentID").value)
            
                    .TextMatrix(i, .ColIndex("project")) = IIf(IsNull(rs.Fields("project_id").value), "", rs.Fields("project_id").value)
            
                    .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(rs.Fields("Emp_Name").value), "", rs.Fields("Emp_Name").value)
               
                    .TextMatrix(i, .ColIndex("Emp_Salary")) = IIf(IsNull(rs.Fields("Emp_Salary").value), "", rs.Fields("Emp_Salary").value)
            
                    .TextMatrix(i, .ColIndex("TotalDiscount")) = IIf(IsNull(rs.Fields("TotalDiscount").value), "", Format(rs.Fields("TotalDiscount").value, SystemOptions.SysDefCurrencyForamt))
                
                    .TextMatrix(i, .ColIndex("Mokafea")) = IIf(IsNull(rs.Fields("Mokafea").value), "", Format(rs.Fields("Mokafea").value, SystemOptions.SysDefCurrencyForamt))
            
                    .TextMatrix(i, .ColIndex("Emp_Salary_sakn")) = IIf(IsNull(rs.Fields("Emp_Salary_sakn").value), "", Format(rs.Fields("Emp_Salary_sakn").value))
            
                    .TextMatrix(i, .ColIndex("Emp_Salary_bus")) = IIf(IsNull(rs.Fields("Emp_Salary_bus").value), "", Format(rs.Fields("Emp_Salary_bus").value))
            
                    .TextMatrix(i, .ColIndex("Emp_Salary_food")) = IIf(IsNull(rs.Fields("Emp_Salary_food").value), "", Format(rs.Fields("Emp_Salary_food").value))
                               
                    .TextMatrix(i, .ColIndex("Emp_Salary_mob")) = IIf(IsNull(rs.Fields("Emp_Salary_mob").value), "", Format(rs.Fields("Emp_Salary_mob").value))
                                    
                    .TextMatrix(i, .ColIndex("Emp_Salary_mang")) = IIf(IsNull(rs.Fields("Emp_Salary_mang").value), "", Format(rs.Fields("Emp_Salary_mang").value))
            
                    .TextMatrix(i, .ColIndex("Emp_Salary_others")) = IIf(IsNull(rs.Fields("Emp_Salary_others").value), "", Format(rs.Fields("Emp_Salary_others").value))
            
                    .TextMatrix(i, .ColIndex("OverTimePrice")) = IIf(IsNull(rs.Fields("OverTimePrice").value), "", Format(rs.Fields("OverTimePrice").value))
            
                    .TextMatrix(i, .ColIndex("SalesCom")) = IIf(IsNull(rs.Fields("SalesCom").value), "", Format(rs.Fields("SalesCom").value))
                    
                    .TextMatrix(i, .ColIndex("total1")) = IIf(IsNull(rs.Fields("total1").value), "", Format(rs.Fields("total1").value))
            
                    .TextMatrix(i, .ColIndex("TotalAdvance")) = IIf(IsNull(rs.Fields("TotalAdvance").value), "", Format(rs.Fields("TotalAdvance").value))
            
                    .TextMatrix(i, .ColIndex("total2")) = IIf(IsNull(rs.Fields("total2").value), "", Format(rs.Fields("total2").value))
            
                    .TextMatrix(i, .ColIndex("EmpTotalNet")) = IIf(IsNull(rs.Fields("EmpTotalNet").value), "", Format(rs.Fields("EmpTotalNet").value))
            
                    rs.MoveNext
            
                Next

                rs.Close
            End If
    
            GetAdvanceValues IntMonth, IntYear
            GetWorkHours
            CalculateNets
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, .ColIndex("Ser")) = "«·√Ã„«·Ï"
            .IsSubtotal(.Rows - 1) = True
            Dim SngTotal As Single
            SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary"), .Rows - 1, .ColIndex("Emp_Salary"))
            .TextMatrix(.Rows - 1, .ColIndex("Emp_Salary")) = SngTotal
    
            SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("EmpTotalNet"), .Rows - 1, .ColIndex("EmpTotalNet"))
            .TextMatrix(.Rows - 1, .ColIndex("EmpTotalNet")) = SngTotal
            net_value1 = SngTotal
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

    End If

ErrTrap:
End Sub

Private Sub GetWorkHours()
    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    Dim LngFindRow As Long
    Dim i As Integer
    Dim X As Long
    Dim Y  As Long
    Dim Z As Long
    Dim IntYear As Integer, IntMonth As Integer
    Dim IntDefWorkHours As Integer

    IntYear = val(Me.CboYear.text)
    IntMonth = Me.CmbMonth.ListIndex + 1

    StrSQL = "SELECT dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name,"
    StrSQL = StrSQL + " dbo.ConvertMintsToHours(sum(dbo.tblPresentTime.WorkHoursCount)) AS WorkHours,"
    StrSQL = StrSQL + " dbo.ConvertMintsToHours(SUM( dbo.tblPresentTime.WorkHoursCount - dbo.tblPresentTime.CurrentWorkMints))as OverTime"
    StrSQL = StrSQL + " FROM  dbo.TblEmployee LEFT OUTER JOIN"
    StrSQL = StrSQL + " dbo.tblPresentTime ON dbo.TblEmployee.Emp_ID = dbo.tblPresentTime.Emp_ID"
    'CONVERT (nvarchar(50),GenPresentTime ,111)
    'StrSQL = StrSQL + " Where CONVERT (nvarchar(50),GenPresentTime ,101) >=" & SQLDate(Me.DtpFrom.Value, True) & " AND " & _
     " CONVERT (nvarchar(50),GenPresentTime ,101) <=" & SQLDate(Me.DtpTO.Value, True)
    StrSQL = StrSQL + " Where Month(GenPresentTime)=" & IntMonth & " AND Year(GenPresentTime)=" & IntYear & ""
    StrSQL = StrSQL + " Group By dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name"
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        Exit Sub
    End If

    IntDefWorkHours = val(Me.TxtMonthHours.text)

    If IntDefWorkHours = 0 Then Exit Sub

    Y = ConvertHoursToMints(IntDefWorkHours & ":00")

    With Me.Grid
        .Cell(flexcpText, .FixedRows, .ColIndex("DefWorkHours"), .Rows - 1, .ColIndex("DefWorkHours")) = IntDefWorkHours

        For i = 1 To rs.RecordCount
            LngFindRow = .FindRow(rs("Emp_ID").value, .FixedRows, .ColIndex("Emp_ID"), False, True)

            If LngFindRow <> -1 Then
                If Not (IsNull(rs("WorkHours").value)) Then
                    .TextMatrix(LngFindRow, .ColIndex("WorkHours")) = rs("WorkHours").value
                    Z = ConvertHoursToMints(rs("WorkHours").value)
                    X = Z - Y

                    If X < 0 Then
                        .TextMatrix(LngFindRow, .ColIndex("OverTime")) = "-" & ConvertMintsToHours(Abs(X))
                    Else
                        .TextMatrix(LngFindRow, .ColIndex("OverTime")) = ConvertMintsToHours(Abs(X))
                    End If
                
                    If InStr(1, .TextMatrix(LngFindRow, .ColIndex("OverTime")), "-", vbTextCompare) <> 0 Then
                        .Cell(flexcpForeColor, LngFindRow, .ColIndex("OverTime")) = vbRed
                    End If

                Else
                    .TextMatrix(LngFindRow, .ColIndex("WorkHours")) = "00:00"
                    .TextMatrix(LngFindRow, .ColIndex("OverTime")) = "00:00"
                End If
            End If

            rs.MoveNext
        Next i

    End With

End Sub

Private Sub CalculateNets()
    Dim i As Integer
    Dim SngHourPrice As Single
    Dim SngOverTimePrice As Single

    Dim NetTotal As Single
    Dim SngTemp As Single
    'On Error GoTo ErrTrap
    On Error Resume Next

    With Me.Grid

        For i = .FixedRows To .Rows - 1
            SngHourPrice = val(.TextMatrix(i, .ColIndex("Emp_Salary"))) / val(.TextMatrix(i, .ColIndex("DefWorkHours")))

            If .TextMatrix(i, .ColIndex("OverTime")) <> "" Then
                SngTemp = ConvertHoursToMints(.TextMatrix(i, .ColIndex("OverTime")))
                SngTemp = SngTemp * (1 / 60)
                SngOverTimePrice = SngTemp * SngHourPrice
                .TextMatrix(i, .ColIndex("OverTimePrice")) = SngOverTimePrice

                If SngOverTimePrice < 0 Then
                    .Cell(flexcpForeColor, i, .ColIndex("OverTimePrice")) = vbRed
                End If
            End If

            .TextMatrix(i, .ColIndex("total1")) = val(.TextMatrix(i, .ColIndex("Emp_Salary"))) + val(.TextMatrix(i, .ColIndex("Emp_Salary_sakn"))) + val(.TextMatrix(i, .ColIndex("Emp_Salary_bus"))) + val(.TextMatrix(i, .ColIndex("Emp_Salary_food"))) + val(.TextMatrix(i, .ColIndex("Emp_Salary_others"))) + val(.TextMatrix(i, .ColIndex("OverTimePrice"))) + val(.TextMatrix(i, .ColIndex("Mokafea"))) + val(.TextMatrix(i, .ColIndex("SalesCom"))) + val(.TextMatrix(i, .ColIndex("Emp_Salary_mob"))) + val(.TextMatrix(i, .ColIndex("Emp_Salary_mang")))
            .TextMatrix(i, .ColIndex("total2")) = val(.TextMatrix(i, .ColIndex("TotalAdvance"))) + val(.TextMatrix(i, .ColIndex("TotalDiscount")))
            .TextMatrix(i, .ColIndex("EmpTotalNet")) = val(.TextMatrix(i, .ColIndex("total1"))) - val(.TextMatrix(i, .ColIndex("total2")))
      
            '.TextMatrix(I, .ColIndex("EmpTotalNet")) = Val(.TextMatrix(I, .ColIndex("EmpTotalNet"))) - Val(.TextMatrix(I, .ColIndex("TotalAdvance")))
            '.TextMatrix(I, .ColIndex("EmpTotalNet")) = Val(.TextMatrix(I, .ColIndex("EmpTotalNet"))) + SngOverTimePrice
            '.TextMatrix(I, .ColIndex("EmpTotalNet")) = Format(Val(.TextMatrix(I, .ColIndex("EmpTotalNet"))), SystemOptions.SysDefCurrencyForamt)
            '.TextMatrix(I, .ColIndex("CorrectEmpTotalNet")) = CorrectCurrency(Val(.TextMatrix(I, .ColIndex("EmpTotalNet"))))
        Next i

    End With

    Exit Sub
ErrTrap:
    'Resume
End Sub

Private Sub Grid_StartPage(ByVal hDC As Long, _
                           ByVal Page As Long, _
                           Cancel As Boolean)
    Dim s As String

    s = "„— »«  «·„ÊŸ›Ì‰ - Page " & Page & " - " & Now
    TextOut hDC, 100, 100, s, Len(s)
End Sub

Private Sub ISButton2_Click()
    FillGridWithData

    DoEvents
    Dim xApp As New CRAXDRT.Application
    Dim rs As New ADODB.Recordset
    Dim My_SQL As String
    Dim xReport As New CRAXDRT.Report

    My_SQL = "SELECT * from emp_salary where m_year='" & CboYear.text & "' and m_month='" & CmbMonth.text & "'"
    rs.Open My_SQL, Cn, adOpenStatic, adLockPessimistic, adCmdText

    Set xReport = xApp.OpenReport(App.path & "\reports\emp\REPORT10.rpt")
    xReport.Database.SetDataSource rs
    Dim FrmReport As New FrmReportViewer
    '   FrmReport = New FrmReportViewer
    FrmReport.CRViewer.ReportSource = xReport
  
    FrmReport.CRViewer.viewReport
    FrmReport.show
    xReport.ParameterFields(4).AddCurrentValue CmbMonth.text
    xReport.ParameterFields(5).AddCurrentValue CboYear.text
     
    Screen.MousePointer = vbDefault
    ' xReport.ReportTitle = X
    SendKeys "{RIGHT}"

End Sub

Private Sub ISButton3_Click()

    Form3.show
    Form3.case_id = 11

End Sub

Private Sub TxtMonthHours_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtMonthHours.text, 1)
End Sub

Private Sub GetAdvanceValues(IntMonth As Integer, _
                             IntYear As Integer)
    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    Dim i As Integer
    Dim LngFindRow As Long
    On Error GoTo hErr
    StrSQL = "Select Emp_ID,Sum(TotalAdvance)as CCC From ( SELECT QryAllEmpAdvance.Emp_ID,QryA" & "llEmpAdvance.TotalAdvance FROM   dbo.QryAllEmpAdvance(" & IntMonth & "," & IntYear & ") QryAllEmpAdvance )" & "Xtable Group By Emp_ID"
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        Exit Sub
    End If

    With Me.Grid
        rs.MoveFirst
        .Cell(flexcpText, .FixedRows, .ColIndex("TotalAdvance"), .Rows - 1, .ColIndex("TotalAdvance")) = 0

        For i = 1 To rs.RecordCount
            LngFindRow = .FindRow(rs("Emp_ID").value, .FixedRows, .ColIndex("Emp_ID"), False, True)

            If LngFindRow <> -1 Then
                If Not (IsNull(rs("CCC").value)) Then
                    .TextMatrix(LngFindRow, .ColIndex("TotalAdvance")) = rs("CCC").value
                End If
            End If

            rs.MoveNext
        Next i

    End With

hErr:
    'Stop
End Sub
