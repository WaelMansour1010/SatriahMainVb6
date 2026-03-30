VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SuiteCtrls.ocx"
Begin VB.Form FrmProjectSearch 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "»ÕÀ"
   ClientHeight    =   7155
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14865
   Icon            =   "FrmProjectSearch.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   7155
   ScaleWidth      =   14865
   Begin C1SizerLibCtl.C1Tab C1Tab1 
      Height          =   7095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14895
      _cx             =   26273
      _cy             =   12515
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
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      FrontTabColor   =   -2147483633
      BackTabColor    =   -2147483633
      TabOutlineColor =   -2147483632
      FrontTabForeColor=   -2147483630
      Caption         =   $"FrmProjectSearch.frx":030A
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
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   6720
         Left            =   45
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   45
         Width           =   14805
         _cx             =   26114
         _cy             =   11853
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
         Begin VB.Frame Fra 
            Height          =   2565
            Index           =   2
            Left            =   96
            RightToLeft     =   -1  'True
            TabIndex        =   2
            Top             =   3480
            Width           =   14595
            Begin VB.CheckBox XPChkSearchType 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·„Ê—œ »«·ﬂ«„· ›ﬁÿ"
               Height          =   375
               Left            =   5760
               RightToLeft     =   -1  'True
               TabIndex        =   21
               Top             =   2760
               Value           =   1  'Checked
               Visible         =   0   'False
               Width           =   2385
            End
            Begin VB.TextBox TxtCompanyName 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   6840
               RightToLeft     =   -1  'True
               TabIndex        =   20
               Top             =   240
               Width           =   6375
            End
            Begin VB.TextBox XPTxtComID 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   10800
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   19
               Top             =   600
               Width           =   2415
            End
            Begin VB.TextBox TxtCustCode 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   3960
               TabIndex        =   18
               Top             =   240
               Width           =   1215
            End
            Begin VB.TextBox TxtCustCode1 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   3960
               TabIndex        =   17
               Top             =   600
               Width           =   1215
            End
            Begin VB.TextBox TxtProjectCosts 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   6840
               TabIndex        =   16
               Top             =   600
               Width           =   2895
            End
            Begin VB.TextBox TxtCustCode2 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   3960
               TabIndex        =   15
               Top             =   960
               Width           =   1215
            End
            Begin VB.TextBox Text10 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   3960
               TabIndex        =   14
               Top             =   1320
               Width           =   1215
            End
            Begin VB.TextBox TxtBand 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   10800
               TabIndex        =   13
               Top             =   960
               Width           =   2415
            End
            Begin VB.Frame Fra 
               Caption         =   " «—ÌŒ »œ«Ì… «·„‘—Ê⁄"
               Height          =   1095
               Index           =   0
               Left            =   11520
               RightToLeft     =   -1  'True
               TabIndex        =   8
               Top             =   1320
               Width           =   2895
               Begin MSComCtl2.DTPicker FrmDTStartDate 
                  Height          =   330
                  Left            =   90
                  TabIndex        =   9
                  Top             =   270
                  Width           =   1590
                  _ExtentX        =   2805
                  _ExtentY        =   582
                  _Version        =   393216
                  CheckBox        =   -1  'True
                  Format          =   177012739
                  CurrentDate     =   38887
               End
               Begin MSComCtl2.DTPicker TODTStartDate 
                  Height          =   330
                  Left            =   90
                  TabIndex        =   10
                  Top             =   630
                  Width           =   1590
                  _ExtentX        =   2805
                  _ExtentY        =   582
                  _Version        =   393216
                  CheckBox        =   -1  'True
                  Format          =   177012739
                  CurrentDate     =   38887
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  Caption         =   "≈·Ï"
                  Height          =   195
                  Index           =   0
                  Left            =   1815
                  RightToLeft     =   -1  'True
                  TabIndex        =   12
                  Top             =   660
                  Width           =   375
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  Caption         =   "„‰"
                  Height          =   195
                  Index           =   5
                  Left            =   1740
                  RightToLeft     =   -1  'True
                  TabIndex        =   11
                  Top             =   330
                  Width           =   420
               End
            End
            Begin VB.Frame Fra 
               Caption         =   " √—ÌŒ ‰Â«Ì… «·„‘—Ê⁄"
               Height          =   1095
               Index           =   3
               Left            =   8520
               RightToLeft     =   -1  'True
               TabIndex        =   3
               Top             =   1320
               Width           =   2895
               Begin MSComCtl2.DTPicker FrmDTEnddate 
                  Height          =   330
                  Left            =   90
                  TabIndex        =   4
                  Top             =   270
                  Width           =   1590
                  _ExtentX        =   2805
                  _ExtentY        =   582
                  _Version        =   393216
                  CheckBox        =   -1  'True
                  Format          =   177012739
                  CurrentDate     =   38887
               End
               Begin MSComCtl2.DTPicker ToDTEnddate 
                  Height          =   330
                  Left            =   90
                  TabIndex        =   5
                  Top             =   630
                  Width           =   1590
                  _ExtentX        =   2805
                  _ExtentY        =   582
                  _Version        =   393216
                  CheckBox        =   -1  'True
                  Format          =   177012739
                  CurrentDate     =   38887
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  Caption         =   "„‰"
                  Height          =   195
                  Index           =   6
                  Left            =   1740
                  RightToLeft     =   -1  'True
                  TabIndex        =   7
                  Top             =   330
                  Width           =   420
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  Caption         =   "≈·Ï"
                  Height          =   195
                  Index           =   7
                  Left            =   1815
                  RightToLeft     =   -1  'True
                  TabIndex        =   6
                  Top             =   660
                  Width           =   375
               End
            End
            Begin MSDataListLib.DataCombo DcAccount2 
               Height          =   315
               Left            =   120
               TabIndex        =   22
               Top             =   240
               Width           =   3855
               _ExtentX        =   6800
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcAccount4 
               Height          =   315
               Left            =   120
               TabIndex        =   23
               Top             =   600
               Width           =   3855
               _ExtentX        =   6800
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcEmp1 
               Height          =   315
               Left            =   120
               TabIndex        =   24
               Top             =   960
               Width           =   3855
               _ExtentX        =   6800
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcEmp 
               Height          =   315
               Left            =   120
               TabIndex        =   25
               Top             =   1320
               Width           =   3855
               _ExtentX        =   6800
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcbDept 
               Height          =   315
               Left            =   120
               TabIndex        =   26
               Top             =   1680
               Width           =   5055
               _ExtentX        =   8916
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcbProcess 
               Height          =   315
               Left            =   6840
               TabIndex        =   27
               Top             =   960
               Width           =   2895
               _ExtentX        =   5106
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label Label16 
               Alignment       =   1  'Right Justify
               Caption         =   "«·⁄„Ì· «·‰Â«∆Ì"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   5280
               TabIndex        =   37
               Top             =   240
               Width           =   1095
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "ﬂÊœ «·„‘—Ê⁄"
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   0
               Left            =   13200
               TabIndex        =   36
               Top             =   600
               Width           =   1095
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               Caption         =   "«”„ «·„‘—Ê⁄"
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   0
               Left            =   13200
               TabIndex        =   35
               Top             =   240
               Width           =   1095
            End
            Begin VB.Label Label23 
               Alignment       =   1  'Right Justify
               Caption         =   "«·⁄„Ì· «·»«ÿ‰"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   5280
               TabIndex        =   34
               Top             =   600
               Width           =   1095
            End
            Begin VB.Label Label20 
               Alignment       =   1  'Right Justify
               Caption         =   "ﬁÌ„… «·„‘—Ê⁄"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   8880
               TabIndex        =   33
               Top             =   600
               Width           =   1815
            End
            Begin VB.Label Label41 
               Alignment       =   1  'Right Justify
               Caption         =   "«·„‰œÊ»"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   5280
               TabIndex        =   32
               Top             =   960
               Width           =   1095
            End
            Begin VB.Label Label35 
               Alignment       =   1  'Right Justify
               Caption         =   "„œÌ— «·„Êﬁ⁄"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   5280
               TabIndex        =   31
               Top             =   1320
               Width           =   1095
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               Caption         =   "Ê’› «·»‰œ"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   13200
               TabIndex        =   30
               Top             =   960
               Width           =   1095
            End
            Begin VB.Label Label43 
               Alignment       =   1  'Right Justify
               Caption         =   "«·«œ«—Â"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   5160
               TabIndex        =   29
               Top             =   1680
               Width           =   1215
            End
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               Caption         =   "«·⁄„·ÌÂ"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   9600
               TabIndex        =   28
               Top             =   960
               Width           =   1095
            End
         End
         Begin VSFlex8UCtl.VSFlexGrid Fg 
            Height          =   3225
            Left            =   120
            TabIndex        =   38
            Top             =   120
            Width           =   14625
            _cx             =   25797
            _cy             =   5689
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
            FormatString    =   $"FrmProjectSearch.frx":0397
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
         Begin ImpulseButton.ISButton Cmd 
            Height          =   375
            Index           =   0
            Left            =   6420
            TabIndex        =   39
            Top             =   6150
            Width           =   2355
            _ExtentX        =   4154
            _ExtentY        =   661
            ButtonPositionImage=   1
            Caption         =   "»ÕÀ ⁄·Ï „” ÊÏ «·„‘—Ê⁄"
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
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            ColorToggledHoverText=   16711680
            LowerToggledContent=   0   'False
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   375
            Index           =   1
            Left            =   1200
            TabIndex        =   40
            Top             =   6150
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   661
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
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            LowerToggledContent=   0   'False
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   375
            Index           =   2
            Left            =   120
            TabIndex        =   41
            Top             =   6150
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   661
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
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            LowerToggledContent=   0   'False
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   375
            Index           =   3
            Left            =   4410
            TabIndex        =   42
            Top             =   6120
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   661
            ButtonPositionImage=   1
            Caption         =   "»ÕÀ ⁄·Ï „” ÊÏ «·»‰Êœ"
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
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            ColorToggledHoverText=   16711680
            LowerToggledContent=   0   'False
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   375
            Index           =   4
            Left            =   2370
            TabIndex        =   43
            Top             =   6150
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   661
            ButtonPositionImage=   1
            Caption         =   "»ÕÀ ⁄·Ï „” ÊÏ «·⁄„·Ì« "
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
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            ColorToggledHoverText=   16711680
            LowerToggledContent=   0   'False
            ColorTextShadow =   4210752
         End
         Begin VB.Label lblSearchtype 
            Alignment       =   1  'Right Justify
            Caption         =   "0"
            Height          =   252
            Left            =   216
            RightToLeft     =   -1  'True
            TabIndex        =   44
            Top             =   2760
            Visible         =   0   'False
            Width           =   252
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic7 
         Height          =   6720
         Left            =   15540
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   45
         Width           =   14805
         _cx             =   26114
         _cy             =   11853
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
         Begin VB.Frame Fra 
            Height          =   1605
            Index           =   1
            Left            =   90
            RightToLeft     =   -1  'True
            TabIndex        =   46
            Top             =   4440
            Width           =   14595
            Begin VB.TextBox Txt_CustomerCode 
               Alignment       =   2  'Center
               Height          =   285
               Left            =   7290
               TabIndex        =   57
               Top             =   360
               Width           =   3885
            End
            Begin VB.Frame Fra 
               Caption         =   " «—ÌŒ"
               Height          =   1095
               Index           =   4
               Left            =   240
               RightToLeft     =   -1  'True
               TabIndex        =   52
               Top             =   240
               Width           =   5895
               Begin MSComCtl2.DTPicker DTP_Me_From 
                  Height          =   336
                  Left            =   3000
                  TabIndex        =   53
                  Top             =   396
                  Width           =   1596
                  _ExtentX        =   2805
                  _ExtentY        =   582
                  _Version        =   393216
                  CheckBox        =   -1  'True
                  Format          =   177078275
                  CurrentDate     =   38887
               End
               Begin MSComCtl2.DTPicker DTP_Me_To 
                  Height          =   336
                  Left            =   216
                  TabIndex        =   54
                  Top             =   396
                  Width           =   1596
                  _ExtentX        =   2805
                  _ExtentY        =   582
                  _Version        =   393216
                  CheckBox        =   -1  'True
                  Format          =   177078275
                  CurrentDate     =   38887
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  Caption         =   "„‰"
                  Height          =   192
                  Index           =   1
                  Left            =   5100
                  RightToLeft     =   -1  'True
                  TabIndex        =   56
                  Top             =   456
                  Width           =   420
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  Caption         =   "≈·Ï"
                  Height          =   192
                  Index           =   2
                  Left            =   2172
                  RightToLeft     =   -1  'True
                  TabIndex        =   55
                  Top             =   420
                  Width           =   372
               End
            End
            Begin VB.TextBox Txt_City 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   6480
               TabIndex        =   51
               Top             =   720
               Width           =   2415
            End
            Begin VB.TextBox Txt_Mobile 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   10200
               TabIndex        =   50
               Top             =   720
               Width           =   3255
            End
            Begin VB.TextBox TXT_District 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   10200
               TabIndex        =   49
               Top             =   1080
               Width           =   3255
            End
            Begin VB.TextBox Txt_MeasureID 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   12120
               RightToLeft     =   -1  'True
               TabIndex        =   48
               Top             =   360
               Width           =   1335
            End
            Begin VB.CheckBox Check1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·„Ê—œ »«·ﬂ«„· ›ﬁÿ"
               Height          =   375
               Left            =   5760
               RightToLeft     =   -1  'True
               TabIndex        =   47
               Top             =   2760
               Value           =   1  'Checked
               Visible         =   0   'False
               Width           =   2385
            End
            Begin MSDataListLib.DataCombo DcbCustmer 
               Bindings        =   "FrmProjectSearch.frx":0698
               Height          =   315
               Left            =   6480
               TabIndex        =   58
               Top             =   360
               Visible         =   0   'False
               Width           =   705
               _ExtentX        =   1244
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
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
               Caption         =   "«·„œÌ‰…"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   9000
               TabIndex        =   63
               Top             =   720
               Width           =   975
            End
            Begin VB.Label Label6 
               Alignment       =   1  'Right Justify
               Caption         =   "«·ÃÊ«·"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   13560
               TabIndex        =   62
               Top             =   720
               Width           =   615
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               Caption         =   "—ﬁ„ «·ÿ·»"
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   8
               Left            =   13440
               TabIndex        =   61
               Top             =   360
               Width           =   735
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "«”„ «·⁄„Ì·"
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   26
               Left            =   11280
               TabIndex        =   60
               Top             =   360
               Width           =   735
            End
            Begin VB.Label Label7 
               Alignment       =   1  'Right Justify
               Caption         =   "«·ÕÏ"
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   13440
               TabIndex        =   59
               Top             =   1080
               Width           =   735
            End
         End
         Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
            Height          =   4305
            Left            =   120
            TabIndex        =   64
            Top             =   120
            Width           =   14625
            _cx             =   25797
            _cy             =   7594
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
            Cols            =   8
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmProjectSearch.frx":06AD
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
         Begin ImpulseButton.ISButton Cmd_Search 
            Height          =   375
            Left            =   8820
            TabIndex        =   65
            Top             =   6195
            Width           =   2355
            _ExtentX        =   4154
            _ExtentY        =   661
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
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            ColorToggledHoverText=   16711680
            LowerToggledContent=   0   'False
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton Cmd_Clear 
            Height          =   375
            Left            =   6270
            TabIndex        =   66
            Top             =   6195
            Width           =   2355
            _ExtentX        =   4154
            _ExtentY        =   661
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
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            LowerToggledContent=   0   'False
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton Cmd_Exit 
            Height          =   375
            Left            =   3720
            TabIndex        =   67
            Top             =   6195
            Width           =   2355
            _ExtentX        =   4154
            _ExtentY        =   661
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
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            LowerToggledContent=   0   'False
            ColorTextShadow =   4210752
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            Caption         =   "0"
            Height          =   252
            Left            =   -624
            RightToLeft     =   -1  'True
            TabIndex        =   68
            Top             =   3960
            Visible         =   0   'False
            Width           =   612
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic2 
         Height          =   6720
         Left            =   15840
         TabIndex        =   69
         TabStop         =   0   'False
         Top             =   45
         Width           =   14805
         _cx             =   26114
         _cy             =   11853
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
         Begin VB.Frame Fra 
            Height          =   2088
            Index           =   5
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   70
            Top             =   3840
            Width           =   14715
            Begin VB.CheckBox Check2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·„Ê—œ »«·ﬂ«„· ›ﬁÿ"
               Height          =   375
               Left            =   5760
               RightToLeft     =   -1  'True
               TabIndex        =   81
               Top             =   2760
               Value           =   1  'Checked
               Visible         =   0   'False
               Width           =   2385
            End
            Begin VB.TextBox TXT_TOrder 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   12720
               RightToLeft     =   -1  'True
               TabIndex        =   80
               Top             =   480
               Width           =   852
            End
            Begin VB.Frame Fra 
               Caption         =   " «—ÌŒ"
               Height          =   1095
               Index           =   6
               Left            =   720
               RightToLeft     =   -1  'True
               TabIndex        =   75
               Top             =   600
               Width           =   6135
               Begin MSComCtl2.DTPicker DTP_T_From 
                  Height          =   336
                  Left            =   3120
                  TabIndex        =   76
                  Top             =   396
                  Width           =   1596
                  _ExtentX        =   2805
                  _ExtentY        =   582
                  _Version        =   393216
                  CheckBox        =   -1  'True
                  Format          =   168427523
                  CurrentDate     =   38887
               End
               Begin MSComCtl2.DTPicker DTP_T_To 
                  Height          =   336
                  Left            =   216
                  TabIndex        =   77
                  Top             =   396
                  Width           =   1596
                  _ExtentX        =   2805
                  _ExtentY        =   582
                  _Version        =   393216
                  CheckBox        =   -1  'True
                  Format          =   168427523
                  CurrentDate     =   38887
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  Caption         =   "≈·Ï"
                  Height          =   192
                  Index           =   3
                  Left            =   2172
                  RightToLeft     =   -1  'True
                  TabIndex        =   79
                  Top             =   420
                  Width           =   372
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  Caption         =   "„‰"
                  Height          =   192
                  Index           =   4
                  Left            =   5100
                  RightToLeft     =   -1  'True
                  TabIndex        =   78
                  Top             =   456
                  Width           =   420
               End
            End
            Begin VB.ComboBox CboPayMentType 
               Height          =   315
               Left            =   7680
               RightToLeft     =   -1  'True
               TabIndex        =   74
               Top             =   960
               Width           =   1995
            End
            Begin VB.TextBox TxtSearchCode 
               Alignment       =   1  'Right Justify
               Height          =   288
               Left            =   12450
               RightToLeft     =   -1  'True
               TabIndex        =   73
               Top             =   1440
               Width           =   1200
            End
            Begin VB.TextBox Txt_MID 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   10560
               RightToLeft     =   -1  'True
               TabIndex        =   72
               Top             =   480
               Width           =   1215
            End
            Begin VB.TextBox TxtTime 
               Alignment       =   2  'Center
               Height          =   315
               Left            =   7680
               TabIndex        =   71
               Top             =   480
               Width           =   1980
            End
            Begin MSDataListLib.DataCombo Dcbranch 
               Bindings        =   "FrmProjectSearch.frx":07D7
               Height          =   315
               Left            =   10560
               TabIndex        =   82
               Top             =   960
               Width           =   3090
               _ExtentX        =   5450
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
            Begin MSDataListLib.DataCombo DcbEmployee2 
               Bindings        =   "FrmProjectSearch.frx":07EC
               Height          =   315
               Left            =   7680
               TabIndex        =   83
               Top             =   1440
               Width           =   4545
               _ExtentX        =   8017
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
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               Caption         =   "—ﬁ„ «·Õ—ﬂ…"
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   1
               Left            =   13800
               TabIndex        =   89
               Top             =   480
               Width           =   735
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               Caption         =   "«·›—⁄"
               Height          =   285
               Index           =   8
               Left            =   14010
               TabIndex        =   88
               Top             =   960
               Width           =   645
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "«·Õ«·…"
               Height          =   165
               Index           =   4
               Left            =   9645
               RightToLeft     =   -1  'True
               TabIndex        =   87
               Top             =   1020
               Width           =   570
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               Caption         =   "«·„ÊŸ›"
               Height          =   285
               Index           =   9
               Left            =   14040
               RightToLeft     =   -1  'True
               TabIndex        =   86
               Top             =   1440
               Width           =   435
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               Caption         =   "—ﬁ„ «·ÿ·»"
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   2
               Left            =   11880
               TabIndex        =   85
               Top             =   480
               Width           =   735
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               Caption         =   "«·Êﬁ "
               Height          =   285
               Index           =   10
               Left            =   9750
               TabIndex        =   84
               Top             =   495
               Width           =   555
            End
         End
         Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid2 
            Height          =   3585
            Left            =   0
            TabIndex        =   90
            Top             =   120
            Width           =   14745
            _cx             =   26009
            _cy             =   6324
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
            Cols            =   8
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmProjectSearch.frx":0801
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
         Begin ImpulseButton.ISButton ISButton1 
            Height          =   375
            Left            =   8820
            TabIndex        =   91
            Top             =   6195
            Width           =   2355
            _ExtentX        =   4154
            _ExtentY        =   661
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
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            ColorToggledHoverText=   16711680
            LowerToggledContent=   0   'False
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton ISButton2 
            Height          =   375
            Left            =   6390
            TabIndex        =   92
            Top             =   6195
            Width           =   2355
            _ExtentX        =   4154
            _ExtentY        =   661
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
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            LowerToggledContent=   0   'False
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton ISButton3 
            Height          =   375
            Left            =   3960
            TabIndex        =   93
            Top             =   6195
            Width           =   2355
            _ExtentX        =   4154
            _ExtentY        =   661
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
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            LowerToggledContent=   0   'False
            ColorTextShadow =   4210752
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            Caption         =   "0"
            Height          =   252
            Left            =   -264
            RightToLeft     =   -1  'True
            TabIndex        =   94
            Top             =   3240
            Visible         =   0   'False
            Width           =   612
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic3 
         Height          =   6720
         Index           =   0
         Left            =   16140
         TabIndex        =   95
         TabStop         =   0   'False
         Top             =   45
         Width           =   14805
         _cx             =   26114
         _cy             =   11853
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
         Begin VB.Frame Fra 
            Height          =   2088
            Index           =   7
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   96
            Top             =   3840
            Width           =   14475
            Begin VB.TextBox Txt_NumberContract 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   7560
               RightToLeft     =   -1  'True
               TabIndex        =   105
               Top             =   480
               Width           =   2175
            End
            Begin VB.Frame Fra 
               Caption         =   " «—ÌŒ"
               Height          =   1095
               Index           =   8
               Left            =   600
               RightToLeft     =   -1  'True
               TabIndex        =   100
               Top             =   480
               Width           =   6012
               Begin MSComCtl2.DTPicker DTP_BDialy_From 
                  Height          =   336
                  Left            =   3120
                  TabIndex        =   101
                  Top             =   396
                  Width           =   1596
                  _ExtentX        =   2805
                  _ExtentY        =   582
                  _Version        =   393216
                  CheckBox        =   -1  'True
                  Format          =   177143811
                  CurrentDate     =   38887
               End
               Begin MSComCtl2.DTPicker DTP_BDialy_To 
                  Height          =   336
                  Left            =   216
                  TabIndex        =   102
                  Top             =   396
                  Width           =   1596
                  _ExtentX        =   2805
                  _ExtentY        =   582
                  _Version        =   393216
                  CheckBox        =   -1  'True
                  Format          =   177143811
                  CurrentDate     =   38887
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  Caption         =   "„‰"
                  Height          =   192
                  Index           =   11
                  Left            =   5100
                  RightToLeft     =   -1  'True
                  TabIndex        =   104
                  Top             =   456
                  Width           =   420
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  Caption         =   "≈·Ï"
                  Height          =   192
                  Index           =   12
                  Left            =   2172
                  RightToLeft     =   -1  'True
                  TabIndex        =   103
                  Top             =   420
                  Width           =   372
               End
            End
            Begin VB.TextBox Txt_BusinessID 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   10995
               RightToLeft     =   -1  'True
               TabIndex        =   99
               Top             =   480
               Width           =   2175
            End
            Begin VB.CheckBox Check3 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·„Ê—œ »«·ﬂ«„· ›ﬁÿ"
               Height          =   375
               Left            =   5760
               RightToLeft     =   -1  'True
               TabIndex        =   98
               Top             =   2760
               Value           =   1  'Checked
               Visible         =   0   'False
               Width           =   2385
            End
            Begin VB.TextBox TxCustmer 
               Alignment       =   2  'Center
               Height          =   288
               Left            =   11880
               TabIndex        =   97
               Top             =   1320
               Width           =   1275
            End
            Begin MSDataListLib.DataCombo DataComboBranch 
               Bindings        =   "FrmProjectSearch.frx":092B
               Height          =   315
               Left            =   7560
               TabIndex        =   106
               Top             =   960
               Width           =   5610
               _ExtentX        =   9895
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
            Begin MSDataListLib.DataCombo DcbCustom 
               Height          =   315
               Left            =   7560
               TabIndex        =   107
               Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·„«·ﬂ"
               Top             =   1320
               Width           =   4155
               _ExtentX        =   7329
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               Caption         =   "—ﬁ„ «·« ›«ﬁÌ…"
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   3
               Left            =   9840
               TabIndex        =   111
               Top             =   480
               Width           =   975
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               Caption         =   "«·›—⁄"
               Height          =   285
               Index           =   15
               Left            =   13290
               TabIndex        =   110
               Top             =   960
               Width           =   645
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               Caption         =   "—ﬁ„"
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   4
               Left            =   13320
               TabIndex        =   109
               Top             =   480
               Width           =   495
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«·⁄„Ì·"
               Height          =   285
               Index           =   2
               Left            =   13320
               RightToLeft     =   -1  'True
               TabIndex        =   108
               Top             =   1320
               Width           =   675
            End
         End
         Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid3 
            Height          =   3585
            Left            =   120
            TabIndex        =   112
            Top             =   120
            Width           =   14505
            _cx             =   25585
            _cy             =   6324
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
            Cols            =   6
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmProjectSearch.frx":0940
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
         Begin ImpulseButton.ISButton ISButton4 
            Height          =   375
            Left            =   8340
            TabIndex        =   113
            Top             =   6195
            Width           =   2355
            _ExtentX        =   4154
            _ExtentY        =   661
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
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            ColorToggledHoverText=   16711680
            LowerToggledContent=   0   'False
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton ISButton5 
            Height          =   375
            Left            =   5910
            TabIndex        =   114
            Top             =   6195
            Width           =   2355
            _ExtentX        =   4154
            _ExtentY        =   661
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
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            LowerToggledContent=   0   'False
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton ISButton6 
            Height          =   375
            Left            =   3480
            TabIndex        =   115
            Top             =   6195
            Width           =   2355
            _ExtentX        =   4154
            _ExtentY        =   661
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
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            LowerToggledContent=   0   'False
            ColorTextShadow =   4210752
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            Caption         =   "0"
            Height          =   252
            Left            =   -384
            RightToLeft     =   -1  'True
            TabIndex        =   116
            Top             =   3240
            Visible         =   0   'False
            Width           =   612
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic4 
         Height          =   6720
         Left            =   16440
         TabIndex        =   117
         TabStop         =   0   'False
         Top             =   45
         Width           =   14805
         _cx             =   26114
         _cy             =   11853
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
         Begin VB.Frame Fra 
            BackColor       =   &H00E2E9E9&
            Height          =   2088
            Index           =   9
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   118
            Top             =   3960
            Width           =   14595
            Begin VB.TextBox txtNoteSerial1 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   810
               RightToLeft     =   -1  'True
               TabIndex        =   223
               Top             =   270
               Width           =   4815
            End
            Begin VB.CheckBox Check4 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·„Ê—œ »«·ﬂ«„· ›ﬁÿ"
               Height          =   375
               Left            =   5760
               RightToLeft     =   -1  'True
               TabIndex        =   128
               Top             =   2760
               Value           =   1  'Checked
               Visible         =   0   'False
               Width           =   2385
            End
            Begin VB.TextBox TxT_TradContract 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   8040
               RightToLeft     =   -1  'True
               TabIndex        =   127
               Top             =   240
               Width           =   4815
            End
            Begin VB.Frame Fra 
               BackColor       =   &H00E2E9E9&
               Caption         =   " «—ÌŒ"
               Height          =   1095
               Index           =   10
               Left            =   810
               RightToLeft     =   -1  'True
               TabIndex        =   122
               Top             =   780
               Width           =   6012
               Begin MSComCtl2.DTPicker DTP_TD_From 
                  Height          =   336
                  Left            =   3120
                  TabIndex        =   123
                  Top             =   396
                  Width           =   1596
                  _ExtentX        =   2805
                  _ExtentY        =   582
                  _Version        =   393216
                  CheckBox        =   -1  'True
                  Format          =   177143811
                  CurrentDate     =   38887
               End
               Begin MSComCtl2.DTPicker DTP_TD_To 
                  Height          =   336
                  Left            =   216
                  TabIndex        =   124
                  Top             =   396
                  Width           =   1596
                  _ExtentX        =   2805
                  _ExtentY        =   582
                  _Version        =   393216
                  CheckBox        =   -1  'True
                  Format          =   177143811
                  CurrentDate     =   38887
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "≈·Ï"
                  Height          =   192
                  Index           =   16
                  Left            =   2172
                  RightToLeft     =   -1  'True
                  TabIndex        =   126
                  Top             =   420
                  Width           =   372
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„‰"
                  Height          =   192
                  Index           =   17
                  Left            =   5100
                  RightToLeft     =   -1  'True
                  TabIndex        =   125
                  Top             =   456
                  Width           =   420
               End
            End
            Begin VB.TextBox TXtCustID 
               Alignment       =   2  'Center
               Height          =   288
               Left            =   11820
               TabIndex        =   121
               Top             =   720
               Width           =   1035
            End
            Begin VB.TextBox TxtAddress 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   8040
               RightToLeft     =   -1  'True
               TabIndex        =   120
               Top             =   1080
               Width           =   4785
            End
            Begin VB.TextBox TxtPhone 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               Height          =   315
               Left            =   8040
               RightToLeft     =   -1  'True
               TabIndex        =   119
               Top             =   1560
               Width           =   4785
            End
            Begin MSDataListLib.DataCombo DcbCusTC 
               Height          =   315
               Left            =   8040
               TabIndex        =   129
               Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·„«·ﬂ"
               Top             =   720
               Width           =   3675
               _ExtentX        =   6482
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ «·›« Ê—…"
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   7
               Left            =   5490
               TabIndex        =   224
               Top             =   300
               Width           =   1215
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ «·« ›«ﬁÌ…"
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   5
               Left            =   12720
               TabIndex        =   133
               Top             =   270
               Width           =   1215
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«·⁄„Ì·"
               Height          =   285
               Index           =   1
               Left            =   13110
               RightToLeft     =   -1  'True
               TabIndex        =   132
               Top             =   720
               Width           =   915
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·⁄‰Ê«‰"
               Height          =   285
               Index           =   18
               Left            =   13110
               RightToLeft     =   -1  'True
               TabIndex        =   131
               Top             =   1095
               Width           =   915
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ÃÊ«·"
               Height          =   285
               Index           =   19
               Left            =   13110
               RightToLeft     =   -1  'True
               TabIndex        =   130
               Top             =   1575
               Width           =   915
            End
         End
         Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid4 
            Height          =   3825
            Left            =   120
            TabIndex        =   134
            Top             =   120
            Width           =   14625
            _cx             =   25797
            _cy             =   6747
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
            FormatString    =   $"FrmProjectSearch.frx":0A28
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
         Begin ImpulseButton.ISButton ISButton7 
            Height          =   375
            Left            =   8700
            TabIndex        =   135
            Top             =   6195
            Width           =   2355
            _ExtentX        =   4154
            _ExtentY        =   661
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
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            ColorToggledHoverText=   16711680
            LowerToggledContent=   0   'False
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton ISButton8 
            Height          =   375
            Left            =   6150
            TabIndex        =   136
            Top             =   6195
            Width           =   2355
            _ExtentX        =   4154
            _ExtentY        =   661
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
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            LowerToggledContent=   0   'False
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton ISButton9 
            Height          =   375
            Left            =   3600
            TabIndex        =   137
            Top             =   6195
            Width           =   2355
            _ExtentX        =   4154
            _ExtentY        =   661
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
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            LowerToggledContent=   0   'False
            ColorTextShadow =   4210752
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            Caption         =   "0"
            Height          =   252
            Left            =   -384
            RightToLeft     =   -1  'True
            TabIndex        =   139
            Top             =   3240
            Visible         =   0   'False
            Width           =   612
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            Caption         =   "0"
            Height          =   252
            Left            =   13440
            RightToLeft     =   -1  'True
            TabIndex        =   138
            Top             =   840
            Visible         =   0   'False
            Width           =   612
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic5 
         Height          =   6720
         Left            =   16740
         TabIndex        =   140
         TabStop         =   0   'False
         Top             =   45
         Width           =   14805
         _cx             =   26114
         _cy             =   11853
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
            Height          =   1725
            Index           =   11
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   141
            Top             =   4320
            Width           =   14595
            Begin VB.TextBox txtContainerNo 
               BackColor       =   &H0000FFFF&
               Height          =   345
               Left            =   8670
               TabIndex        =   227
               Top             =   1230
               Width           =   1725
            End
            Begin VB.Frame Frame1 
               BackColor       =   &H00E2E9E9&
               Caption         =   "‰Ê⁄ «·Õ—ﬂÂ"
               Height          =   645
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   199
               Top             =   120
               Width           =   2835
               Begin XtremeSuiteControls.RadioButton RdAuto_Manual 
                  Height          =   255
                  Index           =   0
                  Left            =   1200
                  TabIndex        =   200
                  Top             =   240
                  Width           =   615
                  _Version        =   786432
                  _ExtentX        =   1085
                  _ExtentY        =   450
                  _StockProps     =   79
                  Caption         =   "ÌœÊÌ"
                  UseVisualStyle  =   -1  'True
                  TextAlignment   =   1
                  RightToLeft     =   -1  'True
               End
               Begin XtremeSuiteControls.RadioButton RdAuto_Manual 
                  Height          =   255
                  Index           =   1
                  Left            =   120
                  TabIndex        =   201
                  Top             =   240
                  Width           =   615
                  _Version        =   786432
                  _ExtentX        =   1085
                  _ExtentY        =   450
                  _StockProps     =   79
                  Caption         =   "«·Ì"
                  UseVisualStyle  =   -1  'True
                  TextAlignment   =   1
                  RightToLeft     =   -1  'True
               End
               Begin XtremeSuiteControls.RadioButton RdAuto_Manual 
                  Height          =   255
                  Index           =   2
                  Left            =   2040
                  TabIndex        =   202
                  Top             =   240
                  Width           =   615
                  _Version        =   786432
                  _ExtentX        =   1085
                  _ExtentY        =   450
                  _StockProps     =   79
                  Caption         =   "«·ﬂ·"
                  UseVisualStyle  =   -1  'True
                  TextAlignment   =   1
                  RightToLeft     =   -1  'True
               End
            End
            Begin VB.TextBox TxtItemCode 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   6180
               TabIndex        =   165
               Top             =   1215
               Width           =   1035
            End
            Begin VB.TextBox Text3 
               Alignment       =   2  'Center
               Height          =   315
               Left            =   6180
               TabIndex        =   153
               Top             =   840
               Width           =   1035
            End
            Begin VB.CheckBox Check5 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·„Ê—œ »«·ﬂ«„· ›ﬁÿ"
               Height          =   375
               Left            =   5760
               RightToLeft     =   -1  'True
               TabIndex        =   152
               Top             =   2760
               Value           =   1  'Checked
               Visible         =   0   'False
               Width           =   2385
            End
            Begin VB.Frame lbprocess 
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ «·Õ—ﬂÂ"
               Height          =   645
               Left            =   8640
               RightToLeft     =   -1  'True
               TabIndex        =   147
               Top             =   120
               Width           =   5835
               Begin VB.TextBox TxtIDFrom 
                  Alignment       =   1  'Right Justify
                  Height          =   345
                  Left            =   2760
                  RightToLeft     =   -1  'True
                  TabIndex        =   149
                  Top             =   240
                  Width           =   1275
               End
               Begin VB.TextBox TxtIDTO 
                  Alignment       =   1  'Right Justify
                  Height          =   345
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   148
                  Top             =   240
                  Width           =   1275
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„‰"
                  Height          =   195
                  Index           =   22
                  Left            =   3975
                  RightToLeft     =   -1  'True
                  TabIndex        =   151
                  Top             =   240
                  Width           =   540
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "≈·Ï"
                  Height          =   195
                  Index           =   23
                  Left            =   1740
                  RightToLeft     =   -1  'True
                  TabIndex        =   150
                  Top             =   240
                  Width           =   525
               End
            End
            Begin VB.Frame lbreg 
               BackColor       =   &H00E2E9E9&
               Caption         =   " «—ÌŒ «·Õ—ﬂÂ"
               Height          =   645
               Left            =   3120
               RightToLeft     =   -1  'True
               TabIndex        =   142
               Top             =   120
               Width           =   5475
               Begin MSComCtl2.DTPicker DtpDateFrom 
                  Height          =   330
                  Left            =   2520
                  TabIndex        =   143
                  Top             =   270
                  Width           =   1590
                  _ExtentX        =   2805
                  _ExtentY        =   582
                  _Version        =   393216
                  CheckBox        =   -1  'True
                  Format          =   177078275
                  CurrentDate     =   38887
               End
               Begin MSComCtl2.DTPicker DtpDateTo 
                  Height          =   330
                  Left            =   330
                  TabIndex        =   144
                  Top             =   240
                  Width           =   1590
                  _ExtentX        =   2805
                  _ExtentY        =   582
                  _Version        =   393216
                  CheckBox        =   -1  'True
                  Format          =   177078275
                  CurrentDate     =   38887
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "≈·Ï"
                  Height          =   195
                  Index           =   13
                  Left            =   1815
                  RightToLeft     =   -1  'True
                  TabIndex        =   146
                  Top             =   300
                  Width           =   480
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„‰"
                  Height          =   195
                  Index           =   14
                  Left            =   4080
                  RightToLeft     =   -1  'True
                  TabIndex        =   145
                  Top             =   330
                  Width           =   540
               End
            End
            Begin MSDataListLib.DataCombo DcbCustomer 
               Height          =   315
               Left            =   120
               TabIndex        =   154
               Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·„«·ﬂ"
               Top             =   840
               Width           =   5955
               _ExtentX        =   10504
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcbBranch 
               Height          =   315
               Left            =   8640
               TabIndex        =   155
               Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·„«·ﬂ"
               Top             =   840
               Width           =   4755
               _ExtentX        =   8387
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcbTypeTransport 
               Height          =   315
               Left            =   11580
               TabIndex        =   163
               Top             =   1200
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcboItems 
               Height          =   315
               Left            =   120
               TabIndex        =   166
               Top             =   1215
               Width           =   5955
               _ExtentX        =   10504
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label Label14 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "—ﬁ„ «—«„ﬂÊ"
               ForeColor       =   &H00000000&
               Height          =   270
               Index           =   0
               Left            =   10380
               TabIndex        =   228
               Top             =   1320
               Width           =   1590
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«·’‰›"
               Height          =   285
               Index           =   5
               Left            =   7320
               RightToLeft     =   -1  'True
               TabIndex        =   168
               Top             =   1200
               Width           =   915
            End
            Begin VB.Label Label15 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   510
               Left            =   4965
               RightToLeft     =   -1  'True
               TabIndex        =   167
               Top             =   1200
               Width           =   1170
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«‰Ê«⁄ «·‰ﬁ·"
               Height          =   315
               Index           =   72
               Left            =   12915
               RightToLeft     =   -1  'True
               TabIndex        =   164
               Top             =   1200
               Width           =   1530
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«·⁄„Ì·"
               Height          =   285
               Index           =   3
               Left            =   7350
               RightToLeft     =   -1  'True
               TabIndex        =   157
               Top             =   840
               Width           =   915
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·›—⁄"
               Height          =   195
               Index           =   20
               Left            =   13800
               RightToLeft     =   -1  'True
               TabIndex        =   156
               Top             =   840
               Width           =   540
            End
         End
         Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid5 
            Height          =   3585
            Left            =   120
            TabIndex        =   158
            Top             =   720
            Width           =   14625
            _cx             =   25797
            _cy             =   6324
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
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmProjectSearch.frx":0B87
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
         Begin ImpulseButton.ISButton ISButton11 
            Height          =   375
            Left            =   6150
            TabIndex        =   159
            Top             =   6195
            Width           =   2355
            _ExtentX        =   4154
            _ExtentY        =   661
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
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            LowerToggledContent=   0   'False
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton ISButton12 
            Height          =   375
            Left            =   3600
            TabIndex        =   160
            Top             =   6195
            Width           =   2355
            _ExtentX        =   4154
            _ExtentY        =   661
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
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            LowerToggledContent=   0   'False
            ColorTextShadow =   4210752
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   645
            Index           =   5
            Left            =   0
            TabIndex        =   169
            TabStop         =   0   'False
            Top             =   0
            Width           =   14835
            _cx             =   26167
            _cy             =   1138
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
            Appearance      =   4
            MousePointer    =   0
            Version         =   801
            BackColor       =   16777215
            ForeColor       =   4210688
            FloodColor      =   6553600
            ForeColorDisabled=   -2147483631
            Picture         =   "FrmProjectSearch.frx":0D0B
            Caption         =   "    »ÕÀ  ›Ê« Ì— «·⁄„·«¡  "
            Align           =   0
            AutoSizeChildren=   7
            BorderWidth     =   0
            ChildSpacing    =   0
            Splitter        =   0   'False
            FloodDirection  =   0
            FloodPercent    =   0
            CaptionPos      =   6
            WordWrap        =   -1  'True
            MaxChildSize    =   0
            MinChildSize    =   0
            TagWidth        =   0
            TagPosition     =   0
            Style           =   0
            TagSplit        =   2
            PicturePos      =   0
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
         Begin ImpulseButton.ISButton ISButton13 
            Height          =   375
            Left            =   8760
            TabIndex        =   170
            Top             =   6195
            Width           =   2355
            _ExtentX        =   4154
            _ExtentY        =   661
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
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            LowerToggledContent=   0   'False
            ColorTextShadow =   4210752
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            Caption         =   "0"
            Height          =   252
            Left            =   13440
            RightToLeft     =   -1  'True
            TabIndex        =   162
            Top             =   840
            Visible         =   0   'False
            Width           =   612
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            Caption         =   "0"
            Height          =   252
            Index           =   1
            Left            =   -384
            RightToLeft     =   -1  'True
            TabIndex        =   161
            Top             =   3240
            Visible         =   0   'False
            Width           =   612
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic6 
         Height          =   6720
         Left            =   17040
         TabIndex        =   171
         TabStop         =   0   'False
         Top             =   45
         Width           =   14805
         _cx             =   26114
         _cy             =   11853
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
            Height          =   1725
            Index           =   12
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   172
            Top             =   4320
            Width           =   14595
            Begin VB.Frame Frame3 
               BackColor       =   &H00E2E9E9&
               Caption         =   " «—ÌŒ «·Õ—ﬂÂ"
               Height          =   645
               Left            =   3120
               RightToLeft     =   -1  'True
               TabIndex        =   180
               Top             =   120
               Width           =   5475
               Begin MSComCtl2.DTPicker DtpDateFrom2 
                  Height          =   330
                  Left            =   2460
                  TabIndex        =   181
                  Top             =   240
                  Width           =   1590
                  _ExtentX        =   2805
                  _ExtentY        =   582
                  _Version        =   393216
                  CheckBox        =   -1  'True
                  Format          =   176947203
                  CurrentDate     =   38887
               End
               Begin MSComCtl2.DTPicker DtpDateTo2 
                  Height          =   330
                  Left            =   330
                  TabIndex        =   182
                  Top             =   240
                  Width           =   1590
                  _ExtentX        =   2805
                  _ExtentY        =   582
                  _Version        =   393216
                  CheckBox        =   -1  'True
                  Format          =   176947203
                  CurrentDate     =   38887
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„‰"
                  Height          =   195
                  Index           =   26
                  Left            =   4080
                  RightToLeft     =   -1  'True
                  TabIndex        =   184
                  Top             =   330
                  Width           =   540
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "≈·Ï"
                  Height          =   195
                  Index           =   25
                  Left            =   1815
                  RightToLeft     =   -1  'True
                  TabIndex        =   183
                  Top             =   300
                  Width           =   480
               End
            End
            Begin VB.Frame Frame2 
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ «·Õ—ﬂÂ"
               Height          =   645
               Left            =   8640
               RightToLeft     =   -1  'True
               TabIndex        =   175
               Top             =   120
               Width           =   5835
               Begin VB.TextBox TxtIDTO2 
                  Alignment       =   1  'Right Justify
                  Height          =   345
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   177
                  Top             =   240
                  Width           =   1275
               End
               Begin VB.TextBox TxtIDFrom2 
                  Alignment       =   1  'Right Justify
                  Height          =   345
                  Left            =   2760
                  RightToLeft     =   -1  'True
                  TabIndex        =   176
                  Top             =   240
                  Width           =   1275
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "≈·Ï"
                  Height          =   195
                  Index           =   24
                  Left            =   1740
                  RightToLeft     =   -1  'True
                  TabIndex        =   179
                  Top             =   240
                  Width           =   525
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„‰"
                  Height          =   195
                  Index           =   21
                  Left            =   3975
                  RightToLeft     =   -1  'True
                  TabIndex        =   178
                  Top             =   240
                  Width           =   540
               End
            End
            Begin VB.CheckBox Check6 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·„Ê—œ »«·ﬂ«„· ›ﬁÿ"
               Height          =   375
               Left            =   5760
               RightToLeft     =   -1  'True
               TabIndex        =   174
               Top             =   2760
               Value           =   1  'Checked
               Visible         =   0   'False
               Width           =   2385
            End
            Begin VB.TextBox Text4 
               Alignment       =   2  'Center
               Height          =   315
               Left            =   6180
               TabIndex        =   173
               Top             =   840
               Width           =   1035
            End
            Begin MSDataListLib.DataCombo DcbCustomer2 
               Height          =   315
               Left            =   120
               TabIndex        =   185
               Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·„«·ﬂ"
               Top             =   840
               Width           =   5955
               _ExtentX        =   10504
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcbBranch2 
               Height          =   315
               Left            =   8640
               TabIndex        =   186
               Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·„«·ﬂ"
               Top             =   840
               Width           =   4755
               _ExtentX        =   8387
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo dcproject 
               Height          =   315
               Left            =   120
               TabIndex        =   198
               Top             =   1200
               Visible         =   0   'False
               Width           =   7095
               _ExtentX        =   12515
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·›—⁄"
               Height          =   195
               Index           =   28
               Left            =   13800
               RightToLeft     =   -1  'True
               TabIndex        =   190
               Top             =   840
               Width           =   540
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«·⁄„Ì·"
               Height          =   285
               Index           =   7
               Left            =   7350
               RightToLeft     =   -1  'True
               TabIndex        =   189
               Top             =   840
               Width           =   915
            End
            Begin VB.Label Label17 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   510
               Left            =   4965
               RightToLeft     =   -1  'True
               TabIndex        =   188
               Top             =   1200
               Width           =   1170
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«·„‘—Ê⁄"
               Height          =   285
               Index           =   6
               Left            =   7320
               RightToLeft     =   -1  'True
               TabIndex        =   187
               Top             =   1200
               Visible         =   0   'False
               Width           =   915
            End
         End
         Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid6 
            Height          =   3585
            Left            =   120
            TabIndex        =   191
            Top             =   720
            Width           =   14625
            _cx             =   25797
            _cy             =   6324
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
            Cols            =   6
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmProjectSearch.frx":19E5
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
         Begin ImpulseButton.ISButton ISButton10 
            Height          =   375
            Left            =   6150
            TabIndex        =   192
            Top             =   6195
            Width           =   2355
            _ExtentX        =   4154
            _ExtentY        =   661
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
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            LowerToggledContent=   0   'False
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton ISButton14 
            Cancel          =   -1  'True
            Height          =   375
            Left            =   3600
            TabIndex        =   193
            Top             =   6195
            Width           =   2355
            _ExtentX        =   4154
            _ExtentY        =   661
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
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            LowerToggledContent=   0   'False
            ColorTextShadow =   4210752
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   645
            Index           =   0
            Left            =   0
            TabIndex        =   194
            TabStop         =   0   'False
            Top             =   0
            Width           =   14835
            _cx             =   26167
            _cy             =   1138
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
            Appearance      =   4
            MousePointer    =   0
            Version         =   801
            BackColor       =   16777215
            ForeColor       =   4210688
            FloodColor      =   6553600
            ForeColorDisabled=   -2147483631
            Picture         =   "FrmProjectSearch.frx":1AD0
            Caption         =   "»ÕÀ «·›Ê« Ì— «·Œœ„Ì… "
            Align           =   0
            AutoSizeChildren=   7
            BorderWidth     =   0
            ChildSpacing    =   0
            Splitter        =   0   'False
            FloodDirection  =   0
            FloodPercent    =   0
            CaptionPos      =   6
            WordWrap        =   -1  'True
            MaxChildSize    =   0
            MinChildSize    =   0
            TagWidth        =   0
            TagPosition     =   0
            Style           =   0
            TagSplit        =   2
            PicturePos      =   0
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
         Begin ImpulseButton.ISButton ISButton15 
            Height          =   375
            Left            =   8760
            TabIndex        =   195
            Top             =   6195
            Width           =   2355
            _ExtentX        =   4154
            _ExtentY        =   661
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
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            LowerToggledContent=   0   'False
            ColorTextShadow =   4210752
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            Caption         =   "0"
            Height          =   252
            Left            =   -384
            RightToLeft     =   -1  'True
            TabIndex        =   197
            Top             =   3240
            Visible         =   0   'False
            Width           =   612
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            Caption         =   "0"
            Height          =   252
            Left            =   13440
            RightToLeft     =   -1  'True
            TabIndex        =   196
            Top             =   840
            Visible         =   0   'False
            Width           =   612
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic3 
         Height          =   6720
         Index           =   1
         Left            =   17340
         TabIndex        =   203
         TabStop         =   0   'False
         Top             =   45
         Width           =   14805
         _cx             =   26114
         _cy             =   11853
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
         Begin VB.Frame Fra 
            Height          =   2088
            Index           =   13
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   204
            Top             =   3840
            Width           =   14475
            Begin VB.TextBox TXTOrDer_no 
               Alignment       =   2  'Center
               Height          =   315
               Left            =   7560
               RightToLeft     =   -1  'True
               TabIndex        =   225
               Top             =   450
               Visible         =   0   'False
               Width           =   2175
            End
            Begin VB.Frame Fra 
               Caption         =   " «—ÌŒ"
               Height          =   1095
               Index           =   14
               Left            =   420
               RightToLeft     =   -1  'True
               TabIndex        =   218
               Top             =   570
               Width           =   6012
               Begin MSComCtl2.DTPicker txtFromDate 
                  Height          =   330
                  Left            =   3240
                  TabIndex        =   219
                  Top             =   390
                  Width           =   1590
                  _ExtentX        =   2805
                  _ExtentY        =   582
                  _Version        =   393216
                  CheckBox        =   -1  'True
                  Format          =   178061315
                  CurrentDate     =   38887
               End
               Begin MSComCtl2.DTPicker txtToDate 
                  Height          =   336
                  Left            =   216
                  TabIndex        =   220
                  Top             =   396
                  Width           =   1596
                  _ExtentX        =   2805
                  _ExtentY        =   582
                  _Version        =   393216
                  CheckBox        =   -1  'True
                  Format          =   178061315
                  CurrentDate     =   38887
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  Caption         =   "≈·Ï"
                  Height          =   192
                  Index           =   29
                  Left            =   2172
                  RightToLeft     =   -1  'True
                  TabIndex        =   222
                  Top             =   420
                  Width           =   372
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  Caption         =   "„‰"
                  Height          =   195
                  Index           =   27
                  Left            =   5100
                  RightToLeft     =   -1  'True
                  TabIndex        =   221
                  Top             =   420
                  Width           =   420
               End
            End
            Begin VB.TextBox txtCustomerCode 
               Alignment       =   2  'Center
               Height          =   288
               Left            =   11880
               TabIndex        =   207
               Top             =   1320
               Width           =   1275
            End
            Begin VB.CheckBox Check7 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·„Ê—œ »«·ﬂ«„· ›ﬁÿ"
               Height          =   375
               Left            =   5760
               RightToLeft     =   -1  'True
               TabIndex        =   206
               Top             =   2760
               Value           =   1  'Checked
               Visible         =   0   'False
               Width           =   2385
            End
            Begin VB.TextBox txtID 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   10860
               RightToLeft     =   -1  'True
               TabIndex        =   205
               Top             =   480
               Width           =   2175
            End
            Begin MSDataListLib.DataCombo DataBranch 
               Height          =   315
               Left            =   7560
               TabIndex        =   208
               Top             =   960
               Width           =   5610
               _ExtentX        =   9895
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   16777215
               ListField       =   ""
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
            Begin MSDataListLib.DataCombo DcbCus 
               Height          =   315
               Left            =   7560
               TabIndex        =   209
               Top             =   1320
               Width           =   4155
               _ExtentX        =   7329
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo cmbPaymentType 
               Height          =   315
               Left            =   7560
               TabIndex        =   229
               Top             =   1710
               Width           =   4185
               _ExtentX        =   7382
               _ExtentY        =   556
               _Version        =   393216
               ListField       =   "6"
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "‰Ê⁄ «·œ›⁄"
               Height          =   255
               Index           =   37
               Left            =   12645
               RightToLeft     =   -1  'True
               TabIndex        =   230
               Top             =   1740
               Width           =   1260
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               Caption         =   "—ﬁ„ «·ﬂ«— "
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   9
               Left            =   9840
               TabIndex        =   226
               Top             =   450
               Visible         =   0   'False
               Width           =   975
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«·⁄„Ì·"
               Height          =   285
               Index           =   8
               Left            =   13320
               RightToLeft     =   -1  'True
               TabIndex        =   212
               Top             =   1320
               Width           =   675
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               Caption         =   "«·›—⁄"
               Height          =   285
               Index           =   30
               Left            =   13290
               TabIndex        =   211
               Top             =   960
               Width           =   645
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               Caption         =   "—ﬁ„ «·Õ—ﬂ…"
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   6
               Left            =   13140
               TabIndex        =   210
               Top             =   480
               Width           =   975
            End
         End
         Begin VSFlex8UCtl.VSFlexGrid GrdF 
            Height          =   3585
            Left            =   120
            TabIndex        =   213
            Top             =   120
            Width           =   14505
            _cx             =   25585
            _cy             =   6324
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
            Cols            =   8
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmProjectSearch.frx":27AA
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
         Begin ImpulseButton.ISButton ISButton16 
            Height          =   375
            Left            =   8340
            TabIndex        =   214
            Top             =   6195
            Width           =   2355
            _ExtentX        =   4154
            _ExtentY        =   661
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
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            ColorToggledHoverText=   16711680
            LowerToggledContent=   0   'False
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton ISButton17 
            Height          =   375
            Left            =   5910
            TabIndex        =   215
            Top             =   6195
            Width           =   2355
            _ExtentX        =   4154
            _ExtentY        =   661
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
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            LowerToggledContent=   0   'False
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton ISButton18 
            Height          =   375
            Left            =   3480
            TabIndex        =   216
            Top             =   6195
            Width           =   2355
            _ExtentX        =   4154
            _ExtentY        =   661
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
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            LowerToggledContent=   0   'False
            ColorTextShadow =   4210752
         End
         Begin VB.Label Label21 
            Alignment       =   1  'Right Justify
            Caption         =   "0"
            Height          =   252
            Left            =   -384
            RightToLeft     =   -1  'True
            TabIndex        =   217
            Top             =   3240
            Visible         =   0   'False
            Width           =   612
         End
      End
   End
End
Attribute VB_Name = "FrmProjectSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Indx As Integer
Public Indx2 As Integer
Dim rs As ADODB.Recordset
Dim cSearchDcbo As clsDCboSearch


Private Sub cmd_clear_Click()

clear_all Me
DTP_Me_From.value = ""
DTP_Me_To.value = ""


End Sub

Private Sub Cmd_Click(Index As Integer)
    Dim Msg As String
    On Error GoTo ErrTrap
Set rs = New ADODB.Recordset
    Select Case Index

        Case 0

             If rs.State = adStateOpen Then
                 rs.Close
             End If

            rs.Open Build_Sql(0), Cn, adOpenStatic, adLockOptimistic, adCmdText

            If rs.RecordCount < 1 Then
                FG.Clear flexClearScrollable, flexClearEverything
                FG.rows = 2
                Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Exit Sub
            End If

            Retrive 0
            FG.SetFocus
               Case 3

            If rs.State = adStateOpen Then
                rs.Close
            End If

            rs.Open Build_Sql(1), Cn, adOpenStatic, adLockOptimistic, adCmdText

            If rs.RecordCount < 1 Then
                FG.Clear flexClearScrollable, flexClearEverything
                FG.rows = 2
                If SystemOptions.UserInterface = ArabicInterface Then
                 Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
                Else
                 Msg = "No Data"
                End If
                
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Exit Sub
            End If

            Retrive 1
            FG.SetFocus
   Case 4

            If rs.State = adStateOpen Then
                rs.Close
            End If

            rs.Open Build_Sql(2), Cn, adOpenStatic, adLockOptimistic, adCmdText

            If rs.RecordCount < 1 Then
                FG.Clear flexClearScrollable, flexClearEverything
                FG.rows = 2
                Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Exit Sub
            End If

            Retrive 2
            FG.SetFocus
        Case 1
            clear_all Me
            FG.Clear flexClearScrollable, flexClearEverything
            FG.rows = 1
  FrmDTStartDate.value = ""
    TODTStartDate.value = ""
    FrmDTEnddate.value = ""
    ToDTEnddate.value = ""
        Case 2
            Unload Me
    End Select

    Exit Sub
ErrTrap:

    If Err.Number = -2147217900 Then
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & CHR(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… „⁄«ÌÌ— «·»ÕÀ Ê√⁄œ «·„Õ«Ê·…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

End Sub

Private Sub Cmd_Exit_Click()
Unload Me
End Sub

Private Sub Cmd_Search_Click()

GetDataMeasure


End Sub

Private Sub GetDataMeasure()
    Dim sql As String
    Dim StrSQL As String
    Dim Begin As Boolean
  '  Public Current_branch As Integer
   ' Public Current_branchSql As String

    Dim StrWhere As String
   ' On Error GoTo ErrTrap


StrSQL = StrSQL & " SELECT dbo.TBL_measureMent.ID, dbo.TBL_measureMent.Cust_name_ID, dbo.TBL_measureMent.Cust_Mobile, dbo.TBL_measureMent.Cust_City,"
StrSQL = StrSQL & "        dbo.TBL_measureMent.Cust_Time, dbo.TBL_measureMent.Cust_District, dbo.TBL_measureMent.Date_Order, dbo.TBL_measureMent.Date_measureMent,"
StrSQL = StrSQL & "  dbo.TBL_measureMent.level1, dbo.TBL_measureMent.WCMen1, dbo.TBL_measureMent.WCWomen1, dbo.TBL_measureMent.WCChildren1,"
StrSQL = StrSQL & "            dbo.TBL_measureMent.WCGirls1, dbo.TBL_measureMent.WCCount1, dbo.TBL_measureMent.WCNote1, dbo.TBL_measureMent.laundryMen1,"
StrSQL = StrSQL & "             dbo.TBL_measureMent.laundryWomen1, dbo.TBL_measureMent.laundryChildren1, dbo.TBL_measureMent.laundryGirls1, dbo.TBL_measureMent.laundryCount1,"
StrSQL = StrSQL & "              dbo.TBL_measureMent.laundryNote1, dbo.TBL_measureMent.laundryareaMen1, dbo.TBL_measureMent.laundryareaWomen1,"
StrSQL = StrSQL & "              dbo.TBL_measureMent.laundryareaChildren1, dbo.TBL_measureMent.laundryareaGirls1, dbo.TBL_measureMent.laundryareaCount1,"
StrSQL = StrSQL & "               dbo.TBL_measureMent.laundryareaNote1, dbo.TBL_measureMent.MainHall1, dbo.TBL_measureMent.MainHallCount1, dbo.TBL_measureMent.MainHallNote1,"
StrSQL = StrSQL & "               dbo.TBL_measureMent.kitchen1, dbo.TBL_measureMent.kitchenCount1, dbo.TBL_measureMent.kitchenNote1, dbo.TBL_measureMent.BoardMen1,"
StrSQL = StrSQL & "                dbo.TBL_measureMent.BoardWomen1, dbo.TBL_measureMent.BoardCount1, dbo.TBL_measureMent.BoardNote1, dbo.TBL_measureMent.MaklatMen1,"
StrSQL = StrSQL & "               dbo.TBL_measureMent.MaklatWomen1, dbo.TBL_measureMent.MaklatCount1, dbo.TBL_measureMent.MaklatNote1, dbo.TBL_measureMent.EntranceMen1,"
StrSQL = StrSQL & "                dbo.TBL_measureMent.EntranceWomen1, dbo.TBL_measureMent.EntranceCount1, dbo.TBL_measureMent.EntranceNote1, dbo.TBL_measureMent.Dorginside1,"
StrSQL = StrSQL & "               dbo.TBL_measureMent.DorgOutside1, dbo.TBL_measureMent.DorgTwacheh1, dbo.TBL_measureMent.DorgCount1, dbo.TBL_measureMent.DorgNote1,"
 StrSQL = StrSQL & "               dbo.TBL_measureMent.ElevatorInside1, dbo.TBL_measureMent.ElevatorOutSide1, dbo.TBL_measureMent.ElevatorCount1, dbo.TBL_measureMent.ElevatorNote1,"
 StrSQL = StrSQL & "               dbo.TBL_measureMent.HoshInside1, dbo.TBL_measureMent.HoshOutSide1, dbo.TBL_measureMent.HoshCount1, dbo.TBL_measureMent.HoshNote1,"
  StrSQL = StrSQL & "              dbo.TBL_measureMent.MainRoom1, dbo.TBL_measureMent.MRoom1, dbo.TBL_measureMent.LRoom1, dbo.TBL_measureMent.MainRoomCount1,"
   StrSQL = StrSQL & "             dbo.TBL_measureMent.MainRoomNote1, dbo.TBL_measureMent.Na3laNormal1, dbo.TBL_measureMent.Na3laDorg1, dbo.TBL_measureMent.Na3laCount1,"
    StrSQL = StrSQL & "            dbo.TBL_measureMent.Na3laNote1, dbo.TBL_measureMent.ClothesInside1, dbo.TBL_measureMent.ClothesOutInside1, dbo.TBL_measureMent.ClothesCount1,"
   StrSQL = StrSQL & "             dbo.TBL_measureMent.ClothesNote1, dbo.TBL_measureMent.ParkingGround1, dbo.TBL_measureMent.ParkingGdar1, dbo.TBL_measureMent.ParkingCount1,"
   StrSQL = StrSQL & "             dbo.TBL_measureMent.ParkingNote1, dbo.TBL_measureMent.Office1, dbo.TBL_measureMent.OfficeCount1, dbo.TBL_measureMent.OfficeNote1,"
   StrSQL = StrSQL & "             dbo.TBL_measureMent.Cust_Mobile2, dbo.TBL_measureMent.Cust_City2, dbo.TBL_measureMent.Cust_Time2, dbo.TBL_measureMent.Cust_District2,"
  StrSQL = StrSQL & "              dbo.TBL_measureMent.Date_Order2, dbo.TBL_measureMent.Date_measureMent2, dbo.TBL_measureMent.level2, dbo.TBL_measureMent.WCMen2,"
   StrSQL = StrSQL & "             dbo.TBL_measureMent.WCWomen2, dbo.TBL_measureMent.WCChildren2, dbo.TBL_measureMent.WCGirls2, dbo.TBL_measureMent.WCCount2,"
  StrSQL = StrSQL & "              dbo.TBL_measureMent.WCNote2, dbo.TBL_measureMent.laundryMen2, dbo.TBL_measureMent.laundryWomen2, dbo.TBL_measureMent.laundryChildren2,"
  StrSQL = StrSQL & "              dbo.TBL_measureMent.laundryGirls2, dbo.TBL_measureMent.laundryCount2, dbo.TBL_measureMent.laundryNote2, dbo.TBL_measureMent.laundryareaMen2,"
 StrSQL = StrSQL & "               dbo.TBL_measureMent.laundryareaWomen2, dbo.TBL_measureMent.laundryareaChildren2, dbo.TBL_measureMent.laundryareaGirls2,"
   StrSQL = StrSQL & "             dbo.TBL_measureMent.laundryareaCount2, dbo.TBL_measureMent.laundryareaNote2, dbo.TBL_measureMent.MainHall2, dbo.TBL_measureMent.MainHallCount2,"
   StrSQL = StrSQL & "             dbo.TBL_measureMent.MainHallNote2, dbo.TBL_measureMent.kitchen2, dbo.TBL_measureMent.kitchenCount2, dbo.TBL_measureMent.kitchenNote2,"
  StrSQL = StrSQL & "              dbo.TBL_measureMent.BoardMen2, dbo.TBL_measureMent.BoardWomen2, dbo.TBL_measureMent.BoardCount2, dbo.TBL_measureMent.BoardNote2,"
   StrSQL = StrSQL & "             dbo.TBL_measureMent.MaklatMen2, dbo.TBL_measureMent.MaklatWomen2, dbo.TBL_measureMent.MaklatCount2, dbo.TBL_measureMent.MaklatNote2,"
    StrSQL = StrSQL & "            dbo.TBL_measureMent.EntranceMen2, dbo.TBL_measureMent.EntranceWomen2, dbo.TBL_measureMent.EntranceCount2,"
    StrSQL = StrSQL & "            dbo.TBL_measureMent.EntranceNote2, dbo.TBL_measureMent.Dorginside2, dbo.TBL_measureMent.DorgOutside2, dbo.TBL_measureMent.DorgTwacheh2,"
    StrSQL = StrSQL & "            dbo.TBL_measureMent.DorgCount2, dbo.TBL_measureMent.DorgNote2, dbo.TBL_measureMent.ElevatorInside2, dbo.TBL_measureMent.ElevatorOutSide2,"
   StrSQL = StrSQL & "             dbo.TBL_measureMent.ElevatorCount2, dbo.TBL_measureMent.ElevatorNote2, dbo.TBL_measureMent.HoshInside2, dbo.TBL_measureMent.HoshOutSide2,"
   StrSQL = StrSQL & "             dbo.TBL_measureMent.HoshCount2, dbo.TBL_measureMent.HoshNote2, dbo.TBL_measureMent.MainRoom2, dbo.TBL_measureMent.MRoom2,"
    StrSQL = StrSQL & "            dbo.TBL_measureMent.LRoom2, dbo.TBL_measureMent.MainRoomCount2, dbo.TBL_measureMent.MainRoomNote2, dbo.TBL_measureMent.Na3laNormal2,"
   StrSQL = StrSQL & "             dbo.TBL_measureMent.Na3laDorg2, dbo.TBL_measureMent.Na3laCount2, dbo.TBL_measureMent.Na3laNote2, dbo.TBL_measureMent.ClothesInside2,"
    StrSQL = StrSQL & "            dbo.TBL_measureMent.ClothesOutInside2, dbo.TBL_measureMent.ClothesCount2, dbo.TBL_measureMent.ClothesNote2,"
    StrSQL = StrSQL & "            dbo.TBL_measureMent.ParkingGround2, dbo.TBL_measureMent.ParkingGdar2, dbo.TBL_measureMent.ParkingCount2, dbo.TBL_measureMent.ParkingNote2,"
   StrSQL = StrSQL & "             dbo.TBL_measureMent.Office2, dbo.TBL_measureMent.OfficeCount2, dbo.TBL_measureMent.OfficeNote2, dbo.TBL_measureMent.UserID,"
     StrSQL = StrSQL & "           dbo.TblCustemers.CusName , dbo.TblCustemers.CusNamee"
   StrSQL = StrSQL & " FROM  dbo.TBL_measureMent INNER JOIN"
    StrSQL = StrSQL & "  dbo.TblCustemers ON dbo.TBL_measureMent.Cust_name_ID = dbo.TblCustemers.CusID"

    StrSQL = StrSQL & " where 1=1 AND (NOT (dbo.TBL_measureMent.ID IS NULL))"
    
    
         
        StrSQL = " SELECT  DISTINCT TBL_measureMent.ID,"
        'IsNull(T2.Level,0) as level1,T2.LevelName ,"
       StrSQL = StrSQL & "  TBL_measureMent.CustomerName as CusName,TBL_measureMent.Cust_Mobile,TBL_measureMent.Cust_City,"
       StrSQL = StrSQL & " TBL_measureMent.Cust_Time,TBL_measureMent.Cust_District,TBL_measureMent.Date_Order,TBL_measureMent.Date_measureMent"
       
       StrSQL = StrSQL & " From TBL_measureMent"
        StrSQL = StrSQL & " where 1=1 AND (NOT (dbo.TBL_measureMent.ID IS NULL))"
    Begin = True
  
   'Sql = StrSQL & " and  dbo.TBL_measureMent.ID in(" & Current_branchSql & ")"

    If Txt_MeasureID.text <> "" Then
        StrWhere = StrWhere + " and dbo.TBL_measureMent.ID like'%" & (Txt_MeasureID.text) & "%'"
      End If
    If Txt_CustomerCode.text <> "" Then
          StrWhere = StrWhere + " and ( dbo.TBL_measureMent.CustomerName LIKE '%" & Trim(Txt_CustomerCode.text) & "%')"
    End If
'     If DcbCustmer.Text <> "" Then
'              StrWhere = StrWhere + " and dbo.TblCustemers.CusName ='" & (DcbCustmer.Text) & "'"
'      End If
   If Txt_Mobile.text <> "" Then
              StrWhere = StrWhere + " and  dbo.TBL_measureMent.Cust_Mobile ='" & (Txt_Mobile.text) & "'"
      End If
     If Txt_City.text <> "" Then
              StrWhere = StrWhere + " and dbo.TBL_measureMent.Cust_City ='" & (Txt_City.text) & "'"
      End If
      If TXT_District.text <> "" Then
              StrWhere = StrWhere + " and dbo.TBL_measureMent.Cust_District ='" & (TXT_District.text) & "'"
      End If
      If Not IsNull(Me.DTP_Me_From.value) Then
             StrWhere = StrWhere & " AND dbo.TBL_measureMent.Date_Order>=" & SQLDate(Me.DTP_Me_From.value, True) & ""
      End If
       If Not IsNull(Me.DTP_Me_To.value) Then
             StrWhere = StrWhere & " AND dbo.TBL_measureMent.Date_Order <=" & SQLDate(Me.DTP_Me_To.value, True) & ""
      End If
      
    
      sql = StrSQL + StrWhere + " order by dbo.TBL_measureMent.ID  "
   
   ''----------------------------------------------------------------------
   
    Dim Num As Integer
    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
    If Not (rs.EOF Or rs.BOF) Then
        VSFlexGrid1.rows = rs.RecordCount + 1
        
        For Num = 1 To rs.RecordCount
       
            With VSFlexGrid1
                .TextMatrix(Num, .ColIndex("Count")) = Num
                .TextMatrix(Num, .ColIndex("id")) = IIf(IsNull(rs("id").value), "", Trim(rs("id").value))
                .TextMatrix(Num, .ColIndex("Name")) = IIf(IsNull(rs("CusName").value), "", (rs("CusName").value))
                .TextMatrix(Num, .ColIndex("Mobile")) = IIf(IsNull(rs("Cust_Mobile").value), "", Trim(rs("Cust_Mobile").value))
                .TextMatrix(Num, .ColIndex("CityName")) = IIf(IsNull(rs("Cust_City").value), "", (rs("Cust_City").value))
                .TextMatrix(Num, .ColIndex("TTime")) = IIf(IsNull(rs("Cust_Time").value), "", (rs("Cust_Time").value))
                .TextMatrix(Num, .ColIndex("District")) = IIf(IsNull(rs("Cust_District").value), "", (rs("Cust_District").value))
                .TextMatrix(Num, .ColIndex("DateOrder")) = IIf(IsNull(rs("Date_Order").value), "", (rs("Date_Order").value))
       
            End With
            rs.MoveNext
        Next Num
      
End If
   
   ''----------------------------------------------------------------------

ErrTrap:

End Sub


Private Sub DcbCusTC_Click(Area As Integer)

TXtCustID.text = ""

    Dim DefaultSalesPersonId As Integer
    Dim fullcode As String

    GetCustomersDetail val(DcbCusTC.BoundText), DefaultSalesPersonId, fullcode

    TXtCustID.text = fullcode


End Sub

Private Sub DcbCustmer_Click(Area As Integer)


Txt_CustomerCode.text = ""

    Dim DefaultSalesPersonId As Integer
    Dim fullcode As String

    GetCustomersDetail val(DcbCustmer.BoundText), DefaultSalesPersonId, fullcode

    Txt_CustomerCode.text = fullcode


End Sub

Private Sub DcbCustomer_Change()
DcbCustomer_Click (0)
End Sub

Private Sub DcbCustomer_Click(Area As Integer)
    Dim fullcode As String
     Dim Dcombos As New ClsDataCombos
    GetCustomersDetail val(DcbCustomer.BoundText), , fullcode, 1
    Me.Text3.text = fullcode
End Sub

Private Sub DcbEmployee2_Click(Area As Integer)
 If val(DcbEmployee2.BoundText) = 0 Then
    Exit Sub
 End If


    Dim EmpCode  As String
 
    GetEmployeeIDFromCode , , DcbEmployee2.BoundText, EmpCode
    TxtSearchCode.text = EmpCode
End Sub

Private Sub fg_Click()
    On Error GoTo ErrTrap

    If Not FG.TextMatrix(FG.Row, 1) = "" Then
    
        If Me.lblSearchtype.Caption = 0 Then
            Projects.Retrive val(FG.TextMatrix(FG.Row, 1))
        ElseIf Me.lblSearchtype.Caption = 1 Then
          FrmCashing.DBCboClientName.BoundText = val(FG.TextMatrix(FG.Row, 1))
  
        ElseIf Me.lblSearchtype.Caption = 2 Then
         
    
        ElseIf Me.lblSearchtype.Caption = 3 Then
          '  FrmPayments.DBCboClientName.BoundText = val(FG.TextMatrix(FG.Row, 1))

        ElseIf Me.lblSearchtype.Caption = 4 Then
           
FrmExpenses5.dcproject.BoundText = (FG.TextMatrix(FG.Row, 15))

        ElseIf Me.lblSearchtype.Caption = 5 Then
           
FrmPayments.DBCboClientName.BoundText = (FG.TextMatrix(FG.Row, 15))
    
        ElseIf Me.lblSearchtype.Caption = 6 Then
            FrmExpenses3.dcproject.BoundText = (FG.TextMatrix(FG.Row, 15))
 ElseIf Me.lblSearchtype.Caption = 666 Then
            frmserviceInvoice.dcproject.BoundText = (FG.TextMatrix(FG.Row, 15))
ElseIf Me.lblSearchtype.Caption = 777 Then
            frmserviceInvoice.dcproject.BoundText = (FG.TextMatrix(FG.Row, 15))
           ' frmserviceInvoice.dcproject2.BoundText = (FG.TextMatrix(FG.row, 1))
            
       ElseIf Me.lblSearchtype.Caption = 7 Then
  frmProjectsReports.dcprojects.BoundText = (FG.TextMatrix(FG.Row, 1))
       ElseIf Me.lblSearchtype.Caption = 8 Then
 projectsbill.DataCombo2.BoundText = (FG.TextMatrix(FG.Row, 1))

       ElseIf Me.lblSearchtype.Caption = 9 Then
 FrmAccountingReport.dcprojects.BoundText = (FG.TextMatrix(FG.Row, 1))
      ElseIf Me.lblSearchtype.Caption = 10 Then
 Projects.Retrive val(FG.TextMatrix(FG.Row, 1))
       
ElseIf Me.lblSearchtype.Caption = 1010 Then
    FrmDestruction.dcproject.BoundText = (FG.TextMatrix(FG.Row, 16))
       ElseIf Me.lblSearchtype.Caption = 11 Then
           

           
 FrmDestruction.dcproject1.BoundText = (FG.TextMatrix(FG.Row, 1))
       ElseIf Me.lblSearchtype.Caption = 111 Then
 FrmDestructionRet.dcproject1.BoundText = (FG.TextMatrix(FG.Row, 1))
 
       ElseIf Me.lblSearchtype.Caption = 12 Then
           
   FrmPO6.dcproject1.BoundText = (FG.TextMatrix(FG.Row, 1))
   
    
       ElseIf Me.lblSearchtype.Caption = 120 Then
           
   'Frm_BusinessDialy.DcbProject.BoundText = (Fg.TextMatrix(Fg.Row, 1))
   
       ElseIf Me.lblSearchtype.Caption = 13 Then
           
 FrmAccEditJournal.dcprojects.BoundText = (FG.TextMatrix(FG.Row, 1))
        
       ElseIf Me.lblSearchtype.Caption = 14 Then
                         With FrmAccEditJournal.Fg_Journal
                          .TextMatrix(.Row, .ColIndex("projectid")) = (FG.TextMatrix(FG.Row, 1))
                         .TextMatrix(.Row, .ColIndex("project")) = (FG.TextMatrix(FG.Row, 3))
                             .TextMatrix(.Row, .ColIndex("pand")) = ""
                             .TextMatrix(.Row, .ColIndex("oper")) = ""
                                FrmAccEditJournal.Fg_Journal.SetFocus
                                        
                      End With

        
        
            ElseIf Me.lblSearchtype.Caption = 15 Then
                         With FrmChangedComponentData.Grid
                         'projectid
'project
                         .TextMatrix(.Row, .ColIndex("projectid2")) = (FG.TextMatrix(FG.Row, 1))
                        .TextMatrix(.Row, .ColIndex("project")) = (FG.TextMatrix(FG.Row, 3))
                   
                               FrmChangedComponentData.Grid.SetFocus
                                        
                    End With
ElseIf Me.lblSearchtype.Caption = 20 Then
           FrmWarrantyOffer.DcbProject.BoundText = (FG.TextMatrix(FG.Row, 1))
ElseIf Me.lblSearchtype.Caption = 21 Then
                         With FrmAccEditJournal4.Fg_Journal
                          .TextMatrix(.Row, .ColIndex("projectid")) = (FG.TextMatrix(FG.Row, 1))
                          .TextMatrix(.Row, .ColIndex("project")) = (FG.TextMatrix(FG.Row, 3))
                          .TextMatrix(.Row, .ColIndex("ProjectCode")) = (FG.TextMatrix(FG.Row, FG.ColIndex("Fullcode")))
                             .TextMatrix(.Row, .ColIndex("pand")) = ""
                             .TextMatrix(.Row, .ColIndex("oper")) = ""
                                FrmAccEditJournal4.Fg_Journal.SetFocus
                           End With
ElseIf Me.lblSearchtype.Caption = 22 Then
                         With FrmAccEditJournal3.Fg_Journal
                          .TextMatrix(.Row, .ColIndex("projectid")) = (FG.TextMatrix(FG.Row, 1))
                          .TextMatrix(.Row, .ColIndex("project")) = (FG.TextMatrix(FG.Row, 3))
                          .TextMatrix(.Row, .ColIndex("ProjectCode")) = (FG.TextMatrix(FG.Row, FG.ColIndex("Fullcode")))
                             .TextMatrix(.Row, .ColIndex("pand")) = ""
                             .TextMatrix(.Row, .ColIndex("oper")) = ""
                             .Fg_Journal.SetFocus
                      End With
ElseIf Me.lblSearchtype.Caption = 23 Then
                         With FrmExpenses5.Fg_Journal
                          .TextMatrix(.Row, .ColIndex("projectid2")) = (FG.TextMatrix(FG.Row, 1))
                          .TextMatrix(.Row, .ColIndex("project")) = (FG.TextMatrix(FG.Row, 3))
                          .TextMatrix(.Row, .ColIndex("PrjectCode")) = (FG.TextMatrix(FG.Row, FG.ColIndex("Fullcode")))
                             .TextMatrix(.Row, .ColIndex("pand")) = ""
                             .TextMatrix(.Row, .ColIndex("oper")) = ""
                             .Fg_Journal.SetFocus
                      End With

        ElseIf Me.lblSearchtype.Caption = 25 Then
                         With FrmExpenses30.Fg_Journal
                          .TextMatrix(.Row, .ColIndex("projectid2")) = (FG.TextMatrix(FG.Row, 1))
                          .TextMatrix(.Row, .ColIndex("project")) = (FG.TextMatrix(FG.Row, 3))
                          .TextMatrix(.Row, .ColIndex("ProjectCode")) = (FG.TextMatrix(FG.Row, FG.ColIndex("Fullcode")))
                             .TextMatrix(.Row, .ColIndex("pand")) = ""
                             .TextMatrix(.Row, .ColIndex("oper")) = ""
                             .Fg_Journal.SetFocus
                      End With
        ElseIf Me.lblSearchtype.Caption = 26 Then
                       FrmExpenses30.dcproject.BoundText = get_project_Account(val(FG.TextMatrix(FG.Row, 1)), "expanses_account")
        ElseIf Me.lblSearchtype.Caption = 27 Then
                         With FrmExpenses30.VSFlexGrid1
                          .TextMatrix(.Row, .ColIndex("projectid")) = (FG.TextMatrix(FG.Row, 1))
                          .TextMatrix(.Row, .ColIndex("project")) = (FG.TextMatrix(FG.Row, 3))
                          .TextMatrix(.Row, .ColIndex("ProjectCode")) = (FG.TextMatrix(FG.Row, FG.ColIndex("Fullcode")))
                             .TextMatrix(.Row, .ColIndex("pand")) = ""
                             .TextMatrix(.Row, .ColIndex("operid")) = ""
                             .VSFlexGrid1.SetFocus
                      End With
        ElseIf Me.lblSearchtype.Caption = 28 Then
          FrmInpout.DcbProject.BoundText = (FG.TextMatrix(FG.Row, 1))
        ElseIf Me.lblSearchtype.Caption = 29 Then
          FrmBillBuy.dcproject.BoundText = (FG.TextMatrix(FG.Row, 1))
        ElseIf Me.lblSearchtype.Caption = 30 Then
          FrmEmpSalary3.dcproject.BoundText = (FG.TextMatrix(FG.Row, 1))
          FrmEmpSalary3.Text1.text = (FG.TextMatrix(FG.Row, FG.ColIndex("Fullcode")))
        ElseIf Me.lblSearchtype.Caption = 31 Then
          FrmEmpSalary3.DcbProject1.BoundText = (FG.TextMatrix(FG.Row, 1))
          FrmEmpSalary3.Text2.text = (FG.TextMatrix(FG.Row, FG.ColIndex("Fullcode")))
        ElseIf Me.lblSearchtype.Caption = 32 Then
          FrmSearchEmpSalary3.dcproject.BoundText = (FG.TextMatrix(FG.Row, 1))
        
       ElseIf Me.lblSearchtype.Caption = 33 Then
          frmsalebill.dcproject.BoundText = (FG.TextMatrix(FG.Row, 1))
               
             ElseIf Me.lblSearchtype.Caption = 34 Then
          FrmPripaidExpenses.dcproject.BoundText = (FG.TextMatrix(FG.Row, 1))
          
          
                 ElseIf Me.lblSearchtype.Caption = 35 Then
                                   With FrmPripaidExpenses.Grid
                          .TextMatrix(.Row, .ColIndex("ProjectID")) = (FG.TextMatrix(FG.Row, 1))
                          .TextMatrix(.Row, .ColIndex("ProjectName")) = (FG.TextMatrix(FG.Row, 3))
                          .TextMatrix(.Row, .ColIndex("PFuLLCode")) = (FG.TextMatrix(FG.Row, FG.ColIndex("Fullcode")))
                          
                             .Grid.SetFocus
                      End With
                      
          
        
                 ElseIf Me.lblSearchtype.Caption = 36 Then
                                   With FrmPaytAmortization.Grid
                          .TextMatrix(.Row, .ColIndex("ProjectID")) = (FG.TextMatrix(FG.Row, 1))
                          .TextMatrix(.Row, .ColIndex("ProjectName")) = (FG.TextMatrix(FG.Row, 3))
                          .TextMatrix(.Row, .ColIndex("PFuLLCode")) = (FG.TextMatrix(FG.Row, FG.ColIndex("Fullcode")))
                          
                             .Grid.SetFocus
                      End With
                      
               
        End If
 
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Fg_DblClick()
    fg_Click
    Cmd_Click (2)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        If Not FG.TextMatrix(FG.Row, FG.ColIndex("Code")) = "" Then
            fg_Click
        Else
            Cmd_Click (0)
        End If
    End If

    If Shift = 2 Then
        If KeyCode = vbKeyX Then
            Cmd_Click (2)
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Load()
    On Error GoTo ErrTrap
     
    C1Tab1.TabVisible(0) = False
    C1Tab1.TabVisible(1) = False
    C1Tab1.TabVisible(2) = False
    C1Tab1.TabVisible(3) = False
    C1Tab1.TabVisible(4) = False
    C1Tab1.TabVisible(5) = False
    RdAuto_Manual(2).value = True
    C1Tab1.TabVisible(Indx) = True
    C1Tab1.CurrTab = Indx
    DtpDateFrom.value = Date
    Dim mFromDate  As String
    Dim mToDate  As String
    mFromDate = "1-1-" & year(Date)
    mToDate = "31-12-" & year(Date)
      GrdF.ColHidden(GrdF.ColIndex("NoteSerial1")) = True
    If Indx2 = 8 Then
        GrdF.ColHidden(GrdF.ColIndex("NoteSerial1")) = False
        GrdF.ColHidden(GrdF.ColIndex("Id")) = True
        GrdF.ColHidden(GrdF.ColIndex("WorkOrder")) = False
        Label2(9).Visible = True
        TXTOrDer_no.Visible = True
    End If
    DtpDateFrom2.value = mFromDate
    DtpDateTo2.value = mToDate
    DtpDateTo.value = Date
    DtpDateFrom.value = ""
    DtpDateTo.value = ""
    
    txtFromDate.value = mFromDate
    txtToDate.value = mToDate
    
 Dim Dcombos As ClsDataCombos

    Set Dcombos = New ClsDataCombos
    Dim BG As New ClsBackGroundPic
    Dim StrSQL As String

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
If Indx = 15 Then
    VSFlexGrid4.ColHidden(VSFlexGrid4.ColIndex("NoteSerial1")) = True
End If



    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Clear").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture

    CenterForm Me
Dim My_SQL As String

My_SQL = "  select Emp_ID,Emp_name  from TblEmployee  "
My_SQL = My_SQL & "  and  BranchId in(" & Current_branchSql & ") "
    fill_combo DcEmp, My_SQL
 If SystemOptions.UserInterface = ArabicInterface Then
With CboPayMentType
.Clear
.AddItem " Ê“Ì⁄ «·ÿ·» „‰ «·„”∆Ê·"
.AddItem "„” ·„ «·ÿ·»"
.AddItem "—œ „” ·„ «·ÿ·»"
.AddItem " Ê“Ì⁄ ‰ ÌÃ… «·ÿ·» ··„’„„Ì‰"
.AddItem "—œ ﬁ”„ «· ’„Ì„"
End With
Else
With CboPayMentType
.Clear
.AddItem " Ê“Ì⁄ «·ÿ·» „‰ «·„”∆Ê·"
.AddItem "„” ·„ «·ÿ·»"
.AddItem "—œ „” ·„ «·ÿ·»"
.AddItem " Ê“Ì⁄ ‰ ÌÃ… «·ÿ·» ··„’„„Ì‰"
.AddItem "—œ ﬁ”„ «· ’„Ì„"
End With
End If

Dcombos.GetSalesRepData Me.DcEmp1

Dcombos.GetSection Me.DcbDept
Dcombos.GetProcess DcbProcess
Dcombos.GetCustomersSuppliers 1, Me.DcbCustmer
Dcombos.GetCustomersSuppliers 1, Me.DcbCus
Dcombos.GetCustomersSuppliers 1, Me.DcbCusTC
Dcombos.GetCustomersSuppliers 1, DcbCustom
Dcombos.GetCustomersSuppliers 1, DcbCustomer2
Dcombos.GetBranches Me.DcbBranch
Dcombos.GetBranches Me.DcbBranch2
Dcombos.GetBranches Me.dcBranch
Dcombos.GetBranches DataComboBranch
    Dcombos.GetPaymentType cmbPaymentType
Dcombos.GetBranches DataBranch




Dcombos.GetUsers Me.DcbEmployee2

Dcombos.GetCustomersSuppliers 1, Me.DcbCustomer
Dcombos.GetTypesTransport Me.DcbTypeTransport
     If SystemOptions.UserInterface = ArabicInterface Then
           StrSQL = "select ItemID,ItemName from tblitems  where GroupID in ( "
     Else
           StrSQL = "select ItemID,ItemNamee from tblitems  where GroupID in ( "
     End If
                StrSQL = StrSQL & " SELECT     GroupID "
                StrSQL = StrSQL & " From dbo.Groups"
                StrSQL = StrSQL & " Where (HoldingMaterials = 1) )"
   fill_combo DcboItems, StrSQL
    If SystemOptions.UserInterface = ArabicInterface Then
        My_SQL = "  select CusID,CusName from TblCustemers  where type=1  "
    Else
        My_SQL = "  select CusID,CusNamee from TblCustemers  where type=1  "
    End If
   My_SQL = My_SQL & "  and  BranchId in(" & Current_branchSql & ") "
    fill_combo DcAccount2, My_SQL



    If SystemOptions.UserInterface = ArabicInterface Then
        My_SQL = "  select CusID,CusName from TblCustemers  where type=1  "
    Else
        My_SQL = "  select CusID,CusNamee from TblCustemers  where type=1  "
    End If
    My_SQL = My_SQL & "  and  BranchId in(" & Current_branchSql & ") "
    fill_combo DcAccount4, My_SQL

   
    
If SystemOptions.UserInterface = ArabicInterface Then
My_SQL = "  select id,name from contract_type  "
Else
My_SQL = "  select id,namee from contract_type  "
End If
   ' fill_combo DataCombo5, My_SQL
    SetDtpickerDate Me.FrmDTStartDate
    SetDtpickerDate Me.TODTStartDate
    SetDtpickerDate Me.FrmDTEnddate
    SetDtpickerDate Me.ToDTEnddate
    FrmDTStartDate.value = ""
    TODTStartDate.value = ""
    FrmDTEnddate.value = ""
    ToDTEnddate.value = ""
    DTP_TD_From.value = ""
    DTP_TD_To.value = ""
    DTP_Me_From.value = ""
    DTP_Me_To.value = ""
    DTP_T_From.value = ""
    DTP_T_To.value = ""
    DTP_BDialy_From.value = ""
    DTP_BDialy_To.value = ""
    FormPostion Me, GetPostion
    FG.WallPaper = BG.SearchWallpaper
    Set rs = New ADODB.Recordset
    Exit Sub
ErrTrap:

End Sub

Private Sub Retrive(Optional Index As Integer)
    Dim Num As Integer
    On Error GoTo ErrTrap
    FG.Clear flexClearScrollable, flexClearEverything

    If Not (rs.EOF Or rs.BOF) Then
        FG.rows = rs.RecordCount + 1

        For Num = 1 To rs.RecordCount

            With FG
                .TextMatrix(Num, .ColIndex("Count")) = Num
                .TextMatrix(Num, .ColIndex("id")) = IIf(IsNull(rs("projectsID").value), "", Trim(rs("projectsID").value))
                .TextMatrix(Num, .ColIndex("Fullcode")) = IIf(IsNull(rs("projectsFullcode").value), "", (rs("projectsFullcode").value))
                .TextMatrix(Num, .ColIndex("value")) = IIf(IsNull(rs("project_cost").value), "", Trim(rs("project_cost").value))
                .TextMatrix(Num, .ColIndex("expanses_account")) = IIf(IsNull(rs("expanses_account").value), "", Trim(rs("expanses_account").value))
                If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(Num, .ColIndex("Name")) = IIf(IsNull(rs("projects_Project_name").value), "", Trim(rs("projects_Project_name").value))
                .TextMatrix(Num, .ColIndex("Sec_name")) = IIf(IsNull(rs("Sec_name").value), "", Trim(rs("Sec_name").value))
                .TextMatrix(Num, .ColIndex("M_Emp_Name")) = IIf(IsNull(rs("M_Emp_Name").value), "", Trim(rs("M_Emp_Name").value))
                .TextMatrix(Num, .ColIndex("De_Emp_Name")) = IIf(IsNull(rs("De_Emp_Name").value), "", Trim(rs("De_Emp_Name").value))
                .TextMatrix(Num, .ColIndex("Cus_CusName")) = IIf(IsNull(rs("Cus_CusName").value), "", Trim(rs("Cus_CusName").value))
                .TextMatrix(Num, .ColIndex("CusName")) = IIf(IsNull(rs("EndCusName").value), "", Trim(rs("EndCusName").value))
                Else
                .TextMatrix(Num, .ColIndex("Name")) = IIf(IsNull(rs("Project_nameE").value), "", Trim(rs("Project_nameE").value))
                .TextMatrix(Num, .ColIndex("Sec_name")) = IIf(IsNull(rs("Sec_nameE").value), "", Trim(rs("Sec_nameE").value))
                .TextMatrix(Num, .ColIndex("M_Emp_Name")) = IIf(IsNull(rs("M_Emp_Namee").value), "", Trim(rs("M_Emp_Namee").value))
                .TextMatrix(Num, .ColIndex("De_Emp_Name")) = IIf(IsNull(rs("De_Emp_NameE").value), "", Trim(rs("De_Emp_NameE").value))
                .TextMatrix(Num, .ColIndex("Cus_CusName")) = IIf(IsNull(rs("Cus_CusNameE").value), "", Trim(rs("Cus_CusNameE").value))
                .TextMatrix(Num, .ColIndex("CusName")) = IIf(IsNull(rs("EndCusNameE").value), "", Trim(rs("EndCusNameE").value))
                End If
                .TextMatrix(Num, .ColIndex("Material_account")) = IIf(IsNull(rs("Material_account").value), "", Trim(rs("Material_account").value))
                
                 If Not (IsNull(rs("StartDate").value)) Then
                    .TextMatrix(Num, .ColIndex("StartDate")) = Format(rs("StartDate").value, "yyyy/M/d")
                End If
                  If Not (IsNull(rs("EndDate").value)) Then
                    .TextMatrix(Num, .ColIndex("EndDate")) = Format(rs("EndDate").value, "yyyy/M/d")
                End If
               If Index <> 0 Then
               .TextMatrix(Num, .ColIndex("des")) = IIf(IsNull(rs("des").value), "", Trim(rs("des").value))
            End If
       If Index = 2 Then
       If SystemOptions.UserInterface = ArabicInterface Then
               .TextMatrix(Num, .ColIndex("ProcessName")) = IIf(IsNull(rs("ProcessName").value), "", Trim(rs("ProcessName").value))
               Else
               .TextMatrix(Num, .ColIndex("ProcessName")) = IIf(IsNull(rs("ProcessNamee").value), "", Trim(rs("ProcessNamee").value))
               End If
            End If
            End With
            rs.MoveNext
        Next Num

        FG.AutoSize 0, FG.Cols - 1, False
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap

    If rs.State = adStateOpen Then
        rs.Close
        Set rs = Nothing
    End If

    Set cSearchDcbo = Nothing

    FormPostion Me, SavePostion
    Exit Sub
ErrTrap:
End Sub

Private Function Build_Sql(Optional TypeSearch As Integer) As String
    Dim StrSQL As String
    Dim Begin As Boolean
    Dim StrWhere As String
    On Error GoTo ErrTrap
    Select Case TypeSearch
    Case 0, 20
   StrSQL = " SELECT     Material_account, dbo.projects.id AS projectsID, dbo.projects.End_user_Account, dbo.projects.End_user_name, dbo.projects.sub_contractor_Account, "
    StrSQL = StrSQL & "                  dbo.projects.sub_contractor_name, dbo.projects.Fullcode AS projectsFullcode, dbo.projects.Project_name AS projects_Project_name, dbo.projects.CurrencyID,"
    StrSQL = StrSQL & "                  dbo.projects.Project_nameE, dbo.projects.project_cost, dbo.projects.general_discount, dbo.projects.cost_after_discount, dbo.projects.DiscountPercentage,"
    StrSQL = StrSQL & "                  dbo.projects.StartDate, dbo.projects.EndDate, dbo.projects.End_user_id, TblCustemers_1.CusName, TblCustemers_1.CusNamee,"
    StrSQL = StrSQL & "                  TblCustemers_1.Fullcode AS Cus_Fullcode, dbo.projects.sub_contractor_id AS projects_sub_contractor_id, TblCustemers_1.CusName AS Cus_CusName,"
    StrSQL = StrSQL & "                  TblCustemers_1.CusNamee AS Cus_CusNameE, TblCustemers_1.Fullcode AS Cus_Fullcode1, dbo.projects.EmpId AS projects_empid,"
    StrSQL = StrSQL & "                  TblEmployee_1.Emp_Name AS M_Emp_Name, TblEmployee_1.Emp_Namee AS M_Emp_Namee, dbo.projects.EmpId1, TblEmployee_2.CustNum AS De_DcEmp1,"
    StrSQL = StrSQL & "                  dbo.projects.Dept_ID, dbo.TblSection.name AS Sec_name, dbo.TblSection.namee AS Sec_nameE, dbo.projects.Remarkss, dbo.projects.DpNearEndDate,"
    StrSQL = StrSQL & "                  TblCustemers_1.CusID, TblCustemers_2.CusName AS EndCusName, TblCustemers_2.CusNamee AS EndCusNameE, TblCustemers_2.Fullcode AS EndFullcode,"
    StrSQL = StrSQL & "                  dbo.projects.expanses_account, TblEmployee_2.Emp_Name AS De_Emp_Name, TblEmployee_2.Emp_Namee AS De_Emp_NameE"
    StrSQL = StrSQL & "  FROM         dbo.projects LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.TblSection ON dbo.projects.Dept_ID = dbo.TblSection.Id LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.TblEmployee TblEmployee_2 ON dbo.projects.EmpId1 = TblEmployee_2.Emp_ID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.TblEmployee TblEmployee_1 ON dbo.projects.EmpId = TblEmployee_1.Emp_ID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.TblCustemers TblCustemers_1 ON dbo.projects.sub_contractor_id = TblCustemers_1.CusID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.TblCustemers TblCustemers_2 ON dbo.projects.End_user_id = TblCustemers_2.CusID"
    
    Case 1
    StrSQL = " SELECT     TOP 100 PERCENT dbo.projects_des.oprid, dbo.projects_des.des, dbo.projects.id AS projectsID, dbo.projects.End_user_Account, dbo.projects.End_user_name, "
    StrSQL = StrSQL & "                  dbo.projects.sub_contractor_Account, dbo.projects.sub_contractor_name, dbo.projects.Fullcode AS projectsFullcode,"
    StrSQL = StrSQL & "                  dbo.projects.Project_name AS projects_Project_name, dbo.projects.CurrencyID, dbo.projects.Project_nameE, dbo.projects.project_cost, dbo.projects.general_discount,"
    StrSQL = StrSQL & "                  dbo.projects.cost_after_discount, dbo.projects.DiscountPercentage, dbo.projects.StartDate, dbo.projects.EndDate, dbo.projects.End_user_id,"
    StrSQL = StrSQL & "                  TblCustemers_1.CusName, TblCustemers_1.CusNamee, TblCustemers_1.Fullcode AS Cus_Fullcode, dbo.projects.sub_contractor_id AS projects_sub_contractor_id,"
    StrSQL = StrSQL & "                  TblCustemers_1.CusName AS Cus_CusName, TblCustemers_1.CusNamee AS Cus_CusNameE, TblCustemers_1.Fullcode AS Cus_Fullcode1,"
    StrSQL = StrSQL & "                  dbo.projects.EmpId AS projects_empid, TblEmployee_1.Emp_Name AS M_Emp_Name, TblEmployee_1.Emp_Namee AS M_Emp_Namee, dbo.projects.EmpId1,"
    StrSQL = StrSQL & "                  TblEmployee_2.CustNum AS De_DcEmp1, dbo.projects.Dept_ID, dbo.TblSection.name AS Sec_name, dbo.TblSection.namee AS Sec_nameE, dbo.projects.Remarkss,"
    StrSQL = StrSQL & "                  dbo.projects.DpNearEndDate, TblCustemers_1.CusID, TblCustemers_2.CusName AS EndCusName, TblCustemers_2.CusNamee AS EndCusNameE,"
    StrSQL = StrSQL & "                  TblCustemers_2.Fullcode AS EndFullcode, TblEmployee_2.Emp_Name AS De_Emp_Name, TblEmployee_2.Emp_Namee AS De_Emp_NameE,"
    StrSQL = StrSQL & "                  dbo.projects_des.project_id , expanses_account "
    StrSQL = StrSQL & "  FROM         dbo.projects LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.TblSection ON dbo.projects.Dept_ID = dbo.TblSection.Id LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.TblEmployee TblEmployee_2 ON dbo.projects.EmpId1 = TblEmployee_2.Emp_ID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.TblEmployee TblEmployee_1 ON dbo.projects.EmpId = TblEmployee_1.Emp_ID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.TblCustemers TblCustemers_1 ON dbo.projects.sub_contractor_id = TblCustemers_1.CusID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.TblCustemers TblCustemers_2 ON dbo.projects.End_user_id = TblCustemers_2.CusID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.projects_des ON dbo.projects.id = dbo.projects_des.project_id"
    Case 2
   StrSQL = " SELECT     TOP 100 PERCENT dbo.projects_des.oprid, dbo.projects_des.des, dbo.terms_operations.id, dbo.projects.id AS projectsID, dbo.projects.End_user_Account, "
   StrSQL = StrSQL & "                   dbo.projects.End_user_name, dbo.projects.sub_contractor_Account, dbo.projects.sub_contractor_name, dbo.projects.Fullcode AS projectsFullcode,"
   StrSQL = StrSQL & "                   dbo.projects.Project_name AS projects_Project_name, dbo.projects.Project_nameE, dbo.projects.project_cost, dbo.projects.general_discount,"
   StrSQL = StrSQL & "                   dbo.projects.cost_after_discount, dbo.projects.DiscountPercentage, dbo.projects.StartDate, dbo.projects.EndDate, dbo.projects.End_user_id,"
   StrSQL = StrSQL & "                   TblCustemers_1.CusName, TblCustemers_1.CusNamee, TblCustemers_1.Fullcode AS Cus_Fullcode, dbo.projects.sub_contractor_id AS projects_sub_contractor_id,"
   StrSQL = StrSQL & "                   TblCustemers_1.CusName AS Cus_CusName, TblCustemers_1.CusNamee AS Cus_CusNameE, TblCustemers_1.Fullcode AS Cus_Fullcode1,"
   StrSQL = StrSQL & "                   dbo.projects.EmpId AS projects_empid, TblEmployee_1.Emp_Name AS M_Emp_Name, TblEmployee_1.Emp_Namee AS M_Emp_Namee, dbo.projects.EmpId1,"
   StrSQL = StrSQL & "                   TblEmployee_2.CustNum AS De_DcEmp1, TblEmployee_2.Emp_Name AS De_Emp_Name, dbo.projects.Dept_ID, dbo.TblSection.name AS Sec_name,"
   StrSQL = StrSQL & "                   dbo.TblSection.namee AS Sec_nameE, dbo.projects.Remarkss, dbo.projects.DpNearEndDate, TblCustemers_1.CusID, TblCustemers_2.CusName AS EndCusName,"
   StrSQL = StrSQL & "                   TblCustemers_2.CusNamee AS EndCusNameE, TblCustemers_2.Fullcode AS EndFullcode, dbo.TblProcessDEF.ProcessName, dbo.TblProcessDEF.ProcessNameE,"
   StrSQL = StrSQL & "                   dbo.terms_operations.OPRIDD , expanses_account "
   StrSQL = StrSQL & " FROM         dbo.terms_operations LEFT OUTER JOIN"
   StrSQL = StrSQL & "                    dbo.TblProcessDEF ON dbo.terms_operations.OPRIDD = dbo.TblProcessDEF.TblProcessDEFID RIGHT OUTER JOIN"
   StrSQL = StrSQL & "                    dbo.projects LEFT OUTER JOIN"
   StrSQL = StrSQL & "                    dbo.TblSection ON dbo.projects.Dept_ID = dbo.TblSection.Id LEFT OUTER JOIN"
   StrSQL = StrSQL & "                    dbo.TblEmployee TblEmployee_2 ON dbo.projects.EmpId1 = TblEmployee_2.Emp_ID LEFT OUTER JOIN"
   StrSQL = StrSQL & "                    dbo.TblEmployee TblEmployee_1 ON dbo.projects.EmpId = TblEmployee_1.Emp_ID LEFT OUTER JOIN"
   StrSQL = StrSQL & "                    dbo.TblCustemers TblCustemers_1 ON dbo.projects.sub_contractor_id = TblCustemers_1.CusID LEFT OUTER JOIN"
   StrSQL = StrSQL & "                    dbo.TblCustemers TblCustemers_2 ON dbo.projects.End_user_id = TblCustemers_2.CusID LEFT OUTER JOIN"
   StrSQL = StrSQL & "                   dbo.projects_des ON dbo.projects.id = dbo.projects_des.project_id ON dbo.terms_operations.ProjectDes_ID = dbo.projects_des.oprid"

    End Select
    StrSQL = StrSQL & " where 1=1 AND (NOT (dbo.projects.Fullcode IS NULL))"
    Begin = True
  
   StrSQL = StrSQL & " and  dbo.projects.branch_no in(" & Current_branchSql & ")"

    If XPTxtComID.text <> "" Then
              StrWhere = StrWhere + " and dbo.projects.fullcode like'%" & (XPTxtComID.text) & "%'"
      End If
    If TxtCompanyName.text <> "" Then
                 StrWhere = StrWhere + " and ( dbo.projects.Project_nameE LIKE '%" & Trim(TxtCompanyName.text) & "%' or dbo.projects.Project_name LIKE '%" & Trim(TxtCompanyName.text) & "%')"
    End If
     If TxtProjectCosts.text <> "" Then
              StrWhere = StrWhere + " and dbo.projects.project_cost =" & (TxtProjectCosts.text) & ""
      End If
   If DcAccount2.text <> "" And val(DcAccount2.BoundText) <> 0 Then
              StrWhere = StrWhere + " and dbo.projects.End_user_id =" & (DcAccount2.BoundText) & ""
      End If
     If DcAccount4.text <> "" And val(DcAccount4.BoundText) <> 0 Then
              StrWhere = StrWhere + " and dbo.projects.sub_contractor_id =" & (DcAccount4.BoundText) & ""
      End If
      If DcEmp1.text <> "" And val(DcEmp1.BoundText) <> 0 Then
              StrWhere = StrWhere + " and dbo.projects.EmpId1 =" & (DcEmp1.BoundText) & ""
      End If
       If DcEmp.text <> "" And val(DcEmp.BoundText) <> 0 Then
              StrWhere = StrWhere + " and dbo.projects.EmpId =" & (DcEmp.BoundText) & ""
      End If
       If DcbDept.text <> "" And val(DcbDept.BoundText) <> 0 Then
              StrWhere = StrWhere + " and dbo.projects.Dept_ID =" & (DcbDept.BoundText) & ""
      End If
      If Not IsNull(Me.FrmDTStartDate.value) Then
             StrWhere = StrWhere & " AND dbo.projects.StartDate>=" & SQLDate(Me.FrmDTStartDate.value, True) & ""
      End If
       If Not IsNull(Me.TODTStartDate.value) Then
             StrWhere = StrWhere & " AND dbo.projects.StartDate <=" & SQLDate(Me.TODTStartDate.value, True) & ""
      End If
       If Not IsNull(Me.FrmDTEnddate.value) Then
             StrWhere = StrWhere & " AND dbo.projects.EndDate>=" & SQLDate(Me.FrmDTEnddate.value, True) & ""
      End If
       If Not IsNull(Me.ToDTEnddate.value) Then
             StrWhere = StrWhere & " AND dbo.projects.EndDate>=" & SQLDate(Me.ToDTEnddate.value, True) & ""
      End If
      If TypeSearch <> 0 Then
        If TxtBand.text <> "" Then
              StrWhere = StrWhere + " and dbo.projects_des.des like'%" & (TxtBand.text) & "%'"
      End If
      End If
      If TypeSearch = 2 Then
         If DcbProcess.text <> "" And val(DcbProcess.BoundText) <> 0 Then
              StrWhere = StrWhere + " and dbo.terms_operations.OPRIDD =" & (DcbProcess.BoundText) & ""
      End If
      End If
      StrSQL = StrSQL & "  AND      branch_no in(" & Current_branchSql & ")"
    Build_Sql = StrSQL + StrWhere + " order by dbo.projects.fullcode  "
    Exit Function
ErrTrap:
End Function

Private Sub ChangeLang()
Me.Caption = "Project Search"
 Label2(0).Caption = "Name"
 Label1(0).Caption = "Code"
 Label20.Caption = "Value"
  Label3.Caption = "Term"
  Label4.Caption = "Opr"
  Fra(0).Caption = "Start Date"
  lbl(5).Caption = "From"
  lbl(0).Caption = "To"
  
  Fra(3).Caption = "End Date"
  lbl(6).Caption = "From"
  lbl(7).Caption = "To"
  Label16.Caption = "End Cust."
  Label23.Caption = "Sub-Cust."
  Label41.Caption = "Sales Man"
  Label35.Caption = "Manger"
  Label43.Caption = "Dept."
  Cmd(0).Caption = "Project Search"
  Cmd(3).Caption = "Term Search"
  Cmd(4).Caption = "OPR Search"
  Cmd(1).Caption = "Clear"
  Cmd(2).Caption = "Exit"
  
      With Me.FG
      
      .TextMatrix(0, .ColIndex("count")) = "I"
        .TextMatrix(0, .ColIndex("fullcode")) = "Code"
        .TextMatrix(0, .ColIndex("Name")) = "Name"
        .TextMatrix(0, .ColIndex("value")) = "Value"
        .TextMatrix(0, .ColIndex("startDate")) = "Start Date"
        .TextMatrix(0, .ColIndex("EndDate")) = "End Date"
        .TextMatrix(0, .ColIndex("Sec_name")) = "Dept."
        
        .TextMatrix(0, .ColIndex("M_Emp_Name")) = "Manger"
        .TextMatrix(0, .ColIndex("CusName")) = "End Cus."
        .TextMatrix(0, .ColIndex("Cus_CusName")) = "Sub Cus.."
        .TextMatrix(0, .ColIndex("De_Emp_Name")) = "Sales pers"
        
        .TextMatrix(0, .ColIndex("des")) = "Terms."
        .TextMatrix(0, .ColIndex("ProcessName")) = "Process"
        
     
 
    End With

With Me.GrdF
      
      .TextMatrix(0, .ColIndex("count")) = "I"
        .TextMatrix(0, .ColIndex("Id")) = "Order No"
        .TextMatrix(0, .ColIndex("NoteSerial1")) = "Note Serial"
        .TextMatrix(0, .ColIndex("branch_name")) = "branch"
        .TextMatrix(0, .ColIndex("CusName")) = "Customer Name"
        .TextMatrix(0, .ColIndex("RecordDate")) = "Record Date"

     
 
    End With

  Label2(6).Caption = "Order No"
  lbl(30).Caption = "branch"
  Label1(8).Caption = "Customer"

  Fra(14).Caption = "Date"
  lbl(27).Caption = "From"
  lbl(29).Caption = "To"
  
 ISButton16.Caption = "Search"
 ISButton17.Caption = "Clear"
  ISButton18.Caption = "Exit"
  Me.Caption = "Search"
End Sub

 
Private Sub GrdF_Click()
If Indx2 = 8 Then
    FrmVizitScreen.FindRec val(GrdF.TextMatrix(GrdF.Row, 1))
Else
 emp_CONTRACT_TYPE.FindRec val(GrdF.TextMatrix(GrdF.Row, 1)), Indx2
End If
 Unload Me
End Sub

Private Sub ISButton1_Click()

GetDataTranOder

End Sub

Private Sub GetDataTranOder()


    Dim sql As String
    Dim StrSQL As String
    Dim Begin As Boolean
  '  Public Current_branch As Integer
   ' Public Current_branchSql As String

    Dim StrWhere As String
    StrWhere = ""
   ' On Error GoTo ErrTrap
StrSQL = ""
  StrSQL = StrSQL & " SELECT dbo.Tbl_TransOrder.ID, dbo.Tbl_TransOrder.TOrder_OrderNum, dbo.Tbl_TransOrder.TOrder_Status, dbo.Tbl_TransOrder.TOrder_BranchID,"
  StrSQL = StrSQL & " dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.Tbl_TransOrder.TOrder_DateOrder, dbo.Tbl_TransOrder.TOrder_EmpID,"
  StrSQL = StrSQL & "  dbo.TblUsers.UserID Emp_Code, dbo.TblUsers.UserName Emp_Name, dbo.TblUsers.UserName Emp_Namee, dbo.Tbl_TransOrder.TOrder_Notes,"
  StrSQL = StrSQL & "  dbo.Tbl_TransOrder.TOrder_DateNote , dbo.Tbl_TransOrder.TOrder_Time, dbo.Tbl_TransOrder.UserID"

   StrSQL = StrSQL & " FROM  dbo.Tbl_TransOrder INNER JOIN"
    StrSQL = StrSQL & "  dbo.TblBranchesData ON dbo.Tbl_TransOrder.TOrder_BranchID = dbo.TblBranchesData.branch_id INNER JOIN"
    StrSQL = StrSQL & "  dbo.TblUsers ON dbo.Tbl_TransOrder.TOrder_EmpID = dbo.TblUsers.Userid"

    StrSQL = StrSQL & " where 1=1 AND (NOT (dbo.Tbl_TransOrder.ID IS NULL))"
       If SystemOptions.usertype = UserNormal Then
    StrSQL = StrSQL & " and Tbl_TransOrder.UserID=" & user_id & " or TOrder_EmpID=" & user_id & ""
 
    
    End If
    
    Begin = True
  
   'Sql = StrSQL & " and  dbo.TBL_measureMent.ID in(" & Current_branchSql & ")"

    If TXT_TOrder.text <> "" Then
              StrWhere = StrWhere + " and dbo.Tbl_TransOrder.ID like'%" & (TXT_TOrder.text) & "%'"
      End If
    If Txt_MID.text <> "" Then
                 StrWhere = StrWhere + " and ( dbo.Tbl_TransOrder.TOrder_OrderNum LIKE '%" & Trim(Txt_MID.text) & "%')"
    End If
     If TxtTime.text <> "" Then
              StrWhere = StrWhere + " and dbo.Tbl_TransOrder.TOrder_Time ='" & (TxtTime.text) & "'"
      End If
   If dcBranch.BoundText <> "" Then
              StrWhere = StrWhere + " and  dbo.Tbl_TransOrder.TOrder_BranchID ='" & (dcBranch.BoundText) & "'"
      End If
     If CboPayMentType.text <> "" Then
              StrWhere = StrWhere + " and dbo.Tbl_TransOrder.TOrder_Status ='" & (CboPayMentType.text) & "'"
      End If
      If TxtSearchCode.text <> "" Then
              StrWhere = StrWhere + " and dbo.TblUsers.Userid =" & (TxtSearchCode.text) & ""
      End If
      
      If DcbEmployee2.BoundText <> "" Then
              StrWhere = StrWhere + " and dbo.TblUsers.Userid=" & (DcbEmployee2.BoundText) & ""
      End If
      
      
      If Not IsNull(Me.DTP_T_From.value) Then
             StrWhere = StrWhere & " AND dbo.Tbl_TransOrder.TOrder_DateOrder>=" & SQLDate(Me.DTP_T_From.value, True) & ""
      End If
       If Not IsNull(Me.DTP_T_To.value) Then
             StrWhere = StrWhere & " AND dbo.Tbl_TransOrder.TOrder_DateOrder <=" & SQLDate(Me.DTP_T_To.value, True) & ""
      End If
      
    
      sql = StrSQL + StrWhere + " order by dbo.Tbl_TransOrder.ID"
   
   ''----------------------------------------------------------------------
  
    Dim Num As Integer
    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    VSFlexGrid2.Clear flexClearScrollable, flexClearEverything
    If Not (rs.EOF Or rs.BOF) Then
        VSFlexGrid2.rows = rs.RecordCount + 1
        
        For Num = 1 To rs.RecordCount
       
            With VSFlexGrid2
                .TextMatrix(Num, .ColIndex("Count")) = Num
                .TextMatrix(Num, .ColIndex("id")) = IIf(IsNull(rs("id").value), "", Trim(rs("id").value))
                .TextMatrix(Num, .ColIndex("IDMeasure")) = IIf(IsNull(rs("TOrder_OrderNum").value), "", (rs("TOrder_OrderNum").value))
                .TextMatrix(Num, .ColIndex("Branch")) = IIf(IsNull(rs("branch_name").value), "", Trim(rs("branch_name").value))
                .TextMatrix(Num, .ColIndex("Status")) = IIf(IsNull(rs("TOrder_Status").value), "", (rs("TOrder_Status").value))
                .TextMatrix(Num, .ColIndex("EmpName")) = IIf(IsNull(rs("Emp_Name").value), "", (rs("Emp_Name").value))
                .TextMatrix(Num, .ColIndex("TTime")) = IIf(IsNull(rs("TOrder_Time").value), "", (rs("TOrder_Time").value))
                .TextMatrix(Num, .ColIndex("DateTOrder")) = IIf(IsNull(rs("TOrder_DateOrder").value), "", (rs("TOrder_DateOrder").value))
       
            End With
            rs.MoveNext
        Next Num
      
End If
   
   ''----------------------------------------------------------------------

ErrTrap:

End Sub

'Private Sub ISButton10_Click()
'
'End Sub

Private Sub ISButton11_Click()
clear_all Me
DtpDateFrom.value = ""
DtpDateTo.value = ""
RdAuto_Manual(2).value = True
End Sub

Private Sub ISButton12_Click()
Unload Me
End Sub

Private Sub ISButton13_Click()
GetBillCustomer
End Sub

Private Sub ISButton15_Click()
GetServiceInvoice
End Sub

Private Sub ISButton16_Click()



    Dim sql As String
    Dim StrSQL As String
    Dim Begin As Boolean
  '  Public Current_branch As Integer
   ' Public Current_branchSql As String

    Dim StrWhere As String
   ' On Error GoTo ErrTrap
    Dim mTableName As String
    If Indx = 1 Then
        mTableName = "ContainerContracts"
    ElseIf Indx = 2 Then
        mTableName = "ContainerContractsRec"
    ElseIf Indx = 4 Then
        mTableName = "ContainerUnloading"
    End If
        If Indx2 = 8 Then
            mTableName = "TblHandWages"
            GetDataHandWages
        Exit Sub
        End If
    
    
            StrSQL = " Select TT.ID,"
    StrSQL = StrSQL & "             TT.CustID,TT.RecordDate,TT.CustTel,TT.BranchID,Tc.CusName,B.branch_name"
   StrSQL = StrSQL & " from " & mTableName & "  TT"
   StrSQL = StrSQL & " LEFT Outer JOIN TblBranchesData AS b ON TT.BranchID = b.branch_id"
   StrSQL = StrSQL & " LEFT Outer JOIN TblCustemers AS tc ON TT.CustID = tc.CusID"
   StrSQL = StrSQL & " where 1=1 AND (NOT (TT.ID IS NULL))"
   
    Begin = True
  

    If txtid.text <> "" Then
              StrWhere = StrWhere + " and TT.ID like '%" & (txtid.text) & "%'"
      End If
    If DcbCus.text <> "" Then
                 StrWhere = StrWhere + " and ( TT.CustID LIKE '%" & Trim(DcbCus.BoundText) & "%')"
    End If
    If DataBranch.BoundText <> "" Then
                 StrWhere = StrWhere + " and ( TT.BranchID LIKE '%" & Trim(DataBranch.BoundText) & "%')"
    End If
    
      
      
      If Not IsNull(Me.txtFromDate.value) Then
             StrWhere = StrWhere & " AND TT.RecordDate>=" & SQLDate(Me.txtFromDate.value, True) & ""
      End If
       If Not IsNull(Me.txtToDate.value) Then
             StrWhere = StrWhere & " AND TT.RecordDate <=" & SQLDate(Me.txtToDate.value, True) & ""
      End If
      
    
      sql = StrSQL + StrWhere + " order by TT.ID"
   
   ''----------------------------------------------------------------------
  
    Dim Num As Integer
    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    GrdF.Clear flexClearScrollable, flexClearEverything
    If Not (rs.EOF Or rs.BOF) Then
        GrdF.rows = rs.RecordCount + 1
        
        For Num = 1 To rs.RecordCount
       
            With GrdF
                .TextMatrix(Num, .ColIndex("Count")) = Num
                .TextMatrix(Num, .ColIndex("id")) = IIf(IsNull(rs("id").value), "", Trim(rs("id").value))
                .TextMatrix(Num, .ColIndex("branch_name")) = IIf(IsNull(rs("branch_name").value), "", (rs("branch_name").value))
                .TextMatrix(Num, .ColIndex("CusName")) = IIf(IsNull(rs("CusName").value), "", Trim(rs("CusName").value))
                .TextMatrix(Num, .ColIndex("CusName")) = IIf(IsNull(rs("CusName").value), "", (rs("CusName").value))
                .TextMatrix(Num, .ColIndex("RecordDate")) = IIf(IsNull(rs("RecordDate").value), "", (rs("RecordDate").value))
                
          
            End With
            rs.MoveNext
        Next Num
      
End If
   
   ''----------------------------------------------------------------------

ErrTrap:


End Sub


Private Sub GetDataHandWages()

    Dim sql As String
    Dim StrSQL As String
    Dim Begin As Boolean
  '  Public Current_branch As Integer
   ' Public Current_branchSql As String

    Dim StrWhere As String
   ' On Error GoTo ErrTrap
    Dim mTableName As String
    
    mTableName = "TblHandWages"

    

StrSQL = " SELECT TT.ID,TT.PaymentId,TblPaymentType.PaymentName, "
StrSQL = StrSQL & "        TT.NoteSerial1,TblCardAuthorizationReform.WorkOrder,"
StrSQL = StrSQL & "        TblCardAuthorizationReform.CusID,"
StrSQL = StrSQL & "        TT.RecordDate,"
StrSQL = StrSQL & "        TT.BranchID,"
StrSQL = StrSQL & "        TblCardAuthorizationReform.ClientName,"
If SystemOptions.UserInterface = EnglishInterface Then
    StrSQL = StrSQL & "       ISNULL(b.branch_namee,b.branch_name) branch_name"
Else
    StrSQL = StrSQL & "       b.branch_name"
End If
StrSQL = StrSQL & " FROM   TblHandWages TT"
StrSQL = StrSQL & "        LEFT OUTER JOIN TblBranchesData AS b"
StrSQL = StrSQL & "             ON  TT.BranchID = b.branch_id"
StrSQL = StrSQL & "        LEFT OUTER JOIN TblCardAuthorizationReform"
StrSQL = StrSQL & "             ON  TblCardAuthorizationReform.WorkOrder = tt.OrDer_no2"
StrSQL = StrSQL & "             Left outer  join TblPaymentType"
StrSQL = StrSQL & "             On TblPaymentType.PaymentId = TT.PaymentId"
StrSQL = StrSQL & " where 1=1 AND (NOT (TT.ID IS NULL))"
   
    Begin = True
  

    If txtid.text <> "" Then
              StrWhere = StrWhere + " and TT.NoteSerial1 like '%" & (txtid.text) & "%'"
      End If
      
    If TXTOrDer_no.text <> "" Then
              StrWhere = StrWhere + " and TblCardAuthorizationReform.WorkOrder = " & (TXTOrDer_no.text)
      End If
      
      
    If DcbCus.text <> "" Then
                 StrWhere = StrWhere + " and ( TblCardAuthorizationReform.ClientName LIKE '%" & Trim(DcbCus.text) & "%')"
    End If
    If DataBranch.BoundText <> "" Then
                 StrWhere = StrWhere + " and ( TT.BranchID LIKE '%" & Trim(DataBranch.BoundText) & "%')"
    End If
    
      
       If cmbPaymentType.BoundText <> "" Then
                 StrWhere = StrWhere + " and ( TT.PaymentId = " & val(cmbPaymentType.BoundText) & ")"
        End If
    
      
      If Not IsNull(Me.txtFromDate.value) Then
             StrWhere = StrWhere & " AND TT.RecordDate>=" & SQLDate(Me.txtFromDate.value, True) & ""
      End If
       If Not IsNull(Me.txtToDate.value) Then
             StrWhere = StrWhere & " AND TT.RecordDate <=" & SQLDate(Me.txtToDate.value, True) & ""
      End If
      
    
      sql = StrSQL + StrWhere + " order by TT.ID"
   
   ''----------------------------------------------------------------------
  
    Dim Num As Integer
    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    GrdF.Clear flexClearScrollable, flexClearEverything
    If Not (rs.EOF Or rs.BOF) Then
        GrdF.rows = rs.RecordCount + 1
        
        For Num = 1 To rs.RecordCount
       
            With GrdF
                .TextMatrix(Num, .ColIndex("Count")) = Num
                .TextMatrix(Num, .ColIndex("id")) = IIf(IsNull(rs("id").value), "", Trim(rs("id").value))
                .TextMatrix(Num, .ColIndex("NoteSerial1")) = IIf(IsNull(rs("NoteSerial1").value), "", Trim(rs("NoteSerial1").value))
                .TextMatrix(Num, .ColIndex("branch_name")) = IIf(IsNull(rs("branch_name").value), "", (rs("branch_name").value))
                .TextMatrix(Num, .ColIndex("CusName")) = IIf(IsNull(rs("ClientName").value), "", Trim(rs("ClientName").value))
                .TextMatrix(Num, .ColIndex("WorkOrder")) = IIf(IsNull(rs("WorkOrder").value), "", Trim(rs("WorkOrder").value))
                .TextMatrix(Num, .ColIndex("PaymentName")) = IIf(IsNull(rs("PaymentName").value), "", Trim(rs("PaymentName").value))
                
                .TextMatrix(Num, .ColIndex("RecordDate")) = IIf(IsNull(rs("RecordDate").value), "", (rs("RecordDate").value))
                
          
            End With
            rs.MoveNext
        Next Num
      
End If
   
   ''----------------------------------------------------------------------

ErrTrap:


End Sub

Private Sub ISButton17_Click()
 clear_all Me
txtFromDate.value = ""
txtToDate.value = ""

End Sub

Private Sub ISButton18_Click()
Unload Me
End Sub

Private Sub ISButton2_Click()
 clear_all Me
 
DTP_T_From.value = ""
DTP_T_To.value = ""


End Sub

Private Sub ISButton3_Click()
Unload Me
End Sub

Private Sub ISButton4_Click()
GetDataBusinessDialy
End Sub
Private Sub GetDataBusinessDialy()


    Dim sql As String
    Dim StrSQL As String
    Dim Begin As Boolean
  '  Public Current_branch As Integer
   ' Public Current_branchSql As String

    Dim StrWhere As String
   ' On Error GoTo ErrTrap

StrSQL = StrSQL & " SELECT dbo.Tbl_BusinessDialy.ID, dbo.Tbl_BusinessDialy.BD_Date,dbo.Tbl_BusinessDialy.BD_BranchID, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee,"
StrSQL = StrSQL & "           dbo.Tbl_BusinessDialy.BD_Notes, dbo.Tbl_BusinessDialy.TradingContractID, dbo.Tbl_TradingContract.TContract_CustID, dbo.TblCustemers.CusName,"
StrSQL = StrSQL & "           dbo.TblCustemers.CusNamee"
StrSQL = StrSQL & " FROM  dbo.Tbl_BusinessDialy INNER JOIN"
StrSQL = StrSQL & "              dbo.Tbl_TradingContract ON dbo.Tbl_BusinessDialy.TradingContractID = dbo.Tbl_TradingContract.ID INNER JOIN"
StrSQL = StrSQL & "              dbo.TblCustemers ON dbo.Tbl_TradingContract.TContract_CustID = dbo.TblCustemers.CusID INNER JOIN"
StrSQL = StrSQL & "             dbo.TblBranchesData ON dbo.Tbl_BusinessDialy.BD_BranchID = dbo.TblBranchesData.branch_id"
    
    
    StrSQL = StrSQL & " where 1=1 AND (NOT (dbo.Tbl_BusinessDialy.ID IS NULL))"
    Begin = True
  
    StrSQL = StrSQL & " And IsNull(IsCanceld,0) <>  1 "
    If Txt_BusinessID.text <> "" Then
              StrWhere = StrWhere + " and dbo.Tbl_BusinessDialy.ID like '%" & (Txt_BusinessID.text) & "%'"
      End If
    If Txt_NumberContract.text <> "" Then
              StrWhere = StrWhere + " and dbo.Tbl_BusinessDialy.TradingContractID ='" & (Txt_NumberContract.text) & "'"
    End If
      
    If DataComboBranch.BoundText <> "" Then
                 StrWhere = StrWhere + " and ( dbo.Tbl_BusinessDialy.BD_BranchID LIKE '%" & Trim(DataComboBranch.BoundText) & "%')"
    End If
     

      If TxCustmer.text <> "" Then
              StrWhere = StrWhere + " and dbo.Tbl_TradingContract.TContract_CustID =" & (TxCustmer.text) & ""
      End If
      
     If DcbCustom.BoundText <> "" Then
                 StrWhere = StrWhere + " and ( dbo.Tbl_TradingContract.TContract_CustID LIKE '%" & Trim(DcbCustom.BoundText) & "%')"
    End If
      
      If Not IsNull(Me.DTP_BDialy_From.value) Then
             StrWhere = StrWhere & " AND dbo.Tbl_BusinessDialy.BD_Date>=" & SQLDate(Me.DTP_BDialy_From.value, True) & ""
      End If
       If Not IsNull(Me.DTP_BDialy_To.value) Then
             StrWhere = StrWhere & " AND dbo.Tbl_BusinessDialy.BD_Date <=" & SQLDate(Me.DTP_BDialy_To.value, True) & ""
      End If
      
    
      sql = StrSQL + StrWhere + " order by dbo.Tbl_BusinessDialy.ID"
   
   ''----------------------------------------------------------------------
  
    Dim Num As Integer
    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    VSFlexGrid3.Clear flexClearScrollable, flexClearEverything
    If Not (rs.EOF Or rs.BOF) Then
        VSFlexGrid3.rows = rs.RecordCount + 1
        
        For Num = 1 To rs.RecordCount
       
            With VSFlexGrid3
                .TextMatrix(Num, .ColIndex("Count")) = Num
                .TextMatrix(Num, .ColIndex("id")) = IIf(IsNull(rs("id").value), "", Trim(rs("id").value))
                .TextMatrix(Num, .ColIndex("IDBusiness")) = IIf(IsNull(rs("TradingContractID").value), "", (rs("TradingContractID").value))
                .TextMatrix(Num, .ColIndex("Branch")) = IIf(IsNull(rs("branch_name").value), "", Trim(rs("branch_name").value))
                .TextMatrix(Num, .ColIndex("BCustomer")) = IIf(IsNull(rs("CusName").value), "", (rs("CusName").value))
                .TextMatrix(Num, .ColIndex("DateBussiness")) = IIf(IsNull(rs("BD_Date").value), "", (rs("BD_Date").value))
       
            End With
            rs.MoveNext
        Next Num
      
End If
   
   ''----------------------------------------------------------------------

ErrTrap:


End Sub
Private Sub ISButton5_Click()
 clear_all Me
DTP_Me_From.value = ""
DTP_Me_To.value = ""


End Sub

Private Sub ISButton6_Click()
Unload Me
End Sub

Private Sub ISButton7_Click()
If Indx = 15 Then
    GetTradingInvContract
Else
    GetTradingContract
End If

End Sub


Private Sub GetBillCustomer()
    Dim sql As String
    Dim StrSQL As String
    Dim Begin As Boolean
    Dim StrWhere As String
   ' On Error GoTo ErrTrap


    StrSQL = " SELECT        dbo.TblTravDueK.ID,TblTravDueK.ContainerNo, dbo.TblTravDueK.recordDate, dbo.TblTravDueK.recordDateH, dbo.TblTravDueK.BranchId, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblTravDueK.CusID, "
    StrSQL = StrSQL & "                      dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode, dbo.TblTravDueK.ItemID, dbo.TblItems.ItemCode, dbo.TblItems.ItemName, dbo.TblItems.ItemNamee,"
    StrSQL = StrSQL & "                     dbo.TblTravDueK.TypeTransportID , dbo.TblTypesTransport.Name, dbo.TblTypesTransport.NameE, dbo.TblTravDueK.NoteSerial1, dbo.TblTravDueK.RdAuto_Manual"
    StrSQL = StrSQL & "     FROM            dbo.TblTravDueK LEFT OUTER JOIN"
    StrSQL = StrSQL & "                     dbo.TblTypesTransport ON dbo.TblTravDueK.TypeTransportID = dbo.TblTypesTransport.ID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                     dbo.TblItems ON dbo.TblTravDueK.ItemID = dbo.TblItems.ItemID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                     dbo.TblCustemers ON dbo.TblTravDueK.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                     dbo.TblBranchesData ON dbo.TblTravDueK.BranchId = dbo.TblBranchesData.branch_id"
     StrSQL = StrSQL & " where 1=1"
    If val(TxtIDFrom.text) <> 0 Then
              StrWhere = StrWhere + " and dbo.TblTravDueK.NoteSerial1 >= " & val(TxtIDFrom.text) & ""
      End If
      
      If Trim(txtContainerNo.text) <> "" Then
              StrWhere = StrWhere + " and dbo.TblTravDueK.ContainerNo = '" & val(txtContainerNo.text) & "'"
      End If
      
      If val(TxtIDTO.text) <> 0 Then
              StrWhere = StrWhere + " and dbo.TblTravDueK.NoteSerial1 <=" & val(TxtIDTO.text) & ""
      End If
        If val(DcbBranch.BoundText) <> 0 And DcbBranch.text <> "" Then
              StrWhere = StrWhere + " and dbo.TblTravDueK.BranchId  = " & val(DcbBranch.BoundText) & ""
      End If
      
    If val(DcbTypeTransport.BoundText) <> 0 And DcbTypeTransport.text <> "" Then
              StrWhere = StrWhere + " and dbo.TblTravDueK.TypeTransportID  = " & val(DcbTypeTransport.BoundText) & ""
      End If
       If val(DcbCustomer.BoundText) <> 0 And DcbCustomer.text <> "" Then
              StrWhere = StrWhere + " and dbo.TblTravDueK.CusID  = " & val(DcbCustomer.BoundText) & ""
      End If
            If val(DcboItems.BoundText) <> 0 And DcboItems.text <> "" Then
              StrWhere = StrWhere + " and dbo.TblTravDueK.ItemID  = " & val(DcboItems.BoundText) & ""
      End If
      
      If Not IsNull(Me.DtpDateFrom.value) Then
             StrWhere = StrWhere & " AND dbo.TblTravDueK.recordDate>=" & SQLDate(Me.DtpDateFrom.value, True) & ""
      End If
       If Not IsNull(Me.DtpDateTo.value) Then
             StrWhere = StrWhere & " AND dbo.TblTravDueK.recordDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
      End If
      If RdAuto_Manual(0).value = True Then
      StrWhere = StrWhere & " AND dbo.TblTravDueK.RdAuto_Manual=0"
      End If
     If RdAuto_Manual(1).value = True Then
      StrWhere = StrWhere & " AND dbo.TblTravDueK.RdAuto_Manual=1"
      End If
    
      sql = StrSQL + StrWhere + " order by dbo.TblTravDueK.NoteSerial1"
   
   ''----------------------------------------------------------------------
  Dim Auto_Man As Integer
    Dim Num As Integer
    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    VSFlexGrid5.Clear flexClearScrollable, flexClearEverything
     VSFlexGrid5.rows = 2
    If rs.RecordCount > 0 Then
        VSFlexGrid5.rows = rs.RecordCount + 1
        
        For Num = 1 To rs.RecordCount
       
            With VSFlexGrid5
                .TextMatrix(Num, .ColIndex("Count")) = Num
                .TextMatrix(Num, .ColIndex("ID")) = IIf(IsNull(rs("ID").value), 0, Trim(rs("ID").value))
                .TextMatrix(Num, .ColIndex("NoteSerial1")) = IIf(IsNull(rs("NoteSerial1").value), "", Trim(rs("NoteSerial1").value))
                .TextMatrix(Num, .ColIndex("recordDate")) = IIf(IsNull(rs("recordDate").value), "", (rs("recordDate").value))
                 Auto_Man = IIf(IsNull(rs("RdAuto_Manual").value), 0, (rs("RdAuto_Manual").value))
                If SystemOptions.UserInterface = ArabicInterface Then
                If Auto_Man = 0 Then
                .TextMatrix(Num, .ColIndex("RdAuto_Manual")) = "ÌœÊÌ"
                Else
                .TextMatrix(Num, .ColIndex("RdAuto_Manual")) = "«·Ì"
                End If
                Else
                 If Auto_Man = 0 Then
                .TextMatrix(Num, .ColIndex("RdAuto_Manual")) = "Manual"
                Else
                .TextMatrix(Num, .ColIndex("RdAuto_Manual")) = "Auto"
                End If
                End If
                If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(Num, .ColIndex("Name")) = IIf(IsNull(rs("Name").value), "", (rs("Name").value))
                .TextMatrix(Num, .ColIndex("CusName")) = IIf(IsNull(rs("CusName").value), "", Trim(rs("CusName").value))
                .TextMatrix(Num, .ColIndex("branch_name")) = IIf(IsNull(rs("branch_name").value), "", (rs("branch_name").value))
                .TextMatrix(Num, .ColIndex("ItemName")) = IIf(IsNull(rs("ItemName").value), "", rs("ItemName").value)
                Else
                .TextMatrix(Num, .ColIndex("Name")) = IIf(IsNull(rs("NameE").value), "", (rs("NameE").value))
                .TextMatrix(Num, .ColIndex("ItemName")) = IIf(IsNull(rs("ItemNamee").value), "", rs("ItemNamee").value)
                .TextMatrix(Num, .ColIndex("CusName")) = IIf(IsNull(rs("CusNamee").value), "", Trim(rs("CusNamee").value))
                .TextMatrix(Num, .ColIndex("branch_name")) = IIf(IsNull(rs("branch_namee").value), "", (rs("branch_namee").value))
                End If
          
            End With
            rs.MoveNext
        Next Num
      
End If
   
   ''----------------------------------------------------------------------

ErrTrap:

End Sub
Private Sub GetTradingContract()


    Dim sql As String
    Dim StrSQL As String
    Dim Begin As Boolean
  '  Public Current_branch As Integer
   ' Public Current_branchSql As String

    Dim StrWhere As String
   ' On Error GoTo ErrTrap


    StrSQL = StrSQL & "  SELECT dbo.Tbl_TradingContract.ID, dbo.Tbl_TradingContract.TContract_DateH, dbo.Tbl_TradingContract.TContract_Date, dbo.TblCustemers.CusID,"
    StrSQL = StrSQL & "  dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.Tbl_TradingContract.TOrder_Address, dbo.Tbl_TradingContract.TOrder_Phone,"
    StrSQL = StrSQL & "  dbo.Tbl_TradingContract.UserID"
    StrSQL = StrSQL & " FROM   dbo.Tbl_TradingContract INNER JOIN"
    StrSQL = StrSQL & "   dbo.TblCustemers ON dbo.Tbl_TradingContract.TContract_CustID = dbo.TblCustemers.CusID"
    StrSQL = StrSQL & " where 1=1 AND (NOT (dbo.Tbl_TradingContract.ID IS NULL))"
    StrSQL = StrSQL & " And IsNull(IsCanceld,0) <> 1"
    Begin = True
  

    If TxT_TradContract.text <> "" Then
              StrWhere = StrWhere + " and dbo.Tbl_TradingContract.ID like '%" & (TxT_TradContract.text) & "%'"
      End If
    If DcbCusTC.BoundText <> "" Then
                 StrWhere = StrWhere + " and ( dbo.TblCustemers.CusID LIKE '%" & Trim(DcbCusTC.BoundText) & "%')"
    End If
     If TxtAddress.text <> "" Then
              StrWhere = StrWhere + " and dbo.Tbl_TradingContract.TOrder_Address ='" & (TxtAddress.text) & "'"
    End If
   
      If TxtPhone.text <> "" Then
              StrWhere = StrWhere + " and dbo.Tbl_TradingContract.TOrder_Phone =" & (TxtPhone.text) & ""
      End If
      
      
      If Not IsNull(Me.DTP_TD_From.value) Then
             StrWhere = StrWhere & " AND dbo.Tbl_TradingContract.TContract_Date>=" & SQLDate(Me.DTP_TD_From.value, True) & ""
      End If
       If Not IsNull(Me.DTP_TD_To.value) Then
             StrWhere = StrWhere & " AND dbo.Tbl_TradingContract.TContract_Date <=" & SQLDate(Me.DTP_TD_To.value, True) & ""
      End If
      
    
      sql = StrSQL + StrWhere + " order by dbo.Tbl_TradingContract.ID"
   
   ''----------------------------------------------------------------------
  
    Dim Num As Integer
    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    VSFlexGrid4.Clear flexClearScrollable, flexClearEverything
    If Not (rs.EOF Or rs.BOF) Then
        VSFlexGrid4.rows = rs.RecordCount + 1
        
        For Num = 1 To rs.RecordCount
       
            With VSFlexGrid4
                .TextMatrix(Num, .ColIndex("Count")) = Num
                .TextMatrix(Num, .ColIndex("id")) = IIf(IsNull(rs("id").value), "", Trim(rs("id").value))
                .TextMatrix(Num, .ColIndex("DateHig")) = IIf(IsNull(rs("TContract_DateH").value), "", (rs("TContract_DateH").value))
                .TextMatrix(Num, .ColIndex("DateTOrder")) = IIf(IsNull(rs("TContract_Date").value), "", Trim(rs("TContract_Date").value))
                .TextMatrix(Num, .ColIndex("CustmerName")) = IIf(IsNull(rs("CusName").value), "", (rs("CusName").value))
                .TextMatrix(Num, .ColIndex("Address")) = IIf(IsNull(rs("TOrder_Address").value), "", (rs("TOrder_Address").value))
                .TextMatrix(Num, .ColIndex("Phone")) = IIf(IsNull(rs("TOrder_Phone").value), "", (rs("TOrder_Phone").value))
          
            End With
            rs.MoveNext
        Next Num
      
End If
   
   ''----------------------------------------------------------------------

ErrTrap:

End Sub


Private Sub GetTradingInvContract()


    Dim sql As String
    Dim StrSQL As String
    Dim Begin As Boolean
  '  Public Current_branch As Integer
   ' Public Current_branchSql As String

    Dim StrWhere As String
   ' On Error GoTo ErrTrap


    StrSQL = StrSQL & "  SELECT dbo.Tbl_TradingContract.ID,Tbl_TradingContractInv.ID InvID, Tbl_TradingContractInv.NoteSerial1, dbo.Tbl_TradingContract.TContract_DateH, dbo.Tbl_TradingContract.TContract_Date, dbo.TblCustemers.CusID,"
    StrSQL = StrSQL & "  dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.Tbl_TradingContract.TOrder_Address, dbo.Tbl_TradingContract.TOrder_Phone,"
    StrSQL = StrSQL & "  dbo.Tbl_TradingContract.UserID"
    StrSQL = StrSQL & " FROM   dbo.Tbl_TradingContract INNER JOIN"
    StrSQL = StrSQL & "   dbo.TblCustemers ON dbo.Tbl_TradingContract.TContract_CustID = dbo.TblCustemers.CusID"
    StrSQL = StrSQL & "   INNER JOIN"
    StrSQL = StrSQL & "   dbo.Tbl_TradingContractInv ON dbo.Tbl_TradingContract.ID = dbo.Tbl_TradingContractInv.TradingContractID"
    
    StrSQL = StrSQL & " where 1=1 AND (NOT (dbo.Tbl_TradingContract.ID IS NULL))"
    StrSQL = StrSQL & " And IsNull(IsCanceld,0) <> 1"
    
    Begin = True
  
    If Trim(TxtNoteSerial1) <> "" Then
          StrWhere = StrWhere + " and dbo.Tbl_TradingContractInv.NoteSerial1 like '%" & (TxtNoteSerial1.text) & "%'"
    End If
    
    If TxT_TradContract.text <> "" Then
              StrWhere = StrWhere + " and dbo.Tbl_TradingContract.ID like '%" & (TxT_TradContract.text) & "%'"
      End If
    If DcbCusTC.BoundText <> "" Then
                 StrWhere = StrWhere + " and ( dbo.TblCustemers.CusID LIKE '%" & Trim(DcbCusTC.BoundText) & "%')"
    End If
     If TxtAddress.text <> "" Then
              StrWhere = StrWhere + " and dbo.Tbl_TradingContract.TOrder_Address ='" & (TxtAddress.text) & "'"
    End If
   
      If TxtPhone.text <> "" Then
              StrWhere = StrWhere + " and dbo.Tbl_TradingContract.TOrder_Phone =" & (TxtPhone.text) & ""
      End If
      
      
      If Not IsNull(Me.DTP_TD_From.value) Then
             StrWhere = StrWhere & " AND dbo.Tbl_TradingContractInv.RecordDate>=" & SQLDate(Me.DTP_TD_From.value, True) & ""
      End If
       If Not IsNull(Me.DTP_TD_To.value) Then
             StrWhere = StrWhere & " AND dbo.Tbl_TradingContractInv.RecordDate <=" & SQLDate(Me.DTP_TD_To.value, True) & ""
      End If
      
    
      sql = StrSQL + StrWhere + " order by dbo.Tbl_TradingContract.ID"
   
   ''----------------------------------------------------------------------
  
    Dim Num As Integer
    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    VSFlexGrid4.Clear flexClearScrollable, flexClearEverything
    If Not (rs.EOF Or rs.BOF) Then
        VSFlexGrid4.rows = rs.RecordCount + 1
        
        For Num = 1 To rs.RecordCount
       
            With VSFlexGrid4
                .TextMatrix(Num, .ColIndex("Count")) = Num
                .TextMatrix(Num, .ColIndex("id")) = IIf(IsNull(rs("id").value), "", Trim(rs("id").value))
                .TextMatrix(Num, .ColIndex("DateHig")) = IIf(IsNull(rs("TContract_DateH").value), "", (rs("TContract_DateH").value))
                .TextMatrix(Num, .ColIndex("DateTOrder")) = IIf(IsNull(rs("TContract_Date").value), "", Trim(rs("TContract_Date").value))
                .TextMatrix(Num, .ColIndex("CustmerName")) = IIf(IsNull(rs("CusName").value), "", (rs("CusName").value))
                .TextMatrix(Num, .ColIndex("Address")) = IIf(IsNull(rs("TOrder_Address").value), "", (rs("TOrder_Address").value))
                .TextMatrix(Num, .ColIndex("Phone")) = IIf(IsNull(rs("TOrder_Phone").value), "", (rs("TOrder_Phone").value))
                .TextMatrix(Num, .ColIndex("NoteSerial1")) = IIf(IsNull(rs("NoteSerial1").value), "", (rs("NoteSerial1").value))
                .TextMatrix(Num, .ColIndex("InvID")) = IIf(IsNull(rs("InvID").value), "", (rs("InvID").value))
          
            End With
            rs.MoveNext
        Next Num
      
End If
   
   ''----------------------------------------------------------------------

ErrTrap:

End Sub

Private Sub ISButton8_Click()
 clear_all Me
DTP_TD_From.value = ""
DTP_TD_To.value = ""


End Sub

Private Sub ISButton9_Click()
Unload Me
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
  Dim CUSTID As Integer
'getemployeeCode

    If KeyAscii = vbKeyReturn Then
        GetCustomersDetail CUSTID, , Text3.text, 1
        Me.DcbCustomer.BoundText = CUSTID
    End If
End Sub

Private Sub Txt_CustomerCode_KeyPress(KeyAscii As Integer)

Dim CUSTID As Integer
If KeyAscii = vbKeyReturn Then
        GetTblCustemersCode Txt_CustomerCode.text, CUSTID
        DcbCustmer.BoundText = CUSTID
    End If

End Sub

Private Sub TxtItemCode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If TxtItemCode.text = "" Then
            Me.DcboItems.BoundText = ""
        Else
            Me.DcboItems.BoundText = GetItemID(Trim$(Me.TxtItemCode.text))
        End If
    End If
End Sub
Private Sub DcboItems_Change()
DcboItems_Click (0)
End Sub

Private Sub DcboItems_Click(Area As Integer)
  Me.TxtItemCode.text = GetItemCode(val(Me.DcboItems.BoundText))
End Sub
Private Sub TxtSearchCode_KeyPress(KeyAscii As Integer)
   Dim EmpID As Integer

       If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode TxtSearchCode.text, EmpID
        DcbEmployee2.BoundText = EmpID
    End If
    
End Sub

Private Sub VSFlexGrid1_Click()

If Label11.Caption = 2 Then

 Frm_TRansOrder.Txt_OrderNumber.text = val(VSFlexGrid1.TextMatrix(VSFlexGrid1.Row, 1))
 Unload Me

ElseIf Label11.Caption = 5 Then

 FrmReportsStudent.Txt_OrderNumber.text = val(VSFlexGrid1.TextMatrix(VSFlexGrid1.Row, 1))
 Unload Me

Else
 Frm_NewMeasure.FindRec val(VSFlexGrid1.TextMatrix(VSFlexGrid1.Row, 1))
 Unload Me
 
 End If


End Sub

Private Sub VSFlexGrid2_Click()
If Label11.Caption = 0 Then
    If Indx2 = 2 Then
        FrmReportsStudent.Txt_OrderNumber.text = val(VSFlexGrid2.TextMatrix(VSFlexGrid2.Row, 2))
        FrmReportsStudent.Txt_OrderNumber2.text = val(VSFlexGrid2.TextMatrix(VSFlexGrid2.Row, 1))
    Else
        FrmReportsStudent.Txt_OrderNumber.text = val(VSFlexGrid2.TextMatrix(VSFlexGrid2.Row, 2))
        FrmReportsStudent.Txt_OrderNumber2.text = ""
    End If
 Unload Me
 Exit Sub
End If
 Frm_TRansOrder.FindRec val(VSFlexGrid2.TextMatrix(VSFlexGrid2.Row, 1))
 Unload Me
End Sub

Private Sub VSFlexGrid3_Click()
 Frm_BusinessDialy.FindRec val(VSFlexGrid3.TextMatrix(VSFlexGrid3.Row, 1))
 Unload Me
End Sub

Private Sub VSFlexGrid4_Click()

 If Label11.Caption = 1 Then
 Frm_BusinessDialy.TxtSearchCode.text = val(VSFlexGrid4.TextMatrix(VSFlexGrid4.Row, 1))
 Unload Me
 ElseIf Label11.Caption = 3 Then
  FrmReportsStudent.TxtSearchCode(0).text = val(VSFlexGrid4.TextMatrix(VSFlexGrid4.Row, 1))
  Unload Me
 ElseIf Label11.Caption = 4 Then
    FrmReportsStudent.TxtSearchCode(1).text = val(VSFlexGrid4.TextMatrix(VSFlexGrid4.Row, 1))
    Unload Me
 ElseIf Label11.Caption = 5 Then
    FrmOut.txtTradingContractID.text = val(VSFlexGrid4.TextMatrix(VSFlexGrid4.Row, 1))
    Unload Me
 Else
 If Indx = 15 Then
        emp_CONTRACT_TYPE.mIndex = 3
        emp_CONTRACT_TYPE.FindRec val(VSFlexGrid4.TextMatrix(VSFlexGrid4.Row, VSFlexGrid4.ColIndex("InvID")))
     
 Else
    Frm_TradingContract.FindRec val(VSFlexGrid4.TextMatrix(VSFlexGrid4.Row, 1))
  End If
 Unload Me
 End If


End Sub

Private Sub VSFlexGrid5_Click()
FrmPaymenTransTrip.Retrive val(VSFlexGrid5.TextMatrix(VSFlexGrid5.Row, VSFlexGrid5.ColIndex("ID")))
End Sub

Private Sub VSFlexGrid6_Click()
frmserviceInvoice.Retrive val(VSFlexGrid6.TextMatrix(VSFlexGrid6.Row, VSFlexGrid6.ColIndex("ID")))
End Sub

Private Sub GetServiceInvoice()
    Dim sql    As String
    Dim StrSQL As String
    Dim Begin As Boolean
    Dim StrWhere As String
    Dim isAr As Boolean
    isAr = (SystemOptions.UserInterface = ArabicInterface)
    ' On Error GoTo ErrTrap

'    StrSQL = " SELECT        dbo.notes_all.NoteID ID,dbo.notes_all.NoteDate,dbo.notes_all.branch_no ,"
'    StrSQL = StrSQL & "                      dbo.TblBranchesData.branch_name,dbo.TblBranchesData.branch_namee,dbo.notes_all.CusID,dbo.TblCustemers.CusName,dbo.TblCustemers.CusNamee,"
'    StrSQL = StrSQL & "                      dbo.notes_all.NoteSerial1"
'    StrSQL = StrSQL & "     FROM            dbo.notes_all LEFT OUTER JOIN"
'    StrSQL = StrSQL & "                     dbo.TblCustemers ON dbo.notes_all.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
'    StrSQL = StrSQL & "                     dbo.TblBranchesData ON dbo.notes_all.branch_no = dbo.TblBranchesData.branch_id"
            StrSQL = " SELECT"
            StrSQL = StrSQL & " ISNULL(dbo.notes_all.NoteID , 0 )  ID,"
            StrSQL = StrSQL & "  ISNULL( dbo.notes_all.NoteSerial1 , '') NoteSerial1,"
            StrSQL = StrSQL & "    dbo.notes_all.NoteDate recordDate,"
            StrSQL = StrSQL & "  dbo.TblBranchesData.branch_name" & IIf(isAr, "", "e") & " branch_name,"
            StrSQL = StrSQL & "   dbo.TblCustemers.CusName" & IIf(isAr, "", "e") & " CusName"
            StrSQL = StrSQL & " From dbo.notes_all"
            StrSQL = StrSQL & "   LEFT OUTER JOIN dbo.TblCustemers"
            StrSQL = StrSQL & "     ON dbo.notes_all.CusID = dbo.TblCustemers.CusID"
            StrSQL = StrSQL & "  LEFT OUTER JOIN dbo.TblBranchesData"
            StrSQL = StrSQL & "    ON dbo.notes_all.branch_no = dbo.TblBranchesData.branch_id"
    
    
    
    StrSQL = StrSQL & " where 1=1 and NoteType = 85 "
    If val(TxtIDFrom2.text) <> 0 Then
        StrWhere = StrWhere + " and dbo.notes_all.NoteSerial1 >= " & val(TxtIDFrom2.text) & ""
    End If
    If val(TxtIDTO2.text) <> 0 Then
        StrWhere = StrWhere + " and dbo.notes_all.NoteSerial1 <=" & val(TxtIDTO2.text) & ""
    End If
    If val(DcbBranch2.BoundText) <> 0 And DcbBranch2.text <> "" Then
        StrWhere = StrWhere + " and dbo.notes_all.branch_no  = " & val(DcbBranch2.BoundText) & ""
    End If
  
    If val(DcbCustomer2.BoundText) <> 0 And DcbCustomer2.text <> "" Then
        StrWhere = StrWhere + " and dbo.notes_all.CusID  = " & val(DcbCustomer2.BoundText) & ""
    End If
      
    If Not IsNull(Me.DtpDateFrom2.value) Then
        StrWhere = StrWhere & " AND dbo.notes_all.NoteDate>=" & SQLDate(Me.DtpDateFrom2.value, True) & ""
    End If
    If Not IsNull(Me.DtpDateTo2.value) Then
        StrWhere = StrWhere & " AND dbo.notes_all.NoteDate <=" & SQLDate(Me.DtpDateTo2.value, True) & ""
    End If
    
    sql = StrSQL + StrWhere + " order by dbo.notes_all.NoteSerial1"
   
    ''----------------------------------------------------------------------
    Dim Auto_Man As Integer
    Dim Num      As Integer
    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
'    VSFlexGrid6.DataMode = flexDMFree
    Set VSFlexGrid6.DataSource = rs
'
'    VSFlexGrid6.Clear flexClearScrollable, flexClearEverything
'    VSFlexGrid6.Rows = 2
'    If rs.RecordCount > 0 Then
'        VSFlexGrid6.Rows = rs.RecordCount + 1
'
'        For Num = 1 To rs.RecordCount
'
'            With VSFlexGrid6
'                .TextMatrix(Num, .ColIndex("Count")) = Num
'                .TextMatrix(Num, .ColIndex("ID")) = IIf(IsNull(rs("ID").value), 0, Trim(rs("ID").value))
'                .TextMatrix(Num, .ColIndex("NoteSerial1")) = IIf(IsNull(rs("NoteSerial1").value), "", Trim(rs("NoteSerial1").value))
'                .TextMatrix(Num, .ColIndex("recordDate")) = IIf(IsNull(rs("NoteDate").value), "", (rs("NoteDate").value))
'                If SystemOptions.UserInterface = ArabicInterface Then
'                    '.TextMatrix(Num, .ColIndex("Name")) = IIf(IsNull(rs("Name").value), "", (rs("Name").value))
'                    .TextMatrix(Num, .ColIndex("CusName")) = IIf(IsNull(rs("CusName").value), "", Trim(rs("CusName").value))
'                    .TextMatrix(Num, .ColIndex("branch_name")) = IIf(IsNull(rs("branch_name").value), "", (rs("branch_name").value))
'                Else
'                    .TextMatrix(Num, .ColIndex("Name")) = IIf(IsNull(rs("NameE").value), "", (rs("NameE").value))
'                    .TextMatrix(Num, .ColIndex("CusName")) = IIf(IsNull(rs("CusNamee").value), "", Trim(rs("CusNamee").value))
'                    .TextMatrix(Num, .ColIndex("branch_name")) = IIf(IsNull(rs("branch_namee").value), "", (rs("branch_namee").value))
'                End If
'
'            End With
'            rs.MoveNext
'        Next Num
'
'    End If
   
    ''----------------------------------------------------------------------

ErrTrap:

End Sub

