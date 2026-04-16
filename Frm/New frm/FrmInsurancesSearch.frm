VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SuiteCtrls.ocx"
Begin VB.Form FrmInsurancesSearch 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7425
   ClientLeft      =   4260
   ClientTop       =   5430
   ClientWidth     =   13470
   Icon            =   "FrmInsurancesSearch.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7425
   ScaleWidth      =   13470
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin C1SizerLibCtl.C1Elastic C1Elastic3 
      Height          =   7410
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   13485
      _cx             =   23786
      _cy             =   13070
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
      Begin VB.Frame Frame1 
         BackColor       =   &H00E2E9E9&
         Height          =   780
         Left            =   0
         TabIndex        =   1
         Top             =   6495
         Width           =   13470
         Begin ImpulseButton.ISButton Cmd 
            Height          =   375
            Index           =   0
            Left            =   9480
            TabIndex        =   2
            Top             =   240
            Width           =   3525
            _ExtentX        =   6218
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
            ButtonImage     =   "FrmInsurancesSearch.frx":6852
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
            Height          =   375
            Index           =   1
            Left            =   5160
            TabIndex        =   3
            Top             =   240
            Width           =   3555
            _ExtentX        =   6271
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
            ButtonImage     =   "FrmInsurancesSearch.frx":D0B4
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
            Height          =   375
            Index           =   2
            Left            =   480
            TabIndex        =   4
            Top             =   240
            Width           =   3735
            _ExtentX        =   6588
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
            ButtonImage     =   "FrmInsurancesSearch.frx":13916
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
      Begin C1SizerLibCtl.C1Tab C1Tab1 
         Height          =   6930
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   13485
         _cx             =   23786
         _cy             =   12224
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial (Arabic)"
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
         ForeColor       =   128
         FrontTabColor   =   14871017
         BackTabColor    =   8454143
         TabOutlineColor =   -2147483632
         FrontTabForeColor=   -2147483630
         Caption         =   "0|1|2|3|4|5|6|»ÕÀ «„—  Õ„Ì·"
         Align           =   0
         CurrTab         =   7
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
         Begin C1SizerLibCtl.C1Elastic C1Elastic7 
            Height          =   6510
            Left            =   -15840
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   45
            Width           =   13395
            _cx             =   23627
            _cy             =   11483
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
            Begin VB.Frame Frame7 
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   855
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   11
               Top             =   0
               Width           =   13350
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "»ÕÀ ‘«‘… «· ﬁÌÌ„"
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
                  Index           =   0
                  Left            =   6600
                  RightToLeft     =   -1  'True
                  TabIndex        =   12
                  Top             =   240
                  Width           =   5400
               End
               Begin VB.Image Image2 
                  Height          =   615
                  Left            =   12360
                  Picture         =   "FrmInsurancesSearch.frx":3D538
                  Stretch         =   -1  'True
                  Top             =   120
                  Width           =   735
               End
            End
            Begin VB.TextBox Emp_Code 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   9720
               RightToLeft     =   -1  'True
               TabIndex        =   10
               Top             =   5595
               Width           =   2160
            End
            Begin VB.ComboBox YearID 
               Height          =   315
               Left            =   5595
               RightToLeft     =   -1  'True
               TabIndex        =   9
               Top             =   5070
               Width           =   1845
            End
            Begin VB.ComboBox MonthID 
               Height          =   315
               Left            =   720
               RightToLeft     =   -1  'True
               TabIndex        =   8
               Top             =   5070
               Width           =   2700
            End
            Begin VB.TextBox ID 
               Alignment       =   1  'Right Justify
               Height          =   330
               Left            =   9735
               MaxLength       =   10
               RightToLeft     =   -1  'True
               TabIndex        =   7
               Top             =   5070
               Width           =   2190
            End
            Begin VSFlex8UCtl.VSFlexGrid fg_Evaluation 
               Height          =   3870
               Left            =   120
               TabIndex        =   13
               Top             =   915
               Width           =   13200
               _cx             =   23283
               _cy             =   6826
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
               BackColor       =   16777215
               ForeColor       =   -2147483640
               BackColorFixed  =   14871017
               ForeColorFixed  =   -2147483630
               BackColorSel    =   -2147483635
               ForeColorSel    =   -2147483634
               BackColorBkg    =   -2147483633
               BackColorAlternate=   16777088
               GridColor       =   -2147483633
               GridColorFixed  =   -2147483633
               TreeColor       =   -2147483632
               FloodColor      =   192
               SheetBorder     =   -2147483633
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
               FormatString    =   $"FrmInsurancesSearch.frx":3ECE5
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
            Begin MSDataListLib.DataCombo BranchID 
               Height          =   315
               Left            =   720
               TabIndex        =   14
               Top             =   5595
               Width           =   2700
               _ExtentX        =   4763
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               BackColor       =   16777215
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo EmployeeID 
               Height          =   315
               Left            =   5565
               TabIndex        =   15
               Top             =   5595
               Width           =   1875
               _ExtentX        =   3307
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               BackColor       =   16777215
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«·”‰…"
               ForeColor       =   &H00000000&
               Height          =   180
               Index           =   1
               Left            =   7800
               TabIndex        =   21
               Top             =   5070
               Width           =   1725
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ﬁ«∆„ »«· ﬁÌÌ„"
               Height          =   225
               Index           =   12
               Left            =   7530
               RightToLeft     =   -1  'True
               TabIndex        =   20
               Top             =   5595
               Width           =   1995
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·›—⁄"
               Height          =   345
               Index           =   24
               Left            =   3720
               RightToLeft     =   -1  'True
               TabIndex        =   19
               Top             =   5595
               Width           =   990
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·—ﬁ„"
               Height          =   255
               Index           =   15
               Left            =   12015
               RightToLeft     =   -1  'True
               TabIndex        =   18
               Top             =   5070
               Width           =   1065
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬂÊœ"
               Height          =   225
               Index           =   16
               Left            =   12045
               RightToLeft     =   -1  'True
               TabIndex        =   17
               Top             =   5595
               Width           =   1065
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«·‘Â—"
               ForeColor       =   &H00000000&
               Height          =   300
               Index           =   3
               Left            =   3585
               TabIndex        =   16
               Top             =   5070
               Width           =   1095
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic1 
            Height          =   6510
            Left            =   -15540
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   45
            Width           =   13395
            _cx             =   23627
            _cy             =   11483
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
            Begin VB.Frame FraHeader 
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   855
               Left            =   -120
               RightToLeft     =   -1  'True
               TabIndex        =   58
               Top             =   0
               Width           =   13665
               Begin VB.Image Image1 
                  Height          =   615
                  Left            =   12360
                  Picture         =   "FrmInsurancesSearch.frx":3EE27
                  Stretch         =   -1  'True
                  Top             =   120
                  Width           =   735
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·»ÕÀ ⁄‰  ”ÃÌ· «À»«  «· √„Ì‰« "
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
                  Left            =   6600
                  RightToLeft     =   -1  'True
                  TabIndex        =   59
                  Top             =   240
                  Width           =   5400
               End
            End
            Begin VB.Frame lbreg 
               BackColor       =   &H00E2E9E9&
               Height          =   810
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   52
               Top             =   3780
               Width           =   7215
               Begin MSComCtl2.DTPicker DtpDateFrom 
                  Height          =   330
                  Left            =   3120
                  TabIndex        =   53
                  Top             =   240
                  Width           =   1575
                  _ExtentX        =   2778
                  _ExtentY        =   582
                  _Version        =   393216
                  CheckBox        =   -1  'True
                  Format          =   192610307
                  CurrentDate     =   38887
               End
               Begin MSComCtl2.DTPicker DtpDateTo 
                  Height          =   330
                  Left            =   240
                  TabIndex        =   54
                  Top             =   240
                  Width           =   1575
                  _ExtentX        =   2778
                  _ExtentY        =   582
                  _Version        =   393216
                  CheckBox        =   -1  'True
                  Format          =   192610307
                  CurrentDate     =   38887
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "≈·Ï"
                  Height          =   315
                  Index           =   3
                  Left            =   1920
                  RightToLeft     =   -1  'True
                  TabIndex        =   57
                  Top             =   240
                  Width           =   1080
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„‰"
                  Height          =   315
                  Index           =   4
                  Left            =   4800
                  RightToLeft     =   -1  'True
                  TabIndex        =   56
                  Top             =   240
                  Width           =   660
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " «—ÌŒ «·⁄„·Ì…"
                  Height          =   195
                  Index           =   13
                  Left            =   5760
                  RightToLeft     =   -1  'True
                  TabIndex        =   55
                  Top             =   240
                  Width           =   1425
               End
            End
            Begin VB.Frame lbprocess 
               BackColor       =   &H00E2E9E9&
               Height          =   810
               Left            =   7320
               RightToLeft     =   -1  'True
               TabIndex        =   46
               Top             =   3780
               Width           =   5835
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
                  Height          =   345
                  Left            =   2400
                  RightToLeft     =   -1  'True
                  TabIndex        =   48
                  Top             =   240
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
                  TabIndex        =   47
                  Top             =   240
                  Width           =   1155
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„‰"
                  Height          =   315
                  Index           =   5
                  Left            =   3720
                  RightToLeft     =   -1  'True
                  TabIndex        =   51
                  Top             =   240
                  Width           =   660
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "≈·Ï"
                  Height          =   315
                  Index           =   6
                  Left            =   1560
                  RightToLeft     =   -1  'True
                  TabIndex        =   50
                  Top             =   240
                  Width           =   645
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "—ﬁ„ «·”‰œ"
                  Height          =   195
                  Index           =   14
                  Left            =   4440
                  RightToLeft     =   -1  'True
                  TabIndex        =   49
                  Top             =   240
                  Width           =   1185
               End
            End
            Begin VB.Frame lblLW 
               BackColor       =   &H00E2E9E9&
               Height          =   1050
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   27
               Top             =   4560
               Width           =   13020
               Begin VB.Frame Frame4 
                  BackColor       =   &H00E2E9E9&
                  Height          =   855
                  Left            =   0
                  RightToLeft     =   -1  'True
                  TabIndex        =   28
                  Top             =   120
                  Visible         =   0   'False
                  Width           =   6015
                  Begin VB.TextBox TxtStay 
                     Alignment       =   2  'Center
                     Height          =   315
                     Left            =   1920
                     TabIndex        =   30
                     Top             =   600
                     Width           =   930
                  End
                  Begin VB.TextBox TxtCivilin 
                     Alignment       =   2  'Center
                     Height          =   315
                     Left            =   1920
                     TabIndex        =   29
                     Top             =   240
                     Width           =   930
                  End
                  Begin ImpulseButton.ISButton Refrish 
                     Height          =   270
                     Left            =   1440
                     TabIndex        =   31
                     TabStop         =   0   'False
                     Top             =   240
                     Width           =   345
                     _ExtentX        =   609
                     _ExtentY        =   476
                     ButtonStyle     =   1
                     ButtonPositionImage=   1
                     Caption         =   ""
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
                     ButtonImage     =   "FrmInsurancesSearch.frx":405D4
                     ColorButton     =   14871017
                     DrawFocusRectangle=   0   'False
                     DisabledImageStyle=   1
                  End
                  Begin ImpulseButton.ISButton Refrish2 
                     Height          =   270
                     Left            =   1440
                     TabIndex        =   32
                     TabStop         =   0   'False
                     Top             =   600
                     Width           =   345
                     _ExtentX        =   609
                     _ExtentY        =   476
                     ButtonStyle     =   1
                     ButtonPositionImage=   1
                     Caption         =   ""
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
                     ButtonImage     =   "FrmInsurancesSearch.frx":46E36
                     ColorButton     =   14871017
                     DrawFocusRectangle=   0   'False
                     DisabledImageStyle=   1
                  End
                  Begin VB.Label lbl 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   " %"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   11.25
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H000040C0&
                     Height          =   315
                     Index           =   9
                     Left            =   1200
                     RightToLeft     =   -1  'True
                     TabIndex        =   36
                     Top             =   600
                     Width           =   300
                  End
                  Begin VB.Label lbl 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   " %"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   11.25
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H000040C0&
                     Height          =   315
                     Index           =   8
                     Left            =   1200
                     RightToLeft     =   -1  'True
                     TabIndex        =   35
                     Top             =   240
                     Width           =   300
                  End
                  Begin VB.Label lbl 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   " ‰”»… «·„ﬁÌ„Ì‰"
                     Height          =   315
                     Index           =   1
                     Left            =   2955
                     RightToLeft     =   -1  'True
                     TabIndex        =   34
                     Top             =   600
                     Width           =   2340
                  End
                  Begin VB.Label lbl 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     BackStyle       =   0  'Transparent
                     Caption         =   " ‰”»… «·„Ê«ÿ‰Ì‰"
                     Height          =   315
                     Index           =   0
                     Left            =   2955
                     RightToLeft     =   -1  'True
                     TabIndex        =   33
                     Top             =   240
                     Width           =   2340
                  End
               End
               Begin MSDataListLib.DataCombo Dcbranch 
                  Height          =   315
                  Left            =   6120
                  TabIndex        =   37
                  Top             =   600
                  Visible         =   0   'False
                  Width           =   4155
                  _ExtentX        =   7329
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DcboBox 
                  Height          =   315
                  Left            =   6120
                  TabIndex        =   38
                  Top             =   600
                  Visible         =   0   'False
                  Width           =   4155
                  _ExtentX        =   7329
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin C1SizerLibCtl.C1Elastic Ele 
                  Height          =   825
                  Index           =   3
                  Left            =   6480
                  TabIndex        =   39
                  TabStop         =   0   'False
                  Top             =   120
                  Width           =   6255
                  _cx             =   11033
                  _cy             =   1455
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
                  Caption         =   " Õœœ «·› —…"
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
                  Begin VB.ComboBox CmbMonth 
                     Height          =   315
                     Left            =   3315
                     RightToLeft     =   -1  'True
                     TabIndex        =   41
                     Top             =   225
                     Width           =   1485
                  End
                  Begin VB.ComboBox CboYear 
                     Height          =   315
                     Left            =   75
                     RightToLeft     =   -1  'True
                     TabIndex        =   40
                     Top             =   225
                     Width           =   1830
                  End
                  Begin VB.Label lbl 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "‘Â—"
                     Height          =   195
                     Index           =   10
                     Left            =   4905
                     RightToLeft     =   -1  'True
                     TabIndex        =   43
                     Top             =   240
                     Width           =   870
                  End
                  Begin VB.Label lbl 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "”‰…"
                     Height          =   240
                     Index           =   11
                     Left            =   2160
                     RightToLeft     =   -1  'True
                     TabIndex        =   42
                     Top             =   240
                     Width           =   900
                  End
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«”„ «·„›—œ"
                  Height          =   285
                  Index           =   7
                  Left            =   10920
                  TabIndex        =   45
                  Top             =   600
                  Visible         =   0   'False
                  Width           =   1845
               End
               Begin VB.Label lblLL 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·›—⁄"
                  Height          =   285
                  Left            =   10920
                  TabIndex        =   44
                  Top             =   240
                  Visible         =   0   'False
                  Width           =   1965
               End
            End
            Begin VB.Frame Frame2 
               BackColor       =   &H00E2E9E9&
               Height          =   750
               Left            =   120
               TabIndex        =   23
               Top             =   5595
               Width           =   13020
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·≈Ã„«·Ì"
                  Height          =   285
                  Index           =   2
                  Left            =   6360
                  RightToLeft     =   -1  'True
                  TabIndex        =   26
                  Top             =   240
                  Width           =   945
               End
               Begin VB.Label lblL 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  ForeColor       =   &H00000080&
                  Height          =   285
                  Index           =   0
                  Left            =   2640
                  RightToLeft     =   -1  'True
                  TabIndex        =   25
                  Top             =   240
                  Width           =   2145
               End
               Begin VB.Label lblL 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  ForeColor       =   &H00000080&
                  Height          =   315
                  Index           =   10
                  Left            =   8520
                  RightToLeft     =   -1  'True
                  TabIndex        =   24
                  Top             =   240
                  Width           =   2775
               End
            End
            Begin VSFlex8UCtl.VSFlexGrid Fg 
               Height          =   2850
               Left            =   120
               TabIndex        =   60
               Top             =   915
               Width           =   13035
               _cx             =   22992
               _cy             =   5027
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
               BackColor       =   16777215
               ForeColor       =   -2147483640
               BackColorFixed  =   14871017
               ForeColorFixed  =   -2147483630
               BackColorSel    =   -2147483635
               ForeColorSel    =   -2147483634
               BackColorBkg    =   -2147483633
               BackColorAlternate=   16777088
               GridColor       =   -2147483633
               GridColorFixed  =   -2147483633
               TreeColor       =   -2147483632
               FloodColor      =   192
               SheetBorder     =   -2147483633
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
               FormatString    =   $"FrmInsurancesSearch.frx":4D698
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
         Begin C1SizerLibCtl.C1Elastic C1Elastic2 
            Height          =   6510
            Left            =   -15240
            TabIndex        =   61
            TabStop         =   0   'False
            Top             =   45
            Width           =   13395
            _cx             =   23627
            _cy             =   11483
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
            Begin VB.Frame Frame8 
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   780
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   85
               Top             =   0
               Width           =   13425
               Begin VB.Image Image3 
                  Height          =   615
                  Left            =   12360
                  Picture         =   "FrmInsurancesSearch.frx":4D82C
                  Stretch         =   -1  'True
                  Top             =   120
                  Width           =   735
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·»ÕÀ  ⁄—Ì› «·„Â«„ Ê«·⁄„·Ì« "
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
                  Index           =   4
                  Left            =   8160
                  RightToLeft     =   -1  'True
                  TabIndex        =   86
                  Top             =   240
                  Width           =   3720
               End
            End
            Begin VB.Frame Frame12 
               BackColor       =   &H00E2E9E9&
               Height          =   735
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   79
               Top             =   3600
               Width           =   7095
               Begin MSComCtl2.DTPicker DateFrom 
                  Height          =   330
                  Left            =   3000
                  TabIndex        =   80
                  Top             =   240
                  Width           =   1575
                  _ExtentX        =   2778
                  _ExtentY        =   582
                  _Version        =   393216
                  CheckBox        =   -1  'True
                  Format          =   191496195
                  CurrentDate     =   38887
               End
               Begin MSComCtl2.DTPicker DateTo 
                  Height          =   330
                  Left            =   600
                  TabIndex        =   81
                  Top             =   240
                  Width           =   1575
                  _ExtentX        =   2778
                  _ExtentY        =   582
                  _Version        =   393216
                  CheckBox        =   -1  'True
                  Format          =   191496195
                  CurrentDate     =   38887
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "≈·Ï"
                  Height          =   315
                  Index           =   23
                  Left            =   1920
                  RightToLeft     =   -1  'True
                  TabIndex        =   84
                  Top             =   240
                  Width           =   1080
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„‰"
                  Height          =   315
                  Index           =   22
                  Left            =   4800
                  RightToLeft     =   -1  'True
                  TabIndex        =   83
                  Top             =   240
                  Width           =   660
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " «—ÌŒ «·Õ—ﬂ…"
                  Height          =   195
                  Index           =   21
                  Left            =   5640
                  RightToLeft     =   -1  'True
                  TabIndex        =   82
                  Top             =   240
                  Width           =   1185
               End
            End
            Begin VB.Frame Frame11 
               BackColor       =   &H00E2E9E9&
               Height          =   735
               Left            =   7200
               RightToLeft     =   -1  'True
               TabIndex        =   73
               Top             =   3600
               Width           =   6195
               Begin VB.TextBox IDFromTxt 
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
                  Left            =   2400
                  RightToLeft     =   -1  'True
                  TabIndex        =   75
                  Top             =   240
                  Width           =   1155
               End
               Begin VB.TextBox IDToTxt 
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
                  TabIndex        =   74
                  Top             =   240
                  Width           =   1155
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„‰"
                  Height          =   315
                  Index           =   20
                  Left            =   3720
                  RightToLeft     =   -1  'True
                  TabIndex        =   78
                  Top             =   240
                  Width           =   660
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "≈·Ï"
                  Height          =   315
                  Index           =   19
                  Left            =   1560
                  RightToLeft     =   -1  'True
                  TabIndex        =   77
                  Top             =   240
                  Width           =   645
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "—ﬁ„ «·Õ—ﬂ…"
                  Height          =   195
                  Index           =   18
                  Left            =   4440
                  RightToLeft     =   -1  'True
                  TabIndex        =   76
                  Top             =   240
                  Width           =   1185
               End
            End
            Begin VB.Frame Frame13 
               BackColor       =   &H00E2E9E9&
               Height          =   1695
               Left            =   90
               RightToLeft     =   -1  'True
               TabIndex        =   62
               Top             =   4200
               Width           =   13305
               Begin VB.TextBox NotesTxt 
                  Alignment       =   1  'Right Justify
                  Height          =   570
                  Left            =   720
                  RightToLeft     =   -1  'True
                  TabIndex        =   66
                  Top             =   840
                  Width           =   4905
               End
               Begin VB.TextBox DesTxt 
                  Alignment       =   1  'Right Justify
                  Height          =   570
                  Left            =   720
                  RightToLeft     =   -1  'True
                  TabIndex        =   65
                  Top             =   240
                  Width           =   4905
               End
               Begin VB.TextBox NameETxt 
                  Alignment       =   1  'Right Justify
                  Height          =   330
                  Left            =   7560
                  RightToLeft     =   -1  'True
                  TabIndex        =   64
                  Top             =   720
                  Width           =   4185
               End
               Begin VB.TextBox NameTxt 
                  Alignment       =   1  'Right Justify
                  Height          =   330
                  Left            =   7560
                  RightToLeft     =   -1  'True
                  TabIndex        =   63
                  Top             =   240
                  Width           =   4185
               End
               Begin MSDataListLib.DataCombo EmpCB 
                  Bindings        =   "FrmInsurancesSearch.frx":4EFD9
                  Height          =   315
                  Left            =   7560
                  TabIndex        =   67
                  Top             =   1200
                  Width           =   4185
                  _ExtentX        =   7382
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
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„·«ÕŸ« "
                  Height          =   195
                  Index           =   28
                  Left            =   5760
                  RightToLeft     =   -1  'True
                  TabIndex        =   72
                  Top             =   960
                  Width           =   1185
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Ê’› «·„Â„…"
                  Height          =   195
                  Index           =   27
                  Left            =   5760
                  RightToLeft     =   -1  'True
                  TabIndex        =   71
                  Top             =   360
                  Width           =   1185
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„œÌ— «·„Â„…"
                  Height          =   195
                  Index           =   26
                  Left            =   11880
                  RightToLeft     =   -1  'True
                  TabIndex        =   70
                  Top             =   1200
                  Width           =   1185
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«”„ «‰Ã·Ì“Ì"
                  Height          =   195
                  Index           =   25
                  Left            =   11880
                  RightToLeft     =   -1  'True
                  TabIndex        =   69
                  Top             =   720
                  Width           =   1185
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«”„ ⁄—»Ì"
                  Height          =   195
                  Index           =   17
                  Left            =   11880
                  RightToLeft     =   -1  'True
                  TabIndex        =   68
                  Top             =   240
                  Width           =   1185
               End
            End
            Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
               Height          =   2745
               Left            =   120
               TabIndex        =   87
               Top             =   840
               Width           =   13275
               _cx             =   23416
               _cy             =   4842
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
               BackColor       =   16777215
               ForeColor       =   -2147483640
               BackColorFixed  =   14871017
               ForeColorFixed  =   -2147483630
               BackColorSel    =   -2147483635
               ForeColorSel    =   -2147483634
               BackColorBkg    =   -2147483633
               BackColorAlternate=   16777088
               GridColor       =   -2147483633
               GridColorFixed  =   -2147483633
               TreeColor       =   -2147483632
               FloodColor      =   192
               SheetBorder     =   -2147483633
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
               FormatString    =   $"FrmInsurancesSearch.frx":4EFEE
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
         Begin C1SizerLibCtl.C1Elastic C1Elastic4 
            Height          =   6510
            Left            =   -14940
            TabIndex        =   88
            TabStop         =   0   'False
            Top             =   45
            Width           =   13395
            _cx             =   23627
            _cy             =   11483
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
            Begin VB.Frame Frame9 
               BackColor       =   &H00E2E9E9&
               Height          =   1455
               Left            =   90
               RightToLeft     =   -1  'True
               TabIndex        =   103
               Top             =   4440
               Width           =   13185
               Begin VB.TextBox EvaNameTxt 
                  Alignment       =   1  'Right Justify
                  Height          =   330
                  Left            =   7560
                  RightToLeft     =   -1  'True
                  TabIndex        =   106
                  Top             =   240
                  Width           =   4185
               End
               Begin VB.TextBox EvaNameeTxt 
                  Alignment       =   1  'Right Justify
                  Height          =   330
                  Left            =   1440
                  RightToLeft     =   -1  'True
                  TabIndex        =   105
                  Top             =   240
                  Width           =   4185
               End
               Begin VB.TextBox NoOfAbcDaysTxt 
                  Alignment       =   1  'Right Justify
                  Height          =   330
                  Left            =   7560
                  RightToLeft     =   -1  'True
                  TabIndex        =   104
                  Top             =   720
                  Width           =   4185
               End
               Begin XtremeSuiteControls.RadioButton Emp_Stude 
                  Height          =   255
                  Index           =   0
                  Left            =   3600
                  TabIndex        =   131
                  Top             =   840
                  Width           =   1815
                  _Version        =   786432
                  _ExtentX        =   3201
                  _ExtentY        =   450
                  _StockProps     =   79
                  Caption         =   "„ÊŸ›"
                  ForeColor       =   8388608
                  BackColor       =   14871017
                  UseVisualStyle  =   -1  'True
                  TextAlignment   =   1
                  RightToLeft     =   -1  'True
               End
               Begin XtremeSuiteControls.RadioButton Emp_Stude 
                  Height          =   255
                  Index           =   1
                  Left            =   1560
                  TabIndex        =   132
                  Top             =   840
                  Width           =   1695
                  _Version        =   786432
                  _ExtentX        =   2990
                  _ExtentY        =   450
                  _StockProps     =   79
                  Caption         =   "„ œ—»"
                  ForeColor       =   8388608
                  BackColor       =   14871017
                  UseVisualStyle  =   -1  'True
                  TextAlignment   =   1
                  RightToLeft     =   -1  'True
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«”„ ⁄—»Ì"
                  Height          =   195
                  Index           =   39
                  Left            =   11880
                  RightToLeft     =   -1  'True
                  TabIndex        =   109
                  Top             =   240
                  Width           =   1185
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·«”„ «‰Ã·Ì“Ì"
                  Height          =   195
                  Index           =   38
                  Left            =   5760
                  RightToLeft     =   -1  'True
                  TabIndex        =   108
                  Top             =   240
                  Width           =   1185
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "⁄œœ «Ì«„ «·€Ì«» "
                  Height          =   195
                  Index           =   35
                  Left            =   11880
                  RightToLeft     =   -1  'True
                  TabIndex        =   107
                  Top             =   720
                  Width           =   1185
               End
            End
            Begin VB.Frame Frame6 
               BackColor       =   &H00E2E9E9&
               Height          =   735
               Left            =   7200
               RightToLeft     =   -1  'True
               TabIndex        =   97
               Top             =   3600
               Width           =   6075
               Begin VB.TextBox EvaIDToTxt 
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
               Begin VB.TextBox EvaIDFromTxt 
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
                  Left            =   2400
                  RightToLeft     =   -1  'True
                  TabIndex        =   98
                  Top             =   240
                  Width           =   1155
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·„”·”·"
                  Height          =   195
                  Index           =   34
                  Left            =   4440
                  RightToLeft     =   -1  'True
                  TabIndex        =   102
                  Top             =   240
                  Width           =   1185
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "≈·Ï"
                  Height          =   315
                  Index           =   33
                  Left            =   1560
                  RightToLeft     =   -1  'True
                  TabIndex        =   101
                  Top             =   240
                  Width           =   645
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„‰"
                  Height          =   315
                  Index           =   32
                  Left            =   3720
                  RightToLeft     =   -1  'True
                  TabIndex        =   100
                  Top             =   240
                  Width           =   660
               End
            End
            Begin VB.Frame Frame5 
               BackColor       =   &H00E2E9E9&
               Height          =   735
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   91
               Top             =   3600
               Width           =   7095
               Begin MSComCtl2.DTPicker EvaDateFrom 
                  Height          =   330
                  Left            =   3120
                  TabIndex        =   92
                  Top             =   240
                  Width           =   1575
                  _ExtentX        =   2778
                  _ExtentY        =   582
                  _Version        =   393216
                  CheckBox        =   -1  'True
                  Format          =   237371395
                  CurrentDate     =   38887
               End
               Begin MSComCtl2.DTPicker EvaDateTo 
                  Height          =   330
                  Left            =   600
                  TabIndex        =   93
                  Top             =   240
                  Width           =   1575
                  _ExtentX        =   2778
                  _ExtentY        =   582
                  _Version        =   393216
                  CheckBox        =   -1  'True
                  Format          =   237371395
                  CurrentDate     =   38887
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«· «—ÌŒ"
                  Height          =   195
                  Index           =   31
                  Left            =   5640
                  RightToLeft     =   -1  'True
                  TabIndex        =   96
                  Top             =   240
                  Width           =   1185
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„‰"
                  Height          =   315
                  Index           =   30
                  Left            =   4800
                  RightToLeft     =   -1  'True
                  TabIndex        =   95
                  Top             =   240
                  Width           =   660
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "≈·Ï"
                  Height          =   315
                  Index           =   29
                  Left            =   1920
                  RightToLeft     =   -1  'True
                  TabIndex        =   94
                  Top             =   240
                  Width           =   1080
               End
            End
            Begin VB.Frame Frame3 
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   780
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   89
               Top             =   0
               Width           =   13425
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·»ÕÀ „⁄«ÌÌ— «· ﬁÌÌ„"
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
                  Index           =   5
                  Left            =   8160
                  RightToLeft     =   -1  'True
                  TabIndex        =   90
                  Top             =   240
                  Width           =   3720
               End
               Begin VB.Image Image4 
                  Height          =   615
                  Left            =   12360
                  Picture         =   "FrmInsurancesSearch.frx":4F122
                  Stretch         =   -1  'True
                  Top             =   120
                  Width           =   735
               End
            End
            Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid2 
               Height          =   2745
               Left            =   120
               TabIndex        =   110
               Top             =   840
               Width           =   13275
               _cx             =   23416
               _cy             =   4842
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
               BackColor       =   16777215
               ForeColor       =   -2147483640
               BackColorFixed  =   14871017
               ForeColorFixed  =   -2147483630
               BackColorSel    =   -2147483635
               ForeColorSel    =   -2147483634
               BackColorBkg    =   -2147483633
               BackColorAlternate=   16777088
               GridColor       =   -2147483633
               GridColorFixed  =   -2147483633
               TreeColor       =   -2147483632
               FloodColor      =   192
               SheetBorder     =   -2147483633
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
               FormatString    =   $"FrmInsurancesSearch.frx":508CF
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
         Begin C1SizerLibCtl.C1Elastic C1Elastic5 
            Height          =   6510
            Left            =   -14640
            TabIndex        =   111
            TabStop         =   0   'False
            Top             =   45
            Width           =   13395
            _cx             =   23627
            _cy             =   11483
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
            Begin VB.CheckBox DetSearch 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "»ÕÀ  ›’Ì·Ì "
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   3000
               RightToLeft     =   -1  'True
               TabIndex        =   135
               Top             =   4560
               Width           =   1935
            End
            Begin VB.Frame Frame16 
               BackColor       =   &H00E2E9E9&
               Enabled         =   0   'False
               Height          =   975
               Left            =   330
               RightToLeft     =   -1  'True
               TabIndex        =   126
               Top             =   5160
               Visible         =   0   'False
               Width           =   12825
               Begin VB.TextBox TotalMarkTxt 
                  Alignment       =   1  'Right Justify
                  Height          =   330
                  Left            =   720
                  RightToLeft     =   -1  'True
                  TabIndex        =   127
                  Top             =   360
                  Width           =   4185
               End
               Begin MSDataListLib.DataCombo EmployeeDC 
                  Bindings        =   "FrmInsurancesSearch.frx":50A0D
                  Height          =   315
                  Left            =   7080
                  TabIndex        =   136
                  Top             =   360
                  Width           =   4185
                  _ExtentX        =   7382
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
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·„ÊŸ›"
                  Height          =   195
                  Index           =   50
                  Left            =   11400
                  RightToLeft     =   -1  'True
                  TabIndex        =   129
                  Top             =   360
                  Width           =   1185
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«Ã„«·Ì «·œ—Ã« "
                  Height          =   195
                  Index           =   49
                  Left            =   5280
                  RightToLeft     =   -1  'True
                  TabIndex        =   128
                  Top             =   360
                  Width           =   1185
               End
            End
            Begin VB.Frame Frame15 
               BackColor       =   &H00E2E9E9&
               Height          =   735
               Left            =   7200
               RightToLeft     =   -1  'True
               TabIndex        =   120
               Top             =   3600
               Width           =   6195
               Begin VB.TextBox EvaClamIDToTxt 
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
                  TabIndex        =   122
                  Top             =   240
                  Width           =   1155
               End
               Begin VB.TextBox EvaClamIDFromTxt 
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
                  Left            =   2400
                  RightToLeft     =   -1  'True
                  TabIndex        =   121
                  Top             =   240
                  Width           =   1155
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "—ﬁ„ «·Õ—ﬂ…"
                  Height          =   195
                  Index           =   45
                  Left            =   4440
                  RightToLeft     =   -1  'True
                  TabIndex        =   125
                  Top             =   240
                  Width           =   1185
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "≈·Ï"
                  Height          =   315
                  Index           =   44
                  Left            =   1560
                  RightToLeft     =   -1  'True
                  TabIndex        =   124
                  Top             =   240
                  Width           =   645
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„‰"
                  Height          =   315
                  Index           =   43
                  Left            =   3720
                  RightToLeft     =   -1  'True
                  TabIndex        =   123
                  Top             =   240
                  Width           =   660
               End
            End
            Begin VB.Frame Frame14 
               BackColor       =   &H00E2E9E9&
               Height          =   735
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   114
               Top             =   3600
               Width           =   7095
               Begin MSComCtl2.DTPicker EvaClamDateFrom 
                  Height          =   330
                  Left            =   3000
                  TabIndex        =   115
                  Top             =   240
                  Width           =   1575
                  _ExtentX        =   2778
                  _ExtentY        =   582
                  _Version        =   393216
                  CheckBox        =   -1  'True
                  Format          =   237371395
                  CurrentDate     =   38887
               End
               Begin MSComCtl2.DTPicker EvaClamDateTo 
                  Height          =   330
                  Left            =   600
                  TabIndex        =   116
                  Top             =   240
                  Width           =   1575
                  _ExtentX        =   2778
                  _ExtentY        =   582
                  _Version        =   393216
                  CheckBox        =   -1  'True
                  Format          =   237371395
                  CurrentDate     =   38887
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " «—ÌŒ «·Õ—ﬂ…"
                  Height          =   195
                  Index           =   42
                  Left            =   5640
                  RightToLeft     =   -1  'True
                  TabIndex        =   119
                  Top             =   240
                  Width           =   1185
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„‰"
                  Height          =   315
                  Index           =   41
                  Left            =   4800
                  RightToLeft     =   -1  'True
                  TabIndex        =   118
                  Top             =   240
                  Width           =   660
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "≈·Ï"
                  Height          =   315
                  Index           =   40
                  Left            =   1920
                  RightToLeft     =   -1  'True
                  TabIndex        =   117
                  Top             =   240
                  Width           =   1080
               End
            End
            Begin VB.Frame Frame10 
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   780
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   112
               Top             =   0
               Width           =   13425
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "»ÕÀ «” Õﬁ«ﬁ «· ﬁÌ„ "
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
                  Index           =   6
                  Left            =   5640
                  RightToLeft     =   -1  'True
                  TabIndex        =   113
                  Top             =   240
                  Width           =   6240
               End
               Begin VB.Image Image5 
                  Height          =   615
                  Left            =   12360
                  Picture         =   "FrmInsurancesSearch.frx":50A22
                  Stretch         =   -1  'True
                  Top             =   120
                  Width           =   735
               End
            End
            Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid3 
               Height          =   2745
               Left            =   120
               TabIndex        =   130
               Top             =   840
               Width           =   13275
               _cx             =   23416
               _cy             =   4842
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
               BackColor       =   16777215
               ForeColor       =   -2147483640
               BackColorFixed  =   14871017
               ForeColorFixed  =   -2147483630
               BackColorSel    =   -2147483635
               ForeColorSel    =   -2147483634
               BackColorBkg    =   -2147483633
               BackColorAlternate=   16777088
               GridColor       =   -2147483633
               GridColorFixed  =   -2147483633
               TreeColor       =   -2147483632
               FloodColor      =   192
               SheetBorder     =   -2147483633
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
               FormatString    =   $"FrmInsurancesSearch.frx":521CF
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
            Begin MSDataListLib.DataCombo BranchDC 
               Bindings        =   "FrmInsurancesSearch.frx":522BF
               Height          =   315
               Left            =   6000
               TabIndex        =   133
               Top             =   4560
               Width           =   5745
               _ExtentX        =   10134
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
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·›—⁄"
               Height          =   195
               Index           =   48
               Left            =   11880
               RightToLeft     =   -1  'True
               TabIndex        =   134
               Top             =   4560
               Width           =   1185
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic6 
            Height          =   6510
            Left            =   -14340
            TabIndex        =   137
            TabStop         =   0   'False
            Top             =   45
            Width           =   13395
            _cx             =   23627
            _cy             =   11483
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
            Begin VB.Frame Frame21 
               BackColor       =   &H00E2E9E9&
               Height          =   735
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   158
               Top             =   3600
               Width           =   3825
               Begin MSDataListLib.DataCombo BranchDCBank 
                  Bindings        =   "FrmInsurancesSearch.frx":522D4
                  Height          =   315
                  Left            =   120
                  TabIndex        =   159
                  Top             =   240
                  Width           =   2865
                  _ExtentX        =   5054
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
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·›—⁄"
                  Height          =   195
                  Index           =   55
                  Left            =   3000
                  RightToLeft     =   -1  'True
                  TabIndex        =   160
                  Top             =   240
                  Width           =   705
               End
            End
            Begin VB.Frame Frame20 
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   780
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   155
               Top             =   0
               Width           =   13425
               Begin VB.Image Image6 
                  Height          =   615
                  Left            =   12360
                  Picture         =   "FrmInsurancesSearch.frx":522E9
                  Stretch         =   -1  'True
                  Top             =   120
                  Width           =   735
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
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
                  Index           =   7
                  Left            =   5640
                  RightToLeft     =   -1  'True
                  TabIndex        =   156
                  Top             =   240
                  Width           =   6240
               End
            End
            Begin VB.Frame Frame19 
               BackColor       =   &H00E2E9E9&
               Height          =   735
               Left            =   3960
               RightToLeft     =   -1  'True
               TabIndex        =   149
               Top             =   3600
               Width           =   5415
               Begin MSComCtl2.DTPicker TransFrom 
                  Height          =   330
                  Left            =   2160
                  TabIndex        =   150
                  Top             =   240
                  Width           =   1575
                  _ExtentX        =   2778
                  _ExtentY        =   582
                  _Version        =   393216
                  CheckBox        =   -1  'True
                  Format          =   191299587
                  CurrentDate     =   38887
               End
               Begin MSComCtl2.DTPicker TransTo 
                  Height          =   330
                  Left            =   120
                  TabIndex        =   151
                  Top             =   240
                  Width           =   1575
                  _ExtentX        =   2778
                  _ExtentY        =   582
                  _Version        =   393216
                  CheckBox        =   -1  'True
                  Format          =   191299587
                  CurrentDate     =   38887
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "≈·Ï"
                  Height          =   315
                  Index           =   54
                  Left            =   1680
                  RightToLeft     =   -1  'True
                  TabIndex        =   154
                  Top             =   240
                  Width           =   480
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„‰"
                  Height          =   315
                  Index           =   53
                  Left            =   3720
                  RightToLeft     =   -1  'True
                  TabIndex        =   153
                  Top             =   240
                  Width           =   540
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " «—ÌŒ «·Õ—ﬂ…"
                  Height          =   195
                  Index           =   52
                  Left            =   4200
                  RightToLeft     =   -1  'True
                  TabIndex        =   152
                  Top             =   240
                  Width           =   1065
               End
            End
            Begin VB.Frame Frame18 
               BackColor       =   &H00E2E9E9&
               Height          =   735
               Left            =   9360
               RightToLeft     =   -1  'True
               TabIndex        =   143
               Top             =   3600
               Width           =   3915
               Begin VB.TextBox BankIDFrom 
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
                  Height          =   405
                  Left            =   1440
                  RightToLeft     =   -1  'True
                  TabIndex        =   145
                  Top             =   240
                  Width           =   795
               End
               Begin VB.TextBox BankIDTo 
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
                  Height          =   405
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   144
                  Top             =   240
                  Width           =   795
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„‰"
                  Height          =   315
                  Index           =   51
                  Left            =   2280
                  RightToLeft     =   -1  'True
                  TabIndex        =   148
                  Top             =   240
                  Width           =   660
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "≈·Ï"
                  Height          =   315
                  Index           =   47
                  Left            =   840
                  RightToLeft     =   -1  'True
                  TabIndex        =   147
                  Top             =   240
                  Width           =   645
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "—ﬁ„ «·Õ—ﬂ…"
                  Height          =   195
                  Index           =   46
                  Left            =   2880
                  RightToLeft     =   -1  'True
                  TabIndex        =   146
                  Top             =   240
                  Width           =   945
               End
            End
            Begin VB.Frame Frame17 
               BackColor       =   &H00E2E9E9&
               Height          =   2055
               Left            =   90
               RightToLeft     =   -1  'True
               TabIndex        =   138
               Top             =   4320
               Width           =   13185
               Begin VB.TextBox Code1 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   10680
                  RightToLeft     =   -1  'True
                  TabIndex        =   168
                  Top             =   1320
                  Width           =   1185
               End
               Begin VB.TextBox beneficiaryTxt 
                  Alignment       =   1  'Right Justify
                  Height          =   330
                  Left            =   480
                  RightToLeft     =   -1  'True
                  TabIndex        =   161
                  Top             =   360
                  Width           =   4425
               End
               Begin VB.TextBox NumberTxt 
                  Alignment       =   1  'Right Justify
                  Height          =   330
                  Left            =   7080
                  RightToLeft     =   -1  'True
                  TabIndex        =   139
                  Top             =   360
                  Width           =   4785
               End
               Begin MSDataListLib.DataCombo applicantDC 
                  Bindings        =   "FrmInsurancesSearch.frx":53A96
                  Height          =   315
                  Left            =   7080
                  TabIndex        =   140
                  Top             =   1320
                  Width           =   3585
                  _ExtentX        =   6324
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
               Begin MSComCtl2.DTPicker ReqDateFrom 
                  Height          =   330
                  Left            =   2730
                  TabIndex        =   163
                  Top             =   1320
                  Width           =   1575
                  _ExtentX        =   2778
                  _ExtentY        =   582
                  _Version        =   393216
                  CheckBox        =   -1  'True
                  Format          =   191299587
                  CurrentDate     =   38887
               End
               Begin MSComCtl2.DTPicker ReqDateTo 
                  Height          =   330
                  Left            =   480
                  TabIndex        =   164
                  Top             =   1320
                  Width           =   1575
                  _ExtentX        =   2778
                  _ExtentY        =   582
                  _Version        =   393216
                  CheckBox        =   -1  'True
                  Format          =   191299587
                  CurrentDate     =   38887
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "≈·Ï"
                  Height          =   315
                  Index           =   59
                  Left            =   2040
                  RightToLeft     =   -1  'True
                  TabIndex        =   167
                  Top             =   1320
                  Width           =   600
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„‰"
                  Height          =   315
                  Index           =   58
                  Left            =   4365
                  RightToLeft     =   -1  'True
                  TabIndex        =   166
                  Top             =   1320
                  Width           =   540
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " «—ÌŒ «·ÿ·»"
                  Height          =   195
                  Index           =   57
                  Left            =   5160
                  RightToLeft     =   -1  'True
                  TabIndex        =   165
                  Top             =   1320
                  Width           =   1185
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«”„ «·„” ›Ìœ"
                  Height          =   195
                  Index           =   56
                  Left            =   5160
                  RightToLeft     =   -1  'True
                  TabIndex        =   162
                  Top             =   360
                  Width           =   1185
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·—ﬁ„"
                  Height          =   195
                  Index           =   37
                  Left            =   11880
                  RightToLeft     =   -1  'True
                  TabIndex        =   142
                  Top             =   360
                  Width           =   1185
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„ﬁœ„ «·ÿ·»"
                  Height          =   195
                  Index           =   36
                  Left            =   11880
                  RightToLeft     =   -1  'True
                  TabIndex        =   141
                  Top             =   1320
                  Width           =   1185
               End
            End
            Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid4 
               Height          =   2745
               Left            =   120
               TabIndex        =   157
               Top             =   840
               Width           =   13275
               _cx             =   23416
               _cy             =   4842
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
               BackColor       =   16777215
               ForeColor       =   -2147483640
               BackColorFixed  =   14871017
               ForeColorFixed  =   -2147483630
               BackColorSel    =   -2147483635
               ForeColorSel    =   -2147483634
               BackColorBkg    =   -2147483633
               BackColorAlternate=   16777088
               GridColor       =   -2147483633
               GridColorFixed  =   -2147483633
               TreeColor       =   -2147483632
               FloodColor      =   192
               SheetBorder     =   -2147483633
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
               FormatString    =   $"FrmInsurancesSearch.frx":53AAB
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
         Begin C1SizerLibCtl.C1Elastic C1Elastic8 
            Height          =   6510
            Left            =   -14040
            TabIndex        =   169
            TabStop         =   0   'False
            Top             =   45
            Width           =   13395
            _cx             =   23627
            _cy             =   11483
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
            Begin VB.Frame Frame26 
               BackColor       =   &H00E2E9E9&
               Height          =   2055
               Left            =   90
               RightToLeft     =   -1  'True
               TabIndex        =   187
               Top             =   4320
               Width           =   13185
               Begin VB.ComboBox PayTypeCB 
                  Height          =   315
                  ItemData        =   "FrmInsurancesSearch.frx":53BE5
                  Left            =   7080
                  List            =   "FrmInsurancesSearch.frx":53BE7
                  RightToLeft     =   -1  'True
                  TabIndex        =   222
                  Top             =   1320
                  Width           =   4785
               End
               Begin VB.TextBox CompName 
                  Alignment       =   1  'Right Justify
                  Height          =   330
                  Left            =   480
                  RightToLeft     =   -1  'True
                  TabIndex        =   221
                  Top             =   240
                  Width           =   4545
               End
               Begin VB.TextBox CompNoTxt 
                  Alignment       =   1  'Right Justify
                  Height          =   330
                  Left            =   480
                  RightToLeft     =   -1  'True
                  TabIndex        =   219
                  Top             =   600
                  Width           =   4545
               End
               Begin VB.TextBox CopyPriceTxt 
                  Alignment       =   1  'Right Justify
                  Height          =   330
                  Left            =   7080
                  RightToLeft     =   -1  'True
                  TabIndex        =   212
                  Top             =   960
                  Width           =   4785
               End
               Begin VB.OptionButton SuppRd 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„Ê—œ"
                  Height          =   195
                  Left            =   10920
                  RightToLeft     =   -1  'True
                  TabIndex        =   209
                  Top             =   600
                  Width           =   945
               End
               Begin VB.OptionButton OtherRd 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "√Œ—Ï"
                  Height          =   195
                  Left            =   10080
                  RightToLeft     =   -1  'True
                  TabIndex        =   208
                  Top             =   600
                  Width           =   705
               End
               Begin VB.TextBox OtherNameTxt 
                  Alignment       =   1  'Right Justify
                  Enabled         =   0   'False
                  Height          =   330
                  Left            =   7080
                  RightToLeft     =   -1  'True
                  TabIndex        =   189
                  Top             =   600
                  Visible         =   0   'False
                  Width           =   2865
               End
               Begin VB.TextBox Text3 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   10680
                  RightToLeft     =   -1  'True
                  TabIndex        =   188
                  Top             =   1680
                  Width           =   1185
               End
               Begin MSDataListLib.DataCombo ResEmpDC 
                  Bindings        =   "FrmInsurancesSearch.frx":53BE9
                  Height          =   315
                  Left            =   7080
                  TabIndex        =   190
                  Top             =   1680
                  Width           =   3585
                  _ExtentX        =   6324
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
               Begin MSComCtl2.DTPicker OpenDateFrom 
                  Height          =   330
                  Left            =   2730
                  TabIndex        =   191
                  Top             =   1680
                  Width           =   1575
                  _ExtentX        =   2778
                  _ExtentY        =   582
                  _Version        =   393216
                  CheckBox        =   -1  'True
                  Format          =   191299587
                  CurrentDate     =   38887
               End
               Begin MSComCtl2.DTPicker OpenDateTo 
                  Height          =   330
                  Left            =   480
                  TabIndex        =   192
                  Top             =   1680
                  Width           =   1575
                  _ExtentX        =   2778
                  _ExtentY        =   582
                  _Version        =   393216
                  CheckBox        =   -1  'True
                  Format          =   191299587
                  CurrentDate     =   38887
               End
               Begin MSDataListLib.DataCombo DepDC 
                  Bindings        =   "FrmInsurancesSearch.frx":53BFE
                  Height          =   315
                  Left            =   7080
                  TabIndex        =   200
                  Top             =   240
                  Width           =   4785
                  _ExtentX        =   8440
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
               Begin MSComCtl2.DTPicker SubDateFrom 
                  Height          =   330
                  Left            =   2730
                  TabIndex        =   206
                  Top             =   1320
                  Width           =   1575
                  _ExtentX        =   2778
                  _ExtentY        =   582
                  _Version        =   393216
                  CheckBox        =   -1  'True
                  Format          =   191299587
                  CurrentDate     =   38887
               End
               Begin MSComCtl2.DTPicker SubDateTo 
                  Height          =   330
                  Left            =   480
                  TabIndex        =   207
                  Top             =   1320
                  Width           =   1575
                  _ExtentX        =   2778
                  _ExtentY        =   582
                  _Version        =   393216
                  CheckBox        =   -1  'True
                  Format          =   191299587
                  CurrentDate     =   38887
               End
               Begin MSDataListLib.DataCombo SuppDC 
                  Bindings        =   "FrmInsurancesSearch.frx":53C13
                  Height          =   315
                  Left            =   7080
                  TabIndex        =   210
                  Top             =   600
                  Visible         =   0   'False
                  Width           =   2865
                  _ExtentX        =   5054
                  _ExtentY        =   556
                  _Version        =   393216
                  Enabled         =   0   'False
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
               Begin MSComCtl2.DTPicker CompDateFrom 
                  Height          =   330
                  Left            =   2730
                  TabIndex        =   216
                  Top             =   960
                  Width           =   1575
                  _ExtentX        =   2778
                  _ExtentY        =   582
                  _Version        =   393216
                  CheckBox        =   -1  'True
                  Format          =   191299587
                  CurrentDate     =   38887
               End
               Begin MSComCtl2.DTPicker CompDateTo 
                  Height          =   330
                  Left            =   480
                  TabIndex        =   217
                  Top             =   960
                  Width           =   1575
                  _ExtentX        =   2778
                  _ExtentY        =   582
                  _Version        =   393216
                  CheckBox        =   -1  'True
                  Format          =   191299587
                  CurrentDate     =   38887
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«”„ «·„‰«›”…"
                  Height          =   210
                  Index           =   83
                  Left            =   5040
                  RightToLeft     =   -1  'True
                  TabIndex        =   220
                  Top             =   240
                  Width           =   1545
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "—ﬁ„ «·„‰«›”…"
                  Height          =   210
                  Index           =   82
                  Left            =   5040
                  RightToLeft     =   -1  'True
                  TabIndex        =   218
                  Top             =   600
                  Width           =   1545
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " «—ÌŒ «·„‰«›”…"
                  Height          =   195
                  Index           =   81
                  Left            =   5040
                  RightToLeft     =   -1  'True
                  TabIndex        =   215
                  Top             =   960
                  Width           =   1545
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„‰"
                  Height          =   195
                  Index           =   80
                  Left            =   4365
                  RightToLeft     =   -1  'True
                  TabIndex        =   214
                  Top             =   960
                  Width           =   540
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "≈·Ï"
                  Height          =   315
                  Index           =   79
                  Left            =   2040
                  RightToLeft     =   -1  'True
                  TabIndex        =   213
                  Top             =   960
                  Width           =   600
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "”⁄— «·‰”Œ…"
                  Height          =   330
                  Index           =   78
                  Left            =   11880
                  RightToLeft     =   -1  'True
                  TabIndex        =   211
                  Top             =   960
                  Width           =   1185
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " «—ÌŒ «· ﬁœÌ„"
                  Height          =   195
                  Index           =   77
                  Left            =   5040
                  RightToLeft     =   -1  'True
                  TabIndex        =   205
                  Top             =   1320
                  Width           =   1545
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„‰"
                  Height          =   195
                  Index           =   76
                  Left            =   4365
                  RightToLeft     =   -1  'True
                  TabIndex        =   204
                  Top             =   1320
                  Width           =   540
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "≈·Ï"
                  Height          =   315
                  Index           =   75
                  Left            =   2040
                  RightToLeft     =   -1  'True
                  TabIndex        =   203
                  Top             =   1320
                  Width           =   600
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬁ”„"
                  Height          =   195
                  Index           =   73
                  Left            =   11880
                  RightToLeft     =   -1  'True
                  TabIndex        =   201
                  Top             =   240
                  Width           =   1185
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·„ÊŸ› «·„”ƒÊ·"
                  Height          =   195
                  Index           =   72
                  Left            =   11880
                  RightToLeft     =   -1  'True
                  TabIndex        =   198
                  Top             =   1680
                  Width           =   1185
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ÿ—Ìﬁ… «·œ›⁄"
                  Height          =   195
                  Index           =   71
                  Left            =   11880
                  RightToLeft     =   -1  'True
                  TabIndex        =   197
                  Top             =   1320
                  Width           =   1185
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ÃÂ… «·ÿ«·»… "
                  Height          =   195
                  Index           =   70
                  Left            =   11880
                  RightToLeft     =   -1  'True
                  TabIndex        =   196
                  Top             =   600
                  Width           =   1185
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " «—ÌŒ › Õ «·„Ÿ«—Ì›"
                  Height          =   195
                  Index           =   69
                  Left            =   5040
                  RightToLeft     =   -1  'True
                  TabIndex        =   195
                  Top             =   1680
                  Width           =   1545
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„‰"
                  Height          =   195
                  Index           =   68
                  Left            =   4365
                  RightToLeft     =   -1  'True
                  TabIndex        =   194
                  Top             =   1680
                  Width           =   540
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "≈·Ï"
                  Height          =   315
                  Index           =   67
                  Left            =   2040
                  RightToLeft     =   -1  'True
                  TabIndex        =   193
                  Top             =   1680
                  Width           =   600
               End
            End
            Begin VB.Frame Frame25 
               BackColor       =   &H00E2E9E9&
               Height          =   735
               Left            =   9360
               RightToLeft     =   -1  'True
               TabIndex        =   181
               Top             =   3600
               Width           =   3915
               Begin VB.TextBox IDDTo 
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
                  Height          =   405
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   183
                  Top             =   240
                  Width           =   795
               End
               Begin VB.TextBox IDDFrom 
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
                  Height          =   405
                  Left            =   1440
                  RightToLeft     =   -1  'True
                  TabIndex        =   182
                  Top             =   240
                  Width           =   795
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "—ﬁ„ «·Õ—ﬂ…"
                  Height          =   195
                  Index           =   66
                  Left            =   2880
                  RightToLeft     =   -1  'True
                  TabIndex        =   186
                  Top             =   240
                  Width           =   945
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "≈·Ï"
                  Height          =   315
                  Index           =   65
                  Left            =   840
                  RightToLeft     =   -1  'True
                  TabIndex        =   185
                  Top             =   240
                  Width           =   645
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„‰"
                  Height          =   315
                  Index           =   64
                  Left            =   2280
                  RightToLeft     =   -1  'True
                  TabIndex        =   184
                  Top             =   240
                  Width           =   660
               End
            End
            Begin VB.Frame Frame24 
               BackColor       =   &H00E2E9E9&
               Height          =   735
               Left            =   3960
               RightToLeft     =   -1  'True
               TabIndex        =   175
               Top             =   3600
               Width           =   5415
               Begin MSComCtl2.DTPicker Date2From 
                  Height          =   330
                  Left            =   2160
                  TabIndex        =   176
                  Top             =   240
                  Width           =   1575
                  _ExtentX        =   2778
                  _ExtentY        =   582
                  _Version        =   393216
                  CheckBox        =   -1  'True
                  Format          =   191299587
                  CurrentDate     =   38887
               End
               Begin MSComCtl2.DTPicker Date2To 
                  Height          =   330
                  Left            =   120
                  TabIndex        =   177
                  Top             =   240
                  Width           =   1575
                  _ExtentX        =   2778
                  _ExtentY        =   582
                  _Version        =   393216
                  CheckBox        =   -1  'True
                  Format          =   235929603
                  CurrentDate     =   38887
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " «—ÌŒ «·Õ—ﬂ…"
                  Height          =   195
                  Index           =   63
                  Left            =   4200
                  RightToLeft     =   -1  'True
                  TabIndex        =   180
                  Top             =   240
                  Width           =   1065
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„‰"
                  Height          =   315
                  Index           =   62
                  Left            =   3720
                  RightToLeft     =   -1  'True
                  TabIndex        =   179
                  Top             =   240
                  Width           =   540
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "≈·Ï"
                  Height          =   315
                  Index           =   61
                  Left            =   1680
                  RightToLeft     =   -1  'True
                  TabIndex        =   178
                  Top             =   240
                  Width           =   480
               End
            End
            Begin VB.Frame Frame23 
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   780
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   173
               Top             =   0
               Width           =   13425
               Begin VB.Label Label1 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "»ÕÀ ‰„«–Ã ‘—«¡ „‰«›”…"
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
                  Index           =   8
                  Left            =   8040
                  RightToLeft     =   -1  'True
                  TabIndex        =   174
                  Top             =   240
                  Width           =   4080
               End
               Begin VB.Image Image7 
                  Height          =   615
                  Left            =   12360
                  Picture         =   "FrmInsurancesSearch.frx":53C28
                  Stretch         =   -1  'True
                  Top             =   120
                  Width           =   735
               End
            End
            Begin VB.Frame Frame22 
               BackColor       =   &H00E2E9E9&
               Height          =   735
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   170
               Top             =   3600
               Width           =   3825
               Begin MSDataListLib.DataCombo BranchDC2 
                  Bindings        =   "FrmInsurancesSearch.frx":553D5
                  Height          =   315
                  Left            =   120
                  TabIndex        =   171
                  Top             =   240
                  Width           =   2865
                  _ExtentX        =   5054
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
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·›—⁄"
                  Height          =   195
                  Index           =   60
                  Left            =   3000
                  RightToLeft     =   -1  'True
                  TabIndex        =   172
                  Top             =   240
                  Width           =   705
               End
            End
            Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid5 
               Height          =   2745
               Left            =   120
               TabIndex        =   199
               Top             =   840
               Width           =   13275
               _cx             =   23416
               _cy             =   4842
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
               BackColor       =   16777215
               ForeColor       =   -2147483640
               BackColorFixed  =   14871017
               ForeColorFixed  =   -2147483630
               BackColorSel    =   -2147483635
               ForeColorSel    =   -2147483634
               BackColorBkg    =   -2147483633
               BackColorAlternate=   16777088
               GridColor       =   -2147483633
               GridColorFixed  =   -2147483633
               TreeColor       =   -2147483632
               FloodColor      =   192
               SheetBorder     =   -2147483633
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
               FormatString    =   $"FrmInsurancesSearch.frx":553EA
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
         Begin C1SizerLibCtl.C1Elastic C1Elastic9 
            Height          =   6510
            Left            =   45
            TabIndex        =   223
            TabStop         =   0   'False
            Top             =   45
            Width           =   13395
            _cx             =   23627
            _cy             =   11483
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
            Begin VB.Frame Frame30 
               BackColor       =   &H00E2E9E9&
               Height          =   2010
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   238
               Top             =   4440
               Width           =   13020
               Begin VB.TextBox txtContainerNo 
                  BackColor       =   &H0000FFFF&
                  Height          =   345
                  Left            =   300
                  TabIndex        =   256
                  Top             =   210
                  Width           =   1725
               End
               Begin VB.TextBox TxtOrderNo 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   3330
                  RightToLeft     =   -1  'True
                  TabIndex        =   254
                  Top             =   240
                  Width           =   1890
               End
               Begin MSDataListLib.DataCombo DBCboClientName1 
                  Height          =   315
                  Left            =   6720
                  TabIndex        =   239
                  Top             =   240
                  Width           =   4395
                  _ExtentX        =   7752
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin C1SizerLibCtl.C1Elastic C1Elastic12 
                  Height          =   510
                  Left            =   240
                  TabIndex        =   242
                  TabStop         =   0   'False
                  Top             =   720
                  Width           =   12645
                  _cx             =   22304
                  _cy             =   900
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
                  Begin VB.TextBox TxtLeaderName 
                     Alignment       =   1  'Right Justify
                     Height          =   300
                     Left            =   150
                     RightToLeft     =   -1  'True
                     TabIndex        =   244
                     Top             =   120
                     Width           =   4860
                  End
                  Begin VB.TextBox Text6 
                     Alignment       =   1  'Right Justify
                     Height          =   300
                     Left            =   10230
                     RightToLeft     =   -1  'True
                     TabIndex        =   243
                     Top             =   120
                     Width           =   660
                  End
                  Begin XtremeSuiteControls.RadioButton ChDrievType 
                     Height          =   255
                     Index           =   0
                     Left            =   10530
                     TabIndex        =   245
                     Top             =   120
                     Width           =   1215
                     _Version        =   786432
                     _ExtentX        =   2143
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "œ«Œ·Ì"
                     BackColor       =   14871017
                     UseVisualStyle  =   -1  'True
                     TextAlignment   =   1
                     RightToLeft     =   -1  'True
                  End
                  Begin XtremeSuiteControls.RadioButton ChDrievType 
                     Height          =   255
                     Index           =   1
                     Left            =   4755
                     TabIndex        =   246
                     Top             =   135
                     Width           =   1590
                     _Version        =   786432
                     _ExtentX        =   2805
                     _ExtentY        =   450
                     _StockProps     =   79
                     Caption         =   "Œ«—ÃÌ"
                     BackColor       =   14871017
                     UseVisualStyle  =   -1  'True
                     TextAlignment   =   1
                     RightToLeft     =   -1  'True
                  End
                  Begin MSDataListLib.DataCombo DcEmployee 
                     Height          =   315
                     Left            =   6480
                     TabIndex        =   247
                     Top             =   120
                     Width           =   3720
                     _ExtentX        =   6562
                     _ExtentY        =   556
                     _Version        =   393216
                     BackColor       =   16777215
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«”„ «·”«∆ﬁ"
                     Height          =   270
                     Index           =   92
                     Left            =   11070
                     RightToLeft     =   -1  'True
                     TabIndex        =   248
                     Top             =   0
                     Width           =   1500
                  End
               End
               Begin C1SizerLibCtl.C1Elastic C1Elastic10 
                  Height          =   510
                  Left            =   240
                  TabIndex        =   249
                  TabStop         =   0   'False
                  Top             =   1320
                  Width           =   12645
                  _cx             =   22304
                  _cy             =   900
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
                  Begin MSDataListLib.DataCombo DcbCar 
                     Height          =   315
                     Left            =   6480
                     TabIndex        =   250
                     Top             =   120
                     Width           =   4440
                     _ExtentX        =   7832
                     _ExtentY        =   556
                     _Version        =   393216
                     BackColor       =   16777215
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin MSDataListLib.DataCombo DcbCar2 
                     Height          =   315
                     Left            =   120
                     TabIndex        =   253
                     Top             =   120
                     Width           =   4860
                     _ExtentX        =   8573
                     _ExtentY        =   556
                     _Version        =   393216
                     BackColor       =   16777215
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin VB.Label Label4 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "”Ì«—… €Ì— „„·Êﬂ…"
                     Height          =   285
                     Left            =   5160
                     TabIndex        =   252
                     Top             =   120
                     Width           =   1245
                  End
                  Begin VB.Label Label3 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "”Ì«—… „„·Êﬂ…"
                     Height          =   285
                     Left            =   11280
                     TabIndex        =   251
                     Top             =   120
                     Width           =   1245
                  End
               End
               Begin VB.Label Label14 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "—ﬁ„ «—«„ﬂÊ"
                  ForeColor       =   &H00000000&
                  Height          =   270
                  Index           =   0
                  Left            =   1710
                  TabIndex        =   257
                  Top             =   300
                  Width           =   1590
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "—ﬁ„ «„— «· Õ„Ì· „‰ «·⁄„Ì·"
                  Height          =   390
                  Index           =   90
                  Left            =   5160
                  RightToLeft     =   -1  'True
                  TabIndex        =   255
                  Top             =   240
                  Width           =   1500
               End
               Begin VB.Label Label2 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«”„ «·⁄„Ì·"
                  Height          =   285
                  Left            =   11760
                  TabIndex        =   240
                  Top             =   240
                  Width           =   1245
               End
            End
            Begin VB.Frame Frame29 
               BackColor       =   &H00E2E9E9&
               Height          =   690
               Left            =   7320
               RightToLeft     =   -1  'True
               TabIndex        =   232
               Top             =   3780
               Width           =   5835
               Begin VB.TextBox Text2 
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
                  Left            =   360
                  RightToLeft     =   -1  'True
                  TabIndex        =   234
                  Top             =   240
                  Width           =   1155
               End
               Begin VB.TextBox Text1 
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
                  Left            =   2760
                  RightToLeft     =   -1  'True
                  TabIndex        =   233
                  Top             =   240
                  Width           =   1155
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "—ﬁ„ «·«„—"
                  Height          =   195
                  Index           =   89
                  Left            =   4440
                  RightToLeft     =   -1  'True
                  TabIndex        =   237
                  Top             =   240
                  Width           =   1185
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "≈·Ï"
                  Height          =   315
                  Index           =   88
                  Left            =   1920
                  RightToLeft     =   -1  'True
                  TabIndex        =   236
                  Top             =   240
                  Width           =   645
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„‰"
                  Height          =   315
                  Index           =   87
                  Left            =   3960
                  RightToLeft     =   -1  'True
                  TabIndex        =   235
                  Top             =   240
                  Width           =   660
               End
            End
            Begin VB.Frame Frame28 
               BackColor       =   &H00E2E9E9&
               Height          =   690
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   226
               Top             =   3780
               Width           =   7215
               Begin MSComCtl2.DTPicker DTPicker1 
                  Height          =   330
                  Left            =   3360
                  TabIndex        =   227
                  Top             =   240
                  Width           =   1815
                  _ExtentX        =   3201
                  _ExtentY        =   582
                  _Version        =   393216
                  CheckBox        =   -1  'True
                  Format          =   235929603
                  CurrentDate     =   38887
               End
               Begin MSComCtl2.DTPicker DTPicker2 
                  Height          =   330
                  Left            =   360
                  TabIndex        =   228
                  Top             =   240
                  Width           =   1815
                  _ExtentX        =   3201
                  _ExtentY        =   582
                  _Version        =   393216
                  CheckBox        =   -1  'True
                  Format          =   235929603
                  CurrentDate     =   38887
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " «—ÌŒ «·⁄„·Ì…"
                  Height          =   195
                  Index           =   86
                  Left            =   5760
                  RightToLeft     =   -1  'True
                  TabIndex        =   231
                  Top             =   240
                  Width           =   1425
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„‰"
                  Height          =   315
                  Index           =   85
                  Left            =   5280
                  RightToLeft     =   -1  'True
                  TabIndex        =   230
                  Top             =   240
                  Width           =   660
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "≈·Ï"
                  Height          =   315
                  Index           =   84
                  Left            =   2280
                  RightToLeft     =   -1  'True
                  TabIndex        =   229
                  Top             =   240
                  Width           =   1080
               End
            End
            Begin VB.Frame Frame27 
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   855
               Left            =   -120
               RightToLeft     =   -1  'True
               TabIndex        =   224
               Top             =   0
               Width           =   13665
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·»ÕÀ ⁄‰ «„—  Õ„Ì·"
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
                  Index           =   9
                  Left            =   6600
                  RightToLeft     =   -1  'True
                  TabIndex        =   225
                  Top             =   240
                  Width           =   5400
               End
               Begin VB.Image Image8 
                  Height          =   615
                  Left            =   12360
                  Picture         =   "FrmInsurancesSearch.frx":5561D
                  Stretch         =   -1  'True
                  Top             =   120
                  Width           =   735
               End
            End
            Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid6 
               Height          =   2850
               Left            =   120
               TabIndex        =   241
               Top             =   915
               Width           =   13035
               _cx             =   22992
               _cy             =   5027
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
               BackColor       =   16777215
               ForeColor       =   -2147483640
               BackColorFixed  =   14871017
               ForeColorFixed  =   -2147483630
               BackColorSel    =   -2147483635
               ForeColorSel    =   -2147483634
               BackColorBkg    =   -2147483633
               BackColorAlternate=   16777088
               GridColor       =   -2147483633
               GridColorFixed  =   -2147483633
               TreeColor       =   -2147483632
               FloodColor      =   192
               SheetBorder     =   -2147483633
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
               FormatString    =   $"FrmInsurancesSearch.frx":56DCA
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
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·—ﬁ„"
      Height          =   195
      Index           =   74
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   202
      Top             =   0
      Width           =   1185
   End
End
Attribute VB_Name = "FrmInsurancesSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim Rs_Temp As ADODB.Recordset
Dim DCboSearch As FrmInsurancesSearch
Public SendForm As Integer
Public BankInx As Integer
Private Sub Code1_KeyPress(KeyAscii As Integer)
    Dim EmpID As Integer
    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode Code1.text, EmpID
        Me.applicantDC.BoundText = EmpID
    End If
End Sub
Private Sub applicantDC_Change()
    applicantDC_Click (0)
End Sub
Private Sub applicantDC_Click(Area As Integer)
    If val(applicantDC.BoundText) = 0 Then Exit Sub
    Dim EmpCode  As String
    GetEmployeeIDFromCode , , applicantDC.BoundText, EmpCode
    Code1.text = EmpCode
End Sub

Private Sub DcEmployee_Change()
DcEmployee_Click (0)
End Sub

Private Sub DcEmployee_Click(Area As Integer)
    If val(DCEmployee.BoundText) = 0 Then Exit Sub
      Dim EmpCode  As String
      GetEmployeeIDFromCode , , DCEmployee.BoundText, EmpCode
      Text6.text = EmpCode
End Sub

Private Sub DetSearch_Click()
    If DetSearch.value = vbChecked Then
        VSFlexGrid3.ColHidden(VSFlexGrid3.ColIndex("Employee")) = False
        VSFlexGrid3.ColHidden(VSFlexGrid3.ColIndex("TotalMark")) = False
        Frame16.Visible = True
        Frame16.Enabled = True
    Else
        VSFlexGrid3.ColHidden(VSFlexGrid3.ColIndex("Employee")) = True
        VSFlexGrid3.ColHidden(VSFlexGrid3.ColIndex("TotalMark")) = True
        Frame16.Visible = False
        Frame16.Enabled = False
    End If
End Sub
Private Sub Emp_Code_Change()
    Dim val1, val2, str As String
    If Emp_Code.text = "" Then Exit Sub
        str = " select * From TblEmployee where  fullcode = '" & Emp_Code.text & "'"
        Set Rs_Temp = New ADODB.Recordset
        Rs_Temp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
        If Rs_Temp.RecordCount > 0 Then
            Rs_Temp.MoveFirst '
            val1 = IIf(IsNull(Rs_Temp("Emp_ID").value), "", Rs_Temp("Emp_ID").value)
        Else
            val1 = ""
        End If
        EmployeeID.BoundText = val1
End Sub
Private Sub EmployeeID_Click(Area As Integer)
    Dim val1, val2, str As String
    If EmployeeID.BoundText = "" Then Exit Sub
        Emp_Code.text = ""
        str = " select * From TblEmployee where  Emp_ID = " & val(EmployeeID.BoundText)
        Set Rs_Temp = New ADODB.Recordset
        Rs_Temp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
        If Rs_Temp.RecordCount > 0 Then
            Rs_Temp.MoveFirst '
            val1 = IIf(IsNull(Rs_Temp("FullCode").value), "", Rs_Temp("FullCode").value)
        Else
            val1 = ""
        End If
End Sub

Private Sub fg_Click()
    FrmInsurances.FindRec val(FG.TextMatrix(FG.Row, 1))
End Sub
Private Sub fg_Evaluation_Click()
    FrmEvaluation.Retrive (val(fg_Evaluation.TextMatrix(fg_Evaluation.Row, fg_Evaluation.ColIndex("ID"))))
End Sub
Sub LodR()
Dim str As String
  If SystemOptions.UserInterface = ArabicInterface Then
      str = " SELECT     dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.FlagDriver, dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmpJobsTypes.JobTypeNamee, "
    str = str & "                   dbo.TblEmployee.Emp_Namee"
   Else
   str = " SELECT     dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Namee, dbo.TblEmployee.FlagDriver, dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmpJobsTypes.JobTypeNamee, "
    str = str & "                   dbo.TblEmployee.Emp_Name"
   End If
    str = str & "    FROM         dbo.TblEmployee LEFT OUTER JOIN"
    str = str & "                    dbo.TblEmpJobsTypes ON dbo.TblEmployee.JobTypeID = dbo.TblEmpJobsTypes.JobTypeID"
    
   If SystemOptions.ShowDriverOnly = True Then
   str = str & "     where  ( JobTypeName like '%”«∆ﬁ%'  or JobTypeNamee like '%driver%' )or (FlagDriver=1) "
   End If
    fill_combo DCEmployee, str

End Sub
Private Sub Form_Load()
    Load_Evaluation
    
    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500
    
    Dim GrdBack As ClsBackGroundPic
    Dim Dcombos As ClsDataCombos
    
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Clear").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    Set Dcombos = New ClsDataCombos
    Dcombos.GetBranches Me.dcBranch
    Dcombos.GetInsurancesCode Me.DcboBox
    Dcombos.GetEmployees EmpCB
    Dcombos.GetBranches BranchDC
    Dcombos.GetEmployees EmployeeDC
    Dcombos.GetBranches BranchDCBank
    Dcombos.GetEmployees applicantDC
    Dcombos.GetEmployees ResEmpDC
    Dcombos.GetCustomersSuppliers 2, SuppDC
    Dcombos.GetCarByVonder DcbCar2
    Dcombos.GetCars Me.DcbCar
    Dcombos.GetCustomersSuppliers 1, Me.DBCboClientName1
    LodR
     With PayTypeCB
        .Clear
        If SystemOptions.UserInterface = ArabicInterface Then
            .AddItem ("‰ﬁœ")
            .AddItem ("‘Ìﬂ „’œﬁ")
            .AddItem ("ÕÊ«·Â »‰ﬂÌ… Õ”» «·„—›ﬁ")
        Else
            .AddItem ("Cash")
            .AddItem ("Certified Check")
            .AddItem ("Bank Transfer")
        End If
    End With
    
    
    Set GrdBack = New ClsBackGroundPic
    With Me.FG
        Set .WallPaper = GrdBack.Picture
        .AutoSize 0, .Cols - 1, False
    End With
    YearMonth

    SetDtpickerDate Me.DtpDateFrom
    SetDtpickerDate Me.DtpDateTo
    EvaClamDateFrom.value = Date
    EvaClamDateTo.value = Date
    EvaClamDateFrom.value = ""
    EvaClamDateTo.value = ""
    EvaDateFrom.value = Date
    EvaDateTo.value = Date
    EvaDateFrom.value = ""
    EvaDateTo.value = ""
    DTPicker1.value = Date
    DTPicker2.value = Date
    DTPicker1.value = ""
    DTPicker2.value = ""
    TransFrom.value = Date
    TransFrom.value = ""
    TransTo.value = Date
    TransTo.value = ""
    ReqDateFrom.value = Date
    ReqDateFrom.value = ""
    ReqDateTo.value = Date
    ReqDateTo.value = ""
    
    Date2From.value = Date
    Date2From.value = ""
    Date2To.value = Date
    Date2To.value = ""
    CompDateFrom.value = Date
    CompDateFrom.value = ""
    CompDateTo.value = Date
    CompDateTo.value = ""
    SubDateFrom.value = Date
    SubDateFrom.value = ""
    SubDateTo.value = Date
    SubDateTo.value = ""
    OpenDateFrom.value = Date
    OpenDateFrom.value = ""
    
    OpenDateTo.value = Date
    OpenDateTo.value = ""

    If SendForm = 5 And BankInx = 1 Then
        Label1(7).Caption = "»ÕÀ ÿ·»«  «·÷„«‰ «·»‰ﬂÌ"
    ElseIf SendForm = 5 And BankInx = 2 Then
        Label1(7).Caption = "»ÕÀ ÿ·»«   „œÌœ «·÷„«‰ «·»‰ﬂÌ "
    ElseIf SendForm = 5 And BankInx = 3 Then
        Label1(7).Caption = "»ÕÀ ÿ·»«  «·÷„«‰ «·»‰ﬂÌ «·‰Â«∆Ì"
     ElseIf SendForm = 7 And BankInx = 703 Then
        Label1(9).Caption = "»ÕÀ « ›«ﬁÌ«  «·⁄„·«¡"
    End If
    
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
    
    C1Tab1.TabVisible(SendForm) = True
    C1Tab1.CurrTab = SendForm
End Sub
Private Sub Cmd_Click(Index As Integer)
    Select Case Index
        Case 0
            If SendForm = 1 Then
                GetData
            ElseIf SendForm = 2 Then
                GetTaskData
            ElseIf SendForm = 3 Then
                GetEvaData
            ElseIf SendForm = 4 Then
                GetEvaClamData
            ElseIf SendForm = 5 Then
                GetBankData
            ElseIf SendForm = 6 Then
                GetCompData
            ElseIf SendForm = 7 And (BankInx = 702 Or BankInx = 701) Then
                GetUploadData
            ElseIf SendForm = 7 And BankInx = 703 Then
                GetUploadData2
            Else
                GetData_Evaluation
            End If
        Case 1
            clear_all Me
            Me.DtpDateFrom.value = ""
            Me.DtpDateTo.value = ""
            EvaClamDateFrom.value = ""
            EvaClamDateTo.value = ""
            EvaDateFrom.value = ""
            EvaDateTo.value = ""
            DTPicker1.value = ""
            DTPicker2.value = ""
            If SystemOptions.UserInterface = ArabicInterface Then
                Me.lbll(0).Caption = "‰ ÌÃ… «·»ÕÀ"
            Else
                Me.lbll(0).Caption = "Search Results"
            End If
         Case 2
         Unload Me
    End Select
End Sub
Private Sub Form_Unload(Cancel As Integer)
    FormPostion Me, SavePostion
    Set DCboSearch = Nothing
End Sub
Private Sub YearMonth()
    Dim i As Integer
    Dim IntDefIndex As Integer
    CmbMonth.Clear
    For i = 1 To 12
        CmbMonth.AddItem MonthName(i)
    Next
    CmbMonth.ListIndex = Month(Date) - 1
    CboYear.Clear
    For i = 2006 To 2050
        CboYear.AddItem i
        If i = year(Date) Then
            IntDefIndex = CboYear.NewIndex
        End If
    Next
    CboYear.ListIndex = IntDefIndex
End Sub
Public Sub GetData()
    Dim sql As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
    
    sql = "SELECT dbo.TBLInsurances.IDINS, dbo.TBLInsurances.DateM, dbo.TBLInsurances.DateH, dbo.TBLInsurances.BranchID, dbo.TBLInsurances.SignalID, dbo.TBLInsurances.Monthe,"
    sql = sql & " dbo.TBLInsurances.SubYear, dbo.TBLInsurances.SudePerce, dbo.TBLInsurances.UnSudePerce, dbo.TBLInsurances.Acount1, dbo.TBLInsurances.Acount2, dbo.TBLInsurances.Totall,"
    sql = sql & " dbo.TblBranchesData.branch_name , dbo.TblBranchesData.branch_nameE, dbo.MOFRAD.Name, dbo.MOFRAD.NameE"
    sql = sql & " FROM dbo.TBLInsurances LEFT OUTER JOIN"
    sql = sql & " dbo.TblBranchesData ON dbo.TBLInsurances.BranchID = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
    sql = sql & " dbo.mofrad ON dbo.TBLInsurances.SignalID = dbo.mofrad.id"
                  
    BolBegine = False
    StrWhere = ""

    If val(Me.TxtIDFrom.text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & "AND  dbo.TBLInsurances.IDINS >=" & val(Me.TxtIDFrom.text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TBLInsurances.IDINS >=" & val(Me.TxtIDFrom.text) & ""
        End If
    End If
    If val(Me.TxtIDTO.text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TBLInsurances.IDINS <=" & val(Me.TxtIDTO.text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TBLInsurances.IDINS <=" & val(Me.TxtIDTO.text) & ""
        End If
    End If
    If Not IsNull(Me.DtpDateFrom.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TBLInsurances.DateM >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TBLInsurances.DateM >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        End If
    End If
    If Not IsNull(Me.DtpDateTo.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TBLInsurances.DateM <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TBLInsurances.DateM <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        End If
    End If
    
    If Me.CmbMonth.text <> "" And (val(CmbMonth.ListIndex) <> -1) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND  dbo.TBLInsurances.Monthe =" & Me.CmbMonth.ListIndex & ""
        Else:
            BolBegine = True
            StrWhere = " Where dbo.TBLInsurances.Monthe =" & Me.CmbMonth.ListIndex & ""
        End If
    End If
    If Me.CboYear.text <> "" And (val(CboYear.ListIndex) <> -1) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND  dbo.TBLInsurances.SubYear =" & Me.CboYear.text & ""
        Else:
            BolBegine = True
            StrWhere = " Where dbo.TBLInsurances.SubYear =" & Me.CboYear.ListIndex & ""
        End If
    End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Me.dcBranch.text <> "" And (val(dcBranch.BoundText) <> 0) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND  dbo.TBLInsurances.BranchID =" & Me.dcBranch.BoundText & ""
        Else:
            BolBegine = True
            StrWhere = " Where dbo.TBLInsurances.BranchID =" & Me.dcBranch.BoundText & ""
        End If
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Me.DcboBox.text <> "" And (val(DcboBox.BoundText) <> 0) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TBLInsurances.SignalID =" & Me.DcboBox.BoundText & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TBLInsurances.SignalID=" & Me.DcboBox.BoundText & ""
        End If
    End If
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If TxtCivilin.text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TBLInsurances.SudePerce like '%" & Me.TxtCivilin.text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TBLInsurances.SudePerce like '%" & Me.TxtCivilin.text & "%'"
        End If
    End If
    If TxtStay.text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TBLInsurances.UnSudePerce like '%" & Me.TxtStay.text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TBLInsurances.UnSudePerce like '%" & Me.TxtStay.text & "%'"
        End If
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    sql = sql & StrWhere
    sql = sql & " Order By dbo.TBLInsurances.IDINS"
    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Me.lbll(10).Caption = "‰ ÌÃ… «·»ÕÀ  =  ’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            Me.lbll(10).Caption = "Search Results=0"
        End If
        MsgBox "⁄›Ê« ...·« ÌÊÃœ »Ì«‰«   ‰«”» ‘—Êÿ «·»ÕÀ ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
        Cmd_Click (1)
        Exit Sub
    Else
        With Me.FG
            .Clear flexClearScrollable, flexClearEverything
            .rows = .FixedRows
            .rows = rs.RecordCount + .FixedRows
            If SystemOptions.UserInterface = ArabicInterface Then
                Me.lbll(10).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
                Me.lbll(10).Caption = "Search Results=" & rs.RecordCount
            End If
            rs.MoveFirst
            For i = .FixedRows To .rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                If Not (IsNull((rs("SubYear").value))) Then
                    .TextMatrix(i, .ColIndex("SubYear")) = rs("SubYear").value + 2006
                End If
                If Not (IsNull((rs("Monthe").value))) Then
                    .TextMatrix(i, .ColIndex("Monthe")) = MonthName(rs("Monthe").value + 1)
                End If
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs("IDINS").value), "", rs("IDINS").value)
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs("IDINS").value), "", rs("IDINS").value)
                If Not (IsNull(rs("DateM").value)) Then
                    .TextMatrix(i, .ColIndex("RecordDate")) = Format(rs("DateM").value, "yyyy/M/d")
                End If
                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(rs("branch_name").value), "", rs("branch_name").value)
                    .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs("Name").value), "", rs("Name").value)
                Else
                    .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(rs("branch_nameE").value), "", rs("branch_nameE").value)
                    .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs("NameE").value), "", rs("NameE").value)
                End If
                .TextMatrix(i, .ColIndex("SudePerce")) = IIf(IsNull(rs("SudePerce").value), "", rs("SudePerce").value)
                .TextMatrix(i, .ColIndex("UnSudePerce")) = IIf(IsNull(rs("UnSudePerce").value), "", rs("UnSudePerce").value)
                .TextMatrix(i, .ColIndex("Totall")) = IIf(IsNull(rs("Totall").value), "", rs("Totall").value)
                rs.MoveNext
            Next i
            .AutoSize 0, .Cols - 1, False
        End With
    End If
End Sub
Private Sub ChangeLang()
    Cmd(1).Caption = "Delete"
    Cmd(0).Caption = "Search"
    Cmd(2).Caption = "Exit"
    Me.Label1(2).Caption = Me.Caption
    Me.lbl(14).Caption = "Operation ID"
    Me.lbl(5).Caption = "From"
    Me.lbl(6).Caption = "To"
    Me.lbl(13).Caption = "Operation Date"
    Me.lbl(4).Caption = "From"
    Me.lbl(3).Caption = "To"
    Me.lblLL.Caption = "Branch Name"
    Me.lbl(7).Caption = "Single Name"
    lbl(0).Caption = "citizens Percentage"
    lbl(1).Caption = "Residents Percentage"
    lbl(2).Caption = "Totall Serch"
    Ele(3).Caption = "Period"
    lbl(10).Caption = "Month"
    lbl(11).Caption = "Year"
    Label1(6).Caption = "Search Maturity Evaluation"
    With Me.FG
        .TextMatrix(0, .ColIndex("Serial")) = "Serial"
        .TextMatrix(0, .ColIndex("id")) = "Code"
        .TextMatrix(0, .ColIndex("RecordDate")) = "Record Date"
        .TextMatrix(0, .ColIndex("branch_name")) = "Branch"
        .TextMatrix(0, .ColIndex("Name")) = "Single Name"
        .TextMatrix(0, .ColIndex("SudePerce")) = "citizens Percentage"
        .TextMatrix(0, .ColIndex("UnSudePerce")) = "Residents Percentage"
        .TextMatrix(0, .ColIndex("Totall")) = "Totall"
        .TextMatrix(0, .ColIndex("Monthe")) = "Month"
        .TextMatrix(0, .ColIndex("SubYear")) = "Year"
    End With
    
    Label1(4).Caption = "Define tasks search "
  
    With VSFlexGrid1
        .TextMatrix(0, .ColIndex("Serial")) = "No."
        .TextMatrix(0, .ColIndex("ID")) = "Process No."
        .TextMatrix(0, .ColIndex("RecordDate")) = "Record Date"
        .TextMatrix(0, .ColIndex("Name")) = "Task Arabic Name"
        .TextMatrix(0, .ColIndex("Namee")) = "Task English Name"
        .TextMatrix(0, .ColIndex("Des")) = "Description"
        .TextMatrix(0, .ColIndex("Employee")) = "Task Manager"
        .TextMatrix(0, .ColIndex("Notes")) = "Notes"
    End With
     
    lbl(18).Caption = "Process No."
    lbl(20).Caption = "From"
    lbl(19).Caption = "To"
    lbl(21).Caption = "Record Date"
    lbl(22).Caption = "From"
    lbl(23).Caption = "To"
    lbl(17).Caption = "Arabiic Name"
    lbl(25).Caption = "English Name"
    lbl(26).Caption = "Task Manager"
    lbl(27).Caption = "Description"
    lbl(28).Caption = "Notes"
     
    lbl(15).Caption = "Number"
    lbl(16).Caption = "Code"
    Label1(1).Caption = "Year"
    lbl(12).Caption = "Evaluations supervisor"
    Label1(3).Caption = "Month"
    lbl(24).Caption = "Branch"
     
    With fg_Evaluation
        .TextMatrix(0, .ColIndex("Serial")) = "No."
        .TextMatrix(0, .ColIndex("id")) = "Number"
        .TextMatrix(0, .ColIndex("SDate")) = "Date"
        .TextMatrix(0, .ColIndex("branch_name")) = "Branch"
        .TextMatrix(0, .ColIndex("EvalEmpCode")) = "Evaluations supervisor Code"
        .TextMatrix(0, .ColIndex("EvalEmpName")) = "Evaluations supervisor"
        .TextMatrix(0, .ColIndex("YearTitle")) = "Year"
        .TextMatrix(0, .ColIndex("MonthTitle")) = "Month"
    End With
     
    Label1(0).Caption = "Evaluation search"
     
    Label1(5).Caption = "Evaluation Standerds Search"
    lbl(34).Caption = "Serial"
    lbl(32).Caption = "From"
    lbl(33).Caption = "To"
    lbl(31).Caption = "Date"
    lbl(30).Caption = "From"
    lbl(29).Caption = "To"
    lbl(39).Caption = "Arabic Name"
    lbl(38).Caption = "English Name"
    lbl(35).Caption = "No. of absence days"
    Emp_Stude(0).RightToLeft = False
    Emp_Stude(1).RightToLeft = False
    Emp_Stude(0).Caption = "Employee"
    Emp_Stude(1).Caption = "Intern"
     
    With VSFlexGrid2
        .TextMatrix(0, .ColIndex("ID")) = "Ser"
        .TextMatrix(0, .ColIndex("RecordDate")) = "Date"
        .TextMatrix(0, .ColIndex("Name")) = "Arabic Name"
        .TextMatrix(0, .ColIndex("Namee")) = "English Name"
        .TextMatrix(0, .ColIndex("EmpOrStud")) = "Employee / Intern"
        .TextMatrix(0, .ColIndex("NoAbcDays")) = "No. of absence days"
        .TextMatrix(0, .ColIndex("MaxMark")) = "Max Mark"
    End With
    
    lbl(45).Caption = "No."
    lbl(43).Caption = "From"
    lbl(44).Caption = "To"
    lbl(42).Caption = "Recored Date"
    lbl(41).Caption = "From"
    lbl(40).Caption = "To"
    lbl(48).Caption = "Branch"
    DetSearch.Caption = "Detailed search"
    lbl(50).Caption = "Employee"
    lbl(49).Caption = "Total Mark"
    
    With VSFlexGrid3
        .TextMatrix(0, .ColIndex("Serial")) = "No."
        .TextMatrix(0, .ColIndex("ID")) = "ID"
        .TextMatrix(0, .ColIndex("RecordDate")) = "Record Date"
        .TextMatrix(0, .ColIndex("Branch")) = "Branch"
        .TextMatrix(0, .ColIndex("Employee")) = "Employee"
        .TextMatrix(0, .ColIndex("TotalMark")) = "Total Mark"
    End With
    
    If SendForm = 5 And BankInx = 1 Then
        Label1(7).Caption = "Bank Pledge Request Search"
    ElseIf SendForm = 5 And BankInx = 2 Then
        Label1(7).Caption = "Bank Pledge Request Renewal Search"
    ElseIf SendForm = 5 And BankInx = 3 Then
        Label1(7).Caption = "Final Bank Pledge Request Search"
    End If
    
    With VSFlexGrid4
        .TextMatrix(0, .ColIndex("Serial")) = "No."
        .TextMatrix(0, .ColIndex("ID")) = "ID"
        .TextMatrix(0, .ColIndex("RecordDate")) = "Record Date"
        .TextMatrix(0, .ColIndex("Branch")) = "Branch"
        .TextMatrix(0, .ColIndex("ReqDate")) = "Request Date"
        .TextMatrix(0, .ColIndex("Number")) = "Number"
        .TextMatrix(0, .ColIndex("Employee")) = "Applicant"
        .TextMatrix(0, .ColIndex("beneficiary")) = "Beneficiary Name"
    End With
    
    lbl(46).Caption = "ID"
    lbl(51).Caption = "From"
    lbl(47).Caption = "To"
    lbl(52).Caption = "Record Date"
    lbl(53).Caption = "From"
    lbl(54).Caption = "To"
    lbl(55).Caption = "Branch"
    lbl(37).Caption = "Number"
    lbl(56).Caption = "Beneficiary Name"
    lbl(36).Caption = "Applicant"
    lbl(57).Caption = "Request Date"
    lbl(58).Caption = "From"
    lbl(59).Caption = "To"
    
    lbl(66).Caption = "ID"
    lbl(64).Caption = "From"
    lbl(65).Caption = "To"
    lbl(63).Caption = "Record Date"
    lbl(62).Caption = "From"
    lbl(61).Caption = "To"
    lbl(60).Caption = "Branch"
    lbl(73).Caption = "Department"
    lbl(83).Caption = "Competition Name"
    lbl(70).Caption = "Requesting Party"
    lbl(82).Caption = "Competition Number"
    lbl(78).Caption = "Copy Price"
    lbl(81).Caption = "Competition Date"
    lbl(71).Caption = "Payment Type"
    lbl(77).Caption = "Submitting Date"
    lbl(72).Caption = "Responsible Employee"
    lbl(69).Caption = "Open Envelope Date"
    lbl(80).Caption = "From"
    lbl(79).Caption = "To"
    lbl(76).Caption = "From"
    lbl(75).Caption = "To"
    lbl(68).Caption = "From"
    lbl(67).Caption = "To"

    With VSFlexGrid5
        .TextMatrix(0, .ColIndex("ID")) = "ID"
        .TextMatrix(0, .ColIndex("RecordDate")) = "Record Date"
        .TextMatrix(0, .ColIndex("Branch")) = "Branch"
        .TextMatrix(0, .ColIndex("Dep")) = "Department"
        .TextMatrix(0, .ColIndex("CompDate")) = "Competition Date"
        .TextMatrix(0, .ColIndex("ReqPart")) = "Requesting Party"
        .TextMatrix(0, .ColIndex("CopyPrice")) = "Copy Price"
        .TextMatrix(0, .ColIndex("CompName")) = "Competition Name"
        .TextMatrix(0, .ColIndex("CompNo")) = "Competition Number"
        .TextMatrix(0, .ColIndex("SubDate")) = "Submitting Date"
        .TextMatrix(0, .ColIndex("OpenEnvDate")) = "Open Envelope Date"
        .TextMatrix(0, .ColIndex("PayType")) = "Payment Type"
        .TextMatrix(0, .ColIndex("ResEmp")) = "Responsible Employee"
    End With
    
    SuppRd.Caption = "Supplier"
    OtherRd.Caption = "Other"
    
    Label1(8).Caption = "Tender Purchase Form Search"
End Sub

Private Sub OtherRd_Click()
    If SuppRd.value = True Then
        SuppDC.Enabled = True
        SuppDC.Visible = True
        OtherNameTxt.Enabled = False
        OtherNameTxt.Visible = False
    ElseIf OtherRd.value = True Then
        SuppDC.Enabled = False
        SuppDC.Visible = False
        OtherNameTxt.Enabled = True
        OtherNameTxt.Visible = True
    End If
End Sub

Private Sub Refrish_Click()
    TxtCivilin.text = 12
End Sub
Private Sub Refrish2_Click()
    TxtStay.text = 2
End Sub
Private Sub Load_Evaluation()
    Dim Dcombos As ClsDataCombos
    Dim str As String
    
    Set Dcombos = New ClsDataCombos
    Dcombos.GetBranches BranchID
    Dcombos.GetEmployees EmployeeID
    
    Dim i As Integer
    For i = 2006 To 2050
        YearID.AddItem i
    Next
   
    For i = 1 To 12
        MonthID.AddItem MonthName(i)
    Next
    
    For i = 2006 To 2050
        YearID.AddItem i
    Next
    For i = 1 To 12
        MonthID.AddItem MonthName(i)
    Next
End Sub
Public Sub GetData_Evaluation()
    Dim sql As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
    Dim MySQL As String

    fg_Evaluation.rows = fg_Evaluation.FixedRows
    
    MySQL = "SELECT dbo.TblEmpEvaluation.ID, dbo.TblEmpEvaluation.SDate, dbo.TblBranchesData.branch_name, dbo.TblEmpEvaluation.EmployeeID,"
    MySQL = MySQL & " dbo.TblEmployee.Emp_Code AS EvalEmpCode, dbo.TblEmployee.Emp_Name AS EvalEmpName, dbo.TblEmpEvaluation.YearID, dbo.TblEmpEvaluation.MonthID,"
    MySQL = MySQL & " TblEmployee_1.Emp_Code AS ECode, TblEmployee_1.Emp_Name AS EName, dbo.TblEmpEvaluation.Branch, B.branch_name AS BName, dbo.TblEmpEvaluation.Project,"
    MySQL = MySQL & " dbo.TblEmpEvaluation.Mangerial, dbo.TblSection.name AS SectionName, dbo.TblEmpEvaluation.opt_AllEmployees, dbo.TblEmpEvaluation.opt_OneEmployee,"
    MySQL = MySQL & " dbo.TblEmpEvaluation.opt_Branch, dbo.TblEmpEvaluation.opt_Project, dbo.TblEmpEvaluation.opt_Managerial, dbo.TblEmpEvaluation.YearTitle,"
    MySQL = MySQL & " dbo.TblEmpEvaluation.MonthTitle"
    MySQL = MySQL & " FROM dbo.projects RIGHT OUTER JOIN"
    MySQL = MySQL & " dbo.TblSection RIGHT OUTER JOIN"
    MySQL = MySQL & " dbo.TblEmpEvaluation ON dbo.TblSection.Id = dbo.TblEmpEvaluation.Mangerial RIGHT OUTER JOIN"
    MySQL = MySQL & " dbo.TblEmployee ON dbo.TblEmpEvaluation.EmployeeID = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
    MySQL = MySQL & " dbo.TblBranchesData AS B ON dbo.TblEmpEvaluation.Branch = B.branch_id RIGHT OUTER JOIN"
    MySQL = MySQL & " dbo.TblBranchesData ON dbo.TblEmpEvaluation.BranchID = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
    MySQL = MySQL & " dbo.TblEmployee AS TblEmployee_1 ON dbo.TblEmpEvaluation.OneEmployee = TblEmployee_1.Emp_ID ON dbo.projects.id = dbo.TblEmpEvaluation.Project"
    MySQL = MySQL & "   where 1 =1  "
    
    If ID.text <> "" Then
        MySQL = MySQL & " and dbo.TblEmpEvaluation.ID =  " & val(ID.text)
    End If

    If YearID.ListIndex <> -1 Then
        MySQL = MySQL & " and dbo.TblEmpEvaluation.YearID =  " & val(YearID.ListIndex)
    End If

    If MonthID.ListIndex <> -1 Then
        MySQL = MySQL & " and dbo.TblEmpEvaluation.MonthID =  " & val(MonthID.ListIndex)
    End If

    If BranchID.BoundText <> "" Then
        MySQL = MySQL & " and  TblEmpEvaluation.BranchID = " & val(BranchID.BoundText)
    End If
    
    If EmployeeID.BoundText <> "" Then
        MySQL = MySQL & "  and TblEmpEvaluation.EmployeeID = " & val(EmployeeID.BoundText)
    End If
    
    MySQL = MySQL & " Order By  dbo.TblEmpEvaluation.ID  "
    Set rs = New ADODB.Recordset
    rs.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    
    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Me.lbll(10).Caption = "‰ ÌÃ… «·»ÕÀ  =  ’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            Me.lbll(10).Caption = "Search Results=0"
        End If
        MsgBox "⁄›Ê« ...·« ÌÊÃœ »Ì«‰«   ‰«”» ‘—Êÿ «·»ÕÀ ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
        Cmd_Click (1)
        Exit Sub
    Else
        With Me.fg_Evaluation
            .Clear flexClearScrollable, flexClearEverything
            .rows = .FixedRows
            .rows = rs.RecordCount + .FixedRows
           
            rs.MoveFirst
            For i = .FixedRows To .rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs("id").value), "", rs("ID").value)
                .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(rs("branch_name").value), "", rs("branch_name").value)
                .TextMatrix(i, .ColIndex("SDate")) = IIf(IsNull(rs("SDate").value), "", rs("SDate").value)
                .TextMatrix(i, .ColIndex("YearTitle")) = IIf(IsNull(rs("YearTitle").value), "", rs("YearTitle").value)
                .TextMatrix(i, .ColIndex("MonthTitle")) = IIf(IsNull(rs("MonthTitle").value), "", rs("MonthTitle").value)
                .TextMatrix(i, .ColIndex("EvalEmpCode")) = IIf(IsNull(rs("EvalEmpCode").value), "", rs("EvalEmpCode").value)
                .TextMatrix(i, .ColIndex("EvalEmpName")) = IIf(IsNull(rs("EvalEmpName").value), "", rs("EvalEmpName").value)
                rs.MoveNext
            Next i
            .AutoSize 0, .Cols - 1, False
          End With
    End If
End Sub
Public Sub GetTaskData()
    Dim sql As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
    Dim MySQL As String


    VSFlexGrid1.rows = VSFlexGrid1.FixedRows
        
    MySQL = "SELECT TblProceeDevelper.ID, TblProceeDevelper.RecoedDate, TblProceeDevelper.EmpID, TblProceeDevelper.Remark, TblProceeDevelper.Description, TblProceeDevelper.Name, TblProceeDevelper.NameE, TblProceeDevelper.Des, "
    MySQL = MySQL & " TblProceeDevelper.NoDay, TblProceeDevelper.StartDate, TblProceeDevelper.EndDate, TblProceeDevelper.UserID, TblProceeDevelper.Priority, TblProceeDevelper.EmpID1, TblProceeDevelper.DesE, TblEmployee.Emp_Name, "
    MySQL = MySQL & " TblEmployee.Emp_Namee "
    MySQL = MySQL & " FROM TblProceeDevelper INNER JOIN "
    MySQL = MySQL & " TblEmployee ON TblProceeDevelper.EmpID = TblEmployee.Emp_ID "
    MySQL = MySQL & "   where 1 = 1  "
    
    If IDFromTxt.text <> "" Then
          MySQL = MySQL & " and TblProceeDevelper.ID >= " & val(IDFromTxt.text)
    End If
    
    If IDToTxt.text <> "" Then
          MySQL = MySQL & " and TblProceeDevelper.ID <= " & val(IDToTxt.text)
    End If
    
    If Not IsNull(Me.DateFrom.value) Then
        MySQL = MySQL & " AND TblProceeDevelper.RecoedDate >=" & SQLDate(Me.DateFrom.value, True) & ""
    End If
    
    If Not IsNull(Me.DateTo.value) Then
        MySQL = MySQL & " AND TblProceeDevelper.RecoedDate <=" & SQLDate(Me.DateTo.value, True) & ""
    End If
    
    If NameTxt.text <> "" Then
          MySQL = MySQL & "and TblProceeDevelper.Name like N'%" & NameTxt.text & "%'"
    End If
    
    If NameeTxt.text <> "" Then
          MySQL = MySQL & "and TblProceeDevelper.NameE like N'%" & NameeTxt.text & "%'"
    End If
    
    If DesTxt.text <> "" Then
          MySQL = MySQL & "and TblProceeDevelper.Description like N'%" & DesTxt.text & "%'"
    End If
    
    If NotesTxt.text <> "" Then
          MySQL = MySQL & "and TblProceeDevelper.Remark like N'%" & NotesTxt.text & "%'"
    End If
    
    If EmpCB.text <> "" And EmpCB.BoundText <> "" Then
          MySQL = MySQL & "and TblProceeDevelper.EmpID1 = " & EmpCB.BoundText & ""
    End If
 
    MySQL = MySQL & " Order By  TblProceeDevelper.ID  "
    Set rs = New ADODB.Recordset
    rs.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    
    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Me.lbll(10).Caption = "‰ ÌÃ… «·»ÕÀ  =  ’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            Me.lbll(10).Caption = "Search Results=0"
        End If
        MsgBox "⁄›Ê« ...·« ÌÊÃœ »Ì«‰«   ‰«”» ‘—Êÿ «·»ÕÀ ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
        Cmd_Click (1)
        Exit Sub
    Else
        With Me.VSFlexGrid1
            .Clear flexClearScrollable, flexClearEverything
            .rows = .FixedRows
            .rows = rs.RecordCount + .FixedRows
           
            rs.MoveFirst
            For i = .FixedRows To .rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("ID")) = IIf(IsNull(rs("ID").value), "", rs("ID").value)
                .TextMatrix(i, .ColIndex("RecordDate")) = IIf(IsNull(rs("RecoedDate").value), "", rs("RecoedDate").value)
                .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs("Name").value), "", rs("Name").value)
                .TextMatrix(i, .ColIndex("Namee")) = IIf(IsNull(rs("NameE").value), "", rs("NameE").value)
                .TextMatrix(i, .ColIndex("Des")) = IIf(IsNull(rs("Description").value), "", rs("Description").value)
                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("Employee")) = IIf(IsNull(rs("Emp_Name").value), "", rs("Emp_Name").value)
                Else
                    .TextMatrix(i, .ColIndex("Employee")) = IIf(IsNull(rs("Emp_Namee").value), "", rs("Emp_Namee").value)
                End If
                .TextMatrix(i, .ColIndex("Notes")) = IIf(IsNull(rs("Remark").value), "", rs("Remark").value)
                rs.MoveNext
            Next i
            .AutoSize 0, .Cols - 1, False
          End With
    End If
End Sub

Private Sub SuppRd_Click()
    If SuppRd.value = True Then
        SuppDC.Enabled = True
        SuppDC.Visible = True
        OtherNameTxt.Enabled = False
        OtherNameTxt.Visible = False
    ElseIf OtherRd.value = True Then
        SuppDC.Enabled = False
        SuppDC.Visible = False
        OtherNameTxt.Enabled = True
        OtherNameTxt.Visible = True
    End If
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
  Dim EmpID As Integer
    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode Text6.text, EmpID
        DCEmployee.BoundText = EmpID
    End If
End Sub

Private Sub VSFlexGrid1_Click()
    FrmOpDevelopment1.FindRec val(VSFlexGrid1.TextMatrix(VSFlexGrid1.Row, 1))
End Sub
Public Sub GetEvaData()
    Dim sql As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
    Dim MySQL As String

    VSFlexGrid2.rows = VSFlexGrid2.FixedRows
        
    MySQL = "select * from TblEvaluationStandered where (1 = 1)"
    
    If EvaIDFromTxt.text <> "" Then
          MySQL = MySQL & " and TblEvaluationStandered.ID >= " & val(EvaIDFromTxt.text)
    End If
    
    If EvaIDToTxt.text <> "" Then
          MySQL = MySQL & " and TblEvaluationStandered.ID  <= " & val(EvaIDToTxt.text)
    End If
    
    If Not IsNull(Me.EvaDateFrom.value) Then
        MySQL = MySQL & " AND TblEvaluationStandered.SDate >=" & SQLDate(Me.EvaDateFrom.value, True) & ""
    End If
    
    If Not IsNull(Me.EvaDateTo.value) Then
        MySQL = MySQL & " AND TblEvaluationStandered.SDate <=" & SQLDate(Me.EvaDateTo.value, True) & ""
    End If
    
    If EvaNameTxt.text <> "" Then
          MySQL = MySQL & "and TblEvaluationStandered.EName like N'%" & EvaNameTxt.text & "%'"
    End If
    
    If EvaNameeTxt.text <> "" Then
          MySQL = MySQL & "and TblEvaluationStandered.ENameE like N'%" & EvaNameeTxt.text & "%'"
    End If
    
    If NoOfAbcDaysTxt.text <> "" Then
          MySQL = MySQL & "and TblEvaluationStandered.NoDayAbcen = " & val(NoOfAbcDaysTxt.text) & " "
    End If
    
    If Emp_Stude(0).value = True Then
          MySQL = MySQL & "and TblEvaluationStandered.Emp_Stude = 0"
    End If
    
    If Emp_Stude(1).value Then
           MySQL = MySQL & "and TblEvaluationStandered.Emp_Stude = 1"
    End If
 
    MySQL = MySQL & " Order By  TblEvaluationStandered.ID  "
    Set rs = New ADODB.Recordset
    rs.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    
    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Me.lbll(10).Caption = "‰ ÌÃ… «·»ÕÀ  =  ’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            Me.lbll(10).Caption = "Search Results=0"
        End If
        MsgBox "⁄›Ê« ...·« ÌÊÃœ »Ì«‰«   ‰«”» ‘—Êÿ «·»ÕÀ ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
        Cmd_Click (1)
        Exit Sub
    Else
        With Me.VSFlexGrid2
            .Clear flexClearScrollable, flexClearEverything
            .rows = .FixedRows
            .rows = rs.RecordCount + .FixedRows
           
            rs.MoveFirst
            For i = .FixedRows To .rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("ID")) = IIf(IsNull(rs("ID").value), "", rs("ID").value)
                .TextMatrix(i, .ColIndex("RecordDate")) = IIf(IsNull(rs("SDate").value), "", rs("SDate").value)
                .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs("EName").value), "", rs("EName").value)
                .TextMatrix(i, .ColIndex("Namee")) = IIf(IsNull(rs("ENameE").value), "", rs("ENameE").value)
                If SystemOptions.UserInterface = ArabicInterface Then
                    If rs("Emp_Stude").value = 0 Then
                        .TextMatrix(i, .ColIndex("EmpOrStud")) = "„ÊŸ›"
                    ElseIf rs("Emp_Stude").value = 1 Then
                        .TextMatrix(i, .ColIndex("EmpOrStud")) = "„ œ—»"
                    Else
                        .TextMatrix(i, .ColIndex("EmpOrStud")) = ""
                    End If
                Else
                    If rs("Emp_Stude").value = 0 Then
                        .TextMatrix(i, .ColIndex("EmpOrStud")) = "Employee"
                    ElseIf rs("Emp_Stude").value = 1 Then
                        .TextMatrix(i, .ColIndex("EmpOrStud")) = "Intern"
                    Else
                        .TextMatrix(i, .ColIndex("EmpOrStud")) = ""
                    End If
                End If
                .TextMatrix(i, .ColIndex("NoAbcDays")) = IIf(IsNull(rs("NoDayAbcen").value), "", rs("NoDayAbcen").value)
                .TextMatrix(i, .ColIndex("MaxMark")) = IIf(IsNull(rs("MaxDgree").value), "", rs("MaxDgree").value)
                rs.MoveNext
            Next i
            .AutoSize 0, .Cols - 1, False
          End With
    End If
End Sub
Private Sub VSFlexGrid2_Click()
    FrmEvaluation_Standerd.Retrive val(VSFlexGrid2.TextMatrix(VSFlexGrid2.Row, 1))
End Sub
Public Sub GetEvaClamData()
    Dim sql As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
    Dim MySQL As String

    VSFlexGrid3.rows = VSFlexGrid3.FixedRows
    If DetSearch.value = vbChecked Then
        MySQL = "SELECT TblEvaluaEntit.ID, TblEvaluaEntit.RecordDate, TblEvaluaEntit.FromDate, TblEvaluaEntit.ToDate, TblEvaluaEntit.UserID, TblEvaluaEntit.BrnchID, TblEvaluaEntit.EmpID, TblEvaluaEntit.DeptID, TblEvaluaEntit.BrnchID1, "
        MySQL = MySQL & " TblEvaluaEntit.ProjectID, TblEvaluaEntit.AllEmp, TblEvaluaEntit.SelDept, TblEvaluaEntit.SelBrnch, TblEvaluaEntit.SelProj, TblEvaluaEntit.Remarks, TblEvaluaEntit.TotalValue, TblEmployee.Emp_Name,"
        MySQL = MySQL & " TblEmployee.Emp_Namee, TblBranchesData.branch_name, TblBranchesData.branch_namee"
        MySQL = MySQL & " FROM TblBranchesData RIGHT OUTER JOIN"
        MySQL = MySQL & " TblEvaluaEntit ON TblBranchesData.branch_id = TblEvaluaEntit.BrnchID FULL OUTER JOIN"
        MySQL = MySQL & " TblEmployee RIGHT OUTER JOIN"
        MySQL = MySQL & " TblEvaluaEntitDet ON TblEmployee.Emp_ID = TblEvaluaEntitDet.EmpID ON TblEvaluaEntit.ID = TblEvaluaEntitDet.EvlaID where (1 = 1)"
    Else
        MySQL = "SELECT TblEvaluaEntit.ID, TblEvaluaEntit.RecordDate, TblEvaluaEntit.FromDate, TblEvaluaEntit.ToDate, TblEvaluaEntit.UserID, TblEvaluaEntit.BrnchID, TblEvaluaEntit.EmpID, TblEvaluaEntit.DeptID, TblEvaluaEntit.BrnchID1, "
        MySQL = MySQL & " TblEvaluaEntit.ProjectID, TblEvaluaEntit.AllEmp, TblEvaluaEntit.SelDept, TblEvaluaEntit.SelBrnch, TblEvaluaEntit.SelProj, TblEvaluaEntit.Remarks, TblEvaluaEntit.TotalValue, TblBranchesData.branch_name,"
        MySQL = MySQL & " TblBranchesData.branch_namee"
        MySQL = MySQL & " FROM TblBranchesData RIGHT OUTER JOIN"
        MySQL = MySQL & " TblEvaluaEntit ON TblBranchesData.branch_id = TblEvaluaEntit.BrnchID where (1 = 1)"
    End If
    
    If val(EvaClamIDFromTxt.text) <> 0 Then
          MySQL = MySQL & " and TblEvaluaEntit.ID >= " & val(EvaClamIDFromTxt.text)
    End If
    
    If val(EvaClamIDToTxt.text) <> 0 Then
          MySQL = MySQL & " and TblEvaluaEntit.ID  <= " & val(EvaClamIDToTxt.text)
    End If
    
    If Not IsNull(Me.EvaClamDateFrom.value) Then
        MySQL = MySQL & " AND TblEvaluaEntit.RecordDate >=" & SQLDate(Me.EvaClamDateFrom.value, True) & ""
    End If
    
    If Not IsNull(Me.EvaClamDateTo.value) Then
        MySQL = MySQL & " AND TblEvaluaEntit.RecordDate <=" & SQLDate(Me.EvaClamDateTo.value, True) & ""
    End If
    
    If BranchDC.BoundText <> "" Then
        MySQL = MySQL & " and  TblEvaluaEntit.BrnchID = " & val(BranchDC.BoundText)
    End If
    
    If DetSearch.value = vbChecked Then
        If EmployeeDC.BoundText <> "" Then
            MySQL = MySQL & " and  TblEvaluaEntitDet.EmpID = " & val(EmployeeDC.BoundText)
        End If
    
        If val(TotalMarkTxt.text) <> 0 Then
            MySQL = MySQL & "and TblEvaluaEntitDet.TotDigree = " & val(TotalMarkTxt.text) & ""
        End If
    End If
    
    MySQL = MySQL & " Order By  TblEvaluaEntit.ID  "
    Set rs = New ADODB.Recordset
    rs.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    
    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Me.lbll(10).Caption = "‰ ÌÃ… «·»ÕÀ  =  ’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            Me.lbll(10).Caption = "Search Results=0"
        End If
        MsgBox "⁄›Ê« ...·« ÌÊÃœ »Ì«‰«   ‰«”» ‘—Êÿ «·»ÕÀ ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
        Cmd_Click (1)
        Exit Sub
    Else
        With Me.VSFlexGrid3
            .Clear flexClearScrollable, flexClearEverything
            .rows = .FixedRows
            .rows = rs.RecordCount + .FixedRows
           
            rs.MoveFirst
            
            For i = .FixedRows To .rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("ID")) = IIf(IsNull(rs("ID").value), "", rs("ID").value)
                .TextMatrix(i, .ColIndex("RecordDate")) = IIf(IsNull(rs("RecordDate").value), "", rs("RecordDate").value)
                
                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("Branch")) = IIf(IsNull(rs("branch_name").value), "", rs("branch_name").value)
                    If DetSearch.value = vbChecked Then
                        .TextMatrix(i, .ColIndex("Employee")) = IIf(IsNull(rs("Emp_Name").value), "", rs("Emp_Name").value)
                    End If
                Else
                    .TextMatrix(i, .ColIndex("Branch")) = IIf(IsNull(rs("branch_namee").value), "", rs("branch_namee").value)
                    If DetSearch.value = vbChecked Then
                        .TextMatrix(i, .ColIndex("Employee")) = IIf(IsNull(rs("Emp_Namee").value), "", rs("Emp_Namee").value)
                    End If
                End If
                
                If DetSearch.value = vbChecked Then
                    .TextMatrix(i, .ColIndex("TotalMark")) = IIf(IsNull(rs("TotalValue").value), "", rs("TotalValue").value)
                End If
                
                rs.MoveNext
            Next i
            .AutoSize 0, .Cols - 1, False
        End With
    End If
End Sub
Private Sub VSFlexGrid3_Click()
    FrmEvaluaEntit.FindRec val(VSFlexGrid3.TextMatrix(VSFlexGrid3.Row, 1))
End Sub
Public Sub GetBankData()

    Dim sql As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
    Dim MySQL As String

    VSFlexGrid4.rows = VSFlexGrid4.FixedRows
    
    If SendForm = 5 And BankInx = 1 Then
        MySQL = "SELECT TblBankPledge.ID, TblBankPledge.RecordDate, TblBankPledge.BranchID, TblBankPledge.ReqDate, TblBankPledge.ReqTime, TblBankPledge.beneficiary, TblBankPledge.Number, TblBankPledge.CompetDate,"
        MySQL = MySQL & " TblBankPledge.Project, TblBankPledge.CompetValue, TblBankPledge.PledgeValue, TblBankPledge.OpenEnvelope, TblBankPledge.AppliedID, TblBankPledge.MangerID, TblBankPledge.GMangerID, TblBankPledge.PledgeMargin,"
        MySQL = MySQL & " TblBankPledge.ReciveDate, TblBankPledge.ThirdPartyFlg, TblBankPledge.ThirdPartyName, TblBankPledge.Notes, TblBankPledge.UserID, TblEmployee.Emp_Name, TblEmployee.Emp_Namee, TblBranchesData.branch_name,"
        MySQL = MySQL & " TblBranchesData.branch_namee"
        MySQL = MySQL & " FROM TblBankPledge LEFT OUTER JOIN"
        MySQL = MySQL & " TblEmployee ON TblBankPledge.AppliedID = TblEmployee.Emp_ID LEFT OUTER JOIN"
        MySQL = MySQL & " TblBranchesData ON TblBankPledge.BranchID = TblBranchesData.branch_id where (1 = 1)"
    ElseIf SendForm = 5 And BankInx = 2 Then
        MySQL = "SELECT TblBankPledge2.ID, TblBankPledge2.RecordDate, TblBankPledge2.BranchID, TblBankPledge2.ReqDate, TblBankPledge2.ReqTime, TblBankPledge2.beneficiary, TblBankPledge2.Number, TblBankPledge2.CompetValue,"
        MySQL = MySQL & " TblBankPledge2.PledgeValue, TblBankPledge2.AppliedID, TblBankPledge2.MangerID, TblBankPledge2.GMangerID, TblBankPledge2.PledgeMargin, TblBankPledge2.ReciveDate, TblBankPledge2.ThirdPartyFlg,"
        MySQL = MySQL & " TblBankPledge2.ThirdPartyName, TblBankPledge2.Notes, TblBankPledge2.UserID, TblEmployee.Emp_Name, TblEmployee.Emp_Namee, TblBranchesData.branch_name, TblBranchesData.branch_namee,"
        MySQL = MySQL & " TblBankPledge2.Compet"
        MySQL = MySQL & " FROM TblBankPledge2 LEFT OUTER JOIN"
        MySQL = MySQL & " TblEmployee ON TblBankPledge2.AppliedID = TblEmployee.Emp_ID LEFT OUTER JOIN"
        MySQL = MySQL & " TblBranchesData ON TblBankPledge2.BranchID = TblBranchesData.branch_id where (1 = 1)"
    ElseIf SendForm = 5 And BankInx = 3 Then
        MySQL = "SELECT TblBankPledge3.ID, TblBankPledge3.RecordDate, TblBankPledge3.BranchID, TblBankPledge3.ReqDate, TblBankPledge3.beneficiary, TblBankPledge3.Number, TblBankPledge3.PledgeValue, TblBankPledge3.AppliedID,"
        MySQL = MySQL & " TblBankPledge3.MangerID, TblBankPledge3.GMangerID, TblBankPledge3.PledgeMargin, TblBankPledge3.ReciveDate, TblBankPledge3.ThirdPartyFlg, TblBankPledge3.ThirdPartyName, TblBankPledge3.Notes,"
        MySQL = MySQL & " TblBankPledge3.UserID, TblEmployee.Emp_Name, TblEmployee.Emp_Namee, TblBranchesData.branch_name, TblBranchesData.branch_namee, TblBankPledge3.Project, TblBankPledge3.ApprovalValue,"
        MySQL = MySQL & " TblBankPledge3.pledgevalidity"
        MySQL = MySQL & " FROM TblBankPledge3 LEFT OUTER JOIN"
        MySQL = MySQL & " TblEmployee ON TblBankPledge3.AppliedID = TblEmployee.Emp_ID LEFT OUTER JOIN"
        MySQL = MySQL & " TblBranchesData ON TblBankPledge3.BranchID = TblBranchesData.branch_id where (1 = 1)"
    End If
    
    If val(BankIDFrom.text) <> 0 Then
          MySQL = MySQL & " and ID >= " & val(BankIDFrom.text)
    End If
    
    If val(BankIDTo.text) <> 0 Then
          MySQL = MySQL & " and ID  <= " & val(BankIDTo.text)
    End If
    
    If Not IsNull(Me.TransFrom.value) Then
        MySQL = MySQL & " AND RecordDate >=" & SQLDate(Me.TransFrom.value, True) & ""
    End If
    
    If Not IsNull(Me.TransTo.value) Then
        MySQL = MySQL & " AND RecordDate <=" & SQLDate(Me.TransTo.value, True) & ""
    End If
    
        If Not IsNull(Me.ReqDateFrom.value) Then
        MySQL = MySQL & " AND ReqDate >=" & SQLDate(Me.ReqDateFrom.value, True) & ""
    End If
    
    If Not IsNull(Me.ReqDateTo.value) Then
        MySQL = MySQL & " AND ReqDate <=" & SQLDate(Me.ReqDateTo.value, True) & ""
    End If
    
    If val(BranchDCBank.BoundText) <> 0 Then
        If BankInx = 1 Then
            MySQL = MySQL & " and  TblBankPledge.BranchID = " & val(BranchDCBank.BoundText)
        ElseIf BankInx = 2 Then
            MySQL = MySQL & " and  TblBankPledge2.BranchID = " & val(BranchDCBank.BoundText)
        ElseIf BankInx = 3 Then
            MySQL = MySQL & " and  TblBankPledge3.BranchID = " & val(BranchDCBank.BoundText)
        End If
    End If
    
    If applicantDC.BoundText <> "" Then
        MySQL = MySQL & " and  AppliedID = " & val(applicantDC.BoundText)
    End If
    
    If NumberTxt.text <> "" Then
          MySQL = MySQL & "and Number like N'%" & NumberTxt.text & "%'"
    End If
    
    If beneficiaryTxt.text <> "" Then
          MySQL = MySQL & "and beneficiary like N'%" & beneficiaryTxt.text & "%'"
    End If
    
    MySQL = MySQL & " Order By ID  "
    Set rs = New ADODB.Recordset
    rs.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    
    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Me.lbll(10).Caption = "‰ ÌÃ… «·»ÕÀ  =  ’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            Me.lbll(10).Caption = "Search Results=0"
        End If
        MsgBox "⁄›Ê« ...·« ÌÊÃœ »Ì«‰«   ‰«”» ‘—Êÿ «·»ÕÀ ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
        Cmd_Click (1)
        Exit Sub
    Else
        With Me.VSFlexGrid4
            .Clear flexClearScrollable, flexClearEverything
            .rows = .FixedRows
            .rows = rs.RecordCount + .FixedRows
           
            rs.MoveFirst
            
            For i = .FixedRows To .rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("ID")) = IIf(IsNull(rs("ID").value), "", rs("ID").value)
                .TextMatrix(i, .ColIndex("RecordDate")) = IIf(IsNull(rs("RecordDate").value), "", rs("RecordDate").value)
                .TextMatrix(i, .ColIndex("ReqDate")) = IIf(IsNull(rs("ReqDate").value), "", rs("ReqDate").value)
                .TextMatrix(i, .ColIndex("Number")) = IIf(IsNull(rs("Number").value), "", rs("Number").value)
                .TextMatrix(i, .ColIndex("beneficiary")) = IIf(IsNull(rs("beneficiary").value), "", rs("beneficiary").value)
                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("Branch")) = IIf(IsNull(rs("branch_name").value), "", rs("branch_name").value)
                    .TextMatrix(i, .ColIndex("Employee")) = IIf(IsNull(rs("Emp_Name").value), "", rs("Emp_Name").value)
                Else
                    .TextMatrix(i, .ColIndex("Branch")) = IIf(IsNull(rs("branch_namee").value), "", rs("branch_namee").value)
                    .TextMatrix(i, .ColIndex("Employee")) = IIf(IsNull(rs("Emp_Namee").value), "", rs("Emp_Namee").value)
                End If
                rs.MoveNext
            Next i
            .AutoSize 0, .Cols - 1, False
        End With
    End If
End Sub
Private Sub VSFlexGrid4_Click()
    If SendForm = 5 And BankInx = 1 Then
        FrmBankPledge1.Retrive val(VSFlexGrid4.TextMatrix(VSFlexGrid4.Row, 1))
    ElseIf SendForm = 5 And BankInx = 2 Then
        FrmBankPledge2.Retrive val(VSFlexGrid4.TextMatrix(VSFlexGrid4.Row, 1))
    ElseIf SendForm = 5 And BankInx = 3 Then
        FrmBankPledge3.Retrive val(VSFlexGrid4.TextMatrix(VSFlexGrid4.Row, 1))
    End If
End Sub
Public Sub GetCompData()

    Dim sql As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
    Dim MySQL As String

    VSFlexGrid5.rows = VSFlexGrid5.FixedRows
    
    MySQL = "SELECT TblBankPledge4.ID, TblBankPledge4.RecordDate, TblBankPledge4.BranchID, TblBankPledge4.DepID, TblBankPledge4.CompetDate, TblBankPledge4.SuppOrOther, TblBankPledge4.SuppID, TblBankPledge4.OtherName,"
    MySQL = MySQL & " TblBankPledge4.CopyPrice, TblBankPledge4.CompetitionName, TblBankPledge4.CompetitionNumber, TblBankPledge4.SubmittingDate, TblBankPledge4.OpenEnvelopeDate, TblBranchesData.branch_name,"
    MySQL = MySQL & " TblBranchesData.branch_namee, TblBankPledge4.AppliedID, TblBankPledge4.MangerID, TblEmployee_1.Emp_Name AS MangerName, TblEmployee_1.Emp_Namee AS MangerNamee,"
    MySQL = MySQL & " TblEmployee.Emp_Name AS AppliedName, TblEmployee.Emp_Namee AS AppliedNamee, TblCustemers.CusName, TblCustemers.CusNamee, TblBankPledge4.Notes, TblBankPledge4.PaymentType, TblBankPledge4.UserID,"
    MySQL = MySQL & " TblEmpDepartments.DepartmentName , TblEmpDepartments.DepartmentNamee"
    MySQL = MySQL & " FROM TblEmployee RIGHT OUTER JOIN"
    MySQL = MySQL & " TblEmployee AS TblEmployee_1 RIGHT OUTER JOIN"
    MySQL = MySQL & " TblCustemers RIGHT OUTER JOIN"
    MySQL = MySQL & " TblBankPledge4 LEFT OUTER JOIN"
    MySQL = MySQL & " TblEmpDepartments ON TblBankPledge4.DepID = TblEmpDepartments.DeparmentID ON TblCustemers.CusID = TblBankPledge4.SuppID ON TblEmployee_1.Emp_ID = TblBankPledge4.MangerID ON"
    MySQL = MySQL & " TblEmployee.Emp_ID = TblBankPledge4.AppliedID LEFT OUTER JOIN"
    MySQL = MySQL & " TblBranchesData ON TblBankPledge4.BranchID = TblBranchesData.branch_id where (1 = 1)"

    If val(IDDFrom.text) <> 0 Then
          MySQL = MySQL & " and ID >= " & val(IDDFrom.text)
    End If
    
    If val(IDDTo.text) <> 0 Then
          MySQL = MySQL & " and ID  <= " & val(IDDTo.text)
    End If
    
    
    If Not IsNull(Me.Date2From.value) Then
        MySQL = MySQL & " AND TblBankPledge4.RecordDate >=" & SQLDate(Me.Date2From.value, True) & ""
    End If
    
    If Not IsNull(Me.Date2To.value) Then
        MySQL = MySQL & " AND TblBankPledge4.RecordDate <=" & SQLDate(Me.Date2To.value, True) & ""
    End If
    
    
    If Not IsNull(Me.CompDateFrom.value) Then
        MySQL = MySQL & " AND CompetDate >=" & SQLDate(Me.CompDateFrom.value, True) & ""
    End If
    
    If Not IsNull(Me.CompDateTo.value) Then
        MySQL = MySQL & " AND CompetDate <=" & SQLDate(Me.CompDateTo.value, True) & ""
    End If
    
    
    
    If Not IsNull(Me.SubDateFrom.value) Then
        MySQL = MySQL & " AND SubmittingDate >=" & SQLDate(Me.SubDateFrom.value, True) & ""
    End If
    
    If Not IsNull(Me.SubDateTo.value) Then
        MySQL = MySQL & " AND SubmittingDate <=" & SQLDate(Me.SubDateTo.value, True) & ""
    End If
    
    
    If Not IsNull(Me.OpenDateFrom.value) Then
        MySQL = MySQL & " AND OpenEnvelopeDate >=" & SQLDate(Me.OpenDateFrom.value, True) & ""
    End If
    
    If Not IsNull(Me.OpenDateTo.value) Then
        MySQL = MySQL & " AND OpenEnvelopeDate <=" & SQLDate(Me.OpenDateTo.value, True) & ""
    End If
    
    If val(BranchDC2.BoundText) <> 0 Then
        MySQL = MySQL & " and  BranchID = " & val(BranchDC2.BoundText)
    End If
    
    If val(DepDC.BoundText) <> 0 Then
        MySQL = MySQL & " and  DepID = " & val(DepDC.BoundText)
    End If
    
    If val(ResEmpDC.BoundText) <> 0 Then
        MySQL = MySQL & " and  AppliedID = " & val(ResEmpDC.BoundText)
    End If
    
    If val(PayTypeCB.ListIndex) <> -1 Then
        MySQL = MySQL & " and  PaymentType = " & val(PayTypeCB.ListIndex)
    End If
    
    If SuppRd.value = True Then
        MySQL = MySQL & " and  SuppOrOther = 0"
        If val(SuppDC.BoundText) <> 0 Then
            MySQL = MySQL & " and  SuppID = " & val(SuppDC.BoundText)
        End If
    ElseIf OtherRd.value = True Then
        MySQL = MySQL & " and  SuppOrOther = 1"
        If OtherNameTxt.text <> "" Then
            MySQL = MySQL & "and OtherName like N'%" & OtherNameTxt.text & "%'"
        End If
    End If
    
    If val(CopyPriceTxt.text) <> 0 Then
        MySQL = MySQL & " and  CopyPrice = " & val(CopyPriceTxt.text)
    End If
    
    If CompName.text <> "" Then
          MySQL = MySQL & "and CompetitionName like N'%" & CompName.text & "%'"
    End If
    
    If CompNoTxt.text <> "" Then
          MySQL = MySQL & "and CompetitionNumber like N'%" & CompNoTxt.text & "%'"
    End If
    
    MySQL = MySQL & " Order By ID  "
    Set rs = New ADODB.Recordset
    rs.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    
    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Me.lbll(10).Caption = "‰ ÌÃ… «·»ÕÀ  =  ’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            Me.lbll(10).Caption = "Search Results=0"
        End If
        MsgBox "⁄›Ê« ...·« ÌÊÃœ »Ì«‰«   ‰«”» ‘—Êÿ «·»ÕÀ ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
        Cmd_Click (1)
        Exit Sub
    Else
        With Me.VSFlexGrid5
            .Clear flexClearScrollable, flexClearEverything
            .rows = .FixedRows
            .rows = rs.RecordCount + .FixedRows
           
            rs.MoveFirst
            
            For i = .FixedRows To .rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("ID")) = IIf(IsNull(rs("ID").value), "", rs("ID").value)
                .TextMatrix(i, .ColIndex("RecordDate")) = IIf(IsNull(rs("RecordDate").value), "", rs("RecordDate").value)
                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("Branch")) = IIf(IsNull(rs("branch_name").value), "", rs("branch_name").value)
                    .TextMatrix(i, .ColIndex("ResEmp")) = IIf(IsNull(rs("AppliedName").value), "", rs("AppliedName").value)
                    .TextMatrix(i, .ColIndex("Dep")) = IIf(IsNull(rs("DepartmentName").value), "", rs("DepartmentName").value)
                Else
                    .TextMatrix(i, .ColIndex("Branch")) = IIf(IsNull(rs("branch_namee").value), "", rs("branch_namee").value)
                    .TextMatrix(i, .ColIndex("ResEmp")) = IIf(IsNull(rs("AppliedNamee").value), "", rs("AppliedNamee").value)
                    .TextMatrix(i, .ColIndex("Dep")) = IIf(IsNull(rs("DepartmentNamee").value), "", rs("DepartmentNamee").value)
                End If
                .TextMatrix(i, .ColIndex("CompDate")) = IIf(IsNull(rs("CompetDate").value), "", rs("CompetDate").value)
                .TextMatrix(i, .ColIndex("SubDate")) = IIf(IsNull(rs("SubmittingDate").value), "", rs("SubmittingDate").value)
                .TextMatrix(i, .ColIndex("OpenEnvDate")) = IIf(IsNull(rs("OpenEnvelopeDate").value), "", rs("OpenEnvelopeDate").value)
                .TextMatrix(i, .ColIndex("CopyPrice")) = IIf(IsNull(rs("CopyPrice").value), "", rs("CopyPrice").value)
                .TextMatrix(i, .ColIndex("CompName")) = IIf(IsNull(rs("CompetitionName").value), "", rs("CompetitionName").value)
                .TextMatrix(i, .ColIndex("CompNo")) = IIf(IsNull(rs("CompetitionNumber").value), "", rs("CompetitionNumber").value)
                
                If rs("SuppOrOther").value = 1 Then
                    .TextMatrix(i, .ColIndex("ReqPart")) = IIf(IsNull(rs("OtherName").value), "", rs("OtherName").value)
                ElseIf rs("SuppOrOther").value = 0 Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        .TextMatrix(i, .ColIndex("ReqPart")) = IIf(IsNull(rs("CusName").value), "", rs("CusName").value)
                    Else
                        .TextMatrix(i, .ColIndex("ReqPart")) = IIf(IsNull(rs("CusNamee").value), "", rs("CusNamee").value)
                    End If
                End If
                
                If SystemOptions.UserInterface = ArabicInterface Then
                    If rs("PaymentType").value = 0 Then
                        .TextMatrix(i, .ColIndex("PayType")) = "‰ﬁœ"
                    ElseIf rs("PaymentType").value = 1 Then
                        .TextMatrix(i, .ColIndex("PayType")) = "‘Ìﬂ „’œﬁ"
                    ElseIf rs("PaymentType").value = 2 Then
                        .TextMatrix(i, .ColIndex("PayType")) = "ÕÊ«·Â »‰ﬂÌ… Õ”» «·„—›ﬁ"
                    Else
                        .TextMatrix(i, .ColIndex("PayType")) = " "
                    End If
                Else
                    If rs("PaymentType").value = 0 Then
                        .TextMatrix(i, .ColIndex("PayType")) = "Cash"
                    ElseIf rs("PaymentType").value = 1 Then
                        .TextMatrix(i, .ColIndex("PayType")) = "Certified Check"
                    ElseIf rs("PaymentType").value = 2 Then
                        .TextMatrix(i, .ColIndex("PayType")) = "Bank Transfer"
                    Else
                        .TextMatrix(i, .ColIndex("PayType")) = " "
                    End If
                End If
                
                
                rs.MoveNext
            Next i
            .AutoSize 0, .Cols - 1, False
        End With
    End If
End Sub
Public Sub GetUploadData()

    Dim sql As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
    Dim MySQL As String
    VSFlexGrid6.rows = VSFlexGrid6.FixedRows
    
    MySQL = " SELECT     dbo.TblOrderUpload.ID, dbo.TblOrderUpload.RecordDate,TblOrderUpload.ContainerNo, dbo.TblOrderUpload.EmpID, dbo.TblOrderUpload.LeaderName, dbo.TblEmployee.Emp_Name, "
    MySQL = MySQL & "                  dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee, dbo.TblOrderUpload.CustId1, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee,"
    MySQL = MySQL & "                  dbo.TblCustemers.Fullcode AS CusFullcode, dbo.TblOrderUpload.OrderNo, dbo.TblOrderUpload.CarID, dbo.TblCarsData.BoardNO, dbo.TblOrderUpload.CarID2,"
    MySQL = MySQL & "                  dbo.TblVendorCars.BoardNo AS BoardNo2, dbo.TblOrderUpload.CarType, dbo.TblOrderUpload.DrievType"
    MySQL = MySQL & " FROM         dbo.TblOrderUpload LEFT OUTER JOIN"
    MySQL = MySQL & "                  dbo.TblVendorCars ON dbo.TblOrderUpload.CarID2 = dbo.TblVendorCars.ID LEFT OUTER JOIN"
    MySQL = MySQL & "                  dbo.TblCarsData ON dbo.TblOrderUpload.CarID = dbo.TblCarsData.id LEFT OUTER JOIN"
    MySQL = MySQL & "                  dbo.TblCustemers ON dbo.TblOrderUpload.CustId1 = dbo.TblCustemers.CusID LEFT OUTER JOIN"
    MySQL = MySQL & "                  dbo.TblEmployee ON dbo.TblOrderUpload.EmpID = dbo.TblEmployee.Emp_ID"
     
    MySQL = MySQL & "  where (1 = 1)"
    MySQL = MySQL & "  and TblOrderUpload.ID Not In  (Select IsNull(BasedNo,0) from notes_all where NoteType = 370)"

    If val(Text1.text) <> 0 Then
          MySQL = MySQL & " and dbo.TblOrderUpload.ID >= " & val(Text1.text)
    End If
    
    If val(Text2.text) <> 0 Then
          MySQL = MySQL & " and dbo.TblOrderUpload.ID  <= " & val(Text2.text)
    End If
       If Trim(txtContainerNo.text) <> "" Then
          MySQL = MySQL & " and dbo.TblOrderUpload.ContainerNo  = '" & Trim(txtContainerNo.text) & "'"
    End If
    
    If Not IsNull(Me.DTPicker1.value) Then
        MySQL = MySQL & " AND dbo.TblOrderUpload.RecordDate >=" & SQLDate(Me.DTPicker1.value, True) & ""
    End If
    
    If Not IsNull(DTPicker2.value) Then
        MySQL = MySQL & " AND dbo.TblOrderUpload.RecordDate<=" & SQLDate(Me.DTPicker2.value, True) & ""
    End If
    
    
    If val(DBCboClientName1.BoundText) <> 0 Then
        MySQL = MySQL & " and   dbo.TblOrderUpload.CustId1 = " & val(DBCboClientName1.BoundText)
    End If

    If ChDrievType(0).value = True And val(DCEmployee.BoundText) <> 0 Then
        MySQL = MySQL & " and  dbo.TblOrderUpload.EmpID = " & val(DCEmployee.BoundText)
    End If
    If ChDrievType(1).value = True And TxtLeaderName.text <> "" Then
        MySQL = MySQL & "and dbo.TblOrderUpload.LeaderName like N'%" & TxtLeaderName.text & "%'"
    End If
    If TxtOrderNo.text <> "" Then
        MySQL = MySQL & "and dbo.TblOrderUpload.OrderNo like N'%" & TxtOrderNo.text & "%'"
    End If

    If val(DcbCar.BoundText) <> 0 Then
        MySQL = MySQL & " and  dbo.TblOrderUpload.CarID = " & val(DcbCar.BoundText)
    End If
    
    If val(DcbCar2.BoundText) <> 0 Then
        MySQL = MySQL & " and  dbo.TblOrderUpload.CarID2 = " & val(DcbCar2.BoundText)
    End If
    
    MySQL = MySQL & " Order By dbo.TblOrderUpload.ID  "
    Set rs = New ADODB.Recordset
    rs.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    
    If rs.BOF Or rs.EOF Then
        MsgBox "⁄›Ê« ...·« ÌÊÃœ »Ì«‰«   ‰«”» ‘—Êÿ «·»ÕÀ ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
        Cmd_Click (1)
        Exit Sub
    Else
        With Me.VSFlexGrid6
            .Clear flexClearScrollable, flexClearEverything
            .rows = .FixedRows
            .rows = rs.RecordCount + .FixedRows
           
            rs.MoveFirst
            
            For i = .FixedRows To .rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("ID")) = IIf(IsNull(rs("ID").value), "", rs("ID").value)
                .TextMatrix(i, .ColIndex("RecordDate")) = IIf(IsNull(rs("RecordDate").value), "", rs("RecordDate").value)
                .TextMatrix(i, .ColIndex("DrievType")) = IIf(IsNull(rs("DrievType").value), 0, rs("DrievType").value)
                .TextMatrix(i, .ColIndex("CarType")) = IIf(IsNull(rs("CarType").value), 0, rs("CarType").value)
                .TextMatrix(i, .ColIndex("ContainerNo")) = IIf(IsNull(rs("ContainerNo").value), "", rs("ContainerNo").value)
                
                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(rs("CusName").value), "", rs("CusName").value)
                    .TextMatrix(i, .ColIndex("LeaderName")) = IIf(IsNull(rs("Emp_Name").value), "", rs("Emp_Name").value)

                Else
                    .TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(rs("CusNamee").value), "", rs("CusNamee").value)
                    .TextMatrix(i, .ColIndex("LeaderName")) = IIf(IsNull(rs("Emp_Namee").value), "", rs("Emp_Namee").value)
                End If
                .TextMatrix(i, .ColIndex("OrderNo")) = IIf(IsNull(rs("OrderNo").value), "", rs("OrderNo").value)
                If val(.TextMatrix(i, .ColIndex("DrievType"))) = 1 Then
                .TextMatrix(i, .ColIndex("LeaderName")) = IIf(IsNull(rs("LeaderName").value), "", rs("LeaderName").value)
                End If
                .TextMatrix(i, .ColIndex("BoardNO")) = IIf(IsNull(rs("BoardNO").value), "", rs("BoardNO").value)
                If val(.TextMatrix(i, .ColIndex("CarType"))) = 1 Then
                .TextMatrix(i, .ColIndex("BoardNO")) = IIf(IsNull(rs("BoardNo2").value), "", rs("BoardNo2").value)
                End If
                rs.MoveNext
            Next i
            .AutoSize 0, .Cols - 1, False
        End With
    End If
End Sub
Public Sub GetUploadData2()

    Dim sql As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
    Dim MySQL As String
    VSFlexGrid6.rows = VSFlexGrid6.FixedRows
    
    MySQL = " SELECT     dbo.TblClientTransContr.ID, dbo.TblClientTransContr.FromDate RecordDate,dbo.TblClientTransContr.ToDate RecordDate2, dbo.TblClientTransContr.CompID,  "
    MySQL = MySQL & "                    dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee,"
    MySQL = MySQL & "                  dbo.TblCustemers.Fullcode AS CusFullcode, "
    MySQL = MySQL & "                  dbo.TblVendorCars.BoardNo AS BoardNo2"
    MySQL = MySQL & " FROM         dbo.TblClientTransContr LEFT OUTER JOIN"
    MySQL = MySQL & "                  dbo.TblVendorCars ON dbo.TblClientTransContr.VehicleType = dbo.TblVendorCars.ID LEFT OUTER JOIN"
    
    MySQL = MySQL & "                  dbo.TblCustemers ON dbo.TblClientTransContr.CompID = dbo.TblCustemers.CusID "
    
    MySQL = MySQL & "  where (1 = 1)"

    If val(Text1.text) <> 0 Then
          MySQL = MySQL & " and dbo.TblClientTransContr.ID >= " & val(Text1.text)
    End If
    
    If val(Text2.text) <> 0 Then
          MySQL = MySQL & " and dbo.TblClientTransContr.ID  <= " & val(Text2.text)
    End If
    
    
    If Not IsNull(Me.DTPicker1.value) Then
        MySQL = MySQL & " AND dbo.TblClientTransContr.FromDate >=" & SQLDate(Me.DTPicker1.value, True) & ""
    End If
    
    If Not IsNull(DTPicker2.value) Then
        MySQL = MySQL & " AND dbo.TblClientTransContr.FromDate<=" & SQLDate(Me.DTPicker2.value, True) & ""
    End If
    
    
    If val(DBCboClientName1.BoundText) <> 0 Then
        MySQL = MySQL & " and   dbo.TblClientTransContr.CompID = " & val(DBCboClientName1.BoundText)
    End If

   
    If val(DcbCar.BoundText) <> 0 Then
        MySQL = MySQL & " and  dbo.TblClientTransContr.VehicleType = " & val(DcbCar.BoundText)
    End If
    
    
    
    MySQL = MySQL & " Order By dbo.TblClientTransContr.ID  "
    Set rs = New ADODB.Recordset
    rs.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    
    If rs.BOF Or rs.EOF Then
        MsgBox "⁄›Ê« ...·« ÌÊÃœ »Ì«‰«   ‰«”» ‘—Êÿ «·»ÕÀ ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
        Cmd_Click (1)
        Exit Sub
    Else
        With Me.VSFlexGrid6
            .Clear flexClearScrollable, flexClearEverything
            .rows = .FixedRows
            .rows = rs.RecordCount + .FixedRows
           
            rs.MoveFirst
            
            For i = .FixedRows To .rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("ID")) = IIf(IsNull(rs("ID").value), "", rs("ID").value)
                .TextMatrix(i, .ColIndex("RecordDate")) = IIf(IsNull(rs("RecordDate").value), "", rs("RecordDate").value)
                '.TextMatrix(i, .ColIndex("DrievType")) = IIf(IsNull(rs("DrievType").value), 0, rs("DrievType").value)
                '.TextMatrix(i, .ColIndex("CarType")) = IIf(IsNull(rs("CarType").value), 0, rs("CarType").value)
                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(rs("CusName").value), "", rs("CusName").value)
                    '.TextMatrix(i, .ColIndex("LeaderName")) = IIf(IsNull(rs("Emp_Name").value), "", rs("Emp_Name").value)

                Else
                    .TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(rs("CusNamee").value), "", rs("CusNamee").value)
                    '.TextMatrix(i, .ColIndex("LeaderName")) = IIf(IsNull(rs("Emp_Namee").value), "", rs("Emp_Namee").value)
                End If
              '  .TextMatrix(i, .ColIndex("OrderNo")) = IIf(IsNull(rs("OrderNo").value), "", rs("OrderNo").value)
              '  If val(.TextMatrix(i, .ColIndex("DrievType"))) = 1 Then
              '  .TextMatrix(i, .ColIndex("LeaderName")) = IIf(IsNull(rs("LeaderName").value), "", rs("LeaderName").value)
             '   End If
                .TextMatrix(i, .ColIndex("BoardNO")) = IIf(IsNull(rs("BoardNo2").value), "", rs("BoardNo2").value)
             '   If val(.TextMatrix(i, .ColIndex("CarType"))) = 1 Then
             '   .TextMatrix(i, .ColIndex("BoardNO")) = IIf(IsNull(rs("BoardNo2").value), "", rs("BoardNo2").value)
             '   End If
                rs.MoveNext
            Next i
            .AutoSize 0, .Cols - 1, False
        End With
    End If
End Sub



Private Sub VSFlexGrid5_Click()
   If BankInx = 606 Then
   FrmTypeExchange.TxtOrderNo.text = val(VSFlexGrid5.TextMatrix(VSFlexGrid5.Row, 1))
   Else
   FrmBankPledge4.Retrive val(VSFlexGrid5.TextMatrix(VSFlexGrid5.Row, 1))
End If
End Sub

Private Sub VSFlexGrid6_Click()
If BankInx = 701 Then
FrmOrderUpload.FindRec val(VSFlexGrid6.TextMatrix(VSFlexGrid6.Row, VSFlexGrid6.ColIndex("id")))
ElseIf BankInx = 702 Then
FrmTravelTransactions.TxtBasedNo.text = val(VSFlexGrid6.TextMatrix(VSFlexGrid6.Row, VSFlexGrid6.ColIndex("id")))
ElseIf BankInx = 703 Then
FrmTravelTransactions.TxtBasedNo.text = val(VSFlexGrid6.TextMatrix(VSFlexGrid6.Row, VSFlexGrid6.ColIndex("id")))
FrmTravelTransactions.RetriveOrders2 val(VSFlexGrid6.TextMatrix(VSFlexGrid6.Row, VSFlexGrid6.ColIndex("id")))
End If
End Sub
