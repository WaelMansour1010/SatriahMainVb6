VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form FrmLC 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9660
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   20940
   HelpContextID   =   580
   Icon            =   "FrmLC.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   9660
   ScaleWidth      =   20940
   Visible         =   0   'False
   Begin C1SizerLibCtl.C1Elastic EleMain 
      Height          =   9660
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   20940
      _cx             =   36936
      _cy             =   17039
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
      AutoSizeChildren=   7
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
      Begin C1SizerLibCtl.C1Tab C1Tab1 
         Height          =   8205
         Left            =   60
         TabIndex        =   1
         Top             =   510
         Width           =   20880
         _cx             =   36830
         _cy             =   14473
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
         Caption         =   "«·»Ì«‰«  «·«”«”Ì…|„’«—Ì› «·› Õ  |«·›Ê« Ì— «·„«·Ì…|revised bond amount|ﬁ—Ê÷ «·«⁄ „«œ« |Refinance|acceptance advice"
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
         Flags(2)        =   2
         Flags(4)        =   2
         Begin C1SizerLibCtl.C1Elastic C1Elastic2 
            Height          =   7785
            Left            =   21825
            TabIndex        =   43
            TabStop         =   0   'False
            Top             =   45
            Width           =   20790
            _cx             =   36671
            _cy             =   13732
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
            Begin VSFlex8UCtl.VSFlexGrid Fg 
               Height          =   6090
               Left            =   0
               TabIndex        =   44
               ToolTipText     =   "«÷€ÿ „— Ì‰ ·› Õ «·›« Ê—…"
               Top             =   0
               Width           =   10335
               _cx             =   18230
               _cy             =   10742
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
               BackColorBkg    =   -2147483636
               BackColorAlternate=   16777215
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
               Cols            =   7
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmLC.frx":038A
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
         Begin C1SizerLibCtl.C1Elastic C1Elastic1 
            Height          =   7785
            Left            =   21525
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   45
            Width           =   20790
            _cx             =   36671
            _cy             =   13732
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
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
            Begin VB.TextBox TxtNoteID 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   18420
               RightToLeft     =   -1  'True
               TabIndex        =   217
               Top             =   6960
               Width           =   1170
            End
            Begin VB.TextBox txtNoteIDRowId 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   14220
               RightToLeft     =   -1  'True
               TabIndex        =   216
               Text            =   "Text4"
               Top             =   7080
               Visible         =   0   'False
               Width           =   2835
            End
            Begin VB.TextBox txtNoteID2RowId 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   14190
               RightToLeft     =   -1  'True
               TabIndex        =   215
               Text            =   "Text3"
               Top             =   6720
               Visible         =   0   'False
               Width           =   2925
            End
            Begin VB.TextBox txtNoteIDOpenRowId 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   14160
               RightToLeft     =   -1  'True
               TabIndex        =   214
               Text            =   "Text2"
               Top             =   6360
               Visible         =   0   'False
               Width           =   2775
            End
            Begin VB.CommandButton cmdAddLine 
               Caption         =   "Õ–› ”ÿ—"
               Height          =   465
               Index           =   6
               Left            =   17820
               RightToLeft     =   -1  'True
               TabIndex        =   213
               Top             =   5910
               Width           =   2235
            End
            Begin VB.CommandButton cmdAddLine 
               Caption         =   "«÷«›… ”ÿ—"
               Height          =   465
               Index           =   1
               Left            =   17850
               RightToLeft     =   -1  'True
               TabIndex        =   210
               Top             =   5400
               Width           =   2235
            End
            Begin VB.CommandButton Command6 
               Caption         =   "Õ–› ﬁÌœ «·› Õ"
               Height          =   375
               Left            =   9180
               RightToLeft     =   -1  'True
               TabIndex        =   206
               Top             =   6570
               Visible         =   0   'False
               Width           =   2310
            End
            Begin VB.TextBox txtNoteIDOpen 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   390
               Left            =   7590
               Locked          =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   205
               Top             =   6540
               Visible         =   0   'False
               Width           =   1575
            End
            Begin VB.TextBox txtNoteSerialOpen 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   390
               Left            =   7590
               Locked          =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   204
               Top             =   6120
               Width           =   1575
            End
            Begin VB.CommandButton Command4 
               Caption         =   "ÿ»«⁄Â «·ﬁÌœ"
               Height          =   465
               Left            =   9180
               RightToLeft     =   -1  'True
               TabIndex        =   203
               Top             =   6090
               Width           =   2295
            End
            Begin VB.CommandButton Command3 
               Caption         =   "≈‰‘«¡ ﬁÌœ «·› Õ"
               Height          =   465
               Left            =   9210
               RightToLeft     =   -1  'True
               TabIndex        =   202
               Top             =   5640
               Visible         =   0   'False
               Width           =   2280
            End
            Begin VB.TextBox txtMarginTotal3 
               Alignment       =   1  'Right Justify
               Height          =   525
               Left            =   13500
               RightToLeft     =   -1  'True
               TabIndex        =   199
               Top             =   5370
               Width           =   2535
            End
            Begin VB.Frame Frame1 
               Caption         =   "LG"
               Height          =   1500
               Left            =   30
               RightToLeft     =   -1  'True
               TabIndex        =   64
               Top             =   -60
               Width           =   13785
               Begin VB.TextBox txtLGExpPeriod 
                  Alignment       =   1  'Right Justify
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Left            =   6660
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   71
                  Top             =   540
                  Width           =   1170
               End
               Begin VB.TextBox txtCostDay 
                  Alignment       =   1  'Right Justify
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Left            =   10470
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   70
                  Top             =   960
                  Width           =   1230
               End
               Begin VB.TextBox txtCostLGYear 
                  Alignment       =   1  'Right Justify
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Left            =   7860
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   69
                  Top             =   960
                  Width           =   1230
               End
               Begin VB.TextBox txtLGExpPeriodEnd 
                  Alignment       =   1  'Right Justify
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Left            =   3480
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   68
                  Top             =   510
                  Width           =   1080
               End
               Begin VB.TextBox txtLGExpPeriodLast 
                  Alignment       =   1  'Right Justify
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Left            =   60
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   67
                  Top             =   510
                  Width           =   1080
               End
               Begin VB.TextBox txtCostLGYearLast 
                  Alignment       =   1  'Right Justify
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Left            =   3480
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   66
                  Top             =   960
                  Width           =   1110
               End
               Begin VB.TextBox txtOPenValue2 
                  Alignment       =   1  'Right Justify
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   9000
                  RightToLeft     =   -1  'True
                  TabIndex        =   65
                  Top             =   180
                  Width           =   2715
               End
               Begin MSComCtl2.DTPicker txtGuaranteeDate 
                  Height          =   435
                  Left            =   10470
                  TabIndex        =   72
                  Top             =   540
                  Width           =   1275
                  _ExtentX        =   2249
                  _ExtentY        =   767
                  _Version        =   393216
                  Format          =   228392961
                  CurrentDate     =   38784
               End
               Begin MSComCtl2.DTPicker txtLGExpiryDate 
                  Height          =   435
                  Left            =   7860
                  TabIndex        =   73
                  Top             =   540
                  Width           =   1275
                  _ExtentX        =   2249
                  _ExtentY        =   767
                  _Version        =   393216
                  Format          =   228392961
                  CurrentDate     =   38784
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " «—ÌŒ «·÷„«‰"
                  Height          =   300
                  Index           =   38
                  Left            =   11520
                  RightToLeft     =   -1  'True
                  TabIndex        =   81
                  Top             =   585
                  Width           =   1440
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "LG Expiry Date"
                  Height          =   300
                  Index           =   41
                  Left            =   9240
                  RightToLeft     =   -1  'True
                  TabIndex        =   80
                  Top             =   615
                  Width           =   1110
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " ﬂ·›… «·ÌÊ„"
                  Height          =   300
                  Index           =   42
                  Left            =   11400
                  RightToLeft     =   -1  'True
                  TabIndex        =   79
                  Top             =   1020
                  Width           =   1515
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " ﬂ·›… «·”‰… «·Õ«·Ì…"
                  Height          =   300
                  Index           =   43
                  Left            =   8910
                  RightToLeft     =   -1  'True
                  TabIndex        =   78
                  Top             =   1020
                  Width           =   1515
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "⁄œœ «·«Ì«„ «·„ »ﬁÌ… ·‰Â«Ì… «·”‰…"
                  Height          =   300
                  Index           =   44
                  Left            =   4620
                  RightToLeft     =   -1  'True
                  TabIndex        =   77
                  Top             =   555
                  Width           =   2040
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "⁄œœ «·«Ì«„ «·„ »ﬁÌ… »⁄œ ‰Â«Ì… «·”‰…"
                  Height          =   300
                  Index           =   45
                  Left            =   1170
                  RightToLeft     =   -1  'True
                  TabIndex        =   76
                  Top             =   555
                  Width           =   2280
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " ﬂ·›… «·«Ì«„ »⁄œ ‰Â«Ì… «·”‰…"
                  Height          =   300
                  Index           =   46
                  Left            =   4470
                  RightToLeft     =   -1  'True
                  TabIndex        =   75
                  Top             =   1020
                  Width           =   1995
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "ﬁÌ„… „’«—Ì› «·› Õ"
                  Height          =   330
                  Index           =   47
                  Left            =   11670
                  RightToLeft     =   -1  'True
                  TabIndex        =   74
                  Top             =   180
                  Width           =   1425
               End
            End
            Begin VB.ComboBox CboPaymentType 
               DataSource      =   "Adodc1"
               Height          =   315
               Left            =   1500
               RightToLeft     =   -1  'True
               TabIndex        =   45
               Top             =   5940
               Width           =   2715
            End
            Begin VB.TextBox txtOPenValue 
               Alignment       =   1  'Right Justify
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   1530
               RightToLeft     =   -1  'True
               TabIndex        =   32
               Top             =   5490
               Width           =   2715
            End
            Begin VB.Frame FraNote 
               Height          =   1665
               Left            =   4560
               RightToLeft     =   -1  'True
               TabIndex        =   22
               Top             =   7980
               Width           =   4155
               Begin VB.TextBox TxtChequeNumber 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   30
                  RightToLeft     =   -1  'True
                  TabIndex        =   23
                  Top             =   840
                  Width           =   2685
               End
               Begin MSComCtl2.DTPicker DtpChequeDueDate 
                  Height          =   315
                  Left            =   30
                  TabIndex        =   24
                  Top             =   1140
                  Width           =   2685
                  _ExtentX        =   4736
                  _ExtentY        =   556
                  _Version        =   393216
                  Format          =   228392961
                  CurrentDate     =   39614
               End
               Begin MSDataListLib.DataCombo DcboBankName 
                  Height          =   288
                  Left            =   0
                  TabIndex        =   25
                  Top             =   480
                  Width           =   2712
                  _ExtentX        =   4789
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DcboBox 
                  Height          =   315
                  Left            =   0
                  TabIndex        =   26
                  Top             =   120
                  Width           =   2715
                  _ExtentX        =   4789
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«”„ «·Œ“‰…"
                  Height          =   285
                  Index           =   16
                  Left            =   2790
                  RightToLeft     =   -1  'True
                  TabIndex        =   30
                  Top             =   180
                  Width           =   1215
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "«”„ «·»‰ﬂ"
                  Height          =   285
                  Index           =   17
                  Left            =   2790
                  RightToLeft     =   -1  'True
                  TabIndex        =   29
                  Top             =   510
                  Width           =   1215
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   "—ﬁ„ «·‘Ìﬂ"
                  Height          =   285
                  Index           =   18
                  Left            =   2760
                  RightToLeft     =   -1  'True
                  TabIndex        =   28
                  Top             =   840
                  Width           =   1215
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BackStyle       =   0  'Transparent
                  Caption         =   " «—ÌŒ «·≈” Õﬁ«ﬁ"
                  Height          =   285
                  Index           =   19
                  Left            =   2820
                  RightToLeft     =   -1  'True
                  TabIndex        =   27
                  Top             =   1140
                  Width           =   1215
               End
            End
            Begin MSDataListLib.DataCombo cmbAccount 
               Height          =   315
               Left            =   30
               TabIndex        =   48
               Top             =   7230
               Width           =   4215
               _ExtentX        =   7435
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo cmbAccountExpProject 
               Height          =   315
               Left            =   30
               TabIndex        =   59
               Top             =   6690
               Visible         =   0   'False
               Width           =   4215
               _ExtentX        =   7435
               _ExtentY        =   556
               _Version        =   393216
               Enabled         =   0   'False
               Style           =   2
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VSFlex8UCtl.VSFlexGrid GrdMargin3 
               Height          =   3600
               Left            =   90
               TabIndex        =   198
               ToolTipText     =   "«÷€ÿ „— Ì‰ ·› Õ «·›« Ê—…"
               Top             =   1650
               Width           =   20085
               _cx             =   35428
               _cy             =   6350
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
               BackColorBkg    =   -2147483636
               BackColorAlternate=   16777215
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
               Cols            =   27
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmLC.frx":04CD
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
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "ﬁÌ„… „’«—Ì› «·› Õ "
               Height          =   330
               Index           =   1
               Left            =   16140
               RightToLeft     =   -1  'True
               TabIndex        =   200
               Top             =   5520
               Width           =   1425
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„’—Ê› «·„‘—Ê⁄"
               Height          =   375
               Index           =   48
               Left            =   4530
               RightToLeft     =   -1  'True
               TabIndex        =   60
               Top             =   6630
               Visible         =   0   'False
               Width           =   1125
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·Õ”«» "
               Height          =   375
               Index           =   33
               Left            =   4530
               RightToLeft     =   -1  'True
               TabIndex        =   49
               Top             =   7170
               Width           =   1125
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "ﬁÌ„… „’«—Ì› «·› Õ"
               Height          =   330
               Index           =   14
               Left            =   4140
               RightToLeft     =   -1  'True
               TabIndex        =   33
               Top             =   5460
               Width           =   1425
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "ÿ—Ìﬁ… «·œ›⁄"
               Height          =   195
               Index           =   15
               Left            =   4260
               RightToLeft     =   -1  'True
               TabIndex        =   31
               Top             =   5940
               Width           =   1245
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   7785
            Index           =   2
            Left            =   45
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   45
            Width           =   20790
            _cx             =   36671
            _cy             =   13732
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
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   7785
               Index           =   1
               Left            =   0
               TabIndex        =   4
               TabStop         =   0   'False
               Top             =   0
               Width           =   20790
               _cx             =   36671
               _cy             =   13732
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
               Begin VB.TextBox Text1 
                  Alignment       =   1  'Right Justify
                  Height          =   330
                  Left            =   2400
                  RightToLeft     =   -1  'True
                  TabIndex        =   50
                  Top             =   -330
                  Width           =   2370
               End
               Begin VB.CommandButton Command5 
                  Caption         =   "≈‰‘«¡ ﬁÌœ «·«” Õﬁ«ﬁ"
                  Height          =   480
                  Left            =   18480
                  RightToLeft     =   -1  'True
                  TabIndex        =   47
                  Top             =   8250
                  Width           =   3405
               End
               Begin VB.Frame Fra 
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Õ«·… «·—’Ìœ «·√›  «ÕÏ"
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
                  Height          =   1350
                  Index           =   1
                  Left            =   18480
                  RightToLeft     =   -1  'True
                  TabIndex        =   34
                  Top             =   7905
                  Visible         =   0   'False
                  Width           =   4425
                  Begin VB.TextBox txtopening_balance_voucher_id 
                     Alignment       =   1  'Right Justify
                     Height          =   285
                     Left            =   0
                     RightToLeft     =   -1  'True
                     TabIndex        =   42
                     Top             =   0
                     Visible         =   0   'False
                     Width           =   615
                  End
                  Begin VB.OptionButton OptType 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "„œÌ‰"
                     Height          =   255
                     Index           =   0
                     Left            =   2190
                     RightToLeft     =   -1  'True
                     TabIndex        =   38
                     Top             =   210
                     Value           =   -1  'True
                     Width           =   765
                  End
                  Begin VB.OptionButton OptType 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "œ«∆‰"
                     Height          =   255
                     Index           =   1
                     Left            =   1320
                     RightToLeft     =   -1  'True
                     TabIndex        =   37
                     Top             =   210
                     Width           =   765
                  End
                  Begin VB.OptionButton OptType 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "€Ì— „Õœœ"
                     Height          =   255
                     Index           =   2
                     Left            =   330
                     RightToLeft     =   -1  'True
                     TabIndex        =   36
                     Top             =   210
                     Width           =   1005
                  End
                  Begin VB.TextBox TxtOpenBalance 
                     Alignment       =   2  'Center
                     Height          =   345
                     Left            =   360
                     RightToLeft     =   -1  'True
                     TabIndex        =   35
                     Top             =   510
                     Width           =   1365
                  End
                  Begin MSComCtl2.DTPicker Dtp 
                     Height          =   330
                     Left            =   360
                     TabIndex        =   39
                     Top             =   870
                     Width           =   1380
                     _ExtentX        =   2434
                     _ExtentY        =   582
                     _Version        =   393216
                     Enabled         =   0   'False
                     CalendarBackColor=   12648447
                     CustomFormat    =   "yyyy/M/d"
                     Format          =   175308803
                     CurrentDate     =   38718
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   " «—ÌŒ «· ”ÃÌ·"
                     Height          =   285
                     Index           =   24
                     Left            =   1800
                     RightToLeft     =   -1  'True
                     TabIndex        =   41
                     Top             =   930
                     Width           =   1215
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "ﬁÌ„… «·—’Ìœ "
                     Height          =   255
                     Index           =   23
                     Left            =   1740
                     RightToLeft     =   -1  'True
                     TabIndex        =   40
                     Top             =   540
                     Width           =   1275
                  End
               End
               Begin VB.OptionButton Option2 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«Œ Ì«— ’‰›"
                  BeginProperty Font 
                     Name            =   "MS Reference Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   390
                  Left            =   23490
                  RightToLeft     =   -1  'True
                  TabIndex        =   17
                  Top             =   2865
                  Value           =   -1  'True
                  Width           =   1575
               End
               Begin VB.OptionButton Option1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "⁄—÷ ﬂ«›Â «·«’‰«›"
                  BeginProperty Font 
                     Name            =   "MS Reference Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   390
                  Left            =   25065
                  RightToLeft     =   -1  'True
                  TabIndex        =   16
                  Top             =   2865
                  Width           =   2250
               End
               Begin VB.CheckBox ChKauto 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·Ì"
                  Enabled         =   0   'False
                  Height          =   255
                  Left            =   17340
                  RightToLeft     =   -1  'True
                  TabIndex        =   15
                  Top             =   7905
                  Visible         =   0   'False
                  Width           =   2145
               End
               Begin VB.TextBox txtType 
                  Alignment       =   1  'Right Justify
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   15780
                  RightToLeft     =   -1  'True
                  TabIndex        =   14
                  Text            =   "0"
                  Top             =   -240
                  Visible         =   0   'False
                  Width           =   570
               End
               Begin VB.TextBox TXTTblLCID 
                  Alignment       =   1  'Right Justify
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   345
                  Left            =   9900
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   13
                  Top             =   -300
                  Visible         =   0   'False
                  Width           =   2265
               End
               Begin VB.CheckBox Check1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "⁄—÷ "
                  Height          =   285
                  Left            =   17430
                  RightToLeft     =   -1  'True
                  TabIndex        =   11
                  Top             =   7770
                  Width           =   3105
               End
               Begin VB.TextBox txtid 
                  Alignment       =   1  'Right Justify
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Index           =   0
                  Left            =   -5280
                  RightToLeft     =   -1  'True
                  TabIndex        =   8
                  Top             =   12300
                  Width           =   2880
               End
               Begin VB.TextBox TxtModFlg 
                  Alignment       =   1  'Right Justify
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   345
                  Left            =   8055
                  RightToLeft     =   -1  'True
                  TabIndex        =   5
                  Top             =   -300
                  Visible         =   0   'False
                  Width           =   2925
               End
               Begin VSFlex8Ctl.VSFlexGrid Grid 
                  Height          =   4350
                  Left            =   21810
                  TabIndex        =   6
                  Top             =   3270
                  Width           =   13755
                  _cx             =   24262
                  _cy             =   7673
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
                  Rows            =   1
                  Cols            =   19
                  FixedRows       =   1
                  FixedCols       =   2
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmLC.frx":0949
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
               Begin MSDataListLib.DataCombo dcopr 
                  Height          =   315
                  Left            =   22095
                  TabIndex        =   9
                  Top             =   2775
                  Width           =   4035
                  _ExtentX        =   7117
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
                  Left            =   22515
                  TabIndex        =   10
                  Top             =   2055
                  Width           =   2235
                  _ExtentX        =   3942
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
               Begin MSDataListLib.DataCombo Dcterm 
                  Height          =   315
                  Left            =   21330
                  TabIndex        =   12
                  Top             =   1080
                  Width           =   4530
                  _ExtentX        =   7990
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
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   405
                  Index           =   20
                  Left            =   21600
                  TabIndex        =   18
                  Top             =   2865
                  Width           =   915
                  _ExtentX        =   1614
                  _ExtentY        =   714
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   "≈÷«›…"
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
                  ButtonImage     =   "FrmLC.frx":0C12
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   405
                  Index           =   21
                  Left            =   21270
                  TabIndex        =   19
                  Top             =   2865
                  Width           =   945
                  _ExtentX        =   1667
                  _ExtentY        =   714
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
                  ButtonImage     =   "FrmLC.frx":0FAC
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin MSDataListLib.DataCombo dcitems 
                  Height          =   315
                  Left            =   21885
                  TabIndex        =   20
                  Top             =   2865
                  Width           =   5955
                  _ExtentX        =   10504
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
               Begin C1SizerLibCtl.C1Elastic Ele 
                  Height          =   3930
                  Index           =   13
                  Left            =   150
                  TabIndex        =   82
                  TabStop         =   0   'False
                  Top             =   60
                  Width           =   20580
                  _cx             =   36301
                  _cy             =   6932
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
                  Begin VB.Frame Frame2 
                     Height          =   450
                     Left            =   6630
                     RightToLeft     =   -1  'True
                     TabIndex        =   161
                     Top             =   3435
                     Width           =   2685
                     Begin VB.OptionButton optTypeLCLG 
                        Alignment       =   1  'Right Justify
                        Caption         =   "«⁄ „«œ"
                        Height          =   195
                        Index           =   0
                        Left            =   1230
                        RightToLeft     =   -1  'True
                        TabIndex        =   163
                        Top             =   150
                        Width           =   1035
                     End
                     Begin VB.OptionButton optTypeLCLG 
                        Alignment       =   1  'Right Justify
                        Caption         =   "÷„«‰"
                        Height          =   195
                        Index           =   1
                        Left            =   180
                        RightToLeft     =   -1  'True
                        TabIndex        =   162
                        Top             =   150
                        Width           =   915
                     End
                  End
                  Begin VB.TextBox TxtItemQty 
                     Alignment       =   1  'Right Justify
                     Height          =   0
                     Index           =   2
                     Left            =   15
                     MaxLength       =   5
                     TabIndex        =   90
                     Top             =   1785
                     Width           =   0
                  End
                  Begin VB.TextBox TxtItemPrice 
                     Alignment       =   1  'Right Justify
                     Height          =   0
                     Index           =   2
                     Left            =   15
                     MaxLength       =   5
                     TabIndex        =   89
                     Top             =   1785
                     Visible         =   0   'False
                     Width           =   0
                  End
                  Begin VB.TextBox TXTLCNO 
                     Alignment       =   1  'Right Justify
                     Height          =   315
                     Left            =   13320
                     RightToLeft     =   -1  'True
                     TabIndex        =   88
                     Top             =   270
                     Width           =   4410
                  End
                  Begin VB.TextBox txtName 
                     Alignment       =   1  'Right Justify
                     Height          =   405
                     Left            =   10500
                     RightToLeft     =   -1  'True
                     TabIndex        =   87
                     Top             =   1230
                     Width           =   7230
                  End
                  Begin VB.TextBox txtNameE 
                     Alignment       =   1  'Right Justify
                     Height          =   405
                     Left            =   345
                     RightToLeft     =   -1  'True
                     TabIndex        =   86
                     Top             =   1230
                     Width           =   6630
                  End
                  Begin VB.TextBox TXTBank2 
                     Alignment       =   1  'Right Justify
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   600
                     Left            =   10545
                     RightToLeft     =   -1  'True
                     TabIndex        =   85
                     Top             =   2265
                     Width           =   7185
                  End
                  Begin VB.TextBox txtProjectName 
                     Alignment       =   1  'Right Justify
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   600
                     Left            =   345
                     RightToLeft     =   -1  'True
                     TabIndex        =   84
                     Top             =   2265
                     Width           =   6630
                  End
                  Begin VB.TextBox txt_Currency_rate 
                     Alignment       =   1  'Right Justify
                     Height          =   345
                     Left            =   10575
                     RightToLeft     =   -1  'True
                     TabIndex        =   83
                     Text            =   "1"
                     Top             =   3105
                     Width           =   1110
                  End
                  Begin MSDataListLib.DataCombo DCPreFix 
                     Height          =   315
                     Left            =   10500
                     TabIndex        =   91
                     Top             =   270
                     Width           =   2820
                     _ExtentX        =   4974
                     _ExtentY        =   556
                     _Version        =   393216
                     BackColor       =   16777215
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin MSDataListLib.DataCombo DCLC 
                     Height          =   315
                     Left            =   345
                     TabIndex        =   92
                     Top             =   210
                     Width           =   6630
                     _ExtentX        =   11695
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
                  Begin MSDataListLib.DataCombo DCBank 
                     Height          =   315
                     Left            =   10500
                     TabIndex        =   93
                     Top             =   780
                     Width           =   7230
                     _ExtentX        =   12753
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
                  Begin MSDataListLib.DataCombo DcBranch 
                     Height          =   315
                     Left            =   345
                     TabIndex        =   94
                     Top             =   750
                     Width           =   6630
                     _ExtentX        =   11695
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
                  Begin MSDataListLib.DataCombo DBCboClientName 
                     Height          =   315
                     Left            =   10515
                     TabIndex        =   95
                     Top             =   1770
                     Width           =   7215
                     _ExtentX        =   12726
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
                  Begin MSDataListLib.DataCombo DataCombo2 
                     DataSource      =   "Adodc1"
                     Height          =   315
                     Left            =   345
                     TabIndex        =   96
                     Top             =   1770
                     Width           =   6630
                     _ExtentX        =   11695
                     _ExtentY        =   556
                     _Version        =   393216
                     ListField       =   ""
                     BoundColumn     =   ""
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
                  Begin MSDataListLib.DataCombo DCCountry 
                     Height          =   315
                     Left            =   14205
                     TabIndex        =   97
                     Top             =   3105
                     Width           =   3525
                     _ExtentX        =   6218
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
                  Begin MSDataListLib.DataCombo DCCUrrency 
                     Height          =   315
                     Left            =   11775
                     TabIndex        =   98
                     Top             =   3105
                     Width           =   1680
                     _ExtentX        =   2963
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
                  Begin MSDataListLib.DataCombo DboParentAccount 
                     Height          =   315
                     Left            =   0
                     TabIndex        =   154
                     Top             =   3180
                     Visible         =   0   'False
                     Width           =   9345
                     _ExtentX        =   16484
                     _ExtentY        =   556
                     _Version        =   393216
                     Style           =   2
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·—ﬁ„"
                     Height          =   270
                     Index           =   4
                     Left            =   17505
                     RightToLeft     =   -1  'True
                     TabIndex        =   110
                     Top             =   270
                     Width           =   2730
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·‰Ê⁄"
                     Height          =   240
                     Index           =   6
                     Left            =   7065
                     RightToLeft     =   -1  'True
                     TabIndex        =   109
                     Top             =   315
                     Width           =   2430
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·«”„ ⁄—»Ì"
                     Height          =   300
                     Index           =   25
                     Left            =   17355
                     RightToLeft     =   -1  'True
                     TabIndex        =   108
                     Top             =   1305
                     Width           =   2880
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·»‰ﬂ"
                     Height          =   270
                     Index           =   9
                     Left            =   18600
                     RightToLeft     =   -1  'True
                     TabIndex        =   107
                     Top             =   780
                     Width           =   1635
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·›—⁄"
                     Height          =   300
                     Index           =   29
                     Left            =   6705
                     RightToLeft     =   -1  'True
                     TabIndex        =   106
                     Top             =   810
                     Width           =   2790
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·«”„ «‰Ã·Ì“Ì"
                     Height          =   285
                     Index           =   52
                     Left            =   6705
                     RightToLeft     =   -1  'True
                     TabIndex        =   105
                     Top             =   1290
                     Width           =   2790
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«”„ «·„Ê—œ/«·⁄„Ì·"
                     Height          =   450
                     Index           =   13
                     Left            =   17970
                     RightToLeft     =   -1  'True
                     TabIndex        =   104
                     Top             =   2490
                     Width           =   2265
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·„Ê—œ/«·⁄„Ì·"
                     Height          =   315
                     Index           =   0
                     Left            =   18420
                     RightToLeft     =   -1  'True
                     TabIndex        =   103
                     Top             =   1875
                     Width           =   1860
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«”„ «·„‘—Ê⁄"
                     Height          =   360
                     Index           =   40
                     Left            =   7275
                     RightToLeft     =   -1  'True
                     TabIndex        =   102
                     Top             =   2475
                     Width           =   2220
                  End
                  Begin VB.Label Label1 
                     Alignment       =   1  'Right Justify
                     BackStyle       =   0  'Transparent
                     Caption         =   "«”„ «·„‘—Ê⁄"
                     Height          =   345
                     Index           =   0
                     Left            =   7515
                     RightToLeft     =   -1  'True
                     TabIndex        =   101
                     Top             =   1845
                     Width           =   1980
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·œÊ·Â"
                     Height          =   300
                     Index           =   12
                     Left            =   19020
                     RightToLeft     =   -1  'True
                     TabIndex        =   100
                     Top             =   3180
                     Width           =   1215
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·⁄„·Â"
                     Height          =   240
                     Index           =   10
                     Left            =   12885
                     RightToLeft     =   -1  'True
                     TabIndex        =   99
                     Top             =   3150
                     Width           =   1245
                  End
               End
               Begin C1SizerLibCtl.C1Elastic Ele 
                  Height          =   1170
                  Index           =   0
                  Left            =   11970
                  TabIndex        =   111
                  TabStop         =   0   'False
                  Top             =   4125
                  Width           =   8745
                  _cx             =   15425
                  _cy             =   2064
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
                  Begin VB.TextBox TxtItemPrice 
                     Alignment       =   1  'Right Justify
                     Height          =   0
                     Index           =   0
                     Left            =   15
                     MaxLength       =   5
                     TabIndex        =   117
                     Top             =   555
                     Visible         =   0   'False
                     Width           =   0
                  End
                  Begin VB.TextBox TxtItemQty 
                     Alignment       =   1  'Right Justify
                     Height          =   0
                     Index           =   0
                     Left            =   15
                     MaxLength       =   5
                     TabIndex        =   116
                     Top             =   555
                     Width           =   0
                  End
                  Begin VB.TextBox TXTValue 
                     Alignment       =   1  'Right Justify
                     Height          =   435
                     Left            =   4695
                     RightToLeft     =   -1  'True
                     TabIndex        =   115
                     Top             =   30
                     Width           =   2340
                  End
                  Begin VB.TextBox txtBondAmt 
                     Alignment       =   1  'Right Justify
                     Height          =   435
                     Left            =   2700
                     RightToLeft     =   -1  'True
                     TabIndex        =   114
                     Top             =   45
                     Width           =   1800
                  End
                  Begin VB.TextBox txtPercentV 
                     Alignment       =   1  'Right Justify
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   435
                     Left            =   4725
                     RightToLeft     =   -1  'True
                     TabIndex        =   113
                     Top             =   555
                     Width           =   2310
                  End
                  Begin VB.TextBox txtAcceptianPeriod 
                     Alignment       =   1  'Right Justify
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   435
                     Left            =   2685
                     RightToLeft     =   -1  'True
                     TabIndex        =   112
                     Top             =   555
                     Width           =   1830
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·ﬁÌ„…"
                     Height          =   390
                     Index           =   28
                     Left            =   7005
                     RightToLeft     =   -1  'True
                     TabIndex        =   122
                     Top             =   90
                     Width           =   1500
                  End
                  Begin VB.Label lbl 
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "Bond Amt"
                     Height          =   420
                     Index           =   34
                     Left            =   135
                     RightToLeft     =   -1  'True
                     TabIndex        =   121
                     Top             =   90
                     Width           =   960
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "‰”»… «· √„Ì‰"
                     Height          =   255
                     Index           =   31
                     Left            =   6975
                     RightToLeft     =   -1  'True
                     TabIndex        =   120
                     Top             =   645
                     Width           =   1575
                  End
                  Begin VB.Label lbl 
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "Deferred Period"
                     Height          =   405
                     Index           =   39
                     Left            =   120
                     TabIndex        =   119
                     Top             =   615
                     Width           =   1500
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "%"
                     Height          =   435
                     Index           =   32
                     Left            =   2595
                     RightToLeft     =   -1  'True
                     TabIndex        =   118
                     Top             =   585
                     Width           =   2100
                  End
               End
               Begin C1SizerLibCtl.C1Elastic Ele 
                  Height          =   2235
                  Index           =   3
                  Left            =   12015
                  TabIndex        =   123
                  TabStop         =   0   'False
                  Top             =   5445
                  Width           =   8715
                  _cx             =   15372
                  _cy             =   3942
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
                  Begin VB.CommandButton Command7 
                     Height          =   255
                     Left            =   7080
                     RightToLeft     =   -1  'True
                     TabIndex        =   218
                     Top             =   1740
                     Width           =   735
                  End
                  Begin VB.TextBox TxtNoteID2 
                     Alignment       =   1  'Right Justify
                     Height          =   315
                     Left            =   -2010
                     RightToLeft     =   -1  'True
                     TabIndex        =   156
                     Top             =   0
                     Visible         =   0   'False
                     Width           =   2475
                  End
                  Begin VB.TextBox TxtItemQty 
                     Alignment       =   1  'Right Justify
                     Height          =   345
                     Index           =   1
                     Left            =   -990
                     MaxLength       =   5
                     TabIndex        =   127
                     Top             =   1065
                     Visible         =   0   'False
                     Width           =   1080
                  End
                  Begin VB.TextBox TxtItemPrice 
                     Alignment       =   1  'Right Justify
                     Height          =   0
                     Index           =   1
                     Left            =   15
                     MaxLength       =   5
                     TabIndex        =   126
                     Top             =   1065
                     Visible         =   0   'False
                     Width           =   0
                  End
                  Begin VB.TextBox TxtNoOfParcil 
                     Alignment       =   1  'Right Justify
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   420
                     Left            =   5280
                     RightToLeft     =   -1  'True
                     TabIndex        =   125
                     Top             =   15
                     Width           =   1875
                  End
                  Begin VB.CheckBox ChkLocked 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   " „ «·«€·«ﬁ"
                     Height          =   735
                     Left            =   120
                     RightToLeft     =   -1  'True
                     TabIndex        =   124
                     Top             =   -135
                     Width           =   1410
                  End
                  Begin MSComCtl2.DTPicker DPLastParcilDate 
                     Height          =   390
                     Left            =   1620
                     TabIndex        =   128
                     Top             =   75
                     Width           =   1920
                     _ExtentX        =   3387
                     _ExtentY        =   688
                     _Version        =   393216
                     Format          =   176816129
                     CurrentDate     =   38784
                  End
                  Begin MSComCtl2.DTPicker dbFromDate 
                     Height          =   330
                     Left            =   5265
                     TabIndex        =   129
                     Top             =   585
                     Width           =   1920
                     _ExtentX        =   3387
                     _ExtentY        =   582
                     _Version        =   393216
                     Format          =   230621185
                     CurrentDate     =   38784
                  End
                  Begin MSComCtl2.DTPicker DpCloseDate 
                     Height          =   360
                     Left            =   1620
                     TabIndex        =   130
                     Top             =   570
                     Width           =   1920
                     _ExtentX        =   3387
                     _ExtentY        =   635
                     _Version        =   393216
                     Format          =   230621185
                     CurrentDate     =   38784
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "⁄œœ «·‘Õ‰« "
                     Height          =   165
                     Index           =   20
                     Left            =   7125
                     RightToLeft     =   -1  'True
                     TabIndex        =   134
                     Top             =   75
                     Width           =   1455
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   " «—ÌŒ «Œ— ‘Õ‰"
                     Height          =   195
                     Index           =   22
                     Left            =   3405
                     RightToLeft     =   -1  'True
                     TabIndex        =   133
                     Top             =   75
                     Width           =   1725
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   " «—ÌŒ «·› Õ"
                     Height          =   180
                     Index           =   5
                     Left            =   7125
                     RightToLeft     =   -1  'True
                     TabIndex        =   132
                     Top             =   615
                     Width           =   1410
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   " «—ÌŒ «·«‰ Â«¡"
                     Height          =   315
                     Index           =   21
                     Left            =   2955
                     RightToLeft     =   -1  'True
                     TabIndex        =   131
                     Top             =   645
                     Width           =   1890
                  End
               End
               Begin C1SizerLibCtl.C1Elastic Ele 
                  Height          =   2175
                  Index           =   4
                  Left            =   210
                  TabIndex        =   135
                  TabStop         =   0   'False
                  Top             =   5475
                  Width           =   11640
                  _cx             =   20532
                  _cy             =   3836
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
                  Begin VB.TextBox TxtItemPrice 
                     Alignment       =   1  'Right Justify
                     Height          =   0
                     Index           =   3
                     Left            =   15
                     MaxLength       =   5
                     TabIndex        =   137
                     Top             =   1005
                     Visible         =   0   'False
                     Width           =   0
                  End
                  Begin VB.TextBox TxtItemQty 
                     Alignment       =   1  'Right Justify
                     Height          =   0
                     Index           =   3
                     Left            =   15
                     MaxLength       =   5
                     TabIndex        =   136
                     Top             =   1005
                     Width           =   0
                  End
                  Begin MSDataListLib.DataCombo cmbAccountMarginParent 
                     Height          =   315
                     Left            =   45
                     TabIndex        =   138
                     Top             =   615
                     Width           =   9360
                     _ExtentX        =   16510
                     _ExtentY        =   556
                     _Version        =   393216
                     Style           =   2
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin MSDataListLib.DataCombo cmbAccountLGParent 
                     Height          =   315
                     Left            =   45
                     TabIndex        =   139
                     Top             =   1170
                     Visible         =   0   'False
                     Width           =   9360
                     _ExtentX        =   16510
                     _ExtentY        =   556
                     _Version        =   393216
                     Style           =   2
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin MSDataListLib.DataCombo cmbAccountAcceptanceParent 
                     Height          =   315
                     Left            =   45
                     TabIndex        =   140
                     Top             =   1680
                     Width           =   9360
                     _ExtentX        =   16510
                     _ExtentY        =   556
                     _Version        =   393216
                     Style           =   2
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin MSDataListLib.DataCombo cmbAccountExpensParent 
                     Height          =   315
                     Left            =   45
                     TabIndex        =   155
                     Top             =   90
                     Width           =   9360
                     _ExtentX        =   16510
                     _ExtentY        =   556
                     _Version        =   393216
                     Style           =   2
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "Acceptance"
                     Height          =   285
                     Index           =   36
                     Left            =   7545
                     RightToLeft     =   -1  'True
                     TabIndex        =   144
                     Top             =   1725
                     Width           =   3735
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "Õ”«» «·÷„«‰"
                     Height          =   300
                     Index           =   26
                     Left            =   7470
                     RightToLeft     =   -1  'True
                     TabIndex        =   143
                     Top             =   1260
                     Visible         =   0   'False
                     Width           =   3810
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "Õ”«» «·ﬂ«‘ „«—Ã‰"
                     Height          =   285
                     Index           =   8
                     Left            =   7365
                     RightToLeft     =   -1  'True
                     TabIndex        =   142
                     Top             =   705
                     Width           =   3915
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "Õ”«» «·„’—Ê›"
                     Height          =   285
                     Index           =   27
                     Left            =   7425
                     RightToLeft     =   -1  'True
                     TabIndex        =   141
                     Top             =   165
                     Width           =   3855
                  End
               End
               Begin C1SizerLibCtl.C1Elastic Ele 
                  Height          =   1245
                  Index           =   6
                  Left            =   150
                  TabIndex        =   145
                  TabStop         =   0   'False
                  Top             =   4125
                  Width           =   11715
                  _cx             =   20664
                  _cy             =   2196
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
                  Begin VB.TextBox TxtItemQty 
                     Alignment       =   1  'Right Justify
                     Height          =   0
                     Index           =   4
                     Left            =   60
                     MaxLength       =   5
                     TabIndex        =   150
                     Top             =   585
                     Width           =   0
                  End
                  Begin VB.TextBox TxtItemPrice 
                     Alignment       =   1  'Right Justify
                     Height          =   0
                     Index           =   4
                     Left            =   60
                     MaxLength       =   5
                     TabIndex        =   149
                     Top             =   585
                     Visible         =   0   'False
                     Width           =   0
                  End
                  Begin VB.TextBox txtGuaranteeNo 
                     Alignment       =   1  'Right Justify
                     Height          =   420
                     Left            =   8790
                     RightToLeft     =   -1  'True
                     TabIndex        =   148
                     Top             =   75
                     Width           =   1110
                  End
                  Begin VB.TextBox TXtPrimaryInvoiceNo 
                     Alignment       =   1  'Right Justify
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   420
                     Left            =   3690
                     RightToLeft     =   -1  'True
                     TabIndex        =   147
                     Top             =   90
                     Width           =   2040
                  End
                  Begin VB.TextBox txtRemarks 
                     Alignment       =   1  'Right Justify
                     Height          =   480
                     Left            =   705
                     MultiLine       =   -1  'True
                     RightToLeft     =   -1  'True
                     ScrollBars      =   2  'Vertical
                     TabIndex        =   146
                     Top             =   615
                     Width           =   9195
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "—ﬁ„ «·÷„«‰"
                     Height          =   315
                     Index           =   37
                     Left            =   9900
                     RightToLeft     =   -1  'True
                     TabIndex        =   153
                     Top             =   135
                     Width           =   1350
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "—ﬁ„ «„— «·‘—«¡ / «·⁄ﬁœ"
                     Height          =   330
                     Index           =   11
                     Left            =   5280
                     RightToLeft     =   -1  'True
                     TabIndex        =   152
                     Top             =   135
                     Width           =   2490
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "‘—Êÿ «· ”·Ì„"
                     Height          =   465
                     Index           =   3
                     Left            =   9915
                     RightToLeft     =   -1  'True
                     TabIndex        =   151
                     Top             =   735
                     Width           =   1425
                  End
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·—ﬁ„"
                  Height          =   225
                  Index           =   35
                  Left            =   -405
                  RightToLeft     =   -1  'True
                  TabIndex        =   51
                  Top             =   45
                  Width           =   1485
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·—ﬁ„"
                  Height          =   90
                  Index           =   7
                  Left            =   19095
                  RightToLeft     =   -1  'True
                  TabIndex        =   7
                  Top             =   3390
                  Width           =   1455
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic3 
            Height          =   7785
            Left            =   22425
            TabIndex        =   52
            TabStop         =   0   'False
            Top             =   45
            Width           =   20790
            _cx             =   36671
            _cy             =   13732
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
            Begin VSFlex8UCtl.VSFlexGrid GrdMargin 
               Height          =   6180
               Left            =   60
               TabIndex        =   159
               ToolTipText     =   "«÷€ÿ „— Ì‰ ·› Õ «·›« Ê—…"
               Top             =   60
               Width           =   20025
               _cx             =   35322
               _cy             =   10901
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
               BackColorBkg    =   -2147483636
               BackColorAlternate=   16777215
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
               Cols            =   21
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmLC.frx":1546
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
         Begin C1SizerLibCtl.C1Elastic C1Elastic4 
            Height          =   7785
            Index           =   0
            Left            =   22125
            TabIndex        =   53
            TabStop         =   0   'False
            Top             =   45
            Width           =   20790
            _cx             =   36671
            _cy             =   13732
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
            Begin VB.CommandButton cmdAddLine 
               Caption         =   "«÷«›… ”ÿ—"
               Height          =   465
               Index           =   0
               Left            =   10770
               RightToLeft     =   -1  'True
               TabIndex        =   209
               Top             =   7050
               Width           =   2235
            End
            Begin VB.TextBox txtTotalBondHistory 
               Alignment       =   1  'Right Justify
               Height          =   525
               Left            =   2490
               RightToLeft     =   -1  'True
               TabIndex        =   160
               Top             =   7080
               Width           =   5085
            End
            Begin VB.TextBox txtMarginTotal 
               Alignment       =   1  'Right Justify
               Height          =   615
               Left            =   5730
               RightToLeft     =   -1  'True
               TabIndex        =   158
               Top             =   7710
               Visible         =   0   'False
               Width           =   5085
            End
            Begin VB.CommandButton CmdCreateV2 
               Caption         =   "≈‰‘«¡ «·ﬁÌœ "
               Height          =   435
               Index           =   0
               Left            =   8130
               RightToLeft     =   -1  'True
               TabIndex        =   58
               Top             =   7110
               Visible         =   0   'False
               Width           =   2280
            End
            Begin VSFlex8UCtl.VSFlexGrid GrdBondHistory 
               Height          =   6330
               Left            =   -30
               TabIndex        =   157
               ToolTipText     =   "«÷€ÿ „— Ì‰ ·› Õ «·›« Ê—…"
               Top             =   660
               Width           =   20775
               _cx             =   36645
               _cy             =   11165
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
               BackColorBkg    =   -2147483636
               BackColorAlternate=   16777215
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
               FormatString    =   $"FrmLC.frx":18D4
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
               RightToLeft     =   0   'False
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
         Begin C1SizerLibCtl.C1Elastic C1Elastic4 
            Height          =   7785
            Index           =   1
            Left            =   22725
            TabIndex        =   54
            TabStop         =   0   'False
            Top             =   45
            Width           =   20790
            _cx             =   36671
            _cy             =   13732
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
            Begin VB.CommandButton cmdAddLine 
               Caption         =   "Õ–› ”ÿ—"
               Height          =   465
               Index           =   5
               Left            =   12840
               RightToLeft     =   -1  'True
               TabIndex        =   212
               Top             =   7200
               Width           =   2235
            End
            Begin VB.CommandButton cmdAddLine 
               Caption         =   "«÷«›… ”ÿ—"
               Height          =   465
               Index           =   2
               Left            =   10500
               RightToLeft     =   -1  'True
               TabIndex        =   208
               Top             =   7200
               Width           =   2235
            End
            Begin VB.CommandButton CmdCreateV2 
               Caption         =   "≈‰‘«¡ «·ﬁÌœ "
               Height          =   465
               Index           =   1
               Left            =   9585
               RightToLeft     =   -1  'True
               TabIndex        =   57
               Top             =   10260
               Width           =   2235
            End
            Begin VB.TextBox txtMarginTotal2 
               Alignment       =   1  'Right Justify
               Height          =   525
               Left            =   3270
               RightToLeft     =   -1  'True
               TabIndex        =   55
               Top             =   7230
               Width           =   4935
            End
            Begin VSFlex8UCtl.VSFlexGrid GrdMargin2 
               Height          =   6900
               Left            =   0
               TabIndex        =   56
               ToolTipText     =   "«÷€ÿ „— Ì‰ ·› Õ «·›« Ê—…"
               Top             =   240
               Width           =   19695
               _cx             =   34740
               _cy             =   12171
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
               BackColorBkg    =   -2147483636
               BackColorAlternate=   16777215
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
               Cols            =   39
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmLC.frx":1B06
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
         Begin C1SizerLibCtl.C1Elastic C1Elastic4 
            Height          =   7785
            Index           =   2
            Left            =   23025
            TabIndex        =   61
            TabStop         =   0   'False
            Top             =   45
            Width           =   20790
            _cx             =   36671
            _cy             =   13732
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
            Begin VB.CommandButton cmdAddLine 
               Caption         =   "Õ–› ”ÿ—"
               Height          =   465
               Index           =   3
               Left            =   14040
               RightToLeft     =   -1  'True
               TabIndex        =   211
               Top             =   7230
               Width           =   2235
            End
            Begin VB.CommandButton cmdAddLine 
               Caption         =   "«÷«›… ”ÿ—"
               Height          =   465
               Index           =   4
               Left            =   11820
               RightToLeft     =   -1  'True
               TabIndex        =   207
               Top             =   7230
               Width           =   2235
            End
            Begin VB.TextBox txtMarginTotal4 
               Alignment       =   1  'Right Justify
               Height          =   525
               Left            =   2985
               RightToLeft     =   -1  'True
               TabIndex        =   63
               Top             =   7170
               Width           =   4995
            End
            Begin VB.CommandButton CmdCreateV2 
               Caption         =   "≈‰‘«¡ «·ﬁÌœ "
               Height          =   465
               Index           =   2
               Left            =   8655
               RightToLeft     =   -1  'True
               TabIndex        =   62
               Top             =   7230
               Visible         =   0   'False
               Width           =   2235
            End
            Begin VSFlex8UCtl.VSFlexGrid GrdMargin4 
               Height          =   6900
               Left            =   60
               TabIndex        =   201
               ToolTipText     =   "«÷€ÿ „— Ì‰ ·› Õ «·›« Ê—…"
               Top             =   90
               Width           =   20580
               _cx             =   36301
               _cy             =   12171
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
               BackColorBkg    =   -2147483636
               BackColorAlternate=   16777215
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
               Cols            =   39
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmLC.frx":21F4
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
      Begin C1SizerLibCtl.C1Elastic EltCont 
         Height          =   825
         Left            =   60
         TabIndex        =   164
         TabStop         =   0   'False
         Top             =   8730
         Width           =   20880
         _cx             =   36830
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
         Begin VB.CommandButton CmdCreateV 
            Caption         =   "≈‰‘«¡ «·ﬁÌœ "
            Height          =   465
            Left            =   8100
            RightToLeft     =   -1  'True
            TabIndex        =   172
            Top             =   30
            Visible         =   0   'False
            Width           =   2280
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Õ–› «·ﬁÌœ "
            Height          =   465
            Left            =   3990
            RightToLeft     =   -1  'True
            TabIndex        =   171
            Top             =   30
            Width           =   1380
         End
         Begin VB.TextBox TxtNoteSerial 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   465
            Left            =   0
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   170
            Top             =   0
            Width           =   3255
         End
         Begin VB.CommandButton Command9 
            Caption         =   "ÿ»«⁄Â «·ﬁÌœ"
            Height          =   465
            Left            =   8070
            RightToLeft     =   -1  'True
            TabIndex        =   169
            Top             =   450
            Width           =   2325
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Õ–› ﬁÌœ «·«€·«ﬁ"
            Height          =   375
            Left            =   10800
            RightToLeft     =   -1  'True
            TabIndex        =   168
            Top             =   510
            Width           =   2070
         End
         Begin VB.CommandButton cmdPrintEntryClose 
            Caption         =   "ÿ»«⁄Â ﬁÌœ «·«€·«ﬁ"
            Height          =   420
            Left            =   12900
            RightToLeft     =   -1  'True
            TabIndex        =   167
            Top             =   480
            Width           =   1485
         End
         Begin VB.TextBox TxtNoteSerial2 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   390
            Left            =   14370
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   166
            Top             =   510
            Width           =   1575
         End
         Begin VB.CommandButton cmdCloseLC 
            Caption         =   "«€·«ﬁ «·÷„«‰"
            Height          =   405
            Left            =   16020
            RightToLeft     =   -1  'True
            TabIndex        =   165
            Top             =   510
            Width           =   1215
         End
         Begin ImpulseButton.ISButton btnQuery 
            Height          =   330
            Left            =   13020
            TabIndex        =   173
            TabStop         =   0   'False
            ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
            Top             =   90
            Visible         =   0   'False
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   582
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "»ÕÀ"
            BackColor       =   14737632
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
            ButtonImage     =   "FrmLC.frx":28F2
            ColorButton     =   14737632
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnUpdate 
            Height          =   330
            Left            =   12765
            TabIndex        =   174
            TabStop         =   0   'False
            ToolTipText     =   " ÕœÌÀ ﬁ«⁄œ… «·»Ì«‰« "
            Top             =   225
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
            ButtonImage     =   "FrmLC.frx":2C8C
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnPrint 
            Height          =   285
            Left            =   13965
            TabIndex        =   175
            TabStop         =   0   'False
            Top             =   150
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
            ButtonImage     =   "FrmLC.frx":3026
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   495
            Index           =   0
            Left            =   7140
            TabIndex        =   176
            Top             =   480
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   873
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
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   495
            Index           =   1
            Left            =   6240
            TabIndex        =   177
            Top             =   480
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   873
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
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   495
            Index           =   2
            Left            =   5400
            TabIndex        =   178
            Top             =   480
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   873
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
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton Cmd 
            CausesValidation=   0   'False
            Height          =   495
            Index           =   3
            Left            =   4395
            TabIndex        =   179
            Top             =   480
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   873
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
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   495
            Index           =   4
            Left            =   3360
            TabIndex        =   180
            Top             =   480
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   873
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
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton Cmd 
            CausesValidation=   0   'False
            Height          =   495
            Index           =   6
            Left            =   480
            TabIndex        =   181
            Top             =   480
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   873
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
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   495
            Index           =   5
            Left            =   2430
            TabIndex        =   182
            Top             =   480
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   873
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
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin ALLButtonS.ALLButton CmdRemove 
            Height          =   375
            Left            =   9120
            TabIndex        =   183
            Tag             =   "Delete Row"
            Top             =   0
            Visible         =   0   'False
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   661
            BTYPE           =   3
            TX              =   "Õ–› ”ÿ—"
            ENAB            =   0   'False
            BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            COLTYPE         =   2
            FOCUSR          =   -1  'True
            BCOL            =   0
            BCOLO           =   0
            FCOL            =   255
            FCOLO           =   255
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "FrmLC.frx":33C0
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   492
            Index           =   7
            Left            =   1440
            TabIndex        =   184
            Top             =   480
            Width           =   768
            _ExtentX        =   1349
            _ExtentY        =   873
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
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin MSDataListLib.DataCombo DCboUserName 
            Height          =   315
            Left            =   5490
            TabIndex        =   185
            Top             =   60
            Width           =   1905
            _ExtentX        =   3360
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   435
            Index           =   11
            Left            =   10830
            TabIndex        =   186
            Top             =   0
            Width           =   2010
            _ExtentX        =   3545
            _ExtentY        =   767
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
         Begin MSComCtl2.DTPicker dbTodate 
            Height          =   405
            Left            =   14340
            TabIndex        =   187
            Top             =   90
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   714
            _Version        =   393216
            Format          =   223477761
            CurrentDate     =   38784
         End
         Begin VB.Label LabCurrRec 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   4920
            RightToLeft     =   -1  'True
            TabIndex        =   192
            Top             =   240
            Width           =   1515
         End
         Begin VB.Label LabCountRec 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   1560
            RightToLeft     =   -1  'True
            TabIndex        =   191
            Top             =   225
            Width           =   1740
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "—ﬁ„ «·ﬁÌœ"
            Height          =   390
            Index           =   35
            Left            =   2745
            RightToLeft     =   -1  'True
            TabIndex        =   190
            Top             =   120
            Width           =   1125
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "Õ—— »Ê«”ÿ…"
            Height          =   435
            Index           =   30
            Left            =   17670
            RightToLeft     =   -1  'True
            TabIndex        =   189
            Top             =   240
            Visible         =   0   'False
            Width           =   1320
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «·€·ﬁ"
            Height          =   285
            Index           =   2
            Left            =   15975
            RightToLeft     =   -1  'True
            TabIndex        =   188
            Top             =   135
            Width           =   1020
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   480
         Index           =   5
         Left            =   0
         TabIndex        =   193
         TabStop         =   0   'False
         Top             =   0
         Width           =   20940
         _cx             =   36936
         _cy             =   847
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
         Picture         =   "FrmLC.frx":33DC
         Caption         =   " › Õ «⁄ „«œ „” ‰œÌ  /  ÷„«‰ »‰ﬂÌ  "
         Align           =   1
         AutoSizeChildren=   0
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
         Begin ImpulseButton.ISButton XPBtnMove 
            Height          =   375
            Index           =   0
            Left            =   1695
            TabIndex        =   194
            Top             =   90
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   661
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
            ButtonImage     =   "FrmLC.frx":40B6
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
            Height          =   375
            Index           =   2
            Left            =   630
            TabIndex        =   195
            Top             =   90
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   661
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
            ButtonImage     =   "FrmLC.frx":4450
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
            Height          =   375
            Index           =   1
            Left            =   2220
            TabIndex        =   196
            Top             =   90
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   661
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
            ButtonImage     =   "FrmLC.frx":47EA
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
            Height          =   375
            Index           =   3
            Left            =   1155
            TabIndex        =   197
            Top             =   90
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   661
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
            ButtonImage     =   "FrmLC.frx":4B84
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
   Begin ImpulseButton.ISButton ISButton1 
      Height          =   345
      Left            =   3360
      TabIndex        =   3
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
      ButtonImage     =   "FrmLC.frx":4F1E
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   285
      Left            =   6450
      TabIndex        =   46
      Top             =   6840
      Width           =   2205
      _ExtentX        =   3889
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
End
Attribute VB_Name = "FrmLC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cSearchDCombo As clsDCboSearch
Dim BKGrndPic As ClsBackGroundPic
Dim net_value As Double
Dim net_value1 As Double
Dim My_SQL  As String
Dim StrSQL  As String
Dim rs As ADODB.Recordset
Dim s As String
Dim rsDummy As ADODB.Recordset
Dim maa_rs As ADODB.Recordset
Public LngRow As Long
Dim FirstPeriodDateInthisYear  As Date
Private Declare Function TextOut _
                Lib "gdi32" _
                Alias "TextOutA" (ByVal hDC As Long, _
                                  ByVal X As Long, _
                                  ByVal Y As Long, _
                                  ByVal lpString As String, _
                                  ByVal nCount As Long) As Long

'ma
Public Sub Search(ID As Integer)


   Set maa_rs = New ADODB.Recordset
 '  StrSQL = StrSQL + " SELECT *  From dbo.project_billl  where 1 =1"
 
 StrSQL = " SELECT *  From  tbllc  where  TblLCID=  " & ID & " Order by TblLCID "
   
    
   maa_rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If maa_rs.RecordCount < 1 Then
        '    XPTxtCurrent.Caption = 0
        '    XPTxtCount.Caption = 0
        Exit Sub
    End If
    If maa_rs.EOF Or maa_rs.BOF Then
        Exit Sub
    Else
    
    'rs.Find "TblLCID= & ID & , , adSearchForward, adBookmarkFirst"
    rs.Find "TblLCID=" & val(ID), , adSearchForward, adBookmarkFirst
    maaRetrive
    
    Retrive
End If
End Sub

Public Sub maaRetrive(Optional LCNO As String = "")
    Dim RsDev As ADODB.Recordset
    Dim StrSQL As String
    Dim i As Integer

    'On Error GoTo ErrTrap
    Grid.Clear flexClearScrollable, flexClearEverything
    Grid.rows = 1

    If maa_rs.RecordCount < 1 Then
        Exit Sub
    End If

    If maa_rs.EOF Or maa_rs.BOF Then
        Exit Sub
    Else

        If LCNO <> "" Then
            maa_rs.Find "LCNO='" & LCNO & "'", , adSearchForward, adBookmarkFirst

            If maa_rs.EOF Or maa_rs.BOF Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·« ÌÊÃœ «⁄ „«œ »Â–« «·—ﬁ„", vbCritical
                Else
                    MsgBox "Lc With This No Not Found", vbCritical
                End If

                Unload Me
                Exit Sub
            
            End If
        End If
    End If
 
    Me.TXTTblLCID.text = IIf(IsNull(maa_rs("TblLCID").value), "", maa_rs("TblLCID").value)
    Me.TxtLcNo.text = IIf(IsNull(maa_rs("LCNO").value), "", maa_rs("LCNO").value)
        Me.TxtName.text = IIf(IsNull(maa_rs("Name").value), "", maa_rs("Name").value)
        
        Me.TxtNameE.text = IIf(IsNull(maa_rs("Namee").value), "", maa_rs("Namee").value)
    DboParentAccount.BoundText = Get_Account_Parent_code(IIf(IsNull(maa_rs("Account_Code").value), "", Trim(maa_rs("Account_Code").value)))
    
    RetriveProformaInvoices TxtLcNo.text
  
    Me.DCLC.BoundText = IIf(IsNull(maa_rs("LCTyperId").value), "", maa_rs("LCTyperId").value)
    Me.Dcbank.BoundText = IIf(IsNull(maa_rs("BankId").value), "", maa_rs("BankId").value)
    Me.TXTBank2.text = IIf(IsNull(maa_rs("Bank2").value), "", maa_rs("Bank2").value)
    Me.TxtValue.text = IIf(Not IsNumeric(maa_rs("Value").value), 0, maa_rs("Value").value)
    Me.txtPercentV.text = IIf(Not IsNumeric(maa_rs("PercentV").value), 0, maa_rs("PercentV").value)
    
    Me.Dccurrency.BoundText = IIf(IsNull(maa_rs("CurrencyId").value), "", maa_rs("CurrencyId").value)
    Me.TXtPrimaryInvoiceNo.text = IIf(IsNull(maa_rs("PrimaryInvoiceNo").value), "", maa_rs("PrimaryInvoiceNo").value)
    Me.DCCountry.BoundText = IIf(IsNull(maa_rs("CountryId").value), "", maa_rs("CountryId").value)
  
    dbFromDate.value = IIf(IsNull(maa_rs("FromDate").value), Date, maa_rs("FromDate").value)
    dbTodate.value = IIf(IsNull(maa_rs("Todate").value), Date, maa_rs("Todate").value)

    DpCloseDate.value = IIf(IsNull(maa_rs("CloseDate").value), Date, maa_rs("CloseDate").value)
    DPLastParcilDate.value = IIf(IsNull(maa_rs("LastParcilDate").value), Date, maa_rs("LastParcilDate").value)
    Me.TxtNoOfParcil.text = IIf(Not IsNumeric(maa_rs("NoOfParcil").value), 0, maa_rs("NoOfParcil").value)

    DBCboClientName.BoundText = IIf(IsNull(maa_rs("VendorId").value), "", maa_rs("VendorId").value)

    TxtRemarks.text = IIf(IsNull(maa_rs("Remarks").value), 0, maa_rs("Remarks").value)

    If IsNull(maa_rs("Locked").value) Then
        ChkLocked.value = vbUnchecked
    Else

        If maa_rs("Locked").value = True Then
            ChkLocked.value = vbChecked
        Else
            ChkLocked.value = vbUnchecked
        End If

    End If

    '    rs("OpenBalanceDate").value = Me.Dtp.value

    txtopening_balance_voucher_id.text = IIf(IsNull(rs("opening_balance_voucher_id").value), "", maa_rs("opening_balance_voucher_id").value)
    Dim FirstPeriodDateInthisYear As Date

    If (IsNull(rs("OpenBalanceDate").value)) Then
        getFirstPeriodDateInthisYear FirstPeriodDateInthisYear

        Me.Dtp.value = FirstPeriodDateInthisYear

        '     Me.Dtp.Enabled = True
    Else
        
        Me.Dtp.value = maa_rs("OpenBalanceDate").value
        '     Me.Dtp.Enabled = False
    End If
    
    If Not IsNull(rs("OpenBalanceType").value) Then
        Me.TxtOpenBalance.text = IIf(IsNull(maa_rs("OpenBalance")), "", Trim(maa_rs("OpenBalance")))

        If maa_rs("OpenBalanceType").value = 0 Then
            OptType(0).value = True
            OptType_Click 0
        ElseIf maa_rs("OpenBalanceType").value = 1 Then
            OptType(1).value = True
            OptType_Click 1
        End If
        
    Else
        Me.TxtOpenBalance.text = 0
        Me.OptType(2).value = True
        OptType_Click 2
    End If
 
    Exit Sub
ErrTrap:
End Sub






Private Sub CboPayMentType_Click()
If CboPayMentType.ListIndex = 0 Then
    DcboBox.Enabled = True
    DcboBankName.Enabled = False
    TxtChequeNumber.Enabled = False
    DtpChequeDueDate.Enabled = False
    cmbAccount.Enabled = False
ElseIf CboPayMentType.ListIndex = 1 Then
     DcboBox.Enabled = False
    DcboBankName.Enabled = True
    TxtChequeNumber.Enabled = True
    DtpChequeDueDate.Enabled = True
    cmbAccount.Enabled = False
ElseIf CboPayMentType.ListIndex = 2 Then
     DcboBox.Enabled = False
    DcboBankName.Enabled = True
    TxtChequeNumber.Enabled = True
    DtpChequeDueDate.Enabled = True
    cmbAccount.Enabled = True
    FraNote.Enabled = False
    DcboBox.text = ""
    DcboBankName.text = ""
    TxtChequeNumber = ""
End If
s = "Select Account_Code from BanksData where BankId = " & val(Dcbank.BoundText)
Set rsDummy = New ADODB.Recordset
rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly
If Not rsDummy.EOF Then
    cmbAccount.BoundText = Trim(rsDummy!Account_code & "")
End If

End Sub

Private Sub cmdAddLine_Click(index As Integer)

Dim mNoteId As Long
Dim mNoteId2 As Long
Dim mNoteId3 As Long
Dim mmID As Long

Select Case index
Case 0
    GrdBondHistory.rows = GrdBondHistory.rows + 1
Case 1
    GrdMargin3.rows = GrdMargin3.rows + 1
Case 2
    GrdMargin2.rows = GrdMargin2.rows + 1
Case 4
    GrdMargin4.rows = GrdMargin4.rows + 1

Case 3
'     If GrdMargin4.row <> 0 Then
'          mNoteId = val(GrdMargin4.TextMatrix(GrdMargin4.row, GrdMargin4.ColIndex("NoteID")))
'          mNoteId2 = val(GrdMargin4.TextMatrix(GrdMargin4.row, GrdMargin4.ColIndex("NoteID2")))
'          mNoteId3 = val(GrdMargin4.TextMatrix(GrdMargin4.row, GrdMargin4.ColIndex("NoteID3")))
'          mmID = val(GrdMargin4.TextMatrix(GrdMargin4.row, GrdMargin4.ColIndex("ID")))
'
'
'            s = "Delete Notes where NoteId = " & mNoteId2
'            Cn.Execute s
'            s = "Delete Notes where NoteId = " & mNoteId3
'            Cn.Execute s
'            If CBool(GrdMargin4.ValueMatrix(GrdMargin4.row, GrdMargin4.ColIndex("IsOpenBalance"))) Then
'                s = "Delete Notes1 where NoteId = " & val(GrdMargin4.TextMatrix(GrdMargin4.row, GrdMargin4.ColIndex("NoteId")))
'                Cn.Execute s
'                s = "Delete DOUBLE_ENTREY_VOUCHERS1 where Notes_ID = " & val(GrdMargin4.TextMatrix(GrdMargin4.row, GrdMargin4.ColIndex("NoteId")))
'
'                Cn.Execute s
'            Else
'                s = "Delete Notes where NoteId = " & val(GrdMargin4.TextMatrix(GrdMargin4.row, GrdMargin4.ColIndex("NoteId")))
'                Cn.Execute s
'            End If
'            s = "Delete TBLLCMargin2 where Id = " & mmID
'            Cn.Execute s
'
'     End If
    If GrdMargin4.row > 0 Then
        GrdMargin4.RemoveItem GrdMargin4.row
    End If
   ' MsgBox " „ Õ–› «·”ÿ—"
Case 5
'     If GrdMargin2.row <> 0 Then
'          mNoteId = val(GrdMargin2.TextMatrix(GrdMargin2.row, GrdMargin2.ColIndex("NoteID")))
'          mNoteId2 = val(GrdMargin2.TextMatrix(GrdMargin2.row, GrdMargin2.ColIndex("NoteID2")))
'          mNoteId3 = val(GrdMargin2.TextMatrix(GrdMargin2.row, GrdMargin2.ColIndex("NoteID3")))
'          mmID = val(GrdMargin2.TextMatrix(GrdMargin2.row, GrdMargin2.ColIndex("ID")))
'
'
'            s = "Delete Notes where NoteId = " & mNoteId2
'            Cn.Execute s
'            s = "Delete Notes where NoteId = " & mNoteId3
'            Cn.Execute s
'
'            s = "Delete Notes where NoteId = " & mNoteId
'            Cn.Execute s
'
'            s = "Delete TBLLCMargin where Id = " & mmID
'            Cn.Execute s
'
'     End If
    If GrdMargin2.row > 0 Then
        GrdMargin2.RemoveItem GrdMargin2.row
    End If
Case 6
    If GrdMargin3.row > 0 Then
        GrdMargin3.RemoveItem GrdMargin3.row
    End If

    
   ' MsgBox " „ Õ–› «·”ÿ—"
End Select
End Sub

Private Sub cmdCloseLC_Click()



If val(TxtNoteSerial2.text) = 0 Then
        createVoucher True
       'FindRec val(TXTTblLCID.text)
      ' rs.Find "TblLCID=" & val(TXTTblLCID.text), , adSearchForward, adBookmarkFirst
       
    '   rs!ToDate = dbTodate.value
    '   rs!NoteID2 = val(TxtNoteID2)
    '   rs!NoteSerial2 = val(TxtNoteSerial2)
    '   rs("Locked").value = 1
       
     '  rs.update
       ChkLocked.value = vbChecked
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox " „ «‰‘«¡ «·ﬁÌœ"
            If val(TxtNoteID2) <> 0 Then
                cmdCloseLC.Enabled = False
                ChkLocked.Enabled = False
                CmdCreateV.Enabled = False
                Command9.Enabled = True
                Command2.Enabled = True
'                Cmd(2).Enabled = False
            Else
                cmdCloseLC.Enabled = True
                ChkLocked.Enabled = True

                cmdCloseLC.Enabled = True
                CmdCreateV.Enabled = True
'                Command9.Enabled = False
                Command2.Enabled = False
            End If
        Else
            MsgBox "Done"
        End If
End If
End Sub

Private Sub cmdPrintEntryClose_Click()
ShowGL_cc Me.TxtNoteSerial2.text, , 200
End Sub

Private Sub Command1_Click()
If Me.TxtModFlg.text = "R" Or Me.TxtModFlg.text = "" Then
Dim X As Integer
Dim Msg As String
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = " √ﬂÌœ Õ–› «·ﬁÌœ "
    Else
        Msg = "Confirm Delete  "
    End If
        X = MsgBox(Msg, vbCritical + vbYesNo)

      If X = vbYes Then
        StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(Me.TxtNoteID2.text)
        Cn.Execute StrSQL, , adExecuteNoRecords
        StrSQL = "Delete From Notes Where NoteID=" & val(Me.TxtNoteID2.text)
        Cn.Execute StrSQL, , adExecuteNoRecords
        'Cn.Execute " Update TblLC set NoteID2=null ,NoteSerial2=null where TblLCID=" & val(TXTTblLCID.text)
        rs!NoteID2 = Null
        rs!NoteSerial2 = Null
        rs.update
        TxtNoteSerial2 = ""
        TxtNoteID2 = ""
       ' rs.Requery
         FindRec val(TXTTblLCID.text)
         TxtModFlg.text = ""
         TxtNoteSerial2 = ""
          If SystemOptions.UserInterface = ArabicInterface Then
            Msg = " „  Õ–› «·ﬁÌœ   "
            
           
            If val(TxtNoteID2) <> 0 Then
                cmdCloseLC.Enabled = False
                Command9.Enabled = True
                Command2.Enabled = True
                Cmd(2).Enabled = False
                Cmd(1).Enabled = False
             Else
                cmdCloseLC.Enabled = True
'                Command9.Enabled = False
                Command2.Enabled = False
            End If
        Else
            Msg = " This voucher deleted  "
        End If
        MsgBox Msg
       End If
 End If


End Sub


Function createVoucher(Optional ByVal IsClose As Boolean = False, Optional ByVal notytype As Integer = 0)
Dim NoteID As Long
Dim NoteDate As Date
Dim NoteSerial As String
Dim Notevalue As Double
'Dim
Dim i
Dim des As String
des = "    Õ”«» «·" & TxtLcNo.text
Dim tablename As String
Dim Filedname As String
Dim NoteSerial1 As Long
Dim BranchID As Integer
Dim sql As String
Dim mRate  As Double

Dim mFieldNoteID As String
Dim mFieldNoteSerial As String
tablename = "TBLLC"
Dim mRowID As String

Filedname = "TblLCID"
NoteSerial1 = val(TXTTblLCID.text)

BranchID = val(dcBranch.BoundText)
mRate = val(txt_Currency_rate)
Dim RowIDField As String
'


' ⁄‰ „ﬂ«‰ Ê÷⁄ «·ÀÊ«»  ÊﬂÌ›Ì… «· —ﬁÌ„   Õ «Ã  Ê÷ÌÕ
' «” ›”«— Ê«∆·
' ·Œ»ÿ… ⁄‰œÏ ›Ï «·„”„Ì«  Ê«·‰Ê   «Ì»
If notytype = 0 Then
    If IsClose Then
        notytype = 22005
        If val(txtBondAmt) = 0 Then txtBondAmt = val(TxtValue)
        
        Notevalue = val(txtBondAmt) * mRate * val(txtPercentV) / 100
        
        'mAccNO = val(DboParentAccount.BoundText)
        NoteDate = (dbTodate.value)
    
    Else
        mRowID = txtNoteIDRowId
        notytype = 22001
        Notevalue = val(TxtValue) * val(txtPercentV)
        
        'mAccNO = val(DboParentAccount.BoundText)
        NoteDate = (dbFromDate.value)
    End If
ElseIf notytype = 22010 Then
        'notytype = 22001
        Notevalue = val(txtOPenValue) * mRate
        
        'mAccNO = val(DboParentAccount.BoundText)
        NoteDate = (dbFromDate.value)

End If


 NoteID = 0
 If IsClose Then
    mFieldNoteID = "NoteID2"
    mFieldNoteSerial = "NoteSerial2"
    NoteSerial = val(TxtNoteSerial2)
    mRowID = txtNoteID2RowId
    RowIDField = "NoteID2RowId"
 Else
 If notytype = 22001 Then
    mFieldNoteID = "NoteID"
    mFieldNoteSerial = "NoteSerial"
    NoteSerial = val(TxtNoteSerial)
    mRowID = txtNoteIDRowId
    RowIDField = "NoteIDRowId"
End If
 End If
 If notytype = 22010 Then
    mFieldNoteID = "NoteIDOpen"
    mFieldNoteSerial = "NoteSerialOpen"
    NoteSerial = val(txtNoteSerialOpen)
    mRowID = txtNoteIDOpenRowId
    RowIDField = "NoteIDOpenRowId"
  '  Cn.Execute "Delete Notes where NoteId = " & val(txtNoteIDOpen.text)
 End If

 
If Notevalue > 0 Then
    If InStr(mRowID, "{") Then
    Else
        mRowID = "{" & mRowID & "}"
    End If
    
    If val(NoteSerial) <> 0 Then
        s = "Select * from Notes where NoteSerial = " & val(NoteSerial) & " and isNull(TblLCID,0)  <> " & val(TXTTblLCID)
        s = "Select * from Notes where NoteSerial = " & val(NoteSerial) & " and RowID  <> '" & Trim(mRowID) & "'"
        Set rsDummy = New ADODB.Recordset
        rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly
        If Not rsDummy.EOF Then
            NoteSerial = 0
        End If
        
    End If
    
    
    
    
    CreateNotes NoteID, NoteDate, BranchID, notytype, Notevalue, NoteSerial, (NoteSerial1), tablename, Filedname, NoteSerial1, des, , TxtLcNo, mFieldNoteID, mFieldNoteSerial, , , val(TXTTblLCID), mRowID, RowIDField
    
If IsClose Then
    
    
    TxtNoteSerial2 = NoteSerial
    TxtNoteID2 = NoteSerial
 Else
 If notytype = 22001 Then
    
    TxtNoteSerial = NoteSerial
    TXTNoteID = NoteID
End If
 End If
 If notytype = 22010 Then
    mFieldNoteID = "NoteIDOpen"
    mFieldNoteSerial = "NoteSerialOpen"
    txtNoteSerialOpen = NoteSerial
    txtNoteIDOpen = NoteID
    
  '  Cn.Execute "Delete Notes where NoteId = " & val(txtNoteIDOpen.text)
 End If
 
    rs.Resync adAffectCurrent
    
    
    If notytype = 22010 Then
            txtNoteIDOpen.text = NoteID
            txtNoteSerialOpen.text = NoteSerial
            
            CREATE_VOUCHER_GE val(txtNoteIDOpen.text), BranchID, val(DCboUserName.BoundText), NoteDate, True, notytype
    '
           
            
           Exit Function
          ' rs.Resync adAffectCurrent
       '   rs.update
          ' rs.Requery
           
    Else
        If IsClose Then
            TxtNoteID2.text = NoteID
            TxtNoteSerial2.text = NoteSerial
            
            CREATE_VOUCHER_GE val(TxtNoteID2.text), BranchID, val(DCboUserName.BoundText), NoteDate, True
    '
         
            rs!locked = 1
            'rs!ToDate = SQLDate(dbTodate.value, True)
            rs!ToDate = dbTodate.value
            rs.update
            
            Exit Function
           ' rs.Resync adAffectCurrent
        Else
'            StrSQL = "delete From DOUBLE_ENTREY_VOUCHERS  where Notes_ID=" & val(TxtNoteID.text)
'            Cn.Execute StrSQL, , adExecuteNoRecords
            
            If notytype = 22001 Then
                TXTNoteID.text = NoteID
                TxtNoteSerial.text = NoteSerial
                CREATE_VOUCHER_GE val(TXTNoteID.text), BranchID, val(DCboUserName.BoundText), NoteDate, False, notytype
            End If
            
            
            
            If notytype <> 2201 And Not IsClose Then
            Else
                Exit Function
            End If
           '       rs!ToDate = dbTodate.value
           
           
            
            
    
           
           
         
    '
        End If
    End If
    
    

            
        
    
  ' rs.Resync adAffectCurrent
'
'
'    Cn.Execute StrSQL
     
  
 
End If

End Function
Public Function CREATE_VOUCHER_GE(general_noteid As Long, BranchID As Integer, UserID As Long _
, NoteDate As Date, Optional ByVal IsClose As Boolean = False, Optional ByVal notytype As Integer = 0)

'     StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & general_noteid
'    Cn.Execute StrSQL, , adExecuteNoRecords
    
    Dim i As Integer
    Dim sql As String
    Dim StoreID6 As Integer
    Dim rs2 As ADODB.Recordset
    Set rs2 = New ADODB.Recordset
    Dim Notevalue As Double
    Dim LngDevID As Long
    Dim Msg As String
    Dim StrAccountCodeDebt As String
    Dim StrAccountCodeCridet As String
    Dim X As Integer
   
    Dim rs As New ADODB.Recordset
    Dim notes_serial As String
    Dim notes_id As String
    Msg = "    Õ”«» " & TxtLcNo.text
    notes_id = general_noteid
    my_branch = val(dcBranch.BoundText)
    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
    Dim line_no As Integer
    line_no = 1
        Dim Notevalue2 As Double
    Dim Notevalue3 As Double
    Dim s As String
    Dim mRate As Double
    mRate = val(txt_Currency_rate)
 Dim mPercent As Double
    
    
    ' «” ›”«— Ê«∆·
    ' ﬁ„  »Ã·» «·Õ”«» «·„‰‘√ ›Ï «·«⁄ „«œ »Â–Â «·ÿ—Ìﬁ… ‰—ÃÊ «·„—«Ã⁄…
    Dim sqlS As String
    Dim rsAcc As New ADODB.Recordset
    
         
     
    sqlS = " Select Account_Code,Account_code2,Account_CodeMargin,AcceptAccount_Code,LCAccount_Code,AccountExpensCode from tblLc Where  TblLCID=" & val(TXTTblLCID.text)
    rsAcc.Open sqlS, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If rsAcc.RecordCount <> 0 Then
       StrAccountCodeDebt = rsAcc!Account_code & ""
    End If
    Notevalue = val(TxtValue.text) * mRate
'    If Notevalue > 0 Then
'
'       ' StrAccountCodeDebt = Trim(DboParentAccount.BoundText)
'        If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCodeDebt, Notevalue, 0, Msg & "    Õ”«»  «·«⁄ „«œ  ", val(notes_id), , , , dbFromDate.value, val(DCboUserName.BoundText), , , , , , CLng(mRate), , , setfoxy_Line, , , , , , , , , _
'        val(dcBranch.BoundText)) = False Then
'            GoTo ErrTrap
'        End If
'        StrAccountCodeCridet = get_bank_Account(Dcbank.BoundText, "Account_Code")
'        GetValueAddedAccount val(TxtValue.text), , StrAccountCodeCridet, 1, 21
'        line_no = line_no + 1
'        If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCodeCridet, Notevalue, 1, Msg & "    Õ”«»  «·»‰ﬂ ", val(notes_id), , , , dbFromDate.value, val(DCboUserName.BoundText), , , , , , CLng(mRate), , , setfoxy_Line, , , , , , , , , val(dcBranch.BoundText)) = False Then
'            GoTo ErrTrap
'        End If
'        line_no = line_no + 1
'    End If

    
    ' «·«ÿ—«›
    ' „‰ Õ”«» „’«—Ì› › Õ «⁄ „«œ
     ' «·Ï Õ”«» «·»‰ﬂ
     
     
     ' ›Ï Õ«·… «‰ «·ﬁÌœ ÂÊ ﬁÌœ «·«€·«ﬁ «Ï «‰Â  „ «·÷€ÿ ⁄·Ï “— «€·«ﬁ «·«⁄ „«œ Ì⁄ﬂ” «·ﬁÌœ Ê«·ﬁÌ„…
        If IsClose Then
            If val(txtBondAmt) = 0 Then txtBondAmt = val(TxtValue)
            Notevalue = val(txtBondAmt.text) * val(txtPercentV) / 100
        Else
            Notevalue = val(TxtValue.text) * val(txtPercentV) / 100
        End If
        If notytype = 22010 Then
            Notevalue = val(txtOPenValue.text)
            
                    
            StrAccountCodeDebt = Trim(rsAcc("AccountExpensCode").value & "")
            
            Notevalue = val(txtOPenValue.text) * mRate
          
            PercentgValueAddedAccount_Transec dbFromDate.value, 22, 0, , mPercent
            Notevalue = Notevalue / (1 + mPercent / 100)
            'Notevalue2 = Notevalue * mPercent / 100
            'Notevalue = Notevalue - Notevalue2
        
            
            If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCodeDebt, Notevalue, 0, Msg & "    Account for the expenses of opening the LC  ", val(notes_id), , , , dbFromDate.value, val(DCboUserName.BoundText), , , , , , CLng(mRate), , "Account for the expenses of opening the LC ", setfoxy_Line, , , , , , , , , _
            val(dcBranch.BoundText)) = False Then
                GoTo ErrTrap
            End If
            
             
            
            GetValueAddedAccount dbFromDate.value, StrAccountCodeCridet, 1, 21
            'StrAccountCodeCridet = get_bank_Account(Dcbank.BoundText, "Account_Code")
            
            Notevalue2 = val(txtOPenValue.text) - Notevalue
            If Notevalue2 <> 0 Then
                line_no = line_no + 1
                If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCodeCridet, val(Notevalue2), 0, " Vat account", val(notes_id), , , , dbFromDate.value, val(DCboUserName.BoundText), , , , , , CLng(mRate), , " Vat account", setfoxy_Line, , , , , , , , , val(dcBranch.BoundText)) = False Then
                    GoTo ErrTrap
                End If
            End If
             
             StrAccountCodeCridet = get_bank_Account(Dcbank.BoundText, "Account_Code")
            Notevalue = val(txtOPenValue.text) * mRate
            line_no = line_no + 1
            If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCodeCridet, val(Notevalue), 1, Msg & "    Bank account", val(notes_id), , , , dbFromDate.value, val(DCboUserName.BoundText), , , , , , CLng(mRate), , "    Bank account", setfoxy_Line, , , , , , , , , val(dcBranch.BoundText)) = False Then
                GoTo ErrTrap
            End If
    
    
     
            
                updateNotesValueAndNobytext (val(notes_id))
                Exit Function
            
        End If
    If Notevalue > 0 Then
        ' «Ì‰ «Ãœ Õ”«» „’«—Ì› › Õ «·«⁄ „«œ Â· ÂÊ ›Ï „·› «·›—Ê⁄ «„ ÌÊÃœ „·› ··«⁄ „«œ« 
        '«” ›”«— Ê«∆·
        
        If IsClose Then
        
        
            StrAccountCodeCridet = get_bank_Account(Dcbank.BoundText, "Account_Code")
            ' ﬁÌ„… «·Ì‰ﬂ „‰ «Ì‰
            
            GetValueAddedAccount dbFromDate.value, StrAccountCodeCridet, 1, 21
            
            StrAccountCodeCridet = get_bank_Account(Dcbank.BoundText, "Account_Code")
            

            line_no = line_no + 1
            If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCodeCridet, val(Notevalue) * mRate, 0, Msg & " Bank account", val(notes_id), , , , dbFromDate.value, val(DCboUserName.BoundText), , , , , , CLng(mRate), , " Bank account", setfoxy_Line, , , , , , , , , val(dcBranch.BoundText)) = False Then
                GoTo ErrTrap
            End If
            
        
            StrAccountCodeDebt = Trim(rsAcc("Account_CodeMargin").value & "")
            
            line_no = line_no + 1
            If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCodeDebt, Notevalue, 1, Msg & "Margin Account.", val(notes_id), , , , dbFromDate.value, val(DCboUserName.BoundText), , , , , , CLng(mRate), , "Margin Account.", setfoxy_Line, , , , , , , , , _
            val(dcBranch.BoundText)) = False Then
                GoTo ErrTrap
            End If
            If val(txtCostLGYear) = 0 Then
                GoTo ExitFunc
            End If

        Else
            StrAccountCodeDebt = Trim(rsAcc("Account_CodeMargin").value & "")
            
            
            If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCodeDebt, Notevalue, 0, Msg & "Margin Account.", val(notes_id), , , , dbFromDate.value, val(DCboUserName.BoundText), , , , , , CLng(mRate), , "Margin Account.", setfoxy_Line, , , , , , , , , _
            val(dcBranch.BoundText)) = False Then
                GoTo ErrTrap
            End If
            StrAccountCodeCridet = get_bank_Account(Dcbank.BoundText, "Account_Code")
            ' ﬁÌ„… «·Ì‰ﬂ „‰ «Ì‰
            
            GetValueAddedAccount dbFromDate.value, StrAccountCodeCridet, 1, 21
            
            StrAccountCodeCridet = get_bank_Account(Dcbank.BoundText, "Account_Code")
            line_no = line_no + 1
            If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCodeCridet, val(Notevalue) * mRate, 1, Msg & " Bank account", val(notes_id), , , , dbFromDate.value, val(DCboUserName.BoundText), , , , , , CLng(mRate), , " Bank account", setfoxy_Line, , , , , , , , , val(dcBranch.BoundText)) = False Then
                GoTo ErrTrap
            End If
            
             
             If Not optTypeLCLG(1) Then
                GoTo ExitFunc
            End If
        End If
        line_no = line_no + 1
        
    End If
     
     
     
    Notevalue = val(txtOPenValue.text) * mRate

    If Notevalue > 0 Or val(txtCostLGYear) <> 0 Then
        ' «Ì‰ «Ãœ Õ”«» „’«—Ì› › Õ «·«⁄ „«œ Â· ÂÊ ›Ï „·› «·›—Ê⁄ «„ ÌÊÃœ „·› ··«⁄ „«œ« 
        '«” ›”«— Ê«∆·
        
          
        PercentgValueAddedAccount_Transec dbFromDate.value, 22, 0, , mPercent
       
        
        
        If val(txtCostLGYear) <> 0 And optTypeLCLG(1) Then
            cmbAccountExpProject.BoundText = Trim(rsAcc("AccountExpensCode").value & "")
            StrAccountCodeDebt = Trim(rsAcc("AccountExpensCode").value & "")
            Notevalue = txtCostLGYear
            
            If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCodeDebt, Notevalue, 0, Msg & "Project Bank Expenses Account until the End of the Year. ", val(notes_id), , , , dbFromDate.value, val(DCboUserName.BoundText), , , , , , CLng(mRate), , "Project Bank Expenses Account until the End of the Year. ", setfoxy_Line, , , , , , , , , _
            val(dcBranch.BoundText)) = False Then
                GoTo ErrTrap
            End If
            
            
      
        
        
            
           ' Notevalue2 = Notevalue
            line_no = line_no + 1
            StrAccountCodeCridet = get_bank_Account(Dcbank.BoundText, "Account_Code")
            If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCodeCridet, val(Notevalue) * mRate, 1, Msg & " Bank account", val(notes_id), , , , dbFromDate.value, val(DCboUserName.BoundText), , , , , , CLng(mRate), , " Bank account", setfoxy_Line, , , , , , , , , val(dcBranch.BoundText)) = False Then
                GoTo ErrTrap
            End If
            
            
                   Notevalue2 = Notevalue * mPercent / 100
            
                     
        GetValueAddedAccount dbFromDate.value, StrAccountCodeCridet, 1, 22
        
            If Notevalue2 <> 0 Then
                line_no = line_no + 1
                If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCodeCridet, val(Notevalue2) * mRate, 0, Msg & " Vat account", val(notes_id), , , , dbFromDate.value, val(DCboUserName.BoundText), , , , , , CLng(mRate), , " Vat account", setfoxy_Line, , , , , , , , , val(dcBranch.BoundText)) = False Then
                    GoTo ErrTrap
                End If
            End If
            
            
             line_no = line_no + 1
            StrAccountCodeCridet = get_bank_Account(Dcbank.BoundText, "Account_Code")
            If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCodeCridet, val(Notevalue2) * mRate, 1, Msg & " Bank account", val(notes_id), , , , dbFromDate.value, val(DCboUserName.BoundText), , , , , , CLng(mRate), , " Bank account", setfoxy_Line, , , , , , , , , val(dcBranch.BoundText)) = False Then
                GoTo ErrTrap
            End If
            
            'Notevalue = Notevalue - Notevalue2
        End If
        
        
        If val(txtCostLGYearLast) <> 0 And optTypeLCLG(1) Then
            line_no = line_no + 1
            StrAccountCodeDebt = get_account_code_branch(227, my_branch)
            Notevalue = txtCostLGYearLast
            
            If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCodeDebt, Notevalue, 0, Msg & "Prepaid Bank Expenses Account.", val(notes_id), , , , dbFromDate.value, val(DCboUserName.BoundText), , , , , , CLng(mRate), , "Prepaid Bank Expenses Account.", setfoxy_Line, , , , , , , , , _
            val(dcBranch.BoundText)) = False Then
                GoTo ErrTrap
            End If
            
            
            Notevalue3 = Notevalue * mPercent / 100
         '   Notevalue = Notevalue - Notevalue3
         '   Notevalue = Notevalue + Notevalue2
            line_no = line_no + 1
            StrAccountCodeCridet = get_bank_Account(Dcbank.BoundText, "Account_Code")
            If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCodeCridet, (val(Notevalue)) * mRate, 1, Msg & " Bank account", val(notes_id), , , , dbFromDate.value, val(DCboUserName.BoundText), , , , , , CLng(mRate), , " Bank account", setfoxy_Line, , , , , , , , , val(dcBranch.BoundText)) = False Then
                GoTo ErrTrap
            End If
            
            
        GetValueAddedAccount dbFromDate.value, StrAccountCodeCridet, 1, 22
        
            If Notevalue3 <> 0 Then
                line_no = line_no + 1
                If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCodeCridet, val(Notevalue3) * mRate, 0, Msg & " Vat account", val(notes_id), , , , dbFromDate.value, val(DCboUserName.BoundText), , , , , , CLng(mRate), , " Vat account", setfoxy_Line, , , , , , , , , val(dcBranch.BoundText)) = False Then
                    GoTo ErrTrap
                End If
            End If
        End If
        
            line_no = line_no + 1
            StrAccountCodeCridet = get_bank_Account(Dcbank.BoundText, "Account_Code")
            If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCodeCridet, (val(Notevalue3)) * mRate, 1, Msg & " Bank account", val(notes_id), , , , dbFromDate.value, val(DCboUserName.BoundText), , , , , , CLng(mRate), , " Bank account", setfoxy_Line, , , , , , , , , val(dcBranch.BoundText)) = False Then
                GoTo ErrTrap
            End If
        
        
'
        

    End If
ExitFunc:

    updateNotesValueAndNobytext (val(notes_id))
    Exit Function
ErrTrap:
  End Function



'-----------------------


Function createVoucher2(ByVal row As Long, ByVal mIsPay As Integer, ByVal TypeGrid As Long)
Dim NoteID As Long
Dim NoteDate As Date
Dim NoteSerial As String
Dim Notevalue As Double
Dim notytype As Integer
Dim des As String
des = "    Õ”«» «·" & TxtLcNo.text
Dim tablename As String
Dim Filedname As String
Dim NoteIDFiled As String
Dim NoteSerial1 As Long
Dim SerialFiled As String
Dim BranchID As Integer
Dim sql As String
Dim mRate  As Double
Dim mRowID As String
If mIsPay Then
    SerialFiled = "NoteSerial2"
    NoteIDFiled = "NoteID2"
Else
    SerialFiled = "NoteSerial"
    NoteIDFiled = "NoteID"

End If
If TypeGrid < 3 Then
    tablename = "TBLLCMargin"
    'Filedname = ""
    'SerialFiled = ""
    'NoteIDFiled = ""
ElseIf TypeGrid = 3 Then
    tablename = "TBLLCHistory"
ElseIf TypeGrid = 4 Then
    tablename = "tblLCOpenB"
    
    Filedname = ""
   
ElseIf TypeGrid = 6 Then
    tablename = "TBLLCMargin2"
    If mIsPay Then
        SerialFiled = "NoteSerial2"
        NoteIDFiled = "NoteID2"
    End If
    'NoteIDFiled = ""
End If

Filedname = "MarginNo"
NoteSerial1 = val(TxtLcNo)

BranchID = val(dcBranch.BoundText)
mRate = val(txt_Currency_rate)

Dim mIsOpenBalance As Boolean


'


' ⁄‰ „ﬂ«‰ Ê÷⁄ «·ÀÊ«»  ÊﬂÌ›Ì… «· —ﬁÌ„   Õ «Ã  Ê÷ÌÕ
' «” ›”«— Ê«∆·
' ·Œ»ÿ… ⁄‰œÏ ›Ï «·„”„Ì«  Ê«·‰Ê   «Ì»

Dim mGrid As Object
If TypeGrid = 1 Then
    Set mGrid = GrdMargin2
ElseIf TypeGrid = 0 Then
    Set mGrid = GrdMargin
ElseIf TypeGrid = 3 Then
    Set mGrid = GrdBondHistory
ElseIf TypeGrid = 4 Then
    Set mGrid = GrdMargin3
ElseIf TypeGrid = 5 Then
    Set mGrid = GrdMargin2
ElseIf TypeGrid = 6 Then
    Set mGrid = GrdMargin4
End If

If TypeGrid < 3 Or TypeGrid = 6 Then
    If mIsPay = 0 Then
         NoteID = val(mGrid.TextMatrix(row, mGrid.ColIndex("NoteID")))
         NoteSerial = val(mGrid.TextMatrix(row, mGrid.ColIndex("NoteSerial")))
        'mGrid.TextMatrix(row, mGrid.ColIndex("NoteSerial")) = NoteSerial
    Else
         NoteID = val(mGrid.TextMatrix(row, mGrid.ColIndex("NoteID2")))
        'mGrid.TextMatrix(row, mGrid.ColIndex("NoteSerial2")) = NoteSerial
        NoteSerial = val(mGrid.TextMatrix(row, mGrid.ColIndex("NoteSerial2")))
    
    End If
ElseIf TypeGrid = 3 Then
        NoteID = val(mGrid.TextMatrix(row, mGrid.ColIndex("NoteID")))
        'mGrid.TextMatrix(row, mGrid.ColIndex("NoteSerial")) = NoteSerial
        NoteSerial = val(mGrid.TextMatrix(row, mGrid.ColIndex("NoteSerial")))
ElseIf TypeGrid = 4 Then
        NoteID = val(mGrid.TextMatrix(row, mGrid.ColIndex("NoteID")))
        NoteSerial = val(mGrid.TextMatrix(row, mGrid.ColIndex("NoteSerial")))
        'mGrid.TextMatrix(row, mGrid.ColIndex("NoteSerial")) = NoteSerial
ElseIf TypeGrid = 5 Then
        NoteID = val(mGrid.TextMatrix(row, mGrid.ColIndex("NoteID3")))
        NoteSerial = val(mGrid.TextMatrix(row, mGrid.ColIndex("NoteSerial3")))
        'mGrid.TextMatrix(row, mGrid.ColIndex("NoteSerial3")) = NoteSerial
End If

'If NoteID <> 0 Then
'    If TypeGrid = 6 Then
'        If CBool(mGrid.ValueMatrix(row, mGrid.ColIndex("IsOpenBalance"))) And mIsPay = 0 Then
'            Cn.Execute "delete notes1 where noteid = " & NoteID
'            Cn.Execute "delete DOUBLE_ENTREY_VOUCHERS1 where Notes_ID = " & NoteID
'
'        End If
'    Else
'        Cn.Execute "delete notes where noteid = " & NoteID
'    End If
'End If
With mGrid
    If TypeGrid < 3 Then
        If mIsPay = 1 Then
            notytype = 22003
            Notevalue = val(.TextMatrix(row, .ColIndex("PayedAmount"))) * mRate
            
        Else
            notytype = 22002
            Notevalue = val(.TextMatrix(row, .ColIndex("Amount"))) * mRate
        End If
        If Notevalue = 0 Then Exit Function
        'mAccNO = val(DboParentAccount.BoundText)
        If mIsPay = 0 Then
            If (.TextMatrix(row, .ColIndex("OrderDate"))) = "" Then
                .TextMatrix(row, .ColIndex("OrderDate")) = Date
            End If
            NoteDate = (.TextMatrix(row, .ColIndex("OrderDate")))
        Else
            NoteDate = (.TextMatrix(row, .ColIndex("PayDate")))
        End If
    ElseIf TypeGrid = 3 Then
             notytype = 22004
             
             Notevalue = (val(.TextMatrix(row, .ColIndex("AmountPlus"))) - val(.TextMatrix(row, .ColIndex("AmountMin")))) * val(txtPercentV) / 100
             
             NoteDate = Date
    ElseIf TypeGrid = 4 Then
        notytype = 22006
        
        If (.TextMatrix(row, .ColIndex("GuaranteeDate"))) = "" Then
            .TextMatrix(row, .ColIndex("GuaranteeDate")) = Date
        End If
        
         'getFirstPeriodDateInthisYear2 FirstPeriodDateInthisYear
           ' Me.DTP_Date.value = FirstPeriodDateInthisYear
            
        NoteDate = (.TextMatrix(row, .ColIndex("GuaranteeDate")))
        Notevalue = (val(.TextMatrix(row, .ColIndex("InsuranceAmount"))) + val(.TextMatrix(row, .ColIndex("ExpAmount"))))
    ElseIf TypeGrid = 5 Then
        notytype = 22007
        NoteDate = (.TextMatrix(row, .ColIndex("OrderDate")))
        Notevalue = val(.TextMatrix(row, .ColIndex("MargenValue")))
    ElseIf TypeGrid = 6 Then
        If mIsPay = 1 Then
            notytype = 22009
            Notevalue = val(.TextMatrix(row, .ColIndex("PayedAmount"))) * mRate
            
        Else
            notytype = 22008
            Notevalue = val(.TextMatrix(row, .ColIndex("Amount"))) * mRate
        End If
        If Notevalue = 0 Then Exit Function
        'mAccNO = val(DboParentAccount.BoundText)
        If mIsPay = 0 Then
            If Trim(.TextMatrix(row, .ColIndex("OrderDate"))) = "" Then
                .TextMatrix(row, .ColIndex("OrderDate")) = Date
                
            End If
            NoteDate = (.TextMatrix(row, .ColIndex("OrderDate")))
        Else
            NoteDate = (.TextMatrix(row, .ColIndex("PayDate")))
        End If

    End If
End With


If Notevalue > 0 Then
    
    If TypeGrid = 6 Then
        If CBool(mGrid.ValueMatrix(row, mGrid.ColIndex("IsOpenBalance"))) And mIsPay = 0 Then
            mIsOpenBalance = True
            notytype = 101
            
        Else
            mIsOpenBalance = False
           
            ', recordDateH.value
        End If
    End If

    NoteSerial1 = val(mGrid.ValueMatrix(row, mGrid.ColIndex("ID")))
    Filedname = "ID"
    mRowID = Trim(mGrid.TextMatrix(row, mGrid.ColIndex("RowId")))
    If InStr(mRowID, "{") Then
    Else
        mRowID = "{" & Trim(mRowID) & "}"
    End If
    If val(NoteSerial) <> 0 Then
        s = "Select * from Notes where NoteSerial = " & NoteSerial & " and isNull(TblLCID,0)  <> " & val(TXTTblLCID) & " And NoteSerial1 <> " & NoteSerial1
        s = "Select * from Notes where NoteSerial = " & NoteSerial & " and RowID  <> '" & Trim(mRowID) & "' "
        Set rsDummy = New ADODB.Recordset
        rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly
        If Not rsDummy.EOF Then
            NoteSerial = 0
        
        End If
        
    End If
    
    CreateNotes NoteID, NoteDate, BranchID, notytype, Abs(Notevalue), NoteSerial, (NoteSerial1), tablename, Filedname, NoteSerial1, des, , TxtLcNo, NoteIDFiled, SerialFiled, , mIsOpenBalance, val(TXTTblLCID), mRowID, "RowId"
    rs.Resync adAffectCurrent
        
    
    CREATE_VOUCHER_GE2 val(NoteID), BranchID, val(DCboUserName.BoundText), NoteDate, row, mIsPay, Notevalue, NoteDate, TypeGrid, mIsOpenBalance
    
   ' rs.Resync adAffectCurrent


If TypeGrid < 3 Or TypeGrid = 6 Then
    If mIsPay = 0 Then
         mGrid.TextMatrix(row, mGrid.ColIndex("NoteID")) = NoteID
        mGrid.TextMatrix(row, mGrid.ColIndex("NoteSerial")) = NoteSerial
    Else
         mGrid.TextMatrix(row, mGrid.ColIndex("NoteID2")) = NoteID
        mGrid.TextMatrix(row, mGrid.ColIndex("NoteSerial2")) = NoteSerial
    
    End If
ElseIf TypeGrid = 3 Then
        mGrid.TextMatrix(row, mGrid.ColIndex("NoteID")) = NoteID
        mGrid.TextMatrix(row, mGrid.ColIndex("NoteSerial")) = NoteSerial
ElseIf TypeGrid = 4 Then
        mGrid.TextMatrix(row, mGrid.ColIndex("NoteID")) = NoteID
        mGrid.TextMatrix(row, mGrid.ColIndex("NoteSerial")) = NoteSerial
ElseIf TypeGrid = 5 Then
        mGrid.TextMatrix(row, mGrid.ColIndex("NoteID3")) = NoteID
        mGrid.TextMatrix(row, mGrid.ColIndex("NoteSerial3")) = NoteSerial
End If

'

If TypeGrid = 1 Then
'  s = "Update TBLLCMargin Set NoteId = " & val(mGrid.TextMatrix(row, mGrid.ColIndex("NoteId")))
'                        s = s & " ,NoteSerial = " & val(mGrid.TextMatrix(row, mGrid.ColIndex("NoteSerial")))
'                        s = s & " ,NoteSerial2 = " & val(mGrid.TextMatrix(row, mGrid.ColIndex("NoteSerial2")))
'                        s = s & " ,NoteId2 = " & val(mGrid.TextMatrix(row, mGrid.ColIndex("NoteId2")))
'                        s = s & " where TblLCID=" & val(TXTTblLCID.text)
'                        s = s & " and ID = " & val(mGrid.TextMatrix(row, mGrid.ColIndex("Id")))
'                        Cn.Execute s
'                        End If
'                        If TypeGrid = 6 Then
'
'  s = "Update TBLLCMargin2 Set NoteId = " & val(mGrid.TextMatrix(row, mGrid.ColIndex("NoteId")))
'                        s = s & " ,NoteSerial = " & val(mGrid.TextMatrix(row, mGrid.ColIndex("NoteSerial")))
'                        s = s & " ,NoteSerial2 = " & val(mGrid.TextMatrix(row, mGrid.ColIndex("NoteSerial2")))
'                        s = s & " ,NoteId2 = " & val(mGrid.TextMatrix(row, mGrid.ColIndex("NoteId2")))
'                        s = s & " where TblLCID=" & val(TXTTblLCID.text)
'                        s = s & " and ID = " & val(mGrid.TextMatrix(row, mGrid.ColIndex("Id")))
'                        Cn.Execute s
'
                End If
''
'
'
'    StrSQL = "update  " & tablename & "   set NoteID=" & NoteID & ",NoteSerial='" & NoteSerial & "'"
'
'    StrSQL = StrSQL & " Where " & Filedname & " = " & NoteSerial1 & ""
'    Cn.Execute StrSQL
     
     
 
End If
End Function
Public Function CREATE_VOUCHER_GE2(general_noteid As Long, BranchID As Integer, UserID As Long _
, NoteDate2 As Date, ByVal row As Long, ByVal mIsPay As Integer, ByVal Notevalue As Double, ByVal NoteDate As Date, ByVal TypeGrid As Long, Optional ByVal mIsOpenBalance As Boolean = False)


Dim mGrid As Object
If TypeGrid = 1 Then
    Set mGrid = GrdMargin2
ElseIf TypeGrid = 0 Then
    Set mGrid = GrdMargin
ElseIf TypeGrid = 3 Then
    Set mGrid = GrdBondHistory
ElseIf TypeGrid = 4 Then
    Set mGrid = GrdMargin3
ElseIf TypeGrid = 5 Then
    Set mGrid = GrdMargin2
ElseIf TypeGrid = 6 Then
    Set mGrid = GrdMargin4
    
End If

If mIsOpenBalance Then
   ' StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS1 Where Notes_ID=" & general_noteid
Else
   '  StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & general_noteid
End If
    Cn.Execute StrSQL, , adExecuteNoRecords
    
    Dim i As Integer
    Dim sql As String
    Dim StoreID6 As Integer
    Dim rs2 As ADODB.Recordset
    Set rs2 = New ADODB.Recordset
    
    Dim LngDevID As Long
    Dim Msg As String
    Dim StrAccountCodeDebt As String
    Dim StrAccountCodeCridet As String
    Dim X As Integer
   
    Dim rs As New ADODB.Recordset
    Dim notes_serial As String
    Dim notes_id As String
    Msg = "    Õ”«» " & TxtLcNo.text
    notes_id = general_noteid
    my_branch = val(dcBranch.BoundText)
    If mIsOpenBalance Then
        LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS1", "Double_Entry_Vouchers_ID", "")
    Else
        LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
    End If
    Dim line_no As Integer
    line_no = 1
    Dim Notevalue2 As Double
                  
                  
    
    Dim s As String
    Dim mRate As Double
    mRate = val(txt_Currency_rate)
 
    
             Dim mPercent As Double
        PercentgValueAddedAccount_Transec dbFromDate.value, 22, 0, , mPercent
        
    ' «” ›”«— Ê«∆·
    ' ﬁ„  »Ã·» «·Õ”«» «·„‰‘√ ›Ï «·«⁄ „«œ »Â–Â «·ÿ—Ìﬁ… ‰—ÃÊ «·„—«Ã⁄…
    Dim sqlS As String
    Dim rsAcc As New ADODB.Recordset
    Dim mDes As String
    Dim mDes2 As String
    
    
    
    If TypeGrid < 3 Or TypeGrid = 6 Then
        If Trim(mGrid.TextMatrix(row, mGrid.ColIndex("MarginAccountCode"))) = "" Then
          mGrid.TextMatrix(row, mGrid.ColIndex("BankAccountCode")) = get_bank_Account(Dcbank.BoundText, "Account_Code")
        End If
        If mIsPay = 0 Then
        
            If TypeGrid <> 6 Then
                StrAccountCodeDebt = Trim(mGrid.TextMatrix(row, mGrid.ColIndex("MarginAccountCode")))
                StrAccountCodeCridet = Trim(mGrid.TextMatrix(row, mGrid.ColIndex("BankAccountCode")))
            Else
                StrAccountCodeDebt = Trim(mGrid.TextMatrix(row, mGrid.ColIndex("MarginAccountCode")))
                StrAccountCodeCridet = Trim(mGrid.TextMatrix(row, mGrid.ColIndex("BankAccountCode")))
            
            End If
                 
            mDes = "LC Financing Account."
            mDes2 = "Bank account"
        Else
            
            StrAccountCodeDebt = Trim(mGrid.TextMatrix(row, mGrid.ColIndex("BankAccountCode"))) 'Trim(mGrid.TextMatrix(row, mGrid.ColIndex("MarginAccountCode")))
           
            
            StrAccountCodeCridet = Trim(mGrid.TextMatrix(row, mGrid.ColIndex("MarginAccountCode"))) ' get_bank_Account(Dcbank.BoundText, "Account_Code")  'Trim(mGrid.TextMatrix(row, mGrid.ColIndex("BankAccountCode")))
            If TypeGrid = 1 Then
                If Trim(mGrid.TextMatrix(row, mGrid.ColIndex("BankAccountCode2"))) = "" Then
                    mGrid.TextMatrix(row, mGrid.ColIndex("BankAccountCode2")) = mGrid.TextMatrix(row, mGrid.ColIndex("MarginAccountCode"))
                    mGrid.TextMatrix(row, mGrid.ColIndex("BankAccount2")) = mGrid.TextMatrix(row, mGrid.ColIndex("MarginAccount"))
                    
                End If
                StrAccountCodeCridet = mGrid.TextMatrix(row, mGrid.ColIndex("BankAccountCode2"))
            End If
            If TypeGrid = 6 Then
                   If Trim(mGrid.TextMatrix(row, mGrid.ColIndex("BankAccountCode2"))) = "" Then
                    mGrid.TextMatrix(row, mGrid.ColIndex("BankAccountCode2")) = get_bank_Account(Dcbank.BoundText, "Account_Code")
                    mGrid.TextMatrix(row, mGrid.ColIndex("BankAccount2")) = Dcbank.text
                   End If
                  StrAccountCodeCridet = Trim(mGrid.TextMatrix(row, mGrid.ColIndex("BankAccountCode2"))) ' get_bank_Account(Dcbank.BoundText, "Account_Code")
                   ' .TextMatrix(row, .ColIndex("MarginAccount")) = Dcbank.text
            End If
            
            mDes = "LC Financing Account."
            mDes2 = "Bank account"
        End If
    ElseIf TypeGrid = 3 Then
            
            
            
             
            
            sqlS = " Select Account_Code,Account_code2,Account_CodeMargin,AcceptAccount_Code,LCAccount_Code,AccountExpensCode from tblLc Where TblLCID=" & val(TXTTblLCID.text) '   and  TblLCid = '" & Trim(Txt) & "'"
            rsAcc.Open sqlS, Cn, adOpenStatic, adLockOptimistic, adCmdText
            
            If rsAcc.RecordCount <> 0 Then
                StrAccountCodeDebt = Trim(rsAcc("Account_CodeMargin").value & "")
                StrAccountCodeCridet = get_bank_Account(Dcbank.BoundText, "Account_Code")
            End If
            
            If Notevalue < 0 Then
                StrAccountCodeDebt = get_bank_Account(Dcbank.BoundText, "Account_Code")
                StrAccountCodeCridet = Trim(rsAcc("Account_CodeMargin").value & "")
            End If
            Notevalue = Abs(Notevalue)
            
            mDes = "Margin Account."
            mDes2 = "Bank account"
            
        ElseIf TypeGrid = 4 Then
            
            
            Dim mAccountCode As String
             
            
            sqlS = " Select Account_Code,Account_code2,Account_CodeMargin,AcceptAccount_Code,LCAccount_Code,AccountExpensCode from tblLc Where TblLCID=" & val(TXTTblLCID.text) '   and  TblLCid = '" & Trim(Txt) & "'"
            rsAcc.Open sqlS, Cn, adOpenStatic, adLockOptimistic, adCmdText
            
            If rsAcc.RecordCount <> 0 Then
                StrAccountCodeDebt = Trim(rsAcc("Account_CodeMargin").value & "")
                StrAccountCodeCridet = get_bank_Account(Dcbank.BoundText, "Account_Code")
                mAccountCode = Trim(rsAcc("AccountExpensCode").value & "")
            End If
            
            If Notevalue < 0 Then
                StrAccountCodeDebt = get_bank_Account(Dcbank.BoundText, "Account_Code")
                StrAccountCodeCridet = Trim(rsAcc("Account_CodeMargin").value & "")
            End If
            Notevalue = Abs(Notevalue)
            

            
            mDes = "Margin Account."
            mDes2 = "Bank account"
            
        ElseIf TypeGrid = 5 Then
            
            
            
             
            
            sqlS = " Select Account_Code,Account_code2,Account_CodeMargin,AcceptAccount_Code,LCAccount_Code,AccountExpensCode from tblLc Where TblLCID=" & val(TXTTblLCID.text) '   and  TblLCid = '" & Trim(Txt) & "'"
            rsAcc.Open sqlS, Cn, adOpenStatic, adLockOptimistic, adCmdText
            
            If rsAcc.RecordCount <> 0 Then
                StrAccountCodeDebt = Trim(mGrid.TextMatrix(row, mGrid.ColIndex("AccountMargen2")))
                StrAccountCodeCridet = get_bank_Account(Dcbank.BoundText, "Account_Code")
             '   mAccountCode = Trim(rsAcc("Account_Code").value & "")
            End If
            
            If Notevalue < 0 Then
                StrAccountCodeDebt = get_bank_Account(Dcbank.BoundText, "Account_Code")
                StrAccountCodeCridet = Trim(rsAcc("Account_CodeMargin").value & "")
            End If
            Notevalue = Abs(Notevalue)
            
        
                    
            mDes = "Margin Account."
            mDes2 = "Bank account"
    End If
         
   
 
     

    If Notevalue <> 0 Then
        ' «Ì‰ «Ãœ Õ”«» „’«—Ì› › Õ «·«⁄ „«œ Â· ÂÊ ›Ï „·› «·›—Ê⁄ «„ ÌÊÃœ „·› ··«⁄ „«œ« 
        '«” ›”«— Ê«∆·
       
        
        If TypeGrid = 4 Then
                Notevalue = val(val(mGrid.TextMatrix(row, mGrid.ColIndex("InsuranceAmount"))))
                NoteDate = mGrid.TextMatrix(row, mGrid.ColIndex("GuaranteeDate"))
                If Notevalue <> 0 Then
              If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCodeDebt, Notevalue, 0, Msg & mDes, val(notes_id), , , , NoteDate, val(DCboUserName.BoundText), , , , , , CLng(mRate), , mDes, setfoxy_Line, , , , mIsOpenBalance, get_opening_balance_voucher_id, , , , _
                  val(dcBranch.BoundText)) = False Then
                  GoTo ErrTrap
              End If
            End If
              ' ﬁÌ„… «·Ì‰ﬂ „‰ «Ì‰
              
             ' GetValueAddedAccount dbFromDate.value, StrAccountCodeCridet, 1, 21
              'Notevalue = val(val(mGrid.TextMatrix(row, mGrid.ColIndex("InsuranceAmount")))) + val(val(mGrid.TextMatrix(row, mGrid.ColIndex("ExpAmount"))))
              Notevalue = val(val(mGrid.TextMatrix(row, mGrid.ColIndex("ExpAmount")))) * mRate
              
              
              
            'Notevalue = val(txtOPenValue.text)
          
            PercentgValueAddedAccount_Transec dbFromDate.value, 22, 0, , mPercent
            Notevalue2 = Notevalue / (1 + mPercent / 100)
            Notevalue = Notevalue - Notevalue2
         
            Dim mVatD
            
              mVatD = Notevalue
              ' Notevalue2 = Notevalue * mPercent / 100
            '  StrAccountCodeCridet = get_bank_Account(DCBank.BoundText, "Account_Code")
              line_no = line_no + 1
'              If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCodeCridet, val(Notevalue), 1, Msg & mDes2, val(notes_id), , , , NoteDate, val(DCboUserName.BoundText), , , , , , CLng(mRate), , , setfoxy_Line, , , , , , , , , val(dcBranch.BoundText)) = False Then
'                  GoTo ErrTrap
'              End If
'              line_no = line_no + 1
              
               GetValueAddedAccount NoteDate, StrAccountCodeCridet, 1, 22


                            If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCodeCridet, val(Notevalue) * mRate, 0, Msg & "Vat account", val(notes_id), , , , NoteDate, val(DCboUserName.BoundText), , , , , , CLng(mRate), , "Vat account", setfoxy_Line, , , , mIsOpenBalance, get_opening_balance_voucher_id, , , , val(dcBranch.BoundText)) = False Then
                                GoTo ErrTrap
                            End If
                            
              Notevalue = val(val(mGrid.TextMatrix(row, mGrid.ColIndex("ExpAmount"))))
              
            '  StrAccountCodeCridet = get_bank_Account(DCBank.BoundText, "Account_Code")
               line_no = line_no + 1
              If ModAccounts.AddNewDev(LngDevID, line_no, mAccountCode, val(Notevalue2), 0, Msg & mDes2, val(notes_id), , , , NoteDate, val(DCboUserName.BoundText), , , , , , CLng(mRate), , mDes2, setfoxy_Line, , , , mIsOpenBalance, get_opening_balance_voucher_id, , , , val(dcBranch.BoundText)) = False Then
                  GoTo ErrTrap
              End If
              line_no = line_no + 1
              
              
              
              Notevalue = Notevalue - Notevalue2
           
                        If Notevalue2 <> 0 Then
                           
                        End If
              
               StrAccountCodeDebt = get_bank_Account(Dcbank.BoundText, "Account_Code")
                Notevalue = val(val(mGrid.TextMatrix(row, mGrid.ColIndex("InsuranceAmount"))))
               If Notevalue2 <> 0 Then
                            line_no = line_no + 1
                            If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCodeDebt, val(Notevalue2) + Notevalue + mVatD, 1, Msg & "Bank account", val(notes_id), , , , NoteDate, val(DCboUserName.BoundText), , , , , , CLng(mRate), , "Bank account", setfoxy_Line, , , , mIsOpenBalance, get_opening_balance_voucher_id, , , , val(dcBranch.BoundText)) = False Then
                                GoTo ErrTrap
                            End If
                        End If
              
              
        ElseIf TypeGrid = 5 Then
                Notevalue = val(val(mGrid.TextMatrix(row, mGrid.ColIndex("MargenValue"))))

              If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCodeDebt, Notevalue, 0, Msg & mDes, val(notes_id), , , , NoteDate, val(DCboUserName.BoundText), , , , , , CLng(mRate), , mDes, setfoxy_Line, , , , mIsOpenBalance, get_opening_balance_voucher_id, , , , _
                  val(dcBranch.BoundText)) = False Then
                  GoTo ErrTrap
              End If
            
              ' ﬁÌ„… «·Ì‰ﬂ „‰ «Ì‰
              
             ' GetValueAddedAccount dbFromDate.value, StrAccountCodeCridet, 1, 21
              
              
            '  StrAccountCodeCridet = get_bank_Account(DCBank.BoundText, "Account_Code")
              line_no = line_no + 1
              If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCodeCridet, val(Notevalue), 1, Msg & mDes2, val(notes_id), , , , NoteDate, val(DCboUserName.BoundText), , , , , , CLng(mRate), , mDes2, setfoxy_Line, , , , mIsOpenBalance, get_opening_balance_voucher_id, , , , val(dcBranch.BoundText)) = False Then
                  GoTo ErrTrap
              End If
              line_no = line_no + 1
              
             

        Else
              If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCodeDebt, Notevalue, 0, Msg & mDes, val(notes_id), , , , NoteDate, val(DCboUserName.BoundText), , , , , , CLng(mRate), , mDes, setfoxy_Line, , , , mIsOpenBalance, get_opening_balance_voucher_id, , , , _
              val(dcBranch.BoundText)) = False Then
                  GoTo ErrTrap
              End If
            
              ' ﬁÌ„… «·Ì‰ﬂ „‰ «Ì‰
              
             ' GetValueAddedAccount dbFromDate.value, StrAccountCodeCridet, 1, 21
              
              
            '  StrAccountCodeCridet = get_bank_Account(DCBank.BoundText, "Account_Code")
              line_no = line_no + 1
              If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCodeCridet, val(Notevalue), 1, Msg & mDes2, val(notes_id), , , , NoteDate, val(DCboUserName.BoundText), , , , , , CLng(mRate), , mDes2, setfoxy_Line, , , , mIsOpenBalance, get_opening_balance_voucher_id, , , , val(dcBranch.BoundText)) = False Then
                  GoTo ErrTrap
              End If
              line_no = line_no + 1
              
              If TypeGrid = 0 Then
                Notevalue = val(val(mGrid.TextMatrix(row, mGrid.ColIndex("Amount"))))
              Else
                
                    Notevalue = val(val(mGrid.TextMatrix(row, mGrid.ColIndex("MargenValue"))))
                    If (TypeGrid = 1 Or TypeGrid = 6) And mIsPay = 0 Then Notevalue = 0
                
                
              End If
              If Notevalue <> 0 Then
                  If TypeGrid = 1 Or TypeGrid = 6 Then
                  
                        
                            
                  
                      '  Notevalue2 = Notevalue * mPercent / 100
                
                           
                        sqlS = " Select Account_Code,Account_code2,Account_CodeMargin,AcceptAccount_Code,LCAccount_Code from tblLc Where TblLCID=" & val(TXTTblLCID.text) '   and  TblLCid = '" & Trim(Txt) & "'"
                        rsAcc.Open sqlS, Cn, adOpenStatic, adLockOptimistic, adCmdText
                        
                        If rsAcc.RecordCount <> 0 Then
                            StrAccountCodeDebt = Trim(mGrid.TextMatrix(row, mGrid.ColIndex("AccountMargen2")))
                            If TypeGrid = 1 Then
                                StrAccountCodeCridet = Trim(mGrid.TextMatrix(row, mGrid.ColIndex("MarginAccountCode"))) 'get_bank_Account(Dcbank.BoundText, "Account_Code")
                                StrAccountCodeCridet = Trim(mGrid.TextMatrix(row, mGrid.ColIndex("BankAccountCode2")))
                            ElseIf TypeGrid = 6 Then
                                StrAccountCodeCridet = Trim(mGrid.TextMatrix(row, mGrid.ColIndex("BankAccountCode2"))) 'get_bank_Account(Dcbank.BoundText, "Account_Code")
                            End If
                         '   mAccountCode = Trim(rsAcc("Account_Code").value & "")
                        End If
                        
                        If Notevalue < 0 Then
                            StrAccountCodeDebt = get_bank_Account(Dcbank.BoundText, "Account_Code")
                            
                             If TypeGrid = 1 Then
                                StrAccountCodeDebt = Trim(mGrid.TextMatrix(row, mGrid.ColIndex("MarginAccountCode"))) 'get_bank_Account(Dcbank.BoundText, "Account_Code")
                            ElseIf TypeGrid = 6 Then
                                StrAccountCodeDebt = Trim(mGrid.TextMatrix(row, mGrid.ColIndex("BankAccountCode2"))) 'get_bank_Account(Dcbank.BoundText, "Account_Code")
                            End If
                            
                            StrAccountCodeCridet = Trim(rsAcc("Account_CodeMargin").value & "")
                        End If
                        Notevalue = Abs(Notevalue)
                        
                        mDes = "Õ”«» «·„«—Ã‰"
                        mDes2 = "Õ”«» «·»‰ﬂ"
                         
                         
                         
                        Notevalue = val(val(mGrid.TextMatrix(row, mGrid.ColIndex("MargenValue"))))
    
                        If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCodeDebt, Notevalue, 0, Msg & mDes, val(notes_id), , , , NoteDate, val(DCboUserName.BoundText), , , , , , CLng(mRate), , mDes, setfoxy_Line, , , , mIsOpenBalance, get_opening_balance_voucher_id, , , , _
                            val(dcBranch.BoundText)) = False Then
                            GoTo ErrTrap
                        End If
                        
                      ' ﬁÌ„… «·Ì‰ﬂ „‰ «Ì‰
                      
                     ' GetValueAddedAccount dbFromDate.value, StrAccountCodeCridet, 1, 21
                      
                      
                    '  StrAccountCodeCridet = get_bank_Account(DCBank.BoundText, "Account_Code")
                      line_no = line_no + 1
                      If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCodeCridet, val(Notevalue) + Notevalue2, 1, Msg & mDes2, val(notes_id), , , , NoteDate, val(DCboUserName.BoundText), , , , , , CLng(mRate), , mDes2, setfoxy_Line, , , , mIsOpenBalance, get_opening_balance_voucher_id, , , , val(dcBranch.BoundText)) = False Then
                          GoTo ErrTrap
                      End If
                      line_no = line_no + 1
                      
                                      
                                      
                         
'                        GetValueAddedAccount dbFromDate.value, StrAccountCodeCridet, 1, 22
'
'                        If Notevalue2 <> 0 Then
'                            line_no = line_no + 1
'                            If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCodeCridet, val(Notevalue2) * mRate, 0, Msg & "    Õ”«»  «·÷—Ì»… ", val(notes_id), , , , dbFromDate.value, val(DCboUserName.BoundText), , , , , , CLng(mRate), , , setfoxy_Line, , , , , , , , , val(DcBranch.BoundText)) = False Then
'                                GoTo ErrTrap
'                            End If
'                        End If
                    End If
                End If
        End If
    End If
     
     
     
    
    

    updateNotesValueAndNobytext (val(notes_id))
    Exit Function
ErrTrap:
  End Function




'--------------------------





Private Sub cmbAccount_KeyUp(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyF3 Then
        Account_search.show
        Account_search.case_id = 988
    End If
End Sub

Private Sub cmbAccountAcceptanceParent_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF3 Then
      Account_search.case_id = 559
        Account_search.show
        
    End If


End Sub

Private Sub cmbAccountExpensParent_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF3 Then
        Account_search.show
        Account_search.case_id = 556
    End If

End Sub

Private Sub cmbAccountExpProject_KeyUp(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyF3 Then
        Account_search.show
        Account_search.case_id = 990
    End If

End Sub

Private Sub cmbAccountLGParent_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF3 Then
        Account_search.show
        Account_search.case_id = 558
    End If


End Sub

Private Sub cmbAccountMarginParent_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF3 Then
        Account_search.show
        Account_search.case_id = 557
    End If

End Sub

Private Sub CmdCreateV_Click()

'If val(TxtNoteSerial.text) = 0 Then
createVoucher False

Dim NoteID As Long
Dim NoteDate As Date
Dim NoteSerial As String
Dim Notevalue As Double
'Dim
Dim i
Dim des As String
des = "    Õ”«» «·" & TxtLcNo.text
Dim tablename As String
Dim Filedname As String
Dim NoteSerial1 As Long
Dim BranchID As Integer
Dim sql As String
Dim mRate  As Double

Dim mFieldNoteID As String
Dim mFieldNoteSerial As String
tablename = "TBLLC"
Dim mRowID As String

Filedname = "TblLCID"
NoteSerial1 = val(TXTTblLCID.text)

BranchID = val(dcBranch.BoundText)
mRate = val(txt_Currency_rate)
Dim RowIDField As String
'  s = "Update TBLLCMargin Set NoteId = " & val(mGrid.TextMatrix(row, mGrid.ColIndex("NoteId")))
'                        s = s & " ,NoteSerial = " & val(mGrid.TextMatrix(row, mGrid.ColIndex("NoteSerial")))
'                        s = s & " ,NoteSerial2 = " & val(mGrid.TextMatrix(row, mGrid.ColIndex("NoteSerial2")))
'                        s = s & " ,NoteId2 = " & val(mGrid.TextMatrix(row, mGrid.ColIndex("NoteId2")))
'                        s = s & " where TblLCID=" & val(TXTTblLCID.text)
'                        s = s & " and ID = " & val(mGrid.TextMatrix(row, mGrid.ColIndex("Id")))
'                        Cn.Execute s
'


'         StrSQL = "sELECT * FROM TBLLCMargin2 Where 1 = -1"
''IncrementID
'    saveGrid StrSQL, GrdMargin4, "MarginNo", "IncrementID", "TblLCID", val(Me.TXTTblLCID.text), "Type", 0
'
'
'
'
'    StrSQL = "SELECT * FROM TBLLCMargin Where 1 = -1"
'   'IncrementID
'    saveGrid StrSQL, GrdMargin2, "MarginNo", "IncrementID", "TblLCID", val(Me.TXTTblLCID.text)
'

       'FindRec val(TXTTblLCID.text)
      'rs.Find "TblLCID=" & val(TXTTblLCID.text), , adSearchForward, adBookmarkFirst
      
      
      
      
          
    Dim row As Long
    With GrdMargin4
        For i = 1 To .rows - 1
            row = i
'            s = "Delete Notes where NoteId = " & val(.TextMatrix(row, .ColIndex("NoteId3")))
'            Cn.Execute s
'            s = "Delete Notes where NoteId = " & val(.TextMatrix(row, .ColIndex("NoteId2")))
'            Cn.Execute s
            If CBool(.ValueMatrix(row, .ColIndex("IsOpenBalance"))) Then
'                s = "Delete Notes1 where NoteId = " & val(.TextMatrix(row, .ColIndex("NoteId")))
'                Cn.Execute s
'                s = "Delete DOUBLE_ENTREY_VOUCHERS1 where Notes_ID = " & val(.TextMatrix(row, .ColIndex("NoteId")))
'
'                Cn.Execute s
            Else
'                s = "Delete Notes where NoteId = " & val(.TextMatrix(row, .ColIndex("NoteId")))
'                Cn.Execute s
            End If
            If row = 6 Then
                row = 6
            End If
            If val(.TextMatrix(row, .ColIndex("Amount"))) <> 0 Then
                CreateEntry row, 6, 0
            End If
            If val(.TextMatrix(row, .ColIndex("PayedAmount"))) <> 0 Then
                CreateEntry row, 6, 1
            Else
                .TextMatrix(row, .ColIndex("NoteSerial2")) = ""
            End If
'            .TextMatrix(row, .ColIndex("NoteId")) = ""
'            .TextMatrix(row, .ColIndex("NoteId2")) = ""
'            .TextMatrix(row, .ColIndex("NoteId3")) = ""
'            .TextMatrix(row, .ColIndex("NoteSerial")) = ""
'            .TextMatrix(row, .ColIndex("NoteSerial2")) = ""
        
        Next
    End With
    
    
    
    With GrdMargin2
        For i = 1 To .rows - 1
            row = i
           
            
            
 
        
            
            If val(.TextMatrix(row, .ColIndex("Amount"))) <> 0 Then
                CreateEntry row, 1, 0
            End If
            If val(.TextMatrix(row, .ColIndex("PayedAmount"))) <> 0 Then
                
                CreateEntry row, 1, 1
            Else
                .TextMatrix(row, .ColIndex("NoteSerial2")) = ""
            End If
'            .TextMatrix(row, .ColIndex("NoteId")) = ""
'            .TextMatrix(row, .ColIndex("NoteId2")) = ""
'            .TextMatrix(row, .ColIndex("NoteId3")) = ""
'            .TextMatrix(row, .ColIndex("NoteSerial")) = ""
'            .TextMatrix(row, .ColIndex("NoteSerial2")) = ""
        
        Next
    End With
    
        
     With GrdMargin3
        For i = 1 To .rows - 1
            row = i
           
           
'            s = "Delete Notes where NoteId = " & val(.TextMatrix(row, .ColIndex("NoteId")))
'            Cn.Execute s
            
            
 
        
            If val(.TextMatrix(row, .ColIndex("TotalAmount"))) <> 0 Then
                CreateEntry row, 4, 0
            End If
'            .TextMatrix(row, .ColIndex("NoteId")) = ""
'            .TextMatrix(row, .ColIndex("NoteId2")) = ""
'            .TextMatrix(row, .ColIndex("NoteId3")) = ""
'            .TextMatrix(row, .ColIndex("NoteSerial")) = ""
'            .TextMatrix(row, .ColIndex("NoteSerial2")) = ""
        
        Next
    End With
        If SystemOptions.UserInterface = ArabicInterface Then
           ' MsgBox " „ «‰‘«¡ «·ﬁÌœ"
            If val(TXTNoteID) <> 0 Then
                CmdCreateV.Enabled = False
                Command9.Enabled = True
                Command2.Enabled = True
                Cmd(2).Enabled = False
            Else
                CmdCreateV.Enabled = True
'                Command9.Enabled = False
                Command2.Enabled = False
            End If
        Else
            MsgBox "Done"
        End If
'End If
End Sub

Private Sub CmdCreateV2_Click(index As Integer)

Dim i As Long

Dim StrSQL As String
Exit Sub

    
    

Dim mGrid As Object
If index = 1 Then
    Set mGrid = GrdMargin2
Else
    Set mGrid = GrdMargin
End If

'
'StrSQL = "delete From TBLLCMargin where TblLCID=" & val(TXTTblLCID.text) & " And Type = " & Index
'Cn.Execute StrSQL, , adExecuteNoRecords
'
'
'StrSQL = "sELECT * FROM TBLLCMargin Where 1 = -1"
'
'saveGrid StrSQL, mGrid, "MarginNo", "", "TblLCID", val(Me.TXTTblLCID.text), "Type", Index
    
    
StrSQL = "delete From TBLLCMargin where TblLCID=" & val(TXTTblLCID.text)
Cn.Execute StrSQL, , adExecuteNoRecords
        

StrSQL = "sELECT * FROM TBLLCMargin Where 1 = -1"
   
saveGrid StrSQL, mGrid, "MarginNo", "", "TblLCID", val(Me.TXTTblLCID.text)
    
    

StrSQL = "sELECT TBLLCMargin.*,Acc.Account_Serial as MarginAccountSerial,Acc.Account_Name as MarginAccount  ,"
StrSQL = StrSQL & " Acc2.Account_Serial as BankAccountSerial,Acc2.Account_Name as BankAccount  "
StrSQL = StrSQL & " from TBLLCMargin Left Outer join Accounts Acc On TBLLCMargin.MarginAccountCode = Acc.Account_Code "
StrSQL = StrSQL & " Left Outer join Accounts Acc2 On TBLLCMargin.BankAccountCode = Acc2.Account_Code"
StrSQL = StrSQL & " Where Type= " & index & "  and  TblLCID = " & val(Me.TXTTblLCID.text)
loadgrid StrSQL, mGrid, True, False
Dim Notevalue As Double

    

With mGrid
    For i = 1 To .rows - 1
        If index = 0 Then
            Notevalue = val(.TextMatrix(i, .ColIndex("Amount")))
        Else
            Notevalue = val(.TextMatrix(i, .ColIndex("PayedAmount")))
        End If
        If val(.TextMatrix(i, .ColIndex("NoteSerial"))) = 0 And Notevalue <> 0 Then
            createVoucher2 i, index, 0
            If index = 1 Then
                s = "Update TBLLCMargin Set PayedAmount = PayedAmount + " & val(.TextMatrix(i, .ColIndex("PayedAmount")))
                s = s & " where TblLCID=" & val(TXTTblLCID.text) & "  and MarginNo = " & val(.TextMatrix(i, .ColIndex("MarginNo")))
              '  Cn.Execute s
                
                
'                s = "Update TBLLCMargin Set NoteId = " & val(.TextMatrix(i, .ColIndex("NoteId")))
'                s = s & " ,NoteSerial = " & val(.TextMatrix(i, .ColIndex("NoteSerial")))
'                s = s & " where TblLCID=" & val(TXTTblLCID.text) & " And Type = " & Index & " and MarginNo = " & val(.TextMatrix(i, .ColIndex("MarginNo")))
'                s = s & " and ID = " & val(.TextMatrix(i, .ColIndex("Id")))
'                Cn.Execute s



                
                s = "Update TBLLCMargin Set StillAmount =Amount - PayedAmount "
                s = s & " where TblLCID=" & val(TXTTblLCID.text) & "   and MarginNo = " & val(.TextMatrix(i, .ColIndex("MarginNo")))
                Cn.Execute s


               
                
               
                'Cn.Execute s
                
            End If
            If val(.TextMatrix(i, .ColIndex("PayedAmount"))) <> 0 Then
            End If
        End If
    Next
End With
End Sub


Private Sub CreateEntry(ByVal mRow As Long, ByVal TypeGrid As Long, ByVal index As Long)

Dim i As Long

Dim StrSQL As String

Dim mNoteSerial As Long
    
Dim mTableName As String

Dim mGrid As Object
If TypeGrid = 1 Then
    Set mGrid = GrdMargin2
    mTableName = "TBLLCMargin"
ElseIf TypeGrid = 0 Then
    Set mGrid = GrdMargin
    mTableName = "TBLLCMargin"
ElseIf TypeGrid = 3 Then
    Set mGrid = GrdBondHistory
    mTableName = "TBLLCHistory"
ElseIf TypeGrid = 4 Then
    Set mGrid = GrdMargin3
    mTableName = "tblLCOpenB"
ElseIf TypeGrid = 5 Then
    Set mGrid = GrdMargin2
    mTableName = "TBLLCMargin"
    
ElseIf TypeGrid = 6 Then
    Set mGrid = GrdMargin4
    mTableName = "TBLLCMargin2"
    
End If


'StrSQL = "delete From TBLLCMargin where TblLCID=" & val(TXTTblLCID.text) & " And Type = " & Index
'Cn.Execute StrSQL, , adExecuteNoRecords
'
'
'StrSQL = "sELECT * FROM TBLLCMargin Where 1 = -1"
'
'saveGrid StrSQL, mGrid, "MarginNo", "", "TblLCID", val(Me.TXTTblLCID.text), "Type", Index
'
'
'
'    StrSQL = "delete From " & mTableName & "  where TblLCID=" & val(TXTTblLCID.text)
'    Cn.Execute StrSQL, , adExecuteNoRecords
'
'
'    StrSQL = "sELECT * FROM " & mTableName & "  Where 1 = -1"
'
'    saveGrid StrSQL, mGrid, "MarginNo", "ID", "TblLCID", val(Me.TXTTblLCID.text)

        
'
'StrSQL = "sELECT TBLLCMargin.*,Acc.Account_Serial as MarginAccountSerial,Acc.Account_Name as MarginAccount  ,"
'StrSQL = StrSQL & " Acc2.Account_Serial as BankAccountSerial,Acc2.Account_Name as BankAccount  "
'StrSQL = StrSQL & " from TBLLCMargin Left Outer join Accounts Acc On TBLLCMargin.MarginAccountCode = Acc.Account_Code "
'StrSQL = StrSQL & " Left Outer join Accounts Acc2 On TBLLCMargin.BankAccountCode = Acc2.Account_Code"
'StrSQL = StrSQL & " Where Type= " & Index & "  and  TblLCID = " & val(Me.TXTTblLCID.text)
'loadgrid StrSQL, mGrid, True, False
Dim Notevalue As Double

    

With mGrid
    For i = mRow To mRow
        mNoteSerial = val(.TextMatrix(i, .ColIndex("NoteSerial")))
        If TypeGrid < 3 Or TypeGrid = 6 Then
            If index = 0 Then
                Notevalue = val(.TextMatrix(i, .ColIndex("Amount")))
            Else
                Notevalue = val(.TextMatrix(i, .ColIndex("PayedAmount")))
                mNoteSerial = val(.TextMatrix(i, .ColIndex("NoteSerial2")))
            End If
        ElseIf TypeGrid = 3 Then
            If val(.TextMatrix(i, .ColIndex("AmountPlus"))) <> 0 Then
                Notevalue = val(.TextMatrix(i, .ColIndex("Total"))) * val(txtPercentV) / 100
                Notevalue = (val(.TextMatrix(i, .ColIndex("AmountPlus"))) - val(.TextMatrix(i, .ColIndex("AmountMin")))) * val(txtPercentV) / 100
            End If
        ElseIf TypeGrid = 4 Then
            If val(.TextMatrix(i, .ColIndex("TotalAmount"))) <> 0 Then
                'Notevalue = val(.TextMatrix(i, .ColIndex("Total"))) * val(txtPercentV) / 100
                Notevalue = (val(.TextMatrix(i, .ColIndex("InsuranceAmount"))) - val(.TextMatrix(i, .ColIndex("ExpAmount"))))
                
            End If
        ElseIf TypeGrid = 5 Then
            If val(.TextMatrix(i, .ColIndex("MargenValue"))) <> 0 Then
                'Notevalue = val(.TextMatrix(i, .ColIndex("Total"))) * val(txtPercentV) / 100
                Notevalue = val(.TextMatrix(i, .ColIndex("MargenValue")))
                
            End If
    
        End If
        
        'If (mNoteSerial = 0 And Notevalue <> 0) Then
        If (Notevalue <> 0) Then
            createVoucher2 i, index, TypeGrid
            
            If TypeGrid < 3 Or TypeGrid = 6 Then
                    If index = 1 Then
                        s = "Update " & mTableName & "  Set PayedAmount = PayedAmount + " & val(.TextMatrix(i, .ColIndex("PayedAmount")))
                       ' s = s & " ,IsFullPayed = " & val(.TextMatrix(i, .ColIndex("IsFullPayed")))
                        s = s & " where TblLCID=" & val(TXTTblLCID.text) & "  and MarginNo = " & val(.TextMatrix(i, .ColIndex("MarginNo")))
                      '  Cn.Execute s
                        
        '
                      
                        
                        s = "Update TBLLCMargin Set StillAmount =Amount - PayedAmount "
                        s = s & " where TblLCID=" & val(TXTTblLCID.text) & "  and MarginNo = " & val(.TextMatrix(i, .ColIndex("MarginNo")))
                        Cn.Execute s
        
        
                        s = "Update TBLLCMargin2 Set StillAmount =Amount - PayedAmount "
                        s = s & " where TblLCID=" & val(TXTTblLCID.text) & "  and MarginNo = " & val(.TextMatrix(i, .ColIndex("MarginNo")))
                        Cn.Execute s
        
'
'                        s = "Update TBLLC Set MarginTotal4 =" & val(txtMarginTotal4)
'                        s = s & " where TblLCID=" & val(TXTTblLCID.text)
'                        Cn.Execute s
'
'
'                        s = "Update TBLLC Set MarginTotal2 =" & val(txtMarginTotal2)
'                        s = s & " where TblLCID=" & val(TXTTblLCID.text)
'                        Cn.Execute s
'
'                        s = "Update TBLLC Set MarginTotal3 =" & val(txtMarginTotal3)
'                        s = s & " where TblLCID=" & val(TXTTblLCID.text)
'                        Cn.Execute s
                        
                        

                        'rs.Resync adAffectCurrent
                      '  rs.update
                        
                    End If
            End If
            
        End If
    Next
End With

'StrSQL = "delete From " & mTableName & "  where TblLCID=" & val(TXTTblLCID.text)
'Cn.Execute StrSQL, , adExecuteNoRecords
'
'
'StrSQL = "sELECT * FROM " & mTableName & "  Where 1 = -1"
'
'saveGrid StrSQL, mGrid, "MarginNo", "", "TblLCID", val(Me.TXTTblLCID.text)

  With mGrid
    For i = mRow To mRow
        
        If index = 1 Then
                If val(.TextMatrix(i, .ColIndex("IsFullPayed"))) = 1 Then
                    s = "Update " & mTableName & "  Set "
                    s = s & " IsFullPayed = " & val(.TextMatrix(i, .ColIndex("IsFullPayed")))
                    s = s & " where TblLCID=" & val(TXTTblLCID.text) & "  and MarginNo = " & val(.TextMatrix(i, .ColIndex("MarginNo")))
                    
                    Cn.Execute s
                
                End If
        End If
    Next
End With
End Sub

Private Sub Command2_Click()
'If Me.TxtModFlg.text = "R" Or Me.TxtModFlg.text = "" Then
Dim X As Integer
Dim Msg As String
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = " √ﬂÌœ Õ–› «·ﬁÌœ "
    Else
        Msg = "Confirm Delete  "
    End If
        X = MsgBox(Msg, vbCritical + vbYesNo)

      If X = vbYes Then
        StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(Me.TXTNoteID.text)
        Cn.Execute StrSQL, , adExecuteNoRecords
        StrSQL = "Delete From Notes Where NoteID=" & val(Me.TXTNoteID.text)
        Cn.Execute StrSQL, , adExecuteNoRecords
       ' Cn.Execute " Update TblLC set NoteID=null ,NoteSerial=null where TblLCID=" & val(TXTTblLCID.text)
        rs!NoteID = Null
        rs!NoteSerial = Null
        rs.update
        
        TxtNoteSerial = ""
        TXTNoteID = ""
        rs.Requery
         
         FindRec val(TXTTblLCID.text)
         TxtModFlg.text = ""
         TxtNoteSerial = ""
          If SystemOptions.UserInterface = ArabicInterface Then
            Msg = " „  Õ–› «·ﬁÌœ   "
            
           
            If val(TXTNoteID) <> 0 Then
                CmdCreateV.Enabled = False
                Command9.Enabled = True
                Command2.Enabled = True
             ' Cmd(2).Enabled = False
             '   Cmd(1).Enabled = False
             Else
                CmdCreateV.Enabled = True
'                Command9.Enabled = False
                Command2.Enabled = False
            End If
        Else
            Msg = " This voucher deleted  "
        End If
        MsgBox Msg
       End If
' End If

End Sub
Public Function FindRec(ByVal RecId As Long)
    On Error GoTo ErrTrap
    
    'rs.Find "ID=" & RecID, , adSearchForward, 1
    rs.Find "TblLCID=" & val(TXTTblLCID.text), , adSearchForward, adBookmarkFirst
    If Not (rs.EOF) Then
        FiLLTXT
        Retrive
        End If
    Exit Function
ErrTrap:
    If rs.EditMode <> adEditNone Then
        rs.CancelUpdate
      '  BtnUndo_Click
    End If
  End Function
  
  Public Sub FiLLTXT()
   On Error GoTo ErrTrap
    'TxtSerial.Text = IIf(IsNull(RsSavRec.Fields("ID").value), "", RsSavRec.Fields("ID").value)
    TXTNoteID = rs!NoteID & ""
    TxtNoteSerial = rs!NoteSerial & ""
    TxtNoteID2 = rs!NoteID2 & ""
    TxtNoteSerial2 = rs!NoteSerial2 & ""
    txtOPenValue = rs!OpenValue & ""
   
    txtNoteIDOpen = rs!NoteIDOpen & ""
    txtNoteSerialOpen = rs!NoteSerialOpen & ""
   If val(txtNoteIDOpen) <> 0 Then
        Command3.Enabled = False
        txtOPenValue.locked = True
        Command6.Enabled = True
    Else
        txtOPenValue.locked = False
        Command6.Enabled = False
        Command3.Enabled = True
   End If
 
ErrTrap:
End Sub

Private Sub Command3_Click()
   If val(txtNoteIDOpen) <> 0 Then
        txtOPenValue.locked = True
    Else
        txtOPenValue.locked = False
   End If
   
   
   
   
'If val(txtNoteSerialOpen.text) = 0 Then
        createVoucher False, 22010
       'FindRec val(TXTTblLCID.text)
      ' rs.Find "TblLCID=" & val(TXTTblLCID.text), , adSearchForward, adBookmarkFirst
       
    '   rs!ToDate = dbTodate.value
    '   rs!NoteID2 = val(TxtNoteID2)
    '   rs!NoteSerial2 = val(TxtNoteSerial2)
    '   rs("Locked").value = 1
       
     '  rs.update
       'ChkLocked.value = vbChecked
        If SystemOptions.UserInterface = ArabicInterface Then
           ' MsgBox " „ «‰‘«¡ «·ﬁÌœ"
            If val(txtNoteIDOpen) <> 0 Then
                Command3.Enabled = False
                
                Command4.Enabled = True
'                Cmd(2).Enabled = False
            Else

                Command3.Enabled = True
                
                Command4.Enabled = False

            End If
        Else
            MsgBox "Done"
        End If
'End If
End Sub

Private Sub Command4_Click()
ShowGL_cc Me.txtNoteSerialOpen.text, , 22010
End Sub

Private Sub Command6_Click()
If Me.TxtModFlg.text = "R" Or Me.TxtModFlg.text = "" Then
Dim X As Integer
Dim Msg As String
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = " √ﬂÌœ Õ–› «·ﬁÌœ "
    Else
        Msg = "Confirm Delete  "
    End If
        X = MsgBox(Msg, vbCritical + vbYesNo)

      If X = vbYes Then
        StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(Me.txtNoteIDOpen.text)
        Cn.Execute StrSQL, , adExecuteNoRecords
        StrSQL = "Delete From Notes Where NoteID=" & val(Me.txtNoteIDOpen.text)
        Cn.Execute StrSQL, , adExecuteNoRecords
       ' Cn.Execute " Update TblLC set NoteIDOpen=null ,NoteSerialOpen=null where TblLCID=" & val(TXTTblLCID.text)
        rs!NoteIDOpen = Null
        rs!NoteSerialOpen = Null
        rs.update
        txtNoteSerialOpen = ""
        txtNoteIDOpen = ""
        rs.Requery
         FindRec val(TXTTblLCID.text)
         TxtModFlg.text = ""
         txtNoteSerialOpen = ""
          If SystemOptions.UserInterface = ArabicInterface Then
            Msg = " „  Õ–› «·ﬁÌœ   "
            
           
            If val(txtNoteIDOpen) <> 0 Then
                Command3.Enabled = False
                Command4.Enabled = True
                Command6.Enabled = True
               
             Else
                
                Command4.Enabled = False
                Command6.Enabled = False
            End If
        Else
            Msg = " This voucher deleted  "
        End If
        MsgBox Msg
       End If
 End If
   If val(txtNoteIDOpen) <> 0 Then
        txtOPenValue.locked = True
    Else
        txtOPenValue.locked = False
   End If

End Sub

Private Sub Command7_Click()
Translatefrm Me
End Sub

Private Sub Command9_Click()
ShowGL_cc Me.TxtNoteSerial.text, , 200
End Sub


Private Sub DBCboClientName_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF3 Then
        FrmCustemerSearch.SearchType = 8987
        FrmCustemerSearch.show vbModal

    End If
 

End Sub

Private Sub dbFromDate_Change()
If Me.TxtModFlg.text = "" Or Me.TxtModFlg.text = "R" Then Exit Sub
    txtGuaranteeDate.value = dbFromDate.value
txtGuaranteeDate_Change
End Sub

Private Sub dcbank_Click(Area As Integer)
s = "Select Account_Code from BanksData where BankId = " & val(Dcbank.BoundText)
Set rsDummy = New ADODB.Recordset
rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly
If Not rsDummy.EOF Then
    cmbAccount.BoundText = Trim(rsDummy!Account_code & "")
End If



End Sub

Private Sub DcCurrency_Change()

    If Me.TxtModFlg.text = "" Or Me.TxtModFlg.text = "R" Then Exit Sub
    If Me.Dccurrency.BoundText <> "" Then
        txt_Currency_rate.text = get_currency_rate(Me.Dccurrency.BoundText)
    Else
        txt_Currency_rate.text = 1
    End If

End Sub

Private Sub DcCurrency_Click(Area As Integer)
    DcCurrency_Change
End Sub


Private Sub DCLC_Change()
If Me.TxtModFlg.text = "" Or Me.TxtModFlg.text = "R" Then Exit Sub
    s = "Select * from LCTypes  Where Id = " & val(DCLC.BoundText)
    Set rsDummy = New ADODB.Recordset
    rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly
    If Not rsDummy.EOF Then
            
        If val(rsDummy!TypeLCLG & "") = 0 Then
            optTypeLCLG(0).value = True
            optTypeLCLG_Click 0
            Frame1.Visible = False
        ElseIf val(rsDummy!TypeLCLG & "") = 1 Then
            optTypeLCLG(1).value = True
            optTypeLCLG_Click 1
            Frame1.Visible = True
        End If
        
        
    End If
   ' "TypeLCLG""
End Sub

Private Sub DCPreFix_Click(Area As Integer)

    Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos
    Dcombos.GetLCTypesName Me.DCLC, " prifix = '" & Trim(DCPreFix.text) & "'"
    
    
End Sub

Private Sub DpCloseDate_Change()
If Me.TxtModFlg.text = "" Or Me.TxtModFlg.text = "R" Then Exit Sub
    txtLGExpiryDate.value = DpCloseDate.value
    txtLGExpPeriod = DateDiff("D", txtGuaranteeDate.value, txtLGExpiryDate.value)

End Sub


Private Sub GrdBondHistory_AfterEdit(ByVal row As Long, ByVal Col As Long)
   Dim StrAccountCode As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long

    With GrdBondHistory
        .TextMatrix(1, .ColIndex("GuaranteeAmount")) = TxtValue
        Select Case .ColKey(Col)
        
         
           Case "AmountPlus"
            End Select
            If row > 1 Then
                .TextMatrix(row, .ColIndex("GuaranteeAmount")) = .TextMatrix(row - 1, .ColIndex("Total"))
            'ElseIf row > 2 Then
            
            End If
            .TextMatrix(row, .ColIndex("Total")) = val(.TextMatrix(row, .ColIndex("GuaranteeAmount"))) + val(.TextMatrix(row, .ColIndex("AmountPlus"))) - val(.TextMatrix(row, .ColIndex("AmountMin")))
            
            Me.txtTotalBondHistory.text = .TextMatrix(.rows - 2, .ColIndex("Total"))   '.Aggregate(flexSTSum, .FixedRows, .ColIndex("Total"), .rows - 1, .ColIndex("Total"))
             If row = .rows - 1 Then
                .rows = .rows + 1
            End If
            txtBondAmt = txtTotalBondHistory
    End With
    
    
    
    
End Sub


Private Sub GrdBondHistory_BeforeEdit(ByVal row As Long, ByVal Col As Long, Cancel As Boolean)
  With GrdBondHistory

        If row > .FixedRows Then
             If val(.TextMatrix(row, .ColIndex("NoteSerial"))) <> 0 And .ColKey(Col) <> "NoteSerial" Then
                  Cancel = True
                  Exit Sub
              End If
        End If

        Select Case .ColKey(Col)
            Case "Vat", "StillAmount"
                 Cancel = True
            Case "Total", "PayedAmount"
                .ComboList = ""
        Case "Vatyo"
             
            Case "value"
               Cancel = True
              Case "Amount", "MarginNo", "GuaranteeDate", "Serial"
                .ComboList = ""
              Case "ChSameCurrncy"
                .ComboList = ""

            Case "des"
                .ComboList = ""
        
            Case "Order_No"
                .ComboList = ""
        
                '  Cancel = True
            
        End Select

    End With

End Sub

Private Sub GrdBondHistory_CellButtonClick(ByVal row As Long, ByVal Col As Long)
With GrdBondHistory
Select Case .ColKey(Col)
       Case "NoteSerial"
'                                  LngRow = Row
        ShowGL_cc val(.TextMatrix(row, .ColIndex("NoteSerial"))), , 22004
 
    Case "CreateNote2"
        If val(.TextMatrix(row, .ColIndex("NoteId2"))) = 0 And val(.TextMatrix(row, .ColIndex("PayedAmount"))) <> 0 Then
            CreateEntry row, 1, 0
        End If
    Case "CreateNote"
        If val(.TextMatrix(row, .ColIndex("NoteId"))) = 0 And val(.TextMatrix(row, .ColIndex("Total"))) <> 0 Then
            CreateEntry row, 3, 0
        End If
    Case "DeleteEntry"
        s = "Delete Notes where NoteId = " & val(.TextMatrix(row, .ColIndex("NoteId")))
        Cn.Execute s
        .TextMatrix(row, .ColIndex("NoteId")) = ""
        .TextMatrix(row, .ColIndex("NoteSerial")) = ""
        
        MsgBox " „ Õ–› «·ﬁÌœ"
 End Select
End With

End Sub




Private Sub GrdMargin3_BeforeEdit(ByVal row As Long, ByVal Col As Long, Cancel As Boolean)
  With GrdMargin3

        If row > .FixedRows Then
             If val(.TextMatrix(row, .ColIndex("NoteSerial"))) <> 0 And .ColKey(Col) <> "NoteSerial" Then
                  Cancel = True
                  Exit Sub
              End If
        End If

        Select Case .ColKey(Col)
            Case "Vat", "StillAmount"
                 Cancel = True
            Case "Total", "PayedAmount"
                .ComboList = ""
        Case "Vatyo"
             
            Case "value"
               Cancel = True
              Case "Amount", "MarginNo", "GuaranteeDate", "Serial", "AmountP", "ExpAmount"
                .ComboList = ""
              Case "ChSameCurrncy"
                .ComboList = ""

            Case "des"
                .ComboList = ""
        
            Case "Order_No"
                .ComboList = ""
        
                '  Cancel = True
            
        End Select

    End With

End Sub

Private Sub GrdMargin3_CellButtonClick(ByVal row As Long, ByVal Col As Long)

Dim Frm As New FrmDateOpProject, mIDD As Long, mUserID As Long, mDate As Date, mTime As Date
With GrdMargin3
Select Case .ColKey(Col)
        Case "GuaranteeDate"
            
            Frm.index = 610
            Me.LngRow = row
            Frm.show 1
            GrdMargin3_AfterEdit row, Col
       Case "NoteSerial"
'                                  LngRow = Row
        ShowGL_cc val(.TextMatrix(row, .ColIndex("NoteSerial"))), , 22006
 
    Case "CreateNote2"
        If val(.TextMatrix(row, .ColIndex("NoteId2"))) = 0 And val(.TextMatrix(row, .ColIndex("PayedAmount"))) <> 0 Then
            CreateEntry row, 1, 0
        End If
    Case "CreateNote"
        If val(.TextMatrix(row, .ColIndex("TotalAmount"))) <> 0 Then
         '   CreateEntry row, 4, 0
        End If
       Case "DeleteEntry"
        s = "Delete Notes where NoteId = " & val(.TextMatrix(row, .ColIndex("NoteId")))
        Cn.Execute s
        .TextMatrix(row, .ColIndex("NoteId")) = ""
        .TextMatrix(row, .ColIndex("NoteSerial")) = ""
        
        MsgBox " „ Õ–› «·ﬁÌœ"

 End Select
End With

End Sub



Private Sub GrdMargin_AfterEdit(ByVal row As Long, ByVal Col As Long)
 
    Dim StrAccountCode As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long

    With GrdMargin
        
        
        If .TextMatrix(row, .ColIndex("GuaranteeDate")) = "" Then
            .TextMatrix(row, .ColIndex("GuaranteeDate")) = txtGuaranteeDate.value
        End If
        
        If .TextMatrix(row, .ColIndex("BankAccountCode")) = "" Then
            .TextMatrix(row, .ColIndex("BankAccountCode")) = txtGuaranteeDate.value
             .TextMatrix(row, .ColIndex("BankAccountCode")) = get_bank_Account(Dcbank.BoundText, "Account_Code")
             s = "Select * from Accounts where Account_Code = '" & Trim(.TextMatrix(row, .ColIndex("BankAccountCode"))) & "'"
             Set rsDummy = New ADODB.Recordset
             rsDummy.Open s, Cn, adOpenForwardOnly, adLockReadOnly
             If Not rsDummy.EOF Then
                .TextMatrix(row, .ColIndex("BankAccountSerial")) = Trim(rsDummy!account_serial & "")
                .TextMatrix(row, .ColIndex("BankAccount")) = Trim(rsDummy!account_name & "")
             End If
             

              
            'BankAccountSerial
            'BankAccount
        End If
        
        Select Case .ColKey(Col)
        
         
          Case "MarginAccountSerial"
                s = "SELECT accounts.Account_Name,accounts.Account_Code,Account_Serial  FROM accounts  WHERE Account_Serial = N'" & Trim(GrdMargin.TextMatrix(row, Col)) & "'"
                Set rsDummy = New ADODB.Recordset
                rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly
                If Not rsDummy.EOF Then


                    GrdMargin.TextMatrix(row, GrdMargin.ColIndex("MarginAccount")) = Trim(rsDummy!account_name & "")
                    GrdMargin.TextMatrix(row, GrdMargin.ColIndex("MarginAccountCode")) = Trim(rsDummy!Account_code & "")
                End If
                
           Case "MarginAccount"
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("MarginAccountCode"), False, True)
                .TextMatrix(row, .ColIndex("MarginAccountCode")) = StrAccountCode
 
                s = "SELECT accounts.Account_Name,accounts.Account_Code,Account_Serial  FROM accounts  WHERE Account_Code = N'" & Trim(StrAccountCode) & "'"
                Set rsDummy = New ADODB.Recordset
                rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly
                If Not rsDummy.EOF Then


                    'GrdMargin2.TextMatrix(row, GrdMargin2.ColIndex("MarginAccount")) = Trim(rsDummy!account_name & "")
                    GrdMargin.TextMatrix(row, GrdMargin.ColIndex("MarginAccountSerial")) = Trim(rsDummy!account_serial & "")
                End If
 
            Case "BankAccountSerial"
                s = "SELECT accounts.Account_Name,accounts.Account_Code,Account_Serial  FROM accounts  WHERE Account_Serial = N'" & Trim(GrdMargin.TextMatrix(row, Col)) & "'"
                Set rsDummy = New ADODB.Recordset
                rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly
                If Not rsDummy.EOF Then


                    GrdMargin.TextMatrix(row, GrdMargin.ColIndex("BankAccount")) = Trim(rsDummy!account_name & "")
                    GrdMargin.TextMatrix(row, GrdMargin.ColIndex("BankAccountCode")) = Trim(rsDummy!Account_code & "")
                End If
           Case "BankAccount"
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("BankAccountCode"), False, True)
                .TextMatrix(row, .ColIndex("BankAccountCode")) = StrAccountCode
                s = "SELECT accounts.Account_Name,accounts.Account_Code,Account_Serial  FROM accounts  WHERE Account_Code = N'" & Trim(StrAccountCode) & "'"
                Set rsDummy = New ADODB.Recordset
                rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly
                If Not rsDummy.EOF Then


                    'GrdMargin2.TextMatrix(row, GrdMargin2.ColIndex("MarginAccount")) = Trim(rsDummy!account_name & "")
                    GrdMargin.TextMatrix(row, GrdMargin.ColIndex("BankAccountSerial")) = Trim(rsDummy!account_serial & "")
                End If
            Case "AccountName"
                '  .TextMatrix(Row, .ColIndex("userid")) = user_id
                        
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("AccountCode"), False, True)
                .TextMatrix(row, .ColIndex("AccountCode")) = StrAccountCode
                .TextMatrix(row, .ColIndex("ExpensesID")) = get_Expenses_id(StrAccountCode)
                .TextMatrix(row, .ColIndex("LineNo1")) = setfoxy_Line

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrSQL = "select * from Expenses_accounts where Account_Code='" & StrAccountCode & "'"
                Else
                    StrSQL = "select * from Expenses_accounts_eng where Account_Code='" & StrAccountCode & "'"
                End If
            
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                     
'                If rs.RecordCount > 0 Then
'                    .TextMatrix(Row, .ColIndex("des")) = IIf(IsNull(rs("parent_account").value), "", rs("parent_account").value)
'                Else
'                    .TextMatrix(Row, .ColIndex("des")) = ""
'                End If

            Case "Amount", "Price", "ChSameCurrncy"
                Dim sgl As String
           
               
                
                
                
                
                '    sgl = "update  marakes_taklefa_temp  set value=0 where  line_no=" & Val(GrdMargin.TextMatrix(GrdMargin.Row, GrdMargin.ColIndex("LineNo1")))
                '     Cn.Execute sgl, , adExecuteNoRecords
        
                '  Me.XPTxtVal.text = Format(Me.XPTxtVal.text, SystemOptions.SysDefCurrencyForamt)
        End Select

   

        ' Me.XPTxtVal.text = Format(Me.XPTxtVal.text, SystemOptions.SysDefCurrencyForamt)
        'to Add new row if needed
        If row = .rows - 1 Then
            .rows = .rows + 1
        End If
          
        ' ReLineGrid
    End With
    Me.txtMarginTotal.text = GrdMargin.Aggregate(flexSTSum, GrdMargin.FixedRows, GrdMargin.ColIndex("Amount"), GrdMargin.rows - 1, GrdMargin.ColIndex("Amount"))
    'ReLineGrid


End Sub

Private Sub GrdMargin_BeforeEdit(ByVal row As Long, ByVal Col As Long, Cancel As Boolean)
  With GrdMargin

        If row > .FixedRows Then
            '  If .TextMatrix(Row - 1, .ColIndex("AccountCode")) = "" Then
            '      Cancel = True
            '  End If
        End If

        Select Case .ColKey(Col)
            Case "Vat", "NoteID", "PayedAmount", "StillAmount"
                 Cancel = True
        Case "Vatyo"
             
            Case "value"
               Cancel = True
              Case "Amount", "MarginNo", "DueDate", "Serial"
                .ComboList = ""
              Case "ChSameCurrncy"
                .ComboList = ""

            Case "des"
                .ComboList = ""
        
            Case "Order_No"
                .ComboList = ""
        
                '  Cancel = True
            
        End Select

    End With
End Sub

Private Sub GrdMargin_CellButtonClick(ByVal row As Long, ByVal Col As Long)

Dim Frm As New FrmDateOpProject, mIDD As Long, mUserID As Long, mDate As Date, mTime As Date

With GrdMargin
Select Case .ColKey(Col)
       
        Case "GuaranteeDate"
            
            Frm.index = 611
            Me.LngRow = row
            Frm.show 1
       
       Case "NoteSerial"
                                 ' LngRow = Row
        ShowGL_cc val(.TextMatrix(row, .ColIndex("NoteSerial"))), , 22002
     Case "CreateNote"
        If val(.TextMatrix(row, .ColIndex("NoteId"))) = 0 And val(.TextMatrix(row, .ColIndex("Amount"))) <> 0 Then
            CreateEntry row, 0, 0
        End If
 End Select
End With

End Sub

Private Sub GrdMargin_KeyUp(KeyCode As Integer, Shift As Integer)
 With GrdMargin

        Select Case .ColKey(.Col)

            Case "Order_No"
                           
                If KeyCode = vbKeyF3 Then
                  
                    Order_no_search.show
                     Order_no_search.RetrunType = 4
                End If

            Case "MarginAccount"
                  If KeyCode = vbKeyF3 Then
                        Account_search.show
                        Account_search.case_id = 350057
                    End If
 
            Case "BankAccount"
                  If KeyCode = vbKeyF3 Then
                        Account_search.show
                        Account_search.case_id = 350058
                    End If
 
        End Select

    End With

End Sub

Private Sub GrdMargin_StartEdit(ByVal row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim rs As New ADODB.Recordset
    Dim Rs1 As New ADODB.Recordset

    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim StrComboList1 As String

    Dim Msg As String

    'Case "DebitName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a1%' Or ACCOUNTS.Account_Code Like 'a3%')"
    'Case "CreditName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a2%' Or ACCOUNTS.Account_Code Like 'a4%')"
    With GrdMargin

        Select Case .ColKey(Col)
          Case "MarginAccount"
                If SystemOptions.UserInterface = EnglishInterface Then
                
                    StrSQL = "SELECT Account_Code, Account_NameEng  from ACCOUNTS    "
           
                        StrSQL = StrSQL + " where (last_account=1)"
                    
                    StrSQL = StrSQL & GetAccountByBarnchUser
                    StrSQL = StrSQL & GetAccountCodeHiding
                Else
                
                    StrSQL = "SELECT   Account_Code  as Account_Code, Account_Name as Account_Name from ACCOUNTS   "
                     StrSQL = StrSQL + " where    (last_account=1)"

                   StrSQL = StrSQL & GetAccountByBarnchUser
                  StrSQL = StrSQL & GetAccountCodeHiding
                
                End If
                Set rs = New ADODB.Recordset
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                If SystemOptions.UserInterface = ArabicInterface Then
                StrComboList = .BuildComboList(rs, "Account_Name", "Account_Code")
                Else
                StrComboList = .BuildComboList(rs, "Account_NameEng", "Account_Code")
                End If
                
                
 
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
                
 Case "BankAccount"
                If SystemOptions.UserInterface = EnglishInterface Then
                
                    StrSQL = "SELECT Account_Code, Account_NameEng  from ACCOUNTS    "
           
                        StrSQL = StrSQL + " where (last_account=1)"
                    
                    StrSQL = StrSQL & GetAccountByBarnchUser
                    StrSQL = StrSQL & GetAccountCodeHiding
                Else
                
                    StrSQL = "SELECT   Account_Code  as Account_Code, Account_Name as Account_Name from ACCOUNTS   "
                     StrSQL = StrSQL + " where    (last_account=1)"

                   StrSQL = StrSQL & GetAccountByBarnchUser
                  StrSQL = StrSQL & GetAccountCodeHiding
                
                End If
                Set rs = New ADODB.Recordset
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                If SystemOptions.UserInterface = ArabicInterface Then
                StrComboList = .BuildComboList(rs, "Account_Name", "Account_Code")
                Else
                StrComboList = .BuildComboList(rs, "Account_NameEng", "Account_Code")
                End If
                
                
 
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
                
            Case "AccountName"

                '      StrSQL = "select * from Expenses_accounts"
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrSQL = "select * from Expenses_accounts"
                Else
                    StrSQL = "select * from Expenses_accounts_eng "
                End If
                 
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                'StrComboList = GrdMargin.BuildComboList(rs, "Account_Name", "Account_Code")
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = GrdMargin.BuildComboList(rs, "Account_Name", "Account_Code")
                Else
                    StrComboList = GrdMargin.BuildComboList(rs, "Account_NameEng", "Account_Code")
                End If
            
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList

            Case "opr_fullcode"
                StrSQL = "  select fullcode,name from terms_operations "
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                StrComboList1 = GrdMargin.BuildComboList(rs, "fullcode", "fullcode")

                If StrComboList1 <> "" Then
                    StrComboList1 = "|" & StrComboList1
                End If

                .ComboList = StrComboList1
         
        End Select

    End With


End Sub

 
 '------------------------
 
 
Private Sub GrdMargin2_AfterEdit(ByVal row As Long, ByVal Col As Long)
 
    Dim StrAccountCode As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long
    
    Dim s As String
    s = "Select Account_code from TblCustemers where CusID = " & val(DBCboClientName.BoundText)
    Set rsDummy = New ADODB.Recordset
    rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly
    If Not rsDummy.EOF Then
'        GrdMargin2.TextMatrix(row, GrdMargin2.ColIndex("MarginAccount")) = DBCboClientName.text ' Trim(rsDummy!Account_code & "")
'        GrdMargin2.TextMatrix(row, GrdMargin2.ColIndex("MarginAccountCode")) = Trim(rsDummy!Account_code & "")
        
    End If
    With GrdMargin2
        If .TextMatrix(row, .ColIndex("GuaranteeDate")) = "" Then
            .TextMatrix(row, .ColIndex("GuaranteeDate")) = txtGuaranteeDate.value
        End If
        
            
            
       
        If .TextMatrix(row, .ColIndex("BankAccountCode")) = "" Then
            .TextMatrix(row, .ColIndex("BankAccountCode")) = txtGuaranteeDate.value
             s = "SELECT AcceptAccount_Code,accounts.Account_Name,accounts.Account_Serial, * FROM TblLC tl Inner join accounts On tl.AcceptAccount_Code = accounts.Account_Code WHERE tl.TblLCID = " & val(TXTTblLCID.text)
            Set rsDummy = New ADODB.Recordset
            rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly
            If Not rsDummy.EOF Then
                
'                GrdMargin2.TextMatrix(row, GrdMargin2.ColIndex("BankAccountSerial")) = Trim(rsDummy!account_serial & "")
'                GrdMargin2.TextMatrix(row, GrdMargin2.ColIndex("BankAccount")) = Trim(rsDummy!account_name & "")
'                GrdMargin2.TextMatrix(row, GrdMargin2.ColIndex("BankAccountCode")) = Trim(rsDummy!AcceptAccount_Code & "")
            End If
             
'             s = "Select * from Accounts where Account_Code = '" & Trim(GrdMargin2.TextMatrix(row, GrdMargin2.ColIndex("BankAccountCode"))) & "'"
'             Set rsDummy = New ADODB.Recordset
'             rsDummy.Open s, Cn, adOpenForwardOnly, adLockReadOnly
'             If Not rsDummy.EOF Then
'                GrdMargin2.TextMatrix(row, GrdMargin2.ColIndex("BankAccountSerial")) = Trim(rsDummy!account_serial & "")
'                GrdMargin2.TextMatrix(row, GrdMargin2.ColIndex("BankAccount")) = Trim(rsDummy!account_name & "")
'             End If
'

              
            'BankAccountSerial
            'BankAccount
        End If

        Select Case .ColKey(Col)
            Case "OrderDate"
                If IsDate(.TextMatrix(row, .ColIndex("OrderDate"))) Then
                    .TextMatrix(row, .ColIndex("GuaranteeDate")) = DateAdd("D", val(txtAcceptianPeriod), .TextMatrix(row, .ColIndex("OrderDate")))
                End If
            Case "MarginNo"
            
                           StrSQL = "sELECT TBLLCMargin.*,BankAccountCode,BankAccountCode2,CC.Account_Name BankAccount,CC.Account_Serial BankAccountSerial, "
                StrSQL = StrSQL & "   accounts.Account_Name MarginAccount ,accounts.Account_Serial  MarginAccountSerial, MarginAccountCode,"
                StrSQL = StrSQL & "   CC4.Account_Name BankAccount2,CC4.Account_Serial BankAccountSerial2, "
                StrSQL = StrSQL & "   accounts.Account_Serial  MarginAccountSerial, MarginAccountCode"
                StrSQL = StrSQL & "   FROM TBLLCMargin Inner join accounts On accounts.Account_Code = TBLLCMargin.MarginAccountCode"
                StrSQL = StrSQL & "   Inner join accounts CC On CC.Account_Code = TBLLCMargin.BankAccountCode"
                
                StrSQL = StrSQL & "   left outer join accounts CC4 On CC4.Account_Code = TBLLCMargin.BankAccountCode2"

                'StrSQL = StrSQL & "   Where MarginNo = " & val(.TextMatrix(row, .ColIndex("MarginNo"))) & " and IsNull(TBLLCMargin.Amount,0) - IsNull(TBLLCMargin.PayedAmount,0) > 0  and TblLCID=" & val(TXTTblLCID.text)
                StrSQL = StrSQL & "   Where MarginNo = " & val(.TextMatrix(row, .ColIndex("MarginNo"))) & " and IsNull(TBLLCMargin.IsFullPayed,0) = 1  and TBLLCMargin.TblLCID=" & val(TXTTblLCID.text)
                StrSQL = StrSQL & "   Order by TBLLCMargin.ID desc"
                Set rsDummy = New ADODB.Recordset
                rsDummy.Open StrSQL, Cn, adOpenStatic, adLockReadOnly
                If Not rsDummy.EOF Then
                    MsgBox "«·›« Ê—… —ﬁ„ " & val(.TextMatrix(row, .ColIndex("MarginNo"))) & " ·Â–« «·«⁄ „«œ  „ ”œ«œÂ« »«·ﬂ«„·"
                     .TextMatrix(row, .ColIndex("MarginNo")) = ""
                     Exit Sub
                End If
                rsDummy.Close
            
                StrSQL = "sELECT TBLLCMargin.*,BankAccountCode,CC.Account_Name BankAccount,CC.Account_Serial BankAccountSerial, "
                StrSQL = StrSQL & "   accounts.Account_Name MarginAccount ,accounts.Account_Serial  MarginAccountSerial, MarginAccountCode,"
                StrSQL = StrSQL & "   CC4.Account_Name BankAccount2,CC4.Account_Serial BankAccountSerial2 "
                StrSQL = StrSQL & "   FROM TBLLCMargin Inner join accounts On accounts.Account_Code = TBLLCMargin.MarginAccountCode"
                StrSQL = StrSQL & "   Inner join accounts CC On CC.Account_Code = TBLLCMargin.BankAccountCode"
                StrSQL = StrSQL & "   left outer join accounts CC4 On CC4.Account_Code = TBLLCMargin.BankAccountCode2"

                'StrSQL = StrSQL & "   Where MarginNo = " & val(.TextMatrix(row, .ColIndex("MarginNo"))) & " and IsNull(TBLLCMargin.Amount,0) - IsNull(TBLLCMargin.PayedAmount,0) > 0  and TblLCID=" & val(TXTTblLCID.text)
                StrSQL = StrSQL & "   Where MarginNo = " & val(.TextMatrix(row, .ColIndex("MarginNo"))) & " and IsNull(TBLLCMargin.IsFullPayed,0) = 0  and TBLLCMargin.TblLCID=" & val(TXTTblLCID.text)
                StrSQL = StrSQL & "   Order by TBLLCMargin.ID desc"
                Set rsDummy = New ADODB.Recordset
                rsDummy.Open StrSQL, Cn, adOpenStatic, adLockReadOnly
                If rsDummy.EOF Then
                 '   MsgBox "Â–« «·ﬁ—÷ €Ì— „”Ã· „‰ ﬁ»·"
                    .TextMatrix(row, .ColIndex("MarginAccountCode")) = get_bank_Account(Dcbank.BoundText, "Account_Code")
                    .TextMatrix(row, .ColIndex("MarginAccount")) = Dcbank.text
                    
                    
                     
                    .TextMatrix(row, .ColIndex("BankAccountCode2")) = .TextMatrix(row, .ColIndex("MarginAccountCode"))
                    .TextMatrix(row, .ColIndex("BankAccount2")) = .TextMatrix(row, .ColIndex("MarginAccount"))
                    
                'End If
                    
       
                Else

                    .TextMatrix(row, .ColIndex("BankAccount")) = rsDummy!BankAccount & ""
                    .TextMatrix(row, .ColIndex("BankAccountSerial")) = rsDummy!BankAccountSerial & ""
                    .TextMatrix(row, .ColIndex("BankAccountCode")) = rsDummy!BankAccountCode & ""

                    .TextMatrix(row, .ColIndex("BankAccount2")) = rsDummy!BankAccount2 & ""
                    .TextMatrix(row, .ColIndex("BankAccountSerial2")) = rsDummy!BankAccountSerial2 & ""
                    .TextMatrix(row, .ColIndex("BankAccountCode2")) = rsDummy!BankAccountCode2 & ""



                    .TextMatrix(row, .ColIndex("MarginAccountCode")) = rsDummy!MarginAccountCode & ""
                    .TextMatrix(row, .ColIndex("MarginAccountSerial")) = rsDummy!MarginAccountSerial & ""
                    .TextMatrix(row, .ColIndex("MarginAccount")) = rsDummy!MarginAccount & ""
                    .TextMatrix(row, .ColIndex("OrderDate")) = rsDummy!OrderDate & ""
                    .TextMatrix(row, .ColIndex("GuaranteeDate")) = rsDummy!GuaranteeDate & ""
                    .TextMatrix(row, .ColIndex("Amount")) = rsDummy!StillAmount & "" 'rsDummy!Amount & ""
                    .TextMatrix(row, .ColIndex("PayedAmount")) = 0 'val(rsDummy!Amount & "") - val(rsDummy!StillAmount & "")
                    .TextMatrix(row, .ColIndex("StillAmount")) = rsDummy!StillAmount & ""
                   ' .TextMatrix(row, .ColIndex("NoteID")) = rsDummy!NoteID & ""
                   ' .TextMatrix(row, .ColIndex("NoteSerial")) = rsDummy!NoteSerial & ""
                    
                    .TextMatrix(row, .ColIndex("IsFullPayed")) = val(rsDummy!IsFullPayed & "")
                    
                    '.TextMatrix(Row, .ColIndex("StillAmount")) = rsDummy!StillAmount & ""
                End If
'

          Case "AccountMargen2Serial"
                s = "SELECT accounts.Account_Name,accounts.Account_Code,Account_Serial  FROM accounts  WHERE Account_Serial = N'" & Trim(GrdMargin2.TextMatrix(row, Col)) & "'"
                Set rsDummy = New ADODB.Recordset
                rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly
                If Not rsDummy.EOF Then


                    GrdMargin2.TextMatrix(row, GrdMargin2.ColIndex("AccountMargen2Name")) = Trim(rsDummy!account_name & "")
                    GrdMargin2.TextMatrix(row, GrdMargin2.ColIndex("AccountMargen2")) = Trim(rsDummy!Account_code & "")
                End If
           Case "MarginAccountSerial"
                s = "SELECT accounts.Account_Name,accounts.Account_Code,Account_Serial  FROM accounts  WHERE Account_Serial = N'" & Trim(GrdMargin2.TextMatrix(row, Col)) & "'"
                Set rsDummy = New ADODB.Recordset
                rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly
                If Not rsDummy.EOF Then


                    GrdMargin2.TextMatrix(row, GrdMargin2.ColIndex("MarginAccount")) = Trim(rsDummy!account_name & "")
                    GrdMargin2.TextMatrix(row, GrdMargin2.ColIndex("MarginAccountCode")) = Trim(rsDummy!Account_code & "")
                End If
                
'                StrAccountCode = .ComboData
'                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("MarginAccountSerial"), False, True)
'                .TextMatrix(row, .ColIndex("MarginAccount")) = StrAccountCode
                
           Case "MarginAccount"
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("MarginAccountCode"), False, True)
                .TextMatrix(row, .ColIndex("MarginAccountCode")) = StrAccountCode
                
                  s = "SELECT accounts.Account_Name,accounts.Account_Code,Account_Serial  FROM accounts  WHERE Account_Code = N'" & Trim(StrAccountCode) & "'"
                Set rsDummy = New ADODB.Recordset
                rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly
                If Not rsDummy.EOF Then


                    'GrdMargin2.TextMatrix(row, GrdMargin2.ColIndex("MarginAccount")) = Trim(rsDummy!account_name & "")
                    GrdMargin2.TextMatrix(row, GrdMargin2.ColIndex("MarginAccountSerial")) = Trim(rsDummy!account_serial & "")
                End If
                
           Case "AccountMargen2Name"
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("AccountMargen2"), False, True)
                .TextMatrix(row, .ColIndex("AccountMargen2")) = StrAccountCode
                
                
                    s = "SELECT accounts.Account_Name,accounts.Account_Code,Account_Serial  FROM accounts  WHERE Account_Code = N'" & Trim(StrAccountCode) & "'"
                Set rsDummy = New ADODB.Recordset
                rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly
                If Not rsDummy.EOF Then


                    'GrdMargin2.TextMatrix(row, GrdMargin2.ColIndex("MarginAccount")) = Trim(rsDummy!account_name & "")
                    GrdMargin2.TextMatrix(row, GrdMargin2.ColIndex("AccountMargen2Serial")) = Trim(rsDummy!account_serial & "")
                End If
 
           Case "BankAccount"
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("BankAccountCode"), False, True)
                .TextMatrix(row, .ColIndex("BankAccountCode")) = StrAccountCode
                
                   s = "SELECT accounts.Account_Name,accounts.Account_Code,Account_Serial  FROM accounts  WHERE Account_Code = N'" & Trim(StrAccountCode) & "'"
                Set rsDummy = New ADODB.Recordset
                rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly
                If Not rsDummy.EOF Then


                    'GrdMargin2.TextMatrix(row, GrdMargin2.ColIndex("MarginAccount")) = Trim(rsDummy!account_name & "")
                    GrdMargin2.TextMatrix(row, GrdMargin2.ColIndex("BankAccountSerial")) = Trim(rsDummy!account_serial & "")
                End If
            Case "BankAccountSerial"
                s = "SELECT accounts.Account_Name,accounts.Account_Code,Account_Serial  FROM accounts  WHERE Account_Serial = N'" & Trim(GrdMargin2.TextMatrix(row, Col)) & "'"
                Set rsDummy = New ADODB.Recordset
                rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly
                If Not rsDummy.EOF Then


                    GrdMargin2.TextMatrix(row, GrdMargin2.ColIndex("BankAccount")) = Trim(rsDummy!account_name & "")
                    GrdMargin2.TextMatrix(row, GrdMargin2.ColIndex("BankAccountCode")) = Trim(rsDummy!Account_code & "")
                End If
                
            Case "BankAccount2"
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("BankAccountCode2"), False, True)
                .TextMatrix(row, .ColIndex("BankAccountCode")) = StrAccountCode
                
                   s = "SELECT accounts.Account_Name,accounts.Account_Code,Account_Serial  FROM accounts  WHERE Account_Code = N'" & Trim(StrAccountCode) & "'"
                Set rsDummy = New ADODB.Recordset
                rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly
                If Not rsDummy.EOF Then


                    'GrdMargin2.TextMatrix(row, GrdMargin2.ColIndex("MarginAccount")) = Trim(rsDummy!account_name & "")
                    GrdMargin2.TextMatrix(row, GrdMargin2.ColIndex("BankAccountSerial2")) = Trim(rsDummy!account_serial & "")
                End If
            Case "BankAccountSerial2"
                s = "SELECT accounts.Account_Name,accounts.Account_Code,Account_Serial  FROM accounts  WHERE Account_Serial = N'" & Trim(GrdMargin2.TextMatrix(row, Col)) & "'"
                Set rsDummy = New ADODB.Recordset
                rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly
                If Not rsDummy.EOF Then


                    GrdMargin2.TextMatrix(row, GrdMargin2.ColIndex("BankAccount2")) = Trim(rsDummy!account_name & "")
                    GrdMargin2.TextMatrix(row, GrdMargin2.ColIndex("BankAccountCode2")) = Trim(rsDummy!Account_code & "")
                End If
                
'                StrAccountCode = .ComboData
'                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("MarginAccountSerial"), False, True)
            Case "AccountName"
                '  .TextMatrix(Row, .ColIndex("userid")) = user_id
                        
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("AccountCode"), False, True)
                .TextMatrix(row, .ColIndex("AccountCode")) = StrAccountCode
                .TextMatrix(row, .ColIndex("ExpensesID")) = get_Expenses_id(StrAccountCode)
                .TextMatrix(row, .ColIndex("LineNo1")) = setfoxy_Line

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrSQL = "select * from Expenses_accounts where Account_Code='" & StrAccountCode & "'"
                Else
                    StrSQL = "select * from Expenses_accounts_eng where Account_Code='" & StrAccountCode & "'"
                End If
            
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                     
'                If rs.RecordCount > 0 Then
'                    .TextMatrix(Row, .ColIndex("des")) = IIf(IsNull(rs("parent_account").value), "", rs("parent_account").value)
'                Else
'                    .TextMatrix(Row, .ColIndex("des")) = ""
'                End If



            Case "Amount"
                
                .TextMatrix(row, .ColIndex("StillAmount")) = val(.TextMatrix(row, .ColIndex("Amount"))) - val(.TextMatrix(row, .ColIndex("PayedAmount")))
                If val(.TextMatrix(row, .ColIndex("StillAmount"))) = 0 Then
                    .TextMatrix(row, .ColIndex("IsFullPayed")) = 1
                    .cell(flexcpBackColor, .rows - 1, 1, .rows - 1, .Cols - 1) = vbGreen
                Else
                    .TextMatrix(row, .ColIndex("IsFullPayed")) = 0
                End If
            Case "PayedAmount"
                If .TextMatrix(row, .ColIndex("PayDate")) = "" Then
                    .TextMatrix(row, .ColIndex("PayDate")) = Date
                End If
                .TextMatrix(row, .ColIndex("StillAmount")) = val(.TextMatrix(row, .ColIndex("Amount"))) - val(.TextMatrix(row, .ColIndex("PayedAmount")))
        
                If val(.TextMatrix(row, .ColIndex("StillAmount"))) = 0 Then
                    .TextMatrix(row, .ColIndex("IsFullPayed")) = 1
                    .cell(flexcpBackColor, row, 1, row, .Cols - 1) = vbGreen
                Else
                    .TextMatrix(row, .ColIndex("IsFullPayed")) = 0
                End If
            
            Case "Amount", "Price", "ChSameCurrncy"
                Dim sgl As String
           
               
                
                
                
                
                '    sgl = "update  marakes_taklefa_temp  set value=0 where  line_no=" & Val(GrdMargin2.TextMatrix(GrdMargin2.Row, GrdMargin2.ColIndex("LineNo1")))
                '     Cn.Execute sgl, , adExecuteNoRecords
        
                '  Me.XPTxtVal.text = Format(Me.XPTxtVal.text, SystemOptions.SysDefCurrencyForamt)
        End Select

   

        ' Me.XPTxtVal.text = Format(Me.XPTxtVal.text, SystemOptions.SysDefCurrencyForamt)
        'to Add new row if needed
        If row = .rows - 1 Then
            .rows = .rows + 1
        End If
          
        ' ReLineGrid
    End With
    Me.txtMarginTotal2.text = GrdMargin2.Aggregate(flexSTSum, GrdMargin2.FixedRows, GrdMargin2.ColIndex("StillAmount"), GrdMargin2.rows - 1, GrdMargin2.ColIndex("StillAmount"))
    'ReLineGrid


End Sub

Private Sub GrdMargin2_BeforeEdit(ByVal row As Long, ByVal Col As Long, Cancel As Boolean)
  With GrdMargin2

        If row > .FixedRows Then
             If val(.TextMatrix(row, .ColIndex("NoteSerial"))) <> 0 And .ColKey(Col) <> "NoteSerial" And val(.TextMatrix(row, .ColIndex("NoteSerial2"))) <> 0 And .ColKey(Col) <> "NoteSerial2" And .ColKey(Col) <> "NoteSerial3" Then
'                  Cancel = True
'                  Exit Sub
              End If
        End If

        Select Case .ColKey(Col)
            Case "Vat", "StillAmount"
                 Cancel = True
            Case "Amount", "PayedAmount", "MargenAmount", "MargenValue"
                .ComboList = ""
        Case "Vatyo"
             
            Case "value"
               Cancel = True
              Case "Amount", "MarginNo", "GuaranteeDate", "Serial"
                .ComboList = ""
              Case "ChSameCurrncy"
                .ComboList = ""

            Case "des"
                .ComboList = ""
        
            Case "Order_No"
                .ComboList = ""
        
                '  Cancel = True
            
        End Select

    End With
End Sub

Private Sub GrdMargin2_CellButtonClick(ByVal row As Long, ByVal Col As Long)

Dim Frm As New FrmDateOpProject, mIDD As Long, mUserID As Long, mDate As Date, mTime As Date



With GrdMargin2
Select Case .ColKey(Col)
       
        Case "OrderDate"
            
            Frm.index = 612
            Me.LngRow = row
            Frm.show 1
            GrdMargin2_AfterEdit row, Col
        
       Case "GuaranteeDate"
            
            Frm.index = 613
            Me.LngRow = row
            Frm.show 1
           
       Case "PayDate"
            
            Frm.index = 617
            Me.LngRow = row
            Frm.show 1
       
       
       Case "NoteSerial"
'                                  LngRow = Row
        ShowGL_cc val(.TextMatrix(row, .ColIndex("NoteSerial"))), , 22002
     Case "NoteSerial2"
'                                  LngRow = Row
        ShowGL_cc val(.TextMatrix(row, .ColIndex("NoteSerial2"))), , 22003
     Case "NoteSerial3"
'                                  LngRow = Row
        ShowGL_cc val(.TextMatrix(row, .ColIndex("NoteSerial3"))), , 22007
        
    Case "CreateNote2"
        'If val(.TextMatrix(row, .ColIndex("NoteId2"))) = 0 And val(.TextMatrix(row, .ColIndex("PayedAmount"))) <> 0 Then
        If val(.TextMatrix(row, .ColIndex("PayedAmount"))) <> 0 Then
          '  CreateEntry row, 1, 1
        End If
    Case "CreateNote"
        'If val(.TextMatrix(row, .ColIndex("NoteId"))) = 0 And val(.TextMatrix(row, .ColIndex("Amount"))) <> 0 Then
        If val(.TextMatrix(row, .ColIndex("Amount"))) <> 0 Then
        '    CreateEntry row, 1, 0
        End If
    Case "CreateNote3"
        If val(.TextMatrix(row, .ColIndex("NoteId3"))) = 0 And val(.TextMatrix(row, .ColIndex("MargenValue"))) <> 0 Then
            CreateEntry row, 5, 0
        End If
    Case "DeleteEntry"
        s = "Delete Notes where NoteId = " & val(.TextMatrix(row, .ColIndex("NoteId")))
        Cn.Execute s
        s = "Delete Notes where NoteId = " & val(.TextMatrix(row, .ColIndex("NoteId3")))
        Cn.Execute s
        s = "Delete Notes where NoteId = " & val(.TextMatrix(row, .ColIndex("NoteId2")))
        Cn.Execute s
        
        .TextMatrix(row, .ColIndex("NoteId")) = ""
        .TextMatrix(row, .ColIndex("NoteId2")) = ""
        .TextMatrix(row, .ColIndex("NoteId3")) = ""
        .TextMatrix(row, .ColIndex("NoteSerial")) = ""
       ' .TextMatrix(row, .ColIndex("NoteSeria2")) = ""
        '.TextMatrix(row, .ColIndex("NoteSeria3")) = ""
        
        MsgBox " „ Õ–› «·ﬁÌÊœ"
   
        
 End Select
End With
End Sub

Private Sub GrdMargin2_KeyUp(KeyCode As Integer, Shift As Integer)
 With GrdMargin2

        Select Case .ColKey(.Col)

            Case "Order_No"
                           
                If KeyCode = vbKeyF3 Then
                  
                    Order_no_search.show
                     Order_no_search.RetrunType = 4
                End If

            Case "MarginAccount"
                  If KeyCode = vbKeyF3 Then
                        Account_search.show
                        Account_search.case_id = 350057
                    End If
 
            Case "AccountMargen2"
                  If KeyCode = vbKeyF3 Then
                        Account_search.show
                        Account_search.case_id = 350059
                    End If
            Case "BankAccount"
                  If KeyCode = vbKeyF3 Then
                        Account_search.show
                        Account_search.case_id = 350058
                    End If
 
 
 
          
            Case "BankAccount2"
                  If KeyCode = vbKeyF3 Then
                        Account_search.show
                        Account_search.case_id = 350070
                    End If
 
             Case "AccountMargen2Name"
                  If KeyCode = vbKeyF3 Then
                        Account_search.show
                        Account_search.case_id = 350060
                    End If
        End Select

    End With

End Sub

Private Sub GrdMargin2_StartEdit(ByVal row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim rs As New ADODB.Recordset
    Dim Rs1 As New ADODB.Recordset

    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim StrComboList1 As String

    Dim Msg As String

    'Case "DebitName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a1%' Or ACCOUNTS.Account_Code Like 'a3%')"
    'Case "CreditName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a2%' Or ACCOUNTS.Account_Code Like 'a4%')"
    With GrdMargin2

        Select Case .ColKey(Col)
        
 
          Case "MarginAccountSerial"
                If SystemOptions.UserInterface = EnglishInterface Then
                
                    StrSQL = "SELECT Account_Code, Account_Serial  from ACCOUNTS    "
           
                        StrSQL = StrSQL + " where (last_account=1)"
                    
                    StrSQL = StrSQL & GetAccountByBarnchUser
                    StrSQL = StrSQL & GetAccountCodeHiding
                Else
                
                    StrSQL = "SELECT   Account_Code  as Account_Code, Account_Serial from ACCOUNTS   "
                     StrSQL = StrSQL + " where    (last_account=1)"

                   StrSQL = StrSQL & GetAccountByBarnchUser
                  StrSQL = StrSQL & GetAccountCodeHiding
                
                End If
                Set rs = New ADODB.Recordset
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                If SystemOptions.UserInterface = ArabicInterface Then
                StrComboList = .BuildComboList(rs, "Account_Serial", "Account_Code")
                Else
                StrComboList = .BuildComboList(rs, "Account_Serial", "Account_Code")
                End If
                
                
 
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
        
        
          Case "MarginAccount"
                If SystemOptions.UserInterface = EnglishInterface Then
                
                    StrSQL = "SELECT Account_Code, Account_NameEng  from ACCOUNTS    "
           
                        StrSQL = StrSQL + " where (last_account=1)"
                    
                    StrSQL = StrSQL & GetAccountByBarnchUser
                    StrSQL = StrSQL & GetAccountCodeHiding
                Else
                
                    StrSQL = "SELECT   Account_Code  as Account_Code, Account_Name as Account_Name from ACCOUNTS   "
                     StrSQL = StrSQL + " where    (last_account=1)"

                   StrSQL = StrSQL & GetAccountByBarnchUser
                  StrSQL = StrSQL & GetAccountCodeHiding
                
                End If
                Set rs = New ADODB.Recordset
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                If SystemOptions.UserInterface = ArabicInterface Then
                StrComboList = .BuildComboList(rs, "Account_Name", "Account_Code")
                Else
                StrComboList = .BuildComboList(rs, "Account_NameEng", "Account_Code")
                End If
                
                
 
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
                
                
   Case "AccountMargen2Name"
                If SystemOptions.UserInterface = EnglishInterface Then
                
                    StrSQL = "SELECT Account_Code, Account_NameEng  from ACCOUNTS    "
           
                        StrSQL = StrSQL + " where (last_account=1)"
                    
                    StrSQL = StrSQL & GetAccountByBarnchUser
                    StrSQL = StrSQL & GetAccountCodeHiding
                Else
                
                    StrSQL = "SELECT   Account_Code  as Account_Code, Account_Name as Account_Name from ACCOUNTS   "
                     StrSQL = StrSQL + " where    (last_account=1)"

                   StrSQL = StrSQL & GetAccountByBarnchUser
                  StrSQL = StrSQL & GetAccountCodeHiding
                
                End If
                Set rs = New ADODB.Recordset
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                If SystemOptions.UserInterface = ArabicInterface Then
                StrComboList = .BuildComboList(rs, "Account_Name", "Account_Code")
                Else
                StrComboList = .BuildComboList(rs, "Account_NameEng", "Account_Code")
                End If
                
                
 
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
                
 Case "BankAccount"
                If SystemOptions.UserInterface = EnglishInterface Then
                
                    StrSQL = "SELECT Account_Code, Account_NameEng  from ACCOUNTS    "
           
                        StrSQL = StrSQL + " where (last_account=1)"
                    
                    StrSQL = StrSQL & GetAccountByBarnchUser
                    StrSQL = StrSQL & GetAccountCodeHiding
                Else
                
                    StrSQL = "SELECT   Account_Code  as Account_Code, Account_Name as Account_Name from ACCOUNTS   "
                     StrSQL = StrSQL + " where    (last_account=1)"

                   StrSQL = StrSQL & GetAccountByBarnchUser
                  StrSQL = StrSQL & GetAccountCodeHiding
                
                End If
                Set rs = New ADODB.Recordset
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                If SystemOptions.UserInterface = ArabicInterface Then
                StrComboList = .BuildComboList(rs, "Account_Name", "Account_Code")
                Else
                StrComboList = .BuildComboList(rs, "Account_NameEng", "Account_Code")
                End If
                
                
 
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
                
 Case "BankAccount2"
                If SystemOptions.UserInterface = EnglishInterface Then
                
                    StrSQL = "SELECT Account_Code, Account_NameEng  from ACCOUNTS    "
           
                        StrSQL = StrSQL + " where (last_account=1)"
                    
                    StrSQL = StrSQL & GetAccountByBarnchUser
                    StrSQL = StrSQL & GetAccountCodeHiding
                Else
                
                    StrSQL = "SELECT   Account_Code  as Account_Code, Account_Name as Account_Name from ACCOUNTS   "
                     StrSQL = StrSQL + " where    (last_account=1)"

                   StrSQL = StrSQL & GetAccountByBarnchUser
                  StrSQL = StrSQL & GetAccountCodeHiding
                
                End If
                Set rs = New ADODB.Recordset
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                If SystemOptions.UserInterface = ArabicInterface Then
                StrComboList = .BuildComboList(rs, "Account_Name", "Account_Code")
                Else
                StrComboList = .BuildComboList(rs, "Account_NameEng", "Account_Code")
                End If
                
                
 
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
                
            Case "AccountName"

                '      StrSQL = "select * from Expenses_accounts"
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrSQL = "select * from Expenses_accounts"
                Else
                    StrSQL = "select * from Expenses_accounts_eng "
                End If
                 
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                'StrComboList = GrdMargin2.BuildComboList(rs, "Account_Name", "Account_Code")
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = GrdMargin2.BuildComboList(rs, "Account_Name", "Account_Code")
                Else
                    StrComboList = GrdMargin2.BuildComboList(rs, "Account_NameEng", "Account_Code")
                End If
            
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList

            Case "opr_fullcode"
                StrSQL = "  select fullcode,name from terms_operations "
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                StrComboList1 = GrdMargin2.BuildComboList(rs, "fullcode", "fullcode")

                If StrComboList1 <> "" Then
                    StrComboList1 = "|" & StrComboList1
                End If

                .ComboList = StrComboList1
         
        End Select

    End With


End Sub


 
Private Sub GrdMargin3_AfterEdit(ByVal row As Long, ByVal Col As Long)
  
   Dim StrAccountCode As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long

    With GrdMargin3
        .TextMatrix(1, .ColIndex("Amount")) = TxtValue
        Select Case .ColKey(Col)
        
         
           Case "AmountP"
            End Select
            If row > 1 Then
                .TextMatrix(row, .ColIndex("Amount")) = .TextMatrix(row - 1, .ColIndex("TotalAmount"))
            'ElseIf row > 2 Then
            
            End If
            .TextMatrix(row, .ColIndex("TotalAmount")) = val(.TextMatrix(row, .ColIndex("Amount"))) + val(.TextMatrix(row, .ColIndex("AmountP")))
            .TextMatrix(row, .ColIndex("PercentA")) = txtPercentV
            .TextMatrix(row, .ColIndex("InsuranceAmount")) = val(.TextMatrix(row, .ColIndex("AmountP"))) * val(txtPercentV) / 100
            
           ' Me.txtMarginTotal3.text = .TextMatrix(.rows - 1, .ColIndex("TotalAmount"))  '.Aggregate(flexSTSum, .FixedRows, .ColIndex("Total"), .rows - 1, .ColIndex("Total"))
            Me.txtMarginTotal3.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("TotalAmount"), .rows - 1, .ColIndex("TotalAmount"))
             If row = .rows - 1 Then
                .rows = .rows + 1
            End If
            'txtBondAmt = txtTotalBondHistory
    End With

    
End Sub

Private Sub optTypeLCLG_Click(index As Integer)
If Me.TxtModFlg.text = "" Or Me.TxtModFlg.text = "R" Then Exit Sub

If index = 0 Then
    If optTypeLCLG(index).value = True Then
        cmbAccountAcceptanceParent.Enabled = True
    Else
        cmbAccountAcceptanceParent.Enabled = False
    End If
ElseIf index = 1 Then
    If optTypeLCLG(index).value = True Then
        cmbAccountAcceptanceParent.Enabled = False
    Else
        cmbAccountAcceptanceParent.Enabled = True
    End If
End If


End Sub

Private Sub txtGuaranteeDate_Change()
    txtLGExpPeriod = DateDiff("D", txtGuaranteeDate.value, txtLGExpiryDate.value)
End Sub

Private Sub txtLGExpiryDate_Change()
    txtLGExpPeriod = DateDiff("D", txtGuaranteeDate.value, txtLGExpiryDate.value)
End Sub

Private Sub txtLGExpPeriod_Change()
If val(txtLGExpPeriod) <> 0 Then
    txtCostDay = Round(val(txtOPenValue) / val(txtLGExpPeriod), 4)
    
    If year(txtGuaranteeDate.value) <> year(txtLGExpiryDate.value) Then
        txtLGExpPeriodEnd = DateDiff("D", txtGuaranteeDate.value, year(txtGuaranteeDate.value) & "-12-31")
    Else
        txtLGExpPeriodEnd = DateDiff("D", txtGuaranteeDate.value, (txtLGExpiryDate.value))
    End If
    txtCostLGYear = Round(val(txtCostDay) * val(txtLGExpPeriodEnd), 3)
    txtLGExpPeriodLast = Round(val(txtLGExpPeriod) - val(txtLGExpPeriodEnd), 3)
    If val(txtLGExpPeriodLast) < 0 Then txtLGExpPeriodLast = 0
    txtCostLGYearLast = Round(val(txtCostDay) * val(txtLGExpPeriodLast), 3)
End If
End Sub

 '------------------------
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 Private Sub TxtName_GotFocus()
    SwitchKeyboardLang LANG_ARABIC
End Sub

Private Sub txtNameE_GotFocus()
    SwitchKeyboardLang LANG_ENGLISH
End Sub

Private Sub ChkDetails_Click()
    FillGridWithData
End Sub
      
Function CuurentLogdata(Optional Currentmode As String)
    LogTextA = "    ‘«‘… " & ScreenNameArabic & CHR(13) & " —ﬁ„ «·«⁄ „«œ     " & TxtLcNo.text & CHR(13) & "   ‰Ê⁄ «·«⁄ „«œ    " & DCLC & CHR(13) & "      «·»‰ﬂ " & Dcbank & CHR(13) & "     ﬁÌ„… «·«⁄ „«œ  " & TxtValue & CHR(13) & "      «·⁄„·… " & Dccurrency & CHR(13) & "       › Õ «·«⁄ „«œ  " & dbFromDate & CHR(13) & "    €·ﬁ «·«⁄ „«œ " & dbTodate & CHR(13) & "      «‰ Â«¡ «·«⁄ „«œ   " & DpCloseDate & CHR(13) & "    «·œÊ·…   " & DCCountry & CHR(13) & "     «·„Ê—œ  " & DBCboClientName & CHR(13) & "    »‰ﬂ «·„Ê—œ   " & TXTBank2 & CHR(13) & "     ⁄œœ «·‘Õ‰«   " & TxtNoOfParcil & CHR(13) & "       «Œ— ‘Õ‰…  " & DPLastParcilDate & CHR(13) & "   ‘—Êÿ «· ”·Ì„    " & TxtRemarks
                    
    If ChkLocked.value = vbChecked Then
        LogTextA = LogTextA & CHR(13) & "  „ «Ìﬁ«› «·«⁄ „«œ "
    End If
                    
    LogTextA = LogTextA & CHR(13) & " ÿ»Ì⁄Â «·—’Ìœ «·«›  «ÕÌ   "

    If OptType(0).value = True Then
        LogTextA = LogTextA & "„œÌ‰"
    ElseIf OptType(1).value = True Then
        LogTextA = LogTextA & "œ«∆‰"
    ElseIf OptType(2).value = True Then
        LogTextA = LogTextA & "€Ì— „Õœœ"
    End If

    LogTextA = LogTextA & CHR(13) & " ﬁÌ„… «·—’Ìœ «·«›  «ÕÌ  " & TxtOpenBalance
                     
    LogTexte = "    Screen  " & ScreenNameEnglish & CHR(13) & "Lc NO" & TxtLcNo.text & CHR(13) & "LC Type" & DCLC & CHR(13) & "Bank" & Dcbank & CHR(13) & "LC Value" & TxtValue & CHR(13) & "Currency" & Dccurrency & CHR(13) & "Open Date" & dbFromDate & CHR(13) & "Close Date " & dbTodate & CHR(13) & " End Date " & DpCloseDate & CHR(13) & " Country" & DCCountry & CHR(13) & "     Supplier  " & DBCboClientName & CHR(13) & "  Supplier Bank" & TXTBank2 & CHR(13) & " No Of Shipments" & TxtNoOfParcil & CHR(13) & "  Last Shipment Data" & DPLastParcilDate & CHR(13) & " Terms of delivery  " & TxtRemarks
                    
    If ChkLocked.value = vbChecked Then
        LogTexte = LogTexte & CHR(13) & "LC Locked "
    End If
                     
    If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, Me.TxtModFlg
    Else
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, "D"
    End If

End Function

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
        
    Msg = "ﬁÌœ «” Õﬁ«ﬁ —Ê« » «·„ÊŸ›Ì‰ ⁄‰ ‘Â— " & "   ”‰… "

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

    rs("NoteType").value = 66
    rs("NoteDate").value = Date
    rs("UserID").value = val(DCboUserName.BoundText)
    rs.update
   
    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
        
    Dim line_no As Integer
    line_no = 1

    With Grid

        For i = .FixedRows To .rows - 2

            If .TextMatrix(i, .ColIndex("project")) = "0" Then
                 
                If ModAccounts.AddNewDev(LngDevID, line_no, Account_Code_dynamic, .TextMatrix(i, .ColIndex("EmpTotalNet")), 0, Msg, val(notes_id), , , , Date, val(DCboUserName.BoundText)) = False Then
                    GoTo ErrTrap
                End If

            Else
                Account_Code_dynamic1 = get_project_Account(.TextMatrix(i, .ColIndex("project")), "Salary_account")

                If ModAccounts.AddNewDev(LngDevID, line_no, Account_Code_dynamic1, .TextMatrix(i, .ColIndex("EmpTotalNet")), 0, Msg, val(notes_id), , , , Date, val(DCboUserName.BoundText)) = False Then
                    GoTo ErrTrap
                End If
            End If
                 
            Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code1")
            StrAccountCode = Employee_account
        
            If ModAccounts.AddNewDev(LngDevID, line_no + 1, StrAccountCode, .TextMatrix(i, .ColIndex("EmpTotalNet")), 1, Msg, val(notes_id), , , , Date, val(DCboUserName.BoundText)) = False Then
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
        
    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
        
    Dim line_no As Integer
    line_no = 1

    With Grid

        For i = .FixedRows To .rows - 2

            If .TextMatrix(i, .ColIndex("project")) = "0" Then
                 
                If ModAccounts.AddNewDev(LngDevID, line_no, Account_Code_dynamic, .TextMatrix(i, .ColIndex("EmpTotalNet")), 0, Msg, , , , , Date, val(DCboUserName.BoundText)) = False Then
                    GoTo ErrTrap
                End If

            Else
                Account_Code_dynamic1 = get_project_Account(.TextMatrix(i, .ColIndex("project")), "Salary_account")

                If ModAccounts.AddNewDev(LngDevID, line_no, Account_Code_dynamic1, .TextMatrix(i, .ColIndex("EmpTotalNet")), 0, Msg, , , , , Date, val(DCboUserName.BoundText)) = False Then
                    GoTo ErrTrap
                End If
            End If
                 
            Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code1")
            StrAccountCode = Employee_account
        
            If ModAccounts.AddNewDev(LngDevID, line_no + 1, StrAccountCode, .TextMatrix(i, .ColIndex("EmpTotalNet")), 1, Msg, , , , , Date, val(DCboUserName.BoundText)) = False Then
                GoTo ErrTrap
            End If
        
            line_no = line_no + 2
   
        Next i

    End With

    Set rs = New ADODB.Recordset
    rs.Open "salary_voucher", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    rs.AddNew
 
    rs("voucher_id").value = LngDevID
  
    rs.update
 
    MsgBox " „ «‰‘«¡ «·ﬁÌœ", vbInformation
    create_report_data

    DoEvents

    Exit Function
ErrTrap:
    MsgBox "ÕœÀ Œÿ√ «À‰«¡ Õ›Ÿ «·»Ì«‰« ", vbExclamation
  
End Function

Private Sub ALLButton2_Click()
    'Dcemp.text = ""

    DCproject.text = ""
    FillGridWithData

    DoEvents
    Create_dev
    CmdOk_Click
End Sub

Private Sub ALLButton3_Click()
 
End Sub

'Private Sub CboPayMentType_Click()
'    'CboPayMentType_Change
'End Sub

Private Sub CboYear_Click()
    CmdOk_Click
End Sub

Private Sub Check1_Click()

    If Check1.value = vbChecked Then
        get_all_employee
    Else

        With Me.Grid
            .rows = 2
            .Clear flexClearScrollable
        End With

    End If

End Sub

Private Sub CmbMonth_Click()
    CmdOk_Click
    'FillGridWithData
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CmdOk_Click()

End Sub

Function create_report_data()

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
 
End Sub

Function RetriveProformaInvoices(LCNO As String)
    On Error GoTo ErrTrap
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim Num As Integer
    Set rs = New ADODB.Recordset
    StrSQL = " SELECT     dbo.Transactions.Transaction_ID, dbo.Transactions.LcNo, dbo.Transactions.Transaction_Date, dbo.Transactions.order_no, dbo.Transactions.CusID, "
    StrSQL = StrSQL & "  dbo.TblCustemers.CusName , dbo.TblCustemers.CusNamee"
    StrSQL = StrSQL & "  FROM         dbo.Transactions INNER JOIN"
    StrSQL = StrSQL & "   dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID"

    StrSQL = StrSQL & "   WHERE     (dbo.Transactions.LcNo = N'" & LCNO & "')"

    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If rs.RecordCount < 1 Then
        FG.Clear flexClearScrollable, flexClearEverything
        FG.rows = 2
  
        Exit Function
    End If

    FG.Clear flexClearScrollable, flexClearEverything
 
    'Set Me.FG.DataSource = rs
    If Not (rs.EOF Or rs.BOF) Then
        FG.rows = rs.RecordCount + 1

        For Num = 1 To rs.RecordCount

            With FG
                .TextMatrix(Num, .ColIndex("Count")) = Num
                .TextMatrix(Num, .ColIndex("noteserial1")) = IIf(IsNull(rs("order_no").value), "", (rs("order_no").value))
                .TextMatrix(Num, .ColIndex("Transaction_ID")) = IIf(IsNull(rs("Transaction_ID").value), "", val(rs("Transaction_ID").value))

                '    .TextMatrix(Num, .ColIndex("Transaction_Serial")) = IIf(IsNull(rs("Transaction_Serial").value), "", rs("Transaction_Serial").value)
                If Not IsNull(rs("Transaction_Date").value) Then
                    .TextMatrix(Num, .ColIndex("BillDate")) = rs("Transaction_Date").value
                Else
                    .TextMatrix(Num, .ColIndex("BillDate")) = ""
                End If

                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(Num, .ColIndex("ClientNmae")) = IIf(IsNull(rs("CusName").value), "", Trim(rs("CusName").value))
                Else
                    .TextMatrix(Num, .ColIndex("ClientNmae")) = IIf(IsNull(rs("CusNamee").value), "", Trim(rs("CusNamee").value))
                End If

                '   .TextMatrix(Num, .ColIndex("StorName")) = IIf(IsNull(rs("StoreName").value), "", Trim(rs("StoreName").value))
            End With

            rs.MoveNext
        Next Num

    End If

ErrTrap:

End Function

Private Sub SaveData()
    Dim Msg As String
    Dim RsTemp As New ADODB.Recordset
    Dim StrSQL As String
    Dim BeginTrans As Boolean
    Dim RsDev As ADODB.Recordset
    Dim LngDevID As Long
    Dim Account_Code_dynamic1 As String
Dim AccountExpensParent As String
    On Error GoTo ErrTrap

    If Me.TxtModFlg.text <> "R" Then
 
        Account_Code_dynamic1 = get_account_code_branch(62, my_branch)
                
        If Account_Code_dynamic1 = "NO branch" Then
            MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
            GoTo ErrTrap
        Else

            If Account_Code_dynamic1 = "NO account" Then
                MsgBox "·„ Ì „  ÕœÌœ Õ”«»  «›  «ÕÌ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                GoTo ErrTrap
            End If
        End If
 
        If Trim(Me.Dcbank.BoundText) = "" Then
            Msg = "ÌÃ» ≈Œ Ì«—   «·»‰ﬂ..!!"
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            Dcbank.SetFocus
            Sendkeys "{F4}"
            Exit Sub
        End If
    
    
        If Trim(Me.dcBranch.BoundText) = "" Then
            Msg = "ÌÃ» ≈Œ Ì«—   «·›—⁄..!!"
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            dcBranch.SetFocus
            Sendkeys "{F4}"
            Exit Sub
        End If
    
    
 
 
    If cmbAccountExpensParent.BoundText = "" Then
            MsgBox "„‰ ›÷·ﬂ «Œ — «·Õ”«» «·—∆Ì”Ì ", vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            cmbAccountExpensParent.SetFocus
            Sendkeys ("{F4}")
            Exit Sub
        End If
        
        If Trim(Me.TxtLcNo.text) = "" Then
            Msg = "ÌÃ» ﬂ «»… —ﬁ„ «·«⁄ „«œ   ..!!"
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            TxtLcNo.SetFocus
            Sendkeys "{F4}"
            Exit Sub
        End If
        
        If Trim(Me.DCLC.BoundText) = "" Then
            Msg = "ÌÃ» ≈Œ Ì«— ‰Ê⁄ «·«⁄ „«œ..!!"
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            DCLC.SetFocus
            Sendkeys "{F4}"
            Exit Sub
        End If
 
        If Trim(Me.Dcbank.BoundText) = "" Then
            Msg = "ÌÃ» ≈Œ Ì«—   «·»‰ﬂ..!!"
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            Dcbank.SetFocus
            Sendkeys "{F4}"
            Exit Sub
        End If
 
        If Trim(Me.Dccurrency.BoundText) = "" Then
            Msg = "ÌÃ» ≈Œ Ì«— ‰Ê⁄ «·⁄„·Â..!!"
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            Dccurrency.SetFocus
            Sendkeys "{F4}"
            Exit Sub
        End If
 
'        If Trim(Me.DCCountry.BoundText) = "" Then
'            Msg = "ÌÃ» ≈Œ Ì«—   «·œÊ·Â..!!"
'            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'            DCCountry.SetFocus
'            SendKeys "{F4}"
'            Exit Sub
'        End If
 
'        If Trim(Me.DBCboClientName.BoundText) = "" Then
'            Msg = "ÌÃ» ≈Œ Ì«— «·„Ê—œ..!!"
'            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'            DBCboClientName.SetFocus
'            SendKeys "{F4}"
'            Exit Sub
'        End If
 
    End If
    
    
       
        If Trim(Me.cmbAccountMarginParent.BoundText) = "" Then
            Msg = "ÌÃ» ≈Œ Ì«—   Õ”«» «·„«—Ã‰..!!"
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            cmbAccountMarginParent.SetFocus
            Sendkeys "{F4}"
            Exit Sub
        End If
'    If Trim(Me.cmbAccountLGParent.BoundText) = "" Then
'            Msg = "ÌÃ» ≈Œ Ì«—   Õ”«» «·÷„«‰..!!"
'            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'            cmbAccountLGParent.SetFocus
'            Sendkeys "{F4}"
'            Exit Sub
'        End If
       
       
If optTypeLCLG(1).value = True Then cmbAccountAcceptanceParent.Enabled = False Else cmbAccountAcceptanceParent.Enabled = True
If cmbAccountAcceptanceParent.Enabled Then
           If Trim(Me.cmbAccountAcceptanceParent.BoundText) = "" Then
            Msg = "ÌÃ» ≈Œ Ì«—   Õ”«» «·ﬁ»Ê·..!!"
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            cmbAccountAcceptanceParent.SetFocus
            Sendkeys "{F4}"
            Exit Sub
        End If
End If
  
    
        Dim mParetnAccount As String
        If SystemOptions.SuppCreat4Acc Then
'
'            s = "Select parent_account,ParetnAccount,PAcceptAccount_Code , PLCAccount_Code,PMarginAccount_Code from BanksData where BankId = " & val(Dcbank.BoundText)
'            Set rsDummy = New ADODB.Recordset
'            rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly
'            If Not rsDummy.EOF Then
'                    mParetnAccount = Trim(rsDummy!ParetnAccount & "")
'            End If
'            If mParetnAccount = "" Then
'                 Msg = "ÌÃ» ≈Œ Ì«— «·»‰ﬂ «Ê «⁄«œ… Õ›Ÿ «·»‰ﬂ „‰ „·› «·»‰Êﬂ ..!!"
'                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'                DCBank.SetFocus
'                Sendkeys "{F4}"
'                Exit Sub
'            End If
            
        End If


    '-------------------------------------------------------------------------------------------
   
    Cn.BeginTrans
    BeginTrans = True
    Dim Account_Code_dynamic As String
    Dim Account_Code_dynamicMargin As String
    
    Dim PAcceptAccount_Code As String
    Dim PLCAccount_Code As String
    Dim AccountExpensCode As String
    Dim des As String

    If SystemOptions.UserInterface = ArabicInterface Then
        des = "«⁄ „«œ   :"
    Else
        des = "LC  :"
    End If
    
        Account_Code_dynamic = get_account_code_branch(225, my_branch)
                
        If Account_Code_dynamic = "NO branch" Then
            MsgBox " ·„ Ì „  ÕœÌœ ›—⁄", vbCritical
            GoTo ErrTrap
        Else

            If Account_Code_dynamic = "NO account" Then
                MsgBox " Õ”«» «·«⁄ „«œ«  «·„” ‰œÌÂ €Ì— „⁄—›  «–Â» «·Ï «·—»ÿ „⁄ «·Õ”«»«  ", vbCritical
                GoTo ErrTrap
                 
            End If
        End If
        
                Account_Code_dynamic = get_account_code_branch(226, my_branch)
                
        If Account_Code_dynamic = "NO branch" Then
            MsgBox " ·„ Ì „  ÕœÌœ ›—⁄", vbCritical
            GoTo ErrTrap
        Else

            If Account_Code_dynamic = "NO account" Then
                MsgBox " Õ”«» «·«⁄ „«œ«  «·„” ‰œÌÂ €Ì— „⁄—›  «–Â» «·Ï «·—»ÿ „⁄ «·Õ”«»«  ", vbCritical
                GoTo ErrTrap
                 
            End If
        End If
    
'
    s = "Select PMarginAccount_Code,BankName,BankNamee,PAcceptAccount_Code , PLCAccount_Code from BanksData where BankId = " & val(Dcbank.BoundText)
    Set rsDummy = New ADODB.Recordset
    rsDummy.Open s, Cn, adOpenKeyset, adLockOptimistic
    Account_Code_dynamic = get_account_code_branch(225, my_branch)
    If Not rsDummy.EOF Then
'        If Trim(rsDummy!PMarginAccount_Code & "") = "" Then
'            Account_Code_dynamic = get_account_code_branch(225, my_branch)
'            rsDummy!PMarginAccount_Code = ModAccounts.AddNewAccount(Account_Code_dynamic, DCPreFix.text & " Margin " & Trim(rsDummy!BankName & "") & "", False, False, Trim(rsDummy!BankName & "") & "  Margin ")
'            rsDummy.update
        End If
'        Account_Code_dynamicMargin = Trim(rsDummy!PMarginAccount_Code & "")
'
'        If Trim(rsDummy!PAcceptAccount_Code & "") = "" Then
'            Account_Code_dynamic = get_account_code_branch(226, my_branch)
'            rsDummy!PAcceptAccount_Code = ModAccounts.AddNewAccount(Account_Code_dynamic, DCPreFix.text & " " & Trim(rsDummy!BankName & "") & " «⁄ „«œ«  „” ‰œÌ… „ﬁ»Ê·… ", False, False, Trim(rsDummy!BankNamee & "") & " Accept ")
'            rsDummy.update
'        End If
'        PAcceptAccount_Code = Trim(rsDummy!PAcceptAccount_Code & "")
'
'
'        If Trim(rsDummy!PLCAccount_Code & "") = "" Then
'            Account_Code_dynamic = get_account_code_branch(51, my_branch)
'            rsDummy!PLCAccount_Code = ModAccounts.AddNewAccount(Account_Code_dynamic, DCPreFix.text & " " & Trim(rsDummy!BankName & "") & "", False, False, DCPreFix.text & " " & Trim(rsDummy!BankNamee & ""))
'            rsDummy.update
'        End If
'        PLCAccount_Code = Trim(rsDummy!PLCAccount_Code & "")
'
'
'
'    End If
    
    
    If TxtModFlg.text = "N" Then
'        Account_Code_dynamic = get_account_code_branch(51, my_branch)
'
'        If Account_Code_dynamic = "NO branch" Then
'            MsgBox " ·„ Ì „  ÕœÌœ ›—⁄", vbCritical
'            GoTo ErrTrap
'        Else
'
'            If Account_Code_dynamic = "NO account" Then
'                MsgBox " Õ”«» «·«⁄ „«œ«  «·„” ‰œÌÂ €Ì— „⁄—›  «–Â» «·Ï «·—»ÿ „⁄ «·Õ”«»«  ", vbCritical
'                GoTo ErrTrap
'
'            End If
'        End If
    StrSQL = "select * From TBLLC  "
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    
        rs.AddNew
        rs!TblLCID = val(TXTTblLCID.text)
        Account_Code_dynamic = DboParentAccount.BoundText
        
        rs("parent_account").value = IIf(DboParentAccount.BoundText = "", Null, (DboParentAccount.text))
                
                
                



    'rs("Account_Code").value = ModAccounts.AddNewAccount(Account_Code_dynamic, DCPreFix.text & "LC " & "  " & TxtName.text & "  »—ﬁ„ :" & Trim$(Me.TxtLcNo.text), True, False, "  LC : " & TxtNameE.text & "  NO :" & Trim$(Me.TxtLcNo.text))
                    
    PAcceptAccount_Code = cmbAccountAcceptanceParent.BoundText
    If cmbAccountAcceptanceParent.Enabled Then
    
        PAcceptAccount_Code = cmbAccountAcceptanceParent.BoundText
        If PAcceptAccount_Code <> "" Then
             rs("AcceptAccount_Code").value = ModAccounts.AddNewAccount(PAcceptAccount_Code, DCPreFix.text & " Accepted" & "  " & Trim(rsDummy!BankName & " ") & TxtName.text & "  »—ﬁ„ :" & Trim$(Me.TxtLcNo.text), True, False, DCPreFix.text & "  Accepted: " & TxtNameE.text & "  NO :" & Trim$(Me.TxtLcNo.text), , , , , , , , , , , , , , , , , , , , , , , , , , , val(TXTTblLCID.text))
         End If
    End If
    
    Account_Code_dynamicMargin = cmbAccountMarginParent.BoundText
    If Account_Code_dynamicMargin <> "" Then
        rs("Account_CodeMargin").value = ModAccounts.AddNewAccount(Account_Code_dynamicMargin, DCPreFix.text & " Margin" & "  " & Trim(rsDummy!BankName & " ") & TxtName.text & "  »—ﬁ„ :" & Trim$(Me.TxtLcNo.text), True, False, DCPreFix.text & "  Margin : " & TxtNameE.text & "  NO :" & Trim$(Me.TxtLcNo.text), , , , , , , , , , , , , , , , , , , , , , , , , , , val(TXTTblLCID.text))
    End If
    
    
     PLCAccount_Code = cmbAccountLGParent.BoundText
   If PLCAccount_Code <> "" Then
        rs("LCAccount_Code").value = ModAccounts.AddNewAccount(PLCAccount_Code, DCPreFix.text & "  " & Trim(rsDummy!BankName & " ") & TxtName.text & "  »—ﬁ„ :" & Trim$(Me.TxtLcNo.text), True, False, DCPreFix.text & "   : " & TxtNameE.text & "  NO :" & Trim$(Me.TxtLcNo.text), , , , , , , , , , , , , , , , , , , , , , , , , , , val(TXTTblLCID.text))
   
    End If
    
    
    
     AccountExpensParent = cmbAccountExpensParent.BoundText
   If AccountExpensParent <> "" Then
        rs("AccountExpensCode").value = ModAccounts.AddNewAccount(AccountExpensParent, DCPreFix.text & "  " & Trim(rsDummy!BankName & " ") & TxtName.text & "  »—ﬁ„ :" & Trim$(Me.TxtLcNo.text), True, False, DCPreFix.text & "   : " & TxtNameE.text & "  NO :" & Trim$(Me.TxtLcNo.text), , , , , , , , , , , , , , , , , , , , , , , , , , , val(TXTTblLCID.text))
        rs("Account_Code").value = rs("AccountExpensCode")
    End If
     
     
    rs("Account_Code").value = rs("LCAccount_Code").value
    
    ElseIf Me.TxtModFlg.text = "E" Then

        
        If Trim(rs("AccountExpensCode") & "") <> "" Then
        
            If Trim(rs("LCAccount_Code").value & "") = "" Then
                rs("LCAccount_Code").value = rs("AccountExpensCode").value
            End If
            If Trim(rs("Account_Code").value & "") = "" Then
                rs("Account_Code").value = rs("AccountExpensCode").value
            End If
            ModAccounts.EditAccount rs("Account_Code").value, DCPreFix.text & "  " & Trim(rsDummy!BankName & " ") & TxtName.text & "  »—ﬁ„ :" & Trim$(Me.TxtLcNo.text), DCPreFix.text & ":  " & TxtNameE.text & "  No :" & Trim$(Me.TxtLcNo.text), , , , , , , , , , , , , , , , , True, val(TXTTblLCID.text)
            ModAccounts.EditAccount rs("LCAccount_Code").value, DCPreFix.text & "  " & Trim(rsDummy!BankName & " ") & TxtName.text & "  »—ﬁ„ :" & Trim$(Me.TxtLcNo.text), DCPreFix.text & " :  " & TxtNameE.text & "  No :" & Trim$(Me.TxtLcNo.text), , , , , , , , , , , , , , , , , True, val(TXTTblLCID.text)
            ModAccounts.EditAccount rs("AccountExpensCode").value, DCPreFix.text & "  " & "   " & Trim(rsDummy!BankName & " ") & TxtName.text & "  »—ﬁ„ :" & Trim$(Me.TxtLcNo.text), DCPreFix.text & "  :  " & TxtNameE.text & "  No :" & Trim$(Me.TxtLcNo.text), , , , , , , , , , , , , , , , , True, val(TXTTblLCID.text)
        Else
        
            AccountExpensParent = cmbAccountExpensParent.BoundText
            rs("AccountExpensCode").value = ModAccounts.AddNewAccount(AccountExpensParent, DCPreFix.text & "  " & "   " & Trim(rsDummy!BankName & " ") & TxtName.text & "  »—ﬁ„ :" & Trim$(Me.TxtLcNo.text), True, False, DCPreFix.text & "  : " & TxtNameE.text & "  NO :" & Trim$(Me.TxtLcNo.text), , , , , , , , , , , , , , , , , , , , , , , , , , , val(TXTTblLCID.text))
            rs("LCAccount_Code").value = rs("AccountExpensCode").value
            rs("Account_Code").value = rs("LCAccount_Code").value
            
            
        End If
    
                 
    
        If Not IsNull(rs("Account_CodeMargin").value) Then
            ModAccounts.EditAccount rs("Account_CodeMargin").value, DCPreFix.text & "  Margin" & "   " & Trim(rsDummy!BankName & " ") & TxtName.text & "  »—ﬁ„ :" & Trim$(Me.TxtLcNo.text), DCPreFix.text & " Margin: " & TxtNameE.text & "  No :" & Trim$(Me.TxtLcNo.text), , , , , , , , , , , , , , , , , True, val(TXTTblLCID.text)
        Else
            rs("Account_CodeMargin").value = ModAccounts.AddNewAccount(Account_Code_dynamicMargin, DCPreFix.text & "  Margin" & "   " & Trim(rsDummy!BankName & " ") & TxtName.text & "  »—ﬁ„ :" & Trim$(Me.TxtLcNo.text), True, False, DCPreFix.text & "   Margin : " & TxtNameE.text & "  NO :" & Trim$(Me.TxtLcNo.text), , , , , , , , , , , , , , , , , , , , , , , , , , , val(TXTTblLCID.text))
        End If
    
        
        
        
        If cmbAccountAcceptanceParent.Enabled Then
        PAcceptAccount_Code = cmbAccountAcceptanceParent.BoundText
        If Not IsNull(rs("AcceptAccount_Code").value) Then
            ModAccounts.EditAccount rs("AcceptAccount_Code").value, DCPreFix.text & "  Accepted" & "   " & Trim(rsDummy!BankName & " ") & TxtName.text & "  »—ﬁ„ :" & Trim$(Me.TxtLcNo.text), DCPreFix.text & "  Accepted:  " & TxtNameE.text & "  No :" & Trim$(Me.TxtLcNo.text), , , , , , , , , , , , , , , , , True, val(TXTTblLCID.text)
        Else
            rs("AcceptAccount_Code").value = ModAccounts.AddNewAccount(PAcceptAccount_Code, DCPreFix.text & "  Accepted" & "   " & Trim(rsDummy!BankName & " ") & TxtName.text & "  »—ﬁ„ :" & Trim$(Me.TxtLcNo.text), True, False, DCPreFix.text & "   Accepted : " & TxtNameE.text & "  NO :" & Trim$(Me.TxtLcNo.text), , , , , , , , , , , , , , , , , , , , , , , , , , , val(TXTTblLCID.text))
        End If
        End If
        
    
       
        
        
'
'            ModAccounts.EditAccount rs("Account_Code2").value, DCPreFix.text & "LC " & TxtName.text & "  »—ﬁ„ :" & Trim$(Me.TxtLcNo.text), "LC:  " & TxtNameE.text & "  No :" & Trim$(Me.TxtLcNo.text), , , , , , , , , , , , , , , , , True
'            ModAccounts.EditAccount
'
       
        
       
    
'
'        StrSQL = "delete From DOUBLE_ENTREY_VOUCHERS1 where opening_balance_voucher_id=" & val(txtopening_balance_voucher_id.text)
'        Cn.Execute StrSQL, , adExecuteNoRecords
        
        
        
        
'
'        StrSQL = "delete From Notes  where NoteId=" & val(TXTNoteID.text)
'        Cn.Execute StrSQL, , adExecuteNoRecords
        
        
'        StrSQL = "delete From Notes  where IsNull(TblLCID,0)=" & val(TXTTblLCID.text)
'        Cn.Execute StrSQL, , adExecuteNoRecords
'
        
       StrSQL = "delete From Notes  where IsNull(TblLCID,0)=" & val(TXTTblLCID.text)
        Cn.Execute StrSQL, , adExecuteNoRecords
        
         
       StrSQL = "delete from DOUBLE_ENTREY_VOUCHERS1 where Notes_ID In (Select  Noteid from Notes1  where IsNull(TblLCID,0)= " & val(TXTTblLCID.text) & " )"
       Cn.Execute StrSQL, , adExecuteNoRecords
       
       
       StrSQL = "delete From Notes1  where IsNull(TblLCID,0)=" & val(TXTTblLCID.text)
        Cn.Execute StrSQL, , adExecuteNoRecords
        
       
        
        StrSQL = "delete From TBLLCHistory where TblLCID=" & val(TXTTblLCID.text)
        Cn.Execute StrSQL, , adExecuteNoRecords
        
        
        StrSQL = "delete From TBLLCMargin where TblLCID=" & val(TXTTblLCID.text)
        Cn.Execute StrSQL, , adExecuteNoRecords
        
        
        StrSQL = "delete From TBLLCMargin2 where TblLCID=" & val(TXTTblLCID.text)
        Cn.Execute StrSQL, , adExecuteNoRecords
        
                
        StrSQL = "delete From tblLCOpenB where TblLCID=" & val(TXTTblLCID.text)
        Cn.Execute StrSQL, , adExecuteNoRecords
       
    End If
    
    '////////////////////////////////////////
    cmbAccountExpProject.BoundText = Trim(rs("AccountExpensCode").value & "")
    
    rs("OpenValue").value = val(txtOPenValue.text)
    rs("PaymentTypeID").value = val(CboPayMentType.ListIndex)
    rs("BoxID").value = val(DcboBox.BoundText)
    rs("BankID2").value = val(DcboBankName.BoundText)
    rs("ChequeNumber").value = val(TxtChequeNumber.text)
    rs("ChequeDueDate").value = DtpChequeDueDate.value
    rs("LGExpiryDate").value = txtLGExpiryDate.value
    
    rs("BranchID").value = val(dcBranch.BoundText)
    
    rs("AccountAcceptanceParent").value = Trim(cmbAccountAcceptanceParent.BoundText)
    rs("AccountLGParent").value = Trim(cmbAccountLGParent.BoundText)
    
    rs("AccountMarginParent").value = Trim(cmbAccountMarginParent.BoundText)
    rs("AccountExpensParent").value = Trim(cmbAccountExpensParent.BoundText)
    
    If optTypeLCLG(0).value = True Then
        rs!TypeLCLG = 0
    ElseIf optTypeLCLG(1).value = True Then
        rs!TypeLCLG = 1
    End If
    
    
    rs("prifix").value = IIf(DCPreFix.text = "", Null, DCPreFix.text)
    rs("project_id").value = val(DataCombo2.BoundText)
    rs("TblLCID").value = val(TXTTblLCID.text)
    rs("LCNO").value = IIf(Me.TxtLcNo.text = "", "", Me.TxtLcNo.text)
    rs("name").value = IIf(Me.TxtName.text = "", "", Me.TxtName.text)
    rs("namee").value = IIf(Me.TxtNameE.text = "", "", Me.TxtNameE.text)
    
     rs("TotalBondHistory").value = val(Me.txtTotalBondHistory.text)
     rs("MarginTotal").value = val(Me.txtMarginTotal.text)
     
     rs("AcceptianPeriod").value = val(Me.txtAcceptianPeriod.text)
      rs("LGExpPeriod").value = val(Me.txtLGExpPeriod.text)
    
    
    rs("UserID").value = val(DCboUserName.BoundText)
    rs("LCTyperId").value = IIf(Me.DCLC.BoundText = "", Null, Me.DCLC.BoundText)
    rs("BankId").value = IIf(Me.Dcbank.BoundText = "", Null, Me.Dcbank.BoundText)
    rs("Bank2").value = IIf(Me.TXTBank2.text = "", "", Me.TXTBank2.text)
    rs("Value").value = IIf(Not IsNumeric(Me.TxtValue.text), 0, Me.TxtValue.text)
    rs("PercentV").value = IIf(Not IsNumeric(Me.txtPercentV.text), 0, Me.txtPercentV.text)
    rs("BondAmt").value = IIf(Not IsNumeric(Me.txtBondAmt.text), 0, Me.txtBondAmt.text)
   
    rs("GuaranteeNo").value = IIf(Me.txtGuaranteeNo.text = "", "", Me.txtGuaranteeNo.text)
    
     rs("Account_CodeExp").value = IIf(cmbAccount.BoundText = "", "", cmbAccount.BoundText)
     rs("AccountExpProject").value = IIf(cmbAccountExpProject.BoundText = "", "", cmbAccountExpProject.BoundText)
    rs("PrimaryInvoiceNo").value = IIf(Me.TXtPrimaryInvoiceNo.text = "", "", Me.TXtPrimaryInvoiceNo.text)
    rs("CountryId").value = IIf(Me.DCCountry.BoundText = "", Null, Me.DCCountry.BoundText)
    rs("CurrencyId").value = IIf(Me.Dccurrency.BoundText = "", Null, Me.Dccurrency.BoundText)
    rs("Currency_rate").value = IIf(Not IsNumeric(txt_Currency_rate.text), 1, txt_Currency_rate.text)
    rs("VendorId").value = IIf(Me.DBCboClientName.BoundText = "", Null, Me.DBCboClientName.BoundText)
    rs("FromDate").value = dbFromDate.value
    rs("Todate").value = dbTodate.value
    rs("GuaranteeDate").value = txtGuaranteeDate.value
     
    
    
If Trim(txtNoteIDRowId & "") = "" Then
  rs!NoteIDRowId = "{" & GenerateGUID & "}"
Else
    If InStr(txtNoteIDRowId, "{") Then
    Else
        txtNoteIDRowId = "{" & Trim(txtNoteIDRowId) & "}"
    End If

  rs!NoteIDRowId = Trim(txtNoteIDRowId)
End If

    
If Trim(rs!NoteID2RowId & "") = "" Then
  rs!NoteID2RowId = "{" & GenerateGUID & "}"
Else
    If InStr(txtNoteID2RowId, "{") Then
    Else
        txtNoteID2RowId = "{" & Trim(txtNoteID2RowId) & "}"
    End If
  rs!NoteID2RowId = Trim(txtNoteID2RowId)
End If


If Trim(rs!NoteIDOpenRowId & "") = "" Then
  rs!NoteIDOpenRowId = "{" & GenerateGUID & "}"
Else

   If InStr(txtNoteIDOpenRowId, "{") Then
    Else
        txtNoteIDOpenRowId = "{" & Trim(txtNoteIDOpenRowId) & "}"
    End If
  rs!NoteIDOpenRowId = Trim(txtNoteIDOpenRowId)
End If






    rs("projectName").value = IIf(Me.txtprojectname.text = "", "", Me.txtprojectname.text)
    rs("Remarks").value = IIf(Me.TxtRemarks.text = "", "", Me.TxtRemarks.text)
 
    If ChkLocked.value = vbChecked Then
        rs("Locked").value = 1
    Else
        rs("Locked").value = 0
    End If
 
    rs("CloseDate").value = DpCloseDate.value
    rs("LastParcilDate").value = DPLastParcilDate.value
    rs("NoOfParcil").value = IIf(Not IsNumeric(Me.TxtNoOfParcil.text), 0, Me.TxtNoOfParcil.text)
    rs!MarginTotal3 = val(txtMarginTotal3.text)
    
    rs!NoteSerial = val(TxtNoteSerial)
    
    If val(TxtOpenBalance.text) = 0 Then
        txtopening_balance_voucher_id = 0
    End If
       
    If Me.OptType(0).value = True Or Me.OptType(1).value = True Then
       
        If val(Me.txtopening_balance_voucher_id.text) = 0 Then
            txtopening_balance_voucher_id.text = get_opening_balance_voucher_id
            
        End If '
    End If '
'    txtopening_balance_voucher_id.text = get_opening_balance_voucher_id
    rs("opening_balance_voucher_id").value = val(txtopening_balance_voucher_id.text)

    If Me.OptType(2).value = True Then
        rs("OpenBalance").value = 0
        rs("OpenBalanceType").value = Null
    ElseIf Me.OptType(0).value = True Then
        rs("OpenBalance").value = val(Me.TxtOpenBalance.text)
        rs("OpenBalanceType").value = 0
    ElseIf Me.OptType(1).value = True Then
        rs("OpenBalance").value = val(Me.TxtOpenBalance.text)
        rs("OpenBalanceType").value = 1
    End If
    rs!OpenValue = val(txtOPenValue)
    rs!MarginTotal2 = val(txtMarginTotal2)
    rs!MarginTotal4 = val(txtMarginTotal4)
    rs!MarginTotal2 = val(txtMarginTotal2)
    rs!MarginTotal3 = val(txtMarginTotal3)
    rs.update
     
    Dim StrDes As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrDes = "«·—’Ìœ «·≈›  «ÕÏ ·‹·«⁄ „«œ —ﬁ„ "
         StrDes = " Opening Balance For LC NO : "
    Else
        StrDes = " Opening Balance For LC NO : "
    End If
        

    
     StrSQL = "sELECT * FROM TBLLCHistory Where 1 = -1"
   
    saveGrid StrSQL, GrdBondHistory, "MarginNo", "", "TblLCID", val(Me.TXTTblLCID.text)
    
    
 

    
         StrSQL = "sELECT * FROM TBLLCMargin2 Where 1 = -1"
'IncrementID
    saveGrid StrSQL, GrdMargin4, "MarginNo", "IncrementID", "TblLCID", val(Me.TXTTblLCID.text), "Type", 0




    StrSQL = "SELECT * FROM TBLLCMargin Where 1 = -1"
   'IncrementID
    saveGrid StrSQL, GrdMargin2, "MarginNo", "IncrementID", "TblLCID", val(Me.TXTTblLCID.text)
    
    
        StrSQL = "SELECT * FROM tblLCOpenB Where 1 = -1"
   
    saveGrid StrSQL, GrdMargin3, "MarginNo", "IncrementID", "TblLCID", val(Me.TXTTblLCID.text)
    
  
    Cn.CommitTrans
    BeginTrans = False
    CuurentLogdata
    
    
Dim row As Long
StrSQL = "sELECT TBLLCMargin.*,Acc.Account_Serial as MarginAccountSerial,Acc.Account_Name as MarginAccount  ,"
StrSQL = StrSQL & " Acc2.Account_Serial as BankAccountSerial,Acc2.Account_Name as BankAccount  "
StrSQL = StrSQL & " from TBLLCMargin Left Outer join Accounts Acc On TBLLCMargin.MarginAccountCode = Acc.Account_Code "
StrSQL = StrSQL & " Left Outer join Accounts Acc2 On TBLLCMargin.BankAccountCode = Acc2.Account_Code"
StrSQL = StrSQL & " Where Type= 0 and  TBLLCMargin.TblLCID = " & val(Me.TXTTblLCID.text)

loadgrid StrSQL, GrdMargin, True, False
    With GrdMargin
    
    For row = 1 To .rows - 1

        If val(.TextMatrix(row, .ColIndex("StillAmount"))) = 0 Then
            .TextMatrix(row, .ColIndex("IsFullPayed")) = 1
            .cell(flexcpBackColor, row, 1, row, .Cols - 1) = vbGreen
        Else
            .TextMatrix(row, .ColIndex("IsFullPayed")) = 0
        End If
    Next
    End With

StrSQL = "sELECT TBLLCMargin.*,Acc.Account_Serial as MarginAccountSerial,Acc.Account_Name as MarginAccount,Acc3.Account_Name AccountMargen2Name  ,"
StrSQL = StrSQL & " Acc2.Account_Serial as BankAccountSerial,Acc2.Account_Name as BankAccount , "
StrSQL = StrSQL & " Acc4.Account_Serial as BankAccountSerial2,Acc4.Account_Name as BankAccount2  "
StrSQL = StrSQL & " from TBLLCMargin Left Outer join Accounts Acc On TBLLCMargin.MarginAccountCode = Acc.Account_Code "
StrSQL = StrSQL & " Left Outer join Accounts Acc2 On TBLLCMargin.BankAccountCode = Acc2.Account_Code"
StrSQL = StrSQL & " Left Outer join Accounts Acc3 On TBLLCMargin.AccountMargen2 = Acc3.Account_Code"
StrSQL = StrSQL & " Left Outer join Accounts Acc4 On TBLLCMargin.BankAccountCode2 = Acc4.Account_Code"
StrSQL = StrSQL & " Where   TBLLCMargin.TblLCID = " & val(Me.TXTTblLCID.text)
loadgrid StrSQL, GrdMargin2, True, False

        If GrdMargin.rows - 1 > 1 Then
            Me.txtMarginTotal.text = GrdMargin.Aggregate(flexSTSum, GrdMargin.FixedRows, GrdMargin.ColIndex("Amount"), GrdMargin.rows - 1, GrdMargin.ColIndex("Amount"))
        End If
        
        If GrdMargin2.rows - 1 > 1 Then
            Me.txtMarginTotal2.text = GrdMargin2.Aggregate(flexSTSum, GrdMargin2.FixedRows, GrdMargin2.ColIndex("StillAmount"), GrdMargin2.rows - 1, GrdMargin2.ColIndex("StillAmount"))
        End If
    
    With GrdMargin2
    
    For row = 1 To .rows - 1

        If val(.TextMatrix(row, .ColIndex("StillAmount"))) = 0 Then
            .TextMatrix(row, .ColIndex("IsFullPayed")) = 1
            .cell(flexcpBackColor, row, 1, row, .Cols - 1) = vbGreen
        Else
            .TextMatrix(row, .ColIndex("IsFullPayed")) = 0
        End If
    Next
    End With

StrSQL = "sELECT TBLLCMargin2.*,Acc.Account_Serial as MarginAccountSerial,Acc.Account_Name as MarginAccount,Acc3.Account_Name AccountMargen2Name  ,"
StrSQL = StrSQL & " Acc2.Account_Serial as BankAccountSerial,Acc2.Account_Name as BankAccount  "
StrSQL = StrSQL & " from TBLLCMargin2 Left Outer join Accounts Acc On TBLLCMargin2.MarginAccountCode = Acc.Account_Code "
StrSQL = StrSQL & " Left Outer join Accounts Acc2 On TBLLCMargin2.BankAccountCode = Acc2.Account_Code"
StrSQL = StrSQL & " Left Outer join Accounts Acc3 On TBLLCMargin2.AccountMargen2 = Acc3.Account_Code"
StrSQL = StrSQL & " Where   TBLLCMargin2.TblLCID = " & val(Me.TXTTblLCID.text)
loadgrid StrSQL, GrdMargin4, True, False

        If GrdMargin4.rows - 1 > 1 Then
            Me.txtMarginTotal4.text = GrdMargin4.Aggregate(flexSTSum, GrdMargin4.FixedRows, GrdMargin4.ColIndex("Amount"), GrdMargin4.rows - 1, GrdMargin4.ColIndex("Amount"))
        End If
        

    With GrdMargin4
    
    For row = 1 To .rows - 1

        If val(.TextMatrix(row, .ColIndex("StillAmount"))) = 0 Then
            .TextMatrix(row, .ColIndex("IsFullPayed")) = 1
            .cell(flexcpBackColor, row, 1, row, .Cols - 1) = vbGreen
        Else
            .TextMatrix(row, .ColIndex("IsFullPayed")) = 0
        End If
    Next
    End With
    
    
     Cn.BeginTrans
    BeginTrans = True
        CmdCreateV_Click
       
        If optTypeLCLG(0).value Then Command3_Click
    
    Cn.CommitTrans
    BeginTrans = False
'    rs.Resync adAffectCurrent
'       rs.Requery
    Select Case Me.TxtModFlg.text

        Case "N"
            Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & CHR(13)
            Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"

            '    GrdMargin.Enabled = False
            If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.Title) = vbYes Then
                Cmd_Click (0)
                Exit Sub
            End If

        Case "E"
            MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            '  GrdMargin.Enabled = False
    End Select

    TxtModFlg.text = "R"
    'End If

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
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
End Sub

Private Sub Cmd_Click(index As Integer)
    On Error GoTo ErrTrap
    Dim FirstPeriodDateInthisYear As Date
    Me.DCboUserName.BoundText = user_id
    Select Case index

        Case 0
     
            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If
            clear_all Me
            TxtModFlg.text = "N"
            clear_all Me
            Me.TXTTblLCID.text = CStr(new_id("TblLC", "TblLCiD", "", True))
       
            Me.dbFromDate.value = Date
            Me.dbTodate.value = Date
       
            'XPDtbTrans.SetFocus
            Grid.Clear flexClearScrollable, flexClearEverything
            Grid.rows = 1
            Grid.Enabled = True
            Option2.value = True
          
            If val(txtNoteIDOpen) <> 0 Then
                Command3.Enabled = False
                txtOPenValue.locked = True
                Command6.Enabled = True
            Else
                txtOPenValue.locked = False
                Command6.Enabled = False
                Command3.Enabled = True
            End If
          GrdBondHistory.Clear flexClearScrollable, flexClearEverything
          GrdBondHistory.rows = GrdBondHistory.rows + 1
          
              
           GrdMargin3.Clear flexClearScrollable, flexClearEverything
          GrdMargin3.rows = GrdMargin3.rows + 1
           
          
           GrdMargin.Clear flexClearScrollable, flexClearEverything
          GrdMargin.rows = GrdMargin.rows + 1
          
          
          GrdMargin2.Clear flexClearScrollable, flexClearEverything
          GrdMargin2.rows = GrdMargin2.rows + 1
          
          
          
          GrdMargin4.Clear flexClearScrollable, flexClearEverything
          GrdMargin4.rows = GrdMargin4.rows + 1
          
            getFirstPeriodDateInthisYear FirstPeriodDateInthisYear

            Me.Dtp.value = FirstPeriodDateInthisYear

            Dim Account_Code_dynamic As String
            Account_Code_dynamic = get_account_code_branch(51, my_branch)
        
            If Account_Code_dynamic = "NO branch" Then
                MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                GoTo ErrTrap
            Else

                If Account_Code_dynamic = "NO account" Then
                    MsgBox "·„ Ì „  ÕœÌœ Õ”«» ··«⁄ „«œ«  Ê «·÷„«‰«  «·»‰ﬂÌ…   ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
       
                End If
            End If
        
            DboParentAccount.BoundText = Account_Code_dynamic
            Me.DCboUserName.BoundText = user_id
            
            CboPayMentType.ListIndex = 2
               Case 11
                    On Error Resume Next
                    ShowAttachments "LC" & TxtLcNo.text, "0901201401"

        Case 1
Ele(3).Enabled = True
            If val(TxtNoteID2) <> 0 Then
                MsgBox "·« Ì„ﬂ‰ «· ⁄œÌ· ⁄·Ï «·«⁄ „«œ »⁄œ «‰‘«¡ «·ﬁÌœ"
                Exit Sub
            End If

            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If

            getFirstPeriodDateInthisYear FirstPeriodDateInthisYear

          
          GrdBondHistory.rows = GrdBondHistory.rows + 1
          
         
          GrdMargin.rows = GrdMargin.rows + 1
          GrdMargin2.rows = GrdMargin2.rows + 1
          GrdMargin3.rows = GrdMargin3.rows + 1
          
                    
          
          GrdMargin4.rows = GrdMargin4.rows + 1
            Me.Dtp.value = FirstPeriodDateInthisYear
 
            TxtModFlg.text = "E"
            '         Grid.Rows = Grid.Rows + 1
            Grid.Enabled = True
         
            CuurentLogdata
            Me.DCboUserName.BoundText = user_id
        Case 2
            C1Tab1.CurrTab = 0
            Dim m1 As Boolean, m2 As Boolean, m3 As Boolean
            Dim Msg As String
        
          
            SaveData
      '     CmdCreateV_Click
        Case 3
            Undo

        Case 4

            If val(TXTNoteID) <> 0 Then
                Msg = "·«Ì„ﬂ‰ «·«·€«¡ »⁄œ «‰‘«¡ «·ﬁÌœ"
                MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Exit Sub
            End If
            
            If DoPremis(Do_Delete, Me.Name, True) = False Then
                Exit Sub
            End If

            Del_Trans

        Case 5

            If DoPremis(Do_Search, Me.Name, True) = False Then
                Exit Sub
            End If

           ' Load FrmNotesSearch
           ' FrmNotesSearch.SearchType = 3
           ' FrmNotesSearch.show vbModalLastParcilDate
            Load FrmLC_Search

        Case 6
            Unload Me

        Case 7
print_report
            '   ViewDataList
        Case 20
            addrow

        Case 21
            RemoveGridRow
    End Select

    Exit Sub
ErrTrap:

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


'MySQL = "  SELECT BanksData_1.BankName, dbo.currency.name AS CurrencyName, dbo.LCTypes.name AS TypeName, dbo.LCTypes.namee AS TypeNameE, dbo.TblLC.CountryId,"
'MySQL = MySQL & "                    dbo.TblLC.BankId, dbo.TblLC.LCTyperId, dbo.TblCountriesData.CountryName, dbo.TblLC.Value, dbo.TblLC.LCNO, dbo.TblLC.Todate, dbo.TblLC.Name, dbo.TblLC.FromDate,"
' MySQL = MySQL & "                            dbo.TblLC.CloseDate, dbo.TblLC.LastParcilDate, dbo.TblLC.VendorId, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblLC.Namee, dbo.TblLC.OpenValue,"
'  MySQL = MySQL & "                           dbo.TblLC.Remarks, dbo.TblLC.NoOfParcil, dbo.TblLC.PaymentTypeID, dbo.TblLC.ChequeNumber, dbo.TblLC.ChequeDueDate, dbo.TblBoxesData.BoxName,"
' MySQL = MySQL & "                            dbo.BanksData.BankName AS BankName2"
'MySQL = MySQL & "           FROM     dbo.TblCountriesData RIGHT OUTER JOIN"
'  MySQL = MySQL & "                           dbo.TblCustemers RIGHT OUTER JOIN"
'     MySQL = MySQL & "                        dbo.TblBoxesData RIGHT OUTER JOIN"
'        MySQL = MySQL & "                     dbo.TblLC LEFT OUTER JOIN"
'    MySQL = MySQL & "                         dbo.BanksData ON dbo.TblLC.BankID2 = dbo.BanksData.BankID ON dbo.TblBoxesData.BoxID = dbo.TblLC.BoxID ON dbo.TblCustemers.CusID = dbo.TblLC.VendorId ON"
'         MySQL = MySQL & "                    dbo.TblCountriesData.CountryID = dbo.TblLC.CountryId LEFT OUTER JOIN"
'     MySQL = MySQL & "                        dbo.currency ON dbo.TblLC.CurrencyId = dbo.currency.id LEFT OUTER JOIN"
'  MySQL = MySQL & "                           dbo.LCTypes ON dbo.TblLC.LCTyperId = dbo.LCTypes.id LEFT OUTER JOIN"
'   MySQL = MySQL & "                          dbo.BanksData AS BanksData_1 ON dbo.TblLC.BankId = BanksData_1.BankID"

  
MySQL = "  SELECT BanksData_1.BankName, dbo.currency.name AS CurrencyName, dbo.LCTypes.name AS TypeName, dbo.LCTypes.namee AS TypeNameE, dbo.TblLC.CountryId,TblLC.prifix,TblLC.PercentV,TblLC.PrimaryInvoiceNo,  projects.Fullcode AS ProjectCode, projects.Project_name,projects.Project_nameE,"
  MySQL = MySQL & "                 TblLC.*, dbo.TblCountriesData.CountryName, "
 MySQL = MySQL & "                          dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, "
   MySQL = MySQL & "                       dbo.TblLC.Remarks, dbo.TblLC.NoOfParcil, dbo.TblLC.PaymentTypeID, dbo.TblLC.ChequeNumber, dbo.TblLC.ChequeDueDate, dbo.TblBoxesData.BoxName,"
   MySQL = MySQL & "                       dbo.BanksData.BankName AS BankName2, dbo.currency.nameE AS CurrencyNameE, dbo.BanksData.BankNamee AS BankNameE2,"
 MySQL = MySQL & "                         BanksData_1.BankNamee AS BankNameE, dbo.TblBoxesData.BoxNameE, projects.Fullcode as ProjectCode,"
 
 MySQL = MySQL & "                       AcceptedValue = ( Select Sum(Note_Value) + Sum(VAT) from notes where CashingType in ( 13,14)   and TradingContractID =  " & val(TXTTblLCID) & ")"
 
 MySQL = MySQL & "       FROM     dbo.TblCountriesData RIGHT OUTER JOIN"
     MySQL = MySQL & "                     dbo.TblCustemers RIGHT OUTER JOIN"
     MySQL = MySQL & "                     dbo.TblBoxesData RIGHT OUTER JOIN"
    MySQL = MySQL & "                      dbo.TblLC LEFT OUTER JOIN"
   MySQL = MySQL & "                       dbo.BanksData ON dbo.TblLC.BankID2 = dbo.BanksData.BankID ON dbo.TblBoxesData.BoxID = dbo.TblLC.BoxID ON dbo.TblCustemers.CusID = dbo.TblLC.VendorId ON"
  MySQL = MySQL & "                        dbo.TblCountriesData.CountryID = dbo.TblLC.CountryId LEFT OUTER JOIN"
     MySQL = MySQL & "                     dbo.currency ON dbo.TblLC.CurrencyId = dbo.currency.id LEFT OUTER JOIN"
    MySQL = MySQL & "                      dbo.LCTypes ON dbo.TblLC.LCTyperId = dbo.LCTypes.id LEFT OUTER JOIN"
      MySQL = MySQL & "                    dbo.BanksData AS BanksData_1 ON dbo.TblLC.BankId = BanksData_1.BankID"
      
      
    MySQL = MySQL & "                       LEFT OUTER JOIN"
      MySQL = MySQL & "                    dbo.projects ON dbo.TblLC.project_id = projects.Id"
    
      
      

  
  
  MySQL = MySQL & "  Where LCNO = " & val(TxtLcNo.text)
    
    If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\rpt_LC_Details2.rpt"
    Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\rpt_LC_Details_E.rpt"
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
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
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
        ' xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
        StrReportTitle = "" '& StrAccountName
      
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.EngCompanyName   ' RPTCompany_Name_Eng
   
    End If

    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault


 
  
 
End Function





Function Del_Trans()
    Dim Msg As String
    Dim IntRes As Integer
    Dim BegainTrans As Boolean
    Dim StrSQL As String

    On Error GoTo ErrTrap

    If TxtLcNo.text <> "" Then
        Msg = "”Ì „ Õ–› »Ì«‰«  «·«⁄ „«œ  —ﬁ„ " & CHR(13)
        Msg = Msg + (TxtLcNo.text) & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"
        IntRes = MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title)
    
        If IntRes = vbYes Then
            If Not rs.RecordCount < 1 Then
                DeleteOpeningBalance
                Cn.BeginTrans
                BegainTrans = True
            
               ' StrSQL = "delete From DOUBLE_ENTREY_VOUCHERS1 where opening_balance_voucher_id=" & val(txtopening_balance_voucher_id.text)
               ' Cn.Execute StrSQL, , adExecuteNoRecords
                
                        
                        
                StrSQL = "delete From Notes where TblLCID=" & val(TXTTblLCID.text)
                Cn.Execute StrSQL, , adExecuteNoRecords
                        
                        
                Cn.Execute "delete DOUBLE_ENTREY_VOUCHERS1 where Notes_ID In (Select NoteID from Notes1  where TblLCID=" & val(TXTTblLCID.text) & ")"
                        
                StrSQL = "delete From Notes1 where TblLCID=" & val(TXTTblLCID.text)
                
                Cn.Execute StrSQL, , adExecuteNoRecords
                
                
                
                StrSQL = "delete From Accounts where TblLCID=" & val(TXTTblLCID.text)
                Cn.Execute StrSQL, , adExecuteNoRecords

                        
        
                StrSQL = "delete From TBLLCHistory where TblLCID=" & val(TXTTblLCID.text)
                Cn.Execute StrSQL, , adExecuteNoRecords
        
                
                StrSQL = "delete From TBLLCMargin where   TblLCID=" & val(TXTTblLCID.text)
                Cn.Execute StrSQL, , adExecuteNoRecords
                
                
                
               
                StrSQL = "delete From TBLLCMargin2 where   TblLCID=" & val(TXTTblLCID.text)
                Cn.Execute StrSQL, , adExecuteNoRecords
                 
                 
                StrSQL = "delete From tblLCOpenB where   TblLCID=" & val(TXTTblLCID.text)
                Cn.Execute StrSQL, , adExecuteNoRecords
                
        
                Dim StrAccountCode As String
                StrAccountCode = rs("Account_Code").value
                StrSQL = "Delete From Accounts Where Account_Code='" & StrAccountCode & "'"
                Cn.Execute StrSQL, , adExecuteNoRecords

                CuurentLogdata ("D")
                rs.delete
                Cn.CommitTrans
                BegainTrans = False
                Msg = " „  ⁄„·Ì… «·Õ–›."
                MsgBox Msg, vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                XPBtnMove_Click 2

                If rs.RecordCount < 1 Then
                    clear_all Me
                    TxtModFlg_Change
                    '  XPTxtCurrent.Caption = 0
                    '  XPTxtCount.Caption = 0
                Else
                    Retrive
                End If
            End If
        End If

    Else
        clear_all Me
        Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        TxtModFlg_Change
        Exit Function
    End If

    TxtModFlg_Change
    Exit Function
ErrTrap:
    'If Err.Number = -2147217887 Then
    Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & CHR(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «·„Ê—œ "
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.Title
    rs.CancelUpdate

    If BegainTrans = True Then
        Cn.RollbackTrans
        BegainTrans = False
    End If

    'End If

End Function

Function DeleteOpeningBalance()
    Cmd_Click (1)
    OptType(2).value = True
    TxtOpenBalance.text = 0
    Cmd_Click (2)

End Function

Private Sub RemoveGridRow()

    With Me.Grid

        If .row <= 0 Then Exit Sub
        .RemoveItem .row
    End With

    ReLineGrid
End Sub

Function addrow()

    Dim wherestr As String

    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim rs2 As ADODB.Recordset
    Set rs2 = New ADODB.Recordset
    Dim RsUnit As ADODB.Recordset
    Set RsUnit = New ADODB.Recordset

    Dim j As Integer

    Dim sql As String
    Dim i As Integer
    Dim Msg  As String
    Dim lastrow As Integer
    Dim LngItemID As Integer

    If Option2.value = True Then
        If dcitems.BoundText = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "ÌÃ»       «Œ Ì«— «·’‰›  ...!!!"
            Else
                Msg = "must Specify item Name ...!!!"
            End If

            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            Exit Function
        End If

        wherestr = "  where ItemID= " & val(dcitems.BoundText)
    End If

    sql = "Select * from TblItems "

    If wherestr <> "" Then
        sql = sql & wherestr
    End If
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then Exit Function

    With Grid
 
        lastrow = .rows
    
        If Rs3.RecordCount > 0 Then
            .rows = Rs3.RecordCount + lastrow
            Rs3.MoveFirst
         
            For i = lastrow To Rs3.RecordCount + lastrow - 1
                .TextMatrix(i, .ColIndex("ItemId")) = IIf(IsNull(Rs3.Fields("ItemId").value), "", Rs3.Fields("ItemId").value)
                LngItemID = IIf(IsNull(Rs3.Fields("ItemId").value), "", Rs3.Fields("ItemId").value)
                       
                .TextMatrix(i, .ColIndex("ItemCode")) = IIf(IsNull(Rs3.Fields("ItemCode").value), "", Rs3.Fields("ItemCode").value)
                .TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(Rs3.Fields("ItemName").value), "", Rs3.Fields("ItemName").value)
                       
                'lllllllllllllll
                StrSQL = "SELECT TblItemsUnits.UnitID, TblUnites.UnitName "
                StrSQL = StrSQL + " FROM TblUnites INNER JOIN TblItemsUnits " & "ON TblUnites.UnitID = TblItemsUnits.UnitID "
                StrSQL = StrSQL + " Where TblItemsUnits.DefaultUnit=1 and  TblItemsUnits.ItemID=" & LngItemID
                StrSQL = StrSQL + " Order BY TblItemsUnits.SecOrder "
                 
                RsUnit.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                If RsUnit.RecordCount > 0 Then
                    RsUnit.MoveFirst
                    .TextMatrix(i, .ColIndex("UnitId")) = IIf(IsNull(RsUnit.Fields("UnitId").value), "", RsUnit.Fields("UnitId").value)
                    .TextMatrix(i, .ColIndex("UnitName")) = IIf(IsNull(RsUnit.Fields("UnitName").value), "", RsUnit.Fields("UnitName").value)
               
                End If

                RsUnit.Close
                       
                Rs3.MoveNext
            Next i
 
            '    .AutoSize 0, .Cols - 1, False
        End If

    End With
 
    Rs3.Close

    ReLineGrid

End Function

Private Sub Undo()
    On Error GoTo ErrTrap

    Select Case TxtModFlg.text

        Case "N"
            clear_all Me
            Me.TxtModFlg.text = "R"
            XPBtnMove_Click (1)

        Case "E"
 
            Retrive
            Me.TxtModFlg.text = "R"
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub Dcdep_Click(Area As Integer)
    CmdOk_Click
End Sub

Private Sub Dcedara_Click(Area As Integer)
    CmdOk_Click
End Sub

Private Sub Dcemp_Click(Area As Integer)
    CmdOk_Click
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

Private Sub CmdRemove_Click()
    Dim X As Integer

    If SystemOptions.UserInterface = EnglishInterface Then
        X = MsgBox("Confirm Delete This line", vbCritical + vbYesNo)
    Else
        X = MsgBox(" √ﬂÌœ «·Õ–›", vbCritical + vbYesNo)
    End If

    If X = vbNo Then Exit Sub
    
    If Grid.rows > 1 Then
        If Grid.rows = 2 Then
            Me.Grid.Clear flexClearScrollable, flexClearEverything
        Else

            If Me.Grid.rows > 1 Then
                If Me.Grid.row <> Me.Grid.FixedRows - 1 Then
                    Me.Grid.RemoveItem (Me.Grid.row)
                End If
            End If
        End If
    End If
            
    With Grid
            
    End With

End Sub

Private Sub DboParentAccount_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF3 Then
        Account_search.show
        Account_search.case_id = 555
    End If
End Sub

Private Sub dcproject_Click(Area As Integer)

    If DCproject.BoundText = "" Then Exit Sub
    My_SQL = " select  fullcode,des from projects_des where project_id=" & val(DCproject.BoundText)
    fill_combo Dcterm, My_SQL

End Sub

Private Sub Dcterm_Click(Area As Integer)

    If Dcterm.BoundText = "" Then Exit Sub

    My_SQL = " select  fullcode,name from terms_operations where term_fullcode='" & Dcterm.BoundText & "'"
    fill_combo dcopr, My_SQL
End Sub

Private Sub fg_Click()
    Dim StrSQL As String
    Dim Num As Integer
    Dim RowNum As Integer
    Dim StrQry As String
    Dim RsDetails As ADODB.Recordset
    Dim DateTemp As Date
    Dim Msg As String

    On Error GoTo ErrTrap
 
    If Not FG.TextMatrix(FG.row, 1) = "" Then
        FrmShowPrice.show
        FrmShowPrice.Retrive val(FG.TextMatrix(FG.row, 1))
     
    End If

ErrTrap:
End Sub

Private Sub Form_Load()


        
    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500
   
    ScreenNameArabic = " «·«⁄ „«œ«  «·„” ‰œÌ…  "
    ScreenNameEnglish = "LC  "
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish, "1"

    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
    Set Cmd(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
    Set Cmd(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
    Set Cmd(5).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    'Set CmdHelp.ButtonImage = MDIFrmMain.ImgLstTree.ListImages("Help").Picture
    'Set Cmd(7).ButtonImage = MDIFrmMain.ImgLstTree.ListImages("FillData").Picture
    Dim My_SQL As String

    Dim GrdBack As ClsBackGroundPic
    Set GrdBack = New ClsBackGroundPic

    With Me.Grid
        Set .WallPaper = GrdBack.Picture
     
    End With
    
    TranslateForm Me, True
    
    Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos
    Set cSearchDCombo = New clsDCboSearch
    LoadCombosData
    Set BKGrndPic = New ClsBackGroundPic
    Dcombos.GetCodeing Me.DCPreFix, 4, "LCTypes"
    
    Dcombos.GetCustomersSuppliers 0, Me.DBCboClientName, True
    'Dcombos.GetItemsNames dcitems
    Dcombos.GetBanks Me.Dcbank
    Dcombos.GetCountriesNames Me.DCCountry
    Dcombos.GetLCTypesName Me.DCLC
    Dcombos.GetCUrrencyNames Me.Dccurrency
   Dcombos.GetBranches dcBranch
    Dcombos.GetBoxes DcboBox
    Dcombos.GetBanks Me.DcboBankName
    
        
 
    My_SQL = "  select id,Project_name from Projects where not (Fullcode is null) and Fullcode <>N'""' "
    My_SQL = My_SQL & "  AND      branch_no in(" & Current_branchSql & ")"
    fill_combo DataCombo2, My_SQL
    
If SystemOptions.UserInterface = ArabicInterface Then
    CboPayMentType.AddItem ("‰ﬁœÏ")
    CboPayMentType.AddItem ("‘Ìﬂ")
    CboPayMentType.AddItem ("Œ’„ „‰ Õ”«»")
    CboPayMentType.AddItem ("ÕÊ«·… »‰ﬂÌ…")
Else
    CboPayMentType.AddItem ("Cash")
    CboPayMentType.AddItem ("cheque")
    CboPayMentType.AddItem ("Account")
    CboPayMentType.AddItem ("Bank")
End If
    
    
    Dcombos.GetAccountingCodes Me.DboParentAccount, False, True
    Dcombos.GetAccountingCodes Me.cmbAccount, True, False
    
    Dcombos.GetAccountingCodes Me.cmbAccountExpProject, True, False
    
    
    Dcombos.GetAccountingCodes Me.cmbAccountExpensParent, False, True
    Dcombos.GetAccountingCodes Me.cmbAccountMarginParent, False, True
    Dcombos.GetAccountingCodes Me.cmbAccountLGParent, False, True
    Dcombos.GetAccountingCodes Me.cmbAccountAcceptanceParent, False, True
    
    
    
    With Me.Grid
        .rows = 1
        .ExplorerBar = flexExSortShowAndMove
        .RowHeightMin = 300
        .ExtendLastCol = True
    End With
      
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    Set rs = New ADODB.Recordset
    StrSQL = "select * From TBLLC  "
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    XPBtnMove_Click 2
    Me.TxtModFlg.text = "R"

    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If

End Sub

Private Sub ChangeLang()
    ChKauto.Caption = "Auto"
    
    
    
    
    TranslateForm Me, True
    
    Cmd(0).Caption = "New"
    Cmd(1).Caption = "Edit"
    Cmd(2).Caption = "Save"
    Cmd(3).Caption = "Undo"
    Cmd(4).Caption = "Delete"
    Cmd(5).Caption = "Search"
    'Cmd(7).Caption = "Print"
    Cmd(6).Caption = "Exit"
    'CmdHelp.Caption = "Help"

lbl(29).Caption = "Branch"

    Dim XPic As IPictureDisp
    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic
    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic

    Me.Caption = "Lc Data"
    Ele(5).Caption = Me.Caption
   lbl(25).Caption = "Arabic Name"
   lbl(26).Caption = "English Name"
   
    lbl(7).Caption = "ID"
    lbl(4).Caption = "Type"
    lbl(4).Caption = "Type"
    lbl(6).Caption = "Bank"
    lbl(9).Caption = "Value"
    lbl(10).Caption = "Currency"
    lbl(5).Caption = "Open date"
    lbl(2).Caption = "Close date"
    lbl(21).Caption = "End date"
 
    lbl(11).Caption = "Performa Inv"
    lbl(12).Caption = "State"
    lbl(0).Caption = "Supplier"
    lbl(13).Caption = "Supplier Bank"
    lbl(20).Caption = "no of Shipments"
    lbl(22).Caption = "Last Shipment Date"
    ChkLocked.Caption = "Locked"
 
    lbl(3).Caption = "Remarks"
    lbl(14).Caption = "LC Expense"
    lbl(15).Caption = "payments Type"
    lbl(16).Caption = "Box"
    lbl(17).Caption = "Bank"
    lbl(18).Caption = "Cheque No"
    lbl(19).Caption = "Due Date"
    lbl(27).Caption = "Main Account"
    Cmd(7).Caption = "Print"

    CmdRemove.Caption = "Remove Line"

    Me.C1Tab1.TabCaption(0) = "Lc Date"
    Me.C1Tab1.TabCaption(1) = "Lc Opening Expenses"
    Me.C1Tab1.TabCaption(2) = "Linked Proforma Invoices"
 
    Me.Fra(1).Caption = "Open Balance State"
    OptType(0).Caption = "Debit"
    OptType(1).Caption = "Credit"
    OptType(2).Caption = "Un Sign"
    lbl(23).Caption = "Balance Value"
    lbl(24).Caption = "Rec Date"

    With Me.FG
        .TextMatrix(0, .ColIndex("count")) = "I"
        .TextMatrix(0, .ColIndex("NoteSerial1")) = "Proforma Invoice#"
        .TextMatrix(0, .ColIndex("BillDate")) = "BillDate"
        .TextMatrix(0, .ColIndex("ClientNmae")) = "Client Name"
  
    End With

    '
End Sub

Public Sub get_all_employee()
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim rs2 As ADODB.Recordset
    Set rs2 = New ADODB.Recordset
    Dim j As Integer

    Dim sql As String
    Dim i As Integer

    sql = "Select * from emp_all_details "
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then Exit Sub
 
    With Grid

        .rows = 2
        .Clear flexClearScrollable

        If Rs3.RecordCount > 0 Then
            .rows = Rs3.RecordCount + 1
            Rs3.MoveFirst
         
            For i = 1 To Rs3.RecordCount
                .TextMatrix(i, .ColIndex("Emp_id")) = IIf(IsNull(Rs3.Fields("Emp_id").value), "", Rs3.Fields("Emp_id").value)
                       
                .TextMatrix(i, .ColIndex("Emp_Code")) = IIf(IsNull(Rs3.Fields("Emp_Code").value), "", Rs3.Fields("Emp_Code").value)
                .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(Rs3.Fields("Emp_Name").value), "", Rs3.Fields("Emp_Name").value)
                .TextMatrix(i, .ColIndex("DepartmentName")) = IIf(IsNull(Rs3.Fields("DepartmentName").value), "", Rs3.Fields("DepartmentName").value)
                .TextMatrix(i, .ColIndex("JobTypeName")) = IIf(IsNull(Rs3.Fields("JobTypeName").value), "", Rs3.Fields("JobTypeName").value)
                .TextMatrix(i, .ColIndex("work_status")) = IIf(IsNull(Rs3.Fields("name").value), "", Rs3.Fields("name").value)
                       
                Rs3.MoveNext
            Next i
 
            .AutoSize 0, .Cols - 1, False
        End If

    End With
 
    Rs3.Close

End Sub

Public Sub FillGridWithData()
    Exit Sub

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

    On Error GoTo ErrTrap
 
    Set rs = New ADODB.Recordset
    Set rs2 = New ADODB.Recordset

    rs.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    With Me.Grid
        .rows = 2
        .Clear flexClearScrollable

        If rs.RecordCount > 0 Then
            .rows = rs.RecordCount + 1
            rs.MoveFirst

            For i = 1 To .rows - 1
        
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

        .rows = .rows + 1
        .TextMatrix(.rows - 1, .ColIndex("Ser")) = "«·√Ã„«·Ï"
        .IsSubtotal(.rows - 1) = True
        Dim SngTotal As Single
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary"), .rows - 1, .ColIndex("Emp_Salary"))
        .TextMatrix(.rows - 1, .ColIndex("Emp_Salary")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("EmpTotalNet"), .rows - 1, .ColIndex("EmpTotalNet"))
        .TextMatrix(.rows - 1, .ColIndex("EmpTotalNet")) = SngTotal
        net_value = SngTotal
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("CorrectEmpTotalNet"), .rows - 1, .ColIndex("CorrectEmpTotalNet"))
        .TextMatrix(.rows - 1, .ColIndex("CorrectEmpTotalNet")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary_sakn"), .rows - 1, .ColIndex("Emp_Salary_sakn"))
        .TextMatrix(.rows - 1, .ColIndex("Emp_Salary_sakn")) = SngTotal
        
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary_bus"), .rows - 1, .ColIndex("Emp_Salary_bus"))
        .TextMatrix(.rows - 1, .ColIndex("Emp_Salary_bus")) = SngTotal
        
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary_food"), .rows - 1, .ColIndex("Emp_Salary_food"))
        .TextMatrix(.rows - 1, .ColIndex("Emp_Salary_food")) = SngTotal
        
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary_others"), .rows - 1, .ColIndex("Emp_Salary_others"))
        .TextMatrix(.rows - 1, .ColIndex("Emp_Salary_others")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("OverTimePrice"), .rows - 1, .ColIndex("OverTimePrice"))
        .TextMatrix(.rows - 1, .ColIndex("OverTimePrice")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Mokafea"), .rows - 1, .ColIndex("Mokafea"))
        .TextMatrix(.rows - 1, .ColIndex("Mokafea")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("SalesCom"), .rows - 1, .ColIndex("SalesCom"))
        .TextMatrix(.rows - 1, .ColIndex("SalesCom")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("TotalAdvance"), .rows - 1, .ColIndex("TotalAdvance"))
        .TextMatrix(.rows - 1, .ColIndex("TotalAdvance")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("TotalDiscount"), .rows - 1, .ColIndex("TotalDiscount"))
        .TextMatrix(.rows - 1, .ColIndex("TotalDiscount")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("total1"), .rows - 1, .ColIndex("total1"))
        .TextMatrix(.rows - 1, .ColIndex("total1")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("total2"), .rows - 1, .ColIndex("total2"))
        .TextMatrix(.rows - 1, .ColIndex("total2")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary_mang"), .rows - 1, .ColIndex("Emp_Salary_mang"))
        .TextMatrix(.rows - 1, .ColIndex("Emp_Salary_mang")) = SngTotal
    
        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary_mob"), .rows - 1, .ColIndex("Emp_Salary_mob"))
        .TextMatrix(.rows - 1, .ColIndex("Emp_Salary_mob")) = SngTotal
    
        .cell(flexcpBackColor, .rows - 1, 1, .rows - 1, .Cols - 1) = vbYellow
        .cell(flexcpFontBold, .rows - 1, 1, .rows - 1, .Cols - 1) = True
        .cell(flexcpFontSize, .rows - 1, 1, .rows - 1, .Cols - 1) = 10
        .cell(flexcpFontName, .rows - 1, 1, .rows - 1, .Cols - 1) = "Tahoma"
        .AutoSize 0, .Cols - 1, False
    End With

ErrTrap:
End Sub
 
Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)

    Dim IntResult As String
    Dim StrMSG As String
    On Error GoTo ErrTrap

    If Me.TxtModFlg.text <> "R" Then

        Select Case Me.TxtModFlg.text

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

        IntResult = MsgBox(StrMSG, vbMsgBoxRight + vbYesNoCancel + vbMsgBoxRtlReading + vbQuestion, App.Title)

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

Private Sub Form_Unload(Cancel As Integer)
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish
End Sub

Private Sub Grid_AfterEdit(ByVal row As Long, _
                           ByVal Col As Long)
    On Error Resume Next
    Dim StrAccountCode As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long
    Dim code  As String

    With Grid

        Select Case .ColKey(Col)
 
            Case "UnitName"
                code = .ComboData
           
                '   LngRow = .FindRow(Code, .FixedRows, .ColIndex("UnitID"), False, True)
                .TextMatrix(row, .ColIndex("UnitID")) = code
                .TextMatrix(row, .ColIndex("UnitName")) = .ComboItem
 
        End Select
   
        If row = .rows - 1 Then
    
            '.Rows = .Rows + 1
        End If

        ReLineGrid
    End With

End Sub

Private Sub ReLineGrid()
    Dim IntCounter As Integer
    IntCounter = 0
    Dim i As Integer

    With Me.Grid

        For i = .FixedRows To .rows - 1
    
            If .TextMatrix(i, .ColIndex("ItemId")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("Ser")) = IntCounter
  
            End If

        Next i
   
    End With

End Sub

Private Sub Grid_BeforeEdit(ByVal row As Long, _
                            ByVal Col As Long, _
                            Cancel As Boolean)

    With Grid

        If .ColKey(Col) <> "UnitName" Then
       
            .ComboList = ""
        End If

    End With

End Sub

Private Sub Grid_StartEdit(ByVal row As Long, _
                           ByVal Col As Long, _
                           Cancel As Boolean)
    Dim rs As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim Msg As String
    Dim LngItemID As Integer
    Dim MyStrList As String

    With Me.Grid

        Select Case .ColKey(Col)

            Case "UnitName"

                LngItemID = val(.TextMatrix(.row, .ColIndex("ItemId")))

                'LngItemID = 1
                If LngItemID = 0 Then
                    Cancel = True
                Else
            
                    StrSQL = "SELECT TblItemsUnits.UnitID, TblUnites.UnitName "
                    StrSQL = StrSQL + " FROM TblUnites INNER JOIN TblItemsUnits " & "ON TblUnites.UnitID = TblItemsUnits.UnitID "
                    StrSQL = StrSQL + " Where TblItemsUnits.ItemID=" & LngItemID
                    StrSQL = StrSQL + " Order BY TblItemsUnits.SecOrder "
                    Set rs = New ADODB.Recordset
                    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                    If Not (rs.BOF Or rs.EOF) Then
                        MyStrList = .BuildComboList(rs, "UnitName", "UnitID")
                        '                    Grid.ColComboList = MyStrList
                        Grid.ColComboList(.ColIndex("UnitName")) = "|" & MyStrList
                    Else
                        Cancel = True
                    End If
                End If
            
        End Select

    End With

End Sub

Public Sub Retrive(Optional LCNO As String = "")
    Dim RsDev As ADODB.Recordset
    Dim StrSQL As String
    Dim i As Integer

    'On Error GoTo ErrTrap
    Grid.Clear flexClearScrollable, flexClearEverything
    Grid.rows = 1

    If rs.RecordCount < 1 Then
        Exit Sub
    End If

    If rs.EOF Or rs.BOF Then
        Exit Sub
    Else

        If LCNO <> "" Then
            rs.Find "LCNO='" & LCNO & "'", , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·« ÌÊÃœ «⁄ „«œ »Â–« «·—ﬁ„", vbCritical
                Else
                    MsgBox "Lc With This No Not Found", vbCritical
                End If

                Unload Me
                Exit Sub
            
            End If
        End If
    End If
 
'///////////////////////////////////

txtOPenValue.text = IIf(IsNull(rs("OpenValue").value), "", rs("OpenValue").value)
CboPayMentType.ListIndex = IIf(IsNull(rs("PaymentTypeID").value), 0, rs("PaymentTypeID").value)
DcboBox.BoundText = IIf(IsNull(rs("BoxID").value), "", rs("BoxID").value)
DcboBankName.BoundText = IIf(IsNull(rs("BankID2").value), "", rs("BankID2").value)
TxtChequeNumber.text = IIf(IsNull(rs("ChequeNumber").value), "", rs("ChequeNumber").value)
txtBondAmt.text = IIf(IsNull(rs("BondAmt").value), "", rs("BondAmt").value)
DtpChequeDueDate.value = IIf(IsNull(rs("ChequeDueDate").value), Date, rs("ChequeDueDate").value)


cmbAccountExpProject.BoundText = Trim(rs("AccountExpensCode").value & "")
cmbAccountAcceptanceParent.BoundText = IIf(IsNull(rs("AccountAcceptanceParent").value), "", rs("AccountAcceptanceParent").value)
cmbAccountLGParent.BoundText = IIf(IsNull(rs("AccountLGParent").value), "", rs("AccountLGParent").value)
cmbAccountMarginParent.BoundText = IIf(IsNull(rs("AccountMarginParent").value), "", rs("AccountMarginParent").value)
cmbAccountExpensParent.BoundText = IIf(IsNull(rs("AccountExpensParent").value), "", rs("AccountExpensParent").value)

txtopening_balance_voucher_id = rs("opening_balance_voucher_id").value & ""
    If val(rs!TypeLCLG & "") = 0 Then
        optTypeLCLG(0).value = True
        
    ElseIf val(rs!TypeLCLG & "") = 1 Then
        optTypeLCLG(1).value = True
        
    End If

    If optTypeLCLG(1).value = True Then cmbAccountAcceptanceParent.Enabled = False Else cmbAccountAcceptanceParent.Enabled = True
    
txtprojectname.text = IIf(IsNull(rs("projectName").value), "", rs("projectName").value)
dcBranch.BoundText = IIf(IsNull(rs("BranchID").value), "", rs("BranchID").value)
cmbAccount.BoundText = IIf(IsNull(rs("Account_CodeExp").value), "", rs("Account_CodeExp").value)
cmbAccountExpProject.BoundText = IIf(IsNull(rs("AccountExpProject").value), "", rs("AccountExpProject").value)

DCPreFix.text = IIf(IsNull(rs("prifix").value), "", rs("prifix").value)
    
    
If Trim(rs!NoteIDRowId & "") = "" Then
  txtNoteIDRowId = GenerateGUID
Else
  txtNoteIDRowId = Trim(rs!NoteIDRowId & "")
End If

If Trim(rs!NoteID2RowId & "") = "" Then
  txtNoteID2RowId = GenerateGUID
Else
  txtNoteID2RowId = Trim(rs!NoteID2RowId & "")
End If


If Trim(rs!NoteIDOpenRowId & "") = "" Then
  txtNoteIDOpenRowId = GenerateGUID
Else
  txtNoteIDOpenRowId = Trim(rs!NoteIDOpenRowId & "")
End If



    DataCombo2.BoundText = IIf(IsNull(rs("project_id").value), "", rs("project_id").value)
Dim s As String

    Me.TXTTblLCID.text = IIf(IsNull(rs("TblLCID").value), "", rs("TblLCID").value)
    txtGuaranteeNo.text = IIf(IsNull(rs("GuaranteeNo").value), "", rs("GuaranteeNo").value)
    Me.TxtLcNo.text = IIf(IsNull(rs("LCNO").value), "", rs("LCNO").value)
        Me.TxtName.text = IIf(IsNull(rs("Name").value), "", rs("Name").value)
        
        
        Me.TxtNameE.text = IIf(IsNull(rs("Namee").value), "", rs("Namee").value)
    DboParentAccount.BoundText = Get_Account_Parent_code(IIf(IsNull(rs("Account_Code").value), "", Trim(rs("Account_Code").value)))
    
    RetriveProformaInvoices TxtLcNo.text
  
    Me.DCLC.BoundText = IIf(IsNull(rs("LCTyperId").value), "", rs("LCTyperId").value)
    Me.Dcbank.BoundText = IIf(IsNull(rs("BankId").value), "", rs("BankId").value)
    Me.TXTBank2.text = IIf(IsNull(rs("Bank2").value), "", rs("Bank2").value)
    Me.TxtValue.text = IIf(Not IsNumeric(rs("Value").value), 0, rs("Value").value)
    Me.txtPercentV.text = IIf(Not IsNumeric(rs("PercentV").value), 0, rs("PercentV").value)
    Me.txtAcceptianPeriod.text = IIf(Not IsNumeric(rs("AcceptianPeriod").value), 0, rs("AcceptianPeriod").value)
    
    Me.DCCountry.BoundText = IIf(IsNull(rs("CountryId").value), "", rs("CountryId").value)
    
    Me.Dccurrency.BoundText = IIf(IsNull(rs("CurrencyId").value), "", rs("CurrencyId").value)
    
    Dim mRate As Double
    
    s = "Select Rate From Currency Where Id  = " & val(Dccurrency.BoundText)

    Dim rsRate As New ADODB.Recordset
    
    rsRate.Open s, Cn, adOpenStatic, adLockOptimistic, adCmdText
    
    
    If rsRate.RecordCount <> 0 Then
        mRate = val(rsRate!Rate & "")
    Else
        mRate = 1
    End If

    
    txt_Currency_rate.text = IIf(IsNull(rs("Currency_rate").value), mRate, (rs("Currency_rate").value))
            
    
    Me.TXtPrimaryInvoiceNo.text = IIf(IsNull(rs("PrimaryInvoiceNo").value), "", rs("PrimaryInvoiceNo").value)
    Me.DCCountry.BoundText = IIf(IsNull(rs("CountryId").value), "", rs("CountryId").value)
    Me.DCboUserName.BoundText = IIf(IsNull(rs("UserID").value), user_id, rs("UserID").value)
    
    dbFromDate.value = IIf(IsNull(rs("FromDate").value), Date, rs("FromDate").value)
    dbTodate.value = IIf(IsNull(rs("Todate").value), Date, rs("Todate").value)
    
    
    

    
    
    TxtNoteSerial2 = rs!NoteSerial2 & ""
    TxtNoteID2 = rs!NoteID2 & ""

    
    
    TxtNoteSerial = rs!NoteSerial & ""
    TXTNoteID = rs!NoteID & ""
 
 
    txtNoteSerialOpen = rs!NoteSerialOpen & ""
    txtNoteIDOpen = rs!NoteIDOpen & ""
    

 

    DpCloseDate.value = IIf(IsNull(rs("CloseDate").value), Date, rs("CloseDate").value)
    DPLastParcilDate.value = IIf(IsNull(rs("LastParcilDate").value), Date, rs("LastParcilDate").value)
    Me.TxtNoOfParcil.text = IIf(Not IsNumeric(rs("NoOfParcil").value), 0, rs("NoOfParcil").value)
    
    Me.txtTotalBondHistory.text = IIf(Not IsNumeric(rs("TotalBondHistory").value), 0, rs("TotalBondHistory").value)
    Me.txtMarginTotal.text = IIf(Not IsNumeric(rs("MarginTotal").value), 0, rs("MarginTotal").value)
    Me.txtMarginTotal2.text = IIf(Not IsNumeric(rs("MarginTotal2").value), 0, rs("MarginTotal2").value)
    Me.txtMarginTotal3.text = IIf(Not IsNumeric(rs("MarginTotal3").value), 0, rs("MarginTotal3").value)

    DBCboClientName.BoundText = IIf(IsNull(rs("VendorId").value), "", rs("VendorId").value)

    TxtRemarks.text = IIf(IsNull(rs("Remarks").value), 0, rs("Remarks").value)

txtGuaranteeDate.value = IIf(IsNull(rs("GuaranteeDate").value), Date, rs("GuaranteeDate").value)

    txtLGExpiryDate.value = IIf(IsNull(rs("LGExpiryDate").value), Date, rs("LGExpiryDate").value)
txtLGExpPeriod.text = IIf(IsNull(rs("LGExpPeriod").value), "", rs("LGExpPeriod").value)

    If IsNull(rs("Locked").value) Then
        ChkLocked.value = vbUnchecked
    Else

        If rs("Locked").value = True Then
            ChkLocked.value = vbChecked
        Else
            ChkLocked.value = vbUnchecked
        End If

    End If
    
StrSQL = "sELECT * from TBLLCHistory  where TblLCID=" & val(Me.TXTTblLCID.text)
loadgrid StrSQL, GrdBondHistory, True, False
'
'
Dim row As Long
StrSQL = "sELECT TBLLCMargin.*,Acc.Account_Serial as MarginAccountSerial,Acc.Account_Name as MarginAccount  ,"
StrSQL = StrSQL & " Acc2.Account_Serial as BankAccountSerial,Acc2.Account_Name as BankAccount  "
StrSQL = StrSQL & " from TBLLCMargin Left Outer join Accounts Acc On TBLLCMargin.MarginAccountCode = Acc.Account_Code "
StrSQL = StrSQL & " Left Outer join Accounts Acc2 On TBLLCMargin.BankAccountCode = Acc2.Account_Code"
StrSQL = StrSQL & " Where Type= 0 and  TBLLCMargin.TblLCID = " & val(Me.TXTTblLCID.text)

loadgrid StrSQL, GrdMargin, True, False
    With GrdMargin
    
    For row = 1 To .rows - 1

        If val(.TextMatrix(row, .ColIndex("StillAmount"))) = 0 Then
            .TextMatrix(row, .ColIndex("IsFullPayed")) = 1
            .cell(flexcpBackColor, row, 1, row, .Cols - 1) = vbGreen
        Else
            .TextMatrix(row, .ColIndex("IsFullPayed")) = 0
        End If
    Next
    End With

StrSQL = "sELECT TBLLCMargin.*,Acc.Account_Serial as MarginAccountSerial,Acc.Account_Name as MarginAccount,Acc3.Account_Name AccountMargen2Name  ,"
StrSQL = StrSQL & " Acc2.Account_Serial as BankAccountSerial,Acc2.Account_Name as BankAccount , "
StrSQL = StrSQL & " Acc4.Account_Serial as BankAccountSerial2,Acc4.Account_Name as BankAccount2  "
StrSQL = StrSQL & " from TBLLCMargin Left Outer join Accounts Acc On TBLLCMargin.MarginAccountCode = Acc.Account_Code "
StrSQL = StrSQL & " Left Outer join Accounts Acc2 On TBLLCMargin.BankAccountCode = Acc2.Account_Code"
StrSQL = StrSQL & " Left Outer join Accounts Acc3 On TBLLCMargin.AccountMargen2 = Acc3.Account_Code"
StrSQL = StrSQL & " Left Outer join Accounts Acc4 On TBLLCMargin.BankAccountCode2 = Acc4.Account_Code"
StrSQL = StrSQL & " Where   TBLLCMargin.TblLCID = " & val(Me.TXTTblLCID.text)
loadgrid StrSQL, GrdMargin2, True, False

        If GrdMargin.rows - 1 > 1 Then
            Me.txtMarginTotal.text = GrdMargin.Aggregate(flexSTSum, GrdMargin.FixedRows, GrdMargin.ColIndex("Amount"), GrdMargin.rows - 1, GrdMargin.ColIndex("Amount"))
        End If
        
        If GrdMargin2.rows - 1 > 1 Then
            Me.txtMarginTotal2.text = GrdMargin2.Aggregate(flexSTSum, GrdMargin2.FixedRows, GrdMargin2.ColIndex("StillAmount"), GrdMargin2.rows - 1, GrdMargin2.ColIndex("StillAmount"))
        End If
    
    With GrdMargin2
    
    For row = 1 To .rows - 1

        If val(.TextMatrix(row, .ColIndex("StillAmount"))) = 0 Then
            .TextMatrix(row, .ColIndex("IsFullPayed")) = 1
            .cell(flexcpBackColor, row, 1, row, .Cols - 1) = vbGreen
        Else
            .TextMatrix(row, .ColIndex("IsFullPayed")) = 0
        End If
    Next
    End With

StrSQL = "sELECT TBLLCMargin2.*,Acc.Account_Serial as MarginAccountSerial,Acc.Account_Name as MarginAccount,Acc3.Account_Name AccountMargen2Name  ,"
StrSQL = StrSQL & " Acc2.Account_Serial as BankAccountSerial,Acc2.Account_Name as BankAccount  "
StrSQL = StrSQL & " from TBLLCMargin2 Left Outer join Accounts Acc On TBLLCMargin2.MarginAccountCode = Acc.Account_Code "
StrSQL = StrSQL & " Left Outer join Accounts Acc2 On TBLLCMargin2.BankAccountCode = Acc2.Account_Code"
StrSQL = StrSQL & " Left Outer join Accounts Acc3 On TBLLCMargin2.AccountMargen2 = Acc3.Account_Code"
StrSQL = StrSQL & " Where   TBLLCMargin2.TblLCID = " & val(Me.TXTTblLCID.text)
loadgrid StrSQL, GrdMargin4, True, False

        If GrdMargin4.rows - 1 > 1 Then
            Me.txtMarginTotal4.text = GrdMargin4.Aggregate(flexSTSum, GrdMargin4.FixedRows, GrdMargin4.ColIndex("Amount"), GrdMargin4.rows - 1, GrdMargin4.ColIndex("Amount"))
        End If
        

    With GrdMargin4
    
    For row = 1 To .rows - 1

        If val(.TextMatrix(row, .ColIndex("StillAmount"))) = 0 Then
            .TextMatrix(row, .ColIndex("IsFullPayed")) = 1
            .cell(flexcpBackColor, row, 1, row, .Cols - 1) = vbGreen
        Else
            .TextMatrix(row, .ColIndex("IsFullPayed")) = 0
        End If
    Next
    End With

  s = "Select * from LCTypes  Where Id = " & val(DCLC.BoundText)
    Set rsDummy = New ADODB.Recordset
    rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly
    If Not rsDummy.EOF Then
            
        If val(rsDummy!TypeLCLG & "") = 0 Then
            optTypeLCLG(0).value = True
            optTypeLCLG_Click 0
            Frame1.Visible = False
        ElseIf val(rsDummy!TypeLCLG & "") = 1 Then
            optTypeLCLG(1).value = True
            optTypeLCLG_Click 1
            Frame1.Visible = True
        End If
   End If


StrSQL = "sELECT tblLCOpenB.*,Acc.Account_Serial as MarginAccountSerial,Acc.Account_Name as MarginAccount  ,"
StrSQL = StrSQL & " Acc2.Account_Serial as BankAccountSerial,Acc2.Account_Name as BankAccount  "
StrSQL = StrSQL & " from tblLCOpenB Left Outer join Accounts Acc On tblLCOpenB.MarginAccountCode = Acc.Account_Code "
StrSQL = StrSQL & " Left Outer join Accounts Acc2 On tblLCOpenB.BankAccountCode = Acc2.Account_Code"
StrSQL = StrSQL & " Where   tblLCOpenB.TblLCID = " & val(Me.TXTTblLCID.text)
loadgrid StrSQL, GrdMargin3, True, False

'        If GrdMargin.rows - 1 > 1 Then
'            Me.txtMarginTotal.text = GrdMargin.Aggregate(flexSTSum, GrdMargin.FixedRows, GrdMargin.ColIndex("Amount"), GrdMargin.rows - 1, GrdMargin.ColIndex("Amount"))
'        End If
'
'        If GrdMargin2.rows - 1 > 1 Then
'            Me.txtMarginTotal2.text = GrdMargin2.Aggregate(flexSTSum, GrdMargin2.FixedRows, GrdMargin2.ColIndex("StillAmount"), GrdMargin2.rows - 1, GrdMargin2.ColIndex("StillAmount"))
'        End If




    '    rs("OpenBalanceDate").value = Me.Dtp.value

    txtopening_balance_voucher_id.text = IIf(IsNull(rs("opening_balance_voucher_id").value), get_opening_balance_voucher_id, rs("opening_balance_voucher_id").value)
    If val(txtopening_balance_voucher_id.text) = 0 Then
        txtopening_balance_voucher_id.text = get_opening_balance_voucher_id
    End If
    Dim FirstPeriodDateInthisYear As Date

    If (IsNull(rs("OpenBalanceDate").value)) Then
        getFirstPeriodDateInthisYear FirstPeriodDateInthisYear

        Me.Dtp.value = FirstPeriodDateInthisYear

        '     Me.Dtp.Enabled = True
    Else
        
        Me.Dtp.value = rs("OpenBalanceDate").value
        '     Me.Dtp.Enabled = False
    End If
    
    If Not IsNull(rs("OpenBalanceType").value) Then
        Me.TxtOpenBalance.text = IIf(IsNull(rs("OpenBalance")), "", Trim(rs("OpenBalance")))

        If rs("OpenBalanceType").value = 0 Then
            OptType(0).value = True
            OptType_Click 0
        ElseIf rs("OpenBalanceType").value = 1 Then
            OptType(1).value = True
            OptType_Click 1
        End If
        
    Else
        Me.TxtOpenBalance.text = 0
        Me.OptType(2).value = True
        OptType_Click 2
    End If
 
 
    If val(TxtNoteID2) <> 0 Then
        CmdCreateV.Enabled = False
        Command9.Enabled = True
        Command2.Enabled = True
'        Cmd(2).Enabled = False
'       Cmd(1).Enabled = False
     Else
        CmdCreateV.Enabled = True
        'Command9.Enabled = False
        Command2.Enabled = False
'        Cmd(2).Enabled = True
'        Cmd(1).Enabled = True
    End If
    
    Exit Sub
ErrTrap:
End Sub
 
Private Sub Text4_Change()

End Sub

Private Sub OptType_Click(index As Integer)
    Me.TxtOpenBalance.Enabled = Not OptType(2).value
    Me.TxtOpenBalance.text = IIf(OptType(2).value = True, 0, Me.TxtOpenBalance.text)

End Sub

Private Sub TxtModFlg_Change()

    If Me.TxtModFlg.text = "N" Then
        CmdRemove.Enabled = True
        Ele(1).Enabled = True
        Cmd(0).Enabled = False
        Cmd(1).Enabled = False
        Cmd(4).Enabled = False
        Cmd(5).Enabled = False

        Cmd(2).Enabled = True
        Cmd(3).Enabled = True
    DboParentAccount.Enabled = True
   
    
    ElseIf Me.TxtModFlg.text = "E" Then
        CmdRemove.Enabled = True
        Ele(1).Enabled = True
        Cmd(2).Enabled = True
        Cmd(3).Enabled = True

        Cmd(0).Enabled = False
        Cmd(1).Enabled = False
        Cmd(4).Enabled = False

        Cmd(5).Enabled = False
DboParentAccount.Enabled = False
        
    Else
        Ele(1).Enabled = False

        CmdRemove.Enabled = False
        Cmd(2).Enabled = False
        Cmd(3).Enabled = False
        Cmd(0).Enabled = True
        Cmd(1).Enabled = True
        Cmd(4).Enabled = True

        Cmd(5).Enabled = True
DboParentAccount.Enabled = False
    End If
    If val(TXTNoteID) <> 0 Then
        CmdCreateV.Enabled = False
        Command9.Enabled = True
        Command2.Enabled = True
        ' Cmd(2).Enabled = False
     Else
        CmdCreateV.Enabled = True
'        Command9.Enabled = False
        Command2.Enabled = False
    End If

End Sub

Private Sub TxtNoOfParcil_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtNoOfParcil.text, 0)

End Sub

Private Sub TxtOpenBalance_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtOpenBalance.text, 0)
End Sub

Private Sub txtOPenValue_Change()

'txtOPenValue2 = txtOPenValue
End Sub

Private Sub txtOPenValue2_Change()
If Me.TxtModFlg.text = "" Or Me.TxtModFlg.text = "R" Then Exit Sub
    'txtOPenValue = txtOPenValue2
    txtLGExpPeriod = DateDiff("D", txtGuaranteeDate.value, txtLGExpiryDate.value)
    txtLGExpPeriod_Change
End Sub

Private Sub txtTotalBondHistory_Change()
txtBondAmt = txtTotalBondHistory
End Sub

Private Sub TxtValue_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtValue.text, 0)
End Sub

Private Sub XPBtnMove_Click(index As Integer)

    If Me.TxtModFlg.text = "N" Then
        clear_all Me
        Me.TxtModFlg.text = "R"
        XPBtnMove_Click (1)
    End If

    On Error GoTo ErrTrap

    Select Case index

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
    FiLLTXT
    Retrive
    Exit Sub
ErrTrap:
End Sub
Function ReloadCombos()
LoadCombosData
End Function

Private Sub LoadCombosData()
    Dim StrSQL As String
    Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos
    Dcombos.GetUsers Me.DCboUserName
End Sub




















'--------------------------------------


Private Sub GrdMargin4_AfterEdit(ByVal row As Long, ByVal Col As Long)
 
    Dim StrAccountCode As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long
    
    Dim s As String
    s = "Select Account_code from TblCustemers where CusID = " & val(DBCboClientName.BoundText)
    Set rsDummy = New ADODB.Recordset
    rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly
    If Not rsDummy.EOF Then
        GrdMargin4.TextMatrix(row, GrdMargin4.ColIndex("MarginAccount")) = DBCboClientName.text ' Trim(rsDummy!Account_code & "")
        GrdMargin4.TextMatrix(row, GrdMargin4.ColIndex("MarginAccountCode")) = Trim(rsDummy!Account_code & "")
        
    End If
    With GrdMargin4
        If .TextMatrix(row, .ColIndex("GuaranteeDate")) = "" Then
            .TextMatrix(row, .ColIndex("GuaranteeDate")) = txtGuaranteeDate.value
        End If
        
        If CBool(.ValueMatrix(row, .ColIndex("IsOpenBalance"))) Then
             getFirstPeriodDateInthisYear2 FirstPeriodDateInthisYear
            .TextMatrix(row, .ColIndex("GuaranteeDate")) = FirstPeriodDateInthisYear
            .TextMatrix(row, .ColIndex("OrderDate")) = FirstPeriodDateInthisYear
        End If
            
       
        If .TextMatrix(row, .ColIndex("BankAccountCode")) = "" Then
            .TextMatrix(row, .ColIndex("BankAccountCode")) = txtGuaranteeDate.value
             s = "SELECT AcceptAccount_Code,accounts.Account_Name,accounts.Account_Serial, * FROM TblLC tl Inner join accounts On tl.AcceptAccount_Code = accounts.Account_Code WHERE tl.TblLCID = " & val(TXTTblLCID.text)
            Set rsDummy = New ADODB.Recordset
            rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly
            If Not rsDummy.EOF Then
                
                GrdMargin4.TextMatrix(row, GrdMargin4.ColIndex("BankAccountSerial")) = Trim(rsDummy!account_serial & "")
                GrdMargin4.TextMatrix(row, GrdMargin4.ColIndex("BankAccount")) = Trim(rsDummy!account_name & "")
                GrdMargin4.TextMatrix(row, GrdMargin4.ColIndex("BankAccountCode")) = Trim(rsDummy!AcceptAccount_Code & "")
            End If
             
'             s = "Select * from Accounts where Account_Code = '" & Trim(GrdMargin4.TextMatrix(row, GrdMargin4.ColIndex("BankAccountCode"))) & "'"
'             Set rsDummy = New ADODB.Recordset
'             rsDummy.Open s, Cn, adOpenForwardOnly, adLockReadOnly
'             If Not rsDummy.EOF Then
'                GrdMargin4.TextMatrix(row, GrdMargin4.ColIndex("BankAccountSerial")) = Trim(rsDummy!account_serial & "")
'                GrdMargin4.TextMatrix(row, GrdMargin4.ColIndex("BankAccount")) = Trim(rsDummy!account_name & "")
'             End If
'

              
            'BankAccountSerial
            'BankAccount
        End If

        Select Case .ColKey(Col)
            Case "OrderDate"
                If IsDate(.TextMatrix(row, .ColIndex("OrderDate"))) Then
                    .TextMatrix(row, .ColIndex("GuaranteeDate")) = DateAdd("D", val(txtAcceptianPeriod), .TextMatrix(row, .ColIndex("OrderDate")))
                End If
                If CBool(.ValueMatrix(row, .ColIndex("IsOpenBalance"))) Then
                    getFirstPeriodDateInthisYear2 FirstPeriodDateInthisYear
                    .TextMatrix(row, .ColIndex("GuaranteeDate")) = FirstPeriodDateInthisYear
                End If

                If CBool(.ValueMatrix(row, .ColIndex("IsOpenBalance"))) Then
                    getFirstPeriodDateInthisYear2 FirstPeriodDateInthisYear
                    .TextMatrix(row, .ColIndex("OrderDate")) = FirstPeriodDateInthisYear
                End If
            Case "MarginNo"
                If CBool(.ValueMatrix(row, .ColIndex("IsOpenBalance"))) Then Exit Sub
                           StrSQL = "sELECT TBLLCMargin2.*,BankAccountCode,CC.Account_Name BankAccount2,CC.Account_Serial BankAccountSerial2, "
                StrSQL = StrSQL & "   accounts.Account_Name MarginAccount ,accounts.Account_Serial  MarginAccountSerial, MarginAccountCode"
                StrSQL = StrSQL & "   FROM TBLLCMargin2 Inner join accounts On accounts.Account_Code = TBLLCMargin2.MarginAccountCode"
                StrSQL = StrSQL & "   Inner join accounts CC On CC.Account_Code = TBLLCMargin2.BankAccountCode"

                'StrSQL = StrSQL & "   Where MarginNo = " & val(.TextMatrix(row, .ColIndex("MarginNo"))) & " and IsNull(TBLLCMargin2.Amount,0) - IsNull(TBLLCMargin2.PayedAmount,0) > 0  and TblLCID=" & val(TXTTblLCID.text)
                StrSQL = StrSQL & "   Where MarginNo = " & val(.TextMatrix(row, .ColIndex("MarginNo"))) & " and IsNull(TBLLCMargin2.IsFullPayed,0) = 1  and TBLLCMargin2.TblLCID=" & val(TXTTblLCID.text)
                StrSQL = StrSQL & "   Order by TBLLCMargin2.ID desc"
                Set rsDummy = New ADODB.Recordset
                rsDummy.Open StrSQL, Cn, adOpenStatic, adLockReadOnly
                If Not rsDummy.EOF Then
                    MsgBox "«·›« Ê—… —ﬁ„ " & val(.TextMatrix(row, .ColIndex("MarginNo"))) & " ·Â–« «·«⁄ „«œ  „ ”œ«œÂ« »«·ﬂ«„·"
                     .TextMatrix(row, .ColIndex("MarginNo")) = ""
                     Exit Sub
                End If
                rsDummy.Close
            
                StrSQL = "sELECT TBLLCMargin2.*,BankAccountCode,CC.Account_Name BankAccount2,CC.Account_Serial BankAccountSerial2, "
                StrSQL = StrSQL & "   accounts.Account_Name MarginAccount ,accounts.Account_Serial  MarginAccountSerial, MarginAccountCode"
                StrSQL = StrSQL & "   FROM TBLLCMargin2 Inner join accounts On accounts.Account_Code = TBLLCMargin2.MarginAccountCode"
                StrSQL = StrSQL & "   Inner join accounts CC On CC.Account_Code = TBLLCMargin2.BankAccountCode"

                'StrSQL = StrSQL & "   Where MarginNo = " & val(.TextMatrix(row, .ColIndex("MarginNo"))) & " and IsNull(TBLLCMargin2.Amount,0) - IsNull(TBLLCMargin2.PayedAmount,0) > 0  and TblLCID=" & val(TXTTblLCID.text)
                StrSQL = StrSQL & "   Where MarginNo = " & val(.TextMatrix(row, .ColIndex("MarginNo"))) & " and IsNull(TBLLCMargin2.IsFullPayed,0) = 0  and TBLLCMargin2.TblLCID=" & val(TXTTblLCID.text)
                StrSQL = StrSQL & "   Order by TBLLCMargin2.ID desc"
                Set rsDummy = New ADODB.Recordset
                rsDummy.Open StrSQL, Cn, adOpenStatic, adLockReadOnly
                If rsDummy.EOF Then
                 
                    .TextMatrix(row, .ColIndex("MarginAccountCode")) = get_bank_Account(Dcbank.BoundText, "Account_Code")
                    .TextMatrix(row, .ColIndex("MarginAccount")) = Dcbank.text
                 
                 '   MsgBox "Â–« «·ﬁ—÷ €Ì— „”Ã· „‰ ﬁ»·"
                Else

                    .TextMatrix(row, .ColIndex("BankAccount")) = rsDummy!BankAccount2 & ""
                    .TextMatrix(row, .ColIndex("BankAccountSerial")) = rsDummy!BankAccountSerial2 & ""
                    .TextMatrix(row, .ColIndex("BankAccountCode")) = rsDummy!BankAccountCode & ""

                    .TextMatrix(row, .ColIndex("MarginAccountCode")) = rsDummy!MarginAccountCode & ""
                    .TextMatrix(row, .ColIndex("MarginAccountSerial")) = rsDummy!MarginAccountSerial & ""
                    .TextMatrix(row, .ColIndex("MarginAccount")) = rsDummy!MarginAccount & ""
                    .TextMatrix(row, .ColIndex("OrderDate")) = rsDummy!OrderDate & ""
                    .TextMatrix(row, .ColIndex("GuaranteeDate")) = rsDummy!GuaranteeDate & ""
                    .TextMatrix(row, .ColIndex("Amount")) = rsDummy!StillAmount & "" 'rsDummy!Amount & ""
                    .TextMatrix(row, .ColIndex("PayedAmount")) = 0 'val(rsDummy!Amount & "") - val(rsDummy!StillAmount & "")
                    .TextMatrix(row, .ColIndex("StillAmount")) = rsDummy!StillAmount & ""
              '      .TextMatrix(row, .ColIndex("NoteID")) = rsDummy!NoteID & ""
              '      .TextMatrix(row, .ColIndex("NoteSerial")) = rsDummy!NoteSerial & ""
                    
                    .TextMatrix(row, .ColIndex("IsFullPayed")) = val(rsDummy!IsFullPayed & "")
                    
                    '.TextMatrix(Row, .ColIndex("StillAmount")) = rsDummy!StillAmount & ""
                End If
'

          Case "AccountMargen2Serial"
                s = "SELECT accounts.Account_Name,accounts.Account_Code,Account_Serial  FROM accounts  WHERE Account_Serial = N'" & Trim(GrdMargin4.TextMatrix(row, Col)) & "'"
                Set rsDummy = New ADODB.Recordset
                rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly
                If Not rsDummy.EOF Then


                    GrdMargin4.TextMatrix(row, GrdMargin4.ColIndex("AccountMargen2Name")) = Trim(rsDummy!account_name & "")
                    GrdMargin4.TextMatrix(row, GrdMargin4.ColIndex("AccountMargen2")) = Trim(rsDummy!Account_code & "")
                End If
           Case "MarginAccountSerial"
                s = "SELECT accounts.Account_Name,accounts.Account_Code,Account_Serial  FROM accounts  WHERE Account_Serial = N'" & Trim(GrdMargin4.TextMatrix(row, Col)) & "'"
                Set rsDummy = New ADODB.Recordset
                rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly
                If Not rsDummy.EOF Then


                    GrdMargin4.TextMatrix(row, GrdMargin4.ColIndex("MarginAccount")) = Trim(rsDummy!account_name & "")
                    GrdMargin4.TextMatrix(row, GrdMargin4.ColIndex("MarginAccountCode")) = Trim(rsDummy!Account_code & "")
                End If
                
'                StrAccountCode = .ComboData
'                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("MarginAccountSerial"), False, True)
'                .TextMatrix(row, .ColIndex("MarginAccount")) = StrAccountCode
                
           Case "MarginAccount"
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("MarginAccountCode"), False, True)
                .TextMatrix(row, .ColIndex("MarginAccountCode")) = StrAccountCode
                
                  s = "SELECT accounts.Account_Name,accounts.Account_Code,Account_Serial  FROM accounts  WHERE Account_Code = N'" & Trim(StrAccountCode) & "'"
                Set rsDummy = New ADODB.Recordset
                rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly
                If Not rsDummy.EOF Then


                    'GrdMargin4.TextMatrix(row, GrdMargin4.ColIndex("MarginAccount")) = Trim(rsDummy!account_name & "")
                    GrdMargin4.TextMatrix(row, GrdMargin4.ColIndex("MarginAccountSerial")) = Trim(rsDummy!account_serial & "")
                End If
                
           Case "AccountMargen2Name"
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("AccountMargen2"), False, True)
                .TextMatrix(row, .ColIndex("AccountMargen2")) = StrAccountCode
                
                
                    s = "SELECT accounts.Account_Name,accounts.Account_Code,Account_Serial  FROM accounts  WHERE Account_Code = N'" & Trim(StrAccountCode) & "'"
                Set rsDummy = New ADODB.Recordset
                rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly
                If Not rsDummy.EOF Then


                    'GrdMargin4.TextMatrix(row, GrdMargin4.ColIndex("MarginAccount")) = Trim(rsDummy!account_name & "")
                    GrdMargin4.TextMatrix(row, GrdMargin4.ColIndex("AccountMargen2Serial")) = Trim(rsDummy!account_serial & "")
                End If
 
           Case "BankAccount"
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("BankAccountCode"), False, True)
                .TextMatrix(row, .ColIndex("BankAccountCode")) = StrAccountCode
                
                   s = "SELECT accounts.Account_Name,accounts.Account_Code,Account_Serial  FROM accounts  WHERE Account_Code = N'" & Trim(StrAccountCode) & "'"
                Set rsDummy = New ADODB.Recordset
                rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly
                If Not rsDummy.EOF Then


                    'GrdMargin4.TextMatrix(row, GrdMargin4.ColIndex("MarginAccount")) = Trim(rsDummy!account_name & "")
                    GrdMargin4.TextMatrix(row, GrdMargin4.ColIndex("BankAccountSerial")) = Trim(rsDummy!account_serial & "")
                End If
            Case "BankAccountSerial"
                s = "SELECT accounts.Account_Name,accounts.Account_Code,Account_Serial  FROM accounts  WHERE Account_Serial = N'" & Trim(GrdMargin4.TextMatrix(row, Col)) & "'"
                Set rsDummy = New ADODB.Recordset
                rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly
                If Not rsDummy.EOF Then


                    GrdMargin4.TextMatrix(row, GrdMargin4.ColIndex("BankAccount")) = Trim(rsDummy!account_name & "")
                    GrdMargin4.TextMatrix(row, GrdMargin4.ColIndex("BankAccountCode")) = Trim(rsDummy!Account_code & "")
                End If
                
                
                
            Case "BankAccount2"
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("BankAccountCode"), False, True)
                .TextMatrix(row, .ColIndex("BankAccountCode2")) = StrAccountCode
                
                   s = "SELECT accounts.Account_Name,accounts.Account_Code,Account_Serial  FROM accounts  WHERE Account_Code = N'" & Trim(StrAccountCode) & "'"
                Set rsDummy = New ADODB.Recordset
                rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly
                If Not rsDummy.EOF Then


                    'GrdMargin4.TextMatrix(row, GrdMargin4.ColIndex("MarginAccount")) = Trim(rsDummy!account_name & "")
                    GrdMargin4.TextMatrix(row, GrdMargin4.ColIndex("BankAccountSerial2")) = Trim(rsDummy!account_serial & "")
                    'GrdMargin4.TextMatrix(row, GrdMargin4.ColIndex("BankAccountCode2")) = Trim(rsDummy!Account_code & "")
                End If
            Case "BankAccountSerial2"
                s = "SELECT accounts.Account_Name,accounts.Account_Code,Account_Serial  FROM accounts  WHERE Account_Serial = N'" & Trim(GrdMargin4.TextMatrix(row, Col)) & "'"
                Set rsDummy = New ADODB.Recordset
                rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly
                If Not rsDummy.EOF Then


                    GrdMargin4.TextMatrix(row, GrdMargin4.ColIndex("BankAccount2")) = Trim(rsDummy!account_name & "")
                    GrdMargin4.TextMatrix(row, GrdMargin4.ColIndex("BankAccountCode2")) = Trim(rsDummy!Account_code & "")
                End If
                
'                StrAccountCode = .ComboData
'                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("MarginAccountSerial"), False, True)
            Case "AccountName"
                '  .TextMatrix(Row, .ColIndex("userid")) = user_id
                        
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("AccountCode"), False, True)
                .TextMatrix(row, .ColIndex("AccountCode")) = StrAccountCode
                .TextMatrix(row, .ColIndex("ExpensesID")) = get_Expenses_id(StrAccountCode)
                .TextMatrix(row, .ColIndex("LineNo1")) = setfoxy_Line

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrSQL = "select * from Expenses_accounts where Account_Code='" & StrAccountCode & "'"
                Else
                    StrSQL = "select * from Expenses_accounts_eng where Account_Code='" & StrAccountCode & "'"
                End If
            
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                     
'                If rs.RecordCount > 0 Then
'                    .TextMatrix(Row, .ColIndex("des")) = IIf(IsNull(rs("parent_account").value), "", rs("parent_account").value)
'                Else
'                    .TextMatrix(Row, .ColIndex("des")) = ""
'                End If



            Case "Amount"
                
                .TextMatrix(row, .ColIndex("StillAmount")) = val(.TextMatrix(row, .ColIndex("Amount"))) - val(.TextMatrix(row, .ColIndex("PayedAmount")))
                If val(.TextMatrix(row, .ColIndex("StillAmount"))) = 0 Then
                    .TextMatrix(row, .ColIndex("IsFullPayed")) = 1
                Else
                    .TextMatrix(row, .ColIndex("IsFullPayed")) = 0
                End If
            Case "PayedAmount"
                If .TextMatrix(row, .ColIndex("PayDate")) = "" Then
                    .TextMatrix(row, .ColIndex("PayDate")) = Date
                End If
                .TextMatrix(row, .ColIndex("StillAmount")) = val(.TextMatrix(row, .ColIndex("Amount"))) - val(.TextMatrix(row, .ColIndex("PayedAmount")))
        
                If val(.TextMatrix(row, .ColIndex("StillAmount"))) = 0 Then
                    .TextMatrix(row, .ColIndex("IsFullPayed")) = 1
                Else
                    .TextMatrix(row, .ColIndex("IsFullPayed")) = 0
                End If
            
            Case "Amount", "Price", "ChSameCurrncy"
                Dim sgl As String
           
               
                
                
                
                
                '    sgl = "update  marakes_taklefa_temp  set value=0 where  line_no=" & Val(GrdMargin4.TextMatrix(GrdMargin4.Row, GrdMargin4.ColIndex("LineNo1")))
                '     Cn.Execute sgl, , adExecuteNoRecords
        
                '  Me.XPTxtVal.text = Format(Me.XPTxtVal.text, SystemOptions.SysDefCurrencyForamt)
        End Select

   

        ' Me.XPTxtVal.text = Format(Me.XPTxtVal.text, SystemOptions.SysDefCurrencyForamt)
        'to Add new row if needed
        If row = .rows - 1 Then
            .rows = .rows + 1
        End If
          
        ' ReLineGrid
    End With
    Me.txtMarginTotal2.text = GrdMargin4.Aggregate(flexSTSum, GrdMargin4.FixedRows, GrdMargin4.ColIndex("StillAmount"), GrdMargin4.rows - 1, GrdMargin4.ColIndex("StillAmount"))
    'ReLineGrid


End Sub

Private Sub GrdMargin4_BeforeEdit(ByVal row As Long, ByVal Col As Long, Cancel As Boolean)
  With GrdMargin4

        If row > .FixedRows Then
             If val(.TextMatrix(row, .ColIndex("NoteSerial"))) <> 0 And .ColKey(Col) <> "NoteSerial" And val(.TextMatrix(row, .ColIndex("NoteSerial2"))) <> 0 And .ColKey(Col) <> "NoteSerial2" And .ColKey(Col) <> "NoteSerial3" Then
                '  Cancel = True
                '  Exit Sub
              End If
        End If

        Select Case .ColKey(Col)
            Case "IsOpenBalance"
                If val(.TextMatrix(row, .ColIndex("NoteSerial"))) <> 0 Or val(.TextMatrix(row, .ColIndex("NoteSerial2"))) <> 0 Then
                    Cancel = True
                End If
            Case "Vat", "StillAmount"
                 Cancel = True
            Case "Amount", "PayedAmount", "MargenAmount", "MargenValue"
                .ComboList = ""
        Case "Vatyo"
             
            Case "value"
               Cancel = True
              Case "Amount", "MarginNo", "GuaranteeDate", "Serial"
                .ComboList = ""
              Case "ChSameCurrncy"
                .ComboList = ""

            Case "des"
                .ComboList = ""
        
            Case "Order_No"
                .ComboList = ""
        
                '  Cancel = True
            
        End Select

    End With
End Sub

Private Sub GrdMargin4_CellButtonClick(ByVal row As Long, ByVal Col As Long)

Dim Frm As New FrmDateOpProject, mIDD As Long, mUserID As Long, mDate As Date, mTime As Date



With GrdMargin4
Select Case .ColKey(Col)

        Case "GuaranteeDate"
            
            Frm.index = 614
            Me.LngRow = row
            Frm.show 1
        Case "OrderDate"
            
            Frm.index = 615
            Me.LngRow = row
            Frm.show 1
            GrdMargin4_AfterEdit row, Col
       Case "PayDate"
            
            Frm.index = 616
            Me.LngRow = row
            Frm.show 1
       Case "NoteSerial"
'                                  LngRow = Row
        If CBool(.ValueMatrix(row, .ColIndex("IsOpenBalance"))) Then
            ShowGL_ccOpening val(.TextMatrix(row, .ColIndex("NoteSerial"))), , 101, val(.TextMatrix(row, .ColIndex("NoteID")))
        Else
        
            ShowGL_cc val(.TextMatrix(row, .ColIndex("NoteSerial"))), , 22008
        End If
     Case "NoteSerial2"
'                                  LngRow = Row
        
        ShowGL_cc val(.TextMatrix(row, .ColIndex("NoteSerial2"))), , 22009
     Case "NoteSerial3"
'                                  LngRow = Row
        ShowGL_cc val(.TextMatrix(row, .ColIndex("NoteSerial3"))), , 22007
        
    Case "CreateNote2"
        If val(.TextMatrix(row, .ColIndex("PayedAmount"))) <> 0 Then
            CreateEntry row, 6, 1
        End If
    Case "CreateNote"
        If val(.TextMatrix(row, .ColIndex("Amount"))) <> 0 Then
            CreateEntry row, 6, 0
        End If
    Case "CreateNote3"
        If val(.TextMatrix(row, .ColIndex("NoteId3"))) = 0 And val(.TextMatrix(row, .ColIndex("MargenValue"))) <> 0 Then
            CreateEntry row, 5, 0
        End If
    Case "DeleteEntry"
        

'        s = "Delete Notes where NoteId = " & val(.TextMatrix(row, .ColIndex("NoteId3")))
'        Cn.Execute s
'        s = "Delete Notes where NoteId = " & val(.TextMatrix(row, .ColIndex("NoteId2")))
'        Cn.Execute s
'        If CBool(.ValueMatrix(row, .ColIndex("IsOpenBalance"))) Then
'            s = "Delete Notes1 where NoteId = " & val(.TextMatrix(row, .ColIndex("NoteId")))
'            Cn.Execute s
'            s = "Delete DOUBLE_ENTREY_VOUCHERS1 where Notes_ID = " & val(.TextMatrix(row, .ColIndex("NoteId")))
'
'            Cn.Execute s
'        Else
'            s = "Delete Notes where NoteId = " & val(.TextMatrix(row, .ColIndex("NoteId")))
'            Cn.Execute s
 '       End If
        .TextMatrix(row, .ColIndex("NoteId")) = ""
        .TextMatrix(row, .ColIndex("NoteId2")) = ""
        .TextMatrix(row, .ColIndex("NoteId3")) = ""
        .TextMatrix(row, .ColIndex("NoteSerial")) = ""
        .TextMatrix(row, .ColIndex("NoteSerial2")) = ""
        '.TextMatrix(row, .ColIndex("NoteSeria3")) = ""
        
        MsgBox " „ Õ–› «·ﬁÌÊœ"
      
        
 End Select
End With
End Sub

Private Sub GrdMargin4_KeyUp(KeyCode As Integer, Shift As Integer)
 With GrdMargin4

        Select Case .ColKey(.Col)

            Case "Order_No"
                           
                If KeyCode = vbKeyF3 Then
                  
                    Order_no_search.show
                     Order_no_search.RetrunType = 4
                End If

            Case "MarginAccount"
                  If KeyCode = vbKeyF3 Then
                        Account_search.show
                        Account_search.case_id = 350057
                    End If
 
            Case "AccountMargen2"
                  If KeyCode = vbKeyF3 Then
                        Account_search.show
                        Account_search.case_id = 350059
                    End If
            Case "BankAccount"
                  If KeyCode = vbKeyF3 Then
                        Account_search.show
                        Account_search.case_id = 350058
                    End If
            Case "BankAccount2"
                  If KeyCode = vbKeyF3 Then
                        Account_search.show
                        Account_search.case_id = 350061
                    End If
        End Select

    End With

End Sub

Private Sub GrdMargin4_StartEdit(ByVal row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim rs As New ADODB.Recordset
    Dim Rs1 As New ADODB.Recordset

    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim StrComboList1 As String

    Dim Msg As String

    'Case "DebitName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a1%' Or ACCOUNTS.Account_Code Like 'a3%')"
    'Case "CreditName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a2%' Or ACCOUNTS.Account_Code Like 'a4%')"
    With GrdMargin4

        Select Case .ColKey(Col)
        
 
          Case "MarginAccountSerial"
                If SystemOptions.UserInterface = EnglishInterface Then
                
                    StrSQL = "SELECT Account_Code, Account_Serial  from ACCOUNTS    "
           
                        StrSQL = StrSQL + " where (last_account=1)"
                    
                    StrSQL = StrSQL & GetAccountByBarnchUser
                    StrSQL = StrSQL & GetAccountCodeHiding
                Else
                
                    StrSQL = "SELECT   Account_Code  as Account_Code, Account_Serial from ACCOUNTS   "
                     StrSQL = StrSQL + " where    (last_account=1)"

                   StrSQL = StrSQL & GetAccountByBarnchUser
                  StrSQL = StrSQL & GetAccountCodeHiding
                
                End If
                Set rs = New ADODB.Recordset
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                If SystemOptions.UserInterface = ArabicInterface Then
                StrComboList = .BuildComboList(rs, "Account_Serial", "Account_Code")
                Else
                StrComboList = .BuildComboList(rs, "Account_Serial", "Account_Code")
                End If
                
                
 
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
        
        
          Case "MarginAccount"
                If SystemOptions.UserInterface = EnglishInterface Then
                
                    StrSQL = "SELECT Account_Code, Account_NameEng  from ACCOUNTS    "
           
                        StrSQL = StrSQL + " where (last_account=1)"
                    
                    StrSQL = StrSQL & GetAccountByBarnchUser
                    StrSQL = StrSQL & GetAccountCodeHiding
                Else
                
                    StrSQL = "SELECT   Account_Code  as Account_Code, Account_Name as Account_Name from ACCOUNTS   "
                     StrSQL = StrSQL + " where    (last_account=1)"

                   StrSQL = StrSQL & GetAccountByBarnchUser
                  StrSQL = StrSQL & GetAccountCodeHiding
                
                End If
                Set rs = New ADODB.Recordset
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                If SystemOptions.UserInterface = ArabicInterface Then
                StrComboList = .BuildComboList(rs, "Account_Name", "Account_Code")
                Else
                StrComboList = .BuildComboList(rs, "Account_NameEng", "Account_Code")
                End If
                
                
 
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
                
                
   Case "AccountMargen2Name"
                If SystemOptions.UserInterface = EnglishInterface Then
                
                    StrSQL = "SELECT Account_Code, Account_NameEng  from ACCOUNTS    "
           
                        StrSQL = StrSQL + " where (last_account=1)"
                    
                    StrSQL = StrSQL & GetAccountByBarnchUser
                    StrSQL = StrSQL & GetAccountCodeHiding
                Else
                
                    StrSQL = "SELECT   Account_Code  as Account_Code, Account_Name as Account_Name from ACCOUNTS   "
                     StrSQL = StrSQL + " where    (last_account=1)"

                   StrSQL = StrSQL & GetAccountByBarnchUser
                  StrSQL = StrSQL & GetAccountCodeHiding
                
                End If
                Set rs = New ADODB.Recordset
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                If SystemOptions.UserInterface = ArabicInterface Then
                StrComboList = .BuildComboList(rs, "Account_Name", "Account_Code")
                Else
                StrComboList = .BuildComboList(rs, "Account_NameEng", "Account_Code")
                End If
                
                
 
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
                
 Case "BankAccount"
                If SystemOptions.UserInterface = EnglishInterface Then
                
                    StrSQL = "SELECT Account_Code, Account_NameEng  from ACCOUNTS    "
           
                        StrSQL = StrSQL + " where (last_account=1)"
                    
                    StrSQL = StrSQL & GetAccountByBarnchUser
                    StrSQL = StrSQL & GetAccountCodeHiding
                Else
                
                    StrSQL = "SELECT   Account_Code  as Account_Code, Account_Name as Account_Name from ACCOUNTS   "
                     StrSQL = StrSQL + " where    (last_account=1)"

                   StrSQL = StrSQL & GetAccountByBarnchUser
                  StrSQL = StrSQL & GetAccountCodeHiding
                
                End If
                Set rs = New ADODB.Recordset
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                If SystemOptions.UserInterface = ArabicInterface Then
                StrComboList = .BuildComboList(rs, "Account_Name", "Account_Code")
                Else
                StrComboList = .BuildComboList(rs, "Account_NameEng", "Account_Code")
                End If
                
                
 
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
                
Case "BankAccount2"
                If SystemOptions.UserInterface = EnglishInterface Then
                
                    StrSQL = "SELECT Account_Code, Account_NameEng  from ACCOUNTS    "
           
                        StrSQL = StrSQL + " where (last_account=1)"
                    
                    StrSQL = StrSQL & GetAccountByBarnchUser
                    StrSQL = StrSQL & GetAccountCodeHiding
                Else
                
                    StrSQL = "SELECT   Account_Code  as Account_Code, Account_Name as Account_Name from ACCOUNTS   "
                     StrSQL = StrSQL + " where    (last_account=1)"

                   StrSQL = StrSQL & GetAccountByBarnchUser
                  StrSQL = StrSQL & GetAccountCodeHiding
                
                End If
                Set rs = New ADODB.Recordset
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                If SystemOptions.UserInterface = ArabicInterface Then
                StrComboList = .BuildComboList(rs, "Account_Name", "Account_Code")
                Else
                StrComboList = .BuildComboList(rs, "Account_NameEng", "Account_Code")
                End If
                
                
 
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
                                
            Case "AccountName"

                '      StrSQL = "select * from Expenses_accounts"
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrSQL = "select * from Expenses_accounts"
                Else
                    StrSQL = "select * from Expenses_accounts_eng "
                End If
                 
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                'StrComboList = GrdMargin4.BuildComboList(rs, "Account_Name", "Account_Code")
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = GrdMargin4.BuildComboList(rs, "Account_Name", "Account_Code")
                Else
                    StrComboList = GrdMargin4.BuildComboList(rs, "Account_NameEng", "Account_Code")
                End If
            
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList

            Case "opr_fullcode"
                StrSQL = "  select fullcode,name from terms_operations "
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                StrComboList1 = GrdMargin4.BuildComboList(rs, "fullcode", "fullcode")

                If StrComboList1 <> "" Then
                    StrComboList1 = "|" & StrComboList1
                End If

                .ComboList = StrComboList1
         
        End Select

    End With


End Sub



