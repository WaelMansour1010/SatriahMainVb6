VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmPay_Garanty_Shipment 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6585
   ClientLeft      =   -15
   ClientTop       =   330
   ClientWidth     =   10065
   Icon            =   "FrmPay_Garanty_Shipment.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6585
   ScaleWidth      =   10065
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin C1SizerLibCtl.C1Tab C1Tab1 
      Height          =   6930
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   10005
      _cx             =   17648
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
      Caption         =   "0|1|2|3|4|5|6|7"
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
      Begin C1SizerLibCtl.C1Elastic C1Elastic100 
         Height          =   6510
         Left            =   -12360
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   45
         Width           =   9915
         _cx             =   17489
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
         Begin VB.Frame Frm2 
            BackColor       =   &H00E2E9E9&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   1365
            Left            =   720
            RightToLeft     =   -1  'True
            TabIndex        =   6
            Top             =   4080
            Width           =   8565
            Begin VB.TextBox TxtUnitID 
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
               Height          =   345
               Left            =   5355
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   10
               Top             =   225
               Width           =   2025
            End
            Begin VB.TextBox TxtVacNamee 
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
               Height          =   345
               Left            =   120
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   9
               Top             =   960
               Width           =   7260
            End
            Begin VB.TextBox TxtVacName 
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
               Height          =   345
               Left            =   120
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   8
               Top             =   585
               Width           =   7260
            End
            Begin VB.ComboBox CmbType 
               BackColor       =   &H80000018&
               Height          =   315
               ItemData        =   "FrmPay_Garanty_Shipment.frx":000C
               Left            =   2280
               List            =   "FrmPay_Garanty_Shipment.frx":001C
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   7
               Top             =   3150
               Visible         =   0   'False
               Width           =   1005
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·«”„ «‰Ã·Ì“Ì"
               Height          =   375
               Index           =   1
               Left            =   7440
               RightToLeft     =   -1  'True
               TabIndex        =   13
               Top             =   1080
               Width           =   1170
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ﬂÊœ"
               Height          =   285
               Index           =   3
               Left            =   7725
               RightToLeft     =   -1  'True
               TabIndex        =   12
               Top             =   210
               Width           =   570
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·«”„ ⁄—»Ì"
               Height          =   255
               Index           =   0
               Left            =   7470
               RightToLeft     =   -1  'True
               TabIndex        =   11
               Top             =   690
               Width           =   1170
            End
         End
         Begin C1SizerLibCtl.C1Elastic EltCont 
            Height          =   990
            Left            =   1620
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   5490
            Width           =   6750
            _cx             =   11906
            _cy             =   1746
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
               Height          =   420
               Left            =   5805
               TabIndex        =   15
               Top             =   495
               Width           =   720
               _ExtentX        =   1270
               _ExtentY        =   741
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":0035
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnSave 
               Height          =   420
               Left            =   4290
               TabIndex        =   16
               Top             =   495
               Width           =   720
               _ExtentX        =   1270
               _ExtentY        =   741
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":03CF
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnModify 
               Height          =   420
               Left            =   4935
               TabIndex        =   17
               Top             =   495
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   741
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":0769
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton BtnUndo 
               Height          =   420
               Left            =   3405
               TabIndex        =   18
               Top             =   495
               Width           =   840
               _ExtentX        =   1482
               _ExtentY        =   741
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":0B03
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnDelete 
               Height          =   420
               Left            =   990
               TabIndex        =   19
               Top             =   495
               Width           =   690
               _ExtentX        =   1217
               _ExtentY        =   741
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":0E9D
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnQuery 
               Height          =   330
               Left            =   2520
               TabIndex        =   20
               TabStop         =   0   'False
               ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
               Top             =   570
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":1437
               ColorButton     =   14737632
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton BtnUpdate 
               Height          =   330
               Left            =   9525
               TabIndex        =   21
               TabStop         =   0   'False
               ToolTipText     =   " ÕœÌÀ ﬁ«⁄œ… «·»Ì«‰« "
               Top             =   105
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":17D1
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton BtnPrint 
               Height          =   285
               Left            =   8145
               TabIndex        =   22
               TabStop         =   0   'False
               Top             =   60
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":1B6B
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnCancel 
               Height          =   420
               Left            =   75
               TabIndex        =   23
               Top             =   495
               Width           =   870
               _ExtentX        =   1535
               _ExtentY        =   741
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":1F05
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton ButPrient 
               Height          =   495
               Left            =   1680
               TabIndex        =   24
               Top             =   480
               Visible         =   0   'False
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   873
               ButtonStyle     =   1
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
               BackStyle       =   0
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":229F
               ColorButton     =   14871017
               DisplayPersistentHover=   0   'False
               DrawFocusRectangle=   0   'False
            End
            Begin VB.Label LabCountRec 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Height          =   210
               Left            =   270
               RightToLeft     =   -1  'True
               TabIndex        =   28
               Top             =   135
               Width           =   540
            End
            Begin VB.Label LabCurrRec 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Height          =   210
               Left            =   2745
               RightToLeft     =   -1  'True
               TabIndex        =   27
               Top             =   165
               Width           =   675
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄œœ «·”Ã·« :"
               Height          =   210
               Index           =   1
               Left            =   960
               RightToLeft     =   -1  'True
               TabIndex        =   26
               Top             =   135
               Width           =   975
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·”Ã· «·Õ«·Ì:"
               Height          =   210
               Index           =   0
               Left            =   3555
               RightToLeft     =   -1  'True
               TabIndex        =   25
               Top             =   135
               Width           =   975
            End
         End
         Begin VSFlex8Ctl.VSFlexGrid Grid 
            Height          =   3285
            Left            =   0
            TabIndex        =   29
            Top             =   675
            Width           =   9960
            _cx             =   17568
            _cy             =   5794
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
            Cols            =   6
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmPay_Garanty_Shipment.frx":2639
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
         Begin C1SizerLibCtl.C1Elastic EleHeader 
            Height          =   915
            Left            =   15
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   -120
            Width           =   9945
            _cx             =   17542
            _cy             =   1614
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial (Arabic)"
               Size            =   21.75
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
            BackColor       =   16777215
            ForeColor       =   4210688
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
            CaptionPos      =   7
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
               Appearance      =   0  'Flat
               BackColor       =   &H0000FF00&
               Enabled         =   0   'False
               Height          =   285
               Left            =   2460
               RightToLeft     =   -1  'True
               TabIndex        =   31
               Top             =   240
               Visible         =   0   'False
               Width           =   945
            End
            Begin ImpulseButton.ISButton btnLast 
               Height          =   345
               Left            =   150
               TabIndex        =   32
               Top             =   270
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   609
               ButtonStyle     =   1
               ButtonPositionImage=   4
               Caption         =   ""
               BackColor       =   16777215
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":271C
               ColorButton     =   16777215
               DrawFocusRectangle=   0   'False
               DisabledImageExtraction=   0
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnNext 
               Height          =   345
               Left            =   615
               TabIndex        =   33
               Top             =   270
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   609
               ButtonStyle     =   1
               ButtonPositionImage=   4
               Caption         =   ""
               BackColor       =   16777215
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":2AB6
               ColorButton     =   16777215
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnPrevious 
               Height          =   345
               Left            =   1065
               TabIndex        =   34
               Top             =   270
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   609
               ButtonStyle     =   1
               ButtonPositionImage=   4
               Caption         =   ""
               BackColor       =   16777215
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":2E50
               ColorButton     =   16777215
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnFirst 
               Height          =   345
               Left            =   1530
               TabIndex        =   35
               Top             =   270
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   609
               ButtonStyle     =   1
               ButtonPositionImage=   4
               Caption         =   ""
               BackColor       =   16777215
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":31EA
               ColorButton     =   16777215
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "ÿ—ﬁ «·œ›⁄"
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
               Index           =   12
               Left            =   5880
               RightToLeft     =   -1  'True
               TabIndex        =   36
               Top             =   240
               Width           =   3960
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   6510
         Left            =   -12060
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   45
         Width           =   9915
         _cx             =   17489
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
         Begin VB.Frame FraHeader 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   780
            Left            =   -45
            RightToLeft     =   -1  'True
            TabIndex        =   46
            Top             =   -240
            Width           =   9945
            Begin VB.TextBox TxtVac_ID 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H000000FF&
               Height          =   240
               Left            =   4110
               RightToLeft     =   -1  'True
               TabIndex        =   51
               Top             =   750
               Visible         =   0   'False
               Width           =   945
            End
            Begin VB.TextBox TxtModFlg1 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H0000FF00&
               Enabled         =   0   'False
               Height          =   285
               Left            =   5820
               RightToLeft     =   -1  'True
               TabIndex        =   50
               Text            =   "modflag"
               Top             =   690
               Visible         =   0   'False
               Width           =   465
            End
            Begin VB.Frame Frmo2 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   375
               Left            =   900
               RightToLeft     =   -1  'True
               TabIndex        =   47
               Top             =   690
               Visible         =   0   'False
               Width           =   3105
               Begin MSDataListLib.DataCombo DCUser 
                  CausesValidation=   0   'False
                  Height          =   315
                  Left            =   -255
                  TabIndex        =   48
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
                  Left            =   2280
                  RightToLeft     =   -1  'True
                  TabIndex        =   49
                  Top             =   -75
                  Width           =   855
               End
            End
            Begin MSComctlLib.ImageList GrdImageList 
               Left            =   5160
               Top             =   720
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
                     Picture         =   "FrmPay_Garanty_Shipment.frx":3584
                     Key             =   "CompanyName"
                  EndProperty
                  BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmPay_Garanty_Shipment.frx":391E
                     Key             =   "Ser"
                  EndProperty
                  BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmPay_Garanty_Shipment.frx":3CB8
                     Key             =   "Vac_Name"
                  EndProperty
                  BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmPay_Garanty_Shipment.frx":4052
                     Key             =   "ShareCount"
                  EndProperty
                  BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmPay_Garanty_Shipment.frx":43EC
                     Key             =   "Dis_Count"
                  EndProperty
                  BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmPay_Garanty_Shipment.frx":4786
                     Key             =   "Bouns"
                  EndProperty
                  BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmPay_Garanty_Shipment.frx":4B20
                     Key             =   "SharesValue"
                  EndProperty
                  BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmPay_Garanty_Shipment.frx":50BA
                     Key             =   "BuyValue"
                  EndProperty
               EndProperty
            End
            Begin ImpulseButton.ISButton btnLast1 
               Height          =   315
               Left            =   90
               TabIndex        =   52
               Top             =   270
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   556
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   ""
               BackColor       =   16777215
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":5454
               ColorButton     =   16777215
               AcclimateGrayTones=   -1  'True
               DrawFocusRectangle=   0   'False
               DisabledImageExtraction=   0
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnNext1 
               Height          =   315
               Left            =   555
               TabIndex        =   53
               Top             =   270
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   556
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   ""
               BackColor       =   16777215
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":57EE
               ColorButton     =   16777215
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnPrevious1 
               Height          =   315
               Left            =   1035
               TabIndex        =   54
               Top             =   270
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   556
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   ""
               BackColor       =   16777215
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":5B88
               ColorButton     =   16777215
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnFirst1 
               Height          =   315
               Left            =   1500
               TabIndex        =   55
               Top             =   270
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   556
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   ""
               BackColor       =   16777215
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":5F22
               ColorButton     =   16777215
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "«‰Ê«⁄ «·÷„«‰« "
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
               Left            =   5655
               RightToLeft     =   -1  'True
               TabIndex        =   56
               Top             =   330
               Width           =   4200
            End
         End
         Begin VB.Frame Frm21 
            BackColor       =   &H00E2E9E9&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   1485
            Left            =   840
            RightToLeft     =   -1  'True
            TabIndex        =   38
            Top             =   4200
            Width           =   8715
            Begin VB.TextBox TxtVacNamee1 
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
               Left            =   105
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   42
               Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«·   ‰Ê⁄ «·÷„«‰"
               Top             =   885
               Width           =   6240
            End
            Begin VB.TextBox TxtVacName1 
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
               Left            =   105
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   41
               Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«·   ‰Ê⁄ «·÷„«‰"
               Top             =   525
               Width           =   6240
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
               Left            =   4080
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   40
               Top             =   165
               Width           =   2265
            End
            Begin VB.ComboBox Combo1 
               BackColor       =   &H80000018&
               Height          =   315
               ItemData        =   "FrmPay_Garanty_Shipment.frx":62BC
               Left            =   2280
               List            =   "FrmPay_Garanty_Shipment.frx":62CC
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   39
               Top             =   2190
               Visible         =   0   'False
               Width           =   1005
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·÷„«‰ «‰Ã·Ì“Ì"
               Height          =   285
               Index           =   4
               Left            =   6480
               RightToLeft     =   -1  'True
               TabIndex        =   45
               Top             =   960
               Width           =   1410
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·÷„«‰ ⁄—»Ì"
               Height          =   285
               Index           =   5
               Left            =   6540
               RightToLeft     =   -1  'True
               TabIndex        =   44
               Top             =   600
               Width           =   1410
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬂÊœ  "
               Height          =   195
               Index           =   6
               Left            =   6705
               RightToLeft     =   -1  'True
               TabIndex        =   43
               Top             =   150
               Width           =   990
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic2 
            Height          =   660
            Left            =   2280
            TabIndex        =   57
            TabStop         =   0   'False
            Top             =   5775
            Width           =   5640
            _cx             =   9948
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
            Begin ImpulseButton.ISButton btnNew1 
               Height          =   330
               Left            =   4575
               TabIndex        =   58
               Top             =   315
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":62E5
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnSave1 
               Height          =   330
               Left            =   3030
               TabIndex        =   59
               Top             =   315
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":667F
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnModify1 
               Height          =   330
               Left            =   3795
               TabIndex        =   60
               Top             =   315
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":6A19
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton BtnUndo1 
               Height          =   330
               Left            =   2265
               TabIndex        =   61
               Top             =   315
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":6DB3
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnDelete1 
               Height          =   330
               Left            =   1500
               TabIndex        =   62
               Top             =   315
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":714D
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton ISButton10 
               Height          =   330
               Left            =   5880
               TabIndex        =   63
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":76E7
               ColorButton     =   14737632
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton BtnUpdate1 
               Height          =   330
               Left            =   6285
               TabIndex        =   64
               TabStop         =   0   'False
               ToolTipText     =   " ÕœÌÀ ﬁ«⁄œ… «·»Ì«‰« "
               Top             =   105
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":7A81
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton ISButton12 
               Height          =   285
               Left            =   6285
               TabIndex        =   65
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":7E1B
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnCancel1 
               Height          =   330
               Left            =   705
               TabIndex        =   66
               Top             =   315
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":81B5
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·”Ã· «·Õ«·Ì:"
               Height          =   210
               Index           =   2
               Left            =   2505
               RightToLeft     =   -1  'True
               TabIndex        =   70
               Top             =   -15
               Width           =   975
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄œœ «·”Ã·« :"
               Height          =   210
               Index           =   3
               Left            =   810
               RightToLeft     =   -1  'True
               TabIndex        =   69
               Top             =   -15
               Width           =   975
            End
            Begin VB.Label LabCurrRec1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Height          =   210
               Left            =   1800
               RightToLeft     =   -1  'True
               TabIndex        =   68
               Top             =   0
               Width           =   675
            End
            Begin VB.Label LabCountRec1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Height          =   210
               Left            =   240
               RightToLeft     =   -1  'True
               TabIndex        =   67
               Top             =   -15
               Width           =   540
            End
         End
         Begin VSFlex8Ctl.VSFlexGrid Grid1 
            Height          =   3435
            Left            =   60
            TabIndex        =   71
            Top             =   570
            Width           =   9825
            _cx             =   17330
            _cy             =   6059
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
            Cols            =   4
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmPay_Garanty_Shipment.frx":854F
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
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic3 
         Height          =   6510
         Left            =   -11760
         TabIndex        =   72
         TabStop         =   0   'False
         Top             =   45
         Width           =   9915
         _cx             =   17489
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
         Begin VB.Frame Frame1 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   825
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   96
            Top             =   -120
            Width           =   10155
            Begin VB.TextBox TxtVac_ID2 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H000000FF&
               Height          =   240
               Left            =   4590
               RightToLeft     =   -1  'True
               TabIndex        =   101
               Top             =   630
               Visible         =   0   'False
               Width           =   945
            End
            Begin VB.TextBox TxtModFlg2 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H0000FF00&
               Enabled         =   0   'False
               Height          =   285
               Left            =   6060
               RightToLeft     =   -1  'True
               TabIndex        =   100
               Text            =   "modflag"
               Top             =   690
               Visible         =   0   'False
               Width           =   465
            End
            Begin VB.Frame Frame2 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   375
               Left            =   540
               RightToLeft     =   -1  'True
               TabIndex        =   97
               Top             =   690
               Visible         =   0   'False
               Width           =   3105
               Begin MSDataListLib.DataCombo DataCombo1 
                  CausesValidation=   0   'False
                  Height          =   315
                  Left            =   -255
                  TabIndex        =   98
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
                  Index           =   7
                  Left            =   2280
                  RightToLeft     =   -1  'True
                  TabIndex        =   99
                  Top             =   45
                  Width           =   855
               End
            End
            Begin MSComctlLib.ImageList GrdImageList2 
               Left            =   3840
               Top             =   600
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
                     Picture         =   "FrmPay_Garanty_Shipment.frx":85E4
                     Key             =   "CompanyName"
                  EndProperty
                  BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmPay_Garanty_Shipment.frx":897E
                     Key             =   "Ser"
                  EndProperty
                  BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmPay_Garanty_Shipment.frx":8D18
                     Key             =   "Vac_Name"
                  EndProperty
                  BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmPay_Garanty_Shipment.frx":90B2
                     Key             =   "ShareCount"
                  EndProperty
                  BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmPay_Garanty_Shipment.frx":944C
                     Key             =   "Dis_Count"
                  EndProperty
                  BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmPay_Garanty_Shipment.frx":97E6
                     Key             =   "Bouns"
                  EndProperty
                  BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmPay_Garanty_Shipment.frx":9B80
                     Key             =   "SharesValue"
                  EndProperty
                  BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmPay_Garanty_Shipment.frx":A11A
                     Key             =   "BuyValue"
                  EndProperty
               EndProperty
            End
            Begin ImpulseButton.ISButton btnLast2 
               Height          =   315
               Left            =   210
               TabIndex        =   102
               Top             =   270
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   556
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   ""
               BackColor       =   16777215
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":A4B4
               ColorButton     =   16777215
               AcclimateGrayTones=   -1  'True
               DrawFocusRectangle=   0   'False
               DisabledImageExtraction=   0
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnNext2 
               Height          =   315
               Left            =   675
               TabIndex        =   103
               Top             =   270
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   556
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   ""
               BackColor       =   16777215
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":A84E
               ColorButton     =   16777215
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnPrevious2 
               Height          =   315
               Left            =   1155
               TabIndex        =   104
               Top             =   270
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   556
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   ""
               BackColor       =   16777215
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":ABE8
               ColorButton     =   16777215
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnFirst2 
               Height          =   315
               Left            =   1620
               TabIndex        =   105
               Top             =   270
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   556
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   ""
               BackColor       =   16777215
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":AF82
               ColorButton     =   16777215
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "«‰Ê«⁄ «·‘Õ‰"
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
               Left            =   6375
               RightToLeft     =   -1  'True
               TabIndex        =   106
               Top             =   210
               Width           =   3480
            End
         End
         Begin VB.Frame Frm22 
            BackColor       =   &H00E2E9E9&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   1290
            Left            =   1095
            RightToLeft     =   -1  'True
            TabIndex        =   73
            Top             =   4485
            Width           =   8190
            Begin VB.TextBox TxtVacNameE2 
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
               Left            =   105
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   77
               Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„  ÿ—Ìﬁ… «·‘Õ‰"
               Top             =   885
               Width           =   6360
            End
            Begin VB.TextBox TxtVacName2 
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
               Left            =   105
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   76
               Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„  ÿ—Ìﬁ… «·‘Õ‰"
               Top             =   525
               Width           =   6360
            End
            Begin VB.TextBox TxtSerial2 
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
               Left            =   4680
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   75
               Top             =   165
               Width           =   1785
            End
            Begin VB.ComboBox Combo2 
               BackColor       =   &H80000018&
               Height          =   315
               ItemData        =   "FrmPay_Garanty_Shipment.frx":B31C
               Left            =   2280
               List            =   "FrmPay_Garanty_Shipment.frx":B32C
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   74
               Top             =   1950
               Visible         =   0   'False
               Width           =   1005
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·«”„ «‰Ã·Ì“Ì"
               Height          =   285
               Index           =   9
               Left            =   6720
               RightToLeft     =   -1  'True
               TabIndex        =   80
               Top             =   840
               Width           =   1170
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·«”„ ⁄—»Ì"
               Height          =   285
               Index           =   10
               Left            =   6660
               RightToLeft     =   -1  'True
               TabIndex        =   79
               Top             =   480
               Width           =   1410
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬂÊœ "
               Height          =   195
               Index           =   11
               Left            =   6945
               RightToLeft     =   -1  'True
               TabIndex        =   78
               Top             =   150
               Width           =   870
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic4 
            Height          =   630
            Left            =   1965
            TabIndex        =   81
            TabStop         =   0   'False
            Top             =   5805
            Width           =   5625
            _cx             =   9922
            _cy             =   1111
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
            Begin ImpulseButton.ISButton btnNew2 
               Height          =   330
               Left            =   4575
               TabIndex        =   82
               Top             =   315
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":B345
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnSave2 
               Height          =   330
               Left            =   3030
               TabIndex        =   83
               Top             =   315
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":B6DF
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnModify2 
               Height          =   330
               Left            =   3795
               TabIndex        =   84
               Top             =   315
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":BA79
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton BtnUndo2 
               Height          =   330
               Left            =   2265
               TabIndex        =   85
               Top             =   315
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":BE13
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnDelete2 
               Height          =   330
               Left            =   1440
               TabIndex        =   86
               Top             =   315
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":C1AD
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton ISButton11 
               Height          =   330
               Left            =   4560
               TabIndex        =   87
               TabStop         =   0   'False
               ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
               Top             =   1410
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":C747
               ColorButton     =   14737632
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton BtnUpdate2 
               Height          =   330
               Left            =   8205
               TabIndex        =   88
               TabStop         =   0   'False
               ToolTipText     =   " ÕœÌÀ ﬁ«⁄œ… «·»Ì«‰« "
               Top             =   1425
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":CAE1
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton ISButton14 
               Height          =   285
               Left            =   6885
               TabIndex        =   89
               TabStop         =   0   'False
               Top             =   1470
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":CE7B
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnCancel2 
               Height          =   330
               Left            =   585
               TabIndex        =   90
               Top             =   315
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":D215
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·”Ã· «·Õ«·Ì:"
               Height          =   210
               Index           =   4
               Left            =   2505
               RightToLeft     =   -1  'True
               TabIndex        =   94
               Top             =   -15
               Width           =   975
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄œœ «·”Ã·« :"
               Height          =   210
               Index           =   5
               Left            =   810
               RightToLeft     =   -1  'True
               TabIndex        =   93
               Top             =   -15
               Width           =   975
            End
            Begin VB.Label LabCurrRec2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Height          =   210
               Left            =   1800
               RightToLeft     =   -1  'True
               TabIndex        =   92
               Top             =   0
               Width           =   675
            End
            Begin VB.Label LabCountRec2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Height          =   210
               Left            =   240
               RightToLeft     =   -1  'True
               TabIndex        =   91
               Top             =   -15
               Width           =   540
            End
         End
         Begin VSFlex8Ctl.VSFlexGrid Grid2 
            Height          =   3705
            Left            =   120
            TabIndex        =   95
            Top             =   855
            Width           =   9675
            _cx             =   17066
            _cy             =   6535
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
            Cols            =   4
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmPay_Garanty_Shipment.frx":D5AF
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
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic5 
         Height          =   6510
         Left            =   -11460
         TabIndex        =   107
         TabStop         =   0   'False
         Top             =   45
         Width           =   9915
         _cx             =   17489
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
         Begin VB.Frame Frm23 
            BackColor       =   &H00E2E9E9&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   1245
            Left            =   1080
            RightToLeft     =   -1  'True
            TabIndex        =   119
            Top             =   4200
            Width           =   8235
            Begin VB.TextBox TxtVacNamee3 
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
               Left            =   105
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   123
               Top             =   765
               Width           =   6240
            End
            Begin VB.TextBox TxtVacName3 
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
               Left            =   105
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   122
               Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·„Ã„Ê⁄Â"
               Top             =   405
               Width           =   6240
            End
            Begin VB.TextBox TxtSerial3 
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
               Left            =   4560
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   121
               Top             =   45
               Width           =   1785
            End
            Begin VB.ComboBox Combo3 
               BackColor       =   &H80000018&
               Height          =   315
               ItemData        =   "FrmPay_Garanty_Shipment.frx":D643
               Left            =   2280
               List            =   "FrmPay_Garanty_Shipment.frx":D653
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   120
               Top             =   1710
               Visible         =   0   'False
               Width           =   1005
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·„Ã„Ê⁄Â «‰Ã·Ì“Ì"
               Height          =   285
               Index           =   14
               Left            =   6480
               RightToLeft     =   -1  'True
               TabIndex        =   126
               Top             =   840
               Width           =   1650
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·„Ã„Ê⁄Â ⁄—»Ì"
               Height          =   285
               Index           =   15
               Left            =   6540
               RightToLeft     =   -1  'True
               TabIndex        =   125
               Top             =   480
               Width           =   1650
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬂÊœ  "
               Height          =   195
               Index           =   16
               Left            =   6945
               RightToLeft     =   -1  'True
               TabIndex        =   124
               Top             =   120
               Width           =   870
            End
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   540
            Left            =   -165
            RightToLeft     =   -1  'True
            TabIndex        =   108
            Top             =   0
            Width           =   10185
            Begin VB.Frame Frame5 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   375
               Left            =   540
               RightToLeft     =   -1  'True
               TabIndex        =   111
               Top             =   450
               Visible         =   0   'False
               Width           =   3105
               Begin MSDataListLib.DataCombo DCUser3 
                  CausesValidation=   0   'False
                  Height          =   315
                  Left            =   -255
                  TabIndex        =   112
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
                  Index           =   17
                  Left            =   2160
                  RightToLeft     =   -1  'True
                  TabIndex        =   113
                  Top             =   45
                  Width           =   855
               End
            End
            Begin VB.TextBox TxtModFlg3 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H0000FF00&
               Enabled         =   0   'False
               Height          =   285
               Left            =   4260
               RightToLeft     =   -1  'True
               TabIndex        =   110
               Text            =   "modflag"
               Top             =   450
               Visible         =   0   'False
               Width           =   465
            End
            Begin VB.TextBox TxtVac_ID3 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H000000FF&
               Height          =   240
               Left            =   3270
               RightToLeft     =   -1  'True
               TabIndex        =   109
               Top             =   390
               Visible         =   0   'False
               Width           =   945
            End
            Begin MSComctlLib.ImageList GrdImageList3 
               Left            =   4920
               Top             =   360
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
                     Picture         =   "FrmPay_Garanty_Shipment.frx":D66C
                     Key             =   "CompanyName"
                  EndProperty
                  BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmPay_Garanty_Shipment.frx":DA06
                     Key             =   "Ser"
                  EndProperty
                  BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmPay_Garanty_Shipment.frx":DDA0
                     Key             =   "Vac_Name"
                  EndProperty
                  BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmPay_Garanty_Shipment.frx":E13A
                     Key             =   "ShareCount"
                  EndProperty
                  BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmPay_Garanty_Shipment.frx":E4D4
                     Key             =   "Dis_Count"
                  EndProperty
                  BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmPay_Garanty_Shipment.frx":E86E
                     Key             =   "Bouns"
                  EndProperty
                  BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmPay_Garanty_Shipment.frx":EC08
                     Key             =   "SharesValue"
                  EndProperty
                  BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmPay_Garanty_Shipment.frx":F1A2
                     Key             =   "BuyValue"
                  EndProperty
               EndProperty
            End
            Begin ImpulseButton.ISButton btnLast3 
               Height          =   315
               Left            =   330
               TabIndex        =   114
               Top             =   30
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   556
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   ""
               BackColor       =   16777215
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":F53C
               ColorButton     =   16777215
               AcclimateGrayTones=   -1  'True
               DrawFocusRectangle=   0   'False
               DisabledImageExtraction=   0
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnNext3 
               Height          =   315
               Left            =   795
               TabIndex        =   115
               Top             =   30
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   556
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   ""
               BackColor       =   16777215
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":F8D6
               ColorButton     =   16777215
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnPrevious3 
               Height          =   315
               Left            =   1275
               TabIndex        =   116
               Top             =   30
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   556
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   ""
               BackColor       =   16777215
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":FC70
               ColorButton     =   16777215
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnFirst3 
               Height          =   315
               Left            =   1740
               TabIndex        =   117
               Top             =   30
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   556
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   ""
               BackColor       =   16777215
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":1000A
               ColorButton     =   16777215
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "„Ã„Ê⁄«  «·„‰«œÌ» ··„‘ —Ì« "
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
               Index           =   18
               Left            =   5775
               RightToLeft     =   -1  'True
               TabIndex        =   118
               Top             =   90
               Width           =   4200
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic6 
            Height          =   780
            Left            =   2160
            TabIndex        =   127
            TabStop         =   0   'False
            Top             =   5640
            Width           =   5520
            _cx             =   9737
            _cy             =   1376
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
            Begin ImpulseButton.ISButton btnNew3 
               Height          =   330
               Left            =   4575
               TabIndex        =   128
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":103A4
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnSave3 
               Height          =   330
               Left            =   3030
               TabIndex        =   129
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":1073E
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnModify3 
               Height          =   330
               Left            =   3795
               TabIndex        =   130
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":10AD8
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton BtnUndo3 
               Height          =   330
               Left            =   2265
               TabIndex        =   131
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":10E72
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnDelete3 
               Height          =   330
               Left            =   1500
               TabIndex        =   132
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":1120C
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton BtnUpdate3 
               Height          =   330
               Left            =   5880
               TabIndex        =   133
               TabStop         =   0   'False
               ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
               Top             =   1410
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":117A6
               ColorButton     =   14737632
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton ISButton15 
               Height          =   330
               Left            =   6045
               TabIndex        =   134
               TabStop         =   0   'False
               ToolTipText     =   " ÕœÌÀ ﬁ«⁄œ… «·»Ì«‰« "
               Top             =   1185
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":11B40
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton ISButton16 
               Height          =   285
               Left            =   4725
               TabIndex        =   135
               TabStop         =   0   'False
               Top             =   1350
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":11EDA
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnCancel3 
               Height          =   330
               Left            =   705
               TabIndex        =   136
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":12274
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin VB.Label LabCountRec3 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Height          =   210
               Left            =   240
               RightToLeft     =   -1  'True
               TabIndex        =   140
               Top             =   105
               Width           =   540
            End
            Begin VB.Label LabCurrRec3 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Height          =   210
               Left            =   1800
               RightToLeft     =   -1  'True
               TabIndex        =   139
               Top             =   105
               Width           =   675
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄œœ «·”Ã·« :"
               Height          =   210
               Index           =   6
               Left            =   810
               RightToLeft     =   -1  'True
               TabIndex        =   138
               Top             =   105
               Width           =   975
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·”Ã· «·Õ«·Ì:"
               Height          =   210
               Index           =   7
               Left            =   2505
               RightToLeft     =   -1  'True
               TabIndex        =   137
               Top             =   105
               Width           =   975
            End
         End
         Begin VSFlex8Ctl.VSFlexGrid Grid3 
            Height          =   3435
            Left            =   60
            TabIndex        =   141
            Top             =   570
            Width           =   9825
            _cx             =   17330
            _cy             =   6059
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
            Cols            =   4
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmPay_Garanty_Shipment.frx":1260E
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
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic7 
         Height          =   6510
         Left            =   -11160
         TabIndex        =   142
         TabStop         =   0   'False
         Top             =   45
         Width           =   9915
         _cx             =   17489
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
         Begin VB.Frame Frm24 
            BackColor       =   &H00E2E9E9&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   1365
            Left            =   855
            RightToLeft     =   -1  'True
            TabIndex        =   143
            Top             =   4080
            Width           =   8430
            Begin VB.TextBox TxtUnitID4 
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
               Height          =   345
               Left            =   4995
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   147
               Top             =   240
               Width           =   2025
            End
            Begin VB.TextBox TxtVacNamee4 
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
               Height          =   345
               Left            =   120
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   146
               Top             =   960
               Width           =   6900
            End
            Begin VB.TextBox TxtVacName4 
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
               Height          =   345
               Left            =   120
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   145
               Top             =   585
               Width           =   6900
            End
            Begin VB.ComboBox Combo4 
               BackColor       =   &H80000018&
               Height          =   315
               ItemData        =   "FrmPay_Garanty_Shipment.frx":1269F
               Left            =   2280
               List            =   "FrmPay_Garanty_Shipment.frx":126AF
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   144
               Top             =   3150
               Visible         =   0   'False
               Width           =   1005
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·«”„ «‰Ã·Ì“Ì"
               Height          =   375
               Index           =   19
               Left            =   7080
               RightToLeft     =   -1  'True
               TabIndex        =   150
               Top             =   1080
               Width           =   1170
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ﬂÊœ"
               Height          =   285
               Index           =   20
               Left            =   7485
               RightToLeft     =   -1  'True
               TabIndex        =   149
               Top             =   330
               Width           =   570
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·«”„ ⁄—»Ì"
               Height          =   255
               Index           =   21
               Left            =   7110
               RightToLeft     =   -1  'True
               TabIndex        =   148
               Top             =   690
               Width           =   1170
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic17 
            Height          =   675
            Left            =   0
            TabIndex        =   151
            TabStop         =   0   'False
            Top             =   0
            Width           =   9915
            _cx             =   17489
            _cy             =   1191
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial (Arabic)"
               Size            =   21.75
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
            BackColor       =   16777215
            ForeColor       =   4210688
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
            CaptionPos      =   7
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
            Begin VB.TextBox TxtModFlg4 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H0000FF00&
               Enabled         =   0   'False
               Height          =   285
               Left            =   2760
               RightToLeft     =   -1  'True
               TabIndex        =   152
               Top             =   240
               Visible         =   0   'False
               Width           =   945
            End
            Begin ImpulseButton.ISButton btnLast4 
               Height          =   345
               Left            =   150
               TabIndex        =   153
               Top             =   150
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   609
               ButtonStyle     =   1
               ButtonPositionImage=   4
               Caption         =   ""
               BackColor       =   16777215
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":126C8
               ColorButton     =   16777215
               DrawFocusRectangle=   0   'False
               DisabledImageExtraction=   0
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnNext4 
               Height          =   345
               Left            =   615
               TabIndex        =   154
               Top             =   150
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   609
               ButtonStyle     =   1
               ButtonPositionImage=   4
               Caption         =   ""
               BackColor       =   16777215
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":12A62
               ColorButton     =   16777215
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnPrevious4 
               Height          =   345
               Left            =   1065
               TabIndex        =   155
               Top             =   150
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   609
               ButtonStyle     =   1
               ButtonPositionImage=   4
               Caption         =   ""
               BackColor       =   16777215
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":12DFC
               ColorButton     =   16777215
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnFirst4 
               Height          =   345
               Left            =   1530
               TabIndex        =   156
               Top             =   150
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   609
               ButtonStyle     =   1
               ButtonPositionImage=   4
               Caption         =   ""
               BackColor       =   16777215
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":13196
               ColorButton     =   16777215
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "ÿ—ﬁ «·‘Õ‰"
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
               Index           =   22
               Left            =   5520
               RightToLeft     =   -1  'True
               TabIndex        =   157
               Top             =   120
               Width           =   4200
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic18 
            Height          =   990
            Left            =   1710
            TabIndex        =   158
            TabStop         =   0   'False
            Top             =   5490
            Width           =   6855
            _cx             =   12091
            _cy             =   1746
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
            Begin ImpulseButton.ISButton btnNew4 
               Height          =   420
               Left            =   5805
               TabIndex        =   159
               Top             =   495
               Width           =   720
               _ExtentX        =   1270
               _ExtentY        =   741
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":13530
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnSave4 
               Height          =   420
               Left            =   4290
               TabIndex        =   160
               Top             =   495
               Width           =   720
               _ExtentX        =   1270
               _ExtentY        =   741
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":138CA
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnModify4 
               Height          =   420
               Left            =   4935
               TabIndex        =   161
               Top             =   495
               Width           =   825
               _ExtentX        =   1455
               _ExtentY        =   741
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":13C64
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton BtnUndo4 
               Height          =   420
               Left            =   3405
               TabIndex        =   162
               Top             =   495
               Width           =   840
               _ExtentX        =   1482
               _ExtentY        =   741
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":13FFE
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnDelete4 
               Height          =   420
               Left            =   990
               TabIndex        =   163
               Top             =   495
               Width           =   690
               _ExtentX        =   1217
               _ExtentY        =   741
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":14398
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnQuery4 
               Height          =   330
               Left            =   2520
               TabIndex        =   164
               TabStop         =   0   'False
               ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
               Top             =   570
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":14932
               ColorButton     =   14737632
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton BtnUpdate4 
               Height          =   330
               Left            =   5805
               TabIndex        =   165
               TabStop         =   0   'False
               ToolTipText     =   " ÕœÌÀ ﬁ«⁄œ… «·»Ì«‰« "
               Top             =   105
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":14CCC
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton ISButton18 
               Height          =   285
               Left            =   4665
               TabIndex        =   166
               TabStop         =   0   'False
               Top             =   60
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":15066
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnCancel4 
               Height          =   420
               Left            =   75
               TabIndex        =   167
               Top             =   495
               Width           =   870
               _ExtentX        =   1535
               _ExtentY        =   741
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":15400
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton ButPrient4 
               Height          =   495
               Left            =   1680
               TabIndex        =   168
               Top             =   480
               Visible         =   0   'False
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   873
               ButtonStyle     =   1
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
               BackStyle       =   0
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":1579A
               ColorButton     =   14871017
               DisplayPersistentHover=   0   'False
               DrawFocusRectangle=   0   'False
            End
            Begin VB.Label LabCountRec4 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Height          =   210
               Left            =   270
               RightToLeft     =   -1  'True
               TabIndex        =   172
               Top             =   135
               Width           =   540
            End
            Begin VB.Label LabCurrRec4 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Height          =   210
               Left            =   2745
               RightToLeft     =   -1  'True
               TabIndex        =   171
               Top             =   165
               Width           =   675
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄œœ «·”Ã·« :"
               Height          =   210
               Index           =   8
               Left            =   960
               RightToLeft     =   -1  'True
               TabIndex        =   170
               Top             =   135
               Width           =   975
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·”Ã· «·Õ«·Ì:"
               Height          =   210
               Index           =   9
               Left            =   3555
               RightToLeft     =   -1  'True
               TabIndex        =   169
               Top             =   135
               Width           =   975
            End
         End
         Begin VSFlex8Ctl.VSFlexGrid Grid4 
            Height          =   3165
            Left            =   0
            TabIndex        =   173
            Top             =   720
            Width           =   9915
            _cx             =   17489
            _cy             =   5583
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
            Cols            =   6
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmPay_Garanty_Shipment.frx":15B34
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
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic8 
         Height          =   6510
         Left            =   -10860
         TabIndex        =   174
         TabStop         =   0   'False
         Top             =   45
         Width           =   9915
         _cx             =   17489
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
         Begin VB.Frame Frame6 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   540
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   183
            Top             =   0
            Width           =   9915
            Begin VB.TextBox TxtVac_ID5 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H000000FF&
               Height          =   240
               Left            =   4830
               RightToLeft     =   -1  'True
               TabIndex        =   188
               Top             =   390
               Visible         =   0   'False
               Width           =   945
            End
            Begin VB.TextBox TxtModFlg5 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H0000FF00&
               Enabled         =   0   'False
               Height          =   285
               Left            =   3060
               RightToLeft     =   -1  'True
               TabIndex        =   187
               Text            =   "modflag"
               Top             =   90
               Visible         =   0   'False
               Width           =   465
            End
            Begin VB.Frame Frame7 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   375
               Left            =   540
               RightToLeft     =   -1  'True
               TabIndex        =   184
               Top             =   450
               Visible         =   0   'False
               Width           =   3105
               Begin MSDataListLib.DataCombo DCUser5 
                  CausesValidation=   0   'False
                  Height          =   315
                  Left            =   -255
                  TabIndex        =   185
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
                  Index           =   23
                  Left            =   2160
                  RightToLeft     =   -1  'True
                  TabIndex        =   186
                  Top             =   45
                  Width           =   855
               End
            End
            Begin MSComctlLib.ImageList GrdImageList5 
               Left            =   4200
               Top             =   480
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
                     Picture         =   "FrmPay_Garanty_Shipment.frx":15C17
                     Key             =   "CompanyName"
                  EndProperty
                  BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmPay_Garanty_Shipment.frx":15FB1
                     Key             =   "Ser"
                  EndProperty
                  BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmPay_Garanty_Shipment.frx":1634B
                     Key             =   "Vac_Name"
                  EndProperty
                  BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmPay_Garanty_Shipment.frx":166E5
                     Key             =   "ShareCount"
                  EndProperty
                  BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmPay_Garanty_Shipment.frx":16A7F
                     Key             =   "Dis_Count"
                  EndProperty
                  BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmPay_Garanty_Shipment.frx":16E19
                     Key             =   "Bouns"
                  EndProperty
                  BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmPay_Garanty_Shipment.frx":171B3
                     Key             =   "SharesValue"
                  EndProperty
                  BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmPay_Garanty_Shipment.frx":1774D
                     Key             =   "BuyValue"
                  EndProperty
               EndProperty
            End
            Begin ImpulseButton.ISButton btnLast5 
               Height          =   315
               Left            =   90
               TabIndex        =   189
               Top             =   30
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   556
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   ""
               BackColor       =   16777215
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":17AE7
               ColorButton     =   16777215
               AcclimateGrayTones=   -1  'True
               DrawFocusRectangle=   0   'False
               DisabledImageExtraction=   0
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnNext5 
               Height          =   315
               Left            =   555
               TabIndex        =   190
               Top             =   30
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   556
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   ""
               BackColor       =   16777215
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":17E81
               ColorButton     =   16777215
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnPrevious5 
               Height          =   315
               Left            =   1155
               TabIndex        =   191
               Top             =   30
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   556
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   ""
               BackColor       =   16777215
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":1821B
               ColorButton     =   16777215
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnFirst5 
               Height          =   315
               Left            =   1620
               TabIndex        =   192
               Top             =   30
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   556
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   ""
               BackColor       =   16777215
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":185B5
               ColorButton     =   16777215
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "«‰Ê«⁄ «·÷„«‰ ·· ﬁ”Ìÿ"
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
               Index           =   24
               Left            =   4695
               RightToLeft     =   -1  'True
               TabIndex        =   193
               Top             =   90
               Width           =   5160
            End
         End
         Begin VB.Frame Frm25 
            BackColor       =   &H00E2E9E9&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   1365
            Left            =   840
            RightToLeft     =   -1  'True
            TabIndex        =   175
            Top             =   4275
            Width           =   8205
            Begin VB.TextBox TxtVacNamee5 
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
               Left            =   75
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   179
               Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· ‰Ê⁄ «·÷„«‰ «‰Ã·Ì“Ì"
               Top             =   765
               Width           =   6840
            End
            Begin VB.TextBox TxtVacName5 
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
               Left            =   75
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   178
               Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· ‰Ê⁄ «·÷„«‰ ⁄—»Ì"
               Top             =   405
               Width           =   6840
            End
            Begin VB.TextBox TxtSerial5 
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
               Left            =   4890
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   177
               Top             =   45
               Width           =   2025
            End
            Begin VB.ComboBox Combo5 
               BackColor       =   &H80000018&
               Height          =   315
               ItemData        =   "FrmPay_Garanty_Shipment.frx":1894F
               Left            =   2280
               List            =   "FrmPay_Garanty_Shipment.frx":1895F
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   176
               Top             =   1470
               Visible         =   0   'False
               Width           =   1005
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·«”„ «‰Ã·Ì“Ì"
               Height          =   285
               Index           =   25
               Left            =   6945
               RightToLeft     =   -1  'True
               TabIndex        =   182
               Top             =   840
               Width           =   1170
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·«”„ ⁄—»Ì"
               Height          =   285
               Index           =   26
               Left            =   6900
               RightToLeft     =   -1  'True
               TabIndex        =   181
               Top             =   480
               Width           =   1290
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬂÊœ  "
               Height          =   195
               Index           =   27
               Left            =   7065
               RightToLeft     =   -1  'True
               TabIndex        =   180
               Top             =   150
               Width           =   870
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic19 
            Height          =   900
            Left            =   2205
            TabIndex        =   194
            TabStop         =   0   'False
            Top             =   5535
            Width           =   5505
            _cx             =   9710
            _cy             =   1588
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
            Begin ImpulseButton.ISButton btnNew5 
               Height          =   330
               Left            =   4575
               TabIndex        =   195
               Top             =   555
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":18978
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnSave5 
               Height          =   330
               Left            =   3030
               TabIndex        =   196
               Top             =   555
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":18D12
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnModify5 
               Height          =   330
               Left            =   3795
               TabIndex        =   197
               Top             =   555
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":190AC
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton BtnUndo5 
               Height          =   330
               Left            =   2265
               TabIndex        =   198
               Top             =   555
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":19446
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnDelete5 
               Height          =   330
               Left            =   1500
               TabIndex        =   199
               Top             =   555
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":197E0
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnQuery5 
               Height          =   330
               Left            =   5880
               TabIndex        =   200
               TabStop         =   0   'False
               ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
               Top             =   930
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":19D7A
               ColorButton     =   14737632
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton BtnUpdate5 
               Height          =   330
               Left            =   5085
               TabIndex        =   201
               TabStop         =   0   'False
               ToolTipText     =   " ÕœÌÀ ﬁ«⁄œ… «·»Ì«‰« "
               Top             =   945
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":1A114
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton ISButton19 
               Height          =   285
               Left            =   4725
               TabIndex        =   202
               TabStop         =   0   'False
               Top             =   990
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":1A4AE
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnCancel5 
               Height          =   330
               Left            =   705
               TabIndex        =   203
               Top             =   555
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":1A848
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·”Ã· «·Õ«·Ì:"
               Height          =   210
               Index           =   10
               Left            =   2505
               RightToLeft     =   -1  'True
               TabIndex        =   207
               Top             =   225
               Width           =   975
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄œœ «·”Ã·« :"
               Height          =   210
               Index           =   11
               Left            =   810
               RightToLeft     =   -1  'True
               TabIndex        =   206
               Top             =   225
               Width           =   975
            End
            Begin VB.Label LabCurrRec5 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Height          =   210
               Left            =   1800
               RightToLeft     =   -1  'True
               TabIndex        =   205
               Top             =   240
               Width           =   675
            End
            Begin VB.Label LabCountRec5 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Height          =   210
               Left            =   240
               RightToLeft     =   -1  'True
               TabIndex        =   204
               Top             =   225
               Width           =   540
            End
         End
         Begin VSFlex8Ctl.VSFlexGrid Grid5 
            Height          =   3435
            Left            =   0
            TabIndex        =   208
            Top             =   570
            Width           =   9915
            _cx             =   17489
            _cy             =   6059
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
            Cols            =   4
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmPay_Garanty_Shipment.frx":1ABE2
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
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic9 
         Height          =   6510
         Left            =   -10560
         TabIndex        =   209
         TabStop         =   0   'False
         Top             =   45
         Width           =   9915
         _cx             =   17489
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
         Begin VB.Frame Frame9 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   540
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   225
            Top             =   0
            Width           =   9915
            Begin VB.TextBox TxtVac_ID6 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H000000FF&
               Height          =   240
               Left            =   4110
               RightToLeft     =   -1  'True
               TabIndex        =   230
               Top             =   30
               Visible         =   0   'False
               Width           =   945
            End
            Begin VB.TextBox TxtModFlg6 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H0000FF00&
               Enabled         =   0   'False
               Height          =   285
               Left            =   3420
               RightToLeft     =   -1  'True
               TabIndex        =   229
               Text            =   "modflag"
               Top             =   -30
               Visible         =   0   'False
               Width           =   465
            End
            Begin VB.Frame Frame10 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   375
               Left            =   1980
               RightToLeft     =   -1  'True
               TabIndex        =   226
               Top             =   330
               Visible         =   0   'False
               Width           =   3105
               Begin MSDataListLib.DataCombo DCUser6 
                  CausesValidation=   0   'False
                  Height          =   315
                  Left            =   -255
                  TabIndex        =   227
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
                  Index           =   28
                  Left            =   2280
                  RightToLeft     =   -1  'True
                  TabIndex        =   228
                  Top             =   45
                  Width           =   735
               End
            End
            Begin MSComctlLib.ImageList GrdImageList6 
               Left            =   5160
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
                     Picture         =   "FrmPay_Garanty_Shipment.frx":1AC76
                     Key             =   "CompanyName"
                  EndProperty
                  BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmPay_Garanty_Shipment.frx":1B010
                     Key             =   "Ser"
                  EndProperty
                  BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmPay_Garanty_Shipment.frx":1B3AA
                     Key             =   "Vac_Name"
                  EndProperty
                  BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmPay_Garanty_Shipment.frx":1B744
                     Key             =   "ShareCount"
                  EndProperty
                  BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmPay_Garanty_Shipment.frx":1BADE
                     Key             =   "Dis_Count"
                  EndProperty
                  BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmPay_Garanty_Shipment.frx":1BE78
                     Key             =   "Bouns"
                  EndProperty
                  BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmPay_Garanty_Shipment.frx":1C212
                     Key             =   "SharesValue"
                  EndProperty
                  BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmPay_Garanty_Shipment.frx":1C7AC
                     Key             =   "BuyValue"
                  EndProperty
               EndProperty
            End
            Begin ImpulseButton.ISButton btnLast6 
               Height          =   315
               Left            =   90
               TabIndex        =   231
               Top             =   30
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   556
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   ""
               BackColor       =   16777215
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":1CB46
               ColorButton     =   16777215
               AcclimateGrayTones=   -1  'True
               DrawFocusRectangle=   0   'False
               DisabledImageExtraction=   0
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnNext6 
               Height          =   315
               Left            =   555
               TabIndex        =   232
               Top             =   30
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   556
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   ""
               BackColor       =   16777215
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":1CEE0
               ColorButton     =   16777215
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnPrevious6 
               Height          =   315
               Left            =   1155
               TabIndex        =   233
               Top             =   30
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   556
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   ""
               BackColor       =   16777215
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":1D27A
               ColorButton     =   16777215
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnFirst6 
               Height          =   315
               Left            =   1620
               TabIndex        =   234
               Top             =   30
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   556
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   ""
               BackColor       =   16777215
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":1D614
               ColorButton     =   16777215
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "»Ì«‰«  «·„‰«œÌ» ··„‘ —Ì« "
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
               Index           =   29
               Left            =   5895
               RightToLeft     =   -1  'True
               TabIndex        =   235
               Top             =   90
               Width           =   3870
            End
         End
         Begin VB.Frame Frm26 
            BackColor       =   &H00E2E9E9&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   1485
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   210
            Top             =   4005
            Width           =   9885
            Begin VB.TextBox TXTCode 
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
               Left            =   7680
               Locked          =   -1  'True
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   214
               Top             =   240
               Width           =   1050
            End
            Begin VB.TextBox TXTDiscounts 
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
               Left            =   7665
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   213
               Tag             =   "⁄›Ê« Ì—ÃÏ   ‰”»… «·Œ’„"
               Top             =   720
               Width           =   1050
            End
            Begin VB.TextBox Text3 
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
               Left            =   8025
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   212
               Top             =   1590
               Visible         =   0   'False
               Width           =   1410
            End
            Begin VB.ComboBox Combo6 
               BackColor       =   &H80000018&
               Height          =   315
               ItemData        =   "FrmPay_Garanty_Shipment.frx":1D9AE
               Left            =   2280
               List            =   "FrmPay_Garanty_Shipment.frx":1D9BE
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   211
               Top             =   1590
               Visible         =   0   'False
               Width           =   1005
            End
            Begin MSDataListLib.DataCombo DCEmP 
               Height          =   315
               Left            =   3750
               TabIndex        =   215
               Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· «Œ Ì«— «·„‰œÊ»"
               Top             =   270
               Width           =   2835
               _ExtentX        =   5001
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DCSalesRepGroups 
               Height          =   315
               Left            =   120
               TabIndex        =   216
               Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·„Ã„Ê⁄Â"
               Top             =   240
               Width           =   2595
               _ExtentX        =   4577
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcBranches 
               Height          =   315
               Left            =   3750
               TabIndex        =   217
               Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·›—⁄"
               Top             =   720
               Width           =   2835
               _ExtentX        =   5001
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DCJob 
               Height          =   315
               Left            =   120
               TabIndex        =   218
               Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«·  «·ÊŸÌ€…"
               Top             =   720
               Visible         =   0   'False
               Width           =   2595
               _ExtentX        =   4577
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ÊŸÌ›Â"
               Height          =   285
               Index           =   30
               Left            =   2850
               RightToLeft     =   -1  'True
               TabIndex        =   224
               Top             =   720
               Visible         =   0   'False
               Width           =   810
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·›—⁄"
               Height          =   285
               Index           =   31
               Left            =   6690
               RightToLeft     =   -1  'True
               TabIndex        =   223
               Top             =   720
               Width           =   930
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„Ã„Ê⁄Â"
               Height          =   285
               Index           =   32
               Left            =   2850
               RightToLeft     =   -1  'True
               TabIndex        =   222
               Top             =   240
               Width           =   810
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "‰”»Â «·Œ’„"
               Height          =   285
               Index           =   33
               Left            =   8820
               RightToLeft     =   -1  'True
               TabIndex        =   221
               Top             =   750
               Width           =   930
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬂÊœ «·„‰œÊ»"
               Height          =   195
               Index           =   34
               Left            =   8805
               RightToLeft     =   -1  'True
               TabIndex        =   220
               Top             =   270
               Width           =   990
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·„‰œÊ»"
               Height          =   285
               Index           =   35
               Left            =   6720
               RightToLeft     =   -1  'True
               TabIndex        =   219
               Top             =   270
               Width           =   930
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic20 
            Height          =   1020
            Left            =   -240
            TabIndex        =   236
            TabStop         =   0   'False
            Top             =   5385
            Width           =   10155
            _cx             =   17912
            _cy             =   1799
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
            Begin ImpulseButton.ISButton btnNew6 
               Height          =   330
               Left            =   9375
               TabIndex        =   237
               Top             =   555
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":1D9D7
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnSave6 
               Height          =   330
               Left            =   7590
               TabIndex        =   238
               Top             =   555
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":1DD71
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnModify6 
               Height          =   330
               Left            =   8475
               TabIndex        =   239
               Top             =   555
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":1E10B
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton BtnUndo6 
               Height          =   330
               Left            =   6705
               TabIndex        =   240
               Top             =   555
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":1E4A5
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnDelete6 
               Height          =   330
               Left            =   1140
               TabIndex        =   241
               Top             =   555
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":1E83F
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton BtnUpdate6 
               Height          =   330
               Left            =   6045
               TabIndex        =   242
               TabStop         =   0   'False
               ToolTipText     =   " ÕœÌÀ ﬁ«⁄œ… «·»Ì«‰« "
               Top             =   105
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":1EDD9
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnCancel6 
               Height          =   330
               Left            =   345
               TabIndex        =   243
               Top             =   555
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":1F173
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnQuery6 
               Height          =   330
               Left            =   5760
               TabIndex        =   244
               TabStop         =   0   'False
               ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
               Top             =   555
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":1F50D
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton ISButton1 
               Height          =   285
               Left            =   4965
               TabIndex        =   245
               TabStop         =   0   'False
               Top             =   555
               Width           =   765
               _ExtentX        =   1349
               _ExtentY        =   503
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":1F8A7
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton ISButton2 
               Height          =   405
               Left            =   3600
               TabIndex        =   246
               TabStop         =   0   'False
               Top             =   480
               Width           =   1245
               _ExtentX        =   2196
               _ExtentY        =   714
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   "ÿ»«⁄…  »Õ”» «·„Ã„Ê⁄…"
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":1FC41
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton ISButton3 
               Height          =   405
               Left            =   1920
               TabIndex        =   247
               TabStop         =   0   'False
               Top             =   555
               Width           =   1485
               _ExtentX        =   2619
               _ExtentY        =   714
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   "ÿ»«⁄… »Õ”» «·›—⁄"
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":1FFDB
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·”Ã· «·Õ«·Ì:"
               Height          =   210
               Index           =   12
               Left            =   2745
               RightToLeft     =   -1  'True
               TabIndex        =   251
               Top             =   225
               Width           =   975
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄œœ «·”Ã·« :"
               Height          =   210
               Index           =   13
               Left            =   1050
               RightToLeft     =   -1  'True
               TabIndex        =   250
               Top             =   225
               Width           =   975
            End
            Begin VB.Label LabCurrRec6 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Height          =   210
               Left            =   2040
               RightToLeft     =   -1  'True
               TabIndex        =   249
               Top             =   240
               Width           =   675
            End
            Begin VB.Label LabCountRec6 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Height          =   210
               Left            =   480
               RightToLeft     =   -1  'True
               TabIndex        =   248
               Top             =   225
               Width           =   540
            End
         End
         Begin VSFlex8UCtl.VSFlexGrid Grid6 
            Height          =   3405
            Left            =   0
            TabIndex        =   252
            Top             =   600
            Width           =   9915
            _cx             =   17489
            _cy             =   6006
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
            Cols            =   8
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmPay_Garanty_Shipment.frx":20375
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
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic10 
         Height          =   6510
         Left            =   45
         TabIndex        =   253
         TabStop         =   0   'False
         Top             =   45
         Width           =   9915
         _cx             =   17489
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
         Begin VB.Frame Frame12 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   540
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   270
            Top             =   0
            Width           =   9915
            Begin VB.TextBox TxtVac_ID7 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H000000FF&
               Height          =   240
               Left            =   4350
               RightToLeft     =   -1  'True
               TabIndex        =   275
               Top             =   390
               Visible         =   0   'False
               Width           =   945
            End
            Begin VB.TextBox TxtModFlg7 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H0000FF00&
               Enabled         =   0   'False
               Height          =   285
               Left            =   2580
               RightToLeft     =   -1  'True
               TabIndex        =   274
               Text            =   "modflag"
               Top             =   90
               Visible         =   0   'False
               Width           =   465
            End
            Begin VB.Frame Frame13 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   375
               Left            =   540
               RightToLeft     =   -1  'True
               TabIndex        =   271
               Top             =   450
               Visible         =   0   'False
               Width           =   3105
               Begin MSDataListLib.DataCombo DCUser7 
                  CausesValidation=   0   'False
                  Height          =   315
                  Left            =   -255
                  TabIndex        =   272
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
                  Index           =   36
                  Left            =   2160
                  RightToLeft     =   -1  'True
                  TabIndex        =   273
                  Top             =   45
                  Width           =   855
               End
            End
            Begin MSComctlLib.ImageList GrdImageList7 
               Left            =   5520
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
                     Picture         =   "FrmPay_Garanty_Shipment.frx":204A3
                     Key             =   "CompanyName"
                  EndProperty
                  BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmPay_Garanty_Shipment.frx":2083D
                     Key             =   "Ser"
                  EndProperty
                  BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmPay_Garanty_Shipment.frx":20BD7
                     Key             =   "Vac_Name"
                  EndProperty
                  BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmPay_Garanty_Shipment.frx":20F71
                     Key             =   "ShareCount"
                  EndProperty
                  BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmPay_Garanty_Shipment.frx":2130B
                     Key             =   "Dis_Count"
                  EndProperty
                  BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmPay_Garanty_Shipment.frx":216A5
                     Key             =   "Bouns"
                  EndProperty
                  BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmPay_Garanty_Shipment.frx":21A3F
                     Key             =   "SharesValue"
                  EndProperty
                  BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmPay_Garanty_Shipment.frx":21FD9
                     Key             =   "BuyValue"
                  EndProperty
               EndProperty
            End
            Begin ImpulseButton.ISButton btnLast7 
               Height          =   315
               Left            =   90
               TabIndex        =   276
               Top             =   30
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   556
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   ""
               BackColor       =   16777215
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":22373
               ColorButton     =   16777215
               AcclimateGrayTones=   -1  'True
               DrawFocusRectangle=   0   'False
               DisabledImageExtraction=   0
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnNext7 
               Height          =   315
               Left            =   555
               TabIndex        =   277
               Top             =   30
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   556
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   ""
               BackColor       =   16777215
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":2270D
               ColorButton     =   16777215
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnPrevious7 
               Height          =   315
               Left            =   1155
               TabIndex        =   278
               Top             =   30
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   556
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   ""
               BackColor       =   16777215
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":22AA7
               ColorButton     =   16777215
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnFirst7 
               Height          =   315
               Left            =   1620
               TabIndex        =   279
               Top             =   30
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   556
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   ""
               BackColor       =   16777215
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":22E41
               ColorButton     =   16777215
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "»Ì«‰«  «·„‰«œÌ» ··„»Ì⁄« "
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
               Index           =   37
               Left            =   6480
               RightToLeft     =   -1  'True
               TabIndex        =   280
               Top             =   90
               Width           =   3390
            End
         End
         Begin VB.Frame Frm27 
            BackColor       =   &H00E2E9E9&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   1365
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   254
            Top             =   4005
            Width           =   12855
            Begin VB.CommandButton Command1 
               Caption         =   "› Õ „·› «·„ÊŸ›Ì‰"
               Height          =   375
               Left            =   360
               RightToLeft     =   -1  'True
               TabIndex        =   259
               Top             =   960
               Visible         =   0   'False
               Width           =   2055
            End
            Begin VB.TextBox TXTCode7 
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
               Left            =   7440
               Locked          =   -1  'True
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   258
               Top             =   150
               Width           =   1170
            End
            Begin VB.TextBox TXTDiscounts7 
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
               Left            =   7425
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   257
               Tag             =   "⁄›Ê« Ì—ÃÏ   ‰”»… «·Œ’„"
               Top             =   480
               Width           =   1170
            End
            Begin VB.TextBox TxtSerial7 
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
               Left            =   8745
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   256
               Top             =   2070
               Visible         =   0   'False
               Width           =   1410
            End
            Begin VB.ComboBox Combo7 
               BackColor       =   &H80000018&
               Height          =   315
               ItemData        =   "FrmPay_Garanty_Shipment.frx":231DB
               Left            =   2280
               List            =   "FrmPay_Garanty_Shipment.frx":231EB
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   255
               Top             =   1590
               Visible         =   0   'False
               Width           =   1005
            End
            Begin MSDataListLib.DataCombo DCEmP7 
               Height          =   315
               Left            =   3390
               TabIndex        =   260
               Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· «Œ Ì«— «·„‰œÊ»"
               Top             =   150
               Width           =   2715
               _ExtentX        =   4789
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DCSalesRepGroups7 
               Height          =   315
               Left            =   120
               TabIndex        =   261
               Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·„Ã„Ê⁄Â"
               Top             =   150
               Width           =   2355
               _ExtentX        =   4154
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcBranches7 
               Height          =   315
               Left            =   3360
               TabIndex        =   262
               Top             =   480
               Width           =   2715
               _ExtentX        =   4789
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DCJob7 
               Height          =   315
               Left            =   120
               TabIndex        =   263
               Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«·  «·ÊŸÌ€…"
               Top             =   480
               Visible         =   0   'False
               Width           =   2355
               _ExtentX        =   4154
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ÊŸÌ›Â"
               Height          =   285
               Index           =   38
               Left            =   2520
               RightToLeft     =   -1  'True
               TabIndex        =   269
               Top             =   480
               Visible         =   0   'False
               Width           =   810
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·›—⁄"
               Height          =   285
               Index           =   39
               Left            =   6090
               RightToLeft     =   -1  'True
               TabIndex        =   268
               Top             =   480
               Width           =   930
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„Ã„Ê⁄Â"
               Height          =   285
               Index           =   40
               Left            =   2490
               RightToLeft     =   -1  'True
               TabIndex        =   267
               Top             =   120
               Width           =   810
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "‰”»Â «·Œ’„"
               Height          =   285
               Index           =   41
               Left            =   8700
               RightToLeft     =   -1  'True
               TabIndex        =   266
               Top             =   510
               Width           =   930
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬂÊœ «·„‰œÊ»"
               Height          =   195
               Index           =   42
               Left            =   8685
               RightToLeft     =   -1  'True
               TabIndex        =   265
               Top             =   150
               Width           =   990
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·„‰œÊ»"
               Height          =   285
               Index           =   43
               Left            =   6120
               RightToLeft     =   -1  'True
               TabIndex        =   264
               Top             =   150
               Width           =   930
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic21 
            Height          =   1020
            Left            =   120
            TabIndex        =   281
            TabStop         =   0   'False
            Top             =   5400
            Width           =   12000
            _cx             =   21167
            _cy             =   1799
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
            Begin ImpulseButton.ISButton btnNew7 
               Height          =   330
               Left            =   9135
               TabIndex        =   282
               Top             =   555
               Width           =   630
               _ExtentX        =   1111
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":23204
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnSave7 
               Height          =   330
               Left            =   7710
               TabIndex        =   283
               Top             =   555
               Width           =   630
               _ExtentX        =   1111
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":2359E
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnModify7 
               Height          =   330
               Left            =   8475
               TabIndex        =   284
               Top             =   555
               Width           =   630
               _ExtentX        =   1111
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":23938
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton BtnUndo7 
               Height          =   330
               Left            =   6945
               TabIndex        =   285
               Top             =   555
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":23CD2
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnDelete7 
               Height          =   330
               Left            =   5460
               TabIndex        =   286
               Top             =   555
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":2406C
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnQuery7 
               Height          =   330
               Left            =   6120
               TabIndex        =   287
               TabStop         =   0   'False
               ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
               Top             =   570
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":24606
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton BtnUpdate7 
               Height          =   330
               Left            =   6765
               TabIndex        =   288
               TabStop         =   0   'False
               ToolTipText     =   " ÕœÌÀ ﬁ«⁄œ… «·»Ì«‰« "
               Top             =   105
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":249A0
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton BtnPrint7 
               Height          =   285
               Left            =   765
               TabIndex        =   289
               TabStop         =   0   'False
               Top             =   630
               Width           =   765
               _ExtentX        =   1349
               _ExtentY        =   503
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":24D3A
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnCancel7 
               Height          =   330
               Left            =   -15
               TabIndex        =   290
               Top             =   555
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":250D4
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton ISButton17 
               Height          =   405
               Left            =   4200
               TabIndex        =   291
               TabStop         =   0   'False
               Top             =   600
               Width           =   1245
               _ExtentX        =   2196
               _ExtentY        =   714
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   "ÿ»«⁄…  »Õ”» «·„Ã„Ê⁄…"
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":2546E
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton ISButton27 
               Height          =   405
               Left            =   2880
               TabIndex        =   292
               TabStop         =   0   'False
               Top             =   600
               Width           =   1245
               _ExtentX        =   2196
               _ExtentY        =   714
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   "ÿ»«⁄… »Õ”» «·›—⁄"
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":25808
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton ISButton107 
               Height          =   405
               Left            =   1560
               TabIndex        =   293
               TabStop         =   0   'False
               Top             =   525
               Width           =   1245
               _ExtentX        =   2196
               _ExtentY        =   714
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   "ÿ»«⁄… ⁄„·«¡ «·„‰«œÌ»"
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
               ButtonImage     =   "FrmPay_Garanty_Shipment.frx":25BA2
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·”Ã· «·Õ«·Ì:"
               Height          =   210
               Index           =   14
               Left            =   2505
               RightToLeft     =   -1  'True
               TabIndex        =   0
               Top             =   225
               Width           =   975
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄œœ «·”Ã·« :"
               Height          =   210
               Index           =   15
               Left            =   810
               RightToLeft     =   -1  'True
               TabIndex        =   1
               Top             =   225
               Width           =   975
            End
            Begin VB.Label LabCurrRec7 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Height          =   210
               Left            =   1800
               RightToLeft     =   -1  'True
               TabIndex        =   2
               Top             =   240
               Width           =   675
            End
            Begin VB.Label LabCountRec7 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Height          =   210
               Left            =   240
               RightToLeft     =   -1  'True
               TabIndex        =   294
               Top             =   225
               Width           =   540
            End
         End
         Begin VSFlex8UCtl.VSFlexGrid Grid7 
            Height          =   3405
            Left            =   0
            TabIndex        =   3
            Top             =   600
            Width           =   9915
            _cx             =   17489
            _cy             =   6006
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
            Cols            =   8
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmPay_Garanty_Shipment.frx":25F3C
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
      End
   End
End
Attribute VB_Name = "FrmPay_Garanty_Shipment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public SendForm As Integer
Dim RsSavRec As ADODB.Recordset
Dim RecId As String
Dim II As Long
'#################################################################################################
Dim RsSavRec1 As ADODB.Recordset
Dim BKGrndPic As ClsBackGroundPic
Dim RecID1 As String
Dim II1 As Long
'##################################################################################################
Dim RsSavRec2 As ADODB.Recordset
Dim BKGrndPic2 As ClsBackGroundPic
Dim RecId2 As String
'Dim II As Long
'##################################################################################################
Dim RsSavRec3 As ADODB.Recordset
Dim BKGrndPic3 As ClsBackGroundPic
Dim RecId3 As String
'Dim II3 As Long
'##################################################################################################
Dim RsSavRec4 As ADODB.Recordset
Dim RecId4 As String
'Dim II4 As Long
'##################################################################################################
Dim RsSavRec5 As ADODB.Recordset
Dim BKGrndPic5 As ClsBackGroundPic
Dim RecId5 As String
'Dim II5 As Long
'##################################################################################################
Dim RsSavRec6 As ADODB.Recordset
Dim BKGrndPic6 As ClsBackGroundPic
Dim RecId6 As String
'Dim II6 As Long
Dim cSearch  As clsDCboSearch
Public chPrinet As Integer
'##################################################################################################
Dim RsSavRec7 As ADODB.Recordset
Dim BKGrndPic7 As ClsBackGroundPic
Dim RecId7 As String
'Dim II7 As Long
Dim cSearch7  As clsDCboSearch
Public chPrinet7 As Integer
Private Sub ChangeLang()
'#####################################################################################################################################################
    Dim XPic As IPictureDisp
    Set XPic = Me.btnFirst.ButtonImage
    Set Me.btnFirst.ButtonImage = Me.btnLast.ButtonImage
    Set Me.btnLast.ButtonImage = XPic
    Set XPic = Me.btnPrevious.ButtonImage
    Set Me.btnPrevious.ButtonImage = Me.btnNext.ButtonImage
    Set Me.btnNext.ButtonImage = XPic
    
    Label1(3).Caption = "Code"
    Label1(0).Caption = "Name Ar"
    Label1(1).Caption = "Name Eng"
    btnQuery.Caption = "Search"
 
    With Grid
        .TextMatrix(0, .ColIndex("UnitID")) = " Code"
        .TextMatrix(0, .ColIndex("UnitName")) = " Name AR"
        .TextMatrix(0, .ColIndex("UnitNameE")) = " Name Eng"
    End With

    Label1(12).Caption = "Payment Methods"
    
    btnNew.Caption = "New"
    btnModify.Caption = "Modify"
    btnSave.Caption = "Save"
    BtnUndo.Caption = "Undo"
    btnDelete.Caption = "Delete"
    btnCancel.Caption = "Exit"
    Label2(0).Caption = "Current Record"
    Label2(1).Caption = "NO Of Record"
'#####################################################################################################################################################
    Dim XPic1 As IPictureDisp
    Set XPic1 = Me.btnFirst1.ButtonImage
    Set Me.btnFirst1.ButtonImage = Me.btnLast1.ButtonImage
    Set Me.btnLast1.ButtonImage = XPic1
    Set XPic1 = Me.btnPrevious1.ButtonImage
    Set Me.btnPrevious1.ButtonImage = Me.btnNext1.ButtonImage
    Set Me.btnNext1.ButtonImage = XPic1

    Label1(2).Caption = "Guranty types"

    With Me.Grid1
        .TextMatrix(0, .ColIndex("Ser")) = "Ser"
        .TextMatrix(0, .ColIndex("id")) = "Code"
        .TextMatrix(0, .ColIndex("name")) = "Arabic Name"
        .TextMatrix(0, .ColIndex("namee")) = "English Name"
    End With

    Label1(6).Caption = "Code"
    Label1(4).Caption = "Name Eng"
    Label1(5).Caption = "Name Ar"

    Label2(2).Caption = "Curr. Rec."
    Label2(3).Caption = "Rec. Count."

    btnNew1.Caption = "New"
    btnModify1.Caption = "Modify"
    btnSave1.Caption = "Save"
    BtnUndo1.Caption = "Undo"
    btnDelete1.Caption = "Delete"
    btnCancel1.Caption = "Exit"
'###################################################################################################################################################
    Dim XPic2 As IPictureDisp
    Set XPic2 = Me.btnFirst2.ButtonImage
    Set Me.btnFirst2.ButtonImage = Me.btnLast2.ButtonImage
    Set Me.btnLast2.ButtonImage = XPic2
    Set XPic2 = Me.btnPrevious2.ButtonImage
    Set Me.btnPrevious2.ButtonImage = Me.btnNext2.ButtonImage
    Set Me.btnNext2.ButtonImage = XPic2

    Label1(8).Caption = "Shipment Mode"

    With Me.GRID2
        .TextMatrix(0, .ColIndex("Ser")) = "Ser"
        .TextMatrix(0, .ColIndex("id")) = "Code"
        .TextMatrix(0, .ColIndex("name")) = "Arabic Name"
        .TextMatrix(0, .ColIndex("namee")) = "English Name"
    End With

    Label1(11).Caption = "Code"
    Label1(10).Caption = "Name Ar"
    Label1(9).Caption = "Name Eng"

    Label2(4).Caption = "Curr. Rec."
    Label2(5).Caption = "Rec. Count."

    btnNew2.Caption = "New"
    btnModify2.Caption = "Modify"
    btnSave2.Caption = "Save"
    BtnUndo2.Caption = "Undo"
    btnDelete2.Caption = "Delete"
    btnCancel2.Caption = "Exit"
'################################################################################################################################################
    Dim XPic3 As IPictureDisp
    Set XPic3 = Me.btnFirst3.ButtonImage
    Set Me.btnFirst3.ButtonImage = Me.btnLast3.ButtonImage
    Set Me.btnLast3.ButtonImage = XPic3
    Set XPic3 = Me.btnPrevious3.ButtonImage
    Set Me.btnPrevious3.ButtonImage = Me.btnNext3.ButtonImage
    Set Me.btnNext3.ButtonImage = XPic3
    
    Label1(18).Caption = "Purchae Rep Groups"
    
    With Me.Grid3
        .TextMatrix(0, .ColIndex("id")) = "Code"
        .TextMatrix(0, .ColIndex("name")) = "Arabic Name"
        .TextMatrix(0, .ColIndex("namee")) = "English Name"
    End With
    
    Label1(16).Caption = "Code"
    Label1(15).Caption = "Name AR"
    Label1(14).Caption = "Name Eng"
    
    Label2(7).Caption = "Curr. Rec."
    Label2(6).Caption = "Rec. Count."
    
    btnNew3.Caption = "New"
    btnModify3.Caption = "Modify"
    btnSave3.Caption = "Save"
    BtnUndo3.Caption = "Undo"
    btnDelete3.Caption = "Delete"
    btnCancel3.Caption = "Exit"
'###################################################################################################################################################
    Dim XPic4 As IPictureDisp
    Set XPic4 = Me.btnFirst4.ButtonImage
    Set Me.btnFirst4.ButtonImage = Me.btnLast4.ButtonImage
    Set Me.btnLast4.ButtonImage = XPic4
    Set XPic4 = Me.btnPrevious4.ButtonImage
    Set Me.btnPrevious4.ButtonImage = Me.btnNext4.ButtonImage
    Set Me.btnNext4.ButtonImage = XPic4
    
    Label1(22).Caption = "Shipping Methods"
    
    Label1(20).Caption = "Code"
    Label1(21).Caption = "Name Ar"
    Label1(19).Caption = "Name Eng"
    btnQuery.Caption = "Search"

    With Grid4
        .TextMatrix(0, .ColIndex("UnitID")) = " Code"
        .TextMatrix(0, .ColIndex("UnitName")) = " Name AR"
        .TextMatrix(0, .ColIndex("UnitNameE")) = " Name Eng"
    End With
    
    btnNew4.Caption = "New"
    btnModify4.Caption = "Modify"
    btnSave4.Caption = "Save"
    BtnUndo4.Caption = "Undo"
    btnDelete4.Caption = "Delete"
    btnCancel4.Caption = "Exit"
    
    Label2(9).Caption = "Current Record"
    Label2(8).Caption = "NO Of Record"
'##################################################################################################################################################
    Dim XPic5 As IPictureDisp
    Set XPic5 = Me.btnFirst5.ButtonImage
    Set Me.btnFirst5.ButtonImage = Me.btnLast5.ButtonImage
    Set Me.btnLast5.ButtonImage = XPic5
    Set XPic5 = Me.btnPrevious5.ButtonImage
    Set Me.btnPrevious5.ButtonImage = Me.btnNext5.ButtonImage
    Set Me.btnNext5.ButtonImage = XPic5


    Me.Label1(24).Caption = "Guarantee for Installments Types"
    
    Label1(27).Caption = "Code"
    Label1(26).Caption = "Name AR"
    Label1(25).Caption = "Name ENG"

    Label2(10).Caption = "Current Record"
    Label2(11).Caption = "NO. Recordes"

    btnNew5.Caption = "New"
    btnModify5.Caption = "Modify"
    btnSave5.Caption = "Save"
    BtnUndo5.Caption = "Undo"
    btnDelete5.Caption = "Delete"
    btnCancel5.Caption = "Exit"

    With Me.Grid5
        .TextMatrix(0, .ColIndex("Ser")) = "Ser"
        .TextMatrix(0, .ColIndex("id")) = "Id"
        .TextMatrix(0, .ColIndex("name")) = "Name AR"
        .TextMatrix(0, .ColIndex("namee")) = "Name ENG"
    End With
'####################################################################################################################################################
    Dim XPic6 As IPictureDisp
    Set XPic6 = Me.btnFirst6.ButtonImage
    Set Me.btnFirst6.ButtonImage = Me.btnLast6.ButtonImage
    Set Me.btnLast6.ButtonImage = XPic6
    Set XPic6 = Me.btnPrevious6.ButtonImage
    Set Me.btnPrevious6.ButtonImage = Me.btnNext6.ButtonImage
    Set Me.btnNext6.ButtonImage = XPic6
    
    ISButton1.Caption = "Prient"
    ISButton2.Caption = "Prient By Group"
    ISButton3.Caption = "Prient By Branch"
    
    btnQuery6.Caption = "Search"

    Me.Label1(29).Caption = "Purchae Rep Data"
    
    Label1(34).Caption = "Code"
    Label1(35).Caption = "Name"
    Label1(30).Caption = "Job"
    Label1(33).Caption = "Discount%"
    Label1(31).Caption = "Branch"
    Label1(32).Caption = "Group"

    Label2(12).Caption = "Current Record"
    Label2(13).Caption = "NO. Recordes"

    btnNew6.Caption = "New"
    btnModify6.Caption = "Modify"
    btnSave6.Caption = "Save"
    BtnUndo6.Caption = "Undo"
    btnDelete6.Caption = "Delete"
    btnCancel6.Caption = "Exit"
 
    With Me.Grid6
        .TextMatrix(0, .ColIndex("ser")) = "Ser"
        .TextMatrix(0, .ColIndex("EmpCode")) = "Code"
        .TextMatrix(0, .ColIndex("EmpName")) = "Name"
        .TextMatrix(0, .ColIndex("JobID")) = "Job"
        .TextMatrix(0, .ColIndex("groupid")) = "Group"
        .TextMatrix(0, .ColIndex("BranchId")) = "Branch"
        .TextMatrix(0, .ColIndex("discountvalue")) = "Discount%"
    End With
'#############################################################################################################################################################################
    Dim XPic7 As IPictureDisp
    Set XPic7 = Me.btnFirst7.ButtonImage
    Set Me.btnFirst7.ButtonImage = Me.btnLast7.ButtonImage
    Set Me.btnLast7.ButtonImage = XPic7
    Set XPic7 = Me.btnPrevious7.ButtonImage
    Set Me.btnPrevious7.ButtonImage = Me.btnNext7.ButtonImage
    Set Me.btnNext7.ButtonImage = XPic7

    BtnPrint7.Caption = "Prient"
    ISButton17.Caption = "Prient By Group"
    ISButton27.Caption = "Prient By Branch"
    btnQuery7.Caption = "Search"
    
    Me.Label1(37).Caption = "Sales Rep Data"
    
    Label1(42).Caption = "Code"
    Label1(43).Caption = "Name"
    ISButton107.Caption = "Print Customer "
    Label1(38).Caption = "Job"

    Label1(41).Caption = "Discount%"
    Label1(39).Caption = "Branch"
    Label1(40).Caption = "Group"

    Label2(14).Caption = "Current Record"
    Label2(15).Caption = "NO. Recordes"

    btnNew7.Caption = "New"
    btnModify7.Caption = "Modify"
    btnSave7.Caption = "Save"
    BtnUndo7.Caption = "Undo"
    btnDelete7.Caption = "Delete"
    btnCancel7.Caption = "Exit"
    Command1.Caption = "Open Employee File"

    With Me.Grid7
        .TextMatrix(0, .ColIndex("ser")) = "Ser"
        .TextMatrix(0, .ColIndex("EmpCode")) = "Code"
        .TextMatrix(0, .ColIndex("EmpName")) = "Name"
        .TextMatrix(0, .ColIndex("JobID")) = "Job"
        .TextMatrix(0, .ColIndex("groupid")) = "Group"
        .TextMatrix(0, .ColIndex("BranchId")) = "Branch"
        .TextMatrix(0, .ColIndex("discountvalue")) = "Discount%"
    End With
End Sub

Private Sub Form_Load()

    On Error GoTo ErrTrap
    
    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
'###################################################################################################################################################
    If SendForm = 0 Then
        ScreenNameArabic = "  «·ÊÕœ«  «·„” Œœ„… ›Ï «·»—‰«„Ã "
        ScreenNameEnglish = " Units Data  "
        RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish, "1"
        Dim cGrdBack As New ClsBackGroundPic
        Set Me.Grid.WallPaper = cGrdBack.Picture
        Dim i As Integer
        Dim My_SQL As String
        My_SQL = "TblPaymetData"
        Set RsSavRec = New ADODB.Recordset
        RsSavRec.CursorLocation = adUseClient
        RsSavRec.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdTableDirect
        Me.TxtModFlg.Text = "R"
        FillGridWithData
        With Me.Grid
            For i = 0 To .Cols - 1
                .Cell(flexcpPictureAlignment, 0, i) = flexPicAlignRightCenter
            Next
            .ExtendLastCol = True
        End With
        BtnFirst_Click
'#############################################################################################################################################
    ElseIf SendForm = 1 Then
        ScreenNameArabic = " «‰Ê«⁄ «·÷„«‰«   "
        ScreenNameEnglish = " Gurantee Types "
        RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish, "1"
        My_SQL = "gurantee_type"
        Set BKGrndPic = New ClsBackGroundPic
        Set RsSavRec1 = New ADODB.Recordset
        RsSavRec1.CursorLocation = adUseClient
        RsSavRec1.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdTableDirect
        Me.TxtModFlg1.Text = "R"
        Resize_Form Me
        My_SQL = "select UserID,UserName From tblUsers "
        fill_combo DCUser, My_SQL
        FillWithData
        With Me.Grid1
            .Cell(flexcpPicture, 0, .ColIndex("name")) = Me.GrdImageList.ListImages("Vac_Name").ExtractIcon
            .Cell(flexcpPicture, 0, .ColIndex("Ser")) = Me.GrdImageList.ListImages("Ser").ExtractIcon
            For i = 0 To .Cols - 1
                .Cell(flexcpPictureAlignment, 0, i) = flexPicAlignRightCenter
            Next
            .ExtendLastCol = True
            .WallPaper = BKGrndPic.Picture
            .RowHeight(-1) = 300
        End With
        btnFirst1_Click
'#########################################################################################################################################################
    ElseIf SendForm = 2 Then
        ScreenNameArabic = " «‰Ê«⁄  «·‘Õ‰  "
        ScreenNameEnglish = "  Shipment Types "
        RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish, "1"
        My_SQL = "Shipment_mode"
        Set BKGrndPic2 = New ClsBackGroundPic
        Set RsSavRec2 = New ADODB.Recordset
        RsSavRec2.CursorLocation = adUseClient
        RsSavRec2.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdTableDirect
        Me.TxtModFlg2.Text = "R"
        Resize_Form Me
        My_SQL = "select UserID,UserName From tblUsers "
        fill_combo DCUser, My_SQL
        FillGridWithData2
        With Me.GRID2
            .Cell(flexcpPicture, 0, .ColIndex("name")) = Me.GrdImageList.ListImages("Vac_Name").ExtractIcon
            .Cell(flexcpPicture, 0, .ColIndex("Ser")) = Me.GrdImageList.ListImages("Ser").ExtractIcon
            For i = 0 To .Cols - 1
                .Cell(flexcpPictureAlignment, 0, i) = flexPicAlignRightCenter
            Next
            .ExtendLastCol = True
            .WallPaper = BKGrndPic2.Picture
            .RowHeight(-1) = 300
        End With
        btnFirst2_Click
'###########################################################################################################################
    ElseIf SendForm = 3 Then
        ScreenNameArabic = " „Ã„Ê⁄«  «·„‰«œÌ» ··„‘ —Ì«  "
        ScreenNameEnglish = " Sales Person Groups "
        RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish, "1"
        My_SQL = "TBLSalesRepGroups2"
        Set BKGrndPic3 = New ClsBackGroundPic
        Set RsSavRec3 = New ADODB.Recordset
        RsSavRec3.CursorLocation = adUseClient
        RsSavRec3.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdTableDirect
        Me.TxtModFlg3.Text = "R"
        Resize_Form Me
        My_SQL = "select UserID,UserName From tblUsers "
        fill_combo DCUser3, My_SQL
        FillGrid3WithData
        With Me.Grid3
            .Cell(flexcpPicture, 0, .ColIndex("name")) = Me.GrdImageList.ListImages("Vac_Name").ExtractIcon
            .Cell(flexcpPicture, 0, .ColIndex("Ser")) = Me.GrdImageList.ListImages("Ser").ExtractIcon
            For i = 0 To .Cols - 1
                .Cell(flexcpPictureAlignment, 0, i) = flexPicAlignRightCenter
            Next
            .ExtendLastCol = True
            .WallPaper = BKGrndPic3.Picture
            .RowHeight(-1) = 300
        End With
        btnFirst3_Click
'###########################################################################################################################
    ElseIf SendForm = 4 Then
        ScreenNameArabic = "  «·ÊÕœ«  «·„” Œœ„… ›Ï «·»—‰«„Ã "
        ScreenNameEnglish = " Units Data  "
        RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish, "1"
        My_SQL = "TblShipingData"
        Set RsSavRec4 = New ADODB.Recordset
        RsSavRec4.CursorLocation = adUseClient
        RsSavRec4.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdTableDirect
        Me.TxtModFlg4.Text = "R"
        FillGrid4WithData
        With Me.Grid4
            For i = 0 To .Cols - 1
                .Cell(flexcpPictureAlignment, 0, i) = flexPicAlignRightCenter
            Next
            .ExtendLastCol = True
        End With
        btnFirst4_Click
'###########################################################################################################################
    ElseIf SendForm = 5 Then
         My_SQL = "Gbasic"
        Set BKGrndPic5 = New ClsBackGroundPic
        Set RsSavRec5 = New ADODB.Recordset
        RsSavRec5.CursorLocation = adUseClient
        RsSavRec5.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdTableDirect
        Me.TxtModFlg5.Text = "R"
        Resize_Form Me

        My_SQL = "select UserID,UserName From tblUsers "
        fill_combo DCUser5, My_SQL

        FillGrid5WithData

        With Me.Grid5
            .Cell(flexcpPicture, 0, .ColIndex("name")) = Me.GrdImageList.ListImages("Vac_Name").ExtractIcon
            .Cell(flexcpPicture, 0, .ColIndex("Ser")) = Me.GrdImageList.ListImages("Ser").ExtractIcon
            For i = 0 To .Cols - 1
                .Cell(flexcpPictureAlignment, 0, i) = flexPicAlignRightCenter
            Next
            .ExtendLastCol = True
            .WallPaper = BKGrndPic5.Picture
            .RowHeight(-1) = 300
        End With
        btnFirst5_Click
'###########################################################################################################################
    ElseIf SendForm = 6 Then
        Dim Dcombos As ClsDataCombos
    
        My_SQL = "TBLSalesRepData2"
        Set BKGrndPic6 = New ClsBackGroundPic
        Set RsSavRec6 = New ADODB.Recordset
        RsSavRec6.CursorLocation = adUseClient
    
        RsSavRec6.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdTableDirect
        Me.TxtModFlg6.Text = "R"
        Resize_Form Me

        My_SQL = "select UserID,UserName From tblUsers "
        fill_combo DCUser, My_SQL
        Set Dcombos = New ClsDataCombos
        Dcombos.GetBranches Me.DcBranches
        Dcombos.GetEmployees Me.DcEmp
        Dcombos.GetEmpJobsTypes Me.DCJob
        Dcombos.GetSalesRepGroupsPurchase Me.DCSalesRepGroups

        Set cSearch = New clsDCboSearch
        Set cSearch.Client = Me.DcEmp
        Set cSearch.Client = Me.DcBranches
        Set cSearch.Client = Me.DCSalesRepGroups
        Set cSearch.Client = Me.DCJob

        ModFgLib.LinkFgColWithDataCombo Grid6, Grid6.ColIndex("EmpName"), Me.DcEmp
        ModFgLib.LinkFgColWithDataCombo Grid6, Grid6.ColIndex("BranchId"), Me.DcBranches
        ModFgLib.LinkFgColWithDataCombo Grid6, Grid6.ColIndex("GroupID"), Me.DCSalesRepGroups
        ModFgLib.LinkFgColWithDataCombo Grid6, Grid6.ColIndex("JobID"), Me.DCJob

        FillGrid6WithData
        With Me.Grid6
            .Cell(flexcpPicture, 0, .ColIndex("DiscountValue")) = Me.GrdImageList.ListImages("Vac_Name").ExtractIcon
            .Cell(flexcpPicture, 0, .ColIndex("Ser")) = Me.GrdImageList.ListImages("Ser").ExtractIcon

            For i = 0 To .Cols - 1
                .Cell(flexcpPictureAlignment, 0, i) = flexPicAlignRightCenter
            Next
            .ExtendLastCol = True
            .WallPaper = BKGrndPic6.Picture
            .RowHeight(-1) = 300
        End With
        btnFirst6_Click
'###########################################################################################################################
    ElseIf SendForm = 7 Then
    My_SQL = "TBLSalesRepData"
    Set BKGrndPic7 = New ClsBackGroundPic
    Set RsSavRec7 = New ADODB.Recordset
    RsSavRec7.CursorLocation = adUseClient
    RsSavRec7.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdTableDirect
    Me.TxtModFlg7.Text = "R"
    Resize_Form Me
    My_SQL = "select UserID,UserName From tblUsers "
    fill_combo DCUser7, My_SQL
    
    Set Dcombos = New ClsDataCombos
    Dcombos.GetBranches Me.DcBranches7
    Dcombos.GetEmployees Me.DCEmP7
    Dcombos.GetEmpJobsTypes Me.DCJob7
    Dcombos.GetSalesRepGroups Me.DCSalesRepGroups7

    Set cSearch7 = New clsDCboSearch
    Set cSearch7.Client = Me.DCEmP7
    Set cSearch7.Client = Me.DcBranches7
    Set cSearch7.Client = Me.DCSalesRepGroups7
    Set cSearch7.Client = Me.DCJob7

    ModFgLib.LinkFgColWithDataCombo Grid7, Grid7.ColIndex("EmpName"), Me.DCEmP7
    ModFgLib.LinkFgColWithDataCombo Grid7, Grid7.ColIndex("BranchId"), Me.DcBranches7
    ModFgLib.LinkFgColWithDataCombo Grid7, Grid7.ColIndex("GroupID"), Me.DCSalesRepGroups7
    ModFgLib.LinkFgColWithDataCombo Grid7, Grid7.ColIndex("JobID"), Me.DCJob7

    FillGrid7WithData

    With Me.Grid7
        .Cell(flexcpPicture, 0, .ColIndex("DiscountValue")) = Me.GrdImageList7.ListImages("Vac_Name").ExtractIcon
        .Cell(flexcpPicture, 0, .ColIndex("Ser")) = Me.GrdImageList7.ListImages("Ser").ExtractIcon

        For i = 0 To .Cols - 1
            .Cell(flexcpPictureAlignment, 0, i) = flexPicAlignRightCenter
        Next
   
        .ExtendLastCol = True
        .WallPaper = BKGrndPic7.Picture
        .RowHeight(-1) = 300
    End With

    btnFirst7_Click
'###########################################################################################################################
    End If
'###########################################################################################################################
    C1Tab1.TabVisible(SendForm) = True
    C1Tab1.CurrTab = SendForm
ErrTrap:
End Sub
Private Sub Form_Unload(Cancel As Integer)

    On Error GoTo ErrTrap
    
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish
    If RsSavRec.State = adStateOpen Then
        If Not (RsSavRec.EOF Or RsSavRec.BOF) Then
            If RsSavRec.EditMode <> adEditNone Then
                RsSavRec.CancelUpdate
            End If
        End If
        RsSavRec.Close
        Set RsSavRec = Nothing
    End If
    Exit Sub
    '################################################################################################
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish
    If RsSavRec1.State = adStateOpen Then
        If Not (RsSavRec1.EOF Or RsSavRec1.BOF) Then
            If RsSavRec1.EditMode <> adEditNone Then
                RsSavRec1.CancelUpdate
            End If
        End If
        RsSavRec1.Close
        Set RsSavRec1 = Nothing
    End If
    '################################################################################################
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish
    If RsSavRec2.State = adStateOpen Then
        If Not (RsSavRec2.EOF Or RsSavRec2.BOF) Then
            If RsSavRec2.EditMode <> adEditNone Then
                RsSavRec2.CancelUpdate
            End If
        End If
        RsSavRec2.Close
        Set RsSavRec2 = Nothing
    End If
    '#################################################################################################
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish
    If RsSavRec3.State = adStateOpen Then
        If Not (RsSavRec3.EOF Or RsSavRec3.BOF) Then
            If RsSavRec3.EditMode <> adEditNone Then
                RsSavRec3.CancelUpdate
            End If
        End If
        RsSavRec3.Close
        Set RsSavRec3 = Nothing
    End If
    '#################################################################################################
        If RsSavRec4.State = adStateOpen Then
        If Not (RsSavRec4.EOF Or RsSavRec4.BOF) Then
            If RsSavRec4.EditMode <> adEditNone Then
                RsSavRec4.CancelUpdate
            End If
        End If
        RsSavRec4.Close
        Set RsSavRec4 = Nothing
    End If
    '#################################################################################################
        If RsSavRec5.State = adStateOpen Then
        If Not (RsSavRec5.EOF Or RsSavRec5.BOF) Then
            If RsSavRec5.EditMode <> adEditNone Then
                RsSavRec5.CancelUpdate
            End If
        End If
        RsSavRec5.Close
        Set RsSavRec5 = Nothing
    End If
    '#################################################################################################
    If RsSavRec6.State = adStateOpen Then
        If Not (RsSavRec6.EOF Or RsSavRec6.BOF) Then
            If RsSavRec6.EditMode <> adEditNone Then
                RsSavRec6.CancelUpdate
            End If
        End If
        RsSavRec6.Close
        Set RsSavRec6 = Nothing
    End If
    Set cSearch = Nothing
    '##################################################################################################
    If RsSavRec7.State = adStateOpen Then
        If Not (RsSavRec7.EOF Or RsSavRec7.BOF) Then
            If RsSavRec7.EditMode <> adEditNone Then
                RsSavRec7.CancelUpdate
            End If
        End If
        RsSavRec7.Close
        Set RsSavRec7 = Nothing
    End If
    Set cSearch7 = Nothing
    '###################################################################################################
ErrTrap:
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim IntResult As String
    
    Dim StrMSG As String
    
    On Error GoTo ErrTrap
    
    If SendForm = 0 Then
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
        End If
    ElseIf SendForm = 1 Then
        If Me.TxtModFlg1.Text <> "R" Then
            Select Case Me.TxtModFlg1.Text
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
        End If
   ElseIf SendForm = 2 Then
        If Me.TxtModFlg2.Text <> "R" Then
            Select Case Me.TxtModFlg2.Text
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
        End If
    ElseIf SendForm = 3 Then
        If Me.TxtModFlg3.Text <> "R" Then
            Select Case Me.TxtModFlg3.Text
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
        End If
    ElseIf SendForm = 4 Then
        If Me.TxtModFlg4.Text <> "R" Then
            Select Case Me.TxtModFlg4.Text
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
        End If
    ElseIf SendForm = 5 Then
        If Me.TxtModFlg5.Text <> "R" Then
            Select Case Me.TxtModFlg5.Text
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
        End If
    ElseIf SendForm = 6 Then
        If Me.TxtModFlg6.Text <> "R" Then
            Select Case Me.TxtModFlg6.Text
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
        End If
    ElseIf SendForm = 7 Then
        If Me.TxtModFlg7.Text <> "R" Then
            Select Case Me.TxtModFlg7.Text
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
        End If
    End If
    If StrMSG <> "" Then
        IntResult = MsgBox(StrMSG, vbMsgBoxRight + vbYesNoCancel + vbMsgBoxRtlReading + vbQuestion, App.title)
        Select Case IntResult
            Case vbYes
                Cancel = True
                Select Case SendForm
                    Case 0
                        btnSave_Click
                    Case 1
                        btnSave1_Click
                    Case 2
                        btnSave2_Click
                    Case 3
                        btnSave3_Click
                    Case 4
                        btnSave4_Click
                    Case 5
                        btnSave5_Click
                    Case 6
                        btnSave6_Click
                    Case 7
                        btnSave7_Click
                End Select
            Case vbCancel
                Cancel = True
        End Select
    End If
    Exit Sub
ErrTrap:
End Sub
'##############################################################################################################################################################
'##############################################################################################################################################################
'##############################################################################################################################################################
Private Sub BtnCancel_Click()
    Unload Me
End Sub
Private Sub btnDelete_Click()

    On Error GoTo ErrTrap

    Dim MSGType As Integer
    Dim BlnRecExist As Boolean
    Dim StrMSG  As String
    Dim Msg As String

    If TxtUnitID.Text <> "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MSGType = MsgBox("Â·  —€» ›Ì Õ–› Â–« «·”Ã·", vbYesNo + vbMsgBoxRight, App.title)
        Else
            MSGType = MsgBox("Delete This Record", vbYesNo + vbMsgBoxRight, App.title)
        End If
        If MSGType = vbYes Then
            RsSavRec.Find "id=" & val(TxtUnitID.Text), , adSearchForward, 1
            If Not (RsSavRec.BOF Or RsSavRec.EOF) Then
                CuurentLogdata ("D")
                RsSavRec.delete
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox " „  ⁄„·Ì… «·Õ–› »‰Ã«Õ", vbOKOnly + vbMsgBoxRight, App.title
                Else
                    MsgBox "Delete Success...", vbOKOnly + vbMsgBoxRight, App.title
                End If
                FillGridWithData
                BtnNext_Click
            End If
        End If
    
    End If

    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217873, -2147467259
            If SystemOptions.UserInterface = ArabicInterface Then
                StrMSG = "⁄›Ê« ·« ÌÃÊ“ Õ–› «·”Ã· ·«— »«ÿÂ »»Ì«‰«  √Œ—Ì"
            Else
                StrMSG = "Sorry .. can't Delete this record , Reason : Data integrity"
            End If
            RsSavRec.CancelUpdate
            MsgBox StrMSG, vbOKOnly + vbMsgBoxRight, App.title
            Cn.Errors.Clear
    End Select

End Sub
Private Sub BtnFirst_Click()

    On Error GoTo ErrTrap
    
    Dim Msg As String
    If Me.TxtModFlg.Text = "N" Then
        FindRec val(Me.TxtUnitID.Text)
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
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
                Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
                Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
                Msg = "Sorry.. this record Already Deleted" & CHR(13)
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
        FindRec val(Me.TxtUnitID.Text)
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
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
                Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
                Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
                Msg = "Sorry.. this record Already Deleted" & CHR(13)
            End If

            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub btnModify_Click()
    On Error GoTo ErrTrap

    Dim Msg As String
    If TxtUnitID.Text <> "" Then
        TxtModFlg = "E"
        Frm2.Enabled = True
        Me.TxtVacName.SetFocus
        CuurentLogdata
    End If
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147467259
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "⁄›Ê«" & CHR(13)
                Msg = Msg & " ·«Ì„ﬂ‰  ⁄œÌ· Â–« «·”Ã· ›Ï «·Êﬁ  «·Õ«·Ï" & CHR(13)
                Msg = Msg & "ÕÌÀ «‰Â ﬁÌœ «· ⁄œÌ· „‰ ﬁ»· „” Œœ„ «Œ— ⁄·Ï «·‘»ﬂ…"
            Else
                Msg = "Sorry" & CHR(13)
                Msg = Msg & " Can't Edit this record now - Another user work with it now" & CHR(13)
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            If RsSavRec.EditMode <> adEditNone Then
                RsSavRec.CancelUpdate
            End If
    End Select

End Sub
Private Sub btnNew_Click()

    On Error GoTo ErrTrap
    
    Dim My_SQL As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Frm2.Enabled = True
    clear_all Me
    TxtModFlg.Text = "N"

    My_SQL = "TblPaymetData"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable

    If rs.RecordCount > 0 Then
        TxtUnitID.Text = rs.RecordCount + 1
    Else
        TxtUnitID.Text = 1
    End If
    rs.Close
    CmbType.ListIndex = 0
    TxtVacName.SetFocus
ErrTrap:
End Sub
Private Sub BtnNext_Click()
    On Error GoTo ErrTrap
    Dim Msg As String

    If Me.TxtModFlg.Text = "N" Then
        FindRec val(Me.TxtUnitID.Text)
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
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
                Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
                Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
                Msg = "Sorry , this Recored was deleted by other user on the network" & CHR(13)
                Msg = Msg & "Date will be updated now" & CHR(13)
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
        FindRec val(Me.TxtUnitID.Text)
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
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
                Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
                Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
                Msg = "Sorry , this Recored was deleted by other user on the network" & CHR(13)
                Msg = Msg & "Date will be updated now" & CHR(13)
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select

End Sub
Private Sub btnSave_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
    Dim StrVacCode As String
    Dim StrVacName As String
    Dim CtrlTxt As Control

    If Trim(Me.TxtVacName.Text) = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "ÌÃ» ﬂ «»… «”„ «·‰Ê⁄ ...!!!"
        Else
            Msg = "Please Enter The name"
        End If
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        TxtVacName.SetFocus
        Exit Sub
    End If

StrVacName = ""
    StrVacName = IsRecExist("TblPaymetData", "name", Trim(TxtVacName.Text), "name", "ID<>'" & Trim(TxtUnitID.Text) & "'")

    If StrVacName <> "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·ﬁœ ”»ﬁ  ”ÃÌ· Â–Â «·‰Ê⁄ „‰ ﬁ»·"
        Else
            Msg = "this Unit Already Exist"
        End If

        MsgBox Msg, vbOKOnly + vbMsgBoxRight, App.title
        TxtVacName.SetFocus
        Exit Sub
    End If
    Select Case Me.TxtModFlg.Text
        Case "N"
            AddNewRec
            BtnLast_Click
        Case "E"
            FiLLRec
    End Select

    Exit Sub
ErrTrap:

    If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "Â‰«ﬂ Œÿ√ „« ›Ì ≈œŒ«· «·»Ì«‰« ", vbOKOnly + vbMsgBoxRight, App.title
    Else
        MsgBox "Error in Enterd data", vbOKOnly + vbMsgBoxRight, App.title
    End If

End Sub
Private Sub BtnUndo_Click()
    FindRec val(TxtUnitID.Text)
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
            Msg = "No. of records before the update" & vbCrLf & FristCount & vbCrLf & "⁄œœ «·”Ã·«  »⁄œ «· ÕœÌÀ" & vbCrLf & LastCount
        
            If LastCount > FristCount Then
                Msg = Msg + vbCrLf & "No. of new records" & vbCrLf & LastCount - FristCount
            Else
                Msg = Msg + vbCrLf & "No. of deleted records" & vbCrLf & FristCount - LastCount
            End If
        End If
    End If
    MsgBox Msg, vbOKOnly + vbMsgBoxRight + vbInformation, App.title
ErrTrap:
End Sub

Public Sub AddNewRec()

    On Error GoTo ErrTrap
    
    Dim StrRecID As String
    StrRecID = new_id("TblPaymetData", "ID", "")
    RsSavRec.AddNew
    RsSavRec.Fields("ID").value = IIf(StrRecID <> "", StrRecID, Null)
    FiLLRec
ErrTrap:
End Sub
Public Sub FiLLRec()

    On Error GoTo ErrTrap
    
    RsSavRec.Fields("name").value = IIf(TxtVacName.Text <> "", Trim(TxtVacName.Text), Null)
    RsSavRec.Fields("namee").value = IIf(TxtVacNamee.Text <> "", Trim(TxtVacNamee.Text), Null)
    RsSavRec.update
    CuurentLogdata
    If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox " „  ⁄„·Ì… «·Õ›Ÿ »‰Ã«Õ", vbOKOnly + vbMsgBoxRight, App.title
    Else
        MsgBox "Saved Successfully", vbOKOnly + vbMsgBoxRight, App.title
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
    TxtUnitID.Text = IIf(IsNull(RsSavRec.Fields("ID").value), "", RsSavRec.Fields("ID").value)
    TxtVacName.Text = IIf(IsNull(RsSavRec.Fields("name").value), "", RsSavRec.Fields("name").value)
    TxtVacNamee.Text = IIf(IsNull(RsSavRec.Fields("nameE").value), "", RsSavRec.Fields("nameE").value)
    LabCurrRec.Caption = RsSavRec.AbsolutePosition
    LabCountRec.Caption = RsSavRec.RecordCount
    With Grid
        For i = 1 To .Rows - 1
            If Trim(TxtUnitID.Text) = .TextMatrix(i, .ColIndex("UnitID")) Then
                .Row = i
                Exit Sub
            End If
        Next
    End With
ErrTrap:

End Sub
Public Sub EditRec(StrTable As String, RecId As String)
    FiLLRec
End Sub
Private Sub Grid_EnterCell()

    On Error GoTo ErrTrap
    
    FindRec val(Me.Grid.TextMatrix(Me.Grid.Row, Me.Grid.ColIndex("UnitID")))
ErrTrap:
End Sub
Private Sub TxtUnitID_Change()
    Dim TxtMod As String
    TxtMod = TxtModFlg.Text
    TxtModFlg.Text = ""
    TxtModFlg = TxtMod
End Sub
Public Function FindRec(ByVal RecId As Long)

    On Error GoTo ErrTrap
    
    RsSavRec.Find "ID=" & RecId, , adSearchForward, 1

    If Not (RsSavRec.EOF) Then
        FiLLTXT
    End If

    Exit Function
ErrTrap:

    If RsSavRec.EditMode <> adEditNone Then
        RsSavRec.CancelUpdate
        BtnUndo_Click
    End If
End Function
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
    ElseIf TxtModFlg.Text = "R" Then
        Frm2.Enabled = False
        Grid.Enabled = True
        btnModify.Enabled = False
        btnDelete.Enabled = False
        If TxtUnitID.Text <> "" Then
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
    ElseIf TxtModFlg.Text = "E" Then
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
    End If
End Sub
Public Sub FillGridWithData()

    On Error GoTo ErrTrap

    Dim i As Integer
    Dim rs As ADODB.Recordset
    Dim My_SQL As String

    Set rs = New ADODB.Recordset
    My_SQL = "select * From TblPaymetData order by id"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText

    With Me.Grid
        .Rows = 2
        .Clear flexClearScrollable
        If rs.RecordCount > 0 Then
            .Rows = rs.RecordCount + 1
            rs.MoveFirst
            For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("Ser")) = i
                .TextMatrix(i, .ColIndex("UnitName")) = IIf(IsNull(rs.Fields("name").value), "", rs.Fields("name").value)
                .TextMatrix(i, .ColIndex("UnitNamee")) = IIf(IsNull(rs.Fields("nameE").value), "", rs.Fields("nameE").value)
                .TextMatrix(i, .ColIndex("UnitID")) = IIf(IsNull(rs.Fields("id").value), "", rs.Fields("id").value)
                rs.MoveNext
            Next
            rs.Close
        End If
        .RowHeight(-1) = 300
    End With
ErrTrap:
End Sub
Function CuurentLogdata(Optional Currentmode As String)
    LogTextA = "    ‘«‘… " & ScreenNameArabic & CHR(13) & " —ﬁ„ «·ÊÕœ…   " & TxtUnitID.Text & CHR(13) & "  «”„ «·ÊÕœ… " & TxtVacName.Text
    LogTextE = "    Screen  " & ScreenNameEnglish & CHR(13) & " Unit No   " & TxtUnitID.Text & CHR(13) & " Unit Name" & TxtVacNamee.Text
    If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTextE, Me.Name, Me.TxtModFlg
    Else
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTextE, Me.Name, "D"
    End If
End Function
Private Sub TxtVacName_GotFocus()
    SwitchKeyboardLang LANG_ARABIC
End Sub
Private Sub TxtVacNamee_GotFocus()
    SwitchKeyboardLang LANG_ENGLISH
End Sub
'##############################################################################################################################################################################
'##############################################################################################################################################################################
'##############################################################################################################################################################################
Private Sub btnCancel1_Click()
    Unload Me
End Sub
Function CuurentLogdata1(Optional Currentmode As String)
    LogTextA = "    ‘«‘… " & ScreenNameArabic & CHR(13) & " ﬂÊœ " & TxtSerial.Text & CHR(13) & " √”„    «·÷„«‰  " & TxtVacName1
    LogTextE = "    Screen  " & ScreenNameEnglish & CHR(13) & "Code  " & TxtSerial.Text & CHR(13) & " Name Of Gurantee " & TxtVacNamee1
    If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTextE, Me.Name, Me.TxtModFlg1
    Else
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTextE, Me.Name, "D"
    End If
End Function
Private Sub btnDelete1_Click()
    Dim MSGType As Integer
    Dim BlnRecExist As Boolean
    Dim StrMSG  As String
    Dim Msg As String
    
    On Error GoTo ErrTrap
    If DoPremis(Do_Delete, Me.Name, True) = False Then
        Exit Sub
    End If
    If TxtSerial.Text <> "" Then
        CuurentLogdata1 ("D")
        RsSavRec1.delete
        If SystemOptions.UserInterface = EnglishInterface Then
            MsgBox "Deleted", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
        Else
            MsgBox " „  ⁄„·Ì… «·Õ–› »‰Ã«Õ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
        End If
        FillWithData
        btnNext1_Click
    End If
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217873, -2147467259
            If SystemOptions.UserInterface = ArabicInterface Then
                StrMSG = "⁄›Ê« ·« ÌÃÊ“ Õ–› «·”Ã· ·«— »«ÿÂ »»Ì«‰«  √Œ—Ì"
            Else
                StrMSG = "sorry, this record cannot be deleted due to data integration"
            End If
            RsSavRec1.CancelUpdate
            MsgBox StrMSG, vbOKOnly + vbMsgBoxRight, App.title
            Cn.Errors.Clear
    End Select
End Sub
Private Sub btnFirst1_Click()

    On Error GoTo ErrTrap

    Dim Msg As String
    If Me.TxtModFlg1.Text = "N" Then
        FindRec1 val(TxtVac_ID.Text)
        Me.TxtModFlg1.Text = "R"
    End If
    TxtModFlg1 = "R"
    If RsSavRec1.RecordCount = 0 Then
        clear_all Me
        Exit Sub
    End If
BegnieWork:
    RsSavRec1.MoveFirst
    FiLLTXT1

    Exit Sub

ErrTrap:
    Select Case Err.Number
        Case -2147217885
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
        Else
            Msg = "Sorry , this Recored was deleted by other user on the network" & CHR(13)
            Msg = Msg & "Date will be updated now" & CHR(13)
        End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec1.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub btnLast1_Click()

    On Error GoTo ErrTrap

    Dim Msg As String
    If Me.TxtModFlg1.Text = "N" Then
        FindRec1 val(TxtVac_ID.Text)
        Me.TxtModFlg1.Text = "R"
    End If
    TxtModFlg1 = "R"
    If RsSavRec1.RecordCount = 0 Then
        clear_all Me
        Exit Sub
    End If
BegnieWork:
    RsSavRec1.MoveLast
    FiLLTXT1
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
                Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
                Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
                Msg = "Sorry , this Recored was deleted by other user on the network" & CHR(13)
                Msg = Msg & "Date will be updated now" & CHR(13)
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec1.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub btnModify1_Click()
    Dim Msg As String
    If DoPremis(Do_Edit, Me.Name, True) = False Then
        Exit Sub
    End If
    
    On Error GoTo ErrTrap

    If TxtVac_ID.Text <> "" Then
        TxtModFlg1 = "E"
        Frm21.Enabled = True
        Me.TxtVacName1.SetFocus
        CuurentLogdata1
    End If

    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147467259
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "⁄›Ê«" & CHR(13)
                Msg = Msg & " ·«Ì„ﬂ‰  ⁄œÌ· Â–« «·”Ã· ›Ï «·Êﬁ  «·Õ«·Ï" & CHR(13)
                Msg = Msg & "ÕÌÀ «‰Â ﬁÌœ «· ⁄œÌ· „‰ ﬁ»· „” Œœ„ «Œ— ⁄·Ï «·‘»ﬂ…"
            Else
                Msg = "Sorry , this Recored was deleted by other user on the network" & CHR(13)
                Msg = Msg & "Date will be updated now" & CHR(13)
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            If RsSavRec1.EditMode <> adEditNone Then
                RsSavRec1.CancelUpdate
            End If
    End Select
End Sub
Private Sub btnNew1_Click()

    Dim My_SQL As String
    Dim rs As ADODB.Recordset

    If DoPremis(Do_New, Me.Name, True) = False Then
        Exit Sub
    End If

    On Error GoTo ErrTrap
    
    Set rs = New ADODB.Recordset
    Frm21.Enabled = True
    clear_all Me
    TxtModFlg1.Text = "N"

    My_SQL = "gurantee_type"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable

    If rs.RecordCount > 0 Then
        TxtSerial.Text = rs.RecordCount + 1
    Else
        TxtSerial.Text = 1
    End If

    rs.Close
    CmbType.ListIndex = 0
    TxtVacName1.SetFocus
ErrTrap:
End Sub
Private Sub btnNext1_Click()

    On Error GoTo ErrTrap
    
    Dim Msg As String
    If Me.TxtModFlg1.Text = "N" Then
        FindRec1 val(TxtVac_ID.Text)
        Me.TxtModFlg1.Text = "R"
    End If

    TxtModFlg1 = "R"

    If RsSavRec1.RecordCount = 0 Then
        clear_all Me
        Exit Sub
    End If

BegnieWork:
    If RsSavRec1.EOF Then
        RsSavRec1.MoveLast
    Else
        RsSavRec1.MoveNext

        If RsSavRec1.EOF Then
            RsSavRec1.MoveLast
        End If
    End If
    FiLLTXT1
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
                Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
                Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
                Msg = "Sorry , this Recored was deleted by other user on the network" & CHR(13)
                Msg = Msg & "Date will be updated now" & CHR(13)
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec1.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub btnPrevious1_Click()

    On Error GoTo ErrTrap
    
    Dim Msg As String
    If Me.TxtModFlg1.Text = "N" Then
        FindRec1 val(TxtVac_ID.Text)
        Me.TxtModFlg1.Text = "R"
    End If
    TxtModFlg1 = "R"
    If RsSavRec1.RecordCount = 0 Then
        clear_all Me
        Exit Sub
    End If
BegnieWork:

    RsSavRec1.MovePrevious

    If RsSavRec1.BOF Then
        RsSavRec1.MoveFirst
    End If

    FiLLTXT1
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
                Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
                Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
                Msg = "Sorry , this Recored was deleted by other user on the network" & CHR(13)
                Msg = Msg & "Date will be updated now" & CHR(13)
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec1.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub btnSave1_Click()

    On Error GoTo ErrTrap
    
    Dim Msg As String
    Dim StrVacCode As String
    Dim StrVacName As String
    Dim CtrlTxt As Control

    For Each CtrlTxt In Me.Controls
        If TypeOf CtrlTxt Is TextBox Or TypeOf CtrlTxt Is ComboBox Then
            If CtrlTxt.Text = "" And CtrlTxt.Tag <> "" And CtrlTxt.Enabled = True Then
                MsgBox CtrlTxt.Tag, vbOKOnly + vbMsgBoxRight, App.title
                CtrlTxt.SetFocus
                Exit Sub
            End If
        End If
    Next
    StrVacName = IsRecExist("gurantee_type", "name", Trim(TxtVacName1.Text), "name", "Vac_ID<>'" & Trim(TxtVac_ID.Text) & "'")
    If StrVacName <> "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·ﬁœ ”»ﬁ  ”ÃÌ· Â–« «·‰Ê⁄ „‰ ﬁ»·"
        Else
            Msg = "This record already exists"
        End If
        MsgBox Msg, vbOKOnly + vbMsgBoxRight, App.title
        TxtVacName1.SetFocus
        Exit Sub

    End If
    Select Case Me.TxtModFlg1.Text
        Case "N"
            AddNewRec1
            btnLast1_Click
        Case "E"
            FiLLRec1
    End Select
    Exit Sub
ErrTrap:
    If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "Â‰«ﬂ Œÿ√ „« ›Ì ≈œŒ«· «·»Ì«‰« ", vbOKOnly + vbMsgBoxRight, App.title
    Else
        MsgBox "Something went wrong while inserting data", vbOKOnly + vbMsgBoxRight, App.title
    End If

End Sub
Private Sub BtnUndo1_Click()
    FindRec1 val(TxtVac_ID.Text)
    Me.TxtModFlg1.Text = "R"
End Sub
Private Sub BtnUpdate1_Click()

    On Error GoTo ErrTrap
    
    Dim Msg As String
    Dim FristCount As Long
    Dim LastCount As Long
    FristCount = RsSavRec1.RecordCount
    RsSavRec1.Requery
    LastCount = RsSavRec1.RecordCount
    BtnUndo1_Click
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
            Msg = "No. of records before the update" & vbCrLf & FristCount & vbCrLf & "⁄œœ «·”Ã·«  »⁄œ «· ÕœÌÀ" & vbCrLf & LastCount
        
            If LastCount > FristCount Then
                Msg = Msg + vbCrLf & "No. of new records" & vbCrLf & LastCount - FristCount
            Else
                Msg = Msg + vbCrLf & "No. of deleted records" & vbCrLf & FristCount - LastCount
            End If
        End If
    End If
    MsgBox Msg, vbOKOnly + vbMsgBoxRight + vbInformation, App.title
ErrTrap:
End Sub
Private Sub Form_QueryUnload1(Cancel As Integer, UnloadMode As Integer)
    Dim IntResult As String
    Dim StrMSG As String
    
    On Error GoTo ErrTrap

    If Me.TxtModFlg1.Text <> "R" Then
        Select Case Me.TxtModFlg1.Text
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
                btnSave1_Click
            Case vbCancel
                Cancel = True
        End Select
    End If
    Exit Sub
ErrTrap:
End Sub
Private Sub Form_Terminate()
    Set FrmVacancy = Nothing
End Sub
Private Sub Form_Activate()
    Me.ZOrder 0
End Sub
Public Sub AddNewRec1()

    On Error GoTo ErrTrap
    
    Dim StrRecID1 As String
    StrRecID1 = new_id("gurantee_type", "id", "")
    RsSavRec1.AddNew
    RsSavRec1.Fields("id").value = IIf(StrRecID1 <> "", StrRecID1, Null)
    FiLLRec1
ErrTrap:
End Sub
Public Sub FiLLRec1()

    On Error GoTo ErrTrap

    RsSavRec1.Fields("name").value = IIf(TxtVacName1.Text <> "", Trim(TxtVacName1.Text), Null)
    RsSavRec1.Fields("namee").value = IIf(TxtVacNamee1.Text <> "", Trim(TxtVacNamee1.Text), Null)
    RsSavRec1.update
    CuurentLogdata1
    If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox " „  ⁄„·Ì… «·Õ›Ÿ »‰Ã«Õ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Else
        MsgBox "record saved successfully", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    End If
    FillWithData
    TxtModFlg1 = "R"

    Exit Sub
ErrTrap:
    If RsSavRec1.EditMode <> adEditNone Then
        RsSavRec1.CancelUpdate
    End If

End Sub
Public Sub FiLLTXT1()

    On Error GoTo ErrTrap
    
    Dim i As Integer
    Frm21.Enabled = False
    TxtVac_ID.Text = IIf(IsNull(RsSavRec1.Fields("id").value), "", RsSavRec1.Fields("id").value)
    TxtVacName1.Text = IIf(IsNull(RsSavRec1.Fields("name").value), "", RsSavRec1.Fields("name").value)
    LabCurrRec.Caption = RsSavRec1.AbsolutePosition
    LabCountRec.Caption = RsSavRec1.RecordCount
    With Grid1
        For i = 1 To .Rows - 1
            If Trim(TxtVac_ID.Text) = .TextMatrix(i, .ColIndex("id")) Then
                TxtSerial.Text = .TextMatrix(i, .ColIndex("Ser"))
                .Row = i
                Exit Sub
            End If
        Next
    End With
ErrTrap:
End Sub
Public Sub EditRec1(StrTable As String, RecID1 As String)
    FiLLRec1
End Sub
Private Sub Grid1_EnterCell()

    On Error GoTo ErrTrap
    
    FindRec1 val(Me.Grid1.TextMatrix(Me.Grid1.Row, Me.Grid1.ColIndex("id")))
ErrTrap:
End Sub
Private Sub TxtVac_ID_Change()
    Dim TxtMod As String
    TxtMod = TxtModFlg1.Text
    TxtModFlg1.Text = ""
    TxtModFlg1 = TxtMod
End Sub
Public Function FindRec1(ByVal RecID1 As Long)

    On Error GoTo ErrTrap
    
    RsSavRec1.Find "id=" & RecID1, , adSearchForward, 1
    If Not (RsSavRec1.EOF) Then
        FiLLTXT1
    End If
    Exit Function
ErrTrap:
    If RsSavRec1.EditMode <> adEditNone Then
        RsSavRec1.CancelUpdate
        BtnUndo1_Click
    End If
End Function
Private Sub TxtModFlg1_Change()
    If TxtModFlg1.Text = "N" Then
        Frm21.Enabled = True
        Me.btnNew1.Enabled = False
        btnModify1.Enabled = False
        btnDelete1.Enabled = False
        Me.btnQuery.Enabled = False
        Grid1.Enabled = False
        BtnUndo1.Enabled = True
        Me.btnSave1.Enabled = True
        BtnUpdate1.Enabled = False
    ElseIf TxtModFlg1.Text = "R" Then
        Frm21.Enabled = False
        Grid1.Enabled = True
        btnModify1.Enabled = False
        btnDelete1.Enabled = False
        If TxtVac_ID.Text <> "" Then
            btnModify1.Enabled = True
            btnDelete1.Enabled = True
        End If
        BtnUpdate1.Enabled = True
        Me.btnQuery.Enabled = True
        Me.btnNew1.Enabled = True
        BtnUndo1.Enabled = False
        Me.btnSave1.Enabled = False
        btnNext1.Enabled = True
        btnPrevious1.Enabled = True
        btnFirst1.Enabled = True
        btnLast1.Enabled = True
    ElseIf TxtModFlg1.Text = "E" Then
        Frm21.Enabled = True
        Me.btnNew1.Enabled = False
        btnModify1.Enabled = False
        btnDelete1.Enabled = False
        Me.btnQuery.Enabled = False
        BtnUpdate1.Enabled = False
        BtnUndo1.Enabled = True
        Me.btnSave1.Enabled = True
        Grid1.Enabled = False
        btnNext1.Enabled = False
        btnPrevious1.Enabled = False
        btnFirst1.Enabled = False
        btnLast1.Enabled = False
    End If
End Sub
Public Sub FillWithData()

    On Error GoTo ErrTrap

    Dim i As Integer
    Dim rs As ADODB.Recordset
    Dim My_SQL As String
    Set rs = New ADODB.Recordset
    My_SQL = "select * From gurantee_type order by id"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText

    With Me.Grid1
        .Rows = 2
        .Clear flexClearScrollable
        If rs.RecordCount > 0 Then
            .Rows = rs.RecordCount + 1
            rs.MoveFirst
            For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("Ser")) = i
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs.Fields("id").value), "", rs.Fields("id").value)
                .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(rs.Fields("name").value), "", rs.Fields("name").value)
                .TextMatrix(i, .ColIndex("namee")) = IIf(IsNull(rs.Fields("namee").value), "", rs.Fields("namee").value)
                rs.MoveNext
            Next
            rs.Close
        End If
        .RowHeight(-1) = 300
    End With
ErrTrap:
End Sub
Private Sub ShowTip1()
    On Error GoTo ErrTrap
    Dim TTP As New clstooltip
    Dim Wrap As String
    Dim Msg As String
    Wrap = CHR(13) + CHR(10)
    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "ÃœÌœ" & Wrap & "·› Õ ”Ã· ÃœÌœ " & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F12 √Ê Enter"
        .AddControl btnNew1, Msg, True
    End With

    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = " ⁄œÌ·" & Wrap & "· ⁄œÌ·  ”Ã· «·Õ«·Ï " & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F11"
        .AddControl btnModify1, Msg, True
    End With

    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Õ›Ÿ" & Wrap & "· ”ÃÌ· «·»Ì«‰«  œ«Œ· ﬁ«⁄œ… " & Wrap & "«·»Ì«‰«  ≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F10"
        .AddControl btnSave1, Msg, True
    End With

    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = " —«Ã⁄" & Wrap & "·· —«Ã⁄ ⁄‰ «·⁄„·Ì… «·Õ«·Ì…" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F9"
        .AddControl BtnUndo1, Msg, True
    End With

    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Õ–› «·”Ã·" & Wrap & "·Õ–› «·”Ã· «·Õ«·Ï" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F18"
        .AddControl btnDelete1, Msg, True
    End With

    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Œ—ÊÃ" & Wrap & "·≈€·«ﬁ Â–Â «·‰«›–…" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " Ctrl+x"
        .AddControl btnCancel1, Msg, True
    End With

    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«·«Ê·" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·«Ê·" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " Home √Ê UpArrow"
        .AddControl btnFirst1, Msg, True
    End With

    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«·”«»ﬁ" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " PageUp √Ê LeftArrow"
        .AddControl btnPrevious1, Msg, True
    End With

    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«· «·Ï" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ï" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " PageDown √Ê RightArrow"
        .AddControl btnNext1, Msg, True
    End With

    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«·«ŒÌ—" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·«ŒÌ—" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " End √Ê DownArrow"
        .AddControl btnLast1, Msg, True
    End With

ErrTrap:
End Sub
Private Sub Form_KeyDown1(KeyCode As Integer, Shift As Integer)
    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        If Me.TxtModFlg1.Text = "R" Then
            btnNew1_Click
        Else
            SendKeys "{TAB}"
        End If
    End If
    'New ---------------------------
    If KeyCode = vbKeyF12 Then
        If btnNew1.Enabled = False Then Exit Sub
        btnNew1_Click
    End If
    'Edit ------------------------
    If KeyCode = vbKeyF11 Then
        If btnModify1.Enabled = False Then Exit Sub
        btnModify1_Click
    End If
    'save --------------------------------------------------------------------------------
    If KeyCode = vbKeyF10 Then
        If btnSave1.Enabled = False Then Exit Sub
        btnSave1_Click
    End If
    'undo ------------------------------------------------------------------------------
    If KeyCode = vbKeyF9 Then
        If BtnUndo1.Enabled = False Then Exit Sub
        BtnUndo1_Click
    End If
    'Delete ---------------------------------------------------------------------------
    If KeyCode = vbKeyF8 Then
        If btnDelete1.Enabled = False Then Exit Sub
        btnDelete1_Click
    End If
    'Exit ----------------------------------------------------------------------
    If Shift = 2 Then
        If KeyCode = vbKeyX Then
            If btnCancel1.Enabled = False Then Exit Sub
            btnCancel1_Click
        End If
    End If
    'Move first --------------------------------------------
    If KeyCode = vbKeyUp Or KeyCode = vbKeyHome Then
        If btnFirst1.Enabled = False Then Exit Sub
        btnFirst1_Click
    End If
    'Move Previous---------------------------------------------------------
    If KeyCode = vbKeyLeft Or KeyCode = vbKeyPageUp Then
        If btnPrevious1.Enabled = False Then Exit Sub
        btnPrevious1_Click
    End If
    'Move Next---------------------------------------------------------
    If KeyCode = vbKeyRight Or KeyCode = vbKeyPageDown Then
        If btnNext1.Enabled = False Then Exit Sub
        btnNext1_Click
    End If
    'Move Last---------------------------------------------------------
    If KeyCode = vbKeyDown Or KeyCode = vbKeyEnd Then
        If btnLast1.Enabled = False Then Exit Sub
        btnLast1_Click
    End If
    Exit Sub
ErrTrap:
End Sub
Private Function CheckDelCountry(Lngid As Long) As Boolean
    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    StrSQL = "Select * From TblEmployee Where id=" & Lngid & ""
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If Not (rs.BOF Or rs.EOF) Then
        CheckDelCountry = False
    Else
        CheckDelCountry = True
    End If
    rs.Close
    Set rs = Nothing
End Function
'###############################################################################################################################################################################
'###############################################################################################################################################################################
'###############################################################################################################################################################################
Private Sub btnCancel2_Click()
    Unload Me
End Sub
Function CuurentLogdata2(Optional Currentmode As String)
    LogTextA = "    ‘«‘… " & ScreenNameArabic & CHR(13) & " ﬂÊœ " & TxtSerial2.Text & CHR(13) & " √”„ ÿ—Ìﬁ… «·‘Õ‰  " & TxtVacName2
    LogTextE = "    Screen  " & ScreenNameEnglish & CHR(13) & "Code  " & TxtSerial2.Text & CHR(13) & " Name Of Shipments " & TxtVacNameE2
    If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTextE, Me.Name, Me.TxtModFlg2
    Else
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTextE, Me.Name, "D"
    End If
End Function
Private Sub btnDelete2_Click()
    Dim MSGType As Integer
    Dim BlnRecExist As Boolean
    Dim StrMSG  As String
    Dim Msg As String
    
    On Error GoTo ErrTrap

    If DoPremis(Do_Delete, Me.Name, True) = False Then
        Exit Sub
    End If
    If TxtVac_ID2.Text <> "" Then
        CuurentLogdata2 ("D")
        RsSavRec2.delete
        If SystemOptions.UserInterface = EnglishInterface Then
            MsgBox "Deleted", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
        Else
            MsgBox " „  ⁄„·Ì… «·Õ–› »‰Ã«Õ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
        End If
        FillGridWithData2
        btnNext2_Click
    End If
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217873, -2147467259
            If SystemOptions.UserInterface = ArabicInterface Then
                StrMSG = "⁄›Ê« ·« ÌÃÊ“ Õ–› «·”Ã· ·«— »«ÿÂ »»Ì«‰«  √Œ—Ì"
            Else
                StrMSG = "sorry, this record cannot be deleted due to data integration"
            End If
            RsSavRec2.CancelUpdate
            MsgBox StrMSG, vbOKOnly + vbMsgBoxRight, App.title
            Cn.Errors.Clear
    End Select
End Sub
Private Sub btnFirst2_Click()

    On Error GoTo ErrTrap

    Dim Msg As String

    If Me.TxtModFlg2.Text = "N" Then
        FindRec2 val(TxtVac_ID2.Text)
        Me.TxtModFlg2.Text = "R"
    End If
    TxtModFlg2 = "R"
    If RsSavRec2.RecordCount = 0 Then
        clear_all Me
        Exit Sub
    End If
BegnieWork:
    RsSavRec2.MoveFirst
    FiLLTXT2
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
                Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
                Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
                Msg = "Sorry , this Recored was deleted by other user on the network" & CHR(13)
                Msg = Msg & "Date will be updated now" & CHR(13)
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec2.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub btnLast2_Click()

    On Error GoTo ErrTrap

    Dim Msg As String
    If Me.TxtModFlg2.Text = "N" Then
        FindRec2 val(TxtVac_ID2.Text)
        Me.TxtModFlg2.Text = "R"
    End If
    TxtModFlg2 = "R"
    If RsSavRec2.RecordCount = 0 Then
        clear_all Me
        Exit Sub
    End If
BegnieWork:
    RsSavRec2.MoveLast
    FiLLTXT2
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
                Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
                Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
                Msg = "sorry, this record cannot be deleted due to data integration"
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec2.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub btnModify2_Click()

    On Error GoTo ErrTrap
    
    Dim Msg As String
    If DoPremis(Do_Edit, Me.Name, True) = False Then
        Exit Sub
    End If
    If TxtVac_ID2.Text <> "" Then
        TxtModFlg2 = "E"
        Frm22.Enabled = True
        Me.TxtVacName2.SetFocus
        CuurentLogdata2
    End If

    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147467259
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "⁄›Ê«" & CHR(13)
                Msg = Msg & " ·«Ì„ﬂ‰  ⁄œÌ· Â–« «·”Ã· ›Ï «·Êﬁ  «·Õ«·Ï" & CHR(13)
                Msg = Msg & "ÕÌÀ «‰Â ﬁÌœ «· ⁄œÌ· „‰ ﬁ»· „” Œœ„ «Œ— ⁄·Ï «·‘»ﬂ…"
            Else
                Msg = "Sorry" & CHR(13)
                Msg = Msg & "This recored can't be edited now" & CHR(13)
                Msg = Msg & "it's under modification by other user on the network"
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            If RsSavRec2.EditMode <> adEditNone Then
                RsSavRec2.CancelUpdate
            End If
    End Select
End Sub
Private Sub btnNew2_Click()

    On Error GoTo ErrTrap
    
    Dim My_SQL As String
    Dim rs As ADODB.Recordset

    If DoPremis(Do_New, Me.Name, True) = False Then
        Exit Sub
    End If

    Set rs = New ADODB.Recordset
    Frm22.Enabled = True
    clear_all Me
    TxtModFlg2.Text = "N"

    My_SQL = "Shipment_mode"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable

    If rs.RecordCount > 0 Then
        TxtSerial2.Text = rs.RecordCount + 1
    Else
        TxtSerial2.Text = 1
    End If

    rs.Close
    CmbType.ListIndex = 0
    TxtVacName2.SetFocus
ErrTrap:
End Sub
Private Sub btnNext2_Click()

    On Error GoTo ErrTrap
    
    Dim Msg As String
    If Me.TxtModFlg2.Text = "N" Then
        FindRec2 val(TxtVac_ID2.Text)
        Me.TxtModFlg2.Text = "R"
    End If

    TxtModFlg2 = "R"

    If RsSavRec2.RecordCount = 0 Then
        clear_all Me
        Exit Sub
    End If

BegnieWork:
    If RsSavRec2.EOF Then
        RsSavRec2.MoveLast
    Else
        RsSavRec2.MoveNext
        If RsSavRec2.EOF Then
            RsSavRec2.MoveLast
        End If
    End If
    FiLLTXT2
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
                Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
                Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
                Msg = "Sorry , this Recored was deleted by other user on the network" & CHR(13)
                Msg = Msg & "Date will be updated now" & CHR(13)
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec2.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub btnPrevious2_Click()

    On Error GoTo ErrTrap
    
    Dim Msg As String
    If Me.TxtModFlg2.Text = "N" Then
        FindRec2 val(TxtVac_ID2.Text)
        Me.TxtModFlg2.Text = "R"
    End If
    TxtModFlg2 = "R"
    If RsSavRec2.RecordCount = 0 Then
        clear_all Me
        Exit Sub
    End If
BegnieWork:
    RsSavRec2.MovePrevious
    If RsSavRec2.BOF Then
        RsSavRec2.MoveFirst
    End If

    FiLLTXT2
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
                Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
                Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
                Msg = "Sorry , this Recored was deleted by other user on the network" & CHR(13)
                Msg = Msg & "Date will be updated now" & CHR(13)
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec2.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub btnSave2_Click()

    On Error GoTo ErrTrap
    
    Dim Msg As String
    Dim StrVacCode As String
    Dim StrVacName As String
    Dim CtrlTxt As Control

    For Each CtrlTxt In Me.Controls
        If TypeOf CtrlTxt Is TextBox Or TypeOf CtrlTxt Is ComboBox Then
            If CtrlTxt.Text = "" And CtrlTxt.Tag <> "" And CtrlTxt.Enabled = True Then
                MsgBox CtrlTxt.Tag, vbOKOnly + vbMsgBoxRight, App.title
                CtrlTxt.SetFocus
                Exit Sub
            End If
        End If
    Next

    StrVacName = IsRecExist("Shipment_mode", "name", Trim(TxtVacName2.Text), "name", "Vac_ID<>'" & Trim(TxtVac_ID2.Text) & "'")
    If StrVacName <> "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·ﬁœ ”»ﬁ  ”ÃÌ· Â–« «·‰Ê⁄ „‰ ﬁ»·"
        Else
            Msg = "This record already exists"
        End If
        
        MsgBox Msg, vbOKOnly + vbMsgBoxRight, App.title
        TxtVacName2.SetFocus
        Exit Sub
    End If

    Select Case Me.TxtModFlg2.Text
        Case "N"
            AddNewRec2
            btnLast2_Click
        Case "E"
            FiLLRec2
    End Select
    Exit Sub
ErrTrap:
    If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "Â‰«ﬂ Œÿ√ „« ›Ì ≈œŒ«· «·»Ì«‰« ", vbOKOnly + vbMsgBoxRight, App.title
    Else
        MsgBox "Something went wrong while inserting data", vbOKOnly + vbMsgBoxRight, App.title
    End If
End Sub
Private Sub BtnUndo2_Click()
    FindRec2 val(TxtVac_ID2.Text)
    Me.TxtModFlg2.Text = "R"
End Sub
Private Sub BtnUpdate2_Click()

    On Error GoTo ErrTrap
    
    Dim Msg As String
    Dim FristCount As Long
    Dim LastCount As Long
    
    FristCount = RsSavRec2.RecordCount
    RsSavRec2.Requery
    LastCount = RsSavRec2.RecordCount
    BtnUndo2_Click
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
            Msg = "No. of records before the update" & vbCrLf & FristCount & vbCrLf & "⁄œœ «·”Ã·«  »⁄œ «· ÕœÌÀ" & vbCrLf & LastCount
        
            If LastCount > FristCount Then
                Msg = Msg + vbCrLf & "No. of new records" & vbCrLf & LastCount - FristCount
            Else
                Msg = Msg + vbCrLf & "No. of deleted records" & vbCrLf & FristCount - LastCount
            End If
        End If
    End If
    MsgBox Msg, vbOKOnly + vbMsgBoxRight + vbInformation, App.title
ErrTrap:
End Sub

Public Sub AddNewRec2()

    On Error GoTo ErrTrap
    
    Dim StrRecId2 As String
    StrRecId2 = new_id("Shipment_mode", "id", "")
    RsSavRec2.AddNew
    RsSavRec2.Fields("id").value = IIf(StrRecId2 <> "", StrRecId2, Null)
    FiLLRec2
ErrTrap:
End Sub
Public Sub FiLLRec2()

    On Error GoTo ErrTrap

    RsSavRec2.Fields("name").value = IIf(TxtVacName2.Text <> "", Trim(TxtVacName2.Text), Null)
    RsSavRec2.Fields("namee").value = IIf(TxtVacNameE2.Text <> "", Trim(TxtVacNameE2.Text), Null)
    RsSavRec2.update
    CuurentLogdata2
    If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox " „  ⁄„·Ì… «·Õ›Ÿ »‰Ã«Õ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Else
        MsgBox "Data Saved", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    End If
    FillGridWithData2
    TxtModFlg2 = "R"
    Exit Sub
ErrTrap:
    If RsSavRec2.EditMode <> adEditNone Then
        RsSavRec2.CancelUpdate
    End If
End Sub
Public Sub FiLLTXT2()

    On Error GoTo ErrTrap
    
    Dim i As Integer
    Frm22.Enabled = False
    TxtVac_ID2.Text = IIf(IsNull(RsSavRec2.Fields("id").value), "", RsSavRec2.Fields("id").value)
    TxtVacName2.Text = IIf(IsNull(RsSavRec2.Fields("name").value), "", RsSavRec2.Fields("name").value)
    TxtVacNameE2.Text = IIf(IsNull(RsSavRec2.Fields("namee").value), "", RsSavRec2.Fields("namee").value)
    LabCurrRec.Caption = RsSavRec2.AbsolutePosition
    LabCountRec.Caption = RsSavRec2.RecordCount
    With GRID2
        For i = 1 To .Rows - 1
            If Trim(TxtVac_ID2.Text) = .TextMatrix(i, .ColIndex("id")) Then
                TxtSerial2.Text = .TextMatrix(i, .ColIndex("Ser"))
                .Row = i
                Exit Sub
            End If
        Next
    End With
ErrTrap:
End Sub

Public Sub EditRec2(StrTable As String, RecId2 As String)
    FiLLRec2
End Sub
Private Sub Grid2_EnterCell()

    On Error GoTo ErrTrap
    
    FindRec2 val(Me.GRID2.TextMatrix(Me.GRID2.Row, Me.GRID2.ColIndex("id")))
ErrTrap:
End Sub
Private Sub TxtVac_ID2_Change()
    Dim TxtMod As String
    TxtMod = TxtModFlg2.Text
    TxtModFlg2.Text = ""
    TxtModFlg2 = TxtMod
End Sub
Public Function FindRec2(ByVal RecId2 As Long)

    On Error GoTo ErrTrap
    
    RsSavRec2.Find "id=" & RecId2, , adSearchForward, 1
    If Not (RsSavRec2.EOF) Then
        FiLLTXT2
    End If
    Exit Function
ErrTrap:
    If RsSavRec2.EditMode <> adEditNone Then
        RsSavRec2.CancelUpdate
        BtnUndo2_Click
    End If
End Function
Private Sub TxtModFlg2_Change()
    If TxtModFlg2.Text = "N" Then
        Frm22.Enabled = True
        Me.btnNew2.Enabled = False
        btnModify2.Enabled = False
        btnDelete2.Enabled = False
        Me.btnQuery.Enabled = False
        GRID2.Enabled = False
        BtnUndo2.Enabled = True
        Me.btnSave2.Enabled = True
        BtnUpdate2.Enabled = False
    ElseIf TxtModFlg2.Text = "R" Then
        Frm22.Enabled = False
        GRID2.Enabled = True
        btnModify2.Enabled = False
        btnDelete2.Enabled = False
        If TxtVac_ID2.Text <> "" Then
            btnModify2.Enabled = True
            btnDelete2.Enabled = True
        End If
        BtnUpdate2.Enabled = True
        Me.btnQuery.Enabled = True
        Me.btnNew2.Enabled = True
        BtnUndo2.Enabled = False
        Me.btnSave2.Enabled = False
        btnNext2.Enabled = True
        btnPrevious2.Enabled = True
        btnFirst2.Enabled = True
        btnLast2.Enabled = True
    ElseIf TxtModFlg2.Text = "E" Then
        Frm22.Enabled = True
        Me.btnNew2.Enabled = False
        btnModify2.Enabled = False
        btnDelete2.Enabled = False
        Me.btnQuery.Enabled = False
        BtnUpdate2.Enabled = False
        BtnUndo2.Enabled = True
        Me.btnSave2.Enabled = True
        GRID2.Enabled = False
        btnNext2.Enabled = False
        btnPrevious2.Enabled = False
        btnFirst2.Enabled = False
        btnLast2.Enabled = False
    End If
End Sub
Public Sub FillGridWithData2()

    On Error GoTo ErrTrap

    Dim i As Integer
    Dim rs As ADODB.Recordset
    Dim My_SQL As String

    Set rs = New ADODB.Recordset
    My_SQL = "select * From Shipment_mode order by id"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText

    With Me.GRID2
        .Rows = 2
        .Clear flexClearScrollable
        If rs.RecordCount > 0 Then
            .Rows = rs.RecordCount + 1
            rs.MoveFirst
            For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("Ser")) = i
                .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(rs.Fields("name").value), "", rs.Fields("name").value)
                .TextMatrix(i, .ColIndex("namee")) = IIf(IsNull(rs.Fields("namee").value), "", rs.Fields("namee").value)
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs.Fields("id").value), "", rs.Fields("id").value)
                rs.MoveNext
            Next
            rs.Close
        End If
        .RowHeight(-1) = 300
    End With
ErrTrap:
End Sub
Private Sub ShowTip2()

    On Error GoTo ErrTrap
    
    Dim TTP As New clstooltip
    Dim Wrap As String
    Dim Msg As String
    Wrap = CHR(13) + CHR(10)

    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "ÃœÌœ" & Wrap & "·› Õ ”Ã· ÃœÌœ " & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F12 √Ê Enter"
            
        .AddControl btnNew2, Msg, True
    End With

    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = " ⁄œÌ·" & Wrap & "· ⁄œÌ·  ”Ã· «·Õ«·Ï " & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F11"
        .AddControl btnModify2, Msg, True
    End With

    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Õ›Ÿ" & Wrap & "· ”ÃÌ· «·»Ì«‰«  œ«Œ· ﬁ«⁄œ… " & Wrap & "«·»Ì«‰«  ≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F10"
        .AddControl btnSave2, Msg, True
    End With

    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = " —«Ã⁄" & Wrap & "·· —«Ã⁄ ⁄‰ «·⁄„·Ì… «·Õ«·Ì…" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F9"
        .AddControl BtnUndo2, Msg, True
    End With

    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Õ–› «·”Ã·" & Wrap & "·Õ–› «·”Ã· «·Õ«·Ï" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F18"
        .AddControl btnDelete2, Msg, True
    End With

    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Œ—ÊÃ" & Wrap & "·≈€·«ﬁ Â–Â «·‰«›–…" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " Ctrl+x"
        .AddControl btnCancel2, Msg, True
    End With

    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«·«Ê·" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·«Ê·" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " Home √Ê UpArrow"
        .AddControl btnFirst2, Msg, True
    End With

    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«·”«»ﬁ" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " PageUp √Ê LeftArrow"
        .AddControl btnPrevious2, Msg, True
    End With

    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«· «·Ï" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ï" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " PageDown √Ê RightArrow"
        .AddControl btnNext2, Msg, True
    End With

    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«·«ŒÌ—" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·«ŒÌ—" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " End √Ê DownArrow"
        .AddControl btnLast2, Msg, True
    End With

ErrTrap:
End Sub
Private Function CheckDelCountry2(Lngid As Long) As Boolean
    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    StrSQL = "Select * From TblEmployee Where id=" & Lngid & ""
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.BOF Or rs.EOF) Then
        CheckDelCountry2 = False
    Else
        CheckDelCountry2 = True
    End If

    rs.Close
    Set rs = Nothing
End Function
'#################################################################################################################################################################
'#################################################################################################################################################################
'#################################################################################################################################################################
Private Sub btnCancel3_Click()
    Unload Me
End Sub
Function CuurentLogdata3(Optional Currentmode As String)
     LogTextA = "    ‘«‘… " & ScreenNameArabic & CHR(13) & "ﬂÊœ " & TxtSerial3.Text & CHR(13) & "   «”„ «·„Ã„Ê⁄Â " & TxtVacName3
        LogTextE = "    Screen  " & ScreenNameEnglish & CHR(13) & "Code  " & TxtSerial3.Text & CHR(13) & "   Group Name " & TxtVacName3
       If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTextE, Me.Name, Me.TxtModFlg3
    Else
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTextE, Me.Name, "D"
    End If
End Function
Private Sub btnDelete3_Click()

    Dim MSGType As Integer
    Dim BlnRecExist As Boolean
    Dim StrMSG  As String
    Dim Msg As String
    
    On Error GoTo ErrTrap

    If DoPremis(Do_Delete, Me.Name, True) = False Then
        Exit Sub
    End If

    If TxtVac_ID3.Text <> "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MSGType = MsgBox("Â·  —€» ›Ì Õ–› Â–« «·”Ã·", vbQuestion + vbMsgBoxRtlReading + vbYesNo + vbMsgBoxRight, App.title)
        Else
            MSGType = MsgBox("Are you sure you want to delete this record", vbQuestion + vbMsgBoxRtlReading + vbYesNo + vbMsgBoxRight, App.title)
        End If
        If MSGType = vbYes Then
            RsSavRec3.Find "id=" & val(TxtVac_ID3.Text), , adSearchForward, 1
            CuurentLogdata3 ("D")
            RsSavRec3.delete
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox " „  ⁄„·Ì… «·Õ–› »‰Ã«Õ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
            Else
                MsgBox "Record deleted successfully", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            End If
            FillGrid3WithData
            btnNext3_Click
        End If
    End If
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217873, -2147467259
            If SystemOptions.UserInterface = ArabicInterface Then
                StrMSG = "⁄›Ê« ·« ÌÃÊ“ Õ–› «·”Ã· ·«— »«ÿÂ »»Ì«‰«  √Œ—Ì"
            Else
                StrMSG = "sorry, this record cannot be deleted due to data integration"
            End If
            RsSavRec3.CancelUpdate
            MsgBox StrMSG, vbOKOnly + vbMsgBoxRight, App.title
            Cn.Errors.Clear
    End Select
End Sub
Private Sub btnFirst3_Click()

    On Error GoTo ErrTrap

    Dim Msg As String

    If Me.TxtModFlg3.Text = "N" Then
        FindRec3 val(TxtVac_ID3.Text)
        Me.TxtModFlg3.Text = "R"
    End If
    TxtModFlg3 = "R"
    If RsSavRec3.RecordCount = 0 Then
        clear_all Me
        Exit Sub
    End If
BegnieWork:
    RsSavRec3.MoveFirst
    FiLLTXT3
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
                Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
                Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
                Msg = "Sorry , this Recored was deleted by other user on the network" & CHR(13)
                Msg = Msg & "Date will be updated now" & CHR(13)
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec3.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub btnLast3_Click()

    On Error GoTo ErrTrap

    Dim Msg As String

    If Me.TxtModFlg3.Text = "N" Then
        FindRec3 val(TxtVac_ID3.Text)
        Me.TxtModFlg3.Text = "R"
    End If
    TxtModFlg3 = "R"
    If RsSavRec3.RecordCount = 0 Then
        clear_all Me
        Exit Sub
    End If
BegnieWork:

    RsSavRec3.MoveLast
    FiLLTXT3
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
                Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
                Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
                Msg = "Sorry , this Recored was deleted by other user on the network" & CHR(13)
                Msg = Msg & "Date will be updated now" & CHR(13)
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec3.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub btnModify3_Click()

    Dim Msg As String
    
    On Error GoTo ErrTrap

    If DoPremis(Do_Edit, Me.Name, True) = False Then
        Exit Sub
    End If
    If TxtVac_ID3.Text <> "" Then
        TxtModFlg3 = "E"
        Frm23.Enabled = True
        Me.TxtVacName3.SetFocus
        CuurentLogdata3
    End If
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147467259
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "⁄›Ê«" & CHR(13)
                Msg = Msg & " ·«Ì„ﬂ‰  ⁄œÌ· Â–« «·”Ã· ›Ï «·Êﬁ  «·Õ«·Ï" & CHR(13)
                Msg = Msg & "ÕÌÀ «‰Â ﬁÌœ «· ⁄œÌ· „‰ ﬁ»· „” Œœ„ «Œ— ⁄·Ï «·‘»ﬂ…"
            Else
                Msg = "Sorry" & CHR(13)
                Msg = Msg & "This recored can't be edited now" & CHR(13)
                Msg = Msg & "it's under modification by other user on the network"
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            If RsSavRec3.EditMode <> adEditNone Then
                RsSavRec3.CancelUpdate
            End If
    End Select
End Sub
Private Sub btnNew3_Click()

    Dim My_SQL As String
    Dim rs As ADODB.Recordset
    
    On Error GoTo ErrTrap

    If DoPremis(Do_New, Me.Name, True) = False Then
        Exit Sub
    End If
    Set rs = New ADODB.Recordset
    Frm23.Enabled = True
    clear_all Me
    TxtModFlg3.Text = "N"
    My_SQL = "TBLSalesRepGroups2"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable
    If rs.RecordCount > 0 Then
        TxtSerial3.Text = rs.RecordCount + 1
    Else
        TxtSerial3.Text = 1
    End If
    rs.Close
    CmbType.ListIndex = 0
    TxtVacName3.SetFocus
ErrTrap:
End Sub
Private Sub btnNext3_Click()

    On Error GoTo ErrTrap
    
    Dim Msg As String

    If Me.TxtModFlg3.Text = "N" Then
        FindRec3 val(TxtVac_ID3.Text)
        Me.TxtModFlg3.Text = "R"
    End If
    TxtModFlg3 = "R"
    If RsSavRec3.RecordCount = 0 Then
        clear_all Me
        Exit Sub
    End If
BegnieWork:
    If RsSavRec3.EOF Then
        RsSavRec3.MoveLast
    Else
        RsSavRec3.MoveNext
        If RsSavRec3.EOF Then
            RsSavRec3.MoveLast
        End If
    End If
    FiLLTXT3
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
                Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
                Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
                Msg = "Sorry , this Recored was deleted by other user on the network" & CHR(13)
                Msg = Msg & "Date will be updated now" & CHR(13)
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec3.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub btnPrevious3_Click()

    On Error GoTo ErrTrap
    
    Dim Msg As String

    If Me.TxtModFlg3.Text = "N" Then
        FindRec3 val(TxtVac_ID3.Text)
        Me.TxtModFlg3.Text = "R"
    End If
    TxtModFlg3 = "R"
    If RsSavRec3.RecordCount = 0 Then
        clear_all Me
        Exit Sub
    End If
BegnieWork:
    RsSavRec3.MovePrevious
    If RsSavRec3.BOF Then
        RsSavRec3.MoveFirst
    End If
    FiLLTXT3
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
                Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
                Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
                Msg = "Sorry , this Recored was deleted by other user on the network" & CHR(13)
                Msg = Msg & "Date will be updated now" & CHR(13)
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec3.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub btnSave3_Click()

    On Error GoTo ErrTrap
    
    Dim Msg As String
    Dim StrVacCode As String
    Dim StrVacName As String
    Dim CtrlTxt As Control

    For Each CtrlTxt In Me.Controls
        If TypeOf CtrlTxt Is TextBox Or TypeOf CtrlTxt Is ComboBox Then
            If CtrlTxt.Text = "" And CtrlTxt.Tag <> "" And CtrlTxt.Enabled = True Then
                MsgBox CtrlTxt.Tag, vbOKOnly + vbMsgBoxRight, App.title
                CtrlTxt.SetFocus
                Exit Sub
            End If
        End If
    Next
    StrVacName = IsRecExist("TBLSalesRepGroups2", "name", Trim(TxtVacName3.Text), "name", "Vac_ID<>'" & Trim(TxtVac_ID3.Text) & "'")
    If StrVacName <> "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·ﬁœ ”»ﬁ  ”ÃÌ· Â–« «·‰Ê⁄ „‰ ﬁ»·"
        Else
            Msg = "This record already exists"
        End If
        MsgBox Msg, vbOKOnly + vbMsgBoxRight, App.title
        TxtVacName3.SetFocus
        Exit Sub
    End If
    Select Case Me.TxtModFlg3.Text
        Case "N"
            AddNewRec3
            btnLast3_Click
        Case "E"
            FiLLRec3
    End Select
    Exit Sub
ErrTrap:
    If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "Â‰«ﬂ Œÿ√ „« ›Ì ≈œŒ«· «·»Ì«‰« ", vbOKOnly + vbMsgBoxRight, App.title
    Else
        MsgBox "Something went wrong while inserting data", vbOKOnly + vbMsgBoxRight, App.title
    End If
End Sub
Private Sub BtnUndo3_Click()
    FindRec3 val(TxtVac_ID3.Text)
    Me.TxtModFlg3.Text = "R"
End Sub
Private Sub BtnUpdate3_Click()

    On Error GoTo ErrTrap
    
    Dim Msg As String
    Dim FristCount As Long
    Dim LastCount As Long
    
    FristCount = RsSavRec3.RecordCount
    RsSavRec3.Requery
    LastCount = RsSavRec3.RecordCount
    BtnUndo3_Click
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
            Msg = "No. of records before the update" & vbCrLf & FristCount & vbCrLf & "⁄œœ «·”Ã·«  »⁄œ «· ÕœÌÀ" & vbCrLf & LastCount
        
            If LastCount > FristCount Then
                Msg = Msg + vbCrLf & "No. of new records" & vbCrLf & LastCount - FristCount
            Else
                Msg = Msg + vbCrLf & "No. of deleted records" & vbCrLf & FristCount - LastCount
            End If
        End If
    End If
    MsgBox Msg, vbOKOnly + vbMsgBoxRight + vbInformation, App.title
ErrTrap:
End Sub
Private Sub Form_QueryUnload3(Cancel As Integer, UnloadMode As Integer)

    Dim IntResult As String
    Dim StrMSG As String
    
    On Error GoTo ErrTrap

    If Me.TxtModFlg3.Text <> "R" Then
        Select Case Me.TxtModFlg3.Text
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
                btnSave3_Click
            Case vbCancel
                Cancel = True
        End Select
    End If
    Exit Sub
ErrTrap:
End Sub
Public Sub AddNewRec3()

    On Error GoTo ErrTrap
    
    Dim StrRecID As String
    StrRecID = new_id("TBLSalesRepGroups2", "id", "")
    RsSavRec3.AddNew
    RsSavRec3.Fields("id").value = IIf(StrRecID <> "", StrRecID, Null)
    FiLLRec3
ErrTrap:
End Sub
Public Sub FiLLRec3()

    On Error GoTo ErrTrap

    RsSavRec3.Fields("name").value = IIf(TxtVacName3.Text <> "", Trim(TxtVacName3.Text), Null)
    RsSavRec3.Fields("namee").value = IIf(TxtVacNamee3.Text <> "", Trim(TxtVacNamee3.Text), Null)
    RsSavRec3.update
    CuurentLogdata3
    If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox " „  ⁄„·Ì… «·Õ›Ÿ »‰Ã«Õ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Else
        MsgBox "Record saved successfully", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    End If
    FillGrid3WithData
    TxtModFlg3 = "R"
    Exit Sub
ErrTrap:
    If RsSavRec3.EditMode <> adEditNone Then
        RsSavRec3.CancelUpdate
    End If
End Sub
Public Sub FiLLTXT3()

    On Error GoTo ErrTrap
    
    Dim i As Integer
    Frm23.Enabled = False
    TxtVac_ID3.Text = IIf(IsNull(RsSavRec3.Fields("id").value), "", RsSavRec3.Fields("id").value)
    TxtVacName3.Text = IIf(IsNull(RsSavRec3.Fields("name").value), "", RsSavRec3.Fields("name").value)
    TxtVacNamee3.Text = IIf(IsNull(RsSavRec3.Fields("namee").value), "", RsSavRec3.Fields("namee").value)
    LabCurrRec3.Caption = RsSavRec3.AbsolutePosition
    LabCountRec3.Caption = RsSavRec3.RecordCount
    With Grid3
        For i = 1 To .Rows - 1
            If Trim(TxtVac_ID3.Text) = .TextMatrix(i, .ColIndex("id")) Then
                TxtSerial3.Text = .TextMatrix(i, .ColIndex("Ser"))
                .Row = i
                Exit Sub
            End If
        Next
    End With
ErrTrap:
End Sub
Public Sub EditRec3(StrTable As String, RecId As String)
    FiLLRec3
End Sub
Private Sub Grid3_EnterCell()
    On Error GoTo ErrTrap
    FindRec3 val(Me.Grid3.TextMatrix(Me.Grid3.Row, Me.Grid3.ColIndex("id")))
ErrTrap:
End Sub
Private Sub TxtVac_ID3_Change()
    Dim TxtMod As String
    TxtMod = TxtModFlg3.Text
    TxtModFlg3.Text = ""
    TxtModFlg3 = TxtMod
End Sub
Public Function FindRec3(ByVal RecId As Long)

    On Error GoTo ErrTrap
    
    RsSavRec3.Find "id=" & RecId, , adSearchForward, 1
    If Not (RsSavRec3.EOF) Then
        FiLLTXT3
    End If
    Exit Function
ErrTrap:
    If RsSavRec3.EditMode <> adEditNone Then
        RsSavRec3.CancelUpdate
        BtnUndo3_Click
    End If
End Function
Private Sub TxtModFlg3_Change()
    If TxtModFlg3.Text = "N" Then
        Frm23.Enabled = True
        Me.btnNew3.Enabled = False
        btnModify3.Enabled = False
        btnDelete3.Enabled = False
        Me.btnQuery.Enabled = False
        Grid3.Enabled = False
        BtnUndo3.Enabled = True
        Me.btnSave3.Enabled = True
        BtnUpdate.Enabled = False
    ElseIf TxtModFlg3.Text = "R" Then
        Frm23.Enabled = False
        Grid3.Enabled = True
        btnModify3.Enabled = False
        btnDelete3.Enabled = False
        If TxtVac_ID3.Text <> "" Then
            btnModify3.Enabled = True
            btnDelete3.Enabled = True
        End If
        BtnUpdate.Enabled = True
        Me.btnQuery.Enabled = True
        Me.btnNew3.Enabled = True
        BtnUndo3.Enabled = False
        Me.btnSave3.Enabled = False
        btnNext3.Enabled = True
        btnPrevious3.Enabled = True
        btnFirst3.Enabled = True
        btnLast3.Enabled = True
    ElseIf TxtModFlg3.Text = "E" Then
        Frm23.Enabled = True
        Me.btnNew3.Enabled = False
        btnModify3.Enabled = False
        btnDelete3.Enabled = False
        Me.btnQuery.Enabled = False
        BtnUpdate.Enabled = False
        BtnUndo3.Enabled = True
        Me.btnSave3.Enabled = True
        Grid3.Enabled = False
        btnNext3.Enabled = False
        btnPrevious3.Enabled = False
        btnFirst3.Enabled = False
        btnLast3.Enabled = False
    End If
End Sub
Public Sub FillGrid3WithData()

    On Error GoTo ErrTrap

    Dim i As Integer
    Dim rs As ADODB.Recordset
    Dim My_SQL As String

    Set rs = New ADODB.Recordset
    My_SQL = "select * From TBLSalesRepGroups2 order by id"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText
    With Me.Grid3
        .Rows = 2
        .Clear flexClearScrollable
        If rs.RecordCount > 0 Then
            .Rows = rs.RecordCount + 1
            rs.MoveFirst
            For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("Ser")) = i
                .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(rs.Fields("name").value), "", rs.Fields("name").value)
                .TextMatrix(i, .ColIndex("namee")) = IIf(IsNull(rs.Fields("namee").value), "", rs.Fields("namee").value)
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs.Fields("id").value), "", rs.Fields("id").value)
                rs.MoveNext
            Next

            rs.Close
        End If
        .RowHeight(-1) = 300
    End With
ErrTrap:
End Sub
Private Function CheckDelCountry3(Lngid As Long) As Boolean
    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    StrSQL = "Select * From TblEmployee Where id=" & Lngid & ""
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.BOF Or rs.EOF) Then
        CheckDelCountry3 = False
    Else
        CheckDelCountry3 = True
    End If

    rs.Close
    Set rs = Nothing
End Function
'##############################################################################################################################################################################################
'##############################################################################################################################################################################################
'##############################################################################################################################################################################################
Private Sub btnCancel4_Click()
    Unload Me
End Sub
Private Sub btnDelete4_Click()

    On Error GoTo ErrTrap

    Dim MSGType As Integer
    Dim BlnRecExist As Boolean
    Dim StrMSG  As String
    Dim Msg As String

    If TxtUnitID4.Text <> "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MSGType = MsgBox("Â·  —€» ›Ì Õ–› Â–« «·”Ã·", vbYesNo + vbMsgBoxRight, App.title)
        Else
            MSGType = MsgBox("Delete This Record", vbYesNo + vbMsgBoxRight, App.title)
        End If
        If MSGType = vbYes Then
            RsSavRec4.Find "id=" & val(TxtUnitID4.Text), , adSearchForward, 1
            If Not (RsSavRec4.BOF Or RsSavRec4.EOF) Then
                CuurentLogdata4 ("D")
                RsSavRec4.delete
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox " „  ⁄„·Ì… «·Õ–› »‰Ã«Õ", vbOKOnly + vbMsgBoxRight, App.title
                Else
                    MsgBox "Delete Success...", vbOKOnly + vbMsgBoxRight, App.title
                End If
                FillGrid4WithData
                btnNext4_Click
            End If
        End If
    
    End If
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217873, -2147467259
            If SystemOptions.UserInterface = ArabicInterface Then
                StrMSG = "⁄›Ê« ·« ÌÃÊ“ Õ–› «·”Ã· ·«— »«ÿÂ »»Ì«‰«  √Œ—Ì"
            Else
                StrMSG = "Sorry .. can't Delete this record , Reason : Data integrity"
            End If
            RsSavRec4.CancelUpdate
            MsgBox StrMSG, vbOKOnly + vbMsgBoxRight, App.title
            Cn.Errors.Clear
    End Select
End Sub
Private Sub btnFirst4_Click()

    On Error GoTo ErrTrap

    Dim Msg As String

    If Me.TxtModFlg4.Text = "N" Then
        FindRec44 val(Me.TxtUnitID4.Text)
        Me.TxtModFlg4.Text = "R"
    End If
    TxtModFlg4 = "R"
    If RsSavRec4.RecordCount = 0 Then
        clear_all Me
        Exit Sub
    End If
BegnieWork:
    RsSavRec4.MoveFirst
    FiLLTXT4
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
                Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
                Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
                Msg = "Sorry.. this record Already Deleted" & CHR(13)
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec4.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub btnLast4_Click()

    On Error GoTo ErrTrap

    Dim Msg As String

    If Me.TxtModFlg4.Text = "N" Then
        FindRec44 val(Me.TxtUnitID4.Text)
        Me.TxtModFlg4.Text = "R"
    End If
    TxtModFlg4 = "R"
    If RsSavRec4.RecordCount = 0 Then
        clear_all Me
        Exit Sub
    End If
BegnieWork:
    RsSavRec4.MoveLast
    FiLLTXT4
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
                Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
                Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
                Msg = "Sorry.. this record Already Deleted" & CHR(13)
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec4.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub btnModify4_Click()

    On Error GoTo ErrTrap

    Dim Msg As String

    If TxtUnitID4.Text <> "" Then
        TxtModFlg4 = "E"
        Frm24.Enabled = True
        Me.TxtVacName4.SetFocus
        CuurentLogdata4
    End If
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147467259
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "⁄›Ê«" & CHR(13)
                Msg = Msg & " ·«Ì„ﬂ‰  ⁄œÌ· Â–« «·”Ã· ›Ï «·Êﬁ  «·Õ«·Ï" & CHR(13)
                Msg = Msg & "ÕÌÀ «‰Â ﬁÌœ «· ⁄œÌ· „‰ ﬁ»· „” Œœ„ «Œ— ⁄·Ï «·‘»ﬂ…"
            Else
                Msg = "Sorry" & CHR(13)
                Msg = Msg & " Can't Edit this record now - Another user work with it now" & CHR(13)
       
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            If RsSavRec4.EditMode <> adEditNone Then
                RsSavRec4.CancelUpdate
            End If
    End Select
End Sub
Private Sub btnNew4_Click()

    On Error GoTo ErrTrap
    
    Dim My_SQL As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    Frm24.Enabled = True
    clear_all Me
    TxtModFlg4.Text = "N"
    My_SQL = "TblShipingData"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable
    If rs.RecordCount > 0 Then
        TxtUnitID4.Text = rs.RecordCount + 1
    Else
        TxtUnitID4.Text = 1
    End If
    rs.Close
    CmbType.ListIndex = 0
    TxtVacName4.SetFocus
ErrTrap:
End Sub
Private Sub btnNext4_Click()

    On Error GoTo ErrTrap
    
    Dim Msg As String

    If Me.TxtModFlg4.Text = "N" Then
        FindRec44 val(Me.TxtUnitID4.Text)
        Me.TxtModFlg4.Text = "R"
    End If
    TxtModFlg4 = "R"
    If RsSavRec4.RecordCount = 0 Then
        clear_all Me
        Exit Sub
    End If

BegnieWork:

    If RsSavRec4.EOF Then
        RsSavRec4.MoveLast
    Else
        RsSavRec4.MoveNext

        If RsSavRec4.EOF Then
            RsSavRec4.MoveLast
        End If
    End If
    FiLLTXT4
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
                Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
                Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
                Msg = "Sorry , this Recored was deleted by other user on the network" & CHR(13)
                Msg = Msg & "Date will be updated now" & CHR(13)
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec4.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub btnPrevious4_Click()

    On Error GoTo ErrTrap
    
    Dim Msg As String

    If Me.TxtModFlg4.Text = "N" Then
        FindRec44 val(Me.TxtUnitID4.Text)
        Me.TxtModFlg4.Text = "R"
    End If

    TxtModFlg4 = "R"

    If RsSavRec4.RecordCount = 0 Then
        clear_all Me
        Exit Sub
    End If

BegnieWork:
    RsSavRec4.MovePrevious

    If RsSavRec4.BOF Then
        RsSavRec4.MoveFirst
    End If

    FiLLTXT4
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
                Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
                Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
                Msg = "Sorry , this Recored was deleted by other user on the network" & CHR(13)
                Msg = Msg & "Date will be updated now" & CHR(13)
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec4.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub btnSave4_Click()

    On Error GoTo ErrTrap
    
    Dim Msg As String
    Dim StrVacCode As String
    Dim StrVacName As String
    Dim CtrlTxt As Control

    If Trim(Me.TxtVacName4.Text) = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "ÌÃ» ﬂ «»… «”„ «·‰Ê⁄ ...!!!"
        Else
            Msg = "Please Enter The Name"
        End If
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        TxtVacName4.SetFocus
        Exit Sub
    End If
    StrVacName = ""
    StrVacName = IsRecExist("TblShipingData", "name", Trim(TxtVacName4.Text), "name", "ID<>'" & Trim(TxtUnitID4.Text) & "'")
    If StrVacName <> "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·ﬁœ ”»ﬁ  ”ÃÌ· Â–Â «·‰Ê⁄ „‰ ﬁ»·"
        Else
            Msg = "this Unit Already Exist"
        End If
        MsgBox Msg, vbOKOnly + vbMsgBoxRight, App.title
        TxtVacName4.SetFocus
        Exit Sub
    End If
    Select Case Me.TxtModFlg4.Text
        Case "N"
            AddNewRec4
            btnLast4_Click
        Case "E"
            FiLLRec4
    End Select
    Exit Sub
ErrTrap:
    If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "Â‰«ﬂ Œÿ√ „« ›Ì ≈œŒ«· «·»Ì«‰« ", vbOKOnly + vbMsgBoxRight, App.title
    Else
        MsgBox "Error in Enterd data", vbOKOnly + vbMsgBoxRight, App.title
    End If
End Sub
Private Sub BtnUndo4_Click()
    FindRec44 val(TxtUnitID4.Text)
    Me.TxtModFlg4.Text = "R"
End Sub
Private Sub BtnUpdate4_Click()

    On Error GoTo ErrTrap
    
    Dim Msg As String
    Dim FristCount As Long
    Dim LastCount As Long
    
    FristCount = RsSavRec4.RecordCount
    RsSavRec4.Requery
    LastCount = RsSavRec4.RecordCount
    BtnUndo4_Click
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
            Msg = "No. of records before the update" & vbCrLf & FristCount & vbCrLf & "⁄œœ «·”Ã·«  »⁄œ «· ÕœÌÀ" & vbCrLf & LastCount
        
            If LastCount > FristCount Then
                Msg = Msg + vbCrLf & "No. of new records" & vbCrLf & LastCount - FristCount
            Else
                Msg = Msg + vbCrLf & "No. of deleted records" & vbCrLf & FristCount - LastCount
            End If
        End If
    End If
    MsgBox Msg, vbOKOnly + vbMsgBoxRight + vbInformation, App.title
ErrTrap:
End Sub
Public Sub AddNewRec4()

    On Error GoTo ErrTrap
    
    Dim StrRecId4 As String
    
    StrRecId4 = new_id("TblShipingData", "ID", "")
    RsSavRec4.AddNew
    RsSavRec4.Fields("ID").value = IIf(StrRecId4 <> "", StrRecId4, Null)
    FiLLRec4
ErrTrap:
End Sub
Public Sub FiLLRec4()

    On Error GoTo ErrTrap

    RsSavRec4.Fields("name").value = IIf(TxtVacName4.Text <> "", Trim(TxtVacName4.Text), Null)
    RsSavRec4.Fields("namee").value = IIf(TxtVacNamee4.Text <> "", Trim(TxtVacNamee4.Text), Null)
    RsSavRec4.update
    CuurentLogdata4
    If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox " „  ⁄„·Ì… «·Õ›Ÿ »‰Ã«Õ", vbOKOnly + vbMsgBoxRight, App.title
    Else
        MsgBox "Saved Successfully", vbOKOnly + vbMsgBoxRight, App.title
    End If

    FillGrid4WithData
    TxtModFlg4 = "R"
    Exit Sub
ErrTrap:
    If RsSavRec4.EditMode <> adEditNone Then
        RsSavRec4.CancelUpdate
    End If

End Sub
Public Sub FiLLTXT4()

    On Error GoTo ErrTrap
    
    Dim i As Integer
    
    Frm24.Enabled = False
    TxtUnitID4.Text = IIf(IsNull(RsSavRec4.Fields("ID").value), "", RsSavRec4.Fields("ID").value)
    TxtVacName4.Text = IIf(IsNull(RsSavRec4.Fields("name").value), "", RsSavRec4.Fields("name").value)
    TxtVacNamee4.Text = IIf(IsNull(RsSavRec4.Fields("nameE").value), "", RsSavRec4.Fields("nameE").value)
    LabCurrRec4.Caption = RsSavRec4.AbsolutePosition
    LabCountRec4.Caption = RsSavRec4.RecordCount

    With Grid4
        For i = 1 To .Rows - 1
            If Trim(TxtUnitID4.Text) = .TextMatrix(i, .ColIndex("UnitID")) Then
                .Row = i
                Exit Sub
            End If
        Next
    End With
ErrTrap:

End Sub
Public Sub EditRec4(StrTable As String, RecId4 As String)
    FiLLRec4
End Sub
Private Sub Grid4_EnterCell()

    On Error GoTo ErrTrap
    
    FindRec44 val(Me.Grid4.TextMatrix(Me.Grid4.Row, Me.Grid4.ColIndex("UnitID")))
ErrTrap:
End Sub
Private Sub TxtUnitID4_Change()
    Dim TxtMod As String
    TxtMod = TxtModFlg4.Text
    TxtModFlg4.Text = ""
    TxtModFlg4 = TxtMod
End Sub
Public Function FindRec44(ByVal RecId4 As Long)

    On Error GoTo ErrTrap
    
    RsSavRec4.Find "ID=" & RecId4, , adSearchForward, 1
    If Not (RsSavRec4.EOF) Then
        FiLLTXT4
    End If
    Exit Function
ErrTrap:
    If RsSavRec4.EditMode <> adEditNone Then
        RsSavRec4.CancelUpdate
        BtnUndo4_Click
    End If
End Function
Private Sub TxtModFlg4_Change()

    If TxtModFlg4.Text = "N" Then
        Frm24.Enabled = True
        Me.btnNew4.Enabled = False
        btnModify4.Enabled = False
        btnDelete4.Enabled = False
        Me.btnQuery4.Enabled = False
        Grid4.Enabled = False
        BtnUndo4.Enabled = True
        Me.btnSave4.Enabled = True
        BtnUpdate4.Enabled = False
    ElseIf TxtModFlg4.Text = "R" Then
        Frm24.Enabled = False
        Grid4.Enabled = True
        btnModify4.Enabled = False
        btnDelete4.Enabled = False

        If TxtUnitID4.Text <> "" Then
            btnModify4.Enabled = True
            btnDelete4.Enabled = True
        End If
        BtnUpdate4.Enabled = True
        Me.btnQuery4.Enabled = True
        Me.btnNew4.Enabled = True
        BtnUndo4.Enabled = False
        Me.btnSave4.Enabled = False
        btnNext4.Enabled = True
        btnPrevious4.Enabled = True
        btnFirst4.Enabled = True
        btnLast4.Enabled = True
    ElseIf TxtModFlg4.Text = "E" Then
        Frm24.Enabled = True
        Me.btnNew4.Enabled = False
        btnModify4.Enabled = False
        btnDelete4.Enabled = False
        Me.btnQuery4.Enabled = False
        BtnUpdate4.Enabled = False
        BtnUndo4.Enabled = True
        Me.btnSave4.Enabled = True
        Grid4.Enabled = False
        btnNext4.Enabled = False
        btnPrevious4.Enabled = False
        btnFirst4.Enabled = False
        btnLast4.Enabled = False
    End If
End Sub
Public Sub FillGrid4WithData()

    On Error GoTo ErrTrap

    Dim i As Integer
    Dim rs As ADODB.Recordset
    Dim My_SQL As String

    Set rs = New ADODB.Recordset
    My_SQL = "select * From TblShipingData order by id"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    With Me.Grid4
        .Rows = 2
        .Clear flexClearScrollable
        If rs.RecordCount > 0 Then
            .Rows = rs.RecordCount + 1
            rs.MoveFirst
            For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("Ser")) = i
                .TextMatrix(i, .ColIndex("UnitName")) = IIf(IsNull(rs.Fields("name").value), "", rs.Fields("name").value)
                .TextMatrix(i, .ColIndex("UnitNamee")) = IIf(IsNull(rs.Fields("nameE").value), "", rs.Fields("nameE").value)
                .TextMatrix(i, .ColIndex("UnitID")) = IIf(IsNull(rs.Fields("id").value), "", rs.Fields("id").value)
                rs.MoveNext
            Next
            rs.Close
        End If
        .RowHeight(-1) = 300
    End With
ErrTrap:
End Sub
Function CuurentLogdata4(Optional Currentmode As String)
    LogTextA = "    ‘«‘… " & ScreenNameArabic & CHR(13) & " —ﬁ„ «·ÊÕœ…   " & TxtUnitID4.Text & CHR(13) & "  «”„ «·ÊÕœ… " & TxtVacName4.Text
    LogTextE = "    Screen  " & ScreenNameEnglish & CHR(13) & " Unit No   " & TxtUnitID4.Text & CHR(13) & " Unit Name" & TxtVacNamee4.Text
    If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTextE, Me.Name, Me.TxtModFlg4
    Else
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTextE, Me.Name, "D"
    End If
End Function
Private Sub TxtVacName4_GotFocus()
    SwitchKeyboardLang LANG_ARABIC
End Sub
Private Sub TxtVacNamee4_GotFocus()
    SwitchKeyboardLang LANG_ENGLISH
End Sub
'##############################################################################################################################################################################################
'##############################################################################################################################################################################################
'##############################################################################################################################################################################################
Private Sub btnCancel5_Click()
    Unload Me
End Sub
Private Sub btnDelete5_Click()
    Dim MSGType As Integer
    Dim BlnRecExist As Boolean
    Dim StrMSG  As String
    Dim Msg As String
    
    On Error GoTo ErrTrap

    If DoPremis(Do_Delete, Me.Name, True) = False Then
        Exit Sub
    End If
    If SystemOptions.UserInterface = ArabicInterface Then
        MSGType = MsgBox("Â·  —€» ›Ì Õ–› Â–« «·”Ã·", vbQuestion + vbMsgBoxRtlReading + vbYesNo + vbMsgBoxRight, App.title)
    Else
        MSGType = MsgBox("Do you want to delete this record", vbQuestion + vbMsgBoxRtlReading + vbYesNo + vbMsgBoxRight, App.title)
    End If
    If MSGType = vbYes Then
        RsSavRec5.Find "id=" & val(TxtVac_ID5.Text), , adSearchForward, 1
        RsSavRec5.delete
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox " „  ⁄„·Ì… «·Õ–› »‰Ã«Õ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
        Else
            MsgBox "Record deleted successfully", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        End If
        FillGrid5WithData
        btnNext5_Click
    End If
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217873, -2147467259
            If SystemOptions.UserInterface = ArabicInterface Then
                StrMSG = "⁄›Ê« ·« ÌÃÊ“ Õ–› «·”Ã· ·«— »«ÿÂ »»Ì«‰«  √Œ—Ì"
            Else
                StrMSG = "sorry, this record cannot be deleted due to data integration"
            End If
            RsSavRec5.CancelUpdate
            MsgBox StrMSG, vbOKOnly + vbMsgBoxRight, App.title
            Cn.Errors.Clear
    End Select
End Sub
Private Sub btnFirst5_Click()

    On Error GoTo ErrTrap

    Dim Msg As String

    If Me.TxtModFlg5.Text = "N" Then
        FindRec5 val(TxtVac_ID5.Text)
        Me.TxtModFlg5.Text = "R"
    End If
    TxtModFlg5 = "R"
    If RsSavRec5.RecordCount = 0 Then
        clear_all Me
        Exit Sub
    End If
BegnieWork:
    RsSavRec5.MoveFirst
    FiLLTXT5
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
                Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
                Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
                Msg = "Sorry , this Recored was deleted by other user on the network" & CHR(13)
                Msg = Msg & "Date will be updated now" & CHR(13)
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec5.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub btnLast5_Click()

    On Error GoTo ErrTrap

    Dim Msg As String

    If Me.TxtModFlg5.Text = "N" Then
        FindRec5 val(TxtVac_ID5.Text)
        Me.TxtModFlg5.Text = "R"
    End If
    TxtModFlg5 = "R"
    If RsSavRec5.RecordCount = 0 Then
        clear_all Me
        Exit Sub
    End If
BegnieWork:

    RsSavRec5.MoveLast
    FiLLTXT5
    Exit Sub

ErrTrap:
    Select Case Err.Number
        Case -2147217885
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
                Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
                Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
                Msg = "Sorry , this Recored was deleted by other user on the network" & CHR(13)
                Msg = Msg & "Date will be updated now" & CHR(13)
            End If
                MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec5.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub btnModify5_Click()

    Dim Msg As String

    If DoPremis(Do_Edit, Me.Name, True) = False Then
        Exit Sub
    End If

    On Error GoTo ErrTrap

    If TxtVac_ID5.Text <> "" Then
        TxtModFlg5 = "E"
        Frm25.Enabled = True
        Me.TxtVacName5.SetFocus
    End If
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147467259
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "⁄›Ê«" & CHR(13)
                Msg = Msg & " ·«Ì„ﬂ‰  ⁄œÌ· Â–« «·”Ã· ›Ï «·Êﬁ  «·Õ«·Ï" & CHR(13)
                Msg = Msg & "ÕÌÀ «‰Â ﬁÌœ «· ⁄œÌ· „‰ ﬁ»· „” Œœ„ «Œ— ⁄·Ï «·‘»ﬂ…"
            Else
                Msg = "Sorry" & CHR(13)
                Msg = Msg & "This recored can't be edited now" & CHR(13)
                Msg = Msg & "it's under modification by other user on the network"
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            If RsSavRec5.EditMode <> adEditNone Then
                RsSavRec5.CancelUpdate
            End If
    End Select
End Sub
Private Sub btnNew5_Click()

    Dim My_SQL As String
    Dim rs As ADODB.Recordset

    If DoPremis(Do_New, Me.Name, True) = False Then
        Exit Sub
    End If

    On Error GoTo ErrTrap

    Set rs = New ADODB.Recordset
    
    Frm25.Enabled = True
    clear_all Me
    TxtModFlg5.Text = "N"
    My_SQL = "Gbasic"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable

    If rs.RecordCount > 0 Then
        TxtSerial5.Text = rs.RecordCount + 1
    Else
        TxtSerial5.Text = 1
    End If
    rs.Close
    CmbType.ListIndex = 0
    TxtVacName5.SetFocus
ErrTrap:
End Sub
Private Sub btnNext5_Click()

    On Error GoTo ErrTrap
    
    Dim Msg As String

    If Me.TxtModFlg5.Text = "N" Then
        FindRec5 val(TxtVac_ID5.Text)
        Me.TxtModFlg5.Text = "R"
    End If

    TxtModFlg5 = "R"
    If RsSavRec5.RecordCount = 0 Then
        clear_all Me
        Exit Sub
    End If
BegnieWork:
    If RsSavRec5.EOF Then
        RsSavRec5.MoveLast
    Else
        RsSavRec5.MoveNext
        If RsSavRec5.EOF Then
            RsSavRec5.MoveLast
        End If
    End If
    FiLLTXT5
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
                Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
                Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
                Msg = "Sorry , this Recored was deleted by other user on the network" & CHR(13)
                Msg = Msg & "Date will be updated now" & CHR(13)
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec5.Requery
            Resume BegnieWork
    End Select
End Sub

Private Sub btnPrevious5_Click()

    On Error GoTo ErrTrap
    
    Dim Msg As String
    If Me.TxtModFlg5.Text = "N" Then
        FindRec5 val(TxtVac_ID5.Text)
        Me.TxtModFlg5.Text = "R"
    End If
    TxtModFlg5 = "R"
    If RsSavRec5.RecordCount = 0 Then
        clear_all Me
        Exit Sub
    End If
BegnieWork:
    RsSavRec5.MovePrevious
    If RsSavRec5.BOF Then
        RsSavRec5.MoveFirst
    End If
    FiLLTXT5
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
                Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
                Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
                Msg = "Sorry , this Recored was deleted by other user on the network" & CHR(13)
                Msg = Msg & "Date will be updated now" & CHR(13)
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec5.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub btnSave5_Click()

    On Error GoTo ErrTrap
    
    Dim Msg As String
    Dim StrVacCode As String
    Dim StrVacName As String
    Dim CtrlTxt As Control
    
    For Each CtrlTxt In Me.Controls
        If TypeOf CtrlTxt Is TextBox Or TypeOf CtrlTxt Is ComboBox Then
            If CtrlTxt.Text = "" And CtrlTxt.Tag <> "" And CtrlTxt.Enabled = True Then
                MsgBox CtrlTxt.Tag, vbOKOnly + vbMsgBoxRight, App.title
                CtrlTxt.SetFocus
                Exit Sub
            End If
        End If

    Next
    StrVacName = IsRecExist("Gbasic", "name", Trim(TxtVacName5.Text), "name", "Vac_ID<>'" & Trim(TxtVac_ID5.Text) & "'")
    If StrVacName <> "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·ﬁœ ”»ﬁ  ”ÃÌ· Â–« «·‰Ê⁄ „‰ ﬁ»·"
        Else
            Msg = "This record already exists"
        End If
        MsgBox Msg, vbOKOnly + vbMsgBoxRight, App.title
        TxtVacName5.SetFocus
        Exit Sub
    End If
    Select Case Me.TxtModFlg5.Text
        Case "N"
            AddNewRec5
            btnLast5_Click

        Case "E"
            FiLLRec5
    End Select

    Exit Sub
ErrTrap:
    If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "Â‰«ﬂ Œÿ√ „« ›Ì ≈œŒ«· «·»Ì«‰« ", vbOKOnly + vbMsgBoxRight, App.title
    Else
        MsgBox "Something went wrong while inserting data", vbOKOnly + vbMsgBoxRight, App.title
    End If
End Sub
Private Sub BtnUndo5_Click()
    FindRec5 val(TxtVac_ID5.Text)
    Me.TxtModFlg5.Text = "R"
End Sub
Private Sub BtnUpdate5_Click()

    On Error GoTo ErrTrap
    
    Dim Msg As String
    Dim FristCount As Long
    Dim LastCount As Long
    
    FristCount = RsSavRec5.RecordCount
    RsSavRec5.Requery
    LastCount = RsSavRec5.RecordCount
    BtnUndo5_Click
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
            Msg = "No. of records before the update" & vbCrLf & FristCount & vbCrLf & "⁄œœ «·”Ã·«  »⁄œ «· ÕœÌÀ" & vbCrLf & LastCount
        
            If LastCount > FristCount Then
                Msg = Msg + vbCrLf & "No. of new records" & vbCrLf & LastCount - FristCount
            Else
                Msg = Msg + vbCrLf & "No. of deleted records" & vbCrLf & FristCount - LastCount
            End If
        End If
    End If
    MsgBox Msg, vbOKOnly + vbMsgBoxRight + vbInformation, App.title
ErrTrap:
End Sub
Public Sub AddNewRec5()
    On Error GoTo ErrTrap
    Dim StrRecId5 As String
    StrRecId5 = new_id("Gbasic", "id", "")
    RsSavRec5.AddNew
    RsSavRec5.Fields("id").value = IIf(StrRecId5 <> "", StrRecId5, Null)
    FiLLRec5
ErrTrap:
End Sub
Public Sub FiLLRec5()

    On Error GoTo ErrTrap

    RsSavRec5.Fields("name").value = IIf(TxtVacName5.Text <> "", Trim(TxtVacName5.Text), Null)
    RsSavRec5.Fields("namee").value = IIf(TxtVacNamee5.Text <> "", Trim(TxtVacNamee5.Text), Null)
    RsSavRec5.update
    If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox " „  ⁄„·Ì… «·Õ›Ÿ »‰Ã«Õ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Else
        MsgBox "Record saved successfully", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    End If
    FillGrid5WithData
    TxtModFlg5 = "R"
    Exit Sub
ErrTrap:

    If RsSavRec5.EditMode <> adEditNone Then
        RsSavRec5.CancelUpdate
    End If
End Sub
Public Sub FiLLTXT5()

    On Error GoTo ErrTrap
    
    Dim i As Integer
    
    Frm25.Enabled = False
    TxtVac_ID5.Text = IIf(IsNull(RsSavRec5.Fields("id").value), "", RsSavRec5.Fields("id").value)
    TxtVacName5.Text = IIf(IsNull(RsSavRec5.Fields("name").value), "", RsSavRec5.Fields("name").value)
    TxtVacNamee5.Text = IIf(IsNull(RsSavRec5.Fields("namee").value), "", RsSavRec5.Fields("namee").value)
    LabCurrRec5.Caption = RsSavRec5.AbsolutePosition
    LabCountRec5.Caption = RsSavRec5.RecordCount
    With Grid5
        For i = 1 To .Rows - 1
            If Trim(TxtVac_ID5.Text) = .TextMatrix(i, .ColIndex("id")) Then
                TxtSerial5.Text = .TextMatrix(i, .ColIndex("Ser"))
                .Row = i
                Exit Sub
            End If
        Next
    End With
ErrTrap:
End Sub
Public Sub EditRec5(StrTable As String, RecId5 As String)
    FiLLRec5
End Sub

Private Sub Grid5_EnterCell()
    On Error GoTo ErrTrap
    FindRec5 val(Me.Grid5.TextMatrix(Me.Grid5.Row, Me.Grid5.ColIndex("id")))
ErrTrap:
End Sub
Private Sub TxtVac_ID5_Change()
    Dim TxtMod As String
    TxtMod = TxtModFlg5.Text
    TxtModFlg5.Text = ""
    TxtModFlg5 = TxtMod
End Sub
Public Function FindRec5(ByVal RecId5 As Long)

    On Error GoTo ErrTrap
    
    RsSavRec5.Find "id=" & RecId5, , adSearchForward, 1
    If Not (RsSavRec5.EOF) Then
        FiLLTXT5
    End If
    Exit Function
ErrTrap:
    If RsSavRec5.EditMode <> adEditNone Then
        RsSavRec5.CancelUpdate
        BtnUndo5_Click
    End If
End Function
Private Sub TxtModFlg5_Change()
    If TxtModFlg5.Text = "N" Then
        Frm25.Enabled = True
        Me.btnNew5.Enabled = False
        btnModify5.Enabled = False
        btnDelete5.Enabled = False
        Me.btnQuery.Enabled = False
        Grid5.Enabled = False
        BtnUndo5.Enabled = True
        Me.btnSave5.Enabled = True
        BtnUpdate5.Enabled = False
    ElseIf TxtModFlg5.Text = "R" Then
        Frm25.Enabled = False
        Grid5.Enabled = True
        btnModify5.Enabled = False
        btnDelete5.Enabled = False
        If TxtVac_ID5.Text <> "" Then
            btnModify5.Enabled = True
            btnDelete5.Enabled = True
        End If
        BtnUpdate5.Enabled = True
        Me.btnQuery.Enabled = True
        Me.btnNew5.Enabled = True
        BtnUndo5.Enabled = False
        Me.btnSave5.Enabled = False
    
        btnNext5.Enabled = True
        btnPrevious5.Enabled = True
        btnFirst5.Enabled = True
        btnLast5.Enabled = True
    ElseIf TxtModFlg5.Text = "E" Then
        Frm25.Enabled = True
        Me.btnNew5.Enabled = False
        btnModify5.Enabled = False
        btnDelete5.Enabled = False
        Me.btnQuery.Enabled = False
        BtnUpdate5.Enabled = False
        BtnUndo5.Enabled = True
        Me.btnSave5.Enabled = True
        Grid5.Enabled = False
        btnNext5.Enabled = False
        btnPrevious5.Enabled = False
        btnFirst5.Enabled = False
        btnLast5.Enabled = False
    
    End If
End Sub
Public Sub FillGrid5WithData()

    On Error GoTo ErrTrap

    Dim i As Integer
    Dim rs As ADODB.Recordset
    Dim My_SQL As String

    Set rs = New ADODB.Recordset
    My_SQL = "select * From Gbasic order by id"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText
    With Me.Grid5
        .Rows = 2
        .Clear flexClearScrollable
        If rs.RecordCount > 0 Then
            .Rows = rs.RecordCount + 1
            rs.MoveFirst

            For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("Ser")) = i
                .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(rs.Fields("name").value), "", rs.Fields("name").value)
                .TextMatrix(i, .ColIndex("namee")) = IIf(IsNull(rs.Fields("namee").value), "", rs.Fields("namee").value)
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs.Fields("id").value), "", rs.Fields("id").value)
                rs.MoveNext
            Next
            rs.Close
        End If
        .RowHeight(-1) = 300
    End With
ErrTrap:
End Sub
Private Function CheckDelCountry5(Lngid As Long) As Boolean

    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    
    StrSQL = "Select * From TblEmployee Where id=" & Lngid & ""
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If Not (rs.BOF Or rs.EOF) Then
        CheckDelCountry5 = False
    Else
        CheckDelCountry5 = True
    End If
    rs.Close
    Set rs = Nothing
End Function
Private Sub TxtVacName5_GotFocus()
    SwitchKeyboardLang LANG_ARABIC
End Sub
Private Sub TxtVacNamee5_GotFocus()
    SwitchKeyboardLang LANG_ENGLISH
End Sub
'################################################################################################################################################
'################################################################################################################################################
'################################################################################################################################################
Private Sub btnCancel6_Click()
    Unload Me
End Sub
Function CuurentLogdata6(Optional Currentmode As String)
     LogTextA = "    ‘«‘… " & ScreenNameArabic & CHR(13) & "ﬂÊœ  «·„‰œÊ»" & txtcode.Text & CHR(13) & "   «”„ «·„‰œÊ» " & DcEmp & CHR(13) & "   ‰”»… «·Œ’„ " & TXTDiscounts & CHR(13) & "   «·›—⁄ " & DcBranches & CHR(13) & " «·„Ã„Ê⁄Â " & DCSalesRepGroups
        LogTextE = "    Screen  " & ScreenNameEnglish & CHR(13) & " Sales Person Code" & txtcode.Text & CHR(13) & "    Sales Person Name   " & DcEmp & CHR(13) & "  Discounts" & TXTDiscounts & CHR(13) & "   Branch " & DcBranches & CHR(13) & "  Group " & DCSalesRepGroups
       If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTextE, Me.Name, Me.TxtModFlg6
    Else
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTextE, Me.Name, "D"
    End If
    
End Function
Private Sub btnDelete6_Click()

    Dim MSGType As Integer
    Dim BlnRecExist As Boolean
    Dim StrMSG  As String
    Dim Msg As String
    
    On Error GoTo ErrTrap

    If DoPremis(Do_Delete, Me.Name, True) = False Then
        Exit Sub
    End If

    If TxtVac_ID6.Text <> "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MSGType = MsgBox("Â·  —€» ›Ì Õ–› Â–« «·”Ã·", vbQuestion + vbMsgBoxRtlReading + vbYesNo + vbMsgBoxRight, App.title)
        Else
            MSGType = MsgBox("Do you want to delete this record", vbQuestion + vbMsgBoxRtlReading + vbYesNo + vbMsgBoxRight, App.title)
        End If
        If MSGType = vbYes Then
            RsSavRec6.Find "id=" & val(TxtVac_ID6.Text), , adSearchForward, 1
            CuurentLogdata6 ("D")
            RsSavRec6.delete
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox " „  ⁄„·Ì… «·Õ–› »‰Ã«Õ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
            Else
                MsgBox "Record deleted successfully", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            End If
            FillGrid6WithData
            btnNext6_Click
        End If
    End If
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217873, -2147467259
            If SystemOptions.UserInterface = ArabicInterface Then
                StrMSG = "⁄›Ê« ·« ÌÃÊ“ Õ–› «·”Ã· ·«— »«ÿÂ »»Ì«‰«  √Œ—Ì"
            Else
                StrMSG = "sorry, this record cannot be deleted due to data integration"
            End If
            RsSavRec6.CancelUpdate
            MsgBox StrMSG, vbOKOnly + vbMsgBoxRight, App.title
            Cn.Errors.Clear
    End Select
End Sub
Private Sub btnFirst6_Click()

    On Error GoTo ErrTrap

    Dim Msg As String
    
    If Me.TxtModFlg6.Text = "N" Then
        FindRec6 val(TxtVac_ID6.Text)
        Me.TxtModFlg6.Text = "R"
    End If
    TxtModFlg6 = "R"
    If RsSavRec6.RecordCount = 0 Then
        clear_all Me
        Exit Sub
    End If
BegnieWork:
    RsSavRec6.MoveFirst
    FiLLTXT6
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
                Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
                Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
                Msg = "Sorry , this Recored was deleted by other user on the network" & CHR(13)
                Msg = Msg & "Date will be updated now" & CHR(13)
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec6.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub btnLast6_Click()

    On Error GoTo ErrTrap

    Dim Msg As String

    If Me.TxtModFlg6.Text = "N" Then
        FindRec6 val(TxtVac_ID6.Text)
        Me.TxtModFlg6.Text = "R"
    End If
    TxtModFlg6 = "R"
    If RsSavRec6.RecordCount = 0 Then
        clear_all Me
        Exit Sub
    End If
BegnieWork:
    RsSavRec6.MoveLast
    FiLLTXT6
    Exit Sub

ErrTrap:
    Select Case Err.Number
        Case -2147217885
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
                Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
                Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
                Msg = "Sorry , this Recored was deleted by other user on the network" & CHR(13)
                Msg = Msg & "Date will be updated now" & CHR(13)
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec6.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub btnModify6_Click()

    Dim Msg As String

    If DoPremis(Do_Edit, Me.Name, True) = False Then
        Exit Sub
    End If

    On Error GoTo ErrTrap

    If TxtVac_ID6.Text <> "" Then
        TxtModFlg6 = "E"
        Frm26.Enabled = True
        Me.TXTDiscounts.SetFocus
        CuurentLogdata6
    End If
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147467259
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "⁄›Ê«" & CHR(13)
                Msg = Msg & " ·«Ì„ﬂ‰  ⁄œÌ· Â–« «·”Ã· ›Ï «·Êﬁ  «·Õ«·Ï" & CHR(13)
                Msg = Msg & "ÕÌÀ «‰Â ﬁÌœ «· ⁄œÌ· „‰ ﬁ»· „” Œœ„ «Œ— ⁄·Ï «·‘»ﬂ…"
            Else
                Msg = "Sorry , this Recored was deleted by other user on the network" & CHR(13)
                Msg = Msg & "Date will be updated now" & CHR(13)
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            If RsSavRec6.EditMode <> adEditNone Then
                RsSavRec6.CancelUpdate
            End If
    End Select
End Sub
Private Sub btnNew6_Click()

    Dim My_SQL As String
    Dim rs As ADODB.Recordset

    If DoPremis(Do_New, Me.Name, True) = False Then
        Exit Sub
    End If

    On Error GoTo ErrTrap
    
    Set rs = New ADODB.Recordset
    Frm26.Enabled = True

    Me.TxtVac_ID6.Text = ""
    Me.TXTDiscounts.Text = ""
    Me.DcBranches.BoundText = ""
    Me.DcEmp.BoundText = ""
    Me.DCJob.BoundText = ""
    Me.DCSalesRepGroups.BoundText = ""
 
    TxtModFlg6.Text = "N"

    My_SQL = "TBLSalesRepData2"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable

    If rs.RecordCount > 0 Then
        TxtSerial.Text = rs.RecordCount + 1
    Else
        TxtSerial.Text = 1
    End If

    rs.Close
    CmbType.ListIndex = 0

    TXTDiscounts.SetFocus
    TXTDiscounts.Text = 0
ErrTrap:
End Sub
Private Sub btnNext6_Click()

    On Error GoTo ErrTrap
    
    Dim Msg As String

    If Me.TxtModFlg6.Text = "N" Then
        FindRec6 val(TxtVac_ID6.Text)
        Me.TxtModFlg6.Text = "R"
    End If
    TxtModFlg6 = "R"
    If RsSavRec6.RecordCount = 0 Then
        clear_all Me
        Exit Sub
    End If
BegnieWork:
    If RsSavRec6.EOF Then
        RsSavRec6.MoveLast
    Else
        RsSavRec6.MoveNext
        If RsSavRec6.EOF Then
            RsSavRec6.MoveLast
        End If
    End If
    FiLLTXT6
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
                Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
                Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
                Msg = "Sorry , this Recored was deleted by other user on the network" & CHR(13)
                Msg = Msg & "Date will be updated now" & CHR(13)
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec6.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub btnPrevious6_Click()

    On Error GoTo ErrTrap
    
    Dim Msg As String

    If Me.TxtModFlg6.Text = "N" Then
        FindRec6 val(TxtVac_ID6.Text)
        Me.TxtModFlg6.Text = "R"
    End If
    TxtModFlg6 = "R"
    If RsSavRec6.RecordCount = 0 Then
        clear_all Me
        Exit Sub
    End If
BegnieWork:
    RsSavRec6.MovePrevious
    If RsSavRec6.BOF Then
        RsSavRec6.MoveFirst
    End If

    FiLLTXT6
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
                Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
                Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
                Msg = "Sorry , this Recored was deleted by other user on the network" & CHR(13)
                Msg = Msg & "Date will be updated now" & CHR(13)
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec6.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub btnQuery6_Click()
    Load FrmSearchSales1
    FrmSearchSales1.show
End Sub
Private Sub btnSave6_Click()

    On Error GoTo ErrTrap
    
    Dim Msg As String
    Dim StrVacCode As String
    Dim StrVacName As String
    Dim CtrlTxt As Control
    
    If txtcode.Text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "«·—Ã«¡ «œŒ«· «·ﬂÊœ", vbOKOnly + vbMsgBoxRight, App.title
        Else
            MsgBox "Please Enter The Code", vbOKOnly + vbMsgBoxRight, App.title
        End If
        Exit Sub
    End If
    
    If TXTDiscounts.Text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "«·—Ã«¡ «œŒ«· ‰”»… «·Œ’„", vbOKOnly + vbMsgBoxRight, App.title
        Else
            MsgBox "Please Enter The discount percentage", vbOKOnly + vbMsgBoxRight, App.title
        End If
        Exit Sub
    End If
    
    If val(DcEmp.BoundText) = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "«·—Ã«¡ «œŒ«· «·„ÊŸ›", vbOKOnly + vbMsgBoxRight, App.title
        Else
            MsgBox "Please chose the employee", vbOKOnly + vbMsgBoxRight, App.title
        End If
        Exit Sub
    End If
    
    'If val(DcBranches.BoundText) = 0 Then
    '    If SystemOptions.UserInterface = ArabicInterface Then
    '        MsgBox "«·—Ã«¡ «œŒ«· «·›—⁄", vbOKOnly + vbMsgBoxRight, App.title
    '    Else
    '        MsgBox "Please chose the branch", vbOKOnly + vbMsgBoxRight, App.title
    '    End If
    '    Exit Sub
    'End If
    
    If val(DCSalesRepGroups.BoundText) = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "«·—Ã«¡ «œŒ«· «·„Ã„Ê⁄…", vbOKOnly + vbMsgBoxRight, App.title
        Else
            MsgBox "Please chose the group", vbOKOnly + vbMsgBoxRight, App.title
        End If
        Exit Sub
    End If
    
    Select Case Me.TxtModFlg6.Text
        Case "N"
            AddNewRec6
            btnLast6_Click

        Case "E"
            FiLLRec6
    End Select

    Exit Sub
ErrTrap:
    If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "Â‰«ﬂ Œÿ√ „« ›Ì ≈œŒ«· «·»Ì«‰« ", vbOKOnly + vbMsgBoxRight, App.title
    Else
        MsgBox "Something went wrong while inserting data", vbOKOnly + vbMsgBoxRight, App.title
    End If
End Sub
Private Sub BtnUndo6_Click()
    FindRec6 val(TxtVac_ID6.Text)
    Me.TxtModFlg6.Text = "R"
End Sub
Private Sub BtnUpdate6_Click()

    On Error GoTo ErrTrap
    
    Dim Msg As String
    Dim FristCount As Long
    Dim LastCount As Long
    
    FristCount = RsSavRec6.RecordCount
    RsSavRec6.Requery
    LastCount = RsSavRec6.RecordCount
    BtnUndo6_Click
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
            Msg = "No. of records before the update" & vbCrLf & FristCount & vbCrLf & "⁄œœ «·”Ã·«  »⁄œ «· ÕœÌÀ" & vbCrLf & LastCount
        
            If LastCount > FristCount Then
                Msg = Msg + vbCrLf & "No. of new records" & vbCrLf & LastCount - FristCount
            Else
                Msg = Msg + vbCrLf & "No. of deleted records" & vbCrLf & FristCount - LastCount
            End If
        End If
    End If
    MsgBox Msg, vbOKOnly + vbMsgBoxRight + vbInformation, App.title
ErrTrap:
End Sub
Private Sub dcEmp_Change()
    If val(Me.DcEmp.BoundText) = 0 Then Exit Sub
    Me.txtcode.Text = get_EMPLOYEE_Data(val(Me.DcEmp.BoundText), "Emp_Code")
End Sub
Private Sub DCEmP_KeyUp(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF3 Then
        FrmEmployeeSearch.lbltype = 32
        FrmEmployeeSearch.show
    End If
End Sub
Function print_report(Optional NoteSerial As String, Optional X As Integer)
    
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
    
    If chPrinet = 0 Then
        MySQL = " SELECT dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee, dbo.TblBranchesData.branch_name, "
        MySQL = MySQL & " dbo.TblBranchesData.branch_namee, dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmpJobsTypes.JobTypeNamee, dbo.TblEmployee.Emp_ID,"
        MySQL = MySQL & " dbo.TBLSalesRepData2.id, dbo.TblBranchesData.branch_id, dbo.TBLSalesRepGroups2.name, dbo.TBLSalesRepGroups2.namee, dbo.TblEmpJobsTypes.JobTypeID,"
        MySQL = MySQL & " dbo.TBLSalesRepData2.discountvalue"
        MySQL = MySQL & " FROM dbo.TblEmployee RIGHT OUTER JOIN"
        MySQL = MySQL & " dbo.TBLSalesRepData2 LEFT OUTER JOIN"
        MySQL = MySQL & " dbo.TblEmpJobsTypes ON dbo.TBLSalesRepData2.JobID = dbo.TblEmpJobsTypes.JobTypeID LEFT OUTER JOIN"
        MySQL = MySQL & " dbo.TBLSalesRepGroups2 ON dbo.TBLSalesRepData2.GroupID = dbo.TBLSalesRepGroups2.id LEFT OUTER JOIN"
        MySQL = MySQL & " dbo.TblBranchesData ON dbo.TBLSalesRepData2.BranchId = dbo.TblBranchesData.branch_id ON dbo.TblEmployee.Emp_ID = dbo.TBLSalesRepData2.EmpID"
  
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "repSalse1.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "repSalse1.rpt"
        End If
    Else
        If chPrinet = 1 Then
            MySQL = " SELECT dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee, dbo.TblBranchesData.branch_name, "
            MySQL = MySQL & " dbo.TblBranchesData.branch_namee, dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmpJobsTypes.JobTypeNamee, dbo.TblEmployee.Emp_ID,"
            MySQL = MySQL & " dbo.TBLSalesRepData2.id, dbo.TblBranchesData.branch_id, dbo.TBLSalesRepGroups2.name, dbo.TBLSalesRepGroups2.namee, dbo.TblEmpJobsTypes.JobTypeID,"
            MySQL = MySQL & " dbo.TBLSalesRepData2.DiscountValue, dbo.TBLSalesRepGroups2.id AS IDgroup"
            MySQL = MySQL & " FROM dbo.TblEmployee RIGHT OUTER JOIN"
            MySQL = MySQL & " dbo.TBLSalesRepData2 LEFT OUTER JOIN"
            MySQL = MySQL & " dbo.TblEmpJobsTypes ON dbo.TBLSalesRepData2.JobID = dbo.TblEmpJobsTypes.JobTypeID LEFT OUTER JOIN"
            MySQL = MySQL & " dbo.TBLSalesRepGroups2 ON dbo.TBLSalesRepData2.GroupID = dbo.TBLSalesRepGroups2.id LEFT OUTER JOIN"
            MySQL = MySQL & " dbo.TblBranchesData ON dbo.TBLSalesRepData2.BranchId = dbo.TblBranchesData.branch_id ON dbo.TblEmployee.Emp_ID = dbo.TBLSalesRepData2.EmpID"
            MySQL = MySQL & " Where (dbo.TBLSalesRepGroups2.id =" & Me.DCSalesRepGroups.BoundText & ")"

            If SystemOptions.UserInterface = ArabicInterface Then
                StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "repSalseGroup 1.rpt"
            Else
                StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "repSalseGroup 1.rpt"
            End If
        Else
            MySQL = " SELECT dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee, dbo.TblBranchesData.branch_name, "
            MySQL = MySQL & " dbo.TblBranchesData.branch_namee, dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmpJobsTypes.JobTypeNamee, dbo.TblEmployee.Emp_ID,"
            MySQL = MySQL & " dbo.TBLSalesRepData2.id, dbo.TblBranchesData.branch_id, dbo.TBLSalesRepGroups2.name, dbo.TBLSalesRepGroups2.namee, dbo.TblEmpJobsTypes.JobTypeID,"
            MySQL = MySQL & " dbo.TBLSalesRepData2.DiscountValue, dbo.TBLSalesRepGroups2.id AS IDgroup"
            MySQL = MySQL & " FROM dbo.TblEmployee RIGHT OUTER JOIN"
            MySQL = MySQL & " dbo.TBLSalesRepData2 LEFT OUTER JOIN"
            MySQL = MySQL & " dbo.TblEmpJobsTypes ON dbo.TBLSalesRepData2.JobID = dbo.TblEmpJobsTypes.JobTypeID LEFT OUTER JOIN"
            MySQL = MySQL & " dbo.TBLSalesRepGroups2 ON dbo.TBLSalesRepData2.GroupID = dbo.TBLSalesRepGroups2.id LEFT OUTER JOIN"
            MySQL = MySQL & " dbo.TblBranchesData ON dbo.TBLSalesRepData2.BranchId = dbo.TblBranchesData.branch_id ON dbo.TblEmployee.Emp_ID = dbo.TBLSalesRepData2.EmpID"
            MySQL = MySQL & " Where (dbo.TblBranchesData.branch_id =" & Me.DcBranches.BoundText & ")"
 
            If SystemOptions.UserInterface = ArabicInterface Then
                StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "repSalseBranch1.rpt"
            Else
                StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "repSalseBranch1.rpt"
            End If
        End If
    End If
    
    If Dir(StrFileName) = "" Then
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        Else
            Msg = "There's no data to show"
        End If
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
        StrReportTitle = "" '& StrAccountName
    Else
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name

    Dim Total As String
    Dim dif As String
    Dim totl As Double

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
Public Sub AddNewRec6()

    On Error GoTo ErrTrap
    
    Dim StrRecId6 As String
    
    StrRecId6 = new_id("TBLSalesRepData2", "id", "")
    RsSavRec6.AddNew
    RsSavRec6.Fields("id").value = IIf(StrRecId6 <> "", StrRecId6, Null)
    FiLLRec6
ErrTrap:
End Sub
Public Sub FiLLRec6()

    On Error GoTo ErrTrap
    
    RsSavRec6.Fields("DiscountValue").value = IIf(IsNumeric(TXTDiscounts.Text), val(TXTDiscounts.Text), Null)
    RsSavRec6.Fields("EmpID").value = IIf(Me.DcEmp.BoundText <> 0, val(Me.DcEmp.BoundText), Null)
    RsSavRec6.Fields("BranchId").value = IIf(val(Me.DcBranches.BoundText) <> 0, val(Me.DcBranches.BoundText), Null)
    RsSavRec6.Fields("GroupID").value = IIf(Me.DCSalesRepGroups.BoundText <> 0, val(Me.DCSalesRepGroups.BoundText), Null)
    RsSavRec6.update
    CuurentLogdata6
    If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox " „  ⁄„·Ì… «·Õ›Ÿ »‰Ã«Õ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Else
        MsgBox "Record saved successfully", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    End If
    FillGrid6WithData
    TxtModFlg6 = "R"
    Exit Sub
ErrTrap:
    If RsSavRec6.EditMode <> adEditNone Then
        RsSavRec6.CancelUpdate
    End If
End Sub
Public Sub FiLLTXT6()

    On Error GoTo ErrTrap
    
    Dim i As Integer
    
    Frm26.Enabled = False
    TxtVac_ID6.Text = IIf(IsNull(RsSavRec6.Fields("id").value), "", RsSavRec6.Fields("id").value)
    TXTDiscounts.Text = IIf(IsNull(RsSavRec6.Fields("DiscountValue").value), 0, RsSavRec6.Fields("DiscountValue").value)
    Me.DcEmp.BoundText = IIf(IsNull(RsSavRec6.Fields("EmpID").value), "", RsSavRec6.Fields("EmpID").value)
    Me.DcBranches.BoundText = IIf(IsNull(RsSavRec6.Fields("BranchId").value), "", RsSavRec6.Fields("BranchId").value)
    Me.DCSalesRepGroups.BoundText = IIf(IsNull(RsSavRec6.Fields("GroupID").value), "", RsSavRec6.Fields("GroupID").value)
    LabCurrRec6.Caption = RsSavRec6.AbsolutePosition
    LabCountRec6.Caption = RsSavRec6.RecordCount
    With Grid6
        For i = 1 To .Rows - 1
            If Trim(TxtVac_ID6.Text) = .TextMatrix(i, .ColIndex("id")) Then
                TxtSerial.Text = .TextMatrix(i, .ColIndex("Ser"))
                .Row = i
                Exit Sub
            End If
        Next
    End With
ErrTrap:
End Sub
Public Sub EditRec6(StrTable As String, RecId6 As String)
    FiLLRec6
End Sub
Private Sub Grid6_EnterCell()
    On Error GoTo ErrTrap
    FindRec6 val(Me.Grid6.TextMatrix(Me.Grid6.Row, Me.Grid6.ColIndex("EmpID")))
ErrTrap:
End Sub
Private Sub ISButton1_Click()
    chPrinet = 0
    print_report
End Sub
Private Sub ISButton2_Click()
    chPrinet = 1
    print_report
End Sub
Private Sub ISButton3_Click()
    chPrinet = 2
    print_report
End Sub
Private Sub TxtVac_ID6_Change()
    Dim TxtMod As String
    TxtMod = TxtModFlg6.Text
    TxtModFlg6.Text = ""
    TxtModFlg6 = TxtMod
End Sub
Public Function FindRec6(ByVal RecId6 As Long)

    On Error GoTo ErrTrap
    
    RsSavRec6.Find "EmpID=" & RecId6, , adSearchForward, 1
    If Not (RsSavRec6.EOF) Then
        FiLLTXT6
    End If
    Exit Function
ErrTrap:
    If RsSavRec6.EditMode <> adEditNone Then
        RsSavRec6.CancelUpdate
        BtnUndo6_Click
    End If
End Function
Private Sub TxtModFlg6_Change()
    If TxtModFlg6.Text = "N" Then
        Frm26.Enabled = True
        Me.btnNew6.Enabled = False
        btnModify6.Enabled = False
        btnDelete6.Enabled = False
        Me.btnQuery6.Enabled = False
        Grid6.Enabled = False
        BtnUndo6.Enabled = True
        Me.btnSave6.Enabled = True
        BtnUpdate6.Enabled = False
    ElseIf TxtModFlg6.Text = "R" Then
        Frm26.Enabled = False
        Grid6.Enabled = True
        btnModify6.Enabled = False
        btnDelete6.Enabled = False
        If TxtVac_ID6.Text <> "" Then
            btnModify6.Enabled = True
            btnDelete6.Enabled = True
        End If
        BtnUpdate6.Enabled = True
        Me.btnQuery6.Enabled = True
        Me.btnNew6.Enabled = True
        BtnUndo6.Enabled = False
        Me.btnSave6.Enabled = False
        btnNext6.Enabled = True
        btnPrevious6.Enabled = True
        btnFirst6.Enabled = True
        btnLast6.Enabled = True
    ElseIf TxtModFlg6.Text = "E" Then
        Frm26.Enabled = True
        Me.btnNew6.Enabled = False
        btnModify6.Enabled = False
        btnDelete6.Enabled = False
        Me.btnQuery6.Enabled = False
        BtnUpdate6.Enabled = False
        BtnUndo6.Enabled = True
        Me.btnSave6.Enabled = True
        Grid6.Enabled = False
        btnNext6.Enabled = False
        btnPrevious6.Enabled = False
        btnFirst6.Enabled = False
        btnLast6.Enabled = False
    End If
End Sub
Public Sub FillGrid6WithData()

    On Error GoTo ErrTrap

    Dim i As Integer
    Dim rs As ADODB.Recordset
    Dim My_SQL As String

    Set rs = New ADODB.Recordset
    My_SQL = "select * From TBLSalesRepData2 order by id"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText
    With Me.Grid6
        .Rows = 2
        .Clear flexClearScrollable
        If rs.RecordCount > 0 Then
            .Rows = rs.RecordCount + 1
            rs.MoveFirst
            For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("Ser")) = i
                .TextMatrix(i, .ColIndex("EmpID")) = IIf(IsNull(rs.Fields("EmpID").value), "", rs.Fields("EmpID").value)
                .TextMatrix(i, .ColIndex("EmpCode")) = get_EMPLOYEE_Data(val(.TextMatrix(i, .ColIndex("EmpID"))), "Emp_Code")
                .TextMatrix(i, .ColIndex("EmpName")) = IIf(IsNull(rs.Fields("EmpID").value), "", rs.Fields("EmpID").value)
                .TextMatrix(i, .ColIndex("BranchId")) = IIf(IsNull(rs.Fields("BranchId").value), "", rs.Fields("BranchId").value)
                .TextMatrix(i, .ColIndex("GroupID")) = IIf(IsNull(rs.Fields("GroupID").value), "", rs.Fields("GroupID").value)
                .TextMatrix(i, .ColIndex("JobID")) = IIf(IsNull(rs.Fields("JobID").value), "", rs.Fields("JobID").value)
                .TextMatrix(i, .ColIndex("DiscountValue")) = IIf(IsNull(rs.Fields("DiscountValue").value), "", rs.Fields("DiscountValue").value)
                rs.MoveNext
            Next
            rs.Close
        End If
        .RowHeight(-1) = 300
    End With
ErrTrap:
End Sub
'#######################################################################################################################################################
'#######################################################################################################################################################
'#######################################################################################################################################################
Private Sub btnCancel7_Click()
    Unload Me
End Sub
Function CuurentLogdata7(Optional Currentmode As String)
     LogTextA = "    ‘«‘… " & ScreenNameArabic & CHR(13) & "ﬂÊœ  «·„‰œÊ»" & TXTCode7.Text & CHR(13) & "   «”„ «·„‰œÊ» " & DCEmP7 & CHR(13) & "   ‰”»… «·Œ’„ " & TXTDiscounts7 & CHR(13) & "   «·›—⁄ " & DcBranches7 & CHR(13) & " «·„Ã„Ê⁄Â " & DCSalesRepGroups7
        LogTextE = "    Screen  " & ScreenNameEnglish & CHR(13) & " Sales Person Code" & TXTCode7.Text & CHR(13) & "    Sales Person Name   " & DCEmP7 & CHR(13) & "  Discounts" & TXTDiscounts7 & CHR(13) & "   Branch " & DcBranches7 & CHR(13) & "  Group " & DCSalesRepGroups7
       If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTextE, Me.Name, Me.TxtModFlg7
    Else
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTextE, Me.Name, "D"
    End If
End Function
Private Sub btnDelete7_Click()

    Dim MSGType As Integer
    Dim BlnRecExist As Boolean
    Dim StrMSG  As String
    Dim Msg As String
    Dim Emp_id As Integer
    Dim StrSQL As String
    Dim RsTemp As New ADODB.Recordset
    
    On Error GoTo ErrTrap

    Emp_id = val(DCEmP7.BoundText)
 
    
    StrSQL = "SELECT  Emp_id  FROM         dbo.Transactions Where (Emp_id = " & Emp_id & ")"

    RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If Not (RsTemp.EOF Or RsTemp.BOF) Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·« Ì„ﬂ‰ Õ–› »Ì«‰«  Â–« «·„ÊŸ›" & CHR(13)
            Msg = Msg + "·«‰… „”Ã· ›Ì »⁄÷ «·Õ—ﬂ«  «· Ã«—Ì… "
        Else
            Msg = "sorry, this record cannot be deleted due to data integration"
        End If
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If
    RsTemp.Close
    StrSQL = " SELECT     EmpId FROM         dbo.Notes WHERE     (EmpId = " & Emp_id & ") "
    RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If Not (RsTemp.EOF Or RsTemp.BOF) Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·« Ì„ﬂ‰ Õ–› »Ì«‰«  Â–« «·„ÊŸ›" & CHR(13)
            Msg = Msg + "·«‰… „”Ã· ›Ì »⁄÷ «·Õ—ﬂ«  «·„«·Ì… "
        Else
            Msg = "sorry, this record cannot be deleted due to data integration"
        End If
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If
    RsTemp.Close
    If DoPremis(Do_Delete, Me.Name, True) = False Then
        Exit Sub
    End If

    If TxtVac_ID7.Text <> "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MSGType = MsgBox("Â·  —€» ›Ì Õ–› Â–« «·”Ã·", vbQuestion + vbMsgBoxRtlReading + vbYesNo + vbMsgBoxRight, App.title)
        Else
            MSGType = MsgBox("Do you want to delete this record", vbQuestion + vbMsgBoxRtlReading + vbYesNo + vbMsgBoxRight, App.title)
        End If
        If MSGType = vbYes Then
            RsSavRec7.Find "id=" & val(TxtVac_ID7.Text), , adSearchForward, 1
            CuurentLogdata7 ("D")
            RsSavRec7.delete
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox " „  ⁄„·Ì… «·Õ–› »‰Ã«Õ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
            Else
                MsgBox "Record deleted successfully", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            End If
            FillGrid7WithData
            btnNext7_Click
        End If
    End If
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217873, -2147467259
            If SystemOptions.UserInterface = ArabicInterface Then
                StrMSG = "⁄›Ê« ·« ÌÃÊ“ Õ–› «·”Ã· ·«— »«ÿÂ »»Ì«‰«  √Œ—Ì"
            Else
                StrMSG = "sorry, this record cannot be deleted due to data integration"
            End If
            RsSavRec7.CancelUpdate
            MsgBox StrMSG, vbOKOnly + vbMsgBoxRight, App.title
            Cn.Errors.Clear
    End Select
End Sub
Private Sub btnFirst7_Click()

    On Error GoTo ErrTrap

    Dim Msg As String

    If Me.TxtModFlg7.Text = "N" Then
        FindRec7 val(TxtVac_ID7.Text)
        Me.TxtModFlg7.Text = "R"
    End If

    TxtModFlg7 = "R"

    If RsSavRec7.RecordCount = 0 Then
        clear_all Me
        Exit Sub
    End If

BegnieWork:
    RsSavRec7.MoveFirst
    FiLLTXT7
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
                Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
                Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
                Msg = "Sorry , this Recored was deleted by other user on the network" & CHR(13)
                Msg = Msg & "Date will be updated now" & CHR(13)
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec7.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub btnLast7_Click()

    On Error GoTo ErrTrap

    Dim Msg As String
    
    If Me.TxtModFlg7.Text = "N" Then
        FindRec7 val(TxtVac_ID7.Text)
        Me.TxtModFlg7.Text = "R"
    End If
    TxtModFlg7 = "R"
    If RsSavRec7.RecordCount = 0 Then
        clear_all Me
        Exit Sub
    End If
BegnieWork:
    RsSavRec7.MoveLast
    FiLLTXT7
    Exit Sub

ErrTrap:
    Select Case Err.Number
        Case -2147217885
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
                Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
                Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
                Msg = "Sorry , this Recored was deleted by other user on the network" & CHR(13)
                Msg = Msg & "Date will be updated now" & CHR(13)
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec7.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub btnModify7_Click()

    Dim Msg As String

    If DoPremis(Do_Edit, Me.Name, True) = False Then
        Exit Sub
    End If

    On Error GoTo ErrTrap

    If TxtVac_ID7.Text <> "" Then
        TxtModFlg7 = "E"
        Frm27.Enabled = True
        Me.TXTDiscounts7.SetFocus
        CuurentLogdata7
    End If
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147467259
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "⁄›Ê«" & CHR(13)
                Msg = Msg & " ·«Ì„ﬂ‰  ⁄œÌ· Â–« «·”Ã· ›Ï «·Êﬁ  «·Õ«·Ï" & CHR(13)
                Msg = Msg & "ÕÌÀ «‰Â ﬁÌœ «· ⁄œÌ· „‰ ﬁ»· „” Œœ„ «Œ— ⁄·Ï «·‘»ﬂ…"
            Else
                Msg = "Sorry , this Recored was deleted by other user on the network" & CHR(13)
                Msg = Msg & "Date will be updated now" & CHR(13)
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            If RsSavRec7.EditMode <> adEditNone Then
                RsSavRec7.CancelUpdate
            End If
    End Select
End Sub
Private Sub btnNew7_Click()

    Dim My_SQL As String
    Dim rs As ADODB.Recordset

    If DoPremis(Do_New, Me.Name, True) = False Then
        Exit Sub
    End If

    On Error GoTo ErrTrap
    
    Set rs = New ADODB.Recordset
    Frm27.Enabled = True
    Me.TxtVac_ID7.Text = ""
    Me.TXTDiscounts7.Text = ""
    Me.DcBranches7.BoundText = ""
    Me.DCEmP7.BoundText = ""
    Me.DCJob7.BoundText = ""
    Me.DCSalesRepGroups7.BoundText = ""
    TxtModFlg7.Text = "N"
    My_SQL = "TBLSalesRepData"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable
    If rs.RecordCount > 0 Then
        TxtSerial.Text = rs.RecordCount + 1
    Else
        TxtSerial.Text = 1
    End If
    rs.Close
    CmbType.ListIndex = 0
    TXTDiscounts7.SetFocus
    TXTDiscounts7.Text = 0
ErrTrap:
End Sub
Private Sub btnNext7_Click()

    On Error GoTo ErrTrap
    
    Dim Msg As String

    If Me.TxtModFlg7.Text = "N" Then
        FindRec7 val(TxtVac_ID7.Text)
        Me.TxtModFlg7.Text = "R"
    End If
    TxtModFlg7 = "R"
    If RsSavRec7.RecordCount = 0 Then
        clear_all Me
        Exit Sub
    End If
BegnieWork:
    If RsSavRec7.EOF Then
        RsSavRec7.MoveLast
    Else
        RsSavRec7.MoveNext
        If RsSavRec7.EOF Then
            RsSavRec7.MoveLast
        End If
    End If
    FiLLTXT7
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
                Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
                Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
                Msg = "Sorry , this Recored was deleted by other user on the network" & CHR(13)
                Msg = Msg & "Date will be updated now" & CHR(13)
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec7.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub btnPrevious7_Click()

    On Error GoTo ErrTrap
    
    Dim Msg As String

    If Me.TxtModFlg7.Text = "N" Then
        FindRec7 val(TxtVac_ID7.Text)
        Me.TxtModFlg7.Text = "R"
    End If
    TxtModFlg7 = "R"
    If RsSavRec7.RecordCount = 0 Then
        clear_all Me
        Exit Sub
    End If
BegnieWork:
    RsSavRec7.MovePrevious
    If RsSavRec7.BOF Then
        RsSavRec7.MoveFirst
    End If
    FiLLTXT7
    Exit Sub
ErrTrap:
    Select Case Err.Number
        Case -2147217885
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
                Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
                Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            Else
                Msg = "Sorry , this Recored was deleted by other user on the network" & CHR(13)
                Msg = Msg & "Date will be updated now" & CHR(13)
            End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec7.Requery
            Resume BegnieWork
    End Select
End Sub
Private Sub BtnPrint7_Click()
    chPrinet7 = 0
    print_report7
End Sub
Private Sub btnQuery7_Click()
    Load FrmSearchSales
    FrmSearchSales.show
End Sub
Private Sub btnSave7_Click()

    On Error GoTo ErrTrap
    
    Dim Msg As String
    Dim StrVacCode As String
    Dim StrVacName As String
    Dim CtrlTxt As Control

    If TXTCode7.Text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "«·—Ã«¡ «œŒ«· «·ﬂÊœ", vbOKOnly + vbMsgBoxRight, App.title
        Else
            MsgBox "Please Enter The Code", vbOKOnly + vbMsgBoxRight, App.title
        End If
        Exit Sub
    End If
    
    If TXTDiscounts7.Text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "«·—Ã«¡ «œŒ«· ‰”»… «·Œ’„", vbOKOnly + vbMsgBoxRight, App.title
        Else
            MsgBox "Please Enter The discount percentage", vbOKOnly + vbMsgBoxRight, App.title
        End If
        Exit Sub
    End If
    
    If val(DCEmP7.BoundText) = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "«·—Ã«¡ «œŒ«· «·„ÊŸ›", vbOKOnly + vbMsgBoxRight, App.title
        Else
            MsgBox "Please chose the employee", vbOKOnly + vbMsgBoxRight, App.title
        End If
        Exit Sub
    End If
    
  ' If val(DcBranches7.BoundText) = 0 Then
  '      If SystemOptions.UserInterface = ArabicInterface Then
  '          MsgBox "«·—Ã«¡ «œŒ«· «·›—⁄", vbOKOnly + vbMsgBoxRight, App.title
  '      Else
  '          MsgBox "Please chose the branch", vbOKOnly + vbMsgBoxRight, App.title
  '      End If
  '      Exit Sub
  '  End If
    
    If val(DCSalesRepGroups7.BoundText) = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "«·—Ã«¡ «œŒ«· «·„Ã„Ê⁄…", vbOKOnly + vbMsgBoxRight, App.title
        Else
            MsgBox "Please chose the group", vbOKOnly + vbMsgBoxRight, App.title
        End If
        Exit Sub
    End If

    Select Case Me.TxtModFlg7.Text
        Case "N"
            AddNewRec7
            btnLast7_Click
        Case "E"
            FiLLRec7
    End Select
    Exit Sub
ErrTrap:
    If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "Â‰«ﬂ Œÿ√ „« ›Ì ≈œŒ«· «·»Ì«‰« ", vbOKOnly + vbMsgBoxRight, App.title
    Else
        MsgBox "Something went wrong while inserting data", vbOKOnly + vbMsgBoxRight, App.title
    End If
End Sub
Private Sub BtnUndo7_Click()
    FindRec7 val(TxtVac_ID7.Text)
    Me.TxtModFlg7.Text = "R"
End Sub
Private Sub BtnUpdate7_Click()

    On Error GoTo ErrTrap
    
    Dim Msg As String
    Dim FristCount As Long
    Dim LastCount As Long
    
    FristCount = RsSavRec7.RecordCount
    RsSavRec7.Requery
    LastCount = RsSavRec7.RecordCount
    BtnUndo7_Click
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
            Msg = "No. of records before the update" & vbCrLf & FristCount & vbCrLf & "⁄œœ «·”Ã·«  »⁄œ «· ÕœÌÀ" & vbCrLf & LastCount
        
            If LastCount > FristCount Then
                Msg = Msg + vbCrLf & "No. of new records" & vbCrLf & LastCount - FristCount
            Else
                Msg = Msg + vbCrLf & "No. of deleted records" & vbCrLf & FristCount - LastCount
            End If
        End If

    End If
    MsgBox Msg, vbOKOnly + vbMsgBoxRight + vbInformation, App.title
ErrTrap:
End Sub
Private Sub Command1_Click()
    Unload FrmEmployee
    If checkApility("FrmEmployee") = False Then
        Exit Sub
    End If
            OpenScreen EmployeesScreen
FrmEmployee.WorkShop_Job = 0
End Sub
Private Sub DCEmP7_Change()
    If val(Me.DCEmP7.BoundText) = 0 Then Exit Sub
    Me.TXTCode7.Text = get_EMPLOYEE_Data(val(Me.DCEmP7.BoundText), "Emp_Code")
End Sub
Private Sub DCEmP7_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        FrmEmployeeSearch.lbltype = 31
        FrmEmployeeSearch.show
    End If
End Sub
Function print_report7(Optional NoteSerial As String, Optional X As Integer)
    
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
    
    If chPrinet7 = 3 Then
        MySQL = " SELECT dbo.TBLSalesRepData.id, dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Fullcode, "
        MySQL = MySQL & " dbo.TblEmployee.Emp_Namee, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode AS CusCode, dbo.TblCustemers.Cus_mobile,"
        MySQL = MySQL & " dbo.TblCustemers.Cus_Phone"
        MySQL = MySQL & " FROM dbo.TBLSalesRepData LEFT OUTER JOIN"
        MySQL = MySQL & " dbo.TblCustemers ON dbo.TBLSalesRepData.EmpID = dbo.TblCustemers.EmpId LEFT OUTER JOIN"
        MySQL = MySQL & " dbo.TblEmployee ON dbo.TBLSalesRepData.EmpID = dbo.TblEmployee.Emp_ID"
        MySQL = MySQL & " Where (dbo.TblEmployee.Emp_ID =" & Me.DCEmP7.BoundText & ")"
        
        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "repSalseCustomer.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "repSalseCustomer.rpt"
        End If
    Else
        If chPrinet7 = 0 Then
            MySQL = "SELECT dbo.TBLSalesRepData.id, dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Fullcode, "
            MySQL = MySQL & " dbo.TblEmployee.Emp_Namee, dbo.TblBranchesData.branch_id, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee,"
            MySQL = MySQL & " dbo.TBLSalesRepGroups.id AS Expr1, dbo.TBLSalesRepGroups.name, dbo.TBLSalesRepGroups.namee, dbo.TblEmpJobsTypes.JobTypeID,"
            MySQL = MySQL & " dbo.TblEmpJobsTypes.JobTypeName , dbo.TblEmpJobsTypes.JobTypeNamee, dbo.TBLSalesRepData.discountvalue"
            MySQL = MySQL & " FROM dbo.TBLSalesRepData INNER JOIN"
            MySQL = MySQL & " dbo.TblEmployee ON dbo.TBLSalesRepData.EmpID = dbo.TblEmployee.Emp_ID INNER JOIN"
            MySQL = MySQL & " dbo.TblBranchesData ON dbo.TBLSalesRepData.BranchId = dbo.TblBranchesData.branch_id INNER JOIN"
            MySQL = MySQL & " dbo.TBLSalesRepGroups ON dbo.TBLSalesRepData.GroupID = dbo.TBLSalesRepGroups.id LEFT OUTER JOIN"
            MySQL = MySQL & " dbo.TblEmpJobsTypes ON dbo.TBLSalesRepData.JobID = dbo.TblEmpJobsTypes.JobTypeID"
            MySQL = MySQL & " Where (dbo.TBLSalesRepData.id =  " & val(TxtVac_ID7.Text) & ")"
            
            If SystemOptions.UserInterface = ArabicInterface Then
                StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "repSalseCustomerN.rpt"
            Else
                StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "repSalseCustomerN.rpt"
            End If
        Else
            If chPrinet7 = 1 Then
                MySQL = " SELECT dbo.TBLSalesRepData.id, dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Fullcode,"
                MySQL = MySQL & " dbo.TblEmployee.Emp_Namee, dbo.TblBranchesData.branch_id, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee,"
                MySQL = MySQL & " dbo.TBLSalesRepGroups.id AS Expr1, dbo.TBLSalesRepGroups.name, dbo.TBLSalesRepGroups.namee, dbo.TblEmpJobsTypes.JobTypeID,"
                MySQL = MySQL & " dbo.TblEmpJobsTypes.JobTypeName , dbo.TblEmpJobsTypes.JobTypeNamee"
                MySQL = MySQL & " FROM dbo.TBLSalesRepData INNER JOIN"
                MySQL = MySQL & " dbo.TblEmployee ON dbo.TBLSalesRepData.EmpID = dbo.TblEmployee.Emp_ID INNER JOIN"
                MySQL = MySQL & " dbo.TblBranchesData ON dbo.TBLSalesRepData.BranchId = dbo.TblBranchesData.branch_id INNER JOIN"
                MySQL = MySQL & " dbo.TBLSalesRepGroups ON dbo.TBLSalesRepData.GroupID = dbo.TBLSalesRepGroups.id LEFT OUTER JOIN"
                MySQL = MySQL & " dbo.TblEmpJobsTypes ON dbo.TBLSalesRepData.JobID = dbo.TblEmpJobsTypes.JobTypeID"
                MySQL = MySQL & " Where (dbo.TBLSalesRepGroups.id =" & Me.DCSalesRepGroups7.BoundText & ")"
                
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "repSalseGroup.rpt"
                Else
                    StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "repSalseGroup.rpt"
                End If
            Else
                MySQL = " SELECT dbo.TBLSalesRepData.id, dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Fullcode,"
                MySQL = MySQL & " dbo.TblEmployee.Emp_Namee, dbo.TblBranchesData.branch_id, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee,"
                MySQL = MySQL & " dbo.TBLSalesRepGroups.id AS Expr1, dbo.TBLSalesRepGroups.name, dbo.TBLSalesRepGroups.namee, dbo.TblEmpJobsTypes.JobTypeID,"
                MySQL = MySQL & " dbo.TblEmpJobsTypes.JobTypeName , dbo.TblEmpJobsTypes.JobTypeNamee"
                MySQL = MySQL & " FROM dbo.TBLSalesRepData INNER JOIN"
                MySQL = MySQL & " dbo.TblEmployee ON dbo.TBLSalesRepData.EmpID = dbo.TblEmployee.Emp_ID INNER JOIN"
                MySQL = MySQL & " dbo.TblBranchesData ON dbo.TBLSalesRepData.BranchId = dbo.TblBranchesData.branch_id INNER JOIN"
                MySQL = MySQL & " dbo.TBLSalesRepGroups ON dbo.TBLSalesRepData.GroupID = dbo.TBLSalesRepGroups.id LEFT OUTER JOIN"
                MySQL = MySQL & " dbo.TblEmpJobsTypes ON dbo.TBLSalesRepData.JobID = dbo.TblEmpJobsTypes.JobTypeID"
                MySQL = MySQL & " Where (dbo.TblBranchesData.branch_id =" & Me.DcBranches7.BoundText & ")"
 
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "repSalseBranch.rpt"
                Else
                    StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "repSalseBranch.rpt"
                End If
            End If
        End If
    End If
 
    If Dir(StrFileName) = "" Then
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        Else
            Msg = "There's no data to show"
        End If
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
        StrReportTitle = "" '& StrAccountName
    Else
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name

    Dim Total As String
    Dim dif As String
    Dim totl As Double

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
Public Sub AddNewRec7()

    On Error GoTo ErrTrap
    
    Dim StrRecId7 As String
    
    StrRecId7 = new_id("TBLSalesRepData", "id", "")
    RsSavRec7.AddNew
    RsSavRec7.Fields("id").value = IIf(StrRecId7 <> "", StrRecId7, Null)
    FiLLRec7
ErrTrap:
End Sub
Public Sub FiLLRec7()

    On Error GoTo ErrTrap

    RsSavRec7.Fields("DiscountValue").value = IIf(IsNumeric(TXTDiscounts7.Text), val(TXTDiscounts7.Text), Null)
    RsSavRec7.Fields("EmpID").value = IIf(val(Me.DCEmP7.BoundText) <> 0, val(Me.DCEmP7.BoundText), Null)
    RsSavRec7.Fields("BranchId").value = IIf(val(Me.DcBranches7.BoundText) <> 0, val(Me.DcBranches7.BoundText), Null)
    RsSavRec7.Fields("GroupID").value = IIf(val(Me.DCSalesRepGroups7.BoundText) <> 0, val(Me.DCSalesRepGroups7.BoundText), Null)

    RsSavRec7.update
    CuurentLogdata7
    If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox " „  ⁄„·Ì… «·Õ›Ÿ »‰Ã«Õ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Else
        MsgBox "Record saved successfully", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    End If
    FillGrid7WithData
    TxtModFlg7 = "R"
    Exit Sub
ErrTrap:
    If RsSavRec7.EditMode <> adEditNone Then
        RsSavRec7.CancelUpdate
    End If
End Sub
Public Sub FiLLTXT7()

    On Error GoTo ErrTrap
    
    Dim i As Integer
    
    Frm27.Enabled = False
    TxtVac_ID7.Text = IIf(IsNull(RsSavRec7.Fields("id").value), "", RsSavRec7.Fields("id").value)
    TXTDiscounts7.Text = IIf(IsNull(RsSavRec7.Fields("DiscountValue").value), 0, RsSavRec7.Fields("DiscountValue").value)
    Me.DCEmP7.BoundText = IIf(IsNull(RsSavRec7.Fields("EmpID").value), "", RsSavRec7.Fields("EmpID").value)
    Me.DcBranches7.BoundText = IIf(IsNull(RsSavRec7.Fields("BranchId").value), "", RsSavRec7.Fields("BranchId").value)
    Me.DCSalesRepGroups7.BoundText = IIf(IsNull(RsSavRec7.Fields("GroupID").value), "", RsSavRec7.Fields("GroupID").value)
    LabCurrRec7.Caption = RsSavRec7.AbsolutePosition
    LabCountRec7.Caption = RsSavRec7.RecordCount

    With Grid7
        For i = 1 To .Rows - 1
            If Trim(TxtVac_ID7.Text) = .TextMatrix(i, .ColIndex("id")) Then
                TxtSerial.Text = .TextMatrix(i, .ColIndex("Ser"))
                .Row = i
                Exit Sub
            End If
        Next
    End With
ErrTrap:
End Sub
Public Sub EditRec7(StrTable As String, RecId7 As String)
    FiLLRec7
End Sub
Private Sub Grid7_EnterCell()

    On Error GoTo ErrTrap
    
    FindRec7 val(Me.Grid7.TextMatrix(Me.Grid7.Row, Me.Grid7.ColIndex("EmpID")))
ErrTrap:
End Sub
Private Sub ISButton17_Click()
    chPrinet7 = 1
    print_report7
End Sub
Private Sub ISButton170_Click()
    chPrinet7 = 3
    print_report7
End Sub
Private Sub ISButton27_Click()
    chPrinet7 = 2
    print_report7
End Sub
Private Sub TxtVac_ID7_Change()

    Dim TxtMod As String
    
    TxtMod = TxtModFlg7.Text
    TxtModFlg7.Text = ""
    TxtModFlg7 = TxtMod
End Sub
Public Function FindRec7(ByVal RecId7 As Long)

    On Error GoTo ErrTrap
    
    RsSavRec7.Find "EmpID=" & RecId7, , adSearchForward, 1
    If Not (RsSavRec7.EOF) Then
        FiLLTXT7
    End If
    Exit Function
ErrTrap:
    If RsSavRec7.EditMode <> adEditNone Then
        RsSavRec7.CancelUpdate
        BtnUndo7_Click
    End If
End Function
Private Sub TxtModFlg7_Change()
    If TxtModFlg7.Text = "N" Then
        Frm27.Enabled = True
        Me.btnNew7.Enabled = False
        btnModify7.Enabled = False
        btnDelete7.Enabled = False
        Me.btnQuery7.Enabled = False
        Grid7.Enabled = False
        BtnUndo7.Enabled = True
        Me.btnSave7.Enabled = True
        BtnUpdate7.Enabled = False
    ElseIf TxtModFlg7.Text = "R" Then
        Frm27.Enabled = False
        Grid7.Enabled = True
        btnModify7.Enabled = False
        btnDelete7.Enabled = False
        If TxtVac_ID7.Text <> "" Then
            btnModify7.Enabled = True
            btnDelete7.Enabled = True
        End If
        BtnUpdate7.Enabled = True
        Me.btnQuery7.Enabled = True
        Me.btnNew7.Enabled = True
        BtnUndo7.Enabled = False
        Me.btnSave7.Enabled = False
        btnNext7.Enabled = True
        btnPrevious7.Enabled = True
        btnFirst7.Enabled = True
        btnLast7.Enabled = True
    ElseIf TxtModFlg7.Text = "E" Then
        Frm27.Enabled = True
        Me.btnNew7.Enabled = False
        btnModify7.Enabled = False
        btnDelete7.Enabled = False
        Me.btnQuery7.Enabled = False
        BtnUpdate7.Enabled = False
        BtnUndo7.Enabled = True
        Me.btnSave7.Enabled = True
        Grid7.Enabled = False
        btnNext7.Enabled = False
        btnPrevious7.Enabled = False
        btnFirst7.Enabled = False
        btnLast7.Enabled = False
    End If

End Sub
Public Sub FillGrid7WithData()

    On Error GoTo ErrTrap

    Dim i As Integer
    Dim rs As ADODB.Recordset
    Dim My_SQL As String
    Set rs = New ADODB.Recordset
    
    My_SQL = "select * From TBLSalesRepData order by id"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText
    With Me.Grid7
        .Rows = 2
        .Clear flexClearScrollable
        If rs.RecordCount > 0 Then
            .Rows = rs.RecordCount + 1
            rs.MoveFirst
            For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("Ser")) = i
                .TextMatrix(i, .ColIndex("EmpID")) = IIf(IsNull(rs.Fields("EmpID").value), "", rs.Fields("EmpID").value)
                .TextMatrix(i, .ColIndex("EmpCode")) = get_EMPLOYEE_Data(val(.TextMatrix(i, .ColIndex("EmpID"))), "Emp_Code")
                .TextMatrix(i, .ColIndex("EmpName")) = IIf(IsNull(rs.Fields("EmpID").value), "", rs.Fields("EmpID").value)
                .TextMatrix(i, .ColIndex("BranchId")) = IIf(IsNull(rs.Fields("BranchId").value), "", rs.Fields("BranchId").value)
                .TextMatrix(i, .ColIndex("GroupID")) = IIf(IsNull(rs.Fields("GroupID").value), "", rs.Fields("GroupID").value)
                .TextMatrix(i, .ColIndex("JobID")) = IIf(IsNull(rs.Fields("JobID").value), "", rs.Fields("JobID").value)
                .TextMatrix(i, .ColIndex("DiscountValue")) = IIf(IsNull(rs.Fields("DiscountValue").value), "", rs.Fields("DiscountValue").value)
                rs.MoveNext
            Next
            rs.Close
        End If
        .RowHeight(-1) = 300
    End With
ErrTrap:
End Sub
