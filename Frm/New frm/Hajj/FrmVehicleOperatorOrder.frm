VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmVehicleOperatorOrder 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "√„—  ‘€Ì· Õ«›·…"
   ClientHeight    =   9840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10290
   Icon            =   "FrmVehicleOperatorOrder.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   9840
   ScaleWidth      =   10290
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   9840
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   10290
      _cx             =   18150
      _cy             =   17357
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
      Begin C1SizerLibCtl.C1Elastic pnlGrid 
         Height          =   2385
         Left            =   120
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   6105
         Width           =   10095
         _cx             =   17806
         _cy             =   4207
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
         Begin VSFlex8Ctl.VSFlexGrid Grid 
            Height          =   2328
            Left            =   0
            TabIndex        =   24
            Top             =   0
            Width           =   10068
            _cx             =   17759
            _cy             =   4106
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
            BackColorAlternate=   16776960
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
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmVehicleOperatorOrder.frx":038A
            ScrollTrack     =   0   'False
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   0   'False
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
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic6 
         Height          =   390
         Left            =   120
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   8565
         Width           =   5610
         _cx             =   9895
         _cy             =   688
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
         Begin VB.Label XPTxtCurrent 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   240
            Left            =   2985
            RightToLeft     =   -1  'True
            TabIndex        =   13
            Top             =   90
            Width           =   795
         End
         Begin VB.Label XPTxtCount 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   240
            Left            =   105
            RightToLeft     =   -1  'True
            TabIndex        =   12
            Top             =   90
            Width           =   645
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «·”Ã· «·Õ«·Ì:"
            Height          =   240
            Index           =   2
            Left            =   3840
            RightToLeft     =   -1  'True
            TabIndex        =   11
            Top             =   90
            Width           =   945
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " ⁄œœ «·”Ã·« :"
            Height          =   240
            Index           =   4
            Left            =   795
            RightToLeft     =   -1  'True
            TabIndex        =   10
            Top             =   90
            Width           =   945
         End
      End
      Begin C1SizerLibCtl.C1Elastic EleHeader 
         Height          =   720
         Left            =   0
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   0
         Width           =   10500
         _cx             =   18521
         _cy             =   1270
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial (Arabic)"
            Size            =   22.5
            Charset         =   178
            Weight          =   700
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
         Caption         =   "     √„—  ‘€Ì· Õ«›·…    "
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
            TabIndex        =   3
            Top             =   180
            Visible         =   0   'False
            Width           =   855
         End
         Begin ImpulseButton.ISButton XPBtnMove 
            Height          =   345
            Index           =   0
            Left            =   1155
            TabIndex        =   4
            Top             =   120
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   609
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
            BackColor       =   -2147483634
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmVehicleOperatorOrder.frx":04EC
            ColorButton     =   -2147483634
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
            TabIndex        =   5
            Top             =   120
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   609
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
            BackColor       =   -2147483634
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmVehicleOperatorOrder.frx":0886
            ColorButton     =   -2147483634
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
            TabIndex        =   6
            Top             =   120
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   609
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
            BackColor       =   -2147483634
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmVehicleOperatorOrder.frx":0C20
            ColorButton     =   -2147483634
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
            TabIndex        =   7
            Top             =   120
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   609
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
            BackColor       =   -2147483634
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmVehicleOperatorOrder.frx":0FBA
            ColorButton     =   -2147483634
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
      Begin C1SizerLibCtl.C1Elastic pnlHeader 
         Height          =   5325
         Left            =   120
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   705
         Width           =   10080
         _cx             =   17780
         _cy             =   9393
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
         Begin C1SizerLibCtl.C1Elastic C1Elastic10 
            Height          =   672
            Left            =   120
            TabIndex        =   79
            TabStop         =   0   'False
            Top             =   4644
            Width           =   9852
            _cx             =   17383
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
            Caption         =   "«·⁄‰Ê«‰"
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
            Begin MSDataListLib.DataCombo MekkaHotelID 
               Height          =   288
               Left            =   6240
               TabIndex        =   80
               Top             =   240
               Width           =   2016
               _ExtentX        =   3545
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               BackColor       =   16777215
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo MadinaHotelID 
               Height          =   288
               Left            =   120
               TabIndex        =   81
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
            Begin MSDataListLib.DataCombo JeddahHotelID 
               Height          =   288
               Left            =   2880
               TabIndex        =   82
               Top             =   240
               Width           =   1896
               _ExtentX        =   3334
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               BackColor       =   16777215
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label Lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "›‰œﬁ «·„œÌ‰…"
               Height          =   288
               Index           =   17
               Left            =   1632
               RightToLeft     =   -1  'True
               TabIndex        =   85
               Top             =   240
               Width           =   1068
            End
            Begin VB.Label Lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "›‰œﬁ Ãœ…"
               Height          =   288
               Index           =   16
               Left            =   4752
               RightToLeft     =   -1  'True
               TabIndex        =   84
               Top             =   240
               Width           =   1068
            End
            Begin VB.Label Lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "›‰œﬁ ›Ï „ﬂ…"
               Height          =   288
               Index           =   15
               Left            =   8472
               RightToLeft     =   -1  'True
               TabIndex        =   83
               Top             =   240
               Width           =   948
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic9 
            Height          =   900
            Left            =   120
            TabIndex        =   66
            TabStop         =   0   'False
            Top             =   3648
            Width           =   9852
            _cx             =   17383
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
            Begin VB.TextBox VehicleNo 
               Alignment       =   1  'Right Justify
               Height          =   312
               Left            =   120
               MaxLength       =   10
               RightToLeft     =   -1  'True
               TabIndex        =   69
               Top             =   144
               Width           =   1740
            End
            Begin VB.TextBox DriverCode 
               Alignment       =   1  'Right Justify
               Height          =   324
               Left            =   6240
               RightToLeft     =   -1  'True
               TabIndex        =   68
               Top             =   504
               Width           =   1932
            End
            Begin VB.TextBox BusNo 
               Alignment       =   1  'Right Justify
               Height          =   324
               Left            =   120
               MaxLength       =   10
               RightToLeft     =   -1  'True
               TabIndex        =   67
               Top             =   504
               Width           =   1740
            End
            Begin MSDataListLib.DataCombo VehicleType 
               Height          =   285
               Left            =   3000
               TabIndex        =   70
               Top             =   150
               Width           =   1890
               _ExtentX        =   3334
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               BackColor       =   16777215
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo ProgrammID 
               Height          =   315
               Left            =   6237
               TabIndex        =   71
               Top             =   150
               Width           =   1935
               _ExtentX        =   3413
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               BackColor       =   16777215
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DriverID 
               Height          =   285
               Left            =   3000
               TabIndex        =   72
               Top             =   510
               Width           =   1890
               _ExtentX        =   3334
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               BackColor       =   16777215
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label Lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "‰Ê⁄ «·Õ«›·« "
               Height          =   312
               Index           =   18
               Left            =   4752
               RightToLeft     =   -1  'True
               TabIndex        =   78
               Top             =   144
               Width           =   1188
            End
            Begin VB.Label Lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·»—‰«„Ã"
               Height          =   312
               Index           =   11
               Left            =   8592
               RightToLeft     =   -1  'True
               TabIndex        =   77
               Top             =   144
               Width           =   1068
            End
            Begin VB.Label Lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄œœ «·Õ«›·« "
               Height          =   315
               Index           =   13
               Left            =   1755
               RightToLeft     =   -1  'True
               TabIndex        =   76
               Top             =   150
               Width           =   1065
            End
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·”«∆ﬁ"
               Height          =   288
               Left            =   4800
               RightToLeft     =   -1  'True
               TabIndex        =   75
               Top             =   504
               Width           =   1092
            End
            Begin VB.Label Label6 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬂÊœ «·”«∆ﬁ"
               Height          =   288
               Left            =   8880
               RightToLeft     =   -1  'True
               TabIndex        =   74
               Top             =   504
               Width           =   732
            End
            Begin VB.Label Lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ «·»«’"
               Height          =   360
               Index           =   19
               Left            =   1875
               RightToLeft     =   -1  'True
               TabIndex        =   73
               Top             =   510
               Width           =   945
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic8 
            Height          =   804
            Left            =   120
            TabIndex        =   53
            TabStop         =   0   'False
            Top             =   2748
            Width           =   9852
            _cx             =   17383
            _cy             =   1429
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
            Begin VB.TextBox FlightNo 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   168
               MaxLength       =   10
               RightToLeft     =   -1  'True
               TabIndex        =   55
               Top             =   72
               Width           =   1728
            End
            Begin VB.TextBox Lounge 
               Alignment       =   1  'Right Justify
               Height          =   300
               Left            =   168
               MaxLength       =   10
               RightToLeft     =   -1  'True
               TabIndex        =   54
               Top             =   432
               Width           =   1728
            End
            Begin MSDataListLib.DataCombo AirPortID 
               Height          =   288
               Left            =   6216
               TabIndex        =   56
               Top             =   72
               Width           =   1956
               _ExtentX        =   3466
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               BackColor       =   16777215
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo AirLineID 
               Height          =   288
               Left            =   2856
               TabIndex        =   57
               Top             =   72
               Width           =   1968
               _ExtentX        =   3493
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               BackColor       =   16777215
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSComCtl2.DTPicker ArriveDate 
               Height          =   300
               Left            =   6216
               TabIndex        =   58
               TabStop         =   0   'False
               Top             =   432
               Width           =   1956
               _ExtentX        =   3440
               _ExtentY        =   529
               _Version        =   393216
               CalendarBackColor=   12648447
               CalendarTitleBackColor=   10383715
               CustomFormat    =   "yyyy/M/d"
               Format          =   115212291
               CurrentDate     =   37140
            End
            Begin MSComCtl2.DTPicker ArriveTime 
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "h:mm:ss AMPM"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   4
               EndProperty
               Height          =   300
               Left            =   2856
               TabIndex        =   59
               TabStop         =   0   'False
               Top             =   432
               Width           =   1968
               _ExtentX        =   3466
               _ExtentY        =   529
               _Version        =   393216
               CalendarBackColor=   12648447
               CalendarTitleBackColor=   10383715
               CustomFormat    =   "yyyy/M/d"
               Format          =   115212290
               CurrentDate     =   37140
            End
            Begin VB.Label Lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„ÿ«—"
               Height          =   300
               Index           =   5
               Left            =   8904
               RightToLeft     =   -1  'True
               TabIndex        =   65
               Top             =   72
               Width           =   756
            End
            Begin VB.Label Lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ŒÿÊÿ «·ÃÊÌ…"
               Height          =   300
               Index           =   7
               Left            =   4824
               RightToLeft     =   -1  'True
               TabIndex        =   64
               Top             =   72
               Width           =   1008
            End
            Begin VB.Label Lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ «·—Õ·…"
               Height          =   300
               Index           =   9
               Left            =   1824
               RightToLeft     =   -1  'True
               TabIndex        =   63
               Top             =   72
               Width           =   768
            End
            Begin VB.Label Lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "“„‰ «·Ê’Ê·"
               Height          =   180
               Index           =   12
               Left            =   5088
               RightToLeft     =   -1  'True
               TabIndex        =   62
               Top             =   432
               Width           =   768
            End
            Begin VB.Label Lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " «—ÌŒ «·Ê’Ê· "
               Height          =   300
               Index           =   10
               Left            =   8700
               RightToLeft     =   -1  'True
               TabIndex        =   61
               Top             =   432
               Width           =   996
            End
            Begin VB.Label Lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·’«·…"
               Height          =   180
               Index           =   14
               Left            =   1800
               RightToLeft     =   -1  'True
               TabIndex        =   60
               Top             =   432
               Width           =   768
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic5 
            Height          =   624
            Left            =   7200
            TabIndex        =   44
            TabStop         =   0   'False
            Top             =   1992
            Width           =   2772
            _cx             =   4895
            _cy             =   1085
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
            Caption         =   "«·„‘—›"
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
            Begin VB.OptionButton emp 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„ÊŸ›"
               Height          =   336
               Left            =   1080
               RightToLeft     =   -1  'True
               TabIndex        =   46
               Top             =   120
               Value           =   -1  'True
               Width           =   972
            End
            Begin VB.OptionButton other 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "√Œ—Ï"
               Height          =   336
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   45
               Top             =   120
               Width           =   852
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic4 
            Height          =   624
            Left            =   120
            TabIndex        =   43
            TabStop         =   0   'False
            Top             =   1992
            Width           =   7092
            _cx             =   12515
            _cy             =   1085
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
            Begin VB.TextBox EmpCode 
               Alignment       =   1  'Right Justify
               Height          =   324
               Left            =   5280
               RightToLeft     =   -1  'True
               TabIndex        =   49
               Top             =   120
               Width           =   1212
            End
            Begin VB.TextBox EmpName 
               Alignment       =   1  'Right Justify
               Height          =   324
               Left            =   2880
               RightToLeft     =   -1  'True
               TabIndex        =   48
               Top             =   120
               Width           =   1932
            End
            Begin VB.TextBox EmpMbile 
               Alignment       =   1  'Right Justify
               Height          =   324
               Left            =   120
               MaxLength       =   10
               RightToLeft     =   -1  'True
               TabIndex        =   47
               Top             =   120
               Width           =   1740
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬂÊœ"
               Height          =   288
               Left            =   6120
               RightToLeft     =   -1  'True
               TabIndex        =   52
               Top             =   120
               Width           =   732
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ "
               Height          =   288
               Left            =   4560
               RightToLeft     =   -1  'True
               TabIndex        =   51
               Top             =   120
               Width           =   612
            End
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "Â« ›"
               Height          =   288
               Left            =   1920
               RightToLeft     =   -1  'True
               TabIndex        =   50
               Top             =   120
               Width           =   612
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic2 
            Height          =   780
            Left            =   120
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   120
            Width           =   9852
            _cx             =   17383
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
            Begin VB.TextBox DependID 
               Alignment       =   1  'Right Justify
               Height          =   288
               Left            =   6120
               RightToLeft     =   -1  'True
               TabIndex        =   27
               Top             =   456
               Width           =   2016
            End
            Begin VB.TextBox ID 
               Alignment       =   1  'Right Justify
               Height          =   276
               Left            =   6120
               Locked          =   -1  'True
               MaxLength       =   10
               RightToLeft     =   -1  'True
               TabIndex        =   26
               Top             =   120
               Width           =   2016
            End
            Begin MSComCtl2.DTPicker SDate 
               Height          =   288
               Left            =   2916
               TabIndex        =   28
               TabStop         =   0   'False
               Top             =   120
               Width           =   1884
               _ExtentX        =   3334
               _ExtentY        =   503
               _Version        =   393216
               CalendarBackColor=   12648447
               CalendarTitleBackColor=   10383715
               CustomFormat    =   "yyyy/M/d"
               Format          =   115212291
               CurrentDate     =   37140
            End
            Begin MSDataListLib.DataCombo BranchID 
               Height          =   288
               Left            =   120
               TabIndex        =   29
               Top             =   120
               Width           =   1764
               _ExtentX        =   3122
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               BackColor       =   16777215
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label Label7 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "»‰«¡« ⁄·Ï  √ﬂÌœ ÕÃ“"
               Height          =   252
               Left            =   8280
               RightToLeft     =   -1  'True
               TabIndex        =   33
               Top             =   456
               Width           =   1332
            End
            Begin VB.Label Lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„”·”·"
               Height          =   312
               Index           =   8
               Left            =   8400
               RightToLeft     =   -1  'True
               TabIndex        =   32
               Top             =   120
               Width           =   1188
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   " «—ÌŒ «·ÌÊ„"
               ForeColor       =   &H00000000&
               Height          =   252
               Left            =   5064
               TabIndex        =   31
               Top             =   120
               Width           =   744
            End
            Begin VB.Label Lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·›—⁄"
               Height          =   288
               Index           =   24
               Left            =   2076
               RightToLeft     =   -1  'True
               TabIndex        =   30
               Top             =   120
               Width           =   528
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic3 
            Height          =   876
            Left            =   120
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   996
            Width           =   9852
            _cx             =   17383
            _cy             =   1561
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
            Begin VB.TextBox TxtCompnyOut 
               Alignment       =   1  'Right Justify
               Height          =   288
               Left            =   2880
               RightToLeft     =   -1  'True
               TabIndex        =   87
               Top             =   120
               Width           =   1896
            End
            Begin VB.TextBox TxtCompnyIn 
               Alignment       =   1  'Right Justify
               Height          =   288
               Left            =   6120
               RightToLeft     =   -1  'True
               TabIndex        =   86
               Top             =   120
               Width           =   2016
            End
            Begin MSDataListLib.DataCombo InClientID 
               Height          =   288
               Left            =   6120
               TabIndex        =   35
               Top             =   120
               Visible         =   0   'False
               Width           =   2016
               _ExtentX        =   3545
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               BackColor       =   16777215
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo OutClientID 
               Height          =   285
               Left            =   120
               TabIndex        =   36
               Top             =   120
               Width           =   1890
               _ExtentX        =   3334
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               BackColor       =   16777215
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo GroupID 
               Height          =   288
               Left            =   6120
               TabIndex        =   37
               Top             =   444
               Width           =   2016
               _ExtentX        =   3545
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               BackColor       =   16777215
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo CompanyID 
               Height          =   288
               Left            =   2880
               TabIndex        =   38
               Top             =   444
               Width           =   1896
               _ExtentX        =   3334
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               BackColor       =   16777215
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label Lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·⁄„Ì·"
               Height          =   315
               Index           =   20
               Left            =   2040
               RightToLeft     =   -1  'True
               TabIndex        =   88
               Top             =   120
               Width           =   705
            End
            Begin VB.Label Lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·‘—ﬂ… „‰ «·Œ«—Ã"
               Height          =   312
               Index           =   0
               Left            =   4752
               RightToLeft     =   -1  'True
               TabIndex        =   42
               Top             =   120
               Width           =   1188
            End
            Begin VB.Label Lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·‘—ﬂ… «·”⁄ÊœÌ…"
               Height          =   312
               Index           =   6
               Left            =   8520
               RightToLeft     =   -1  'True
               TabIndex        =   41
               Top             =   120
               Width           =   1188
            End
            Begin VB.Label Lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·„Ã„Ê⁄…"
               Height          =   288
               Index           =   1
               Left            =   8592
               RightToLeft     =   -1  'True
               TabIndex        =   40
               Top             =   444
               Width           =   1068
            End
            Begin VB.Label Lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·‘—ﬂ…"
               Height          =   288
               Index           =   3
               Left            =   4872
               RightToLeft     =   -1  'True
               TabIndex        =   39
               Top             =   444
               Width           =   1068
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic7 
         Height          =   720
         Left            =   120
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   9045
         Width           =   10035
         _cx             =   17701
         _cy             =   1270
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
         Begin ImpulseButton.ISButton Cmd 
            Height          =   480
            Index           =   0
            Left            =   8832
            TabIndex        =   14
            Top             =   120
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   847
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
            ButtonImage     =   "FrmVehicleOperatorOrder.frx":1354
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
            Height          =   480
            Index           =   1
            Left            =   7752
            TabIndex        =   15
            Top             =   120
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   847
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
            ButtonImage     =   "FrmVehicleOperatorOrder.frx":7BB6
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
            Height          =   480
            Index           =   2
            Left            =   6528
            TabIndex        =   16
            Top             =   120
            Width           =   1212
            _ExtentX        =   2143
            _ExtentY        =   847
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
            ButtonImage     =   "FrmVehicleOperatorOrder.frx":E418
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
            Height          =   480
            Index           =   3
            Left            =   5508
            TabIndex        =   17
            Top             =   120
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   847
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
            ButtonImage     =   "FrmVehicleOperatorOrder.frx":14C7A
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
            Height          =   480
            Index           =   4
            Left            =   4176
            TabIndex        =   18
            Top             =   120
            Width           =   1332
            _ExtentX        =   2355
            _ExtentY        =   847
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
            ButtonImage     =   "FrmVehicleOperatorOrder.frx":1B4DC
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
            Height          =   480
            Index           =   6
            Left            =   1140
            TabIndex        =   19
            Top             =   120
            Width           =   876
            _ExtentX        =   1535
            _ExtentY        =   847
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
            ButtonImage     =   "FrmVehicleOperatorOrder.frx":21D3E
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
            Height          =   480
            Left            =   108
            TabIndex        =   20
            Top             =   120
            Width           =   972
            _ExtentX        =   1720
            _ExtentY        =   847
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
            ButtonImage     =   "FrmVehicleOperatorOrder.frx":4B960
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
            Height          =   480
            Index           =   7
            Left            =   3300
            TabIndex        =   21
            Top             =   120
            Width           =   852
            _ExtentX        =   1508
            _ExtentY        =   847
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
            ButtonImage     =   "FrmVehicleOperatorOrder.frx":521C2
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
            Height          =   480
            Index           =   9
            Left            =   2040
            TabIndex        =   22
            Top             =   120
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   847
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
            ButtonImage     =   "FrmVehicleOperatorOrder.frx":58A24
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
   End
End
Attribute VB_Name = "FrmVehicleOperatorOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim Rs_Temp As ADODB.Recordset
Dim TTP As clstooltip

Private Sub Cmd_Click(Index As Integer)
'    On Error GoTo ErrTrap
    Select Case Index
        Case 0

            If DoPremis(Do_New, Me.name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.text = "N"
            clear_all Me
             BranchID.BoundText = Current_branch
            ID.text = CStr(new_id("TblVehicleOperatorOrder", "ID", "", True))
           emp.value = True
           Grid.Rows = Grid.FixedRows
           Grid.Rows = Grid.FixedRows + 10
        Case 1

            If DoPremis(Do_Edit, Me.name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.text = "E"
            Grid.Rows = Grid.Rows + 1
        Case 2

            SaveData

        Case 3
            Undo

        Case 4

            If DoPremis(Do_Delete, Me.name, True) = False Then
                Exit Sub
            End If

            Del_Action

        Case 5

        Case 6
            Unload Me
         Case 7
         print_report2
         Case 9
         
                Unload FrmSearch_Hajj
                FrmSearch_Hajj.SendForm = "VehicleOperatorOrder"
                FrmSearch_Hajj.show
         
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub CmdHelp_Click()
    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
    SystemOptions.SysHelp.HHDisplayTopicID Me.hWnd
End Sub

 

Private Sub DriverCode_Change()
Dim val1, val2
If DriverCode.text = "" Then Exit Sub
Dim str As String
    str = " select * from TblEmployee  where Emp_Code =  '" & DriverCode.text & "'"
    Set Rs_Temp = New ADODB.Recordset
    Rs_Temp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If Rs_Temp.RecordCount > 0 Then
        Rs_Temp.MoveFirst '
        'val1 = IIf(IsNull(rs_temp("NumEkama").value), "", rs_temp("NumEkama").value)
        val2 = IIf(IsNull(Rs_Temp("Emp_ID").value), "", Rs_Temp("Emp_ID").value)
    End If
   ' txtID.text = val1
    DriverID.BoundText = val(val2)
    
End Sub

Private Sub DriverID_Change()

Dim val1 As String, val2 As String
If DriverID.BoundText = "" Then Exit Sub
Dim I  As Integer, str As String
I = DriverID.BoundText
If I > 0 Then
    str = " select * from TblEmployee  where Emp_ID =  " & I
    Set Rs_Temp = New ADODB.Recordset
    Rs_Temp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
   
    If Rs_Temp.RecordCount > 0 Then
        Rs_Temp.MoveFirst '
        val1 = IIf(IsNull(Rs_Temp("Emp_Code").value), "", Rs_Temp("Emp_Code").value)
        'val2 = IIf(IsNull(rs_temp("NumEkama").value), "", rs_temp("NumEkama").value)
    End If
End If
DriverCode.text = val1
'txtID.text = val2
End Sub


Private Sub EmpCode_Change()
    
    If emp.value = True Then
        Dim val1, val2
        If EmpCode.text = "" Then Exit Sub
        Dim str As String, name As String, Mobile As String
        name = ""
        Mobile = ""
        
            str = " select  * from TblEmployee  where  fullcode = '" & EmpCode.text & "'"
            Set Rs_Temp = New ADODB.Recordset
            Rs_Temp.Open str, Cn, adOpenStatic, adLockOptimistic, adCmdText
            If Rs_Temp.RecordCount > 0 Then
                Rs_Temp.MoveFirst '
                name = IIf(IsNull(Rs_Temp("emp_name").value), "", Rs_Temp("emp_name").value)
                Mobile = IIf(IsNull(Rs_Temp("emp_Mobile").value), "", Rs_Temp("emp_Mobile").value)
             Else
                EmpName.text = ""
                EmpMbile.text = ""
            End If
            
            EmpName.text = name
            EmpMbile.text = Mobile
    End If

End Sub



Private Sub FlightNo_KeyPress(KeyAscii As Integer)
'KeyAscii = KeyAscii_Num(KeyAscii, Me.FlightNo.text, 1)
End Sub

Private Sub Form_Activate()
'    txtid.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        If Me.TxtModFlg.text = "R" Then
            Cmd_Click (0)
        Else
            SendKeys "{TAB}"
        End If
    End If

    If Me.TxtModFlg.text = "R" Then
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

Private Sub Fill_Combos()
 Dim Dcombos As ClsDataCombos
  Dim str As String
  
   Set Dcombos = New ClsDataCombos
   
   Dcombos.GetBranches BranchID
   Dcombos.GetCompany InClientID, 0, 1
   Dcombos.GetCompany OutClientID, 2, 2
    
   str = "select ID, Name from tblcompaniesgroup"
   fill_combo GroupID, str
    
   
   str = "select Id , Name from TblTourismCompanies "
   fill_combo CompanyID, str
   
    str = "select Id , name  from tblairlines"
   fill_combo AirLineID, str
   
    str = "select id , name from TblAirport "
   fill_combo AirPortID, str
   
    str = "select id ,name from TblProgrammTypes "
   fill_combo ProgrammID, str
   
    str = "select id , name from tblhotels"
   fill_combo MekkaHotelID, str
    
      str = "select id , name from tblhotels"
   fill_combo JeddahHotelID, str
   
     str = "select id , name from tblhotels"
   fill_combo MadinaHotelID, str
   
   Dcombos.GetTblCarsDataGroup VehicleType
   
   
   str = "  select   e.Emp_ID Emp_ID , e.Emp_Name   Emp_Name  from TblEmployee e, TblEmpJobsTypes  j"
  str = str & "   Where e.JobTypeID = j.JobTypeID"
 str = str & "     and  ( j.JobTypeName like '%”«∆ﬁ%'  or j.JobTypeNamee like '%driver%')"
   
    fill_combo DriverID, str
   
   ' Dcombos.getCountriesGovernments Me.inCity
End Sub


Private Sub Form_Load()
   On Error GoTo ErrTrap


        Fill_Combos
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    LogTextA = "   «·œŒÊ· «·Ì ‘«‘… " & "  „·› «·„œ«—”  "
   LogTextE = " Open Window " & "  Boxes Data "
   AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTextE, Me.name, "O", "", ""



    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
   Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
    Set Cmd(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
    Set Cmd(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
    Set Cmd(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
 '   Set CmdHelp.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Help").Picture
    Resize_Form Me
    AddTip
    Set rs = New ADODB.Recordset
    
    
    
    Dim StrSQL As String
    StrSQL = ""
    
     If SystemOptions.usertype <> UserAdminAll Then
      
StrSQL = "SELECT  *  From TblVehicleOperatorOrder    "
  Else
 StrSQL = "SELECT  *  From TblVehicleOperatorOrder "
    End If
  rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
      
      


        
    Me.TxtModFlg.text = "R"
    XPBtnMove_Click 2

    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If

    Exit Sub

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
                    StrMSG = "You will close this screen before save " & Chr(13)
                    StrMSG = StrMSG & " the new data  " & Chr(13)
                    StrMSG = StrMSG & " do you want save before exit" & Chr(13)
                    StrMSG = StrMSG & "yes" & "-" & "save the new data" & Chr(13)
                    StrMSG = StrMSG & "no" & "-" & "Don't save" & Chr(13)
                    StrMSG = StrMSG & "cancel" & "-" & "Cancel Exit" & Chr(13)
    
                Else
                    StrMSG = "”Ê› Ì „ €·ﬁ «·‘«‘… Ê·„  ‰ Â „‰  ”ÃÌ·" & Chr(13)
                    StrMSG = StrMSG & " «·»Ì«‰«  «·ÃœÌœ… «·Õ«·Ì… " & Chr(13)
                    StrMSG = StrMSG & " Â·  —Ìœ «·Õ›Ÿ ﬁ»· «·Œ—ÊÃ" & Chr(13)
                    StrMSG = StrMSG & "‰⁄„" & "-" & "Ì „ Õ›Ÿ «·»Ì«‰«  «·ÃœÌœ…" & Chr(13)
                    StrMSG = StrMSG & "·«" & "-" & "·‰ Ì „ «·Õ›Ÿ" & Chr(13)
                    StrMSG = StrMSG & "≈·€«¡ «·√„—" & "-" & "≈·€«¡ ⁄„·Ì… «·Œ—ÊÃ" & Chr(13)
        
                End If
        
            Case "E"

                If SystemOptions.UserInterface = EnglishInterface Then
                    StrMSG = "You will close this screen before save  " & Chr(13)
                    StrMSG = StrMSG & " the Modifications  " & Chr(13)
                    StrMSG = StrMSG & " do you want save before exit" & Chr(13)
                    StrMSG = StrMSG & "yes" & "-" & "save the new data" & Chr(13)
                    StrMSG = StrMSG & "no" & "-" & "Don't save" & Chr(13)
                    StrMSG = StrMSG & "cancel" & "-" & "Cancel Exit" & Chr(13)
    
                Else
                    StrMSG = "”Ê› Ì „ €·ﬁ «·‘«‘… Ê·„  ‰ Â „‰  ”ÃÌ·" & Chr(13)
                    StrMSG = StrMSG & " «· ⁄œÌ·«  «·ÃœÌœ… ⁄·Ï «·”Ã· «·Õ«·Ï " & Chr(13)
                    StrMSG = StrMSG & " Â·  —Ìœ «·Õ›Ÿ ﬁ»· «·Œ—ÊÃ" & Chr(13)
                    StrMSG = StrMSG & "‰⁄„" & "-" & "Ì „ Õ›Ÿ «· ⁄œÌ·«   «·ÃœÌœ…" & Chr(13)
                    StrMSG = StrMSG & "·«" & "-" & "·‰ Ì „ «·Õ›Ÿ" & Chr(13)
                    StrMSG = StrMSG & "≈·€«¡ «·√„—" & "-" & "≈·€«¡ ⁄„·Ì… «·Œ—ÊÃ" & Chr(13)
                
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

    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic
    
   
    lbl(7).Caption = " Name En"
    lbl(3).Caption = " Name Ar"
    lbl(8).Caption = "Process No"
    lbl(0).Caption = "Minister No."
    Label3.Caption = "School Manager"
    Label1.Caption = "Managerial Area"
    Label2.Caption = "City"
    lbl(5).Caption = "Student Count"
    lbl(6).Caption = "Custom"
    lbl(1).Caption = "School Type"
    lbl(10).Caption = "Telephone"
    Label4.Caption = "Supervisor Code"
    Label6.Caption = "Supervisor"
    Label5.Caption = "Student Gender"
    Me.Caption = "School Data"
    EleHeader.Caption = Me.Caption
    
    lbl(2).Caption = "Current Record"
    lbl(4).Caption = "NO. Recordes"

    Me.Cmd(0).Caption = "New"
    Me.Cmd(1).Caption = "Edit"
    Me.Cmd(2).Caption = "Save"
    Me.Cmd(3).Caption = "Undo"
    Me.Cmd(4).Caption = "Delete"
    'Me.Cmd(5).Caption = "Search"
    Me.Cmd(6).Caption = "Exit"
    Me.Cmd(7).Caption = "Print"
   CmdAttach.Caption = "Attachment"

lbl(9).Caption = "Last Contract"



End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap
    LogTextA = "     «·Œ—ÊÃ „‰ ‘«‘… " & "  »Ì«‰«  √„—  ‘€Ì· Õ«›·…   "
    LogTextE = " Exit Window " & "  Boxes Data "
    AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTextE, Me.name, "O", "", ""

    If rs.State = adStateOpen Then
        If Not (rs.EOF Or rs.BOF) Then
            If rs.EditMode <> adEditNone Then
                rs.CancelUpdate
            End If
        End If

        rs.Close
    End If

    Set rs = Nothing
    Set TTP = Nothing
    Exit Sub
ErrTrap:
End Sub

 

Private Sub Grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)

Dim StrAccountCode As String
Dim Msg As String
'  Dim rs As New ADODB.Recordset
Dim StrSQL As String
Dim ClsAcc As New ClsAccounts
Dim LngRow As Long
Dim Sql As String
Dim count As Integer
Dim rate As Double
 
    With Grid

     Select Case .ColKey(Col)

             Case "FromCity"
                        StrAccountCode = .ComboData
                        .TextMatrix(Row, .ColIndex("FromcityId")) = StrAccountCode
                        Grid.Rows = Grid.Rows + 1
             Case "ToCity"
                         StrAccountCode = .ComboData
                        .TextMatrix(Row, .ColIndex("tocityId")) = StrAccountCode
     End Select
End With




End Sub


 



Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

With Grid

If TxtModFlg.text = "R" Then
          .ComboList = ""
          Cancel = True
End If


Select Case .ColKey(Col)
    Case "Date"
         .ComboList = ""
          Cancel = True
    Case "Time"
            .ComboList = ""
            Cancel = True
    Case "Remark"
            .ComboList = ""
            Cancel = True
    End Select
 End With

End Sub

Private Sub Grid_Click()
If Grid.Row > 0 And TxtModFlg.text <> "R" Then
Select Case Grid.ColKey(Grid.Col)
Case "Date"
            Unload FrmRegesterDateProject
            FrmRegesterDateProject.SendForm = "VehicleOperatorOrder"
            FrmRegesterDateProject.show vbModal
Case "Time"
            Unload FrmRegesterDateProject
            FrmRegesterDateProject.SendForm = "VehicleOperatorOrder"
            FrmRegesterDateProject.show vbModal
End Select
End If
End Sub

Private Sub Grid_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)


'Dim rs As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim Msg As String
    
     With Grid
     
     Select Case .ColKey(Col)
     
    
 
          
       Case "FromCity"
          Set Rs_Temp = New ADODB.Recordset
          StrSQL = " Select GovernmentID,GovernmentName From TblCountriesGovernments  "
          Rs_Temp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
          StrComboList = Grid.BuildComboList(Rs_Temp, "GovernmentName", "GovernmentID")
           If StrComboList <> "" Then
                 StrComboList = "|" & StrComboList
           End If
          .ComboList = StrComboList
                            
         Case "ToCity"
           Set Rs_Temp = New ADODB.Recordset
          StrSQL = " Select GovernmentID,GovernmentName From TblCountriesGovernments  "
          Rs_Temp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
          StrComboList = Grid.BuildComboList(Rs_Temp, "GovernmentName", "GovernmentID")
           If StrComboList <> "" Then
                 StrComboList = "|" & StrComboList
           End If
          .ComboList = StrComboList
           
     End Select
   End With
   



End Sub

Private Sub GroupID_Change()

Dim str As String
Set Rs_Temp = New ADODB.Recordset
Set CompanyID.RowSource = Rs_Temp
If SystemOptions.UserInterface = ArabicInterface Then
    str = " select  ID , Name   from TblTourismCompanies where GroupID   = " & val(GroupID.BoundText)
Else
    str = " Select ID , NameE   TblTourismCompanies where GroupID  = " & val(GroupID.BoundText)
End If
fill_combo CompanyID, str
CompanyID.Refresh

End Sub

Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap

    Select Case Me.TxtModFlg.text

        Case "R"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«  √„—  ‘€Ì· Õ«›·…"
            Else
                Me.Caption = "School  Data"
            End If

            Me.Cmd(2).Enabled = False
            Me.Cmd(3).Enabled = False
        
            Me.Cmd(0).Enabled = True
            Me.Cmd(1).Enabled = True
            Me.Cmd(4).Enabled = True
            Me.Cmd(9).Enabled = True
        
            Me.XPBtnMove(0).Enabled = True
            Me.XPBtnMove(1).Enabled = True
            Me.XPBtnMove(2).Enabled = True
            Me.XPBtnMove(3).Enabled = True
        
            ID.locked = True
      

            If rs.RecordCount < 1 Then
                Me.XPBtnMove(0).Enabled = False
                Me.XPBtnMove(1).Enabled = False
                Me.XPBtnMove(2).Enabled = False
                Me.XPBtnMove(3).Enabled = False
                Me.Cmd(1).Enabled = False
                Me.Cmd(4).Enabled = False
            End If
            
            pnlHeader.Enabled = False
            
        Case "N"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«  √„—  ‘€Ì· Õ«›·… ( ÃœÌœ )"
            Else
                Me.Caption = "Booking Request Data(New)"
            End If
        
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«   √„—  ‘€Ì· Õ«›·… ( ÃœÌœ )"
            Else
                Me.Caption = "Booking Request Data(New)"
            End If
        
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(9).Enabled = False
            ID.locked = True
            pnlHeader.Enabled = True
        Case "E"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«   √„—  ‘€Ì· Õ«›·… (  ⁄œÌ· )"
            Else
                Me.Caption = "Booking Request Data(Edit)"
            End If
        
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(9).Enabled = False
            
            Me.XPBtnMove(0).Enabled = False
            Me.XPBtnMove(1).Enabled = False
            Me.XPBtnMove(2).Enabled = False
            Me.XPBtnMove(3).Enabled = False
        
            ID.locked = True
           pnlHeader.Enabled = True
    End Select

    Exit Sub
ErrTrap:
End Sub
Function print_report2(Optional NoteSerial As String)
    
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
    



MySQL = " SELECT     dbo.TblVehicleOperatorOrder.ID, dbo.TblVehicleOperatorOrder.ProgrammID, dbo.TblVehicleOperatorOrder.AirLineID, dbo.TblVehicleOperatorOrder.AirPortID, "
MySQL = MySQL & "                       dbo.TblVehicleOperatorOrder.CompanyID, dbo.TblVehicleOperatorOrder.MekkaHotelID, TblHotels_2.Name AS MekkaHotelName,"
MySQL = MySQL & "                       dbo.TblVehicleOperatorOrder.MadinaHotelID, TblHotels_1.Name AS MadinaHotelName, dbo.TblVehicleOperatorOrder.JeddahHotelID,"
MySQL = MySQL & "                       TblHotels_2.Name AS JeddahHotelName, dbo.TblVehicleOperatorOrder.InClientID, TblCustemers_1.CusName AS InClientName,"
MySQL = MySQL & "                       dbo.TblVehicleOperatorOrder.OutClientID, dbo.TblCustemers.CusName AS OutClientName, dbo.TblAirport.Name AS AirPortName,"
MySQL = MySQL & "                       dbo.TblAirlines.Name AS AirLineName, dbo.TblTourismCompanies.Name AS CompanyName, dbo.TblBranchesData.branch_name AS BranchName,"
MySQL = MySQL & "                       dbo.TblCompaniesGroup.Name AS GroupName, dbo.TblProgrammTypes.Name AS ProgrammName, dbo.TblVehicleOperatorOrder.SDate,"
MySQL = MySQL & "                       dbo.TblVehicleOperatorOrder.BranchID, dbo.TblVehicleOperatorOrder.FlightNo, dbo.TblVehicleOperatorOrder.emp, dbo.TblVehicleOperatorOrder.GroupID,"
MySQL = MySQL & "                       dbo.TblVehicleOperatorOrder.other, dbo.TblVehicleOperatorOrder.EmpID, dbo.TblVehicleOperatorOrder.EmpName, dbo.TblVehicleOperatorOrder.EmpCode,"
MySQL = MySQL & "                       dbo.TblVehicleOperatorOrder.EmpMbile, CONVERT(char(10), dbo.TblVehicleOperatorOrder.ArriveTime, 108) AS ArriveTime, dbo.TblVehicleOperatorOrder.ArriveDate,"
MySQL = MySQL & "                       dbo.TblVehicleOperatorOrder.VehicleNo, dbo.TblVehicleOperatorOrder.Model, dbo.TblVehicleOperatorOrder.VehicleType, dbo.TblVehicleOperatorOrder.CreationUserID,"
MySQL = MySQL & "                       dbo.TblVehicleOperatorOrder.CreationDate, dbo.TblVOODetails.FromCity, dbo.TblVOODetails.TOCity, dbo.TblVOODetails.[Date], dbo.TblVOODetails.HID,"
MySQL = MySQL & "                       CONVERT(char(10), dbo.TblVOODetails.[Time], 108) AS Time, dbo.TblVOODetails.CreationUserID AS Expr1, dbo.TblVOODetails.CreationDate AS Expr2,"
MySQL = MySQL & "                       dbo.TblVOODetails.Remarks, dbo.TblCountriesGovernments.GovernmentName AS FromCityName, TblCountriesGovernments_1.GovernmentName AS ToCityName,"
MySQL = MySQL & "                       dbo.TblVehicleOperatorOrder.Lounge, dbo.TblVehicleOperatorOrder.DependID, dbo.TblVehicleOperatorOrder.DriverID, dbo.TblVehicleOperatorOrder.BusNo,"
MySQL = MySQL & "                       dbo.TblEmployee.Emp_Name AS DriverName, dbo.TblEmployee.Fullcode AS DriverCode, dbo.TblVehicleOperatorOrder.CompnyIn,"
MySQL = MySQL & "                       dbo.TblVehicleOperatorOrder.CompnyOut"
MySQL = MySQL & "  FROM         dbo.TblAirlines RIGHT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblVehicleOperatorOrder LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblCustemers ON dbo.TblVehicleOperatorOrder.OutClientID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblCustemers TblCustemers_1 ON dbo.TblVehicleOperatorOrder.InClientID = TblCustemers_1.CusID LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblEmployee ON dbo.TblVehicleOperatorOrder.DriverID = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblProgrammTypes ON dbo.TblVehicleOperatorOrder.ProgrammID = dbo.TblProgrammTypes.ID LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblBranchesData ON dbo.TblVehicleOperatorOrder.BranchID = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblCompaniesGroup ON dbo.TblVehicleOperatorOrder.GroupID = dbo.TblCompaniesGroup.ID LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblHotels TblHotels_3 ON dbo.TblVehicleOperatorOrder.JeddahHotelID = TblHotels_3.ID LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblHotels TblHotels_2 ON dbo.TblVehicleOperatorOrder.MekkaHotelID = TblHotels_2.ID LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblAirport ON dbo.TblVehicleOperatorOrder.AirPortID = dbo.TblAirport.ID ON dbo.TblAirlines.ID = dbo.TblVehicleOperatorOrder.AirLineID LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblCountriesGovernments TblCountriesGovernments_1 RIGHT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblCountriesGovernments RIGHT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblVOODetails ON dbo.TblCountriesGovernments.GovernmentID = dbo.TblVOODetails.FromCity ON"
MySQL = MySQL & "                       TblCountriesGovernments_1.GovernmentID = dbo.TblVOODetails.TOCity ON dbo.TblVehicleOperatorOrder.ID = dbo.TblVOODetails.HID LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblHotels TblHotels_1 ON dbo.TblVehicleOperatorOrder.MadinaHotelID = TblHotels_1.ID LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblTourismCompanies ON dbo.TblVehicleOperatorOrder.CompanyID = dbo.TblTourismCompanies.ID"
MySQL = MySQL & "  where  dbo.TblVehicleOperatorOrder.ID  =  " & val(ID.text)
 If SystemOptions.UserInterface = ArabicInterface Then
          StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rpt_VehicleOperatorOrder.rpt"
     Else
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rpt_VehicleOperatorOrder.rpt"
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
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName
        StrReportTitle = "" '& StrAccountName
    Else
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
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
Public Sub Retrive(Optional Lngid As Long = 0)

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
            rs.find "ID =" & Lngid, , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Exit Sub
            End If
        End If
    End If
   
    ID.text = IIf(IsNull(rs("ID").value), "", (rs("ID").value))
    SDate.value = IIf(IsNull(rs("Sdate").value), Date, rs("Sdate").value)
    BranchID.BoundText = IIf(IsNull(rs("BranchID").value), "", Trim(rs("BranchID").value))
    InClientID.BoundText = IIf(IsNull(rs("InClientID").value), "", Trim(rs("InClientID").value))
    OutClientID.BoundText = IIf(IsNull(rs("OutClientID").value), "", Trim(rs("OutClientID").value))
    AirLineID.BoundText = IIf(IsNull(rs("AirLineID").value), "", Trim(rs("AirLineID").value))
    AirPortID.BoundText = IIf(IsNull(rs("AirPortID").value), "", Trim(rs("AirPortID").value))
    emp.value = IIf(IsNull(rs("emp").value), False, Trim(rs("emp").value))
    other.value = IIf(IsNull(rs("other").value), False, Trim(rs("other").value))
    EmpCode.text = IIf(IsNull(rs("EmpCode").value), "", Trim(rs("EmpCode").value))
    EmpName.text = IIf(IsNull(rs("EmpName").value), "", Trim(rs("EmpName").value))
    EmpMbile.text = IIf(IsNull(rs("EmpMbile").value), "", Trim(rs("EmpMbile").value))
    FlightNo.text = IIf(IsNull(rs("FlightNo").value), "", Trim(rs("FlightNo").value))
    ArriveDate.value = IIf(IsNull(rs("ArriveDate").value), Date, Trim(rs("ArriveDate").value))
    ArriveTime.value = IIf(IsNull(rs("ArriveTime").value), Date, Trim(rs("ArriveTime").value))
    ProgrammID.BoundText = IIf(IsNull(rs("ProgrammID").value), "", Trim(rs("ProgrammID").value))
    VehicleNo.text = IIf(IsNull(rs("VehicleNo").value), 0, Trim(rs("VehicleNo").value))
    TxtCompnyOut.text = IIf(IsNull(rs("CompnyOut").value), "", Trim(rs("CompnyOut").value))
    TxtCompnyIn.text = IIf(IsNull(rs("CompnyIn").value), "", Trim(rs("CompnyIn").value))
    
 '   Model.text = IIf(IsNull(rs("Model").value), "", Trim(rs("Model").value))
    MekkaHotelID.BoundText = IIf(IsNull(rs("MekkaHotelID").value), "", Trim(rs("MekkaHotelID").value))
    MadinaHotelID.BoundText = IIf(IsNull(rs("MadinaHotelID").value), "", Trim(rs("MadinaHotelID").value))
    JeddahHotelID.BoundText = IIf(IsNull(rs("JeddahHotelID").value), "", Trim(rs("JeddahHotelID").value))
    VehicleType.BoundText = IIf(IsNull(rs("VehicleType").value), "", Trim(rs("VehicleType").value))
    GroupID.BoundText = IIf(IsNull(rs("GroupID").value), "", Trim(rs("GroupID").value))
    CompanyID.BoundText = IIf(IsNull(rs("CompanyID").value), "", Trim(rs("CompanyID").value))
    
     DriverID.BoundText = IIf(IsNull(rs("DriverID").value), "", Trim(rs("DriverID").value))
     DependID.text = IIf(IsNull(rs("DependID").value), "", Trim(rs("DependID").value))
     BusNo.text = IIf(IsNull(rs("BusNo").value), "", Trim(rs("BusNo").value))
     Lounge.text = IIf(IsNull(rs("Lounge").value), "", Trim(rs("Lounge").value))
     
    Set Rs_Temp = New ADODB.Recordset
    Dim StrSQL As String
    StrSQL = " SELECT   TblVOODetails.remarks , dbo.TblVOODetails.FromCity , dbo.TblCountriesGovernments.GovernmentName AS FromCityName, dbo.TblVOODetails.TOCity , "
    StrSQL = StrSQL & "   TblCountriesGovernments_1.GovernmentName AS ToCityName, dbo.TblVOODetails.*"
    StrSQL = StrSQL & "   FROM     dbo.TblVOODetails INNER JOIN"
    StrSQL = StrSQL & "   dbo.TblCountriesGovernments ON dbo.TblVOODetails.FromCity = dbo.TblCountriesGovernments.GovernmentID INNER JOIN"
    StrSQL = StrSQL & "   dbo.TblCountriesGovernments AS TblCountriesGovernments_1 ON dbo.TblVOODetails.TOCity = TblCountriesGovernments_1.GovernmentID"
    StrSQL = StrSQL & "  where TblVOODetails.HID = " & val(ID.text)
    
    Rs_Temp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If Rs_Temp.RecordCount > 0 Then
     Rs_Temp.MoveFirst
     With Grid
        .Rows = Rs_Temp.RecordCount + 1
        Dim j As Integer
        For j = 1 To .Rows - 1
                .TextMatrix(j, .ColIndex("id")) = IIf(IsNull(Rs_Temp("id").value), "", Rs_Temp("id").value)
                .TextMatrix(j, .ColIndex("hid")) = IIf(IsNull(Rs_Temp("hid").value), 0, Rs_Temp("hid").value)
                .TextMatrix(j, .ColIndex("fromcityid")) = IIf(IsNull(Rs_Temp("fromcity").value), "", Rs_Temp("fromcity").value)
                .TextMatrix(j, .ColIndex("tocityid")) = IIf(IsNull(Rs_Temp("tocity").value), "", Rs_Temp("tocity").value)
                .TextMatrix(j, .ColIndex("fromcity")) = IIf(IsNull(Rs_Temp("fromcityname").value), "", Rs_Temp("fromcityname").value)
                .TextMatrix(j, .ColIndex("tocity")) = IIf(IsNull(Rs_Temp("tocityname").value), "", Rs_Temp("tocityname").value)
                .TextMatrix(j, .ColIndex("Remark")) = IIf(IsNull(Rs_Temp("Remarks").value), "", Rs_Temp("Remarks").value)
                .TextMatrix(j, .ColIndex("date")) = IIf(IsNull(Rs_Temp("date").value), "", Rs_Temp("date").value)
                .TextMatrix(j, .ColIndex("time")) = IIf(IsNull(Rs_Temp("time").value), "", Rs_Temp("time").value)
                Rs_Temp.MoveNext
         Next
        End With
    End If
    
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    Exit Sub
ErrTrap:
End Sub

 



Private Sub VehicleNo_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.VehicleNo.text, 1)
End Sub

Private Sub XPBtnMove_Click(Index As Integer)

    'On Error GoTo ErrTrap
    If Me.TxtModFlg.text = "N" Then
        clear_all Me
        Grid.Rows = Grid.FixedRows
        Me.TxtModFlg.text = "R"
        XPBtnMove_Click (1)
        Grid.Rows = Grid.FixedRows
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

Function CuurentLogdata(Optional Currentmode As String)
 

End Function
 
Private Sub SaveData()
    Dim Msg As String
    Dim StrSQL As String
    Dim RsTemp As New ADODB.Recordset
    Dim RsTempM As New ADODB.Recordset
    Dim BeginTrans As Boolean
   ' On Error GoTo ErrTrap

    If Me.TxtModFlg.text <> "R" Then
    
        If Trim(BranchID.BoundText) = "" Then
            If SystemOptions.UserInterface = EnglishInterface Then
                Msg = "Specify Managerial Area"
            Else
                Msg = "Õœœ «·›—⁄ «Ê·« "
            End If

            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            BranchID.SetFocus
   '         SendKeys "{F4}"
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    
        Cn.BeginTrans
        BeginTrans = True

        Select Case Me.TxtModFlg.text

           Case "N"
                rs.AddNew
                ID.text = CStr(new_id("TblVehicleOperatorOrder", "ID", "", True))
           Case "E"
                StrSQL = "delete From TblVOODetails  where  HID =" & val(ID.text)
                Cn.Execute StrSQL, , adExecuteNoRecords
                 
           End Select
   
        rs("ID").value = val(ID.text)
        rs("SDate").value = SDate.value
        rs("BranchID").value = IIf(BranchID.BoundText = "", Null, BranchID.BoundText)
        rs("InClientID").value = IIf(InClientID.BoundText = "", Null, InClientID.BoundText)
        rs("OutClientID").value = IIf(OutClientID.BoundText = "", Null, OutClientID.BoundText)
        rs("GroupID").value = IIf(GroupID.BoundText = "", Null, GroupID.BoundText)
        rs("CompanyID").value = IIf(CompanyID.BoundText = "", Null, CompanyID.BoundText)
        rs("AirLineID").value = IIf(AirLineID.BoundText = "", Null, AirLineID.BoundText)
        rs("AirPortID").value = IIf(AirPortID.BoundText = "", Null, AirPortID.BoundText)
        rs("FlightNo").value = IIf(FlightNo.text = "", Null, Trim(FlightNo.text))
        rs("ArriveDate").value = ArriveDate.value
        rs("ArriveTime").value = ArriveTime.value
        rs("CompnyIn").value = TxtCompnyIn.text
        rs("CompnyOut").value = TxtCompnyOut.text
        
        rs("ProgrammID").value = IIf(ProgrammID.BoundText = "", Null, (ProgrammID.BoundText))
        rs("VehicleNo").value = IIf(VehicleNo.text = "", 0, val(VehicleNo.text))
  '      rs("Model").value = IIf(Model.text = "", 0, Model.text)
        rs("MekkaHotelID").value = IIf(MekkaHotelID.BoundText = "", Null, (MekkaHotelID.BoundText))
        rs("JeddahHotelID").value = IIf(JeddahHotelID.BoundText = "", Null, (JeddahHotelID.BoundText))
        rs("MadinaHotelID").value = IIf(MadinaHotelID.BoundText = "", Null, (MadinaHotelID.BoundText))
        rs("VehicleType").value = IIf(VehicleType.BoundText = "", Null, (VehicleType.BoundText))
        rs("EmpCode").value = IIf(EmpCode.text = "", Null, (EmpCode.text))
        rs("EmpName").value = IIf(EmpName.text = "", Null, (EmpName.text))
        rs("EmpMbile").value = IIf(EmpMbile.text = "", Null, (EmpMbile.text))
        rs("VehicleType").value = IIf(VehicleType.BoundText = "", Null, (VehicleType.BoundText))
        rs("emp").value = emp.value
        rs("other").value = other.value
        rs("FlightNo").value = FlightNo.text
        rs("creationdate").value = Date
        rs("creationuserID").value = user_id
        rs("GroupID").value = IIf(GroupID.BoundText = "", Null, (GroupID.BoundText))
        rs("CompanyID").value = IIf(CompanyID.BoundText = "", Null, (CompanyID.BoundText))
        
        rs("DriverID").value = IIf(DriverID.BoundText = "", Null, (DriverID.BoundText))
        rs("DependID").value = IIf(DependID.text = "", Null, (DependID.text))
        rs("BusNo").value = val(BusNo.text)
        rs("Lounge").value = Lounge.text
        rs.Update
        
        
       Dim Rs_Temp As ADODB.Recordset
        Set Rs_Temp = New ADODB.Recordset
        StrSQL = " select * from TblVOODetails  where 1 = -1 "
        Rs_Temp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
        With Grid
        Dim j As Integer
        For j = 1 To Grid.Rows - 1
           If .TextMatrix(j, .ColIndex("fromcity")) <> "" Then
                    Rs_Temp.AddNew
                    Rs_Temp("ID") = CStr(new_id("TblVOODetails", "ID", "", True))
                    Rs_Temp("HID") = val(ID.text)
                    Rs_Temp("FromCity") = .TextMatrix(j, .ColIndex("FromCityid"))
                    Rs_Temp("ToCity") = .TextMatrix(j, .ColIndex("ToCityid"))
                    Rs_Temp("Date") = IIf(.TextMatrix(j, .ColIndex("Date")) = "", Date, .TextMatrix(j, .ColIndex("Date")))
                    Rs_Temp("Time") = IIf(.TextMatrix(j, .ColIndex("Time")) = "", Date, .TextMatrix(j, .ColIndex("Time")))
                    Rs_Temp("Remarks") = .TextMatrix(j, .ColIndex("Remark"))
                    Rs_Temp("creationdate").value = Date
                    Rs_Temp("creationuserID").value = user_id
                    Rs_Temp.Update
                 End If
           Next
        End With
         
        
        
    
        Dim StrDes As String

     

        Cn.CommitTrans
        BeginTrans = False
        XPTxtCurrent.Caption = rs.AbsolutePosition
        XPTxtCount.Caption = rs.RecordCount
        'CuurentLogdata

        Select Case Me.TxtModFlg.text

            Case "N"

                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â √„—  ‘€Ì· Õ«›·… " & Chr(13)
                    Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
                Else
                    Msg = "Saved" & Chr(13)
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

        TxtModFlg.text = "R"
    End If

    Exit Sub
ErrTrap:

    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If

    If Err.Number = -2147217900 Then
        Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & Chr(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
End Sub

Private Sub Undo()
    On Error GoTo ErrTrap

    Select Case TxtModFlg.text

        Case "N"
            clear_all Me
            Me.TxtModFlg.text = "R"
            XPBtnMove_Click (1)
Grid.Rows = Grid.FixedRows
        Case "E"
            rs.find " ID='" & val(ID.text) & "'", , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Me.TxtModFlg.text = "R"
                Exit Sub
            End If

            Retrive
            Me.TxtModFlg.text = "R"
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub Del_Action()
  
        Dim Msg As String
    Dim StrSQL As String
    Dim RsTemp As New ADODB.Recordset
    Dim StrAccountCode As String
    Dim StrAccountCode1 As String
    Dim ParentAccount As String
    '
 
    On Error GoTo ErrTrap
            
        If ID.text <> "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "”Ì „ Õ–› »Ì«‰«  ÿ·» «·ÕÃ“  —ﬁ„ " & Chr(13)
        Msg = Msg + (ID.text) & Chr(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"
        Else
        Msg = "Delete Booking Request File ? " & Chr(13)
        Msg = Msg + (ID.text) & Chr(13)
        Msg = Msg + "  Are you sure you want to delete ?"
        End If
        
        
        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            If Not rs.RecordCount < 1 Then
                                
                 StrSQL = "delete From TblVOODetails where  HID =" & val(ID.text)
                Cn.Execute StrSQL, , adExecuteNoRecords
                           
                StrSQL = "delete From TblVehicleOperatorOrder where  ID =" & val(ID.text)
                Cn.Execute StrSQL, , adExecuteNoRecords
                 
                 rs.MoveFirst
                    
                   StrSQL = "SELECT  *  From TblVehicleOperatorOrder "
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
        Grid.Rows = Grid.FixedRows
        If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
        Else
         Msg = "this process Not Aailable"
        End If
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        TxtModFlg_Change
        Exit Sub
    End If

    TxtModFlg_Change
    Exit Sub
ErrTrap:
    'If Err.Number = -2147217887 Then
    Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & Chr(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« √„—  ‘€Ì· Õ«›·… "
    Msg = Msg & Chr(13) & Err.description
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
    rs.CancelUpdate
    'End If

End Sub



Private Sub AddTip()
    Dim Wrap As String
    On Error GoTo ErrTrap
    Set TTP = New clstooltip
    Wrap = Chr(13) + Chr(10)

    With TTP
        .Create Me.hWnd, "»Ì«‰«  √„—  ‘€Ì· Õ«›·…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«  √„—  ‘€Ì· Õ«›·… «·ÕÃ“ ÃœÌœ…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  √„—  ‘€Ì· Õ«›·…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  √„—  ‘€Ì· Õ«›·…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  √„—  ‘€Ì· Õ«›·…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  √„—  ‘€Ì· Õ«›·… «·ÃœÌœ" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  √„—  ‘€Ì· Õ«›·…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  √„—  ‘€Ì· Õ«›·…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  Â–« √„—  ‘€Ì· Õ«›·…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  √„—  ‘€Ì· Õ«›·…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(5), "»ÕÀ ..." & Wrap & "···»ÕÀ ⁄‰ Œ“‰…" & Wrap & "Ì‰ÿ»ﬁ ⁄·ÌÂ ‘—Êÿ „⁄Ì‰…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  √„—  ‘€Ì· Õ«›·…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  √„—  ‘€Ì· Õ«›·…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  √„—  ‘€Ì· Õ«›·…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  √„—  ‘€Ì· Õ«›·…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  √„—  ‘€Ì· Õ«›·…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hWnd, "»Ì«‰«  √„—  ‘€Ì· Õ«›·…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
       ' .AddControl CmdHelp, "„”«⁄œ… ..." & Wrap & "·· ⁄—› ⁄·Ï ÊŸÌ›… Â–Â «·‰«›–…" & Wrap & "ÊﬂÌ›Ì… «· ⁄«„· „⁄Â«" & Wrap & "≈÷€ÿ Â‰«" & Wrap, True
    End With

    Exit Sub
ErrTrap:
End Sub


