VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmEditCost 
   ClientHeight    =   10905
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   15915
   LinkTopic       =   "Form1"
   ScaleHeight     =   10905
   ScaleWidth      =   15915
   StartUpPosition =   3  'Windows Default
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   WindowState     =   2  'Maximized
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   10905
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   15915
      _cx             =   28072
      _cy             =   19235
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
      Frame           =   0
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
      Begin C1SizerLibCtl.C1Elastic C1Elastic3 
         Height          =   8280
         Left            =   0
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   1440
         Width           =   15885
         _cx             =   28019
         _cy             =   14605
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
         Begin VB.CheckBox chkBigCost 
            Caption         =   "«’‰«› » ﬂ«·Ì› ﬂ»Ì—…"
            Height          =   255
            Left            =   1080
            TabIndex        =   54
            Top             =   510
            Width           =   1935
         End
         Begin VB.CheckBox chkIsCost 
            Caption         =   "ÌÕ›Ÿ «· ﬂ·›…"
            Height          =   225
            Left            =   4470
            TabIndex        =   53
            Top             =   1230
            Width           =   2325
         End
         Begin VB.CheckBox AllItemHaveOneUnit 
            Caption         =   "«·«’‰«› ﬂ·Â« ÊÕœÂ Ê«ÕœÂ"
            Height          =   375
            Left            =   1080
            TabIndex        =   52
            Top             =   840
            Width           =   2175
         End
         Begin VB.Frame Fra 
            Caption         =   " «—ÌŒ"
            Height          =   645
            Index           =   14
            Left            =   8880
            RightToLeft     =   -1  'True
            TabIndex        =   47
            Top             =   480
            Width           =   6270
            Begin MSComCtl2.DTPicker FrmDate 
               Height          =   330
               Index           =   1
               Left            =   3240
               TabIndex        =   48
               Top             =   210
               Width           =   1590
               _ExtentX        =   2805
               _ExtentY        =   582
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   178782211
               CurrentDate     =   38887
            End
            Begin MSComCtl2.DTPicker ToDate 
               Height          =   330
               Index           =   1
               Left            =   210
               TabIndex        =   49
               Top             =   210
               Width           =   1590
               _ExtentX        =   2805
               _ExtentY        =   582
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   178782211
               CurrentDate     =   38887
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "„‰"
               Height          =   195
               Index           =   27
               Left            =   5100
               RightToLeft     =   -1  'True
               TabIndex        =   51
               Top             =   240
               Width           =   420
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               Caption         =   "≈·Ï"
               Height          =   195
               Index           =   3
               Left            =   2175
               RightToLeft     =   -1  'True
               TabIndex        =   50
               Top             =   240
               Width           =   375
            End
         End
         Begin VB.TextBox ItemCodeTxt 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   6975
            MaxLength       =   40
            RightToLeft     =   -1  'True
            TabIndex        =   40
            Top             =   135
            Width           =   555
         End
         Begin VB.CommandButton ShowItemsData 
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄—÷ «·√’‰«›"
            Height          =   315
            Left            =   4800
            RightToLeft     =   -1  'True
            TabIndex        =   39
            Top             =   600
            Width           =   3825
         End
         Begin MSDataListLib.DataCombo itemNameComp 
            Height          =   315
            Left            =   2760
            TabIndex        =   41
            Top             =   120
            Width           =   4200
            _ExtentX        =   7408
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCboStoreName 
            Height          =   315
            Left            =   8520
            TabIndex        =   42
            Top             =   120
            Width           =   6450
            _ExtentX        =   11377
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VSFlex8Ctl.VSFlexGrid Grid 
            Height          =   6585
            Left            =   120
            TabIndex        =   46
            Top             =   1560
            Width           =   15585
            _cx             =   27490
            _cy             =   11615
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
            SelectionMode   =   3
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   1
            Cols            =   23
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmEditCost.frx":0000
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
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„Œ“‰  "
            Height          =   405
            Index           =   1
            Left            =   14400
            RightToLeft     =   -1  'True
            TabIndex        =   45
            Top             =   135
            Width           =   1260
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·’‰›"
            Height          =   285
            Index           =   30
            Left            =   7200
            RightToLeft     =   -1  'True
            TabIndex        =   44
            Top             =   120
            Width           =   1110
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬂÊœ «·’‰›"
            Height          =   225
            Index           =   35
            Left            =   13950
            RightToLeft     =   -1  'True
            TabIndex        =   43
            Top             =   135
            Visible         =   0   'False
            Width           =   1710
         End
      End
      Begin VB.TextBox txtcode 
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
         Height          =   285
         Left            =   12810
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   990
         Width           =   2205
      End
      Begin VB.Frame FraHeader 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   0
         Width           =   15915
         Begin VB.Frame Frmo2 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   375
            Left            =   4560
            RightToLeft     =   -1  'True
            TabIndex        =   4
            Top             =   120
            Visible         =   0   'False
            Width           =   3105
            Begin MSDataListLib.DataCombo DCUser 
               CausesValidation=   0   'False
               Height          =   315
               Left            =   5865
               TabIndex        =   5
               Tag             =   "„‰ ›÷·ﬂ √œŒ· —ﬁ„ «·ﬁ÷Ì…"
               Top             =   375
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
               TabIndex        =   6
               Top             =   -435
               Width           =   855
            End
         End
         Begin VB.TextBox TxtModFlg 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H0000FF00&
            Enabled         =   0   'False
            Height          =   285
            Left            =   3900
            RightToLeft     =   -1  'True
            TabIndex        =   3
            Text            =   "modflag"
            Top             =   90
            Visible         =   0   'False
            Width           =   465
         End
         Begin VB.TextBox TxtVac_ID 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
            Height          =   240
            Left            =   7920
            RightToLeft     =   -1  'True
            TabIndex        =   2
            Top             =   240
            Visible         =   0   'False
            Width           =   945
         End
         Begin MSComctlLib.ImageList GrdImageList 
            Left            =   3120
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
                  Picture         =   "frmEditCost.frx":0384
                  Key             =   "CompanyName"
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmEditCost.frx":071E
                  Key             =   "Ser"
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmEditCost.frx":0AB8
                  Key             =   "Vac_Name"
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmEditCost.frx":0E52
                  Key             =   "ShareCount"
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmEditCost.frx":11EC
                  Key             =   "Dis_Count"
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmEditCost.frx":1586
                  Key             =   "Bouns"
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmEditCost.frx":1920
                  Key             =   "SharesValue"
               EndProperty
               BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmEditCost.frx":1EBA
                  Key             =   "BuyValue"
               EndProperty
            EndProperty
         End
         Begin ImpulseButton.ISButton btnLast 
            Height          =   315
            Left            =   330
            TabIndex        =   7
            Top             =   150
            Width           =   405
            _ExtentX        =   714
            _ExtentY        =   556
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   ""
            BackColor       =   14871017
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
            ButtonImage     =   "frmEditCost.frx":2254
            ColorButton     =   14871017
            AcclimateGrayTones=   -1  'True
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnNext 
            Height          =   315
            Left            =   795
            TabIndex        =   8
            Top             =   150
            Width           =   405
            _ExtentX        =   714
            _ExtentY        =   556
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   ""
            BackColor       =   14871017
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
            ButtonImage     =   "frmEditCost.frx":25EE
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnPrevious 
            Height          =   315
            Left            =   1395
            TabIndex        =   9
            Top             =   150
            Width           =   405
            _ExtentX        =   714
            _ExtentY        =   556
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   ""
            BackColor       =   14871017
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
            ButtonImage     =   "frmEditCost.frx":2988
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btnFirst 
            Height          =   315
            Left            =   1860
            TabIndex        =   10
            Top             =   150
            Width           =   405
            _ExtentX        =   714
            _ExtentY        =   556
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   ""
            BackColor       =   14871017
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
            ButtonImage     =   "frmEditCost.frx":2D22
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   " ⁄œÌ· ”⁄— «· ﬂ·›… ··√’‰«› "
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
            Height          =   495
            Index           =   2
            Left            =   9840
            RightToLeft     =   -1  'True
            TabIndex        =   11
            Top             =   120
            Width           =   4560
         End
      End
      Begin MSComCtl2.DTPicker DtRecord 
         Height          =   285
         Left            =   8610
         TabIndex        =   13
         Top             =   960
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         CalendarBackColor=   12648447
         CustomFormat    =   "yyyy/M/d"
         Format          =   178782211
         CurrentDate     =   38718
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   510
         Index           =   1
         Left            =   30
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   10440
         Width           =   16125
         _cx             =   28443
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
         Appearance      =   0
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
            Height          =   390
            Index           =   0
            Left            =   14835
            TabIndex        =   15
            Top             =   60
            Width           =   1110
            _ExtentX        =   1958
            _ExtentY        =   688
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
            Height          =   390
            Index           =   1
            Left            =   13065
            TabIndex        =   16
            Top             =   60
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   688
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
            Height          =   390
            Index           =   2
            Left            =   11265
            TabIndex        =   17
            Top             =   60
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   688
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
            Height          =   390
            Index           =   3
            Left            =   9570
            TabIndex        =   18
            Top             =   60
            Width           =   1110
            _ExtentX        =   1958
            _ExtentY        =   688
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
            Height          =   390
            Index           =   4
            Left            =   7605
            TabIndex        =   19
            Top             =   60
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   688
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
            Height          =   390
            Index           =   5
            Left            =   5865
            TabIndex        =   20
            Top             =   60
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   688
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
         Begin ImpulseButton.ISButton Cmd 
            Height          =   390
            Index           =   6
            Left            =   555
            TabIndex        =   21
            Top             =   60
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   688
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
            Height          =   390
            Index           =   7
            Left            =   4365
            TabIndex        =   22
            Top             =   60
            Width           =   1020
            _ExtentX        =   1799
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
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
         End
         Begin ImpulseButton.ISButton CmdHelp 
            Height          =   390
            Left            =   2340
            TabIndex        =   23
            Top             =   60
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   688
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "„”«⁄œ…"
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
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic2 
         Height          =   645
         Left            =   150
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   9630
         Width           =   15915
         _cx             =   28072
         _cy             =   1138
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
         CaptionPos      =   7
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
         Frame           =   0
         FrameStyle      =   5
         FrameWidth      =   1
         FrameColor      =   -2147483628
         FrameShadow     =   -2147483632
         FloodStyle      =   1
         _GridInfo       =   ""
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   9
         Begin MSDataListLib.DataCombo DCboUserName 
            Height          =   315
            Left            =   10785
            TabIndex        =   25
            Top             =   210
            Width           =   2985
            _ExtentX        =   5265
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            ListField       =   "6"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label XPTxtCurrent 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Height          =   345
            Left            =   8340
            RightToLeft     =   -1  'True
            TabIndex        =   30
            Top             =   210
            Width           =   675
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "⁄œœ «·”Ã·« :"
            Height          =   315
            Index           =   2
            Left            =   6960
            RightToLeft     =   -1  'True
            TabIndex        =   29
            Top             =   225
            Width           =   1440
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·”Ã· «·Õ«·Ì:"
            Height          =   315
            Index           =   0
            Left            =   9120
            RightToLeft     =   -1  'True
            TabIndex        =   28
            Top             =   225
            Width           =   1395
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "»Ê«”ÿ…"
            Height          =   300
            Index           =   4
            Left            =   13770
            RightToLeft     =   -1  'True
            TabIndex        =   27
            Top             =   210
            Width           =   1530
         End
         Begin VB.Label XPTxtCount 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Height          =   315
            Left            =   5775
            RightToLeft     =   -1  'True
            TabIndex        =   26
            Top             =   210
            Width           =   765
         End
      End
      Begin MSDataListLib.DataCombo Dcbranch 
         Bindings        =   "frmEditCost.frx":30BC
         Height          =   315
         Left            =   2865
         TabIndex        =   31
         Top             =   960
         Width           =   4605
         _ExtentX        =   8123
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
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " «—ÌŒ «·Õ—ﬂ…"
         Height          =   255
         Index           =   0
         Left            =   10755
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   1005
         Width           =   1530
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„"
         Height          =   255
         Index           =   3
         Left            =   15225
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   990
         Width           =   585
      End
      Begin VB.Label lblBr 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«·›—⁄"
         Height          =   255
         Left            =   7095
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   960
         Width           =   1035
      End
   End
   Begin C1SizerLibCtl.C1Elastic Ele 
      Height          =   1065
      Index           =   2
      Left            =   1320
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   1200
      Width           =   2355
      _cx             =   4154
      _cy             =   1879
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
      CaptionPos      =   7
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
      Frame           =   0
      FrameStyle      =   5
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
      Height          =   30
      Index           =   4
      Left            =   0
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   0
      Width           =   2355
      _cx             =   4154
      _cy             =   53
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
      CaptionPos      =   7
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
      Frame           =   0
      FrameStyle      =   5
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
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·„Œ“‰  "
      Height          =   285
      Index           =   29
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   37
      Top             =   2880
      Width           =   1740
   End
End
Attribute VB_Name = "frmEditCost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'########################## Genrale Var Declaration ######################
Dim StrSQL  As String
Dim rs As ADODB.Recordset

Private Sub ChangeLang() '%%%%% Convert window object lung to english %%%%%%
Label1(2).Caption = "Edit Items Costs"
Label1(3).Caption = "Ser"
Label1(0).Caption = "Process date"
ShowItemsData.Caption = "Show Items"
With Grid
    .TextMatrix(0, .ColIndex("TransactionTypeName")) = "Transaction Name"
    .TextMatrix(0, .ColIndex("Transaction_Date")) = "Date"
    .TextMatrix(0, .ColIndex("Quantity")) = "Quantity"
    .TextMatrix(0, .ColIndex("Price")) = "Price"
    .TextMatrix(0, .ColIndex("OldQty")) = "Old Qty"
    .TextMatrix(0, .ColIndex("OldCost")) = "Old Cost"
    .TextMatrix(0, .ColIndex("NewQty")) = "New Qty"
    .TextMatrix(0, .ColIndex("NewCost")) = "New Cost"
    
    .TextMatrix(0, .ColIndex("Fullcode")) = "Item Code"
    .TextMatrix(0, .ColIndex("ItemName")) = "Item Name"
    .TextMatrix(0, .ColIndex("ItemNamee")) = "Item Name"
End With
'lblde.Caption = "Filter results by"
lbl(0).Caption = "Current Record"
lbl(2).Caption = "No. of Records"
lbl(35).Caption = "Item Code"
lbl(30).Caption = "Item Name"
lbl(1).Caption = "Store"
lblBr.Caption = "Branch"
Label1(4).Caption = "By"
Cmd(0).Caption = "New"
Cmd(1).Caption = "Edit"
Cmd(2).Caption = "Save"
Cmd(3).Caption = "Cancel"
Cmd(4).Caption = "Delete"
Cmd(5).Caption = "Search"
Cmd(7).Caption = "Print"
Cmd(6).Caption = "Exit"
CmdHelp.Caption = "Help"
End Sub

Private Sub Undo()    '%%%%%%%% Undo Enteries and clear all fields also set text mode to R %%%%%%%%
    '@@@@@@@@ declear Var @@@@@@@@@
    Dim Msg As String
    '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    
    'Handling an exception
    On Error GoTo ErrTrap
    
    Select Case TxtModFlg.text
        Case "N"
        
              If SystemOptions.UserInterface = EnglishInterface Then
                Msg = "This process will be undone."
                Msg = Msg & CHR(13) & "do you want to continue"
            Else
                Msg = "”Ê› Ì „ «· —«Ã⁄ ›Ï  ”ÃÌ· Â–Â «·⁄„·Ì… .."
                Msg = Msg & CHR(13) & "›Â· «‰  „ «ﬂœ „‰ «·√” „—«— ..!!"
            End If
          
            If MsgBox(Msg, vbQuestion + vbYesNo + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                clear_all Me
                Me.TxtModFlg.text = "R"
                BtnLast_Click
            End If
        Case "E"
        
            If SystemOptions.UserInterface = EnglishInterface Then
                Msg = "This process will be undone."
                Msg = Msg & CHR(13) & "do you want to continue"
            Else
                Msg = "”Ê› Ì „ «· —«Ã⁄ ›Ï  ”ÃÌ· Â–Â «·⁄„·Ì… .."
                Msg = Msg & CHR(13) & "›Â· «‰  „ «ﬂœ „‰ «·√” „—«— ..!!"
            End If
            
            If MsgBox(Msg, vbQuestion + vbYesNo + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                'rs.Find "Transaction_ID='" & val(XPTxtBillID.Text) & "'", , adSearchForward, adBookmarkFirst
                'If Not rs.EOF Or rs.BOF Then
                    Me.TxtModFlg.text = "R"
                    BtnLast_Click
                    'Retrive
                'End If
            End If
    End Select
    
    'get data again
    Retrive
    
ErrTrap:
End Sub

'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'///////////////////////////////////////////////////// 10/07/2017 ///////////////////////////////////////////////
Private Sub GetItemData(Optional LngItemID As Long = 0, _
                        Optional StrItemCode As String = "")

    Dim StrSQL As String
    Dim RsTemp As New ADODB.Recordset
    On Error GoTo ErrTrap

    If LngItemID = 0 And StrItemCode <> "" Then
        StrSQL = "select * From TblItems where ItemCode='" & StrItemCode & "'"
        RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (RsTemp.EOF Or RsTemp.BOF) Then
                itemNameComp.BoundText = RsTemp("ItemID").value
            'Cmd_Click (0)
        Else
            DCboItemsName.BoundText = ""
        End If

        If Me.Tag <> "" Then
            'Cmd_Click (0)
            Me.Tag = ""
        End If

    ElseIf LngItemID <> 0 And StrItemCode = "" Then
        StrSQL = "select * From TblItems where ItemID=" & LngItemID
        RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (RsTemp.EOF Or RsTemp.BOF) Then
                itemNameComp.BoundText = RsTemp("ItemID").value
        End If

        'Cmd_Click (0)
    End If

    Exit Sub
ErrTrap:
End Sub



Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With Grid
Select Case .ColKey(Col)
Case "TransactionTypeName"
Cancel = True
Case "Transaction_Date"
Cancel = True
Case "Fullcode"
Cancel = True
Case "ItemName"
Cancel = True
Case "OldQty"
Cancel = True
Case "OldCost"
Cancel = True
End Select
End With

End Sub

Private Sub itemNameComp_Click(Area As Integer)
On Error Resume Next
Dim StrItemCode As String
    Dim rs As ADODB.Recordset
    Dim StrSQL As String

    If Me.itemNameComp.BoundText = "" Then
        Me.ItemCodeTxt.text = ""
    Else
        StrSQL = "Select * From TblItems Where ItemID =" & Me.itemNameComp.BoundText & " "
        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (rs.BOF Or rs.EOF) Then
            StrItemCode = IIf(IsNull(rs("ItemCode").value), "", rs("ItemCode").value)
        End If
        If StrItemCode <> Trim(Me.ItemCodeTxt.text) Then
            Me.ItemCodeTxt.text = StrItemCode
        End If
        rs.Close
        Set rs = Nothing
    End If

End Sub

Private Sub ItemCodeTxt_Change()
GetItemData 0, Trim(Me.ItemCodeTxt.text)
End Sub
'#################################################################
Private Sub Form_Load() ' %%%%%% windows start %%%%%%%
    Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos
    '############################### Set Icons for bottom Bar #############################
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
    Set Cmd(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
    Set Cmd(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
    Set Cmd(5).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    Set Cmd(7).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Print").Picture
    Set CmdHelp.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Help").Picture
    'Set CmdConvert.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Excute").Picture
    'Set CmdTemplate.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Excute").Picture
    'Set Accredit.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Required").Picture
    '######################################################################################
    
    '################################## Change The lung ###################################
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
    '######################################################################################
    '############################### Defult window state ##################################
    
    '######################################################################################
    
    '########################## change windows state to read ##############################
    TxtModFlg.text = "R"
    '######################################################################################
    
    '########################## Get data for all list and combos ##########################
     Dcombos.GetItemsNames Me.itemNameComp, 0
     Dcombos.GetStores Me.DCboStoreName
     Dcombos.GetBranches Me.Dcbranch
     Dcombos.GetUsers Me.DCboUserName
    '######################################################################################
     Dcbranch.BoundText = Current_branch
    '################################# Get the last recored ###############################
    Set rs = New ADODB.Recordset
    StrSQL = "select * From TblEditItemCost"
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If Not (rs.BOF Or rs.EOF) Then
        rs.MoveLast
        txtcode.text = IIf(IsNull(rs("ID").value), 0, rs("ID").value)
    End If
    Retrive
    
    '######################################################################################
End Sub
Function UpdateInputCost(WorkOrderNO As String) As Double
'    StrSQL = "SELECT    dbo.Transactions.NoteSerial,  dbo.Transactions.NoteSerial1,    dbo.Transactions.Transaction_Serial, dbo.Transactions.Transaction_ID, dbo.Transactions.Transaction_Type, dbo.Transaction_Details.Item_ID, "
'    StrSQL = StrSQL + " dbo.Transactions.WorkOrderNO, dbo.TblItems.ItemCode, dbo.TblItems.ItemName, dbo.Transaction_Details.ColorID, dbo.Transaction_Details.ItemSize,"
'    StrSQL = StrSQL + " dbo.Transaction_Details.UnitId, dbo.Transaction_Details.ShowQty,showPrice, dbo.Transaction_Details.QtyBySmalltUnit, dbo.Transaction_Details.ClassId,"
'    StrSQL = StrSQL + " dbo.Transaction_Details.Price ,Transaction_Details.ItemID2 , dbo.TblUnites.UnitName"
'    StrSQL = StrSQL + "  ,ShowQty*showPrice  as Costs,ItemName2 = (Select ItemName From TblItems AS ti2 Where ti2.ItemId =Transaction_Details.ItemID2 ) "
'    StrSQL = StrSQL + "  FROM         dbo.Transactions INNER JOIN"
'    StrSQL = StrSQL + " dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID INNER JOIN"
'    StrSQL = StrSQL + " dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID INNER JOIN"
'    StrSQL = StrSQL + " dbo.TblUnites ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID"
'
'

If val(WorkOrderNO) = 0 Then UpdateInputCost = 0: Exit Function
    If WorkOrderNO = "" Then UpdateInputCost = 0: Exit Function
    Dim RsDev As New ADODB.Recordset
    Dim RsDev2 As New ADODB.Recordset
    StrSQL = " SELECT"
    StrSQL = StrSQL + "     FactoryExpenses"
    
   
    StrSQL = StrSQL + " From dbo.transactions"
  
 

    StrSQL = StrSQL + " where  (dbo.transactions.Transaction_Type = 26"
    StrSQL = StrSQL + "   and  IsNull(Transactions.NoteSerial1,'') = '" & Trim(WorkOrderNO) & "' ) "
      
    RsDev2.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
     Dim FactoryExpenses As Double
  If Not RsDev2.EOF Then
        FactoryExpenses = val(RsDev2!FactoryExpenses & "")
    Else
        FactoryExpenses = 0
        Exit Function
    
   End If
     
    StrSQL = " SELECT"
    StrSQL = StrSQL + "     dbo.transactions.Transaction_Type"
    StrSQL = StrSQL + "    ,dbo.Transactions.WorkOrderNO"
    StrSQL = StrSQL + "    ,SUM(dbo.Transaction_Details.ShowQty * dbo.Transaction_Details.showPrice) AS cost"
    StrSQL = StrSQL + " From dbo.transactions"
    StrSQL = StrSQL + " INNER JOIN dbo.Transaction_Details"
    StrSQL = StrSQL + "     ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID"
    StrSQL = StrSQL + " GROUP BY dbo.Transactions.Transaction_Type"
    StrSQL = StrSQL + "         ,dbo.Transactions.WorkOrderNO"
    StrSQL = StrSQL + " Having (dbo.transactions.Transaction_Type = 27"
    StrSQL = StrSQL + "   and  IsNull(Transactions.WorkOrderNO,'') = '" & Trim(WorkOrderNO) & "' ) "
     
 
    
 
 
    
    RsDev.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    
   
    
    If Not RsDev.EOF Then
        UpdateInputCost = val(RsDev!cost & "") + FactoryExpenses
    Else
        UpdateInputCost = FactoryExpenses
    End If
 End Function
Private Sub ShowItemsData_Click()
    '@@@@@@@@ declear Var @@@@@@@@@
    Dim RsDev As New ADODB.Recordset
    Dim i As Double
    Dim Item_ID As Double
    Dim OLDItem_ID As Double
    Dim StockEffect As Double
    Dim sql As String
    Dim UnitFactor As Double
    Dim QtyBySmalltUnit As Double
    Dim SecOrder As Integer
    Dim rsItemCostUnit As ADODB.Recordset
    Dim SecOrderCurrent As Integer
    Dim UnitFactorCurrent As Double
                                
    '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'    sql = "Update Transaction_Details set OldQty = 0,OldCost = 0 ,NewQty = 0,NewCost = 0,Price = 0"
'    sql = sql & " where Transaction_Details.Transaction_ID In (Select Transactions.Transaction_ID from Transactions  "
'    sql = sql & " Inner join dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type "
'    sql = sql & " where 1 = 1 "
'    sql = sql & "  and (dbo.TransactionTypes.StockEffect = -1  )"
'    If DCboStoreName.BoundText <> "" Then
'         sql = sql & " and (dbo.Transactions.StoreID = " & DCboStoreName.BoundText & " ) "
'    End If
'
'    If Not IsNull(FrmDate(1).value) Then
'        sql = sql & "  And dbo.transactions.Transaction_Date >=" & SQLDate(FrmDate(1).value, True) & ""
'    End If
'    If Not IsNull(ToDate(1).value) Then
'        sql = sql & "  And dbo.transactions.Transaction_Date <=" & SQLDate(ToDate(1).value, True) & ""
'    End If
'    sql = sql & "  )"
'
'    If itemNameComp.BoundText <> "" And (itemNameComp.Text) <> "" Then
'         sql = sql & " and (dbo.Transaction_Details.Item_ID = " & itemNameComp.BoundText & " ) "
'    End If
    
'    Cn.Execute sql
            
            
sql = "Update Transaction_Details set OldQty = 0,OldCost = 0 ,NewQty = 0,NewCost = 0"
    sql = sql & " where Transaction_Details.Transaction_ID In (Select Transactions.Transaction_ID from Transactions  "
    sql = sql & " Inner join dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type "
    sql = sql & " where 1 = 1 "
    sql = sql & "  and (dbo.TransactionTypes.StockEffect <> 0  )"
    If DCboStoreName.BoundText <> "" Then
         sql = sql & " and (dbo.Transactions.StoreID = " & DCboStoreName.BoundText & " ) "
    End If
    
    If Not IsNull(FrmDate(1).value) Then
        sql = sql & "  And dbo.transactions.Transaction_Date >=" & SQLDate(FrmDate(1).value, True) & ""
    End If
    If Not IsNull(ToDate(1).value) Then
        sql = sql & "  And dbo.transactions.Transaction_Date <=" & SQLDate(ToDate(1).value, True) & ""
    End If
    sql = sql & "  )"
    
    If itemNameComp.BoundText <> "" And (itemNameComp.text) <> "" Then
         sql = sql & " and (dbo.Transaction_Details.Item_ID = " & itemNameComp.BoundText & " ) "
    End If
    
    Cn.Execute sql
            
If Me.TxtModFlg.text <> "R" Then
    Me.Grid.Rows = 1
    
          sql = "  SELECT   Transactions.WorkOrderNO ,Transactions.CBoBasedON, Transactions.NoteSerial1 ,Transaction_Details.QtyBySmalltUnit ,dbo.TransactionTypes.StockEffect, dbo.Transactions.Transaction_ID, dbo.Transactions.Transaction_Date, dbo.Transaction_Details.Item_ID, "
    sql = sql & " dbo.Transactions.StoreID, dbo.Transaction_Details.Quantity, dbo.Transaction_Details.Price, dbo.Transaction_Details.OldQty, dbo.Transaction_Details.OldCost,"
    sql = sql & "  dbo.Transaction_Details.showprice, dbo.Transaction_Details.NewQty, dbo.Transaction_Details.NewCost, Transaction_Details.UnitID,dbo.Transactions.Transaction_Type, dbo.TransactionTypes.TransactionTypeName,"
    sql = sql & " dbo.TblItems.fullcode , dbo.TblItems.itemname, dbo.TblItems.ItemNamee"
    sql = sql & " FROM dbo.Transactions INNER JOIN"
    sql = sql & " dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID INNER JOIN"
    sql = sql & " dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type INNER JOIN"
    sql = sql & " dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID"
     
        sql = sql & " Where (dbo.TransactionTypes.StockEffect <> 0   ) "
        'sql = sql & " Where (dbo.TransactionTypes.StockEffect <> 0   and ItemID=205) "
       If chkBigCost.value = vbChecked Then
            sql = sql & " and (dbo.Transaction_Details.Price > 9000 ) "
       End If
    If itemNameComp.BoundText <> "" Then
         sql = sql & " and (dbo.Transaction_Details.Item_ID = " & itemNameComp.BoundText & " ) "
    End If
    If DCboStoreName.BoundText <> "" Then
         sql = sql & " and (dbo.Transactions.StoreID = " & DCboStoreName.BoundText & " ) "
    End If
    
    If Not IsNull(FrmDate(1).value) Then
sql = sql & "  And dbo.transactions.Transaction_Date >=" & SQLDate(FrmDate(1).value, True) & ""
End If
If Not IsNull(ToDate(1).value) Then
sql = sql & "  And dbo.transactions.Transaction_Date <=" & SQLDate(ToDate(1).value, True) & ""
End If
  If SystemOptions.UserInterface = ArabicInterface Then
        X = MsgBox("Â·  —Ìœ   — Ì» ÿ»ﬁ« · — Ì» «·Õ—ﬂÂ    ", vbExclamation + vbYesNoCancel)
    Else
        X = MsgBox("Do you want to upload photo from file", vbExclamation + vbYesNo)
    End If
    If X = vbYes Then
     sql = sql & " ORDER BY  item_ID, dbo.Transactions.Transaction_Date , dbo.Transactions.Transaction_ID "
    Else
     sql = sql & " ORDER BY  item_ID, dbo.Transactions.Transaction_Date,dbo.TransactionTypes.StockEffect desc, dbo.Transactions.Transaction_ID "
    End If
    

   
    
    RsDev.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
    
    Dim rsItemCost As New ADODB.Recordset
    Dim mmUnitPurPrice As Double
    If Not (RsDev.BOF Or RsDev.EOF) Then
        RsDev.MoveFirst
        With Me.Grid
            .Rows = .FixedRows + RsDev.RecordCount
            For i = .FixedRows To .Rows - 1
            If val(RsDev("UnitID").value) = 3 Then
                i = i
            End If
         '  If val(RsDev!Transaction_ID & "") = 146353 Then
         '            i = i
         '   End If
          
                s = " Select UnitPurPrice,UnitFactor,SecOrder from TblItemsUnits where ItemID = " & val(RsDev!Item_ID & "") & "  and UnitId = " & val(RsDev("UnitID").value & "")
                Set rsItemCostUnit = New ADODB.Recordset
                rsItemCostUnit.Open s, Cn, adOpenStatic, adLockReadOnly
                If Not rsItemCostUnit.EOF Then
                    mmUnitPurPrice = val(rsItemCostUnit!UnitPurPrice & "")
                    SecOrderCurrent = val(rsItemCostUnit!SecOrder & "")
                    UnitFactorCurrent = val(rsItemCostUnit!UnitFactor & "")
                    QtyBySmalltUnit = UnitFactorCurrent
                     SecOrder = SecOrderCurrent
                End If
              .TextMatrix(i, .ColIndex("Index")) = i

                .TextMatrix(i, .ColIndex("Transaction_Type")) = IIf(IsNull(RsDev("Transaction_Type").value), 0, (RsDev("Transaction_Type").value))
                .TextMatrix(i, .ColIndex("NoteSerial1")) = IIf(IsNull(RsDev("NoteSerial1").value), 0, (RsDev("NoteSerial1").value))
                .TextMatrix(i, .ColIndex("QtyBySmalltUnit")) = IIf(IsNull(RsDev("QtyBySmalltUnit").value), 1, (RsDev("QtyBySmalltUnit").value))
                .TextMatrix(i, .ColIndex("UnitID")) = IIf(IsNull(RsDev("UnitID").value), 1, (RsDev("UnitID").value))

                .TextMatrix(i, .ColIndex("TransactionTypeName")) = IIf(IsNull(RsDev("TransactionTypeName").value), 0, (RsDev("TransactionTypeName").value))
                .TextMatrix(i, .ColIndex("StockEffect")) = IIf(IsNull(RsDev("StockEffect").value), 0, (RsDev("StockEffect").value))
                StockEffect = val(.TextMatrix(i, .ColIndex("StockEffect")))
                .TextMatrix(i, .ColIndex("Transaction_Date")) = IIf(IsNull(RsDev("Transaction_Date").value), 0, (RsDev("Transaction_Date").value))
                .TextMatrix(i, .ColIndex("Transaction_ID")) = IIf(IsNull(RsDev("Transaction_ID").value), 0, (RsDev("Transaction_ID").value))
                .TextMatrix(i, .ColIndex("Item_ID")) = IIf(IsNull(RsDev("Item_ID").value), 0, (RsDev("Item_ID").value))
                Item_ID = .TextMatrix(i, .ColIndex("Item_ID"))
                .TextMatrix(i, .ColIndex("StoreID")) = IIf(IsNull(RsDev("StoreID").value), 0, (RsDev("StoreID").value))
                .TextMatrix(i, .ColIndex("Quantity")) = IIf(IsNull(RsDev("Quantity").value), 0, (RsDev("Quantity").value))
                .TextMatrix(i, .ColIndex("Price")) = IIf(IsNull(RsDev("Price").value), 0, (RsDev("Price").value))
                .TextMatrix(i, .ColIndex("Price")) = Round(.TextMatrix(i, .ColIndex("Price")), 3)
                .TextMatrix(i, .ColIndex("OldQty")) = IIf(IsNull(RsDev("OldQty").value), 0, (RsDev("OldQty").value))
                .TextMatrix(i, .ColIndex("OldCost")) = IIf(IsNull(RsDev("OldCost").value), 0, (RsDev("OldCost").value))
               ' .TextMatrix(i, .ColIndex("NewQty")) = IIf(IsNull(RsDev("NewQty").Value), 0, (RsDev("NewQty").Value))
               ' .TextMatrix(i, .ColIndex("NewCost")) = IIf(IsNull(RsDev("NewCost").Value), 0, (RsDev("NewCost").Value))
                .TextMatrix(i, .ColIndex("Fullcode")) = IIf(IsNull(RsDev("Fullcode").value), 0, (RsDev("Fullcode").value))
                .TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(RsDev("ItemName").value), 0, (RsDev("ItemName").value))
                .TextMatrix(i, .ColIndex("ItemNamee")) = IIf(IsNull(RsDev("ItemNamee").value), 0, (RsDev("ItemNamee").value))
                 .TextMatrix(i, .ColIndex("CBoBasedON")) = IIf(IsNull(RsDev("CBoBasedON").value), 0, (RsDev("CBoBasedON").value))
                 .TextMatrix(i, .ColIndex("WorkOrderNO")) = IIf(IsNull(RsDev("WorkOrderNO").value), "", (RsDev("WorkOrderNO").value))
                 

              
                If OLDItem_ID <> Item_ID Then
  
              .TextMatrix(i, .ColIndex("OldQty")) = 0
            .TextMatrix(i, .ColIndex("OldCost")) = 0
'newwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwwww
   .TextMatrix(i, .ColIndex("OldQty")) = ItemOpeningBalances(CLng(Item_ID), val(DCboStoreName.BoundText), , , , FrmDate(1).value)

 .TextMatrix(i, .ColIndex("OldCost")) = ModItemCostPrice.GetCostItemPrice(CLng(Item_ID), 0, "", , SystemOptions.SysMainStockCostMethod, , DateAdd("d", -1, CDate(RsDev!Transaction_Date & "")), DateAdd("d", -1, CDate(RsDev!Transaction_Date & "")), , val(.TextMatrix(i, .ColIndex("UnitID"))), val(DCboStoreName.BoundText))

        GoTo newItem
  Else
        If UnitFactorCurrent = 1 And val(.TextMatrix(i - 1, .ColIndex("QtyBySmalltUnit"))) <> 1 Then
        ElseIf UnitFactorCurrent <> 1 And val(.TextMatrix(i - 1, .ColIndex("QtyBySmalltUnit"))) = 1 Then
            
        ElseIf UnitFactorCurrent <> 1 And val(.TextMatrix(i - 1, .ColIndex("QtyBySmalltUnit"))) <> 1 Then
        ElseIf UnitFactorCurrent = 1 And val(.TextMatrix(i - 1, .ColIndex("QtyBySmalltUnit"))) = 1 Then
        End If
        
        
        
        .TextMatrix(i, .ColIndex("OldQty")) = .TextMatrix(i - 1, .ColIndex("NewQty"))
            
            
            
            .TextMatrix(i, .ColIndex("OldCost")) = .TextMatrix(i - 1, .ColIndex("NewCost"))
newItem:
            If StockEffect = 1 Then
             .TextMatrix(i, .ColIndex("NewQty")) = val(.TextMatrix(i, .ColIndex("Quantity"))) + val(.TextMatrix(i, .ColIndex("OldQty")))
             
  'If val(.TextMatrix(i, .ColIndex("Price"))) = 0 Then
  '                   .TextMatrix(i, .ColIndex("Price")) = val(.TextMatrix(i, .ColIndex("OldCost")))
  '              End If
                
                If val(.TextMatrix(i, .ColIndex("Transaction_Type"))) = 11 Or val(.TextMatrix(i, .ColIndex("Transaction_Type"))) = 15 Or (val(.TextMatrix(i, .ColIndex("Transaction_Type"))) = 20 And val(.TextMatrix(i, .ColIndex("CBoBasedON"))) <> 5) Then
                     If val(.TextMatrix(i, .ColIndex("Transaction_Type"))) <> 15 Then
                         If val(.TextMatrix(i, .ColIndex("OldCost"))) <> 0 Then
                        .TextMatrix(i, .ColIndex("Price")) = val(.TextMatrix(i, .ColIndex("OldCost")))
                        End If
                    ElseIf val(.TextMatrix(i, .ColIndex("Transaction_Type"))) = 15 Then
                            .TextMatrix(i, .ColIndex("NewCost")) = val(.TextMatrix(i, .ColIndex("Price")))
                    
                    End If
                     
                     
                     
                     If val(.TextMatrix(i, .ColIndex("Price"))) = 0 Then
                      .TextMatrix(i, .ColIndex("Price")) = getcostbuylastinvoice(CDbl(val(RsDev("Item_ID").value & "")), CDate(RsDev!Transaction_Date & ""), val(RsDev("UnitID").value & ""), UnitFactor, SecOrder)
                        If SecOrderCurrent > SecOrder Then
                            If UnitFactorCurrent < 1 Then
                                .TextMatrix(i, .ColIndex("Price")) = val(.TextMatrix(i, .ColIndex("Price"))) * UnitFactorCurrent
                            Else
                                .TextMatrix(i, .ColIndex("Price")) = val(.TextMatrix(i, .ColIndex("Price"))) / UnitFactorCurrent
                            End If
                        ElseIf SecOrderCurrent < SecOrder Then
                            .TextMatrix(i, .ColIndex("Price")) = val(.TextMatrix(i, .ColIndex("Price"))) * UnitFactor
                        Else
                            .TextMatrix(i, .ColIndex("Price")) = val(.TextMatrix(i, .ColIndex("Price"))) * QtyBySmalltUnit
                        End If
                      End If
                     
                     
                     If val(.TextMatrix(i, .ColIndex("Price"))) = 0 Then
                     .TextMatrix(i, .ColIndex("Price")) = ModItemCostPrice.GetCostItemPrice(CLng(Item_ID), 0, "", , SystemOptions.SysMainStockCostMethod, , DateAdd("d", -1, CDate(RsDev!Transaction_Date & "")), DateAdd("d", -1, FrmDate(1).value), , val(.TextMatrix(i, .ColIndex("UnitID"))), val(DCboStoreName.BoundText))
                     End If
                     
                       
                   
 End If
    
   '«·«‰ «Ã «· «„
            If val(.TextMatrix(i, .ColIndex("Transaction_Type"))) = 28 Then
                     .TextMatrix(i, .ColIndex("Price")) = UpdateInputCost(Trim(.TextMatrix(i, .ColIndex("WorkOrderNO"))))
                     If val(.TextMatrix(i, .ColIndex("Quantity"))) <> 0 Then
                     
                     .TextMatrix(i, .ColIndex("Price")) = Round(.TextMatrix(i, .ColIndex("Price")) / val(.TextMatrix(i, .ColIndex("Quantity"))), 3)
        End If

If .TextMatrix(i, .ColIndex("Price")) = 0 Then
 .TextMatrix(i, .ColIndex("Price")) = val(.TextMatrix(i, .ColIndex("OldCost")))
 End If
 
            End If
               
               
           
               
                        If val(.TextMatrix(i, .ColIndex("OldQty"))) > 0 Then
                            If val(.TextMatrix(i, .ColIndex("Transaction_Type"))) = 15 Then
                            
                            .TextMatrix(i, .ColIndex("NewCost")) = val(.TextMatrix(i, .ColIndex("Price")))
                            Else
                          .TextMatrix(i, .ColIndex("NewCost")) = (val(.TextMatrix(i, .ColIndex("Quantity"))) * val(.TextMatrix(i, .ColIndex("Price"))) + val(.TextMatrix(i, .ColIndex("OldQty"))) * val(.TextMatrix(i, .ColIndex("OldCost")))) / (val(.TextMatrix(i, .ColIndex("Quantity"))) + val(.TextMatrix(i, .ColIndex("OldQty"))))
                          End If
                       ElseIf val(.TextMatrix(i, .ColIndex("OldQty"))) = 0 Then
                          .TextMatrix(i, .ColIndex("NewCost")) = val(.TextMatrix(i, .ColIndex("Price")))
                        Else
                             .TextMatrix(i, .ColIndex("NewCost")) = 0
                        End If
'              Debug.Print .TextMatrix(i, .ColIndex("NewCost"))
               
               ' If val(.TextMatrix(i, .ColIndex("Transaction_Type"))) = 15 And val(.TextMatrix(i, .ColIndex("NewCost"))) = 0 Then
               '      .TextMatrix(i, .ColIndex("NewCost")) = val(.TextMatrix(i, .ColIndex("OldCost")))
               ' End If
                
               ' If (.TextMatrix(i, .ColIndex("OldQty"))) < 0 Then
               '     .TextMatrix(i, .ColIndex("NewCost")) = -9999
               ' End If


            Else
            
            .TextMatrix(i, .ColIndex("NewQty")) = val(.TextMatrix(i, .ColIndex("OldQty"))) - val(.TextMatrix(i, .ColIndex("Quantity")))
              .TextMatrix(i, .ColIndex("NewCost")) = val(.TextMatrix(i, .ColIndex("OldCost")))
               If SecOrderCurrent > SecOrder Then
                If UnitFactorCurrent < 1 Then
                    .TextMatrix(i, .ColIndex("NewCost")) = val(.TextMatrix(i, .ColIndex("NewCost"))) * UnitFactorCurrent
                    .TextMatrix(i, .ColIndex("Price")) = val(.TextMatrix(i, .ColIndex("OldCost"))) * UnitFactorCurrent
                Else
                    .TextMatrix(i, .ColIndex("NewCost")) = val(.TextMatrix(i, .ColIndex("NewCost"))) / UnitFactorCurrent
                    .TextMatrix(i, .ColIndex("Price")) = val(.TextMatrix(i, .ColIndex("OldCost"))) / UnitFactorCurrent
                End If
            ElseIf SecOrderCurrent < SecOrder Then
                .TextMatrix(i, .ColIndex("NewCost")) = val(.TextMatrix(i, .ColIndex("NewCost"))) * UnitFactor
                .TextMatrix(i, .ColIndex("Price")) = val(.TextMatrix(i, .ColIndex("OldCost"))) * UnitFactor
            Else
                .TextMatrix(i, .ColIndex("NewCost")) = val(.TextMatrix(i, .ColIndex("NewCost"))) * QtyBySmalltUnit
                .TextMatrix(i, .ColIndex("Price")) = val(.TextMatrix(i, .ColIndex("OldCost"))) * QtyBySmalltUnit
            End If
              If .TextMatrix(i, .ColIndex("NewQty")) = 0 Then
                .TextMatrix(i, .ColIndex("NewCost")) = 0
                
                'Salimhere
                
                  
                
                
              End If
             ' .TextMatrix(i, .ColIndex("Price")) = val(.TextMatrix(i, .ColIndex("OldCost")))
            End If
  
  
  End If
  
  OLDItem_ID = Item_ID
' If val(.TextMatrix(i, .ColIndex("NewCost"))) > 1000 Then
'            i = i
'  End If
'here here
       If val(.TextMatrix(i, .ColIndex("NewCost"))) <= 0 Then
                .TextMatrix(i, .ColIndex("NewCost")) = getcostbuylastinvoice(CDbl(val(RsDev("Item_ID").value & "")), CDate(RsDev!Transaction_Date & ""), val(RsDev("UnitID").value & ""), UnitFactor, SecOrder)
                
            If SecOrderCurrent > SecOrder Then
                If UnitFactorCurrent < 1 Then
                    .TextMatrix(i, .ColIndex("NewCost")) = val(.TextMatrix(i, .ColIndex("NewCost"))) * UnitFactorCurrent
                Else
                    .TextMatrix(i, .ColIndex("NewCost")) = val(.TextMatrix(i, .ColIndex("NewCost"))) / UnitFactorCurrent
                End If
            ElseIf SecOrderCurrent < SecOrder Then
                .TextMatrix(i, .ColIndex("NewCost")) = val(.TextMatrix(i, .ColIndex("NewCost"))) * UnitFactor
            Else
                .TextMatrix(i, .ColIndex("NewCost")) = val(.TextMatrix(i, .ColIndex("NewCost"))) * QtyBySmalltUnit
            End If
            
        End If
            
            If val(.TextMatrix(i, .ColIndex("NewCost"))) = 0 Then
'                s = " Select UnitPurPrice from TblItemsUnits where ItemID = " & val(RsDev("Item_ID").value & "") & "  and UnitId = " & val(RsDev("UnitID").value & "")
'                Set rsItemCost = New ADODB.Recordset
'                rsItemCost.Open s, Cn, adOpenStatic, adLockReadOnly
'                If Not rsItemCost.EOF Then
                    .TextMatrix(i, .ColIndex("NewCost")) = mmUnitPurPrice '  val(rsItemCost!UnitPurPrice & "")
            '    End If
                
            End If
   .TextMatrix(i, .ColIndex("NewCost")) = Round(.TextMatrix(i, .ColIndex("NewCost")), 3)
    .TextMatrix(i, .ColIndex("OldCost")) = Round(.TextMatrix(i, .ColIndex("OldCost")), 3)
    If val(.TextMatrix(i, .ColIndex("NewCost"))) < 0 Then
        .TextMatrix(i, .ColIndex("NewCost")) = .TextMatrix(i, .ColIndex("NewCost"))
    End If
   RsDev("OldQty").value = .TextMatrix(i, .ColIndex("OldQty"))
  
' If AllItemHaveOneUnit.value = vbChecked Then
'   RsDev("showprice").value = .TextMatrix(i, .ColIndex("Price"))


' End If
    RsDev("OldCost").value = .TextMatrix(i, .ColIndex("OldCost"))
    RsDev("NewCost").value = .TextMatrix(i, .ColIndex("NewCost"))
    If chkIsCost.value = vbChecked Then
            
            RsDev("showprice").value = val(IIf((.TextMatrix(i, .ColIndex("Price")) = ""), Null, val(.TextMatrix(i, .ColIndex("Price"))))) ' * val(.TextMatrix(i, .ColIndex("QtyBySmalltUnit")))
            RsDev("Price").value = .TextMatrix(i, .ColIndex("Price"))
            
    End If
    Debug.Print i & "---" & RsDev("Price").value
      RsDev("NewQty").value = .TextMatrix(i, .ColIndex("NewQty"))

       RsDev.update
   ' RsDev.Resync
       
                RsDev.MoveNext
            Next i
        End With
    End If
    
    Else
    If SystemOptions.UserInterface = ArabicInterface Then
              MsgBox "·« ÌÊÃœ »Ì«‰«  ·⁄—÷Â«", vbInformation
  Else
            MsgBox "No Data ", vbInformation
  End If
End If
End Sub
Private Sub TxtModFlg_Change() ' %%%%%%%% Set Windows Stutes %%%%%%%%%
    If Me.TxtModFlg.text = "N" Then
    '################### case new recored ########################
    Grid.Enabled = True
    Cmd(0).Enabled = False
    Cmd(1).Enabled = False
    Cmd(2).Enabled = True
    Cmd(3).Enabled = True
    Cmd(4).Enabled = False
    Cmd(5).Enabled = False
    Cmd(6).Enabled = True
    Cmd(7).Enabled = True
    btnFirst.Enabled = False
    btnPrevious.Enabled = False
    btnNext.Enabled = False
    btnLast.Enabled = False
    ShowItemsData.Enabled = True
    txtcode.Enabled = True
    ItemCodeTxt.Enabled = True
    itemNameComp.Enabled = True
    DCboStoreName.Enabled = True
    Me.DCboUserName.BoundText = user_id
    '#############################################################
    ElseIf Me.TxtModFlg.text = "E" Then
    '################### case edit recored #######################
    Grid.Enabled = True
    Cmd(0).Enabled = False
    Cmd(1).Enabled = False
    Cmd(2).Enabled = True
    Cmd(3).Enabled = True
    Cmd(4).Enabled = False
    Cmd(5).Enabled = False
    btnFirst.Enabled = False
    btnPrevious.Enabled = False
    btnNext.Enabled = False
    btnLast.Enabled = False
    ShowItemsData.Enabled = True
    txtcode.Enabled = True
    ItemCodeTxt.Enabled = True
    itemNameComp.Enabled = True
    DCboStoreName.Enabled = True
    Me.DCboUserName.BoundText = user_id
    '#############################################################
    ElseIf Me.TxtModFlg.text = "R" Then
    '################### case read recored #######################
    ' lock all fields show only
    Cmd(0).Enabled = True
    Cmd(1).Enabled = True
    Cmd(2).Enabled = False
    Cmd(3).Enabled = False
    Cmd(4).Enabled = True
    Cmd(5).Enabled = True
    Cmd(6).Enabled = True
    Cmd(7).Enabled = True
    btnFirst.Enabled = True
    btnPrevious.Enabled = True
    btnNext.Enabled = True
    btnLast.Enabled = True
    ShowItemsData.Enabled = False
    txtcode.Enabled = False
    ItemCodeTxt.Enabled = False
    itemNameComp.Enabled = False
    DCboStoreName.Enabled = False
    '#############################################################
    End If
End Sub
Private Sub BtnFirst_Click()
    If Not (rs.EOF Or rs.BOF) Then
        rs.MoveFirst
        txtcode.text = IIf(IsNull(rs("ID").value), 0, rs("ID").value)
    End If
    Retrive
End Sub
Private Sub BtnLast_Click()
    If Not (rs.EOF Or rs.BOF) Then
        rs.MoveLast
        txtcode.text = IIf(IsNull(rs("ID").value), 0, rs("ID").value)
    End If
    Retrive
End Sub

Private Sub BtnNext_Click()
    If Not (rs.EOF Or rs.BOF) Then
        rs.MoveNext
        If rs.EOF Then rs.MoveLast
        txtcode.text = IIf(IsNull(rs("ID").value), 0, rs("ID").value)
    End If
    Retrive
End Sub

Private Sub BtnPrevious_Click()
    If Not (rs.EOF Or rs.BOF) Then
        rs.MovePrevious
        If rs.BOF Then rs.MoveFirst
        txtcode.text = IIf(IsNull(rs("ID").value), 0, rs("ID").value)
    End If
    Retrive
End Sub
Private Sub SaveData()
    '@@@@@@@@ declear Var @@@@@@@@@
    Dim StrSQL As String
    Dim i As Double
    Dim Note_Value1 As Double
    Dim RsHeader As ADODB.Recordset
    Set RsHeader = New ADODB.Recordset
    Dim RsDetails As ADODB.Recordset
    Set RsDetails = New ADODB.Recordset
    Dim rs2 As ADODB.Recordset
    Set rs2 = New ADODB.Recordset
    '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    
    'Handling an exception
    'On Error GoTo ErrTrap
    
    '################################################################## Header Part ##################################################################
    If Me.TxtModFlg.text = "N" Then
        'get the last id and add one
        Dim str As String
        str = new_id("TblEditItemCost", "ID", "", True)
        rs.AddNew
        rs("ID").value = str
        txtcode.text = str
    ElseIf Me.TxtModFlg.text = "E" Then
        StrSQL = "Delete From TblEditItemCostDet Where TblEditItemCostID=" & val(Me.txtcode.text)
        Cn.Execute StrSQL, , adExecuteNoRecords
    End If
    rs("RecDate").value = DtRecord.value
    rs("BranchID").value = Current_branch
    rs("ItemID").value = val(itemNameComp.BoundText)
    rs("StoreID").value = val(DCboStoreName.BoundText)
    rs("UserID").value = user_id
    rs.update
        
    '#################################################################################################################################################

    '############################################################## Det Part (Grid part) #############################################################
    StrSQL = "SELECT * from TblEditItemCostDet Where (1 = -1)"
    RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    'RsDetails("NoteDate").value = IIf((.TextMatrix(i, .ColIndex("NoteDate"))) = "", Null, (.TextMatrix(i, .ColIndex("NoteDate"))))
    With Grid
        For i = .FixedRows To .Rows - 1
            RsDetails.AddNew
            RsDetails("TblEditItemCostID").value = IIf((txtcode.text) = "", 0, (txtcode.text))
            RsDetails("TransactionTypeName").value = IIf((.TextMatrix(i, .ColIndex("TransactionTypeName"))) = "", Null, (.TextMatrix(i, .ColIndex("TransactionTypeName"))))
            RsDetails("StockEffect").value = IIf((.TextMatrix(i, .ColIndex("StockEffect"))) = "", 0, (.TextMatrix(i, .ColIndex("StockEffect"))))
            RsDetails("Transaction_Date").value = IIf((.TextMatrix(i, .ColIndex("Transaction_Date"))) = "", Date, (.TextMatrix(i, .ColIndex("Transaction_Date"))))
            RsDetails("Item_ID").value = IIf((.TextMatrix(i, .ColIndex("Item_ID"))) = "", 0, (.TextMatrix(i, .ColIndex("Item_ID"))))
            RsDetails("StoreID").value = IIf((.TextMatrix(i, .ColIndex("StoreID"))) = "", 0, (.TextMatrix(i, .ColIndex("StoreID"))))
            RsDetails("Quantity").value = IIf((.TextMatrix(i, .ColIndex("Quantity"))) = "", 0, (.TextMatrix(i, .ColIndex("Quantity"))))
            RsDetails("Price").value = IIf((.TextMatrix(i, .ColIndex("Price"))) = "", 0, (.TextMatrix(i, .ColIndex("Price"))))
            RsDetails("OldQty").value = IIf((.TextMatrix(i, .ColIndex("OldQty"))) = "", 0, (.TextMatrix(i, .ColIndex("OldQty"))))
            RsDetails("OldCost").value = IIf((.TextMatrix(i, .ColIndex("OldCost"))) = "", 0, (.TextMatrix(i, .ColIndex("OldCost"))))
            RsDetails("NewQty").value = IIf((.TextMatrix(i, .ColIndex("NewQty"))) = "", 0, (.TextMatrix(i, .ColIndex("NewQty"))))
            RsDetails("NewCost").value = IIf((.TextMatrix(i, .ColIndex("NewCost"))) = "", 0, (.TextMatrix(i, .ColIndex("NewCost"))))
            RsDetails("Fullcode").value = IIf((.TextMatrix(i, .ColIndex("Fullcode"))) = "", " ", (.TextMatrix(i, .ColIndex("Fullcode"))))
            RsDetails("ItemName").value = IIf((.TextMatrix(i, .ColIndex("ItemName"))) = "", " ", (.TextMatrix(i, .ColIndex("ItemName"))))
            RsDetails("ItemNamee").value = IIf((.TextMatrix(i, .ColIndex("ItemNamee"))) = "", " ", (.TextMatrix(i, .ColIndex("ItemNamee"))))
            RsDetails.update
        Next i
    End With
    RsDetails.Close
    Set RsDetails = Nothing
    '#############################################################################################################################################
    If TxtModFlg.text = "N" Then
        If SystemOptions.UserInterface = EnglishInterface Then
            Msg = "Data Saved Successfully" & CHR(13)
        Else
            Msg = " „ Õ›Ÿ «·»Ì«‰« " & CHR(13)
        End If
        
    ElseIf TxtModFlg.text = "E" Then
        If SystemOptions.UserInterface = EnglishInterface Then
            Msg = "Data Edited Successfully" & CHR(13)
        Else
            Msg = " „  ⁄œÌ· «·»Ì«‰« " & CHR(13)
        End If
    End If
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Me.TxtModFlg.text = "R"
    XPTxtCurrent.Caption = rs.RecordCount
    XPTxtCount.Caption = rs.RecordCount
ErrTrap:
'******************************** show Error Message *******************************

End Sub
Public Sub Retrive(Optional Lngid As Long = 0) '%%%%%%%%%% Get the last Recored %%%%%%%%%%
    '@@@@@@@@ declear Var @@@@@@@@@
    Dim sql As String
    Dim i As Double
    Dim StrSQL As String
    
    'Grid Part
    Dim RsDetails As ADODB.Recordset
    Set RsDetails = New ADODB.Recordset
    
    'Header part
    Dim RsHeader As ADODB.Recordset
    Set RsHeader = New ADODB.Recordset
    
    '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    
    'Handling an exception
    On Error GoTo ErrTrap
    
    '########################################################### Header Part ##################################################
    'StrSQL = "select * from TblEndDebtAgingInv where ID = " & val(RecId.Text)
    'RsHeader.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    
    txtcode.text = IIf(IsNull(rs("ID").value), 0, rs("ID").value)
    DtRecord.value = IIf(IsNull(rs("RecDate").value), Date, rs("RecDate").value)
    itemNameComp.BoundText = IIf(IsNull(rs("ItemID").value), Date, rs("ItemID").value)
    DCboStoreName.BoundText = IIf(IsNull(rs("StoreID").value), Date, rs("StoreID").value)
    DCboUserName.BoundText = IIf(IsNull(rs("UserID").value), Date, rs("UserID").value)
    
    '##########################################################################################################################
    
    '################################################## Det Part (Grid Part) ##################################################
    StrSQL = "SELECT * from TblEditItemCostDet where TblEditItemCostID = " & txtcode.text & ""
    RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    
    With Grid
        .Clear flexClearScrollable, flexClearEverything
        .Rows = .FixedRows
        If Not (RsDetails.BOF Or RsDetails.EOF) Then
            RsDetails.MoveFirst
            .Rows = .FixedRows + RsDetails.RecordCount
            For i = .FixedRows To RsDetails.RecordCount
            
            .TextMatrix(i, .ColIndex("Index")) = i
                .TextMatrix(i, .ColIndex("TransactionTypeName")) = IIf(IsNull(RsDetails("TransactionTypeName").value), 0, RsDetails("TransactionTypeName").value)
                .TextMatrix(i, .ColIndex("StockEffect")) = IIf(IsNull(RsDetails("StockEffect").value), 0, RsDetails("StockEffect").value)
                .TextMatrix(i, .ColIndex("Transaction_Date")) = IIf(IsNull(RsDetails("Transaction_Date").value), 0, RsDetails("Transaction_Date").value)
                .TextMatrix(i, .ColIndex("Transaction_ID")) = IIf(IsNull(RsDetails("Transaction_ID").value), 0, RsDetails("Transaction_ID").value)
                .TextMatrix(i, .ColIndex("Item_ID")) = IIf(IsNull(RsDetails("Item_ID").value), 0, RsDetails("Item_ID").value)
                .TextMatrix(i, .ColIndex("StoreID")) = IIf(IsNull(RsDetails("StoreID").value), 0, RsDetails("StoreID").value)
                .TextMatrix(i, .ColIndex("Quantity")) = IIf(IsNull(RsDetails("Quantity").value), 0, RsDetails("Quantity").value)
                .TextMatrix(i, .ColIndex("Price")) = IIf(IsNull(RsDetails("Price").value), 0, RsDetails("Price").value)
                .TextMatrix(i, .ColIndex("OldQty")) = IIf(IsNull(RsDetails("OldQty").value), 0, RsDetails("OldQty").value)
                .TextMatrix(i, .ColIndex("OldCost")) = IIf(IsNull(RsDetails("OldCost").value), 0, RsDetails("OldCost").value)
                .TextMatrix(i, .ColIndex("NewQty")) = IIf(IsNull(RsDetails("NewQty").value), 0, RsDetails("NewQty").value)
                .TextMatrix(i, .ColIndex("NewCost")) = IIf(IsNull(RsDetails("NewCost").value), 0, RsDetails("NewCost").value)
                .TextMatrix(i, .ColIndex("Fullcode")) = IIf(IsNull(RsDetails("Fullcode").value), 0, RsDetails("Fullcode").value)
                .TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(RsDetails("ItemName").value), 0, RsDetails("ItemName").value)
                .TextMatrix(i, .ColIndex("ItemNamee")) = IIf(IsNull(RsDetails("ItemNamee").value), 0, RsDetails("ItemNamee").value)
                RsDetails.MoveNext
            Next i
        End If
    End With
    
        
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    RsDetails.Close
    Set RsDetails = Nothing
    '#################################################################################################################
ErrTrap:
'******************************** show Error Message *******************************
End Sub
Private Sub Cmd_Click(Index As Integer) '%%%%%%%%% Command Bar %%%%%%%%%
    '@@@@@@@@ declear Var @@@@@@@@@
    Dim intDef As Integer
    Dim StrSQL As String
    '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    
    'Handling an exception
    'On Error GoTo ErrTrap

    Select Case Index
        Case 0
        '######################### New Bottom ###########################
            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If
            TxtModFlg.text = "N"
            clear_all Me
             Me.DCboUserName.BoundText = user_id
             Dcbranch.BoundText = branch_id
             DtRecord.value = Date
            txtcode.text = new_id("TblEditItemCost", "ID", "", True)
            Grid.Rows = 1
        '################################################################
        Case 1
        '######################## Edit Bottom ###########################
        'check if user have permission to EDIT recored
            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If
            TxtModFlg.text = "E"
        Case 2
        '######################## Save Bottom ###########################
            'call save function
            SaveData
        '################################################################
        Case 3
        '######################## Undo Bottom ###########################
            'call undo function
            Undo
        '################################################################
        Case 4
        '######################## Delete Bottom ###########################
            ' check if user have permission to DELETE recored
            If DoPremis(Do_Delete, Me.Name, True) = False Then
                Exit Sub
            End If
            
            If SystemOptions.UserInterface = EnglishInterface Then
            Else
            End If
           If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "”Ê› Ì „ Õ–› «·⁄„·Ì… .."
            Msg = Msg & CHR(13) & "›Â· «‰  „ «ﬂœ „‰ «·√” „—«— ..!!"
            Else
            Msg = "Confirm Delete"
            End If
            If MsgBox(Msg, vbQuestion + vbYesNo + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                'call Delete function
                DelRecored
                Me.TxtModFlg.text = "R"
                BtnLast_Click
            End If
        '################################################################
        Case 5
        '######################## Search Bottom #########################
        '################################################################
        Case 7
        '######################## Print Bottom ##########################
            'call print report function
            print_report
        '################################################################
        Case 6
        '######################## Exit Bottom ##########################
            'clear all function and get the last recored
            Unload Me
        '################################################################
    End Select
ErrTrap:
'******************************** show Error Message *******************************
End Sub
Private Sub DelRecored() '%%%%%%%% Delete current recored %%%%%%%%
    '@@@@@@@@ declear Var @@@@@@@@@
    Dim Msg As String
    Dim StrSQL As String
    '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
    'Handling an exception
    On Error GoTo ErrTrap
    If rs.RecordCount > 0 Then
        rs.delete
        StrSQL = "Delete From TblEditItemCost Where Id = " & val(txtcode.text)
        Cn.Execute StrSQL, , adExecuteNoRecords
        StrSQL = "Delete From TblEditItemCostDet Where TblEditItemCostID = " & val(txtcode.text)
        Cn.Execute StrSQL, , adExecuteNoRecords
        rs.MovePrevious
    End If
    
    If rs.RecordCount < 1 Then
        Grid.Rows = 1
        clear_all Me
        XPTxtCurrent.Caption = 0
        XPTxtCount.Caption = 0
    Else
        Grid.Rows = 1
        clear_all Me
        txtcode.text = IIf(IsNull(rs("ID").value), 0, rs("ID").value)
        Retrive
    End If
                    
ErrTrap:
'******************************** show Error Message *******************************

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
  
    MySQL = "SELECT * from TblEditItemCostDet where TblEditItemCostID = " & txtcode.text & ""
 
    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "EditCostReport.rpt"
    Else
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "EditCostReportE.rpt"
    End If

    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        Else
            Msg = "No Data"
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
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName  ' RPTCompany_Name_Eng
        StrReportTitle = ""
    End If
    xReport.ParameterFields(3).AddCurrentValue user_name
    
    xReport.ParameterFields(4).AddCurrentValue IIf(IsNull(txtcode.text), 0, txtcode.text)
    xReport.ParameterFields(5).AddCurrentValue DtRecord.value
    xReport.ParameterFields(6).AddCurrentValue IIf((Dcbranch.text = ""), " ", Dcbranch.text)
    xReport.ParameterFields(7).AddCurrentValue IIf((itemNameComp.text = ""), "No", itemNameComp.text)
    xReport.ParameterFields(8).AddCurrentValue IIf((DCboStoreName.text = ""), "No ", DCboStoreName.text)
    
    
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



