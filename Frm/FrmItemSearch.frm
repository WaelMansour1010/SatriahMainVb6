VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SuiteCtrls.ocx"
Begin VB.Form FrmItemSearch 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«·»ÕÀ ⁄‰ ’‰›..."
   ClientHeight    =   7635
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15000
   Icon            =   "FrmItemSearch.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   7635
   ScaleWidth      =   15000
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin C1SizerLibCtl.C1Tab TabMain 
      Height          =   8265
      Left            =   120
      TabIndex        =   0
      Top             =   90
      Width           =   16200
      _cx             =   28575
      _cy             =   14579
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
      Caption         =   "⁄—÷ ‘Ã—Ï|⁄—÷ ÃœÊ·Ï"
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
         Height          =   7890
         Index           =   1
         Left            =   45
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   45
         Width           =   16110
         _cx             =   28416
         _cy             =   13917
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
         Begin VB.TextBox XPTxtItemCode 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   3930
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   27
            Top             =   2505
            Width           =   1635
         End
         Begin VB.TextBox TxtItemName 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   3930
            RightToLeft     =   -1  'True
            TabIndex        =   26
            Top             =   2880
            Width           =   4665
         End
         Begin VB.ComboBox CboSerial 
            Height          =   315
            Left            =   990
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   3690
            Width           =   1515
         End
         Begin VB.TextBox TxtItemID 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   6960
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   24
            Top             =   2520
            Width           =   1635
         End
         Begin VB.ComboBox CboItemType 
            Height          =   315
            Left            =   5850
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   4050
            Width           =   2745
         End
         Begin VB.ComboBox CboAssbliedItem 
            Height          =   315
            Left            =   990
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   4050
            Width           =   1515
         End
         Begin VB.ComboBox CboAttachedItem 
            Height          =   315
            Left            =   990
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   4440
            Width           =   1515
         End
         Begin VB.ComboBox CboNameSearch 
            Height          =   315
            Left            =   990
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   2880
            Width           =   1515
         End
         Begin VB.ComboBox CboGuar 
            Height          =   315
            Left            =   3930
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   3705
            Width           =   1305
         End
         Begin VB.ComboBox CboArchive 
            Height          =   315
            Left            =   3930
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   4050
            Width           =   1305
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·’‰› «·„—«œ «·»ÕÀ ⁄‰Â ÌÕ ÊÏ ⁄·Ï Â–« «·’‰› ﬂ«Õœ „ﬂÊ‰« Â"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   840
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   17
            Top             =   7470
            Width           =   6495
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·’‰› «·„—«œ «·»ÕÀ ⁄‰Â ÌÕ ÊÏ ⁄·Ï Â–« «·’‰› ﬂ«Õœ „·Õﬁ« Â"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   840
            Left            =   -60
            RightToLeft     =   -1  'True
            TabIndex        =   16
            Top             =   7410
            Width           =   6495
         End
         Begin VB.ComboBox CboItemCodeSearch 
            Height          =   315
            Left            =   990
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   2520
            Width           =   1515
         End
         Begin VB.TextBox TxtbarCodeNO 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   3930
            RightToLeft     =   -1  'True
            TabIndex        =   14
            Top             =   3285
            Width           =   4665
         End
         Begin VB.TextBox TxtPartNo 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   960
            RightToLeft     =   -1  'True
            TabIndex        =   13
            Top             =   3285
            Width           =   1545
         End
         Begin VB.CheckBox Check17 
            Alignment       =   1  'Right Justify
            Caption         =   " ÕœÌœ «·ﬂ·"
            Height          =   360
            Left            =   9840
            RightToLeft     =   -1  'True
            TabIndex        =   12
            Top             =   2490
            Width           =   1305
         End
         Begin VB.TextBox TxtItemPrice 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   7500
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   11
            Top             =   4440
            Width           =   1095
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00E2E9E9&
            BorderStyle     =   0  'None
            Height          =   360
            Left            =   3960
            RightToLeft     =   -1  'True
            TabIndex        =   5
            Top             =   4410
            Width           =   3495
            Begin XtremeSuiteControls.RadioButton RdTotal 
               Height          =   255
               Index           =   0
               Left            =   3000
               TabIndex        =   6
               Top             =   0
               Width           =   495
               _Version        =   786432
               _ExtentX        =   873
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   ">"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton RdTotal 
               Height          =   255
               Index           =   1
               Left            =   2400
               TabIndex        =   7
               Top             =   0
               Width           =   495
               _Version        =   786432
               _ExtentX        =   873
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "<"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton RdTotal 
               Height          =   255
               Index           =   2
               Left            =   1800
               TabIndex        =   8
               Top             =   0
               Width           =   495
               _Version        =   786432
               _ExtentX        =   873
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "="
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton RdTotal 
               Height          =   255
               Index           =   3
               Left            =   1080
               TabIndex        =   9
               Top             =   0
               Width           =   615
               _Version        =   786432
               _ExtentX        =   1085
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   ">="
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton RdTotal 
               Height          =   255
               Index           =   4
               Left            =   120
               TabIndex        =   10
               Top             =   0
               Width           =   855
               _Version        =   786432
               _ExtentX        =   1508
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "<="
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               UseVisualStyle  =   -1  'True
            End
         End
         Begin VSFlex8UCtl.VSFlexGrid FgItems 
            Height          =   7800
            Index           =   0
            Left            =   22425
            TabIndex        =   2
            Top             =   675
            Width           =   15885
            _cx             =   28019
            _cy             =   13758
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
            FormatString    =   $"FrmItemSearch.frx":030A
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
         Begin VSFlex8UCtl.VSFlexGrid Fg 
            Height          =   2370
            Left            =   0
            TabIndex        =   28
            Top             =   60
            Width           =   14775
            _cx             =   26061
            _cy             =   4180
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
            Cols            =   13
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmItemSearch.frx":03CA
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
         Begin ImpulseButton.ISButton Cmd 
            Height          =   360
            Index           =   0
            Left            =   6000
            TabIndex        =   29
            Top             =   5040
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   635
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
            Height          =   360
            Index           =   1
            Left            =   5010
            TabIndex        =   30
            Top             =   5040
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   635
            ButtonStyle     =   1
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
            Cancel          =   -1  'True
            Height          =   360
            Index           =   2
            Left            =   4110
            TabIndex        =   31
            Top             =   5040
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   635
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
         Begin MSDataListLib.DataCombo DCboGroupName 
            Height          =   315
            Left            =   5850
            TabIndex        =   32
            Top             =   3705
            Width           =   2745
            _ExtentX        =   4842
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   360
            Index           =   3
            Left            =   2880
            TabIndex        =   33
            Top             =   5085
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   635
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "«Œ Ì«— „ ⁄œœ"
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
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·„Ã„Ê⁄…"
            Height          =   300
            Index           =   3
            Left            =   8610
            RightToLeft     =   -1  'True
            TabIndex        =   50
            Top             =   3705
            Width           =   1065
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬂÊœ «·’‰›"
            Height          =   330
            Index           =   0
            Left            =   5820
            RightToLeft     =   -1  'True
            TabIndex        =   49
            Top             =   2520
            Width           =   795
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·’‰›"
            Height          =   300
            Index           =   1
            Left            =   8610
            RightToLeft     =   -1  'True
            TabIndex        =   48
            Top             =   2880
            Width           =   1065
         End
         Begin VB.Label LblRes 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            ForeColor       =   &H000000C0&
            Height          =   255
            Left            =   10050
            RightToLeft     =   -1  'True
            TabIndex        =   47
            Top             =   4170
            Width           =   915
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ÿ«„ «·”Ì—Ì«·"
            Height          =   300
            Index           =   2
            Left            =   2880
            RightToLeft     =   -1  'True
            TabIndex        =   46
            Top             =   3705
            Width           =   975
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·’‰›"
            Height          =   300
            Index           =   4
            Left            =   8610
            RightToLeft     =   -1  'True
            TabIndex        =   45
            Top             =   2520
            Width           =   1065
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·’‰›"
            Height          =   300
            Index           =   5
            Left            =   8610
            RightToLeft     =   -1  'True
            TabIndex        =   44
            Top             =   4065
            Width           =   1065
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " Ã„Ì⁄"
            Height          =   270
            Index           =   6
            Left            =   2880
            RightToLeft     =   -1  'True
            TabIndex        =   43
            Top             =   4020
            Width           =   915
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„·Õﬁ"
            Height          =   300
            Index           =   7
            Left            =   3150
            RightToLeft     =   -1  'True
            TabIndex        =   42
            Top             =   4410
            Width           =   705
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„Ã«· «·»ÕÀ"
            Height          =   330
            Index           =   8
            Left            =   2550
            RightToLeft     =   -1  'True
            TabIndex        =   41
            Top             =   2880
            Width           =   975
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·÷„«‰"
            Height          =   255
            Index           =   9
            Left            =   4920
            RightToLeft     =   -1  'True
            TabIndex        =   40
            Top             =   3705
            Width           =   885
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·√—‘Ì›"
            Height          =   285
            Index           =   10
            Left            =   5190
            RightToLeft     =   -1  'True
            TabIndex        =   39
            Top             =   4065
            Width           =   705
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„Ã«· «·»ÕÀ"
            Height          =   330
            Index           =   11
            Left            =   2550
            RightToLeft     =   -1  'True
            TabIndex        =   38
            Top             =   2520
            Width           =   975
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·»«—ﬂÊœ"
            Height          =   300
            Index           =   12
            Left            =   8610
            RightToLeft     =   -1  'True
            TabIndex        =   37
            Top             =   3285
            Width           =   1065
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackColor       =   &H00E2E9E9&
            Height          =   180
            Index           =   13
            Left            =   5430
            RightToLeft     =   -1  'True
            TabIndex        =   36
            Top             =   3285
            Width           =   45
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·ﬁÿ⁄Â/«·„ÊœÌ·"
            Height          =   345
            Index           =   0
            Left            =   2520
            RightToLeft     =   -1  'True
            TabIndex        =   35
            Top             =   3285
            Width           =   1335
         End
         Begin VB.Label XPLbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "”⁄— «·’‰›"
            Height          =   300
            Index           =   7
            Left            =   8610
            RightToLeft     =   -1  'True
            TabIndex        =   34
            Top             =   4410
            Width           =   1065
         End
      End
      Begin C1SizerLibCtl.C1Elastic ELe 
         Height          =   7890
         Index           =   0
         Left            =   16845
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   45
         Width           =   16110
         _cx             =   28416
         _cy             =   13917
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
         Begin VSFlex8UCtl.VSFlexGrid FgItems 
            Height          =   7800
            Index           =   1
            Left            =   22425
            TabIndex        =   4
            Top             =   675
            Width           =   15885
            _cx             =   28019
            _cy             =   13758
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
            FormatString    =   $"FrmItemSearch.frx":05C2
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
         Begin C1SizerLibCtl.C1Elastic C1Elastic3 
            Height          =   6720
            Index           =   1
            Left            =   0
            TabIndex        =   51
            TabStop         =   0   'False
            Top             =   0
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
               TabIndex        =   52
               Top             =   3840
               Width           =   14475
               Begin VB.TextBox txtID 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   10860
                  RightToLeft     =   -1  'True
                  TabIndex        =   60
                  Top             =   480
                  Width           =   2175
               End
               Begin VB.CheckBox Check7 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«”„ «·„Ê—œ »«·ﬂ«„· ›ﬁÿ"
                  Height          =   375
                  Left            =   5760
                  RightToLeft     =   -1  'True
                  TabIndex        =   59
                  Top             =   2760
                  Value           =   1  'Checked
                  Visible         =   0   'False
                  Width           =   2385
               End
               Begin VB.TextBox txtCustomerCode 
                  Alignment       =   2  'Center
                  Height          =   288
                  Left            =   11880
                  TabIndex        =   58
                  Top             =   1320
                  Width           =   1275
               End
               Begin VB.Frame Fra 
                  Caption         =   " «—ÌŒ"
                  Height          =   1095
                  Index           =   14
                  Left            =   420
                  RightToLeft     =   -1  'True
                  TabIndex        =   53
                  Top             =   570
                  Width           =   6012
                  Begin MSComCtl2.DTPicker txtFromDate 
                     Height          =   330
                     Left            =   3240
                     TabIndex        =   54
                     Top             =   390
                     Width           =   1590
                     _ExtentX        =   2805
                     _ExtentY        =   582
                     _Version        =   393216
                     CheckBox        =   -1  'True
                     Format          =   240910339
                     CurrentDate     =   38887
                  End
                  Begin MSComCtl2.DTPicker txtToDate 
                     Height          =   336
                     Left            =   216
                     TabIndex        =   55
                     Top             =   396
                     Width           =   1596
                     _ExtentX        =   2805
                     _ExtentY        =   582
                     _Version        =   393216
                     CheckBox        =   -1  'True
                     Format          =   240910339
                     CurrentDate     =   38887
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     Caption         =   "„‰"
                     Height          =   195
                     Index           =   27
                     Left            =   5100
                     RightToLeft     =   -1  'True
                     TabIndex        =   57
                     Top             =   420
                     Width           =   420
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     Caption         =   "≈·Ï"
                     Height          =   192
                     Index           =   29
                     Left            =   2172
                     RightToLeft     =   -1  'True
                     TabIndex        =   56
                     Top             =   420
                     Width           =   372
                  End
               End
               Begin MSDataListLib.DataCombo DataBranch 
                  Height          =   315
                  Left            =   7560
                  TabIndex        =   61
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
                  TabIndex        =   62
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
                  Caption         =   "—ﬁ„ «·Õ—ﬂ…"
                  ForeColor       =   &H00000000&
                  Height          =   255
                  Index           =   6
                  Left            =   13140
                  TabIndex        =   65
                  Top             =   480
                  Width           =   975
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  Caption         =   "«·›—⁄"
                  Height          =   285
                  Index           =   30
                  Left            =   13290
                  TabIndex        =   64
                  Top             =   960
                  Width           =   645
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
                  TabIndex        =   63
                  Top             =   1320
                  Width           =   675
               End
            End
            Begin VSFlex8UCtl.VSFlexGrid GrdF 
               Height          =   3585
               Left            =   120
               TabIndex        =   66
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
               Cols            =   7
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmItemSearch.frx":0682
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
               TabIndex        =   67
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
               TabIndex        =   68
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
               TabIndex        =   69
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
               TabIndex        =   70
               Top             =   3240
               Visible         =   0   'False
               Width           =   612
            End
         End
      End
   End
End
Attribute VB_Name = "FrmItemSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim cSearchDcbo As clsDCboSearch
Private m_DcboItems As DataCombo
Private m_RetrunType As Integer
Dim ItemsIDes As String
Public Indx As Long
Public mRow As Long
Private Sub Cmd_Click(index As Integer)

    On Error GoTo ErrTrap

    Select Case index
        Case 0
            If rs.State = adStateOpen Then
                rs.Close
            End If
            rs.Open Build_Sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
            If SystemOptions.UserInterface = ArabicInterface Then
                LblRes.Caption = "‰ ÌÃ… «·»ÕÀ = " & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
                LblRes.Caption = "Search Result=" & rs.RecordCount
            End If
            If rs.RecordCount < 1 Then
                FG.Clear flexClearScrollable, flexClearEverything
                FG.rows = 2
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Else
                    Msg = "NO Search Results Found...!!!"
                    MsgBox Msg, vbOKOnly + vbExclamation, App.Title
                End If
                Exit Sub
            End If
            Retrive
            FG.SetFocus
        Case 1
            clear_all Me
            FG.Clear flexClearScrollable, flexClearEverything
        Case 2
            Unload Me
        Case 3
          If Me.RetrunType = 5 Then
            FillItemsIDes
            frmsalebill.TxtItemsIDes.text = ItemsIDes
            frmsalebill.Retrive_Items_data
        ElseIf Me.RetrunType = 21915 Then
            FillItemsIDes
          'sa  FrmCustomerContract.TxtItemsIDes.Text = ItemsIDes
         'sa    FrmCustomerContract.Retrive_Items_data
         ElseIf Me.RetrunType = 3 Then
            FillItemsIDes
            FrmBillBuy.TxtItemsIDes.text = ItemsIDes
            FrmBillBuy.Retrive_Items_data1
            
               ElseIf Me.RetrunType = 909 Then
            FillItemsIDes
            FrmReturnpurchases.TxtItemsIDes.text = ItemsIDes
            FrmReturnpurchases.Retrive_Items_data1
            
            
              ElseIf Me.RetrunType = 9 Then
            FillItemsIDes
            FrmReturnSalling.TxtItemsIDes.text = ItemsIDes
            FrmReturnSalling.Retrive_Items_data1
            
          ElseIf Me.RetrunType = 6 Then
            FillItemsIDes
            FrmOut.TxtItemsIDes.text = ItemsIDes
            FrmOut.Retrive_Items_data1
          ElseIf Me.RetrunType = 222 Then
            FillItemsIDes
            FrmPO11.TxtItemsIDes.text = ItemsIDes
            FrmPO11.Retrive_Items_data1
          ElseIf Me.RetrunType = 12 Then
            FillItemsIDes
            FrmNewGard.TxtItemsIDes.text = ItemsIDes
            FrmNewGard.Retrive_Items_data1
          ElseIf Me.RetrunType = 9878 Then
            FillItemsIDes
            FrmStudentCalling.FG4.TextMatrix(mRow, FrmStudentCalling.FG4.ColIndex("ItemName")) = ItemsIDes

          ElseIf Me.RetrunType = 9888 Then
                FillItemsIDes
                FrmSallingPlan.TxtItemsIDes = ItemsIDes
                FrmSallingPlan.Retrive_Items_data1
            
       
            
          End If
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

Private Sub FG_BeforeEdit(ByVal row As Long, ByVal Col As Long, Cancel As Boolean)
With FG
If .ColKey(Col) <> "Send" Then
Cancel = True
End If
End With
End Sub

Private Sub fg_Click()

    On Error GoTo ErrTrap

    If Not FG.TextMatrix(FG.row, 2) = "" Then
        If Me.RetrunType = 0 Then
            FrmItems.Retrive val(FG.TextMatrix(FG.row, 2))
        ElseIf Me.RetrunType = 1 Then
            If Not Me.DcboItems Is Nothing Then
                Me.DcboItems.BoundText = val(FG.TextMatrix(FG.row, 2))
            End If
        ElseIf Me.RetrunType = 2 Then
            FrmShowPrice.DCboItemsCode.BoundText = val(FG.TextMatrix(FG.row, 2))
        ElseIf Me.RetrunType = 3 Then
            FrmBillBuy.DCboItemsCode.BoundText = val(FG.TextMatrix(FG.row, 2))
        ElseIf Me.RetrunType = 2020 Then
            FrmDefinCompItem.DcboItemID1.BoundText = val(FG.TextMatrix(FG.row, 2))
        ElseIf Me.RetrunType = 2026 Then
            FrmDefinCompItem.DcboItemID3.BoundText = val(FG.TextMatrix(FG.row, 2))
            
        ElseIf Me.RetrunType = 4 Then
            FrmInpout.DCboItemsCode.BoundText = val(FG.TextMatrix(FG.row, 2))
        ElseIf Me.RetrunType = 5 Then
            frmsalebill.DCboItemsCode.BoundText = val(FG.TextMatrix(FG.row, 2))
                    ElseIf Me.RetrunType = 50 Then
            frmsalebill4.DCboItemsCode.BoundText = val(FG.TextMatrix(FG.row, 2))
            
        ElseIf Me.RetrunType = 7715 Then
            frmsalebill2.DCboItemsCode.BoundText = val(FG.TextMatrix(FG.row, 2))
    frmsalebill2.txtItemCodeSearch2.text = FG.TextMatrix(FG.row, 3)
        ElseIf Me.RetrunType = 6 Then
            FrmOut.DCboItemsCode.BoundText = val(FG.TextMatrix(FG.row, 2))
            
       ElseIf Me.RetrunType = 61 Then
            FrmOut1.DCboItemsCode.BoundText = val(FG.TextMatrix(FG.row, 2))
                 
        ElseIf Me.RetrunType = 222 Then
            FrmPO11.DCboItemsCode.BoundText = val(FG.TextMatrix(FG.row, 2))
            ElseIf Me.RetrunType = 30719 Then
            FrmPO10.DCboItemsCode.BoundText = val(FG.TextMatrix(FG.row, 2))
            
        ElseIf Me.RetrunType = 7 Then
            FrmDestruction.DCboItemsCode.BoundText = val(FG.TextMatrix(FG.row, 2))
            
       ElseIf Me.RetrunType = 70 Then
            FrmDestructionRet.DCboItemsCode.BoundText = val(FG.TextMatrix(FG.row, 2))
            
            
        ElseIf Me.RetrunType = 8 Then
            FrmOpeningBalance.DCboItemsCode.BoundText = val(FG.TextMatrix(FG.row, 2))
        ElseIf Me.RetrunType = 9 Then
            FrmReturnSalling.DCboItemsCode.BoundText = val(FG.TextMatrix(FG.row, 2))
        ElseIf Me.RetrunType = 909 Then
            FrmReturnpurchases.DCboItemsCode.BoundText = val(FG.TextMatrix(FG.row, 2))
        ElseIf Me.RetrunType = 10 Then
            FrmMoving.DCboItemsCode.BoundText = val(FG.TextMatrix(FG.row, 2))
        
        ElseIf Me.RetrunType = 11 Then
            FrmSallingPlan.dcitems.BoundText = val(FG.TextMatrix(FG.row, 2))
        ElseIf Me.RetrunType = 9888 Then
            FrmSallingPlan.dcitems.BoundText = val(FG.TextMatrix(FG.row, 2))
        ElseIf Me.RetrunType = 12 Then
            FrmNewGard.DCboItemsCode = val(FG.TextMatrix(FG.row, 2))
            
        ElseIf Me.RetrunType = 13 Then
            FrmOutProductionOrder.DCboItemsCode = val(FG.TextMatrix(FG.row, 2))
        ElseIf Me.RetrunType = 14 Then
            FrmInpoutWorkOrder.DCboItemsCode = val(FG.TextMatrix(FG.row, 2))
        ElseIf Me.RetrunType = 15 Then
            FrmProductionOrder1.DCboItemsCode = val(FG.TextMatrix(FG.row, 2))
        ElseIf Me.RetrunType = 16 Then
             FrmProductionOrder1.DCboItemsCode = val(FG.TextMatrix(FG.row, 2))
        ElseIf Me.RetrunType = 18 Then
            Order_no_search2.DcboItems.BoundText = val(FG.TextMatrix(FG.row, 2))
        ElseIf Me.RetrunType = 17 Then
        
        ElseIf Me.RetrunType = 666 Then
            FrmBeforeInventoryK.DcbItem.BoundText = val(FG.TextMatrix(FG.row, 2))
        ElseIf Me.RetrunType = 19 Then
            FrmPO7.DCboItemsCode = val(FG.TextMatrix(FG.row, 2))
        ElseIf Me.RetrunType = 20 Then
            FrmPO4.DCboItemsCode = val(FG.TextMatrix(FG.row, 2))
        ElseIf Me.RetrunType = 21 Then
            FrmPO5.DCboItemsCode = val(FG.TextMatrix(FG.row, 2))
        ElseIf Me.RetrunType = 22 Then
            FrmPO8.DCboItemsCode = val(FG.TextMatrix(FG.row, 2))
        ElseIf Me.RetrunType = 23 Then
            FrmPO.DCboItemsCode = val(FG.TextMatrix(FG.row, 2))
        ElseIf Me.RetrunType = 24 Then
            FrmPO1.DCboItemsCode = val(FG.TextMatrix(FG.row, 2))
        ElseIf Me.RetrunType = 25 Then
            FrmPO2.DCboItemsCode = val(FG.TextMatrix(FG.row, 2))
        ElseIf Me.RetrunType = 26 Then
            FrmPO3.DCboItemsCode = val(FG.TextMatrix(FG.row, 2))
        ElseIf Me.RetrunType = 27 Then
            FrmDistriExpensItems.DcItem1.BoundText = val(FG.TextMatrix(FG.row, 2))
        ElseIf Me.RetrunType = 28 Then
            FrmDestriEpensItemSearch.DCItem.BoundText = val(FG.TextMatrix(FG.row, 2))
        ElseIf Me.RetrunType = 29 Then
            FrmProcessDef.dcitems.BoundText = val(FG.TextMatrix(FG.row, 2))
        ElseIf Me.RetrunType = 30 Then
            FrmManAddNew.DCboItemsCode1.BoundText = val(FG.TextMatrix(FG.row, 2))
        ElseIf Me.RetrunType = 31 Then
            FrmItems.DcboItems.BoundText = val(FG.TextMatrix(FG.row, 2))
                 ElseIf Me.RetrunType = 310 Then
            FrmCarAuthontication.DcboItems.BoundText = val(FG.TextMatrix(FG.row, 2))
                             
            FrmCarAuthontication.cmbItems.BoundText = val(FG.TextMatrix(FG.row, 2))

         ElseIf Me.RetrunType = 9878 Then
            FrmStudentCalling.FG4.TextMatrix(mRow, FrmStudentCalling.FG4.ColIndex("ItemId")) = val(FG.TextMatrix(FG.row, 2))
            FrmStudentCalling.FG4.TextMatrix(mRow, FrmStudentCalling.FG4.ColIndex("ItemName")) = Trim(FG.TextMatrix(FG.row, FG.ColIndex("KindNme")))
            
        ElseIf Me.RetrunType = 200 Then
            FrmPO6.DCboItemsCode.BoundText = val(FG.TextMatrix(FG.row, 2))
        ElseIf Me.RetrunType = 11815 Then
            FrmProductionOrder.DCboItemsCode.BoundText = val(FG.TextMatrix(FG.row, 2))
        ElseIf Me.RetrunType = 20915 Then
            FrmVendorContract.dcitems.BoundText = val(FG.TextMatrix(FG.row, 2))
            FrmVendorContract.TxtItemCode.text = (FG.TextMatrix(FG.row, 3))
        ElseIf Me.RetrunType = 21915 Then
            FrmCustomerContract.dcitems.BoundText = val(FG.TextMatrix(FG.row, 1))
            FrmCustomerContract.TxtItemCode.text = (FG.TextMatrix(FG.row, 3))
            FrmCustomerContract.dcitems.BoundText = GetItemID(Trim$((FG.TextMatrix(FG.row, 3))))
        ElseIf Me.RetrunType = 1302 Then
            FrmDefinCompItem.DcboItemID2.BoundText = val(FG.TextMatrix(FG.row, 2))
            FrmDefinCompItem.TxtAttachedItemCode2.text = (FG.TextMatrix(FG.row, 3))
        ElseIf Me.RetrunType = 100 Then
            FrmTravelTransactions.DcboItems.BoundText = val(FG.TextMatrix(FG.row, 2))
            FrmTravelTransactions.TxtItemCode.text = (FG.TextMatrix(FG.row, 3))
        ElseIf Me.RetrunType = 101 Then
            FrmPaymenTransTrip.DcboItems.BoundText = val(FG.TextMatrix(FG.row, 2))
            FrmPaymenTransTrip.TxtItemCode.text = (FG.TextMatrix(FG.row, 3))
            
        ElseIf Me.RetrunType = 2711 Then
            Frmovers.DcbItemDit.BoundText = val(FG.TextMatrix(FG.row, 2))
            Frmovers.TxtCode.text = (FG.TextMatrix(FG.row, 3))
            'Order_no_search2.DcboItems.BoundText = val(Fg.TextMatrix(Fg.Row, 1))
            Select Case Frmovers.Item
                Case 1
                    Frmovers.DcbItem.BoundText = val(FG.TextMatrix(FG.row, 2))
                Case 2
                    Frmovers.DcbItemDit.BoundText = val(FG.TextMatrix(FG.row, 2))
                Case 3
                    Frmovers.DcbItemDDis.BoundText = val(FG.TextMatrix(FG.row, 2))
                Case 4
                    Frmovers.DcbItemDis.BoundText = val(FG.TextMatrix(FG.row, 2))
                Case 5
                    Frmovers.DcItem1.BoundText = val(FG.TextMatrix(FG.row, 2))
            End Select
        End If
    End If
    Exit Sub
ErrTrap:
End Sub
Private Sub Retrive()

    Dim Num As Integer
    
    On Error GoTo ErrTrap
    
    FG.Clear flexClearScrollable, flexClearEverything

    If Not (rs.EOF Or rs.BOF) Then
        FG.rows = rs.RecordCount + 1
        For Num = 1 To rs.RecordCount
            With FG
                .TextMatrix(Num, .ColIndex("NumIndex")) = Num
                .TextMatrix(Num, .ColIndex("ItemNum")) = IIf(IsNull(rs("ItemID").value), "", val(rs("ItemID").value))
                .TextMatrix(Num, .ColIndex("KindCode")) = IIf(IsNull(rs("ItemCode").value), "", Trim(rs("ItemCode").value))
                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(Num, .ColIndex("KindNme")) = IIf(IsNull(rs("ItemName").value), "", Trim(rs("ItemName").value))
                Else
                    .TextMatrix(Num, .ColIndex("KindNme")) = IIf(IsNull(rs("ItemNamee").value), "", Trim(rs("ItemNamee").value))
                End If
                .TextMatrix(Num, .ColIndex("barCodeNO")) = IIf(IsNull(rs("barCodeNO").value), "", Trim(rs("barCodeNO").value))
                .TextMatrix(Num, .ColIndex("PartNo")) = IIf(IsNull(rs("PartNo").value), "", (rs("PartNo").value))
                If rs("ItemType").value = 0 Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        .TextMatrix(Num, .ColIndex("ItemType")) = "”·⁄…"
                    Else
                        .TextMatrix(Num, .ColIndex("ItemType")) = "Goods"
                    End If
                Else
                    If SystemOptions.UserInterface = ArabicInterface Then
                        .TextMatrix(Num, .ColIndex("ItemType")) = "Œœ„…"
                    Else
                        .TextMatrix(Num, .ColIndex("ItemType")) = "Service"
                    End If
                End If
                If rs("HaveSerial").value Eqv True Then
                    .TextMatrix(Num, .ColIndex("HaveSerial")) = 1
                Else
                    .TextMatrix(Num, .ColIndex("HaveSerial")) = 0
                End If

                If rs("HaveGuarantee").value Eqv True Then
                    .TextMatrix(Num, .ColIndex("HaveGuarantee")) = 1
                Else
                    .TextMatrix(Num, .ColIndex("HaveGuarantee")) = 0
                End If

                If rs("IsArchive").value Eqv True Then
                    .TextMatrix(Num, .ColIndex("IsArchive")) = 1
                Else
                    .TextMatrix(Num, .ColIndex("IsArchive")) = 0
                End If
            
                If rs("AssbliedItem").value Eqv True Then
                    .TextMatrix(Num, .ColIndex("AssbliedItem")) = 1
                Else
                    .TextMatrix(Num, .ColIndex("AssbliedItem")) = 0
                End If
                        
                If rs("RelatedItem").value Eqv True Then
                    .TextMatrix(Num, .ColIndex("RelatedItem")) = 1
                Else
                    .TextMatrix(Num, .ColIndex("RelatedItem")) = 0
                End If
            End With
            rs.MoveNext
        Next Num
        FG.AutoSize 0, FG.Cols - 1, False
    End If
    Exit Sub
ErrTrap:
End Sub
Private Sub Fg_DblClick()
    fg_Click
    Unload Me
End Sub

Private Sub Form_Activate()
Check17.Visible = False
Cmd(3).Visible = False
FG.ColHidden(FG.ColIndex("Send")) = True
If Me.RetrunType = 9 Or Me.RetrunType = 909 Or RetrunType = 5 Or Me.RetrunType = 3 Or Me.RetrunType = 6 Or Me.RetrunType = 222 Or Me.RetrunType = 12 Or Me.RetrunType = 21915 Or Me.RetrunType = 9888 Then
Check17.Visible = True
Cmd(3).Visible = True
FG.ColHidden(FG.ColIndex("Send")) = False
End If
End Sub

Private Sub Form_Load()

    On Error GoTo ErrTrap
    
    Dim StrSQL As String
    Dim BG As New ClsBackGroundPic
    Dim Dcombos As ClsDataCombos

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Clear").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture

    Set Dcombos = New ClsDataCombos
    Dcombos.GetItemSGroups Me.DCboGroupName
    Set cSearchDcbo = New clsDCboSearch
    'cSearchDcbo.AllowWriting = False
    Set cSearchDcbo.Client = Me.DCboGroupName
    TabMain.TabVisible(0) = True
    TabMain.TabVisible(1) = False

    If Indx = 1 Then
        TabMain.TabVisible(0) = False
        TabMain.TabVisible(1) = True
        Me.Caption = "»ÕÀ «·ÕÃÊ“« "
        txtFromDate.value = Date
        txtFromDate.value = Date
        txtToDate.value = Date
    End If
    If SystemOptions.UserInterface = ArabicInterface Then
        With Me.CboItemCodeSearch
            .Clear
            .AddItem "»ÕÀ „ÿ«»ﬁ"
            .AddItem "»ÕÀ „‰ «·»œ«Ì…"
            .AddItem "»ÕÀ „‰ «·‰Â«Ì…"
            .AddItem "»ÕÀ ›Ï «Ï „ﬂ«‰"
        End With
        With Me.CboSerial
            .Clear
            .AddItem "«·ﬂ·"
            .ItemData(0) = 0
            .AddItem "·Â ”Ì—Ì«·"
            .ItemData(1) = 1
            .AddItem "·Ì” ·Â ”Ì—Ì«·"
            .ItemData(2) = 2
        End With
        With Me.CboNameSearch
            .Clear
            .AddItem "„‰ «Ê· «·√”„"
            .AddItem "›Ï «Ï Ã“¡ „‰ «·√”„"
        End With
        With Me.CboItemType
            .Clear
            .AddItem "”·⁄…"
            .AddItem "Œœ„…"
            .AddItem "«·ﬂ·"
        End With

        With Me.CboGuar
            .Clear
            .AddItem "·Â ÷„«‰"
            .AddItem "·Ì” ·Â ÷„«‰"
            .AddItem "«·ﬂ·"
        End With

        With Me.CboArchive
            .Clear
            .AddItem "›Ï «·√—‘Ì›"
            .AddItem "·Ì” ›Ï «·√—‘Ì›"
            .AddItem "«·ﬂ·"
        End With

        With Me.CboAssbliedItem
            .Clear
            .AddItem "’‰› „Ã„⁄"
            .AddItem "’‰› ⁄«œÏ"
            .AddItem "«·ﬂ·"
        End With

        With Me.CboAttachedItem
            .Clear
            .AddItem "·Â «’‰«› „·Õﬁ…"
            .AddItem "·Ì” ·Â «’‰«› „·Õﬁ…"
            .AddItem "«·ﬂ·"
        End With

    ElseIf SystemOptions.UserInterface = EnglishInterface Then

        With Me.CboItemCodeSearch
            .Clear
            .AddItem "Typical Search"
            .AddItem "From The Start"
            .AddItem "From The End"
            .AddItem "Any Where"
        End With

        With Me.CboSerial
            .Clear
            .AddItem "All"
            .ItemData(0) = 0
            .AddItem "Has Serial"
            .ItemData(1) = 1
            .AddItem "NO Serial"
            .ItemData(2) = 2
        End With

        With Me.CboNameSearch
            .Clear
            .AddItem "Start Name"
            .AddItem "Any Part of Name"
        End With

        With Me.CboItemType
            .Clear
            .AddItem "Goods"
            .AddItem "Services"
            .AddItem "All"
        End With

        With Me.CboGuar
            .Clear
            .AddItem "YES"
            .AddItem "NO"
            .AddItem "ALL"
        End With

        With Me.CboArchive
            .Clear
            .AddItem "YES"
            .AddItem "NO"
            .AddItem "ALL"
        End With

        With Me.CboAssbliedItem
            .Clear
            .AddItem "YES"
            .AddItem "NO"
            .AddItem "ALL"
        End With

        With Me.CboAttachedItem
            .Clear
            .AddItem "YES"
            .AddItem "NO"
            .AddItem "ALL"
        End With

    End If

    CenterForm Me
 
    FormPostion Me, GetPostion
    FG.WallPaper = BG.SearchWallpaper
    Set rs = New ADODB.Recordset

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
    Set m_DcboItems = Nothing
    Exit Sub
ErrTrap:
End Sub
Private Function Build_Sql()

    Dim StrSQL As String
    Dim Begin As Boolean
    Dim StrWhere As String
    Dim BolHaveSerial As Boolean
    Dim IntHaveSerial As Integer

    On Error GoTo ErrTrap

    StrSQL = "Select * From TblItems "
    StrSQL = StrSQL + " Where ItemID <> 0 "


  If SystemOptions.WorkWithLINKEDiActivity = True Then
    StrSQL = StrSQL & "  and dbo.TblItems.GroupID in(   "
     StrSQL = StrSQL & " select GroupID from fullgroups ()  )"
 End If



    If val(Me.TxtItemID.text) <> 0 Then
        StrSQL = StrSQL + " AND ItemID =" & val(Me.TxtItemID.text)
    End If

    If XPTxtItemCode.text <> "" Then
        If Me.CboItemCodeSearch.ListIndex = 0 Then
            StrWhere = StrWhere + " and ItemCode ='" & Trim(XPTxtItemCode.text) & "'"
        ElseIf Me.CboItemCodeSearch.ListIndex = 1 Then
            StrWhere = StrWhere + " and ItemCode like '" & Trim(XPTxtItemCode.text) & "%'"
        ElseIf Me.CboItemCodeSearch.ListIndex = 2 Then
            StrWhere = StrWhere + " and ItemCode like '%" & Trim(XPTxtItemCode.text) & "'"
        ElseIf Me.CboItemCodeSearch.ListIndex = 3 Then
            StrWhere = StrWhere + " and ItemCode like '%" & Trim(XPTxtItemCode.text) & "%'"
        ElseIf Me.CboItemCodeSearch.ListIndex = -1 Then
            StrWhere = StrWhere + " and ItemCode like '%" & Trim(XPTxtItemCode.text) & "%'"
        End If
    End If
''///////////

    If val(TxtItemPrice.text) <> 0 Then
        If RdTotal(0).value = True Then
            StrWhere = StrWhere + " and ItemID in(select ItemID from TblItemsUnits where TblItemsUnits.ItemID=TblItems.ItemID and UnitSalesPrice > " & val(TxtItemPrice.text) & ")"
        ElseIf RdTotal(1).value = True Then
            StrWhere = StrWhere + " and ItemID in(select ItemID from TblItemsUnits where TblItemsUnits.ItemID=TblItems.ItemID and UnitSalesPrice < " & val(TxtItemPrice.text) & " and UnitSalesPrice > 0 ) "
        ElseIf RdTotal(3).value = True Then
            StrWhere = StrWhere + " and ItemID in(select ItemID from TblItemsUnits where TblItemsUnits.ItemID=TblItems.ItemID and UnitSalesPrice >= " & val(TxtItemPrice.text) & ")"
        ElseIf RdTotal(4).value = True Then
            StrWhere = StrWhere + " and ItemID in(select ItemID from TblItemsUnits where TblItemsUnits.ItemID=TblItems.ItemID and UnitSalesPrice <= " & val(TxtItemPrice.text) & " and UnitSalesPrice > 0 )"
        Else
            StrWhere = StrWhere + " and ItemID in(select ItemID from TblItemsUnits where TblItemsUnits.ItemID=TblItems.ItemID and UnitSalesPrice=" & val(TxtItemPrice.text) & ")"
        End If
    End If
''///////
    If TxtbarCodeNO.text <> "" Then
        If Me.CboItemCodeSearch.ListIndex = 0 Then
            StrWhere = StrWhere + " and barCodeNO ='" & Trim(TxtbarCodeNO.text) & "'"
        ElseIf Me.CboItemCodeSearch.ListIndex = 1 Then
            StrWhere = StrWhere + " and barCodeNO like '" & Trim(TxtbarCodeNO.text) & "%'"
        ElseIf Me.CboItemCodeSearch.ListIndex = 2 Then
            StrWhere = StrWhere + " and barCodeNO like '%" & Trim(TxtbarCodeNO.text) & "'"
        ElseIf Me.CboItemCodeSearch.ListIndex = 3 Then
            StrWhere = StrWhere + " and barCodeNO like '%" & Trim(TxtbarCodeNO.text) & "%'"
        ElseIf Me.CboItemCodeSearch.ListIndex = -1 Then
            StrWhere = StrWhere + " and barCodeNO like '%" & Trim(TxtbarCodeNO.text) & "%'"
        End If
    End If
    If TxtPartNo.text <> "" Then
        StrWhere = StrWhere + " and PartNo like '%" & Trim(TxtPartNo.text) & "%'"
    End If

    If Me.CboSerial.ListIndex > 0 Then
        If Me.CboSerial.ItemData(CboSerial.ListIndex) = 1 Then
            BolHaveSerial = True
        ElseIf Me.CboSerial.ItemData(CboSerial.ListIndex) = 2 Then
            BolHaveSerial = False
        End If

        If SystemOptions.SysDataBaseType = AccessDataBase Then
            StrWhere = StrWhere + " and HaveSerial =" & BolHaveSerial & ""
        ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
            IntHaveSerial = IIf(BolHaveSerial = True, 1, 0)
            StrWhere = StrWhere + " and HaveSerial =" & IntHaveSerial & ""
        End If
    End If

'    If SystemOptions.UserInterface = ArabicInterface Then
        If Trim(Me.txtItemName.text) <> "" Then
            If Me.CboNameSearch.ListIndex = 0 Then
                StrWhere = StrWhere + " and ItemName Like '" & Trim(Me.txtItemName.text) & "%'"
            ElseIf (Me.CboNameSearch.ListIndex = 1 Or Me.CboNameSearch.ListIndex = -1) Then
                StrWhere = StrWhere + " and ItemName like '%" & Trim(Me.txtItemName.text) & "%'"
            End If
        End If
'    Else
        If Trim(Me.txtItemName.text) <> "" Then
            If Me.CboNameSearch.ListIndex = 0 Then
                StrWhere = StrWhere + " or ItemNamee Like '" & Trim(Me.txtItemName.text) & "%'"
            ElseIf (Me.CboNameSearch.ListIndex = 1 Or Me.CboNameSearch.ListIndex = -1) Then
                StrWhere = StrWhere + " or ItemNamee like '%" & Trim(Me.txtItemName.text) & "%'"
            End If
        End If
'    End If

    If Me.DCboGroupName.BoundText <> "" And Me.DCboGroupName.text <> "" Then
 '       StrWhere = StrWhere + " and GroupID =" & Me.DCboGroupName.BoundText & ""
        CreateRecusiveGroup val(DCboGroupName.BoundText), CInt(user_id)
       StrWhere = StrWhere & " and  GroupID in ( " & " select GroupID from fullgroups" & user_id & " () " & ")"
       
        
        
    End If

    If Me.CboItemType.ListIndex <> -1 Then
        If Me.CboItemType.ListIndex = 0 Then
            StrSQL = StrSQL + " AND ItemType =0"
        ElseIf Me.CboItemType.ListIndex = 1 Then
            StrSQL = StrSQL + " AND ItemType =1"
        End If
    End If

    If Me.CboGuar.ListIndex <> -1 Then
        If Me.CboGuar.ListIndex = 0 Then
            If SystemOptions.SysDataBaseType = SQLServerDataBase Then
                StrWhere = StrWhere + " and HaveGuarantee =1"
            Else
                StrWhere = StrWhere + " and HaveGuarantee =True"
            End If

        ElseIf Me.CboGuar.ListIndex = 1 Then
            If SystemOptions.SysDataBaseType = SQLServerDataBase Then
                StrWhere = StrWhere + " and HaveGuarantee =0"
            Else
                StrWhere = StrWhere + " and HaveGuarantee =False"
            End If
        End If
    End If
    If Me.CboArchive.ListIndex <> -1 Then
        If Me.CboArchive.ListIndex = 0 Then
            If SystemOptions.SysDataBaseType = SQLServerDataBase Then
                StrWhere = StrWhere + " and IsArchive =1"
            Else
                StrWhere = StrWhere + " and IsArchive =True"
            End If
        ElseIf Me.CboArchive.ListIndex = 1 Then

            If SystemOptions.SysDataBaseType = SQLServerDataBase Then
                StrWhere = StrWhere + " and IsArchive =0"
            Else
                StrWhere = StrWhere + " and IsArchive =False"
            End If
        End If
    End If

    If Me.CboAssbliedItem.ListIndex <> -1 Then
        If Me.CboAssbliedItem.ListIndex = 0 Then
            If SystemOptions.SysDataBaseType = SQLServerDataBase Then
                StrWhere = StrWhere + " and AssbliedItem =1"
            Else
                StrWhere = StrWhere + " and AssbliedItem =True"
            End If

        ElseIf Me.CboAssbliedItem.ListIndex = 1 Then

            If SystemOptions.SysDataBaseType = SQLServerDataBase Then
                StrWhere = StrWhere + " and AssbliedItem =0"
            Else
                StrWhere = StrWhere + " and AssbliedItem =False"
            End If
        End If
    End If

    If Me.CboAttachedItem.ListIndex <> -1 Then
        If Me.CboAttachedItem.ListIndex = 0 Then
            If SystemOptions.SysDataBaseType = SQLServerDataBase Then
                StrWhere = StrWhere + " and RelatedItem =1"
            Else
                StrWhere = StrWhere + " and RelatedItem =True"
            End If

        ElseIf Me.CboAttachedItem.ListIndex = 1 Then

            If SystemOptions.SysDataBaseType = SQLServerDataBase Then
                StrWhere = StrWhere + " and RelatedItem =0"
            Else
                StrWhere = StrWhere + " and RelatedItem =False"
            End If
        End If
    End If

    Build_Sql = StrSQL + StrWhere
    Exit Function
ErrTrap:
End Function
Sub FillItemsIDes()
Dim i As Integer
ItemsIDes = "0"
With FG
For i = 1 To .rows - 1
If .cell(flexcpChecked, i, .ColIndex("Send")) = flexChecked Then
ItemsIDes = ItemsIDes & "," & val(.TextMatrix(i, 2))
End If
Next i
End With
End Sub
Private Sub Check17_Click()
    Dim i As Integer

    If Check17.value = vbChecked Then

        With Me.FG
 
            For i = 1 To .rows - 1
        
                .TextMatrix(i, .ColIndex("Send")) = True
            Next i

        End With

    Else

        With Me.FG

            For i = 1 To .rows - 1
        
                .TextMatrix(i, .ColIndex("Send")) = False
            Next i

        End With

    End If

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        If Me.ActiveControl Is FG Then
            If Not FG.TextMatrix(FG.row, 1) = "" Then
                fg_Click
                Unload Me
            End If
        Else
            Cmd_Click (0)
        End If
    End If

    On Error GoTo ErrTrap

    If Shift = 2 Then
        If KeyCode = vbKeyX Then
            Cmd_Click (2)
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Public Property Get DcboItems() As DataCombo
    Set DcboItems = m_DcboItems
End Property

Public Property Set DcboItems(ByVal vNewValue As DataCombo)
    Set m_DcboItems = vNewValue
End Property

Public Property Get RetrunType() As Integer
    RetrunType = m_RetrunType
End Property

Public Property Let RetrunType(ByVal vNewValue As Integer)
    m_RetrunType = vNewValue
    ' 0 = Retrun in the Items Screen
    ' 1 = Retrun in the Data Combo
End Property

Private Sub ChangeLang()
Check17.RightToLeft = False
Check17.Caption = "Select All"
    Me.Caption = "Search For Item"
    lbl(0).Caption = "Item Code"
    lbl(1).Caption = "Item Name"
    lbl(2).Caption = "Serial Type"
    lbl(3).Caption = "Group Name"
    lbl(4).Caption = "Item ID"
    lbl(5).Caption = "Item Type"
    lbl(6).Caption = "Assembled"
    lbl(7).Caption = "Attached"
    lbl(8).Caption = "Match Type"
    lbl(9).Caption = "Guarantee"
    lbl(10).Caption = "Archives"
    lbl(11).Caption = "Match Type"
    lbl(12).Caption = "BarCode"
     Cmd(3).Caption = "Multiple Choice"
    'Label1.Caption = "Part No"
    XPLbl(7).Caption = "Sale Price"
    Cmd(0).Caption = "Search"
    Cmd(1).Caption = "Clear"
    Cmd(2).Caption = "Exit"

    'OptType(0).Caption = "Start of the name"
    'OptType(1).Caption = "any part of the name"
    With Me.FG
        .TextMatrix(0, .ColIndex("NumIndex")) = "Serial"
        .TextMatrix(0, .ColIndex("ItemNum")) = "Item ID"
        .TextMatrix(0, .ColIndex("KindCode")) = "Item Code"
        .TextMatrix(0, .ColIndex("barCodeNO")) = "Bar Code"
        
        .TextMatrix(0, .ColIndex("PartNo")) = "Part No"
        
        .TextMatrix(0, .ColIndex("KindNme")) = "Item Name"
        .TextMatrix(0, .ColIndex("ItemType")) = "Item Type"
        .TextMatrix(0, .ColIndex("HaveSerial")) = "Have Serial"
        .TextMatrix(0, .ColIndex("IsArchive")) = "Archive"
        .TextMatrix(0, .ColIndex("HaveGuarantee")) = "Guarantee"
        .TextMatrix(0, .ColIndex("AssbliedItem")) = "Assblied"
        .TextMatrix(0, .ColIndex("RelatedItem")) = "Attached Items"
        .AutoSize 0, .Cols - 1, False
    End With

End Sub


Private Sub GrdF_Click()
If Indx = 1 Then
    frmsalebill3.txtCallingID = val(GrdF.TextMatrix(GrdF.row, 1))
Else
 
End If
 Unload Me

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
        mTableName = "TblStudCalling"

    End If
    
    
    
        StrSQL = " SELECT TT.ID           ,"
        StrSQL = StrSQL & " TT.EnterDate RecordDate,"
        StrSQL = StrSQL & " TblCustemers.CusName,"
        StrSQL = StrSQL & " tbd.branch_name,"
        StrSQL = StrSQL & " TblCustemers.CusID"
        StrSQL = StrSQL & " From TblStudCalling as TT"
        StrSQL = StrSQL & " INNER JOIN TblCustemers"
        StrSQL = StrSQL & "            ON  TblCustemers.CusID = TT.CompID"
        StrSQL = StrSQL & " INNER JOIN TblBranchesData  AS tbd"
        StrSQL = StrSQL & "            ON  TT.BranchID = tbd.branch_id"
        StrSQL = StrSQL & " where 1=1 "
   
    Begin = True
  

    If txtID.text <> "" Then
              StrWhere = StrWhere + " and TT.ID like '%" & (txtID.text) & "%'"
      End If
    If DcbCus.text <> "" Then
                 StrWhere = StrWhere + " and ( TblCustemers.CustID LIKE '%" & Trim(DcbCus.BoundText) & "%')"
    End If
    If DataBranch.BoundText <> "" Then
                 StrWhere = StrWhere + " and ( TT.BranchID LIKE '%" & Trim(DataBranch.BoundText) & "%')"
    End If
    
      
      
      If Not IsNull(Me.txtFromDate.value) Then
             StrWhere = StrWhere & " AND TT.EnterDate>=" & SQLDate(Me.txtFromDate.value, True) & ""
      End If
       If Not IsNull(Me.txtToDate.value) Then
             StrWhere = StrWhere & " AND TT.EnterDate <=" & SQLDate(Me.txtToDate.value, True) & ""
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

Private Sub ISButton17_Click()
 clear_all Me
txtFromDate.value = ""
txtToDate.value = ""

End Sub

Private Sub ISButton18_Click()
Unload Me
End Sub

Private Sub TxtItemName_Change()
    If Trim$(Me.txtItemName.text) = "" Then
        Me.lbl(8).Enabled = False
        Me.CboNameSearch.Enabled = False
    Else
        Me.lbl(8).Enabled = True
        Me.CboNameSearch.Enabled = True
    End If
End Sub
Private Sub XPTxtItemCode_Change()
    If Trim$(Me.XPTxtItemCode.text) = "" Then
        Me.lbl(11).Enabled = False
        Me.CboItemCodeSearch.Enabled = False
    Else
        Me.lbl(11).Enabled = True
        Me.CboItemCodeSearch.Enabled = True
    End If
End Sub

