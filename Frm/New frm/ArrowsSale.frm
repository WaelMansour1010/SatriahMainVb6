VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form ArrowsSale 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "»Ì⁄ √”Â„"
   ClientHeight    =   6975
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   13335
   Icon            =   "ArrowsSale.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   6975
   ScaleWidth      =   13335
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.CommandButton Cmd 
      Caption         =   " Õ„Ì· «·«”⁄«— „‰ «·«‰ —‰ "
      Height          =   315
      Index           =   0
      Left            =   8400
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   1935
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   9240
      Width           =   10215
      ExtentX         =   18018
      ExtentY         =   2566
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser2 
      Height          =   1455
      Left            =   0
      TabIndex        =   2
      Top             =   9480
      Width           =   10815
      ExtentX         =   19076
      ExtentY         =   2566
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin C1SizerLibCtl.C1Elastic EleTop 
      Height          =   660
      Left            =   0
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   13335
      _cx             =   23521
      _cy             =   1164
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial (Arabic)"
         Size            =   20.25
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
      BackColor       =   12648447
      ForeColor       =   8421376
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   "»Ì⁄ √”Â„"
      Align           =   1
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
      Style           =   0
      TagSplit        =   2
      PicturePos      =   7
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
   Begin C1SizerLibCtl.C1Tab XPTab301 
      Height          =   6270
      Left            =   5880
      TabIndex        =   4
      Top             =   720
      Width           =   7410
      _cx             =   13070
      _cy             =   11060
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
      ForeColor       =   0
      FrontTabColor   =   14871017
      BackTabColor    =   12648447
      TabOutlineColor =   -2147483632
      FrontTabForeColor=   16711680
      Caption         =   "«·«”Â„ «·„„·ÊﬂÂ|»Ì«‰«  ”Â„"
      Align           =   0
      CurrTab         =   0
      FirstTab        =   0
      Style           =   3
      Position        =   0
      AutoSwitch      =   -1  'True
      AutoScroll      =   -1  'True
      TabPreview      =   -1  'True
      ShowFocusRect   =   -1  'True
      TabsPerPage     =   0
      BorderWidth     =   0
      BoldCurrent     =   0   'False
      DogEars         =   0   'False
      MultiRow        =   0   'False
      MultiRowOffset  =   200
      CaptionStyle    =   0
      TabHeight       =   0
      TabCaptionPos   =   4
      TabPicturePos   =   1
      CaptionEmpty    =   ""
      Separators      =   0   'False
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   37
      Picture(0)      =   "ArrowsSale.frx":000C
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   5805
         Index           =   0
         Left            =   45
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   420
         Width           =   7320
         _cx             =   12912
         _cy             =   10239
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
         Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid1 
            Height          =   3060
            Left            =   120
            TabIndex        =   7
            Top             =   2640
            Width           =   6915
            _cx             =   12197
            _cy             =   5397
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
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   16777215
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
            Rows            =   2
            Cols            =   26
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"ArrowsSale.frx":03A6
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
         Begin VSFlex8UCtl.VSFlexGrid Grid 
            Height          =   1605
            Left            =   120
            TabIndex        =   9
            Top             =   600
            Width           =   6990
            _cx             =   12330
            _cy             =   2831
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
            FormatString    =   $"ArrowsSale.frx":071B
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
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«Œ — «·„Õ›Ÿ…"
            Height          =   255
            Left            =   4800
            RightToLeft     =   -1  'True
            TabIndex        =   10
            Top             =   240
            Width           =   2175
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«Œ — «·”Â„ «·–Ì  —Ìœ »Ì⁄Â"
            Height          =   255
            Left            =   4920
            RightToLeft     =   -1  'True
            TabIndex        =   8
            Top             =   2400
            Width           =   2175
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Height          =   375
            Left            =   4560
            RightToLeft     =   -1  'True
            TabIndex        =   6
            Top             =   840
            Width           =   7575
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   5805
         Index           =   1
         Left            =   8055
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   420
         Width           =   7320
         _cx             =   12912
         _cy             =   10239
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
         Begin VB.TextBox Text16 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1560
            RightToLeft     =   -1  'True
            TabIndex        =   59
            Top             =   4560
            Width           =   3615
         End
         Begin VB.TextBox Text15 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1560
            RightToLeft     =   -1  'True
            TabIndex        =   57
            Top             =   4200
            Width           =   3615
         End
         Begin VB.TextBox Text14 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1560
            RightToLeft     =   -1  'True
            TabIndex        =   55
            Top             =   3840
            Width           =   3615
         End
         Begin VB.TextBox Text13 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1560
            RightToLeft     =   -1  'True
            TabIndex        =   53
            Top             =   3480
            Width           =   3615
         End
         Begin VB.TextBox Text12 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1560
            RightToLeft     =   -1  'True
            TabIndex        =   51
            Top             =   3120
            Width           =   3615
         End
         Begin VB.TextBox Text11 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1560
            RightToLeft     =   -1  'True
            TabIndex        =   49
            Top             =   2760
            Width           =   3615
         End
         Begin VB.TextBox Text10 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1560
            RightToLeft     =   -1  'True
            TabIndex        =   47
            Top             =   2400
            Width           =   3615
         End
         Begin VB.TextBox Text9 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1560
            RightToLeft     =   -1  'True
            TabIndex        =   45
            Top             =   2040
            Width           =   3615
         End
         Begin VB.TextBox Text8 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1560
            RightToLeft     =   -1  'True
            TabIndex        =   43
            Top             =   1680
            Width           =   3615
         End
         Begin VB.TextBox Text7 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1560
            RightToLeft     =   -1  'True
            TabIndex        =   41
            Top             =   1320
            Width           =   3615
         End
         Begin VB.TextBox Text6 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1560
            RightToLeft     =   -1  'True
            TabIndex        =   39
            Top             =   960
            Width           =   3615
         End
         Begin VB.TextBox Text3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1560
            RightToLeft     =   -1  'True
            TabIndex        =   37
            Top             =   600
            Width           =   3615
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "„»·€ «·—»Õ/«·Œ”«—…"
            Height          =   255
            Index           =   17
            Left            =   5280
            RightToLeft     =   -1  'True
            TabIndex        =   60
            Top             =   4560
            Width           =   1335
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "ﬁÌ„… «·—»Õ/«·Œ”«—…"
            Height          =   255
            Index           =   16
            Left            =   5280
            RightToLeft     =   -1  'True
            TabIndex        =   58
            Top             =   4200
            Width           =   1335
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·ﬁÌ„Â «·”ÊﬁÌ…"
            Height          =   255
            Index           =   15
            Left            =   5520
            RightToLeft     =   -1  'True
            TabIndex        =   56
            Top             =   3840
            Width           =   1095
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«Œ— ”⁄—"
            Height          =   255
            Index           =   14
            Left            =   5520
            RightToLeft     =   -1  'True
            TabIndex        =   54
            Top             =   3480
            Width           =   1095
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "’«›Ì «·‘—«¡"
            Height          =   255
            Index           =   13
            Left            =   5520
            RightToLeft     =   -1  'True
            TabIndex        =   52
            Top             =   3120
            Width           =   1095
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "⁄„Ê·Â «·‘—«¡"
            Height          =   255
            Index           =   12
            Left            =   5520
            RightToLeft     =   -1  'True
            TabIndex        =   50
            Top             =   2760
            Width           =   1095
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«Ã„«·Ì «·‘—«¡"
            Height          =   255
            Index           =   11
            Left            =   5520
            RightToLeft     =   -1  'True
            TabIndex        =   48
            Top             =   2400
            Width           =   1095
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "”⁄— ‘—«¡ «·”Â„"
            Height          =   255
            Index           =   10
            Left            =   5400
            RightToLeft     =   -1  'True
            TabIndex        =   46
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·«”Â„ «·„„·ÊﬂÂ"
            Height          =   255
            Index           =   9
            Left            =   5400
            RightToLeft     =   -1  'True
            TabIndex        =   44
            Top             =   1680
            Width           =   1215
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«”„ «·‘—ﬂÂ"
            Height          =   255
            Index           =   8
            Left            =   5520
            RightToLeft     =   -1  'True
            TabIndex        =   42
            Top             =   1320
            Width           =   1095
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "—„“ «·‘—ﬂÂ"
            Height          =   255
            Index           =   7
            Left            =   5520
            RightToLeft     =   -1  'True
            TabIndex        =   40
            Top             =   960
            Width           =   1095
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   " «—ÌŒ «·‘—«¡"
            Height          =   255
            Index           =   3
            Left            =   5520
            RightToLeft     =   -1  'True
            TabIndex        =   38
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Height          =   375
            Left            =   4560
            RightToLeft     =   -1  'True
            TabIndex        =   29
            Top             =   840
            Width           =   7575
         End
      End
   End
   Begin C1SizerLibCtl.C1Elastic Ele 
      Height          =   2685
      Index           =   2
      Left            =   120
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   600
      Width           =   5760
      _cx             =   10160
      _cy             =   4736
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
      Begin VB.CommandButton Cmd 
         Caption         =   "Œ—ÊÃ"
         Height          =   315
         Index           =   4
         Left            =   240
         TabIndex        =   33
         Top             =   2280
         Width           =   855
      End
      Begin VB.CommandButton Cmd 
         Caption         =   "  ⁄œÌ· »Ì«‰«  «·»Ì⁄"
         Height          =   315
         Index           =   3
         Left            =   1320
         TabIndex        =   32
         Top             =   2280
         Width           =   1455
      End
      Begin VB.CommandButton Cmd 
         Caption         =   " Õ–› »Ì«‰«  «·»Ì⁄"
         Height          =   315
         Index           =   2
         Left            =   2760
         TabIndex        =   31
         Top             =   2280
         Width           =   1335
      End
      Begin VB.CommandButton Cmd 
         Caption         =   " Õ›Ÿ  »Ì«‰«  «·»Ì⁄"
         Height          =   315
         Index           =   1
         Left            =   4200
         TabIndex        =   30
         Top             =   2280
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   1920
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "Õ”«» «·⁄„Ê·Â „‰ ‰”»… «·‘»ﬂÂ"
         Height          =   255
         Index           =   0
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   1200
         Width           =   2655
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "Õ”«» «·⁄„Ê·Â „‰ ‰”»… «·»‰ﬂ"
         Height          =   255
         Index           =   1
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   1560
         Width           =   2655
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "Õ”«» «·⁄„Ê·Â „‰ «ﬁ· „»·€ ⁄„Ê·Â"
         Height          =   255
         Index           =   2
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   1920
         Width           =   2655
      End
      Begin MSComCtl2.DTPicker DTP_Date 
         Height          =   270
         Left            =   2520
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   480
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   476
         _Version        =   393216
         CalendarBackColor=   12648447
         CalendarTitleBackColor=   10383715
         CustomFormat    =   "yyyy/M/d"
         Format          =   104595459
         CurrentDate     =   37140
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "»Ì«‰«  «·»Ì⁄"
         Height          =   255
         Left            =   3240
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   120
         Width           =   2175
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " «—ÌŒ «·»Ì⁄"
         Height          =   255
         Index           =   0
         Left            =   4320
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "⁄œœ «·«”Â„"
         Height          =   255
         Index           =   1
         Left            =   4320
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "”⁄— «·»Ì⁄"
         Height          =   255
         Index           =   2
         Left            =   4320
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "‰”»… «·⁄„Ê·…"
         Height          =   255
         Index           =   4
         Left            =   4320
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "„»·€ «·⁄„Ê·…"
         Height          =   255
         Index           =   5
         Left            =   4320
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Õ”«» «·⁄„Ê·…"
         Height          =   255
         Index           =   6
         Left            =   1560
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   960
         Width           =   1095
      End
   End
   Begin C1SizerLibCtl.C1Elastic Ele 
      Height          =   3765
      Index           =   3
      Left            =   120
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   3240
      Width           =   5760
      _cx             =   10160
      _cy             =   6641
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
      Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid2 
         Height          =   2925
         Left            =   0
         TabIndex        =   34
         Top             =   600
         Width           =   5670
         _cx             =   10001
         _cy             =   5159
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
         FormatString    =   $"ArrowsSale.frx":0845
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
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "»Ì«‰«  ⁄„·Ì«  «·»Ì⁄ «·”«»ﬁ…"
         Height          =   255
         Left            =   3480
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   4560
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   840
         Width           =   7575
      End
   End
End
Attribute VB_Name = "ArrowsSale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim path As String
Dim NEW_interface As Boolean

Private Sub Cmd_Click(Index As Integer)
    Me.VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
    Me.VSFlexGrid1.Rows = 2

    Select Case Index

        Case 0
            NEW_interface = True
            path = "http://www.tadawul.com.sa/wps/portal/!ut/p/c1/04_SB8K8xLLM9MSSzPy8xBz9CP0os3g_A-ewIE8TIwMLj2AXA0_vQGNzY18g18cQKB-JJO8eEGZq4GniE2wUHOBlbOBpREB3cGKRvp9Hfm6qfkFuRDkAgpcLJw!!/dl2/d1/L2dJQSEvUUt3QS9ZQnB3LzZfTjBDVlJJNDIwMFM1MDBJNExWVENMRzMwMjY!/"
            WebBrowser1.Navigate2 path
            path = "http://www.tadawul.com.sa/wps/portal/!ut/p/c1/04_SB8K8xLLM9MSSzPy8xBz9CP0os3g_A-ewIE8TIwODYFMDA08Tn7AQZx93YwMjM6B8JG55AwOSdLsHhJmC5IONggO8jA08jQjoDk4s0vfzyM9N1S_IDY0od1RUBAD6Iu2e/dl2/d1/L2dJQSEvUUt3QS9ZQnB3LzZfTjBDVlJJNDIwR05QOTBJSzZFSUlEUjAwVDY!/"
            WebBrowser2.Navigate2 path
    End Select

End Sub

Private Sub Form_Load()
    Resize_Form Me
    NEW_interface = False
    'WebBrowser1.Navigate2 "http://www.tadawul.com.sa/Resources/Reports/DetailedDaily_ar.html"
End Sub

Private Sub VSFlexGrid1_Click()

    With VSFlexGrid1

        If Not .TextMatrix(.Row, .ColIndex("HyperLink")) = "" Then
            ArrowsCompanyDetails.show
            ArrowsCompanyDetails.LoadPage .TextMatrix(.Row, .ColIndex("HyperLink")), .TextMatrix(.Row, .ColIndex("Symbol")), .TextMatrix(.Row, .ColIndex("Name"))
        End If

    End With

End Sub

Private Sub WebBrowser1_DocumentComplete(ByVal pDisp As Object, _
                                         URL As Variant)
    On Error GoTo ErrTrap

    If NEW_interface = False Then Exit Sub
    Dim i As Integer

    Dim objTable As Object

    'The ninth table in the page is the Companies List
    Dim startLoad As Integer
    Dim Cols As Integer
    'On Error Resume Next
    startLoad = 75
    Set objTable = WebBrowser1.Document.getElementsByTagName("table").Item(12)

    With Me.VSFlexGrid1
 
        .Rows = objTable.getElementsByTagName("tr").Length - 1
 
        For i = startLoad To .Rows
            Cols = objTable.getElementsByTagName("tr").Item(i).getElementsByTagName("td").Length
            Dim HyperLink  As String
            Dim SymbolNo As Integer

            If Cols >= 2 Then
                .TextMatrix((i - startLoad) + 1, .ColIndex("LineNo")) = (i - startLoad) + 1
                .TextMatrix((i - startLoad) + 1, .ColIndex("Name")) = objTable.getElementsByTagName("tr").Item(i).getElementsByTagName("td").Item(0).innerText
      
            End If
     
            If Cols = 14 Then
                HyperLink = objTable.getElementsByTagName("tr").Item(i).getElementsByTagName("a")
                SymbolNo = right(HyperLink, 4)
                .TextMatrix((i - startLoad) + 1, .ColIndex("Symbol")) = SymbolNo
                .TextMatrix((i - startLoad) + 1, .ColIndex("HyperLink")) = HyperLink
                .TextMatrix((i - startLoad) + 1, .ColIndex("LastPrice")) = objTable.getElementsByTagName("tr").Item(i).getElementsByTagName("td").Item(1).innerText
                .TextMatrix((i - startLoad) + 1, .ColIndex("Change")) = objTable.getElementsByTagName("tr").Item(i).getElementsByTagName("td").Item(3).innerText
                .TextMatrix((i - startLoad) + 1, .ColIndex("ChangePercentage")) = objTable.getElementsByTagName("tr").Item(i).getElementsByTagName("td").Item(4).innerText
                .TextMatrix((i - startLoad) + 1, .ColIndex("NoOfDeals")) = objTable.getElementsByTagName("tr").Item(i).getElementsByTagName("td").Item(5).innerText
                .TextMatrix((i - startLoad) + 1, .ColIndex("Qty")) = objTable.getElementsByTagName("tr").Item(i).getElementsByTagName("td").Item(6).innerText
                .TextMatrix((i - startLoad) + 1, .ColIndex("Opening")) = objTable.getElementsByTagName("tr").Item(i).getElementsByTagName("td").Item(11).innerText
                .TextMatrix((i - startLoad) + 1, .ColIndex("Max")) = objTable.getElementsByTagName("tr").Item(i).getElementsByTagName("td").Item(12).innerText
                .TextMatrix((i - startLoad) + 1, .ColIndex("Min")) = objTable.getElementsByTagName("tr").Item(i).getElementsByTagName("td").Item(13).innerText
            End If

        Next i

        .AutoSize 0, .Cols - 1, False
        Dim j As Integer
        Dim lastindex As Integer

        For j = .Rows - 1 To 2 Step -1

            If .TextMatrix(j, .ColIndex("Name")) <> "" Then
                lastindex = j + 1
                GoTo LL
            End If

        Next j

LL:
        .Rows = lastindex + 1
    End With

    Set objTable = Nothing
    Exit Sub
ErrTrap:
    MsgBox "·«»œ „‰ «·« ’«· »«·«‰ —‰  «Ê·«"

End Sub

Private Sub WebBrowser2_DocumentComplete(ByVal pDisp As Object, _
                                         URL As Variant)
    On Error GoTo ErrTrap
    Dim i As Integer
    Dim objTable As Object

    If NEW_interface = False Then Exit Sub
    'The ninth table in the page is the Companies List
    Set objTable = WebBrowser2.Document.getElementsByTagName("table").Item(5)

    'Now enumerate all TR tags within the table
 
    Label1.Caption = objTable.getElementsByTagName("tr").Item(0).getElementsByTagName("td").Item(1).innerText & vbCrLf

    Set objTable = Nothing
    Exit Sub
ErrTrap:
    MsgBox "·«»œ „‰ «·« ’«· »«·«‰ —‰  «Ê·«"

End Sub

