VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmImport 
   Caption         =   "«” Ì—«œ „‰ «·«ﬂ”·"
   ClientHeight    =   10545
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15840
   LinkTopic       =   "Form1"
   ScaleHeight     =   10545
   ScaleWidth      =   15840
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      Caption         =   "Command6"
      Height          =   195
      Left            =   12390
      TabIndex        =   60
      Top             =   750
      Width           =   615
   End
   Begin VB.TextBox Text2 
      Height          =   405
      Left            =   9090
      TabIndex        =   58
      Text            =   "Text2"
      Top             =   1350
      Width           =   2745
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   6600
      TabIndex        =   57
      Text            =   "Text1"
      Top             =   1170
      Width           =   2205
   End
   Begin VB.TextBox txtBDNAMe 
      Height          =   435
      Left            =   9990
      TabIndex        =   56
      Text            =   "Byte"
      Top             =   150
      Width           =   1395
   End
   Begin VB.CommandButton Command5 
      Caption         =   "«·ﬁÌœ «·«›  «ÕÏ"
      Height          =   435
      Left            =   9810
      TabIndex        =   55
      Top             =   810
      Width           =   2115
   End
   Begin VB.CommandButton cmdFromAccount 
      Caption         =   "Ã·» „‰ «·Õ”«»« "
      Height          =   315
      Left            =   6840
      TabIndex        =   32
      Top             =   570
      Width           =   1425
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Õ›Ÿ"
      Height          =   315
      Left            =   3810
      TabIndex        =   10
      Top             =   240
      Width           =   1545
   End
   Begin VB.Frame Frame1 
      Caption         =   "Server Data"
      Height          =   1455
      Left            =   120
      TabIndex        =   5
      Top             =   60
      Width           =   3375
      Begin VB.TextBox TxtServerDataBaseName 
         Height          =   375
         Left            =   1200
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Text            =   "byte"
         Top             =   840
         Width           =   1815
      End
      Begin VB.TextBox DestinationServer 
         Height          =   375
         Left            =   1200
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Server name"
         Height          =   375
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "DBname"
         Height          =   375
         Left            =   -360
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   960
         Width           =   1335
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   " Õ„Ì· «·„·›..."
      Height          =   285
      Left            =   5490
      TabIndex        =   3
      Top             =   270
      Width           =   1485
   End
   Begin VB.TextBox txtFile 
      Height          =   285
      Left            =   4980
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   630
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   " Õ„Ì· «·„·›..."
      Height          =   255
      Left            =   4110
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   1170
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   " ÕœÌœ «·„·›..."
      Height          =   255
      Left            =   6990
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   300
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VSFlex8Ctl.VSFlexGrid tmpGrd 
      Height          =   1830
      Left            =   13320
      TabIndex        =   4
      Top             =   -120
      Visible         =   0   'False
      Width           =   2265
      _cx             =   3995
      _cy             =   3228
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
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
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
   Begin C1SizerLibCtl.C1Tab TabMain 
      Height          =   8835
      Left            =   60
      TabIndex        =   11
      Top             =   1680
      Width           =   15570
      _cx             =   27464
      _cy             =   15584
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
      Caption         =   "«·„ÊŸ›Ì‰|«·⁄„·«¡ |«·„ÃÊ⁄« |«·ÊÕœ« |«·«’‰«›|„Ã„Ê⁄«  «·⁄„·«¡|«·»‰Êﬂ|«·⁄Âœ|«·„’—Ê›« "
      Align           =   0
      CurrTab         =   4
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
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   8460
         Index           =   1
         Left            =   -17025
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   45
         Width           =   15480
         _cx             =   27305
         _cy             =   14923
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
            Height          =   8340
            Index           =   0
            Left            =   21510
            TabIndex        =   13
            Top             =   735
            Width           =   15345
            _cx             =   27067
            _cy             =   14711
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
            FormatString    =   $"frmImport.frx":0000
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
         Begin VSFlex8Ctl.VSFlexGrid Grd 
            Height          =   8265
            Left            =   0
            TabIndex        =   16
            Top             =   150
            Width           =   15075
            _cx             =   26591
            _cy             =   14579
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
            GridLines       =   3
            GridLinesFixed  =   2
            GridLineWidth   =   5
            Rows            =   2
            Cols            =   69
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmImport.frx":00C0
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
            ExplorerBar     =   3
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
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   8460
         Index           =   0
         Left            =   -16725
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   45
         Width           =   15480
         _cx             =   27305
         _cy             =   14923
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
         Begin VB.OptionButton Option1 
            Caption         =   "„Ê—œÌ‰"
            Height          =   345
            Left            =   3150
            TabIndex        =   31
            Top             =   0
            Width           =   1920
         End
         Begin VB.OptionButton Option2 
            Caption         =   "⁄„·«¡"
            Height          =   315
            Left            =   5205
            TabIndex        =   30
            Top             =   60
            Value           =   -1  'True
            Width           =   825
         End
         Begin VSFlex8UCtl.VSFlexGrid FgItems 
            Height          =   8340
            Index           =   1
            Left            =   21510
            TabIndex        =   15
            Top             =   735
            Width           =   15345
            _cx             =   27067
            _cy             =   14711
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
            FormatString    =   $"frmImport.frx":0DAE
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
         Begin VSFlex8Ctl.VSFlexGrid grdMan 
            Height          =   7815
            Left            =   270
            TabIndex        =   24
            Top             =   510
            Width           =   15075
            _cx             =   26591
            _cy             =   13785
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
            GridLines       =   3
            GridLinesFixed  =   2
            GridLineWidth   =   5
            Rows            =   2
            Cols            =   16
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmImport.frx":0E6E
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
            ExplorerBar     =   3
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
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   8460
         Index           =   2
         Left            =   -16425
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   45
         Width           =   15480
         _cx             =   27305
         _cy             =   14923
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
            Height          =   8340
            Index           =   2
            Left            =   21510
            TabIndex        =   18
            Top             =   735
            Width           =   15345
            _cx             =   27067
            _cy             =   14711
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
            FormatString    =   $"frmImport.frx":1170
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
         Begin VSFlex8Ctl.VSFlexGrid grdGroups 
            Height          =   8265
            Left            =   135
            TabIndex        =   23
            Top             =   120
            Width           =   15075
            _cx             =   26591
            _cy             =   14579
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
            GridLines       =   3
            GridLinesFixed  =   2
            GridLineWidth   =   5
            Rows            =   2
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmImport.frx":1230
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
            ExplorerBar     =   3
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
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   8460
         Index           =   3
         Left            =   -16125
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   45
         Width           =   15480
         _cx             =   27305
         _cy             =   14923
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
            Height          =   8340
            Index           =   3
            Left            =   21510
            TabIndex        =   20
            Top             =   735
            Width           =   15345
            _cx             =   27067
            _cy             =   14711
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
            FormatString    =   $"frmImport.frx":13D7
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
         Begin VSFlex8Ctl.VSFlexGrid grdUnits 
            Height          =   8265
            Left            =   0
            TabIndex        =   25
            Top             =   0
            Width           =   15075
            _cx             =   26591
            _cy             =   14579
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
            GridLines       =   3
            GridLinesFixed  =   2
            GridLineWidth   =   5
            Rows            =   2
            Cols            =   5
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmImport.frx":1497
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
            ExplorerBar     =   3
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
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   8460
         Index           =   4
         Left            =   45
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   45
         Width           =   15480
         _cx             =   27305
         _cy             =   14923
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
         Begin VB.TextBox XPTxtBillID 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   0
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   51
            Top             =   0
            Visible         =   0   'False
            Width           =   1920
         End
         Begin VB.CommandButton Command4 
            Caption         =   "«œŒ«· «·Ã—œ"
            Enabled         =   0   'False
            Height          =   495
            Left            =   3150
            TabIndex        =   50
            Top             =   840
            Width           =   2190
         End
         Begin VB.TextBox TxtTransSerial 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   4665
            RightToLeft     =   -1  'True
            TabIndex        =   48
            Top             =   60
            Width           =   540
         End
         Begin VSFlex8UCtl.VSFlexGrid FgItems 
            Height          =   8340
            Index           =   4
            Left            =   21510
            TabIndex        =   22
            Top             =   735
            Width           =   15345
            _cx             =   27067
            _cy             =   14711
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
            FormatString    =   $"frmImport.frx":1557
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
         Begin VSFlex8Ctl.VSFlexGrid grdItems 
            Height          =   6735
            Left            =   -135
            TabIndex        =   26
            Top             =   1680
            Width           =   15060
            _cx             =   26564
            _cy             =   11880
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
            GridLines       =   3
            GridLinesFixed  =   2
            GridLineWidth   =   5
            Rows            =   2
            Cols            =   20
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmImport.frx":1617
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
            ExplorerBar     =   3
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
         Begin MSComCtl2.DTPicker XPDtbBill 
            Height          =   345
            Left            =   11235
            TabIndex        =   39
            Top             =   375
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   609
            _Version        =   393216
            Format          =   139591683
            CurrentDate     =   38784
         End
         Begin MSDataListLib.DataCombo DCboStoreName 
            Height          =   315
            Left            =   6990
            TabIndex        =   41
            Top             =   870
            Width           =   2460
            _ExtentX        =   4339
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   705
            Index           =   8
            Left            =   7950
            TabIndex        =   42
            TabStop         =   0   'False
            Top             =   30
            Width           =   2460
            _cx             =   4339
            _cy             =   1244
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
            ForeColor       =   16711680
            FloodColor      =   6553600
            ForeColorDisabled=   -2147483631
            Caption         =   " ÕœÌœ «·› —… «·“„‰Ì…"
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
            Begin MSComCtl2.DTPicker DTPickerAccFrom 
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "dd/MM/yyyy"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   11265
                  SubFormatType   =   3
               EndProperty
               Height          =   345
               Left            =   2850
               TabIndex        =   43
               ToolTipText     =   "„‰  «—ÌŒ ﬁœÌ„"
               Top             =   720
               Visible         =   0   'False
               Width           =   1500
               _ExtentX        =   2646
               _ExtentY        =   609
               _Version        =   393216
               CalendarBackColor=   -2147483624
               CalendarTitleBackColor=   10383715
               CheckBox        =   -1  'True
               CustomFormat    =   "yyyy/M/d"
               Format          =   139591683
               CurrentDate     =   37357
            End
            Begin MSComCtl2.DTPicker DTPickerAccTo 
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   "dd/MM/yyyy"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   11265
                  SubFormatType   =   3
               EndProperty
               Height          =   345
               Left            =   90
               TabIndex        =   44
               ToolTipText     =   " ≈·Ï  «—ÌŒ √ÕœÀ"
               Top             =   240
               Width           =   1500
               _ExtentX        =   2646
               _ExtentY        =   609
               _Version        =   393216
               CalendarBackColor=   -2147483624
               CalendarTitleBackColor=   10383715
               CheckBox        =   -1  'True
               CustomFormat    =   "yyyy/M/d"
               Format          =   139591683
               CurrentDate     =   37357
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " «—ÌŒ «·Ã—œ"
               ForeColor       =   &H00FF8080&
               Height          =   285
               Index           =   11
               Left            =   1620
               RightToLeft     =   -1  'True
               TabIndex        =   46
               Top             =   240
               Width           =   795
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„‰"
               Height          =   285
               Index           =   10
               Left            =   4350
               RightToLeft     =   -1  'True
               TabIndex        =   45
               Top             =   765
               Visible         =   0   'False
               Width           =   555
            End
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„”·”·"
            Height          =   375
            Index           =   1
            Left            =   4800
            RightToLeft     =   -1  'True
            TabIndex        =   49
            Top             =   75
            Width           =   1770
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·„Œ“‰"
            Height          =   375
            Index           =   2
            Left            =   9585
            RightToLeft     =   -1  'True
            TabIndex        =   47
            Top             =   840
            Width           =   960
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «·«œŒ«·"
            Height          =   375
            Index           =   0
            Left            =   12735
            RightToLeft     =   -1  'True
            TabIndex        =   40
            Top             =   360
            Width           =   690
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   8460
         Index           =   5
         Left            =   16215
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   45
         Width           =   15480
         _cx             =   27305
         _cy             =   14923
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
            Height          =   8340
            Index           =   5
            Left            =   21510
            TabIndex        =   28
            Top             =   735
            Width           =   15345
            _cx             =   27067
            _cy             =   14711
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
            FormatString    =   $"frmImport.frx":1932
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
         Begin VSFlex8Ctl.VSFlexGrid grdGroups2 
            Height          =   8265
            Left            =   135
            TabIndex        =   29
            Top             =   120
            Width           =   15075
            _cx             =   26591
            _cy             =   14579
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
            GridLines       =   3
            GridLinesFixed  =   2
            GridLineWidth   =   5
            Rows            =   2
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmImport.frx":19F2
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
            ExplorerBar     =   3
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
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   8460
         Index           =   6
         Left            =   16515
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   45
         Width           =   15480
         _cx             =   27305
         _cy             =   14923
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
            Height          =   8340
            Index           =   6
            Left            =   21510
            TabIndex        =   34
            Top             =   735
            Width           =   15345
            _cx             =   27067
            _cy             =   14711
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
            FormatString    =   $"frmImport.frx":1BA2
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
         Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid1 
            Height          =   8265
            Left            =   135
            TabIndex        =   35
            Top             =   120
            Width           =   15075
            _cx             =   26591
            _cy             =   14579
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
            GridLines       =   3
            GridLinesFixed  =   2
            GridLineWidth   =   5
            Rows            =   2
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmImport.frx":1C62
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
            ExplorerBar     =   3
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
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   8460
         Index           =   7
         Left            =   16815
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   45
         Width           =   15480
         _cx             =   27305
         _cy             =   14923
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
            Height          =   8340
            Index           =   7
            Left            =   21510
            TabIndex        =   37
            Top             =   735
            Width           =   15345
            _cx             =   27067
            _cy             =   14711
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
            FormatString    =   $"frmImport.frx":1E12
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
         Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid2 
            Height          =   8265
            Left            =   135
            TabIndex        =   38
            Top             =   120
            Width           =   15075
            _cx             =   26591
            _cy             =   14579
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
            GridLines       =   3
            GridLinesFixed  =   2
            GridLineWidth   =   5
            Rows            =   2
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmImport.frx":1ED2
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
            ExplorerBar     =   3
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
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   8460
         Index           =   9
         Left            =   17115
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   45
         Width           =   15480
         _cx             =   27305
         _cy             =   14923
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
            Height          =   8340
            Index           =   8
            Left            =   21510
            TabIndex        =   53
            Top             =   735
            Width           =   15345
            _cx             =   27067
            _cy             =   14711
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
            FormatString    =   $"frmImport.frx":2082
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
         Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid3 
            Height          =   8265
            Left            =   135
            TabIndex        =   54
            Top             =   120
            Width           =   15075
            _cx             =   26591
            _cy             =   14579
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
            GridLines       =   3
            GridLinesFixed  =   2
            GridLineWidth   =   5
            Rows            =   2
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmImport.frx":2142
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
            ExplorerBar     =   3
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
   Begin MSComCtl2.DTPicker txtDate 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11265
         SubFormatType   =   3
      EndProperty
      Height          =   345
      Left            =   11550
      TabIndex        =   59
      ToolTipText     =   " ≈·Ï  «—ÌŒ √ÕœÀ"
      Top             =   180
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   609
      _Version        =   393216
      CalendarBackColor=   -2147483624
      CalendarTitleBackColor=   10383715
      CheckBox        =   -1  'True
      CustomFormat    =   "yyyy/M/d"
      Format          =   139657219
      CurrentDate     =   37357
   End
End
Attribute VB_Name = "frmImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mIndex As Long

Private Sub GetFromAccount(ByVal mTable As String, ByVal mType As Integer)
Dim s As String
Dim rsDummy As New ADODB.Recordset
Dim rsData As New ADODB.Recordset
Dim astrSplit2tems2() As String
Dim mCode As String
    Dim mSer As Long
    Dim mMaxId As Long
Select Case mTable
Case "TblEmployee"
   
    
    s = "SELECT Max(Emp_ID) MaxID  FROM " & mTable & " AS te "
    rsDummy.Open s, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    If Not rsDummy.EOF Then
        mMaxId = Val(rsDummy!MaxId & "")
    End If
    rsDummy.Close
    
    s = " SELECT REPLACE( Replace(Account_Name,'–„„',''),'/','') Name ,Account_Code FROM ACCOUNTS WHERE Parent_Account_Code ="
    s = s & " (SELECT TOP 1 account_code  FROM ACCOUNTS AS a WHERE a.Account_Name LIKE '%–„„ «·⁄«„·Ì‰%'"
    s = s & " AND a.last_account = 0"
    s = s & " ORDER BY a.Account_Code DESC)"
    s = s & " AND last_account = 1"
    s = s & " and Account_Code Not In (Select Account_Code From " & mTable & " ) "
    rsDummy.Open s, Cn, adOpenForwardOnly, adLockReadOnly
    
    s = "Select * from TblEmployee Where 1 = -1"
    rsData.Open s, Cn, adOpenKeyset, adLockOptimistic
    
    mSer = mMaxId
    Do While Not rsDummy.EOF
        mSer = mSer + 1
        rsData.AddNew
        rsData!Fullcode = GetCode(mSer)
        rsData!Emp_Code = GetCode(mSer)
        rsData!Emp_ID = mSer
        rsData!emp_Name = Trim(rsDummy!Name & "")
        astrSplit2tems2 = Split(Trim(rsDummy!Name & ""), " ")
        rsData!BranchID = 1
        rsData!Emp_Name1 = astrSplit2tems2(0)
         If UBound(astrSplit2tems2) > 0 Then
            rsData!Emp_Name2 = astrSplit2tems2(1)
        End If
        If UBound(astrSplit2tems2) > 1 Then
            rsData!Emp_Name3 = astrSplit2tems2(2)
         End If
        If UBound(astrSplit2tems2) > 2 Then
            rsData!Emp_Name4 = astrSplit2tems2(3)
         End If
        rsData!Account_Code = Trim(rsDummy!Account_Code & "")
        rsData.Update
        rsDummy.MoveNext
    Loop
    s = "UPDATE     TblEmployee SET JobTypeID = 4 WHERE ISNULL(JobTypeID,0) = 0 ;" & vbNewLine
    s = s & " UPDATE     TblEmployee SET JobTypeID = 4 WHERE ISNULL(JobTypeID,0) = 0 ;" & vbNewLine
    s = s & " UPDATE     TblEmployee SET dean = N'„”·„' WHERE ISNULL(dean,'') = '' ;" & vbNewLine
    s = s & " UPDATE     TblEmployee SET DepartmentID  = 1 WHERE ISNULL(DepartmentID,0) = 0 ;" & vbNewLine
    s = s & " UPDATE     TblEmployee SET jopstatusid = 1  WHERE ISNULL(jopstatusid,0) = 0 ;" & vbNewLine
    s = s & " UPDATE     TblEmployee SET workstate  = 1  WHERE ISNULL(workstate,0) = 0 ;" & vbNewLine
    s = s & " UPDATE     TblEmployee SET workstate  = 1  WHERE ISNULL(workstate,0) = 0 ;" & vbNewLine
    s = s & " UPDATE     TblEmployee SET Account_code1 =Account_code     WHERE ISNULL(Account_code1,'') = '' ;" & vbNewLine
    s = s & " UPDATE     TblEmployee SET Account_code2 =Account_code     WHERE ISNULL(Account_code2,'') = '' ;" & vbNewLine
    s = s & " UPDATE     TblEmployee SET Account_code3 =Account_code     WHERE ISNULL(Account_code3,'') = '' ;" & vbNewLine
    s = s & " UPDATE     TblEmployee SET Account_code4 =Account_code     WHERE ISNULL(Account_code4,'') = '' ;" & vbNewLine
    s = s & " UPDATE     TblEmployee SET Account_Code5  =Account_code     WHERE ISNULL(Account_Code5,'') = ''"
    s = s & " UPDATE     TblEmployee SET Account_codeTEMP   =Account_code     WHERE ISNULL(Account_codeTEMP,'') = '' ;" & vbNewLine
    s = s & " UPDATE     TblEmployee SET Account_code1TEMP    =Account_code     WHERE ISNULL(Account_code1TEMP,'') = '' ;" & vbNewLine
    s = s & " UPDATE     TblEmployee SET Account_code2TEMP    =Account_code     WHERE ISNULL(Account_code2TEMP,'') = '' ;" & vbNewLine
    s = s & " UPDATE     TblEmployee SET Account_code3TEMP=Account_code     WHERE ISNULL(Account_code3TEMP,'') = '' ;" & vbNewLine
    s = s & " UPDATE     TblEmployee SET BranchId =1 WHERE ISNULL(BranchId,0) = 0  ;" & vbNewLine
    Cn.Execute s
    
    s = ""
Case "TblCustemers"
    
    s = "SELECT Max(CusID) MaxID  FROM " & mTable & " AS te "
    rsDummy.Open s, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    If Not rsDummy.EOF Then
        mMaxId = Val(rsDummy!MaxId & "")
    End If
    rsDummy.Close

    If mType = 1 Then
            s = " SELECT  REPLACE( Replace(Account_Name,'⁄„·«¡',''),'/','') NAME,Account_Code"
        s = s & " FROM ACCOUNTS WHERE Parent_Account_Code In"
        s = s & " (SELECT  Account_Code  FROM ACCOUNTS AS a WHERE a.Account_Name LIKE '%⁄„·«¡%' Or a.Account_Name LIKE '%„œÌ‰Ê‰ „ ‰Ê⁄Ê‰%'"
        s = s & " AND a.last_account = 0"
        s = s & " )  AND last_account = 1"
        s = s & " and Account_Code Not In (Select IsNull(Account_Code,'') From " & mTable & " ) "
        


    Else
         s = " SELECT  REPLACE( Replace(Account_Name,'⁄„·«¡',''),'/','') NAME,Account_Code"
        s = s & " FROM ACCOUNTS WHERE Parent_Account_Code In"
        s = s & " (SELECT  Account_Code  FROM ACCOUNTS AS a WHERE a.Account_Name LIKE '%„Ê—œÊ‰%' Or a.Account_Name LIKE '%œ«∆‰Ê‰ „ ‰Ê⁄Ê‰%'"
        s = s & " AND a.last_account = 0"
        s = s & " )  AND last_account = 1"
        s = s & " and Account_Code Not In (Select IsNull(Account_Code,'') From " & mTable & " ) "
    End If
        rsDummy.Open s, Cn, adOpenForwardOnly, adLockReadOnly
    
    s = "Select * from TblCustemers Where 1 = -1"
    rsData.Open s, Cn, adOpenKeyset, adLockOptimistic
    
    mSer = mMaxId
    Do While Not rsDummy.EOF
        mSer = mSer + 1
        rsData.AddNew
        rsData!Fullcode = GetCode(mSer)
        rsData!code = GetCode(mSer)
        rsData!CusID = mSer
        rsData!CusName = Trim(rsDummy!Name & "")
        rsData!CusNamee = Trim(rsDummy!Name & "")
        rsData!Type = mType
        rsData!CreditlimitCredit = 0
        rsData!SaleType = 0
        rsData!Locked = 0
        rsData!CreditlimitCredit = 0
        rsData!CreditlimitCredit = 0
        rsData!BranchID = 1
        rsData!Account_Code = Trim(rsDummy!Account_Code & "")
        rsData.Update
        rsDummy.MoveNext
    Loop
    s = " Update TblCustemers"
    s = s & " Set TblCustemers.parent_account = ACCOUNTS.Parent_Account_Code"
    s = s & " From dbo.TblCustemers"
    s = s & "        INNER JOIN dbo.ACCOUNTS"
    s = s & "                    ON  dbo.TblCustemers.Account_Code = dbo.ACCOUNTS.Account_Code"
    Cn.Execute s
    
Case "TblBoxesData"
    s = " SELECT  REPLACE( Replace(Account_Name,'',''),'/','') NAME,Account_Code"
    s = s & "   FROM ACCOUNTS WHERE Parent_Account_Code In"
    s = s & " (SELECT  account_code  FROM ACCOUNTS AS a WHERE a.Account_Name LIKE '%⁄Âœ%'"
    s = s & " AND a.last_account = 0)"
    s = s & " and last_account = 1 and Account_Code Not In (Select IsNull(Account_Code,'') From " & mTable & " ) "
      rsDummy.Open s, Cn, adOpenForwardOnly, adLockReadOnly
    
    s = "Select * from TblBoxesData Where 1 = -1"
    rsData.Open s, Cn, adOpenKeyset, adLockOptimistic
    
    mSer = mMaxId
    Do While Not rsDummy.EOF
        mSer = mSer + 1
        rsData.AddNew
        'rsData!Fullcode = GetCode(mSer)
        'rsData!Code = GetCode(mSer)
        rsData!BoxID = mSer
        rsData!BoxName = Trim(rsDummy!Name & "")
        rsData!BoxNamee = Trim(rsDummy!Name & "")
        rsData!Type = 1
        rsData!Account_Code = Trim(rsDummy!Account_Code & "")
        rsData.Update
        rsDummy.MoveNext
    Loop
    s = " Update TblBoxesData"
    s = s & " Set TblBoxesData.parent_account = ACCOUNTS.Parent_Account_Code"
    s = s & " From dbo.TblBoxesData"
    s = s & "        INNER JOIN dbo.ACCOUNTS"
    s = s & "             ON  dbo.TblBoxesData.Account_Code = dbo.ACCOUNTS.Account_Code"
    Cn.Execute s
Case "BanksData"
    s = " SELECT  REPLACE( Replace(Account_Name,'',''),'/','') NAME,Account_Code"
    s = s & "   FROM ACCOUNTS WHERE Parent_Account_Code In"
    s = s & " (SELECT  account_code  FROM ACCOUNTS AS a WHERE a.Account_Name LIKE '%»‰ﬂ%'"
    s = s & " AND a.last_account = 0)"
    s = s & " and last_account = 1 and Account_Code Not In (Select IsNull(Account_Code,'') From " & mTable & " ) "
      rsDummy.Open s, Cn, adOpenForwardOnly, adLockReadOnly
    
    s = "Select * from BanksData Where 1 = -1"
    rsData.Open s, Cn, adOpenKeyset, adLockOptimistic
    
    mSer = mMaxId
    Do While Not rsDummy.EOF
        mSer = mSer + 1
        rsData.AddNew
        'rsData!Fullcode = GetCode(mSer)
        'rsData!Code = GetCode(mSer)
        rsData!BankId = mSer
        If Len(Trim(rsDummy!Name & "")) > 50 Then
            rsData!BankName = Right(Trim(rsDummy!Name & ""), 50)
        Else
            rsData!BankName = Trim(rsDummy!Name & "")
        End If
        
        rsData!BankNamee = Right(Trim(rsDummy!Name & ""), 50)
        
        rsData!Account_Code = Trim(rsDummy!Account_Code & "")
        rsData.Update
        rsDummy.MoveNext
    Loop
        s = " Update BanksData"
    s = s & " Set BanksData.parent_account = ACCOUNTS.Parent_Account_Code"
    s = s & " From dbo.BanksData"
    s = s & "        INNER JOIN dbo.ACCOUNTS"
    s = s & "             ON  dbo.BanksData.Account_Code = dbo.ACCOUNTS.Account_Code"
    Cn.Execute s
Case "ExpensesType"
    s = " SELECT  REPLACE( Replace(Account_Name,'',''),'/','') NAME,Account_Code"
    s = s & "   FROM ACCOUNTS WHERE Parent_Account_Code In"
    s = s & " (SELECT  account_code  FROM ACCOUNTS AS a WHERE a.Account_Name LIKE N'%„’—Ê›« %'  Or a.Account_Name LIKE N'% ﬂ«·Ì› «·„‘«—Ì⁄%'  Or a.Account_Name LIKE N'%„’«—Ì›%'"
    s = s & " AND a.last_account = 0)"
    s = s & " and last_account = 1 and Account_Code Not In (Select IsNull(Account_Code,'') From " & mTable & " ) "
      rsDummy.Open s, Cn, adOpenForwardOnly, adLockReadOnly
    
    s = "Select * from ExpensesType Where 1 = -1"
    rsData.Open s, Cn, adOpenKeyset, adLockOptimistic
    
    mSer = mMaxId
    Do While Not rsDummy.EOF
        mSer = mSer + 1
        rsData.AddNew
        'rsData!Fullcode = GetCode(mSer)
        'rsData!Code = GetCode(mSer)
        rsData!ID = mSer
        rsData!Name = Trim(rsDummy!Name & "")
        rsData!Namee = Trim(rsDummy!Name & "")
        
        rsData!TypicalProduction = 0
        rsData!IndirectCosts = 0
        rsData!Account_Code = Trim(rsDummy!Account_Code & "")
        rsData.Update
        rsDummy.MoveNext
    Loop
    
        s = " Update ExpensesType"
    s = s & " Set ExpensesType.parent_account = ACCOUNTS.Parent_Account_Code"
    s = s & " From dbo.ExpensesType"
    s = s & "        INNER JOIN dbo.ACCOUNTS"
    s = s & "             ON  dbo.ExpensesType.Account_Code = dbo.ACCOUNTS.Account_Code"
    Cn.Execute s
Case ""
End Select

s = ""

End Sub
Private Function GetCode(ByVal mValue As Long) As String
If Len(CStr(mValue)) = 1 Then
    GetCode = "0000" & mValue
ElseIf Len(CStr(mValue)) = 2 Then
    GetCode = "000" & mValue
ElseIf Len(CStr(mValue)) = 3 Then
    GetCode = "00" & mValue
ElseIf Len(CStr(mValue)) = 4 Then
    GetCode = "0" & mValue
    
End If

End Function

Private Sub cmdFromAccount_Click()
Select Case mIndex
Case 0
    GetFromAccount "TblEmployee", 0
Case 1
    GetFromAccount "TblCustemers", IIf(Option2, 1, 2)
Case 6
    GetFromAccount "BanksData", 0
Case 7
    GetFromAccount "TblBoxesData", 0
Case 8
    GetFromAccount "ExpensesType", 0
    
End Select
MsgBox " „ ‰ﬁ· «·»Ì«‰« "
End Sub

Private Sub cmdSave_Click()



Dim i As Long
Dim mGrd As Object

Select Case mIndex
Case 0
    Set mGrd = Grd
Case 1
    Set mGrd = grdMan

Case 2
    Set mGrd = grdGroups

Case 3
    Set mGrd = grdUnits
Case 4
    Set mGrd = grdItems
Case 5
    Set mGrd = grdGroups2
End Select

For i = 0 To mGrd.Cols - 1
    If mGrd.ColEditMask(i) <> "" Then
        mGrd.ColHidden(i) = False
    End If
    'Grd.ColComboList(i) = ""
Next
Dim s As String

Select Case mIndex
Case 0
    s = "Select * from TblEmployee Where Emp_ID =  -1"
    saveGridExcel s, mGrd, "Fullcode", "Emp_ID", "TblEmployee"
    s = "update TblEmployee set TblEmployee.workstate =1 where jopstatusid=1"
    Cn.Execute s
    
    s = "UPDATE TblEmployee SET Emp_Name = Emp_Namee WHERE ISNULL(Emp_Name,'') = ''"
    Cn.Execute s
    
    s = "UPDATE TblEmployee SET BranchId = 1 where IsNull(BranchId ,0) = 0 "
    
    Cn.Execute s
    
    s = "UPDATE TblEmployee SET InsuranceState = 1 where IsNull(InsuranceNO,0) <> 0 "
   Cn.Execute s
    s = "UPDATE TblEmployee SET DepartmentID = 1 where IsNull(DepartmentID ,0) = 0 "
    Cn.Execute s
Case 3
    s = "Select * from TblUnites Where UnitID =  -1"
    saveGridExcel s, mGrd, "UnitName", "UnitID", "TblUnites"
Case 2
    
        s = "Select * from groups Where GroupID =  -1"
    saveGridExcel s, mGrd, "Fullcode", "GroupID", "groups"
Case 1
    s = "Select * from TblCustemers Where CusID =  -1"
    saveGridExcel s, mGrd, "Fullcode", "CusID ", "TblCustemers"
    s = " UPDATE TblCustemers SET code = Fullcode,BranchId = null"
    Cn.Execute s
Case 4
    s = "Select * from tblItems Where ItemID =  -1"
    saveGridExcel s, mGrd, "Fullcode", "ItemID ", "tblItems"
    
    s = "Select * from TblItemsUnits Where ItemID =  -1"
    saveGridExcel s, mGrd, "ItemID", "ItemID ", "TblItemsUnits"
    Command4.Enabled = True
    
    
s = " Update TblItems"
s = s & " SET prifix = (SELECT TOP 1  g.Fullcode FROM  Groups AS g WHERE g.GroupID = TblItems.GroupID )"
s = s & " ,Code =  (SELECT TOP 1  REPLACE(TblItems.code,g.Fullcode,'') FROM  Groups AS g WHERE g.GroupID = TblItems.GroupID )"
s = s & " WHERE ISNULL(TblItems.prifix,'') = ''"
Cn.Execute s
s = " UPDATE tblItemsUnits SET DefaultUnit = 1"
Cn.Execute s
s = " UPDATE Groups SET ParentID = 1  WHERE ISNULL(ParentID,0) = 0  and GroupID <> 1"
 Cn.Execute s
Case 5
    
        s = "Select * from GroupsCustomers Where GroupID =  -1"
    saveGridExcel s, mGrd, "Fullcode", "GroupID", "GroupsCustomers"
End Select
MsgBox " „ «·Õ›Ÿ"
For i = 0 To mGrd.Cols - 1
    If mGrd.ColEditMask(i) <> "" Then
        mGrd.ColHidden(i) = True
        
    End If
    If mGrd.ColEditMask(i) = "Date" Then
        mGrd.ColHidden(i) = False
    End If
    'Grd.ColComboList(i) = ""
Next



cmdSave.Enabled = False



End Sub

Private Sub Command1_Click()
CD1.ShowOpen
txtFile.Text = CD1.FileName
End Sub

Private Sub Command2_Click()

ExportToExcel Me, Grd, , , Me.Caption
'FillItem
End Sub




Sub FillItem()
Dim error_string  As String
  error_string = ""
If txtFile.Text = "" Then MsgBox "Õœœ grdGroups·› «Ê·«": Exit Sub
    Dim ExcelObj As Object
    Dim ExcelBook As Object
    Dim ExcelSheet As Object
    Dim i As Integer
    Dim currentvalue As String, mDesc As String
    Dim Name As String
    Dim itemcode As String
    Dim itemqty As Double
    Dim mEqu As String
    Dim des As String
    Dim DebitValue As String
    Dim CreditValue As String
   Grd.Rows = 1
    Set ExcelObj = CreateObject("Excel.Application")
'        Set ExcelSheet = Nothing
'    Set ExcelBook = Nothing
'    Set ExcelObj = Nothing
'
    Set ExcelSheet = CreateObject("Excel.Sheet")
    ExcelObj.Workbooks.Open txtFile.Text   ' App.Path & "\TrialBalance.xls"
DoEvents
    Set ExcelBook = ExcelObj.Workbooks(1)
    Set ExcelSheet = ExcelBook.Worksheets(1)
    IsFromExcel = True
    With ExcelSheet
    i = 2
    Dim j As Long
    Do Until .cells(i, 1) & "" = ""
        
         '  For j = 1 To Grd.Cols - 1
                       
           itemcode = .cells(i, 1)
           itemqty = .cells(i, 2)
           Name = .cells(i, 3)
           mEqu = .cells(i, 4)
           'mDesc = .cells(i, 5)
           If Val(mEqu) = 0 Then
               mEqu = 0
           End If
    addrow2 itemcode, itemqty, Name, mEqu, Name
          i = i + 1
     '  NewGrid.CountItems
    Loop
        End With
    ExcelObj.Workbooks.Close

    Set ExcelSheet = Nothing
    Set ExcelBook = Nothing
    Set ExcelObj = Nothing

        If error_string <> "" Then
            'CreatLog_File_for_error (error_string)
       End If
       IsFromExcel = False
       Me.Grd.Rows = Me.Grd.Rows + 1
'GetNotinGard
'Coloring
End Sub



Function addrow2(Fullcode As String, Qty As Double, Optional Name As String, Optional Eque As String, Optional des As String)
    Dim StrSQL As String
    Dim Rs2 As ADODB.Recordset
    Set Rs2 = New ADODB.Recordset
    Dim UnitID As Double
    Dim LngItemID As Long
    Dim LngUnitID As Long
    Dim ColorID As Integer
    Dim sizeid As Integer
    Dim ClassId As Integer
    Dim ParrtNoCode As String
    Dim ItemDetailedCode As String
 
    Dim Price As Double
  '  UnitID = GetUnitID(Name)
   If Fullcode <> "" Then
   
        LngItemID = 1
    If LngItemID <> 0 Then
    Dim mRow As Long
    
    With Me.Grd
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, .ColIndex("Fullcode")) = Fullcode
        .TextMatrix(.Rows - 1, .ColIndex("Fullcode")) = Qty
       ' .TextMatrix(.Rows - 1, .ColIndex("FixCode")) = IIf(Eque = 0, "", Eque)
       ' .TextMatrix(.Rows - 1, .ColIndex("des")) = des
        
        .Row = .Rows - 1
  
'        If .TextMatrix(.Rows - 1, .ColIndex("FixCode")) <> "" Then
'           ' .Rows = .Rows + 1
'            Grd_AfterEdit .Rows - 1, .ColIndex("FixCode")
'
'        End If
'
'
'
'
'
'        If .TextMatrix(.Rows - 1, .ColIndex("Account_Serial")) <> "" Then
'           ' .Rows = .Rows + 1
'            Grd_AfterEdit .Rows - 1, .ColIndex("Account_Serial")
'
'        End If
'             If Val(.TextMatrix(.Rows - 1, .ColIndex("value"))) <> 0 Then
'           ' .Rows = .Rows + 1
'            Grd_AfterEdit .Rows - 1, .ColIndex("value")
'
'        End If
'        If Trim(.TextMatrix(.Rows - 1, .ColIndex("Account_Serial"))) = "" Then
'            .Rows = .Rows - 1
'        End If
'If SystemOptions.UserInterface = ArabicInterface Then
'             fg.TextMatrix(.Rows - 1, fg.ColIndex("UnitID")) = IIf(IsNull(rs2("UnitName")), "", (rs2("UnitName").value))
'Else
'    fg.TextMatrix(.Rows - 1, fg.ColIndex("UnitID")) = IIf(IsNull(rs2("UnitNamee")), "", (rs2("UnitNamee").value))
'End If

     End With
    '      Me.TxtItemCodeB.Text = ""
     
    '\      Unload FrmItemSearch2
     ' Me.TxtItemCodeB.SetFocus
         
    Else
         
    End If
    
    Else
           error_string = error_string & Trim(Fullcode) & "," & Qty & "," & Name & vbCrLf

End If
'End If

End Function

Private Sub Command3_Click()
'ExportToExcel Me, Grd, "TT", , "grdItems"
tmpGrd.Rows = 1

If mIndex = 0 Then
    Grd.Rows = 1
    FromExcel Grd, tmpGrd, Me, , , txtFile.Text, "TblEmployee"
ElseIf mIndex = 1 Then
    grdMan.Rows = 1
    FromExcel grdMan, tmpGrd, Me, , , txtFile.Text, "TblCustemers"
ElseIf mIndex = 2 Then
    grdGroups.Rows = 1
    FromExcel grdGroups, tmpGrd, Me, , , txtFile.Text, "Groups"
ElseIf mIndex = 3 Then
    grdUnits.Rows = 1
    FromExcel grdUnits, tmpGrd, Me, , , txtFile.Text, "TblUnites"
 ElseIf mIndex = 4 Then
    grdItems.Rows = 1
    FromExcel grdItems, tmpGrd, Me, , , txtFile.Text, "TBLITEMS"
       
 ElseIf mIndex = 5 Then
    grdGroups2.Rows = 1
    FromExcel grdGroups2, tmpGrd, Me, , , txtFile.Text, "GroupsCustomers"
       
       
End If
cmdSave.Enabled = True


Dim i As Long
Dim j As Long
Dim mJob1 As Long
Dim mJobName1 As String
Dim mJob2 As Long
Dim mJobName2 As String
Dim mJob3 As Long
Dim mJobName3 As String

'For i = 0 To Grd.Cols - 1
'    If Grd.ColEditMask(i) <> "" Then
'        Grd.ColHidden(i) = False
'    End If
'    'Grd.ColComboList(i) = ""
'Next
If mIndex = 0 Then
    For i = 1 To Grd.Rows - 1
        For j = 1 To Grd.Cols - 1
            Select Case Grd.ColKey(j)
            Case "JobTypeID"
                mJob1 = Val(Grd.TextMatrix(i, j))
                mJobName1 = Trim(Grd.TextMatrix(i, (j - 1)))
            Case "JobTypeID3"
                mJob2 = Val(Grd.TextMatrix(i, (j)))
                
                If mJob2 = 0 Then
                    mJob2 = mJob1
                    mJobName2 = mJobName1
                    Grd.TextMatrix(i, (j)) = mJob2
                    Grd.TextMatrix(i, (j - 1)) = mJobName2
                    
                End If
            Case "JobTypeID2"
                mJob3 = Val(Grd.TextMatrix(i, (j)))
                If mJob3 = 0 Then
                
                    mJob3 = mJob2
                    mJobName3 = mJobName2
                    Grd.TextMatrix(i, (j)) = mJob3
                    Grd.TextMatrix(i, (j - 1)) = mJobName3
                End If
            Case ""
            Case ""
            End Select
        Next j
    Next
End If

If mIndex = 1 Then
   For i = 1 To grdMan.Rows - 1
     grdMan.TextMatrix(i, grdMan.ColIndex("Type")) = IIf(Option2.Value = True, 1, 2)
   Next
End If

End Sub

Private Sub Command4_Click()
     Dim rs As ADODB.Recordset
     Dim RSTransDetails As ADODB.Recordset
        StrSQL = "Select * From Transactions where Transaction_Type=30"
    StrSQL = StrSQL & "  AND     1 = -1"
    
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    
         Cn.BeginTrans
        BegineTrans = True
     XPTxtBillID.Text = CStr(new_id("Transactions", "Transaction_ID", "", True))
            Me.TxtTransSerial.Text = CStr(new_id("Transactions", "Transaction_Serial", "", True, "Transaction_Type=30"))
            rs.AddNew
            rs("Transaction_ID").Value = Val(XPTxtBillID.Text)
            
        StrSQL = "SELECT     dbo.Transaction_Details.* from dbo.Transaction_Details Where (Transaction_ID = -1)"
        Set RSTransDetails = New ADODB.Recordset
   RSTransDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
        
        
        rs("BranchId").Value = 1
        rs("opening_balance_voucher_id").Value = get_opening_balance_voucher_id
        rs("Transaction_Serial").Value = Me.TxtTransSerial.Text
        rs("Transaction_Date").Value = XPDtbBill.Value
    
        rs("GardFromDate").Value = DTPickerAccFrom.Value
        rs("GardTodate").Value = DTPickerAccTo.Value
        rs("GardEntryType").Value = 0
         rs("Transaction_Type").Value = 30
        rs("UserID").Value = 1
        rs("StoreID").Value = IIf(DCboStoreName.BoundText = "", Null, DCboStoreName.BoundText)
        rs.Update
        
            For RowNum = 1 To grdItems.Rows - 1

            If grdItems.TextMatrix(RowNum, grdItems.ColIndex("Code")) <> "" Then
                RSTransDetails.AddNew
                RSTransDetails("Transaction_ID").Value = XPTxtBillID.Text
                
                RSTransDetails("AutoDetect").Value = 0
                RSTransDetails("Item_ID").Value = IIf((grdItems.TextMatrix(RowNum, grdItems.ColIndex("ItemID")) = ""), Null, Val(grdItems.TextMatrix(RowNum, grdItems.ColIndex("ItemID"))))
                RSTransDetails("Quantity").Value = IIf((grdItems.TextMatrix(RowNum, grdItems.ColIndex("TotalQty")) = ""), Null, Val(grdItems.TextMatrix(RowNum, grdItems.ColIndex("TotalQty"))))

               
                'RSTransDetails("ParrtNoCode").Value = IIf((grdItems.TextMatrix(RowNum, grdItems.ColIndex("ParrtNoCode")) = ""), Null, (grdItems.TextMatrix(RowNum, grdItems.ColIndex("ParrtNoCode"))))
                '    RSTransDetails("ItemDetailedCode").Value = IIf((grdItems.TextMatrix(RowNum, grdItems.ColIndex("ItemDetailedCode")) = ""), Null, (grdItems.TextMatrix(RowNum, grdItems.ColIndex("ItemDetailedCode"))))
'
                'RSTransDetails("ItemCase").Value = IIf((grdItems.TextMatrix(RowNum, grdItems.ColIndex("ItemCase")) = ""), Null, Val(grdItems.TextMatrix(RowNum, grdItems.ColIndex("ItemCase"))))
                RSTransDetails("Price").Value = Val(grdItems.TextMatrix(RowNum, grdItems.ColIndex("UnitSalesPrice")))
            
'                RSTransDetails("ColorID").Value = IIf((grdItems.TextMatrix(RowNum, grdItems.ColIndex("ColorID")) = ""), 1, Val(grdItems.TextMatrix(RowNum, grdItems.ColIndex("ColorID"))))
'
'                RSTransDetails("ItemSize").Value = IIf((grdItems.TextMatrix(RowNum, grdItems.ColIndex("ItemSize")) = ""), 1, Val(grdItems.TextMatrix(RowNum, grdItems.ColIndex("ItemSize"))))
'
'                RSTransDetails("ClassId").Value = IIf((grdItems.TextMatrix(RowNum, grdItems.ColIndex("ClassId")) = ""), 1, Val(grdItems.TextMatrix(RowNum, grdItems.ColIndex("ClassId"))))
            
                RSTransDetails("BranchId").Value = 1
                ' IIf((grdItems.TextMatrix(RowNum, grdItems.ColIndex("BranchId")) = ""), 1, Val(grdItems.TextMatrix(RowNum, grdItems.ColIndex("BranchId"))))
               
                ' RSTransDetails("ItemSize").value = _
                  IIf((grdItems.TextMatrix(RowNum, grdItems.ColIndex("ItemSize")) = ""), "", Trim$(grdItems.TextMatrix(RowNum, grdItems.ColIndex("ItemSize"))))
                'RSTransDetails("LotNO").Value = IIf(grdItems.TextMatrix(RowNum, grdItems.ColIndex("LotNO")) = "", Null, grdItems.TextMatrix(RowNum, grdItems.ColIndex("LotNO")))
              
                RSTransDetails("UnitID").Value = IIf((grdItems.TextMatrix(RowNum, grdItems.ColIndex("UnitID")) = ""), 1, Val(grdItems.TextMatrix(RowNum, grdItems.ColIndex("UnitID"))))
                RSTransDetails("ShowQty").Value = IIf((grdItems.TextMatrix(RowNum, grdItems.ColIndex("TotalQty")) = ""), Null, Val(grdItems.TextMatrix(RowNum, grdItems.ColIndex("TotalQty"))))

                Dim RsUnitData As ADODB.Recordset
                Dim LngCurItemID As Long
                Dim LngUnitID As Long
                Dim DblQty As Double
        
                LngCurItemID = Val(grdItems.TextMatrix(RowNum, grdItems.ColIndex("ItemID")))
                LngUnitID = Val(grdItems.TextMatrix(RowNum, grdItems.ColIndex("UnitID")))
                DblQty = Val(grdItems.TextMatrix(RowNum, grdItems.ColIndex("TotalQty")))

                StrSQL = "Select * From TblItemsUnits Where ItemID=" & LngCurItemID
                StrSQL = StrSQL + " AND UnitID=" & LngUnitID
                Set RsUnitData = New ADODB.Recordset
                RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If Not (RsUnitData.BOF Or RsUnitData.EOF) Then
                    RSTransDetails("QtyBySmalltUnit").Value = RsUnitData("UnitFactor").Value
                    RSTransDetails("Quantity").Value = RSTransDetails("QtyBySmalltUnit").Value * RSTransDetails("showqty").Value
                End If

                RSTransDetails("Price").Value = Val(grdItems.TextMatrix(RowNum, grdItems.ColIndex("UnitSalesPrice")))
                              
                RSTransDetails("showprice").Value = Val(grdItems.TextMatrix(RowNum, grdItems.ColIndex("UnitPurPrice")))
                RSTransDetails.Update
            End If

        Next RowNum
        
        Cn.CommitTrans
        BegineTrans = False
        
        MsgBox " „ «·Ã—œ"
        Command4.Enabled = False
End Sub
Public Function get_opening_balance_voucher_id() As Double
  Dim newSeril As Double
    On Error Resume Next
 
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String

'    sql = "select max(opening_balance_voucher_id) As id from DOUBLE_ENTREY_VOUCHERS1"
 
'    rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
'    If rs3.RecordCount > 0 Then
'        get_opening_balance_voucher_id = IIf(IsNull(rs3("id").value), 0, rs3("id").value) + 1
'
'    Else
'        get_opening_balance_voucher_id = 1
'    End If
    Dim LngDevID As Long
'LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS1", "Double_Entry_Vouchers_ID", "", True)

'  Cn.Execute "insert into  DOUBLE_ENTREY_VOUCHERS1 (opening_balance_voucher_id,DEV_ID_Line_No,Double_Entry_Vouchers_ID)  values (" & get_opening_balance_voucher_id & ",0," & LngDevID & ")"
get_opening_balance_voucher_id = MyTime
End Function

Private Sub Command5_Click()
Dim s As String

Dim rsDummy As New ADODB.Recordset
Dim rsDummyData As New ADODB.Recordset
Dim NoteID As Long
Dim EntryID As Long
s = "Select Max(Notes_Id) NoteID,Max(Double_Entry_Vouchers_ID) as EntryID from DOUBLE_ENTREY_VOUCHERS1 "
rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly
If Not rsDummy.EOF Then
    EntryID = Val(rsDummy!EntryID & "")
    NoteID = Val(rsDummy!NoteID & "")
End If
rsDummy.Close
Set rsDummyData = New ADODB.Recordset
s = "Select * from DOUBLE_ENTREY_VOUCHERS1 Where Notes_Id = " & NoteID
rsDummyData.Open s, Cn, adOpenKeyset, adLockReadOnly
If Not rsDummyData.EOF Then
'    EntryID = Val(rsDummy!EntryID & "")
'    NoteID = Val(rsDummy!NoteID & "")
End If
'rsDummy.Close

s = " SELECT * FROM ("
s = s & " SELECT TOP 100 PERCENT                Account_Code,"
s = s & "        Account_Name,"
s = s & "        last_account,"
s = s & "        Account_NameEng,"
s = s & "        ISNULL(CreditBalance, 0)    AS CreditBalance,"
s = s & "        ISNULL(DepitBalance, 0)     AS DepitBalance,"
s = s & "        ISNULL(opening_balance, 0)  AS opening_balance,"
s = s & "        ISNULL(Balance, 0)          AS Balance,"
s = s & "        ISNULL(opening_balance, 0)+ ISNULL(DepitBalance, 0)  + ISNULL(CreditBalance, 0) as balance2,"
s = s & "        Account_Serial,"
s = s & "        Parent_Account_Code"
s = s & " From " & Trim(txtBDNAMe) & ".dbo.ACCOUNTS"
s = s & " Where 1 = 1"
s = s & "        AND (last_account = 1)"
s = s & "        AND NOT ("
s = s & "                opening_balance = 0"
s = s & "                AND DepitBalance = 0"
s = s & "                AND CreditBalance = 0"
s = s & "            )"
s = s & "        AND ("
s = s & "                ACCOUNTS.Account_Code IN (SELECT TblAccountBranch.Account_Code"
s = s & "                                          From " & Trim(txtBDNAMe) & ".dbo.TblAccountBranch"
s = s & "                                          WHERE  TblAccountBranch.BranchID  IN (SELECT BranchID"
s = s & "                                                                                From " & Trim(txtBDNAMe) & ".dbo.TblUsersBranches"
s = s & "                                                                                WHERE  (UserID = 1))"
s = s & "                                                 AND ("
s = s & "                                                         ACCOUNTS.Account_Code IN (SELECT TblAccountUser.Account_Code"
s = s & "                                                                                   From " & Trim(txtBDNAMe) & ".dbo.TblAccountUser"
s = s & "                                                                                   WHERE  TblAccountUser.UserID = 1)"
s = s & "                                                         OR ACCOUNTS.Account_Code NOT IN (SELECT"
s = s & "                                                                                                 TblAccountUser.Account_Code"
s = s & "                                                                                          FROM   " & Trim(txtBDNAMe) & ".dbo.TblAccountUser)"
s = s & "                                                     ))"
s = s & "                OR ("
s = s & "                       ACCOUNTS.Account_Code NOT IN (SELECT TblAccountBranch.Account_Code"
s = s & "                                                     FROM   " & Trim(txtBDNAMe) & ".dbo.TblAccountBranch)"
s = s & "                       AND ACCOUNTS.Account_Code NOT IN (SELECT TblAccountUser.Account_Code"
s = s & "                                                         FROM   " & Trim(txtBDNAMe) & ".dbo.TblAccountUser)"
s = s & "                   )"
s = s & "            )"
s = s & " Order By"
s = s & "        Account_Serial) T Where (opening_balance) + (DepitBalance) + (CreditBalance) <> 0"
'Where balance2 <> 0"

Text1 = s

rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly


s = "Select * from DOUBLE_ENTREY_VOUCHERS1 Where Double_Entry_Vouchers_ID = -5"
Dim rsData As New ADODB.Recordset
rsData.Open s, Cn, adOpenKeyset, adLockOptimistic
Dim mLine As Integer
mLine = 0
Do While Not rsDummy.EOF
If Abs(Val(rsDummy!balance2 & "")) = 0 Then GoTo MoveNext1
    rsData.AddNew
    rsData!Double_Entry_Vouchers_ID = EntryID + 1
    If mLine = 0 Or mLine = 2 Then mLine = 1 Else mLine = 2
    rsData!DEV_ID_Line_No = mLine
    rsData!Account_Code = rsDummy!Account_Code
    rsData!Value = Abs(Val(rsDummy!balance2 & ""))
    If Val(rsDummy!balance2 & "") > 0 Then
        rsData!Credit_Or_Debit = 0
    Else
        rsData!Credit_Or_Debit = 1
    End If
    'rsData!Credit_Or_Debit = 0
    rsData!branch_id = rsDummyData!branch_id
    rsData!RecordDate = rsDummyData!RecordDate
    rsData!Notes_ID = rsDummyData!Notes_ID
    rsData!UserID = rsDummyData!UserID
    rsData!Account_Interval_ID = rsDummyData!Account_Interval_ID
    rsData!DEV_Serial = rsDummyData!DEV_Serial
    rsData!Rate = rsDummyData!Rate
    rsData!Notes_ID = rsDummyData!Notes_ID
    rsData!opening_balance_voucher_id = Val(rsData!opening_balance_voucher_id & "") + 1
    rsData!DEV_ID_Line_No1 = rsData!DEV_ID_Line_No1
    rsData!Remarks2 = 3
    rsData.Update
MoveNext1:
    rsDummy.MoveNext
Loop

'
's = " SELECT * FROM ("
's = s & " SELECT TOP 100 PERCENT                Account_Code,"
's = s & "        Account_Name,"
's = s & "        last_account,"
's = s & "        Account_NameEng,"
's = s & "        ISNULL(CreditBalance, 0)    AS CreditBalance,"
's = s & "        ISNULL(DepitBalance, 0)     AS DepitBalance,"
's = s & "        ISNULL(opening_balance, 0)  AS opening_balance,"
's = s & "        ISNULL(Balance, 0)          AS Balance,"
's = s & "        ISNULL(opening_balance, 0)+ ISNULL(DepitBalance, 0)  + ISNULL(CreditBalance, 0) as balance2,"
's = s & "        Account_Serial,"
's = s & "        Parent_Account_Code"
's = s & " From " & Trim(txtBDNAMe) & ".dbo.ACCOUNTS"
's = s & " Where 1 = 1"
's = s & "        AND (last_account = 1)"
's = s & "        AND NOT ("
's = s & "                opening_balance = 0"
's = s & "                AND DepitBalance = 0"
's = s & "                AND CreditBalance = 0"
's = s & "            )"
's = s & "        AND ("
's = s & "                ACCOUNTS.Account_Code IN (SELECT TblAccountBranch.Account_Code"
's = s & "                                          From dbo.TblAccountBranch"
's = s & "                                          WHERE  TblAccountBranch.BranchID  IN (SELECT BranchID"
's = s & "                                                                                From " & Trim(txtBDNAMe) & ".dbo.TblUsersBranches"
's = s & "                                                                                WHERE  (UserID = 1))"
's = s & "                                                 AND ("
's = s & "                                                         ACCOUNTS.Account_Code IN (SELECT TblAccountUser.Account_Code"
's = s & "                                                                                   From " & Trim(txtBDNAMe) & ".dbo.TblAccountUser"
's = s & "                                                                                   WHERE  TblAccountUser.UserID = 1)"
's = s & "                                                         OR ACCOUNTS.Account_Code NOT IN (SELECT"
's = s & "                                                                                                 TblAccountUser.Account_Code"
's = s & "                                                                                          FROM   " & Trim(txtBDNAMe) & ".dbo.TblAccountUser)"
's = s & "                                                     ))"
's = s & "                OR ("
's = s & "                       ACCOUNTS.Account_Code NOT IN (SELECT TblAccountBranch.Account_Code"
's = s & "                                                     FROM   " & Trim(txtBDNAMe) & ".dbo.TblAccountBranch)"
's = s & "                       AND ACCOUNTS.Account_Code NOT IN (SELECT TblAccountUser.Account_Code"
's = s & "                                                         FROM   " & Trim(txtBDNAMe) & ".dbo.TblAccountUser)"
's = s & "                   )"
's = s & "            )"
's = s & " Order By"
's = s & "        Account_Serial) T Where (opening_balance) + (DepitBalance) + (CreditBalance) < 0"
''Where balance2 <> 0"
'
'Text2 = s
'Set rsDummy = New ADODB.Recordset
'rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly
'
'
's = "Select * from DOUBLE_ENTREY_VOUCHERS1 Where Double_Entry_Vouchers_ID = -5"
'Set rsData = New ADODB.Recordset
'rsData.Open s, Cn, adOpenKeyset, adLockOptimistic
'
''mLine = 0
'Do While Not rsDummy.EOF
'    If Abs(Val(rsDummy!balance2 & "")) = 0 Then GoTo MoveNext2
'    rsData.AddNew
'    rsData!Double_Entry_Vouchers_ID = EntryID + 1
'    If mLine = 0 Or mLine = 2 Then mLine = 1 Else mLine = 2
'    rsData!DEV_ID_Line_No = mLine
'    rsData!Account_Code = rsDummy!Account_Code
'    rsData!Value = Abs(Val(rsDummy!balance2 & ""))
''    If Val(rsDummy!balance2 & "") > 0 Then
''        rsData!Credit_Or_Debit = 0
''    Else
''        rsData!Credit_Or_Debit = 1
''    End If
'    rsData!Credit_Or_Debit = 1
'    rsData!branch_id = rsDummyData!branch_id
'    rsData!RecordDate = rsDummyData!RecordDate
'    rsData!Notes_ID = rsDummyData!Notes_ID
'    rsData!UserID = rsDummyData!UserID
'    rsData!Account_Interval_ID = rsDummyData!Account_Interval_ID
'    rsData!DEV_Serial = rsDummyData!DEV_Serial
'    rsData!Rate = rsDummyData!Rate
'    rsData!Notes_ID = rsDummyData!Notes_ID
'    rsData!opening_balance_voucher_id = Val(rsData!opening_balance_voucher_id & "") + 1
'    rsData!DEV_ID_Line_No1 = rsData!DEV_ID_Line_No1
'    rsData!Remarks2 = 3
'    rsData.Update
'MoveNext2:
'    rsDummy.MoveNext
'Loop


MsgBox " „"
End Sub

Private Sub Command6_Click()
    MsgBox Weekday(txtDate, 0)
    
    MsgBox WeekdayName(Weekday(txtDate, 0))
    MsgBox vbWednesday
End Sub

Private Sub Form_Load()
txtDbPath = GetSetting("ConvertToAccess", "Setting", "DbPath", "DatabasePath")
TxtTableName = GetSetting("ConvertToAccess", "Setting", "TableName", "TableName")
TxtUSERID = GetSetting("ConvertToAccess", "Setting", "USERID", "USERID")
TxtCHECKTIME = GetSetting("ConvertToAccess", "Setting", "CHECKTIME", "CHECKTIME")
'DcTime.Value = GetSetting("ConvertToAccess", "Setting", "UpdateHours", "00")
dbRecordDate = Date
TxtServerDataBaseName = SysSQLServerDataBaseName
DestinationServer = SysSQLServerName
ServerDb = TxtServerDataBaseName.Text
ConnectionFirst
XPDtbBill = Date
DTPickerAccTo = Date
mIndex = TabMain.CurrTab
'BranchDigit = 1
Dim Msg As String
If Dir(App.Path & "\pos.txt", vbNormal) = "" Then
            Msg = "„·›  ”ÃÌ· «·ﬁÊ«⁄œ €Ì— „ÊÃÊœ ...!!!"
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            
           End
           
        End If
        
    Open App.Path & "\pos.txt" For Input As #1
    
   Dim StrSQL As String

    
 

        StrSQL = "SELECT StoreID,StoreName From TblStore where 1=1"
 




    GetComboData DCboStoreName, StrSQL
    
    cmdSave.Enabled = False
End Sub
Private Sub GetComboData(My_Combo As DataCombo, _
                         My_SQL As String)
    Dim rs As ADODB.Recordset
    Dim StrTemp As String
    Dim Msg As String
    On Error GoTo ErrorHandler

    If InStr(1, My_SQL, "SELECT", vbTextCompare) = 0 Then
        Exit Sub
    End If

    My_Combo.Tag = My_SQL
    Set rs = New ADODB.Recordset

    
        rs.CursorLocation = adUseClient
   

    rs.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    'Populate the ADO datacombo by setting its properties
    With My_Combo
        StrTemp = .BoundText
        Set .RowSource = rs
        .BoundColumn = rs(0).Name
        .ListField = rs(1).Name

        If Trim(StrTemp) <> "" Then
            .BoundText = StrTemp
        Else
            .BoundText = ""
            .Text = ""
        End If

    End With

Exit_Sub:
    Set rs = Nothing
    Exit Sub
ErrorHandler:

    'MsgBox "ERROR! Err# " & Err.Number & " Desc: " & Err.Description, vbCritical + vbOKOnly
    Resume Exit_Sub
End Sub

Private Sub TabMain_Click()
mIndex = TabMain.CurrTab
End Sub
