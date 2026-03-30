VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmImport 
   Caption         =   "«” Ì—«œ „‰ «·«ﬂ”·"
   ClientHeight    =   9465
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16305
   LinkTopic       =   "Form1"
   ScaleHeight     =   9465
   ScaleWidth      =   16305
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkIsRepeatName 
      Caption         =   " ﬂ—«— «·«”„"
      Height          =   195
      Left            =   3750
      TabIndex        =   171
      Top             =   990
      Width           =   2040
   End
   Begin VB.CheckBox chkIsRepeatCode 
      Caption         =   " ﬂ—«— «·ﬂÊœ"
      Height          =   195
      Left            =   6060
      TabIndex        =   169
      Top             =   930
      Width           =   2040
   End
   Begin VB.CommandButton Command14 
      Caption         =   " ’œÌ— „·› «·«ﬂ”Ì· «·–Ï ”Ì” Œœ„ "
      Height          =   285
      Left            =   6120
      TabIndex        =   147
      Top             =   630
      Width           =   2625
   End
   Begin VB.TextBox txtTableName 
      Height          =   405
      Left            =   9240
      TabIndex        =   143
      Text            =   "TblCustemers"
      Top             =   630
      Visible         =   0   'False
      Width           =   2745
   End
   Begin VB.Frame Frame3 
      Caption         =   "Ã·» »Ì«‰«  „‰ «·Õ”«»«  «·Ï «·„·›«  «·«”«”Ì…"
      Height          =   1275
      Left            =   150
      TabIndex        =   84
      Top             =   1140
      Width           =   10065
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   3840
         TabIndex        =   173
         Top             =   1020
         Width           =   1485
      End
      Begin VB.OptionButton optGetFromAcc 
         Alignment       =   1  'Right Justify
         Caption         =   "«·„Ê—œÌ‰"
         Height          =   195
         Index           =   2
         Left            =   4110
         RightToLeft     =   -1  'True
         TabIndex        =   96
         Top             =   480
         Width           =   885
      End
      Begin VB.OptionButton optGetFromAcc 
         Alignment       =   1  'Right Justify
         Caption         =   "«·„Œ«“‰"
         Height          =   195
         Index           =   6
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   95
         Top             =   480
         Width           =   855
      End
      Begin VB.OptionButton optGetFromAcc 
         Alignment       =   1  'Right Justify
         Caption         =   "«·„’—Ê›« "
         Height          =   195
         Index           =   5
         Left            =   1050
         RightToLeft     =   -1  'True
         TabIndex        =   94
         Top             =   480
         Width           =   1035
      End
      Begin VB.OptionButton optGetFromAcc 
         Alignment       =   1  'Right Justify
         Caption         =   "«·⁄Âœ Ê«·Œ“‰"
         Height          =   195
         Index           =   4
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   93
         Top             =   480
         Width           =   1155
      End
      Begin VB.OptionButton optGetFromAcc 
         Alignment       =   1  'Right Justify
         Caption         =   "«·»‰Êﬂ"
         Height          =   195
         Index           =   3
         Left            =   3360
         RightToLeft     =   -1  'True
         TabIndex        =   92
         Top             =   480
         Width           =   735
      End
      Begin VB.OptionButton optGetFromAcc 
         Alignment       =   1  'Right Justify
         Caption         =   "«·⁄„·«¡"
         Height          =   195
         Index           =   1
         Left            =   4980
         RightToLeft     =   -1  'True
         TabIndex        =   91
         Top             =   480
         Width           =   795
      End
      Begin VB.OptionButton optGetFromAcc 
         Alignment       =   1  'Right Justify
         Caption         =   "«·„ÊŸ›Ì‰"
         Height          =   195
         Index           =   0
         Left            =   5820
         RightToLeft     =   -1  'True
         TabIndex        =   90
         Top             =   480
         Width           =   1035
      End
      Begin VB.TextBox txtKeySerach 
         Height          =   315
         Left            =   390
         TabIndex        =   87
         Text            =   "«·»‰Êﬂ"
         Top             =   750
         Width           =   5865
      End
      Begin VB.CommandButton cmdFromAccount 
         Caption         =   "Ã·» „‰ «·Õ”«»« "
         Height          =   315
         Left            =   8550
         TabIndex        =   85
         Top             =   480
         Width           =   1395
      End
      Begin MSDataListLib.DataCombo DboParentAccount2 
         Height          =   315
         Left            =   330
         TabIndex        =   172
         Top             =   1020
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label lblCapt 
         Height          =   345
         Left            =   150
         TabIndex        =   89
         Top             =   960
         Width           =   6345
      End
      Begin VB.Label Label3 
         Caption         =   "ﬂ·„… „› «ÕÌ…"
         Height          =   285
         Left            =   6360
         TabIndex        =   88
         Top             =   780
         Width           =   945
      End
      Begin VB.Label Label2 
         Caption         =   "ﬁ„ »«Œ Ì«— «· «» „‰ «”›· «Ê·« · ÕœÌœ ‰Ê⁄ «·»Ì«‰«  «·„—«œ ‰ﬁ·Â« „‰ «·Õ”«»« "
         Height          =   255
         Left            =   4920
         TabIndex        =   86
         Top             =   150
         Width           =   5085
      End
   End
   Begin VB.CommandButton Command8 
      Caption         =   "«‰‘«¡ «·Õ”«»«  «·‰Ÿ«„Ì…"
      Height          =   285
      Left            =   10440
      TabIndex        =   78
      Top             =   1560
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Height          =   405
      Left            =   10440
      TabIndex        =   57
      Text            =   "Text2"
      Top             =   1110
      Visible         =   0   'False
      Width           =   2745
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   10320
      TabIndex        =   56
      Text            =   "Text1"
      Top             =   2100
      Visible         =   0   'False
      Width           =   2205
   End
   Begin VB.TextBox txtBDNAMe 
      Height          =   435
      Left            =   9990
      TabIndex        =   55
      Text            =   "Byte"
      Top             =   150
      Width           =   1395
   End
   Begin VB.CommandButton Command5 
      Caption         =   "«‰‘«¡ «·ﬁÌœ «·«›  «ÕÏ Õ«·… œŒÊ· «·»Ì«‰«  ·ÃœÊ· «·Õ”«»« "
      Height          =   405
      Left            =   13320
      TabIndex        =   54
      Top             =   1590
      Visible         =   0   'False
      Width           =   2835
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
      Height          =   1005
      Left            =   240
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
         Top             =   570
         Width           =   1815
      End
      Begin VB.TextBox DestinationServer 
         Height          =   375
         Left            =   1200
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   210
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
         Top             =   210
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
         Top             =   690
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
      Left            =   -120
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   " ÕœÌœ «·„·›..."
      Height          =   255
      Left            =   6990
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   300
      Width           =   1305
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
      Left            =   13950
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
      Cols            =   40
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
      Height          =   6765
      Left            =   120
      TabIndex        =   11
      Top             =   2520
      Width           =   16110
      _cx             =   28416
      _cy             =   11933
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
      Caption         =   $"frmImport.frx":0000
      Align           =   0
      CurrTab         =   13
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
         Height          =   6390
         Index           =   1
         Left            =   -20265
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   45
         Width           =   16020
         _cx             =   28258
         _cy             =   11271
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
            Height          =   6300
            Index           =   0
            Left            =   22485
            TabIndex        =   13
            Top             =   555
            Width           =   15720
            _cx             =   27728
            _cy             =   11112
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
            FormatString    =   $"frmImport.frx":00E4
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
            Height          =   6240
            Left            =   0
            TabIndex        =   16
            Top             =   120
            Width           =   15405
            _cx             =   27173
            _cy             =   11007
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
            Cols            =   94
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmImport.frx":01A4
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
         Height          =   6390
         Index           =   0
         Left            =   -19965
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   45
         Width           =   16020
         _cx             =   28258
         _cy             =   11271
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
         Begin VB.CheckBox Check1 
            Caption         =   "⁄„·«¡ »œÊ‰ Õ”«»«  ›ﬁÿ"
            Height          =   195
            Left            =   2775
            TabIndex        =   102
            Top             =   360
            Width           =   2160
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Õ–› «·„ÊÃÊœ"
            Height          =   195
            Left            =   1230
            TabIndex        =   101
            Top             =   360
            Width           =   1230
         End
         Begin VB.CheckBox Check3 
            Caption         =   "Õ–› «·„ÊÃÊœ"
            Height          =   210
            Left            =   1230
            TabIndex        =   100
            Top             =   720
            Width           =   1230
         End
         Begin VB.CheckBox Check6 
            Caption         =   "Õ–› «·«»"
            Height          =   195
            Left            =   1230
            TabIndex        =   99
            Top             =   915
            Width           =   1230
         End
         Begin VB.CheckBox Check8 
            Caption         =   "Õ–› «·«»"
            Height          =   180
            Left            =   1230
            TabIndex        =   98
            Top             =   555
            Width           =   1230
         End
         Begin VB.CommandButton CmdRecalcAccountSupp 
            Caption         =   "«⁄«œ… «‰‘«¡  Õ”«»«  «·„Ê—œÌ‰"
            Height          =   285
            Left            =   315
            TabIndex        =   97
            Top             =   45
            Width           =   2760
         End
         Begin VB.CheckBox chkBalanceOnly 
            Caption         =   "«—’œ… ›ﬁÿ"
            Height          =   210
            Left            =   10785
            TabIndex        =   81
            Top             =   45
            Width           =   2160
         End
         Begin VB.CommandButton Command9 
            Caption         =   "Ã·» «·«—’œ… ›ﬁÿ"
            Height          =   225
            Left            =   8625
            TabIndex        =   79
            Top             =   30
            Width           =   1845
         End
         Begin VB.OptionButton Option1 
            Caption         =   "„Ê—œÌ‰"
            Height          =   255
            Left            =   3075
            TabIndex        =   31
            Top             =   0
            Width           =   2160
         End
         Begin VB.OptionButton Option2 
            Caption         =   "⁄„·«¡"
            Height          =   240
            Left            =   5235
            TabIndex        =   30
            Top             =   45
            Value           =   -1  'True
            Width           =   1230
         End
         Begin VSFlex8UCtl.VSFlexGrid FgItems 
            Height          =   6300
            Index           =   1
            Left            =   22485
            TabIndex        =   15
            Top             =   555
            Width           =   15720
            _cx             =   27728
            _cy             =   11112
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
            FormatString    =   $"frmImport.frx":12CB
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
            Height          =   4770
            Left            =   270
            TabIndex        =   24
            Top             =   1380
            Width           =   15390
            _cx             =   27146
            _cy             =   8414
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
            Cols            =   36
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmImport.frx":138B
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
         Begin VSFlex8Ctl.VSFlexGrid grdManBal 
            Height          =   4920
            Left            =   315
            TabIndex        =   80
            Top             =   1320
            Visible         =   0   'False
            Width           =   15390
            _cx             =   27146
            _cy             =   8678
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
            Cols            =   7
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
         Height          =   6390
         Index           =   2
         Left            =   -19665
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   45
         Width           =   16020
         _cx             =   28258
         _cy             =   11271
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
            Height          =   6300
            Index           =   2
            Left            =   22485
            TabIndex        =   18
            Top             =   555
            Width           =   15720
            _cx             =   27728
            _cy             =   11112
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
            FormatString    =   $"frmImport.frx":1AF0
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
            Height          =   6240
            Left            =   315
            TabIndex        =   23
            Top             =   90
            Width           =   15390
            _cx             =   27146
            _cy             =   11007
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
            Cols            =   11
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmImport.frx":1BB0
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
         Height          =   6390
         Index           =   3
         Left            =   -19365
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   45
         Width           =   16020
         _cx             =   28258
         _cy             =   11271
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
            Height          =   6300
            Index           =   3
            Left            =   22485
            TabIndex        =   20
            Top             =   555
            Width           =   15720
            _cx             =   27728
            _cy             =   11112
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
            FormatString    =   $"frmImport.frx":1D8E
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
            Height          =   6240
            Left            =   0
            TabIndex        =   25
            Top             =   0
            Width           =   15405
            _cx             =   27173
            _cy             =   11007
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
            FormatString    =   $"frmImport.frx":1E4E
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
         Height          =   6390
         Index           =   4
         Left            =   -19065
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   45
         Width           =   16020
         _cx             =   28258
         _cy             =   11271
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
         Begin VB.Frame fra 
            BackColor       =   &H00E2E9E9&
            Caption         =   "›Ì «·› —…"
            ForeColor       =   &H00FF0000&
            Height          =   1125
            Index           =   63
            Left            =   4620
            RightToLeft     =   -1  'True
            TabIndex        =   163
            Top             =   660
            Width           =   2460
            Begin MSComCtl2.DTPicker DTPFrom 
               Height          =   345
               Left            =   120
               TabIndex        =   164
               Top             =   240
               Width           =   1665
               _ExtentX        =   2937
               _ExtentY        =   609
               _Version        =   393216
               CheckBox        =   -1  'True
               CustomFormat    =   "dd/m/yyyy"
               DateIsNull      =   -1  'True
               Format          =   151912449
               CurrentDate     =   36494
            End
            Begin MSComCtl2.DTPicker DTPTo 
               Height          =   345
               Left            =   120
               TabIndex        =   165
               Top             =   630
               Width           =   1665
               _ExtentX        =   2937
               _ExtentY        =   609
               _Version        =   393216
               CheckBox        =   -1  'True
               DateIsNull      =   -1  'True
               Format          =   151912449
               CurrentDate     =   38784
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "≈·Ï"
               Height          =   285
               Index           =   23
               Left            =   1830
               RightToLeft     =   -1  'True
               TabIndex        =   167
               Top             =   675
               Width           =   465
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„‰"
               Height          =   285
               Index           =   24
               Left            =   1740
               RightToLeft     =   -1  'True
               TabIndex        =   166
               Top             =   255
               Width           =   525
            End
         End
         Begin VB.CommandButton Command16 
            Caption         =   "«‰‘«¡ —’Ìœ «›  «ÕÏ „‰ œ«Œ· «·»—‰«„Ã"
            Height          =   495
            Left            =   2460
            TabIndex        =   150
            Top             =   960
            Width           =   2160
         End
         Begin VB.Frame Frame5 
            Height          =   345
            Left            =   2460
            TabIndex        =   130
            Top             =   30
            Width           =   3390
            Begin VB.OptionButton optOpenBalanceItem 
               Caption         =   "Ã—œ"
               Height          =   195
               Index           =   1
               Left            =   240
               TabIndex        =   132
               Top             =   180
               Width           =   1095
            End
            Begin VB.OptionButton optOpenBalanceItem 
               Caption         =   "—’Ìœ «›  «ÕÏ"
               Height          =   195
               Index           =   0
               Left            =   2280
               TabIndex        =   131
               Top             =   180
               Value           =   -1  'True
               Width           =   1545
            End
         End
         Begin MSDataListLib.DataCombo DataCombo1 
            Height          =   315
            Left            =   315
            TabIndex        =   59
            Top             =   450
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   556
            _Version        =   393216
            Text            =   "DataCombo1"
         End
         Begin VB.TextBox XPTxtBillID 
            Alignment       =   1  'Right Justify
            Height          =   255
            Left            =   0
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   50
            Top             =   0
            Visible         =   0   'False
            Width           =   2160
         End
         Begin VB.CommandButton Command4 
            Caption         =   "«œŒ«· «·Ã—œ"
            Height          =   375
            Left            =   315
            TabIndex        =   49
            Top             =   1050
            Width           =   1845
         End
         Begin VB.TextBox TxtTransSerial 
            Alignment       =   1  'Right Justify
            Height          =   255
            Left            =   6465
            RightToLeft     =   -1  'True
            TabIndex        =   47
            Top             =   90
            Width           =   315
         End
         Begin VSFlex8UCtl.VSFlexGrid FgItems 
            Height          =   6300
            Index           =   4
            Left            =   22485
            TabIndex        =   22
            Top             =   555
            Width           =   15720
            _cx             =   27728
            _cy             =   11112
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
            FormatString    =   $"frmImport.frx":1F0E
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
            Height          =   4380
            Left            =   0
            TabIndex        =   26
            Top             =   2010
            Width           =   16335
            _cx             =   28813
            _cy             =   7726
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
            Cols            =   24
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmImport.frx":1FCE
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
            Height          =   270
            Left            =   11700
            TabIndex        =   38
            Top             =   285
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   476
            _Version        =   393216
            Format          =   113115139
            CurrentDate     =   38784
         End
         Begin MSDataListLib.DataCombo DCboStoreName 
            Height          =   315
            Left            =   7395
            TabIndex        =   40
            Top             =   660
            Width           =   2160
            _ExtentX        =   3810
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   525
            Index           =   8
            Left            =   8010
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   30
            Width           =   2775
            _cx             =   4895
            _cy             =   926
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
               TabIndex        =   42
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
               Format          =   113115139
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
               TabIndex        =   43
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
               Format          =   113115139
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
               TabIndex        =   45
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
               TabIndex        =   44
               Top             =   765
               Visible         =   0   'False
               Width           =   555
            End
         End
         Begin MSDataListLib.DataCombo DCboItemName10 
            Height          =   315
            Left            =   9555
            TabIndex        =   151
            Top             =   1620
            Width           =   4005
            _ExtentX        =   7064
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCboGroup10 
            Height          =   315
            Left            =   9555
            TabIndex        =   152
            Top             =   1290
            Width           =   4005
            _ExtentX        =   7064
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcStore10 
            Height          =   315
            Left            =   9555
            TabIndex        =   153
            Top             =   930
            Width           =   4005
            _ExtentX        =   7064
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCItemsColors 
            Height          =   315
            Left            =   11700
            TabIndex        =   154
            Top             =   2010
            Width           =   1860
            _ExtentX        =   3281
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCItemsSizes 
            Height          =   315
            Left            =   9555
            TabIndex        =   155
            Top             =   2010
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCItemsClasses 
            Height          =   315
            Left            =   11700
            TabIndex        =   156
            Top             =   2370
            Width           =   1860
            _ExtentX        =   3281
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo dcCustomer 
            Height          =   315
            Left            =   0
            TabIndex        =   168
            Top             =   1830
            Width           =   4005
            _ExtentX        =   7064
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   " «·’‰›"
            Height          =   315
            Index           =   108
            Left            =   13560
            RightToLeft     =   -1  'True
            TabIndex        =   162
            Top             =   1680
            Width           =   615
         End
         Begin VB.Label XPLbl10 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·„Ã„Ê⁄…"
            Height          =   285
            Index           =   4
            Left            =   13560
            RightToLeft     =   -1  'True
            TabIndex        =   161
            Top             =   1290
            Width           =   615
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·„Œ“‰"
            Height          =   315
            Index           =   109
            Left            =   13560
            RightToLeft     =   -1  'True
            TabIndex        =   160
            Top             =   930
            Width           =   615
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«··Ê‰"
            Height          =   315
            Index           =   123
            Left            =   13245
            RightToLeft     =   -1  'True
            TabIndex        =   159
            Top             =   2010
            Width           =   930
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·„ﬁ«”"
            Height          =   315
            Index           =   124
            Left            =   11085
            RightToLeft     =   -1  'True
            TabIndex        =   158
            Top             =   2010
            Width           =   615
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·›—“"
            Height          =   315
            Index           =   125
            Left            =   13245
            RightToLeft     =   -1  'True
            TabIndex        =   157
            Top             =   2370
            Width           =   930
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„”·”·"
            Height          =   285
            Index           =   1
            Left            =   6165
            RightToLeft     =   -1  'True
            TabIndex        =   48
            Top             =   120
            Width           =   1845
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·„Œ“‰"
            Height          =   285
            Index           =   2
            Left            =   9855
            RightToLeft     =   -1  'True
            TabIndex        =   46
            Top             =   645
            Width           =   930
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «·«œŒ«·"
            Height          =   285
            Index           =   0
            Left            =   13245
            RightToLeft     =   -1  'True
            TabIndex        =   39
            Top             =   270
            Width           =   315
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   6390
         Index           =   5
         Left            =   -18765
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   45
         Width           =   16020
         _cx             =   28258
         _cy             =   11271
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
            Height          =   6300
            Index           =   5
            Left            =   22485
            TabIndex        =   28
            Top             =   555
            Width           =   15720
            _cx             =   27728
            _cy             =   11112
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
            FormatString    =   $"frmImport.frx":23F4
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
            Height          =   6240
            Left            =   315
            TabIndex        =   29
            Top             =   90
            Width           =   15390
            _cx             =   27146
            _cy             =   11007
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
            FormatString    =   $"frmImport.frx":24B4
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
         Height          =   6390
         Index           =   6
         Left            =   -18465
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   45
         Width           =   16020
         _cx             =   28258
         _cy             =   11271
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
            Height          =   6300
            Index           =   6
            Left            =   22485
            TabIndex        =   33
            Top             =   555
            Width           =   15720
            _cx             =   27728
            _cy             =   11112
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
            FormatString    =   $"frmImport.frx":2664
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
            Height          =   6240
            Left            =   315
            TabIndex        =   34
            Top             =   90
            Width           =   15390
            _cx             =   27146
            _cy             =   11007
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
            FormatString    =   $"frmImport.frx":2724
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
         Height          =   6390
         Index           =   7
         Left            =   -18165
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   45
         Width           =   16020
         _cx             =   28258
         _cy             =   11271
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
            Height          =   6300
            Index           =   7
            Left            =   22485
            TabIndex        =   36
            Top             =   555
            Width           =   15720
            _cx             =   27728
            _cy             =   11112
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
            FormatString    =   $"frmImport.frx":28D4
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
         Begin VSFlex8Ctl.VSFlexGrid grdBox 
            Height          =   6240
            Left            =   315
            TabIndex        =   37
            Top             =   90
            Width           =   15390
            _cx             =   27146
            _cy             =   11007
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
            Cols            =   11
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmImport.frx":2994
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
         Height          =   6390
         Index           =   9
         Left            =   -17865
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   45
         Width           =   16020
         _cx             =   28258
         _cy             =   11271
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
            Height          =   6300
            Index           =   8
            Left            =   22485
            TabIndex        =   52
            Top             =   555
            Width           =   15720
            _cx             =   27728
            _cy             =   11112
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
            FormatString    =   $"frmImport.frx":2B46
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
            Height          =   6240
            Left            =   315
            TabIndex        =   53
            Top             =   90
            Width           =   15390
            _cx             =   27146
            _cy             =   11007
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
            FormatString    =   $"frmImport.frx":2C06
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
         Height          =   6390
         Index           =   10
         Left            =   -17565
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   45
         Width           =   16020
         _cx             =   28258
         _cy             =   11271
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
            Height          =   6300
            Index           =   9
            Left            =   22485
            TabIndex        =   61
            Top             =   555
            Width           =   15720
            _cx             =   27728
            _cy             =   11112
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
            FormatString    =   $"frmImport.frx":2DB6
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
         Begin VSFlex8Ctl.VSFlexGrid GrdAccount 
            Height          =   6240
            Left            =   315
            TabIndex        =   62
            Top             =   90
            Width           =   15390
            _cx             =   27146
            _cy             =   11007
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
            FormatString    =   $"frmImport.frx":2E76
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
      Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid4 
         Height          =   6390
         Left            =   -17265
         TabIndex        =   63
         Top             =   45
         Width           =   16020
         _cx             =   28257
         _cy             =   11271
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
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmImport.frx":305B
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
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   6390
         Index           =   11
         Left            =   -16965
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   45
         Width           =   16020
         _cx             =   28258
         _cy             =   11271
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
         Begin VB.CommandButton cmdCheckExcel 
            Caption         =   "›Õ’ „·› «·«ﬂ”Ì· ﬁ»· «·Õ›Ÿ"
            Height          =   225
            Left            =   6165
            TabIndex        =   71
            Top             =   30
            Width           =   3690
         End
         Begin VSFlex8UCtl.VSFlexGrid FgItems 
            Height          =   6300
            Index           =   10
            Left            =   22485
            TabIndex        =   65
            Top             =   555
            Width           =   15720
            _cx             =   27728
            _cy             =   11112
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
            FormatString    =   $"frmImport.frx":3165
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
         Begin VSFlex8Ctl.VSFlexGrid GrdAccount2 
            Height          =   6255
            Left            =   315
            TabIndex        =   66
            Top             =   270
            Width           =   15390
            _cx             =   27146
            _cy             =   11033
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
            Cols            =   15
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmImport.frx":3225
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
            Begin VB.Frame Frame2 
               Height          =   6555
               Left            =   180
               TabIndex        =   67
               Top             =   900
               Visible         =   0   'False
               Width           =   14475
               Begin VB.CommandButton Command7 
                  Caption         =   "x"
                  Height          =   195
                  Left            =   12930
                  TabIndex        =   69
                  Top             =   30
                  Width           =   825
               End
               Begin VSFlex8Ctl.VSFlexGrid GrdAccount3 
                  Height          =   5955
                  Left            =   120
                  TabIndex        =   68
                  Top             =   480
                  Width           =   14280
                  _cx             =   25188
                  _cy             =   10504
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
                  FormatString    =   $"frmImport.frx":34F3
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
               Begin VB.Label Label1 
                  Caption         =   "Õ”«»«   Õ «Ã ·„—«Ã⁄… ‰ ÌÃ… «Œÿ«¡ ›Ï „·› «·«ﬂ”Ì· ›Ï «·«ﬂÊ«œ"
                  Height          =   255
                  Left            =   4770
                  TabIndex        =   70
                  Top             =   240
                  Width           =   6615
               End
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   6390
         Index           =   12
         Left            =   -16665
         TabIndex        =   72
         TabStop         =   0   'False
         Top             =   45
         Width           =   16020
         _cx             =   28258
         _cy             =   11271
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
            Height          =   6300
            Index           =   11
            Left            =   22485
            TabIndex        =   73
            Top             =   555
            Width           =   15720
            _cx             =   27728
            _cy             =   11112
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
            FormatString    =   $"frmImport.frx":38AD
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
         Begin VSFlex8Ctl.VSFlexGrid grdFAGroups 
            Height          =   6240
            Left            =   315
            TabIndex        =   77
            Top             =   90
            Width           =   15390
            _cx             =   27146
            _cy             =   11007
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
            Cols            =   17
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmImport.frx":396D
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
         Height          =   6390
         Index           =   13
         Left            =   45
         TabIndex        =   74
         TabStop         =   0   'False
         Top             =   45
         Width           =   16020
         _cx             =   28258
         _cy             =   11271
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
         Begin VB.OptionButton Opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "·Ì” ·Â «Â·«ﬂ"
            Enabled         =   0   'False
            Height          =   255
            Index           =   1
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   83
            Top             =   0
            Value           =   -1  'True
            Visible         =   0   'False
            Width           =   2460
         End
         Begin VB.OptionButton Opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "·Â «Â·«ﬂ"
            Enabled         =   0   'False
            Height          =   225
            Index           =   0
            Left            =   2775
            RightToLeft     =   -1  'True
            TabIndex        =   82
            Top             =   30
            Visible         =   0   'False
            Width           =   1545
         End
         Begin VSFlex8UCtl.VSFlexGrid FgItems 
            Height          =   6270
            Index           =   12
            Left            =   22485
            TabIndex        =   75
            Top             =   720
            Width           =   15720
            _cx             =   27728
            _cy             =   11060
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
            FormatString    =   $"frmImport.frx":3CB5
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
         Begin VSFlex8Ctl.VSFlexGrid grdFa2 
            Height          =   5940
            Left            =   -30
            TabIndex        =   76
            Top             =   1020
            Visible         =   0   'False
            Width           =   16020
            _cx             =   28257
            _cy             =   10477
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
            Cols            =   23
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmImport.frx":3D75
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
         Begin VSFlex8Ctl.VSFlexGrid grdFa 
            Height          =   5940
            Left            =   0
            TabIndex        =   149
            Top             =   180
            Width           =   16020
            _cx             =   28257
            _cy             =   10477
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
            Cols            =   22
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmImport.frx":418C
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
         Height          =   6390
         Index           =   14
         Left            =   16755
         TabIndex        =   103
         TabStop         =   0   'False
         Top             =   45
         Width           =   16020
         _cx             =   28258
         _cy             =   11271
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
         Begin VB.TextBox txtTotalCredit 
            Height          =   300
            Left            =   8010
            TabIndex        =   129
            Text            =   "Text3"
            Top             =   6075
            Width           =   1230
         End
         Begin VB.TextBox txtTotalDebit 
            Height          =   300
            Left            =   9240
            TabIndex        =   128
            Text            =   "Text3"
            Top             =   6075
            Width           =   1545
         End
         Begin VB.CommandButton Command10 
            Caption         =   "›Õ’ „·› «·«ﬂ”Ì· ﬁ»· «·Õ›Ÿ"
            Height          =   300
            Left            =   12015
            TabIndex        =   116
            Top             =   30
            Width           =   3690
         End
         Begin VB.Frame Frame4 
            Height          =   4665
            Left            =   315
            TabIndex        =   112
            Top             =   1125
            Visible         =   0   'False
            Width           =   15090
            Begin VB.CommandButton Command12 
               Caption         =   "÷»ÿ «·«”„«¡ "
               Height          =   285
               Left            =   2490
               TabIndex        =   121
               Top             =   180
               Width           =   1365
            End
            Begin VB.CommandButton Command11 
               Caption         =   "«‰‘«¡ «·Õ”«» "
               Height          =   285
               Left            =   3990
               TabIndex        =   120
               Top             =   180
               Width           =   1365
            End
            Begin VB.CommandButton Command6 
               Caption         =   "x"
               Height          =   195
               Left            =   12930
               TabIndex        =   113
               Top             =   30
               Width           =   825
            End
            Begin VSFlex8Ctl.VSFlexGrid GrdAccount4 
               Height          =   5955
               Left            =   120
               TabIndex        =   114
               Top             =   480
               Width           =   14280
               _cx             =   25188
               _cy             =   10504
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
               Cols            =   9
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"frmImport.frx":452F
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
            Begin VB.Label Label4 
               Caption         =   "Õ”«»«   Õ «Ã ·„—«Ã⁄… ‰ ÌÃ… «Œÿ«¡ ›Ï „·› «·«ﬂ”Ì· ›Ï «·«ﬂÊ«œ"
               Height          =   255
               Left            =   5610
               TabIndex        =   115
               Top             =   210
               Width           =   6615
            End
         End
         Begin VB.CommandButton cmdCheckDataOpen 
            Caption         =   "Ã·» «Œ— Õ—ﬂ… —’Ìœ «›  «ÕÏ ··⁄„· ⁄·ÌÂ«"
            Height          =   300
            Left            =   2775
            TabIndex        =   111
            Top             =   30
            Width           =   2775
         End
         Begin VB.CommandButton cmdCreateopenEntry 
            Caption         =   "«‰‘«¡ «·ﬁÌœ «·«›  «ÕÏ"
            Enabled         =   0   'False
            Height          =   240
            Left            =   315
            TabIndex        =   110
            Top             =   90
            Width           =   1845
         End
         Begin VSFlex8UCtl.VSFlexGrid FgItems 
            Height          =   6270
            Index           =   13
            Left            =   22485
            TabIndex        =   104
            Top             =   720
            Width           =   15720
            _cx             =   27728
            _cy             =   11060
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
            FormatString    =   $"frmImport.frx":4695
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
         Begin VSFlex8Ctl.VSFlexGrid GrdAccountOpen 
            Height          =   5625
            Left            =   315
            TabIndex        =   105
            Top             =   405
            Width           =   15390
            _cx             =   27146
            _cy             =   9922
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
            Cols            =   12
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmImport.frx":4755
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
         Begin MSDataListLib.DataCombo cmbBranch 
            Height          =   315
            Left            =   8325
            TabIndex        =   106
            Top             =   60
            Width           =   2460
            _ExtentX        =   4339
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSComCtl2.DTPicker txtOPenDate 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11265
               SubFormatType   =   3
            EndProperty
            Height          =   330
            Left            =   5850
            TabIndex        =   108
            ToolTipText     =   " ≈·Ï  «—ÌŒ √ÕœÀ"
            Top             =   30
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   582
            _Version        =   393216
            CalendarBackColor=   -2147483624
            CalendarTitleBackColor=   10383715
            CheckBox        =   -1  'True
            CustomFormat    =   "yyyy/M/d"
            Format          =   151977987
            CurrentDate     =   37357
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «·ﬁÌœ"
            Height          =   270
            Index           =   4
            Left            =   7395
            RightToLeft     =   -1  'True
            TabIndex        =   109
            Top             =   60
            Width           =   615
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·›—⁄"
            Height          =   240
            Index           =   3
            Left            =   11085
            RightToLeft     =   -1  'True
            TabIndex        =   107
            Top             =   120
            Width           =   930
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   6390
         Index           =   15
         Left            =   17055
         TabIndex        =   117
         TabStop         =   0   'False
         Top             =   45
         Width           =   16020
         _cx             =   28258
         _cy             =   11271
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
            Height          =   6270
            Index           =   14
            Left            =   22485
            TabIndex        =   118
            Top             =   720
            Width           =   15720
            _cx             =   27728
            _cy             =   11060
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
            FormatString    =   $"frmImport.frx":493E
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
         Begin VSFlex8Ctl.VSFlexGrid GrdEmpFee 
            Height          =   6105
            Left            =   0
            TabIndex        =   119
            Top             =   240
            Width           =   16020
            _cx             =   28257
            _cy             =   10769
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
            Cols            =   13
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmImport.frx":49FE
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
         Height          =   6390
         Index           =   16
         Left            =   17355
         TabIndex        =   122
         TabStop         =   0   'False
         Top             =   45
         Width           =   16020
         _cx             =   28258
         _cy             =   11271
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
            Height          =   6270
            Index           =   15
            Left            =   22485
            TabIndex        =   123
            Top             =   720
            Width           =   15720
            _cx             =   27728
            _cy             =   11060
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
            FormatString    =   $"frmImport.frx":4C0C
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
         Begin VSFlex8Ctl.VSFlexGrid grdSchools 
            Height          =   5940
            Left            =   0
            TabIndex        =   124
            Top             =   405
            Width           =   16020
            _cx             =   28257
            _cy             =   10477
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
            Cols            =   13
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmImport.frx":4CCC
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
         Height          =   6390
         Index           =   17
         Left            =   17655
         TabIndex        =   125
         TabStop         =   0   'False
         Top             =   45
         Width           =   16020
         _cx             =   28258
         _cy             =   11271
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
            Height          =   6270
            Index           =   16
            Left            =   22485
            TabIndex        =   126
            Top             =   720
            Width           =   15720
            _cx             =   27728
            _cy             =   11060
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
            FormatString    =   $"frmImport.frx":4F1F
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
         Begin VSFlex8Ctl.VSFlexGrid grdCars 
            Height          =   5940
            Left            =   0
            TabIndex        =   127
            Top             =   405
            Width           =   16020
            _cx             =   28257
            _cy             =   10477
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
            FormatString    =   $"frmImport.frx":4FDF
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
         Height          =   6390
         Index           =   18
         Left            =   17955
         TabIndex        =   133
         TabStop         =   0   'False
         Top             =   45
         Width           =   16020
         _cx             =   28258
         _cy             =   11271
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
         Begin VB.CommandButton Command15 
            Caption         =   "Õ· „‘ﬂ·… «·„”«›«  ··„ÊŸ›Ì‰ Ê Ê“Ì⁄ «·«”„ ⁄·Ï «· ·«  Œ«‰« "
            Height          =   345
            Left            =   3390
            TabIndex        =   148
            Top             =   60
            Width           =   2160
         End
         Begin VB.TextBox txtSql 
            Height          =   135
            Left            =   4935
            MultiLine       =   -1  'True
            TabIndex        =   146
            Text            =   "frmImport.frx":526C
            Top             =   1470
            Width           =   6465
         End
         Begin VB.CommandButton Command13 
            Caption         =   "„‘ﬂ·… «·„’«—Ì› Ê«·Õ”«»« "
            Height          =   195
            Left            =   12015
            TabIndex        =   145
            Top             =   1410
            Width           =   1845
         End
         Begin VB.TextBox txtWhere 
            Height          =   390
            Left            =   315
            TabIndex        =   144
            Top             =   1215
            Width           =   2460
         End
         Begin VB.TextBox txtFeildNameE 
            Height          =   330
            Left            =   2775
            TabIndex        =   141
            Text            =   "cusNamee"
            Top             =   840
            Width           =   2775
         End
         Begin VB.TextBox txtFeildName 
            Height          =   345
            Left            =   2775
            TabIndex        =   138
            Text            =   "cusName"
            Top             =   495
            Width           =   2775
         End
         Begin VB.TextBox txtTableName2 
            Height          =   330
            Left            =   7695
            TabIndex        =   137
            Text            =   "TblCustemers"
            Top             =   525
            Width           =   2475
         End
         Begin VB.CommandButton cmdResolveLongName 
            Caption         =   "Õ· „‘ﬂ·… «·„”«›« "
            Height          =   300
            Left            =   12330
            TabIndex        =   136
            Top             =   525
            Width           =   2145
         End
         Begin VSFlex8UCtl.VSFlexGrid FgItems 
            Height          =   6270
            Index           =   17
            Left            =   22485
            TabIndex        =   134
            Top             =   720
            Width           =   15720
            _cx             =   27728
            _cy             =   11060
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
            FormatString    =   $"frmImport.frx":5272
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
         Begin VSFlex8Ctl.VSFlexGrid Grdtmp 
            Height          =   4740
            Left            =   0
            TabIndex        =   135
            Top             =   1605
            Width           =   16020
            _cx             =   28257
            _cy             =   8361
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
            Cols            =   12
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmImport.frx":5332
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
         Begin VB.Label Label6 
            Caption         =   "«”„ Õﬁ· «·«”„ E"
            Height          =   210
            Index           =   1
            Left            =   5550
            TabIndex        =   142
            Top             =   870
            Width           =   2145
         End
         Begin VB.Label Label6 
            Caption         =   "«”„ Õﬁ· «·«”„"
            Height          =   240
            Index           =   0
            Left            =   5550
            TabIndex        =   140
            Top             =   555
            Width           =   2145
         End
         Begin VB.Label Label5 
            Caption         =   "«”„ «·ÃœÊ·"
            Height          =   225
            Left            =   10170
            TabIndex        =   139
            Top             =   585
            Width           =   2160
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid grdExcel 
         Height          =   6390
         Index           =   1
         Left            =   18255
         TabIndex        =   170
         Top             =   45
         Width           =   16020
         _cx             =   28257
         _cy             =   11271
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
         Cols            =   14
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmImport.frx":54E8
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
      TabIndex        =   58
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
      Format          =   113115139
      CurrentDate     =   37357
   End
End
Attribute VB_Name = "frmImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mIndex As Long
Dim FirstPeriodDateInthisYear  As Date

Dim mBranchName As String
Dim mBranchId As Long
Dim mPurchasePrice As Double
Dim mAccDepreciation As Double
'Dim mNewId As Long
Dim mLast As Boolean
Dim TxtnoOfInst As Double
Dim TxtAge As Double
Dim TxtCurrentValue As Double
Dim txtinstallValue As Double
Dim txtinstallmentresult As Double
Dim txtinstallDo As Double
Dim TxtPurchasePrice As Double
Dim DCPreFix As Double
Dim TXT24 As String
Dim TXT26 As String
Dim TXT25 As String
Dim TXT31 As String
Dim TXT40 As String
Dim TXtPercentage1 As String
Dim txtPercentage2 As String


    Dim AccountName As String
    Dim Percentage1 As Integer
    Dim Percentage2 As Integer
    Dim DepType As Boolean
    Dim Account_code As String
    Dim Account_code1 As String
    Dim Account_code2 As String
    Dim Account_code3 As String
    Dim Account_code4 As String
   Dim noOfInstallments As Integer
        Dim Age As Integer
        Dim currentvalue As Double
        Dim installValue As Double
        Dim RemainInstallments As Double
        Dim EXEInstallments As Double
        Dim DCGroup As Long
        
        Dim txtopening_balance_voucher_id As String
        
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
    
    s = " SELECT REPLACE( Replace(Account_Name,'–„„',''),'/','') Name ,Account_Code,BranchID FROM ACCOUNTS WHERE Parent_Account_Code ="
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
        rsData!Emp_code = GetCode(mSer)
        rsData!Emp_id = mSer
        rsData!emp_name = Trim(rsDummy!Name & "")
        astrSplit2tems2 = Split(Trim(rsDummy!Name & ""), " ")
        rsData!BranchID = Val(rsDummy!BranchID & "")
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
        rsData!Account_code = Trim(rsDummy!Account_code & "")
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
            s = " SELECT  REPLACE( Replace(Account_Name,'⁄„·«¡',''),'/','') NAME,Account_Code,opening_balance,BranchID"
        s = s & " FROM ACCOUNTS WHERE Parent_Account_Code In"
        s = s & " (SELECT  Account_Code  FROM ACCOUNTS AS a WHERE a.Account_Name LIKE '%⁄„·«¡%' Or a.Account_Name LIKE '%„œÌ‰Ê‰ „ ‰Ê⁄Ê‰%'"
        s = s & " AND a.last_account = 0"
        s = s & " )  AND last_account = 1"
        s = s & " and Account_Code Not In (Select IsNull(Account_Code,'') From " & mTable & " ) "
        


    Else
         s = " SELECT  REPLACE( Replace(Account_Name,'„Ê—œÌ‰',''),'/','') NAME,Account_Code,BranchID,opening_balance "
        s = s & " FROM ACCOUNTS WHERE Parent_Account_Code In"
        s = s & " (SELECT  Account_Code  FROM ACCOUNTS AS a WHERE a.Account_Name LIKE '%„Ê—œÊ‰%' Or a.Account_Name LIKE '%„Ê—œÌ‰%' Or a.Account_Name LIKE '%œ«∆‰Ê‰  Ã«—ÌÊ‰' "
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
        rsData!cusName = Trim(rsDummy!Name & "")
        rsData!CusNamee = Trim(rsDummy!Name & "")
        rsData!Type = mType
        rsData!CreditlimitCredit = 0
        rsData!OpenBalance = Val(rsDummy!opening_balance & "")
    
        
         If Val(rsDummy!opening_balance & "") < 0 Then
            rsData!OpenBalanceType = 1
        Else
            rsData!OpenBalanceType = 0
        End If
        rsData!SaleType = 0
        rsData!Locked = 0
        rsData!CreditlimitCredit = 0
        rsData!CreditlimitCredit = 0
        rsData!BranchID = Val(rsDummy!BranchID & "")
        rsData!Account_code = Trim(rsDummy!Account_code & "")
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

  s = "SELECT Max(BoxID) MaxID  FROM " & mTable & " AS te "
    rsDummy.Open s, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    If Not rsDummy.EOF Then
        mMaxId = Val(rsDummy!MaxId & "")
    End If
    rsDummy.Close
    
    
    's = " SELECT  REPLACE( Replace(Account_Name,'',''),'/','') NAME,Account_Code,BranchID "
    s = " SELECT Account_Name as  NAME,Account_Code,BranchID "
    s = s & "         ,ParntName = (SELECT Account_Name FROM ACCOUNTS AS a WHERE a.Account_Code =ACCOUNTS.Parent_Account_Code )"
    s = s & "   FROM ACCOUNTS WHERE Parent_Account_Code In"
    s = s & " (SELECT  account_code  FROM ACCOUNTS AS a WHERE a.Account_Name LIKE '%⁄Âœ%' Or a.Account_Name LIKE '%‰ﬁœÌ… »«·Œ“Ì‰…%'"
    s = s & " AND a.last_account = 0)"
    s = s & " and last_account = 1 and Account_Code Not In (Select IsNull(Account_Code,'') From " & mTable & " ) "
      
      If Trim(DboParentAccount2.Text) <> "" Then
    
            
            s = " SELECT a.last_account, REPLACE( Replace(a.Account_Name,'',''),'/','') NAME,a.Account_Code,a.BranchID,"
        s = s & "         a.Account_Code,"
        s = s & "                 a.Account_Name,"
        s = s & "                 a2.Account_Name      AS parantName,"
        s = s & "                 a.Parent_Account_Code"
        s = s & "          FROM   ACCOUNTS             AS a"
        s = s & "                 INNER JOIN ACCOUNTS  AS a2"
        s = s & "                      ON  a.Parent_Account_Code = a2.Account_Code"
        s = s & "          WHERE  a.Account_Code IN (SELECT Code"
        s = s & "                                    FROM   [FN_MAIN_ACCOUNT_SUB_CODES]('" & Trim(DboParentAccount2.BoundText) & "', '" & Trim(DboParentAccount2.BoundText) & "', 0))"
        s = s & " OR (a.Account_Code = '" & Trim(DboParentAccount2.BoundText) & "')"
        s = s & "          Order By"
        s = s & "                 a.Parent_Account_Code,"
        s = s & "                 a.last_account"
 
 End If
      
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
          If Len(Trim(rsDummy!Name & "") & " " & Trim(rsDummy!parantName & "")) > 50 Then
            rsData!BoxName = Right(Trim(rsDummy!Name & "") & " " & Trim(rsDummy!parantName & ""), 50)
        Else
            rsData!BoxName = Trim(Trim(rsDummy!Name & "") & " " & Trim(rsDummy!parantName & ""))
        End If
        rsData!BoxName = Trim(rsDummy!Name & "")
        rsData!BoxNamee = rsData!BoxName
        rsData!Type = 1
        rsData!Account_code = Trim(rsDummy!Account_code & "")
        rsData!BranchID = Val(rsDummy!BranchID & "")
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
     s = "SELECT Max(BankId) MaxID  FROM " & mTable & " AS te "
    rsDummy.Open s, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    If Not rsDummy.EOF Then
        mMaxId = Val(rsDummy!MaxId & "")
    End If
    rsDummy.Close
    
    s = " SELECT  REPLACE( Replace(Account_Name,'',''),'/','') NAME,Account_Code,BranchID "
    s = s & " , BranchName = (SELECT branch_name FROM TblBranchesData WHERE branch_id = ACCOUNTS.BranchID)"
    s = s & "   FROM ACCOUNTS WHERE Parent_Account_Code In"
    s = s & " (SELECT  account_code  FROM ACCOUNTS AS a WHERE  (a.Account_Name Like '»‰ﬂ%' OR a.Account_Name = '»‰Êﬂ%' OR a.Account_Name like '«·»‰ﬂ%' OR a.Account_Name like 'Õ”«» »‰ﬂ%')"
    s = s & " AND Not (a.Account_Name LIKE '%„’—Ê›« %')"
    s = s & " AND a.last_account = 0)"
    s = s & " and last_account = 1 and Account_Code Not In (Select IsNull(Account_Code,'') From " & mTable & " ) "
    
    If Trim(DboParentAccount2.Text) <> "" Then
    
            
            s = " SELECT a.last_account, REPLACE( Replace(a.Account_Name,'',''),'/','') NAME,a.Account_Code,a.BranchID,"
        s = s & "         a.Account_Code,"
        s = s & "                 a.Account_Name,"
        s = s & "                 a2.Account_Name      AS parantName,"
        s = s & "                 a.Parent_Account_Code"
        s = s & "          FROM   ACCOUNTS             AS a"
        s = s & "                 INNER JOIN ACCOUNTS  AS a2"
        s = s & "                      ON  a.Parent_Account_Code = a2.Account_Code"
        s = s & "          WHERE  a.Account_Code IN (SELECT Code"
        s = s & "                                    FROM   [FN_MAIN_ACCOUNT_SUB_CODES]('" & Trim(DboParentAccount2.BoundText) & "', '" & Trim(DboParentAccount2.BoundText) & "', 0))"
        s = s & " OR (a.Account_Code = '" & Trim(DboParentAccount2.BoundText) & "')"
        s = s & "          Order By"
        s = s & "                 a.Parent_Account_Code,"
        s = s & "                 a.last_account"
 
 End If
 
      rsDummy.Open s, Cn, adOpenForwardOnly, adLockReadOnly
    
    s = "Select * from BanksData Where 1 = -1"
    rsData.Open s, Cn, adOpenKeyset, adLockOptimistic

    mSer = mMaxId
    Do While Not rsDummy.EOF
        mSer = mSer + 1
        rsData.AddNew
        'rsData!Fullcode = GetCode(mSer)
        'rsData!Code = GetCode(mSer)
        rsData!BankID = mSer
       If Val(rsDummy!BranchID & "") <> 1 Then
        rsData!BankID = mSer
       End If
       rsData!BranchID = Val(rsDummy!BranchID & "")
      
            rsData!BankName = Trim(rsDummy!Name & "") & " - " & Trim(rsDummy!BranchName & "")
     
        
        rsData!BankNamee = Right(Trim(rsDummy!Name & ""), 50)
        
        rsData!Account_code = Trim(rsDummy!Account_code & "")
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

    s = "SELECT Max(Id) MaxID  FROM " & mTable & " AS te "
    Set rsDummy = New ADODB.Recordset
    rsDummy.Open s, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    If Not rsDummy.EOF Then
        mMaxId = Val(rsDummy!MaxId & "")
    End If
    rsDummy.Close


    s = " SELECT  REPLACE( Replace(Account_Name,'',''),'/','') NAME,Account_Code,BranchID "
    s = s & "   FROM ACCOUNTS WHERE Parent_Account_Code In"
    s = s & " (SELECT  account_code  FROM ACCOUNTS AS a WHERE "
  
    s = s & "                                                                a.last_account = 0"
    s = s & "                                                                 AND"
    s = s & "                                                                 a.Account_Name NOT LIKE '%„’—Ê›«  „œ›Ê⁄… „ﬁœ„«%'"
    s = s & "                                                                 AND ("
    s = s & "                                                                 "
    s = s & " a.Account_Name LIKE N'%„’—Ê›« %'  Or a.Account_Name LIKE N'% ﬂ«·Ì› «·„‘«—Ì⁄%'  Or a.Account_Name LIKE N'%„’«—Ì›%' Or a.Account_Name LIKE N'%Âœ«Ì« Ê⁄Ì‰« %' Or a.Account_Name LIKE N'%Âœ«Ì« Ê⁄Ì‰« %'  "
    s = s & " Or a.Account_Name LIKE N'% ﬂ«·Ì› «·«‘ —«ﬂ%'"
    s = s & " Or a.Account_Name LIKE N'%« ⁄«» Ê«” ‘«—« %'"
    s = s & " Or a.Account_Name LIKE N'%œ„€… ° »—Ìœ°  ·Ì›Ê‰« %'"
    s = s & " Or a.Account_Name LIKE N'% √„Ì‰ Õ—Ìﬁ Ê”—ﬁ…%'"
    s = s & " Or a.Account_Name LIKE N'%√⁄„«· ’Ì«‰…%'"
    s = s & " Or a.Account_Name LIKE N'% ﬂ·›… «·√ÃÊ—%'"
    s = s & " Or a.Account_Name LIKE N'% ﬂ·›… «·√ÃÊ—%'"
    s = s & " AND a.last_account = 0))"
    s = s & " and last_account = 1 and Account_Code Not In (Select IsNull(Account_Code,'') From " & mTable & " ) "
    
    
    If Trim(DboParentAccount2.Text) <> "" Then
    
            
            s = " SELECT a.last_account, REPLACE( Replace(a.Account_Name,'',''),'/','') NAME,a.Account_Code,a.BranchID,"
        s = s & "         a.Account_Code,"
        s = s & "                 a.Account_Name,"
        s = s & "                 a2.Account_Name      AS parantName,"
        s = s & "                 a.Parent_Account_Code"
        s = s & "          FROM   ACCOUNTS             AS a"
        s = s & "                 INNER JOIN ACCOUNTS  AS a2"
        s = s & "                      ON  a.Parent_Account_Code = a2.Account_Code"
        s = s & "          WHERE  a.Account_Code IN (SELECT Code"
        s = s & "                                    FROM   [FN_MAIN_ACCOUNT_SUB_CODES]('" & Trim(DboParentAccount2.BoundText) & "', '" & Trim(DboParentAccount2.BoundText) & "', 0))"
        s = s & " OR (a.Account_Code = '" & Trim(DboParentAccount2.BoundText) & "')"
        s = s & "          Order By"
        s = s & "                 a.Parent_Account_Code,"
        s = s & "                 a.last_account"
 
 End If
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
        If Len(Trim(rsDummy!Name & "")) > 50 Then
            rsData!Name = Right(Trim(rsDummy!Name & ""), 50)
        Else
            rsData!Name = Trim(rsDummy!Name & "")
        End If

        rsData!NameE = Right(Trim(rsDummy!Name & ""), 50)
        
        rsData!TypicalProduction = 0
        rsData!IndirectCosts = 0
        rsData!Account_code = Trim(rsDummy!Account_code & "")
        'rsData!BranchID = val(rsDummy!BranchID & "")
        rsData.Update
        rsDummy.MoveNext
    Loop
    
        s = " Update ExpensesType"
    s = s & " Set ExpensesType.parent_account = ACCOUNTS.Parent_Account_Code"
    s = s & " From dbo.ExpensesType"
    s = s & "        INNER JOIN dbo.ACCOUNTS"
    s = s & "             ON  dbo.ExpensesType.Account_Code = dbo.ACCOUNTS.Account_Code"
    Cn.Execute s
Case "TblStore"
    s = " SELECT  REPLACE( Replace(Account_Name,'',''),'/','') NAME,Account_Code,BranchID "
    s = s & "   FROM ACCOUNTS WHERE Parent_Account_Code In"
    s = s & " (SELECT  account_code  FROM ACCOUNTS AS a WHERE a.Account_Name LIKE '%«·„Œ“Ê‰%'"
    s = s & " AND a.last_account = 0)"
    s = s & " and last_account = 1 and Account_Code Not In (Select IsNull(Account_Code,'') From " & mTable & " ) "
      rsDummy.Open s, Cn, adOpenForwardOnly, adLockReadOnly
    
    s = "Select * from TblStore Where 1 = -1"
    rsData.Open s, Cn, adOpenKeyset, adLockOptimistic
    
    mSer = mMaxId
    Do While Not rsDummy.EOF
        mSer = mSer + 1
        rsData.AddNew
        'rsData!Fullcode = GetCode(mSer)
        'rsData!Code = GetCode(mSer)
        rsData!StoreId = mSer
        If Len(Trim(rsDummy!Name & "")) > 50 Then
            rsData!StoreName = Right(Trim(rsDummy!Name & ""), 50)
        Else
            rsData!StoreName = Trim(rsDummy!Name & "")
        End If
        
        rsData!StoreNamee = Right(Trim(rsDummy!Name & ""), 50)
        
        rsData!Account_code = Trim(rsDummy!Account_code & "")
        rsData!Account_code1 = Trim(rsDummy!Account_code & "")
        rsData!BranchID = Val(rsDummy!BranchID & "")
        rsData.Update
        rsDummy.MoveNext
    Loop

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

Private Sub chkBalanceOnly_Click()
If chkBalanceOnly.Value = vbChecked Then

    grdManBal.Visible = True
    grdMan.Visible = False
    cmdSave.Enabled = False
Else
    grdManBal.Visible = False
    grdMan.Visible = True
End If
End Sub

Private Sub chkIsRepeatCode_Click()
If chkIsRepeatCode.Value = vbChecked Then
    IsRepeatCode = True
Else
    IsRepeatCode = False
End If
End Sub
Private Sub chkIsRepeatName_Click()
If chkIsRepeatName.Value = vbChecked Then
    IsRepeatName = True
Else
    IsRepeatName = False
End If
End Sub
Private Sub cmdCheckDataOpen_Click()
Dim rsDummy As New ADODB.Recordset
Dim rsDummyData As New ADODB.Recordset
Dim mAccountName As String
Dim mAccount_Serial As String
Dim mAccountCode As String
Dim mBranchId As Long

Dim mDebit As Double
Dim mCredit As Double

Dim NoteID As Long
Dim EntryID As Long
Dim i As Long
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
    mBranchId = Val(rsDummyData!branch_id & "")
    cmbBranch.BoundText = mBranchId
    txtOPenDate.Value = rsDummyData!RecordDate & ""
'    EntryID = Val(rsDummy!EntryID & "")
'    NoteID = Val(rsDummy!NoteID & "")
    cmdCreateopenEntry.Enabled = True
End If
End Sub

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
Case 9
    GetFromAccount "TblStore", 0
    
End Select
MsgBox " „ ‰ﬁ· «·»Ì«‰« "
End Sub

Private Sub CmdRecalcAccountSupp_Click()
Dim rs As ADODB.Recordset
Dim Account_Code_dynamic As String
Dim sql As String
Dim Current_account As String
Dim ParnetAccount As String
Dim Account_Code_dynamic1 As String
Dim my_branch As String

Dim ss As String

Dim rsDummy As New ADODB.Recordset

    Set rs = New ADODB.Recordset
    Account_Code_dynamic1 = get_account_code_branch(9, my_branch)
        If Account_Code_dynamic1 = "NO branch" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·« ÌÊÃœ Õ”«»«    Ê”Ìÿ «›  «ÕÌ ··„Ê—œÌ‰ ·Â–« «·›—⁄"
        Else
            Msg = "No Accounts For This Branch"
        End If

        MsgBox Msg, vbCritical
        
        Exit Sub

    ElseIf Account_Code_dynamic1 = "NO account" Then

        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "Õ”«» Ê”Ìÿ «›  «ÕÌ ··„Ê—œÌ‰ €Ì— „Õœœ ›Ï «·›—⁄"
        Else
            Msg = "No Accounts For This Branch"
        End If

        MsgBox Msg, vbCritical
        
        Exit Sub
    End If
    
    sql = " SELECT     *"
    sql = sql & " From dbo.TblCustemers"
    sql = sql & " WHERE     (Type = 2)and (CusID<>1) and (CusID<>2) "
    If Check1.Value = vbChecked Then
        sql = sql & " and IsNull(Account_Code,'') = ''"
    End If
    rs.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
        
 Dim strA As String
 Dim Stre As String
 Dim FlgDele As Boolean
 
 'Account_Code_dynamic1 = "a1a1a2a1a1"
If Check2.Value = vbChecked Then
        For i = 1 To rs.RecordCount
        
           ss = "Select Account_Code From ACCOUNTS Where Branch = " & Val(rs!BranchID & "")
        Set rsDummy = New ADODB.Recordset
        rsDummy.Open ss, Cn, adOpenKeyset, adLockReadOnly
        
        If Not rsDummy.EOF Then
            Account_Code_dynamic = Trim(rsDummy!Account_code & "")
        End If
        
        FlgDele = False
        If DeleteAccount(IIf(IsNull(rs("Account_Code").Value), "", rs("Account_Code").Value)) = True Then
        rs("Account_Code").Value = ""
        Else
        FlgDele = True
        End If
        If Check8.Value = vbChecked And FlgDele = False Then
        rs("parent_account").Value = ""
        End If
        rs.MoveNext
        Next i
        
  End If
 
  If rs.RecordCount > 0 Then
  rs.MoveFirst
  End If
        For i = 1 To rs.RecordCount

           Current_account = IIf(IsNull(rs("Account_Code").Value), "", rs("Account_Code").Value)
           ParnetAccount = IIf(IsNull(rs("parent_account").Value), "", rs("parent_account").Value)
           If check_account_exist(ParnetAccount) = True Then
           Account_Code_dynamic = ParnetAccount
           Else
           Account_Code_dynamic = Account_Code_dynamic1
           End If
           
                   ss = "Select Account_Code From ACCOUNTS Where Branch = " & Val(rs!BranchID & "")
        Set rsDummy = New ADODB.Recordset
        rsDummy.Open ss, Cn, adOpenKeyset, adLockReadOnly
        
        If Not rsDummy.EOF Then
            Account_Code_dynamic = Trim(rsDummy!Account_code & "")
        End If
                 If check_account_exist(Account_Code_dynamic) = True Then
                     If Current_account = "" Then 'new
                      rs("Account_Code").Value = AddNewAccount(Account_Code_dynamic, IIf(IsNull(rs("CusName")), rs("Fullcode").Value, rs("CusName")), True, False, IIf(IsNull(rs("CusNamee")), rs("Fullcode").Value, rs("CusNamee")), , , , , , , , , , 1, 1, 1, 0, 0, , , , Val(rs!BranchID & ""))          '
                      'IIf(IsNull(rs("parent_account")), Account_Code_dynamic, rs("parent_account")) = 1
                                     ' rs("Account_Code").value = ModAccounts.AddNewAccount(Account_Code_dynamic, IIf(IsNull(rs("CusName")), rs("Fullcode").value, rs("CusName")), True, False, IIf(IsNull(rs("CusNamee")), rs("Fullcode").value, rs("CusNamee")), , , , , , , , , , 1, 1, 1, 0, 0)

                     Else 'check
                            If check_account_exist(Current_account) = False Then
                              rs("Account_Code").Value = AddNewAccount(Account_Code_dynamic, IIf(IsNull(rs("CusName")), rs("Fullcode").Value, rs("CusName")), True, False, IIf(IsNull(rs("CusNamee")), rs("Fullcode").Value, rs("CusNamee")), , , , , , , , , , 1, 1, 1, 0, 0, , , , Val(rs!BranchID & ""))  '
                             Else
                             '  ModAccounts.EditAccount rs("Account_Code").value, IIf(IsNull(rs("Emp_Name")), rs("fullcode").value, rs("Emp_Name")), IIf(IsNull(rs("Emp_Namee")), IIf(IsNull(rs("Emp_Name")), rs("fullcode").value, rs("Emp_Name")), rs("Emp_Namee")), , , , , , , , , , , , , , , , , True
                               EditAccount rs("Account_Code").Value, IIf(IsNull(rs("CusName")), rs("Fullcode").Value, rs("CusName")), IIf(IsNull(rs("CusNamee")), rs("Fullcode").Value, rs("CusNamee")), , , , , , , , , 1, 1, 1, 0, 0, , , , , Val(rs!BranchID & "")
                             End If
                     
                     End If
         
                  End If

                  
                  
       rs.MoveNext
        Next i

     
    MsgBox " „"
   
End Sub

Private Sub cmdResolveLongName_Click()
Dim rsDummy As New ADODB.Recordset
Dim s As String
Dim X As String
Dim xx As Variant
Dim xxx As Variant
Dim mtmpName As String
Dim mName As String
Dim mNamee As String
Dim II As Integer
Dim i As Long
If txtFeildNameE = "" Then txtFeildNameE = txtFeildName & "E"
    s = "Select " & txtFeildName & " ," & txtFeildNameE & ""
    's = s & " ,cusId "
    If UCase(txtTableName2) = "ACCOUNTS" Then
        s = s & " ,Account_Code"
    ElseIf UCase(txtTableName2) = "TBLCUSTEMERS" Then
        s = s & " ,CUSID "
    ElseIf UCase(txtTableName2) = "TBLITEMS" Then
        s = s & " ,ItemId "
    End If
    s = s & " from  " & txtTableName2 & " "
    'where   Len(" & txtFeildName & ")  "
    If Trim(txtWhere) <> "" Then
        s = s & " Where " & txtWhere
    End If
    rsDummy.Open s, Cn, adOpenKeyset, adLockOptimistic
    Do While Not rsDummy.EOF
        
'        If rsDummy!ItemID = 67 Then
'            X = 67
'        End If
    
        X = Trim(rsDummy("" & txtFeildName & ""))
        i = 0
        xx = Split(X)
        mName = ""
        
        
        
        For i = 0 To UBound(xx)
            xxx = Split(xx(i), vbTab)
            II = 0
            mtmpName = ""
            For II = 0 To UBound(xxx)
                If xxx(II) <> "" Then
                    mtmpName = xxx(II)
                    Exit For
                End If
            Next
            
            If mtmpName = "" Then mtmpName = xx(i)
            If i = UBound(xx) - 1 And Len(RTrim(LTrim(mtmpName))) > 20 Then
                
            Else
                If mtmpName <> "" Then
                    
                    If Val(mtmpName) <> 0 Then
                        mtmpName = " " & CStr(mtmpName)
                    End If
                    xxx = Split(mtmpName, vbent)
                    If UBound(xxx) > 0 Then
                       ' mtmpName = ""
                        
                        For jj = 0 To UBound(xxx)
                        '    mtmpName = mtmpName & xxx(jj)
                        Next
                    End If

                    mName = mName & " " & Trim(mtmpName)
                End If
                
            End If
            
        Next
        X = rsDummy("" & txtFeildNameE & "") & ""
        i = 0
        xx = Split(X)
        mNamee = ""
        For i = 0 To UBound(xx) - 1
            If i = UBound(xx) - 1 And Len(RTrim(LTrim(xx(i)))) > 20 Then
                
            Else
                mNamee = mNamee & " " & xx(i)
            End If
        Next
    
        
        rsDummy("" & txtFeildName & "") = Trim(mName)
        rsDummy("" & txtFeildNameE & "") = Trim(IIf(mNamee = "", mName, mNamee))
        rsDummy.Update
        
        rsDummy.MoveNext
    Loop
    
    MsgBox " „"
End Sub

Private Sub cmdSave_Click()

    Dim i    As Long
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
        Case 7
    
            For i = 1 To grdBox.Rows - 1
                grdBox.TextMatrix(i, grdBox.ColIndex("BoxName")) = Trim(grdBox.TextMatrix(i, grdBox.ColIndex("BoxName"))) & " " & grdBox.TextMatrix(i, grdBox.ColIndex("ref"))
            Next
            Set mGrd = grdBox
        Case 9
            Set mGrd = GrdAccount
        Case 11
            Set mGrd = GrdAccount2
        Case 12
            Set mGrd = grdFAGroups
        Case 13
            Set mGrd = grdFa
        Case 15
            Set mGrd = GrdEmpFee
        Case 16
            Set mGrd = grdSchools
        Case 17
            Set mGrd = grdCars
        Case 18
            Set mGrd = Grdtmp
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
            s = "update TblEmployee set TblEmployee.workstate =1 where jopstatusid=1 and TblEmployee.workstate = null"
            Cn.Execute s
    
            s = "UPDATE TblEmployee SET Emp_Name = Emp_Namee WHERE ISNULL(Emp_Name,'') = ''"
            Cn.Execute s
    
            s = "UPDATE TblEmployee SET BranchId = 1 where IsNull(BranchId ,0) = 0 "
    
            Cn.Execute s
    
            s = "UPDATE TblEmployee SET InsuranceState = 1 where IsNull(InsuranceNO,0) <> 0 "
            Cn.Execute s
            s = "UPDATE TblEmployee SET DepartmentID = 1 where IsNull(DepartmentID ,0) = 0 "
            Cn.Execute s
            s = " UPDATE TblEmployee SET Nationality = (SELECT NAME FROM Nationality AS n WHERE id = TblEmployee.NationlID)"
            Cn.Execute s
        Case 3
            s = "Select * from TblUnites Where UnitID =  -1"
            saveGridExcel s, mGrd, "UnitName", "UnitID", "TblUnites"
        Case 2
    
                        s = "Select * from groups Where GroupID =  -1"
                        saveGridExcel s, mGrd, "Fullcode", "GroupID", "groups"
                        s = " UPDATE Groups SET LastGroup = 1 WHERE GroupID NOT IN (SELECT ISNULL(ff.ParentID,0) FROM Groups ff )"
                        Cn.Execute s
      '      LoadExcelToGroups grdGroups
        Case 1
            s = "Select * from TblCustemers Where CusID =  -1"
            saveGridExcel s, mGrd, "Fullcode", "CusID ", "TblCustemers"
            s = " UPDATE TblCustemers SET code = Fullcode"
            Cn.Execute s
    
            s = " UPDATE TblCustemers SET OpenBalanceType = 1 WHERE IsNull(OpenBalance,0) < 0"
            Cn.Execute s
            s = " UPDATE TblCustemers SET OpenBalanceType = 0 WHERE IsNull(OpenBalance,0) > 0"
            Cn.Execute s

            s = "UPDATE TblCustemers SET OpenBalance = ABS(OpenBalance)"
            Cn.Execute s
            s = " UPDATE TblCustemers SET BranchId = 1 WHERE ISNULL(BranchId,0) = 0"
            Cn.Execute s
        Case 4
            s = "Select * from tblItems Where ItemID =  -1"
            saveGridExcel s, mGrd, "code", "ItemID ", "tblItems"
    
            s = "Select * from TblItemsUnits Where ItemID =  -1"
            Dim rsItemsUnits As New ADODB.Recordset
            rsItemsUnits.Open s, Cn, adOpenKeyset, adLockOptimistic
            Dim mUnitID      As Long
            Dim UnitFactor   As Double
            Dim mSecondEntry As Integer
        
            For i = 1 To mGrd.Rows - 1
                mSecondEntry = 0
                If Val(grdItems.TextMatrix(i, grdItems.ColIndex("UnitId"))) = 0 Then
                    mUnitID = 1
                Else
                    mUnitID = Val(grdItems.TextMatrix(i, grdItems.ColIndex("UnitId")))
                End If
                UnitFactor = 1
        
AddNew:
                rsItemsUnits.AddNew
                rsItemsUnits!UnitID = mUnitID
                rsItemsUnits!ItemID = Val(grdItems.TextMatrix(i, grdItems.ColIndex("ItemID")))
                rsItemsUnits!UnitFactor = UnitFactor
                ' ,UnitSalesPrice,UnitPurPrice,FactorByDefaultUnit,FactorBySmallUnit
        
                rsItemsUnits!UnitSalesPrice = Val(grdItems.TextMatrix(i, grdItems.ColIndex("UnitSalesPrice"))) * UnitFactor
                rsItemsUnits!UnitPurPrice = Val(grdItems.TextMatrix(i, grdItems.ColIndex("UnitPurPrice"))) * UnitFactor
                rsItemsUnits!FactorByDefaultUnit = 1
                If mSecondEntry = 0 Then
                    rsItemsUnits!FactorByDefaultUnit = 1
                    rsItemsUnits!DefaultUnit = 1
                    If Val(grdItems.TextMatrix(i, grdItems.ColIndex("UnitId2"))) <> 0 Then
                        rsItemsUnits!FactorBySmallUnit = Val(grdItems.TextMatrix(i, grdItems.ColIndex("UnitFactor2")))
                    Else
                        rsItemsUnits!FactorBySmallUnit = 1
                    End If
                Else
                    rsItemsUnits!DefaultUnit = 0
                End If
                rsItemsUnits.Update
                If mSecondEntry = 1 Then GoTo NextRow
                If Val(grdItems.TextMatrix(i, grdItems.ColIndex("UnitId2"))) <> 0 And Val(grdItems.TextMatrix(i, grdItems.ColIndex("UnitFactor2"))) <> 0 Then
                    mUnitID = Val(grdItems.TextMatrix(i, grdItems.ColIndex("UnitId2")))
                    UnitFactor = Val(grdItems.TextMatrix(i, grdItems.ColIndex("UnitFactor2")))
                    mSecondEntry = 1
                    GoTo AddNew
                End If
NextRow:
                mSecondEntry = 0
            Next
            '    s = "Select * from TblItemsUnits Where ItemID =  -1"
            '    saveGridExcel s, mGrd, "ItemID", "ItemID ", "TblItemsUnits"
            Command4.Enabled = True
    
            s = " Update TblItems"
            s = s & " SET prifix = (SELECT TOP 1  g.Fullcode FROM  Groups AS g WHERE g.GroupID = TblItems.GroupID )"
            s = s & " ,Code =  (SELECT TOP 1  REPLACE(TblItems.code,g.Fullcode,'') FROM  Groups AS g WHERE g.GroupID = TblItems.GroupID )"

            s = s & " WHERE ISNULL(TblItems.prifix,'') = ''"
            Cn.Execute s
            s = " Update TblItems Set ItemCode = FullCode where IsNull(ItemCode,'') = ''  "
            Cn.Execute s
            's = " UPDATE tblItemsUnits SET DefaultUnit = 1,UnitFactor = 1"
            'Cn.Execute s

            s = " UPDATE Groups SET ParentID = 1  WHERE ISNULL(ParentID,0) = 0  and GroupID <> 1"
            Cn.Execute s
 
            s = " UPDATE groups SET Code = fullCode WHERE ISNULL(code,'') = ''"
            Cn.Execute s
 
        Case 5
    
            s = "Select * from GroupsCustomers Where GroupID =  -1"
            saveGridExcel s, mGrd, "Fullcode", "GroupID", "GroupsCustomers"

        Case 7
    
            s = "Select * from TblBoxesData Where BoxID =  -1"
            saveGridExcel s, mGrd, "BoxName", "BoxID", "TblBoxesData"

        Case 18

            s = "Select * from EmpTempData Where ID =  '-1'"
            saveGridExcel s, mGrd, "Name", "ID", "EmpTempData"

        Case 9
            SaveAccount
        Case 11
            SaveAccount2
        Case 12
    
            s = "Select * from FixedAssetsGroup Where GroupID =  -1"
            saveGridExcel s, mGrd, "GroupName", "GroupID", "FixedAssetsGroup"
    
            s = " UPDATE FixedAssetsGroup SET ParentID = 1  WHERE ISNULL(ParentID,0) = 0  and GroupID <> 1"
            Cn.Execute s
 
            s = " UPDATE FixedAssetsGroup SET Code = fullCode WHERE ISNULL(code,'') = ''"
            Cn.Execute s
 
            SaveFAGropsAccount
        Case 13
            s = "Select * from FixedAssets Where ID =  -1"
            saveGridExcel s, mGrd, "code", "ID", "FixedAssets"
            ' s = " UPDATE FixedAssets SET code = Fullcode"
            ' Cn.Execute s
            'New_or_opening = 1
            s = " UPDATE FixedAssets SET OpenBalanceType = 1 WHERE IsNull(OpenBalance,0) < 0"
            ' Cn.Execute s
            SaveFA
        Case 15
  
            Dim rsDummy As New ADODB.Recordset
            For i = 1 To GrdEmpFee.Rows - 1
                s = "Select * from TblEmployee Where Emp_Code = N'" & Trim(GrdEmpFee.TextMatrix(i, GrdEmpFee.ColIndex("Emp_Code"))) & "'"
                Set rsDummy = New ADODB.Recordset
                rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly
                If Not rsDummy.EOF Then
                    '  GrdEmpFee.TextMatrix(i, GrdEmpFee.ColIndex("BranchID")) = Val(rsDummy!BranchID & "")
                    GrdEmpFee.TextMatrix(i, GrdEmpFee.ColIndex("Emp_ID")) = Val(rsDummy!Emp_id & "")
                    GrdEmpFee.TextMatrix(i, GrdEmpFee.ColIndex("Account_Code2")) = Trim(rsDummy!Account_code2 & "")
                    GrdEmpFee.TextMatrix(i, GrdEmpFee.ColIndex("Account_Code4")) = Trim(rsDummy!Account_code4 & "")
                    GrdEmpFee.TextMatrix(i, GrdEmpFee.ColIndex("Account_Code5")) = Trim(rsDummy!Account_Code5 & "")
                End If
            Next
            s = "Select * from empfees Where ID =  -1"
            saveGridExcel s, mGrd, "Emp_Code", "Id", "empfees"
        Case 16
            s = "Select * from TblSchooleFile Where ID =  -1"
            saveGridExcel s, mGrd, "Name", "ID", "TblSchooleFile"

        Case 17
            s = "Select * from TblCarsData Where ID =  -1"
            saveGridExcel s, mGrd, "code", "ID", "TblCarsData"

            s = " UPDATE TblCarsData SET Branch_NO = 1"

            Cn.Execute s

            s = "UPDATE TblCarsData SET fixedAssetid = (SELECT id FROM FixedAssets AS fa WHERE fa.Name =  TblCarsData.BoardNO)"
            Cn.Execute s
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

Private Sub SaveFAGropsAccount()
Dim s As String
Dim i As Long
Dim rsDummy As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim mCode As String
Dim mName As String
Dim mNameP As String
Dim mSql As String
Dim StrNewAccountCode  As String
Dim mBranchName As String
Dim mBranchId As Long
Dim my_branch As String
mBranchId = 1
my_branch = mBranchId
'Dim mNewId As Long
Dim mLast As Boolean
For i = 1 To grdFAGroups.Rows - 1
    mCode = Trim(grdFAGroups.TextMatrix(i, grdFAGroups.ColIndex("FullCode")))
    mName = Trim(grdFAGroups.TextMatrix(i, grdFAGroups.ColIndex("GroupName")))
    mLast = CBool(grdFAGroups.ValueMatrix(i, grdFAGroups.ColIndex("LastGroup")))
    DepType = CBool(grdFAGroups.ValueMatrix(i, grdFAGroups.ColIndex("DepType")))
    mNewId = Val(grdFAGroups.TextMatrix(i, grdFAGroups.ColIndex("NewID")))
    
       s = "Select * from FixedAssetsGroup Where GroupID = " & mNewId
    Set rs = New ADODB.Recordset
    rs.Open s, Cn, adOpenKeyset, adLockOptimistic
    If rs.EOF Then GoTo ErrTrap
    
    If rs!ParentEAssetAccount & "" = "" Then
        Account_Code_dynamic = get_account_code_branch(24, my_branch)
        
        
        
        If Account_Code_dynamic = "NO branch" Then
            MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
            GoTo ErrTrap
        Else
    
            If Account_Code_dynamic = "NO account" Then
                MsgBox "·„ Ì „  ÕœÌœ Õ”«»   ﬁÌ„… «·«’Ê· «·À«» … ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                GoTo ErrTrap
             
            End If
        End If
           
         rs!ParentEAssetAccount = Account_Code_dynamic
           
        
    End If
    
    
        If rs!ParentExpensesAccount & "" = "" Then
        Account_Code_dynamic = get_account_code_branch(25, my_branch)
        
        
        
        If Account_Code_dynamic = "NO branch" Then
            MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
            GoTo ErrTrap
        Else
    
            If Account_Code_dynamic = "NO account" Then
            MsgBox "·„ Ì „  ÕœÌœ      Õ”«» „’—Ê› «·«Â·«ﬂ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                GoTo ErrTrap
             
            End If
        End If
           
         rs!ParentExpensesAccount = Account_Code_dynamic
           
        
    End If
    rs.Update
    
    'mNameP = Trim(grdFAGroups.TextMatrix(i, grdFAGroups.ColIndex("AccountNamePar")))
    'mBranchName = Trim(grdFAGroups.TextMatrix(i, grdFAGroups.ColIndex("BranchName")))
    'Set rsDummy = New ADODB.Recordset
    's = "Select branch_id from TblBranchesData Where branch_name Like '%" & mBranchName & "'"
    
    'rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly
    'If Not rsDummy.EOF Then
    '    mBranchID = Val(rsDummy!branch_id & "")
    'Else
    '    mBranchID = 1
    'End If
    
    
'
'    sql = " select * from ACCOUNTS Where Account_Serial = '" & mCode & "' Or Account_Name Like '" & mNameP & "' ORDER BY Parent_Account_Code Desc"
'    Set rs = New ADODB.Recordset
'    rs.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    If Not rs.EOF Then
'        StrNewAccountCode = AddNewAccount(Trim(rs!Account_Code & ""), Trim$(mName), True, False, Trim$(mName), 1, False, False, False, , "", 0, 0, 0, 1, 0, IIf(rs("DepitOrCredit").Value = 0, 0, 1), 0, 0, 0, 1, False, , True, mBranchID)
'        SaveBransh_UserAccount StrNewAccountCode
'        'mSql = GetSqlQueryInsert(rs, ServerDb, "ACCOUNTS", "Account_ID", "", "", 0, 0, True)
'    End If
'
 
  If mLast Then '«· √ﬂÌœ  ⁄·Ï «·Õ”«»«  ›Ì Õ«·Â „Ã„Ê⁄Â ‰Â«∆Ì… ·Â« «Â·«ﬂ
            If create_accounts(CInt(mNewId), mName, True, DepType) = False Then
                Exit Sub
            End If
        End If
  
  
  
ErrTrap:

Next

End Sub



Private Sub SaveFA()
Dim s As String
Dim i As Long
Dim rsDummy As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim mCode As String
Dim mName As String
Dim mNameP As String
Dim mSql As String
Dim StrNewAccountCode  As String

'Â‰« »Ì»œ√ ÌÕ”» «·ﬁÌ„Â «·œ› —Ì… ··«’· Ê„Ã„⁄ «·«Â·«ﬂ Ê⁄œœ «·«ﬁ”««ÿ «·„‰›–Â ÊﬁÌ„Â «·ﬁ”ÿ Ê„ »ﬁÌ  ﬂ«„ ﬁ”ÿ Ê⁄œœ «·«ﬁ”«ÿ «·«Ã„«·ÌÂ
        

        TxtnoOfInst = noOfInstallments
        TxtAge = Age

            TxtCurrentValue = currentvalue '«·ﬁÌ„Â «·œ› —Ì…
            txtinstallValue = installValue 'ﬁÌ„Â «·ﬁ”ÿ
            txtinstallmentresult = RemainInstallments '«·«ﬁ”«ÿ «·„ »ﬁÌ…
            txtinstallDo = Val(EXEInstallments) '«·«ﬁ”«ÿ «·„‰›–…







For i = 1 To grdFa.Rows - 1
    mCode = Trim(grdFa.TextMatrix(i, grdFa.ColIndex("Code")))
    mName = Trim(grdFa.TextMatrix(i, grdFa.ColIndex("Name")))
    'mLast = CBool(grdFa.TextMatrix(i, grdFa.ColIndex("LastGroup")))
    'DepType = CBool(grdFa.TextMatrix(i, grdFa.ColIndex("DepType")))
    
    mNewId = Val(grdFa.TextMatrix(i, grdFa.ColIndex("NewID")))
    mPurchasePrice = Val(grdFa.TextMatrix(i, grdFa.ColIndex("PurchasePrice")))
    mAccDepreciation = Val(grdFa.TextMatrix(i, grdFa.ColIndex("AccDepreciation")))
    TxtPurchasePrice = mPurchasePrice
    TxtAccDepreciation = mAccDepreciation
    
     s = "Select * from FixedAssets Where ID = " & mNewId
    Set rs = New ADODB.Recordset
    rs.Open s, Cn, adOpenKeyset, adLockOptimistic
    If rs.EOF Then GoTo ErrTrap
    DCGroup = Val(rs!group_id & "")
    mBranchId = Val(rs!branch_no & "")
    If mBranchId = 0 Then mBranchId = 1
    
    
    rs!New_or_opening = 1
    rs("ISEQUP").Value = 0
    
    DCPreFix = Val(GetPrefix(Val(DCGroup), "FixedAssetsGroup"))

   Dim DepType As Integer
 
'Â‰« »ÌÃÌ» Õ”«»«  «·„Ã„Ê⁄Â Ê‰”» «·«Â·«ﬂ Ê»ÌÕ”» ⁄„—  «·«’· »«·‘Â—
    GetFixedAssetsGroupAccount Val(DCGroup), , Val(mBranchId), , , Percentage1, Percentage2, DepType, Account_code, Account_code1, Account_code2, Account_code3, Account_code4
 'Â‰« »ÌÃÌ» Õ”«»«  «·„Ã„Ê⁄Â
    TXT24 = Get_Account_name(, Account_code)
    TXT26 = Get_Account_name(, Account_code2)
    TXT25 = Get_Account_name(, Account_code1)
    TXT31 = Get_Account_name(, Account_code3)
    TXT40 = Get_Account_name(, Account_code4)
    TXtPercentage1 = Percentage1
    txtPercentage2 = Percentage2
  If TXtPercentage1 <> 0 Then
    TxtAge = Round(100 / Val(TXtPercentage1) * 12, 0)
End If
  
    
  If DepType = 1 Then ' Â· «·«’· ·Â «Â·«ﬂ
        Opt(0).Value = True ' ·Â «Â·«ﬂ
        
    Else
        Opt(1).Value = True ' ·Ì” ·Â «Â·«ﬂ
        
    End If

       
        If Opt(1).Value = True Then
            TxtCurrentValue = mPurchasePrice
            txtinstallValue = 0
            txtinstallmentresult = 0
            txtinstallDo = 0
            TxtAccDepreciation = 0
            TxtKhordaPrice = 0
        End If
        
 
 
    
    GetAndCalculateAll CInt(mNewId), Val(TXtPercentage1), noOfInstallments, Age, Val(mPurchasePrice), Val(TxtKhordaPrice), Val(TxtAccDepreciation), TxtCurrentValue, installValue, EXEInstallments, RemainInstallments

    TxtnoOfInst = noOfInstallments
    rs("CurrentValue").Value = IIf(Val(TxtCurrentValue) = 0, 0, Val(TxtCurrentValue))
       ' rs("AccDepreciation").Value = IIf(Val(TxtAccDepreciation.Text) = 0, 0, Val(TxtAccDepreciatio))
        'rs("Status_id").value = GetStatus_id
       ' rs("Status_id").Value = cStatus.ListIndex
       ' rs("Depreciation_Type_id").Value = CBoDepreciation_Type_id.ListIndex
        'rs("DefaultAge").value = GetDefaultAge
        rs("DefaultAge").Value = Val(TxtAge)
        
        rs("Fullcode").Value = DCPreFix & mCode
        rs("prifix").Value = DCPreFix

        If Me.Opt(0).Value = True Then
            rs("HaveDepreciation").Value = 1
        Else
        
            rs("HaveDepreciation").Value = 0
        End If
        
          If Option1.Value = True Then
            txtopening_balance_voucher_id = 0
        End If
         rs!Status_id = 0
        
          'TxtCurrentValue = currentvalue '«·ﬁÌ„Â «·œ› —Ì…
            txtinstallValue = installValue 'ﬁÌ„Â «·ﬁ”ÿ
            txtinstallmentresult = RemainInstallments '«·«ﬁ”«ÿ «·„ »ﬁÌ…
            txtinstallDo = Val(EXEInstallments) '«·«ﬁ”«ÿ «
            txtopening_balance_voucher_id = get_opening_balance_voucher_id
        


   rs("NoOfInstallments").Value = Val(TxtnoOfInst)
        rs("RemainInstallments").Value = Val(txtinstallmentresult)
        rs("InstallmentValue").Value = Val(txtinstallValue)
        rs("EXEInstallments").Value = Val(txtinstallDo)

        '   Dim PurchasePrice As Double
        '    Dim PurchaseDate As Data
        '    Dim PurchaseBillId As String
        '
        '    getPurchaseInformations Val(Me.XPTxtID), PurchaseDate, PurchasePrice, PurchaseBillId
        '     If Option1.value = True Then
        '     rs("PurchasePrice").value = PurchasePrice
        '     rs("PurchaseDate").value = PurchaseDate
        ''     rs("PurchaseBillId") = PurchaseBillId
        '     Else
        'rs("PurchasePrice").value = Val(TxtPurchasePrice.text)
        'rs("PurchaseDate").value = DpPurchaseDate.value
        'rs("PurchaseBillId") = txtPurchaseBillId.text
        'End If
        rs("PurchasePrice").Value = Val(TxtPurchasePrice)
        'rs("PurchaseDate").Value = DpPurchaseDate.Value
        rs("PurchaseBillId") = 0 'txtPurchaseBillId.Text
        rs("KhordaPrice").Value = 0 'IIf(Val(TxtKhordaPrice.Text) = 0, 0, Val(TxtKhordaPrice.Text))
        
        rs("opening_balance_voucher_id").Value = Val(txtopening_balance_voucher_id)
        rs!Depreciation_Type_id = 0
        getFirstPeriodDateInthisYear2 FirstPeriodDateInthisYear
        txtDate.Value = FirstPeriodDateInthisYear
        
        

'
 rs.Update
    updateFixedAsseTInstallmentInformations Val(mNewId), Val(TxtPurchasePrice), Val(TxtCurrentValue), , Me.txtDate.Value, mAccDepreciation, , , False, True
  
    If CreateJL(DCGroup, mBranchId, mNewId, mName, mPurchasePrice) = False Then '«‰‘«¡ «·ﬁÌÊœ «·«›  «ÕÌ…
            GoTo ErrTrap
        End If
   
  SaveAssest Val(mNewId), mCode, mName
    
Next
ErrTrap:

End Sub
Sub SaveAssest(Optional FexdID As Double = 0, Optional ByVal mCode As String = "", Optional ByVal mName As String = "")
Dim sql As String
Dim StrSQL As String
Dim Msg As String
Dim Rs5 As ADODB.Recordset
Set Rs5 = New ADODB.Recordset
sql = "Select * from TblAssestes where 1=-1"
Rs5.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
Rs5.AddNew
Rs5("CarsDataID").Value = Val(FexdID)
Rs5("FlgCarNotFixed").Value = 3
If SystemOptions.UserInterface = ArabicInterface Then
Msg = "„‰ „·› «·«’Ê·"
Else
Msg = "From Fixed Assest File"
End If
Rs5("AsFixedID").Value = FexdID
Rs5("AsDes").Value = Msg
If SystemOptions.UserInterface = ArabicInterface Then
Rs5("AsName").Value = mName
Else
Rs5("AsName").Value = mName
End If
Rs5("AsCode").Value = Val(mCode)
Rs5.Update

End Sub

Public Function updateFixedAsseTInstallmentInformations(FixedAsseId As Integer, _
                                                        Optional PurcahsePrice As Double, _
                                                        Optional currentvalue As Double, _
                                                        Optional ByRef InstallmentID As Integer, _
                                                        Optional ByRef InstallmentDate As Date, _
                                                        Optional ByRef AccDepreciation As Double, _
                                                        Optional ByRef Installmentvalue As Double, _
                                                        Optional ByRef RemainInstallments As Double, _
                                                        Optional NewAsset As Boolean, _
                                                        Optional First_Installment As Boolean)
    Dim KhordaPrice As Double
    Dim noOfInstallments As Double
    Dim delsql As String
    Dim InstallmentProduct As Integer

    If NewAsset = True And First_Installment = True Then    'ÃœÌœ «Ê· ﬁ”ÿ
        delsql = "Delete FixedAssetInstallmentsDetails where FixedAssetID=" & FixedAsseId & "and InstallmentID=0"
        Cn.Execute delsql
        GetAllDataAboutFixedAsset FixedAsseId, , , , , , currentvalue, AccDepreciation, , , , , , noOfInstallments, , RemainInstallments, PurcahsePrice, , , KhordaPrice, Installmentvalue
        InstallmentID = 0
        InstallmentProduct = 0
        AccDepreciation = 0
        RemainInstallments = noOfInstallments
        Installmentvalue = Round((PurcahsePrice - KhordaPrice) / noOfInstallments, 2)
        Installmentvalue = 0
        AddInstallment 0, FixedAsseId, currentvalue, InstallmentID, InstallmentDate, AccDepreciation, Installmentvalue, RemainInstallments, InstallmentProduct
    ElseIf NewAsset = True And First_Installment = False Then 'ÃœÌœ Ê·Ì” «Ê· ﬁ”ÿ
   
    ElseIf NewAsset = False And First_Installment = True Then ' «›  «ÕÌ «Ê· ﬁ”ÿ
        delsql = "Delete FixedAssetInstallmentsDetails where FixedAssetID=" & FixedAsseId & "and InstallmentID=0"
        Cn.Execute delsql
        GetAllDataAboutFixedAsset FixedAsseId, , , , , , currentvalue, AccDepreciation, , , , , InstallmentDate, noOfInstallments, , RemainInstallments, PurcahsePrice, , , KhordaPrice, Installmentvalue
        InstallmentID = 0
        Installmentvalue = AccDepreciation
        InstallmentProduct = noOfInstallments - RemainInstallments
        AddInstallment 0, FixedAsseId, currentvalue, InstallmentID, InstallmentDate, AccDepreciation, Installmentvalue, RemainInstallments, InstallmentProduct
    ElseIf NewAsset = False And First_Installment = False Then ' «›  «ÕÌÊ·Ì”  «Ê· ﬁ”ÿ
    
    End If
   
End Function
 Public Function AddInstallment(FixedAssetInstallmentsid As Integer, _
                               FixedAsseId As Integer, _
                               Optional currentvalue As Double, _
                               Optional InstallmentID As Integer, _
                               Optional InstallmentDate As Date, _
                               Optional ByRef AccDepreciation As Double, _
                               Optional ByRef Installmentvalue As Double, _
                               Optional ByRef RemainInstallments As Double, _
                               Optional ByRef InstallmentProduct As Integer)
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    rs.Open "FixedAssetInstallmentsDetails", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    rs.AddNew
    rs("FixedAssetInstallmentsid").Value = FixedAssetInstallmentsid
    rs("FixedAssetID").Value = FixedAsseId
    rs("CurrentValue").Value = currentvalue
    rs("InstallmentID").Value = InstallmentID
    rs("InstallmentValue").Value = Installmentvalue
    rs("InstallmentDate").Value = InstallmentDate
    rs("AccDepreciation").Value = AccDepreciation
    rs("RemainInstallments").Value = RemainInstallments
    rs("month").Value = Month(InstallmentDate)
    rs("year").Value = year(InstallmentDate)
    rs("InstallmentProduct").Value = InstallmentProduct

    rs.Update
 
End Function

 
Public Function GetAllDataAboutFixedAsset(FixedAsseId As Integer, _
                                          Optional ByRef Name As String, _
                                          Optional ByRef group_id As Integer, _
                                          Optional ByRef branch_no As Integer, _
                                          Optional ByRef Emp_id As Integer, _
                                          Optional ByRef ReceiveDate As Date, _
                                          Optional ByRef currentvalue As Double, _
                                          Optional ByRef AccDepreciation As Double, _
                                          Optional ByRef Status_id As Integer, _
                                          Optional ByRef Depreciation_Type_id As Integer, _
                                          Optional ByRef DefaultAge As Integer, _
                                          Optional ByRef StartDepreciationDate As Date, _
                                          Optional ByRef LastDepreciationDate As Date, _
                                          Optional ByRef noOfInstallments As Double, _
                                          Optional ByRef EXEInstallments As Double, _
                                          Optional ByRef RemainInstallments As Double, _
                                          Optional ByRef purchaseprice As Double, _
                                          Optional ByRef PurchaseDate As Date, _
                                          Optional ByRef PurchaseBillId As Double, _
                                          Optional ByRef KhordaPrice As Double, _
                                          Optional ByRef Installmentvalue As Double, _
                                          Optional ByRef New_or_opening As Integer, _
                                          Optional ByRef Notes As String, _
                                          Optional ByRef Fullcode As String, Optional DepitAccount As String, Optional CreditAccount As String, Optional Account_Code5 As String, Optional ParetnAccount As String, Optional GroupName As String)
  
    Dim sql As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
'    sql = "Select * from  FixedAssets where id=" & FixedAsseId
sql = "SELECT     dbo.FixedAssets.*,FixedAssets.fullcode as fafullcode, dbo.FixedAssetsGroup.*"
sql = sql & " FROM         dbo.FixedAssets INNER JOIN"
sql = sql & "  dbo.FixedAssetsGroup ON dbo.FixedAssets.group_id = dbo.FixedAssetsGroup.GroupID"
sql = sql & " WHERE     (dbo.FixedAssets.id = " & FixedAsseId & ")"
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        Name = IIf(IsNull(rs("Name").Value), "", rs("Name").Value)
        group_id = IIf(IsNull(rs("group_id").Value), 0, Val(rs("group_id").Value))
        branch_no = IIf(IsNull(rs("Branch_NO").Value & ""), 1, Val(rs("Branch_NO").Value & ""))
        Emp_id = IIf(IsNull(rs("Emp_id").Value), 0, (rs("Emp_id").Value))
        ReceiveDate = IIf(IsNull(rs("ReceiveDate").Value), Date, rs("ReceiveDate").Value)
        currentvalue = IIf(IsNull(rs("CurrentValue").Value), 0, Val(rs("CurrentValue").Value))
 DepitAccount = IIf(IsNull(rs("Account_Code1").Value), "", (rs("Account_Code1").Value))
 CreditAccount = IIf(IsNull(rs("Account_Code2").Value), "", (rs("Account_Code2").Value))
 Account_Code5 = IIf(IsNull(rs("Account_Code5").Value), "", (rs("Account_Code5").Value))
 ParetnAccount = IIf(IsNull(rs("ParetnAccount").Value), "", (rs("ParetnAccount").Value))
 
 GroupName = IIf(IsNull(rs("GroupName").Value), "", (rs("GroupName").Value))
 GroupNamee = IIf(IsNull(rs("GroupNamee").Value), "", (rs("GroupNamee").Value))
 
        AccDepreciation = Val(rs!AccDepreciation & "")

        Status_id = IIf(IsNull(rs("Status_id").Value), 0, Val(rs("Status_id").Value))
        Depreciation_Type_id = IIf(IsNull(rs("Depreciation_Type_id").Value), 0, Val(rs("Depreciation_Type_id").Value))
        DefaultAge = IIf(IsNull(rs("DefaultAge").Value), 0, Val(rs("DefaultAge").Value))
        StartDepreciationDate = IIf(IsNull(rs("StartDepreciationDate").Value), Date, rs("StartDepreciationDate").Value)
        LastDepreciationDate = IIf(IsNull(rs("LastDepreciationDate").Value), Date, rs("LastDepreciationDate").Value)
        noOfInstallments = IIf(IsNull(rs("NoOfInstallments").Value), 0, Val(rs("NoOfInstallments").Value))
        EXEInstallments = IIf(IsNull(rs("EXEInstallments").Value), 0, Val(rs("EXEInstallments").Value))
        RemainInstallments = IIf(IsNull(rs("RemainInstallments").Value), 0, Val(rs("RemainInstallments").Value))
        purchaseprice = IIf(IsNull(rs("PurchasePrice").Value), 0, Val(rs("PurchasePrice").Value))
        PurchaseDate = IIf(IsNull(rs("PurchaseDate").Value), Date, (rs("PurchaseDate").Value))
        PurchaseBillId = IIf(IsNull(rs("PurchaseBillId").Value), 0, Val(rs("PurchaseBillId").Value))
        KhordaPrice = IIf(IsNull(rs("KhordaPrice").Value), 0, Val(rs("KhordaPrice").Value))
        Installmentvalue = IIf(IsNull(rs("InstallmentValue").Value), 0, Val(rs("InstallmentValue").Value))
        New_or_opening = IIf(IsNull(rs("New_or_opening").Value), 0, Val(rs("New_or_opening").Value))
      Fullcode = IIf(IsNull(rs("fafullcode").Value), "", rs("fafullcode").Value)
        Notes = IIf(IsNull(rs("Notes").Value), "", rs("Notes").Value)
    End If

End Function
Public Function CheCkInstallmentCount(FixedassetId As Integer) As Integer
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String
    sql = "Select  count (FixedAssetID ) As InstallmentCount from FixedAssetInstallmentsDetails where FixedAssetID=" & FixedassetId
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then CheCkInstallmentCount = 0: Exit Function
    If IsNull(Rs3("InstallmentCount").Value) Then CheCkInstallmentCount = 0: Exit Function
    If Not IsNull(Rs3("InstallmentCount").Value) Then CheCkInstallmentCount = Rs3("InstallmentCount").Value - 1: Exit Function
    Rs3.Close

End Function

Public Function CheckLastInstallmentDate(Month As Integer, _
                                         year As Integer, Optional BranchID As Integer) As Boolean
    CheckLastInstallmentDate = False
    Dim sql As String
    Dim rs As ADODB.Recordset

  '  sql = "Select max(Month) As LastMonth  From FixedAssetInstallments where year =" & year
    sql = "Select max(Month) As LastMonth  From FixedAssetInstallments where year(RecordDate) =" & year
    
  '
    sql = sql & "  and BranchId=" & BranchID
    
    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount = 0 Then
        CheckLastInstallmentDate = True
    ElseIf IsNull(rs("LastMonth").Value) Then
        CheckLastInstallmentDate = True
    ElseIf Month - rs("LastMonth").Value > 1 Then
        CheckLastInstallmentDate = False
    ElseIf Month - rs("LastMonth").Value = 1 Then
        CheckLastInstallmentDate = True
    ElseIf rs("LastMonth").Value - Month >= 0 Then
        CheckLastInstallmentDate = False

        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·«Ì„ﬂ‰ «’œ«— «ﬁ”«ÿ ·Â–« «·‘Â—  ·«‰Â „’œ—„‰ ﬁ»·"
            MsgBox Msg, vbInformation
        Else
            Msg = "Cant Create Depreciation Installment For this Month , already Created"
            MsgBox Msg, vbInformation
        End If

        Exit Function
    End If

    If CheckLastInstallmentDate = False Then

        'CboYear.ListIndex = -1
        'CmbMonth.ListIndex = -1
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·«Ì„ﬂ‰ «’œ«— «ﬁ”«ÿ ·Â–« «·‘Â— ÌÊÃœ «ﬁ”«ÿ ”«»ﬁ… €Ì— „’œ—…"
            MsgBox Msg, vbInformation
        Else
            Msg = "Cant Create Depreciation Installment For this Month , Check Last  Installment Date"
            MsgBox Msg, vbInformation
        End If

    End If

End Function


Public Function getFirstPeriodDateInthisYear2(Optional ByRef FirstPeriodDateInthisYear As Date)
 
    Dim rs As ADODB.Recordset
    Dim sql As String
 
    sql = "SELECT     Min(OpeneingbalancesDate) AS OpeningBalanceDate FROM         dbo.TblyearsData"
 
    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If rs.RecordCount > 0 Then
        FirstPeriodDateInthisYear = IIf(IsDate(rs("OpeningBalanceDate").Value), rs("OpeningBalanceDate").Value, Date)
    End If

End Function



Public Function GetPrefix(group_id As Double, _
                          Optional tablename As String) As String
    On Error Resume Next
    Dim fullCodeAll As String
    
    Dim GroupCode As String
    Dim ParentGroupCode As String
    Dim ParentId As Double
    fullCodeAll = ""
    txtid.Text = ""

 If SystemOptions.WorkWithBarCodeParent = False Then
            GetGroupData group_id, GroupCode, , ParentGroupCode, ParentId, tablename
            GetPrefix = GroupCode
            Exit Function
 End If
 
 If SystemOptions.WorkWithBarCodeParent = True Then
 
      GetGroupData group_id, GroupCode, , ParentGroupCode, ParentId, tablename
      
     If group_id = 0 Or group_id = 1 Then
GetPrefix = GroupCode
 Exit Function
End If
 

  '      fullCodeAll = fullCodeAll & SystemOptions.itemSeprator & GroupCode
         GetPrefix = GetPrefix(ParentId, tablename) & SystemOptions.itemSeprator & GroupCode
    

  
 
 End If
 
    

 
End Function

Public Function GetGroupData(GroupID As Double, _
                             Optional ByRef GroupCode As String, _
                             Optional ByRef GroupName As String, _
                             Optional ByRef ParentGroupCode As String, _
                             Optional ByRef ParentId As Double, _
                             Optional tablename As String, _
                             Optional ByRef EXpirType As Integer, _
                             Optional ByRef EXpireValue As Integer, Optional ByRef OverHead As Double)
    Dim rs As ADODB.Recordset
    Dim Rs1 As ADODB.Recordset

    Dim sql As String
    Dim str As String
    If GroupID = 1 Or GroupID = 0 Then Exit Function
    Set Rs1 = New ADODB.Recordset
    sql = "SELECT * from " & tablename & " where GroupID=" & GroupID
    Rs1.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Rs1.RecordCount > 0 Then
        GroupCode = IIf(IsNull(Rs1("Fullcode").Value), "", Rs1("Fullcode").Value)
        GroupName = IIf(IsNull(Rs1("GroupName").Value), "", Rs1("GroupName").Value)
        ParentId = IIf(IsNull(Rs1("ParentID").Value), 0, Rs1("ParentID").Value)
        EXpirType = IIf(IsNull(Rs1("EXpirType").Value), -1, Rs1("EXpirType").Value)
        EXpireValue = IIf(IsNull(Rs1("EXpireValue").Value), -1, Rs1("EXpireValue").Value)
        
        
         
        
        OverHead = IIf(IsNull(Rs1("OverHead").Value), 0, Rs1("OverHead").Value)
                
    End If

    Rs1.Close
   Exit Function ' xxxxxxxxxxxxxxxxxxxxxxxx
     If ParentId = 1 Then Exit Function
    sql = "SELECT * from Groups where GroupID=" & ParentId
    Rs1.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Rs1.RecordCount > 0 Then
        ParentGroupCode = IIf(IsNull(Rs1("Fullcode").Value), "", Rs1("Fullcode").Value)
       ParentId = IIf(IsNull(Rs1("ParentID").Value), 0, Rs1("ParentID").Value)
    End If
Rs1.Close
 
End Function




Public Function get_account_code_branch(account_index As Integer, _
                                        branch_id As String) As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim sql As String
    Dim AccountName As String
    branch_id = 1
    AccountName = "a" & account_index
    sql = "Select * from branches " 'where branch_id='" & branch_id & "'"
 
    rs.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If rs.RecordCount = 0 Then get_account_code_branch = "NO branch": Exit Function
    rs.MoveFirst
    
    If IsNull(rs(AccountName).Value) Or rs(AccountName).Value = "" Then get_account_code_branch = "NO account": Exit Function
  
    If Not IsNull(rs(AccountName).Value) Then
        If CheckAccountToJE(rs(AccountName).Value) = True Then
            get_account_code_branch = rs(AccountName).Value: Exit Function
        Else
            get_account_code_branch = "NO account": Exit Function
        End If
  
    End If
  
    rs.Close
End Function

Public Function CheckAccountToJE(Account_code As String) As Boolean
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim sql As String
    Dim AccountName As String
    branch_id = 1
    AccountName = "a" & account_index
    sql = "Select * from ACCOUNTS where   Account_Code='" & Account_code & "'"
 
    rs.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If rs.RecordCount = 0 Then CheckAccountToJE = False: Exit Function
  
    If IsNull(rs("Account_Code").Value) Or rs("Account_Code").Value = "" Then CheckAccountToJE = False: Exit Function
  
    If Not IsNull(rs("Account_Code").Value) Then
        CheckAccountToJE = True: Exit Function
  
    End If
  
    rs.Close

End Function
Function create_accounts(group_id As Integer, group_name As String, Optional Checkonly As Boolean = False, Optional ByVal DepType As Boolean = False) As Boolean
    Dim rsOut As New ADODB.Recordset
    Dim Current_case As Integer
    Dim Rs3 As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String
    Dim i As Integer
    s = "Select * from FixedAssetsGroup Where GroupID = " & group_id
    Set rs = New ADODB.Recordset
    rs.Open s, Cn, adOpenKeyset, adLockOptimistic
    Dim my_branch As String
    my_branch = 1
 
    Dim Account_Code_dynamic As String
        If Trim(rs!ParentEAssetAccount & "") = "" Then
    Account_Code_dynamic = get_account_code_branch(24, my_branch)
        
    If Account_Code_dynamic = "NO branch" Then
        MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
        GoTo ErrTrap
    Else

        If Account_Code_dynamic = "NO account" Then
            MsgBox "·„ Ì „  ÕœÌœ Õ”«»   ﬁÌ„… «·«’Ê· «·À«» … ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
            GoTo ErrTrap
         
        End If
    End If
       Else
       Account_Code_dynamic = Trim(rs!ParentEAssetAccount & "")
       
       End If
       
       
    If Trim(rs!ParentExpensesAccount & "") = "" Then
    Dim Account_Code_dynamic1 As String
    Account_Code_dynamic1 = get_account_code_branch(25, my_branch)
        
    If Account_Code_dynamic1 = "NO branch" Then
        MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
        GoTo ErrTrap
    Else

        If Account_Code_dynamic1 = "NO account" Then
            MsgBox "·„ Ì „  ÕœÌœ      Õ”«» „’—Ê› «·«Â·«ﬂ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
            GoTo ErrTrap
         
        End If
    End If
    
    Else
    
    Account_Code_dynamic1 = Trim(rs!ParentExpensesAccount & "")
    
    End If
    
    
    
    Dim Account_Code_dynamic2 As String
    Account_Code_dynamic2 = get_account_code_branch(26, my_branch)
        
    If Account_Code_dynamic2 = "NO branch" Then
        MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
        GoTo ErrTrap
    Else

        If Account_Code_dynamic2 = "NO account" Then
            MsgBox "·„ Ì „  ÕœÌœ   Õ”«» „Ã„⁄ «·«Â·«ﬂ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
            GoTo ErrTrap
         
        End If
    End If
        
    Dim Account_Code_dynamic3 As String
    Dim Account_Code_dynamic4 As String
           
    If SystemOptions.AssetAccount1 = True Then
        Account_Code_dynamic3 = get_account_code_branch(31, my_branch)
        
        If Account_Code_dynamic3 = "NO branch" Then
            MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
            GoTo ErrTrap
        Else

            If Account_Code_dynamic3 = "NO account" Then
                MsgBox "·„ Ì „  ÕœÌœ     Õ”«» «—»«Õ »Ì⁄ «.À«» … ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                GoTo ErrTrap
         
            End If
        End If
           
        Account_Code_dynamic4 = get_account_code_branch(40, my_branch)
        
        If Account_Code_dynamic4 = "NO branch" Then
            MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
            GoTo ErrTrap
        Else

            If Account_Code_dynamic4 = "NO account" Then
                MsgBox "·„ Ì „  ÕœÌœ  Õ”«» Œ”«—… »Ì⁄ «.À«» … ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                GoTo ErrTrap
         
            End If
        End If
    End If
        
    If Checkonly = True Then
        'GoTo ll
    End If
       
    Dim X As String

    If SystemOptions.AssetAccount = True Then
        X = AddNewAccount(Account_Code_dynamic, group_name, False, False, group_name)
        rs("ParetnAccount").Value = X
        rs("Account_Code").Value = AddNewAccount(X, " ﬁÌ„Â " & group_name, True, False, group_name)
If DepType Then
        rs("Account_Code2").Value = AddNewAccount(X, "  „Ã„⁄ «Â·«ﬂ   " & group_name, True, False, group_name & " Accumulated depreciation")
 End If
    Else
        rs("Account_Code").Value = AddNewAccount(Account_Code_dynamic, " ﬁÌ„Â " & group_name, True, False, group_name & "Value")
       If DepType Then
        rs("Account_Code2").Value = AddNewAccount(Account_Code_dynamic2, "   „Ã„⁄ «Â·«ﬂ   " & group_name, True, False, group_name & " Accumulated depreciation")
        End If
    End If
     If DepType Then
    rs("Account_Code1").Value = AddNewAccount(Account_Code_dynamic1, "  „’—Ê›«    " & group_name, True, False, group_name & " Expenses ")
       End If
       
    If SystemOptions.AssetAccount1 = True Then
        rs("Account_Code3").Value = AddNewAccount(Account_Code_dynamic3, "  «—»«Õ »Ì⁄   " & group_name, True, False, group_name & " Sale Profit ")
        rs("Account_Code4").Value = AddNewAccount(Account_Code_dynamic4, " Œ”«—… »Ì⁄   " & group_name, True, False, group_name & " Sale Loss ")
    End If
    rs.Update
    
        
ll:
  
    create_accounts = True
    Exit Function
ErrTrap:

    create_accounts = False

End Function



Private Sub SaveAccount()
Dim s As String
Dim i As Long
Dim rsDummy As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim mCode As String
Dim mName As String
Dim mNameP As String
Dim mSql As String
Dim StrNewAccountCode  As String
Dim mBranchName As String
Dim mBranchId As Long
Dim last_account As Boolean
Dim opening_balance As Double
Dim RSDDD As New ADODB.Recordset
For i = 1 To GrdAccount.Rows - 1
    mCode = Trim(GrdAccount.TextMatrix(i, GrdAccount.ColIndex("Account_Serial")))
    mName = Trim(GrdAccount.TextMatrix(i, GrdAccount.ColIndex("Account_Name")))
    mNameP = Trim(GrdAccount.TextMatrix(i, GrdAccount.ColIndex("AccountNamePar")))
    last_account = Trim(GrdAccount.ValueMatrix(i, GrdAccount.ColIndex("last_account")))
    If last_account Then
        last_account = False
    Else
        last_account = True
    End If
    If last_account = False Then
        last_account = False
        
    End If
    
    mCode = Trim(GrdAccount.TextMatrix(i, GrdAccount.ColIndex("Account_Serial")))
    
    opening_balance = Val(GrdAccount.TextMatrix(i, GrdAccount.ColIndex("opening_balance")))
    mBranchName = Trim(GrdAccount.TextMatrix(i, GrdAccount.ColIndex("BranchName")))
    If mBranchName <> "«·›—⁄ «·—∆Ì”Ì" Then
        mBranchName = mBranchName
    End If
    Set rsDummy = New ADODB.Recordset
    s = "Select branch_id from TblBranchesData Where branch_name Like '%" & mBranchName & "'"
    
    rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly
    If Not rsDummy.EOF Then
        mBranchId = Val(rsDummy!branch_id & "")
    Else
        mBranchId = 1
    End If
    
    sql = " select * from ACCOUNTS Where Account_Serial = '" & mCode & "' Or Account_Name Like '" & mNameP & "' ORDER BY Parent_Account_Code Desc"
    sql = " select * from ACCOUNTS Where  Account_Name Like '" & mNameP & "' ORDER BY Parent_Account_Code Desc"
    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
    
    
    If Not rs.EOF Then
        If CBool(rs!last_account & "") Then
            Cn.Execute "Update Accounts set last_account = 0 where Account_code = N'" & Trim(rs!Account_code & "") & "'"
        End If
11111:
        If last_account = False Then
            sql = " select * from ACCOUNTS Where  Account_Name Like '" & mName & "' and last_account = 0 ORDER BY Parent_Account_Code Desc"
            Set RSDDD = New ADODB.Recordset
            RSDDD.Open sql, Cn, adOpenStatic, adLockReadOnly
            If Not RSDDD.EOF Then
                GoTo 2222
            End If
            
        End If
        StrNewAccountCode = AddNewAccount(Trim(rs!Account_code & ""), Trim$(mName), last_account, False, Trim$(mName), 1, False, False, False, , "", 0, 0, 0, 1, 0, IIf(rs("DepitOrCredit").Value = 0, 0, 1), 0, 0, 0, 1, False, , True, , , mBranchId, opening_balance)
        
        SaveBransh_UserAccount StrNewAccountCode
2222:
        'mSql = GetSqlQueryInsert(rs, ServerDb, "ACCOUNTS", "Account_ID", "", "", 0, 0, True)
    End If
    
 
  
  
  
    
Next


End Sub


Private Sub SaveAccount2()
Dim s As String
Dim i As Long
Dim rsDummy As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim mCode As String
Dim mName As String
Dim mNameP As String

Dim mCode2 As String
Dim mName2 As String
Dim mNameP2 As String


Dim mCode3 As String
Dim mName3 As String
Dim mNameP3 As String


Dim mCode4 As String
Dim mName4 As String
Dim mNameP4 As String

Dim mCode5 As String
Dim mName5 As String
Dim mNameP5 As String
Dim AccountTypes As Integer
Dim AccountTab As Integer
Dim DepitOrCreditv As Integer
Dim Differenttypev As Integer
Dim Authorityv As Integer

Dim mDebitCredit As Boolean
Dim mLastAccount As Boolean
Dim mSql As String
Dim StrNewAccountCode  As String
Dim h As Long
For i = 1 To GrdAccount2.Rows - 1
 

    
      
    
    mCode = Trim(GrdAccount2.TextMatrix(i, GrdAccount2.ColIndex("Account_Serial")))
    If mCode = "" Then Exit Sub
    Select Case Val(mCode)
    Case 1
        AccountTypes = 1
        AccountTab = 0
        DepitOrCreditv = 1
        Differenttypev = 1
        Authorityv = 0
    Case 2
        AccountTypes = 1
        AccountTab = 1
        DepitOrCreditv = 1
        Differenttypev = 1
        Authorityv = 0
    Case 3
        AccountTypes = 2
        AccountTab = 2
        DepitOrCreditv = 1
        Differenttypev = 1
        Authorityv = 0
    Case 4
        AccountTypes = 2
        AccountTab = 3
        DepitOrCreditv = 1
        Differenttypev = 1
        Authorityv = 0
    Case 5
        AccountTypes = 0
        AccountTab = 4
        DepitOrCreditv = 1
        Differenttypev = 1
        Authorityv = 0
    End Select
    
    Dim mLevel As Integer
    mLastAccount = False
    mNameP = Trim(GrdAccount2.TextMatrix(i, GrdAccount2.ColIndex("AccountNamePar")))
    If mNameP = "«·Œ’Ê„" Then
        mNameP = "«·Œ’Ê„"
    End If
    mLevel = 1
    sql = " select * from ACCOUNTS Where Level = " & mLevel & " and (  Account_Name Like '" & mNameP & "')"
    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If rs.EOF Then
        StrNewAccountCode = AddNewAccount("r", Trim$(mNameP), True, False, Trim$(mNameP), 1, False, False, False, , "", 0, 0, 0, 1, 0, IIf(Not mDebitCredit, 0, 1), 0, 0, 0, 1, False, True, mLastAccount, Val(mCode), mLevel)
        SaveBransh_UserAccount StrNewAccountCode
        'mSql = GetSqlQueryInsert(rs, ServerDb, "ACCOUNTS", "Account_ID", "", "", 0, 0, True)
    Else
        StrNewAccountCode = Trim(rs!Account_code & "")
    End If
    
    
    

    
    mLevel = 2
    mCode2 = Trim(GrdAccount2.TextMatrix(i, GrdAccount2.ColIndex("Account_Serial2")))
    mNameP2 = Trim(GrdAccount2.TextMatrix(i, GrdAccount2.ColIndex("AccountNamePar2")))
    
    sql = " select * from ACCOUNTS Where Level = " & mLevel & " and (  Account_Name Like '" & mNameP2 & "') and Parent_Account_Code = N'" & StrNewAccountCode & "'"
    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If rs.EOF Then

    
        sql = " select * from ACCOUNTS Where Account_Code = '" & StrNewAccountCode & "' "
        Set rs = New ADODB.Recordset
        rs.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
        If Not rs.EOF Then
            StrNewAccountCode = AddNewAccount(Trim(rs!Account_code & ""), Trim$(mNameP2), True, False, Trim$(mNameP2), 1, False, False, False, , "", 0, 0, 0, 1, 0, IIf(rs("DepitOrCredit").Value = 0, 0, 1), 0, 0, 0, 1, False, False, False, , mLevel)
            SaveBransh_UserAccount StrNewAccountCode
            'mSql = GetSqlQueryInsert(rs, ServerDb, "ACCOUNTS", "Account_ID", "", "", 0, 0, True)
        End If
    Else
            StrNewAccountCode = Trim(rs!Account_code & "")
    End If
    
    mCode3 = Trim(GrdAccount2.TextMatrix(i, GrdAccount2.ColIndex("Account_Serial3")))
    mNameP3 = Trim(GrdAccount2.TextMatrix(i, GrdAccount2.ColIndex("AccountNamePar3")))
    mLevel = 3

    sql = " select * from ACCOUNTS Where Level = " & mLevel & " and (  Account_Name Like '" & mNameP3 & "') and Parent_Account_Code = N'" & StrNewAccountCode & "'"
    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If rs.EOF Then

        sql = " select * from ACCOUNTS Where Account_Code = '" & StrNewAccountCode & "' "
        Set rs = New ADODB.Recordset
        rs.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
        If Not rs.EOF Then
            StrNewAccountCode = AddNewAccount(Trim(rs!Account_code & ""), Trim$(mNameP3), True, False, Trim$(mNameP3), 1, False, False, False, , "", 0, 0, 0, 1, 0, IIf(rs("DepitOrCredit").Value = 0, 0, 1), 0, 0, 0, 1, False, False, False, , mLevel)
            SaveBransh_UserAccount StrNewAccountCode
            'mSql = GetSqlQueryInsert(rs, ServerDb, "ACCOUNTS", "Account_ID", "", "", 0, 0, True)
        End If
    Else
            StrNewAccountCode = Trim(rs!Account_code & "")
    End If
    
    mCode4 = Trim(GrdAccount2.TextMatrix(i, GrdAccount2.ColIndex("Account_Serial4")))
    mNameP4 = Trim(GrdAccount2.TextMatrix(i, GrdAccount2.ColIndex("AccountNamePar4")))
     mLevel = 4
    sql = " select * from ACCOUNTS Where Level = " & mLevel & " and (  Account_Name Like '" & mNameP4 & "') and Parent_Account_Code = N'" & StrNewAccountCode & "'"
    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If rs.EOF Then

           sql = " select * from ACCOUNTS Where Account_Code = '" & StrNewAccountCode & "' "
        Set rs = New ADODB.Recordset
        rs.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
        If Not rs.EOF Then
            StrNewAccountCode = AddNewAccount(Trim(rs!Account_code & ""), Trim$(mNameP4), True, False, Trim$(mNameP4), 1, False, False, False, , "", 0, 0, 0, 1, 0, IIf(rs("DepitOrCredit").Value = 0, 0, 1), 0, 0, 0, 1, False, False, False, , mLevel)
            SaveBransh_UserAccount StrNewAccountCode
            'mSql = GetSqlQueryInsert(rs, ServerDb, "ACCOUNTS", "Account_ID", "", "", 0, 0, True)
        End If
    Else
            StrNewAccountCode = Trim(rs!Account_code & "")
    End If
    
    
    mCode5 = Trim(GrdAccount2.TextMatrix(i, GrdAccount2.ColIndex("Account_Serial5")))
    mNameP5 = Trim(GrdAccount2.TextMatrix(i, GrdAccount2.ColIndex("AccountNamePar5")))
    mLevel = 5
    mLastAccount = CBool(GrdAccount2.TextMatrix(i, GrdAccount2.ColIndex("last_account")))
 
    sql = " select * from ACCOUNTS Where Level = " & mLevel & " and (  Account_Name Like '" & mNameP5 & "')"
    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
   ' If rs.EOF Then
    
        sql = " select * from ACCOUNTS Where Account_Code = '" & StrNewAccountCode & "' "
        Set rs = New ADODB.Recordset
        rs.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
        If Not rs.EOF Then
            StrNewAccountCode = AddNewAccount(Trim(rs!Account_code & ""), Trim$(mNameP5), True, False, Trim$(mNameP5), 1, False, False, False, , "", 0, 0, 0, 1, 0, IIf(rs("DepitOrCredit").Value = 0, 0, 1), 0, 0, 0, 1, False, False, mLastAccount, , mLevel)
            SaveBransh_UserAccount StrNewAccountCode
            'mSql = GetSqlQueryInsert(rs, ServerDb, "ACCOUNTS", "Account_ID", "", "", 0, 0, True)
            If StrNewAccountCode <> "" Then
            
                GrdAccount2.TextMatrix(i, GrdAccount2.ColIndex("IsCreated")) = 1
            Else
                GrdAccount2.TextMatrix(i, GrdAccount2.ColIndex("IsCreated")) = 0
                
            
            End If
        Else
            MsgBox "«·«» «·—«»⁄ ”ÿ— —ﬁ„" & i & "·„ Ì „ «‰‘«∆Â »—Ã«¡ „—«Ã⁄… «·„·›"
            GrdAccount2.TextMatrix(i, GrdAccount2.ColIndex("IsCreated")) = 0
        End If
        
   ' End If

  
    
Next


End Sub




Private Sub SaveAccount3()
Dim s As String
Dim i As Long
Dim rsDummy As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim mCode As String
Dim mName As String
Dim mNameP As String

Dim mCode2 As String
Dim mName2 As String
Dim mNameP2 As String


Dim mCode3 As String
Dim mName3 As String
Dim mNameP3 As String


Dim mCode4 As String
Dim mName4 As String
Dim mNameP4 As String

Dim mCode5 As String
Dim mName5 As String
Dim mNameP5 As String
Dim AccountTypes As Integer
Dim AccountTab As Integer
Dim DepitOrCreditv As Integer
Dim Differenttypev As Integer
Dim Authorityv As Integer

Dim mDebitCredit As Boolean
Dim mLastAccount As Boolean
Dim mSql As String
Dim StrNewAccountCode  As String
Dim h As Long
For i = 1 To GrdAccount2.Rows - 1
 

    
      
    
    mCode = Trim(GrdAccount2.TextMatrix(i, GrdAccount2.ColIndex("Account_Serial")))
    If mCode = "" Then Exit Sub
    Select Case Val(mCode)
    Case 1
        AccountTypes = 1
        AccountTab = 0
        DepitOrCreditv = 1
        Differenttypev = 1
        Authorityv = 0
    Case 2
        AccountTypes = 1
        AccountTab = 1
        DepitOrCreditv = 1
        Differenttypev = 1
        Authorityv = 0
    Case 3
        AccountTypes = 2
        AccountTab = 2
        DepitOrCreditv = 1
        Differenttypev = 1
        Authorityv = 0
    Case 4
        AccountTypes = 2
        AccountTab = 3
        DepitOrCreditv = 1
        Differenttypev = 1
        Authorityv = 0
    Case 5
        AccountTypes = 0
        AccountTab = 4
        DepitOrCreditv = 1
        Differenttypev = 1
        Authorityv = 0
    End Select
    
    Dim mLevel As Integer
    mLastAccount = False
    mNameP = Trim(GrdAccount2.TextMatrix(i, GrdAccount2.ColIndex("AccountNamePar")))
    If mNameP = "«·Œ’Ê„" Then
        mNameP = "«·Œ’Ê„"
    End If
    mLevel = 1
    sql = " select * from ACCOUNTS Where Level = " & mLevel & " and (  Account_Name Like '" & mNameP & "')"
    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If rs.EOF Then
        StrNewAccountCode = AddNewAccount("r", Trim$(mNameP), True, False, Trim$(mNameP), 1, False, False, False, , "", 0, 0, 0, 1, 0, IIf(Not mDebitCredit, 0, 1), 0, 0, 0, 1, False, True, mLastAccount, Val(mCode), mLevel)
        SaveBransh_UserAccount StrNewAccountCode
        'mSql = GetSqlQueryInsert(rs, ServerDb, "ACCOUNTS", "Account_ID", "", "", 0, 0, True)
    Else
        StrNewAccountCode = Trim(rs!Account_code & "")
    End If
    
    
    

    
    mLevel = 2
    mCode2 = Trim(GrdAccount2.TextMatrix(i, GrdAccount2.ColIndex("Account_Serial2")))
    mNameP2 = Trim(GrdAccount2.TextMatrix(i, GrdAccount2.ColIndex("AccountNamePar2")))
    
    sql = " select * from ACCOUNTS Where Level = " & mLevel & " and (  Account_Name Like '" & mNameP2 & "') and Parent_Account_Code = N'" & StrNewAccountCode & "'"
    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If rs.EOF Then

    
        sql = " select * from ACCOUNTS Where Account_Code = '" & StrNewAccountCode & "' "
        Set rs = New ADODB.Recordset
        rs.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
        If Not rs.EOF Then
            StrNewAccountCode = AddNewAccount(Trim(rs!Account_code & ""), Trim$(mNameP2), True, False, Trim$(mNameP2), 1, False, False, False, , "", 0, 0, 0, 1, 0, IIf(rs("DepitOrCredit").Value = 0, 0, 1), 0, 0, 0, 1, False, False, False, , mLevel)
            SaveBransh_UserAccount StrNewAccountCode
            'mSql = GetSqlQueryInsert(rs, ServerDb, "ACCOUNTS", "Account_ID", "", "", 0, 0, True)
        End If
    Else
            StrNewAccountCode = Trim(rs!Account_code & "")
    End If
    
    mCode3 = Trim(GrdAccount2.TextMatrix(i, GrdAccount2.ColIndex("Account_Serial3")))
    mNameP3 = Trim(GrdAccount2.TextMatrix(i, GrdAccount2.ColIndex("AccountNamePar3")))
    mLevel = 3

    sql = " select * from ACCOUNTS Where Level = " & mLevel & " and (  Account_Name Like '" & mNameP3 & "') and Parent_Account_Code = N'" & StrNewAccountCode & "'"
    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If rs.EOF Then

        sql = " select * from ACCOUNTS Where Account_Code = '" & StrNewAccountCode & "' "
        Set rs = New ADODB.Recordset
        rs.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
        If Not rs.EOF Then
            StrNewAccountCode = AddNewAccount(Trim(rs!Account_code & ""), Trim$(mNameP3), True, False, Trim$(mNameP3), 1, False, False, False, , "", 0, 0, 0, 1, 0, IIf(rs("DepitOrCredit").Value = 0, 0, 1), 0, 0, 0, 1, False, False, False, , mLevel)
            SaveBransh_UserAccount StrNewAccountCode
            'mSql = GetSqlQueryInsert(rs, ServerDb, "ACCOUNTS", "Account_ID", "", "", 0, 0, True)
        End If
    Else
            StrNewAccountCode = Trim(rs!Account_code & "")
    End If
    
    mCode4 = Trim(GrdAccount2.TextMatrix(i, GrdAccount2.ColIndex("Account_Serial4")))
    mNameP4 = Trim(GrdAccount2.TextMatrix(i, GrdAccount2.ColIndex("AccountNamePar4")))
     mLevel = 4
    sql = " select * from ACCOUNTS Where Level = " & mLevel & " and (  Account_Name Like '" & mNameP4 & "') and Parent_Account_Code = N'" & StrNewAccountCode & "'"
    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If rs.EOF Then

           sql = " select * from ACCOUNTS Where Account_Code = '" & StrNewAccountCode & "' "
        Set rs = New ADODB.Recordset
        rs.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
        If Not rs.EOF Then
            StrNewAccountCode = AddNewAccount(Trim(rs!Account_code & ""), Trim$(mNameP4), True, False, Trim$(mNameP4), 1, False, False, False, , "", 0, 0, 0, 1, 0, IIf(rs("DepitOrCredit").Value = 0, 0, 1), 0, 0, 0, 1, False, False, False, , mLevel)
            SaveBransh_UserAccount StrNewAccountCode
            'mSql = GetSqlQueryInsert(rs, ServerDb, "ACCOUNTS", "Account_ID", "", "", 0, 0, True)
        End If
    Else
            StrNewAccountCode = Trim(rs!Account_code & "")
    End If
    
    
    mCode5 = Trim(GrdAccount2.TextMatrix(i, GrdAccount2.ColIndex("Account_Serial5")))
    mNameP5 = Trim(GrdAccount2.TextMatrix(i, GrdAccount2.ColIndex("AccountNamePar5")))
    mLevel = 5
    mLastAccount = CBool(GrdAccount2.TextMatrix(i, GrdAccount2.ColIndex("last_account")))
 
    sql = " select * from ACCOUNTS Where Level = " & mLevel & " and (  Account_Name Like '" & mNameP5 & "')"
    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
   ' If rs.EOF Then
    
        sql = " select * from ACCOUNTS Where Account_Code = '" & StrNewAccountCode & "' "
        Set rs = New ADODB.Recordset
        rs.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
        If Not rs.EOF Then
            StrNewAccountCode = AddNewAccount(Trim(rs!Account_code & ""), Trim$(mNameP5), True, False, Trim$(mNameP5), 1, False, False, False, , "", 0, 0, 0, 1, 0, IIf(rs("DepitOrCredit").Value = 0, 0, 1), 0, 0, 0, 1, False, False, mLastAccount, , mLevel)
            SaveBransh_UserAccount StrNewAccountCode
            'mSql = GetSqlQueryInsert(rs, ServerDb, "ACCOUNTS", "Account_ID", "", "", 0, 0, True)
            If StrNewAccountCode <> "" Then
            
                GrdAccount2.TextMatrix(i, GrdAccount2.ColIndex("IsCreated")) = 1
            Else
                GrdAccount2.TextMatrix(i, GrdAccount2.ColIndex("IsCreated")) = 0
                
            
            End If
        Else
            MsgBox "«·«» «·—«»⁄ ”ÿ— —ﬁ„" & i & "·„ Ì „ «‰‘«∆Â »—Ã«¡ „—«Ã⁄… «·„·›"
            GrdAccount2.TextMatrix(i, GrdAccount2.ColIndex("IsCreated")) = 0
        End If
        
   ' End If

  
    
Next


End Sub



Private Sub Command1_Click()
CD1.ShowOpen
txtFile.Text = CD1.filename
End Sub

Private Sub Command10_Click()
Dim mCode As String
Dim mName As String
Dim mNameP As String

Dim mCode2 As String
Dim mName2 As String
Dim mNameP2 As String


Dim mCode3 As String
Dim mName3 As String
Dim mNameP3 As String


Dim mCode4 As String
Dim mName4 As String
Dim mNameP4 As String

Dim mCode5 As String
Dim mName5 As String
Dim mNameP5 As String
Dim AccountTypes As Integer
Dim AccountTab As Integer
Dim DepitOrCreditv As Integer
Dim Differenttypev As Integer
Dim Authorityv As Integer
Dim mCredit As Double
Dim mDebit As Double
Dim mBalacne As Double
Dim i As Long
Dim j As Long
Dim mEmpCode As String
cmdCheckDataOpen_Click
GrdAccount4.Rows = 1
For i = 1 To GrdAccountOpen.Rows - 1
    mCode = Trim(GrdAccountOpen.TextMatrix(i, GrdAccountOpen.ColIndex("Account_Serial")))
    mNameP = Trim(GrdAccountOpen.TextMatrix(i, GrdAccountOpen.ColIndex("Account_Name")))
    mCredit = Val(GrdAccountOpen.TextMatrix(i, GrdAccountOpen.ColIndex("Credit")))
    mDebit = Val(GrdAccountOpen.TextMatrix(i, GrdAccountOpen.ColIndex("Debit")))
    mEmpCode = Trim(GrdAccountOpen.TextMatrix(i, GrdAccountOpen.ColIndex("EmpCode")))
    
    mBalacne = mDebit - mCredit
    If mBalacne = 0 Then GoTo NextRow
    If mEmpCode <> "" Then
        s = "Select Account_code1 as Account_code From TblEmployee Where Emp_Code = N'" & mEmpCode & "'"
        Set rsDummy = New ADODB.Recordset
        rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly
        If Not rsDummy.EOF Then
            GoTo Finish
        End If
    End If
    
    
    s = " SELECT Account_Serial,Account_Code,Account_Name,last_account"
    s = s & " FROM ACCOUNTS Where last_account = 1 and  (Account_Name Like '%" & Trim(mNameP) & "%' ) "
    
    's = s & " and (IsNull(BranchId,0) = 0 Or BranchId = " & Val(cmbBranch.BoundText) & ")"
    Set rsDummy = New ADODB.Recordset
    rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly
    If Not rsDummy.EOF Then
    Else
        s = " SELECT Account_Serial,Account_Code,Account_Name,last_account"
        s = s & " FROM ACCOUNTS Where last_account = 1 and  ( Account_Serial = '" & mCode & "' and  Account_Name Like '%" & Trim(mNameP) & "%' )"
    '    s = s & " and (IsNull(BranchId,0) = 0 Or BranchId = " & Val(cmbBranch.BoundText) & ")"
        Set rsDummy = New ADODB.Recordset
        rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly
        If rsDummy.EOF Then
            GoTo copyLine
            
        End If
    End If
    If rsDummy.RecordCount > 1 Then
        s = " SELECT Account_Serial,Account_Code,Account_Name,last_account"
        s = s & " FROM ACCOUNTS Where last_account = 1 and  (Account_Name Like '%" & Trim(mNameP) & "%' ) "
    
        s = s & " and (IsNull(BranchId,0) = 0 Or BranchId = " & Val(cmbBranch.BoundText) & ")"
        Set rsDummy = New ADODB.Recordset
        rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly
    End If
    If rsDummy.EOF Then
    GoTo copyLine
    End If
Finish:
    GrdAccountOpen.TextMatrix(i, GrdAccountOpen.ColIndex("Account_Code")) = rsDummy!Account_code & ""
           
        
        
     
    
    If Not isFound Then GoTo NextRow
copyLine:
    GrdAccount4.Rows = GrdAccount4.Rows + 1
    GrdAccount4.TextMatrix(GrdAccount4.Rows - 1, GrdAccount4.ColIndex("Account_Serial")) = mCode
    GrdAccount4.TextMatrix(GrdAccount4.Rows - 1, GrdAccount4.ColIndex("Account_Name")) = mNameP
    GrdAccount4.TextMatrix(GrdAccount4.Rows - 1, GrdAccount4.ColIndex("Line1")) = Line1
    GrdAccount4.TextMatrix(GrdAccount4.Rows - 1, GrdAccount4.ColIndex("DebitValue")) = mDebit
        GrdAccount4.TextMatrix(GrdAccount4.Rows - 1, GrdAccount4.ColIndex("CreditValue")) = mCredit
            GrdAccount4.TextMatrix(GrdAccount4.Rows - 1, GrdAccount4.ColIndex("Ser")) = i
    GrdAccount4.TextMatrix(GrdAccount4.Rows - 1, GrdAccount4.ColIndex("Account_SerialP")) = GrdAccountOpen.TextMatrix(i, GrdAccountOpen.ColIndex("Account_SerialP"))

NextRow:
 Next i
If GrdAccount4.Rows > 1 Then Frame4.Visible = True

End Sub

Private Sub Command11_Click()
Dim s As String
Dim i As Long
Dim rsDummy As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim mCode As String
Dim mName As String
Dim mNameP As String
Dim mSql As String
Dim StrNewAccountCode  As String
Dim mBranchName As String
Dim mBranchId As Long
Dim opening_balance As Double
For i = 1 To GrdAccount4.Rows - 1
    mCode = Trim(GrdAccount4.TextMatrix(i, GrdAccount4.ColIndex("Account_SerialP")))
    mName = Trim(GrdAccount4.TextMatrix(i, GrdAccount4.ColIndex("Account_Name")))
    
   
    If mCode <> "" Then
    mBranchId = Val(cmbBranch.BoundText)
 
    
    sql = " select * from ACCOUNTS Where Account_Serial = '" & mCode & "' and last_account = 0 "
    Set rs = New ADODB.Recordset
    rs.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If Not rs.EOF Then
        StrNewAccountCode = AddNewAccount(Trim(rs!Account_code & ""), Trim$(mName), True, False, Trim$(mName), 1, False, False, False, , "", 0, 0, 0, 1, 0, IIf(rs("DepitOrCredit").Value = 0, 0, 1), 0, 0, 0, 1, False, , True, , , mBranchId)
        SaveBransh_UserAccount StrNewAccountCode
        'mSql = GetSqlQueryInsert(rs, ServerDb, "ACCOUNTS", "Account_ID", "", "", 0, 0, True)
    End If
    End If
 
  
  
  
    
Next
End Sub

Private Sub Command12_Click()
Dim s As String
Dim i As Long
Dim rsDummy As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim mCode As String
Dim mName As String
Dim mNameP As String
Dim mSql As String
Dim StrNewAccountCode  As String
Dim mBranchName As String
Dim mBranchId As Long
Dim opening_balance As Double
Dim mSer As Long
For i = 1 To GrdAccount4.Rows - 1
    mCode = Trim(GrdAccount4.TextMatrix(i, GrdAccount4.ColIndex("Account_Serial")))
    mName = Trim(GrdAccount4.TextMatrix(i, GrdAccount4.ColIndex("Account_Name")))
    mSer = Val(GrdAccount4.TextMatrix(i, GrdAccount4.ColIndex("Ser")))
    
    
    
    GrdAccountOpen.TextMatrix(mSer, GrdAccountOpen.ColIndex("Account_SerialP")) = GrdAccount4.TextMatrix(i, GrdAccount4.ColIndex("Account_SerialP"))
    GrdAccountOpen.TextMatrix(mSer, GrdAccountOpen.ColIndex("Account_Name")) = GrdAccount4.TextMatrix(i, GrdAccount4.ColIndex("Account_Name"))
       
  
   
    
  
    
 
  
  
  
    
Next

End Sub

Private Sub Command13_Click()
Dim s As String
Dim i As Long
Dim rsDummy As New ADODB.Recordset
Dim rs As New ADODB.Recordset
Dim mCode As String
Dim mName As String
Dim mNameP As String
Dim mSql As String
Dim StrNewAccountCode  As String
Dim mBranchName As String
Dim mBranchId As Long
Dim opening_balance As Double

s = txtSql
Set rsDummy = New ADODB.Recordset
rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly
Do While Not rsDummy.EOF
    
    StrNewAccountCode = AddNewAccount(Trim(rsDummy!parent_account & ""), Trim$(rsDummy!Name & ""), True, False, Trim$(rsDummy!Name & ""), 1, False, False, False, , "", 0, 0, 0, 1, 0, IIf(rs("DepitOrCredit").Value = 0, 0, 1), 0, 0, 0, 1, False, , True, , , 1, 0)
    SaveBransh_UserAccount StrNewAccountCode

    rsDummy.MoveNext
Loop
   
   

End Sub

Private Sub Command14_Click()

tmpGrd.Cols = 0

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
Case 7

    
    
    For i = 1 To grdBox.Rows - 1
        grdBox.TextMatrix(i, grdBox.ColIndex("BoxName")) = Trim(grdBox.TextMatrix(i, grdBox.ColIndex("BoxName"))) & " " & grdBox.TextMatrix(i, grdBox.ColIndex("ref"))
    Next
   Set mGrd = grdBox
Case 9
    Set mGrd = GrdAccount
Case 11
    Set mGrd = GrdAccount2
Case 12
    Set mGrd = grdFAGroups
Case 13
    Set mGrd = grdFa
Case 15
    Set mGrd = GrdEmpFee
Case 16
    Set mGrd = grdSchools
Case 17
    Set mGrd = grdCars
 Case 18
End Select

For i = 0 To mGrd.Cols - 1
    If Not mGrd.ColHidden(i) Then
        tmpGrd.Cols = tmpGrd.Cols + 1
        tmpGrd.ColKey(tmpGrd.Cols - 1) = mGrd.ColKey(i)
        tmpGrd.TextMatrix(0, tmpGrd.Cols - 1) = mGrd.TextMatrix(0, i)
    End If
Next
'tmpGrd = Grd

ExportToExcel Me, mGrd, tmpGrd, , , Me.Caption
End Sub

Private Sub Command15_Click()
Dim rsDummy As New ADODB.Recordset
Dim s As String
Dim X As String
Dim xx As Variant
Dim xxx As Variant
Dim mtmpName As String
Dim mName As String
Dim mNamee As String
Dim II As Integer
Dim i As Long

    s = "Select * from tblemployee "
    's = s & " ,cusId "

    
    'where   Len(" & txtFeildName & ")  "
    Dim mShortName As String
    rsDummy.Open s, Cn, adOpenKeyset, adLockOptimistic
    Do While Not rsDummy.EOF
        
        If Trim(rsDummy!Emp_code) = "60077" Then
            X = 0
        
        End If
        X = Trim(rsDummy!emp_name & "")
        i = 0
        xx = Split(X)
        mName = ""
        
        rsDummy!Emp_Name4 = ""
        mShortName = ""
        For i = 0 To UBound(xx)
            xxx = Split(xx(i), vbTab)
            II = 0
            mtmpName = ""
            For II = 0 To UBound(xxx)
                If xxx(II) <> "" Then
                    mtmpName = xxx(II)
                    If Len(mtmpName) <= 2 Then
                        mShortName = mtmpName & " "
                        GoTo NextS
                    
                        
                    End If
                    Exit For
                End If
            Next
            
            If mtmpName = "" Then mtmpName = xx(i)
            If i = UBound(xx) - 1 And Len(RTrim(LTrim(mtmpName))) > 20 Then
                
            Else
                If mtmpName <> "" Then
                    
                    If Val(mtmpName) <> 0 Then
                        mtmpName = " " & CStr(mtmpName)
                    End If
                    xxx = Split(mtmpName, vbent)
                    If UBound(xxx) > 0 Then
                       ' mtmpName = ""
                        
                        For jj = 0 To UBound(xxx)
                        '    mtmpName = mtmpName & xxx(jj)
                        Next
                    End If

                    mName = Trim(mtmpName)
                End If
                
            End If
            
            If i = 0 Then
                rsDummy!Emp_Name1 = mShortName & mName
            ElseIf i = 1 Then
                rsDummy!Emp_Name2 = mShortName & mName
            ElseIf i = 2 Then
                rsDummy!Emp_Name3 = mShortName & mName
            Else
                rsDummy!Emp_Name4 = IIf((rsDummy!Emp_Name4 & "") <> "", Trim(rsDummy!Emp_Name4 & "") & " ", "") & mShortName & mName
            End If
           mShortName = ""
           rsDummy!Emp_Namee1 = rsDummy!Emp_Name1
           rsDummy!Emp_Namee2 = rsDummy!Emp_Name2
           rsDummy!Emp_Namee3 = rsDummy!Emp_Name3
           rsDummy!Emp_Namee4 = rsDummy!Emp_Name4
           
NextS:
        Next
'        X = rsDummy("" & txtFeildNameE & "")
'        i = 0
'        xx = Split(X)
'        mNamee = ""
'        For i = 0 To UBound(xx) - 1
'            If i = UBound(xx) - 1 And Len(RTrim(LTrim(xx(i)))) > 20 Then
'
'            Else
'                mNamee = mNamee & " " & xx(i)
'            End If
'        Next
    
          
       
        rsDummy.Update
        
        rsDummy.MoveNext
    Loop
    
    MsgBox " „"

End Sub

Private Sub Command16_Click()
    Dim rs As ADODB.Recordset
     Dim RSTransDetails As ADODB.Recordset
     Dim RSDetails As ADODB.Recordset
     
     Dim rsItems As New ADODB.Recordset
     Dim mUnitID As Long
 

Dim s As String
Dim rsDummy As ADODB.Recordset
        s = "SELECT * FROM TblStore AS ts"
        Set rsDummy = New ADODB.Recordset
        rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly
        Dim fromdateS As Variant
        Dim todateS As Variant
        Dim openingdate As Variant
        Dim StrSQL As String
    
        fromdateS = Replace(Format$(DTPFrom.Value, "MM/DD/yyyy"), "-", "/")
        todateS = Replace(Format$(DTPTo.Value, "MM/DD/yyyy"), "-", "/")
        openingdate = DateAdd("D", -1, DTPFrom.Value)
        openingdate = Replace(Format$(openingdate, "MM/DD/yyyy"), "-", "/")
          Cn.BeginTrans
                BegineTrans = True

    
        Do While Not rsDummy.EOF
                    DcStore10.BoundText = Val(rsDummy!StoreId & "")
                    DCboStoreName.BoundText = Val(rsDummy!StoreId & "")
                    
                            StrSQL = "Select * From Transactions where Transaction_Type=30"
    StrSQL = StrSQL & "  AND     1 = -1"
    
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    
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
                If optOpenBalanceItem(0) Then
                rs("Transaction_Type").Value = 3
                Else
                rs("Transaction_Type").Value = 30
                End If
                rs("UserID").Value = 1
                rs("StoreID").Value = IIf(DCboStoreName.BoundText = "", Null, DCboStoreName.BoundText)
                rs.Update
                
                    
                    StrSQL = " SELECT * FROM ( SELECT     isnull(QryItemsInventry3.outputValue,0) as outputValue,"
                    StrSQL = StrSQL + " isnull(QryItemsInventry3.inputvalue,0) as inputvalue,"
                    StrSQL = StrSQL + " isnull( QryItemsInventry3.openingValue,0) as openingValue ,"
                    
                    
                   
                    StrSQL = StrSQL + " isnull(  dbo.GetItemqtytodatenew('" & openingdate & "', dbo.TblItems.ItemID, "
                    StrSQL = StrSQL + IIf(DcStore10.BoundText = "", "null", (DcStore10.BoundText))
                    StrSQL = StrSQL + ","
                    StrSQL = StrSQL + IIf(DCItemsColors.BoundText = "", "null", (DCItemsColors.BoundText))
                    StrSQL = StrSQL + ","
                    StrSQL = StrSQL + IIf(DCItemsSizes.BoundText = "", "null", (DCItemsSizes.BoundText))
                    StrSQL = StrSQL + ","
                    StrSQL = StrSQL + IIf(DCItemsClasses.BoundText = "", "null", (DCItemsClasses.BoundText))
            
                    StrSQL = StrSQL + ", NULL "
                    StrSQL = StrSQL + ","
                    StrSQL = StrSQL + IIf(dcCustomer.BoundText = "", "null", (dcCustomer.BoundText))
                    StrSQL = StrSQL + " ) ,0) "
                    StrSQL = StrSQL + "  + ISNULL(openingValue, 0) + ISNULL(inputvalue, 0) + ISNULL(outputValue, 0) balance ,"
            
            
             
                    StrSQL = StrSQL + " isnull( dbo.GetItemCostPrice('" & "01/01/2000" & "', ' " & todateS & " ',  dbo.TblItems.ItemID) ,  0)AS Cost,"
            
'                    StrSQL = StrSQL + " isnull(  dbo.GetItemqtytodatenew('" & openingdate & "', dbo.TblItems.ItemID, "
'                    StrSQL = StrSQL + IIf(DcStore10.BoundText = "", "null", (DcStore10.BoundText))
'                    StrSQL = StrSQL + ","
'                    StrSQL = StrSQL + IIf(DCItemsColors.BoundText = "", "null", (DCItemsColors.BoundText))
'                    StrSQL = StrSQL + ","
'                    StrSQL = StrSQL + IIf(DCItemsSizes.BoundText = "", "null", (DCItemsSizes.BoundText))
'                    StrSQL = StrSQL + ","
'                    StrSQL = StrSQL + IIf(DCItemsClasses.BoundText = "", "null", (DCItemsClasses.BoundText))
'
'                    StrSQL = StrSQL + ", NULL "
'                    StrSQL = StrSQL + ","
'                    StrSQL = StrSQL + IIf(dcCustomer.BoundText = "", "null", (dcCustomer.BoundText))
'                    StrSQL = StrSQL + " ) ,0) AS oldopening"
            
                    StrSQL = StrSQL + "   dbo.TblItems.barcodeno,  dbo.TblItems.ItemCode, dbo.TblItems.ItemName, dbo.TblItems.ItemID, dbo.TblItems.ItemNamee"
                    StrSQL = StrSQL + "  , dbo.Groups.GroupCode, dbo.Groups.GroupName, dbo.Groups.GroupNamee  "
                    StrSQL = StrSQL + " FROM         dbo.QryItemsInventry3('" & fromdateS & "', '" & todateS & "',"
            
                    StrSQL = StrSQL + IIf(DcStore10.BoundText = "", "null", (DcStore10.BoundText))
                    StrSQL = StrSQL + ","
                    StrSQL = StrSQL + IIf(DCItemsColors.BoundText = "", "null", (DCItemsColors.BoundText))
                    StrSQL = StrSQL + ","
                    StrSQL = StrSQL + IIf(DCItemsSizes.BoundText = "", "null", (DCItemsSizes.BoundText))
                    StrSQL = StrSQL + ","
                    StrSQL = StrSQL + IIf(DCItemsClasses.BoundText = "", "null", (DCItemsClasses.BoundText))
                    StrSQL = StrSQL + " ,Null "
                    StrSQL = StrSQL + ","
                    StrSQL = StrSQL + IIf(dcCustomer.BoundText = "", "null", (dcCustomer.BoundText))
                    StrSQL = StrSQL + ") QryItemsInventry3 RIGHT OUTER JOIN"
                    StrSQL = StrSQL + " dbo.TblItems ON QryItemsInventry3.Item_ID = dbo.TblItems.ItemID "
            StrSQL = StrSQL + " INNER JOIN  dbo.Groups  ON dbo.Groups.GroupID = dbo.TblItems.GroupID  "
             StrSQL = StrSQL + "  where 1=1"
             
             StrSQL = StrSQL + "  and  TblItems.ItemID  in ( "
             StrSQL = StrSQL + " SELECT DISTINCT dbo.Transaction_Details.Item_ID"
             StrSQL = StrSQL + " FROM         dbo.Transactions INNER JOIN"
             StrSQL = StrSQL + " dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type INNER JOIN"
             StrSQL = StrSQL + " dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID"
             StrSQL = StrSQL + " WHERE     (dbo.TransactionTypes.StockEffect <> 0) AND (dbo.Transactions.StoreID =" & DcStore10.BoundText & "))"
             StrSQL = StrSQL + "  )FF WHERE FF.balance <> 0"

            StrSQL = StrSQL + "Order By"
            StrSQL = StrSQL + "       itemcode"
            
            Set RSDetails = New ADODB.Recordset
            RSDetails.Open StrSQL, Cn, adOpenStatic, adLockReadOnly
            Do While Not RSDetails.EOF
                
                
                RSTransDetails.AddNew
                RSTransDetails("Transaction_ID").Value = XPTxtBillID.Text
                
                RSTransDetails("AutoDetect").Value = 0
                RSTransDetails("Item_ID").Value = Val(RSDetails!ItemID & "")
                RSTransDetails("Quantity").Value = Val(RSDetails!Balance & "")
                
                
                'RSTransDetails("ParrtNoCode").Value = IIf((grdItems.TextMatrix(RowNum, grdItems.ColIndex("ParrtNoCode")) = ""), Null, (grdItems.TextMatrix(RowNum, grdItems.ColIndex("ParrtNoCode"))))
                '    RSTransDetails("ItemDetailedCode").Value = IIf((grdItems.TextMatrix(RowNum, grdItems.ColIndex("ItemDetailedCode")) = ""), Null, (grdItems.TextMatrix(RowNum, grdItems.ColIndex("ItemDetailedCode"))))
                '
                RSTransDetails("ItemCase").Value = 1
                RSTransDetails("Price").Value = Val(RSDetails!cost & "")
                
                                RSTransDetails("ColorID").Value = 1
                '
                                RSTransDetails("ItemSize").Value = 1
                '
                                RSTransDetails("ClassId").Value = 1
                
                RSTransDetails("BranchId").Value = 1
              
                
                GetDefaultItemUnit Val(RSDetails!ItemID & ""), mUnitID
                RSTransDetails("UnitID").Value = mUnitID
                RSTransDetails("ShowQty").Value = Val(RSDetails!Balance & "")
                
                Dim RsUnitData As ADODB.Recordset
                Dim LngCurItemID As Long
                Dim LngUnitID As Long
                Dim DblQty As Double
                
                LngCurItemID = Val(RSDetails!ItemID & "")
                LngUnitID = mUnitID
                DblQty = Val(RSDetails!Balance & "")
                
                StrSQL = "Select * From TblItemsUnits Where ItemID=" & LngCurItemID
                StrSQL = StrSQL + " AND UnitID=" & LngUnitID
                Set RsUnitData = New ADODB.Recordset
                RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                
                If Not (RsUnitData.BOF Or RsUnitData.EOF) Then
                    RSTransDetails("QtyBySmalltUnit").Value = RsUnitData("UnitFactor").Value
                    RSTransDetails("Quantity").Value = RSTransDetails("QtyBySmalltUnit").Value * RSTransDetails("showqty").Value
                End If
                
                RSTransDetails("Price").Value = Val(RSDetails!cost & "")
               
                
                
               
                RSTransDetails("showprice").Value = Val(RSDetails!cost & "")
              
                RSTransDetails.Update
            
                    
                RSDetails.MoveNext
            Loop
            
       rsDummy.MoveNext
    Loop
    s = "Delete FROM Transactions WHERE Transaction_ID NOT IN (SELECT Transaction_ID FROM Transaction_Details ) AND Transaction_Type = 3"
    Cn.Execute s
            Cn.CommitTrans
        BegineTrans = False
        
        MsgBox " „ «·Ã—œ"

End Sub

Private Sub Command2_Click()


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
Function LoadExcelToGroups(mGrd As Object)
    '<EhHeader>
    On Error GoTo LoadExcelToGroups_Err
    '</EhHeader>
   
    Dim i             As Integer
    Dim Fullcode      As String
    Dim GroupName     As String
    Dim GroupNamee    As String
    Dim MainGroumName As String
    Dim GroupCode     As String
    Dim ParentId      As String
    Dim tmpTable      As String
    tmpTable = "#tmpGrp"
    Dim sqlStr As String
    sqlStr = " IF OBJECT_ID(N'tempdb.." & tmpTable & "') IS NOT NULL "
    sqlStr = sqlStr & " BEGIN "
    sqlStr = sqlStr & "DROP TABLE " & tmpTable & " "
    sqlStr = sqlStr & " End "
    Cn.Execute sqlStr
    sqlStr = ""
    sqlStr = sqlStr & "CREATE TABLE " & tmpTable & " "
    sqlStr = sqlStr & "("
    sqlStr = sqlStr & " id INT IDENTITY NOT NULL PRIMARY KEY, "
    sqlStr = sqlStr & " Fullcode NVARCHAR(255) NULL, "
    sqlStr = sqlStr & " GroupName NVARCHAR(50) NULL, "
    sqlStr = sqlStr & " GroupNamee NVARCHAR(255) NULL, "
    sqlStr = sqlStr & " MainGroupName NVARCHAR(255) NULL, "
    sqlStr = sqlStr & " GroupCode NVARCHAR(255) NULL, "
    sqlStr = sqlStr & " ParentId NVARCHAR(50) NULL , "
    sqlStr = sqlStr & " dbId INT NULL , "
    sqlStr = sqlStr & " dbMainId INT NULL "
    sqlStr = sqlStr & ") "
    Cn.Execute sqlStr
    Dim rs As New ADODB.Recordset
    rs.Open "Select * FRom " & tmpTable & "", Cn, adOpenKeyset, adLockOptimistic
    For i = 1 To mGrd.Rows - 1
        Fullcode = Trim(mGrd.TextMatrix(i, mGrd.ColIndex("Fullcode")))
        GroupName = Trim(mGrd.TextMatrix(i, mGrd.ColIndex("GroupName")))
        GroupNamee = Trim(mGrd.TextMatrix(i, mGrd.ColIndex("GroupNamee")))
        MainGroumName = Trim(mGrd.TextMatrix(i, mGrd.ColIndex("MainGroumName")))
        GroupCode = Trim(mGrd.TextMatrix(i, mGrd.ColIndex("GroupCode")))
      
        If Fullcode = "" Then GoTo NextRow
        rs.AddNew
        rs!Fullcode = Fullcode
        rs!GroupName = GroupName
        rs!GroupNamee = GroupNamee
        rs!MainGroupName = MainGroumName
        rs!GroupCode = GroupCode
        rs!ParentId = ParentId
        rs.Update
NextRow:
    Next
    rs.Requery
    '    Rs.Close
    '    Rs.Open "Select *   From " & tmpTable & "", Cn, adOpenKeyset, adLockOptimistic
    If rs.EOF Then
        Exit Function
    End If
    
    Dim rsCheck As New ADODB.Recordset
    Dim rstmp   As New ADODB.Recordset
    Dim sql2    As String
    Dim s       As String
    Dim firstG  As Integer
    Do While Not rs.EOF
        sql2 = " SELECT top 1 * FROM Groups "
        sql2 = sql2 & " WHERE Fullcode = '" & rs!Fullcode & "'"
        rsCheck.Open sql2, Cn, adOpenKeyset, adLockOptimistic
        If Not rsCheck.EOF Then
            rs!dbid = rsCheck!GroupID
            rs!dbMainId = rsCheck!ParentId
            rsCheck!GroupName = rs!GroupName
            rsCheck!GroupNamee = rs!GroupNamee
        Else
            Dim NewId As Integer
            s = "SELECT Max(GroupID) MaxID  FROM Groups "
            rstmp.Open s, Cn, adOpenForwardOnly, adLockReadOnly, adCmdText
            If Not rstmp.EOF Then
                NewId = Val(rstmp!MaxId & "")
                
            End If
            NewId = NewId + 1
            If firstG = 0 Then
                firstG = NewId
            End If
            rsCheck.AddNew
            rsCheck!GroupID = NewId
            rsCheck!Fullcode = rs!Fullcode
            rsCheck!GroupName = rs!GroupName
            rsCheck!GroupNamee = rs!GroupNamee
            rs!dbid = NewId
            rstmp.Close
        End If
            
        rsCheck.Update
        rs.Update
        rsCheck.Close
        rs.MoveNext
    Loop
    
    rs.Requery
    Dim rsMain As New ADODB.Recordset
    Do While Not rs.EOF
        If rs!GroupCode & "" = "" Then GoTo NextMRow
        sql2 = " SELECT top 1 * FROM Groups "
        sql2 = sql2 & "WHERE GroupID = '" & rs!dbid & "'"
        rsCheck.Open sql2, Cn, adOpenKeyset, adLockOptimistic
        '*****************
        sql2 = "SELECT TOP 1  GroupID  FROM Groups "
        sql2 = sql2 & " WHERE FullCode = '" & rs!GroupCode & "'"
      
        rsMain.Open sql2, Cn, adOpenForwardOnly, adLockReadOnly
        If rsMain.EOF Then
            rsMain.Close
            sql2 = "SELECT TOP 1  GroupID  FROM Groups "
            sql2 = sql2 & " WHERE   GroupNamee = '" & rs!MainGroupName & "'  OR GroupNamee = '" & rs!MainGroupName & "' "
            rsMain.Open sql2, Cn, adOpenForwardOnly, adLockReadOnly
        End If
        Dim mGroupID As Integer
        If Not rsMain.EOF Then
            mGroupID = Val(rsMain!GroupID & "")
        End If
        If mGroupID > 0 And Val(rsCheck!ParentId & "") = 0 Then
         
            rsCheck!ParentId = mGroupID
            rsCheck!GroupCode = GetNewGroupCode(Val(mGroupID), "Groups")
        End If
        '******************
            
        rsCheck.Update
        rsMain.Close
        rsCheck.Close
NextMRow:
        rs.MoveNext
        
    Loop
    MsgBox "All Done" & firstG
    '<EhFooter>
    Exit Function

LoadExcelToGroups_Err:
    MsgBox Err.Description & vbCrLf & _
       "in ImportExportData.frmImport.LoadExcelToGroups " & _
       "at line " & Erl, _
       vbExclamation + vbOKOnly, "Application Error"
    '</EhFooter>
End Function
Private Sub Command3_Click()
    'ExportToExcel Me, Grd, "TT", , "grdItems"
    tmpGrd.Rows = 1

    Dim i As Long
    If mIndex = 0 Then
        Grd.Rows = 1
        FromExcel Grd, tmpGrd, Me, , , txtFile.Text, "TblEmployee"
    ElseIf mIndex = 1 Then
        If chkBalanceOnly.Value = vbChecked Then
            grdManBal.Rows = 1
            FromExcel grdManBal, tmpGrd, Me, , , txtFile.Text, "TblCustemers"

        Else
            grdMan.Rows = 1
            FromExcel grdMan, tmpGrd, Me, , , txtFile.Text, "TblCustemers"
        End If
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
    ElseIf mIndex = 7 Then
        grdBox.Rows = 1
        FromExcel grdBox, tmpGrd, Me, , , txtFile.Text, "TblBoxesData"
             
    ElseIf mIndex = 9 Then
        GrdAccount.Rows = 1
        FromExcel GrdAccount, tmpGrd, Me, , , txtFile.Text, "ACCOUNTS"

    ElseIf mIndex = 11 Then
        GrdAccount2.Rows = 1
        FromExcel GrdAccount2, tmpGrd, Me, , , txtFile.Text, "ACCOUNTS"

    ElseIf mIndex = 12 Then
        grdFAGroups.Rows = 1
        FromExcel grdFAGroups, tmpGrd, Me, , , txtFile.Text, "FixedAssetsGroup"

    ElseIf mIndex = 13 Then
        grdFa.Rows = 1
        FromExcel grdFa, tmpGrd, Me, , , txtFile.Text, "FixedAssets"
    ElseIf mIndex = 14 Then
        GrdAccountOpen.Rows = 1
        FromExcel GrdAccountOpen, tmpGrd, Me, , , txtFile.Text, "ACCOUNTS"
        txtTotalCredit = 0
        txtTotalDebit = 0
        For i = 1 To GrdAccountOpen.Rows - 1
            txtTotalDebit = Val(txtTotalDebit) + Val(GrdAccountOpen.TextMatrix(i, GrdAccountOpen.ColIndex("Debit")))
            txtTotalCredit = Val(txtTotalCredit) + Val(GrdAccountOpen.TextMatrix(i, GrdAccountOpen.ColIndex("Credit")))
        Next
    ElseIf mIndex = 15 Then
        GrdEmpFee.Rows = 1
        FromExcel GrdEmpFee, tmpGrd, Me, , , txtFile.Text, "empfees"

    ElseIf mIndex = 16 Then
        grdSchools.Rows = 1
        FromExcel grdSchools, tmpGrd, Me, , , txtFile.Text, "TblSchooleFile"

    ElseIf mIndex = 17 Then
        grdCars.Rows = 1
        FromExcel grdCars, tmpGrd, Me, , , txtFile.Text, "TblCarsData"

    ElseIf mIndex = 18 Then
        grdCars.Rows = 1
        FromExcel Grdtmp, tmpGrd, Me, , , txtFile.Text, "TblCarsData"

    End If

    cmdSave.Enabled = True

    Dim j         As Long
    Dim mJob1     As Long
    Dim mJobName1 As String
    Dim mJob2     As Long
    Dim mJobName2 As String
    Dim mJob3     As Long
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
    Command4.Enabled = True
End Sub

Private Sub Command4_Click()
     Dim rs As ADODB.Recordset
     Dim RSTransDetails As ADODB.Recordset
     Dim s As String
     Dim rsItems As New ADODB.Recordset
     
     
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
        If optOpenBalanceItem(0) Then
            rs("Transaction_Type").Value = 3
        Else
            rs("Transaction_Type").Value = 30
        End If
        rs("UserID").Value = 1
        rs("StoreID").Value = IIf(DCboStoreName.BoundText = "", Null, DCboStoreName.BoundText)
        rs.Update
        
            For RowNum = 1 To grdItems.Rows - 1

            If grdItems.TextMatrix(RowNum, grdItems.ColIndex("Code")) <> "" Then
                
                s = "Select * from tblItems Where Code =  '" & Trim(grdItems.TextMatrix(RowNum, grdItems.ColIndex("Code"))) & "'"
                 Set rsItems = New ADODB.Recordset
                rsItems.Open s, Cn, adOpenKeyset, adLockReadOnly
                If rsItems.EOF Then
                    s = "Select * from tblItems Where ItemName Like  '" & Trim(grdItems.TextMatrix(RowNum, grdItems.ColIndex("ItemName"))) & "'"
                    Set rsItems = New ADODB.Recordset
                    rsItems.Open s, Cn, adOpenKeyset, adLockReadOnly
                End If
                If Not rsItems.EOF Then
                grdItems.TextMatrix(RowNum, grdItems.ColIndex("ItemID")) = rsItems!ItemID & ""
                
                RSTransDetails.AddNew
                RSTransDetails("Transaction_ID").Value = XPTxtBillID.Text
                
                RSTransDetails("AutoDetect").Value = 0
                RSTransDetails("Item_ID").Value = IIf((grdItems.TextMatrix(RowNum, grdItems.ColIndex("ItemID")) = ""), Null, Val(grdItems.TextMatrix(RowNum, grdItems.ColIndex("ItemID"))))
                RSTransDetails("Quantity").Value = IIf((grdItems.TextMatrix(RowNum, grdItems.ColIndex("TotalQty")) = ""), Null, Val(grdItems.TextMatrix(RowNum, grdItems.ColIndex("TotalQty"))))

               
                'RSTransDetails("ParrtNoCode").Value = IIf((grdItems.TextMatrix(RowNum, grdItems.ColIndex("ParrtNoCode")) = ""), Null, (grdItems.TextMatrix(RowNum, grdItems.ColIndex("ParrtNoCode"))))
                '    RSTransDetails("ItemDetailedCode").Value = IIf((grdItems.TextMatrix(RowNum, grdItems.ColIndex("ItemDetailedCode")) = ""), Null, (grdItems.TextMatrix(RowNum, grdItems.ColIndex("ItemDetailedCode"))))
'
                'RSTransDetails("ItemCase").Value = IIf((grdItems.TextMatrix(RowNum, grdItems.ColIndex("ItemCase")) = ""), Null, Val(grdItems.TextMatrix(RowNum, grdItems.ColIndex("ItemCase"))))
                RSTransDetails("Price").Value = Val(grdItems.TextMatrix(RowNum, grdItems.ColIndex("UnitPurPrice")))
            
                RSTransDetails("ColorID").Value = 0 'IIf((grdItems.TextMatrix(RowNum, grdItems.ColIndex("ColorID")) = ""), 1, Val(grdItems.TextMatrix(RowNum, grdItems.ColIndex("ColorID"))))
'
                RSTransDetails("ItemSize").Value = 0 'IIf((grdItems.TextMatrix(RowNum, grdItems.ColIndex("ItemSize")) = ""), 1, Val(grdItems.TextMatrix(RowNum, grdItems.ColIndex("ItemSize"))))
'
                RSTransDetails("ClassId").Value = 0 'IIf((grdItems.TextMatrix(RowNum, grdItems.ColIndex("ClassId")) = ""), 1, Val(grdItems.TextMatrix(RowNum, grdItems.ColIndex("ClassId"))))
            
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
                If Val(grdItems.TextMatrix(RowNum, grdItems.ColIndex("UnitPurPrice"))) = 0 Then
                    RSTransDetails("Price").Value = Val(grdItems.TextMatrix(RowNum, grdItems.ColIndex("UnitSalesPrice")))
                Else
                    RSTransDetails("Price").Value = Val(grdItems.TextMatrix(RowNum, grdItems.ColIndex("UnitPurPrice")))
                
                End If
                If Val(grdItems.TextMatrix(RowNum, grdItems.ColIndex("UnitPurPrice"))) = 0 Then
                    RSTransDetails("showprice").Value = Val(grdItems.TextMatrix(RowNum, grdItems.ColIndex("UnitSalesPrice")))
                Else
                    RSTransDetails("showprice").Value = Val(grdItems.TextMatrix(RowNum, grdItems.ColIndex("UnitPurPrice")))
                End If
                RSTransDetails.Update
            End If
            End If
NextRow:
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
    rsData!Account_code = rsDummy!Account_code
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

Private Sub cmdCreateopenEntry_Click()
Dim s As String

Dim rsDummy As New ADODB.Recordset
Dim rsDummyData As New ADODB.Recordset
Dim mAccountName As String
Dim mAccount_Serial As String
Dim mAccountCode As String
Dim mBranchId As Long
Dim mBalacne As Double
Dim mDebit As Double
Dim mCredit As Double

Dim NoteID As Long
Dim EntryID As Long
Dim i As Long
Command10_Click
If Frame4.Visible = True Then
    MsgBox "ÌÊÃœ Õ”«»«  €Ì— „ÊÃÊœ… Ì—ÃÏ «· √ﬂœ „‰Â« «Ê·«"
    Exit Sub
End If


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
    mBranchId = Val(rsDummyData!branch_id & "")
    cmbBranch.BoundText = mBranchId
'    EntryID = Val(rsDummy!EntryID & "")
'    NoteID = Val(rsDummy!NoteID & "")
End If
'rsDummy.Close


Dim mLine As Integer
mLine = 0

For i = 1 To GrdAccountOpen.Rows - 1
  
    mAccountName = Trim(GrdAccountOpen.TextMatrix(i, GrdAccountOpen.ColIndex("Account_Name")))
    mAccount_Serial = Trim(GrdAccountOpen.TextMatrix(i, GrdAccountOpen.ColIndex("Account_Serial")))
    mAccount_Code = Trim(GrdAccountOpen.TextMatrix(i, GrdAccountOpen.ColIndex("Account_Code")))
    mCredit = Val(GrdAccountOpen.TextMatrix(i, GrdAccountOpen.ColIndex("Credit")))
    mDebit = Val(GrdAccountOpen.TextMatrix(i, GrdAccountOpen.ColIndex("Debit")))
    mDouble_Entry_Vouchers_Description = Trim(GrdAccountOpen.TextMatrix(i, GrdAccountOpen.ColIndex("Double_Entry_Vouchers_Description")))
    
    mBalacne = mDebit - mCredit
      If mAccountName = "" Then GoTo MoveNext1
'    If mBranchId <> Val(GrdAccountOpen.TextMatrix(i, GrdAccountOpen.ColIndex("BranchID"))) Then
'        MsgBox "—«Ã⁄ „·› «·«ﬂ”Ì· Ê«·›—⁄ ·«‰Â ÌÊÃœ «Œ ·«› »Ì‰ ›—⁄ «Œ— Õ—ﬂ… —’Ìœ Ê„·› «·«ﬂ”Ì·"
'        Exit Sub
'    End If
    If mLine = 0 Or mLine = 2 Then mLine = 1 Else mLine = 2

    If Abs(Val(mBalacne)) = 0 Then GoTo MoveNext1
    CreateOpenEntry EntryID, mBalacne, mAccount_Code, mLine, rsDummyData, mDouble_Entry_Vouchers_Description, mAccountName
   

MoveNext1:
Next i

s = "Select * from EmpFees Where BranchId = " & Val(cmbBranch.BoundText)
Set rsDummy = New ADODB.Recordset
rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly
Do While Not rsDummy.EOF
    mAccount_Code = Trim(rsDummy!Account_code2 & "")
    mBalacne = Val(rsDummy!Fee1 & "")
    If Abs(Val(mBalacne)) <> 0 Then
        If mLine = 0 Or mLine = 2 Then mLine = 1 Else mLine = 2
        mBalacne = mBalacne * -1
        CreateOpenEntry EntryID, mBalacne, mAccount_Code, mLine, rsDummyData
    End If
    
    mAccount_Code = Trim(rsDummy!Account_Code5 & "")
    mBalacne = Val(rsDummy!Fee2 & "")
    If Abs(Val(mBalacne)) <> 0 Then
        If mLine = 0 Or mLine = 2 Then mLine = 1 Else mLine = 2
        mBalacne = mBalacne * -1
        CreateOpenEntry EntryID, mBalacne, mAccount_Code, mLine, rsDummyData
    End If
    
    mAccount_Code = Trim(rsDummy!Account_code4 & "")
    mBalacne = Val(rsDummy!Fee3 & "")
    If Abs(Val(mBalacne)) <> 0 Then
        If mLine = 0 Or mLine = 2 Then mLine = 1 Else mLine = 2
        mBalacne = mBalacne * -1
        CreateOpenEntry EntryID, mBalacne, mAccount_Code, mLine, rsDummyData
    End If
    rsDummy.MoveNext
Loop


MsgBox " „ «‰‘«¡ «·ﬁÌœ"
cmdCreateopenEntry.Enabled = False
End Sub


Private Sub CreateOpenEntry(ByVal EntryID As Long, ByVal mBalacne As Double, ByVal mAccount_Code As String, ByVal mLine As Integer, ByRef rsDummyData As ADODB.Recordset, Optional ByVal mDouble_Entry_Vouchers_Description As String = "", Optional ByVal mAccountName As String = "")
   Dim mProjectId As Long
   If Len(mDouble_Entry_Vouchers_Description) > 10 Then
    mDouble_Entry_Vouchers_Description = mDouble_Entry_Vouchers_Description
   End If
   If mDouble_Entry_Vouchers_Description <> "" And (mDouble_Entry_Vouchers_Description <> "⁄„·«¡" And mDouble_Entry_Vouchers_Description <> "„Ê—œÌ‰" And mDouble_Entry_Vouchers_Description <> "–„„" And mDouble_Entry_Vouchers_Description <> "«’Ê·") Then
        Dim rsProj As New ADODB.Recordset
        Dim IsCostCenter As Boolean
         s = "SELECT * FROM projects AS p WHERE p.Project_name LIKE '%" & mDouble_Entry_Vouchers_Description & "%'"
         rsProj.Open s, Cn, adOpenKeyset, adLockReadOnly
         If Not rsProj.EOF Then
             mProjectId = Val(rsProj!ID & "")
             IsCostCenter = False
         Else
             mProjectId = 0
             s = "Select * from markaas_taklefa where account_name Like '%" & Trim(mDouble_Entry_Vouchers_Description) & "%'"
             Set rsProj = New ADODB.Recordset
             rsProj.Open s, Cn, adOpenStatic, adLockReadOnly
             If Not rsProj.EOF Then
                IsCostCenter = True
             End If
             
         
         End If
   End If
   s = "Select * from DOUBLE_ENTREY_VOUCHERS1 Where Double_Entry_Vouchers_ID = -5"
    Dim rsData As New ADODB.Recordset
    rsData.Open s, Cn, adOpenKeyset, adLockOptimistic
    rsData.AddNew
    rsData!Double_Entry_Vouchers_ID = EntryID + 1
    
    rsData!DEV_ID_Line_No = mLine
    rsData!Account_code = mAccount_Code
    rsData!Value = Abs(Val(mBalacne))
    If Val(mBalacne) > 0 Then
        rsData!Credit_Or_Debit = 0
    Else
        rsData!Credit_Or_Debit = 1
    End If
    'rsData!Credit_Or_Debit = 0
    
    
    rsData!branch_id = rsDummyData!branch_id
    rsData!RecordDate = rsDummyData!RecordDate
    rsData!Notes_ID = rsDummyData!Notes_ID
    rsData!UserID = rsDummyData!UserID
    rsData!projectid = mProjectId
    rsData!project_id = mProjectId
    
    rsData!Account_Interval_ID = rsDummyData!Account_Interval_ID
    rsData!DEV_Serial = rsDummyData!DEV_Serial
    rsData!Rate = rsDummyData!Rate
    rsData!Double_Entry_Vouchers_Description = mDouble_Entry_Vouchers_Description
    rsData!Notes_ID = rsDummyData!Notes_ID
    rsData!opening_balance_voucher_id = Val(rsData!opening_balance_voucher_id & "") + 1
    rsData!DEV_ID_Line_No1 = rsData!DEV_ID_Line_No1
    rsData!Remarks2 = 3
    rsData.Update
    
    If IsCostCenter Then
        s = " SELECT * FROM NOTES1 Where NoteID = " & Val(rsDummyData!Notes_ID & "")
        Dim rsNoteId As New ADODB.Recordset
        Dim mFoxId As Long
        Dim mNoteSerial As Long
        rsNoteId.Open s, Cn, adOpenForwardOnly, adLockReadOnly
        If rsNoteId.EOF Then
            Exit Sub
        End If
        
        Dim rsCostCenter As New ADODB.Recordset
        s = "Select * from marakes_taklefa_temp where cost_center_id = '-1'"
        rsCostCenter.Open s, Cn, adOpenKeyset, adLockOptimistic
        rsCostCenter.AddNew
        rsCostCenter!cost_center_id = rsProj!account_no
        rsCostCenter!cost_center = rsProj!account_name
        rsCostCenter!Value = Abs(Val(mBalacne))
        rsCostCenter!opr_id = Val(rsNoteId!foxy_no & "")
        rsCostCenter!kedno = Val(rsNoteId!foxy_no & "")
        rsCostCenter!NoteSerial = Val(rsNoteId!NoteSerial & "")
        rsCostCenter!NoteDate = (rsNoteId!NoteDate & "")
        rsCostCenter!record_date = (rsNoteId!NoteDate & "")
        rsCostCenter!line_no = rsData!DEV_ID_Line_No1
        
        rsCostCenter!account_name = mAccountName
        rsCostCenter!account_no = mAccount_Code
        rsCostCenter!opr_type = "”‰œ ﬁÌœ «›  «ÕÌ"
        
        If Val(mBalacne) > 0 Then
            rsCostCenter!depit_or_credit = "„œÌ‰"
        Else
            rsCostCenter!depit_or_credit = "œ«∆‰"
        End If
        rsCostCenter!ok = 1
        
        rsCostCenter.Update
        
        
         
        
    End If
End Sub
Private Sub DataList1_Click()

End Sub

Private Sub Command6_Click()
Frame4.Visible = False
End Sub

Private Sub Command7_Click()
Frame2.Visible = False
End Sub

Private Sub cmdCheckExcel_Click()
Dim mCode As String
Dim mName As String
Dim mNameP As String

Dim mCode2 As String
Dim mName2 As String
Dim mNameP2 As String


Dim mCode3 As String
Dim mName3 As String
Dim mNameP3 As String


Dim mCode4 As String
Dim mName4 As String
Dim mNameP4 As String

Dim mCode5 As String
Dim mName5 As String
Dim mNameP5 As String
Dim AccountTypes As Integer
Dim AccountTab As Integer
Dim DepitOrCreditv As Integer
Dim Differenttypev As Integer
Dim Authorityv As Integer

Dim i As Long
Dim j As Long

GrdAccount3.Rows = 1
For i = 1 To GrdAccount2.Rows - 1
    mCode = Trim(GrdAccount2.TextMatrix(i, GrdAccount2.ColIndex("Account_Serial")))
    mNameP = Trim(GrdAccount2.TextMatrix(i, GrdAccount2.ColIndex("AccountNamePar")))
    
    
    mCode2 = Trim(GrdAccount2.TextMatrix(i, GrdAccount2.ColIndex("Account_Serial2")))
    mNameP2 = Trim(GrdAccount2.TextMatrix(i, GrdAccount2.ColIndex("AccountNamePar2")))
       
    mCode3 = Trim(GrdAccount2.TextMatrix(i, GrdAccount2.ColIndex("Account_Serial3")))
    mNameP3 = Trim(GrdAccount2.TextMatrix(i, GrdAccount2.ColIndex("AccountNamePar3")))
     
    mCode4 = Trim(GrdAccount2.TextMatrix(i, GrdAccount2.ColIndex("Account_Serial4")))
    mNameP4 = Trim(GrdAccount2.TextMatrix(i, GrdAccount2.ColIndex("AccountNamePar4")))
    
    mCode5 = Trim(GrdAccount2.TextMatrix(i, GrdAccount2.ColIndex("Account_Serial5")))
    mNameP5 = Trim(GrdAccount2.TextMatrix(i, GrdAccount2.ColIndex("AccountNamePar5")))
    Dim Line1 As Long
    Dim Line2 As Long
    Dim Line3 As Long
    Dim Line4 As Long
    Dim Line5 As Long
    Dim isFound As Boolean
    isFound = False
    For j = 1 To GrdAccount2.Rows - 1
        Line1 = 0
        Line2 = 0
        Line3 = 0
        Line4 = 0
        Line5 = 0
        If mCode = Trim(GrdAccount2.TextMatrix(j, GrdAccount2.ColIndex("Account_Serial"))) And mNameP <> Trim(GrdAccount2.TextMatrix(j, GrdAccount2.ColIndex("AccountNamePar"))) Then
            isFound = True
            Line1 = i

            GoTo copyLine
        End If
        
        If mCode2 = Trim(GrdAccount2.TextMatrix(j, GrdAccount2.ColIndex("Account_Serial2"))) And mNameP2 <> Trim(GrdAccount2.TextMatrix(j, GrdAccount2.ColIndex("AccountNamePar2"))) Then
            isFound = True


            Line2 = i
            GoTo copyLine
        End If
        
        If mCode3 = Trim(GrdAccount2.TextMatrix(j, GrdAccount2.ColIndex("Account_Serial3"))) And mNameP3 <> Trim(GrdAccount2.TextMatrix(j, GrdAccount2.ColIndex("AccountNamePar3"))) Then
            isFound = True
            Line3 = i
      

            GoTo copyLine
        End If
        
        If mCode4 = Trim(GrdAccount2.TextMatrix(j, GrdAccount2.ColIndex("Account_Serial4"))) And mNameP4 <> Trim(GrdAccount2.TextMatrix(j, GrdAccount2.ColIndex("AccountNamePar4"))) Then
            isFound = True
            

            Line4 = i
            GoTo copyLine
        End If
        
        If mCode5 = Trim(GrdAccount2.TextMatrix(j, GrdAccount2.ColIndex("Account_Serial5"))) And mNameP5 <> Trim(GrdAccount2.TextMatrix(j, GrdAccount2.ColIndex("AccountNamePar5"))) Then
            Line5 = i
            isFound = True
            GoTo copyLine
        End If
    Next j
    If Not isFound Then GoTo NextRow
copyLine:
    GrdAccount3.Rows = GrdAccount3.Rows + 1
    GrdAccount3.TextMatrix(GrdAccount3.Rows - 1, GrdAccount3.ColIndex("Account_Serial")) = mCode
    GrdAccount3.TextMatrix(GrdAccount3.Rows - 1, GrdAccount3.ColIndex("AccountNamePar")) = mNameP
    GrdAccount3.TextMatrix(GrdAccount3.Rows - 1, GrdAccount3.ColIndex("Line1")) = Line1
    GrdAccount3.TextMatrix(GrdAccount3.Rows - 1, GrdAccount3.ColIndex("Account_Serial2")) = mCode2
    GrdAccount3.TextMatrix(GrdAccount3.Rows - 1, GrdAccount3.ColIndex("AccountNamePar2")) = mNameP2
    GrdAccount3.TextMatrix(GrdAccount3.Rows - 1, GrdAccount3.ColIndex("Line2")) = Line2
    
    GrdAccount3.TextMatrix(GrdAccount3.Rows - 1, GrdAccount3.ColIndex("Account_Serial3")) = mCode3
    GrdAccount3.TextMatrix(GrdAccount3.Rows - 1, GrdAccount3.ColIndex("AccountNamePar3")) = mNameP3
    GrdAccount3.TextMatrix(GrdAccount3.Rows - 1, GrdAccount3.ColIndex("Line3")) = Line3
    
    GrdAccount3.TextMatrix(GrdAccount3.Rows - 1, GrdAccount3.ColIndex("Account_Serial4")) = mCode4
    GrdAccount3.TextMatrix(GrdAccount3.Rows - 1, GrdAccount3.ColIndex("AccountNamePar4")) = mNameP4
    GrdAccount3.TextMatrix(GrdAccount3.Rows - 1, GrdAccount3.ColIndex("Line4")) = Line4
    GrdAccount3.TextMatrix(GrdAccount3.Rows - 1, GrdAccount3.ColIndex("Account_Serial5")) = mCode5
    GrdAccount3.TextMatrix(GrdAccount3.Rows - 1, GrdAccount3.ColIndex("AccountNamePar5")) = mNameP5
    GrdAccount3.TextMatrix(GrdAccount3.Rows - 1, GrdAccount3.ColIndex("Line5")) = Line5
NextRow:
 Next i
If GrdAccount3.Rows > 1 Then Frame2.Visible = True
End Sub

Private Sub Command8_Click()
Dim s As String

Dim rsDummy1 As New ADODB.Recordset
Dim mNo As String
s = " SELECT MAX(a.Account_Serial) TT FROM ACCOUNTS AS a WHERE a.Parent_Account_Code = 'r'"
rsDummy1.Open s, Cn, adOpenKeyset, adLockReadOnly
If Not rsDummy1.EOF Then
    mNo = Val(rsDummy1!TT & "") + 1
End If


s = " INSERT INTO [ACCOUNTS]( [Account_Code], [Account_Name], [Parent_Account_Code], [last_account], [cannot_del], [Account_Serial], [BasicAccount], [DateCreated], [Account_NameEng], [zmam], [mowazna], [currenct_code], [rate], [cost_center], [Sum_account], [opening_balance], [opening_balance_type], [cost_center_id], [cost_center_type], [ShowInBlanceSheet], [ActivityTypeId], [AccountTypes], [AccountTab], [DepitOrCredit], [Differenttype], [Authority], [Block], [UserGroupId], [UserId], [Branch], [Balance], [DepitBalance], [CreditBalance], [interval1], [interval2], [interval3], [interval4], [interval5], [interval6], [interval7], [interval8], [interval9], [interval10], [interval11], [interval12], [ProfitBalance], [BranchID], [Level])"
s = s & "     VALUES( 'a" & mNo & "', '«·Õ”«»«  «·‰Ÿ«„Ì…', 'r', 0, 1, '" & mNo & "', 1, '2007-10-26 03:18:00', 'Legal Accounts', NULL, 0, '1', NULL, 0, 1, 0, '1', NULL, NULL, NULL, NULL, 0, 4, NULL, NULL, 0, NULL, NULL, NULL, NULL, NULL, NULL, NULL, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, NULL, NULL);"
 Cn.Execute s
 
s = "  INSERT INTO [ACCOUNTS] ([Account_Code], [Account_Name], [Parent_Account_Code], [last_account], [cannot_del], [Account_Serial], [BasicAccount], [DateCreated], [Account_NameEng], [zmam], [mowazna], [currenct_code], [rate], [cost_center], [Sum_account], [opening_balance], [opening_balance_type], [cost_center_id], [cost_center_type], [ShowInBlanceSheet], [ActivityTypeId], [AccountTypes], [AccountTab], [DepitOrCredit], [Differenttype], [Authority], [Block], [UserGroupId], [UserId], [Branch], [Balance], [DepitBalance], [CreditBalance], [interval1], [interval2], [interval3], [interval4], [interval5], [interval6], [interval7], [interval8], [interval9], [interval10], [interval11], [interval12], [ProfitBalance], [BranchID], [Level])"
s = s & "     VALUES( 'a" & mNo & "a2', 'Ê”Ìÿ «›  «ÕÌ',  'a" & mNo & "', 0, 0, '" & mNo & "2', 0, '2012-04-30 00:00:00', 'Legal Account', NULL, 0, '1', NULL, 0, 1, 0, NULL, NULL, 0, NULL, 0, 0, 4, 0, 1, 0, 0, 0, 0, '0', NULL, NULL, NULL, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, NULL, NULL);"
 
 Cn.Execute s
 
 s = " INSERT INTO [ACCOUNTS]( [Account_Code], [Account_Name], [Parent_Account_Code], [last_account], [cannot_del], [Account_Serial], [BasicAccount], [DateCreated], [Account_NameEng], [zmam], [mowazna], [currenct_code], [rate], [cost_center], [Sum_account], [opening_balance], [opening_balance_type], [cost_center_id], [cost_center_type], [ShowInBlanceSheet], [ActivityTypeId], [AccountTypes], [AccountTab], [DepitOrCredit], [Differenttype], [Authority], [Block], [UserGroupId], [UserId], [Branch], [Balance], [DepitBalance], [CreditBalance], [interval1], [interval2], [interval3], [interval4], [interval5], [interval6], [interval7], [interval8], [interval9], [interval10], [interval11], [interval12], [ProfitBalance], [BranchID], [Level])"
s = s & "     VALUES( 'a" & mNo & "a2a1', 'Ê”Ìÿ «›  «ÕÌ ··Õ”«»« ', 'a" & mNo & "a2', 0, 0, '" & mNo & "201', 0, '2014-01-26 00:00:00', 'Opening broker accounts', NULL, 0, '1', NULL, 0, 1, 0, NULL, NULL, 0, NULL, 0, 0, 4, 0, 1, 0, 0, 0, 0, '0', NULL, NULL, NULL, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, NULL, NULL);"
Cn.Execute s

s = " INSERT INTO [ACCOUNTS]( [Account_Code], [Account_Name], [Parent_Account_Code], [last_account], [cannot_del], [Account_Serial], [BasicAccount], [DateCreated], [Account_NameEng], [zmam], [mowazna], [currenct_code], [rate], [cost_center], [Sum_account], [opening_balance], [opening_balance_type], [cost_center_id], [cost_center_type], [ShowInBlanceSheet], [ActivityTypeId], [AccountTypes], [AccountTab], [DepitOrCredit], [Differenttype], [Authority], [Block], [UserGroupId], [UserId], [Branch], [Balance], [DepitBalance], [CreditBalance], [interval1], [interval2], [interval3], [interval4], [interval5], [interval6], [interval7], [interval8], [interval9], [interval10], [interval11], [interval12], [ProfitBalance], [BranchID], [Level])"
s = s & "    VALUES( 'a" & mNo & "a2a1a1', 'Ê”Ìÿ «›  «ÕÌ Õ”«»« ', 'a" & mNo & "a2a1', 0, 0, '" & mNo & "20101', 0, '2014-01-26 00:00:00', 'Opening broker accounts', NULL, 0, '1', NULL, 0, 1, 0, NULL, NULL, 0, NULL, 0, 0, 4, 0, 1, 0, 0, 0, 0, '0', NULL, NULL, NULL, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, NULL, NULL);"
Cn.Execute s




    s = " INSERT INTO [ACCOUNTS]( [Account_Code], [Account_Name], [Parent_Account_Code], [last_account], [cannot_del], [Account_Serial], [BasicAccount], [DateCreated], [Account_NameEng], [zmam], [mowazna], [currenct_code], [rate], [cost_center], [Sum_account], [opening_balance], [opening_balance_type], [cost_center_id], [cost_center_type], [ShowInBlanceSheet], [ActivityTypeId], [AccountTypes], [AccountTab], [DepitOrCredit], [Differenttype], [Authority], [Block], [UserGroupId], [UserId], [Branch], [Balance], [DepitBalance], [CreditBalance], [interval1], [interval2], [interval3], [interval4], [interval5], [interval6], [interval7], [interval8], [interval9], [interval10], [interval11], [interval12], [ProfitBalance], [BranchID], [Level])"
    s = s & "    VALUES ('a" & mNo & "a2a1a1a1', 'Ê”Ìÿ «›  «ÕÏ Õ”«»« ', 'a" & mNo & "a2a1a1', 1, 0, '" & mNo & "20101001', 0, '2014-01-26 00:00:00', 'Opening broker accounts', NULL, 0, '1', NULL, 0, 1, 0, NULL, NULL, 0, NULL, 0, 0, 4, 0, 1, 0, 0, 0, 0, '0', NULL, NULL, NULL, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, NULL, NULL);"
 Cn.Execute s

 s = " INSERT INTO [ACCOUNTS]( [Account_Code], [Account_Name], [Parent_Account_Code], [last_account], [cannot_del], [Account_Serial], [BasicAccount], [DateCreated], [Account_NameEng], [zmam], [mowazna], [currenct_code], [rate], [cost_center], [Sum_account], [opening_balance], [opening_balance_type], [cost_center_id], [cost_center_type], [ShowInBlanceSheet], [ActivityTypeId], [AccountTypes], [AccountTab], [DepitOrCredit], [Differenttype], [Authority], [Block], [UserGroupId], [UserId], [Branch], [Balance], [DepitBalance], [CreditBalance], [interval1], [interval2], [interval3], [interval4], [interval5], [interval6], [interval7], [interval8], [interval9], [interval10], [interval11], [interval12], [ProfitBalance], [BranchID], [Level])"
s = s & "        VALUES('a" & mNo & "a2a1a1a2', 'Ê”Ìÿ «›  «ÕÏ „Œ“Ê‰', 'a" & mNo & "a2a1a1', 1, 0, '" & mNo & "20101002', 0, '2014-01-26 00:00:00', 'Opening Broker Inventoory', NULL, 0, '1', NULL, 0, 1, 0, NULL, NULL, 0, NULL, 0, 0, 4, 0, 1, 0, 0, 0, 0, '0', NULL, NULL, NULL, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, NULL, NULL);"
 Cn.Execute s

 s = " INSERT INTO [ACCOUNTS]( [Account_Code], [Account_Name], [Parent_Account_Code], [last_account], [cannot_del], [Account_Serial], [BasicAccount], [DateCreated], [Account_NameEng], [zmam], [mowazna], [currenct_code], [rate], [cost_center], [Sum_account], [opening_balance], [opening_balance_type], [cost_center_id], [cost_center_type], [ShowInBlanceSheet], [ActivityTypeId], [AccountTypes], [AccountTab], [DepitOrCredit], [Differenttype], [Authority], [Block], [UserGroupId], [UserId], [Branch], [Balance], [DepitBalance], [CreditBalance], [interval1], [interval2], [interval3], [interval4], [interval5], [interval6], [interval7], [interval8], [interval9], [interval10], [interval11], [interval12], [ProfitBalance], [BranchID], [Level])"
 s = s & "      VALUES('a" & mNo & "a2a1a1a3', 'Ê”Ìÿ «›  «ÕÏ «’Ê· À«» …', 'a" & mNo & "a2a1a1', 1, 0, '" & mNo & "20101003', 0, '2014-01-26 00:00:00', 'Opening Broker Fixed Assets', NULL, 0, '1', NULL, 0, 1, 0, NULL, NULL, 0, NULL, 0, 0, 4, 0, 1, 0, 0, 0, 0, '0', NULL, NULL, NULL, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, NULL, NULL);"
 Cn.Execute s

 s = " INSERT INTO [ACCOUNTS]( [Account_Code], [Account_Name], [Parent_Account_Code], [last_account], [cannot_del], [Account_Serial], [BasicAccount], [DateCreated], [Account_NameEng], [zmam], [mowazna], [currenct_code], [rate], [cost_center], [Sum_account], [opening_balance], [opening_balance_type], [cost_center_id], [cost_center_type], [ShowInBlanceSheet], [ActivityTypeId], [AccountTypes], [AccountTab], [DepitOrCredit], [Differenttype], [Authority], [Block], [UserGroupId], [UserId], [Branch], [Balance], [DepitBalance], [CreditBalance], [interval1], [interval2], [interval3], [interval4], [interval5], [interval6], [interval7], [interval8], [interval9], [interval10], [interval11], [interval12], [ProfitBalance], [BranchID], [Level])"
s = s & "   VALUES('a" & mNo & "a2a1a1a4', 'Ê”Ìÿ «›  «ÕÏ ··‰ﬁœÌ…', 'a" & mNo & "a2a1a1', 1, 0, '" & mNo & "20101004', 0, '2014-01-26 00:00:00', 'Opening Broker Cash', NULL, 0, '1', NULL, 0, 1, 0, NULL, NULL, 0, NULL, 0, 0, 4, 0, 1, 0, 0, 0, 0, '0', NULL, NULL, NULL, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, NULL, NULL);"
Cn.Execute s

    s = " INSERT INTO [ACCOUNTS]( [Account_Code], [Account_Name], [Parent_Account_Code], [last_account], [cannot_del], [Account_Serial], [BasicAccount], [DateCreated], [Account_NameEng], [zmam], [mowazna], [currenct_code], [rate], [cost_center], [Sum_account], [opening_balance], [opening_balance_type], [cost_center_id], [cost_center_type], [ShowInBlanceSheet], [ActivityTypeId], [AccountTypes], [AccountTab], [DepitOrCredit], [Differenttype], [Authority], [Block], [UserGroupId], [UserId], [Branch], [Balance], [DepitBalance], [CreditBalance], [interval1], [interval2], [interval3], [interval4], [interval5], [interval6], [interval7], [interval8], [interval9], [interval10], [interval11], [interval12], [ProfitBalance], [BranchID], [Level])"
    s = s & "       VALUES('a" & mNo & "a2a1a1a5', 'Ê”Ìÿ «›  «ÕÏ ··»‰Êﬂ', 'a" & mNo & "a2a1a1', 1, 0, '" & mNo & "20101005', 0, '2014-01-26 00:00:00', 'Opening Broker Banks', NULL, 0, '1', NULL, 0, 1, 0, NULL, NULL, 0, NULL, 0, 0, 4, 0, 1, 0, 0, 0, 0, '0', NULL, NULL, NULL, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, NULL, NULL);"
 Cn.Execute s

 s = " INSERT INTO [ACCOUNTS]( [Account_Code], [Account_Name], [Parent_Account_Code], [last_account], [cannot_del], [Account_Serial], [BasicAccount], [DateCreated], [Account_NameEng], [zmam], [mowazna], [currenct_code], [rate], [cost_center], [Sum_account], [opening_balance], [opening_balance_type], [cost_center_id], [cost_center_type], [ShowInBlanceSheet], [ActivityTypeId], [AccountTypes], [AccountTab], [DepitOrCredit], [Differenttype], [Authority], [Block], [UserGroupId], [UserId], [Branch], [Balance], [DepitBalance], [CreditBalance], [interval1], [interval2], [interval3], [interval4], [interval5], [interval6], [interval7], [interval8], [interval9], [interval10], [interval11], [interval12], [ProfitBalance], [BranchID], [Level])"
s = s & "       VALUES('a" & mNo & "a2a1a1a6', 'Ê”Ìÿ «›  «ÕÏ ⁄„·«¡', 'a" & mNo & "a2a1a1', 1,  0, '" & mNo & "20101006', 0, '2014-01-26 00:00:00', 'Opening Broker Customers', NULL, 0, '1', NULL, 0, 1, 0, NULL, NULL, 0, NULL, 0, 0, 4, 0, 1, 0, 0, 0, 0, '0', NULL, NULL, NULL, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, NULL, NULL);"
 Cn.Execute s

 s = " INSERT INTO [ACCOUNTS]( [Account_Code], [Account_Name], [Parent_Account_Code], [last_account], [cannot_del], [Account_Serial], [BasicAccount], [DateCreated], [Account_NameEng], [zmam], [mowazna], [currenct_code], [rate], [cost_center], [Sum_account], [opening_balance], [opening_balance_type], [cost_center_id], [cost_center_type], [ShowInBlanceSheet], [ActivityTypeId], [AccountTypes], [AccountTab], [DepitOrCredit], [Differenttype], [Authority], [Block], [UserGroupId], [UserId], [Branch], [Balance], [DepitBalance], [CreditBalance], [interval1], [interval2], [interval3], [interval4], [interval5], [interval6], [interval7], [interval8], [interval9], [interval10], [interval11], [interval12], [ProfitBalance], [BranchID], [Level])"
 s = s & "      VALUES('a" & mNo & "a2a1a1a7', 'Ê”Ìÿ «›  «ÕÏ „Ê—œÌ‰', 'a" & mNo & "a2a1a1', 1, 0, '" & mNo & "20101007', 0, '2014-01-26 00:00:00', 'Opening Broker Vendors', NULL, 0, '1', NULL, 0, 1, 0, NULL, NULL, 0, NULL, 0, 0, 4, 0, 1, 0, 0, 0, 0, '0', NULL, NULL, NULL, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, NULL, NULL);"
 Cn.Execute s

 s = " INSERT INTO [ACCOUNTS]( [Account_Code], [Account_Name], [Parent_Account_Code], [last_account], [cannot_del], [Account_Serial], [BasicAccount], [DateCreated], [Account_NameEng], [zmam], [mowazna], [currenct_code], [rate], [cost_center], [Sum_account], [opening_balance], [opening_balance_type], [cost_center_id], [cost_center_type], [ShowInBlanceSheet], [ActivityTypeId], [AccountTypes], [AccountTab], [DepitOrCredit], [Differenttype], [Authority], [Block], [UserGroupId], [UserId], [Branch], [Balance], [DepitBalance], [CreditBalance], [interval1], [interval2], [interval3], [interval4], [interval5], [interval6], [interval7], [interval8], [interval9], [interval10], [interval11], [interval12], [ProfitBalance], [BranchID], [Level])"
s = s & "       VALUES('a" & mNo & "a2a1a1a8', 'Ê”Ìÿ «›  «ÕÏ „ÊŸ›Ì‰', 'a" & mNo & "a2a1a1', 1, 0,'" & mNo & "20101008', 0, '2014-01-26 00:00:00', 'Opening Broker Emloyee', NULL, 0, '1', NULL, 0, 1,0, NULL, NULL, 0, NULL, 0, 0, 4, 0, 1, 0, 0, 0, 0, '0', NULL, NULL, NULL, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, NULL, NULL);"
 Cn.Execute s


End Sub

Private Sub Command9_Click()

chkBalanceOnly.Value = vbChecked
grdMan.Visible = False
grdManBal.Visible = True

Dim rsDummy As New ADODB.Recordset
Dim mName As String
Dim mNamee As String
Dim mCusId As Long
Dim i As Long
Dim OpenBalance As Double
Dim rsDummy2 As New ADODB.Recordset
For i = 1 To grdManBal.Rows - 1
    mName = Trim(grdManBal.TextMatrix(i, grdManBal.ColIndex("CusName")))
    mNamee = Trim(grdManBal.TextMatrix(i, grdManBal.ColIndex("CusNamee")))
    mOpenBalance = Val(grdManBal.TextMatrix(i, grdManBal.ColIndex("OpenBalance")))
    s = "Select * from TblCustemers WHere (CusName Like '%" & mName & "%' Or CusNamee Like '%" & mNamee & "%' )"
    Set rsDummy = New ADODB.Recordset
    rsDummy.Open s, Cn, adOpenKeyset, adLockOptimistic
   ' «Ê· Õ«·… «·⁄„Ì· „ÊÃÊœ ›Ï «·«ﬂ”Ì· Ê€Ì— „ÊÃÊœ ›Ï «·œ« « Ì „ «÷«› Â «·Ï ÃœÊ· „ƒﬁ  ··‰Ÿ— ›ÌÂ
    If rsDummy.EOF Then
        s = "Select * from TblCustTemp Where Id = -1"
        Set rsDummy2 = New ADODB.Recordset
       
 
        
        rsDummy2.Open s, Cn, adOpenKeyset, adLockOptimistic
        rsDummy2.AddNew
        rsDummy2!Name = mName
        rsDummy2!NameE = mNamee
        If OpenBalance < 0 Then
            rsDummy2!OpenBalanceType = 1
        Else
            rsDummy2!OpenBalanceType = 0
        End If
        rsDummy2!OpenBalance = Abs(mOpenBalance)
        rsDummy2.Update
        
    Else
    '⁄„Ì· „ÊÃÊœ ›Ï «·œ« « ÌﬁÊ„ » ÕœÌÀ —’ÌœÂ „‰ «·«ﬂ”Ì·
        rsDummy!OpenBalance = Abs(mOpenBalance)
        If OpenBalance < 0 Then
            rsDummy!OpenBalanceType = 1
        Else
            rsDummy!OpenBalanceType = 0
        End If
        
        rsDummy.Update
    End If
Next
End Sub

Private Sub Form_Load()
txtDbPath = GetSetting("ConvertToAccess", "Setting", "DbPath", "DatabasePath")
txtTableName = GetSetting("ConvertToAccess", "Setting", "TableName", "TableName")
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

    
 txtSql = "Select  (Select Max(Account_Code) from ACCOUNTS acc where acc.Parent_Account_Code  =ExpensesType.parent_account ), * from ExpensesType " & vbNewLine
 txtSql = txtSql & " Inner Join accounts On accounts.account_code =  ExpensesType.parent_account " & vbNewLine
 txtSql = txtSql & " where ExpensesType.account_code Not In (Select account_code from accounts)" & vbNewLine
 txtSql = txtSql & " Order By parent_account"


        StrSQL = "SELECT StoreID,StoreName From TblStore where 1=1"
 




    GetComboData DCboStoreName, StrSQL
     GetComboData DcStore10, StrSQL
    
        StrSQL = "SELECT branch_id,branch_name FROM TblBranchesData  where 1=1"
 


    

    GetComboData cmbBranch, StrSQL
    'StrSQL = "Select ACCOUNTS.Account_Code,ACCOUNTS.Account_Name,*  from ACCOUNTS where last_account = 0 " & ' and BasicAccount = 1 "
    StrSQL = "Select ACCOUNTS.Account_Code,ACCOUNTS.Account_Name,*  from ACCOUNTS where last_account = 0 "
    GetComboData DboParentAccount2, StrSQL
      
    
   StrComboList = "#0;«·„—Õ·… «·«» œ«∆Ì…|#1;«·„—Õ·… «·„ Ê”ÿ…|#2;«·„—Õ·… «·À«‰ÊÌ…|#3;«·„—Õ·… «·„Ã„⁄…"
'    StrSQL = "select id,prifix from coding  where  FIELD_no=" & FIELD_no & "  and branch_no=" & branch_no & " Order By prifix"
    
    LoadMainSystemOptions2
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

Private Sub optGetFromAcc_Click(Index As Integer)
Dim ss As String
Select Case Index
Case 0
  TabMain.CurrTab = 0
  ss = "–„„ «·⁄«„·Ì‰" & "-" & "–„„"
Case 1
    TabMain.CurrTab = 1
    
    Option2.Value = True
    Option1.Value = False
    ss = "⁄„·«¡" & "-" & "„œÌ‰Ê‰ „ ‰Ê⁄Ê‰"
Case 2
    TabMain.CurrTab = 1
    Option2.Value = False
    Option1.Value = True
    ss = "„Ê—œÌ‰" & "-" & "„Ê—œÊ‰" & "-" & "„œÌ‰Ê‰ „ ‰Ê⁄Ê‰"
Case 3
    TabMain.CurrTab = 6
   ss = "»‰ﬂ" & "-" & "»‰Êﬂ" & "-" & "«·»‰ﬂ" & "-" & "«·»‰ﬂ"
Case 4
    TabMain.CurrTab = 7
    
    ss = "⁄Âœ" & "-" & "‰ﬁœÌ… »«·Œ“Ì‰…" & "-" & "„œÌ‰Ê‰ „ ‰Ê⁄Ê‰"
Case 5
    TabMain.CurrTab = 8
    ss = "„’—Ê›« " & "-" & " ﬂ«·Ì› «·„‘«—Ì⁄" & "-" & "„’«—Ì›" & "-" & "Âœ«Ì« Ê⁄Ì‰« " & " ﬂ«·Ì› «·«‘ —«ﬂ" & "-" & "« ⁄«» Ê«” ‘«—« " & "-" & "œ„€… ° »—Ìœ°  ·Ì›Ê‰« " & "-" & " √„Ì‰ Õ—Ìﬁ Ê”—ﬁ…" & "√⁄„«· ’Ì«‰…" & "-" & " ﬂ·›… «·√ÃÊ—"
Case 6
    TabMain.CurrTab = 9
    ss = "«·„Œ“Ê‰" & "-" & "«·„Œ«“‰"
End Select
txtKeySerach = ss
mIndex = TabMain.CurrTab
    
End Sub

Private Sub TabMain_Click()
mIndex = TabMain.CurrTab



End Sub



Public Function GetFixedAssetsGroupAccount(GroupID As Integer, _
                                           Optional account_type_code As Integer, _
                                           Optional branch_id As Integer = 0, _
                                           Optional ByRef Account_Codex As String, _
                                           Optional ByRef account_name As String, _
                                           Optional ByRef Percentage1 As Integer = 0, _
                                           Optional ByRef Percentage2 As Integer = 0, _
                                           Optional ByRef DepType As Integer = 0, _
                                           Optional ByRef Account_code As String, _
                                           Optional ByRef Account_code1 As String, _
                                           Optional ByRef Account_code2 As String, _
                                           Optional ByRef Account_code3 As String, _
                                           Optional ByRef Account_code4 As String)
    Dim rs As ADODB.Recordset
    Dim Rs1 As ADODB.Recordset

    Dim sql As String
    Dim str As String
    Set Rs1 = New ADODB.Recordset
    sql = "SELECT * from FixedAssetsGroup where GroupID=" & GroupID
    Rs1.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Rs1.RecordCount > 0 Then
        Percentage1 = IIf(Not IsNumeric(Rs1("Percentage1").Value), 0, Rs1("Percentage1").Value)
        Percentage2 = IIf(Not IsNumeric(Rs1("Percentage2").Value), 0, Rs1("Percentage2").Value)
        DepType = IIf(IsNull(Rs1("DepType").Value), 0, Rs1("DepType").Value)
        Account_code = IIf(IsNull(Rs1("Account_Code").Value), "", Rs1("Account_Code").Value)
        Account_code1 = IIf(IsNull(Rs1("Account_Code1").Value), "", Rs1("Account_Code1").Value)
        Account_code2 = IIf(IsNull(Rs1("Account_Code2").Value), "", Rs1("Account_Code2").Value)
        Account_code3 = IIf(IsNull(Rs1("Account_Code3").Value), "", Rs1("Account_Code3").Value)
        Account_code4 = IIf(IsNull(Rs1("Account_Code4").Value), "", Rs1("Account_Code4").Value)
  
    End If
 
    'Set rs = New ADODB.Recordset
    'sql = "SELECT     dbo.ACCOUNTS.Account_Name, dbo.FixedAssetsGroupsAccount.account_code, dbo.FixedAssetsGroupsAccount.branch_id, dbo.FixedAssetsGroupsAccount.group_id, " & _
    '               "       dbo.FixedAssetsGroupsAccount.account_type_code " & _
    '" FROM         dbo.FixedAssetsGroupsAccount INNER JOIN" & _
    ' "                     dbo.ACCOUNTS ON dbo.FixedAssetsGroupsAccount.account_code = dbo.ACCOUNTS.Account_Code"
 
    'sql = sql & " where group_id =" & GroupID & " and account_type_code='" & account_type_code & "' "
    'If branch_id <> 0 Then
    'sql = sql & " and  branch_id=" & branch_id
    'End If
    'rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    'If rs.RecordCount > 0 Then
    'account_name = IIf(IsNull(rs("Account_Name").value), "", rs("Account_Name").value)
    ' Account_Code = IIf(IsNull(rs("account_code").value), "", rs("account_code").value)
    ' End If
    'rs.Close
    'Rs1.Close
End Function






Function CreateJL(Optional ByVal DCGroup As Long = 0, Optional ByVal dcBranch As Long = 0, Optional ByVal mIID As Long = 0, Optional ByVal mName As String = "", Optional ByVal TxtPurchasePrice As Double = 0) As Boolean
    CreateJL = False
    Dim LngDevID As Long
    Dim DepitAccount As String
    Dim CreditAccount1 As String
    Dim CreditAccount2 As String
    Dim Msg As String
    'GetFixedAssetsGroupAccount Val(DCGroup.BoundText), 24, Val(Me.DcBranch.BoundText), DepitAccount    'Õ”«» «·«’·
    'GetFixedAssetsGroupAccount Val(DCGroup.BoundText), 26, Val(Me.DcBranch.BoundText), CreditAccount1    '„Ã„⁄ «·«Â·«ﬂ

    Dim Account_code As String
    Dim Account_code2 As String

    GetFixedAssetsGroupAccount Val(DCGroup), , Val(dcBranch), , , , , , Account_code, , Account_code2
    DepitAccount = Account_code
    CreditAccount1 = Account_code2

    CreditAccount2 = get_account_code_branch(41, Val(dcBranch))

    If CreditAccount2 = "NO branch" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·« ÌÊÃœ Õ”«»«  ·Â–« «·›—⁄"
        Else
            Msg = "No Accounts For This Branch"
        End If

        MsgBox Msg, vbCritical
        CreateJL = False
        Exit Function

    ElseIf CreditAccount2 = "NO account" Then

        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "Õ”«» Ê”Ìÿ «›  «ÕÌ ··«’Ê· €Ì— „Õœœ ›Ï «·›—⁄"
        Else
            Msg = "Fixed Asset Opening Balance Account Not Defined In this Branch"
        End If

        MsgBox Msg, vbCritical
        CreateJL = False
        Exit Function
    End If

    Dim sql As String

    'sql = "Delete   from notes where NoteID=" & Val(TxtNoteID.text)
    'Cn.Execute sql
    '«‰‘«¡ «·ﬁÌÊœ
'    If Option1.Value = True Or (TXT24.Text) = "" Then    'ÃœÌœ
'        CreateJL = True
'        Exit Function
'    Else
        '   Dim RsNotes As ADODB.Recordset
        '   Dim RsDev As ADODB.Recordset
        '   Dim NoteID As String
        '   Set RsNotes = New ADODB.Recordset
        Dim StrSQL As String
   
        '   RsNotes.Open "Notes", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    
        '        Set RsDev = New ADODB.Recordset
        '        RsDev.Open "DOUBLE_ENTREY_VOUCHERS", Cn, adOpenStatic, adLockOptimistic, adCmdTable
'        If Me.TxtModFlg.Text = "N" Then
'            '                RsNotes.AddNew
'            '    my_branch = Val(Me.dcBranch.BoundText)
'            '                RsNotes("NoteID").value = CStr(TXTNoteID.text)
'            '                RsNotes("Note_Value").value = Val(TxtPurchasePrice.text)
'            '               RsNotes("branch_no").value = Val(Me.dcBranch.BoundText)
'            '                RsNotes("Remark").value = ""
'            '                RsNotes("NoteType").value = 90
'            '                RsNotes("NoteDate").value = XPDtbTrans.value
'            '                RsNotes("UserID").value = user_id
'            '                RsNotes("NoteSerial").value = Trim$(Me.TxtNoteSerial.text) '„”·”· «·ﬁÌœ
'            '                RsNotes("numbering_type").value = sand_numbering_type(0) '‰Ê⁄  —ﬁÌ„ ”‰œ «·ﬁÌœ
'            '                RsNotes("sanad_year").value = year(Date)
'            '                RsNotes("sanad_month").value = Month(Date)
'            ''                RsNotes("note_value_by_characters").value = WriteNo(Format(Val(TxtPurchasePrice.text), "0.00"), 0, True, ".")
'            '                RsNotes.update
'        Else
'       '     Cn.Execute "Delete DOUBLE_ENTREY_VOUCHERS  Where Notes_ID=" & val(txtNoteID.text)
'            StrSQL = "delete From DOUBLE_ENTREY_VOUCHERS1 where opening_balance_voucher_id=" & Val(txtopening_balance_voucher_id.Text)
'            Cn.Execute StrSQL, , adExecuteNoRecords
'
'        End If

        Dim des As String
        Dim LngOpenID  As Long
        LngOpenID = 1

        'If Option2.Value = True And Opt(1).Value = True And Me.cStatus.ListIndex = -1 Then '«›  «ÕÌ Ê ·Ì” ·Â «Â·«ﬂ
            ' LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)
  
            LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS1", "Double_Entry_Vouchers_ID", "", True)
            
            If SystemOptions.UserInterface <> ArabicInterface Then
                des = "Fixed Asset Opening Balance For Asset  " & mName & "  And have No Depreciation '"
            Else
                des = "»‰«¡ ⁄·Ï —’Ìœ «›  «ÕÌ ··«’· " & mName & "  Ê·Ì” ·Â «Â·«ﬂ -ﬁÌ„… «·«’·'"
            End If
            Dim user_id As Long
            user_id = 1
            If AddNewDev(LngDevID, 0, DepitAccount, Val(TxtPurchasePrice), 0, des, LngOpenID, , , SystemOptions.SysCurrentAccountIntervalID, Me.txtDate.Value, user_id, , , , , , , , , , , , , True, Val(txtopening_balance_voucher_id), Val(mIID), Val(DCGroup), Val(dcBranch), Val(dcBranch)) = False Then
                GoTo ErrTrap
            End If
 
            If AddNewDev(LngDevID, 1, CreditAccount2, Val(TxtPurchasePrice), 1, des, LngOpenID, , , SystemOptions.SysCurrentAccountIntervalID, Me.txtDate.Value, user_id, , , , , , , , , , , , , True, Val(txtopening_balance_voucher_id), Val(mIID), Val(DCGroup), Val(dcBranch), Val(dcBranch)) = False Then
                GoTo ErrTrap
            End If
 
            '„œÌ‰
            '     If ModAccounts.AddNewDev(LngDevID, 0, _
                  DepitAccount, Val(TxtPurchasePrice.text), 0, _
                  des, Val(Me.TxtNoteID), , , _
                  SystemOptions.SysCurrentAccountIntervalID, Me.XPDtbTrans.value, user_id, , , , , , , , , , _
                  , , , , , Val(mIID), Val(Me.DCGroup.BoundText), Val(Me.DcBranch.BoundText)) = False Then
            '         GoTo ErrTrap
                    
            '    End If
            '            œ«∆‰ 1
            '  If ModAccounts.AddNewDev(LngDevID, 1, _
               CreditAccount2, Val(TxtPurchasePrice.text), 1, _
               des, Val(Me.TXTNoteID), , , _
               SystemOptions.SysCurrentAccountIntervalID, Me.XPDtbTrans.value, user_id, , , , , , , , , , _
               , , , , , Val(mIID), Val(Me.DCGroup.BoundText), Val(Me.dcBranch.BoundText)) = False Then
            '         GoTo ErrTrap
                    
            '    End If
'
'        ElseIf Option2.Value = True And Opt(0).Value = True And Me.cStatus.ListIndex = 0 Then '  Ê«·Õ«·… Ã«—Ì «·«Â·«ﬂ' '«›  «ÕÌ Ê   ·Â «Â·«ﬂ
'
'            ' LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)
'            LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS1", "Double_Entry_Vouchers_ID", "", True)
'
'            If SystemOptions.UserInterface <> ArabicInterface Then
'                des = "Fixed Asset Opening Balance For Asset  " & mName & "  And have Depreciation '"
'            Else
'                des = "»‰«¡ ⁄·Ï —’Ìœ «›  «ÕÌ ··«’· " & mName & "    ·Â «Â·«ﬂ '"
'            End If
'
'            If ModAccounts.AddNewDev(LngDevID, 1, DepitAccount, Val(Me.TxtPurchasePrice.Text), 0, des, LngOpenID, , , SystemOptions.SysCurrentAccountIntervalID, Me.XPDtbTrans.Value, user_id, , , , , , , , , , , , , True, Val(Me.txtopening_balance_voucher_id.Text), Val(mIID), Val(Me.DCGroup.BoundText), Val(Me.dcBranch.BoundText), Val(Me.dcBranch.BoundText)) = False Then
'                GoTo ErrTrap
'            End If
'
'            '„œÌ‰
'            '            If ModAccounts.AddNewDev(LngDevID, 1, _
'                         DepitAccount, Val(TxtPurchasePrice.text), 0, _
'                         des, Val(Me.TxtNoteID), , , _
'                         SystemOptions.SysCurrentAccountIntervalID, Me.XPDtbTrans.value, user_id, , , , , , , , , , _
'                         , , , , , Val(mIID), Val(Me.DCGroup.BoundText), Val(Me.DcBranch.BoundText)) = False Then
'            '       GoTo ErrTrap
'
'            '            End If
'            '             „Ã„⁄ «·«Â·«ﬂ œ«∆‰ 1
'            If Val(TxtAccDepreciation.Text) > 0 Then
'                If ModAccounts.AddNewDev(LngDevID, 2, CreditAccount1, Val(Me.TxtAccDepreciation.Text), 1, des, LngOpenID, , , SystemOptions.SysCurrentAccountIntervalID, Me.XPDtbTrans.Value, user_id, , , , , , , , , , , , , True, Val(Me.txtopening_balance_voucher_id.Text), Val(mIID), Val(Me.DCGroup.BoundText), Val(Me.dcBranch.BoundText), Val(Me.dcBranch.BoundText)) = False Then
'                    GoTo ErrTrap
'                End If
'
'                '          If ModAccounts.AddNewDev(LngDevID, 2, _
'                           CreditAccount1, Val(TxtAccDepreciation.text), 1, _
'                           des, Val(Me.TxtNoteID), , , _
'                           SystemOptions.SysCurrentAccountIntervalID, Me.XPDtbTrans.value, user_id, , , , , , , , , , _
'                           , , , , , Val(mIID), Val(Me.DCGroup.BoundText), Val(Me.DcBranch.BoundText)) = False Then
'                '    GoTo ErrTrap
'
'                '            End If
'            End If
'
'            '            Ê”Ìÿ «›  «ÕÌ 2
'            If Val(Me.TxtPurchasePrice.Text) - Val(Me.TxtAccDepreciation.Text) > 0 Then
'                If ModAccounts.AddNewDev(LngDevID, 3, CreditAccount2, Val(Me.TxtPurchasePrice.Text) - Val(Me.TxtAccDepreciation.Text), 1, des, LngOpenID, , , SystemOptions.SysCurrentAccountIntervalID, Me.XPDtbTrans.Value, user_id, , , , , , , , , , , , , True, Val(Me.txtopening_balance_voucher_id.Text), Val(mIID), Val(Me.DCGroup.BoundText), Val(Me.dcBranch.BoundText), Val(Me.dcBranch.BoundText)) = False Then
'                    GoTo ErrTrap
'                End If
'            End If
'
'
''ﬁÌ„… «·«’· ﬂŒ—œ…
'      If Val(TxtKhordaPrice) > 0 Then
'   '             If ModAccounts.AddNewDev(LngDevID, 3, DepitAccount, val(Me.TxtKhordaPrice.text), 1, des, LngOpenID, , , SystemOptions.SysCurrentAccountIntervalID, Me.XPDtbTrans.value, user_id, , , , , , , , , , , , , True, val(Me.txtopening_balance_voucher_id.text), val(mIID), val(Me.DCGroup.BoundText), val(Me.DcBranch.BoundText), val(Me.DcBranch.BoundText)) = False Then
'   '                 GoTo ErrTrap
'   '             End If
'            End If
'
'
'            '    If ModAccounts.AddNewDev(LngDevID, 3, _
'                 CreditAccount2, Val(TxtPurchasePrice.text) - Val(TxtAccDepreciation.text), 1, _
'                 des, Val(Me.TxtNoteID), , , _
'                 SystemOptions.SysCurrentAccountIntervalID, Me.XPDtbTrans.value, user_id, , , , , , , , , , _
'                 , , , , , Val(mIID), Val(Me.DCGroup.BoundText), Val(Me.DcBranch.BoundText)) = False Then
'            '   GoTo ErrTrap
'
'            '      End If
'
'        ElseIf Option2.Value = True And Opt(0).Value = True And Me.cStatus.ListIndex = 2 Then  '«›  «ÕÌ Ê·Â «Â·«ﬂ  Ê  „   «·«Â·«ﬂ
'            '     LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "", True)
'            LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS1", "Double_Entry_Vouchers_ID", "", True)
'
'            If SystemOptions.UserInterface <> ArabicInterface Then
'                des = "Fixed Asset Fully Depreciation , Name IS  " & mName
'            Else
'                des = "«’· «›  «ÕÌ Ê „ «Â·«ﬂ «·«’· " & mName
'            End If
'
'            If ModAccounts.AddNewDev(LngDevID, 0, DepitAccount, Val(Me.TxtKhordaPrice.Text), 0, des, LngOpenID, , , SystemOptions.SysCurrentAccountIntervalID, Me.XPDtbTrans.Value, user_id, , , , , , , , , , , , , True, Val(Me.txtopening_balance_voucher_id.Text), Val(mIID), Val(Me.DCGroup.BoundText), Val(Me.dcBranch.BoundText), Val(Me.dcBranch.BoundText)) = False Then
'                GoTo ErrTrap
'            End If
'
'            '„œÌ‰
'            '            If ModAccounts.AddNewDev(LngDevID, 0, _
'                         DepitAccount, Val(TxtKhordaPrice.text), 0, _
'                         des, Val(Me.TxtNoteID), , , _
'                         SystemOptions.SysCurrentAccountIntervalID, Me.XPDtbTrans.value, user_id, , , , , , , , , , _
'                         , , , , , Val(mIID), Val(Me.DCGroup.BoundText), Val(Me.dcBranch.BoundText)) = False Then
'            '                 GoTo ErrTrap
'
'            '            End If
'            If ModAccounts.AddNewDev(LngDevID, 3, CreditAccount2, Val(Me.TxtKhordaPrice.Text), 1, des, LngOpenID, , , SystemOptions.SysCurrentAccountIntervalID, Me.XPDtbTrans.Value, user_id, , , , , , , , , , , , , True, Val(Me.txtopening_balance_voucher_id.Text), Val(mIID), Val(Me.DCGroup.BoundText), Val(Me.dcBranch.BoundText), Val(Me.dcBranch.BoundText)) = False Then
'                GoTo ErrTrap
'            End If
'
'            '            œ«∆‰ 1
'            '          If ModAccounts.AddNewDev(LngDevID, 1, _
'                       CreditAccount2, Val(TxtKhordaPrice.text), 1, _
'                       des, Val(Me.TxtNoteID), , , _
'                       SystemOptions.SysCurrentAccountIntervalID, Me.XPDtbTrans.value, user_id, , , , , , , , , , _
'                       , , , , , Val(mIID), Val(Me.DCGroup.BoundText), Val(Me.dcBranch.BoundText)) = False Then
'            '                 GoTo ErrTrap
'
'            '            End If
'
'        End If
'    End If

    CreateJL = True
    Exit Function
ErrTrap:
    CreateJL = False
End Function

Function GetStatus_id() As Integer
    GetStatus_id = 1
End Function

Function GetDefaultAge() As Integer
    GetDefaultAge = 10
End Function

Function getLastDepreciationDate() As Date
    getLastDepreciationDate = "05-05-2012"
End Function

Function GetInstallmentsInformations(FixedassetId As Integer, Optional ByRef noOfInstallments As Integer, Optional ByRef EXEInstallments As Integer, Optional ByRef RemainInstallments As Integer, Optional ByRef purchaseprice As Double, Optional ByRef KhordaPrice As Double, Optional Depreciation_Percentage As Double, Optional ByRef Installmentvalue As Double)
    noOfInstallments = 10
    EXEInstallments = 4
    RemainInstallments = 6
    Installmentvalue = 700
End Function




Public Function Get_Account_name(Optional serial As String, _
                                 Optional Account_code As String) As String
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String
    sql = "Select * from ACCOUNTS where Account_Serial='" & serial & "'"

    If Account_code <> "" Then
        sql = "Select * from ACCOUNTS where Account_Code='" & Account_code & "'"
    End If
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then Get_Account_name = "": Exit Function
    If IsNull(Rs3("Account_Name").Value) And IsNull(Rs3("Account_NameEng").Value) Then Get_Account_name = "": Exit Function
  
    If SystemOptions.UserInterface = EnglishInterface Then
        If Not IsNull(Rs3("Account_NameEng").Value) Then Get_Account_name = Rs3("Account_NameEng").Value: Exit Function
    Else

        If Not IsNull(Rs3("Account_Name").Value) Then Get_Account_name = Rs3("Account_Name").Value: Exit Function
    End If
  
    Rs3.Close

End Function




Public Function GetAndCalculateAll(FixedassetId As Integer, _
                                   Optional DepreciationPercentag As Double, _
                                   Optional ByRef noOfInstallments As Integer, _
                                   Optional ByRef Age As Integer, _
                                   Optional purchaseprice As Double, _
                                   Optional KhordaPrice As Double, _
                                   Optional AccDepreciation As Double, _
                                   Optional ByRef currentvalue As Double, _
                                   Optional ByRef installValue As Double, _
                                   Optional ByRef EXEInstallments As Double, _
                                   Optional ByRef RemainInstallments As Double, _
                                   Optional MinusValue As Double)
    On Error Resume Next

    If DepreciationPercentag = 0 Then
        noOfInstallments = 0
        AccDepreciation = 0
        KhordaPrice = 0
        currentvalue = purchaseprice
        noOfInstallments = 0
        installValue = 0
        EXEInstallments = 0
        MinusValue = 0
        Exit Function
 
    End If
 
    noOfInstallments = 100 / DepreciationPercentag * 12
    Age = noOfInstallments
'    currentvalue = (purchaseprice - (AccDepreciation + KhordaPrice)) - MinusValue
currentvalue = (purchaseprice - (AccDepreciation + 0)) - MinusValue
    installValue = Round((purchaseprice - KhordaPrice) / noOfInstallments, 2)
    EXEInstallments = Round(AccDepreciation / installValue, 0)
    RemainInstallments = noOfInstallments - EXEInstallments

End Function






      Public Function AddNewDev(LngDevID As Variant, _
                          IntLineNO As Variant, _
                          StrAccountCode As String, _
                          SngValue As Variant, _
                          Credit_Or_Debit As Integer, _
                          Optional StrDes As String, _
                          Optional LngNoteID As Variant = 0, _
                          Optional LngReceiptID As Long = 0, _
                          Optional LngOperaID As Long = 0, _
                          Optional IntAccInterval As Long = 0, _
                          Optional RecordDate As Date, _
                          Optional LngUserID As Long = 0, _
                          Optional LngTransaction_ID As Long = 0, _
                          Optional StrDEV_Serial As String = "", _
                          Optional LngAdvancedID As Long = 0, _
                          Optional valuee As Variant = 0, _
                          Optional curr As String = "", _
                          Optional Rate As Long = 1, Optional ExpensesID As Double, Optional StrDese As String, Optional IntLineNO1 As Double, Optional notes_all As Double, Optional project_id As Integer, Optional opr_fullcode As String, _
                          Optional opening_balance As Boolean = False, Optional opening_balance_voucher_id As Double, Optional FixedassetId As Integer, Optional FixedAssetgroupid As Integer, Optional FixedAssetbranch_id As Integer, _
                          Optional branch_id As Integer = 1, Optional CarID As Double, Optional ShowQty1 As Double = 0, Optional showPrice1 As Double = 0, Optional showPrice2 As Double = 0, Optional Salaries1 As Double = 0, Optional Salaries2 As Double = 0, _
                          Optional Departementid As Double = 0, Optional NEmpid As Double, Optional ContNo As Integer, Optional Aqarid As Integer, Optional unittype As Integer, Optional unitno As Integer, Optional BillNo As String, Optional project_id1 As Integer, Optional pand_id As Integer, _
                          Optional oper_id As Integer, Optional Remarks2 As String, Optional hideline As Integer = 0, Optional ToTrans As Integer, Optional BankID As Integer = 0, Optional BoxID As Integer = 0, Optional StoreId As Integer = 0, Optional EmpID As Integer = 0, Optional CusID As Integer = 0, _
                          Optional Posted As Integer = 0, Optional FLgBranch As Integer, Optional OtherInformation As ClsGLOther) As Boolean

    Dim RsDev As ADODB.Recordset
    Dim RsSerial As ADODB.Recordset
    Dim StrSQL As String
    Dim LngSerialCount As Long
    Dim DblValue As Double
 
    'On Local Error GoTo ErrTrap
    
DblValue = Val(Format(SngValue, "." & String(Abs(SystemOptions.SysDefCurrencyForamt), "#")))
'DblValue = val(Format(SngValue, "." & String(Abs(SystemOptions.SysDefCurrencyForamt), "#")))

If DblValue = 0 Then

AddNewDev = True
Exit Function
End If
    If IsMissing(RecordDate) Then
        RecordDate = Date
    End If

    Set RsDev = New ADODB.Recordset

    If opening_balance = False Then
     '  RsDev.Open "DOUBLE_ENTREY_VOUCHERS", Cn, adOpenForwardOnly, adLockOptimistic, adCmdTable
            StrSQL = "SELECT     dbo.DOUBLE_ENTREY_VOUCHERS.* from dbo.DOUBLE_ENTREY_VOUCHERS Where (Double_Entry_Vouchers_ID = -1)"

          RsDev.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText

    Else
        check_opening_balance_notes
       ' RsDev.Open "DOUBLE_ENTREY_VOUCHERS1", Cn, adOpenForwardOnly, adLockOptimistic, adCmdTable
       StrSQL = "SELECT     dbo.DOUBLE_ENTREY_VOUCHERS1.* from dbo.DOUBLE_ENTREY_VOUCHERS1 Where (Double_Entry_Vouchers_ID = -1)"
       RsDev.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText

    End If

    RsDev.AddNew
    If Posted = 1 Then
    RsDev("Posted").Value = Posted
    Else
    RsDev("Posted").Value = Null
    End If
    If OtherInformation Is Nothing Then
    GoTo d
    End If
    If OtherInformation.FlgVat = 1 Then
    RsDev("FlgVat").Value = 1
    Else
    RsDev("FlgVat").Value = Null
    End If
    RsDev("CurrRow").Value = OtherInformation.CurrRow
    RsDev("Vatyo").Value = OtherInformation.Vatyo
    RsDev("Vat").Value = OtherInformation.Vat
    RsDev("TotalValue").Value = OtherInformation.TotalValue
    RsDev("DescAccount").Value = OtherInformation.DescAccount
    RsDev("AccountCode2").Value = OtherInformation.AccountCode2
    RsDev("SupplierID").Value = OtherInformation.SupplierID
    RsDev("CusVATNO").Value = OtherInformation.CusVATNO
    RsDev("SupplierName").Value = OtherInformation.SupplierName
    RsDev("PriceTotal").Value = OtherInformation.PriceTotal
    RsDev("Rate2").Value = OtherInformation.Rate
    RsDev("NextAccount_Code").Value = OtherInformation.NextAccount_Code
   ' RsDev("BillNo").value = OtherInformation.BillNo
d:
    If opening_balance = True Then
     RsDev("opening_balance_voucher_id").Value = opening_balance_voucher_id
     RsDev("ShowQty1").Value = ShowQty1
     RsDev("showPrice1").Value = showPrice1
     RsDev("showPrice2").Value = showPrice2
     RsDev("Salaries1").Value = Salaries1
     RsDev("Salaries2").Value = Salaries2
     RsDev("opening_balance_voucher_id").Value = opening_balance_voucher_id
    End If
    
 If hideline = 0 Then
 RsDev("hideline").Value = Null
 Else
RsDev("hideline").Value = hideline
End If

RsDev("ToTrans").Value = ToTrans

'If opening_balance = False Then
RsDev("Remarks2").Value = Remarks2
RsDev("projectid").Value = project_id1
RsDev("pandid").Value = pand_id
RsDev("operid").Value = oper_id
If FLgBranch = 1 Then
RsDev("FLgBranch").Value = FLgBranch
Else
  RsDev("FLgBranch").Value = Null
End If

'End If
    RsDev("branch_id").Value = branch_id
    RsDev("Double_Entry_Vouchers_ID").Value = LngDevID
    RsDev("DEV_ID_Line_No").Value = IntLineNO
    RsDev("DEV_ID_Line_No1").Value = IntLineNO1
    
    RsDev("Account_Code").Value = StrAccountCode
    DblValue = Val(Format(SngValue, "." & String(Abs(SystemOptions.SysDefCurrencyForamt), "#")))
    '    DBLValue = Round(SngValue, SystemOptions.SysDefCurrencyForamt)
    RsDev("Value").Value = DblValue
    RsDev("valuee").Value = valuee
       
        
    '   RsDev("Value").value = Round(RsDev("Value").value, SystemOptions.SysDefCurrencyForamt)
    '   RsDev("Value").value = Round(RsDev("Value").value, SystemOptions.SysDefCurrencyForamt)
      
    '     RsDev("ExpensesID").value = ExpensesID
     
    RsDev("currency").Value = curr
    RsDev("rate").Value = Rate
    RsDev("Credit_Or_Debit").Value = Credit_Or_Debit
    RsDev("Double_Entry_Vouchers_Description").Value = StrDes
    RsDev("Double_Entry_Vouchers_Descriptione").Value = StrDese
    
    If LngNoteID = 0 Then
        RsDev("Notes_ID").Value = Null
    Else
        RsDev("Notes_ID").Value = LngNoteID
    End If
    
    '  If Branch_Id = 0 Then
    '     rsdev("branch_id").value = Null
    ' Else
    '     rsdev("branch_id").value = LngNoteID
    ' End If
    
    If LngReceiptID = 0 Then
        RsDev("ReceiptID").Value = Null
    Else
        RsDev("ReceiptID").Value = LngReceiptID
    End If

    If LngOperaID = 0 Then
        RsDev("OperaID").Value = Null
    Else
        RsDev("OperaID").Value = LngOperaID
    End If

    If IntAccInterval = 0 Then
        RsDev("Account_Interval_ID").Value = SystemOptions.SysCurrentAccountIntervalID
    Else
        RsDev("Account_Interval_ID").Value = IntAccInterval
    End If

    RsDev("RecordDate").Value = RecordDate
    RsDev("RecordDateH").Value = ToHijriDate(RecordDate)
     
    If LngUserID = 0 Then
        RsDev("UserID") = user_id
    Else
        RsDev("UserID") = LngUserID
    End If

    If LngTransaction_ID = 0 Then
        RsDev("Transaction_ID").Value = Null
    Else
        RsDev("Transaction_ID").Value = LngTransaction_ID
    End If

    If LngAdvancedID = 0 Then
        RsDev("AdvanceID").Value = Null
    Else
        RsDev("AdvanceID").Value = LngAdvancedID
    End If

    If StrDEV_Serial <> "" Then
        RsDev("DEV_Serial").Value = StrDEV_Serial
    Else
        RsDev("DEV_Serial").Value = GetNewDEV_Serial(RecordDate)
    End If
    
    RsDev("notes_all").Value = Val(notes_all)
    RsDev("project_id").Value = project_id
    RsDev("opr_fullcode").Value = opr_fullcode
    
    RsDev("FixedAssetId").Value = FixedassetId
    RsDev("FixedAssetgroupid").Value = FixedAssetgroupid
    RsDev("FixedAssetbranch_id").Value = FixedAssetbranch_id
    RsDev("Departementid").Value = Departementid
    RsDev("NEmpid").Value = NEmpid
    If opening_balance = True Then
    RsDev("ContNo").Value = ContNo
End If
    If opening_balance = False Then
        RsDev("CarId").Value = CarID
      
        RsDev("Billno").Value = BillNo
          RsDev("Aqarid").Value = Aqarid
        RsDev("unittype").Value = unittype
        RsDev("unitno").Value = unitno
 
    End If

    If LngNoteID = 1 And opening_balance = True Then
        updateAutoOpeningBalanceVoucherValuebyCharacttex
    
    End If

      RsDev.Update

    AddNewDev = True
    RsDev.Close
    Set RsDev = Nothing
    Exit Function
ErrTrap:
    AddNewDev = False

    If RsDev.EditMode <> adEditNone Then
        RsDev.CancelUpdate
    End If

    RsDev.Close
    Set RsDev = Nothing
End Function


Public Function updateAutoOpeningBalanceVoucherValuebyCharacttex()
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim Total As Double
    sql = "SELECT     SUM([Value]) AS Total from dbo.DOUBLE_ENTREY_VOUCHERS1 WHERE  Credit_Or_Debit=0 and    (Notes_ID = 1)"
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount = 0 Then
 
    Else
        Total = IIf(IsNull(rs(Total).Value), 0, rs(Total).Value)
        sql = "update Notes1 set note_value_by_characters='" & WriteNo(Format(Total, "0.00"), 0, True, ".") & "' where NoteID=1"
        Cn.Execute sql

    End If

    rs.Close
    Set rs = Nothing

End Function


Public Function GET_DEFAULT_CURRENCY_INF(Optional ID As Integer = 0) As Boolean
    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    GET_DEFAULT_CURRENCY_INF = True
    
  If ID = 0 Then
    StrSQL = "Select * From currency Where basic=1"
  Else
  StrSQL = "Select * From currency Where id=" & ID
  End If
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText

    If rs.RecordCount > 0 Then
        DEFALUT_CURRENCY = IIf(IsNull(rs("NAME").Value), "", rs("NAME").Value)
        DEFALUT_CURRENCYE = IIf(IsNull(rs("NAMEE").Value), "", rs("NAMEE").Value)

        DEFALUT_CURRENCY_DIV = IIf(IsNull(rs("divname").Value), "", rs("divname").Value)
        DEFALUT_CURRENCY_DIVE = IIf(IsNull(rs("divnameE").Value), "", rs("divnameE").Value)
    Else

        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "·„ Ì „  ÕœÌœ ⁄„·… «· ⁄«„· «·«› —«÷Ì…"
        Else
            MsgBox "Please Define the default Currency"
        End If

        GET_DEFAULT_CURRENCY_INF = False
    End If

End Function


Function WriteNo(NO As String, Sex As Integer, Optional DecimalBracktes As Boolean = False, Optional DeciamlSymbol As String = "", Optional GroupingSymbol As String = "", Optional IntLang As Integer = 0, Optional DECIMAL_FOUND As Boolean = False _
, Optional currencycode As Integer) As String


    If GET_DEFAULT_CURRENCY_INF(currencycode) = False Then
        FRMcurrency.Show
        Exit Function
    End If

    If SystemOptions.UserInterface = EnglishInterface Or IntLang = 1 Then
        WriteNo = ConvertNumbersToWords(NO, currencycode)
        Exit Function
    End If

    Static FirstArray(9, 1) As String
    Static FirstArray1(2, 1)  As String
    Static SecondArray(9, 1) As String
    Static ThirdArray(9) As String

    ReDim Parts(4) As String
    ReDim PartStr(-1 To 3) As String

    Dim Length As Integer, i As Integer, TempLength As Integer
    Dim NoString As String, pos  As Integer
    Dim AfterPoint As String
    Dim Txt As String
    Dim StrSysDecSymbol As String
    Dim StrSysGroupSymbol As String
    Dim BolNegativeNumber As Boolean
    'sex=0 „–ﬂ—
    'sex= 1 „ƒ‰À

    'IntLang=0 Arabic
    'IntLang=1 English
    If SystemOptions.UserInterface = ArabicInterface Then
        FirstArray(1, 0) = "Ê«Õœ ": FirstArray(2, 0) = "«À‰«‰ ": FirstArray(3, 0) = "À·«À… "
        FirstArray(4, 0) = "√—»⁄… ": FirstArray(5, 0) = "Œ„”… ": FirstArray(6, 0) = "” … "
        FirstArray(7, 0) = "”»⁄… ": FirstArray(8, 0) = "À„«‰Ì… ": FirstArray(9, 0) = " ”⁄… "
    
        FirstArray(1, 1) = "Ê«Õœ… ": FirstArray(2, 1) = "«À‰ «‰ ": FirstArray(3, 1) = "À·«À "
        FirstArray(4, 1) = "√—»⁄ ": FirstArray(5, 1) = "Œ„” ": FirstArray(6, 1) = "”  "
        FirstArray(7, 1) = "”»⁄ ": FirstArray(8, 1) = "À„«‰ ": FirstArray(9, 1) = " ”⁄ "
                                                          
        FirstArray1(1, 0) = "√Õœ ": FirstArray1(2, 0) = "≈À‰« "
    
        FirstArray1(1, 1) = "≈ÕœÏ ": FirstArray1(2, 1) = "≈À‰ « "
                       
        SecondArray(1, 0) = "⁄‘—… ": SecondArray(2, 0) = "⁄‘—Ê‰ ": SecondArray(3, 0) = "À·«ÀÊ‰ "
        SecondArray(4, 0) = "√—»⁄Ê‰ ": SecondArray(5, 0) = "Œ„”Ê‰ ": SecondArray(6, 0) = "” Ê‰ "
        SecondArray(7, 0) = "”»⁄Ê‰ ": SecondArray(8, 0) = "À„«‰Ê‰ ": SecondArray(9, 0) = " ”⁄Ê‰ "
    
        SecondArray(1, 1) = "⁄‘—… ": SecondArray(2, 1) = "⁄‘—Ê‰ ": SecondArray(3, 1) = "À·«ÀÊ‰ "
        SecondArray(4, 1) = "√—»⁄Ê‰ ": SecondArray(5, 1) = "Œ„”Ê‰ ": SecondArray(6, 1) = "” Ê‰ "
        SecondArray(7, 1) = "”»⁄Ê‰ ": SecondArray(8, 1) = "À„«‰Ê‰ ": SecondArray(9, 1) = " ”⁄Ê‰ "
    
        ThirdArray(1) = "„«∆… ": ThirdArray(2) = "„«∆ «‰ ": ThirdArray(3) = "À·«À„«∆… "
        ThirdArray(4) = "√—»⁄„«∆… ": ThirdArray(5) = "Œ„”„«∆… ": ThirdArray(6) = "” „«∆… "
        ThirdArray(7) = "”»⁄„«∆… ": ThirdArray(8) = "À„«‰„«∆… ": ThirdArray(9) = " ”⁄„«∆… "
    ElseIf SystemOptions.UserInterface = EnglishInterface Then
        FirstArray(1, 0) = "One ": FirstArray(2, 0) = "Two ": FirstArray(3, 0) = "Three "
        FirstArray(4, 0) = "Four ": FirstArray(5, 0) = "Five ": FirstArray(6, 0) = "Six "
        FirstArray(7, 0) = "Seven ": FirstArray(8, 0) = "Eight ": FirstArray(9, 0) = "Nine "
    
        FirstArray(1, 1) = "One ": FirstArray(2, 1) = "Two ": FirstArray(3, 1) = "Three "
        FirstArray(4, 1) = "Four ": FirstArray(5, 1) = "Five ": FirstArray(6, 1) = "Six "
        FirstArray(7, 1) = "Seven ": FirstArray(8, 1) = "Eight ": FirstArray(9, 1) = "Nine "
                                                          
        FirstArray1(1, 0) = "√Õœ ": FirstArray1(2, 0) = "≈À‰« "
    
        FirstArray1(1, 1) = "≈ÕœÏ ": FirstArray1(2, 1) = "≈À‰ « "
                       
        SecondArray(1, 0) = "Ten ": SecondArray(2, 0) = "twenty ": SecondArray(3, 0) = "Thirty "
        SecondArray(4, 0) = "Forty ": SecondArray(5, 0) = "Fifty ": SecondArray(6, 0) = "Sixty "
        SecondArray(7, 0) = "Seventy ": SecondArray(8, 0) = "Eighty ": SecondArray(9, 0) = "Ninety "
    
        SecondArray(1, 1) = "Ten ": SecondArray(2, 1) = "Twenty ": SecondArray(3, 1) = "Thirty "
        SecondArray(4, 1) = "Forty ": SecondArray(5, 1) = "Fifty ": SecondArray(6, 1) = "Sixty "
        SecondArray(7, 1) = "Seventy ": SecondArray(8, 1) = "Eighty ": SecondArray(9, 1) = "Ninety "
    
        ThirdArray(1) = "One hundred": ThirdArray(2) = "two hundred ": ThirdArray(3) = "Three hundred "
        ThirdArray(4) = "Four hundred ": ThirdArray(5) = "Five hundred ": ThirdArray(6) = "Six hundred "
        ThirdArray(7) = "Seven hundred ": ThirdArray(8) = "Eight ": ThirdArray(9) = "Nine hundred "

    End If

    Txt = "": i = -1

    If Val(NO) = 0 Then 'Â· «·⁄œœ «·„œŒ· ’›—
        WriteNo = "’›—"
        Exit Function
    End If

    '«Õ–› «·›—«€«  «·Ì„Ì‰Ì… Ê«·Ì”«—Ì… «·“«∆œ… ›Ì Õ«· ÊÃÊœÂ«
    NoString = Trim(NO)

    '----------------------
    'ÌÃ»  ÕœÌœ Â· «·—ﬁ„ ”«·» «„ „ÊÃ»
    If Val(NO) > 0 Then
        BolNegativeNumber = False
    Else
        NO = Abs(Val(NO))
        BolNegativeNumber = True
    End If

    '----------------------
    'ÌÃ» „⁄—›… ‰Ê⁄ «·›«’·… «·⁄‘—Ì…
    '«·„” Œœ„… ›Ï «·—ﬁ„ «·„—”·
    If DeciamlSymbol = "" Then
        StrSysDecSymbol = GetDeciamlSymbol
    Else
        StrSysDecSymbol = DeciamlSymbol
    End If

    '----------------------
    '·Ê «‰ «·—ﬁ„ «·„—”· ≈·Ï «·œ«·… »Â ⁄·«„… ⁄‘—Ì…
    '»œ·« „‰ «·⁄·«ﬁ… «·√› —«÷Ì… «·„Œ’’… ›Ï «·ÃÂ«“
    '›Ï Â–Â «·Õ«·… ·«»œ „‰  »œÌ· «·⁄·«„… «·⁄‘—Ì… «·„—”·…
    '›Ï «·—ﬁ„ ‰›”Â »«·⁄·«„… «·⁄‘—Ì… «·√› —«÷Ì… Ê–·ﬂ
    pos = InStr(NoString, ".")

    If pos > 0 Then
        If StrSysDecSymbol <> "." Then
            NoString = Replace(NoString, ".", StrSysDecSymbol, , , vbBinaryCompare)
        End If
    End If

    '----------------------
    'ÌÃ» „⁄—›… ‰Ê⁄ ⁄·«„…  ÃÌ„⁄ «·¬·«›
    If GroupingSymbol = "" Then
        StrSysGroupSymbol = GetGroupingSymbol
    Else
        StrSysGroupSymbol = GroupingSymbol
    End If

    '----------------------
    pos = InStr(NoString, ",")

    If pos > 0 Then
        If StrSysGroupSymbol <> "," Then
            NoString = Replace(NoString, ",", "", , , vbBinaryCompare)
        End If
    End If

    '----------------------
    pos = InStr(NoString, StrSysGroupSymbol)

    If pos > 0 Then
        If StrSysDecSymbol <> "." Then
            NoString = Replace(NoString, StrSysGroupSymbol, "", , , vbBinaryCompare)
        End If
    End If

    '----------------------
'    If CheckTheSendNumber(NoString, StrSysDecSymbol, StrSysGroupSymbol) = False Then
'        If IntLang = 0 Then
'            WriteNo = "Œÿ√ ›Ï «·—ﬁ„ «·„—”· ...!!!"
'        Else
'            WriteNo = "Error in the Number...!!!"
'        End If
'
'        Exit Function
'    End If

    '----------------------
    ' «Õ’· ⁄·Ï ÿÊ· ”·”·… «·⁄œœ
    Length = Len(NoString)
    '«Õ›Ÿ „ﬂ«‰ ÊÃÊœ «·›«’·… «·⁄‘—Ì…
    pos = InStr(NoString, StrSysDecSymbol)

    '«ﬁ”„ ”·”·… «·⁄œœ≈·Ï „«ﬁ»· «·›«’·… Ê„«»⁄œ «·›«’·…
    If pos > 0 Then
        AfterPoint = Right$(NoString, Length - pos)
        NoString = Left$(NoString, pos - 1)
        Length = Len(NoString)
    Else
        pos = InStr(NoString, ",")

        If pos > 0 Then
            AfterPoint = Right$(NoString, Length - pos)
            NoString = Left$(NoString, pos - 1)
            Length = Len(NoString)
        End If
    End If

    'Ã“¡ «·⁄œœ ≈·Ï ”·«”· Õ—›Ì… „ƒ·›… „‰ À·«À Œ«‰«  ⁄‘—Ì… √Ê √ﬁ·
    TempLength = Length
    Parts(0) = NoString

    Do While TempLength >= 3
        TempLength = TempLength - 3
        i = i + 1
        Parts(i) = Right$(NoString, 3)
        NoString$ = Left$(NoString, TempLength)
    Loop

    Parts(i + 1) = NoString

    '«” œ⁄ «· «»⁄ «·›—⁄Ì Ê«Õ›Ÿ «·‰ «∆Ã ›Ì «·„’›Ê›…
    For i = 0 To 3

        If Len(Parts(i)) > 0 Then
            PartStr(i) = GetNo(Parts(i), Sex, i, FirstArray(), FirstArray1(), SecondArray(), ThirdArray(), IntLang)
        Else
            Exit For
        End If

    Next

    '««Ã„⁄ «·ﬂ·„«  «·Ã“∆Ì… «·‰« Ã… ›Ì ⁄«—… Ê«Õœ…
    For i = 3 To 0 Step -1

        If Len(PartStr(i)) > 0 Then
            If Len(PartStr(i - 1)) > 0 Then
                Txt = Txt & " " & PartStr(i) & IIf(IntLang = 0, "Ê", "and")
            Else
                Txt = Txt & " " & PartStr(i) & " "
            End If
        End If

    Next i

    If DECIMAL_FOUND = False Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Txt = "  ›ﬁÿ " & Txt & " " & DEFALUT_CURRENCY
        Else
            Txt = " ONLY " & Txt & " " & DEFALUT_CURRENCYE
        End If
    End If

    If Val(AfterPoint) > 0 Then
        Dim StrTemp As String
        StrTemp = GetAfterPoint(AfterPoint)
        DecimalBracktes = False

        If DecimalBracktes = True Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Txt = Txt & " Ê ) " & WriteNo(AfterPoint, Sex, , , , , True) & " " & StrTemp & " ("
            Else
                Txt = Txt & " and (" & WriteNo(AfterPoint, Sex, , , , , True) & " " & StrTemp & ")"
            End If

        Else

            If SystemOptions.UserInterface = ArabicInterface Then
                Txt = Txt & " Ê " & WriteNo(AfterPoint, Sex, , , , , True) & " " & StrTemp
            Else
                Txt = Txt & " and " & WriteNo(AfterPoint, Sex, , , , , True) & " " & StrTemp
            End If
        End If
    End If

    If BolNegativeNumber = True Then
        If IntLang = 0 Then
            Txt = "”«·» " & Txt
        Else
            Txt = "negative " & Txt
        End If
    End If

    If SystemOptions.UserInterface = ArabicInterface Then
        'WriteNo = " ›ﬁÿ " & Txt & "" & Get_currency_txt & "  ·«€Ì— "
        WriteNo = Txt & "  ·«€Ì— "

    Else
        'WriteNo = Txt & "" & Get_currency_txt
        WriteNo = Txt

    End If

    If DECIMAL_FOUND = True Then
        WriteNo = Txt
        Exit Function
    End If
 
End Function









Public Function check_opening_balance_notes()
    On Error Resume Next
    Dim departement_name As String
    departement_name = 1
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String
    Dim i As Integer

    sql = "select * from notes1  where NoteID=1"
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then
 
        Dim RsNetes As ADODB.Recordset
        Set RsNetes = New ADODB.Recordset
        RsNetes.Open "notes1", Cn, adOpenStatic, adLockOptimistic, adCmdTable

        RsNetes.AddNew
        RsNetes("NoteID").Value = 1
        RsNetes("NoteType").Value = 101
 
        RsNetes("NoteSerial").Value = Mid(Format$(Now, "dd/mm/yyyy"), 7, 4) & 1
        RsNetes("NoteSerial1").Value = Mid(Format$(Now, "dd/mm/yyyy"), 7, 4) & 1
    
        RsNetes("numbering_type").Value = sand_numbering_type(0) ' „”·”· «·ﬁÌœ
        RsNetes("numbering_type1").Value = sand_numbering_type(3) ' „”·”· «·”‰œ
    
        RsNetes("sanad_year").Value = year(Now)
        RsNetes("sanad_month").Value = Month(Now)
        RsNetes("foxy_no").Value = 1
        RsNetes("NoteDate").Value = Now
        RsNetes("Note_Value").Value = 0
        RsNetes("Double_Entry_Vouchers_ID").Value = 0
        RsNetes("DAWRY").Value = chkIsRepeatCode.Value
        RsNetes("KALEB").Value = Check3.Value
    
        RsNetes("Remark").Value = "Opening Balance"
    
        RsNetes.Update
 
    End If
  
End Function

Public Function sand_numbering_type(Sanad_No As Integer, Optional Current_branch As Long = 1) As Integer
    On Error Resume Next
    Dim departement_name As String
    departement_name = 1
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String
    Dim i As Integer

    sql = "select * from sanad_numbering where branch_no=" & Current_branch & " and  sanad_no=" & Sanad_No
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then numbering_id = 0: Exit Function
    sand_numbering_type = IIf(IsNull(Rs3("numbering_id").Value), 0, Rs3("numbering_id").Value): Exit Function
    Rs3.Close

End Function




Public Function GetNewDEV_Serial(RecordDate As Date) As String
    Dim StrSQL As String
    Dim RsSerial As ADODB.Recordset
    Dim LngSerialCount As Long
VBA.Calendar = vbCalGreg
    StrSQL = "Select Distinct Double_Entry_Vouchers_ID From " & " DOUBLE_ENTREY_VOUCHERS"
    StrSQL = StrSQL & " Where RecordDate=" & SQLDate(RecordDate, True) & ""
    Set RsSerial = New ADODB.Recordset
    RsSerial.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not RsSerial.BOF Or RsSerial.EOF Then
        LngSerialCount = 1
    Else
        LngSerialCount = RsSerial.RecordCount + 1
    End If

    GetNewDEV_Serial = year(RecordDate) & IIf(Len(Month(RecordDate)) = 1, "0" & Month(RecordDate), Month(RecordDate)) & IIf(Len(Day(RecordDate)) = 1, "0" & Day(RecordDate), Day(RecordDate)) & "0" & LngSerialCount
End Function




Public Function ConvertNumbersToWords(ByVal strInput As String, Optional currencycode As Integer) As String

    Dim strCleaned As String
    Dim nLoop As Integer
    Dim nLoop2 As Integer
    Dim strCents As String
    Dim strDollars As String
    Dim strConverted As String
    Dim strConvertedAll As String
    Dim strSection As String
    Dim strSectionValue As String
    Dim strSubSection As String
    Dim strNbrValue As String
    Dim strNbrValue2 As String

GET_DEFAULT_CURRENCY_INF currencycode
    ' Remove any characters not numeric or decimal
    For nLoop = 1 To Len(strInput)

        Select Case Mid$(strInput, nLoop, 1)

            Case "0" To "9", "."
                strCleaned = strCleaned & Mid$(strInput, nLoop, 1)
        End Select

    Next

    ' Check for cents
    nLoop = InStr(strCleaned, ".")

    ' Pad both with zeros
    If nLoop > 0 Then
        strCents = Right$("00" & Mid$(strCleaned, nLoop + 1), 3)
        strDollars = Right$(String$(12, "0") & Left$(strCleaned, nLoop - 1), 12)
    Else
        strDollars = Right$(String$(12, "0") & strCleaned, 12)
        strCents = "00"
    End If

    ' Put back together
    strCleaned = strDollars & "." & strCents

    ' Start making words
    For nLoop = 1 To Len(strCleaned)

        ' Which section of the number are we on?
        Select Case nLoop

            Case 1
                strConverted = vbNullString
                strSectionValue = vbNullString
                strSection = "Billion "

            Case 4

                If Trim$(strSectionValue) <> vbNullString Then
                    strConvertedAll = strConvertedAll & strSectionValue & " " & strSection
                End If

                strSectionValue = vbNullString
                strConverted = vbNullString
                strSection = "Million "

            Case 7

                If Trim$(strSectionValue) <> vbNullString Then
                    strConvertedAll = strConvertedAll & strSectionValue & " " & strSection
                End If

                strSectionValue = vbNullString
                strConverted = vbNullString
                strSection = "Thousand "

            Case 10

                If Trim$(strSectionValue) <> vbNullString Then
                    strConvertedAll = strConvertedAll & strSectionValue & " " & strSection
                End If

                strSectionValue = vbNullString
                strConverted = vbNullString
                strSection = vbNullString

            Case 14

                If Trim$(strSectionValue) <> vbNullString Then
                    strConvertedAll = strConvertedAll & strSectionValue & " " & strSection
                End If

                strSectionValue = vbNullString
                strConverted = vbNullString
                strSection = "Cents"
        End Select
  
        If Mid$(strCleaned, nLoop, 1) = "." Then
        Else

            For nLoop2 = 1 To 3

                ' Number value
                Select Case nLoop2

                    Case 1, 3

                        Select Case Val(Mid$(strCleaned, nLoop2 + nLoop - 1, 1))

                            Case 1: strNbrValue = "One"

                            Case 2: strNbrValue = "Two"

                            Case 3: strNbrValue = "Three"

                            Case 4: strNbrValue = "Four"

                            Case 5: strNbrValue = "Five"

                            Case 6: strNbrValue = "Six"

                            Case 7: strNbrValue = "Seven"

                            Case 8: strNbrValue = "Eight"

                            Case 9: strNbrValue = "Nine"

                            Case Else: strNbrValue = vbNullString
                        End Select
                             
                        Select Case nLoop2

                            Case 1

                                If strNbrValue <> vbNullString Then
                                    strSectionValue = strNbrValue & " Hundred"
                                End If

                            Case 3

                                If strNbrValue <> vbNullString Then
                                    If Right$(strSectionValue, 3) = "Ten" Then

                                        Select Case strNbrValue

                                            Case "One":                           strSectionValue = Left$(strSectionValue, Len(strSectionValue) - 3) & "Eleven"

                                            Case "Two":                           strSectionValue = Left$(strSectionValue, Len(strSectionValue) - 3) & "Twelve"

                                            Case "Three":                         strSectionValue = Left$(strSectionValue, Len(strSectionValue) - 3) & "Thirteen"

                                            Case "Four":                          strSectionValue = Left$(strSectionValue, Len(strSectionValue) - 3) & "Fourteen"

                                            Case "Five":                          strSectionValue = Left$(strSectionValue, Len(strSectionValue) - 3) & "Fifteen"

                                            Case "Six", "Seven", "Eight", "Nine": strSectionValue = Left$(strSectionValue, Len(strSectionValue) - 3) & strNbrValue & "teen"

                                            Case Else:                            strSectionValue = strSectionValue & " " & strNbrValue
                                        End Select

                                    Else
                                        strSectionValue = strSectionValue & " " & strNbrValue
                                    End If
                                End If

                        End Select

                    Case 2

                        Select Case Val(Mid$(strCleaned, nLoop2 + nLoop - 1, 1))

                            Case 1: strNbrValue2 = "Ten"

                            Case 2: strNbrValue2 = "Twenty"

                            Case 3: strNbrValue2 = "Thirty"

                            Case 4: strNbrValue2 = "Fourty"

                            Case 5: strNbrValue2 = "Fifty"

                            Case 6: strNbrValue2 = "Sixty"

                            Case 7: strNbrValue2 = "Seventy"

                            Case 8: strNbrValue2 = "Eighty"

                            Case 9: strNbrValue2 = "Ninety"

                            Case Else: strNbrValue2 = vbNullString
                        End Select

                        If strNbrValue2 <> vbNullString Then
                            strSectionValue = strSectionValue & " " & strNbrValue2
                        End If

                End Select

            Next

            nLoop = nLoop + 2
        End If

    Next

    ' Check for cents
    If strConvertedAll = "" Then strConvertedAll = "No "
    If Trim$(strSectionValue) = vbNullString Then
        If strConvertedAll = " One " Then
            strConvertedAll = strConvertedAll & DEFALUT_CURRENCYE
        Else
            strConvertedAll = strConvertedAll & DEFALUT_CURRENCYE
        End If

    Else

        If strConvertedAll = " One " Then
            strConvertedAll = strConvertedAll & DEFALUT_CURRENCYE & "And" & strSectionValue & " " & DEFALUT_CURRENCY_DIVE
        Else
            strConvertedAll = strConvertedAll & DEFALUT_CURRENCYE & " And" & strSectionValue & " " & DEFALUT_CURRENCY_DIVE
        End If
    End If

    ConvertNumbersToWords = strConvertedAll

End Function





Function GetNo(ns As String, Sex As Integer, Power As Integer, frst() As String, frst1() As String, scnd() As String, thrd() As String, Optional IntLang As Integer = 0) As String

    Dim Lngth As Integer, InvSex  As Integer
    ReDim Indx(3) As Integer
    ReDim TmpArray(2) As String
    Dim tms As String

    If Sex = 0 Then
        InvSex = 1
    Else
        InvSex = 0
    End If

    '«·Õ· „‰ √Ã· À·«À… √—ﬁ«„
    Lngth = Len(ns)
    '«·¬Õ«œ
    Indx(1) = Val(Mid$(ns, Lngth, 1))
    TmpArray(0) = frst(Indx(1), Sex)
    Lngth = Lngth - 1

    If Lngth > 0 Then
        '«·⁄‘—« 
        Indx(2) = Val(Mid$(ns, Lngth, 1))

        If TmpArray(0) <> "" Then
            TmpArray(1) = scnd(Indx(2), InvSex)
        Else
            TmpArray(1) = scnd(Indx(2), Sex)
        End If

        If (Indx(2) > 1) And (TmpArray(0) <> "") Then '«·⁄‘—«  „‰ 1 ≈·Ï  ”⁄…
            TmpArray(0) = TmpArray(0) & IIf(IntLang = 0, " Ê ", " and ")
        ElseIf (Indx(1) = 1) And (Indx(2) = 1) Then  '√Õœ ⁄‘—
            TmpArray(0) = frst1(1, Sex)
        ElseIf (Indx(1) = 2) And (Indx(2) = 1) Then ' «À‰« ⁄‘—
            TmpArray(0) = frst1(2, Sex)
        End If

        Lngth = Lngth - 1

        If Lngth > 0 Then
            '«·„∆« 
            Indx(3) = Val(Mid$(ns, Lngth, 1))
            TmpArray(2) = thrd(Indx(3))

            If (Indx(3) > 0) And ((TmpArray(0) <> "") Or (TmpArray(1) <> "")) Then
                TmpArray(2) = TmpArray(2) & IIf(IntLang = 0, " Ê ", " and ")
            End If

        Else
            GoTo last
        End If

    Else
        GoTo last
    End If

    '≈÷«›… ﬂ·„… «·„— »…(„∆…,√·›,...)Õ”» „— »… «·√—ﬁ«„
last:

    Select Case Power

        Case Is = -1
            tms = TmpArray(2) & TmpArray(0) & TmpArray(1)

            If (TmpArray(0) <> "") And (TmpArray(1) = "") And (TmpArray(2) = "") Then
                GetNo = tms & ""
            ElseIf (TmpArray(0) <> "") And (TmpArray(1) <> "") And (TmpArray(2) = "") Then
                GetNo = tms & ""
            ElseIf (TmpArray(0) <> "") And (TmpArray(1) <> "") And (TmpArray(2) <> "") Then
                GetNo = tms & ""
            End If

        Case Is = 0
            GetNo = TmpArray(2) & TmpArray(0) & TmpArray(1)

        Case Is = 1

            If (Indx(1) = 1) And (Indx(2) = 0) And (Indx(3) = 0) Then
                'GetNo = " √·› "
                GetNo = IIf(IntLang = 0, " √·› ", " Thousand ")
            ElseIf (Indx(1) = 2) And (Indx(2) = 0) And (Indx(3) = 0) Then
                'GetNo = " √·›«‰ "
                GetNo = IIf(IntLang = 0, " √·›«‰ ", " Two Thousand ")
            ElseIf (Indx(1) > 2) And (Indx(2) = 0) And (Indx(3) = 0) Then
                'GetNo = TmpArray(0) & " ¬·«› "
                GetNo = IIf(IntLang = 0, TmpArray(0) & " ¬·«› ", TmpArray(0) & " Thousands")
            ElseIf (Indx(1) = 0) And (Indx(2) = 1) And (Indx(3) = 0) Then

                'GetNo = TmpArray(2) & TmpArray(0) & TmpArray(1) & " ¬·«› "
                If IntLang = 0 Then
                    GetNo = TmpArray(2) & TmpArray(0) & TmpArray(1) & " ¬·«› "
                ElseIf IntLang = 1 Then
                    GetNo = TmpArray(2) & TmpArray(0) & TmpArray(1) & " Thousands "
                End If

            ElseIf (Indx(1) = 0) And (Indx(2) = 0) And (Indx(3) = 0) Then
                GetNo = TmpArray(2) & TmpArray(0) & TmpArray(1)
            Else
                'GetNo = TmpArray(2) & TmpArray(0) & TmpArray(1) & " √·› "
                GetNo = TmpArray(2) & TmpArray(0) & TmpArray(1) & IIf(IntLang = 0, " √·› ", " Thousand ")
            End If

        Case Is = 2

            If (Indx(1) = 1) And (Indx(2) = 0) And (Indx(3) = 0) Then
                'GetNo = " „·ÌÊ‰ "
                GetNo = IIf(IntLang = 0, " „·ÌÊ‰ ", " One Million ")
            ElseIf (Indx(1) = 2) And (Indx(2) = 0) And (Indx(3) = 0) Then
                'GetNo = " „·ÌÊ‰«‰ "
                GetNo = IIf(IntLang = 0, " „·ÌÊ‰«‰ ", " Two Millions ")
            ElseIf (Indx(1) > 2) And (Indx(2) = 0) And (Indx(3) = 0) Then
                'GetNo = TmpArray(0) & " „·«ÌÌ‰ "
                GetNo = TmpArray(0) & IIf(IntLang = 0, " „·«ÌÌ‰ ", " Millions ")
            ElseIf (Indx(1) = 0) And (Indx(2) = 1) And (Indx(3) = 0) Then
                GetNo = TmpArray(2) & TmpArray(0) & TmpArray(1) & IIf(IntLang = 0, " „·«ÌÌ‰ ", " Millions ")
            ElseIf (Indx(1) = 0) And (Indx(2) = 0) And (Indx(3) = 0) Then
                GetNo = TmpArray(2) & TmpArray(0) & TmpArray(1)
            Else
                'GetNo = TmpArray(2) & TmpArray(0) & TmpArray(1) & " „·ÌÊ‰ "
                GetNo = TmpArray(2) & TmpArray(0) & TmpArray(1) & IIf(IntLang = 0, " „·ÌÊ‰ ", " Millions ")
            End If

        Case Is = 3

            If (Indx(1) = 1) And (Indx(2) = 0) And (Indx(3) = 0) Then
                'GetNo = " „·Ì«— "
                GetNo = IIf(IntLang = 0, " „·Ì«— ", " Milliard ")
            ElseIf (Indx(1) = 2) And (Indx(2) = 0) And (Indx(3) = 0) Then
                'GetNo = " „·Ì«—«‰ "
                GetNo = IIf(IntLang = 0, " „·Ì«—«‰ ", " Two Milliard ")
            ElseIf (Indx(1) > 2) And (Indx(2) = 0) And (Indx(3) = 0) Then
                'GetNo = TmpArray(0) & " „·Ì«—«  "
                GetNo = TmpArray(0) & IIf(IntLang = 0, " „·Ì«—«  ", " Milliard ")
            ElseIf (Indx(1) = 0) And (Indx(2) = 1) And (Indx(3) = 0) Then
                GetNo = TmpArray(2) & TmpArray(0) & TmpArray(1) & IIf(IntLang = 0, " „·Ì«—«  ", " Milliard ")
            Else
                GetNo = TmpArray(2) & TmpArray(0) & TmpArray(1) & IIf(IntLang = 0, " „·Ì«— ", " Milliard ")
            End If

    End Select

End Function


Public Function GetAfterPoint(AfPont As String) As String
    Dim StrTemp As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrTemp = DEFALUT_CURRENCY_DIV
    Else
        StrTemp = DEFALUT_CURRENCY_DIVE
    End If

    GoTo ll

    If Len(AfPont) = 1 Then
        StrTemp = "„‰ ⁄‘—…"
    ElseIf Len(AfPont) = 2 Then
        StrTemp = "„‰ „«∆…"
    ElseIf Len(AfPont) = 3 Then
        StrTemp = "„‰ «·›"
    ElseIf Len(AfPont) = 4 Then
        StrTemp = "„‰ ⁄‘—… √·«›"
    ElseIf Len(AfPont) = 5 Then
        StrTemp = "„‰ „«∆… «·›"
    ElseIf Len(AfPont) = 6 Then
        StrTemp = "„‰ „·ÌÊ‰"
    ElseIf Len(AfPont) = 7 Then
        StrTemp = "„‰ ⁄‘—… „·«Ì‰"
    ElseIf Len(AfPont) = 8 Then
        StrTemp = "„‰ „«∆… „·ÌÊ‰"
    ElseIf Len(AfPont) = 9 Then
        StrTemp = "„‰ „·Ì«—"
    ElseIf Len(AfPont) = 10 Then
        StrTemp = "„‰ ⁄‘—… „·Ì«—"
    ElseIf Len(AfPont) = 11 Then
        StrTemp = "„‰ „«∆… „·Ì«—"
    ElseIf Len(AfPont) = 12 Then
        StrTemp = "„‰  —Ì·ÌÊ‰"
    ElseIf Len(AfPont) = 13 Then
        StrTemp = "„‰ ⁄‘—…  —Ì·ÌÊ‰"
    ElseIf Len(AfPont) = 14 Then
        StrTemp = "„‰ „«∆…  —Ì·ÌÊ‰"
    Else
        StrTemp = "€Ì— „Õœœ"
    End If

ll:
    GetAfterPoint = StrTemp
End Function

Private Function GetDeciamlSymbol() As String
    Dim i As Single
    Dim StrTemp As String
    i = 1 / 2
    StrTemp = FormatNumber(i, , vbUseDefault, vbUseDefault, vbTrue)
    GetDeciamlSymbol = Mid$(StrTemp, 2, 1)
End Function

Public Function GetGroupingSymbol() As String
    Dim i As Single
    Dim StrTemp As String
    i = 8819
    StrTemp = FormatNumber(i, , vbUseDefault, vbUseDefault, vbTrue)
    GetGroupingSymbol = Mid$(StrTemp, 2, 1)
End Function


 Public Function GetDefaultItemUnit(ItemID As Long, _
                                   Optional ByRef UnitID As Long, _
                                   Optional ByRef UnitName As String, Optional ByRef UnitFactor As Double)
    Dim RsUnitData As New ADODB.Recordset
    
    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = " SELECT TblItemsUnits.ItemID, TblItemsUnits.UnitID, TblUnites.UnitName,TblItemsUnits.UnitFactor, TblItemsUnits.SecOrder, TblItemsUnits.DefaultUnit," & "TblItemsUnits.UnitSalesPrice, TblItemsUnits.UnitPurPrice, TblItemsUnits.FactorByDefaultUnit," & "TblItemsUnits.FactorBySmallUnit "
    Else
        StrSQL = " SELECT TblItemsUnits.ItemID, TblItemsUnits.UnitID, TblUnites.UnitNamee UnitName," & "TblItemsUnits.UnitFactor, TblItemsUnits.SecOrder, TblItemsUnits.DefaultUnit," & "TblItemsUnits.UnitSalesPrice, TblItemsUnits.UnitPurPrice, TblItemsUnits.FactorByDefaultUnit," & "TblItemsUnits.FactorBySmallUnit "
    End If
    StrSQL = StrSQL + " FROM TblItemsUnits INNER JOIN TblUnites ON TblItemsUnits.UnitID =" & "TblUnites.UnitID"
    StrSQL = StrSQL + " Where TblItemsUnits.ItemID=" & ItemID
    StrSQL = StrSQL + " AND DefaultUnit=1"
    RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsUnitData.BOF Or RsUnitData.EOF) Then
        UnitID = IIf(IsNull(RsUnitData("UnitID").Value), 0, RsUnitData("UnitID").Value)
        UnitName = IIf(IsNull(RsUnitData("UnitName").Value), "", RsUnitData("UnitName").Value)
        UnitFactor = IIf(IsNull(RsUnitData("UnitFactor").Value), 0, RsUnitData("UnitFactor").Value)
    End If

    RsUnitData.Close
    Set RsUnitData = Nothing
         
End Function




Private Sub Text3_Change()
                Dim s As String
                s = "SELECT accounts.Account_Name,accounts.Account_Code,Account_Serial  FROM accounts  WHERE Account_Serial = N'" & Trim(Text3) & "'"
                Set rsDummy = New ADODB.Recordset
                rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly
                If Not rsDummy.EOF Then


                    DboParentAccount2.BoundText = Trim(rsDummy!Account_code & "")
                    
                End If
End Sub
