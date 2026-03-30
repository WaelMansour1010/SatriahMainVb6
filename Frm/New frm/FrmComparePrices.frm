VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form FrmComparePrices 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "„ﬁ«—‰Â ⁄—Ê÷ «·«”⁄«— "
   ClientHeight    =   9480
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   15405
   HelpContextID   =   580
   Icon            =   "FrmComparePrices.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   9480
   ScaleWidth      =   15405
   Visible         =   0   'False
   Begin C1SizerLibCtl.C1Elastic EleMain 
      Height          =   8865
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   15435
      _cx             =   27226
      _cy             =   15637
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
      AutoSizeChildren=   8
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
      GridRows        =   4
      GridCols        =   4
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"FrmComparePrices.frx":038A
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin C1SizerLibCtl.C1Tab C1Tab1 
         Height          =   7830
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   15375
         _cx             =   27120
         _cy             =   13811
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
         Caption         =   "»Ì«‰«  «”«”Ì…|»Ì«‰«   Õ·Ì·Ì…"
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
            Height          =   7410
            Left            =   16020
            TabIndex        =   83
            TabStop         =   0   'False
            Top             =   45
            Width           =   15285
            _cx             =   26961
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
            Begin VSFlex8Ctl.VSFlexGrid GridOvers 
               Height          =   1935
               Left            =   6960
               TabIndex        =   84
               Top             =   600
               Width           =   8265
               _cx             =   14579
               _cy             =   3413
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
               Cols            =   6
               FixedRows       =   1
               FixedCols       =   2
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmComparePrices.frx":0410
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
            Begin VSFlex8Ctl.VSFlexGrid GridItems 
               Height          =   1935
               Left            =   1800
               TabIndex        =   88
               Top             =   5280
               Width           =   6585
               _cx             =   11615
               _cy             =   3413
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
               Cols            =   4
               FixedRows       =   1
               FixedCols       =   2
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmComparePrices.frx":0511
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
            Begin VSFlex8Ctl.VSFlexGrid GridItemsGroup 
               Height          =   1935
               Left            =   8760
               TabIndex        =   90
               Top             =   5280
               Width           =   6465
               _cx             =   11404
               _cy             =   3413
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
               Cols            =   4
               FixedRows       =   1
               FixedCols       =   2
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmComparePrices.frx":05B0
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
            Begin VSFlex8Ctl.VSFlexGrid gridVendor 
               Height          =   1935
               Left            =   8760
               TabIndex        =   91
               Top             =   3000
               Width           =   6465
               _cx             =   11404
               _cy             =   3413
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
               Cols            =   4
               FixedRows       =   1
               FixedCols       =   2
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmComparePrices.frx":0657
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
            Begin VB.Label Label6 
               Alignment       =   1  'Right Justify
               Caption         =   "«·«’‰«› «·„Õœœ…"
               Height          =   375
               Left            =   6240
               RightToLeft     =   -1  'True
               TabIndex        =   89
               Top             =   5040
               Width           =   1935
            End
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               Caption         =   "«·„Ê—œÌ‰ «·„ÕœœÌ‰"
               Height          =   375
               Left            =   13200
               RightToLeft     =   -1  'True
               TabIndex        =   87
               Top             =   2640
               Width           =   1935
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               Caption         =   "„Ã„Ê⁄«  «·«’‰«› «·„Õœœ…"
               Height          =   375
               Left            =   13200
               RightToLeft     =   -1  'True
               TabIndex        =   86
               Top             =   5040
               Width           =   1935
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "«·⁄—Ê÷ ««·„Õœœ…"
               Height          =   375
               Left            =   12960
               RightToLeft     =   -1  'True
               TabIndex        =   85
               Top             =   240
               Width           =   1935
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   7410
            Index           =   2
            Left            =   45
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   45
            Width           =   15285
            _cx             =   26961
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
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   765
               Index           =   5
               Left            =   0
               TabIndex        =   33
               TabStop         =   0   'False
               Top             =   0
               Width           =   15315
               _cx             =   27014
               _cy             =   1349
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
               Picture         =   "FrmComparePrices.frx":06F6
               Caption         =   "„ﬁ«—‰Â ⁄—Ê÷ «·«”⁄«— "
               Align           =   0
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
                  TabIndex        =   34
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
                  ButtonImage     =   "FrmComparePrices.frx":13D0
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
                  TabIndex        =   35
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
                  ButtonImage     =   "FrmComparePrices.frx":176A
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
                  TabIndex        =   36
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
                  ButtonImage     =   "FrmComparePrices.frx":1B04
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
                  TabIndex        =   37
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
                  ButtonImage     =   "FrmComparePrices.frx":1E9E
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
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   7395
               Index           =   1
               Left            =   0
               TabIndex        =   5
               TabStop         =   0   'False
               Top             =   0
               Width           =   15345
               _cx             =   27067
               _cy             =   13044
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
                  Caption         =   "‘—Êÿ «·⁄—÷"
                  Height          =   1935
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   48
                  Top             =   1920
                  Width           =   15135
                  Begin VB.Frame Frame2 
                     Caption         =   "”Ì«”Â «›÷· ”⁄—"
                     Height          =   1215
                     Left            =   480
                     RightToLeft     =   -1  'True
                     TabIndex        =   77
                     Top             =   240
                     Width           =   3855
                     Begin VB.ListBox List2 
                        Height          =   840
                        ItemData        =   "FrmComparePrices.frx":2238
                        Left            =   600
                        List            =   "FrmComparePrices.frx":223A
                        RightToLeft     =   -1  'True
                        TabIndex        =   80
                        Top             =   240
                        Width           =   3135
                     End
                     Begin VB.CommandButton cmdUP 
                        Caption         =   "/\"
                        Height          =   255
                        Left            =   240
                        TabIndex        =   79
                        Top             =   360
                        Width           =   375
                     End
                     Begin VB.CommandButton CMDDown 
                        Caption         =   "\/"
                        Height          =   255
                        Left            =   240
                        TabIndex        =   78
                        Top             =   600
                        Width           =   375
                     End
                  End
                  Begin VB.CheckBox ChKOvers 
                     Alignment       =   1  'Right Justify
                     Caption         =   "·⁄—Ê÷ „Õœœ"
                     Height          =   195
                     Left            =   10920
                     RightToLeft     =   -1  'True
                     TabIndex        =   72
                     Top             =   240
                     Width           =   1455
                  End
                  Begin VB.Frame Frame6 
                     Height          =   1215
                     Left            =   9840
                     RightToLeft     =   -1  'True
                     TabIndex        =   70
                     Top             =   240
                     Width           =   2535
                     Begin VB.CommandButton CmdOvers 
                        Caption         =   "..."
                        BeginProperty Font 
                           Name            =   "MS Sans Serif"
                           Size            =   9.75
                           Charset         =   178
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        Height          =   240
                        Left            =   960
                        RightToLeft     =   -1  'True
                        TabIndex        =   73
                        Top             =   360
                        Width           =   495
                     End
                     Begin VB.Label lbl 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00E2E9E9&
                        BackStyle       =   0  'Transparent
                        Caption         =   "Õœœ «·⁄—Ê÷"
                        Height          =   195
                        Index           =   14
                        Left            =   1320
                        RightToLeft     =   -1  'True
                        TabIndex        =   71
                        Top             =   360
                        Width           =   1080
                     End
                  End
                  Begin VB.CheckBox Chkdates 
                     Alignment       =   1  'Right Justify
                     Caption         =   "· «—ÌŒ „Õœœ"
                     Height          =   195
                     Left            =   13320
                     RightToLeft     =   -1  'True
                     TabIndex        =   64
                     Top             =   240
                     Width           =   1695
                  End
                  Begin VB.Frame Frame4 
                     Caption         =   "«·„Ê—œÌ‰"
                     Height          =   1215
                     Left            =   4440
                     RightToLeft     =   -1  'True
                     TabIndex        =   59
                     Top             =   240
                     Width           =   2535
                     Begin VB.CommandButton CmdVendors 
                        Caption         =   "..."
                        Height          =   195
                        Left            =   120
                        RightToLeft     =   -1  'True
                        TabIndex        =   76
                        Top             =   480
                        Width           =   375
                     End
                     Begin VB.OptionButton OptVendor 
                        Alignment       =   1  'Right Justify
                        Caption         =   "Option3"
                        Height          =   255
                        Index           =   0
                        Left            =   2040
                        RightToLeft     =   -1  'True
                        TabIndex        =   61
                        Top             =   240
                        Value           =   -1  'True
                        Width           =   255
                     End
                     Begin VB.OptionButton OptVendor 
                        Alignment       =   1  'Right Justify
                        Caption         =   "Option3"
                        Height          =   255
                        Index           =   2
                        Left            =   2040
                        RightToLeft     =   -1  'True
                        TabIndex        =   60
                        Top             =   480
                        Width           =   255
                     End
                     Begin VB.Label lbl 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00E2E9E9&
                        BackStyle       =   0  'Transparent
                        Caption         =   "·ﬂ· «·„Ê—œÌ‰"
                        Height          =   315
                        Index           =   12
                        Left            =   360
                        RightToLeft     =   -1  'True
                        TabIndex        =   63
                        Top             =   240
                        Width           =   1560
                     End
                     Begin VB.Label lbl 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00E2E9E9&
                        BackStyle       =   0  'Transparent
                        Caption         =   "·„Ê—œÌ‰ „ÕœœÌ‰"
                        Height          =   315
                        Index           =   11
                        Left            =   480
                        RightToLeft     =   -1  'True
                        TabIndex        =   62
                        Top             =   480
                        Width           =   1440
                     End
                  End
                  Begin VB.Frame Frame3 
                     Caption         =   "«·«’‰«›"
                     Height          =   1215
                     Left            =   7080
                     RightToLeft     =   -1  'True
                     TabIndex        =   52
                     Top             =   240
                     Width           =   2535
                     Begin VB.CommandButton CmDItems 
                        Caption         =   "..."
                        Height          =   195
                        Left            =   240
                        RightToLeft     =   -1  'True
                        TabIndex        =   75
                        Top             =   720
                        Width           =   375
                     End
                     Begin VB.CommandButton CmdGroups 
                        Caption         =   "..."
                        Height          =   195
                        Left            =   240
                        RightToLeft     =   -1  'True
                        TabIndex        =   74
                        Top             =   480
                        Width           =   375
                     End
                     Begin VB.OptionButton OptItems 
                        Alignment       =   1  'Right Justify
                        Caption         =   "Option3"
                        Height          =   315
                        Index           =   2
                        Left            =   1920
                        RightToLeft     =   -1  'True
                        TabIndex        =   58
                        Top             =   720
                        Width           =   255
                     End
                     Begin VB.OptionButton OptItems 
                        Alignment       =   1  'Right Justify
                        Caption         =   "Option3"
                        Height          =   255
                        Index           =   1
                        Left            =   1920
                        RightToLeft     =   -1  'True
                        TabIndex        =   57
                        Top             =   480
                        Width           =   255
                     End
                     Begin VB.OptionButton OptItems 
                        Alignment       =   1  'Right Justify
                        Caption         =   "Option3"
                        Height          =   255
                        Index           =   0
                        Left            =   1920
                        RightToLeft     =   -1  'True
                        TabIndex        =   56
                        Top             =   240
                        Value           =   -1  'True
                        Width           =   255
                     End
                     Begin VB.Label lbl 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00E2E9E9&
                        BackStyle       =   0  'Transparent
                        Caption         =   "·«’‰«› „Õœœ…"
                        Height          =   315
                        Index           =   6
                        Left            =   360
                        RightToLeft     =   -1  'True
                        TabIndex        =   55
                        Top             =   720
                        Width           =   1560
                     End
                     Begin VB.Label lbl 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00E2E9E9&
                        BackStyle       =   0  'Transparent
                        Caption         =   "·ﬂ· «·«’‰«›"
                        Height          =   315
                        Index           =   4
                        Left            =   240
                        RightToLeft     =   -1  'True
                        TabIndex        =   54
                        Top             =   240
                        Width           =   1560
                     End
                     Begin VB.Label lbl 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00E2E9E9&
                        BackStyle       =   0  'Transparent
                        Caption         =   "·„Ã„Ê⁄«  „Õœœ…"
                        Height          =   315
                        Index           =   10
                        Left            =   360
                        RightToLeft     =   -1  'True
                        TabIndex        =   53
                        Top             =   480
                        Width           =   1560
                     End
                  End
                  Begin VB.Frame Frame5 
                     Height          =   1215
                     Left            =   12480
                     RightToLeft     =   -1  'True
                     TabIndex        =   65
                     Top             =   240
                     Width           =   2535
                     Begin MSComCtl2.DTPicker dbFromDate 
                        Height          =   315
                        Left            =   120
                        TabIndex        =   67
                        Top             =   240
                        Width           =   1335
                        _ExtentX        =   2355
                        _ExtentY        =   556
                        _Version        =   393216
                        Format          =   103677953
                        CurrentDate     =   38784
                     End
                     Begin MSComCtl2.DTPicker dbTodate 
                        Height          =   315
                        Left            =   120
                        TabIndex        =   69
                        Top             =   600
                        Width           =   1335
                        _ExtentX        =   2355
                        _ExtentY        =   556
                        _Version        =   393216
                        Format          =   103677953
                        CurrentDate     =   38784
                     End
                     Begin VB.Label lbl 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00E2E9E9&
                        BackStyle       =   0  'Transparent
                        Caption         =   "«·Ï  «—ÌŒ"
                        Height          =   195
                        Index           =   2
                        Left            =   1440
                        RightToLeft     =   -1  'True
                        TabIndex        =   68
                        Top             =   600
                        Width           =   960
                     End
                     Begin VB.Label lbl 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00E2E9E9&
                        BackStyle       =   0  'Transparent
                        Caption         =   "„‰  «—ÌŒ"
                        Height          =   315
                        Index           =   5
                        Left            =   1560
                        RightToLeft     =   -1  'True
                        TabIndex        =   66
                        Top             =   240
                        Width           =   840
                     End
                  End
                  Begin ALLButtonS.ALLButton cmdAdd 
                     Height          =   300
                     Left            =   480
                     TabIndex        =   82
                     Tag             =   "Delete Row"
                     Top             =   1560
                     Width           =   1260
                     _ExtentX        =   2223
                     _ExtentY        =   529
                     BTYPE           =   3
                     TX              =   " ‰›Ì– «·„ﬁ«—‰…"
                     ENAB            =   -1  'True
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
                     BCOL            =   65280
                     BCOLO           =   65280
                     FCOL            =   0
                     FCOLO           =   0
                     MCOL            =   12632256
                     MPTR            =   1
                     MICON           =   "FrmComparePrices.frx":223C
                     UMCOL           =   -1  'True
                     SOFT            =   0   'False
                     PICPOS          =   0
                     NGREY           =   0   'False
                     FX              =   0
                     HAND            =   0   'False
                     CHECK           =   0   'False
                     VALUE           =   0   'False
                  End
               End
               Begin VB.TextBox txtRemarks 
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
                  Height          =   615
                  Left            =   6480
                  MultiLine       =   -1  'True
                  RightToLeft     =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   47
                  Top             =   1320
                  Width           =   7680
               End
               Begin VB.CheckBox ChkLocked 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«Ìﬁ«› «· ⁄«„·"
                  Height          =   195
                  Left            =   14760
                  RightToLeft     =   -1  'True
                  TabIndex        =   41
                  Top             =   1680
                  Width           =   2310
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
                  Height          =   375
                  Left            =   16080
                  RightToLeft     =   -1  'True
                  TabIndex        =   40
                  Top             =   1980
                  Value           =   -1  'True
                  Width           =   1095
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
                  Height          =   375
                  Left            =   17160
                  RightToLeft     =   -1  'True
                  TabIndex        =   39
                  Top             =   1980
                  Width           =   1695
               End
               Begin VB.CheckBox ChKauto 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·Ì"
                  Enabled         =   0   'False
                  Height          =   255
                  Left            =   15600
                  RightToLeft     =   -1  'True
                  TabIndex        =   38
                  Top             =   3180
                  Width           =   1590
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
                  Height          =   315
                  Left            =   16320
                  RightToLeft     =   -1  'True
                  TabIndex        =   31
                  Text            =   "0"
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   495
               End
               Begin VB.TextBox TxtTblVendorContractD 
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
                  Height          =   255
                  Left            =   12960
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   30
                  Top             =   900
                  Width           =   1200
               End
               Begin VB.CheckBox Check1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "⁄—÷ "
                  Height          =   255
                  Left            =   15480
                  RightToLeft     =   -1  'True
                  TabIndex        =   20
                  Top             =   3300
                  Width           =   2310
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
                  Height          =   315
                  Index           =   0
                  Left            =   -3930
                  RightToLeft     =   -1  'True
                  TabIndex        =   10
                  Top             =   8580
                  Width           =   2175
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
                  Height          =   315
                  Left            =   5835
                  RightToLeft     =   -1  'True
                  TabIndex        =   6
                  Top             =   420
                  Visible         =   0   'False
                  Width           =   2160
               End
               Begin VSFlex8Ctl.VSFlexGrid Grid 
                  Height          =   3495
                  Left            =   17175
                  TabIndex        =   7
                  Top             =   3750
                  Width           =   9945
                  _cx             =   17542
                  _cy             =   6165
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
                  FormatString    =   $"FrmComparePrices.frx":2258
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
                  Left            =   16440
                  TabIndex        =   12
                  Top             =   1860
                  Width           =   4365
                  _ExtentX        =   7699
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
                  Left            =   16080
                  TabIndex        =   13
                  Top             =   1440
                  Width           =   1605
                  _ExtentX        =   2831
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
                  Left            =   16680
                  TabIndex        =   28
                  Top             =   780
                  Width           =   3285
                  _ExtentX        =   5794
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
                  Height          =   390
                  Index           =   20
                  Left            =   16680
                  TabIndex        =   43
                  Top             =   1980
                  Width           =   720
                  _ExtentX        =   1270
                  _ExtentY        =   688
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
                  ButtonImage     =   "FrmComparePrices.frx":2521
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   390
                  Index           =   21
                  Left            =   15960
                  TabIndex        =   44
                  Top             =   1980
                  Width           =   690
                  _ExtentX        =   1217
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
                  ButtonImage     =   "FrmComparePrices.frx":28BB
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin MSDataListLib.DataCombo DBCboClientName 
                  Height          =   315
                  Left            =   6480
                  TabIndex        =   45
                  Top             =   900
                  Width           =   2805
                  _ExtentX        =   4948
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
               Begin MSDataListLib.DataCombo dcitems 
                  Height          =   315
                  Left            =   16920
                  TabIndex        =   46
                  Top             =   1980
                  Width           =   4365
                  _ExtentX        =   7699
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
               Begin MSComCtl2.DTPicker DTPicker1 
                  Height          =   255
                  Left            =   10560
                  TabIndex        =   49
                  Top             =   900
                  Width           =   1335
                  _ExtentX        =   2355
                  _ExtentY        =   450
                  _Version        =   393216
                  Format          =   103677953
                  CurrentDate     =   38784
               End
               Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid1 
                  Height          =   3135
                  Left            =   0
                  TabIndex        =   51
                  Top             =   3840
                  Width           =   15225
                  _cx             =   26855
                  _cy             =   5530
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
                  Cols            =   16
                  FixedRows       =   1
                  FixedCols       =   2
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmComparePrices.frx":2E55
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
                  BackColor       =   &H00C0FFFF&
                  Caption         =   " ﬁÊ„ Â–… «·‘«‘… »„ﬁ«—‰… ⁄—Ê÷ «·«”⁄«— ÿ»ﬁ« ··‘—Êÿ Ê «·„⁄«ÌÌ—— «· Ì  ÕœœÂ« "
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
                  Height          =   540
                  Index           =   38
                  Left            =   480
                  RightToLeft     =   -1  'True
                  TabIndex        =   81
                  Top             =   1080
                  Width           =   5775
               End
               Begin VB.Shape Shape1 
                  BorderWidth     =   2
                  FillColor       =   &H00C0FFFF&
                  FillStyle       =   0  'Solid
                  Height          =   975
                  Left            =   120
                  Top             =   840
                  Width           =   6255
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«· «—ÌŒ"
                  Height          =   255
                  Index           =   9
                  Left            =   11880
                  RightToLeft     =   -1  'True
                  TabIndex        =   50
                  Top             =   900
                  Width           =   840
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„·«ÕŸ« "
                  Height          =   315
                  Index           =   3
                  Left            =   14280
                  RightToLeft     =   -1  'True
                  TabIndex        =   42
                  Top             =   1320
                  Width           =   840
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬁ«∆„ »«·⁄„·Ì…"
                  Height          =   255
                  Index           =   0
                  Left            =   8925
                  RightToLeft     =   -1  'True
                  TabIndex        =   29
                  Top             =   900
                  Width           =   1440
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "»œ«Ì… «· Œ’Ì’"
                  Height          =   270
                  Index           =   8
                  Left            =   15360
                  RightToLeft     =   -1  'True
                  TabIndex        =   11
                  Top             =   2400
                  Width           =   1785
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "—ﬁ„ «·⁄„·»…"
                  Height          =   210
                  Index           =   7
                  Left            =   13980
                  RightToLeft     =   -1  'True
                  TabIndex        =   9
                  Top             =   900
                  Width           =   1185
               End
               Begin VB.Label Label5 
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
                  Height          =   315
                  Left            =   13800
                  RightToLeft     =   -1  'True
                  TabIndex        =   8
                  Top             =   900
                  Width           =   855
               End
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·„ÊŸ›"
               Height          =   315
               Index           =   1
               Left            =   8400
               RightToLeft     =   -1  'True
               TabIndex        =   3
               Top             =   90
               Width           =   1125
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic EltCont 
         Height          =   960
         Left            =   30
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   7875
         Width           =   15375
         _cx             =   27120
         _cy             =   1693
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
         Begin ImpulseButton.ISButton btnQuery 
            Height          =   330
            Left            =   11880
            TabIndex        =   15
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
            ButtonImage     =   "FrmComparePrices.frx":30C3
            ColorButton     =   14737632
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnUpdate 
            Height          =   330
            Left            =   12765
            TabIndex        =   16
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
            ButtonImage     =   "FrmComparePrices.frx":345D
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton BtnPrint 
            Height          =   285
            Left            =   13965
            TabIndex        =   17
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
            ButtonImage     =   "FrmComparePrices.frx":37F7
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   495
            Index           =   0
            Left            =   11220
            TabIndex        =   21
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
            Left            =   10320
            TabIndex        =   22
            Top             =   510
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
            Left            =   9480
            TabIndex        =   23
            Top             =   510
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
            Left            =   8475
            TabIndex        =   24
            Top             =   510
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
            Left            =   7440
            TabIndex        =   25
            Top             =   510
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
            Left            =   4560
            TabIndex        =   26
            Top             =   510
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
            Left            =   6510
            TabIndex        =   27
            Top             =   510
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
            Left            =   14040
            TabIndex        =   32
            Tag             =   "Delete Row"
            Top             =   120
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
            MICON           =   "FrmComparePrices.frx":3B91
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
            Height          =   495
            Index           =   7
            Left            =   5520
            TabIndex        =   92
            Top             =   480
            Width           =   765
            _ExtentX        =   1349
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
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   4210752
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   4210752
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
            TabIndex        =   19
            Top             =   225
            Width           =   1740
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
            TabIndex        =   18
            Top             =   240
            Width           =   1515
         End
      End
   End
   Begin ImpulseButton.ISButton ISButton1 
      Height          =   345
      Left            =   3360
      TabIndex        =   4
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
      ButtonImage     =   "FrmComparePrices.frx":3BAD
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
End
Attribute VB_Name = "FrmComparePrices"
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
Public StrOrder   As String
Public StrItemID As String
Public StrGroupID As String
Public StrCusID   As String
Public policy   As String

Private Declare Function TextOut _
                Lib "gdi32" _
                Alias "TextOutA" (ByVal hDC As Long, _
                                  ByVal X As Long, _
                                  ByVal Y As Long, _
                                  ByVal lpString As String, _
                                  ByVal nCount As Long) As Long

Function print_report(Optional NoteSerial As String)
    
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String



MySQL = " SELECT     dbo.TblComparPrice.ID, dbo.TblComparPrice.EmpID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Name1, dbo.TblEmployee.Emp_Name2, "
MySQL = MySQL & "                      dbo.TblEmployee.Emp_Name3, dbo.TblEmployee.Emp_Name4, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee4, dbo.TblEmployee.Emp_Namee3,"
MySQL = MySQL & "                       dbo.TblEmployee.Emp_Namee2, dbo.TblEmployee.Emp_Namee1, dbo.TblEmployee.Emp_Namee, dbo.TblComparPrice.RecordDate, dbo.TblComparPrice.FromDate,"
MySQL = MySQL & "                       dbo.TblComparPrice.Todate, dbo.TblComparPrice.Remarks, dbo.TblComparPrice.Chkdates, dbo.TblComparPrice.ChKOvers, dbo.TblComparPrice.ChItem,"
MySQL = MySQL & "                       dbo.TblComparPrice.ChCustomer, dbo.TblComparPrice.StrOrder, dbo.TblComparPrice.StrItemID, dbo.TblComparPrice.StrGroupID, dbo.TblComparPrice.StrCusID,"
MySQL = MySQL & "                       dbo.TblComparPrice.policy, dbo.TblComparPriceDet.Transaction_ID, dbo.TblComparPriceDet.GroupID, dbo.Groups.GroupName, dbo.Groups.GroupNamee,"
MySQL = MySQL & "                       dbo.TblComparPriceDet.ItemId, dbo.TblItems.ItemName, dbo.TblItems.ItemNamee, dbo.TblComparPriceDet.NoteSerial1, dbo.TblComparPriceDet.Quantity,"
MySQL = MySQL & "                       dbo.TblComparPriceDet.ShowPrice, dbo.TblComparPriceDet.PODays, dbo.TblComparPriceDet.Transaction_Date, dbo.TblComparPriceDet.CusID,"
MySQL = MySQL & "                       dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode AS CusFullcode, dbo.TblComparPriceDet.CountryID, dbo.Nationality.name,"
MySQL = MySQL & "                       dbo.nationality.NameE , dbo.nationality.Quality"
MySQL = MySQL & "  FROM         dbo.Nationality RIGHT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblComparPriceDet ON dbo.Nationality.id = dbo.TblComparPriceDet.CountryID LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblCustemers ON dbo.TblComparPriceDet.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblItems ON dbo.TblComparPriceDet.ItemId = dbo.TblItems.ItemID LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.Groups ON dbo.TblComparPriceDet.GroupID = dbo.Groups.GroupID RIGHT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblComparPrice ON dbo.TblComparPriceDet.CoPriceID = dbo.TblComparPrice.ID LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.TblEmployee ON dbo.TblComparPrice.EmpID = dbo.TblEmployee.Emp_ID"
MySQL = MySQL & "  Where (dbo.TblComparPrice.id =" & val(TxtTblVendorContractD.Text) & ")"

 


  If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepComparQutaion.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepComparQutaionE.rpt"
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
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
        ' xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
        StrReportTitle = "" '& StrAccountName
        'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        '    StrReportTitle = StrReportTitle + " »œ«Ì… „‰ " & Format(Me.DTPickerAccFrom.value, "yyyy/M/d") & ""
        'End If
        'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        '    StrReportTitle = StrReportTitle + " ≈·Ï " & Format(Me.DTPickerAccTo.value, "yyyy/M/d") & " "
        'End If
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        'xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""
        'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        '    StrReportTitle = StrReportTitle + " From Date " & (Me.DTPickerAccFrom.value) & ""
        'End If
        'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        '    StrReportTitle = StrReportTitle + " To Date :  " & (Me.DTPickerAccTo.value) & ""
        'End If
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
      '  xReport.ParameterFields(4).AddCurrentValue WriteNo(Format(val(TxtAdvanceValue.text), "0.00"), 0, True, ".")
       ' xReport.ParameterFields(6).AddCurrentValue val(lbl(23).Caption)
        ' xReport.ParameterFields(7).AddCurrentValue DBIssueDate.value
    'xReport.ParameterFields(8).AddCurrentValue IIf(IsNumeric(Fg.TextMatrix(Me.Fg.FixedRows, Fg.ColIndex("PartValue"))), val(Fg.TextMatrix(Me.Fg.FixedRows, Fg.ColIndex("PartValue"))), 0)
' xReport.ParameterFields(9).AddCurrentValue val(lbl(22).Caption)
 ' xReport.ParameterFields(10).AddCurrentValue val(TxtDiscount.text)
  ' xReport.ParameterFields(11).AddCurrentValue txtDiscountDES.text
   
'    xReport.ParameterFields(5).AddCurrentValue ToHijriDate(RsData("notedate").value)
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
Sub FillListPolicy()
If SystemOptions.UserInterface = ArabicInterface Then
   List2.AddItem "1 ÿ»ﬁ« ··”⁄—"
   List2.AddItem "2 ÿ»ﬁ« · «—ÌŒ «· ”·Ì„"
   List2.AddItem "3 ÿ»ﬁ« ··œ›⁄ "
   List2.AddItem "4 ÿ»ﬁ« ·»·œ «·„‰‘√"
   Else
   List2.AddItem "1 According to price"
   List2.AddItem "2 According to the date of delivery"
   List2.AddItem "3 According to the payment "
   List2.AddItem "4 According to the country of origin "
  End If
End Sub


Sub retrivepolicy(Optional policy As String)
If policy = "" Then
Exit Sub
End If
 Dim astrSplit2tems2() As String
   List2.Clear
Dim j As Integer
For j = 0 To 3
           astrSplit2tems2 = Split(policy, "#")
      If SystemOptions.UserInterface = ArabicInterface Then
         Select Case astrSplit2tems2(j)
         Case "0"
  List2.AddItem "1 ÿ»ﬁ« ··”⁄—"
 Case "1"
   List2.AddItem "2 ÿ»ﬁ« · «—ÌŒ «· ”·Ì„"
   Case "2"
  List2.AddItem "3 ÿ»ﬁ« ··œ›⁄ "
   Case "3"
  List2.AddItem "4 ÿ»ﬁ« ·»·œ «·„‰‘√"
 
  End Select
  Else
       Select Case astrSplit2tems2(j)
         Case "0"
  List2.AddItem "1 According to price"
 Case "1"
   List2.AddItem "2 According to the date of delivery"
   Case "2"
  List2.AddItem "3 According to the payment "
    Case "3"
  List2.AddItem "4 According to the country of origin "
 
  End Select
  End If
      
      Next j
End Sub
'Private Sub ChkDetails_Click()
'    FillGridWithData
'End Sub

'Private Sub ALLButton1_Click()
'    FrmShowCol1.show
'End Sub

'Function check_previous_dev(year As String, Month As String) As Boolean
'    Dim rs As ADODB.Recordset
'    Set rs = New ADODB.Recordset
''    Dim sql As String
 '   sql = "Select * from notes where salary=" & year & Month
 '
 '   rs.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
 '
 '   If rs.RecordCount = 0 Then
 '       check_previous_dev = False
 '   Else
 '       check_previous_dev = True
 '   End If
 
'End Function

'Function check_previous_dev1(year As String, Month As String) As Boolean
'    Dim rs As ADODB.Recordset
''    Set rs = New ADODB.Recordset
 '   Dim sql As String
 '   sql = "Select * from salary_voucher where m_year='" & year & "' and m_month='" & Month & "'"
 '
 '   rs.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
 '
 '   If rs.RecordCount = 0 Then
 '       check_previous_dev1 = False
 '   Else
 '       check_previous_dev1 = True
 '   End If
 '
'End Function

'Function Create_dev()
'    Dim i As Integer
'    Dim LngDevID As Long
'    Dim Msg As String
'    Dim Account_Code_dynamic As String
'    Dim Account_Code_dynamic1 As String
'
'    Dim Employee_account As String
'    Dim StrAccountCode As String
'    Dim X As Integer
'    Dim rs As ADODB.Recordset
'    Dim notes_serial As String
'    Dim notes_id As String
''
 '   Account_Code_dynamic = get_account_code_branch(16, my_branch)
'
'    If Account_Code_dynamic = "NO branch" Then
'        MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
'        GoTo ErrTrap
'    Else
'
'        If Account_Code_dynamic = "NO account" Then
'            MsgBox "·„ Ì „  ÕœÌœ Õ”«»  ··«ÃÊ—   ··„ÊŸ›Ì‰ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
'            GoTo ErrTrap
'
'        End If
'    End If
'
'    Msg = "ﬁÌœ «” Õﬁ«ﬁ —Ê« » «·„ÊŸ›Ì‰ ⁄‰ ‘Â— " & "   ”‰… "
'
'    Dim StrSQL As String
'    Set rs = New ADODB.Recordset
'    StrSQL = "select * From Notes where NoteType=66 order by NoteID"
'
'    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    notes_id = CStr(new_id("Notes", "NoteID", "", True))
'    notes_serial = CStr(new_id("Notes", "NoteSerial", "", True, "NoteType=66"))
'
'    rs.AddNew
'    rs("NoteID").value = notes_id
'    rs("NoteSerial").value = notes_serial '
'    rs("Note_Value").value = Null
'    rs("Remark").value = Msg
'
'    rs("NoteType").value = 66
'    rs("NoteDate").value = Date
'    rs("UserID").value = user_id
'    rs.update
'
'    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
'
'    Dim line_no As Integer
'    line_no = 1
'
'    With Grid
'
'        For i = .FixedRows To .Rows - 2
'
'            If .TextMatrix(i, .ColIndex("project")) = "0" Then
'
'                If ModAccounts.AddNewDev(LngDevID, line_no, Account_Code_dynamic, .TextMatrix(i, .ColIndex("EmpTotalNet")), 0, Msg, val(notes_id), , , , Date, user_id) = False Then
'                    GoTo ErrTrap
'                End If
'
'            Else
'                Account_Code_dynamic1 = get_project_Account(.TextMatrix(i, .ColIndex("project")), "Salary_account")
'
'                If ModAccounts.AddNewDev(LngDevID, line_no, Account_Code_dynamic1, .TextMatrix(i, .ColIndex("EmpTotalNet")), 0, Msg, val(notes_id), , , , Date, user_id) = False Then
'                    GoTo ErrTrap
''                End If
 '           End If
 '
 '           Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code1")
 '           StrAccountCode = Employee_account
 '
 '           If ModAccounts.AddNewDev(LngDevID, line_no + 1, StrAccountCode, .TextMatrix(i, .ColIndex("EmpTotalNet")), 1, Msg, val(notes_id), , , , Date, user_id) = False Then
 '               GoTo ErrTrap
 '           End If
 '
 '           line_no = line_no + 2
 '
 '       Next i
'
'    End With
'
'    MsgBox " „ «‰‘«¡ «·ﬁÌœ", vbInformation
'    create_report_data
'
'    DoEvents
'
'    Exit Function
'ErrTrap:
'    MsgBox "ÕœÀ Œÿ√ «À‰«¡ Õ›Ÿ «·»Ì«‰« ", vbExclamation
'
'End Function
'
'Function Create_dev1()
'    Dim i As Integer
'    Dim LngDevID As Long
'    Dim Msg As String
'    Dim Account_Code_dynamic As String
'    Dim Account_Code_dynamic1 As String
'
'    Dim Employee_account As String
'    Dim StrAccountCode As String
'    Dim X As Integer
'    Dim rs As ADODB.Recordset
'
'    Account_Code_dynamic = get_account_code_branch(16, my_branch)
'
'    If Account_Code_dynamic = "NO branch" Then
'        MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
'        GoTo ErrTrap
'    Else
'
'        If Account_Code_dynamic = "NO account" Then
''            MsgBox "·„ Ì „  ÕœÌœ Õ”«»  ··«ÃÊ—   ··„ÊŸ›Ì‰ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
 '           GoTo ErrTrap
 '
 '       End If
 '   End If
 '
 '   'StrAccountCode = Account_Code_dynamic
 '
 '   LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
 '
 '   Dim line_no As Integer
 '   line_no = 1
'
'    With Grid
'
'        For i = .FixedRows To .Rows - 2
'
'            If .TextMatrix(i, .ColIndex("project")) = "0" Then
'
'                If ModAccounts.AddNewDev(LngDevID, line_no, Account_Code_dynamic, .TextMatrix(i, .ColIndex("EmpTotalNet")), 0, Msg, , , , , Date, user_id) = False Then
'                    GoTo ErrTrap
'                End If
'
'            Else
'                Account_Code_dynamic1 = get_project_Account(.TextMatrix(i, .ColIndex("project")), "Salary_account")
'
'                If ModAccounts.AddNewDev(LngDevID, line_no, Account_Code_dynamic1, .TextMatrix(i, .ColIndex("EmpTotalNet")), 0, Msg, , , , , Date, user_id) = False Then
'                    GoTo ErrTrap
'                End If
'            End If
'
'            Employee_account = get_EMPLOYEE_Account(val(.TextMatrix(i, .ColIndex("Emp_ID"))), "Account_Code1")
'            StrAccountCode = Employee_account
'
'            If ModAccounts.AddNewDev(LngDevID, line_no + 1, StrAccountCode, .TextMatrix(i, .ColIndex("EmpTotalNet")), 1, Msg, , , , , Date, user_id) = False Then
'                GoTo ErrTrap
'            End If
'
'            line_no = line_no + 2
'
'        Next i
'
'    End With

'    Set rs = New ADODB.Recordset
'    rs.Open "salary_voucher", Cn, adOpenStatic, adLockOptimistic, adCmdTable
'    rs.AddNew
'
'    rs("voucher_id").value = LngDevID
'
'    rs.update
'
'    MsgBox " „ «‰‘«¡ «·ﬁÌœ", vbInformation
'    create_report_data
'
'    DoEvents
'
'    Exit Function
'ErrTrap:
'    MsgBox "ÕœÀ Œÿ√ «À‰«¡ Õ›Ÿ «·»Ì«‰« ", vbExclamation
  
'End Function

'Private Sub ALLButton2_Click()
'    'Dcemp.text = ""

'    dcproject.text = ""
'    FillGridWithData
'
'    DoEvents
'    Create_dev
'    CMDOK_Click
'End Sub



'Private Sub CboPayMentType_Click()
'    CboPayMentType_Change
'End Sub

'Private Sub CboYear_Click()
'    CMDOK_Click
'End Sub



'Private Sub Check1_Click()
'
'    If Check1.value = vbChecked Then
'        get_all_employee
''    Else
'
'        With Me.Grid
'            .Rows = 2
'            .Clear flexClearScrollable
'        End With
'
'    End If

'End Sub

'Private Sub CmbMonth_Click()
'    CMDOK_Click
'    'FillGridWithData
'End Sub

'Private Sub CmdExit_Click()
'    Unload Me
'End Sub



'Private Sub CmdPrint_Click()
'    On Error Resume Next
'    Dim GrdBack As ClsBackGroundPic
    'Grid.ExtendLastCol = True
'    Grid.WallPaper = Nothing
    'Grid.AutoSize  0, Grid.Cols - 1, False
'    Printer.Orientation = VBRUN.PrinterObjectConstants.vbPRORLandscape
 
    'Printer.RightToLeft = True
    'Printer.Print ("Employee Salary Report")

'    Me.Grid.PrintGrid " ﬁ—Ì— —Ê« » «·„ÊŸ›Ì‰", True, 2, 1, 1500

    'Me.Grid.PrintGrid , True, 2, 0, 2

    'Grid.ExtendLastCol = False
    'Grid.AutoSize 0, Grid.Cols - 1, False
    'Set GrdBack = New ClsBackGroundPic
    'Set Grid.WallPaper = GrdBack.Picture
    'Grid.ExtendLastCol = True
'End Sub



Private Sub SaveData()
Dim i As Integer
    Dim Msg As String
    Dim RsTemp As New ADODB.Recordset
    Dim StrSQL As String
    Dim BeginTrans As Boolean
    Dim RsDev As ADODB.Recordset
    Dim LngDevID As Long
 
 '   On Error GoTo ErrTrap

    If Me.TxtModFlg.Text <> "R" Then
 
        If Trim(Me.DBCboClientName.BoundText) = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "ÌÃ» ≈Œ Ì«— «·ﬁ«∆„ »«·⁄„·ÌÂ..!!"
            Else
             Msg = "Please Select Based Process"
            
            End If
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            DBCboClientName.SetFocus
            SendKeys "{F4}"
            Exit Sub
        End If
 
    End If

    '-------------------------------------------------------------------------------------------
  If TxtModFlg.Text = "N" Then
   TxtTblVendorContractD.Text = CStr(new_id("TblComparPrice", "ID", "", True))
   End If

  Cn.BeginTrans
    BeginTrans = True

    If TxtModFlg.Text = "N" Then
   
        rs.AddNew
        rs("ID").value = val(TxtTblVendorContractD.Text)
        
    ElseIf Me.TxtModFlg.Text = "E" Then
        Cn.Execute "delete TblComparPriceDet where CoPriceID=" & val(Me.TxtTblVendorContractD.Text)
   
    End If
     
    rs("ID").value = val(TxtTblVendorContractD.Text)
    rs("EmpID").value = IIf(Me.DBCboClientName.BoundText = "", Null, Me.DBCboClientName.BoundText)
    rs("RecordDate").value = DTPicker1.value
    rs("FromDate").value = dbFromDate.value
    rs("Todate").value = dbTodate.value
    rs("Remarks").value = IIf(Me.txtRemarks.Text = "", "", Me.txtRemarks.Text)
    rs("StrOrder").value = StrOrder
    rs("StrItemID").value = StrItemID
    rs("StrGroupID").value = StrGroupID
    rs("StrCusID").value = StrCusID


 Dim policy As String
   For i = 0 To List2.ListCount - 1
  Select Case List2.List(i)
  Case "1 ÿ»ﬁ« ··”⁄—", "1 According to price"
  policy = policy & "0" & "#"
   Case "2 ÿ»ﬁ« · «—ÌŒ «· ”·Ì„", "2 According to the date of delivery"
   policy = policy & "1" & "#"
  Case "3 ÿ»ﬁ« ··œ›⁄ ", "3 According to the payment "
   policy = policy & "2" & "#"
    Case "4 ÿ»ﬁ« ·»·œ «·„‰‘√", "4 According to the country of origin "
   policy = policy & "3" & "#"
  End Select
    Next i
     rs("policy").value = policy
 ''/
    If Chkdates.value = vbChecked Then
        rs("Chkdates").value = 1
    Else
        rs("Chkdates").value = 0
    End If
    If ChKOvers.value = vbChecked Then
        rs("ChKOvers").value = 1
    Else
        rs("ChKOvers").value = 0
    End If
   If OptItems(0).value = True Then
        rs("ChItem").value = 1
    ElseIf OptItems(1).value Then
        rs("ChItem").value = 2
       ElseIf OptItems(2).value Then
        rs("ChItem").value = 3
       Else
       rs("ChItem").value = 0
    End If
       If OptVendor(0).value = True Then
        rs("ChCustomer").value = 1
    ElseIf OptVendor(2).value Then
        rs("ChCustomer").value = 2
      Else
      rs("ChCustomer").value = 0
      End If
    rs.Update
   ' CuurentLogdata
     Set RsDev = New ADODB.Recordset
       StrSQL = "SELECT     *  from dbo.TblComparPriceDet Where (1 = -1)"
   RsDev.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
 
  
    With Me.VSFlexGrid1

        For i = 1 To .Rows - 1

            If .TextMatrix(i, .ColIndex("ItemId")) <> "" Then
         
                RsDev.AddNew
                RsDev("CoPriceID").value = Me.TxtTblVendorContractD.Text
                RsDev("CountryID").value = val(.TextMatrix(i, .ColIndex("CountryID")))
                RsDev("ItemId").value = val(.TextMatrix(i, .ColIndex("ItemId")))
                RsDev("Transaction_ID").value = val(.TextMatrix(i, .ColIndex("Transaction_ID")))
                RsDev("GroupID").value = val(.TextMatrix(i, .ColIndex("GroupID")))
                RsDev("CusID").value = val(.TextMatrix(i, .ColIndex("CusID")))
                RsDev("PODays").value = val(.TextMatrix(i, .ColIndex("PODays")))
                RsDev("ShowPrice").value = val(.TextMatrix(i, .ColIndex("ShowPrice")))
                RsDev("Quantity").value = val(.TextMatrix(i, .ColIndex("Quantity")))
                RsDev("NoteSerial1").value = .TextMatrix(i, .ColIndex("NoteSerial1"))
                RsDev("Transaction_Date").value = IIf(.TextMatrix(i, .ColIndex("Transaction_Date")) = "", Null, .TextMatrix(i, .ColIndex("Transaction_Date")))
                RsDev.Update
                    
            End If
            
            '
        Next i

   '     AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTextE, Me.name, Me.TxtModFlg
            
    End With
 
    Cn.CommitTrans
    BeginTrans = False
 
    Select Case Me.TxtModFlg.Text

        Case "N"
   If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì…" & Chr(13)
                    Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
                Else
                    Msg = " Saved Successfully" & Chr(13)
                    Msg = Msg + "do you new Operation?"
        
                End If
            '    Fg_Journal.Enabled = False
            If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                Cmd_Click (0)
                Exit Sub
            End If

        Case "E"
          If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Else
                    MsgBox "Saved Changes Successfully", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                End If
    End Select

    TxtModFlg.Text = "R"
    'End If

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



Private Sub Cmd_Click(Index As Integer)
    On Error GoTo ErrTrap

    Select Case Index

        Case 0
     
            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.Text = "N"
            clear_all Me
            'Me.TxtTblVendorContractD.text = CStr(new_id("TblComparPrice", "ID", "", True))
       
            Me.dbFromDate.value = Date
            Me.dbTodate.value = Date
       List2.Clear
       Me.FillListPolicy
            'XPDtbTrans.SetFocus
            VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
            VSFlexGrid1.Rows = 1
            VSFlexGrid1.Enabled = True
            Option2.value = True
       OptItems(0).value = True
        OptVendor(0).value = True
        
StrOrder = ""
StrItemID = ""
StrGroupID = ""
 StrCusID = ""
  policy = ""
        Case 1

            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If

            TxtModFlg.Text = "E"
            '         Grid.Rows = Grid.Rows + 1
            Grid.Enabled = True
         
           ' CuurentLogdata

        Case 2
    
            SaveData
           
        Case 3
            Undo

        Case 4

            If DoPremis(Do_Delete, Me.Name, True) = False Then
                Exit Sub
            End If

            Del_Trans

        Case 5

            If DoPremis(Do_Search, Me.Name, True) = False Then
                Exit Sub
            End If
            Unload FrmSearchComparQuotation
           Load FrmSearchComparQuotation
            
            FrmSearchComparQuotation.show vbModal

        Case 6
            Unload Me
           Case 7

            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

            If val(Me.TxtTblVendorContractD.Text) <> 0 Then
                print_report val(Me.TxtTblVendorContractD.Text)
        
        
            End If

        Case 21
            RemoveGridRow
    End Select

    Exit Sub
ErrTrap:

End Sub

Private Sub Del_Trans()
    On Error GoTo ErrTrap
    Dim Msg  As String

    If TxtTblVendorContractD.Text <> "" Then
        Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & Chr(13)
        Msg = Msg + (TxtTblVendorContractD.Text) & Chr(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            If Not rs.RecordCount < 1 Then
             Cn.Execute "delete TblComparPrice where ID=" & val(Me.TxtTblVendorContractD.Text)
                Cn.Execute "delete TblComparPriceDet where CoPriceID=" & val(Me.TxtTblVendorContractD.Text)
               ' CuurentLogdata ("D")
                rs.delete
                rs.MoveFirst

                If rs.RecordCount < 1 Then
                    clear_all Me
                    VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
            VSFlexGrid1.Rows = 1
            VSFlexGrid1.Enabled = True
                    TxtModFlg_Change
                    '   XPTxtCurrent.Caption = 0
                    '   XPTxtCount.Caption = 0
                Else
                    Retrive
                End If
            End If
        End If

    Else
        clear_all Me
        Msg = "Â–Â «·⁄„·Ì… €Ì— „ «Õ… ÕÌÀ √‰Â ·«ÌÊÃœ √Ì ”Ã·« "
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        TxtModFlg_Change
        Exit Sub
    End If

    TxtModFlg_Change
    Exit Sub
ErrTrap:

    If Err.Number = -2147217887 Then
        Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & Chr(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «·„Ê—œ "
        MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
        rs.CancelUpdate
    End If

End Sub

Private Sub RemoveGridRow()

    With Me.VSFlexGrid1

        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With

    ReLineGrid
End Sub

'Function addrow()
'
'    Dim wherestr As String
'
'    Dim Rs3 As ADODB.Recordset
'    Set Rs3 = New ADODB.Recordset
'    Dim rs2 As ADODB.Recordset
'    Set rs2 = New ADODB.Recordset
'    Dim RsUnit As ADODB.Recordset
'    Set RsUnit = New ADODB.Recordset
'
'    Dim j As Integer
'
'    Dim sql As String
''    Dim i As Integer
 '   Dim Msg  As String
 '   Dim lastrow As Integer
 '   Dim LngItemID As Integer
'
'    If Option2.value = True Then
'        If dcitems.BoundText = "" Then
'            If SystemOptions.UserInterface = ArabicInterface Then
'                Msg = "ÌÃ»       «Œ Ì«— «·’‰›  ...!!!"
'            Else
'                Msg = "must Specify item Name ...!!!"
''            End If
'
'            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
'            Exit Function
'        End If
'
'        wherestr = "  where ItemID= " & val(dcitems.BoundText)
'    End If
'
'    sql = "Select * from TblItems "
'
'    If wherestr <> "" Then
'        sql = sql & wherestr
'    End If
'
'    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
'
'    If Rs3.RecordCount = 0 Then Exit Function
'
'    With Grid
'
'        lastrow = .Rows
'
'        If Rs3.RecordCount > 0 Then
'            .Rows = Rs3.RecordCount + lastrow
'            Rs3.MoveFirst
'
'            For i = lastrow To Rs3.RecordCount + lastrow - 1
'                .TextMatrix(i, .ColIndex("ItemId")) = IIf(IsNull(Rs3.Fields("ItemId").value), "", Rs3.Fields("ItemId").value)
'                LngItemID = IIf(IsNull(Rs3.Fields("ItemId").value), "", Rs3.Fields("ItemId").value)
'
'                .TextMatrix(i, .ColIndex("ItemCode")) = IIf(IsNull(Rs3.Fields("ItemCode").value), "", Rs3.Fields("ItemCode").value)
'                .TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(Rs3.Fields("ItemName").value), "", Rs3.Fields("ItemName").value)
'
'                'lllllllllllllll
'                StrSQL = "SELECT TblItemsUnits.UnitID, TblUnites.UnitName "
'                StrSQL = StrSQL + " FROM TblUnites INNER JOIN TblItemsUnits " & "ON TblUnites.UnitID = TblItemsUnits.UnitID "
'                StrSQL = StrSQL + " Where TblItemsUnits.DefaultUnit=1 and  TblItemsUnits.ItemID=" & LngItemID
'                StrSQL = StrSQL + " Order BY TblItemsUnits.SecOrder "
'
'                RsUnit.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
'
'                If RsUnit.RecordCount > 0 Then
'                    RsUnit.MoveFirst
'                    .TextMatrix(i, .ColIndex("UnitId")) = IIf(IsNull(RsUnit.Fields("UnitId").value), "", RsUnit.Fields("UnitId").value)
'                    .TextMatrix(i, .ColIndex("UnitName")) = IIf(IsNull(RsUnit.Fields("UnitName").value), "", RsUnit.Fields("UnitName").value)
'
'                End If
'
'                RsUnit.Close
'
'                Rs3.MoveNext
'            Next i
'
'            '    .AutoSize 0, .Cols - 1, False
'        End If
'
'    End With
 
'    Rs3.Close

'    ReLineGrid

'End Function

Private Sub Undo()
    On Error GoTo ErrTrap

    Select Case TxtModFlg.Text

        Case "N"
            clear_all Me
            Me.TxtModFlg.Text = "R"
            XPBtnMove_Click (1)

        Case "E"
 
            Retrive
            Me.TxtModFlg.Text = "R"
    End Select

    Exit Sub
ErrTrap:
End Sub

'Private Sub Dcdep_Click(Area As Integer)
'    CMDOK_Click
'End Sub

'Private Sub Dcedara_Click(Area As Integer)
'    CMDOK_Click
'End Sub
'
'Private Sub Dcemp_Click(Area As Integer)
'    CMDOK_Click
'End Sub

'Private Sub DCmboEmp_Click(Area As Integer)
'    FillGridWithData
'End Sub

'Function SHow_grig_col()
'    Dim rs2 As ADODB.Recordset
'    Set rs2 = New ADODB.Recordset
'    rs2.Open "Employee_salary_col", Cn, adOpenStatic, adLockOptimistic, adCmdTable
'
'    With Grid
'
'        If rs2("s1").value = True Then
'            .ColHidden(.ColIndex("Emp_Code")) = False
'        Else
'            .ColHidden(.ColIndex("Emp_Code")) = True
'        End If
'
'        If rs2("s2").value = True Then
'            .ColHidden(.ColIndex("Emp_Name")) = False
'        Else
'            .ColHidden(.ColIndex("Emp_Name")) = True
'        End If
'
'        If rs2("s3").value = True Then
'            .ColHidden(.ColIndex("Emp_Salary")) = False
'        Else
'            .ColHidden(.ColIndex("Emp_Salary")) = True
'        End If
'
'        If rs2("s4").value = True Then
'            .ColHidden(.ColIndex("Emp_Salary_sakn")) = False
'        Else
'            .ColHidden(.ColIndex("Emp_Salary_sakn")) = True
'        End If
'
'        If rs2("s5").value = True Then
'            .ColHidden(.ColIndex("Emp_Salary_bus")) = False
'        Else
'            .ColHidden(.ColIndex("Emp_Salary_bus")) = True
'        End If
'
'        If rs2("s6").value = True Then
'            .ColHidden(.ColIndex("Emp_Salary_food")) = False
'        Else
'            .ColHidden(.ColIndex("Emp_Salary_food")) = True
'        End If
'
'        If rs2("s7").value = True Then
'            .ColHidden(.ColIndex("Emp_Salary_mob")) = False
'        Else
'            .ColHidden(.ColIndex("Emp_Salary_mob")) = True
'        End If
'
'        If rs2("s8").value = True Then
'            .ColHidden(.ColIndex("Emp_Salary_mang")) = False
'        Else
'            .ColHidden(.ColIndex("Emp_Salary_mang")) = True
'        End If
'
'        If rs2("s9").value = True Then
'            .ColHidden(.ColIndex("Emp_Salary_others")) = False
'        Else
'            .ColHidden(.ColIndex("Emp_Salary_others")) = True
'        End If
'
'        If rs2("s10").value = True Then
'            .ColHidden(.ColIndex("OverTimePrice")) = False
'        Else
'            .ColHidden(.ColIndex("OverTimePrice")) = True
'        End If
'
'        If rs2("s11").value = True Then
'            .ColHidden(.ColIndex("Mokafea")) = False
'        Else
'            .ColHidden(.ColIndex("Mokafea")) = True
'        End If
'
'        If rs2("s12").value = True Then
'            .ColHidden(.ColIndex("SalesCom")) = False
'        Else
'            .ColHidden(.ColIndex("SalesCom")) = True
'        End If
'
'        If rs2("s13").value = True Then
'            .ColHidden(.ColIndex("total1")) = False
'        Else
'            .ColHidden(.ColIndex("total1")) = True
'        End If
'
'        If rs2("s14").value = True Then
'            .ColHidden(.ColIndex("TotalAdvance")) = False
'        Else
'            .ColHidden(.ColIndex("TotalAdvance")) = True
'        End If
'
'        If rs2("s15").value = True Then
'            .ColHidden(.ColIndex("TotalDiscount")) = False
'        Else
'            .ColHidden(.ColIndex("TotalDiscount")) = True
'        End If
'
'        If rs2("s16").value = True Then
 '           .ColHidden(.ColIndex("total2")) = False
'        Else
'            .ColHidden(.ColIndex("total2")) = True
'        End If
'
'        If rs2("s17").value = True Then
'            .ColHidden(.ColIndex("EmpTotalNet")) = False
'        Else
'            .ColHidden(.ColIndex("EmpTotalNet")) = True
'        End If
'
'        If rs2("s18").value = True Then
'            .ColHidden(.ColIndex("sgn")) = False
'        Else
'            .ColHidden(.ColIndex("sgn")) = True
'        End If
'
'    End With
'
'End Function
'
'

Sub GetDatat()
 Dim StrWhere As String
 Dim Rs1 As ADODB.Recordset
 Dim i As Integer
Dim sql As String
Dim Msg As String
Dim Order As String
Order = ""
sql = "SELECT     TOP 100 PERCENT dbo.Transactions.Transaction_ID, dbo.Transactions.Transaction_Date, dbo.Transactions.NoteSerial1, dbo.TblCustemers.CusName, "
sql = sql & "                      dbo.TblCustemers.CusNamee, dbo.Transactions.PODays, dbo.Transactions.PaymentType, dbo.Transactions.Transaction_HijriDate, dbo.Transaction_Details.Item_ID,"
sql = sql & "                      dbo.TblItems.code, dbo.TblItems.Fullcode, dbo.TblItems.ItemCode, dbo.TblItems.ItemName, dbo.TblItems.ItemNamee, dbo.Transaction_Details.Quantity,"
sql = sql & "                      dbo.Transaction_Details.Price, dbo.Transaction_Details.ShowQty, dbo.Transaction_Details.showPrice, dbo.Transaction_Details.ItemCase,"
sql = sql & "                      dbo.Transaction_Details.ItemSerial, dbo.Transaction_Details.ItemDiscountType, dbo.Transaction_Details.ItemDiscount, dbo.Transaction_Details.guaranteeTime,"
sql = sql & "                      dbo.Transaction_Details.CostPrice, dbo.Transaction_Details.CostTransID, dbo.Transaction_Details.ItemProfit, dbo.Transaction_Details.Remarks,"
sql = sql & "                      dbo.Transaction_Details.ItemSize, dbo.Transaction_Details.QtyBySmalltUnit, dbo.Transaction_Details.Remarks1, dbo.Transaction_Details.Remarks2,"
sql = sql & "                      dbo.Transaction_Details.sallReturnPrice, dbo.Transaction_Details.ToTAlELSHahn, dbo.TblItems.GroupID, dbo.Groups.GroupName, dbo.Groups.GroupCode,"
sql = sql & "                      dbo.Groups.Fullcode AS GroupFullcode, dbo.Groups.GroupNamee, dbo.Transaction_Details.UnitId, dbo.TblUnites.UnitName, dbo.TblUnites.UnitNamee,"
sql = sql & "                      dbo.Transaction_Details.[Catalog], dbo.Transaction_Details.ItemBalance, dbo.Transaction_Details.ShipedQty, dbo.Transactions.CusID, dbo.Transactions.countryid,"
sql = sql & "                      dbo.nationality.Quality , dbo.nationality.NameE, dbo.nationality.name"
sql = sql & " FROM         dbo.Nationality RIGHT OUTER JOIN"
sql = sql & "                      dbo.Transactions ON dbo.Nationality.id = dbo.Transactions.countryid LEFT OUTER JOIN"
sql = sql & "                      dbo.TblUnites RIGHT OUTER JOIN"
sql = sql & "                      dbo.Transaction_Details ON dbo.TblUnites.UnitID = dbo.Transaction_Details.UnitId LEFT OUTER JOIN"
sql = sql & "                      dbo.TblItems LEFT OUTER JOIN"
sql = sql & "                      dbo.Groups ON dbo.TblItems.GroupID = dbo.Groups.GroupID ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID ON"
sql = sql & "                      dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID LEFT OUTER JOIN"
sql = sql & "                      dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID"
sql = sql & "  Where (dbo.Transactions.Transaction_Type = 46)"
If Chkdates.value = vbChecked Then
  If Not IsNull(Me.dbFromDate.value) Then
            StrWhere = StrWhere & " AND  dbo.Transactions.Transaction_Date >=" & SQLDate(Me.dbFromDate.value, True) & ""
      End If
     If Not IsNull(Me.dbTodate.value) Then
            StrWhere = StrWhere & " AND  dbo.Transactions.Transaction_Date <=" & SQLDate(Me.dbTodate.value, True) & ""
      End If
      
 End If
 If ChKOvers.value = vbChecked And StrOrder <> "" Then
            StrWhere = StrWhere & " AND dbo.Transactions.Transaction_ID in( " & StrOrder & " ) "
      End If
  If OptItems(0).value = True Then
          StrWhere = StrWhere & " AND dbo.Transaction_Details.Item_ID <> -1"
  End If
   If OptItems(1).value = True And StrGroupID <> "" Then
          StrWhere = StrWhere & " AND dbo.TblItems.GroupID in( " & StrGroupID & " ) "
  End If
  
     If OptItems(2).value = True And StrItemID <> "" Then
          StrWhere = StrWhere & " AND dbo.Transaction_Details.Item_ID in( " & StrItemID & " ) "
  End If
       If OptVendor(0).value = True Then
          StrWhere = StrWhere & " AND dbo.Transactions.CusID <>-1"
  End If
       If OptVendor(2).value = True And StrCusID <> "" Then
          StrWhere = StrWhere & " AND dbo.Transactions.CusID in( " & StrCusID & " ) "
  End If

  Order = " order by "
  For i = 0 To List2.ListCount - 1
  Select Case List2.List(i)
  Case "1 ÿ»ﬁ« ··”⁄—", "1 According to price"
  Order = Order & "dbo.Transaction_Details.showPrice" & ","
   Case "2 ÿ»ﬁ« · «—ÌŒ «· ”·Ì„", "2 According to the date of delivery"
    Order = Order & "dbo.Transactions.PODays" & ","
  Case "3 ÿ»ﬁ« ··œ›⁄ ", "3 According to the payment "
  Order = Order & " dbo.Transactions.PaymentType" & ","
   Case "4 ÿ»ﬁ« ·»·œ «·„‰‘√", "4 According to the country of origin "
   Order = Order & " dbo.nationality.Quality" & ","
  End Select
    Next i
      Order = Order & "dbo.Transactions.Transaction_ID "
   
    '-----------------------------------

    sql = sql & StrWhere
    sql = sql & " " & Order
    Set Rs1 = New ADODB.Recordset
    Rs1.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Rs1.BOF Or Rs1.EOF Then
          Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «·»ÕÀ"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    Else
            With VSFlexGrid1
            .Clear flexClearScrollable, flexClearEverything
            .Rows = .FixedRows
            .Rows = Rs1.RecordCount + .FixedRows
            Rs1.MoveFirst
        
            For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .ColIndex("Ser")) = i
                .TextMatrix(i, .ColIndex("CountryID")) = IIf(IsNull(Rs1("CountryID").value), "", Rs1("CountryID").value)
                .TextMatrix(i, .ColIndex("CusID")) = IIf(IsNull(Rs1("CusID").value), "", Rs1("CusID").value)
                 .TextMatrix(i, .ColIndex("Transaction_ID")) = IIf(IsNull(Rs1("Transaction_ID").value), "", Rs1("Transaction_ID").value)
                 .TextMatrix(i, .ColIndex("NoteSerial1")) = IIf(IsNull(Rs1("NoteSerial1").value), "", Rs1("NoteSerial1").value)
                 .TextMatrix(i, .ColIndex("GroupID")) = IIf(IsNull(Rs1("GroupID").value), "", Rs1("GroupID").value)
                 .TextMatrix(i, .ColIndex("ItemCode")) = IIf(IsNull(Rs1("Fullcode").value), "", Rs1("Fullcode").value)
                 .TextMatrix(i, .ColIndex("ItemId")) = IIf(IsNull(Rs1("Item_ID").value), "", Rs1("Item_ID").value)
                 .TextMatrix(i, .ColIndex("ShowPrice")) = IIf(IsNull(Rs1("ShowPrice").value), "", Rs1("ShowPrice").value)
                 .TextMatrix(i, .ColIndex("Quantity")) = IIf(IsNull(Rs1("Quantity").value), "", Rs1("Quantity").value)
                 .TextMatrix(i, .ColIndex("PODays")) = IIf(IsNull(Rs1("PODays").value), "", Rs1("PODays").value)
                If Not (IsNull(Rs1("Transaction_Date").value)) Then
                    .TextMatrix(i, .ColIndex("Transaction_Date")) = Format(Rs1("Transaction_Date").value, "yyyy/M/d")
                End If
                
                 If SystemOptions.UserInterface = ArabicInterface Then
                 .TextMatrix(i, .ColIndex("Country")) = IIf(IsNull(Rs1("name").value), "", Rs1("name").value)
                 .TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(Rs1("CusName").value), "", Rs1("CusName").value)
                 .TextMatrix(i, .ColIndex("GroupName")) = IIf(IsNull(Rs1("GroupName").value), "", Rs1("GroupName").value)
                 .TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(Rs1("ItemName").value), "", Rs1("ItemName").value)
               
                Else
                 .TextMatrix(i, .ColIndex("Country")) = IIf(IsNull(Rs1("nameE").value), "", Rs1("nameE").value)
                 .TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(Rs1("CusNamee").value), "", Rs1("CusNamee").value)
                 .TextMatrix(i, .ColIndex("GroupName")) = IIf(IsNull(Rs1("GroupNamee").value), "", Rs1("GroupNamee").value)
                .TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(Rs1("ItemNamee").value), "", Rs1("ItemNamee").value)
                End If
                Rs1.MoveNext
    Next i
    End With
    End If
End Sub

Private Sub cmdAdd_Click()
GetDatat
End Sub

Private Sub CmdGroups_Click()
FrmSelectItemsGroups.show
End Sub

Private Sub CmDItems_Click()
FrmSelectItems.show

End Sub

Private Sub CmdOvers_Click()
 FrmSelectOrders.show
End Sub

Private Sub CmdRemove_Click()
    Dim X As Integer

    If SystemOptions.UserInterface = EnglishInterface Then
        X = MsgBox("Confirm Delete This line", vbCritical + vbYesNo)
    Else
        X = MsgBox(" √ﬂÌœ «·Õ–›", vbCritical + vbYesNo)
    End If

    If X = vbNo Then Exit Sub
    
    If VSFlexGrid1.Rows > 1 Then
        If VSFlexGrid1.Rows = 2 Then
            Me.VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
        Else

            If Me.VSFlexGrid1.Rows > 1 Then
                If Me.VSFlexGrid1.Row <> Me.VSFlexGrid1.FixedRows - 1 Then
                    Me.VSFlexGrid1.RemoveItem (Me.VSFlexGrid1.Row)
                End If
            End If
        End If
    End If
            


End Sub

Private Sub cmdUP_Click()
MoveUpDown List2, 0
End Sub

Private Sub CMDDown_Click()
MoveUpDown List2, 1
End Sub

Private Sub CmdVendors_Click()
FrmSelectVendor.show
FrmSelectVendor.Indxx = 0
End Sub

 

Private Sub DBCboClientName_KeyUp(KeyCode As Integer, _
                                  Shift As Integer)

    If KeyCode = vbKeyF5 Then
        Dim Dcombos As ClsDataCombos
        
        Set Dcombos = New ClsDataCombos
        Dcombos.GetCustomersSuppliers 0, Me.DBCboClientName, True
 
    End If

End Sub

Private Sub dcitems_KeyUp(KeyCode As Integer, _
                          Shift As Integer)

    If KeyCode = vbKeyF5 Then
        Dim Dcombos As ClsDataCombos
        
        Set Dcombos = New ClsDataCombos
           
        Dcombos.GetItemsNames dcitems
    End If

End Sub

Private Sub dcproject_Click(Area As Integer)

    If dcproject.BoundText = "" Then Exit Sub
    My_SQL = " select  fullcode,des from projects_des where project_id=" & val(dcproject.BoundText)
    fill_combo Dcterm, My_SQL

End Sub

Private Sub Dcterm_Click(Area As Integer)

    If Dcterm.BoundText = "" Then Exit Sub

    My_SQL = " select  fullcode,name from terms_operations where term_fullcode='" & Dcterm.BoundText & "'"
    fill_combo dcopr, My_SQL
End Sub

Private Sub Form_Load()

    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500
    ScreenNameArabic = " „ﬁ«—‰Â ⁄—Ê÷ «·«”⁄«—  "
    ScreenNameEnglish = "  Compare quotations "
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
FillListPolicy
    'My_SQL = " select id,Project_name from projects"
    'fill_combo dcproject, My_SQL
    '
    'My_SQL = " select  fullcode,des from projects_des"
    'fill_combo Dcterm, My_SQL

    'My_SQL = " select  fullcode,name from terms_operations"
    'fill_combo dcopr, My_SQL

    Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos
    Set cSearchDCombo = New clsDCboSearch
 
    Set BKGrndPic = New ClsBackGroundPic

    Dcombos.GetEmployees Me.DBCboClientName, True
   ' Dcombos.GetItemsNames dcitems

    With Me.Grid
        .Rows = 1
        .ExplorerBar = flexExSortShowAndMove
        .RowHeightMin = 300
        .ExtendLastCol = True
    End With
      
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    Set rs = New ADODB.Recordset
    StrSQL = "select * From TblComparPrice   "
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
  '   rs.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    XPBtnMove_Click 2
    Me.TxtModFlg.Text = "R"

    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If

End Sub

Private Sub ChangeLang()
    ChKauto.Caption = "Auto"
    Cmd(0).Caption = "New"
    Cmd(1).Caption = "Edit"
    Cmd(2).Caption = "Save"
    Cmd(3).Caption = "Undo"
    Cmd(4).Caption = "Delete"
    Cmd(5).Caption = "Search"
    Cmd(7).Caption = "Print"
    Cmd(6).Caption = "Exit"
    'CmdHelp.Caption = "Help"

    Dim XPic As IPictureDisp
    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic
    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic

    Me.Caption = "Compare Quotations "
    Ele(5).Caption = Me.Caption
    lbl(7).Caption = "Trans ID"
    lbl(9).Caption = "Date"
    lbl(0).Caption = "Based process"
    lbl(3).Caption = "Remarks"
    lbl(38).Caption = "This screen is Comparing Quotations"
    Frame1.Caption = "Terms of the Offer"
    Chkdates.Caption = "Select Date"
    lbl(5).Caption = "From Date"
    lbl(2).Caption = "To Date"
    ChKOvers.Caption = "Select Offer"
    lbl(14).Caption = "Select Offer"
    Frame3.Caption = "Items"
    lbl(4).Caption = "All Items"
    lbl(10).Caption = "Select Groups"
    lbl(6).Caption = "Select Items"
    Frame4.Caption = "Suppliers"
    lbl(12).Caption = "All Suppliers "
    lbl(11).Caption = "Select Suppliers "
   Frame2.Caption = "Best Price Policy"
   cmdAdd.Caption = "Execution"
   C1Tab1.Caption = "Basic Data|Analytical Data"
   ' Cmd(20).Caption = "Add"
   ' Cmd(21).Caption = "Remove"

   ' CmdRemove.Caption = "Remove Line"
Label1.Caption = "Selected Offers"
    With Me.VSFlexGrid1
        .TextMatrix(0, .ColIndex("Ser")) = "Serial"
        .TextMatrix(0, .ColIndex("NoteSerial1")) = "Offer No"
        .TextMatrix(0, .ColIndex("Transaction_Date")) = "Date"
        .TextMatrix(0, .ColIndex("CusName")) = "Supplier"
        .TextMatrix(0, .ColIndex("GroupName")) = "Group Name"
        .TextMatrix(0, .ColIndex("ItemCode")) = "Item Code"
        .TextMatrix(0, .ColIndex("ItemName")) = "Item Name"
        .TextMatrix(0, .ColIndex("Quantity")) = "Quantity"
        .TextMatrix(0, .ColIndex("ShowPrice")) = "Price"
        .TextMatrix(0, .ColIndex("PODays")) = "Due Date"
        .TextMatrix(0, .ColIndex("Country")) = "Country"
        

    End With
        With Me.GridOvers
        .TextMatrix(0, .ColIndex("Ser")) = "Serial"
        .TextMatrix(0, .ColIndex("NoteSerial1")) = "Offer No"
        .TextMatrix(0, .ColIndex("Transaction_Date")) = "Date"
        .TextMatrix(0, .ColIndex("CusName")) = "Supplier"
        .TextMatrix(0, .ColIndex("PODays")) = "No Day Offer"
    End With
            With Me.GridOvers
        .TextMatrix(0, .ColIndex("Ser")) = "Serial"
        .TextMatrix(0, .ColIndex("NoteSerial1")) = "Offer No"
        .TextMatrix(0, .ColIndex("Transaction_Date")) = "Date"
        .TextMatrix(0, .ColIndex("CusName")) = "Supplier"
        .TextMatrix(0, .ColIndex("PODays")) = "No Day Offer"
    End With
    Label4.Caption = "Selected Suppliers "
            With Me.gridVendor
        .TextMatrix(0, .ColIndex("Ser")) = "Serial"
        .TextMatrix(0, .ColIndex("Fullcode")) = "Code"
        .TextMatrix(0, .ColIndex("CusName")) = "Supplier"
       End With
    Label2.Caption = "Selected Groups"
         With Me.GridItemsGroup
        .TextMatrix(0, .ColIndex("Ser")) = "Serial"
        .TextMatrix(0, .ColIndex("Fullcode")) = "Group Code"
        .TextMatrix(0, .ColIndex("GroupName")) = "Group Name"
       End With
   Label6.Caption = "Selected Items"
      With Me.GridItems
        .TextMatrix(0, .ColIndex("Ser")) = "Serial"
        .TextMatrix(0, .ColIndex("ItemCode")) = "Item Code "
        .TextMatrix(0, .ColIndex("ItemName")) = "Item Name"
       End With

End Sub

'Public Sub get_all_employee()
'    Dim Rs3 As ADODB.Recordset
'    Set Rs3 = New ADODB.Recordset
'    Dim rs2 As ADODB.Recordset
''    Set rs2 = New ADODB.Recordset
 '   Dim j As Integer
'
'    Dim sql As String
'    Dim i As Integer
'
''    sql = "Select * from emp_all_details "
 '
 '   Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
 '
 '   If Rs3.RecordCount = 0 Then Exit Sub
 '
 '   With Grid
'
'        .Rows = 2
'        .Clear flexClearScrollable
'
'        If Rs3.RecordCount > 0 Then
'            .Rows = Rs3.RecordCount + 1
'            Rs3.MoveFirst
'
'            For i = 1 To Rs3.RecordCount
'                .TextMatrix(i, .ColIndex("Emp_id")) = IIf(IsNull(Rs3.Fields("Emp_id").value), "", Rs3.Fields("Emp_id").value)
'
'                .TextMatrix(i, .ColIndex("Emp_Code")) = IIf(IsNull(Rs3.Fields("Emp_Code").value), "", Rs3.Fields("Emp_Code").value)
'                .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(Rs3.Fields("Emp_Name").value), "", Rs3.Fields("Emp_Name").value)
'                .TextMatrix(i, .ColIndex("DepartmentName")) = IIf(IsNull(Rs3.Fields("DepartmentName").value), "", Rs3.Fields("DepartmentName").value)
'                .TextMatrix(i, .ColIndex("JobTypeName")) = IIf(IsNull(Rs3.Fields("JobTypeName").value), "", Rs3.Fields("JobTypeName").value)
''                .TextMatrix(i, .ColIndex("work_status")) = IIf(IsNull(Rs3.Fields("name").value), "", Rs3.Fields("name").value)
 '
 '               Rs3.MoveNext
 '           Next i
 '
 '           .AutoSize 0, .Cols - 1, False
 '       End If
'
'    End With
'
'    Rs3.Close
'
'End Sub
'
'Public Sub FillGridWithData()
'    Exit Sub
'
'    Dim i As Integer
'    Dim rs As ADODB.Recordset
'    Dim rs2 As ADODB.Recordset
'    Dim LstDay As Date
'    Dim FrstDay As Date
'    Dim StrTxt As String
'    Dim My_SQL As String
'    Dim StrWhere As String
'    Dim StrGrp As String
'    Dim IntMonth As Integer
'    Dim IntYear As Integer
'    Dim Msg As String
'
'    On Error GoTo ErrTrap
'
'    Set rs = New ADODB.Recordset
'    Set rs2 = New ADODB.Recordset
'
'    rs.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    With Me.Grid
'        .Rows = 2
'        .Clear flexClearScrollable
'
'        If rs.RecordCount > 0 Then
'            .Rows = rs.RecordCount + 1
'            rs.MoveFirst
'
'            For i = 1 To .Rows - 1
'
'                .TextMatrix(i, .ColIndex("Ser")) = i
'                ',DepartmentID,project_id
'
'                .TextMatrix(i, .ColIndex("dep")) = IIf(IsNull(rs.Fields("DepartmentID").value), "", rs.Fields("DepartmentID").value)
'
'                .TextMatrix(i, .ColIndex("project")) = IIf(IsNull(rs.Fields("project_id").value), "", rs.Fields("project_id").value)
'
'                .TextMatrix(i, .ColIndex("Emp_ID")) = IIf(IsNull(rs.Fields("Emp_ID").value), "", rs.Fields("Emp_ID").value)
'
'                .TextMatrix(i, .ColIndex("Emp_Code")) = IIf(IsNull(rs.Fields("Emp_Code").value), "", rs.Fields("Emp_Code").value)
'
'                .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(rs.Fields("Emp_Name").value), "", rs.Fields("Emp_Name").value)
'
'                .TextMatrix(i, .ColIndex("Emp_Salary")) = IIf(IsNull(rs.Fields("Emp_Salary").value), "", rs.Fields("Emp_Salary").value)
'
'                .TextMatrix(i, .ColIndex("TotalDiscount")) = IIf(IsNull(rs.Fields("TotalDiscount").value), "", Format(rs.Fields("TotalDiscount").value, SystemOptions.SysDefCurrencyForamt))
'
'                .TextMatrix(i, .ColIndex("Mokafea")) = IIf(IsNull(rs.Fields("TotalMokafea").value), "", Format(rs.Fields("TotalMokafea").value, SystemOptions.SysDefCurrencyForamt))
'
'                '.TextMatrix(I, .ColIndex("TotalAdvance")) = IIf(IsNull(Rs.Fields("TotalAdvance").Value), _
'                 "", Format(Rs.Fields("TotalAdvance").Value, SystemOptions.SysDefCurrencyForamt))
'
'                '   .TextMatrix(I, .ColIndex("EmpTotalNet")) = IIf(IsNull(Rs.Fields("EmpTotalNet").value), _
'                '      "", Format(Rs.Fields("EmpTotalNet").value, SystemOptions.SysDefCurrencyForamt))
'
'                .TextMatrix(i, .ColIndex("Emp_Salary_sakn")) = IIf(IsNull(rs.Fields("Emp_Salary_sakn").value), "", Format(rs.Fields("Emp_Salary_sakn").value))
'
'                .TextMatrix(i, .ColIndex("Emp_Salary_bus")) = IIf(IsNull(rs.Fields("Emp_Salary_bus").value), "", Format(rs.Fields("Emp_Salary_bus").value))
'
'                .TextMatrix(i, .ColIndex("Emp_Salary_food")) = IIf(IsNull(rs.Fields("Emp_Salary_food").value), "", Format(rs.Fields("Emp_Salary_food").value))
'
'                .TextMatrix(i, .ColIndex("Emp_Salary_mob")) = IIf(IsNull(rs.Fields("Emp_Salary_mob").value), "", Format(rs.Fields("Emp_Salary_mob").value))
'
'                .TextMatrix(i, .ColIndex("Emp_Salary_mang")) = IIf(IsNull(rs.Fields("Emp_Salary_mang").value), "", Format(rs.Fields("Emp_Salary_mang").value))
'
'                .TextMatrix(i, .ColIndex("Emp_Salary_others")) = IIf(IsNull(rs.Fields("Emp_Salary_others").value), "", Format(rs.Fields("Emp_Salary_others").value))
'
'                rs.MoveNext
'
'            Next
'
'            rs.Close
'        End If
'
'        .Rows = .Rows + 1
'        .TextMatrix(.Rows - 1, .ColIndex("Ser")) = "«·√Ã„«·Ï"
'        .IsSubtotal(.Rows - 1) = True
'        Dim SngTotal As Single
'        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary"), .Rows - 1, .ColIndex("Emp_Salary"))
'        .TextMatrix(.Rows - 1, .ColIndex("Emp_Salary")) = SngTotal
'
'        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("EmpTotalNet"), .Rows - 1, .ColIndex("EmpTotalNet"))
'        .TextMatrix(.Rows - 1, .ColIndex("EmpTotalNet")) = SngTotal
'        net_value = SngTotal
'        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("CorrectEmpTotalNet"), .Rows - 1, .ColIndex("CorrectEmpTotalNet"))
'        .TextMatrix(.Rows - 1, .ColIndex("CorrectEmpTotalNet")) = SngTotal
'
'        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary_sakn"), .Rows - 1, .ColIndex("Emp_Salary_sakn"))
'        .TextMatrix(.Rows - 1, .ColIndex("Emp_Salary_sakn")) = SngTotal
'
'        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary_bus"), .Rows - 1, .ColIndex("Emp_Salary_bus"))
'        .TextMatrix(.Rows - 1, .ColIndex("Emp_Salary_bus")) = SngTotal
'
'        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary_food"), .Rows - 1, .ColIndex("Emp_Salary_food"))
'        .TextMatrix(.Rows - 1, .ColIndex("Emp_Salary_food")) = SngTotal
'
'        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary_others"), .Rows - 1, .ColIndex("Emp_Salary_others"))
'        .TextMatrix(.Rows - 1, .ColIndex("Emp_Salary_others")) = SngTotal
'
'        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("OverTimePrice"), .Rows - 1, .ColIndex("OverTimePrice"))
'        .TextMatrix(.Rows - 1, .ColIndex("OverTimePrice")) = SngTotal
'
'        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Mokafea"), .Rows - 1, .ColIndex("Mokafea"))
'        .TextMatrix(.Rows - 1, .ColIndex("Mokafea")) = SngTotal
'
'        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("SalesCom"), .Rows - 1, .ColIndex("SalesCom"))
'        .TextMatrix(.Rows - 1, .ColIndex("SalesCom")) = SngTotal
'
'        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("TotalAdvance"), .Rows - 1, .ColIndex("TotalAdvance"))
'        .TextMatrix(.Rows - 1, .ColIndex("TotalAdvance")) = SngTotal
'
'        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("TotalDiscount"), .Rows - 1, .ColIndex("TotalDiscount"))
'        .TextMatrix(.Rows - 1, .ColIndex("TotalDiscount")) = SngTotal
'
'        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("total1"), .Rows - 1, .ColIndex("total1"))
'        .TextMatrix(.Rows - 1, .ColIndex("total1")) = SngTotal
'
'        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("total2"), .Rows - 1, .ColIndex("total2"))
'        .TextMatrix(.Rows - 1, .ColIndex("total2")) = SngTotal
    
'        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary_mang"), .Rows - 1, .ColIndex("Emp_Salary_mang"))
'        .TextMatrix(.Rows - 1, .ColIndex("Emp_Salary_mang")) = SngTotal
'
'        SngTotal = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Emp_Salary_mob"), .Rows - 1, .ColIndex("Emp_Salary_mob"))
'        .TextMatrix(.Rows - 1, .ColIndex("Emp_Salary_mob")) = SngTotal
'
'        .Cell(flexcpBackColor, .Rows - 1, 1, .Rows - 1, .Cols - 1) = vbYellow
'        .Cell(flexcpFontBold, .Rows - 1, 1, .Rows - 1, .Cols - 1) = True
'        .Cell(flexcpFontSize, .Rows - 1, 1, .Rows - 1, .Cols - 1) = 10
'        .Cell(flexcpFontName, .Rows - 1, 1, .Rows - 1, .Cols - 1) = "Tahoma"
'        .AutoSize 0, .Cols - 1, False
'    End With
'
'ErrTrap:
'End Sub
 
Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)

    Dim IntResult As String
    Dim StrMSG As String
    On Error GoTo ErrTrap

    If Me.TxtModFlg.Text <> "R" Then

        Select Case Me.TxtModFlg.Text

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

'Function CuurentLogdata(Optional Currentmode As String)
'
'    LogTextA = "    ‘«‘… " & ScreenNameArabic & Chr(13) & " —ﬁ„ «·« ›«ﬁÌ…    " & TxtTblVendorContractD.text & Chr(13) & " «·„Ê—œ " & DBCboClientName.text & Chr(13) & "  „œ Â« „‰  " & dbFromDate & Chr(13) & "  «·Ï " & dbTodate & Chr(13) & "  „·«ÕŸ«  " & txtRemarks
'
'    If ChkLocked.value = Checked Then
'        LogTextA = LogTextA & Chr(13) & "   „ «Ìﬁ«› «· ⁄«„· "
'    End If
'
'    LogTextE = "    Screen  " & ScreenNameEnglish & Chr(13) & " Contract No    " & TxtTblVendorContractD.text & Chr(13) & " Supplier " & DBCboClientName.text & Chr(13) & " From   " & dbFromDate & Chr(13) & "  To  " & dbTodate & Chr(13) & "  Remarks " & txtRemarks
'
'    If ChkLocked.value = Checked Then
'        LogTextA = LogTextA & Chr(13) & " Locked "
'    End If
'
'    If Currentmode <> "D" Then
'        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTextE, Me.name, Me.TxtModFlg
'    Else
'        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTextE, Me.name, "D"
'    End If
    
'End Function

Private Sub Form_Unload(Cancel As Integer)
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish
End Sub

Private Sub Grid_AfterEdit(ByVal Row As Long, _
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
                .TextMatrix(Row, .ColIndex("UnitID")) = code
                .TextMatrix(Row, .ColIndex("UnitName")) = .ComboItem
 
        End Select
   
        If Row = .Rows - 1 Then
    
            '.Rows = .Rows + 1
        End If

        ReLineGrid
    End With

    If Me.TxtModFlg <> "E" Then Exit Sub

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\LogFile Saving//////////////////////////////////////////////////////////////////////////
    'Grid.TextMatrix(Row, Grid.ColIndex("Code"))
    'Grid.TextMatrix(Row, Grid.ColIndex("Name"))
    If Col = Grid.ColIndex("ItemCode") Or Col = Grid.ColIndex("ItemName") Then
        RegisterItemData Me.Name, Me.TxtModFlg, Grid.TextMatrix(Row, Grid.ColIndex("ItemCode")), Grid.TextMatrix(Row, Grid.ColIndex("ItemName")), , , , , , , , , , , Me.TxtTblVendorContractD
    ElseIf Col = Grid.ColIndex("UnitName") Then
        RegisterItemData Me.Name, Me.TxtModFlg, Grid.TextMatrix(Row, Grid.ColIndex("ItemCode")), Grid.TextMatrix(Row, Grid.ColIndex("ItemName")), Grid.TextMatrix(Row, Grid.ColIndex("UnitName")), , , , , , , , , , Me.TxtTblVendorContractD
    ElseIf Col = Grid.ColIndex("Price") Then
        RegisterItemData Me.Name, Me.TxtModFlg, Grid.TextMatrix(Row, Grid.ColIndex("ItemCode")), Grid.TextMatrix(Row, Grid.ColIndex("ItemName")), , , (Grid.TextMatrix(Row, Grid.ColIndex("Price"))), , , , , , , , Me.TxtTblVendorContractD
    ElseIf Col = Grid.ColIndex("Discount") Then
        RegisterItemData Me.Name, Me.TxtModFlg, Grid.TextMatrix(Row, Grid.ColIndex("ItemCode")), Grid.TextMatrix(Row, Grid.ColIndex("ItemName")), , , , , , , , , Grid.TextMatrix(Row, Grid.ColIndex("Discount")), , Me.TxtTblVendorContractD

    End If

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\LogFile Saving//////////////////////////////////////////////////////////////////////////

End Sub

Private Sub ReLineGrid()
    Dim IntCounter As Integer
    IntCounter = 0
    Dim i As Integer

    With Me.VSFlexGrid1

        For i = .FixedRows To .Rows - 1
    
            If .TextMatrix(i, .ColIndex("ItemId")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("Ser")) = IntCounter
  
            End If

        Next i
   
    End With

End Sub

'Private Sub Grid_BeforeEdit(ByVal Row As Long, _
'                            ByVal Col As Long, _
'                            Cancel As Boolean)
'
'    With Grid
'
'        If .ColKey(Col) <> "UnitName" Then
'
'            .ComboList = ""
'        End If
'
'    End With
'
'End Sub
'
'Private Sub Grid_StartEdit(ByVal Row As Long, _
'                           ByVal Col As Long, _
'                           Cancel As Boolean)
'    Dim rs As New ADODB.Recordset
'    Dim StrSQL  As String
'    Dim StrAccountType As String
'    Dim StrComboList As String
'    Dim Msg As String
'    Dim LngItemID As Integer
'    Dim MyStrList As String
'
'    With Me.Grid
'
'        Select Case .ColKey(Col)
'
'            Case "UnitName"
'
'                LngItemID = val(.TextMatrix(.Row, .ColIndex("ItemId")))
'
'                'LngItemID = 1
'                If LngItemID = 0 Then
'                    Cancel = True
'                Else
'
'                    StrSQL = "SELECT TblItemsUnits.UnitID, TblUnites.UnitName "
'                    StrSQL = StrSQL + " FROM TblUnites INNER JOIN TblItemsUnits " & "ON TblUnites.UnitID = TblItemsUnits.UnitID "
'                    StrSQL = StrSQL + " Where TblItemsUnits.ItemID=" & LngItemID
'                    StrSQL = StrSQL + " Order BY TblItemsUnits.SecOrder "
'                    Set rs = New ADODB.Recordset
'                    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
'
'                    If Not (rs.BOF Or rs.EOF) Then
'                        MyStrList = .BuildComboList(rs, "UnitName", "UnitID")
'                        '                    Grid.ColComboList = MyStrList
'                        Grid.ColComboList(.ColIndex("UnitName")) = "|" & MyStrList
'                    Else
'                        Cancel = True
''                    End If
 '               End If
 '
 '       End Select
'
'    End With
'
'End Sub
'
Public Sub Retrive(Optional Lngid As Long = 0)
    Dim RsDev As ADODB.Recordset
    Dim StrSQL As String
    Dim i As Integer

    'On Error GoTo ErrTrap
    VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
    VSFlexGrid1.Rows = 1
          
    If rs.RecordCount < 1 Then
        Exit Sub
    End If

       If Lngid <> 0 Then
            rs.find "ID=" & Lngid, , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Exit Sub
            End If
        End If
 
    Me.TxtTblVendorContractD.Text = IIf(IsNull(rs("ID").value), "", rs("ID").value)
 
    DTPicker1.value = IIf(IsNull(rs("RecordDate").value), Date, rs("RecordDate").value)
     dbFromDate.value = IIf(IsNull(rs("FromDate").value), Date, rs("FromDate").value)
    dbTodate.value = IIf(IsNull(rs("Todate").value), Date, rs("Todate").value)

    DBCboClientName.BoundText = IIf(IsNull(rs("EmpID").value), "", rs("EmpID").value)

    txtRemarks.Text = IIf(IsNull(rs("Remarks").value), "", rs("Remarks").value)
    StrOrder = IIf(IsNull(rs("StrOrder").value), "", rs("StrOrder").value)
    StrItemID = IIf(IsNull(rs("StrItemID").value), "", rs("StrItemID").value)
    StrGroupID = IIf(IsNull(rs("StrGroupID").value), "", rs("StrGroupID").value)
    StrCusID = IIf(IsNull(rs("StrCusID").value), "", rs("StrCusID").value)
    policy = IIf(IsNull(rs("policy").value), "", rs("policy").value)
    retrivepolicy policy
    If rs("Chkdates").value = False Then
        Chkdates.value = vbUnchecked
    Else
        Chkdates.value = vbChecked
    End If
    If rs("ChKOvers").value = False Then
        ChKOvers.value = vbUnchecked
    Else
        ChKOvers.value = vbChecked
    End If
    If rs("ChItem").value = 1 Then
        OptItems(0).value = True
    ElseIf rs("ChItem").value = 2 Then
        OptItems(1).value = True
     ElseIf rs("ChItem").value = 3 Then
       OptItems(2).value = True
    End If
    
      If rs("ChCustomer").value = 1 Then
        OptVendor(0).value = True
    ElseIf rs("ChCustomer").value = 2 Then
        OptVendor(2).value = True
      End If
 
    ''//
StrSQL = " SELECT     dbo.TblComparPriceDet.ID, dbo.TblComparPriceDet.CoPriceID, dbo.TblComparPriceDet.GroupID, dbo.Groups.GroupName, dbo.Groups.GroupCode, "
StrSQL = StrSQL & "                      dbo.Groups.GroupNamee, dbo.Groups.Fullcode, dbo.TblComparPriceDet.CusID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee,"
StrSQL = StrSQL & "                      dbo.TblCustemers.Fullcode AS CusFullcode, dbo.TblComparPriceDet.ItemId, dbo.TblItems.Fullcode AS ItemFullcode, dbo.TblItems.ItemCode, dbo.TblItems.ItemName,"
StrSQL = StrSQL & "                      dbo.TblItems.ItemNamee, dbo.TblComparPriceDet.Transaction_ID, dbo.TblComparPriceDet.NoteSerial1, dbo.TblComparPriceDet.Quantity,"
StrSQL = StrSQL & "                      dbo.TblComparPriceDet.ShowPrice, dbo.TblComparPriceDet.PODays, dbo.TblComparPriceDet.Transaction_Date, dbo.TblComparPriceDet.CountryID,"
StrSQL = StrSQL & "                      dbo.nationality.name , dbo.nationality.NameE, dbo.nationality.Quality"
StrSQL = StrSQL & " FROM         dbo.TblComparPriceDet LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.Nationality ON dbo.TblComparPriceDet.CountryID = dbo.Nationality.id LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblItems ON dbo.TblComparPriceDet.ItemId = dbo.TblItems.ItemID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblCustemers ON dbo.TblComparPriceDet.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.Groups ON dbo.TblComparPriceDet.GroupID = dbo.Groups.GroupID"
StrSQL = StrSQL & " Where dbo.TblComparPriceDet.CoPriceID = " & val(Me.TxtTblVendorContractD.Text)
    Set RsDev = New ADODB.Recordset
    RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsDev.BOF Or RsDev.EOF) Then
        RsDev.MoveFirst
    
        With Me.VSFlexGrid1
    
            .Rows = .FixedRows + RsDev.RecordCount

            For i = .FixedRows To .Rows - 1
             .TextMatrix(i, .ColIndex("Ser")) = i
                .TextMatrix(i, .ColIndex("CusID")) = IIf(IsNull(RsDev("CusID").value), "", RsDev("CusID").value)
                 .TextMatrix(i, .ColIndex("Transaction_ID")) = IIf(IsNull(RsDev("Transaction_ID").value), "", RsDev("Transaction_ID").value)
                 .TextMatrix(i, .ColIndex("NoteSerial1")) = IIf(IsNull(RsDev("NoteSerial1").value), "", RsDev("NoteSerial1").value)
                 .TextMatrix(i, .ColIndex("GroupID")) = IIf(IsNull(RsDev("GroupID").value), "", RsDev("GroupID").value)
                 .TextMatrix(i, .ColIndex("ItemCode")) = IIf(IsNull(RsDev("Fullcode").value), "", RsDev("Fullcode").value)
                 .TextMatrix(i, .ColIndex("ItemId")) = IIf(IsNull(RsDev("ItemId").value), "", RsDev("ItemId").value)
                 .TextMatrix(i, .ColIndex("ShowPrice")) = IIf(IsNull(RsDev("ShowPrice").value), "", RsDev("ShowPrice").value)
                 .TextMatrix(i, .ColIndex("Quantity")) = IIf(IsNull(RsDev("Quantity").value), "", RsDev("Quantity").value)
                 .TextMatrix(i, .ColIndex("PODays")) = IIf(IsNull(RsDev("PODays").value), "", RsDev("PODays").value)
                 .TextMatrix(i, .ColIndex("CountryID")) = IIf(IsNull(RsDev("CountryID").value), "", RsDev("CountryID").value)
                If Not (IsNull(RsDev("Transaction_Date").value)) Then
                    .TextMatrix(i, .ColIndex("Transaction_Date")) = Format(RsDev("Transaction_Date").value, "yyyy/M/d")
                End If
                 If SystemOptions.UserInterface = ArabicInterface Then
                 .TextMatrix(i, .ColIndex("Country")) = IIf(IsNull(RsDev("name").value), "", RsDev("name").value)
                 .TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(RsDev("CusName").value), "", RsDev("CusName").value)
                 .TextMatrix(i, .ColIndex("GroupName")) = IIf(IsNull(RsDev("GroupName").value), "", RsDev("GroupName").value)
                 .TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(RsDev("ItemName").value), "", RsDev("ItemName").value)
               
                Else
                .TextMatrix(i, .ColIndex("Country")) = IIf(IsNull(RsDev("nameE").value), "", RsDev("nameE").value)
                 .TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(RsDev("CusNamee").value), "", RsDev("CusNamee").value)
                 .TextMatrix(i, .ColIndex("GroupName")) = IIf(IsNull(RsDev("GroupNamee").value), "", RsDev("GroupNamee").value)
                 .TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(RsDev("ItemNamee").value), "", RsDev("ItemNamee").value)
                End If
            
                RsDev.MoveNext
            Next i
 
        End With

    End If
 
    ReLineGrid
    Exit Sub
ErrTrap:
End Sub
 
 

Private Sub OptItems_Click(Index As Integer)
Select Case Index
Case 0
CmdGroups.Enabled = False
CmDItems.Enabled = False
Case 1
CmdGroups.Enabled = True
CmDItems.Enabled = False
Case 2
CmdGroups.Enabled = False
CmDItems.Enabled = True

End Select
End Sub

Private Sub OptVendor_Click(Index As Integer)
Select Case Index
Case 0
 
CmdVendors.Enabled = False
Case 1
 
CmdVendors.Enabled = False
Case 2
 
CmdVendors.Enabled = True

End Select
End Sub

Private Sub TxtModFlg_Change()

    If Me.TxtModFlg.Text = "N" Then
        CmdRemove.Enabled = True
        Ele(1).Enabled = True
        Cmd(0).Enabled = False
        Cmd(1).Enabled = False
        Cmd(4).Enabled = False
        Cmd(5).Enabled = False

        Cmd(2).Enabled = True
        Cmd(3).Enabled = True

    ElseIf Me.TxtModFlg.Text = "E" Then
        CmdRemove.Enabled = True
        Ele(1).Enabled = True
        Cmd(2).Enabled = True
        Cmd(3).Enabled = True

        Cmd(0).Enabled = False
        Cmd(1).Enabled = False
        Cmd(4).Enabled = False

        Cmd(5).Enabled = False

    Else
        Ele(1).Enabled = False

        CmdRemove.Enabled = False
        Cmd(2).Enabled = False
        Cmd(3).Enabled = False
        Cmd(0).Enabled = True
        Cmd(1).Enabled = True
        Cmd(4).Enabled = True

        Cmd(5).Enabled = True

    End If

End Sub

Private Sub XPBtnMove_Click(Index As Integer)

    If Me.TxtModFlg.Text = "N" Then
        clear_all Me
        Me.TxtModFlg.Text = "R"
        XPBtnMove_Click (1)
    End If

    On Error GoTo ErrTrap

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
