VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Begin VB.Form FrmManStore 
   Caption         =   "„Œ“‰ «·’Ì«‰…"
   ClientHeight    =   8550
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11010
   Icon            =   "FrmManStore.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   8550
   ScaleWidth      =   11010
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin C1SizerLibCtl.C1Elastic EleMain 
      Height          =   8550
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   11010
      _cx             =   19420
      _cy             =   15081
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
      AutoSizeChildren=   8
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
      GridRows        =   4
      GridCols        =   4
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"FrmManStore.frx":038A
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   750
         Index           =   7
         Left            =   15
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   15
         Width           =   10980
         _cx             =   19368
         _cy             =   1323
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
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈÷€ÿ œ«∆„«⁄·Ï “—  ÕœÌÀ · ÕœÌÀ «·»Ì«‰« "
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
            Height          =   225
            Left            =   30
            RightToLeft     =   -1  'True
            TabIndex        =   22
            Top             =   480
            Visible         =   0   'False
            Width           =   2835
         End
      End
      Begin C1SizerLibCtl.C1Tab TabMain 
         Height          =   7095
         Left            =   15
         TabIndex        =   6
         Top             =   780
         Width           =   10980
         _cx             =   19368
         _cy             =   12515
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
         Caption         =   "„Œ“‰ «·’Ì«‰…|÷„«‰«  Ê ’·ÌÕ Œ«—ÃÌ|√ÃÂ“… „ÿ·Ê»  Ã„Ì⁄Â«|Ã«Â“ ·· ”·Ì„|√ÃÂ“…  „  Ã„Ì⁄Â«"
         Align           =   0
         CurrTab         =   3
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
         Picture(0)      =   "FrmManStore.frx":040B
         Picture(1)      =   "FrmManStore.frx":07A5
         Picture(2)      =   "FrmManStore.frx":0B3F
         Picture(3)      =   "FrmManStore.frx":0ED9
         Picture(4)      =   "FrmManStore.frx":1273
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   6630
            Index           =   2
            Left            =   11625
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   45
            Width           =   10890
            _cx             =   19209
            _cy             =   11695
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
            Begin VSFlex8UCtl.VSFlexGrid FgCompAssblied 
               Height          =   6555
               Left            =   0
               TabIndex        =   19
               Top             =   30
               Width           =   10890
               _cx             =   19209
               _cy             =   11562
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
               Cols            =   9
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"FrmManStore.frx":180D
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
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   6630
            Index           =   6
            Left            =   45
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   45
            Width           =   10890
            _cx             =   19209
            _cy             =   11695
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
            Begin VSFlex8UCtl.VSFlexGrid FgReady 
               Height          =   6540
               Left            =   0
               TabIndex        =   20
               Top             =   30
               Width           =   10890
               _cx             =   19209
               _cy             =   11536
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
               FloodColor      =   0
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
               Cols            =   16
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   345
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmManStore.frx":197F
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
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   6630
            Index           =   5
            Left            =   -11535
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   45
            Width           =   10890
            _cx             =   19209
            _cy             =   11695
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
            Begin VSFlex8UCtl.VSFlexGrid FgRequired 
               Height          =   6555
               Left            =   0
               TabIndex        =   17
               Top             =   30
               Width           =   10890
               _cx             =   19209
               _cy             =   11562
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
               FormatString    =   $"FrmManStore.frx":1BDE
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
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   6630
            Index           =   4
            Left            =   -11835
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   45
            Width           =   10890
            _cx             =   19209
            _cy             =   11695
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
            Begin VSFlex8UCtl.VSFlexGrid FgOutMan 
               Height          =   6555
               Left            =   0
               TabIndex        =   15
               Top             =   30
               Width           =   10890
               _cx             =   19209
               _cy             =   11562
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
               Cols            =   8
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"FrmManStore.frx":1D74
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
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   6630
            Index           =   1
            Left            =   -12135
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   45
            Width           =   10890
            _cx             =   19209
            _cy             =   11695
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
            BackColor       =   4210752
            ForeColor       =   -2147483630
            FloodColor      =   6553600
            ForeColorDisabled=   -2147483631
            Caption         =   ""
            Align           =   0
            AutoSizeChildren=   7
            BorderWidth     =   2
            ChildSpacing    =   6
            Splitter        =   -1  'True
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
               Height          =   6570
               Index           =   3
               Left            =   0
               TabIndex        =   8
               TabStop         =   0   'False
               Top             =   30
               Width           =   10890
               _cx             =   19209
               _cy             =   11589
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
               GridRows        =   3
               GridCols        =   4
               Frame           =   3
               FrameStyle      =   0
               FrameWidth      =   1
               FrameColor      =   -2147483628
               FrameShadow     =   -2147483632
               FloodStyle      =   1
               _GridInfo       =   $"FrmManStore.frx":1EBB
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   9
               Begin VSFlex8UCtl.VSFlexGrid Fg 
                  Height          =   6540
                  Left            =   15
                  TabIndex        =   12
                  Top             =   15
                  Width           =   10860
                  _cx             =   19156
                  _cy             =   11536
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
                  Cols            =   15
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   345
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"FrmManStore.frx":1F32
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
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   645
         Index           =   0
         Left            =   15
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   7890
         Width           =   10980
         _cx             =   19368
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
            Height          =   375
            Index           =   0
            Left            =   8520
            TabIndex        =   2
            Top             =   150
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   661
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
            ColorTextShadow =   -2147483637
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   375
            Index           =   4
            Left            =   5790
            TabIndex        =   3
            Top             =   150
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
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
            ColorTextShadow =   -2147483637
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   375
            Index           =   5
            Left            =   4440
            TabIndex        =   4
            Top             =   150
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   661
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
            ColorTextShadow =   -2147483637
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   375
            Index           =   6
            Left            =   570
            TabIndex        =   5
            Top             =   150
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   661
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
            ColorTextShadow =   -2147483637
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   375
            Index           =   1
            Left            =   7260
            TabIndex        =   13
            Top             =   150
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   661
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
            DisabledImageExtraction=   0
            ColorToggledHoverText=   16711680
            ColorTextShadow =   -2147483637
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   375
            Index           =   2
            Left            =   3150
            TabIndex        =   14
            Top             =   150
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   661
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
            ColorTextShadow =   -2147483637
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   375
            Index           =   3
            Left            =   1860
            TabIndex        =   16
            Top             =   150
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   661
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   " ÕœÌÀ"
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
            ColorTextShadow =   -2147483637
         End
      End
   End
End
Attribute VB_Name = "FrmManStore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Cmd_Click(Index As Integer)

    Select Case Index

        Case 0
            Load FrmManAddNew
            FrmManAddNew.TxtModFlg.text = "N"
            FrmManAddNew.Show

        Case 3
            'Refresh
            LoadData

        Case 4
            DelRecord

        Case 6
            Unload Me
    End Select

End Sub

Private Sub Fg_DblClick()

    With Me.FG

        If .Col <= 0 Then Exit Sub
        If .Row <= 0 Then Exit Sub
        If .IsSubtotal(.Row) = True Then
            'ModFgLib.ExpandCollapsRow Fg, Fg.Row
        End If

        Select Case .ColKey(.Col)

            Case "MaintananceID", "ReciptNumber", "TicketNO"
        End Select

    End With

End Sub

Private Sub Fg_MouseMove(Button As Integer, _
                         Shift As Integer, _
                         x As Single, _
                         Y As Single)
    Dim StrTip As String
    On Error GoTo hErr

    With FG

        If .MouseRow <= 0 Then
            .MousePointer = flexDefault
            .ToolTipText = vbNullString
            Exit Sub
        End If

        If .MouseCol <= 0 Then
            .MousePointer = flexDefault
            .ToolTipText = vbNullString
            Exit Sub
        End If

        If .IsSubtotal(.MouseRow) = True Then
            If SystemOptions.UserInterface = ArabicInterface Then
                StrTip = "≈÷€ÿ „— Ì‰ „  «· Ì‰ (·⁄—÷ √Ê ≈Œ›«¡)  ﬁ«—Ì— «·„ «»⁄… «·Œ«’… »«·’‰›"
            Else
                StrTip = " Double Click To show and Hide Data  "
            End If

            .ToolTipText = StrTip
            .MousePointer = flexHand
        Else
            .MousePointer = flexDefault
            .ToolTipText = vbNullString
        End If

    End With

    Exit Sub
hErr:
    FG.MousePointer = flexDefault
    FG.ToolTipText = vbNullString
End Sub

Private Sub Fg_MouseUp(Button As Integer, _
                       Shift As Integer, _
                       x As Single, _
                       Y As Single)
    Dim LngManID As Long
    Dim LngMouseRow As Long
    Dim LngTicktNO As Long
    Dim StrTemp  As String
    Dim LngItemID As Long
    Dim StrItemSerial As String
    Dim LngReciptNO As Long

    mdifrmmain.MnuManTools.Tag = ""

    If Button = vbRightButton Then

        With Me.FG

            If .MouseRow <> 0 Then
                LngMouseRow = .MouseRow

                If LngMouseRow <= 0 Then Exit Sub
                If .IsSubtotal(LngMouseRow) = True Then
                    LngManID = val(.TextMatrix(LngMouseRow, .ColIndex("MaintananceID")))
                    LngTicktNO = val(.TextMatrix(LngMouseRow, .ColIndex("TicketNO")))
                 
                    LngReciptNO = val(.TextMatrix(LngMouseRow, .ColIndex("ReciptNumber")))
                 
                    StrTemp = LngManID & "-" & LngTicktNO & "-" & LngReciptNO
                    LngItemID = val(.TextMatrix(LngMouseRow, .ColIndex("ItemID")))
                    StrItemSerial = Trim$(.TextMatrix(LngMouseRow, .ColIndex("Item_Serial")))
                Else
                    Exit Sub
                End If

                mdifrmmain.MnuManTools.Tag = StrTemp
                StrTemp = LngItemID & ";" & StrItemSerial

                If StrItemSerial <> "" Then
                    mdifrmmain.MnuManToolsSub6.Tag = StrTemp
                    mdifrmmain.MnuManToolsSub6.Enabled = True

                    If SystemOptions.UserInterface = ArabicInterface Then
                
                    Else
                        mdifrmmain.MnuManToolsSub6.Caption = "Show Details"
                        mdifrmmain.MnuManTools.Caption = "Show Detalis"
                    End If

                Else
                    mdifrmmain.MnuManToolsSub6.Enabled = False
                End If

                Me.PopupMenu mdifrmmain.MnuManTools
            End If

        End With

    End If

End Sub

Private Sub FgCompAssblied_MouseUp(Button As Integer, _
                                   Shift As Integer, _
                                   x As Single, _
                                   Y As Single)
    Dim LngMouseRow As Long
    Dim LngTableID As Long
    Dim StrTemp As String

    With Me.FgRequired
        mdifrmmain.MnuManTools2.Tag = ""

        If Button = vbRightButton Then

            With Me.FgCompAssblied

                If .MouseRow <> 0 Then
                    LngMouseRow = .MouseRow

                    If LngMouseRow <= 0 Then Exit Sub
                    If .IsSubtotal(LngMouseRow) = True Then
                        LngTableID = val(.TextMatrix(LngMouseRow, .ColIndex("TableID")))
                        StrTemp = LngTableID
                    Else
                        Exit Sub
                    End If

                    mdifrmmain.MnuManTools2.Tag = StrTemp
                    mdifrmmain.MnuManTools2Sub1.Visible = False
                    mdifrmmain.MnuManTools2Sub2.Visible = True
                    Me.PopupMenu mdifrmmain.MnuManTools2
                    LoadManAlrams
                    LoadManAlramsDone
                End If

            End With

        End If

    End With

End Sub

Private Sub FgRequired_Click()
    Dim LngTrans As Long

    With Me.FgRequired

        If .Col <= -1 Then Exit Sub
        If .Row <= -1 Then Exit Sub
        If .Col = .ColIndex("View") Then
            LngTrans = val(.TextMatrix(.Row, .ColIndex("TransID")))

            If LngTrans <> 0 Then
                OpenScreen InvoiceScreen, LngTrans
            End If
        End If

    End With

End Sub

Private Sub FgRequired_MouseMove(Button As Integer, _
                                 Shift As Integer, _
                                 x As Single, _
                                 Y As Single)
    Dim LngMouseRow As Long
    Dim LngMouseCol As Long

    With Me.FgRequired
        .MousePointer = flexDefault

        If .Rows = .FixedRows Then Exit Sub
    
        .Cell(flexcpFontUnderline, .FixedRows, .ColIndex("View"), .Rows - 1, .ColIndex("View")) = False

        If .MouseRow <= -1 Or .MouseCol <= -1 Then
            Exit Sub
        ElseIf .ColKey(.MouseCol) = "View" Then
            .Cell(flexcpFontUnderline, .FixedRows, .ColIndex("View"), .Rows - 1, .ColIndex("View")) = False
            LngMouseRow = .MouseRow

            If val(.TextMatrix(LngMouseRow, .ColIndex("TransID"))) <> 0 Then
                .Cell(flexcpFontUnderline, LngMouseRow, .ColIndex("View"), LngMouseRow, .ColIndex("View")) = True
                .MousePointer = flexHand
            End If
        End If

    End With

End Sub

Private Sub FgRequired_MouseUp(Button As Integer, _
                               Shift As Integer, _
                               x As Single, _
                               Y As Single)
    Dim LngMouseRow As Long
    Dim LngTableID As Long
    Dim StrTemp As String

    With Me.FgRequired
        mdifrmmain.MnuManTools2.Tag = ""

        If Button = vbRightButton Then

            With Me.FgRequired

                If .MouseRow <> 0 Then
                    LngMouseRow = .MouseRow

                    If LngMouseRow <= 0 Then Exit Sub
                    If .IsSubtotal(LngMouseRow) = True Then
                        LngTableID = val(.TextMatrix(LngMouseRow, .ColIndex("TableID")))
                        StrTemp = LngTableID
                    Else
                        Exit Sub
                    End If

                    mdifrmmain.MnuManTools2.Tag = StrTemp
                    mdifrmmain.MnuManTools2Sub1.Visible = True
                    mdifrmmain.MnuManTools2Sub2.Visible = False
                    Me.PopupMenu mdifrmmain.MnuManTools2
                    LoadManAlrams
                    LoadManAlramsDone
                End If

            End With

        End If

    End With

End Sub

Private Sub ChangeLang()

    Me.Caption = "Maintenance Follow"
    Ele(7).Caption = Me.Caption

    Cmd(0).Caption = "New"
    Cmd(1).Caption = "Modify"
    Cmd(5).Caption = "Search"
    Cmd(2).Caption = "Print"
    Cmd(4).Caption = "Delete"
    Cmd(3).Caption = "Refresh"
    Cmd(6).Caption = "Exit"

    TabMain.TabCaption(0) = "Maintenance Store"
    TabMain.TabCaption(1) = " Warranty and repair external"
    TabMain.TabCaption(2) = "Devices required assembled"
    TabMain.TabCaption(3) = "Ready for delivery"
    TabMain.TabCaption(4) = "Devices were assembled"

    With FgOutMan
        .TextMatrix(0, .ColIndex("TicketNO")) = "Ticket NO"
        .TextMatrix(0, .ColIndex("ItemCode")) = "Item Code"
        .TextMatrix(0, .ColIndex("itemname")) = "Item Name"
        .TextMatrix(0, .ColIndex("ItemSerial")) = "Item_Serial"

        .TextMatrix(0, .ColIndex("outDate")) = "Out Date"
        .TextMatrix(0, .ColIndex("expectedReturndate")) = "Expected Return Date"
    End With

    With Me.FG
        .TextMatrix(0, .ColIndex("ReciptNumber")) = "Recipt Number"
        .TextMatrix(0, .ColIndex("TicketNO")) = "Ticket NO"
        .TextMatrix(0, .ColIndex("ItemCode")) = "Item Code"
        .TextMatrix(0, .ColIndex("itemname")) = "Item Name"
        .TextMatrix(0, .ColIndex("Item_Serial")) = "Item_Serial"
        .TextMatrix(0, .ColIndex("Qty")) = "Quantity"
        .TextMatrix(0, .ColIndex("MType")) = "Maint. Type"
        .TextMatrix(0, .ColIndex("CusName")) = "Cus. Name"
        .TextMatrix(0, .ColIndex("CustomerNotes")) = "Customer complaint "
        .TextMatrix(0, .ColIndex("EmpNotes")) = "Employee Notes"
        .TextMatrix(0, .ColIndex("DateGoIN")) = "Recived Date"
        .TextMatrix(0, .ColIndex("StoreName")) = "Store Name"
      
    End With

    With Me.FgRequired
        .TextMatrix(0, .ColIndex("TableID")) = "ID"
        .TextMatrix(0, .ColIndex("AlramDate")) = "Alram Date"
        .TextMatrix(0, .ColIndex("AlramPriorityType")) = "Alram Priority "
        .TextMatrix(0, .ColIndex("Transaction_serial")) = "Transaction Serial"
        .TextMatrix(0, .ColIndex("CusName")) = "Cus. Name"
        .TextMatrix(0, .ColIndex("AlramMsg")) = "Alram Msg"
        .TextMatrix(0, .ColIndex("View")) = "View"
        
    End With

    With Me.FgCompAssblied
        .TextMatrix(0, .ColIndex("TableID")) = "ID"
        .TextMatrix(0, .ColIndex("Transaction_serial")) = "Transaction Serial"
        .TextMatrix(0, .ColIndex("CusName")) = "Cus. Name"
        .TextMatrix(0, .ColIndex("DoneDate")) = "DoneDate"
        .TextMatrix(0, .ColIndex("DoneUserID")) = "DoneUserID"
        .TextMatrix(0, .ColIndex("AlramMsg")) = "AlramMsg"
        .TextMatrix(0, .ColIndex("View")) = "View"
    End With

    With Me.FgReady
        .TextMatrix(0, .ColIndex("ReciptNumber")) = "Recipt Number"
        .TextMatrix(0, .ColIndex("TicketNO")) = "Ticket NO"
        .TextMatrix(0, .ColIndex("ItemCode")) = "Item Code"
        .TextMatrix(0, .ColIndex("itemname")) = "Item Name"
        .TextMatrix(0, .ColIndex("Item_Serial")) = "Item_Serial"
        .TextMatrix(0, .ColIndex("Qty")) = "Quantity"
        .TextMatrix(0, .ColIndex("MType")) = "Maint. Type"
        .TextMatrix(0, .ColIndex("CusName")) = "Cus. Name"
        ' .TextMatrix(0, .ColIndex("CustomerNotes")) = "Customer complaint "
        '    .TextMatrix(0, .ColIndex("EmpNotes")) = "Employee Notes"
        .TextMatrix(0, .ColIndex("DateGoIN")) = "Recived Date"
        .TextMatrix(0, .ColIndex("StoreName")) = "Store Name"
        .TextMatrix(0, .ColIndex("Cost")) = "Maintenence Cost"
      
    End With

End Sub

Private Sub Form_Load()

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    Dim GrdBack As ClsBackGroundPic
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Print").Picture
    Set Cmd(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Refresh").Picture
    Set Cmd(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
    Set Cmd(5).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    Me.TabMain.CurrTab = 0

    With Me.FG
        .GridLines = flexGridNone
        .RowHeightMin = 345
        Set GrdBack = New ClsBackGroundPic
        .ExtendLastCol = True
        .WallPaper = GrdBack.Picture
        .AutoSize 0, .Cols - 1, False
    End With

    With Me.FgRequired
        .GridLines = flexGridNone
        .RowHeightMin = 345
        Set GrdBack = New ClsBackGroundPic
        .ExtendLastCol = True
        .WallPaper = GrdBack.Picture
        .AutoSize 0, .Cols - 1, False
    End With

    With Me.FgCompAssblied
        .GridLines = flexGridNone
        .RowHeightMin = 345
        Set GrdBack = New ClsBackGroundPic
        .ExtendLastCol = True
        .WallPaper = GrdBack.Picture
        .AutoSize 0, .Cols - 1, False
    End With

    Me.Width = 12000
    Me.Height = 9500
    Resize_Form Me
    LoadData
End Sub

Public Sub LoadManStore()
    Dim rs As ADODB.Recordset, RsTemp As ADODB.Recordset
    Dim StrSQL As String
    Dim i As Long
    Dim LngLastRow As Long
    Dim LngFindRow As Long
    Dim XNode As VSFlex8UCtl.VSFlexNode
    Dim StrTemp  As String
    Dim Msg As String
    Dim SngTemp As Single

    On Error GoTo ErrTrap

    With Me.FG
        .Redraw = flexRDNone
        .Rows = .FixedRows
        .Clear flexClearScrollable, flexClearEverything
        .ExtendLastCol = True
        .MergeCells = flexMergeFree
        .OutlineBar = flexOutlineBarComplete
        .AllowUserResizing = flexResizeColumns
        .ExtendLastCol = True
        '.NodeClosedPicture = MDIFrmMain.ImgLstMenuIcons.ListImages("NodeClose").Picture
        '.NodeOpenPicture = MDIFrmMain.ImgLstMenuIcons.ListImages("NodeOpen").Picture
        .RowHeightMin = 345
        .ScrollTrack = False
        .ScrollTips = True
        .SheetBorder = vbWhite
        StrSQL = "SELECT dbo.TblMainteneceNew.MaintananceID, dbo.TblMainteneceNew.ReciptNumber," & _
           "dbo.TblMainteneceNew.Transaction_ID,dbo.Transactions.Transaction_Serial," & _
           "dbo.TblMainteneceNew.CusID, dbo.TblCustemers.CusName,dbo.TblCustemers.CusNamee, dbo.TblMainteneceNew.CashCustomerName," & _
           "dbo.TblMainteneceNew.CashCust" & _
           "omerPhone, dbo.TblMainteneceNew.CashCustomerMobile, dbo.TblMainteneceNew.CashCus" & _
           "tomerEmail,dbo.TblMainteneceNew.CashCustomerAddress, dbo" & _
           ".TblMainteneceNew.DateGoIN, dbo.TblMainteneceNew.DateGoOUT, dbo.TblMainteneceNew" & _
           ".EmpID,dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Nam" & _
           "e, dbo.TblMainteneceNew.StoreID, dbo.TblStore.StoreName,dbo.TblStore.StoreNamee,                          " & _
           "dbo.TblMainteneceNew.Remarks, dbo.TblMainteneceNew.GoOut, dbo.TblMainteneceNew.U" & _
           "serID, dbo.TblUsers.UserName,dbo.TblMainteneceNew.Paymen" & _
           "tType, dbo.TblMainteneceNew.MType, dbo.TblMainteneceNew.ManOperationTypeID,     " & _
           "dbo.TblManOperations.ManOperationTypeName,.TblManOperations.ManOperationTypeNamee, dbo.TblMainteneceN" & _
           "ew.ItemID, dbo.TblItems.ItemCode, dbo.TblItems.ItemName,   dbo.TblItems.ItemNamee,                      " & _
           "dbo.TblMainteneceNew.ItemSerial, dbo.TblMainteneceNew.Quantity, dbo.TblMaintenec" & _
           "eNew.TicketNO, dbo.TblMainteneceNew.CustomerNotes,                        dbo.Tb" & _
           "lMainteneceNew.EmpNotes FROM         dbo.Transactions RIGHT OUTER JOIN          " & _
           "             dbo.TblUsers RIGHT OUTER JOIN                       dbo.TblManOpera" & _
           "tions INNER JOIN                       dbo.TblMainteneceNew INNER JOIN          " & _
           "             dbo.TblCustemers ON dbo.TblMainteneceNew.CusID = dbo.TblCustemers.C" & _
           "usID INNER JOIN                       dbo.TblItems ON dbo.TblMainteneceNew.ItemI" & _
           "D = dbo.TblItems.ItemID INNER JOIN                       dbo.TblStore ON dbo.Tbl" & _
           "MainteneceNew.StoreID = dbo.TblStore.StoreID INNER JOIN                       db" & _
           "o.TblEmployee ON dbo.TblMainteneceNew.EmpID = dbo.TblEmployee.Emp_ID ON  " & _
               "dbo.TblManOperations.ManOperationTypeID = dbo.TblMainteneceNew.ManOperationTypeID ON"
        StrSQL = StrSQL + "      dbo.TblUsers.UserID = dbo.TblMainteneceNew.UserID ON dbo.Transactions.Transaction_ID =" & "dbo.TblMainteneceNew.Transaction_ID"
        StrSQL = StrSQL + " Where dbo.TblMainteneceNew.ManOperationTypeID =1" 'œŒÊ· ··’Ì«‰…    '
        StrSQL = StrSQL + " AND dbo.TblMainteneceNew.TicketNO IN (Select TicketNO From dbo.QryManStockComplete(0)QryManStockComplete)"
        StrSQL = StrSQL + " Order by dbo.TblMainteneceNew.MaintananceID DESC"
        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        For i = 1 To rs.RecordCount
            .Rows = .Rows + 1
            LngLastRow = .Rows - 1
            .IsSubtotal(LngLastRow) = True
            .TextMatrix(LngLastRow, .ColIndex("MaintananceID")) = IIf(IsNull(rs("MaintananceID").value), "", rs("MaintananceID").value)
            .TextMatrix(LngLastRow, .ColIndex("ReciptNumber")) = IIf(IsNull(rs("ReciptNumber").value), "", rs("ReciptNumber").value)
            .TextMatrix(LngLastRow, .ColIndex("TicketNO")) = IIf(IsNull(rs("TicketNO").value), "", rs("TicketNO").value)
            .TextMatrix(LngLastRow, .ColIndex("ItemID")) = IIf(IsNull(rs("ItemID").value), "", rs("ItemID").value)
            .TextMatrix(LngLastRow, .ColIndex("ItemCode")) = IIf(IsNull(rs("ItemCode").value), "", rs("ItemCode").value)

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(LngLastRow, .ColIndex("ItemName")) = IIf(IsNull(rs("ItemName").value), "", rs("ItemName").value)
            Else
                .TextMatrix(LngLastRow, .ColIndex("ItemName")) = IIf(IsNull(rs("ItemNamee").value), "", rs("ItemNamee").value)
            End If

            .TextMatrix(LngLastRow, .ColIndex("Item_Serial")) = IIf(IsNull(rs("ItemSerial").value), "", rs("ItemSerial").value)
            .TextMatrix(LngLastRow, .ColIndex("Qty")) = IIf(IsNull(rs("Quantity").value), "", rs("Quantity").value)

            If Not IsNull(rs("MType").value) Then
            
                If rs("MType").value = 0 Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        .TextMatrix(LngLastRow, .ColIndex("MType")) = "œ«Œ· «·÷„«‰"
                    Else
                        .TextMatrix(LngLastRow, .ColIndex("MType")) = "Local Gurantee"
                    End If

                ElseIf rs("MType").value = 1 Then

                    If SystemOptions.UserInterface = ArabicInterface Then
                        .TextMatrix(LngLastRow, .ColIndex("MType")) = "Œ«—Ã «·÷„«‰"
                    Else
                        .TextMatrix(LngLastRow, .ColIndex("MType")) = "Out Side"
                    End If
                End If
        
            End If

            .TextMatrix(LngLastRow, .ColIndex("CustomerNotes")) = IIf(IsNull(rs("CustomerNotes").value), "", rs("CustomerNotes").value)
            .TextMatrix(LngLastRow, .ColIndex("EmpNotes")) = IIf(IsNull(rs("EmpNotes").value), "", rs("EmpNotes").value)

            If Not IsNull(rs("DateGoIN").value) Then
                .TextMatrix(LngLastRow, .ColIndex("DateGoIN")) = DisplayDate(rs("DateGoIN").value)
            End If

            If rs("CusID").value = 2 Then

                '⁄„Ì· ‰ﬁœÌ
                If Not IsNull(rs("CashCustomerName").value) Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        StrTemp = rs("CusName").value & IIf(Len(rs("CashCustomerName").value) > 0, " -- " & rs("CashCustomerName").value, "")
                    Else
                        StrTemp = rs("CusNamee").value & IIf(Len(rs("CashCustomerName").value) > 0, " -- " & rs("CashCustomerName").value, "")
                    End If

                    .TextMatrix(LngLastRow, .ColIndex("CusName")) = StrTemp
                Else

                    If SystemOptions.UserInterface = ArabicInterface Then
                        .TextMatrix(LngLastRow, .ColIndex("CusName")) = IIf(IsNull(rs("CusName").value), "", rs("CusName").value)
                    Else
                        .TextMatrix(LngLastRow, .ColIndex("CusName")) = IIf(IsNull(rs("CusNamee").value), "", rs("CusNamee").value)
                    End If
                End If
            
            Else

                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(LngLastRow, .ColIndex("CusName")) = IIf(IsNull(rs("CusName").value), "", rs("CusName").value)
                Else
                    .TextMatrix(LngLastRow, .ColIndex("CusName")) = IIf(IsNull(rs("CusNamee").value), "", rs("CusNamee").value)
                End If
            End If

            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(LngLastRow, .ColIndex("StoreName")) = IIf(IsNull(rs("StoreName").value), "", rs("StoreName").value)
            Else
                .TextMatrix(LngLastRow, .ColIndex("StoreName")) = IIf(IsNull(rs("StoreNamee").value), "", rs("StoreNamee").value)
            End If

            '---------------------------------------------
            .Cell(flexcpFontBold, LngLastRow, 0, LngLastRow, .Cols - 1) = True
            '---------------------------------------------
            rs.MoveNext
        Next i

        '-------------------------------------------------------------------------
        '⁄„·Ì«  «·≈” »œ«· «·›Ê—Ï
        rs.Close
        StrSQL = "SELECT dbo.TblMainteneceNew.MaintananceID, dbo.TblItems.ItemCode, dbo.TblItems.ItemName," & "dbo.TblMainteneceNew.ReItemID,dbo.TblMainteneceNew.ReItemSerial , dbo.TblMainteneceNew.ReItemQuantity," & "dbo.TblMainteneceNew.ReItemStore, dbo.TblStore.StoreName"
        StrSQL = StrSQL + " FROM         dbo.TblMainteneceNew INNER JOIN dbo.TblItems ON " & "dbo.TblMainteneceNew.ReItemID = dbo.TblItems.ItemID INNER JOIN dbo.TblStore ON " & "dbo.TblMainteneceNew.ReItemStore = dbo.TblStore.StoreID"
        StrSQL = StrSQL + " Where dbo.TblMainteneceNew.ManOperationTypeID =1" 'œŒÊ· ··’Ì«‰…
        StrSQL = StrSQL + " AND dbo.TblMainteneceNew.ReItemID IS NOT NULL "
        StrSQL = StrSQL + " Order by dbo.TblMainteneceNew.MaintananceID DESC"
        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (rs.BOF Or rs.EOF) Then
            rs.MoveFirst

            For i = 1 To rs.RecordCount
                LngFindRow = .FindRow(rs("MaintananceID").value, .FixedRows, .ColIndex("MaintananceID"), False, True)

                If LngFindRow <> -1 Then
                    .AddItem "", (LngFindRow + 1)
                    LngLastRow = (LngFindRow + 1)
                    .RowOutlineLevel(LngLastRow) = .RowOutlineLevel(LngFindRow) + 1
                    .MergeRow(LngLastRow) = True
                    .TextMatrix(LngLastRow, .ColIndex("MaintananceID")) = IIf(IsNull(rs("MaintananceID").value), "", rs("MaintananceID").value)
                    .Cell(flexcpText, LngLastRow, 2, LngLastRow, 3) = "≈” »œ«· ›Ê—Ì"
                    '‰⁄—÷ ’Ê—… „„Ì“… ·⁄„·Ì… «·≈” »œ«·
                    .Cell(flexcpPicture, LngLastRow, 0, LngLastRow, 0) = mdifrmmain.ImgLstMenuIcons.ListImages("Execute").ExtractIcon
                    .Cell(flexcpPictureAlignment, LngLastRow, 0, LngLastRow, 0) = flexPicAlignRightCenter
                        
                    .TextMatrix(LngLastRow, .ColIndex("ItemID")) = IIf(IsNull(rs("ReItemID").value), "", rs("ReItemID").value)
                    .TextMatrix(LngLastRow, .ColIndex("ItemCode")) = IIf(IsNull(rs("ItemCode").value), "", rs("ItemCode").value)
                    .TextMatrix(LngLastRow, .ColIndex("ItemName")) = IIf(IsNull(rs("ItemName").value), "", rs("ItemName").value)
                    .TextMatrix(LngLastRow, .ColIndex("Item_Serial")) = IIf(IsNull(rs("ReItemSerial").value), "", rs("ReItemSerial").value)
                    .TextMatrix(LngLastRow, .ColIndex("Qty")) = IIf(IsNull(rs("ReItemQuantity").value), "", rs("ReItemQuantity").value)
                    .TextMatrix(LngLastRow, .ColIndex("StoreName")) = IIf(IsNull(rs("StoreName").value), "", rs("StoreName").value)
                    '--------------
                    .Cell(flexcpFontName, LngLastRow, 0, LngLastRow, .Cols - 1) = "Tahoma"
                    .Cell(flexcpForeColor, LngLastRow, 0, LngLastRow, .Cols - 1) = &HC0&
                    .Cell(flexcpFontBold, LngLastRow, 0, LngLastRow, .Cols - 1) = False
                    '-------------
                End If

                rs.MoveNext
            Next i

        End If

        rs.Close
        '------------⁄—÷  ﬁ«—Ì— «·„ «»⁄… «·Œ«’… »«·’Ì«‰…
        StrSQL = " SELECT TOP 100 PERCENT dbo.TblMainteneceNew.MaintananceID, dbo.TblMainteneceNew.CusID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee," & "dbo.TblMainteneceNew.DateGoIN, dbo.TblMainteneceNew.DateGoOUT, dbo.TblMainteneceNew.EmpID, dbo.TblEmployee.Emp_Name," & "dbo.TblMainteneceNew.RetrunOrgID, dbo.TblManOperations.ManOperationTypeName,dbo.TblManOperations.ManOperationTypeNamee, dbo.TblManOperations.ManOperationTypeID," & "dbo.TblMainteneceNew.SupDeci ,dbo.TblMainteneceNew.Cost,dbo.TblManSupDecs.SupDecName,dbo.TblManSupDecs.SupDecNamee "
        StrSQL = StrSQL + " FROM         dbo.TblMainteneceNew LEFT OUTER JOIN dbo.TblCustemers ON dbo.TblMainteneceNew.CusID =" & "dbo.TblCustemers.CusID INNER JOIN dbo.TblManOperations ON dbo.TblMainteneceNew.ManOperationTypeID = " & "dbo.TblManOperations.ManOperationTypeID INNER JOIN dbo.TblEmployee ON dbo.TblMainteneceNew.EmpID = " & "dbo.TblEmployee.Emp_ID LEFT OUTER JOIN dbo.TblManSupDecs ON dbo.TblMainteneceNew.SupDeci = dbo.TblManSupDecs.SupDecID"
        StrSQL = StrSQL + " Where (dbo.TblMainteneceNew.RetrunOrgID Is Not Null)"
        StrSQL = StrSQL + " ORDER BY dbo.TblMainteneceNew.RetrunOrgID, dbo.TblMainteneceNew.MaintananceID"
        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (rs.BOF Or rs.EOF) Then
            rs.MoveFirst

            For i = 1 To rs.RecordCount
                LngFindRow = .FindRow(rs("RetrunOrgID").value, .FixedRows, .ColIndex("MaintananceID"), False, True)

                If LngFindRow <> -1 Then
                    If .IsSubtotal(LngFindRow) = True Then
                        Set XNode = .GetNode(LngFindRow)
                        LngLastRow = LngFindRow + ModFgLib.GetNodeChildTotal(FG, XNode, flexSTCount) + 1
                        .AddItem "", LngLastRow
                        .MergeRow(LngLastRow) = True
                        .RowOutlineLevel(LngLastRow) = .RowOutlineLevel(LngFindRow) + 1
                        ' ÕœÌœ ·Ê‰ «·”ÿ—
                        .Cell(flexcpForeColor, LngLastRow, 0, LngLastRow, .Cols - 1) = &HC0&
                        .TextMatrix(LngLastRow, .ColIndex("MaintananceID")) = IIf(IsNull(rs("MaintananceID").value), "", rs("MaintananceID").value)

                        If SystemOptions.UserInterface = ArabicInterface Then
                            .Cell(flexcpText, LngLastRow, 2, LngLastRow, 4) = rs("ManOperationTypeName").value
                        Else
                            .Cell(flexcpText, LngLastRow, 2, LngLastRow, 4) = rs("ManOperationTypeNamee").value
                        End If

                        If rs("ManOperationTypeID").value = 2 Or rs("ManOperationTypeID").value = 5 Then

                            '›Ï Õ«·… Œ—ÊÃ ÷„«‰ ··„Ê—œ «Ê Œ—ÊÃ ·‘—ﬂ… ’Ì«‰… Œ«—ÃÌ…
                            '›«‰‰« ‰⁄—÷ «”„ «·„Ê—œ  «Ê «”„ ‘—ﬂ… «·’Ì«‰… «·Œ«—ÃÌ…
                            If SystemOptions.UserInterface = ArabicInterface Then
                                .Cell(flexcpText, LngLastRow, 5, LngLastRow, 6) = rs("CusName").value
                            Else
                                .Cell(flexcpText, LngLastRow, 5, LngLastRow, 6) = rs("CusNamee").value
                            End If

                            '‰⁄—÷ ’Ê—… „„Ì“… ·⁄„·Ì… «·Œ—ÊÃ
                            .Cell(flexcpPicture, LngLastRow, 0, LngLastRow, 0) = mdifrmmain.ImgLstMenuIcons.ListImages("TabRight").ExtractIcon
                            .Cell(flexcpPictureAlignment, LngLastRow, 0, LngLastRow, 0) = flexPicAlignRightCenter

                            '‰⁄—÷  «—ÌŒ «·Œ—ÊÃ Ê «—ÌŒ «·—ÃÊ⁄ «·„ Êﬁ⁄ „‰ ⁄‰œ «·„Ê—œ «Ê ‘—ﬂ… «·’Ì«‰… «·Œ«—ÃÌ…
                            If SystemOptions.UserInterface = ArabicInterface Then
                                StrTemp = " «—ÌŒ «·Œ—ÊÃ " & DisplayDate(rs("DateGoIN").value) & "     " & " «—ÌŒ «·—ÃÊ⁄ «·„ Êﬁ⁄ " & DisplayDate(rs("DateGoOUT").value)
                            Else
                                StrTemp = "Out Date " & DisplayDate(rs("DateGoIN").value) & "     " & "Expected Return Date " & DisplayDate(rs("DateGoOUT").value)

                            End If

                            .Cell(flexcpText, LngLastRow, 7, LngLastRow, .Cols - 1) = StrTemp
                            .Cell(flexcpFontUnderline, LngLastRow, 7, LngLastRow, .Cols - 1) = True

                            'Ê÷⁄ ’Ê—…  ‰»ÌÂ ›Ï Õ«·…  √ŒÌ— «·’‰› ⁄‰œ ‘—ﬂ… «·’Ì«‰… «Ê «·„Ê—œ
                            If DateDiff("d", Date, rs("DateGoOUT").value) < 0 Then
                                .Cell(flexcpPicture, LngLastRow, 7, LngLastRow, 7) = mdifrmmain.ImgLstMenuIcons.ListImages("Warn").Picture
                                .Cell(flexcpPictureAlignment, LngLastRow, 7, LngLastRow, 7) = flexPicAlignRightCenter
                            End If

                        ElseIf rs("ManOperationTypeID").value = 3 Or rs("ManOperationTypeID").value = 6 Then

                            '›Ï Õ«·… «·—ÃÊ⁄ „‰  ÷„«‰ „Ê—œ «Ê «·—ÃÊ⁄ „‰ ‘—ﬂ… «·’Ì«‰… «·Œ«—ÃÌ…
                            '›«‰‰« ‰⁄—÷ ﬁ—«— «·„Ê—œ «Ê ‘—ﬂ… «·’Ì«‰… «·Œ«—ÃÌ…
                            If SystemOptions.UserInterface = ArabicInterface Then
                                .Cell(flexcpText, LngLastRow, 5, LngLastRow, 6) = rs("SupDecName").value
                            Else
                                .Cell(flexcpText, LngLastRow, 5, LngLastRow, 6) = rs("SupDecNamee").value
                            End If

                            '‰⁄—÷ ’Ê—… „„Ì“… ·⁄„·Ì… «·—ÃÊ⁄
                            .Cell(flexcpPicture, LngLastRow, 0, LngLastRow, 0) = mdifrmmain.ImgLstMenuIcons.ListImages("TabLeft").ExtractIcon
                            .Cell(flexcpPictureAlignment, LngLastRow, 0, LngLastRow, 0) = flexPicAlignRightCenter

                            '‰⁄—÷  «—ÌŒ «·—ÃÊ⁄ „‰ ⁄‰œ «·„Ê—œ «Ê ‘—ﬂ… «·’Ì«‰… «·Œ«—ÃÌ…
                            If SystemOptions.UserInterface = ArabicInterface Then
                                StrTemp = " «—ÌŒ «·—ÃÊ⁄ " & DisplayDate(rs("DateGoIN").value) ' & "     " & "„œ… «·»ﬁ«¡ " & DisplayDate(Rs("DateGoOUT").Value)
                            Else
                                StrTemp = "Return Date   " & DisplayDate(rs("DateGoIN").value) ' & "     " & "„œ… «·»ﬁ«¡ " & DisplayDate(Rs("DateGoOUT").Value)
                            End If

                            .Cell(flexcpText, LngLastRow, 7, LngLastRow, .Cols - 1) = StrTemp
                            .Cell(flexcpFontUnderline, LngLastRow, 7, LngLastRow, .Cols - 1) = True

                            '---------⁄—÷ „⁄·Ê„«  „„Ì“… ·ﬁ—«— «·„Ê—œ «Ê ‘—ﬂ… «·’Ì«‰…
                            If Not IsNull(rs("SupDeci").value) Then
                                If rs("SupDeci").value = 8 Then
                                    '---------›Ï Õ«·… «‰ ÌﬂÊ‰ «·„Ê—œ ﬁœ ﬁ«„ »≈” »œ«· «·’‰› ›«‰‰« ‰⁄—÷  »Ì«‰«  «·’‰› «·„” »œ·
                                    AddReplacedItem rs("MaintananceID").value, LngLastRow
                                ElseIf rs("SupDeci").value = 12 Then

                                    ' Œ’Ì„ ⁄·Ï «·„Ê—œ
                                    '‰⁄—÷ «·ﬁ—«— »«·≈÷«›… ≈·Ï ﬁÌ„… «· Œ’Ì„
                                    If SystemOptions.UserInterface = ArabicInterface Then
                                        .Cell(flexcpText, LngLastRow, 5, LngLastRow, 6) = rs("SupDecName").value & "(ﬁÌ„… «· Œ’Ì„ : " & rs("Cost").value & ")"
                                    Else
                                        .Cell(flexcpText, LngLastRow, 5, LngLastRow, 6) = rs("SupDecNamee").value & "(Discount Value   : " & rs("Cost").value & ")"
                                    End If

                                ElseIf rs("SupDeci").value = 13 Then
                                    ' „ «· ’·ÌÕ ·Â–« ‰⁄—÷ ﬁÌ„… «· ’·ÌÕ Ê‰⁄—÷ «·Œ«‰… »«··Ê‰ «·√Œ÷—
                                    SngTemp = IIf(IsNull(rs("Cost").value), 0, rs("Cost").value)
                                    SngTemp = val(Format(SngTemp, SystemOptions.SysDefCurrencyForamt))

                                    If SystemOptions.UserInterface = ArabicInterface Then
                                        .Cell(flexcpText, LngLastRow, 5, LngLastRow, 6) = rs("SupDecName").value & "(ﬁÌ„… «· ’·ÌÕ : " & SngTemp & ")"
                                    Else
                                        .Cell(flexcpText, LngLastRow, 5, LngLastRow, 6) = rs("SupDecNamee").value & "( Repair Value :" & SngTemp & ")"
                                    End If
                               
                                    .Cell(flexcpForeColor, LngLastRow, 5, LngLastRow, 6) = &H8000&
                                    SngTemp = .GetNodeRow(LngLastRow, flexNTParent)
                                    .Cell(flexcpForeColor, SngTemp, 0, SngTemp, .Cols - 1) = &HC000&
                                End If
                            End If

                            '-----------------------------------------------------------------------
                        ElseIf rs("ManOperationTypeID").value = 7 Then

                            '›Ï Õ«·… „ «»⁄… „ÊŸ› «·’Ì«‰…
                            '›«‰‰« ‰⁄—÷ ﬁ—«— „ÊŸ› «·’Ì«‰…
                            If SystemOptions.UserInterface = ArabicInterface Then
                                .Cell(flexcpText, LngLastRow, 5, LngLastRow, 6) = rs("SupDecName").value
                            Else
                                .Cell(flexcpText, LngLastRow, 5, LngLastRow, 6) = rs("SupDecNamee").value
                            End If

                            If Not IsNull(rs("Cost").value) Then
                                If SystemOptions.UserInterface = ArabicInterface Then
                                    StrTemp = rs("SupDecName").value
                                Else
                                    StrTemp = rs("SupDecNamee").value
                                End If

                                StrTemp = StrTemp & "(" & rs("Cost").value & ")"
                                .Cell(flexcpText, LngLastRow, 5, LngLastRow, 6) = StrTemp
                            End If

                        ElseIf rs("ManOperationTypeID").value = 4 Then

                            '›Ï Õ«·…  ”·Ì„ «·⁄„Ì·
                            '›«‰‰« ‰⁄—÷ ﬁ—«— „ÊŸ› «·’Ì«‰…
                            If SystemOptions.UserInterface = ArabicInterface Then
                                .Cell(flexcpText, LngLastRow, 5, LngLastRow, 6) = rs("SupDecName").value
                            Else
                                .Cell(flexcpText, LngLastRow, 5, LngLastRow, 6) = rs("SupDecNamee").value
                            End If

                            If Not IsNull(rs("Cost").value) Then
                                If SystemOptions.UserInterface = ArabicInterface Then
                                    StrTemp = rs("SupDecName").value
                                Else
                                    StrTemp = rs("SupDecNamee").value
                                End If

                                StrTemp = StrTemp & "(" & rs("Cost").value & ")"
                                .Cell(flexcpText, LngLastRow, 5, LngLastRow, 6) = StrTemp
                            End If
                        End If

                        If .Cell(flexcpForeColor, LngLastRow, 5, LngLastRow, 6) <> &H8000& Then
                            .Cell(flexcpForeColor, LngLastRow, 5, LngLastRow, 6) = vbBlue
                        End If

                        .Cell(flexcpFontBold, LngLastRow, 0, LngLastRow, .Cols - 1) = False
                        .Cell(flexcpFontName, LngLastRow, 0, LngLastRow, .Cols - 1) = "Tahoma"
                    End If
                End If

                rs.MoveNext
            Next i

        End If

        '-------------------------------------------------------------------------
        For i = .FixedRows To .Rows - 1

            If .IsSubtotal(i) = True Then
                .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = &HC0C0C0
            Else
                .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = &HE2E9E9
            End If

        Next i

        .AutoSize 0, .Cols - 1, False
        .Redraw = flexRDDirect
    End With

    Exit Sub
ErrTrap:
    FG.Redraw = flexRDDirect
    Msg = "ÕœÀ Œÿ« √À‰«¡  Õ„Ì· «·»Ì«‰« ...!!!"
    Msg = Msg & Chr(13) & Err.description
    Msg = Msg & Chr(13) & Err.Number
    Msg = Msg & Chr(13) & Err.Source
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
End Sub

Private Sub LoadManOut()
    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    Dim i As Long
    Dim LngLastRow As Long
    Dim LngFindRow As Long
    Dim XNode As VSFlex8UCtl.VSFlexNode
    Dim StrTemp  As String
    Dim GrdBack As ClsBackGroundPic
    Dim LngOldCusID As Long

    With Me.FgOutMan
        .Rows = .FixedRows
        .Clear flexClearScrollable, flexClearEverything
        .GridLines = flexGridNone
        .RowHeightMin = 345
        Set GrdBack = New ClsBackGroundPic
        .ExtendLastCol = True
        .WallPaper = GrdBack.Picture
        .MergeCells = flexMergeFree
        .OutlineBar = flexOutlineBarComplete
        .AllowUserResizing = flexResizeColumns
        .ScrollTrack = False
        .ScrollTips = True
        .SheetBorder = vbWhite
        .AutoSize 0, .Cols - 1, False
    End With

    StrSQL = "SELECT TOP 100 PERCENT QryManSupStockComplete.QTY, QryManSupStockComplete.ItemID, QryManSupStockComplete.ItemCode, "
    StrSQL = StrSQL + " QryManSupStockComplete.ItemName, QryManSupStockComplete.ItemSerial, QryManSupStockComplete.TicketNO,"
    StrSQL = StrSQL + " QryManSupStockComplete.HaveSerial , QryManSupStockComplete.CusID, dbo.TblCustemers.CusName, dbo.TblCustemers.Type"
    StrSQL = StrSQL + " FROM dbo.QryManSupStockComplete(0) QryManSupStockComplete INNER JOIN "
    StrSQL = StrSQL + " dbo.TblCustemers ON QryManSupStockComplete.CusID = dbo.TblCustemers.CusID "
    StrSQL = StrSQL + " ORDER BY QryManSupStockComplete.CusID"
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If

    rs.MoveFirst

    With Me.FgOutMan

        Do While Not rs.EOF

            If LngOldCusID <> rs("CusID").value Then
                LngOldCusID = rs("CusID").value
                'New Cus
                .AddItem ""
                LngLastRow = .Rows - 1
                .Cell(flexcpText, LngLastRow, .ColIndex("TicketNO"), LngLastRow, .ColIndex("ItemID")) = rs("CusName").value
                .Cell(flexcpFontBold, LngLastRow, .ColIndex("TicketNO"), LngLastRow, .Cols - 1) = True
                .IsSubtotal(LngLastRow) = True
                .MergeRow(LngLastRow) = True
            End If

            .AddItem ""
            LngLastRow = .Rows - 1
            .IsSubtotal(LngLastRow) = False
            .RowOutlineLevel(LngLastRow) = 1
            .TextMatrix(LngLastRow, .ColIndex("TicketNO")) = IIf(IsNull(rs("TicketNO").value), "", rs("TicketNO").value)
            .TextMatrix(LngLastRow, .ColIndex("ItemID")) = IIf(IsNull(rs("ItemID").value), "", rs("ItemID").value)
            .TextMatrix(LngLastRow, .ColIndex("ItemCode")) = IIf(IsNull(rs("ItemCode").value), "", rs("ItemCode").value)
            .TextMatrix(LngLastRow, .ColIndex("ItemName")) = IIf(IsNull(rs("ItemName").value), "", rs("ItemName").value)
            .TextMatrix(LngLastRow, .ColIndex("ItemSerial")) = IIf(IsNull(rs("ItemSerial").value), "", rs("ItemSerial").value)
            .Cell(flexcpFontName, LngLastRow, 0, LngLastRow, .Cols - 1) = "Tahoma"
            .Cell(flexcpForeColor, LngLastRow, 0, LngLastRow, .Cols - 1) = &HC0&
            rs.MoveNext
        Loop

        .AutoSize 0, .Cols - 1, False
    End With

    rs.Close
    Set rs = Nothing
End Sub

Private Sub LoadManAlrams()
    Dim rs As ADODB.Recordset, RsItems As ADODB.Recordset
    Dim StrSQL As String
    Dim i As Long
    Dim LngLastRow As Long
    Dim LngFindRow As Long
    Dim XNode As VSFlex8UCtl.VSFlexNode
    Dim StrTemp  As String
    Dim GrdBack As ClsBackGroundPic
    Dim LngTransID As Long

    With Me.FgRequired
        .Rows = .FixedRows
        .Clear flexClearScrollable, flexClearEverything
        .GridLines = flexGridNone
        .RowHeightMin = 345
        .ExtendLastCol = True
        .MergeCells = flexMergeFree
        .OutlineBar = flexOutlineBarComplete
        .AllowUserResizing = flexResizeColumns
        .ScrollTrack = False
        .ScrollTips = True
        .SheetBorder = vbWhite
        .AutoSize 0, .Cols - 1, False
    End With

    StrSQL = "SELECT dbo.TblManAlram.TableID, dbo.TblManAlram.TransID,dbo.Transactions.Transaction_Serial," & "dbo.TblCustemers.CusName, dbo.TblManAlram.AlramDate,dbo.TblManAlram.AlramPriority,dbo.TblManAlram.AlramMsg," & "dbo.TblManAlram.UserID, dbo.TblUsers.UserName, dbo.TblManAlram.State "
    StrSQL = StrSQL + " FROM dbo.Transactions INNER JOIN dbo.TblCustemers ON dbo.Transactions.CusID = " & "dbo.TblCustemers.CusID INNER JOIN dbo.TblManAlram ON dbo.Transactions.Transaction_ID =" & "dbo.TblManAlram.TransID INNER JOIN dbo.TblUsers ON dbo.TblManAlram.UserID = dbo.TblUsers.UserID "
    StrSQL = StrSQL + " Where dbo.TblManAlram.State=0 "

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If

    rs.MoveFirst

    With Me.FgRequired

        For i = 1 To rs.RecordCount
            .AddItem ""
            LngLastRow = .Rows - 1
            .IsSubtotal(LngLastRow) = True
            .RowOutlineLevel(LngLastRow) = 0
            LngTransID = rs("TransID").value
            .TextMatrix(LngLastRow, .ColIndex("TableID")) = IIf(IsNull(rs("TableID").value), "", rs("TableID").value)
            .TextMatrix(LngLastRow, .ColIndex("TransID")) = IIf(IsNull(rs("TransID").value), "", rs("TransID").value)
            .TextMatrix(LngLastRow, .ColIndex("Transaction_Serial")) = IIf(IsNull(rs("Transaction_Serial").value), "", rs("Transaction_Serial").value)
            .TextMatrix(LngLastRow, .ColIndex("CusName")) = IIf(IsNull(rs("CusName").value), "", rs("CusName").value)
            .TextMatrix(LngLastRow, .ColIndex("AlramMsg")) = IIf(IsNull(rs("AlramMsg").value), "", rs("AlramMsg").value)

            If rs("AlramPriority").value = 1 Then
                .TextMatrix(LngLastRow, .ColIndex("AlramPriorityType")) = "÷⁄Ì›"
            ElseIf rs("AlramPriority").value = 2 Then
                .TextMatrix(LngLastRow, .ColIndex("AlramPriorityType")) = "⁄«œÏ"
            ElseIf rs("AlramPriority").value = 3 Then
                .TextMatrix(LngLastRow, .ColIndex("AlramPriorityType")) = "Â«„ Ãœ«"
            End If

            If Not IsNull(rs("AlramDate").value) Then
                .TextMatrix(LngLastRow, .ColIndex("AlramDate")) = DisplayDate(rs("AlramDate").value)
            End If

            .TextMatrix(LngLastRow, .ColIndex("View")) = "„‘«Âœ… «·›« Ê—…"
            .Cell(flexcpForeColor, LngLastRow, .ColIndex("View"), LngLastRow, .ColIndex("View")) = vbBlue
            '-------------------------------------------------------------
            StrSQL = "SELECT     dbo.Transaction_Details.Transaction_ID, dbo.Transaction_Details.Item_ID," & "dbo.TblItems.ItemCode , dbo.TblItems.ItemName,dbo.Transaction_Details.ItemCase," & "dbo.Transaction_Details.ItemSerial, dbo.Transaction_Details.Quantity "
            StrSQL = StrSQL + " FROM         dbo.Transaction_Details INNER JOIN "
            StrSQL = StrSQL + " dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID "
            StrSQL = StrSQL + " Where  dbo.Transaction_Details.Transaction_ID=" & LngTransID
            Set RsItems = New ADODB.Recordset
            RsItems.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

            If Not (RsItems.BOF Or RsItems.EOF) Then
                RsItems.MoveFirst

                Do While Not RsItems.EOF
                    .AddItem ""
                    LngLastRow = .Rows - 1
                    .RowOutlineLevel(LngLastRow) = 1
                    .TextMatrix(LngLastRow, .ColIndex("AlramPriorityType")) = IIf(IsNull(RsItems("ItemCode").value), "", RsItems("ItemCode").value)
                    .TextMatrix(LngLastRow, .ColIndex("Transaction_Serial")) = IIf(IsNull(RsItems("ItemName").value), "", RsItems("ItemName").value)
                    .TextMatrix(LngLastRow, .ColIndex("CusName")) = IIf(IsNull(RsItems("ItemSerial").value), "", RsItems("ItemSerial").value)
                    .TextMatrix(LngLastRow, .ColIndex("AlramMsg")) = IIf(IsNull(RsItems("Quantity").value), "", RsItems("Quantity").value)
                    RsItems.MoveNext
                    .Cell(flexcpFontName, LngLastRow, 0, LngLastRow, .Cols - 1) = "Tahoma"
                    .Cell(flexcpForeColor, LngLastRow, 0, LngLastRow, .Cols - 1) = &HC0&
                Loop

            End If

            '-------------------------------------------------------------
            rs.MoveNext
        Next i

        '--------------------------------------------------------------------------------------------
        .AutoSize 0, .Cols - 1, False
    End With

    rs.Close
    Set rs = Nothing

End Sub

Private Sub LoadManAlramsDone()
    Dim rs As ADODB.Recordset, RsItems As ADODB.Recordset
    Dim StrSQL As String
    Dim i As Long
    Dim LngLastRow As Long
    Dim LngFindRow As Long
    Dim XNode As VSFlex8UCtl.VSFlexNode
    Dim StrTemp  As String
    Dim GrdBack As ClsBackGroundPic
    Dim LngTransID As Long

    With Me.FgCompAssblied
        .Rows = .FixedRows
        .Clear flexClearScrollable, flexClearEverything
        .GridLines = flexGridNone
        .RowHeightMin = 345
        .ExtendLastCol = True
        .MergeCells = flexMergeFree
        .OutlineBar = flexOutlineBarComplete
        .AllowUserResizing = flexResizeColumns
        .ScrollTrack = False
        .ScrollTips = True
        .SheetBorder = vbWhite
        .AutoSize 0, .Cols - 1, False
    End With

    StrSQL = "SELECT dbo.TblManAlram.TableID, dbo.TblManAlram.TransID,dbo.Transactions.Transaction_Serial," & "dbo.TblCustemers.CusName, dbo.TblManAlram.AlramDate,dbo.TblManAlram.AlramPriority,dbo.TblManAlram.AlramMsg," & "dbo.TblManAlram.UserID,dbo.TblManAlram.DoneDate, dbo.TblUsers.UserName, dbo.TblManAlram.State "
    StrSQL = StrSQL + " FROM dbo.Transactions INNER JOIN dbo.TblCustemers ON dbo.Transactions.CusID = " & "dbo.TblCustemers.CusID INNER JOIN dbo.TblManAlram ON dbo.Transactions.Transaction_ID =" & "dbo.TblManAlram.TransID INNER JOIN dbo.TblUsers ON dbo.TblManAlram.DoneUserID = dbo.TblUsers.UserID "
    StrSQL = StrSQL + " Where dbo.TblManAlram.State=2 "

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If

    rs.MoveFirst

    With Me.FgCompAssblied

        For i = 1 To rs.RecordCount
            .AddItem ""
            LngLastRow = .Rows - 1
            .IsSubtotal(LngLastRow) = True
            .RowOutlineLevel(LngLastRow) = 0
            LngTransID = rs("TransID").value
            .TextMatrix(LngLastRow, .ColIndex("TableID")) = IIf(IsNull(rs("TableID").value), "", rs("TableID").value)
            .TextMatrix(LngLastRow, .ColIndex("TransID")) = IIf(IsNull(rs("TransID").value), "", rs("TransID").value)
            .TextMatrix(LngLastRow, .ColIndex("Transaction_Serial")) = IIf(IsNull(rs("Transaction_Serial").value), "", rs("Transaction_Serial").value)
            .TextMatrix(LngLastRow, .ColIndex("CusName")) = IIf(IsNull(rs("CusName").value), "", rs("CusName").value)
            .TextMatrix(LngLastRow, .ColIndex("AlramMsg")) = IIf(IsNull(rs("AlramMsg").value), "", rs("AlramMsg").value)
            .TextMatrix(LngLastRow, .ColIndex("DoneDate")) = IIf(IsNull(rs("DoneDate").value), "", rs("DoneDate").value)
            .TextMatrix(LngLastRow, .ColIndex("DoneUserID")) = IIf(IsNull(rs("AlramMsg").value), "", rs("UserName").value)
            .TextMatrix(LngLastRow, .ColIndex("View")) = "„‘«Âœ… «·›« Ê—…"
            .Cell(flexcpForeColor, LngLastRow, .ColIndex("View"), LngLastRow, .ColIndex("View")) = vbBlue
            '-------------------------------------------------------------
            rs.MoveNext
        Next i

        '--------------------------------------------------------------------------------------------
        .AutoSize 0, .Cols - 1, False
    End With

    rs.Close
    Set rs = Nothing
End Sub

Private Sub LoadManReady()
    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    Dim i As Long
    Dim LngLastRow As Long
    Dim LngFindRow As Long
    Dim XNode As VSFlex8UCtl.VSFlexNode
    Dim StrTemp  As String
    Dim GrdBack As ClsBackGroundPic
    Dim LngOldCusID As Long

    With Me.FgReady
        .Rows = .FixedRows
        .Clear flexClearScrollable, flexClearEverything
        '.GridLines = flexGridNone
        .RowHeightMin = 345
        Set GrdBack = New ClsBackGroundPic
        .ExtendLastCol = True
        .WallPaper = GrdBack.Picture
        .MergeCells = flexMergeFree
        .OutlineBar = flexOutlineBarComplete
        .AllowUserResizing = flexResizeColumns
        .ScrollTrack = False
        .ScrollTips = True
        .SheetBorder = vbWhite
        .AutoSize 0, .Cols - 1, False
    End With

    StrSQL = " SELECT     dbo.TblMainteneceNew.ReciptNumber, dbo.TblMainteneceNew.Transaction_ID, dbo.TblMainteneceNew.CusID, dbo.TblMainteneceNew.CashCustomerName, "
    StrSQL = StrSQL & " dbo.TblMainteneceNew.CashCustomerPhone, dbo.TblMainteneceNew.CashCustomerMobile, dbo.TblMainteneceNew.CashCustomerEmail,"
    StrSQL = StrSQL & "  dbo.TblMainteneceNew.CashCustomerAddress, dbo.TblMainteneceNew.OperationDate, dbo.TblMainteneceNew.DateGoIN, dbo.TblMainteneceNew.DateGoOUT,"
    StrSQL = StrSQL & " dbo.TblMainteneceNew.EmpID, dbo.TblMainteneceNew.StoreID, dbo.TblMainteneceNew.Remarks, dbo.TblMainteneceNew.GoOut, dbo.TblMainteneceNew.UserID,"
    StrSQL = StrSQL & " dbo.TblMainteneceNew.PaymentType, dbo.TblMainteneceNew.MType, dbo.TblMainteneceNew.ManOperationTypeID, dbo.TblMainteneceNew.ItemID,"
    StrSQL = StrSQL & " dbo.TblMainteneceNew.ItemSerial, dbo.TblMainteneceNew.Quantity, dbo.TblMainteneceNew.TicketNO, dbo.TblMainteneceNew.CustomerNotes,"
    StrSQL = StrSQL & " dbo.TblMainteneceNew.EmpNotes, dbo.TblMainteneceNew.Cost, dbo.TblMainteneceNew.SupDeci, dbo.TblMainteneceNew.MainOperationID,"
    StrSQL = StrSQL & " dbo.TblMainteneceNew.RetrunOrgID, dbo.TblMainteneceNew.FastReplace, dbo.TblMainteneceNew.FastReplaceType, dbo.TblMainteneceNew.ReItemID,"
    StrSQL = StrSQL & " dbo.TblMainteneceNew.ReItemSerial, dbo.TblMainteneceNew.ReItemQuantity, dbo.TblMainteneceNew.ReItemPrice, dbo.TblMainteneceNew.ReItemStore,"
    StrSQL = StrSQL & " dbo.TblItems.ItemCode, dbo.TblItems.ItemName, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblItems.ItemNamee, dbo.TblStore.StoreName,"
    StrSQL = StrSQL & " dbo.TblStore.StoreNamee"
    StrSQL = StrSQL & " FROM         dbo.TblMainteneceNew INNER JOIN"
    StrSQL = StrSQL & " dbo.TblItems ON dbo.TblMainteneceNew.ItemID = dbo.TblItems.ItemID LEFT OUTER JOIN"
    StrSQL = StrSQL & " dbo.TblCustemers ON dbo.TblMainteneceNew.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
    StrSQL = StrSQL & " dbo.TblStore ON dbo.TblMainteneceNew.StoreID = dbo.TblStore.StoreID"
    StrSQL = StrSQL & "  Where (dbo.TblMainteneceNew.ManOperationTypeID = 4) And (dbo.TblMainteneceNew.GoOut Is Null)"

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        rs.Close
        Set rs = Nothing
        Exit Sub
    End If

    rs.MoveFirst

    With Me.FgReady

        Do While Not rs.EOF

            If LngOldCusID <> rs("CusID").value Then
                LngOldCusID = rs("CusID").value
                'New Cus
                .AddItem ""
                LngLastRow = .Rows - 1
                .Cell(flexcpText, LngLastRow, .ColIndex("TicketNO"), LngLastRow, .ColIndex("ItemID")) = rs("CusName").value
                .Cell(flexcpFontBold, LngLastRow, .ColIndex("TicketNO"), LngLastRow, .Cols - 1) = True
                .IsSubtotal(LngLastRow) = True
                .MergeRow(LngLastRow) = True
            End If

            .AddItem ""
            LngLastRow = .Rows - 1
            .IsSubtotal(LngLastRow) = False
            .RowOutlineLevel(LngLastRow) = 1
            .TextMatrix(LngLastRow, .ColIndex("ReciptNumber")) = IIf(IsNull(rs("ReciptNumber").value), "", rs("ReciptNumber").value)
            .TextMatrix(LngLastRow, .ColIndex("ItemID")) = IIf(IsNull(rs("ItemID").value), "", rs("ItemID").value)
            .TextMatrix(LngLastRow, .ColIndex("ItemCode")) = IIf(IsNull(rs("ItemCode").value), "", rs("ItemCode").value)
            .TextMatrix(LngLastRow, .ColIndex("Item_Serial")) = IIf(IsNull(rs("ItemSerial").value), "", rs("ItemSerial").value)
            .TextMatrix(LngLastRow, .ColIndex("Qty")) = IIf(IsNull(rs("Quantity").value), "", rs("Quantity").value)
            .TextMatrix(LngLastRow, .ColIndex("DateGoIN")) = IIf(IsNull(rs("DateGoIN").value), "", rs("DateGoIN").value)
            
            .TextMatrix(LngLastRow, .ColIndex("Cost")) = CostForMaintenance(.TextMatrix(LngLastRow, .ColIndex("ReciptNumber")))
           
            '                     If rs("MType").value = 0 Then
            '                   If SystemOptions.UserInterface = ArabicInterface Then
            '                      .TextMatrix(LngLastRow, .ColIndex("MType")) = "œ«Œ· «·÷„«‰"
            '                  Else
            '                  .TextMatrix(LngLastRow, .ColIndex("MType")) = "Local Gurantee"
            '                  End If
            '          ElseIf rs("MType").value = 1 Then
            '                If SystemOptions.UserInterface = ArabicInterface Then
            '                  .TextMatrix(LngLastRow, .ColIndex("MType")) = "Œ«—Ã «·÷„«‰"
            '               Else
            '               .TextMatrix(LngLastRow, .ColIndex("MType")) = "Out Side"
            '               End If
            '          End If
            
            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(LngLastRow, .ColIndex("ItemName")) = IIf(IsNull(rs("ItemName").value), "", rs("ItemName").value)
                .TextMatrix(LngLastRow, .ColIndex("CusName")) = IIf(IsNull(rs("CusName").value), "", rs("CusName").value)
                .TextMatrix(LngLastRow, .ColIndex("StoreName")) = IIf(IsNull(rs("StoreName").value), "", rs("StoreName").value)
  
            Else
                .TextMatrix(LngLastRow, .ColIndex("ItemName")) = IIf(IsNull(rs("ItemNamee").value), "", rs("ItemNamee").value)
                .TextMatrix(LngLastRow, .ColIndex("CusName")) = IIf(IsNull(rs("CusNamee").value), "", rs("CusNamee").value)
                .TextMatrix(LngLastRow, .ColIndex("StoreName")) = IIf(IsNull(rs("StoreNamee").value), "", rs("StoreNamee").value)
            End If
        
            .Cell(flexcpFontName, LngLastRow, 0, LngLastRow, .Cols - 1) = "Tahoma"
            .Cell(flexcpForeColor, LngLastRow, 0, LngLastRow, .Cols - 1) = &HC0&
            rs.MoveNext
        Loop

        .AutoSize 0, .Cols - 1, False
    End With

    rs.Close
    Set rs = Nothing

End Sub

Private Sub AddReplacedItem(ManID As Long, _
                            LngRow As Long)
    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    Dim LngLastRow As Long

    Set rs = New ADODB.Recordset
    StrSQL = "SELECT dbo.TblMainteneceNew.MaintananceID, dbo.TblItems.ItemCode, dbo.TblItems.ItemName," & "dbo.TblMainteneceNew.ReItemID,dbo.TblMainteneceNew.ReItemSerial , dbo.TblMainteneceNew.ReItemQuantity," & "dbo.TblMainteneceNew.ReItemStore,dbo.TblMainteneceNew.ReItemPrice, dbo.TblStore.StoreName"
    StrSQL = StrSQL + " FROM         dbo.TblMainteneceNew INNER JOIN dbo.TblItems ON " & "dbo.TblMainteneceNew.ReItemID = dbo.TblItems.ItemID INNER JOIN dbo.TblStore ON " & "dbo.TblMainteneceNew.ReItemStore = dbo.TblStore.StoreID"
    StrSQL = StrSQL + " Where dbo.TblMainteneceNew.MaintananceID =" & ManID
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.BOF Or rs.EOF) Then

        With Me.FG
            LngLastRow = LngRow + 1
            .AddItem "", LngLastRow
            .IsSubtotal(LngLastRow) = False
            .RowOutlineLevel(LngLastRow) = .RowOutlineLevel(LngRow)
            .MergeRow(LngLastRow) = True
                
            .Cell(flexcpFontName, LngLastRow, 0, LngLastRow, .Cols - 1) = "Tahoma"
            .Cell(flexcpForeColor, LngLastRow, 0, LngLastRow, .Cols - 1) = vbBlue
            .Cell(flexcpFontBold, LngLastRow, 0, LngLastRow, .Cols - 1) = False
        
            .Cell(flexcpText, LngLastRow, 2, LngLastRow, 4) = "»Ì«‰«  «·’‰› «·„” »œ·"
            .TextMatrix(LngLastRow, .ColIndex("ItemID")) = IIf(IsNull(rs("ReItemID").value), "", rs("ReItemID").value)
            .TextMatrix(LngLastRow, .ColIndex("ItemCode")) = IIf(IsNull(rs("ItemCode").value), "", rs("ItemCode").value)
            .TextMatrix(LngLastRow, .ColIndex("ItemName")) = IIf(IsNull(rs("ItemName").value), "", rs("ItemName").value)
            .TextMatrix(LngLastRow, .ColIndex("Item_Serial")) = IIf(IsNull(rs("ReItemSerial").value), "", rs("ReItemSerial").value)
            .TextMatrix(LngLastRow, .ColIndex("Qty")) = IIf(IsNull(rs("ReItemQuantity").value), "", rs("ReItemQuantity").value)
            'MType
            'Money2
            .Cell(flexcpPicture, LngLastRow, .ColIndex("MType"), LngLastRow, .ColIndex("MType")) = mdifrmmain.ImgLstMenuIcons.ListImages("Money2").ExtractIcon
            .Cell(flexcpPictureAlignment, LngLastRow, .ColIndex("MType"), LngLastRow, .ColIndex("MType")) = flexPicAlignRightCenter
            .TextMatrix(LngLastRow, .ColIndex("MType")) = "”⁄— «·≈” »œ«· : " & IIf(IsNull(rs("ReItemPrice").value), "0", rs("ReItemPrice").value)
        
            .TextMatrix(LngLastRow, .ColIndex("StoreName")) = IIf(IsNull(rs("StoreName").value), "", rs("StoreName").value)
        End With

    End If

    rs.Close
    Set rs = Nothing
End Sub

Private Sub DelRecord()
    Dim Msg As String
    Dim IntRes As Integer
    Dim StrSQL  As String
    Dim LngManID As Long
    Dim BegineTrans As Boolean

    On Error GoTo hErr

    With Me.FG

        If .Row <= 0 Then
            Msg = "ÌÃ»  ÕœÌœ «·⁄„·Ì… «·„—«œ Õ–›Â«...!!!"
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            Exit Sub
        End If

        If .Col <= 0 Then
            Msg = "ÌÃ»  ÕœÌœ «·⁄„·Ì… «·„—«œ Õ–›Â«...!!!"
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            Exit Sub
        End If

        LngManID = val(.TextMatrix(.Row, .ColIndex("MaintananceID")))

        If LngManID = 0 Then
            Msg = "·« ÊÃœ «Ï ⁄„·Ì… ··Õ–›."
            MsgBox Msg, vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            Exit Sub
        End If

        Msg = "”Ê› Ì „ Õ–› «·⁄„·Ì… —ﬁ„ .."
        Msg = Msg & Chr(13) & LngManID
        Msg = Msg & Chr(13) & "›Â· «‰  „ «ﬂœ „‰ ⁄„·Ì… «·Œ–›...øø"
        IntRes = MsgBox(Msg, vbQuestion + vbYesNo + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title)

        If IntRes = vbYes Then
            Cn.BeginTrans
            BegineTrans = True
            StrSQL = "Delete From TblMainteneceNew Where MaintananceID=" & LngManID
            Cn.Execute StrSQL, , adExecuteNoRecords
            Cn.CommitTrans
            BegineTrans = False
            Msg = " „  ⁄„·Ì… «·Õ–›"
            MsgBox Msg, vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            LoadData
        Else
            Exit Sub
        End If

    End With

    Exit Sub
hErr:

    If BegineTrans = True Then
        Cn.RollbackTrans
        BegineTrans = False
    End If

    Msg = "ÕœÀ Œÿ√ «ıÀ‰«¡ ⁄„·Ì… «·Õ–› "
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
End Sub

Private Sub LoadData()
    LoadManStore
    LoadManOut
    LoadManAlrams
    LoadManAlramsDone
    LoadManReady
End Sub
