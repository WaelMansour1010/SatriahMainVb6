VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{CFC0A331-9521-11D5-B9E6-5A06F6000000}#1.0#0"; "VDSCombo.DLL"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmBillBuyComposite 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "›« Ê—… „‘ —Ì«  „Ã„⁄…"
   ClientHeight    =   9000
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15870
   HelpContextID   =   100
   Icon            =   "FrmBillBuyComposite.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "FrmBillBuyComposite.frx":038A
   RightToLeft     =   -1  'True
   ScaleHeight     =   9000
   ScaleWidth      =   15870
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   51
      Top             =   8625
      Width           =   15870
      _ExtentX        =   27993
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin C1SizerLibCtl.C1Elastic C1ElasticMain 
      Height          =   8625
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   15870
      _cx             =   27993
      _cy             =   15214
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
      GridRows        =   5
      GridCols        =   1
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"FrmBillBuyComposite.frx":2B2C
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin C1SizerLibCtl.C1Tab XPTab301 
         Height          =   5265
         Left            =   15
         TabIndex        =   1
         Top             =   2340
         Width           =   15840
         _cx             =   27940
         _cy             =   9287
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
         BackColor       =   14871017
         ForeColor       =   0
         FrontTabColor   =   14871017
         BackTabColor    =   12648447
         TabOutlineColor =   -2147483632
         FrontTabForeColor=   16711680
         Caption         =   "«·√’‰«›|«·√Ê—«ﬁ «·„«·Ì…|„·«ÕŸ«  ⁄·Ï «·›« Ê—…|”‰œ«  «·’—›|«·ÿ·»Ì« |›Ê« Ì— „«·Ì…|„’—Ê›«   ﬁœÌ—ÌÂ|«·„—›ﬁ« "
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
         Picture(0)      =   "FrmBillBuyComposite.frx":2B93
         Picture(1)      =   "FrmBillBuyComposite.frx":2F2D
         Begin VB.Frame Frame2 
            Height          =   4800
            Left            =   17685
            RightToLeft     =   -1  'True
            TabIndex        =   105
            Top             =   45
            Width           =   15750
            Begin VB.CommandButton Command5 
               Caption         =   " Œ’Ì’"
               Height          =   480
               Left            =   12120
               RightToLeft     =   -1  'True
               TabIndex        =   109
               Top             =   3240
               Visible         =   0   'False
               Width           =   2220
            End
            Begin VB.CommandButton Command4 
               Caption         =   "⁄—÷ «·›Ê« Ì— «·„«·Ì…"
               Height          =   480
               Left            =   7560
               RightToLeft     =   -1  'True
               TabIndex        =   108
               Top             =   2880
               Width           =   2220
            End
            Begin VB.TextBox txt_total_bill 
               Height          =   405
               Left            =   10200
               Locked          =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   106
               Top             =   2880
               Width           =   1770
            End
            Begin VSFlex8UCtl.VSFlexGrid grid4 
               Height          =   2325
               Left            =   240
               TabIndex        =   113
               Tag             =   "1"
               Top             =   480
               Width           =   14055
               _cx             =   24791
               _cy             =   4101
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
               AllowUserResizing=   0
               SelectionMode   =   0
               GridLines       =   1
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   50
               Cols            =   11
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"FrmBillBuyComposite.frx":32C7
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
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«·›Ê« Ì— «·„«·ÌÂ"
               Height          =   285
               Index           =   64
               Left            =   10800
               RightToLeft     =   -1  'True
               TabIndex        =   183
               Top             =   120
               Width           =   3120
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«Ã„«·Ì «·›Ê« Ì—"
               Height          =   285
               Index           =   59
               Left            =   12150
               RightToLeft     =   -1  'True
               TabIndex        =   107
               Top             =   3000
               Width           =   2040
            End
         End
         Begin VB.Frame Frame4 
            Height          =   4800
            Left            =   17385
            RightToLeft     =   -1  'True
            TabIndex        =   67
            Top             =   45
            Width           =   15750
            Begin VB.CommandButton Command7 
               Caption         =   "Command7"
               Height          =   195
               Left            =   1800
               RightToLeft     =   -1  'True
               TabIndex        =   115
               Top             =   1320
               Visible         =   0   'False
               Width           =   855
            End
            Begin VB.CommandButton Command3 
               Caption         =   "⁄—÷ ÿ·»«  «·‘—«¡"
               Height          =   480
               Left            =   10680
               RightToLeft     =   -1  'True
               TabIndex        =   68
               Top             =   3000
               Width           =   2010
            End
            Begin VSFlex8UCtl.VSFlexGrid GRID2 
               Height          =   2205
               Left            =   5040
               TabIndex        =   73
               Tag             =   "1"
               Top             =   600
               Width           =   7695
               _cx             =   13573
               _cy             =   3889
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
               AllowUserResizing=   0
               SelectionMode   =   0
               GridLines       =   1
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   3
               Cols            =   6
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"FrmBillBuyComposite.frx":348B
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
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "ÿ·»«  «·‘—«¡  Ê «·›Ê« Ì— «·„»œ∆ÌÂ"
               Height          =   285
               Index           =   57
               Left            =   8280
               RightToLeft     =   -1  'True
               TabIndex        =   182
               Top             =   240
               Width           =   4440
            End
         End
         Begin VB.Frame Frame1 
            Height          =   4800
            Left            =   17085
            RightToLeft     =   -1  'True
            TabIndex        =   66
            Top             =   45
            Width           =   15750
            Begin VB.CommandButton Command6 
               Caption         =   "Command6"
               Height          =   375
               Left            =   6840
               RightToLeft     =   -1  'True
               TabIndex        =   114
               Top             =   3000
               Visible         =   0   'False
               Width           =   1935
            End
            Begin VB.TextBox TXTToTAlELSHahn 
               Alignment       =   1  'Right Justify
               Height          =   405
               Left            =   1200
               Locked          =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   104
               Text            =   "0"
               Top             =   3000
               Visible         =   0   'False
               Width           =   1935
            End
            Begin VB.TextBox Txt_EXport 
               Alignment       =   1  'Right Justify
               Height          =   405
               Left            =   9600
               Locked          =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   70
               Top             =   2880
               Width           =   1890
            End
            Begin VB.CommandButton Command2 
               Caption         =   "⁄—÷ «·„’—Ê›« "
               Height          =   480
               Left            =   12000
               RightToLeft     =   -1  'True
               TabIndex        =   69
               Top             =   3240
               Visible         =   0   'False
               Width           =   2220
            End
            Begin VSFlex8UCtl.VSFlexGrid Grid 
               Height          =   2325
               Left            =   240
               TabIndex        =   71
               Tag             =   "1"
               Top             =   480
               Width           =   14055
               _cx             =   24791
               _cy             =   4101
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
               AllowUserResizing=   0
               SelectionMode   =   0
               GridLines       =   1
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   50
               Cols            =   10
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"FrmBillBuyComposite.frx":357E
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
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«Ã„«·Ì «·„’—Ê›« "
               Height          =   285
               Index           =   60
               Left            =   2880
               RightToLeft     =   -1  'True
               TabIndex        =   110
               Top             =   3000
               Visible         =   0   'False
               Width           =   1800
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "”‰œ«  «·’—›"
               Height          =   285
               Index           =   54
               Left            =   11520
               RightToLeft     =   -1  'True
               TabIndex        =   101
               Top             =   240
               Width           =   2640
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«Ã„«·Ì  ”‰œ«  «·„’—Ê›« "
               Height          =   285
               Index           =   51
               Left            =   11670
               RightToLeft     =   -1  'True
               TabIndex        =   72
               Top             =   3000
               Width           =   1920
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   4800
            Index           =   0
            Left            =   45
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   45
            Width           =   15750
            _cx             =   27781
            _cy             =   8467
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
            GridRows        =   3
            GridCols        =   1
            Frame           =   3
            FrameStyle      =   0
            FrameWidth      =   1
            FrameColor      =   -2147483628
            FrameShadow     =   -2147483632
            FloodStyle      =   1
            _GridInfo       =   $"FrmBillBuyComposite.frx":36FE
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin VSFlex8UCtl.VSFlexGrid FG 
               Height          =   3660
               Left            =   30
               TabIndex        =   3
               Top             =   735
               Width           =   15690
               _cx             =   27675
               _cy             =   6456
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
               Rows            =   2
               Cols            =   22
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmBillBuyComposite.frx":374E
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
               WallPaperAlignment=   0
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   24
            End
            Begin MSComctlLib.Toolbar TBar 
               Height          =   630
               Left            =   30
               TabIndex        =   4
               Top             =   4410
               Width           =   15690
               _ExtentX        =   27675
               _ExtentY        =   1111
               ButtonWidth     =   609
               ButtonHeight    =   1005
               Appearance      =   1
               _Version        =   393216
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   690
               Index           =   4
               Left            =   30
               TabIndex        =   142
               TabStop         =   0   'False
               Top             =   30
               Width           =   15690
               _cx             =   27675
               _cy             =   1217
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
               Begin VB.TextBox TxtPrice 
                  Alignment       =   1  'Right Justify
                  Height          =   345
                  Left            =   765
                  MaxLength       =   10
                  RightToLeft     =   -1  'True
                  TabIndex        =   146
                  Top             =   315
                  Width           =   1410
               End
               Begin VB.TextBox TxtSerial 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFFFFF&
                  Enabled         =   0   'False
                  Height          =   345
                  Left            =   3960
                  MaxLength       =   20
                  RightToLeft     =   -1  'True
                  TabIndex        =   145
                  Top             =   315
                  Width           =   3525
               End
               Begin VB.TextBox TxtQuantity 
                  Alignment       =   1  'Right Justify
                  Height          =   345
                  Left            =   2175
                  MaxLength       =   10
                  RightToLeft     =   -1  'True
                  TabIndex        =   144
                  Top             =   315
                  Width           =   1785
               End
               Begin VB.ComboBox CboItemCase 
                  Height          =   315
                  Left            =   7545
                  RightToLeft     =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   143
                  Top             =   315
                  Width           =   2430
               End
               Begin MSDataListLib.DataCombo DCboItemsName 
                  Height          =   315
                  Left            =   10050
                  TabIndex        =   147
                  Top             =   315
                  Width           =   3585
                  _ExtentX        =   6324
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin MSDataListLib.DataCombo DCboItemsCode 
                  Height          =   315
                  Left            =   13680
                  TabIndex        =   148
                  Top             =   315
                  Width           =   1965
                  _ExtentX        =   3466
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin ImpulseButton.ISButton CmdAdd 
                  Height          =   345
                  Left            =   45
                  TabIndex        =   149
                  Top             =   315
                  Width           =   720
                  _ExtentX        =   1270
                  _ExtentY        =   609
                  ButtonStyle     =   1
                  ButtonPositionImage=   4
                  Caption         =   ""
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
                  ButtonImage     =   "FrmBillBuyComposite.frx":3ADA
                  ColorButton     =   14871017
                  ColorHighlight  =   16777215
                  ColorHoverText  =   16711680
                  ColorShadow     =   -2147483637
                  ColorOutline    =   0
                  DrawFocusRectangle=   0   'False
                  ColorToggledHoverText=   16711680
                  LowerToggledContent=   0   'False
                  ColorTextShadow =   -2147483637
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·”⁄—"
                  Height          =   270
                  Index           =   26
                  Left            =   765
                  RightToLeft     =   -1  'True
                  TabIndex        =   155
                  Top             =   30
                  Width           =   1410
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬂ„Ì…"
                  Height          =   270
                  Index           =   27
                  Left            =   2175
                  RightToLeft     =   -1  'True
                  TabIndex        =   154
                  Top             =   30
                  Width           =   1785
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·”Ì—Ì«·"
                  Height          =   270
                  Index           =   28
                  Left            =   3960
                  RightToLeft     =   -1  'True
                  TabIndex        =   153
                  Top             =   30
                  Width           =   3525
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Õ«·… «·’‰›"
                  Height          =   270
                  Index           =   29
                  Left            =   7545
                  RightToLeft     =   -1  'True
                  TabIndex        =   152
                  Top             =   30
                  Width           =   2430
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "≈”„ «·’‰›"
                  Height          =   270
                  Index           =   30
                  Left            =   10050
                  RightToLeft     =   -1  'True
                  TabIndex        =   151
                  Top             =   30
                  Width           =   3585
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬂÊœ «·’‰›"
                  Height          =   270
                  Index           =   31
                  Left            =   13680
                  RightToLeft     =   -1  'True
                  TabIndex        =   150
                  Top             =   30
                  Width           =   1965
               End
            End
            Begin VB.Label LblItemsCount 
               Alignment       =   2  'Center
               BackColor       =   &H00404040&
               ForeColor       =   &H0000FFFF&
               Height          =   360
               Left            =   30
               RightToLeft     =   -1  'True
               TabIndex        =   5
               Top             =   4410
               Width           =   15690
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   4800
            Index           =   2
            Left            =   16485
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   45
            Width           =   15750
            _cx             =   27781
            _cy             =   8467
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
            BackColor       =   255
            ForeColor       =   -2147483630
            FloodColor      =   6553600
            ForeColorDisabled=   -2147483631
            Caption         =   ""
            Align           =   0
            AutoSizeChildren=   8
            BorderWidth     =   0
            ChildSpacing    =   0
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
            _GridInfo       =   $"FrmBillBuyComposite.frx":3E74
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   2190
               Index           =   10
               Left            =   0
               TabIndex        =   7
               TabStop         =   0   'False
               Top             =   2610
               Width           =   15750
               _cx             =   27781
               _cy             =   3863
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
               GridRows        =   10
               GridCols        =   1
               Frame           =   3
               FrameStyle      =   0
               FrameWidth      =   1
               FrameColor      =   -2147483628
               FrameShadow     =   -2147483632
               FloodStyle      =   1
               _GridInfo       =   $"FrmBillBuyComposite.frx":3EE5
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   9
               Begin C1SizerLibCtl.C1Elastic Ele 
                  Height          =   405
                  Index           =   14
                  Left            =   15
                  TabIndex        =   8
                  TabStop         =   0   'False
                  Top             =   15
                  Width           =   15300
                  _cx             =   26988
                  _cy             =   714
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
                  Begin VB.CheckBox XPChkPayType 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "‘Ìﬂ« "
                     Height          =   330
                     Index           =   2
                     Left            =   12315
                     RightToLeft     =   -1  'True
                     TabIndex        =   9
                     Top             =   75
                     Width           =   1095
                  End
                  Begin ImpulseButton.ISButton CmdCheque 
                     Height          =   330
                     Left            =   6180
                     TabIndex        =   10
                     Top             =   75
                     Width           =   1170
                     _ExtentX        =   2064
                     _ExtentY        =   582
                     Caption         =   " ”ÃÌ· «·‘Ìﬂ« "
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
                     DrawFocusRectangle=   0   'False
                  End
                  Begin MSDataListLib.DataCombo dcbanks 
                     Height          =   315
                     Left            =   0
                     TabIndex        =   64
                     Top             =   0
                     Width           =   2280
                     _ExtentX        =   4022
                     _ExtentY        =   556
                     _Version        =   393216
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin VB.Label Label2 
                     Alignment       =   1  'Right Justify
                     Caption         =   "«·»‰ﬂ"
                     Height          =   300
                     Left            =   2280
                     RightToLeft     =   -1  'True
                     TabIndex        =   65
                     Top             =   150
                     Width           =   690
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Height          =   330
                     Index           =   18
                     Left            =   6960
                     RightToLeft     =   -1  'True
                     TabIndex        =   14
                     Top             =   75
                     Width           =   750
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "≈Ã„«·Ï ﬁÌ„… «·‘Ìﬂ« "
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
                     Height          =   330
                     Index           =   16
                     Left            =   7650
                     RightToLeft     =   -1  'True
                     TabIndex        =   13
                     Top             =   75
                     Visible         =   0   'False
                     Width           =   1725
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "⁄œœ «·‘Ìﬂ« "
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
                     Height          =   330
                     Index           =   17
                     Left            =   10695
                     RightToLeft     =   -1  'True
                     TabIndex        =   12
                     Top             =   75
                     Visible         =   0   'False
                     Width           =   1035
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Height          =   330
                     Index           =   19
                     Left            =   9735
                     RightToLeft     =   -1  'True
                     TabIndex        =   11
                     Top             =   75
                     Width           =   960
                  End
               End
               Begin VSFlex8UCtl.VSFlexGrid FgCheques 
                  Height          =   1275
                  Left            =   15
                  TabIndex        =   181
                  Top             =   435
                  Width           =   15300
                  _cx             =   26987
                  _cy             =   2249
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
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"FrmBillBuyComposite.frx":3F83
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
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   4545
               Index           =   7
               Left            =   0
               TabIndex        =   15
               TabStop         =   0   'False
               Top             =   255
               Width           =   15750
               _cx             =   27781
               _cy             =   8017
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
               GridRows        =   10
               GridCols        =   1
               Frame           =   3
               FrameStyle      =   0
               FrameWidth      =   1
               FrameColor      =   -2147483628
               FrameShadow     =   -2147483632
               FloodStyle      =   1
               _GridInfo       =   $"FrmBillBuyComposite.frx":40B7
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   9
               Begin C1SizerLibCtl.C1Elastic Ele 
                  Height          =   435
                  Index           =   12
                  Left            =   15
                  TabIndex        =   16
                  TabStop         =   0   'False
                  Top             =   15
                  Width           =   15300
                  _cx             =   26988
                  _cy             =   767
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
                  Begin VB.CheckBox XPChkPayType 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "¬Ã· "
                     Height          =   720
                     Index           =   1
                     Left            =   12345
                     RightToLeft     =   -1  'True
                     TabIndex        =   20
                     Top             =   -210
                     Width           =   900
                  End
                  Begin VB.TextBox XPTxtValue 
                     Alignment       =   1  'Right Justify
                     Height          =   645
                     Index           =   1
                     Left            =   9975
                     MaxLength       =   10
                     RightToLeft     =   -1  'True
                     TabIndex        =   19
                     Top             =   45
                     Width           =   1275
                  End
                  Begin VB.TextBox XPTxtSerial 
                     Alignment       =   1  'Right Justify
                     Height          =   780
                     Index           =   1
                     Left            =   7245
                     Locked          =   -1  'True
                     RightToLeft     =   -1  'True
                     TabIndex        =   18
                     Top             =   45
                     Width           =   1755
                  End
                  Begin VB.CheckBox ChkInstall 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   " ﬁ”Ìÿ"
                     Height          =   210
                     Left            =   2700
                     RightToLeft     =   -1  'True
                     TabIndex        =   17
                     Top             =   75
                     Width           =   1170
                  End
                  Begin ImpulseButton.ISButton CmdINSTALLMENT 
                     Height          =   330
                     Left            =   195
                     TabIndex        =   21
                     Top             =   75
                     Width           =   1575
                     _ExtentX        =   2778
                     _ExtentY        =   582
                     ButtonPositionImage=   1
                     Caption         =   "Õ”«» «·√ﬁ”«ÿ"
                     BackColor       =   14871017
                     Enabled         =   0   'False
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   178
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ButtonImage     =   "FrmBillBuyComposite.frx":4158
                     ColorButton     =   14871017
                     ColorHighlight  =   16777215
                     ColorHoverText  =   16711680
                     ColorShadow     =   4210752
                     ColorOutline    =   0
                     DrawFocusRectangle=   0   'False
                     ColorToggledHoverText=   16711680
                     ColorTextShadow =   4210752
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   " «—ÌŒ «·«” Õﬁ«ﬁ"
                     Height          =   1005
                     Index           =   21
                     Left            =   4350
                     RightToLeft     =   -1  'True
                     TabIndex        =   52
                     Top             =   1080
                     Width           =   960
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "„”·”·"
                     Height          =   660
                     Index           =   14
                     Left            =   9120
                     RightToLeft     =   -1  'True
                     TabIndex        =   23
                     Top             =   75
                     Width           =   615
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·ﬁÌ„…"
                     Height          =   720
                     Index           =   15
                     Left            =   11505
                     RightToLeft     =   -1  'True
                     TabIndex        =   22
                     Top             =   75
                     Width           =   420
                  End
               End
               Begin VSFlex8UCtl.VSFlexGrid FgInstallments 
                  Height          =   870
                  Left            =   15
                  TabIndex        =   156
                  Top             =   465
                  Width           =   15300
                  _cx             =   26987
                  _cy             =   1535
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
                  Cols            =   5
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"FrmBillBuyComposite.frx":44F2
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
               Begin C1SizerLibCtl.C1Elastic Ele 
                  Height          =   435
                  Index           =   13
                  Left            =   15
                  TabIndex        =   165
                  TabStop         =   0   'False
                  Top             =   1785
                  Width           =   15300
                  _cx             =   26988
                  _cy             =   767
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
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·ﬁÌ„… «·„»œ∆Ì…"
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
                     Height          =   285
                     Index           =   37
                     Left            =   420
                     RightToLeft     =   -1  'True
                     TabIndex        =   180
                     Top             =   75
                     Width           =   1590
                  End
                  Begin VB.Label LblStartValue 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Height          =   285
                     Left            =   75
                     RightToLeft     =   -1  'True
                     TabIndex        =   179
                     Top             =   75
                     Width           =   345
                  End
                  Begin VB.Label LblInstallSeprator 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     ForeColor       =   &H00FF0000&
                     Height          =   285
                     Left            =   3135
                     RightToLeft     =   -1  'True
                     TabIndex        =   178
                     Top             =   75
                     Width           =   375
                  End
                  Begin VB.Label LblPrecenValue 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Height          =   285
                     Left            =   11850
                     RightToLeft     =   -1  'True
                     TabIndex        =   177
                     Top             =   75
                     Width           =   450
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "‰”»… «·›«∆œ…"
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
                     Height          =   285
                     Index           =   35
                     Left            =   12300
                     RightToLeft     =   -1  'True
                     TabIndex        =   176
                     Top             =   75
                     Width           =   705
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "‰Ê⁄ «·›«∆œ…"
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
                     Height          =   285
                     Index           =   34
                     Left            =   14085
                     RightToLeft     =   -1  'True
                     TabIndex        =   175
                     Top             =   75
                     Width           =   1140
                  End
                  Begin VB.Label LblPrecenType 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Height          =   285
                     Left            =   13005
                     RightToLeft     =   -1  'True
                     TabIndex        =   174
                     Top             =   75
                     Width           =   1005
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·„»·€ «·ﬂ·Ï"
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
                     Height          =   285
                     Index           =   36
                     Left            =   10455
                     RightToLeft     =   -1  'True
                     TabIndex        =   173
                     Top             =   75
                     Width           =   1335
                  End
                  Begin VB.Label LblInstallTotal 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Height          =   285
                     Left            =   9495
                     RightToLeft     =   -1  'True
                     TabIndex        =   172
                     Top             =   75
                     Width           =   885
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "⁄œœ «·√ﬁ”«ÿ"
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
                     Height          =   285
                     Index           =   38
                     Left            =   8010
                     RightToLeft     =   -1  'True
                     TabIndex        =   171
                     Top             =   75
                     Width           =   1410
                  End
                  Begin VB.Label LblInstallCount 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Height          =   285
                     Left            =   7425
                     RightToLeft     =   -1  'True
                     TabIndex        =   170
                     Top             =   75
                     Width           =   525
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«Ê· ﬁ”ÿ"
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
                     Height          =   285
                     Index           =   40
                     Left            =   6300
                     RightToLeft     =   -1  'True
                     TabIndex        =   169
                     Top             =   75
                     Width           =   1125
                  End
                  Begin VB.Label LblFirstInstallDate 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Height          =   285
                     Left            =   5055
                     RightToLeft     =   -1  'True
                     TabIndex        =   168
                     Top             =   75
                     Width           =   1245
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "› —… «· ﬁ”Ìÿ"
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
                     Height          =   285
                     Index           =   42
                     Left            =   3510
                     RightToLeft     =   -1  'True
                     TabIndex        =   167
                     Top             =   75
                     Width           =   1545
                  End
                  Begin VB.Label LblInstallmentType 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     ForeColor       =   &H00FF0000&
                     Height          =   285
                     Left            =   2085
                     RightToLeft     =   -1  'True
                     TabIndex        =   166
                     Top             =   75
                     Width           =   990
                  End
               End
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   255
               Index           =   11
               Left            =   0
               TabIndex        =   157
               TabStop         =   0   'False
               Top             =   0
               Width           =   15750
               _cx             =   27781
               _cy             =   450
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
               Begin VB.TextBox XPTxtValue 
                  Alignment       =   1  'Right Justify
                  Height          =   375
                  Index           =   0
                  Left            =   9945
                  MaxLength       =   10
                  RightToLeft     =   -1  'True
                  TabIndex        =   161
                  Top             =   0
                  Width           =   1260
               End
               Begin VB.TextBox XPTxtSerial 
                  Alignment       =   1  'Right Justify
                  Height          =   375
                  Index           =   0
                  Left            =   7320
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   160
                  Top             =   60
                  Width           =   1725
               End
               Begin VB.CheckBox XPChkPayType 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "‰ﬁœ«"
                  Height          =   345
                  Index           =   0
                  Left            =   12150
                  RightToLeft     =   -1  'True
                  TabIndex        =   159
                  Top             =   0
                  Width           =   1140
               End
               Begin MSDataListLib.DataCombo DcboCurrency 
                  Height          =   315
                  Left            =   1530
                  TabIndex        =   158
                  Top             =   120
                  Visible         =   0   'False
                  Width           =   1815
                  _ExtentX        =   3201
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·⁄„·…"
                  Height          =   225
                  Index           =   20
                  Left            =   3075
                  RightToLeft     =   -1  'True
                  TabIndex        =   164
                  Top             =   150
                  Visible         =   0   'False
                  Width           =   810
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬁÌ„…"
                  Height          =   345
                  Index           =   13
                  Left            =   11295
                  RightToLeft     =   -1  'True
                  TabIndex        =   163
                  Top             =   90
                  Width           =   630
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "„”·”·"
                  Height          =   345
                  Index           =   12
                  Left            =   9225
                  RightToLeft     =   -1  'True
                  TabIndex        =   162
                  Top             =   90
                  Width           =   630
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   4800
            Index           =   15
            Left            =   16785
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   45
            Width           =   15750
            _cx             =   27781
            _cy             =   8467
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial (Arabic)"
               Size            =   12
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
            BackColor       =   14871017
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
            GridRows        =   7
            GridCols        =   1
            Frame           =   3
            FrameStyle      =   0
            FrameWidth      =   1
            FrameColor      =   -2147483628
            FrameShadow     =   -2147483632
            FloodStyle      =   1
            _GridInfo       =   $"FrmBillBuyComposite.frx":45C3
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin VB.TextBox TxtBillComment 
               Alignment       =   1  'Right Justify
               Height          =   2340
               Left            =   15
               MultiLine       =   -1  'True
               RightToLeft     =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   40
               Top             =   2445
               Width           =   15720
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   1275
               Index           =   18
               Left            =   15
               TabIndex        =   25
               TabStop         =   0   'False
               Top             =   1155
               Width           =   15720
               _cx             =   27728
               _cy             =   2249
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
               Appearance      =   5
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
               Begin VB.CheckBox ChkTaxSerivce 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "÷—»Ì… Œœ„…"
                  Height          =   420
                  Left            =   6390
                  RightToLeft     =   -1  'True
                  TabIndex        =   27
                  Top             =   240
                  Visible         =   0   'False
                  Width           =   990
               End
               Begin VB.TextBox TxtTaxServiceValue 
                  Alignment       =   1  'Right Justify
                  Height          =   660
                  Left            =   4980
                  MaxLength       =   4
                  RightToLeft     =   -1  'True
                  TabIndex        =   26
                  Top             =   120
                  Visible         =   0   'False
                  Width           =   705
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Enabled         =   0   'False
                  Height          =   450
                  Index           =   49
                  Left            =   1200
                  RightToLeft     =   -1  'True
                  TabIndex        =   50
                  Top             =   165
                  Visible         =   0   'False
                  Width           =   285
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬁÌ„…"
                  Enabled         =   0   'False
                  Height          =   615
                  Index           =   43
                  Left            =   5550
                  RightToLeft     =   -1  'True
                  TabIndex        =   29
                  Top             =   165
                  Visible         =   0   'False
                  Width           =   585
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "%"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   495
                  Index           =   47
                  Left            =   4440
                  RightToLeft     =   -1  'True
                  TabIndex        =   28
                  Top             =   165
                  Visible         =   0   'False
                  Width           =   210
               End
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   1275
               Index           =   17
               Left            =   15
               TabIndex        =   30
               TabStop         =   0   'False
               Top             =   1155
               Visible         =   0   'False
               Width           =   15720
               _cx             =   27728
               _cy             =   2249
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
               Appearance      =   5
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
               Begin VB.CheckBox ChkTaxStamp 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "œ„€…"
                  Height          =   285
                  Left            =   6540
                  RightToLeft     =   -1  'True
                  TabIndex        =   32
                  Top             =   75
                  Width           =   840
               End
               Begin VB.TextBox TxtTaxStampValue 
                  Alignment       =   1  'Right Justify
                  Height          =   345
                  Left            =   4980
                  MaxLength       =   4
                  RightToLeft     =   -1  'True
                  TabIndex        =   31
                  Top             =   75
                  Visible         =   0   'False
                  Width           =   705
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Enabled         =   0   'False
                  Height          =   285
                  Index           =   33
                  Left            =   1260
                  RightToLeft     =   -1  'True
                  TabIndex        =   49
                  Top             =   75
                  Visible         =   0   'False
                  Width           =   990
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬁÌ„…"
                  Enabled         =   0   'False
                  Height          =   300
                  Index           =   41
                  Left            =   5550
                  RightToLeft     =   -1  'True
                  TabIndex        =   34
                  Top             =   75
                  Width           =   585
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "$"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   48
                  Left            =   4440
                  RightToLeft     =   -1  'True
                  TabIndex        =   33
                  Top             =   75
                  Width           =   150
               End
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   465
               Index           =   16
               Left            =   15
               TabIndex        =   35
               TabStop         =   0   'False
               Top             =   480
               Visible         =   0   'False
               Width           =   15720
               _cx             =   27728
               _cy             =   820
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
               Appearance      =   5
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
               Begin VB.CheckBox ChkTaxAdd 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "÷—»Ì»… Œ’„ Ê≈÷«›… (√—»«Õ  Ã«—Ì…)"
                  Height          =   585
                  Left            =   6135
                  RightToLeft     =   -1  'True
                  TabIndex        =   37
                  Top             =   45
                  Visible         =   0   'False
                  Width           =   1335
               End
               Begin VB.TextBox TxtTaxAddValue 
                  Alignment       =   1  'Right Justify
                  Height          =   480
                  Left            =   4980
                  MaxLength       =   4
                  RightToLeft     =   -1  'True
                  TabIndex        =   36
                  Top             =   75
                  Width           =   705
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Enabled         =   0   'False
                  Height          =   345
                  Index           =   32
                  Left            =   420
                  RightToLeft     =   -1  'True
                  TabIndex        =   48
                  Top             =   120
                  Width           =   135
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬁÌ„…"
                  Enabled         =   0   'False
                  Height          =   390
                  Index           =   39
                  Left            =   5685
                  RightToLeft     =   -1  'True
                  TabIndex        =   39
                  Top             =   120
                  Visible         =   0   'False
                  Width           =   450
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "%"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   345
                  Index           =   46
                  Left            =   4080
                  RightToLeft     =   -1  'True
                  TabIndex        =   38
                  Top             =   120
                  Visible         =   0   'False
                  Width           =   510
               End
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   450
               Index           =   8
               Left            =   15
               TabIndex        =   41
               TabStop         =   0   'False
               Top             =   15
               Visible         =   0   'False
               Width           =   15720
               _cx             =   27728
               _cy             =   794
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
               Appearance      =   5
               MousePointer    =   0
               Version         =   801
               BackColor       =   14871017
               ForeColor       =   -2147483630
               FloodColor      =   6553600
               ForeColorDisabled=   -2147483631
               Caption         =   ""
               Align           =   0
               AutoSizeChildren=   7
               BorderWidth     =   0
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
               Begin VB.TextBox XPTxtTaxValue 
                  Alignment       =   1  'Right Justify
                  Height          =   240
                  Left            =   4980
                  MaxLength       =   4
                  RightToLeft     =   -1  'True
                  TabIndex        =   43
                  Top             =   45
                  Width           =   705
               End
               Begin VB.CheckBox XPChkTAX 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "÷—»Ì»… «·„»Ì⁄« "
                  Height          =   150
                  Left            =   6135
                  RightToLeft     =   -1  'True
                  TabIndex        =   42
                  Top             =   90
                  Visible         =   0   'False
                  Width           =   1335
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Enabled         =   0   'False
                  Height          =   150
                  Index           =   25
                  Left            =   420
                  RightToLeft     =   -1  'True
                  TabIndex        =   47
                  Top             =   90
                  Width           =   135
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·ﬁÌ„…"
                  Enabled         =   0   'False
                  Height          =   165
                  Index           =   22
                  Left            =   5430
                  RightToLeft     =   -1  'True
                  TabIndex        =   45
                  Top             =   90
                  Visible         =   0   'False
                  Width           =   630
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "%"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9.75
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   165
                  Index           =   45
                  Left            =   3960
                  RightToLeft     =   -1  'True
                  TabIndex        =   44
                  Top             =   90
                  Visible         =   0   'False
                  Width           =   690
               End
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "≈÷«›… √Ì… „·«ÕŸ«  ⁄·Ï «·›« Ê—…"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   12
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   1275
               Index           =   44
               Left            =   15
               RightToLeft     =   -1  'True
               TabIndex        =   46
               Top             =   1155
               Width           =   15720
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   4800
            Index           =   9
            Left            =   17985
            TabIndex        =   120
            TabStop         =   0   'False
            Top             =   45
            Width           =   15750
            _cx             =   27781
            _cy             =   8467
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
            Begin VB.TextBox TXTFactoryExpenses 
               Alignment       =   2  'Center
               Height          =   405
               Left            =   7920
               Locked          =   -1  'True
               RightToLeft     =   -1  'True
               TabIndex        =   121
               Top             =   2880
               Width           =   1215
            End
            Begin VSFlex8Ctl.VSFlexGrid Fg_Journal 
               Height          =   2340
               Left            =   1800
               TabIndex        =   122
               Top             =   480
               Width           =   12600
               _cx             =   22225
               _cy             =   4128
               Appearance      =   1
               BorderStyle     =   1
               Enabled         =   0   'False
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
               Rows            =   1
               Cols            =   9
               FixedRows       =   1
               FixedCols       =   0
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"FrmBillBuyComposite.frx":463A
               ScrollTrack     =   0   'False
               ScrollBars      =   2
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
               Begin VB.PictureBox PicDes 
                  BorderStyle     =   0  'None
                  Height          =   1635
                  Left            =   240
                  RightToLeft     =   -1  'True
                  ScaleHeight     =   1635
                  ScaleWidth      =   2925
                  TabIndex        =   123
                  Top             =   960
                  Visible         =   0   'False
                  Width           =   2925
                  Begin VB.TextBox TxtDes 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H80000018&
                     BorderStyle     =   0  'None
                     Height          =   1125
                     Left            =   30
                     MultiLine       =   -1  'True
                     RightToLeft     =   -1  'True
                     ScrollBars      =   3  'Both
                     TabIndex        =   124
                     Top             =   360
                     Visible         =   0   'False
                     Width           =   2115
                  End
                  Begin VB.Label LblDes 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H8000000C&
                     Caption         =   "Ì„ﬂ‰ﬂ ﬂ «»…  ⁄·Ìﬁ Â‰«:"
                     ForeColor       =   &H0000C8FF&
                     Height          =   315
                     Left            =   0
                     RightToLeft     =   -1  'True
                     TabIndex        =   125
                     Top             =   0
                     Width           =   2445
                  End
               End
               Begin VDSCOMBOLibCtl.SmartCombo CboDes 
                  Height          =   315
                  Left            =   240
                  TabIndex        =   126
                  ToolTipText     =   "ﬂ «»…  ⁄·Ìﬁ"
                  Top             =   600
                  Visible         =   0   'False
                  Width           =   2955
                  _cx             =   1973752924
                  _cy             =   1973748268
                  Alignment       =   0
                  Appearance      =   3
                  AutoSearch      =   0   'False
                  BackColor       =   -2147483624
                  BackgroundColor =   -2147483633
                  BorderColor     =   0
                  BorderVisible   =   -1  'True
                  Caption         =   "SmartCombo1"
                  CaptionAlignment=   4
                  CaptionBackColor=   -2147483633
                  BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  CaptionForeColor=   -2147483630
                  CaptionHeight   =   15
                  CaptionOnTop    =   0   'False
                  CaptionMultiLine=   0
                  Checkbox3D      =   0   'False
                  CheckboxAlignment=   5
                  CheckboxBackColor=   16777215
                  CheckboxSize    =   13
                  CheckboxValue   =   0
                  BrowsePictureAlignment=   5
                  BrowsePictureStretchH=   0
                  BrowsePictureStretchV=   0
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
                  ForeColor       =   0
                  Gap             =   0
                  HideSelection   =   -1  'True
                  Locked          =   0   'False
                  MaxLength       =   0
                  MultiLine       =   0
                  OnFocus         =   3
                  PasswordChar    =   ""
                  Picture         =   "FrmBillBuyComposite.frx":479A
                  PictureAlignment=   5
                  PictureBackColor=   -2147483624
                  PictureStretchH =   0
                  PictureStretchV =   0
                  Redraw          =   -1  'True
                  ScrollBar       =   0
                  Style           =   0
                  Text            =   ""
                  UnderLine       =   0   'False
                  Enabled0        =   -1  'True
                  Position0       =   0
                  Tip0            =   "Caption"
                  Visible0        =   0   'False
                  Width0          =   90
                  Enabled1        =   -1  'True
                  Position1       =   1
                  Tip1            =   ""
                  Visible1        =   -1  'True
                  Width1          =   32
                  Enabled2        =   -1  'True
                  Position2       =   2
                  Tip2            =   "Check Box (Space, Ctrl + Space)"
                  Visible2        =   0   'False
                  Width2          =   16
                  Enabled3        =   -1  'True
                  Position3       =   3
                  Tip3            =   "ﬂ «»…  ⁄·Ìﬁ"
                  Visible3        =   -1  'True
                  Width3          =   145
                  Enabled4        =   -1  'True
                  Position4       =   4
                  Tip4            =   "Left Spinner (Alt + Left)"
                  Visible4        =   0   'False
                  Width4          =   16
                  Enabled5        =   -1  'True
                  Position5       =   5
                  Tip5            =   "Right Spinner (Alt + Right)"
                  Visible5        =   0   'False
                  Width5          =   16
                  Enabled6        =   -1  'True
                  Position6       =   6
                  Tip6            =   "Up Spinner (Ctrl + Up)"
                  Visible6        =   0   'False
                  Width6          =   16
                  Enabled7        =   -1  'True
                  Position7       =   7
                  Tip7            =   "Down Spinner (Ctrl + Down)"
                  Visible7        =   0   'False
                  Width7          =   16
                  Enabled8        =   -1  'True
                  Position8       =   8
                  Tip8            =   "Browse (Alt + Enter)"
                  Visible8        =   0   'False
                  Width8          =   16
                  Enabled9        =   -1  'True
                  Position9       =   9
                  Tip9            =   " (Alt + Down, F4)"
                  Visible9        =   -1  'True
                  Width9          =   16
                  Enabled10       =   -1  'True
                  Position10      =   10
                  Tip10           =   "Right Arrow (Alt + >)"
                  Visible10       =   0   'False
                  Width10         =   16
               End
            End
            Begin ImpulseButton.ISButton Cmd 
               Height          =   390
               Index           =   9
               Left            =   13200
               TabIndex        =   127
               Top             =   2880
               Width           =   1050
               _ExtentX        =   1852
               _ExtentY        =   688
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   "Õ–› ”ÿ—"
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
               ButtonImage     =   "FrmBillBuyComposite.frx":4D34
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin VB.Label Label19 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«Œ Ì«— «·„’—Ê›«  «· ﬁœÌ—ÌÂ"
               Height          =   255
               Left            =   10440
               RightToLeft     =   -1  'True
               TabIndex        =   129
               Top             =   120
               Width           =   3855
            End
            Begin VB.Label Label18 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«Ã„«·Ì  «·„’«—Ì› «· ﬁœÌ—ÌÂ"
               Height          =   375
               Left            =   9120
               RightToLeft     =   -1  'True
               TabIndex        =   128
               Top             =   3000
               Width           =   2055
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   4800
            Index           =   19
            Left            =   18285
            TabIndex        =   130
            TabStop         =   0   'False
            Top             =   45
            Width           =   15750
            _cx             =   27781
            _cy             =   8467
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
            Begin VB.TextBox Text4 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   4320
               RightToLeft     =   -1  'True
               TabIndex        =   137
               Top             =   0
               Visible         =   0   'False
               Width           =   1425
            End
            Begin VB.OptionButton BillBasedOn 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "”‰œ«  «·«” ·«„"
               Height          =   195
               Index           =   1
               Left            =   10200
               RightToLeft     =   -1  'True
               TabIndex        =   136
               Top             =   600
               Width           =   4215
            End
            Begin VB.OptionButton BillBasedOn 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "√Ê«„— «·‘—«¡"
               Height          =   195
               Index           =   2
               Left            =   10320
               RightToLeft     =   -1  'True
               TabIndex        =   135
               Top             =   3000
               Visible         =   0   'False
               Width           =   4095
            End
            Begin VB.OptionButton BillBasedOn 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "›« Ê—… „‘ —Ì« "
               Height          =   195
               Index           =   0
               Left            =   10080
               RightToLeft     =   -1  'True
               TabIndex        =   134
               Top             =   360
               Value           =   -1  'True
               Visible         =   0   'False
               Width           =   4335
            End
            Begin VB.Frame Frame3 
               BackColor       =   &H00E2E9E9&
               Caption         =   "»Ì«‰«  ﬁÌœ «·›« Ê—Â"
               Height          =   1575
               Left            =   240
               RightToLeft     =   -1  'True
               TabIndex        =   131
               Top             =   720
               Width           =   4695
               Begin VB.TextBox TxtNoteSerial 
                  Alignment       =   1  'Right Justify
                  Height          =   315
                  Left            =   1560
                  RightToLeft     =   -1  'True
                  TabIndex        =   141
                  Top             =   600
                  Width           =   2625
               End
               Begin ImpulseButton.ISButton Cmd 
                  CausesValidation=   0   'False
                  Height          =   375
                  Index           =   10
                  Left            =   120
                  TabIndex        =   132
                  Top             =   600
                  Width           =   1215
                  _ExtentX        =   2143
                  _ExtentY        =   661
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   "ÿ»«⁄Â «·ﬁÌœ"
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
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "—ﬁ„ «·ﬁÌœ ··›« Ê—Â"
                  Height          =   195
                  Index           =   62
                  Left            =   1920
                  TabIndex        =   133
                  Top             =   240
                  Width           =   2175
               End
            End
            Begin VSFlex8UCtl.VSFlexGrid GRID1 
               Height          =   2085
               Left            =   5160
               TabIndex        =   138
               Tag             =   "1"
               Top             =   840
               Width           =   9255
               _cx             =   16325
               _cy             =   3678
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
               AllowUserResizing=   0
               SelectionMode   =   0
               GridLines       =   1
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   3
               Cols            =   9
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"FrmBillBuyComposite.frx":52CE
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
            Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
               Height          =   1725
               Left            =   5160
               TabIndex        =   139
               Tag             =   "1"
               Top             =   3240
               Visible         =   0   'False
               Width           =   9255
               _cx             =   16325
               _cy             =   3043
               Appearance      =   1
               BorderStyle     =   1
               Enabled         =   0   'False
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
               AllowUserResizing=   0
               SelectionMode   =   0
               GridLines       =   1
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   3
               Cols            =   6
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmBillBuyComposite.frx":541B
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
            Begin ImpulseButton.ISButton CmdAttach 
               Height          =   375
               Left            =   3720
               TabIndex        =   235
               Top             =   2280
               Width           =   1155
               _ExtentX        =   2037
               _ExtentY        =   661
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
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·›« Ê—Â »‰«¡ ⁄·Ï"
               Height          =   300
               Index           =   61
               Left            =   12240
               TabIndex        =   140
               Top             =   120
               Width           =   2160
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   600
         Index           =   6
         Left            =   15
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   15
         Width           =   15840
         _cx             =   27940
         _cy             =   1058
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
         Caption         =   "›« Ê—… „‘ —Ì«  „Ã„⁄…"
         Align           =   0
         AutoSizeChildren=   7
         BorderWidth     =   0
         ChildSpacing    =   0
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
         Begin VB.TextBox oldtxtNoteSerial1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   2955
            RightToLeft     =   -1  'True
            TabIndex        =   191
            Top             =   0
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.TextBox TxtFillData 
            Alignment       =   1  'Right Justify
            BackColor       =   &H000000FF&
            Height          =   315
            Left            =   8565
            RightToLeft     =   -1  'True
            TabIndex        =   56
            Top             =   105
            Visible         =   0   'False
            Width           =   600
         End
         Begin VB.TextBox TxtModFlg 
            Alignment       =   1  'Right Justify
            BackColor       =   &H000000FF&
            Height          =   345
            Left            =   9345
            RightToLeft     =   -1  'True
            TabIndex        =   55
            Top             =   90
            Visible         =   0   'False
            Width           =   915
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   10575
            RightToLeft     =   -1  'True
            TabIndex        =   54
            Top             =   120
            Visible         =   0   'False
            Width           =   630
         End
         Begin ImpulseButton.ISButton CmdNotes 
            Height          =   390
            Left            =   5055
            TabIndex        =   57
            Top             =   90
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   688
            ButtonStyle     =   1
            ButtonPositionImage=   3
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
            ButtonImage     =   "FrmBillBuyComposite.frx":550E
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton XPBtnMove 
            Height          =   345
            Index           =   0
            Left            =   2700
            TabIndex        =   58
            Top             =   105
            Width           =   1050
            _ExtentX        =   1852
            _ExtentY        =   609
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
            ButtonImage     =   "FrmBillBuyComposite.frx":58A8
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
            Left            =   1440
            TabIndex        =   59
            Top             =   105
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   609
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
            ButtonImage     =   "FrmBillBuyComposite.frx":5C42
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
            Left            =   3810
            TabIndex        =   60
            Top             =   105
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   609
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
            ButtonImage     =   "FrmBillBuyComposite.frx":5FDC
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
            Left            =   165
            TabIndex        =   61
            Top             =   105
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   609
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
            ButtonImage     =   "FrmBillBuyComposite.frx":6376
            ColorHighlight  =   4194304
            ColorHoverText  =   16777215
            ColorShadow     =   -2147483631
            ColorOutline    =   -2147483631
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
            ColorToggledHoverText=   16777215
            ColorTextShadow =   16777215
         End
         Begin ImpulseButton.ISButton CmdRetruns 
            Height          =   390
            Left            =   6315
            TabIndex        =   62
            Top             =   90
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   688
            ButtonStyle     =   1
            ButtonPositionImage=   3
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
            ButtonImage     =   "FrmBillBuyComposite.frx":6710
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton CmdInfo 
            Height          =   480
            Left            =   11295
            TabIndex        =   63
            Top             =   0
            Width           =   645
            _ExtentX        =   1138
            _ExtentY        =   847
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "FrmBillBuyComposite.frx":6CAA
            ButtonImageHover=   "FrmBillBuyComposite.frx":7984
            ColorButton     =   16777215
            DrawFocusRectangle=   0   'False
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   435
            Index           =   67
            Left            =   4935
            RightToLeft     =   -1  'True
            TabIndex        =   192
            Top             =   120
            Width           =   6855
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   1695
         Index           =   5
         Left            =   15
         TabIndex        =   74
         TabStop         =   0   'False
         Top             =   630
         Width           =   15840
         _cx             =   27940
         _cy             =   2990
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
         Begin VB.TextBox TxtManualNo1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H0080FFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   3915
            TabIndex        =   201
            Top             =   1200
            Width           =   1290
         End
         Begin VB.TextBox TxtSearchCode 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   13290
            RightToLeft     =   -1  'True
            TabIndex        =   200
            Top             =   840
            Width           =   1305
         End
         Begin VB.TextBox TxtLCNO 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   555
            Locked          =   -1  'True
            MaxLength       =   55
            RightToLeft     =   -1  'True
            TabIndex        =   195
            Top             =   390
            Width           =   1635
         End
         Begin VB.TextBox txtManualNO 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   11160
            RightToLeft     =   -1  'True
            TabIndex        =   189
            Top             =   0
            Width           =   1260
         End
         Begin VB.ComboBox CBoBasedON 
            Height          =   315
            ItemData        =   "FrmBillBuyComposite.frx":865E
            Left            =   8730
            List            =   "FrmBillBuyComposite.frx":8660
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   184
            Top             =   0
            Width           =   1185
         End
         Begin VB.TextBox TXTNoteID 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   5595
            RightToLeft     =   -1  'True
            TabIndex        =   112
            Top             =   -225
            Visible         =   0   'False
            Width           =   1155
         End
         Begin VB.TextBox TxtNoteSerial1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   13290
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   111
            Top             =   0
            Width           =   1305
         End
         Begin VB.TextBox TXT_order_no 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   6555
            MaxLength       =   55
            RightToLeft     =   -1  'True
            TabIndex        =   100
            Top             =   0
            Width           =   1230
         End
         Begin VB.ComboBox CboPayMentType 
            Height          =   315
            Left            =   11160
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   85
            Top             =   375
            Width           =   1230
         End
         Begin VB.ComboBox XPCboDiscountType 
            Height          =   315
            Left            =   3000
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   84
            Top             =   840
            Width           =   2565
         End
         Begin VB.TextBox XPTxtDiscountVal 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   555
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   83
            Top             =   735
            Width           =   1635
         End
         Begin VB.TextBox TxtTransSerial 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   14490
            RightToLeft     =   -1  'True
            TabIndex        =   82
            Top             =   1620
            Visible         =   0   'False
            Width           =   1320
         End
         Begin VB.TextBox XPTxtBillID 
            Alignment       =   1  'Right Justify
            Height          =   405
            Left            =   10350
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   81
            Top             =   -360
            Visible         =   0   'False
            Width           =   1275
         End
         Begin VB.TextBox TxtStoreID 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   13290
            RightToLeft     =   -1  'True
            TabIndex        =   80
            Top             =   1215
            Width           =   1305
         End
         Begin VB.TextBox TxtCusID 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   13395
            RightToLeft     =   -1  'True
            TabIndex        =   79
            Top             =   840
            Width           =   1050
         End
         Begin VB.CommandButton Command1 
            Caption         =   " ÕÊÌ· «·Ï  «–‰ «÷«›… "
            Height          =   330
            Left            =   -345
            RightToLeft     =   -1  'True
            TabIndex        =   78
            Top             =   1215
            Visible         =   0   'False
            Width           =   2385
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   675
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   76
            Top             =   735
            Visible         =   0   'False
            Width           =   1350
         End
         Begin VB.TextBox txt_Currency_rate 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   495
            RightToLeft     =   -1  'True
            TabIndex        =   75
            Text            =   "1"
            Top             =   15
            Width           =   855
         End
         Begin MSDataListLib.DataCombo DcCurrency 
            Height          =   315
            Left            =   1350
            TabIndex        =   77
            Top             =   0
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DBCboClientName 
            Height          =   315
            Left            =   9720
            TabIndex        =   86
            Top             =   840
            Width           =   3525
            _ExtentX        =   6218
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCboStoreName 
            Height          =   315
            Left            =   9720
            TabIndex        =   87
            Top             =   1215
            Width           =   3525
            _ExtentX        =   6218
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSComCtl2.DTPicker XPDtbBill 
            Height          =   315
            Left            =   13290
            TabIndex        =   88
            Top             =   360
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   556
            _Version        =   393216
            Format          =   99418113
            CurrentDate     =   38784
         End
         Begin ImpulseButton.ISButton XPBtnNewClients 
            Height          =   420
            Left            =   14505
            TabIndex        =   89
            TabStop         =   0   'False
            Top             =   795
            Visible         =   0   'False
            Width           =   435
            _ExtentX        =   767
            _ExtentY        =   741
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
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
            ButtonImage     =   "FrmBillBuyComposite.frx":8662
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorShadow     =   -2147483631
            ColorOutline    =   -2147483631
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin MSDataListLib.DataCombo DCproject 
            Height          =   315
            Left            =   2955
            TabIndex        =   103
            Top             =   330
            Width           =   2640
            _ExtentX        =   4657
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSComCtl2.DTPicker DTArrivalDate 
            Height          =   330
            Left            =   6525
            TabIndex        =   116
            Top             =   1200
            Width           =   1740
            _ExtentX        =   3069
            _ExtentY        =   582
            _Version        =   393216
            Format          =   99418113
            CurrentDate     =   38784
         End
         Begin MSDataListLib.DataCombo dcBranch 
            Height          =   315
            Left            =   2985
            TabIndex        =   118
            Top             =   0
            Width           =   2670
            _ExtentX        =   4710
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcboBox 
            Height          =   315
            Left            =   6525
            TabIndex        =   187
            Top             =   840
            Width           =   2400
            _ExtentX        =   4233
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCDocTypes 
            Height          =   315
            Left            =   6570
            TabIndex        =   193
            Top             =   360
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin ImpulseButton.ISButton ISButton1 
            Height          =   270
            Left            =   180
            TabIndex        =   197
            Top             =   375
            Width           =   360
            _ExtentX        =   635
            _ExtentY        =   476
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "⁄—÷"
            BackColor       =   12632256
            ForeColor       =   16711680
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColorButton     =   12632256
            ColorHighlight  =   16777215
            ColorHoverText  =   255
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledText=   16711680
            ColorToggledHoverText=   255
            ColorTextShadow =   -2147483637
         End
         Begin MSComCtl2.DTPicker DtpDelayDate 
            Height          =   330
            Left            =   8715
            TabIndex        =   198
            Top             =   390
            Width           =   1245
            _ExtentX        =   2196
            _ExtentY        =   582
            _Version        =   393216
            Format          =   99418113
            CurrentDate     =   38784
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "—ﬁ„ «–‰ «·«” ·«„ «·ÌœÊÌ"
            Height          =   195
            Index           =   69
            Left            =   5055
            TabIndex        =   202
            Top             =   1200
            Width           =   1365
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   " «—ÌŒ «·«” Õﬁ«ﬁ"
            Height          =   525
            Left            =   9885
            RightToLeft     =   -1  'True
            TabIndex        =   199
            Top             =   360
            Width           =   1170
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·«⁄ „«œ"
            Height          =   255
            Index           =   68
            Left            =   2160
            RightToLeft     =   -1  'True
            TabIndex        =   196
            Top             =   495
            Width           =   690
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "‰Ê⁄ «·”‰œ"
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   7620
            TabIndex        =   194
            Top             =   360
            Width           =   870
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "›« Ê—… «·„Ê—œ"
            Height          =   390
            Index           =   53
            Left            =   12360
            RightToLeft     =   -1  'True
            TabIndex        =   190
            Top             =   0
            Width           =   870
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·Œ“Ì‰Â"
            Height          =   330
            Index           =   2
            Left            =   8895
            RightToLeft     =   -1  'True
            TabIndex        =   188
            Top             =   855
            Width           =   735
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·—ﬁ„"
            Height          =   270
            Index           =   66
            Left            =   7890
            RightToLeft     =   -1  'True
            TabIndex        =   186
            Top             =   0
            Width           =   780
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "»‰«¡ ⁄·Ï"
            Height          =   195
            Index           =   65
            Left            =   9855
            RightToLeft     =   -1  'True
            TabIndex        =   185
            Top             =   30
            Width           =   1155
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·›—⁄"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   5880
            TabIndex        =   119
            Top             =   120
            Width           =   435
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "  Ê’Ê· «·‘Õ‰Â"
            Height          =   270
            Index           =   56
            Left            =   8355
            RightToLeft     =   -1  'True
            TabIndex        =   117
            Top             =   1320
            Width           =   1305
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·„‘—Ê⁄"
            Height          =   285
            Index           =   58
            Left            =   5625
            RightToLeft     =   -1  'True
            TabIndex        =   102
            Top             =   330
            Width           =   810
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬁÌ„… «·Œ’„"
            Height          =   300
            Index           =   11
            Left            =   1995
            RightToLeft     =   -1  'True
            TabIndex        =   99
            Top             =   840
            Width           =   900
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÿ—Ìﬁ… «·œ›⁄"
            Height          =   405
            Index           =   10
            Left            =   12495
            RightToLeft     =   -1  'True
            TabIndex        =   98
            Top             =   375
            Width           =   690
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «·⁄„·Ì…"
            Height          =   300
            Index           =   7
            Left            =   14445
            RightToLeft     =   -1  'True
            TabIndex        =   97
            Top             =   480
            Width           =   1305
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "««·⁄„·…"
            Height          =   270
            Index           =   9
            Left            =   2250
            RightToLeft     =   -1  'True
            TabIndex        =   96
            Top             =   120
            Width           =   705
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·›« Ê—…"
            Height          =   255
            Index           =   8
            Left            =   14070
            RightToLeft     =   -1  'True
            TabIndex        =   95
            Top             =   15
            Width           =   1680
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„Ê—œ «·⁄«„"
            Height          =   270
            Index           =   6
            Left            =   14070
            RightToLeft     =   -1  'True
            TabIndex        =   94
            Top             =   855
            Width           =   1680
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·Œ’„"
            Height          =   360
            Index           =   5
            Left            =   5280
            RightToLeft     =   -1  'True
            TabIndex        =   93
            Top             =   840
            Width           =   1185
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «·„Œ“‰"
            Height          =   210
            Index           =   4
            Left            =   14070
            RightToLeft     =   -1  'True
            TabIndex        =   92
            Top             =   1215
            Width           =   1680
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   55
            Left            =   315
            RightToLeft     =   -1  'True
            TabIndex        =   91
            Top             =   720
            Width           =   195
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "—ﬁ„ «–‰ «·«÷«›…"
            Height          =   315
            Index           =   52
            Left            =   2535
            RightToLeft     =   -1  'True
            TabIndex        =   90
            Top             =   1560
            Visible         =   0   'False
            Width           =   1320
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   435
         Index           =   3
         Left            =   15
         TabIndex        =   203
         TabStop         =   0   'False
         Top             =   7620
         Width           =   15840
         _cx             =   27940
         _cy             =   767
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
         BorderWidth     =   0
         ChildSpacing    =   0
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
         Begin VB.TextBox XPTxtSum 
            Alignment       =   2  'Center
            BackColor       =   &H000000FF&
            Height          =   390
            Left            =   3780
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   204
            TabStop         =   0   'False
            Top             =   -90
            Visible         =   0   'False
            Width           =   330
         End
         Begin MSDataListLib.DataCombo DCboUserName 
            Height          =   315
            Left            =   2760
            TabIndex        =   205
            Top             =   90
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label LblCommision 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   345
            Left            =   9480
            TabIndex        =   234
            Top             =   -120
            Width           =   990
         End
         Begin VB.Label LblCommisionV 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   345
            Left            =   9105
            RightToLeft     =   -1  'True
            TabIndex        =   233
            Top             =   0
            Width           =   960
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄„Ê·« "
            Height          =   255
            Index           =   70
            Left            =   10470
            RightToLeft     =   -1  'True
            TabIndex        =   232
            Top             =   120
            Width           =   630
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·”Ã·"
            Height          =   495
            Index           =   0
            Left            =   1755
            RightToLeft     =   -1  'True
            TabIndex        =   221
            Top             =   90
            Width           =   810
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«Ã„«·Ì «·ﬂ„ÌÂ"
            Height          =   300
            Index           =   63
            Left            =   4080
            TabIndex        =   220
            Top             =   135
            Visible         =   0   'False
            Width           =   750
         End
         Begin VB.Label LblTotalQty 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   345
            Left            =   1995
            TabIndex        =   219
            Top             =   0
            Visible         =   0   'False
            Width           =   1080
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·’«›Ì"
            Height          =   210
            Index           =   24
            Left            =   8145
            RightToLeft     =   -1  'True
            TabIndex        =   218
            Top             =   135
            Width           =   525
         End
         Begin VB.Label LblDiscountsTotal 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   345
            Left            =   11220
            RightToLeft     =   -1  'True
            TabIndex        =   217
            Top             =   0
            Width           =   975
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "Œ’Ê„« "
            Height          =   255
            Index           =   50
            Left            =   12300
            RightToLeft     =   -1  'True
            TabIndex        =   216
            Top             =   90
            Width           =   630
         End
         Begin VB.Label LblTotalAll 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   405
            Left            =   13050
            RightToLeft     =   -1  'True
            TabIndex        =   215
            Top             =   0
            Width           =   1470
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "/"
            Height          =   255
            Index           =   23
            Left            =   930
            RightToLeft     =   -1  'True
            TabIndex        =   214
            Top             =   90
            Width           =   180
         End
         Begin VB.Label LblTotal 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   405
            Left            =   6480
            RightToLeft     =   -1  'True
            TabIndex        =   213
            Top             =   15
            Width           =   1470
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„” Œœ„"
            Height          =   255
            Index           =   1
            Left            =   4875
            RightToLeft     =   -1  'True
            TabIndex        =   212
            Top             =   90
            Width           =   615
         End
         Begin VB.Label XPTxtCount 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Height          =   255
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   211
            Top             =   90
            Width           =   855
         End
         Begin VB.Label XPTxtCurrent 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Height          =   255
            Left            =   1155
            RightToLeft     =   -1  'True
            TabIndex        =   210
            Top             =   90
            Width           =   645
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·≈Ã„«·Ï"
            Height          =   255
            Index           =   3
            Left            =   14700
            RightToLeft     =   -1  'True
            TabIndex        =   209
            Top             =   90
            Width           =   735
         End
         Begin VB.Label LblTotalAllview 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   405
            Left            =   13035
            TabIndex        =   208
            Top             =   0
            Width           =   1440
         End
         Begin VB.Label LblDiscountsTotalview 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   345
            Left            =   11250
            TabIndex        =   207
            Top             =   0
            Width           =   990
         End
         Begin VB.Label LblTotalview 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   405
            Left            =   6495
            TabIndex        =   206
            Top             =   0
            Width           =   1470
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   540
         Index           =   1
         Left            =   15
         TabIndex        =   222
         TabStop         =   0   'False
         Top             =   8070
         Width           =   15840
         _cx             =   27940
         _cy             =   953
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
            Height          =   540
            Index           =   0
            Left            =   14145
            TabIndex        =   223
            Top             =   0
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   953
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
            Height          =   540
            Index           =   1
            Left            =   12360
            TabIndex        =   224
            Top             =   0
            Width           =   1620
            _ExtentX        =   2858
            _ExtentY        =   953
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
            Height          =   540
            Index           =   2
            Left            =   10575
            TabIndex        =   225
            Top             =   0
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   953
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
            Height          =   540
            Index           =   3
            Left            =   8865
            TabIndex        =   226
            Top             =   0
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   953
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
            Height          =   540
            Index           =   4
            Left            =   7065
            TabIndex        =   227
            Top             =   0
            Width           =   1650
            _ExtentX        =   2910
            _ExtentY        =   953
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
            Height          =   540
            Index           =   5
            Left            =   5310
            TabIndex        =   228
            Top             =   0
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   953
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
            Height          =   540
            Index           =   6
            Left            =   45
            TabIndex        =   229
            Top             =   0
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   953
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
            Height          =   540
            Index           =   7
            Left            =   3495
            TabIndex        =   230
            Top             =   0
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   953
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
            Height          =   540
            Left            =   1770
            TabIndex        =   231
            Top             =   0
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   953
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
   End
End
Attribute VB_Name = "FrmBillBuyComposite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim NewGrid As ClsGrid
Dim TTP As clstooltip
Dim BuyReport As ClsBuyReport
Dim cSearchDcbo(3) As clsDCboSearch

Public BolPrint As Boolean
Dim WithEvents m_MnuShowNewItemsPrices As Menu
Attribute m_MnuShowNewItemsPrices.VB_VarHelpID = -1
Dim WithEvents m_MenuViewList As Menu
Attribute m_MenuViewList.VB_VarHelpID = -1
Dim WithEvents m_MenuShowItemCostEffect As Menu
Attribute m_MenuShowItemCostEffect.VB_VarHelpID = -1
Dim WithEvents m_FrmSearch As Form
Attribute m_FrmSearch.VB_VarHelpID = -1
Dim bank_account As String
Dim general_noteid As Long
Dim CurrentVoucherNo As String
Dim CurrentVoucherSerialNo As String
Dim DateChanged As Boolean
Dim TxtNoteSerial1V As String

Public Sub RetriveSerials(ItemID As String, _
                          itemname As String, _
                          seriallist As String, _
                          currentrow As Long)
    Dim RsDetails As New ADODB.Recordset
    Dim strSQL As String
    Dim RsNotes As New ADODB.Recordset
    Dim RsTemp As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Dim Num As Long
    On Error GoTo ErrTrap
    Dim strInputString As String
    Dim strFilterText As String
    Dim astrSplitItems() As String
    Dim astrFilteredItems() As String
    Dim strFilteredString As String
    Dim intX As Integer
    strInputString = seriallist
    strFilterText = ","
 
    astrSplitItems = Split(strInputString, strFilterText)
    Dim i As Integer
    ' For i = 1 To Fg.Rows - 2
    '        If Fg.TextMatrix(i, Fg.ColIndex("Code")) = ItemID Then
    '         Me.Fg.RemoveItem (i)
    '         i = 1
    '        End If
    'NewGrid.Grid_AfterEdit Num, Fg.ColIndex("Code")
    ' Next i
   
    Num = currentrow

    '  For Num = currentrow To UBound(astrSplitItems)+currentrow
    For intX = 0 To UBound(astrSplitItems)
   
        FG.TextMatrix(Num, FG.ColIndex("Code")) = ItemID
        NewGrid.Grid_AfterEdit Num, FG.ColIndex("Code")
        ' FG.TextMatrix(Num, FG.ColIndex("Name")) = itemname
        FG.TextMatrix(Num, FG.ColIndex("Count")) = 1
        FG.TextMatrix(Num, FG.ColIndex("Serial")) = astrSplitItems(intX)
  
        '      RsDetails.MoveNext
        '      Debug.Print Num
        FG.Rows = FG.Rows + 1
 
        Num = Num + 1
    Next
 
    TxtFillData.text = "F"
    TxtFillData_Change
    Exit Sub
ErrTrap:
    Screen.MousePointer = vbDefault
End Sub

Function CheckMyData() As Boolean
    CheckMyData = True

    If TxtNoteSerial.text = "" Then
        If Notes_coding(val(my_branch), XPDtbBill.value) = "error" Then
            MsgBox " ·« Ì„ﬂ‰ «÷«›… ﬁÌÊœ ·Â–… «·⁄„·Ì… ·«‰ﬂ  ⁄œÌ  «·Õœ «·«ﬁ’Ì ··ﬁÌÊœ ﬂ„« Õœœ  ›Ì  —ﬁÌ„ «·”‰œ«  ": GoTo ErrTrap
        Else
                       
            If Notes_coding(val(my_branch), XPDtbBill.value) = "" Then
                MsgBox " ·«Ì„ﬂ‰ «‰‘«¡ «·ﬁÌœ ·Â–« «·„” ‰œ ·«‰ﬂ Õœœ   —ﬁÌ„ ﬁÌÊœ ÌœÊÌ  ": GoTo ErrTrap
            Else
                TxtNoteSerial.text = Notes_coding(val(my_branch), XPDtbBill.value)
            End If
        End If
    End If
        
    If TxtNoteSerial1.text = "" Then
        If Voucher_coding(val(my_branch), XPDtbBill.value, 6, 150, , 22) = "error" Then
            MsgBox " ·« Ì„ﬂ‰ «÷«›… ”‰œ ›« Ê—… „‘ —Ì«  ÃœÌœ ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": GoTo ErrTrap
        Else
                       
            If Voucher_coding(val(my_branch), XPDtbBill.value, 6, 150, , 22) = "" Then
                MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ ›« Ê—… „‘ —Ì«   ÌœÊÌ« ﬂ„« Õœœ   ": GoTo ErrTrap
            Else
                TxtNoteSerial1.text = Voucher_coding(val(my_branch), XPDtbBill.value, 6, 150, , 22)
            End If
        End If
    End If
 
    If BillBasedOn(0).value = True Then

        If Voucher_coding(val(my_branch), XPDtbBill.value, 9, 160, , 20) = "" Then
                                
            If Trim$(TxtManualNo1) = "" Then
                MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ ”‰œ «·«” ·«„ ÌœÊÌ« Õ Ï Ì „ «‰‘«¡ «·”‰œ  ":  GoTo ErrTrap
            
            Else
                TxtNoteSerial1V = TxtManualNo1
            End If
            
        End If
                       
    End If

    Exit Function
ErrTrap:
    CheckMyData = False
End Function

Public Sub Convert()
    Cmd_Click (0)
End Sub

Public Sub Cala()
    NewGrid.Calculate 1, , , True
End Sub

Private Sub BillBasedOn_Click(Index As Integer)

    Select Case Index

        Case 0

            If BillBasedOn(0).value = True Then
                
                FillVoucherGrid (0)
                GRID1.Enabled = True
            End If

        Case 1

            If BillBasedOn(1).value = True Then
                
                FillVoucherGrid (1)
                GRID1.Enabled = True
            End If

        Case 2

            If BillBasedOn(2).value = True Then
                
                '            FillOrderGrid
                '            GRID2.Enabled = True
            End If

    End Select

End Sub

Function FillVoucherGrid(Optional OPtype As Integer = 0)
    ' ⁄»∆…  ”‰œ«   «·’—›
    On Error Resume Next

    With Me.GRID1
        .Rows = .FixedRows
        .ExtendLastCol = True
        .RowHeightMin = 300
        .Editable = flexEDKbdMouse
        .ExplorerBar = flexExSortShowAndMove

        '    .AutoSize 0, .Cols - 1, False
    End With

    Dim i As Integer
    Dim RsExp As ADODB.Recordset
    Dim My_SQL As String

    Set RsExp = New ADODB.Recordset

    'My_SQL = "SELECT dbo.Transactions.closed,dbo.Transactions.NoteSerial1, dbo.Transactions.Transaction_ID,dbo.Transactions.order_no , dbo.Transactions.Transaction_Date,dbo.Transactions.CusID, dbo.TblCustemers.CusName FROM dbo.Transactions  INNER JOIN dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID Where  Transaction_Type=20   and   dbo.TblCustemers.CusID=" & Val(DBCboClientName.BoundText)
    If OPtype = 0 Then
        My_SQL = "SELECT dbo.Transactions.closed,dbo.Transactions.NoteSerial1,dbo.Transactions.NoteSerial, dbo.Transactions.Transaction_ID,dbo.Transactions.order_no , dbo.Transactions.Transaction_Date,dbo.Transactions.CusID, dbo.TblCustemers.CusName FROM dbo.Transactions  INNER JOIN dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID Where  Transaction_ID= " & val(Text1.text)
    Else
        'My_SQL = "SELECT dbo.Transactions.closed,dbo.Transactions.NoteSerial1,dbo.Transactions.NoteSerial, dbo.Transactions.Transaction_ID,dbo.Transactions.order_no , dbo.Transactions.Transaction_Date,dbo.Transactions.CusID, dbo.TblCustemers.CusName FROM dbo.Transactions  INNER JOIN dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID Where  Transaction_ID= " & Val(Text1.text)
        My_SQL = "SELECT dbo.Transactions.closed,dbo.Transactions.NoteSerial1,dbo.Transactions.NoteSerial, dbo.Transactions.Transaction_ID,dbo.Transactions.order_no , dbo.Transactions.Transaction_Date,dbo.Transactions.CusID, dbo.TblCustemers.CusName FROM dbo.Transactions  INNER JOIN dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID Where   ( (nots='" & Me.XPTxtBillID.text & "' and  Transaction_Type=20) or ( Transaction_Type=20   and  closed =0 and (nots='' or nots is null) ) and  (dbo.TblCustemers.CusID=" & val(DBCboClientName.BoundText) & ")) "
    End If

    RsExp.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText

    With Me.GRID1
        .Rows = 2
        .Clear flexClearScrollable

        If RsExp.RecordCount > 0 Then
            .Rows = RsExp.RecordCount + 1
            RsExp.MoveFirst

            For i = 1 To .Rows - 1
             
                .TextMatrix(i, .ColIndex("Select")) = IIf(IsNull(RsExp.Fields("closed").value), 0, RsExp.Fields("closed").value)
              
                .TextMatrix(i, .ColIndex("NoteSerial")) = IIf(IsNull(RsExp.Fields("NoteSerial").value), "", RsExp.Fields("NoteSerial").value)
              
                .TextMatrix(i, .ColIndex("NoteSerial1")) = IIf(IsNull(RsExp.Fields("NoteSerial1").value), "", RsExp.Fields("NoteSerial1").value)
               
                .TextMatrix(i, .ColIndex("Transaction_Date")) = IIf(IsNull(RsExp.Fields("Transaction_Date").value), "", RsExp.Fields("Transaction_Date").value)
           
                .TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(RsExp.Fields("CusName").value), "", RsExp.Fields("CusName").value)
                .TextMatrix(i, .ColIndex("Transaction_ID")) = IIf(IsNull(RsExp.Fields("Transaction_ID").value), "", RsExp.Fields("Transaction_ID").value)

                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("P1")) = "⁄—÷ «·”‰œ"
                    .TextMatrix(i, .ColIndex("P2")) = "ÿ»«⁄Â  «·ﬁÌœ"
                Else
                    .TextMatrix(i, .ColIndex("P1")) = "View VCHR"
                    .TextMatrix(i, .ColIndex("P2")) = "Print GE"

                End If

                RsExp.MoveNext
            Next

            RsExp.Close
        End If

        .RowHeight(-1) = 300
    End With

    GRID1.Visible = True

End Function

Function CloseIssueVoucher()
    On Error Resume Next
    Dim i As Integer
    Dim sql As String
 
    If BillBasedOn(1).value = False Then Exit Function

    With GRID1

        For i = 1 To .Rows - 1
     
            If .Cell(flexcpChecked, i, .ColIndex("select")) = flexChecked Then
           
                sql = "update transactions set closed=1" & ",nots=" & val(Me.XPTxtBillID.text) & ",nots2=" & Me.TxtNoteSerial1.text & " where  Transaction_ID= " & val(.TextMatrix(i, .ColIndex("Transaction_ID")))
            Else
                sql = "update transactions set closed=0 ,nots='' ,nots2='' where  Transaction_ID=" & val(.TextMatrix(i, .ColIndex("Transaction_ID"))) & "nots=" & "" & "nots2=" & ""
               
            End If
       
            Cn.Execute sql
 
        Next
       
    End With
       
End Function

Private Sub CBoBasedON_Change()

    If Me.CBoBasedON.ListIndex = 0 Then

    ElseIf Me.CBoBasedON.ListIndex = 1 Then

        If SystemOptions.UserInterface = ArabicInterface Then
            lbl(66).Caption = "—ﬁ„ «·«„—  "
        Else
            lbl(66).Caption = "Order NO"
        End If

    ElseIf Me.CBoBasedON.ListIndex = 2 Then

        If SystemOptions.UserInterface = ArabicInterface Then
            lbl(66).Caption = "—ﬁ„Â«"
        Else
            lbl(66).Caption = "NO:"
        End If
    End If

    If Txt_order_no.text <> "" Then
        Txt_order_no_Change
    End If

End Sub

Private Sub CBoBasedON_Click()
    CBoBasedON_Change
End Sub

Private Sub ChkInstall_Click()

    If ChkInstall.value = vbChecked Then
        Me.CmdINSTALLMENT.Enabled = True
    Else
        Me.CmdINSTALLMENT.Enabled = False

        With Me.FgInstallments
            .Clear flexClearScrollable, flexClearEverything
            .Rows = .FixedRows
            LblPrecenType.Caption = ""
            LblPrecenValue.Caption = ""
            LblInstallTotal.Caption = ""
            LblInstallCount.Caption = ""
            LblFirstInstallDate.Caption = ""
            LblInstallmentType.Caption = ""
        End With

    End If

End Sub

Private Sub ChkTaxAdd_Click()

    If ChkTaxAdd.value = Checked Then
        TxtTaxAddValue.Enabled = True
        lbl(39).Enabled = True
        lbl(46).Enabled = True
    Else
        TxtTaxAddValue.text = ""
        TxtTaxAddValue.Enabled = False
        lbl(39).Enabled = False
        lbl(46).Enabled = False
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub ChkTaxSerivce_Click()
    On Error GoTo ErrTrap

    If ChkTaxSerivce.value = Checked Then
        TxtTaxServiceValue.Enabled = True
        lbl(43).Enabled = True
        lbl(47).Enabled = True
    Else
        TxtTaxServiceValue.text = ""
        TxtTaxServiceValue.Enabled = False
        lbl(43).Enabled = False
        lbl(47).Enabled = False
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub ChkTaxStamp_Click()

    If ChkTaxStamp.value = Checked Then
        TxtTaxStampValue.Enabled = True
        lbl(41).Enabled = True
        lbl(48).Enabled = True
    Else
        TxtTaxStampValue.text = ""
        TxtTaxStampValue.Enabled = False
        lbl(41).Enabled = False
        lbl(48).Enabled = False
    End If

    Exit Sub
ErrTrap:
End Sub

Function RemoveFactoryExpenses()

    With Me.Fg_Journal
  
        If .Row <= 0 Then Exit Function
        .RemoveItem .Row
    End With

    With Me.Fg_Journal
        Me.TXTFactoryExpenses.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .Rows - 1, .ColIndex("Value"))
    End With

    ReLineGrid

End Function

Private Sub Cmd_Click(Index As Integer)
    On Error GoTo ErrTrap
    Dim AskOption As Boolean
    Dim intDef As Integer
    Dim Msg As String

    BolPrint = True
 
    Select Case Index
    
        Case 0

            If DoPremis(Do_New, Me.name, True) = False Then
                Exit Sub
            End If

            Command2.Enabled = True
            Txt_EXport.Enabled = True
            '  Grid.Visible = True
            clear_all Me
            TxtModFlg.text = "N"
            ' Me.TxtNoteID.text = CStr(new_id("Notes", "NoteID", "", True))
            BillBasedOn(0).value = True
            XPTxtBillID.text = CStr(new_id("Transactions", "Transaction_ID", "", True))

            If BillType = 22 Then
                TxtTransSerial.text = CStr(new_id("Transactions", "Transaction_Serial", "", True, "Transaction_Type=22"))
            End If
        
            If BillType = 1 Then
                TxtTransSerial.text = CStr(new_id("Transactions", "Transaction_Serial", "", True, "Transaction_Type=1"))
            End If

            '      TxtTransSerial.text = CStr(new_id("Transactions", "Transaction_Serial", "",  True  )
            SetDefaults
            NewGrid.GridDefaultValue 1
            Me.DCboUserName.BoundText = user_id
            intDef = GetSetting(StrAppRegPath, "DefaultOptions", "DefaultSup", 1)
            DBCboClientName.BoundText = intDef
            intDef = GetSetting(StrAppRegPath, "DefaultOptions", "DefaultPurchaseStore", 1)
            DCboStoreName.BoundText = intDef
            XPTab301.CurrTab = 0
            '        FG.SetFocus
            FG.Col = FG.ColIndex("Code")
            FG.Row = FG.Rows - 1
            Command2_Click
            
            
Dim dstore As Integer
            Dim dBox As Integer
            Dim usertype As Integer
            Dim empid As Integer
            Dim userbranchid As Integer
            'GetBranchData branch_id, dstore, dBox
                 
            GetUserData user_id, usertype, userbranchid, dstore, dBox, , empid
     
            If usertype <> 0 Then 'admin
                dcBranch.Enabled = False
 
                DCboStoreName.Enabled = False
              '  TxtStoreID.Enabled = False
                Me.DCboStoreName.BoundText = dstore
            Else
                dcBranch.Enabled = True
 
                DCboStoreName.Enabled = True
 
                Me.dcBranch.BoundText = ""
                Me.DCboStoreName.BoundText = ""
'                TxtStoreID.Enabled = True
            End If
                    
                    
        

      If SystemOptions.usertype <> UserAdminAll Then
                            If checkmanyBranches = False Then
                                   Me.dcBranch.Enabled = False
                                   Else
                                    Me.dcBranch.Enabled = True
                             End If
                    
                      If checkmanyStores = False Then
                                   Me.DCboStoreName.Enabled = False
                                    
                                   Else
                                   Me.DCboStoreName.Enabled = True
 
                             End If
                                  
           End If

            
            Me.dcBranch.BoundText = Current_branch
            Me.CBoBasedON.ListIndex = 0
            Fg_Journal.Clear flexClearScrollable, flexClearEverything
            Fg_Journal.Rows = 2
            Fg_Journal.Enabled = True
            GRID1.Rows = 1
            GRID1.Enabled = True
          
            Dccurrency.BoundText = 1
 
        Case 1

            If DoPremis(Do_Edit, Me.name, True) = False Then
                Exit Sub
            End If

            If SystemOptions.usertype = UserNormal Then
                If AvailableDeal = False Then
                    Exit Sub
                End If
            End If

            TxtModFlg.text = "E"
            Me.DCboUserName.BoundText = user_id
        
            Fg_Journal.Rows = Fg_Journal.Rows + 1
            Fg_Journal.Enabled = True
            DateChanged = False
            CuurentLogdata
    
        Case 2
    
            If Trim(dcBranch.BoundText) = "" Then
                If SystemOptions.UserInterface = EnglishInterface Then
                    Msg = "Specify Departement"
                Else
                    Msg = "Õœœ «·›—⁄ «Ê·«  "
                End If

                MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                dcBranch.SetFocus
                SendKeys "{F4}"
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
 
            my_branch = Me.dcBranch.BoundText

            '   If Me.TxtModFlg.text = "N" Then
             
            ' End If
      
            SaveData

        Case 3
            Undo

        Case 4

            If DoPremis(Do_Delete, Me.name, True) = False Then
                Exit Sub
            End If

            If SystemOptions.usertype = UserNormal Then
                Msg = "·Ì” ·ﬂ Õﬁ Õ–› ›Ï «·›Ê« Ì—"
                MsgBox Msg, vbOKOnly + vbMsgBoxRight + vbExclamation + vbMsgBoxRtlReading, App.title
                Exit Sub
            End If

            Del_TransAction

        Case 5

            If DoPremis(Do_Search, Me.name, True) = False Then
                Exit Sub
            End If

            If m_FrmSearch Is Nothing Then
                Set m_FrmSearch = New FrmBuySearch
                m_FrmSearch.DealingForm = PurchaseTransaction
            
                If SystemOptions.UserInterface = ArabicInterface Then
                    m_FrmSearch.Caption = "Search About Purchase Invoice"
                Else
            
                End If

                Set m_FrmSearch.RetrunFrm = Me
                m_FrmSearch.show vbModal
            Else
                Msg = "Â‰«ﬂ ‘«‘… »ÕÀ Œ«’À… »‘«‘… ›« Ê—… «·‘—«¡ «·Õ«·Ì…"
                Msg = Msg & Chr(13) & "Ÿ«Â—… «„«„ﬂ ›⁄·«...·«Ì„ﬂ‰ ⁄—÷ «ﬂÀ— „‰ ‘«‘… »ÕÀ ·ﬂ· ‘«‘… ›« Ê—…"
                MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                m_FrmSearch.Visible = True
                m_FrmSearch.ZOrder 0
                m_FrmSearch.SetFocus
            End If

        Case 6
            Unload Me

        Case 7

            If DoPremis(Do_Print, Me.name, True) = False Then
                Exit Sub
            End If

            AskOption = GetSetting(StrAppRegPath, "View_Type", "ShowMe", False)

            If AskOption = False Then
                FrmPrintOptions.show vbModal
            End If

            If BolPrint = False Then
                Exit Sub
            End If
        
            printing
        
        Case 9
            RemoveFactoryExpenses

        Case 10
            ShowGL_cc TxtNoteSerial.text, , 200
        
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub CmdAttach_Click()
     On Error Resume Next
ShowAttachments TxtNoteSerial1, "0812201404"

End Sub

Private Sub CmdCheque_Click()

    If Me.TxtModFlg.text = "R" Then
        Exit Sub
    End If

    Load FrmChecks
    FrmChecks.TxtModFlg.text = Me.TxtModFlg.text
    FrmChecks.XPTxtBillID.text = Me.XPTxtBillID.text
    Set FrmChecks.PutFg = Me.FgCheques
    FrmChecks.show vbModal
    SumChecks

End Sub

Private Sub SumChecks()

    With Me.FgCheques

        If .Rows > 1 Then
            Me.lbl(19).Caption = .Aggregate(flexSTCount, .FixedRows, .ColIndex("CheckNumber"), .Rows - 1, .ColIndex("CheckNumber"))
            Me.lbl(18).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("CheckValue"), .Rows - 1, .ColIndex("CheckValue"))
        Else
            Me.lbl(19).Caption = 0
            Me.lbl(18).Caption = 0
        End If

    End With

End Sub

Private Sub CmdHelp_Click()
    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
    SystemOptions.SysHelp.HHDisplayTopicID Me.hwnd
End Sub

Private Sub CmdInfo_Click()
    Me.PopupMenu mdifrmmain.MnuInvPurchase
End Sub

Private Sub CmdINSTALLMENT_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
    Dim i As Integer

    If XPTxtValue(1).text = "" Then
        Msg = "ÌÃ»  ÕœÌœ «·ﬁÌ„… «·¬Ã·… ﬁ»·  ”ÃÌ· «·√ﬁ”«ÿ"
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title

        If XPTxtValue(1).Enabled = True Then
            XPTxtValue(1).SetFocus
        End If

        Exit Sub
    End If

    Load FrmInstallMent
    Set FrmInstallMent.Frm = Me

    With FrmInstallMent

        If Me.TxtModFlg.text = "R" Then
            .Tag = "R"
            .Retrive val(XPTxtValue(1).Tag)
        Else
            .Tag = "N"
            .Txt(1).text = XPTxtValue(1).text
            .LblNoteID.Caption = XPTxtSerial(1).text
            .CboPrecenType.ListIndex = val(Me.LblPrecenType.Tag)
            .Txt(3).text = val(LblPrecenValue.Caption)
            .Txt(5).text = val(LblInstallCount.Caption)

            If IsDate(Me.LblFirstInstallDate.Caption) Then
                .Dtp_First.value = Me.LblFirstInstallDate.Caption
            End If

            .Txt(7).text = val(LblInstallSeprator.Caption)

            If val(LblInstallmentType.Tag) = 0 Then
                .OptInt(0).value = True
            ElseIf val(LblInstallmentType.Tag) = 1 Then
                .OptInt(1).value = True
            ElseIf val(LblInstallmentType.Tag) = 2 Then
                .OptInt(2).value = True
            End If

            With .FG
                .Rows = Me.FgInstallments.Rows

                For i = 1 To Me.FgInstallments.Rows - 1
                    .TextMatrix(i, .ColIndex("Serial")) = i
                    .TextMatrix(i, .ColIndex("Value")) = Me.FgInstallments.TextMatrix(i, Me.FgInstallments.ColIndex("Value"))
                    .TextMatrix(i, .ColIndex("Due_Date")) = Me.FgInstallments.TextMatrix(i, Me.FgInstallments.ColIndex("Due_Date"))
                Next i

                .AutoSize 0, .Cols - 1, False
            End With

        End If

        .show vbModal
    End With

    Exit Sub
ErrTrap:
End Sub

Private Sub CmdNotes_Click()
    ShowRelatedNotes val(Me.XPTxtBillID.text), 1
End Sub

Private Sub CmdNotes_MouseMove(Button As Integer, _
                               Shift As Integer, _
                               X As Single, _
                               Y As Single)
    Dim StrTemp As String

    If val(Me.CmdNotes.Tag) = 0 Then
        Me.CmdNotes.ToolTipText = ""
    Else
        StrTemp = " ÊÃœ ⁄·Ï Â–Â «·Õ—ﬂ… ⁄„·Ì«  „«·Ì… „ﬁœ«—Â« : " & val(Me.CmdNotes.Tag)
        Me.CmdNotes.ToolTipText = StrTemp
    End If

End Sub

Private Sub CmdRetruns_MouseMove(Button As Integer, _
                                 Shift As Integer, _
                                 X As Single, _
                                 Y As Single)
    Dim StrTemp As String

    If val(Me.CmdRetruns.Tag) = 0 Then
        Me.CmdRetruns.ToolTipText = ""
    Else
        StrTemp = " ÊÃœ ⁄·Ï Â–Â «·Õ—ﬂ… Õ—ﬂ«   Ã«—Ì… √Œ—Ï ·Â« ⁄·«ﬁ… »Â« ≈Ã„«·ÌÂ«: " & val(Me.CmdRetruns.Tag)
        Me.CmdRetruns.ToolTipText = StrTemp
    End If

End Sub

Function CREATE_VOUCHER_GE(Transaction_ID As Long, TxtNoteSerialV As String, TxtNoteSerial1V As String, general_noteid As Long, BranchId As Integer)
    Dim LngDevID As Long
    Dim LngDevNO  As Integer
    Dim StrTempAccountCode As String
    Dim StrTempDes As String
    Dim SngTemp  As Variant
    Dim Account_Code_dynamic As String
    Dim i As Integer
    Dim total_shahn As Single
        
    Dim usedaccount As Integer

    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
    '«·ÿ—› «·„œÌ‰
    SngTemp = ((NewGrid.GetItemsTotal(ItemsGoodType) - val(LblDiscountsTotal.Caption)) * val(txt_Currency_rate.text) + val(TXTToTAlELSHahn.text))

    If SngTemp > 0 Then
        If detect_inventory_work_type = 1 Then

            Account_Code_dynamic = get_account_code_branch(0, my_branch)
        
            If Account_Code_dynamic = "NO branch" Then
                MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                GoTo ErrTrap
            Else

                If Account_Code_dynamic = "NO account" Then
                    MsgBox "·„ Ì „  ÕœÌœ Õ”«»  «·„Œ“Ê‰ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                    GoTo ErrTrap
         
                End If
            End If

            If val(DCDocTypes.BoundText) > 0 Then
                getDocAccounts val(DCDocTypes.BoundText), , , , StrTempAccountCode, , , , , usedaccount

                If StrTempAccountCode = "" And usedaccount = 1 Then
                    MsgBox "ÌÊÃœ Œÿ√ ›Ì «·Õ”«» «·„œÌ‰ ·”‰œ «·«” ·«„", vbCritical
                    GoTo ErrTrap
                ElseIf StrTempAccountCode <> "" And usedaccount = 1 Then
        
                ElseIf usedaccount = 0 Then
        
                    StrTempAccountCode = Account_Code_dynamic '„Œ“Ê‰ «·»÷«⁄…
                End If

            Else
                StrTempAccountCode = Account_Code_dynamic '„Œ“Ê‰ «·»÷«⁄…
            End If

            ' StrTempAccountCode = "a1a2a5" '„Œ“Ê‰ «·»÷«⁄…
            If SystemOptions.UserInterface = ArabicInterface Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrTempDes = "”‰œ «” ·«„ —ﬁ„ " & TxtNoteSerial1V & " »‰«¡ ⁄·Ï ›« Ê—… „‘ —Ì«  —ﬁ„ " & Me.TxtNoteSerial1.text
                Else
                    StrTempDes = "Ò Recieve Voucher No. " & TxtNoteSerial1V & " Based On Purchase Invoice NO:" & Me.TxtNoteSerial1.text
                End If
            
            Else
                StrTempDes = "Ò Recieve Voucher No. " & TxtNoteSerial1V
            End If

            LngDevNO = LngDevNO + 1

            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, 0, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Transaction_ID, , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                GoTo ErrTrap
            End If

        ElseIf detect_inventory_work_type = 2 Then
            '«·„Œ“Ê‰ «·”·⁄Ì ⁄·Ï „” ÊÏ «·„Œ“‰
    
            If val(DCDocTypes.BoundText) > 0 Then
                getDocAccounts val(DCDocTypes.BoundText), , , , StrTempAccountCode, , , , , usedaccount

                If StrTempAccountCode = "" And usedaccount = 1 Then
                    MsgBox "ÌÊÃœ Œÿ√ ›Ì «·Õ”«» «·„œÌ‰ ·”‰œ «·«” ·«„", vbCritical
                    GoTo ErrTrap
                ElseIf StrTempAccountCode <> "" And usedaccount = 1 Then
                    Account_Code_dynamic = StrTempAccountCode
                ElseIf usedaccount = 0 Then
        
                    Account_Code_dynamic = get_store_Account(DCboStoreName.BoundText, "Account_Code")
                End If

            Else
                Account_Code_dynamic = get_store_Account(DCboStoreName.BoundText, "Account_Code")
            End If

            If Account_Code_dynamic = "" Then
                MsgBox "·„ Ì „  ÕœÌœ Õ”«»  ··„Œ“Ê‰ «·”·⁄Ì ·Â–« «·„Œ“‰ ›Ì Â–« «·›—⁄    ", vbCritical
                GoTo ErrTrap
            End If
    
            StrTempAccountCode = Account_Code_dynamic  '„Õ“Ê‰ «·”·⁄Ì ··„Œ“‰

            ' StrTempAccountCode = "a1a2a5" '„Õ“Ê‰ «·»÷«⁄…
            If SystemOptions.UserInterface = ArabicInterface Then
                StrTempDes = "”‰œ «” ·«„ —ﬁ„ " & TxtNoteSerial1V & " »‰«¡ ⁄·Ï ›« Ê—… „‘ —Ì«  —ﬁ„ " & Me.TxtNoteSerial1.text
            Else
                StrTempDes = "Ò Recieve Voucher No. " & TxtNoteSerial1V & " Based On Purchase Invoice NO:" & Me.TxtNoteSerial1.text
            End If
            
            LngDevNO = LngDevNO + 1

            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, 0, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Transaction_ID, , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                GoTo ErrTrap
            End If

        ElseIf detect_inventory_work_type = 3 Then
            Dim groupAccount As String
             
            Dim line_value As Single

            With FG

                For i = 1 To FG.Rows - 1

                    If FG.TextMatrix(i, FG.ColIndex("Code")) <> "" Then
    
                        ' groupAccount = get_item_group_account(FG.TextMatrix(i, FG.ColIndex("Code")), DCboStoreName.BoundText, 2)
                        groupAccount = get_item_group_account_inventory(FG.TextMatrix(i, FG.ColIndex("Code")), DCboStoreName.BoundText, 0)

                        If groupAccount = "Error" Then
                            If SystemOptions.UserInterface = ArabicInterface Then
                                MsgBox "«·’‰› ›Ì «·”ÿ— —ﬁ„" & i & " €Ì— „Õœœ —ﬁ„ Õ”«»  «·„Œ“Ê‰ «·”⁄·⁄Ì ··„Œ“‰ «·„Õœœ   ·„Ã„Ê⁄ …"
                            Else
                                MsgBox "Item in line no " & i & "Group Name Account Not Defined"
                            End If

                            GoTo ErrTrap
                        End If

                        line_value = 0

                        line_value = FG.TextMatrix(i, FG.ColIndex("Price")) * FG.TextMatrix(i, FG.ColIndex("Count")) * val(txt_Currency_rate.text)
    
                        total_shahn = Round((((line_value) / (val(LblTotal.Caption) * val(txt_Currency_rate.text))) * val(TXTToTAlELSHahn.text)), 2)  'ﬁÌ„… «Ã„«·Ì ‘Õ‰ ”ÿ—
                        line_value = line_value + total_shahn + val(FG.TextMatrix(i, FG.ColIndex("LineShahn")))
                        line_value = Round(line_value, 2)
     
                        If SystemOptions.UserInterface = ArabicInterface Then
                            StrTempDes = "”‰œ «” ·«„ —ﬁ„ " & TxtNoteSerial1V & " »‰«¡ ⁄·Ï ›« Ê—… „‘ —Ì«  —ﬁ„ " & Me.TxtNoteSerial1.text
                        Else
                            StrTempDes = "Ò Recieve Voucher No. " & TxtNoteSerial1V & " Based On Purchase Invoice NO:" & Me.TxtNoteSerial1.text
                        End If
   
                        LngDevNO = LngDevNO + 1

                        If ModAccounts.AddNewDev(LngDevID, LngDevNO, groupAccount, line_value, 0, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Transaction_ID, , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                            GoTo ErrTrap
                        End If
    
                    End If

                Next i

            End With

        End If

        '«·ÿ—› «·œ«∆‰
        SngTemp = (NewGrid.GetItemsTotal(ItemsGoodType) - val(Me.LblDiscountsTotal.Caption)) * val(txt_Currency_rate.text) '+ Val(TXTToTAlELSHahn.text)

        If SngTemp > 0 Then
            If detect_inventory_work_type = 1 Or detect_inventory_work_type = 2 Then

                Account_Code_dynamic = get_account_code_branch(4, my_branch)
        
                If Account_Code_dynamic = "NO branch" Then
                    MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                    GoTo ErrTrap
                Else

                    If Account_Code_dynamic = "NO account" Then
                        MsgBox "·„ Ì „  ÕœÌœ Õ”«»  «·„‘ —Ì«  ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                        GoTo ErrTrap
         
                    End If
                End If

                If val(DCDocTypes.BoundText) > 0 Then
                    getDocAccounts val(DCDocTypes.BoundText), , , , , StrTempAccountCode, , , , , usedaccount

                    If StrTempAccountCode = "" And usedaccount = 1 Then
                        MsgBox "ÌÊÃœ Œÿ√ ›Ì «·Õ”«» «·œ«∆‰ ·”‰œ «·«” ·«„", vbCritical
                        GoTo ErrTrap
                    ElseIf StrTempAccountCode <> "" And usedaccount = 1 Then
        
                    ElseIf usedaccount = 0 Then
        
                        StrTempAccountCode = Account_Code_dynamic '«·„‘ —Ì« 
                    End If

                Else
                    StrTempAccountCode = Account_Code_dynamic '«·„‘ —Ì« 
                End If
            
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrTempDes = "”‰œ «” ·«„ —ﬁ„ " & TxtNoteSerial1V & " »‰«¡ ⁄·Ï ›« Ê—… „‘ —Ì«  —ﬁ„ " & Me.TxtNoteSerial1.text
                Else
                    StrTempDes = "Ò Recieve Voucher No. " & TxtNoteSerial1V & " Based On Purchase Invoice NO:" & Me.TxtNoteSerial1.text
                End If
            
                LngDevNO = LngDevNO + 1

                If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, 1, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Transaction_ID, , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                    GoTo ErrTrap
                End If
         
            ElseIf detect_inventory_work_type = 3 Then

                With FG

                    For i = 1 To FG.Rows - 1

                        If FG.TextMatrix(i, FG.ColIndex("Code")) <> "" Then
    
                            groupAccount = get_item_group_account_in_branch(FG.TextMatrix(i, FG.ColIndex("Code")), val(my_branch), 4)

                            '  groupAccount = get_item_group_account_inventory(FG.TextMatrix(I, FG.ColIndex("Code")), DCboStoreName.BoundText, 4)
                            If groupAccount = "Error" Then
                                If SystemOptions.UserInterface = ArabicInterface Then
                                    MsgBox "«·’‰› ›Ì «·”ÿ— —ﬁ„" & i & " €Ì— „Õœœ —ﬁ„ Õ”«»   «·„‘ —Ì«    ··„Œ“‰ «·„Õœœ   ·„Ã„Ê⁄ …"
                                Else
                                    MsgBox "Item in line no " & i & "Group Name Account Not Defined"
                                End If

                                GoTo ErrTrap
                            End If

                            line_value = 0
                            line_value = FG.TextMatrix(i, FG.ColIndex("Price")) * FG.TextMatrix(i, FG.ColIndex("Count")) * val(txt_Currency_rate.text)
                            '  total_shahn = Round((((line_value) / (Val(LblTotal.Caption) * Val(txt_Currency_rate.text))) * Val(TXTToTAlELSHahn.text)), 2)  'ﬁÌ„… «Ã„«·Ì ‘Õ‰ ”ÿ—
                            '  line_value = line_value + total_shahn + Val(FG.TextMatrix(I, FG.ColIndex("LineShahn")))
                            line_value = Round(line_value, 2)
     
                            If SystemOptions.UserInterface = ArabicInterface Then
                                StrTempDes = "”‰œ «” ·«„ —ﬁ„ " & TxtNoteSerial1V & " »‰«¡ ⁄·Ï ›« Ê—… „‘ —Ì«  —ﬁ„ " & Me.TxtNoteSerial1.text
                            Else
                                StrTempDes = "Ò Recieve Voucher No. " & TxtNoteSerial1V & " Based On Purchase Invoice NO:" & Me.TxtNoteSerial1.text
                            End If
            
                            LngDevNO = LngDevNO + 1

                            If ModAccounts.AddNewDev(LngDevID, LngDevNO, groupAccount, line_value, 1, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Transaction_ID, , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                                GoTo ErrTrap
                            End If
    
                        End If

                    Next i

                End With

            End If
        End If

        'ﬁÌœ «·„’—Ê›« 
        Dim Account_Code As String
        Dim Note_Value As Single

        With Grid

            For i = 1 To Grid.Rows - 1

                If .Cell(flexcpChecked, i, .ColIndex("select")) = flexChecked Then
            
                    If SystemOptions.UserInterface = ArabicInterface Then
                        StrTempDes = "”‰œ «” ·«„ —ﬁ„ " & TxtNoteSerial1V & " »‰«¡ ⁄·Ï ›« Ê—… „‘ —Ì«  —ﬁ„ " & Me.TxtNoteSerial1.text
                    Else
                        StrTempDes = "Ò Recieve Voucher No. " & TxtNoteSerial1V & " Based On Purchase Invoice NO:" & Me.TxtNoteSerial1.text
                    End If
            
                    LngDevNO = LngDevNO + 1
                    Account_Code = Grid.TextMatrix(i, Grid.ColIndex("Account_code"))
                    Note_Value = Grid.TextMatrix(i, Grid.ColIndex("Note_value"))

                    If ModAccounts.AddNewDev(LngDevID, LngDevNO, Account_Code, Note_Value, 1, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Transaction_ID, , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                End If
        
            Next

        End With

        'ﬁÌœ «·›Ê« Ì—
        With grid4

            For i = 1 To grid4.Rows - 1

                If .Cell(flexcpChecked, i, .ColIndex("select")) = flexChecked Then
                                            
                    If SystemOptions.UserInterface = ArabicInterface Then
                        StrTempDes = "”‰œ «” ·«„ —ﬁ„ " & TxtNoteSerial1V & " »‰«¡ ⁄·Ï ›« Ê—… „‘ —Ì«  —ﬁ„ " & Me.TxtNoteSerial1.text
                    Else
                        StrTempDes = "Ò Recieve Voucher No. " & TxtNoteSerial1V & " Based On Purchase Invoice NO:" & Me.TxtNoteSerial1.text
                    End If
                                                        
                    LngDevNO = LngDevNO + 1
                    Account_Code = grid4.TextMatrix(i, grid4.ColIndex("Account_code"))
                    Note_Value = grid4.TextMatrix(i, grid4.ColIndex("Note_value"))

                    If ModAccounts.AddNewDev(LngDevID, LngDevNO, Account_Code, Note_Value, 1, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Transaction_ID, , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                End If
         
            Next
   
        End With

        '«·„’—Ê›«  «·„»«‘—…
        With Fg_Journal

            For i = 1 To .Rows - 1

                If .TextMatrix(i, .ColIndex("AccountCode")) <> "" And val(.TextMatrix(i, .ColIndex("value"))) <> 0 Then
            
                    If SystemOptions.UserInterface = ArabicInterface Then
                        StrTempDes = "”‰œ «” ·«„ —ﬁ„ " & TxtNoteSerial1V & " »‰«¡ ⁄·Ï ›« Ê—… „‘ —Ì«  —ﬁ„ " & Me.TxtNoteSerial1.text
                    Else
                        StrTempDes = "Ò Recieve Voucher No. " & TxtNoteSerial1V & " Based On Purchase Invoice NO:" & Me.TxtNoteSerial1.text
                    End If
            
                    LngDevNO = LngDevNO + 1
                    Account_Code = .TextMatrix(i, .ColIndex("AccountCode"))
                    Note_Value = val(.TextMatrix(i, .ColIndex("value")))

                    If ModAccounts.AddNewDev(LngDevID, LngDevNO, Account_Code, Note_Value, 1, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Transaction_ID, , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                        GoTo ErrTrap
                    End If
                End If
        
            Next

        End With

    End If

ErrTrap:
End Function

Function CreateRecieveVouchers()

    If BillBasedOn(1).value = True Then Exit Function
    'On Error GoTo errortrap
    Dim MYWAER As String
    Dim strSQL As String
    Dim RsNotes As ADODB.Recordset
    Dim MYinvnum As String
    Dim note_id As Long

    Dim RSTransDetails As ADODB.Recordset
    Dim RsTemp As New ADODB.Recordset
    Dim RowNum As Integer
    Dim StrSqlDel As String
    Dim SearchResault As Integer
    'Dim Note_ID As Long
    Dim RsDetalis  As ADODB.Recordset
    Dim BeginTrans As Boolean
    Dim LnItemID As Long
    Dim i As Long
    Dim StrCurrentItemName As String
    Dim DblNotesTotal As Double
    Dim rs As ADODB.Recordset
    Dim IntLineNO As Integer
    Dim StrAccountCode As String
    '  Dim RowNum As Integer
    Dim Frm As Form
    Dim Msg As String
    Dim mytext As Integer
    '>>>>>>>>>>>>>>>>>>>>>>>>>
    CurrentVoucherNo = ""
    CurrentVoucherSerialNo = ""
    CurrentVoucherNo = GetVoucherGLNO(val(Text1.text), CurrentVoucherSerialNo)

    DeleteTransactiomsVoucher val(Text1.text)

    ' rs.Close

    '        rs.Open "select * from Transactions where nots = " & TxtTransSerial.text & " and Transaction_type = 20"
    '          If rs.RecordCount > 0 Then
    '        If rs!nots <> "" Then
    '        If SystemOptions.UserInterface = ArabicInterface Then
    '             Msg = "·ﬁœ  „  ÕÊÌ· Â–… «·›« Ê—… «·Ï «–‰ «÷«›…    .."
    '             Msg = " »«·«–‰ —ﬁ„ " + Text1.text & Chr(13)
    '            Msg = Msg & Chr(13) & "Ê·«Ì„ﬂ‰  ÕÊÌ·… „—… «Œ—Ï  ..!!"
    '        Else
    '          Msg = "This bill already converted" & Chr(13)
    '          Msg = Msg + " Voucher No " + Text1.text & Chr(13)
    'End If
    ''          MsgBox Msg, vbOKOnly, App.Title
    '       Exit Function
    '       End If
    '       End If

    '        rs.Close

    Set rs = New ADODB.Recordset

    strSQL = "select * from Transactions where Transaction_Serial = " & TxtTransSerial.text & " and Transaction_type = 22"
    rs.Open strSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    '      If SystemOptions.UserInterface = ArabicInterface Then
    '        Msg = "”Ê› Ì „ «‰‘«¡ «–‰ «÷«›… „‰ Â–… «·›« Ê—…   .."
    '        Msg = Msg & Chr(13) & "›Â· «‰  „ «ﬂœ „‰ «·√” „—«— ..!!"
    '      Else
    '       Msg = "Create Recieve Voucher to this bill ?"
    '        End If
    ' On Error GoTo ErrTrap

    ' If MsgBox(Msg, vbYesNo, App.Title) = vbYes Then
    '    Screen.MousePointer = vbArrowHourglass
    '    Set Frm = New FrmInpout

    '       If Rs.EOF Then
    mytext = CStr(new_id("Transactions", "Transaction_Serial", "", True, "Transaction_Type=20"))
    '      rs!nots = mytext
    '      rs.update

    Dim Transaction_ID As Long
    Transaction_ID = CStr(new_id("Transactions", "Transaction_ID", "", True))
 
    rs!nots = Transaction_ID
    rs.update
         
    'set rs!Transaction_Serial=  where Transaction_Type=20

    '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    my_branch = val(Me.dcBranch.BoundText)
    Dim general_noteid As Long
    Dim RsNotesGeneral As ADODB.Recordset
    Dim TxtNoteSerialV As String
    ' Dim TxtNoteSerial1V As String
            
    my_branch = val(Me.dcBranch.BoundText)

    If TxtNoteSerialV = "" Then
        If Notes_coding(val(my_branch), XPDtbBill.value) = "error" Then
            MsgBox " ·« Ì„ﬂ‰ «÷«›… ﬁÌÊœ ·Â–… «·⁄„·Ì… ·«‰ﬂ  ⁄œÌ  «·Õœ «·«ﬁ’Ì ··ﬁÌÊœ ﬂ„« Õœœ  ›Ì  —ﬁÌ„ «·”‰œ«  ": Exit Function
        Else
                       
            If Notes_coding(val(my_branch), XPDtbBill.value) = "" Then
                MsgBox " ·«Ì„ﬂ‰ «‰‘«¡ «·ﬁÌœ ·Â–« «·„” ‰œ ·«‰ﬂ Õœœ   —ﬁÌ„ ﬁÌÊœ ÌœÊÌ  ": Exit Function
            Else
                TxtNoteSerialV = Notes_coding(val(my_branch), XPDtbBill.value)
            End If
        End If
    End If
        
    If TxtNoteSerial1V = "" Then
        If Voucher_coding(val(my_branch), XPDtbBill.value, 9, 160, , 20) = "error" Then
            MsgBox " ·« Ì„ﬂ‰ «÷«›… ”‰œ «÷«›… ÃœÌœ ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": Exit Function
        Else
                       
            If Voucher_coding(val(my_branch), XPDtbBill.value, 9, 160, , 20) = "" Then
                MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ ”‰œ «÷«›…  ÌœÊÌ« ﬂ„« Õœœ   ": Exit Function
            Else
                TxtNoteSerial1V = Voucher_coding(val(my_branch), XPDtbBill.value, 9, 160, , 20)
            End If
        End If
    End If
 
    If Trim(CurrentVoucherNo) <> "" And DateChanged <> True Then
        TxtNoteSerialV = CurrentVoucherNo '—ﬁ„ «·ﬁÌœ
        TxtNoteSerial1V = Trim(CurrentVoucherSerialNo)
    End If
           
    Cn.Execute "INSERT INTO  Transactions (Transaction_ID ,Transaction_Serial,Transaction_Date,Transaction_Type ,CusID,StoreID,UserID,Emp_ID,nots,NoteSerial,NoteSerial1,NoteId,BranchId,nots2)SELECT " & Transaction_ID & "," & mytext & ",Transaction_Date,Transaction_Type = 20,CusID,StoreID,UserID,Emp_ID,nots='" & TxtTransSerial.text & "',NoteSerial=' " & TxtNoteSerialV & "',NoteSerial1='" & TxtNoteSerial1V & "',NoteId=" & general_noteid & ",BranchId ," & TxtNoteSerial1 & " From Transactions Where Transaction_ID =" & XPTxtBillID.text + " And Transaction_Type = 22"
    'Create big notes
    Set RsNotesGeneral = New ADODB.Recordset
'    RsNotesGeneral.Open "[Notes]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    
    
       strSQL = "SELECT     dbo.Notes.* from dbo.Notes Where (NoteID = -1)"
   RsNotesGeneral.Open strSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText


    If Me.TxtModFlg.text = "N" Then
    
    Else
 
        '   general_noteid = Val(TxtNoteID.text)
    End If

    RsNotesGeneral.AddNew
    RsNotesGeneral("NoteID").value = CStr(new_id("Notes", "NoteID", "", True))
    general_noteid = RsNotesGeneral("NoteID").value
    '   TxtNoteID.text = general_noteid
    RsNotesGeneral("Transaction_ID").value = Transaction_ID
    RsNotesGeneral("NoteDate").value = XPDtbBill.value
    RsNotesGeneral("NoteType").value = 160
    RsNotesGeneral("Note_Value").value = Null
    RsNotesGeneral("NoteSerial").value = IIf(Trim(TxtNoteSerialV) = "", Null, Trim(TxtNoteSerialV))
    RsNotesGeneral("NoteSerial1").value = IIf(Trim(TxtNoteSerial1V) = "", Null, Trim(TxtNoteSerial1V))
    RsNotesGeneral("numbering_type").value = sand_numbering_type(0) '”‰œ «·ﬁÌœ
    RsNotesGeneral("numbering_type1").value = sand_numbering_type(9) '«–‰ «÷«›…
    RsNotesGeneral("sanad_year").value = year(XPDtbBill.value)
    RsNotesGeneral("sanad_month").value = Month(XPDtbBill.value)
    RsNotesGeneral("NoteID").value = CStr(new_id("Notes", "NoteID", "", True))
    RsNotesGeneral("branch_no").value = val(Me.dcBranch.BoundText)
    general_noteid = RsNotesGeneral("NoteID").value
    'RsNotes("note_value_by_characters").value = Trim$(Me.lbl(18).Caption)
    RsNotesGeneral.update
    '%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
        
    '
    Dim sql As String
    sql = "INSERT INTO  dbo.Transaction_Details(showPrice,guaranteeTime,Transaction_ID,Item_ID,ItemCase,ItemSerial,Quantity,Price,ColorID,ItemSize,UnitId,ShowQty,QtyBySmalltUnit,BranchId,FoxyNo,OrderArrivalDate,order_no,ClassID,ProductionDate,ExpiryDate,LotNO) " & "SELECT   ( ( (showPrice-(discountvalue+TotalDiscountPerLine)*QtyBySmalltUnit)*" & val(txt_Currency_rate.text) & ")+(ToTAlELSHahn+LineShahn)*QtyBySmalltUnit) ,guaranteeTime," & Transaction_ID & ",Item_ID,ItemCase,ItemSerial , Quantity, (( (Price-(discountvalue+TotalDiscountPerLine))*" & val(txt_Currency_rate.text) & ")+(ToTAlELSHahn+LineShahn) ), ColorID,ItemSize, UnitId, ShowQty, QtyBySmalltUnit,BranchId,FoxyNo,OrderArrivalDate,order_no,ClassID,ProductionDate,ExpiryDate,LotNO  From dbo.Transaction_Details Where Transaction_ID = " & XPTxtBillID.text
       
    'Cn.Execute "INSERT INTO  dbo.Transaction_Details(showPrice,guaranteeTime,Transaction_ID,Item_ID,ItemCase,ItemSerial,Quantity,Price,ColorID,UnitId,ShowQty,QtyBySmalltUnit)SELECT round(showPrice + ToTAlELSHahn/ShowQty,2),guaranteeTime," & Transaction_ID & ",Item_ID,ItemCase,ItemSerial , Quantity, Price*rate+ToTAlELSHahn, ColorID, UnitId, ShowQty, QtyBySmalltUnit From dbo.Transaction_Details Where Transaction_ID = " & XPTxtBillID.text
    Cn.Execute sql
    '"INSERT INTO  dbo.Transaction_Details(showPrice,guaranteeTime,Transaction_ID,Item_ID,ItemCase,ItemSerial,Quantity,Price,ColorID,ItemSize,UnitId,ShowQty,QtyBySmalltUnit,BranchId,FoxyNo,OrderArrivalDate,order_no,ClassID) " & _
     "SELECT   ( ( (showPrice-discountvalue)*" & Val(txt_Currency_rate.text) & ")+(ToTAlELSHahn+LineShahn)*QtyBySmalltUnit) ,guaranteeTime," & Transaction_ID & ",Item_ID,ItemCase,ItemSerial , Quantity, ((Price*" & Val(txt_Currency_rate.text) & ")+(ToTAlELSHahn+LineShahn) ), ColorID,ItemSize, UnitId, ShowQty, QtyBySmalltUnit,BranchId,FoxyNo,OrderArrivalDate,order_no,ClassID  From dbo.Transaction_Details Where Transaction_ID = " & XPTxtBillID.text
        
    'Cn.Execute "update Transactions Set Transaction_Serial = Transaction_Serial Where Transaction_Type = 20"

    CREATE_VOUCHER_GE Transaction_ID, TxtNoteSerialV, TxtNoteSerial1V, general_noteid, val(Me.dcBranch.BoundText)

    'End If

    '>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
  
    'If Text1.text <> "" Then
    '    Msg = " „  ÕÊÌ· Â–… «·›« Ê—… „‰ ﬁ»· Ê·« Ì„ﬂ‰  ÕÊÌ·Â«  "
    '            MsgBox Msg, vbOKOnly + vbMsgBoxRight + vbExclamation + vbMsgBoxRtlReading, App.Title
    '            Exit Sub
    '        End If
    'On Error GoTo ErrTrap
    'Screen.MousePointer = vbArrowHourglass
    '    Set Frm = New FrmInpout
    'With Frm
    '    .Convert
    ''    .XPTxtBillID.Text = XPTxtBillID.Text
    '    .XPDtbBill.Value = XPDtbBill.Value
    '    .DBCboClientName.BoundText = DBCboClientName.BoundText
    '    .DCboStoreName.BoundText = DCboStoreName.BoundText
    '    .CboPayMentType.ListIndex = CboPayMentType.ListIndex
    '    .Text1.text = TxtTransSerial.text
    '    .Text2.text = XPTxtBillID.text
    '
    '
    '    For RowNum = 1 To FG.Rows - 1
    '        If .FG.TextMatrix(.FG.Rows - 1, .FG.ColIndex("Code")) <> "" Then
    '           .FG.Rows = .FG.Rows + 1
    '        End If
    '        .FG.TextMatrix(.FG.Rows - 1, .FG.ColIndex("Name")) = IIf(FG.TextMatrix(RowNum, FG.ColIndex("Name")) = "", "", FG.TextMatrix(RowNum, FG.ColIndex("Name")))
    '        .FG.TextMatrix(.FG.Rows - 1, .FG.ColIndex("Code")) = IIf(FG.TextMatrix(RowNum, FG.ColIndex("Code")) = "", "", FG.TextMatrix(RowNum, FG.ColIndex("Code")))
    '        .FG.TextMatrix(.FG.Rows - 1, .FG.ColIndex("ItemCase")) = IIf(FG.TextMatrix(RowNum, FG.ColIndex("ItemCase")) = "", "", FG.TextMatrix(RowNum, FG.ColIndex("ItemCase")))
    '        .FG.TextMatrix(.FG.Rows - 1, .FG.ColIndex("HaveSerial")) = IIf(FG.TextMatrix(RowNum, FG.ColIndex("HaveSerial")) = "", "", FG.TextMatrix(RowNum, FG.ColIndex("HaveSerial")))
    '        .FG.TextMatrix(.FG.Rows - 1, .FG.ColIndex("Count")) = IIf(FG.TextMatrix(RowNum, FG.ColIndex("Count")) = "", "", FG.TextMatrix(RowNum, FG.ColIndex("Count")))
    '        .FG.TextMatrix(.FG.Rows - 1, .FG.ColIndex("Price")) = IIf(FG.TextMatrix(RowNum, FG.ColIndex("Price")) = "", "", FG.TextMatrix(RowNum, FG.ColIndex("Price")))
    '        .FG.TextMatrix(.FG.Rows - 1, .FG.ColIndex("DiscountType")) = IIf(FG.TextMatrix(RowNum, FG.ColIndex("DiscountType")) = "", "", FG.TextMatrix(RowNum, FG.ColIndex("DiscountType")))
    ''        Dim StrSQL As String
    '        Dim RsUnit As New ADODB.Recordset
    'StrSQL = "SELECT TOP 100 PERCENT dbo.TblItemsUnits.UnitID, dbo.TblUnites.UnitName, dbo.Transactions.Transaction_Serial,dbo.Transactions.Transaction_Type FROM dbo.Transaction_Details INNER JOIN dbo.Transactions ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID INNER JOIN dbo.TblUnites INNER JOIN dbo.TblItemsUnits ON dbo.TblUnites.UnitID = dbo.TblItemsUnits.UnitID ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID AND dbo.Transaction_Details.Item_ID = dbo.TblItemsUnits.ItemID WHERE (dbo.Transactions.Transaction_Serial = '" & TxtTransSerial & "') AND (dbo.Transactions.Transaction_Type = 22) AND (dbo.TblItemsUnits.ItemID = " & FG.TextMatrix(RowNum, FG.ColIndex("Code")) & ") ORDER BY dbo.TblItemsUnits.SecOrder"
    'Set RsUnit = New ADODB.Recordset
    'RsUnit.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    '
    '
    '
    '        .FG.Cell(flexcpData, RowNum, .FG.ColIndex("UnitID")) = IIf(IsNull(RsUnit("UnitID")), "", (RsUnit("UnitID").Value))
    '        .FG.TextMatrix(RowNum, .FG.ColIndex("UnitID")) = IIf(IsNull(RsUnit("UnitName")), "", (RsUnit("UnitName").Value))
    '         Rs!nots = TxtTransSerial.text
    '         Rs.update
    '
    '
    ''        FG.Cell(flexcpData, I, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitID")), "", (RsDetails("UnitID").Value))
    ''        FG.TextMatrix(I, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitName")), "", (RsDetails("UnitName").Value))
    ''           StrSQL = "SELECT dbo.Transactions.Transaction_Type, dbo.Transaction_Details.UnitId, dbo.TblUnites.UnitName, dbo.Transactions.Transaction_Serial FROM dbo.Transaction_Details INNER JOIN dbo.Transactions ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID INNER JOIN dbo.TblUnites ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID WHERE (dbo.Transactions.Transaction_Type = 19) AND (dbo.Transactions.Transaction_Serial = '" & TxtTransSerial & "')"
    ''        .FG.Cell(flexcpData, .FG.Rows - 1, FG.ColIndex("UnitID")) = 1 'FG.Cell(flexcpData, RowNum, FG.ColIndex("UnitID")) ' IIf(IsNull(RsUnit("UnitID")), "", (RsUnit("UnitID").Value))
    ''        .FG.TextMatrix(.FG.Rows - 1, FG.ColIndex("UnitID")) = "Ã—«„" 'FG.TextMatrix(RowNum, FG.ColIndex("UnitID")) ' IIf(IsNull(RsUnit("UnitName")), "", (RsUnit("UnitName").Value))
    '
    '    Next RowNum
    '    .Cala
    'End With
    'Screen.MousePointer = vbDefault
    'Cmd_Click (2)
    'Frm.Hide
    'Exit Sub
    'errortrap:
    'Screen.MousePointer = vbDefault
    'MsgBox " „  ÕÊÌ· Â–… «·›« Ê—… „‰ ﬁ»·", vbCritical
ErrTrap:

End Function

Private Sub Command1_Click()
    CreateRecieveVouchers
End Sub

Private Sub Command2_Click()

    '    ⁄»∆… «·«–Ê‰ «·„’—Ê›« 
    If CBoBasedON.ListIndex = 0 Or CBoBasedON.ListIndex = 1 Or Txt_order_no.text = "" Then

        With Me.Grid
            .Rows = .FixedRows
   
        End With

        Exit Sub

    End If

    With Me.Grid
        .Rows = .FixedRows
        .ExtendLastCol = True
        .RowHeightMin = 300
        .Editable = flexEDKbdMouse
        .ExplorerBar = flexExSortShowAndMove

        '    .AutoSize 0, .Cols - 1, False
    End With

    Dim i As Integer
    Dim RsExp As ADODB.Recordset
    Dim My_SQL As String

    Set RsExp = New ADODB.Recordset

    'My_SQL = "SELECT dbo.Notes.NoteID,dbo.Notes.buy,dbo.Notes.NoteSerial,dbo.notes.ItemID , dbo.Notes.Note_Value, dbo.ExpensesType.Name ,  dbo.ExpensesType.Account_Code FROM dbo.Notes INNER JOIN dbo.ExpensesType ON dbo.Notes.ExpensesID = dbo.ExpensesType.ID Where (dbo.Notes.NoteType = 3   and order_no='" & Me.Txt_order_no.text & "' and(Transaction_ID1 is null or Transaction_ID1=" & Val(Me.XPTxtBillID.text) & ")  )  "
    My_SQL = "SELECT dbo.Notes.NoteID,dbo.Notes.buy,dbo.Notes.NoteSerial,dbo.notes.ItemID , dbo.Notes.Note_Value, dbo.ExpensesType.Name ,  dbo.ExpensesType.Account_Code FROM dbo.Notes INNER JOIN dbo.ExpensesType ON dbo.Notes.ExpensesID = dbo.ExpensesType.ID Where  dbo.Notes.NoteType = 3   and order_no='" & Me.Txt_order_no.text & "'"
    'My_SQL = ""

    RsExp.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText
    Dim strSQL  As String

    With Me.Grid
        .Rows = 1
        .Clear flexClearScrollable

        If RsExp.RecordCount > 0 Then
            .Rows = RsExp.RecordCount + 1
            RsExp.MoveFirst

            For i = 1 To .Rows - 1
                   
                .TextMatrix(i, .ColIndex("ItemID")) = IIf(IsNull(RsExp.Fields("ItemID").value), "", RsExp.Fields("ItemID").value)
    
                strSQL = "select * from TblItems where ItemID=" & val(.TextMatrix(i, .ColIndex("ItemID")))
                Dim rs As New ADODB.Recordset
                Set rs = Nothing
                rs.Open strSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If Not (rs.BOF Or rs.EOF) Then
    
                    .TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(rs("ItemName").value), "", rs("ItemName").value)
                    .TextMatrix(i, .ColIndex("ItemCode")) = IIf(IsNull(rs("ItemCode").value), "", rs("ItemCode").value)
                    
                Else
            
                    .TextMatrix(i, .ColIndex("ItemName")) = ""
                    .TextMatrix(i, .ColIndex("ItemCode")) = ""
 
                End If
               
                .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(RsExp.Fields("Name").value), "", RsExp.Fields("Name").value)
               
                .TextMatrix(i, .ColIndex("NoteSerial")) = IIf(IsNull(RsExp.Fields("NoteSerial").value), "", RsExp.Fields("NoteSerial").value)
            
                .TextMatrix(i, .ColIndex("NoteID")) = IIf(IsNull(RsExp.Fields("NoteID").value), "", RsExp.Fields("NoteID").value)
           
                .TextMatrix(i, .ColIndex("Note_Value")) = IIf(IsNull(RsExp.Fields("Note_Value").value), "", RsExp.Fields("Note_Value").value)
                .TextMatrix(i, .ColIndex("Account_Code")) = IIf(IsNull(RsExp.Fields("Account_Code").value), "", RsExp.Fields("Account_Code").value)
            
                If IsNull(RsExp.Fields("buy").value) Then
                    .TextMatrix(i, .ColIndex("Select")) = 0
                Else

                    If RsExp.Fields("buy").value = False Then
                        .TextMatrix(i, .ColIndex("Select")) = 0
                    ElseIf RsExp.Fields("buy").value = True Then
                        .TextMatrix(i, .ColIndex("Select")) = 1
                    Else
                        .TextMatrix(i, .ColIndex("Select")) = 0
                    End If
           
                End If
           
                .TextMatrix(i, .ColIndex("Select")) = 1
               
                RsExp.MoveNext
            Next

            RsExp.Close
        End If

        .RowHeight(-1) = 300
    End With

    Grid.Visible = True

    '    ⁄»∆… «·«–Ê‰ «·„œ›Ê⁄« 

    Expenses_update_total

End Sub

Private Sub Command3_Click()

    ' ⁄»∆… «Ê«„— «·‘—«¡ Ê «·»Ì⁄

    With Me.GRID2
        .Rows = .FixedRows
        .ExtendLastCol = True
        .RowHeightMin = 300
        .Editable = flexEDKbdMouse
        .ExplorerBar = flexExSortShowAndMove

        .AutoSize 0, .Cols - 1, False
    End With

    Dim i As Integer
    Dim RsExp As ADODB.Recordset
    Dim My_SQL As String

    Set RsExp = New ADODB.Recordset
    My_SQL = "SELECT dbo.Transactions.Transaction_ID,dbo.Transactions.order_no , dbo.Transactions.Transaction_Date,dbo.Transactions.CusID, dbo.TblCustemers.CusName FROM dbo.Transactions  INNER JOIN dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID Where  (Transaction_Type=6  or Transaction_Type=29)and NOT(ORDER_NO IS NULL) AND CLOSED= 0 and   dbo.TblCustemers.CusID=" & val(DBCboClientName.BoundText)

    RsExp.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText

    With Me.GRID2
        .Rows = 2
        .Clear flexClearScrollable

        If RsExp.RecordCount > 0 Then
            .Rows = RsExp.RecordCount + 1
            RsExp.MoveFirst

            For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("order_no")) = IIf(IsNull(RsExp.Fields("order_no").value), "", RsExp.Fields("order_no").value)
               
                .TextMatrix(i, .ColIndex("Transaction_Date")) = IIf(IsNull(RsExp.Fields("Transaction_Date").value), "", RsExp.Fields("Transaction_Date").value)
           
                .TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(RsExp.Fields("CusName").value), "", RsExp.Fields("CusName").value)
                .TextMatrix(i, .ColIndex("Transaction_ID")) = IIf(IsNull(RsExp.Fields("Transaction_ID").value), "", RsExp.Fields("Transaction_ID").value)

                RsExp.MoveNext
            Next

            RsExp.Close
        End If

        .RowHeight(-1) = 300
    End With

    GRID2.Visible = True

End Sub

Private Sub Command4_Click()
    'If Not Fg.TextMatrix(Fg.Row, Fg.ColIndex("Code")) = "" Then
    '    ⁄»∆… «·«–Ê‰ «·„’—Ê›« 

    'Frame2.Caption = FG.TextMatrix(FG.Row, FG.ColIndex("name"))

    If CBoBasedON.ListIndex = 0 Or CBoBasedON.ListIndex = 1 Or Txt_order_no.text = "" Then

        With Me.grid4
            .Rows = .FixedRows
   
        End With

        Exit Sub

    End If

    With Me.grid4
        .Rows = .FixedRows
        .ExtendLastCol = True
        .RowHeightMin = 300
        .Editable = flexEDKbdMouse
        .ExplorerBar = flexExSortShowAndMove
        '
        '    .AutoSize 0, .Cols - 1, False
    End With

    Dim i As Integer
    Dim RsExp As ADODB.Recordset
    Dim My_SQL As String

    Set RsExp = New ADODB.Recordset

    'My_SQL = "SELECT dbo.Notes.Item_id,dbo.Notes.NoteID,dbo.Notes.buy,dbo.Notes.NoteSerial , dbo.Notes.Note_Value, dbo.ExpensesType.Name ,  dbo.ExpensesType.Account_Code FROM dbo.Notes INNER JOIN dbo.ExpensesType ON dbo.Notes.ExpensesID = dbo.ExpensesType.ID Where (dbo.Notes.NoteType = 3 and order_no='" & Me.TXT_order_no.text & "' " & "AND (ITEM_ID=" & Val(FG.TextMatrix(FG.Row, FG.ColIndex("Code"))) & " or  ITEM_ID is null)  and(Transaction_ID1 is null or Transaction_ID1=" & Val(Me.XPTxtBillID.text) & "))  "
    My_SQL = "SELECT     dbo.Notes.NoteType, dbo.Notes.NoteSerial1, dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code, dbo.DOUBLE_ENTREY_VOUCHERS.[Value], "
    My_SQL = My_SQL + " dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit, dbo.ACCOUNTS.Account_Name, dbo.ACCOUNTS.Account_NameEng, dbo.ACCOUNTS.Account_Serial,"
    My_SQL = My_SQL + " dbo.Notes.order_no, dbo.DOUBLE_ENTREY_VOUCHERS.ItemID ,  dbo.Notes.NoteID, dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_ID, dbo.DOUBLE_ENTREY_VOUCHERS.buy,dbo.DOUBLE_ENTREY_VOUCHERS.Transaction_ID1 "
    My_SQL = My_SQL + " FROM         dbo.Notes INNER JOIN"
    My_SQL = My_SQL + " dbo.DOUBLE_ENTREY_VOUCHERS ON dbo.Notes.NoteID = dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID INNER JOIN"
    My_SQL = My_SQL + "  dbo.ACCOUNTS ON dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code = dbo.ACCOUNTS.Account_Code"
    'My_SQL = My_SQL + " WHERE      (dbo.DOUBLE_ENTREY_VOUCHERS.Transaction_ID1 is null or dbo.DOUBLE_ENTREY_VOUCHERS.Transaction_ID1=" & Val(Me.XPTxtBillID.text) & ") and  (dbo.Notes.NoteType = 80) AND (dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit = 0) AND (dbo.Notes.ORDER_NO = '" & Me.Txt_order_no.text & "')"
    My_SQL = My_SQL + " WHERE       (dbo.Notes.NoteType = 80) AND (dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit = 0) AND (dbo.Notes.ORDER_NO = '" & Me.Txt_order_no.text & "')"

    RsExp.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText

    Dim strSQL As String
    Dim rs As New ADODB.Recordset

    With Me.grid4
        .Rows = 1
        .Clear flexClearScrollable

        If RsExp.RecordCount > 0 Then
            .Rows = RsExp.RecordCount + 1
            RsExp.MoveFirst

            For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("Double_Entry_Vouchers_ID")) = IIf(IsNull(RsExp.Fields("Double_Entry_Vouchers_ID").value), 0, RsExp.Fields("Double_Entry_Vouchers_ID").value)
           
                .TextMatrix(i, .ColIndex("ItemID")) = IIf(IsNull(RsExp.Fields("ItemID").value), "", RsExp.Fields("ItemID").value)
    
                strSQL = "select * from TblItems where ItemID=" & val(.TextMatrix(i, .ColIndex("ItemID")))
                Set rs = Nothing
                rs.Open strSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If Not (rs.BOF Or rs.EOF) Then
    
                    .TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(rs("ItemName").value), "", rs("ItemName").value)
                    .TextMatrix(i, .ColIndex("ItemCode")) = IIf(IsNull(rs("ItemCode").value), "", rs("ItemCode").value)
                    
                Else
            
                    .TextMatrix(i, .ColIndex("ItemName")) = ""
                    .TextMatrix(i, .ColIndex("ItemCode")) = ""
 
                End If
               
                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(RsExp.Fields("Account_Name").value), "", RsExp.Fields("Account_Name").value)
 
                Else
                    .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(RsExp.Fields("Account_NameEng").value), "", RsExp.Fields("Account_NameEng").value)
                End If
 
                .TextMatrix(i, .ColIndex("NoteSerial1")) = IIf(IsNull(RsExp.Fields("NoteSerial1").value), "", RsExp.Fields("NoteSerial1").value)
 
                .TextMatrix(i, .ColIndex("NoteID")) = IIf(IsNull(RsExp.Fields("NoteID").value), "", RsExp.Fields("NoteID").value)
 
                .TextMatrix(i, .ColIndex("Note_Value")) = IIf(IsNull(RsExp.Fields("Value").value), "", RsExp.Fields("Value").value)
 
                .TextMatrix(i, .ColIndex("Account_Code")) = IIf(IsNull(RsExp.Fields("Account_Code").value), "", RsExp.Fields("Account_Code").value)
 
                If IsNull(RsExp.Fields("buy").value) Then
                    .TextMatrix(i, .ColIndex("Select")) = 0
                Else

                    If RsExp.Fields("buy").value = False Then
                        .TextMatrix(i, .ColIndex("Select")) = 0
                    ElseIf RsExp.Fields("buy").value = True Then
                        .TextMatrix(i, .ColIndex("Select")) = 1
                    Else
                        .TextMatrix(i, .ColIndex("Select")) = 0
                    End If
           
                End If
 
                ' .TextMatrix(i, .ColIndex("Select")) = IIf(IsNull(RsExp.Fields("buy").value), _
                  0, RsExp.Fields("buy").value)
                .TextMatrix(i, .ColIndex("Select")) = 1
                RsExp.MoveNext
            Next

            RsExp.Close
        End If

        .RowHeight(-1) = 300
    End With

    grid4.Visible = True

    ' End If
  
    update_finincial_invoice_total
       
End Sub

Private Sub save_expenses()
    Dim Item_ID As Integer
    Dim i As Integer
    Dim sql As String
    ' ﬁÊ„ » ÕÌÀ ﬂ· ”ÿ— ›Ì «·ﬁÌœ »Ê÷⁄ —ﬁ„ «·⁄„·Ì… Ê—ﬁ„ «·’‰› Ê buy - Notes
    'Item_ID = Val(FG.TextMatrix(FG.Row, FG.ColIndex("Code")))
    ' ›—Ì€  ﬂ·›… «·‘Õ‰ ⁄·Ï „” ÊÏ «·”ÿ—

    With Grid

        For i = 1 To Grid.Rows - 1
      
            Cn.BeginTrans
 
            If .Cell(flexcpChecked, i, .ColIndex("select")) = flexChecked Then
                check_item_Exist_in_Grid val(.TextMatrix(i, .ColIndex("ItemID"))), val(.TextMatrix(i, .ColIndex("Note_value"))), True
        
                sql = "update notes set Transaction_ID1=" & val(Me.XPTxtBillID.text) & " , buy='1',itemid=" & IIf(val(.TextMatrix(i, .ColIndex("itemid"))) = 0, "Null", val(.TextMatrix(i, .ColIndex("itemid")))) & " where NoteID=" & val(.TextMatrix(i, .ColIndex("NoteID")))
        
            Else
                sql = "update notes set Transaction_ID1=null ,  buy=Null,itemid=Null  where NoteID=" & val(.TextMatrix(i, .ColIndex("NoteID")))

            End If

            Cn.Execute sql

            Cn.CommitTrans

        Next

    End With

    Expenses_update_total

End Sub

Function Expenses_update_total()
    Dim i As Integer
    On Error Resume Next
    Txt_EXport.text = 0

    If Grid.Rows = 1 Then Exit Function

    With Grid

        For i = 1 To Grid.Rows - 1
        
            If .Cell(flexcpChecked, i, .ColIndex("select")) = flexChecked And Grid.TextMatrix(i, Grid.ColIndex("ItemID")) = "" Then
            
                Txt_EXport.text = val(Txt_EXport.text) + val(Grid.TextMatrix(i, Grid.ColIndex("note_value")))
            End If
            
            If val(Grid.TextMatrix(i, Grid.ColIndex("select"))) = 0 Then
                Grid.TextMatrix(i, Grid.ColIndex("ItemID")) = ""
                Grid.TextMatrix(i, Grid.ColIndex("ItemCode")) = ""
                Grid.TextMatrix(i, Grid.ColIndex("ItemName")) = ""
            
            End If
            
        Next
 
    End With
       
End Function

Private Sub Save_Financial_invoice()
    'FG.TextMatrix(FG.Row, FG.ColIndex("LineShahn")) = Val(Me.txt_item_expenses.text)
    ' ﬁÊ„ » ÕÌÀ ﬂ· ”ÿ— ›Ì «·ﬁÌœ »Ê÷⁄ —ﬁ„ «·⁄„·Ì… Ê—ﬁ„ «·’‰› Ê buy - Double entry Voucher
    Dim Item_ID As Integer
    Dim i As Integer
    Dim sql As String

    'Item_ID = Val(FG.TextMatrix(FG.Row, FG.ColIndex("Code")))
    ' ›—Ì€  ﬂ·›… «·‘Õ‰ ⁄·Ï „” ÊÏ «·”ÿ—
    With FG

        For i = 1 To FG.Rows - 1
        
            .TextMatrix(i, .ColIndex("LineShahn")) = 0
      
        Next i

    End With

    With grid4
 
        For i = 1 To grid4.Rows - 1
      
            Cn.BeginTrans
 
            If .Cell(flexcpChecked, i, .ColIndex("select")) = flexChecked Then
                check_item_Exist_in_Grid val(.TextMatrix(i, .ColIndex("ItemID"))), val(.TextMatrix(i, .ColIndex("Note_value")))
        
                sql = "update DOUBLE_ENTREY_VOUCHERS set Transaction_ID1=" & val(Me.XPTxtBillID.text) & " , buy='1',itemid=" & IIf(val(grid4.TextMatrix(i, grid4.ColIndex("itemid"))) = 0, "Null", val(grid4.TextMatrix(i, grid4.ColIndex("itemid")))) & " where Double_Entry_Vouchers_ID=" & val(grid4.TextMatrix(i, grid4.ColIndex("Double_Entry_Vouchers_ID")))
        
            Else
                sql = "update DOUBLE_ENTREY_VOUCHERS set Transaction_ID1=null , buy=Null,itemid=Null where Double_Entry_Vouchers_ID=" & val(grid4.TextMatrix(i, grid4.ColIndex("Double_Entry_Vouchers_ID")))

            End If

            Cn.Execute sql

            Cn.CommitTrans

        Next

    End With

    update_finincial_invoice_total

    '    DoEvents
    '    Command4_Click
End Sub

Function update_finincial_invoice_total()
    On Error Resume Next
    Dim i As Integer
    txt_total_bill.text = 0

    If grid4.Rows = 1 Then Exit Function

    With grid4

        For i = 1 To grid4.Rows - 1
        
            If .Cell(flexcpChecked, i, .ColIndex("select")) = flexChecked And grid4.TextMatrix(i, grid4.ColIndex("ItemID")) = "" Then
                txt_total_bill.text = val(txt_total_bill.text) + val(grid4.TextMatrix(i, grid4.ColIndex("note_value")))
  
            End If
            
            If val(grid4.TextMatrix(i, grid4.ColIndex("select"))) = 0 Then
                grid4.TextMatrix(i, grid4.ColIndex("ItemID")) = ""
                grid4.TextMatrix(i, grid4.ColIndex("ItemCode")) = ""
                grid4.TextMatrix(i, grid4.ColIndex("ItemName")) = ""
            
            End If

        Next

    End With

End Function

Private Sub Command5_Click()

    Save_Financial_invoice
       
End Sub

Private Sub Command6_Click()
    save_expenses
End Sub

Private Sub DBCboClientName_Change()
    Dim strSQL As String
    Dim RsTemp As ADODB.Recordset

    On Error GoTo ErrTrap
    Dim fullcode As String
 
    GetCustomersDetail val(DBCboClientName.BoundText), , fullcode, 2
    TxtSearchCode.text = fullcode

    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
        If DBCboClientName.BoundText <> "" Then
            If DBCboClientName.BoundText = 1 Or DBCboClientName.BoundText = 2 Then
                '   CboPayMentType.locked = True
                '   CboPayMentType.ListIndex = 0
            Else
                '   CboPayMentType.locked = False
            End If
        End If
    End If

    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
        strSQL = "Select * From TblCustemers Where CusID=" & val(DBCboClientName.BoundText)
        Set RsTemp = New ADODB.Recordset
        RsTemp.Open strSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (RsTemp.BOF Or RsTemp.EOF) Then
            If Not (IsNull(RsTemp("Trans_DiscountTypePur").value)) Then
                If RsTemp("Trans_DiscountTypePur").value = 0 Then
                    '     mina           Me.XPCboDiscountType.ListIndex = 0
                    '   mina             Me.XPTxtDiscountVal.text = 0
                ElseIf RsTemp("Trans_DiscountTypePur").value = 1 Then
                    Me.XPCboDiscountType.ListIndex = 1
                    Me.XPTxtDiscountVal.text = IIf(IsNull(RsTemp("Trans_DiscountPur").value), "", RsTemp("Trans_DiscountPur").value)
                ElseIf RsTemp("Trans_DiscountTypePur").value = 2 Then
                    Me.XPCboDiscountType.ListIndex = 2
                    Me.XPTxtDiscountVal.text = IIf(IsNull(RsTemp("Trans_DiscountPur").value), "", RsTemp("Trans_DiscountPur").value)
                End If

            Else
                Me.XPCboDiscountType.ListIndex = 0
                Me.XPTxtDiscountVal.text = 0
            End If

        Else
            Me.XPCboDiscountType.ListIndex = 0
            '     mina   Me.XPTxtDiscountVal.text = 0
        End If

        RsTemp.Close
        Set RsTemp = Nothing
                
    End If

    If BillBasedOn(1).value = True Then
                
        FillVoucherGrid (1)
                
    End If
    
    Exit Sub
ErrTrap:
End Sub

Private Sub DBCboClientName_Click(Area As Integer)
    DBCboClientName_Change
End Sub

Private Sub DBCboClientName_KeyUp(KeyCode As Integer, _
                                  Shift As Integer)

    If KeyCode = vbKeyF3 Then
        FrmCompanySearch.show vbModal
        FrmCompanySearch.lblSearchtype.Caption = 1
    End If

End Sub

Private Sub DCboItemsCode_KeyUp(KeyCode As Integer, _
                                Shift As Integer)

    If KeyCode = vbKeyF3 Then
        Load FrmItemSearch
        FrmItemSearch.RetrunType = 3
        FrmItemSearch.show vbModal
    End If

End Sub

Private Sub Dcbranch_Change()
    Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos

    If Me.TxtModFlg = "N" Or Me.TxtModFlg = "E" Then
        Dcombos.GetDocTypebyid Me.DCDocTypes, 22, val(Me.dcBranch.BoundText)
    End If

    If dcBranch.BoundText = "" Then TxtNoteSerial1.locked = True: Exit Sub

    If Voucher_coding(val(Me.dcBranch.BoundText), XPDtbBill.value, 6, 150, 22) = "" Then
        TxtNoteSerial1.locked = True
    Else
        TxtNoteSerial1.locked = False
 
    End If

End Sub

Private Sub Dcbranch_Click(Area As Integer)
    Dcbranch_Change

    If Voucher_coding(val(val(Me.dcBranch.BoundText)), XPDtbBill.value, 6, 150, , 22) = "" Then Exit Sub
    TxtNoteSerial.text = ""
    TxtNoteSerial1.text = ""
End Sub

Private Sub Dcbranch_KeyUp(KeyCode As Integer, _
                           Shift As Integer)

    If KeyCode = vbKeyF5 Then
        Dim Dcombos As ClsDataCombos
        
        Set Dcombos = New ClsDataCombos
        Dcombos.GetBranches dcBranch
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

Private Sub DCCurrency_KeyUp(KeyCode As Integer, _
                             Shift As Integer)

    If KeyCode = vbKeyF5 Then
        Dim strSQL As String
        strSQL = " select id,code from currency"
 
        fill_combo Me.Dccurrency, strSQL
    End If

End Sub

Private Sub DCDocTypes_KeyUp(KeyCode As Integer, _
                             Shift As Integer)

    If KeyCode = vbKeyF5 Then
        Dim Dcombos As ClsDataCombos
        
        Set Dcombos = New ClsDataCombos
        Dcombos.GetDocTypebyid Me.DCDocTypes, 22, val(Me.dcBranch.BoundText)
    End If

End Sub

Private Sub Ele_DblClick(Index As Integer)
    On Error GoTo ErrTrap

    Select Case Index

        Case 6

            If Me.WindowState = vbNormal Then
                Me.WindowState = vbMaximized
            Else
                Me.WindowState = vbNormal
            End If

    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub FG_AfterEdit(ByVal Row As Long, _
                         ByVal Col As Long)
    fill_bill_items_table

    If Me.TxtModFlg <> "E" Then Exit Sub

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\LogFile Saving//////////////////////////////////////////////////////////////////////////
    If Col = FG.ColIndex("Code") Or Col = FG.ColIndex("Name") Then
        RegisterItemData Me.name, Me.TxtModFlg, FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , , , , , , Me.TxtNoteSerial, Me.TxtNoteSerial1, 150
    ElseIf Col = FG.ColIndex("UnitID") Then
        RegisterItemData Me.name, Me.TxtModFlg, FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("UnitID")), , , , , , , , , Me.TxtNoteSerial, Me.TxtNoteSerial1, 150
    ElseIf Col = FG.ColIndex("Count") Then
        RegisterItemData Me.name, Me.TxtModFlg, FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , (FG.TextMatrix(Row, FG.ColIndex("Count"))), , , , , , , , Me.TxtNoteSerial, Me.TxtNoteSerial1, 150
    ElseIf Col = FG.ColIndex("Price") Then
        RegisterItemData Me.name, Me.TxtModFlg, FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , (FG.TextMatrix(Row, FG.ColIndex("Price"))), , , , , , , Me.TxtNoteSerial, Me.TxtNoteSerial1, 150
    ElseIf Col = FG.ColIndex("ColorID") Then
        RegisterItemData Me.name, Me.TxtModFlg, FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("ColorID")), , , , , Me.TxtNoteSerial, Me.TxtNoteSerial1, 150
    ElseIf Col = FG.ColIndex("ItemSize") Then
        RegisterItemData Me.name, Me.TxtModFlg, FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , , FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("ItemSize")), , , , Me.TxtNoteSerial, Me.TxtNoteSerial1, 150
    ElseIf Col = FG.ColIndex("ClassId") Then
        RegisterItemData Me.name, Me.TxtModFlg, FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , , , FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("ClassId")), , , Me.TxtNoteSerial, Me.TxtNoteSerial1, 150
    ElseIf Col = FG.ColIndex("DiscountType") Then
        RegisterItemData Me.name, Me.TxtModFlg, FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , , , , FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("DiscountType")), , Me.TxtNoteSerial, Me.TxtNoteSerial1, 150
    ElseIf Col = FG.ColIndex("DiscountVal") Then
        RegisterItemData Me.name, Me.TxtModFlg, FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Code")), FG.Cell(flexcpTextDisplay, Row, FG.ColIndex("Name")), , , , , , , , , FG.TextMatrix(Row, FG.ColIndex("DiscountVal")), Me.TxtNoteSerial, Me.TxtNoteSerial1, 150

    End If

    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\LogFile Saving//////////////////////////////////////////////////////////////////////////
End Sub

Private Sub Fg_Click()
    'Command4_Click
End Sub

Function fill_bill_items_table() ' ﬁÊ„ Â–… «·œ«·Â »Õ›Ÿ «’‰«› «·›« Ê—… ›Ì ÃœÊ· „ƒﬁ  ·«” Œœ«„Â« ›Ì «· Ê“Ì⁄ ⁄·Ï «·„’—Ê›«   Ê«·›Ê« Ì—
    Dim bill_items As ADODB.Recordset
    Set bill_items = New ADODB.Recordset
    Dim StrSqlDel As String
    Dim RowNum As Integer
    bill_items.Open "[temp_bill_items]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    
    StrSqlDel = "delete From temp_bill_items"
    Cn.Execute StrSqlDel, , adExecuteNoRecords
 
    With FG

        For RowNum = 1 To FG.Rows - 1

            If FG.TextMatrix(RowNum, FG.ColIndex("Code")) <> "" Then
                bill_items.AddNew
                bill_items("ItemID").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Code")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Code"))))
                bill_items.update
            End If

        Next RowNum

    End With

End Function

Public Sub Fg_Journal_AfterEdit(ByVal Row As Long, _
                                ByVal Col As Long)
 
    Dim StrAccountCode As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long

    With Fg_Journal

        Select Case .ColKey(Col)
 
            Case "AccountName"
                '  .TextMatrix(Row, .ColIndex("userid")) = user_id
                        
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("AccountCode"), False, True)
                .TextMatrix(Row, .ColIndex("AccountCode")) = StrAccountCode
                .TextMatrix(Row, .ColIndex("ExpensesID")) = get_Expenses_id(StrAccountCode)
                .TextMatrix(Row, .ColIndex("LineNo1")) = setfoxy_Line

                If SystemOptions.UserInterface = ArabicInterface Then
                    strSQL = "select * from Expenses_accounts where Account_Code='" & StrAccountCode & "'"
                Else
                    strSQL = "select * from Expenses_accounts_eng where Account_Code='" & StrAccountCode & "'"
                End If
            
                rs.Open strSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                     
                If rs.RecordCount > 0 Then
                    .TextMatrix(Row, .ColIndex("des")) = IIf(IsNull(rs("parent_account").value), "", rs("parent_account").value)
                Else
                    .TextMatrix(Row, .ColIndex("des")) = ""
                End If

            Case "value"
                Dim sgl As String
               
                Me.TXTFactoryExpenses.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .Rows - 1, .ColIndex("Value"))
                '    sgl = "update  marakes_taklefa_temp  set value=0 where  line_no=" & Val(Fg_Journal.TextMatrix(Fg_Journal.Row, Fg_Journal.ColIndex("LineNo1")))
                '     Cn.Execute sgl, , adExecuteNoRecords
        
                '  Me.XPTxtVal.text = Format(Me.XPTxtVal.text, SystemOptions.SysDefCurrencyForamt)
        End Select

        Me.TXTFactoryExpenses.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .Rows - 1, .ColIndex("Value"))

        ' Me.XPTxtVal.text = Format(Me.XPTxtVal.text, SystemOptions.SysDefCurrencyForamt)
        'to Add new row if needed
        If Row = .Rows - 1 Then
            .Rows = .Rows + 1
        End If

        ' ReLineGrid
    End With

    ReLineGrid
End Sub

Private Sub ReLineGrid()
    Dim i As Integer
    Dim IntCounter As Integer

    With Fg_Journal

        For i = .FixedRows To .Rows - 1

            If .TextMatrix(i, .ColIndex("AccountCode")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("LineNo")) = IntCounter
            End If

        Next i

    End With

End Sub

Private Sub Fg_Journal_BeforeEdit(ByVal Row As Long, _
                                  ByVal Col As Long, _
                                  Cancel As Boolean)

    With Fg_Journal

        If Row > .FixedRows Then
            '  If .TextMatrix(Row - 1, .ColIndex("AccountCode")) = "" Then
            '      Cancel = True
            '  End If
        End If

        Select Case .ColKey(Col)

            Case "value"
                .ComboList = ""

            Case "des"
                .ComboList = ""
        
            Case "Order_No"
                .ComboList = ""
        
                '  Cancel = True
            
        End Select

    End With

End Sub

Private Sub Fg_Journal_DblClick()
    Exit Sub
  
    Static lNoteRow&, lNoteCol&, r&, c&

    With Fg_Journal
        ' clicking? no work
        'If Button <> 0 Then Exit Sub
        ' get mouse coordinates
        r = Fg_Journal.Row
        c = Fg_Journal.Col

        If Fg_Journal.ColKey(c) <> "Des" Then
            CboDes.Visible = False
            Exit Sub
        End If

        If Fg_Journal.TextMatrix(r, c) = "" Then
            'Exit Sub
        End If

        If .TextMatrix(r, .ColIndex("AccountCode")) = "" Then
            Exit Sub
        End If

        ' same cell or neighbour? no work
        '    If r = lNoteRow And C = lNoteCol Then Exit Sub
        '    If r = lNoteRow And C = lNoteCol + 1 Then Exit Sub

        ' other cell, hide current note, if any
        If lNoteRow >= 0 And lNoteCol >= 0 Then
            Fg_Journal.SetFocus
            lNoteRow = -1
            lNoteCol = -1
        End If

        ' no note to show? then bail out
        If r <= 0 Or c <= 0 Then Exit Sub
        If typename(Fg_Journal.Cell(flexcpData, r, c)) <> "String" Then
            TxtDes.text = ""
        Else
            '
            TxtDes.text = Fg_Journal.Cell(flexcpData, r, c)
        End If

        ' show new note
        CboDes.Move .CellLeft, .CellTop, .CellWidth, .CellHeight
        CboDes.Visible = True
        CboDes.ZOrder 0
        CboDes.SetFocus
        'save coordinates for next time
        lNoteRow = r
        lNoteCol = c
    End With

End Sub

Private Sub Fg_Journal_KeyUp(KeyCode As Integer, _
                             Shift As Integer)

    With Fg_Journal

        Select Case .ColKey(.Col)

            Case "Order_No"
                           
                If KeyCode = vbKeyF3 Then
                    Order_no_search.show
                    Order_no_search.RetrunType = 4
                End If

            Case "AccountName"

                If KeyCode = vbKeyF3 Then
                    FrmExpensesSearch.show
                    FrmExpensesSearch.RetrunType = 3
                End If
 
        End Select

    End With

End Sub

Private Sub Fg_Journal_StartEdit(ByVal Row As Long, _
                                 ByVal Col As Long, _
                                 Cancel As Boolean)
    Dim rs As New ADODB.Recordset
    Dim Rs1 As New ADODB.Recordset

    Dim strSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim StrComboList1 As String

    Dim Msg As String

    'Case "DebitName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a1%' Or ACCOUNTS.Account_Code Like 'a3%')"
    'Case "CreditName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a2%' Or ACCOUNTS.Account_Code Like 'a4%')"
    With Fg_Journal

        Select Case .ColKey(Col)

            Case "AccountName"

                '      StrSQL = "select * from Expenses_accounts"
                If SystemOptions.UserInterface = ArabicInterface Then
                    strSQL = "select * from Expenses_accounts"
                Else
                    strSQL = "select * from Expenses_accounts_eng "
                End If
                 
                rs.Open strSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                'StrComboList = Fg_Journal.BuildComboList(rs, "Account_Name", "Account_Code")
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = Fg_Journal.BuildComboList(rs, "Account_Name", "Account_Code")
                Else
                    StrComboList = Fg_Journal.BuildComboList(rs, "Account_NameEng", "Account_Code")
                End If
            
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList

            Case "opr_fullcode"
                strSQL = "  select fullcode,name from terms_operations "
                rs.Open strSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                StrComboList1 = Fg_Journal.BuildComboList(rs, "fullcode", "fullcode")

                If StrComboList1 <> "" Then
                    StrComboList1 = "|" & StrComboList1
                End If

                .ComboList = StrComboList1
         
        End Select

    End With

End Sub

Function fillExpensesFactoryGrid()
 
    '  «·’‰«⁄Ì…   ⁄»∆… «·«–Ê‰ «·„’—Ê›« 
    With Me.Fg_Journal
        .Rows = .FixedRows
        .ExtendLastCol = True
        .RowHeightMin = 300
        .Editable = flexEDKbdMouse
        .ExplorerBar = flexExSortShowAndMove

        '    .AutoSize 0, .Cols - 1, False
    End With

    Dim i As Integer
    Dim RsExp As ADODB.Recordset
    Dim My_SQL As String

    Set RsExp = New ADODB.Recordset
    My_SQL = "SELECT * from TblProductOrderFactoryexpenses where Transaction_ID=" & val(Me.XPTxtBillID.text)

    RsExp.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText
    Dim strSQL  As String

    With Me.Fg_Journal
        .Rows = 1
        .Clear flexClearScrollable

        If RsExp.RecordCount > 0 Then
            .Rows = RsExp.RecordCount + 1
            RsExp.MoveFirst

            For i = 1 To .Rows - 1
                   
                .TextMatrix(i, .ColIndex("LineNo")) = i
                
                .TextMatrix(i, .ColIndex("Accountcode")) = IIf(IsNull(RsExp.Fields("Accountcode").value), "", RsExp.Fields("Accountcode").value)
            
                .TextMatrix(i, .ColIndex("AccountName")) = IIf(IsNull(RsExp.Fields("AccountName").value), "", RsExp.Fields("AccountName").value)
               
                .TextMatrix(i, .ColIndex("value")) = IIf(Not IsNumeric(RsExp.Fields("value").value), 0, RsExp.Fields("value").value)
            
                .TextMatrix(i, .ColIndex("des")) = IIf(IsNull(RsExp.Fields("des").value), "", RsExp.Fields("des").value)
                RsExp.MoveNext
            Next

            RsExp.Close
        End If

        .RowHeight(-1) = 300
    End With

    With Me.Fg_Journal
        Me.TXTFactoryExpenses.text = .Aggregate(flexSTSum, .FixedRows - 1, .ColIndex("Value"), .Rows - 1, .ColIndex("Value"))
    End With
 
End Function

Private Sub Form_Activate()
    Set m_MnuShowNewItemsPrices = mdifrmmain.MnuInvPurchaseMnu2
    Set m_MenuViewList = mdifrmmain.MnuInvPurchaseMnu1
    Set m_MenuShowItemCostEffect = mdifrmmain.MnuInvPurchaseMnu4

End Sub

Private Sub CmdRetruns_Click()
    ShowRelatedTransactions val(Me.XPTxtBillID.text), 1
End Sub

Private Sub Grid_AfterEdit(ByVal Row As Long, _
                           ByVal Col As Long)
    Dim StrAccountCode As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long
    Dim sql As String
       
    With Grid

        Select Case .ColKey(Col)
   
            Case "ItemID"
          
                .TextMatrix(Row, Col) = Trim(.TextMatrix(Row, Col))
    
                strSQL = "select * from QRY_temp_bill_items where ItemID=" & Trim(.TextMatrix(Row, Col))
                Set rs = Nothing
                rs.Open strSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
            
                If Not (rs.BOF Or rs.EOF) Then
    
                    .TextMatrix(Row, .ColIndex("ItemName")) = IIf(IsNull(rs("ItemName").value), "", rs("ItemName").value)
                    .TextMatrix(Row, .ColIndex("ItemCode")) = IIf(IsNull(rs("ItemCode").value), "", rs("ItemCode").value)
                    
                Else
            
                    .TextMatrix(Row, .ColIndex("ItemName")) = ""
                    .TextMatrix(Row, .ColIndex("ItemCode")) = ""
                    .TextMatrix(Row, .ColIndex("ItemID")) = ""
 
                End If
 
                check_item_Exist_in_Grid val(.TextMatrix(Row, .ColIndex("ItemID"))), val(.TextMatrix(Row, .ColIndex("Note_value")))

            Case "ItemCode"
          
                .TextMatrix(Row, Col) = Trim(.TextMatrix(Row, Col))

                If .TextMatrix(Row, Col) = "" Then
                    Exit Sub
                End If

                strSQL = "select * from QRY_temp_bill_items where ItemCode='" & Trim(.TextMatrix(Row, Col)) & "'"
                Set rs = Nothing
                rs.Open strSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If Not (rs.BOF Or rs.EOF) Then
    
                    .TextMatrix(Row, .ColIndex("ItemName")) = IIf(IsNull(rs("ItemName").value), "", rs("ItemName").value)
                    .TextMatrix(Row, .ColIndex("ItemID")) = IIf(IsNull(rs("ItemID").value), "", rs("ItemID").value)
                    
                Else
                    .TextMatrix(Row, .ColIndex("ItemCode")) = ""
                    .TextMatrix(Row, .ColIndex("ItemName")) = ""
                    .TextMatrix(Row, .ColIndex("ItemID")) = ""
 
                End If
 
            Case "ItemName"
                  
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("ItemID"), False, True)
    
                Set ClsAcc = New ClsAccounts
      
                .TextMatrix(Row, .ColIndex("ItemID")) = StrAccountCode
                 
                strSQL = "select * from QRY_temp_bill_items where ItemID= " & val(StrAccountCode)
                Set rs = Nothing
                rs.Open strSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If Not (rs.BOF Or rs.EOF) Then
            
                    .TextMatrix(Row, .ColIndex("ItemCode")) = rs("ItemCode").value
                Else
                    .TextMatrix(Row, .ColIndex("ItemCode")) = ""
                    .TextMatrix(Row, .ColIndex("ItemID")) = ""
                    .TextMatrix(Row, .ColIndex("ItemName")) = ""
                   
                End If

        End Select

        'to Add new row if needed
        If Row = .Rows - 1 Then
            '    .Rows = .Rows + 1
        End If

    End With

    Expenses_update_total
End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Long, _
                            ByVal Col As Long, _
                            Cancel As Boolean)

    With Grid

        If .ColKey(Col) <> "ItemName" Then
            .ComboList = ""
        End If
   
    End With

End Sub

Private Sub Grid_Click()
    ' Expenses_update_total

End Sub

Public Function close_order2(order_no As String)
    Dim strSQL  As String
    Dim RsDetails As ADODB.Recordset
    Set RsDetails = New ADODB.Recordset
    Dim i As Integer
    Dim result As Integer
    result = 1
    strSQL = "select * from  items_qty_not_recieved_in_order where  order_no='" & order_no & "'"
    RsDetails.Open strSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Not (RsDetails.EOF Or RsDetails.BOF) Then

        For i = 1 To RsDetails.RecordCount

            If IsNull(RsDetails("net").value) Then result = 0: GoTo LL
            If RsDetails("net").value <> 0 Then
                result = 0
                GoTo LL
            End If

            RsDetails.MoveNext
        Next i
 
    End If

LL:
    Dim sql As String
    sql = "update Transactions Set closed = " & result & " Where Transaction_Type = 6 and order_no='" & Me.Txt_order_no & "'"
    Cn.Execute sql

End Function

Public Function items_qty_not_recieved_in_order(Item_ID As Integer, _
                                                order_no As String) As Integer
    Dim strSQL  As String
    Dim RsDetails As ADODB.Recordset
    Set RsDetails = New ADODB.Recordset
    strSQL = "select * from  items_qty_not_recieved_in_order where Item_ID=" & Item_ID & " and order_no='" & order_no & "'"
    RsDetails.Open strSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Not (RsDetails.EOF Or RsDetails.BOF) Then

        items_qty_not_recieved_in_order = IIf(IsNull(RsDetails("net").value), IIf(IsNull(RsDetails("sum_qty").value), 0, RsDetails("sum_qty").value), RsDetails("net").value)

    Else
        items_qty_not_recieved_in_order = 0
    End If

End Function

Function Retrive_orders_data(Transaction_ID As Integer)
    Dim strSQL  As String
    Dim RsDetails As ADODB.Recordset
    Set RsDetails = New ADODB.Recordset
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset

    Dim row_count As Integer
    Dim Num As Integer

    strSQL = "Select * from transactions where Transaction_ID=" & Transaction_ID
    Set rs = New ADODB.Recordset
    rs.Open strSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount < 1 Then
 
        Exit Function
    Else
        DBCboClientName.BoundText = IIf(IsNull(rs("CusID").value), "", rs("CusID").value)
        Me.Dccurrency.BoundText = IIf(IsNull(rs("Currency_id").value), "", rs("Currency_id").value)
        'txt_Currency_rate.text = IIf(IsNull(rs("Currency_rate").value), 1, (rs("Currency_rate").value))
        TxtLcNo.text = IIf(IsNull(rs("LcNo").value), "", (rs("LcNo").value))
    End If

    If rs.EOF Or rs.BOF Then
        Exit Function
    End If

    strSQL = "SELECT TblItems.HaveSerial, * FROM TblItems INNER JOIN Transaction_Details " & "ON TblItems.ItemID = Transaction_Details.Item_ID INNER JOIN dbo.TblUnites ON dbo.Transaction_Details.UnitID = dbo.TblUnites.UnitID"
    strSQL = strSQL + " where Transaction_ID=" & Transaction_ID

    RsDetails.Open strSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    XPTxtSum.text = ""

    FG.Rows = 2
    FG.Clear flexClearScrollable, flexClearEverything

    If Not (RsDetails.EOF Or RsDetails.BOF) Then
        row_count = FG.Rows
    
        If FG.TextMatrix(row_count - 1, FG.ColIndex("Code")) = "" Then
            row_count = row_count - 1
        End If
     
        FG.Rows = RsDetails.RecordCount + row_count

        For Num = row_count To FG.Rows - 1 'RsDetails.RecordCount
    
            FG.TextMatrix(Num, FG.ColIndex("order_no")) = IIf(IsNull(RsDetails("order_no")), "", (RsDetails("order_no").value))
            FG.TextMatrix(Num, FG.ColIndex("OrderArrivalDate")) = DTArrivalDate.value
         
            FG.TextMatrix(Num, FG.ColIndex("Code")) = IIf(IsNull(RsDetails("Item_ID")), "", (RsDetails("Item_ID").value))
            FG.TextMatrix(Num, FG.ColIndex("Name")) = IIf(IsNull(RsDetails("Item_ID")), "", Trim(RsDetails("Item_ID").value))
        
            '          FG.TextMatrix(Num, FG.ColIndex("Count")) = items_qty_not_recieved_in_order(FG.TextMatrix(Num, FG.ColIndex("Code")), FG.TextMatrix(Num, FG.ColIndex("order_no")))
            FG.TextMatrix(Num, FG.ColIndex("Count")) = IIf(IsNull(RsDetails("Showqty")), "", (RsDetails("Showqty").value))
        
            FG.TextMatrix(Num, FG.ColIndex("Price")) = IIf(IsNull(RsDetails("ShowPrice")), "", (RsDetails("ShowPrice").value))
            FG.TextMatrix(Num, FG.ColIndex("ItemCase")) = IIf(IsNull(RsDetails("ItemCase")), "", (RsDetails("ItemCase").value))
            FG.TextMatrix(Num, FG.ColIndex("DiscountType")) = IIf(IsNull(RsDetails("ItemDiscountType")), "", (RsDetails("ItemDiscountType").value))
            FG.TextMatrix(Num, FG.ColIndex("DiscountVal")) = IIf(IsNull(RsDetails("ItemDiscount")), "", (RsDetails("ItemDiscount").value))
            FG.TextMatrix(Num, FG.ColIndex("ColorID")) = IIf(IsNull(RsDetails("ColorID")), 1, (RsDetails("ColorID").value))
        
            FG.TextMatrix(Num, FG.ColIndex("ItemSize")) = IIf(IsNull(RsDetails("ItemSize")), 1, (RsDetails("ItemSize").value))
            FG.TextMatrix(Num, FG.ColIndex("ClassId")) = IIf(IsNull(RsDetails("ClassId")), 1, (RsDetails("ClassId").value))
        
            If RsDetails("HaveSerial") = True Then
                FG.TextMatrix(Num, FG.ColIndex("HaveSerial")) = True
            End If
        
            FG.Cell(flexcpData, Num, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitID")), "", (RsDetails("UnitID").value))
            FG.TextMatrix(Num, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitName")), "", (RsDetails("UnitName").value))
        
            RsDetails.MoveNext
            ' Debug.Print Num
            ' If FG.Rows > 10 Then
            '     If Num = 8 Then FG.Refresh
            ' End If
        Next Num

    End If

End Function

Private Sub Grid_StartEdit(ByVal Row As Long, _
                           ByVal Col As Long, _
                           Cancel As Boolean)
    Dim rs As New ADODB.Recordset
    Dim strSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim Msg As String

    'Case "DebitName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a1%' Or ACCOUNTS.Account_Code Like 'a3%')"
    'Case "CreditName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a2%' Or ACCOUNTS.Account_Code Like 'a4%')"
    With Grid

        Select Case .ColKey(Col)

            Case "ItemName"
       
                strSQL = "Select * from QRY_temp_bill_items"
                
                rs.Open strSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                StrComboList = grid4.BuildComboList(rs, "ItemName", "ItemID")
                Debug.Print strSQL
 
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
        End Select

    End With

End Sub

Private Sub Grid1_Click()

    With GRID1

        Select Case .Col

            Case 2
 
                '       If .Cell(flexcpChecked, .Row, .ColIndex("select")) = flexChecked Then
                '            Retrive_orders_data (Val(.TextMatrix(.Row, .ColIndex("Transaction_ID"))))
                '
                '
                '        End If

                With FG
                    .Clear flexClearScrollable, flexClearEverything
                    .Rows = 1
       
                End With
 
                fillVchr

            Case 7
                FrmInpout.Retrive val(.TextMatrix(.Row, 1))

            Case 8
                ShowGL_cc .TextMatrix(.Row, .ColIndex("NoteSerial")), , 200

        End Select

    End With

End Sub

Function fillVchr()
    Dim i As Integer
        
    With GRID1

        For i = 1 To .Rows - 1

            If .Cell(flexcpChecked, i, .ColIndex("select")) = flexChecked Then
                Retrive_orders_data (val(.TextMatrix(i, .ColIndex("Transaction_ID"))))
            
            End If

        Next i

    End With

End Function

Private Sub GRID2_Click()

    With GRID2

        If .Cell(flexcpChecked, .Row, .ColIndex("select")) = flexChecked Then
            Retrive_orders_data (val(GRID2.TextMatrix(GRID2.Row, GRID2.ColIndex("Transaction_ID"))))
            
        End If

    End With

End Sub
 
Private Function check_item_Exist_in_Grid(ItemID As Integer, _
                                          value As Single, _
                                          Optional addition As Boolean)
    Dim i As Integer
    On Error Resume Next

    With FG

        For i = 1 To FG.Rows - 1

            If .TextMatrix(i, .ColIndex("Code")) = CStr(ItemID) Then
                If addition = False Then
                    .TextMatrix(i, .ColIndex("LineShahn")) = value
                Else
                    .TextMatrix(i, .ColIndex("LineShahn")) = val(.TextMatrix(i, .ColIndex("LineShahn"))) + value
                End If

                Exit Function
    
            End If

        Next i

    End With
 
End Function

Private Sub grid4_AfterEdit(ByVal Row As Long, _
                            ByVal Col As Long)
    Dim StrAccountCode As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long
    Dim sql As String
       
    With grid4

        Select Case .ColKey(Col)
   
            Case "ItemID"
          
                .TextMatrix(Row, Col) = Trim(.TextMatrix(Row, Col))

                If .TextMatrix(Row, Col) = "" Then
                    Exit Sub
                End If

                strSQL = "select * from QRY_temp_bill_items where ItemID=" & Trim(.TextMatrix(Row, Col))
                Set rs = Nothing
                rs.Open strSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If Not (rs.BOF Or rs.EOF) Then
    
                    .TextMatrix(Row, .ColIndex("ItemName")) = IIf(IsNull(rs("ItemName").value), "", rs("ItemName").value)
                    .TextMatrix(Row, .ColIndex("ItemCode")) = IIf(IsNull(rs("ItemCode").value), "", rs("ItemCode").value)
                    
                Else
            
                    .TextMatrix(Row, .ColIndex("ItemName")) = ""
                    .TextMatrix(Row, .ColIndex("ItemCode")) = ""
                    .TextMatrix(Row, .ColIndex("ItemID")) = ""
 
                End If
 
                check_item_Exist_in_Grid val(.TextMatrix(Row, .ColIndex("ItemID"))), val(.TextMatrix(Row, .ColIndex("Note_value")))

            Case "ItemCode"
          
                .TextMatrix(Row, Col) = Trim(.TextMatrix(Row, Col))
         
                strSQL = "select * from QRY_temp_bill_items where ItemCode='" & Trim(.TextMatrix(Row, Col)) & "'"
                Set rs = Nothing
                rs.Open strSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If Not (rs.BOF Or rs.EOF) Then
    
                    .TextMatrix(Row, .ColIndex("ItemName")) = IIf(IsNull(rs("ItemName").value), "", rs("ItemName").value)
                    .TextMatrix(Row, .ColIndex("ItemID")) = IIf(IsNull(rs("ItemID").value), "", rs("ItemID").value)
                    
                Else
            
                    .TextMatrix(Row, .ColIndex("ItemName")) = ""
                    .TextMatrix(Row, .ColIndex("ItemID")) = ""
                    
                    .TextMatrix(Row, .ColIndex("ItemCode")) = ""
                End If
 
            Case "ItemName"
                  
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("ItemID"), False, True)
    
                Set ClsAcc = New ClsAccounts
      
                .TextMatrix(Row, .ColIndex("ItemID")) = StrAccountCode
                 
                strSQL = "select * from QRY_temp_bill_items where ItemID= " & val(StrAccountCode)
                Set rs = Nothing
                rs.Open strSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If Not (rs.BOF Or rs.EOF) Then
            
                    .TextMatrix(Row, .ColIndex("ItemCode")) = rs("ItemCode").value
                    .TextMatrix(Row, .ColIndex("ItemID")) = rs("ItemID").value
                Else
                    .TextMatrix(Row, .ColIndex("ItemCode")) = ""
                    .TextMatrix(Row, .ColIndex("ItemID")) = ""
                    .TextMatrix(Row, .ColIndex("ItemName")) = ""
                   
                End If

        End Select

        'to Add new row if needed
        If Row = .Rows - 1 Then
            '    .Rows = .Rows + 1
        End If

    End With

    update_finincial_invoice_total
End Sub

Private Sub grid4_BeforeEdit(ByVal Row As Long, _
                             ByVal Col As Long, _
                             Cancel As Boolean)

    With grid4

        If .ColKey(Col) <> "ItemName" Then
            .ComboList = ""
        End If
   
    End With

End Sub

Private Sub grid4_Click()
    'update_finincial_invoice_total
       
End Sub

Private Sub grid4_StartEdit(ByVal Row As Long, _
                            ByVal Col As Long, _
                            Cancel As Boolean)
    Dim rs As New ADODB.Recordset
    Dim strSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim Msg As String

    'Case "DebitName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a1%' Or ACCOUNTS.Account_Code Like 'a3%')"
    'Case "CreditName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a2%' Or ACCOUNTS.Account_Code Like 'a4%')"
    With grid4

        Select Case .ColKey(Col)

            Case "ItemName"
       
                strSQL = "Select * from QRY_temp_bill_items"
                
                rs.Open strSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                StrComboList = grid4.BuildComboList(rs, "ItemName", "ItemID")
                Debug.Print strSQL
 
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
        End Select

    End With

End Sub

Private Sub ISButton1_Click()
    FrmLC.show
    FrmLC.Retrive Trim(Me.TxtLcNo.text)

End Sub

Private Sub LblDiscountsTotal_Change()
    LblDiscountsTotalView.Caption = Format(val(LblDiscountsTotal.Caption), "#,###." & String(Abs(SystemOptions.SysDefCurrencyForamt), "0"))
End Sub

Private Sub LblTotal_Change()

    If CboPayMentType.ListIndex = 1 Then
        XPTxtValue(1).text = LblTotal.Caption
    ElseIf CboPayMentType.ListIndex = 0 Then
        XPTxtValue(0).text = LblTotal.Caption
    End If
         
    LblTotalView.Caption = Format(val(LblTotal.Caption), "#,###." & String(Abs(SystemOptions.SysDefCurrencyForamt), "0"))
End Sub

Private Sub LblTotal_MouseMove(Button As Integer, _
                               Shift As Integer, _
                               X As Single, _
                               Y As Single)
    LblTotal.ToolTipText = WriteNo(LblTotal.Caption, 0, True)
End Sub

Private Sub LblTotalAll_Change()
    LblTotalAllview.Caption = Format(val(LblTotalAll.Caption), "#,###." & String(Abs(SystemOptions.SysDefCurrencyForamt), "0"))
End Sub

Private Sub m_FrmSearch_Unload(Cancel As Integer)
    Set m_FrmSearch = Nothing
End Sub

Private Sub m_MenuShowItemCostEffect_Click()

    If Me.TxtModFlg.text = "R" Then
        ShowItemCostEffectForTrans 1, , Trim$(Me.TxtTransSerial.text)
    End If

End Sub

Private Sub m_MenuViewList_Click()
    Dim FrmView As FrmViewList
    Dim FG As VSFlex8UCtl.vsFlexGrid
    Dim strSQL As String
    Dim rs As ADODB.Recordset
    Dim StrComboList As String
    Dim GrdBack As ClsBackGroundPic
    Dim cProgress As ClsProgress
    Dim BolFrmLoaded As Boolean
    Set FrmView = New FrmViewList
    Set FG = FrmView.vsfGroup1.vsFlexGrid

    With FG
        .Cols = 9
        .RowHeightMin = 320
        .TextMatrix(0, 0) = "—ﬁ„ «·»—‰«„Ã"
        .ColKey(0) = "Transaction_ID"
        .TextMatrix(0, 1) = "—ﬁ„ «·›« Ê—…"
        .TextMatrix(0, 2) = " «—ÌŒ «·›« Ê—…"
        .ColDataType(2) = flexDTDate
        .TextMatrix(0, 3) = "«”„ «·„Ê—œ"
        .TextMatrix(0, 4) = "ÿ—Ìﬁ… «·œ›⁄"
        StrComboList = "#0;‰ﬁœÏ|#1;√Ã·"
        .ColComboList(4) = StrComboList
    
        .TextMatrix(0, 5) = "«”„ «·„Œ“‰"
        .TextMatrix(0, 6) = "‰Ê⁄ «·Œ’„"
        .TextMatrix(0, 7) = "ﬁÌ„… «·Œ’„"
        .TextMatrix(0, 8) = "≈Ã„«·Ï «·›« Ê—…"

        ',
        'QryTransactionsTotal.TransSum
        'QryTransactionsTotal.TransNet,
        If SystemOptions.SysDataBaseType = SQLServerDataBase Then
            strSQL = "SELECT TOP 100 PERCENT QryTransactionsTotal.Transaction_ID," & "QryTransactionsTotal.Transaction_Serial, QryTransactionsTotal.Transaction_Date, " & "dbo.TblCustemers.CusName, QryTransactionsTotal.PaymentType, dbo.TblStore.StoreName," & "QryTransactionsTotal.Trans_DiscountType,QryTransactionsTotal.Trans_Discount ," & "QryTransactionsTotal.TotalAfterTax "
            strSQL = strSQL + " FROM dbo.QryTransactionsTotal() QryTransactionsTotal LEFT OUTER JOIN "
            strSQL = strSQL + "dbo.TblStore ON QryTransactionsTotal.StoreID = dbo.TblStore.StoreID " & "LEFT OUTER JOIN dbo.TblCustemers ON QryTransactionsTotal.CusID = dbo.TblCustemers.CusID"
            strSQL = strSQL + " Where (QryTransactionsTotal.Transaction_Type = 1)"
            strSQL = strSQL + " ORDER BY QryTransactionsTotal.Transaction_ID "
        ElseIf SystemOptions.SysDataBaseType = AccessDataBase Then
            strSQL = "SELECT QryTransactionsTotal.Transaction_ID , QryTransactionsTotal.Transaction_Serial," & "QryTransactionsTotal.Transaction_Date,TblCustemers.CusName, QryTransactionsTotal.PaymentType," & "TblStore.StoreName,TblEmployee.Emp_Name ,QryTransactionsTotal.Trans_DiscountType," & "QryTransactionsTotal.Trans_Discount,QryTransactionsTotal.TotalAfterTax "
            strSQL = strSQL + "FROM (TblEmployee RIGHT JOIN (TblCustemers RIGHT JOIN QryTransactionsTotal " & "ON TblCustemers.CusID = QryTransactionsTotal.CusID) ON TblEmployee.Emp_ID = QryTransactionsTotal.Emp_ID) " & "LEFT JOIN TblStore ON QryTransactionsTotal.StoreID = TblStore.StoreID "
            strSQL = strSQL + " WHERE QryTransactionsTotal.Transaction_Type=1 "
            strSQL = strSQL + " Order  By QryTransactionsTotal.Transaction_ID"
        End If

        Set rs = New ADODB.Recordset
        rs.Open strSQL, Cn, adOpenKeyset, adLockReadOnly, adAsyncExecute + adAsyncFetch
        Set cProgress = New ClsProgress
        BolFrmLoaded = True
        cProgress.ProgressType = Waiting
        cProgress.StartProgress

        Do While rs.State = adStateExecuting
            DoEvents
        Loop

        If BolFrmLoaded = True Then
            cProgress.StopProgess
            Set cProgress = Nothing
        End If

        Set .DataSource = rs
        .TextMatrix(0, 0) = "—ﬁ„ «·»—‰«„Ã"
        .ColKey(0) = "Transaction_ID"
        .TextMatrix(0, 1) = "—ﬁ„ «·›« Ê—…"
        .TextMatrix(0, 2) = " «—ÌŒ «·›« Ê—…"
        .ColDataType(2) = flexDTDate
        .TextMatrix(0, 3) = "«”„ «·„Ê—œ"
        .TextMatrix(0, 4) = "ÿ—Ìﬁ… «·œ›⁄"
        StrComboList = "#0;‰ﬁœÏ|#1;√Ã·"
        .ColComboList(4) = StrComboList
        .TextMatrix(0, 5) = "«”„ «·„Œ“‰"
        .TextMatrix(0, 6) = "‰Ê⁄ «·Œ’„"
        .TextMatrix(0, 7) = "ﬁÌ„… «·Œ’„"
        .TextMatrix(0, 8) = "≈Ã„«·Ï «·›« Ê—…"
        .ColKey(8) = "TotalAfterTax"
        'Rs.Close
        'Set Rs = Nothing
    End With

    Set GrdBack = New ClsBackGroundPic
    FrmView.vsfGroup1.vsFlexGrid.WallPaper = GrdBack.Picture
    FrmView.vsfGroup1.SetRTL = True
    FrmView.vsfGroup1.TotalOnColKey = "TotalAfterTax"
    FrmView.vsfGroup1.update
    FrmView.BolRetrunOnDblClick = True
    FrmView.SetDblClickRetrun Me, "Transaction_ID"
    FrmView.Caption = "⁄—÷ ‘Ã—Ï ÃœÊ·Ï ·›Ê« Ì— «·„‘ —Ì« "
    FrmView.show
End Sub

Private Sub m_MnuShowNewItemsPrices_Click()

    If Not NewGrid Is Nothing Then
        NewGrid.ShowNewItemsPrice
    End If

End Sub

Private Sub Txt_EXport_Change()
    Me.TXTToTAlELSHahn.text = IIf(Not IsNumeric(Txt_EXport.text), 0, val(Txt_EXport.text)) + IIf(Not IsNumeric(txt_total_bill.text), 0, val(txt_total_bill.text)) + IIf(Not IsNumeric(TXTFactoryExpenses.text), 0, val(TXTFactoryExpenses.text))
End Sub

Private Sub Txt_order_no_Change()

    With Me.grid4
        .Rows = .FixedRows
 
    End With

    With Me.Grid
        .Rows = .FixedRows
 
    End With

    If Txt_order_no.text = "" Then
        txt_total_bill.text = ""
        Txt_EXport.text = ""
    End If

    Command4_Click
    Command2_Click
    Command3_Click
    Dim Transaction_ID As String
    Dim Transaction_Type As Integer

    If CBoBasedON.ListIndex = 1 Then
        Transaction_Type = 29
    ElseIf CBoBasedON.ListIndex = 2 Then
        Transaction_Type = 17
    Else
        Transaction_Type = 0
        Exit Sub
    End If

    Transaction_ID = get_transactionData("order_no", Txt_order_no.text, "Transaction_ID", Transaction_Type)

    If Me.TxtModFlg <> "R" And Me.TxtModFlg <> "" Then
        Retrive_orders_data (val(Transaction_ID))
    End If

End Sub

Private Sub TXT_total_payments_Change()
    'Me.TXTToTAlELSHahn.text = IIf(Not IsNumeric(Txt_EXport.text), 0, Val(Txt_EXport.text)) + IIf(Not IsNumeric(TXT_total_payments.text), 0, Val(TXT_total_payments.text))
End Sub

Private Sub txt_ORDER_NO_KeyUp(KeyCode As Integer, _
                               Shift As Integer)

    If KeyCode = vbKeyF3 Then

        If CBoBasedON.ListIndex = 0 Then
            Exit Sub
                
        Else
                
            Txt_order_no.text = ""
            Order_no_search.show
            Order_no_search.RetrunType = 3
            Order_no_search.lblSpecificsearch.Caption = val(CBoBasedON.ListIndex)
        End If

    End If

End Sub

Private Sub txt_total_bill_Change()
    Me.TXTToTAlELSHahn.text = IIf(Not IsNumeric(Txt_EXport.text), 0, val(Txt_EXport.text)) + IIf(Not IsNumeric(txt_total_bill.text), 0, val(txt_total_bill.text)) + IIf(Not IsNumeric(TXTFactoryExpenses.text), 0, val(TXTFactoryExpenses.text))

End Sub

Private Sub TXTFactoryExpenses_Change()
    Me.TXTToTAlELSHahn.text = IIf(Not IsNumeric(Txt_EXport.text), 0, val(Txt_EXport.text)) + IIf(Not IsNumeric(txt_total_bill.text), 0, val(txt_total_bill.text)) + IIf(Not IsNumeric(TXTFactoryExpenses.text), 0, val(TXTFactoryExpenses.text))

End Sub

Private Sub TxtFillData_Change()

    If TxtFillData.text = "F" Then
        NewGrid.Calculate 1, , , True
    End If

End Sub

Private Sub TxtLcNo_KeyUp(KeyCode As Integer, _
                          Shift As Integer)
       
    If KeyCode = vbKeyF3 Then
        Order_no_search3.show
        Order_no_search3.RetrunType = 2
         
    End If

End Sub

Private Sub TxtSearchCode_KeyPress(KeyAscii As Integer)

    Dim CUSTID As Integer

    If KeyAscii = vbKeyReturn Then
        GetCustomersDetail CUSTID, , TxtSearchCode.text, 2
        DBCboClientName.BoundText = CUSTID
    End If

End Sub

Private Sub TxtStoreID_KeyDown(KeyCode As Integer, Shift As Integer)
 Dim StoreId As Integer

    If KeyCode = vbKeyReturn Then
    StoreId = getStoreInformatin(TxtStoreID)
        DCboStoreName.BoundText = StoreId
    End If
End Sub

Private Sub XPBtnMove_Click(Index As Integer)

    'On Error GoTo ErrTrap
    If Me.TxtModFlg.text = "N" Then
        clear_all Me
        Me.TxtModFlg.text = "R"
        XPBtnMove_Click (1)
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

    Command2_Click
    Command4_Click
    Exit Sub
ErrTrap:
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        If Me.TxtModFlg.text = "R" Then
            '        Cmd_Click (0)
        Else
            '      SendKeys "{TAB}"
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

    If KeyCode = vbKeyF3 Then
        If Cmd(5).Enabled = False Then Exit Sub
        Cmd_Click (5)
    End If

    If KeyCode = vbKeyF6 Then
        If Cmd(7).Enabled = False Then Exit Sub
        Cmd_Click (7)
    End If

    If KeyCode = vbKeyF2 Then
        If TxtModFlg.text = "N" Or TxtModFlg.text = "E" Then
            'XPBtnAdd_Click
        End If
    End If

    If KeyCode = vbKeyF3 Then
        If TxtModFlg.text = "N" Or TxtModFlg.text = "E" Then
            'XPBtnRemove_Click
        End If
    End If

    If KeyCode = vbKeyF5 Then
        If TxtModFlg.text = "N" Or TxtModFlg.text = "E" Then
            XPBtnNewClients_Click
        End If
    End If

    If Shift = 2 Then
        If KeyCode = vbKeySpace Then
            If TxtModFlg.text = "N" Or TxtModFlg.text = "E" Then
                'XPFillData_Click
            End If
        End If
    End If

    If Shift = 2 Then
        XPTab301.SetFocus

        If KeyCode = vbKeyTab Then
            If XPTab301.CurrTab = 0 Then
                XPTab301.CurrTab = 1

                If XPChkPayType(0).Enabled = True Then
                    XPChkPayType(0).SetFocus
                End If

            Else
                XPTab301.CurrTab = 0
                FG.SetFocus
            End If
        End If
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

Private Sub Form_Load()
    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500
       
    ScreenNameArabic = "  ›« Ê—… „‘—Ì«  "
    ScreenNameEnglish = " Purchase Invoice  "
    RegisterLogInOut Me.name, ScreenNameArabic, ScreenNameEnglish, "1", 150

    Dim RsClients As New ADODB.Recordset
    Dim strSQL As String
    Dim Num As Integer
    Dim StrList As String
    Dim Dcombos As ClsDataCombos
    Dim BGround As New ClsBackGroundPic
    Dim RsNote As New ADODB.Recordset

    On Error GoTo ErrTrap
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
    Set Cmd(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
    Set Cmd(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
    Set Cmd(5).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    Set Cmd(7).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Print").Picture
    Set CmdHelp.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Help").Picture
    Dim My_SQL As String
    'My_SQL = "  select branch_id,branch_name from TblBranchesData   "
    'fill_combo dcBranch, My_SQL
  
    'dcBranch
    If SystemOptions.usertype <> UserAdminAll Then
 
        Me.dcBranch.Enabled = False
    End If

    SetDtpickerDate XPDtbBill
    Set NewGrid = New ClsGrid
    NewGrid.GridTrans = PurchaseTransaction
    Set NewGrid.Grid = Me.FG
    Set NewGrid.TxtInvID = Me.XPTxtBillID
    Set NewGrid.TxtModFlag = Me.TxtModFlg
    Set NewGrid.TxtTotal = Me.XPTxtSum
    Set NewGrid.CboDiscount_Type = XPCboDiscountType
    Set NewGrid.TxtDiscount_Val = XPTxtDiscountVal
    Set NewGrid.TxtValueCash = XPTxtValue(0)
    Set NewGrid.TxtValueDelay = XPTxtValue(1)
    Set NewGrid.TxtValuechque = XPTxtValue(2)
    Set NewGrid.TxtFillData = TxtFillData
    Set NewGrid.DtpBillDate = Me.XPDtbBill
    Set NewGrid.STORENAME = Me.DCboStoreName
    Set NewGrid.GrdTBar = Me.TBar
    '-----------------------------------------------------------------------------
    Set NewGrid.TxtTaxValue = Me.XPTxtTaxValue
    Set NewGrid.TxtAddTax = Me.TxtTaxAddValue
    Set NewGrid.TxtStampTax = Me.TxtTaxStampValue
    Set NewGrid.TxtServiceTax = Me.TxtTaxServiceValue
    Set NewGrid.LblTotalQty = Me.LblTotalQty
    '-----------------------------------------------------------------------------
    ' ⁄»∆… »Ì«‰«  «·√’‰«›
    Set NewGrid.DCboItemName = DCboItemsName
    Set NewGrid.DCboItemCode = DCboItemsCode
    Set NewGrid.CboItemCase = CboItemCase
    Set NewGrid.CmdAddData = CmdAdd
    Set NewGrid.TxtSerial = TxtSerial
    Set NewGrid.TxtQuantity = TxtQuantity
    Set NewGrid.TxtPrice = TxtPrice
    Set NewGrid.LblItemsCount = Me.LblItemsCount
    Set NewGrid.LblTotalAll = Me.LblTotalAll
    Set NewGrid.LblDiscountsTotal = Me.LblDiscountsTotal
    
    Set NewGrid.LblCommision = Me.LblCommision
    
    Set NewGrid.LblTaxSalesValue = Me.lbl(25)
    Set NewGrid.LblTaxAddValue = Me.lbl(32)
    Set NewGrid.LblTaxStampValue = Me.lbl(33)
    Set NewGrid.LblTaxServiceValue = Me.lbl(49)

    FG.WallPaper = BGround.Picture

    AddTip
    XPTab301.CurrTab = 0
    XPDtbBill.value = Date

    With XPCboDiscountType
        .AddItem "·«ÌÊÃœ Œ’„"
        .AddItem "Œ’„ »ﬁÌ„…"
        .AddItem "Œ’„ »‰”»…"
    End With

    With CboPayMentType
        .AddItem "‰ﬁœ«"
        .AddItem "¬Ã·"
    End With

    With Me.CBoBasedON
        .Clear
        .AddItem "»·«"
        .AddItem "√„— ‘—¡"
        .AddItem "›« Ê—… „»œ∆ÌÂ"
    End With

    NewGrid.fillgrid
    Set Dcombos = New ClsDataCombos
    Dcombos.GetBoxes Me.DcboBox
    Dcombos.GetUsers Me.DCboUserName
    Dcombos.GetCustomersSuppliers 3, Me.DBCboClientName, True
    Dcombos.GetBranches dcBranch
    Dcombos.GetDocTypebyid Me.DCDocTypes, 22, val(Me.dcBranch.BoundText)
    strSQL = "  select  BankID,BankName  from BanksData   "
    fill_combo dcbanks, strSQL
 
    strSQL = " select id,code from currency"
 
    fill_combo Me.Dccurrency, strSQL

    strSQL = " select id,Project_name from projects"
 
    fill_combo Me.DCproject, strSQL
      
    Set cSearchDcbo(0) = New clsDCboSearch
    Set cSearchDcbo(0).Client = Me.DBCboClientName
    cSearchDcbo(0).SetBuddyText Me.TxtCusID
    Dcombos.GetStores Me.DCboStoreName
    Set cSearchDcbo(2) = New clsDCboSearch
    Set cSearchDcbo(2).Client = Me.DCboStoreName
    cSearchDcbo(2).SetBuddyText Me.TxtStoreID
    '-----------------------------------------
    SetDtpickerDate Me.DtpDelayDate
    '≈⁄œ«œ Ã—œ «·√ﬁ”«ÿ
    ChkInstall.value = Unchecked
    ChkInstall.Enabled = False

    With Me.FgInstallments
        .Rows = .FixedRows
        Set .WallPaper = BGround.Picture
        .RowHeightMin = 300
        .AutoSize 0, .Cols - 1, False
    End With

    With Me.FgCheques
        .Rows = .FixedRows
        Set .WallPaper = BGround.Picture
        .RowHeightMin = 300
        .AutoSize 0, .Cols - 1, False
    End With

    Me.XPChkTAX.value = vbUnchecked
    XPChkTAX_Click
    Me.ChkTaxAdd.value = vbUnchecked
    ChkTaxAdd_Click
    Me.ChkTaxStamp.value = vbUnchecked
    ChkTaxStamp_Click
    Me.ChkTaxSerivce.value = vbUnchecked
    ChkTaxSerivce_Click

    If SystemOptions.UserInterface = EnglishInterface Then
     
        SetInterface Me
        ChangeLang
    End If
 
    '-----------------------------------------------------------------------------
    Dim rsOut As New ADODB.Recordset
    Dim Msg As String
    Set rsOut = New ADODB.Recordset
    rsOut.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable

    If Not (rsOut.EOF Or rsOut.BOF) Then
 
        If rsOut!checkinpo = True Then
            strSQL = "SELECT * FROM Transactions WHERE Transaction_Type=22 Order by Transaction_ID"
            Set rs = New ADODB.Recordset
            rs.Open strSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText

            If Not (rs.BOF Or rs.EOF) Then
                rs.MoveLast
            End If

            Retrive
            Me.TxtModFlg.text = "R"
            ' Resize_Form Me, TransactionSize
            BillType = 22
    
        Else
            strSQL = "SELECT * FROM Transactions WHERE Transaction_Type=1"
            Set rs = New ADODB.Recordset
            rs.Open strSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

            If Not (rs.BOF Or rs.EOF) Then
                rs.MoveLast
            End If

            Retrive
            Me.TxtModFlg.text = "R"
            '  Resize_Form Me, TransactionSize
            BillType = 1
            Exit Sub
        End If
    End If

    Me.TxtModFlg.text = "R"
    Command2_Click
    Command4_Click

    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If

ErrTrap:
End Sub

Private Sub ChangeLang()
    Dim XPic As IPictureDisp
    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic
CmdAttach.Caption = "Attachments"


    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic
    lbl(69).Caption = "I. Manual No"
    lbl(68).Caption = "LC No"
    Label4.Caption = "Doc Type"
    'Label3.Caption = "Shioment No."
    'Label4.Caption = "Order No."
    Ele(13).Visible = False
    Command4.Caption = "Financial Invoice"
    XPCboDiscountType.Clear
    XPCboDiscountType.AddItem "NO"
    XPCboDiscountType.AddItem "Value"
    XPCboDiscountType.AddItem "Percent"
    CboPayMentType.Clear
    CboPayMentType.AddItem "Cash"
    CboPayMentType.AddItem "Credit"
    Me.XPTab301.TabCaption(0) = "Items"
    Me.XPTab301.TabCaption(1) = "Securities"
    Me.XPTab301.TabCaption(2) = "Comment On Invoice"
    Me.XPTab301.TabCaption(5) = "Expenses Vouchers"
    Me.XPTab301.TabCaption(4) = "Purchase Orders and Performa Invoices"
    Me.XPTab301.TabCaption(3) = "Fn invoices"
    Me.XPTab301.TabCaption(6) = "Estimated Expenses"
    Me.XPTab301.TabCaption(7) = " Linked voucher"
    '«·ÊﬁÊ› ⁄‰œ «›Ê« Ì— «·„«·Ì… ⁄‘«‰ «··€Â «·«‰Ã·Ì“Ì…
    lbl(57).Caption = "Purcahase order and Performa Invoices"
    lbl(64).Caption = "Financial Invoices"
    Label19.Caption = "Discretionary Expenses"

    lbl(52).Caption = "RCV VCHR No."
    lbl(58).Caption = "Project"
    Label3.Caption = "Branch"
    lbl(65).Caption = " Based On"
    lbl(56).Caption = "O. Arival Date"
    lbl(66).Caption = "NO."
    lbl(63).Caption = "Total Qty "

    With CBoBasedON
        CBoBasedON.Clear
        CBoBasedON.AddItem "WithOut"
        CBoBasedON.AddItem "Purchase Order"
        CBoBasedON.AddItem "Performa Invoices"

    End With

    ' lbl(53).Caption = "Order No:"
    lbl(54).Caption = "Expenses"
    '  lbl(56).Caption = "Payment Voucher"
    '  lbl(57).Caption = "Total Payment"
    lbl(60).Caption = "Total"
 
    lbl(51).Caption = "Total Expenses"
    Command3.Caption = "View P.O. For Vendor"
    Me.Caption = "Purchase Invoice"
    Ele(6).Caption = Me.Caption
    Cmd(0).Caption = "New"
    Cmd(1).Caption = "Edit"
    Cmd(2).Caption = "Save"
    Cmd(3).Caption = "Undo"
    Cmd(4).Caption = "Delete"
    Cmd(5).Caption = "Search"
    Cmd(7).Caption = "Print"
    Cmd(6).Caption = "Exit"
    CmdHelp.Caption = "Help"

    With FG
        .TextMatrix(0, .ColIndex("NewItem")) = "New ItemID"
 
    End With

    lbl(3).Caption = "Total"
    lbl(50).Caption = "Discount"
    lbl(24).Caption = "Net"
    lbl(1).Caption = "By"
    lbl(0).Caption = "Record#"
    lbl(31).Caption = "Item Code"
    lbl(30).Caption = "Item Name"
    lbl(29).Caption = "Status"
    lbl(28).Caption = "Serial "
    lbl(27).Caption = "Qty"
    lbl(26).Caption = "Price"
    lbl(4).Caption = "Inventory"
    lbl(6).Caption = "Vendor"
    lbl(7).Caption = "Date"
    lbl(8).Caption = "Bill#"
    lbl(9).Caption = "Currency"

    lbl(5).Caption = "Discount type"
    lbl(11).Caption = "Discount Value"
    lbl(10).Caption = "Payment Method"
    Command1.Caption = "Convert to Recived VCHR"
    Command2.Caption = "Show Payment VCHR"
    lbl(44).Caption = "Comment"
    XPChkPayType(0).Caption = "Cash"
    lbl(13).Caption = "Value"
    lbl(12).Caption = "ID"
    lbl(2).Caption = "Box"
    lbl(20).Caption = "Currency"
    XPChkPayType(1).Caption = "Credit"
    lbl(15).Caption = "Value"
    lbl(14).Caption = "ID"
    Label1.Caption = "Due Date"
    ChkInstall.Caption = "Installment"
    CmdINSTALLMENT.Caption = "Calc"
    Label2.Caption = "Bank"
    CmdCheque.Caption = "Register"

    With FgInstallments
        .TextMatrix(0, .ColIndex("QestID")) = "ID"
        .TextMatrix(0, .ColIndex("value")) = "value"
        .TextMatrix(0, .ColIndex("Due_Date")) = "Due_Date"
 
    End With

    With FgCheques
 
        .TextMatrix(0, .ColIndex("CheckValue")) = "Value"
        .TextMatrix(0, .ColIndex("CheckNumber")) = "Cheque Number"
        .TextMatrix(0, .ColIndex("BankName")) = "Bank Name"
        .TextMatrix(0, .ColIndex("DueDate")) = "Due Date"
        .TextMatrix(0, .ColIndex("ReleaseDate")) = "Release Date"
 
    End With

    With Me.FG
 
        .TextMatrix(0, .ColIndex("order_no")) = "P/O NO."
    End With

    lbl(53).Caption = "Vendor Bill"

    With Me.Grid
 
        .TextMatrix(0, .ColIndex("Select")) = "Select"
        .TextMatrix(0, .ColIndex("NoteID")) = "NoteID"
        .TextMatrix(0, .ColIndex("NoteSerial")) = "NoteID"

        .TextMatrix(0, .ColIndex("Note_Value")) = "Note_Value"
        .TextMatrix(0, .ColIndex("name")) = "name"

        .TextMatrix(0, .ColIndex("ItemCode")) = "Item Code"
        .TextMatrix(0, .ColIndex("ItemName")) = "Item Name"
    End With

    With Me.GRID2
 
        .TextMatrix(0, .ColIndex("Select")) = "Select"
        .TextMatrix(0, .ColIndex("order_no")) = "order No"
        .TextMatrix(0, .ColIndex("Transaction_Date")) = "Order Date"
        .TextMatrix(0, .ColIndex("CusName")) = "Customer Name"

    End With

    'With Me.grid
 
    '.TextMatrix(0, .ColIndex("Select")) = "Select"
    '.TextMatrix(0, .ColIndex("NoteID")) = "NoteID"
    '.TextMatrix(0, .ColIndex("NoteSerial")) = "NoteID"
    '
    '.TextMatrix(0, .ColIndex("Note_Value")) = "Note_Value"
    '.TextMatrix(0, .ColIndex("name")) = "Based ON"
    '
 
    'End With

    With Me.grid4
        '

        .TextMatrix(0, .ColIndex("Select")) = "Select"
        .TextMatrix(0, .ColIndex("NoteSerial1")) = "NoteID"
        .TextMatrix(0, .ColIndex("name")) = "Account Name"

        .TextMatrix(0, .ColIndex("Note_Value")) = "Note_Value"
        .TextMatrix(0, .ColIndex("ItemCode")) = "Item Code"
        .TextMatrix(0, .ColIndex("ItemName")) = "Item Name"

    End With
 
    Cmd(9).Caption = "Delete Row"
    Label18.Caption = "Total"
 
    With Me.Fg_Journal
        .TextMatrix(0, .ColIndex("LineNo")) = "I"
        .TextMatrix(0, .ColIndex("AccountName")) = "Expenses Name"
        .TextMatrix(0, .ColIndex("value")) = "value"

        .TextMatrix(0, .ColIndex("des")) = "des"
    End With

    lbl(61).Caption = "Bill type"

    BillBasedOn(0).Caption = "Direct Purchase Invoice"
    BillBasedOn(1).Caption = "From Recieve Vouchers"
    BillBasedOn(2).Caption = "Purchase Orders"

    With Me.Fg_Journal
        .TextMatrix(0, .ColIndex("LineNo")) = "I"
        .TextMatrix(0, .ColIndex("AccountName")) = "Expenses Name"
        .TextMatrix(0, .ColIndex("value")) = "value"

        .TextMatrix(0, .ColIndex("des")) = "des"
    End With

    With Me.GRID1
        .TextMatrix(0, .ColIndex("Select")) = "Select"
        .TextMatrix(0, .ColIndex("noteserial1")) = "Voucher NO"
        .TextMatrix(0, .ColIndex("Transaction_Date")) = "Voucher Date"
        .TextMatrix(0, .ColIndex("NoteSerial")) = "JE Voucher NO"
    End With

    With Me.GRID1
        .TextMatrix(0, .ColIndex("Select")) = "Select"
        .TextMatrix(0, .ColIndex("noteserial1")) = "Voucher NO"
        .TextMatrix(0, .ColIndex("Transaction_Date")) = "Voucher Date"
        .TextMatrix(0, .ColIndex("NoteSerial")) = "JE Voucher NO"
    End With

    With Me.VSFlexGrid1
        .TextMatrix(0, .ColIndex("Select")) = "Select"
        .TextMatrix(0, .ColIndex("order_no")) = "Order No"
        .TextMatrix(0, .ColIndex("Transaction_Date")) = "Voucher Date"
        .TextMatrix(0, .ColIndex("CusName")) = "Customer Name"
    End With

    Frame3.Caption = "JE Voucher NO"
    lbl(62).Caption = "JE Voucher NO"
    Cmd(10).Caption = "Print JE"
 
    lbl(59).Caption = "Total Financial Invoice"
    Command5.Caption = "Save"
    XPChkPayType(2).Caption = "Cheques"

End Sub

Function CuurentLogdata(Optional Currentmode As String)
    LogTextA = "    ‘«‘… " & ScreenNameArabic & Chr(13) & " —ﬁ„ «·›« Ê—…   " & TxtNoteSerial1.text & Chr(13) & " —ﬁ„ ›« Ê—… «·„Ê—œ   " & txtManualNO.text & Chr(13) & " «· «—ÌŒ " & XPDtbBill.value & Chr(13) & " «·Œ“Ì‰… " & DcboBox.text & Chr(13) & " «·„Œ“‰  " & DCboStoreName.text & Chr(13) & "  «·⁄„Ì· / «·„Ê—œ   " & DBCboClientName.text & Chr(13) & "‰Ê⁄ «·”‰œ " & DCDocTypes & Chr(13) & "»‰«¡ ⁄·Ï " & CBoBasedON & "»—ﬁ„   " & Txt_order_no & Chr(13) & "ÿ—Ìﬁ… «·œ›⁄ " & CboPayMentType & Chr(13) & "‰Ê⁄ «·Œ’„ " & XPCboDiscountType & Chr(13) & "ﬁÌ„… «·Œ’„ " & XPTxtDiscountVal & Chr(13) & "  Ê’Ê· «·‘Õ‰… " & DTArrivalDate & Chr(13) & "  «·«” Õﬁ«ﬁ " & DtpDelayDate & Chr(13) & " «·⁄„·Â " & Dccurrency & Chr(13) & "—ﬁ„ «·ﬁÌœ " & TxtNoteSerial
                     
    LogTextE = "    Screen  " & ScreenNameEnglish & Chr(13) & " Bill No " & TxtNoteSerial1.text & Chr(13) & "Supplier Bill No " & txtManualNO.text & Chr(13) & " Date " & XPDtbBill.value & Chr(13) & " Box " & DcboBox.text & Chr(13) & " Store  " & DCboStoreName.text & Chr(13) & " Supplier/Cuxtomer" & DBCboClientName.text & Chr(13) & "Doc Type" & DCDocTypes & Chr(13) & "Based On" & CBoBasedON & "No :   " & Txt_order_no & Chr(13) & "Payment Type" & CboPayMentType & Chr(13) & "Discount Type  " & XPCboDiscountType & Chr(13) & " Discount Vaalue   " & XPTxtDiscountVal & Chr(13) & " Shipment Arival Date" & DTArrivalDate & Chr(13) & "Due Date " & DtpDelayDate & Chr(13) & " Currency " & Dccurrency & Chr(13) & " GE NO" & TxtNoteSerial
                     
    If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), 150, Date, Time, LogTextA, LogTextE, Me.name, Me.TxtModFlg, "", , TxtNoteSerial, TxtNoteSerial1
    Else
        AddToLogFile CInt(user_id), 150, Date, Time, LogTextA, LogTextE, Me.name, "D", "", , TxtNoteSerial, TxtNoteSerial1
    End If
    
End Function

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap
    Dim i As Integer
    RegisterLogInOut Me.name, ScreenNameArabic, ScreenNameEnglish, 150

    If rs.State = adStateOpen Then
        If Not (rs.EOF Or rs.BOF) Then
            If rs.EditMode <> adEditNone Then
                rs.CancelUpdate
            End If
        End If

        rs.Close
    End If

    For i = LBound(cSearchDcbo) To UBound(cSearchDcbo)
        Set cSearchDcbo(i) = Nothing
    Next i

    Set rs = Nothing
    Set TTP = Nothing
    NewGrid.Class_Terminate
    Set NewGrid = Nothing
    Set BuyReport = Nothing
    Set m_MnuShowNewItemsPrices = Nothing

    If Not m_FrmSearch Is Nothing Then
        Unload m_FrmSearch
        Set m_FrmSearch = Nothing
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap

    Select Case Me.TxtModFlg.text

        Case "R"

            If SystemOptions.UserInterface = EnglishInterface Then
                Me.Caption = "Purcahase Invoice"
        
            Else
                Me.Caption = "›« Ê—… ‘—«¡"
            End If
    
            Me.Cmd(2).Enabled = False
            Me.Cmd(3).Enabled = False
     
            Me.Cmd(0).Enabled = True
            Me.Cmd(1).Enabled = True
            Me.Cmd(4).Enabled = True
            Me.Cmd(5).Enabled = True
            Me.Cmd(7).Enabled = True
        
            Me.XPBtnMove(0).Enabled = True
            Me.XPBtnMove(1).Enabled = True
            Me.XPBtnMove(2).Enabled = True
            Me.XPBtnMove(3).Enabled = True
            XPBtnNewClients.Enabled = False
        
            XPCboDiscountType.locked = True
            Me.XPDtbBill.Enabled = False
            Me.DBCboClientName.locked = True
            Me.DCboStoreName.locked = True
            Me.XPTxtDiscountVal.locked = True
        
            XPChkPayType(0).Enabled = False
            XPChkPayType(1).Enabled = False
            XPChkPayType(2).Enabled = False
        
            XPTxtValue(0).Enabled = False
            XPTxtSerial(0).Enabled = False
            XPTxtValue(1).Enabled = False
            XPTxtSerial(1).Enabled = False
            FG.Editable = flexEDNone

            If rs.RecordCount < 1 Then
                Me.XPBtnMove(0).Enabled = False
                Me.XPBtnMove(1).Enabled = False
                Me.XPBtnMove(2).Enabled = False
                Me.XPBtnMove(3).Enabled = False
                Me.Cmd(1).Enabled = False
                Me.Cmd(4).Enabled = False
                Me.Cmd(5).Enabled = False
                Me.Cmd(7).Enabled = False
            
            End If
        
            CboPayMentType.locked = True
            DtpDelayDate.Enabled = False
            Ele(4).Enabled = False
        
            XPChkTAX.Enabled = False
            ChkTaxAdd.Enabled = False
            ChkTaxSerivce.Enabled = False
            ChkTaxStamp.Enabled = False
        
        Case "N"
            '   Me.Caption = "›« Ê—… ‘—«¡( ÃœÌœ )"
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
        
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(5).Enabled = False
            Me.Cmd(7).Enabled = False
        
            '        Me.XPBtnMove(0).Enabled = False
            '        Me.XPBtnMove(1).Enabled = False
            '        Me.XPBtnMove(2).Enabled = False
            '        Me.XPBtnMove(3).Enabled = False
            XPBtnNewClients.Enabled = True
            FG.Enabled = True
            FG.Rows = 2
            XPCboDiscountType.locked = False
            Me.XPDtbBill.Enabled = True
            Me.DBCboClientName.locked = False
            Me.DCboStoreName.locked = False
            Me.XPTxtDiscountVal.locked = False
            XPChkPayType(0).Enabled = True
            XPChkPayType(1).Enabled = True
            XPChkPayType(2).Enabled = True
            XPChkPayType(0).value = Unchecked
            XPChkPayType(1).value = Unchecked
            XPChkPayType(2).value = Unchecked
            FG.Editable = flexEDKbdMouse
            XPDtbBill.value = Date
            '        XPFillData.Enabled = True
            XPCboDiscountType.ListIndex = 0
            CboPayMentType.ListIndex = 0
            CboPayMentType.locked = False
            DtpDelayDate.Enabled = True
            DtpDelayDate.value = Date
            Ele(4).Enabled = True
        
            CboItemCase.ListIndex = 0
        
            XPChkTAX.Enabled = True
            ChkTaxAdd.Enabled = True
            ChkTaxSerivce.Enabled = True
            ChkTaxStamp.Enabled = True

        Case "E"
            '   Me.Caption = "›« Ê—… ‘—«¡(  ⁄œÌ· )"
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
        
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(5).Enabled = False
            Me.Cmd(7).Enabled = False
        
            Me.XPBtnMove(0).Enabled = False
            Me.XPBtnMove(1).Enabled = False
            Me.XPBtnMove(2).Enabled = False
            Me.XPBtnMove(3).Enabled = False
            XPBtnNewClients.Enabled = True
        
            FG.Enabled = True
            XPCboDiscountType.locked = False
            Me.XPDtbBill.Enabled = True
            Me.DBCboClientName.locked = False
            Me.DCboStoreName.locked = False
            Me.XPTxtDiscountVal.locked = False
        
            XPChkPayType(0).Enabled = True
            XPChkPayType(1).Enabled = True
            XPChkPayType(2).Enabled = True
            DtpDelayDate.Enabled = True
        
            If XPChkPayType(0).value = Checked Then
                XPChkPayType_Click (0)
            End If

            If XPChkPayType(1).value = Checked Then
                XPChkPayType_Click (1)
            End If

            If XPChkPayType(2).value = Checked Then
                XPChkPayType_Click (2)
            End If

            If CboPayMentType.ListIndex = 0 Then
                CboPayMentType_Change
            End If

            FG.Editable = flexEDKbdMouse
        
            CboPayMentType.locked = False
            DBCboClientName_Change
            Ele(4).Enabled = True
            XPChkTAX.Enabled = True
            ChkTaxAdd.Enabled = True
            ChkTaxSerivce.Enabled = True
            ChkTaxStamp.Enabled = True
    End Select

    Exit Sub
ErrTrap:
    Stop
End Sub

Public Sub Retrive(Optional Lngid As Long = 0)
    Dim RsDetails As ADODB.Recordset
    Dim strSQL As String
    Dim RsNotes As ADODB.Recordset
    Dim RsTemp As ADODB.Recordset
    Dim RsTest As ADODB.Recordset
    Dim Num As Long
    Dim Msg As String
    Dim i As Integer
    Dim LngPartID As Long
    Dim RsPartDetails As ADODB.Recordset

    ' On Error GoTo ErrTrap
    '---------------------------------------------
    'Here We Reset all Setting
    Me.CmdNotes.Visible = False
    Me.CmdNotes.Tag = ""
    Me.CmdRetruns.Visible = False
    Me.CmdRetruns.Tag = ""
    ChkTaxAdd.value = vbUnchecked
    Me.TxtTaxAddValue.text = ""
    ChkTaxStamp.value = vbUnchecked
    Me.TxtTaxStampValue.text = ""
    ChkTaxStamp.value = vbUnchecked
    Me.TxtTaxStampValue.text = ""
    ChkTaxSerivce.value = vbUnchecked
    Me.TxtTaxServiceValue.text = ""

    '---------------------------------------------
    '---------------------------------------------
    If rs.RecordCount < 1 Then
        XPTxtCurrent.Caption = 0
        XPTxtCount.Caption = 0
        Exit Sub
    End If

    If rs.EOF Or rs.BOF Then
        Exit Sub
    End If

    If Lngid <> 0 Then
        rs.find "Transaction_ID=" & Lngid, , adSearchForward, adBookmarkFirst

        If rs.BOF Or rs.EOF Then
            Exit Sub
        End If
    End If

    TxtFillData.text = "T"
    Screen.MousePointer = vbArrowHourglass
    dcBranch.BoundText = IIf(IsNull(rs("BranchId").value), "", rs("BranchId").value)
    Me.DCproject.BoundText = IIf(IsNull(rs("project_id").value), "", rs("project_id").value)
    Me.TxtNoteSerial.text = IIf(IsNull(rs("NoteSerial").value), "", (rs("NoteSerial").value))
    Me.TxtNoteSerial1.text = IIf(IsNull(rs("NoteSerial1").value), "", (rs("NoteSerial1").value))
    Me.oldtxtNoteSerial1.text = IIf(IsNull(rs("OldNoteSerial1").value), IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value), rs("OldNoteSerial1").value)

    lbl(67).Caption = showLabel(TxtNoteSerial1, oldtxtNoteSerial1)

    Me.TXTNoteID.text = IIf(IsNull(rs("NoteID").value), "", (rs("NoteID").value))

    Txt_order_no.text = IIf(IsNull(rs("order_no").value), "", (rs("order_no").value))
    txtManualNO.text = IIf(IsNull(rs("ManualNO").value), "", (rs("ManualNO").value))

    TxtManualNo1.text = IIf(IsNull(rs("ManualNo1").value), "", (rs("ManualNo1").value))

    txt_Currency_rate.text = IIf(IsNull(rs("Currency_rate").value), 1, (rs("Currency_rate").value))

    XPTxtBillID.text = IIf(IsNull(rs("Transaction_ID").value), "", (rs("Transaction_ID").value))
    TxtTransSerial.text = IIf(IsNull(rs("Transaction_Serial").value), "", (rs("Transaction_Serial").value))
    XPDtbBill.value = IIf(IsNull(rs("Transaction_Date").value), "", (rs("Transaction_Date").value))
    DTArrivalDate.value = IIf(IsNull(rs("ArrivalDate").value), Date, (rs("ArrivalDate").value))

    XPCboDiscountType.ListIndex = IIf(IsNull(rs("Trans_DiscountType").value), 0, rs("Trans_DiscountType").value)

    XPTxtDiscountVal.text = IIf(IsNull(rs("Trans_Discount").value), "", Trim(rs("Trans_Discount").value))

    TXTToTAlELSHahn.text = IIf(Not IsNumeric(rs("ToTAlELSHahn").value), 0, rs("ToTAlELSHahn").value)
    Txt_EXport.text = IIf(Not IsNumeric(rs("total_expenses").value), 0, rs("total_expenses").value)
    txt_total_bill.text = IIf(Not IsNumeric(rs("total_payments").value), 0, rs("total_payments").value)
    DCDocTypes.BoundText = IIf(IsNull(rs("Doctype").value), "", rs("Doctype").value)

    TxtBillComment.text = IIf(IsNull(rs("TransactionComment").value), "", (rs("TransactionComment").value))

    CboPayMentType.ListIndex = IIf(IsNull(rs("PaymentType").value), 0, rs("PaymentType").value)
    Me.DBCboClientName.BoundText = IIf(IsNull(rs("CusID").value), "", rs("CusID").value)
    Me.DCboUserName.BoundText = IIf(IsNull(rs("UserID").value), "", rs("UserID").value)
    Me.DCboStoreName.BoundText = IIf(IsNull(rs("StoreID").value), "", rs("StoreID").value)
    Me.DcboBox.BoundText = IIf(IsNull(rs("BoxID").value), "", rs("BoxID").value)

    Text1.text = IIf(IsNull(rs("nots").value), "", (rs("nots").value))
    'Text1.text = IIf(IsNull(Rs("nots").Value), "", (Rs("nots").Value))

    'txt_Shipment_no.text = IIf(IsNull(Rs("Shipment_no").value), "", Trim(Rs("Shipment_no").value))
    'Txt_order_no.text = IIf(IsNull(Rs("order_no").value), "", Trim(Rs("order_no").value))
    Me.Dccurrency.BoundText = IIf(IsNull(rs("Currency_id").value), "", rs("Currency_id").value)

    TxtLcNo.text = IIf(IsNull(rs("LcNo").value), "", (rs("LcNo").value))

    '÷—»Ì… «·„»Ì⁄« 
    XPTxtTaxValue.text = IIf(IsNull(rs("TaxValue").value), "", (rs("TaxValue").value))
    XPChkTAX.value = IIf(rs("TaxFound") = True, Checked, Unchecked)

    '÷—»Ì… «·Œ’„ Ê«·≈÷«›…
    If Not IsNull(rs("TaxAddValue").value) Then
        If rs("TaxAddValue").value > 0 Then
            ChkTaxAdd.value = vbChecked
            Me.TxtTaxAddValue.text = rs("TaxAddValue").value
        End If
    End If

    '÷—»Ì… «·œ„€…
    If Not IsNull(rs("TaxStampValue").value) Then
        If rs("TaxStampValue").value > 0 Then
            ChkTaxStamp.value = vbChecked
            Me.TxtTaxStampValue.text = rs("TaxStampValue").value
        End If
    End If

    '÷—»Ì… «·Œœ„…
    If Not IsNull(rs("TaxServiceValue").value) Then
        If rs("TaxServiceValue").value > 0 Then
            ChkTaxSerivce.value = vbChecked
            Me.TxtTaxServiceValue.text = rs("TaxServiceValue").value
        End If
    End If

    FG.Rows = 2
    FG.Clear flexClearScrollable, flexClearEverything
    XPTxtSum.text = ""
    CBoBasedON.ListIndex = IIf(IsNull(rs("CBoBasedON").value), 0, (rs("CBoBasedON").value))

    If Not IsNull(rs("BillBasedOn").value) Then

        If rs("BillBasedOn").value = 0 Then
            BillBasedOn(0).value = True
            BillBasedOn_Click (0)
        ElseIf rs("BillBasedOn").value = 1 Then
            BillBasedOn(1).value = True
            BillBasedOn_Click (1)
        ElseIf rs("BillBasedOn").value = 2 Then
            BillBasedOn(2).value = True
            BillBasedOn_Click (2)
        End If
    
    Else

        BillBasedOn(0).value = True
        BillBasedOn_Click (0)
    End If

    strSQL = "SELECT TblItems.HaveSerial, * FROM TblItems INNER JOIN Transaction_Details " & "ON TblItems.ItemID = Transaction_Details.Item_ID INNER JOIN dbo.TblUnites ON dbo.Transaction_Details.UnitID = dbo.TblUnites.UnitID"
    strSQL = strSQL + "  where Transaction_ID=" & val(rs("Transaction_ID").value)
    Set RsDetails = New ADODB.Recordset
    RsDetails.Open strSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Not (RsDetails.EOF Or RsDetails.BOF) Then
        FG.Rows = RsDetails.RecordCount + 1

        For Num = 1 To RsDetails.RecordCount
     
            FG.TextMatrix(Num, FG.ColIndex("LineShahn")) = IIf(IsNull(RsDetails("LineShahn")), 0, (RsDetails("LineShahn").value))
            FG.TextMatrix(Num, FG.ColIndex("Code")) = IIf(IsNull(RsDetails("Item_ID")), "", (RsDetails("Item_ID").value))
            FG.TextMatrix(Num, FG.ColIndex("Name")) = IIf(IsNull(RsDetails("Item_ID")), "", Trim(RsDetails("Item_ID").value))
            'FG.TextMatrix(Num, FG.ColIndex("Count")) = IIf(IsNull(RsDetails("Quantity")), "", (RsDetails("Quantity").Value))
            FG.TextMatrix(Num, FG.ColIndex("Count")) = IIf(IsNull(RsDetails("ShowQty")), "", (RsDetails("ShowQty").value))

            FG.TextMatrix(Num, FG.ColIndex("Serial")) = IIf(IsNull(RsDetails("ItemSerial")), "", Trim(RsDetails("ItemSerial").value))

            If RsDetails("HaveSerial") = True Then
                FG.TextMatrix(Num, FG.ColIndex("HaveSerial")) = True
            End If

            ' FG.TextMatrix(Num, FG.ColIndex("Price")) = IIf(IsNull(RsDetails("Price")), "", (RsDetails("Price").Value))
            FG.TextMatrix(Num, FG.ColIndex("Price")) = IIf(IsNull(RsDetails("showPrice")), "", (RsDetails("showPrice").value))

            If SystemOptions.SysDataBaseType = SQLServerDataBase Then
                FG.TextMatrix(Num, FG.ColIndex("ItemCase")) = IIf(IsNull(RsDetails("ItemCase")), "", (RsDetails("ItemCase").value))
            Else
                FG.TextMatrix(Num, FG.ColIndex("ItemCase")) = IIf(IsNull(RsDetails("Transaction_Details.ItemCase")), "", (RsDetails("Transaction_Details.ItemCase").value))
            End If

            FG.TextMatrix(Num, FG.ColIndex("DiscountType")) = IIf(IsNull(RsDetails("ItemDiscountType")), "", (RsDetails("ItemDiscountType").value))
            FG.TextMatrix(Num, FG.ColIndex("DiscountVal")) = IIf(IsNull(RsDetails("ItemDiscount")), "", (RsDetails("ItemDiscount").value))
            FG.TextMatrix(Num, FG.ColIndex("guaranteeTime")) = IIf(IsNull(RsDetails("guaranteeTime")), "", (RsDetails("guaranteeTime").value))
            FG.TextMatrix(Num, FG.ColIndex("Remarks")) = IIf(IsNull(RsDetails("Remarks")), "", (RsDetails("Remarks").value))
        
            FG.TextMatrix(Num, FG.ColIndex("ColorID")) = IIf(IsNull(RsDetails("ColorID")), 1, (RsDetails("ColorID").value))
            FG.TextMatrix(Num, FG.ColIndex("ItemSize")) = IIf(IsNull(RsDetails("ItemSize")), 0, (RsDetails("ItemSize").value))
        
            FG.Cell(flexcpData, Num, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitID")), "", (RsDetails("UnitID").value))
            FG.TextMatrix(Num, FG.ColIndex("UnitID")) = IIf(IsNull(RsDetails("UnitName")), "", (RsDetails("UnitName").value))
            FG.TextMatrix(Num, FG.ColIndex("order_no")) = IIf(IsNull(RsDetails("order_no")), "", (RsDetails("order_no").value))
            FG.TextMatrix(Num, FG.ColIndex("OrderArrivalDate")) = IIf(IsNull(RsDetails("OrderArrivalDate")), "", (RsDetails("OrderArrivalDate").value))
            
            FG.TextMatrix(Num, FG.ColIndex("FoxyNo")) = IIf(IsNull(RsDetails("FoxyNo")), "", (RsDetails("FoxyNo").value))
        
            FG.TextMatrix(Num, FG.ColIndex("ColorID")) = IIf(IsNull(RsDetails("ColorID")), 1, (RsDetails("ColorID").value))
            FG.TextMatrix(Num, FG.ColIndex("ItemSize")) = IIf(IsNull(RsDetails("ItemSize")), 1, (RsDetails("ItemSize").value))
            FG.TextMatrix(Num, FG.ColIndex("ClassID")) = IIf(IsNull(RsDetails("ClassID")), 1, (RsDetails("ClassID").value))
            'FG.TextMatrix(Num, FG.ColIndex("OpeningBurcahseQty")) = IIf(IsNull(RsDetails("OpeningBurcahseQty").value), "", RsDetails("OpeningBurcahseQty").value)
            '      FG.TextMatrix(Num, FG.ColIndex("OpeningBurcahseValue")) = IIf(IsNull(RsDetails("OpeningBurcahseValue").value), "", RsDetails("OpeningBurcahseValue").value)
            '       FG.TextMatrix(Num, FG.ColIndex("OpeningSalesQty")) = IIf(IsNull(RsDetails("OpeningSalesQty").value), "", RsDetails("OpeningSalesQty").value)
            '        FG.TextMatrix(Num, FG.ColIndex("OpeningSalesValue")) = IIf(IsNull(RsDetails("OpeningSalesValue").value), "", RsDetails("OpeningSalesValue").value)
            FG.TextMatrix(Num, FG.ColIndex("OrderArrivalDate")) = IIf(IsNull(RsDetails("OrderArrivalDate").value), "", RsDetails("OrderArrivalDate").value)
            FG.TextMatrix(Num, FG.ColIndex("order_no")) = IIf(IsNull(RsDetails("order_no").value), "", RsDetails("order_no").value)
            FG.TextMatrix(Num, FG.ColIndex("FoxyNo")) = IIf(IsNull(RsDetails("FoxyNo")), "", RsDetails("FoxyNo").value)

            FG.TextMatrix(Num, FG.ColIndex("ProductionDate")) = IIf(IsNull(RsDetails("ProductionDate")), "", (RsDetails("ProductionDate").value))
            FG.TextMatrix(Num, FG.ColIndex("ExpiryDate")) = IIf(IsNull(RsDetails("ExpiryDate")), "", (RsDetails("ExpiryDate").value))
            FG.TextMatrix(Num, FG.ColIndex("LotNO")) = IIf(IsNull(RsDetails("LotNO")), "", (RsDetails("LotNO").value))
        '***********************************************************
        FG.TextMatrix(Num, FG.ColIndex("SBillNO")) = IIf(IsNull(RsDetails("SBillNO").value), "", RsDetails("SBillNO").value)
            FG.TextMatrix(Num, FG.ColIndex("ExtraType")) = IIf(IsNull(RsDetails("ExtraType")), "", (RsDetails("ExtraType").value))
            FG.TextMatrix(Num, FG.ColIndex("ExtraVal")) = IIf(IsNull(RsDetails("ExtraVal")), "", (RsDetails("ExtraVal").value))
        
        'FG.Cell(flexcpData, Num, FG.ColIndex("SupplierID")) = IIf(IsNull(RsDetails("SupplierID")), "", (RsDetails("SupplierID").value))
            FG.TextMatrix(Num, FG.ColIndex("SupplierID")) = IIf(IsNull(RsDetails("SupplierID")), "", (RsDetails("SupplierID").value))
            
        
         '***********************************************************
            RsDetails.MoveNext

            If FG.Rows > 10 Then
                If Num = 8 Then FG.Refresh
            End If

        Next Num

        '  FG.AutoSize 0, FG.Cols - 1, False
    End If

    XPChkPayType(0).value = Unchecked
    XPChkPayType(1).value = Unchecked
    XPChkPayType(2).value = Unchecked
    XPTxtValue(0).text = ""
    XPTxtValue(1).text = ""

    XPTxtSerial(0).text = ""
    XPTxtSerial(1).text = ""
    DtpDelayDate.value = Date
    strSQL = "select * From Notes where Transaction_ID=" & val(rs("Transaction_ID").value)
    Set RsNotes = New ADODB.Recordset
    RsNotes.Open strSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Not (RsNotes.EOF Or RsNotes.BOF) Then

        For Num = 1 To RsNotes.RecordCount

            If RsNotes("NoteType").value = 0 Then
                XPChkPayType(0).value = Checked
                XPChkPayType_Click (0)
                'Me.TxtNoteID(0).text = IIf(IsNull(RsNotes("NOTEID").Value), "", (RsNotes("NOTEID").Value))
                XPTxtValue(0).text = IIf(IsNull(RsNotes("Note_Value").value), "", (RsNotes("Note_Value").value))
                XPTxtSerial(0).text = IIf(IsNull(RsNotes("NoteSerial").value), "", Trim(RsNotes("NoteSerial").value))
                Me.DcboBox.BoundText = IIf(IsNull(RsNotes("BoxID").value), "", RsNotes("BoxID").value)
            End If

            If RsNotes("NoteType").value = 1 Then
                XPChkPayType(1).value = Checked
                XPChkPayType_Click (1)
                'Me.TxtNoteID(1).text = IIf(IsNull(RsNotes("NOTEID").Value), "", (RsNotes("NOTEID").Value))
                XPTxtValue(1).text = IIf(IsNull(RsNotes("Note_Value").value), "", (RsNotes("Note_Value").value))
                XPTxtValue(1).Tag = IIf(IsNull(RsNotes("NoteID").value), "", (RsNotes("NoteID").value))
                XPTxtSerial(1).text = IIf(IsNull(RsNotes("NoteSerial").value), "", Trim(RsNotes("NoteSerial").value))
                DtpDelayDate.value = IIf(IsNull(RsNotes("DueDate").value), "", (RsNotes("DueDate").value))
            End If

            If RsNotes("NoteType").value = 13 Then
                XPChkPayType(2).value = Checked
                XPChkPayType_Click (2)
            End If
        
            RsNotes.MoveNext
        Next Num

    End If

    Set RsNotes = New ADODB.Recordset
    strSQL = "SELECT Notes.NoteID, Notes.NoteDate, Notes.NoteType, Notes.NoteSerial," & "Notes.Note_Value, Notes.BankID,BanksData.BankName , Notes.ChqueNum, Notes.DueDate "
    strSQL = strSQL + " FROM Notes INNER JOIN BanksData ON Notes.BankID = BanksData.BankID "
    strSQL = strSQL + " Where NoteType=13 AND NOTES.Transaction_ID=" & val(rs("Transaction_ID").value)
    strSQL = strSQL + " Order BY Notes.NoteID"
    RsNotes.Open strSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    With Me.FgCheques
        .Rows = .FixedRows

        If Not (RsNotes.BOF Or RsNotes.EOF) Then
            .Rows = .FixedRows + RsNotes.RecordCount

            For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .ColIndex("CheckValue")) = IIf(IsNull(RsNotes("Note_Value").value), "", RsNotes("Note_Value").value)
                .TextMatrix(i, .ColIndex("CheckNumber")) = IIf(IsNull(RsNotes("ChqueNum").value), "", RsNotes("ChqueNum").value)
                .TextMatrix(i, .ColIndex("BankID")) = IIf(IsNull(RsNotes("BankID").value), "", RsNotes("BankID").value)
                .TextMatrix(i, .ColIndex("BankName")) = IIf(IsNull(RsNotes("BankName").value), "", RsNotes("BankName").value)

                If Not IsNull(RsNotes("DueDate").value) Then
                    .TextMatrix(i, .ColIndex("DueDate")) = DisplayDate(RsNotes("DueDate").value)
                Else
                    .TextMatrix(i, .ColIndex("DueDate")) = ""
                End If

                RsNotes.MoveNext
            Next i

        End If

        .AutoSize 0, .Cols - 1, False
        SumChecks
    End With

    '⁄—÷ «·√ﬁ”«ÿ «·„”Ã·… ⁄·Ï «·›« Ê—…
    If XPTxtValue(1).Tag <> "" Then
        strSQL = "Select * From InstallMent where NoteID=" & XPTxtValue(1).Tag
        Set RsTest = New ADODB.Recordset
        RsTest.Open strSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (RsTest.EOF Or RsTest.BOF) Then
            CmdINSTALLMENT.Enabled = True

            If SystemOptions.UserInterface = ArabicInterface Then
                CmdINSTALLMENT.Caption = "⁄—÷ «·√ﬁ”«ÿ «·„”Ã·…"
            Else
                CmdINSTALLMENT.Caption = "View"
            End If

            LngPartID = RsTest("PartID").value
            Me.LblPrecenType.Tag = RsTest("InterestType").value

            If RsTest("InterestType").value = 0 Then
                LblPrecenType.Caption = "‰”»… „∆ÊÌ…"
            ElseIf RsTest("InterestType").value = 1 Then
                LblPrecenType.Caption = "ﬁÌ„… À«» …"
            ElseIf RsTest("InterestType").value = 2 Then
                LblPrecenType.Caption = "·«ÌÊÃœ"
            End If

            Me.LblPrecenValue.Caption = RsTest("InterestVal").value
            Me.LblInstallTotal.Caption = RsTest("Total").value
            Me.LblInstallCount.Caption = RsTest("InstallCount").value
            Me.LblFirstInstallDate.Caption = DisplayDate(RsTest("FirstInstallDate").value)
            Me.LblInstallmentType.Tag = RsTest("InstallmentType").value

            If RsTest("InstallmentType").value = 0 Then
                LblInstallmentType.Caption = "ÌÊ„"
            ElseIf RsTest("InstallmentType").value = 1 Then
                LblInstallmentType.Caption = "‘Â—"
            ElseIf RsTest("InstallmentType").value = 2 Then
                LblInstallmentType.Caption = "”‰…"
            End If

            Me.LblInstallSeprator.Caption = RsTest("InstallSeprator").value
            Me.LblStartValue.Caption = IIf(IsNull(RsTest("StartValue").value), "", RsTest("StartValue").value)
            Set RsPartDetails = New ADODB.Recordset
            strSQL = "Select * From InstallMentDetails Where PartID=" & LngPartID
            RsPartDetails.Open strSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

            If Not (RsPartDetails.BOF Or RsPartDetails.EOF) Then
                RsPartDetails.MoveFirst

                With Me.FgInstallments
                    .Rows = .FixedRows + RsPartDetails.RecordCount

                    For i = .FixedRows To .Rows - 1
                        .TextMatrix(i, .ColIndex("QestID")) = IIf(IsNull(RsPartDetails("QestID").value), "", RsPartDetails("QestID").value)
                        .TextMatrix(i, .ColIndex("Value")) = IIf(IsNull(RsPartDetails("Value").value), "", RsPartDetails("Value").value)

                        If Not IsNull(RsPartDetails("DueDate").value) Then
                            .TextMatrix(i, .ColIndex("Due_Date")) = DisplayDate(RsPartDetails("DueDate").value)
                        Else
                            .TextMatrix(i, .ColIndex("Due_Date")) = ""
                        End If
 
                        RsPartDetails.MoveNext
                    Next i

                End With

            End If

        Else
            CmdINSTALLMENT.Enabled = False

            If SystemOptions.UserInterface = ArabicInterface Then
                CmdINSTALLMENT.Caption = " ﬁ”Ìÿ «·ﬁÌ„… «·¬Ã·…"
            Else
                CmdINSTALLMENT.Caption = "calc"
            End If
        End If
    End If

    NewGrid.Calculate 1, , , True
    Dim SngRelatedNotesValues As Single
    Me.CmdNotes.Visible = ShowRelatedNotes(val(Me.XPTxtBillID.text), 0, SngRelatedNotesValues)
    Me.CmdNotes.Tag = SngRelatedNotesValues

    SngRelatedNotesValues = 0
    Me.CmdRetruns.Visible = ShowRelatedTransactions(val(Me.XPTxtBillID.text), 0, SngRelatedNotesValues)
    Me.CmdRetruns.Tag = SngRelatedNotesValues
    '-----------------------------------------------------------------------------------------------
    Screen.MousePointer = vbDefault
    TxtFillData.text = "F"
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    fill_bill_items_table
    Command3_Click
    '«” —Ã«⁄ «·„’—Ê›«  «· ﬁœÌ—ÌÂ
    fillExpensesFactoryGrid
    '  FillVoucherGrid

    Exit Sub
ErrTrap:
    Msg = "Œÿ« ›Ï ≈” —Ã«⁄ «·»Ì«‰« ..!!!"
    Msg = Msg & Chr(13) & Err.description
    Msg = Msg & Chr(13) & Err.Number
    Msg = Msg & Chr(13) & Err.Source
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Screen.MousePointer = vbDefault

End Sub

Private Sub Undo()
    Dim Msg As String

    On Error GoTo ErrTrap

    Select Case TxtModFlg.text

        Case "N"

            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "”Ê› Ì „ «· —«Ã⁄ ›Ï  ”ÃÌ· Â–Â «·›« Ê—… .."
                Msg = Msg & Chr(13) & "›Â· «‰  „ «ﬂœ „‰ «·√” „—«— ..!!"
            Else
                Msg = "Confirm Undo"
            End If

            If MsgBox(Msg, vbQuestion + vbYesNo + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                clear_all Me
                Me.TxtModFlg.text = "R"
                XPBtnMove_Click (1)
            End If

        Case "E"

            If SystemOptions.UserInterface = ArabicInterface Then
        
                Msg = "”Ê› Ì „ «· —«Ã⁄ ›Ï  ⁄œÌ· Â–Â «·›« Ê—… .."
                Msg = Msg & Chr(13) & "›Â· «‰  „ «ﬂœ „‰ «·√” „—«— ..!!"
            Else
                Msg = "Confirm Undo"
            End If
  
            If MsgBox(Msg, vbQuestion + vbYesNo + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                rs.find "Transaction_ID='" & val(XPTxtBillID.text) & "'", , adSearchForward, adBookmarkFirst

                If rs.EOF Or rs.BOF Then
                    Me.TxtModFlg.text = "R"
                    Exit Sub
                End If

                If Not rs.EOF Or rs.BOF Then
                    Me.TxtModFlg.text = "R"
                    Retrive
                End If
            End If

    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub Del_TransAction()
    Dim Msg As String
    Dim strSQL As String
    Dim BegainTrans As Boolean
    Dim order_no As String
    order_no = Me.Txt_order_no
    On Error GoTo ErrTrap

    If XPTxtBillID.text <> "" Then
        Msg = "”Ì „ Õ–› »Ì«‰«  «·›« Ê—…  —ﬁ„ " & Chr(13)
        Msg = Msg + TxtNoteSerial1.text & Chr(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            If AvailableDeal = True Or AvailableDeal = False Then
                If Not rs.RecordCount < 1 Then
                    Cn.BeginTrans
                    BegainTrans = True
                
                    strSQL = "Delete From DOUBLE_ENTREY_VOUCHERS  " & "Where DOUBLE_ENTREY_VOUCHERS.Transaction_ID=" & rs("Transaction_ID").value
                    Cn.Execute strSQL, , adExecuteNoRecords
                
                    strSQL = "Delete From DOUBLE_ENTREY_VOUCHERS  " & "Where DOUBLE_ENTREY_VOUCHERS.Transaction_ID=" & get_transaction_id(rs("nots").value, 20)
                    Cn.Execute strSQL, , adExecuteNoRecords
                
                    strSQL = "Delete From Transactions  " & "Where Transaction_ID=" & get_transaction_id(rs("nots").value, 20)
                    Cn.Execute strSQL, , adExecuteNoRecords
                
                    strSQL = "update Notes set  Transaction_ID1=Null , ItemID=NUll, buy = null Where   (Transaction_ID1=" & val(Me.XPTxtBillID.text) & ")"
                    Cn.Execute strSQL
            
                    strSQL = "update DOUBLE_ENTREY_VOUCHERS set Transaction_ID1=Null ,  ItemID=NUll, buy = null Where  ( Transaction_ID1=" & val(Me.XPTxtBillID.text) & ")"
                    Cn.Execute strSQL
            
                    strSQL = "delete From Notes where noteid=" & val(TXTNoteID.text)
    
                    Cn.Execute strSQL, , adExecuteNoRecords
                
                    DeleteTransactiomsVoucher val(Text1.text)
                    CuurentLogdata ("D")
                    rs.delete
                    Cn.CommitTrans
                    BegainTrans = False
                    rs.MoveFirst
                    close_order2 order_no
                 
                    If rs.RecordCount < 1 Then
                        clear_all Me
                        TxtModFlg_Change
                        XPTxtCurrent.Caption = 0
                        XPTxtCount.Caption = 0
                    Else
                        Retrive
                    End If
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
    Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & Chr(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «·”Ã· "
    Msg = Msg & Chr(13) & Err.Number
    Msg = Msg & Chr(13) & Err.description
    Msg = Msg & Chr(13) & Err.Source
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title

    If BegainTrans = True Then
        rs.CancelUpdate
        Cn.RollbackTrans
        BegainTrans = False
    End If

End Sub

Private Sub AddTip()
    Dim Wrap As String
    Set TTP = New clstooltip
    On Error GoTo ErrTrap
    Wrap = Chr(13) + Chr(10)

    With TTP
        .Create Me.hwnd, "»Ì«‰«  ›« Ê—… «·‘—«¡", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«  ⁄„·Ì… ‘—«¡ ÃœÌœ…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«" & Wrap & "„›« ÌÕ «·«Œ ’«— F12 OR Enter", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  ›« Ê—… «·‘—«¡", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(7), "ÿ»«⁄… ..." & Wrap & "·⁄—÷ «·»Ì«‰«  «·Õ«·Ì… ›Ì  ﬁ—Ì— " & Wrap & " Ì„ﬂ‰ ÿ»«⁄ Â ⁄‰ ÿ—Ìﬁ «·ÿ«»⁄…" & Wrap & "„›« ÌÕ «·«Œ ’«— F6", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  ›« Ê—… «·‘—«¡", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  ⁄„·Ì… «·‘—«¡" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«" & Wrap & "„›« ÌÕ «·«Œ ’«— F11", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  ›« Ê—… «·‘—«¡", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  ⁄„·Ì… «·‘—«¡ «·ÃœÌœ…" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«" & Wrap & "„›« ÌÕ «·«Œ ’«— F10", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  ›« Ê—… «·‘—«¡", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·‘—«¡" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«" & Wrap & "„›« ÌÕ «·«Œ ’«— F9", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  ⁄„·Ì«  «·‘—«¡", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  ⁄„·Ì… ‘—«¡" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«" & Wrap & "„›« ÌÕ «·«Œ ’«— F8", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  ›« Ê—… «·‘—«¡", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(5), "»ÕÀ ..." & Wrap & "···»ÕÀ ⁄‰ ⁄„·Ì… ‘—«¡" & Wrap & "Ì‰ÿ»ﬁ ⁄·ÌÂ« ‘—Êÿ „⁄Ì‰…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«" & Wrap & "„›« ÌÕ «·«Œ ’«— F7", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  ›« Ê—… «·‘—«¡", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«" & Wrap & "„›« ÌÕ «·«Œ ’«— Ctrl + X", True
    End With

    'With TTP
    '   .Create Me.hwnd, "»Ì«‰«  ›« Ê—… «·‘—«¡", 1, 15204351, -2147483630
    '   .MaxWidth = 4000
    '   .VisibleTime = 9000
    '   .DelayTime = 600
    '   .AddControl XPBtnAdd, _
    '    "≈÷«›… «·√’‰«› ..." & Wrap & _
    '    " ·«÷«›… ’‰› ÃœÌœ" & Wrap & _
    '    " ›ﬁÿ ≈÷€ÿ Â‰«" & Wrap & _
    '    "„›« ÌÕ «·«Œ ’«— F2", True
    'End With
    'With TTP
    '   .Create Me.hwnd, "»Ì«‰«  ›« Ê—… «·‘—«¡", 1, 15204351, -2147483630
    '   .MaxWidth = 4000
    '   .VisibleTime = 9000
    '   .DelayTime = 600
    '   .AddControl XPBtnRemove, _
    '    "Õ–› ’‰› ..." & Wrap & _
    '    "·Õ–› √Õœ «·√’‰«›" & Wrap & _
    '    " ÕœœÂ Ê«÷€ÿ Â‰«" & Wrap & _
    '    "„›« ÌÕ «·«Œ ’«— F3", True
    'End With
    With TTP
        .Create Me.hwnd, "»Ì«‰«  ›« Ê—… «·‘—«¡", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnNewClients, "≈÷«›… ⁄„Ì· ÃœÌœ ..." & Wrap & "· ”ÃÌ· »Ì«‰«  ⁄„Ì· ÃœÌœ" & Wrap & " «÷€ÿ Â‰«" & Wrap & "„›« ÌÕ «·«Œ ’«— F5", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  ›« Ê—… «·‘—«¡", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  ›« Ê—… «·‘—«¡", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  ›« Ê—… «·‘—«¡", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  ›« Ê—… «·‘—«¡", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    'With TTP
    '   .Create Me.hwnd, "»Ì«‰«  ›« Ê—… «·‘—«¡", 1, 15204351, -2147483630
    '   .MaxWidth = 4000
    '   .VisibleTime = 9000
    '   .DelayTime = 600
    '   .AddControl XPFillData, _
    '    " ⁄»∆… »Ì«‰«  «·√’‰«›" & Wrap & _
    '    "· ⁄»∆… »Ì«‰«  «·√’‰«› ›Ì" & Wrap & _
    '    "›Ì ‰«›–… ÕÊ«—" & Wrap & _
    '    "  ≈÷€ÿ Â‰«" & Wrap & _
    '    "„›« ÌÕ «·«Œ ’«— Ctrl + Space", True
    'End With
    With TTP
        .Create Me.hwnd, "»Ì«‰«  ›« Ê—… «·‘—«¡", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl CmdHelp, "„”«⁄œ… ..." & Wrap & "·· ⁄—› ⁄·Ï ÊŸÌ›… Â–Â «·‰«›–…" & Wrap & "ÊﬂÌ›Ì… «· ⁄«„· „⁄Â«" & Wrap & "≈÷€ÿ Â‰«" & Wrap, True
    End With

    Exit Sub
ErrTrap:
End Sub

Function Closeorders()
    On Error Resume Next
    Dim i As Integer
    Dim rs3 As ADODB.Recordset
    Set rs3 = New ADODB.Recordset

    Dim Rs4 As ADODB.Recordset
    Set Rs4 = New ADODB.Recordset

    Dim sql As String
 
    Dim differnt As Integer
    Dim order_qty As Integer
    Dim QTYRecived As Integer
    Dim close_order As Boolean

    Dim j As Integer

    With GRID2

        For i = 1 To GRID2.Rows - 1
            close_order = True

            If .Cell(flexcpChecked, i, .ColIndex("select")) = flexChecked Then
           
                sql = "select * from QRY_items_orders_data where order_no='" & GRID2.TextMatrix(i, GRID2.ColIndex("order_no")) & "'"
                rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText

                If rs3.RecordCount = 0 Then GoTo LL

                For j = 1 To rs3.RecordCount
                    order_qty = IIf(IsNull(rs3("Quantity").value), 0, rs3("Quantity").value)
                    QTYRecived = IIf(IsNull(rs3("QTYRSV").value), 0, rs3("QTYRSV").value)
                    differnt = order_qty - QTYRecived

                    If differnt <= 0 Then
                        close_order = False
                    End If
                
                    rs3.MoveNext
                Next j
           
                If close_order = True Then
                    sql = "select * from Transactions where Transaction_Type=6 and order_no='" & GRID2.TextMatrix(i, GRID2.ColIndex("order_no")) & "'"
                    Rs4.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
                    Rs4("Closed").value = 1
                    Rs4.update
                    Rs4.Close
                   
                End If
               
            End If
       
            rs3.Close
LL:
        Next
       
    End With
       
End Function

Private Sub SaveData()
    Dim usedaccount As Integer
    Dim RSTransDetails As ADODB.Recordset
    Dim RsNotes As ADODB.Recordset
    Dim RsNotesGeneral As ADODB.Recordset
    Dim RsTemp As New ADODB.Recordset
    Dim Msg As String
    Dim RowNum As Integer
    Dim strSQL As String
    Dim StrSqlDel As String
    Dim SearchResault As Integer
    Dim note_id As Long
    Dim RsDetalis  As ADODB.Recordset
    Dim BeginTrans As Boolean
    Dim LnItemID As Long
    Dim i As Long
    Dim StrCurrentItemName As String
    Dim DblNotesTotal As Double

    Dim IntLineNO As Integer
    Dim StrAccountCode As String
    Dim TotalBillDiscount As Double
    Dim TotalDiscountPerLine As Double
    On Error GoTo ErrTrap

    If CboPayMentType.ListIndex = 1 Then
        XPChkPayType(1).value = 1
        '  XPTxtValue(1).text = Val(LblTotalAll.Caption)
        XPTxtValue(1).text = val(LblTotal.Caption)

    Else
        XPChkPayType(0).value = 1
        '  XPTxtValue(0).text = Val(LblTotalAll.Caption)
        XPTxtValue(0).text = val(LblTotal.Caption)

    End If
        
    If Dccurrency.BoundText = "" Then
    
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "«Œ — «·⁄„·… «Ê·« "
        Else
            Msg = "Select Currency First"
    
        End If

        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Dccurrency.SetFocus
        SendKeys "{F4}"
        Screen.MousePointer = vbDefault
        Exit Sub
    
    End If

    If Due_Date > DtpDelayDate.value Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "ÌÃ» «‰ ÌﬂÊ‰  «—ÌŒ «·«” Õﬁ«ﬁ «ﬂ»—   „‰ «Ê Ì”«ÊÌ     «—ÌŒ «Œ— ﬁ”ÿ"
        Else
            MsgBox "installment Date Must be Graeter than  or equal todya"
    
        End If

        Exit Sub
    End If

    If CboPayMentType.ListIndex = 1 Then
        Me.XPChkPayType(1).value = 1
        ' hany  XPTxtValue(1).text = Val(LblTotalAll.Caption)
    End If

    If Trim(Me.TxtTransSerial.text) = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "ÌÃ» ﬂ «»… —ﬁ„ ›« Ê—… «·‘—«¡..!!!"
        Else
            Msg = "Must Enter Bill No."
    
        End If
    
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Me.TxtTransSerial.SetFocus
        Exit Sub
    End If

    If Me.TxtModFlg.text = "N" Then
        If RepeatSerial(Trim(Me.TxtTransSerial.text), 1, 0, val(Me.DBCboClientName.BoundText)) = True Then
            Exit Sub
        End If

    ElseIf Me.TxtModFlg.text = "E" Then

        If RepeatSerial(Trim(Me.TxtTransSerial.text), 1, val(Me.XPTxtBillID.text), val(Me.DBCboClientName.BoundText)) = True Then
            Exit Sub
        End If
    End If

    '«· √ﬂœ „‰ ⁄œ„  ﬂ—«— —ﬁ„ «·”‰œ
    Dim BolTemp As Boolean

    If Voucher_coding(val(my_branch), XPDtbBill.value, 6, 150, , 22) = "" Then
        If Me.TxtModFlg.text = "N" Then
    
            BolTemp = UniqueNoteSerial1(Trim(Me.TxtNoteSerial1.text), 22, , val(dcBranch.BoundText))
        ElseIf Me.TxtModFlg.text = "E" Then
            BolTemp = UniqueNoteSerial1(Trim(Me.TxtNoteSerial1.text), 22, val(Me.XPTxtBillID.text), val(dcBranch.BoundText))
        End If
 
        If BolTemp = False Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "—ﬁ„ «·”‰œ „”Ã· „”»ﬁ« ›Ï «·»—‰«„Ã.." & Chr(13)
                Msg = Msg & "Ê·«Ì„ﬂ‰  ﬂ—«— —ﬁ„ «·”‰œ"
            Else
                Msg = "This Bill No Already Exist" & Chr(13)
        
            End If

            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            TxtNoteSerial1.SetFocus
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    
    End If

    '‰Â«Ì… «· √ﬂœ

    Screen.MousePointer = vbArrowHourglass

    If DBCboClientName.text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "„‰ ›÷·ﬂ √œŒ· «”„ «·⁄„Ì·"
        Else
            Msg = "Select Customer Name"
    
        End If

        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        DBCboClientName.SetFocus
        SendKeys "{F4}"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    If DCboStoreName.text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "„‰ ›÷·ﬂ Õœœ «”„ «·„Œ“‰"
        Else
            Msg = "Select Inventory First"
        End If

        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        DCboStoreName.SetFocus
        SendKeys "{F4}"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    If XPCboDiscountType.ListIndex = 1 Or XPCboDiscountType.ListIndex = 2 Then
        If XPTxtDiscountVal.text = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "ÌÃ»  ÕœÌœ ﬁÌ„… «·Œ’„ «·ﬂ·Ì ⁄·Ï «·›« Ê—…"
            Else
        
                Msg = "Specify Total Discount"
            End If

            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            XPTxtDiscountVal.SetFocus
            Screen.MousePointer = vbDefault
            Exit Sub
        End If

        If Not IsNumeric(XPTxtDiscountVal.text) Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "ﬁÌ„… «·Œ’„ «·ﬂ·Ì ⁄·Ï «·›« Ê—… ÌÃ» √‰  ﬂÊ‰ ﬁÌ„… —ﬁ„Ì…"
            Else
                Msg = "Discount Value Must be Numeric"
            End If

            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            XPTxtDiscountVal.SetFocus
            Screen.MousePointer = vbDefault
            Exit Sub
        End If

        XPTxtDiscountVal.SetFocus
    End If

    If CboPayMentType.ListIndex = -1 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "ÌÃ»  ÕœÌœ ÿ—Ìﬁ… «·œ›⁄"
        Else
            Msg = "Specify Payment Method"
    
        End If

        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        CboPayMentType.SetFocus
        SendKeys "{F4}"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    If XPChkPayType(0).value = vbChecked Then
        If Me.DcboBox.BoundText = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "ÌÃ»  ÕœÌœ «”„ «·Œ“‰…...!!!"
            Else
                Msg = "Specify Box Name "
            End If

            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Screen.MousePointer = vbDefault
            Exit Sub
        End If

        If Me.TxtModFlg.text = "N" Then
            If CheckBoxAccount(Me.DcboBox.BoundText, val(Me.XPTxtValue(0).text), Me.XPDtbBill.value) = False Then
                Screen.MousePointer = vbDefault
                Exit Sub
            End If

        ElseIf Me.TxtModFlg.text = "E" Then

            If CheckBoxAccount(Me.DcboBox.BoundText, val(Me.XPTxtValue(0).text), Me.XPDtbBill.value, , , val(Me.XPTxtValue(0).Tag)) = False Then
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
        End If
    End If

    If val(Me.XPTxtValue(1).text) > 0 Then
        If ChkInstall.value = vbChecked Then
            If val(Me.LblInstallTotal.Caption) = 0 Then
                Msg = "ÌÃ» Õ”«» «·√ﬁ”«ÿ ﬁ»· ⁄„·Ì… «·Õ›Ÿ..!!!"
                MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Me.XPTab301.CurrTab = 1
                Screen.MousePointer = vbDefault
                Exit Sub
            End If

            '  hany      If Val(Me.LblInstallTotal.Caption) <> Val(Me.XPTxtValue(1).text) Then
            '            Me.XPTxtValue(1).text = Val(Me.LblInstallTotal.Caption)
            '        End If
        End If
    End If

    If XPChkPayType(2).value = vbChecked Then
        If val(Me.lbl(18).Caption) = 0 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "ÌÃ» ≈œŒ«· «·‘Ìﬂ«  ﬁ»· ⁄„·Ì… «·Õ›Ÿ..!!!"
            Else
                Msg = "Enter Cheques Data Before Save"
            End If

            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Me.XPTab301.CurrTab = 1
            Screen.MousePointer = vbDefault
            Exit Sub
        End If

        If dcbanks.BoundText = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = Msg + "ÌÃ»  ÕœÌœ «”„ «·»‰ﬂ     " & Chr(13)
            Else
                Msg = Msg + " Specify Bank NAme     " & Chr(13)
            End If
        
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            '            Dcbanks.SetFocus
            SendKeys "{F4}"
            Screen.MousePointer = vbDefault
            Exit Sub
        Else
    
            Dim rsbank As New ADODB.Recordset
            Set rsbank = New ADODB.Recordset
            rsbank.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
       
            If Not (rsbank.EOF Or rsbank.BOF) Then
                If rsbank!banks_Accounts = True Then
                    bank_account = get_bank_Account(val(Me.dcbanks.BoundText), "Account_Code2")
                Else
                    bank_account = get_bank_Account(val(Me.dcbanks.BoundText), "Account_Code")
                End If
            End If
        
        End If
    
    End If

    XPTab301.CurrTab = 0

    If NewGrid.CheckDataEntered = False Then
        Exit Sub
    End If

    If CheckMyData = False Then
        Exit Sub
    End If

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '    If Me.TxtModFlg.text = "E" Then
    '        If EditTransStatus(Val(Me.XPTxtBillID.text), "E", NewGrid) = False Then
    '            Exit Sub
    '        End If
    '    End If
    '---------------------------------------------------------------
    Cn.Execute "delete DOUBLE_ENTREY_VOUCHERS where Transaction_ID = " & val(Text2.text)

    If NewGrid.Calculate(1, , False, True) = False Then
        Screen.MousePointer = vbDefault
        '  Exit Sub
    End If

    '-------------------------------
    '    If Me.XPChkPayType(0).value = vbChecked Then
    '        DblNotesTotal = Val(Me.XPTxtValue(0).text)
    '    End If

    '    If Me.XPChkPayType(1).value = vbChecked Then
    '        Me.XPTxtValue(1).text = LblTotal.Caption
    DblNotesTotal = DblNotesTotal + val(Me.XPTxtValue(1).text)
    '    End If

    '    If Me.XPChkPayType(2).value = vbChecked Then
    '        DblNotesTotal = DblNotesTotal + Val(Me.lbl(18).Caption)
    '    End If
    DblNotesTotal = val(Me.XPTxtValue(0).text) + val(Me.XPTxtValue(1).text) + val(lbl(18).Caption)

    If CboPayMentType.ListIndex = 1 Then
        Me.XPChkPayType(1).value = 1
        '  XPTxtValue(1).text = Val(LblTotalAll.Caption)
    End If

    '   If CboPayMentType.ListIndex = 0 Then
    '       Me.XPChkPayType(0).value = 1
    '         XPTxtValue(0).text = Val(LblTotalAll.Caption)
    '   End If
     
    If DblNotesTotal <> val(LblTotal.Caption) Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "≈Ã„«·Ï «·√Ê—«ﬁ «·„«·Ì… €Ì— „ ”«ÊÏ „⁄ ≈Ã„«·Ï «·›« Ê—…...!!!"
        Else
            Msg = "Error In total ...!!!"
        End If

        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

    '---------Start Saving------------------------------------------------
    '    Me.TxtNoteSerial.text = IIf(IsNull(rs("NoteSerial").value), "", (rs("NoteSerial").value))
    'Me.TxtNoteSerial1.text = IIf(IsNull(rs("NoteSerial1").value), "", (rs("NoteSerial1").value))
    'Me.TXTNoteID.text = IIf(IsNull(rs("NoteID").value), "", (rs("NoteID").value))

    'Õ›Ÿ «·„’—Ê›«  «·«÷«›Ì… Ê«·›Ê« Ì— «·„«·Ì…
    Save_Financial_invoice
    save_expenses

    '---------Notes ID ------------------------------------------------
    'Create big notes
    GoTo xll

    If TxtNoteSerial.text = "" Then
        If Notes_coding(val(my_branch), XPDtbBill.value) = "error" Then
            MsgBox " ·« Ì„ﬂ‰ «÷«›… ﬁÌÊœ ·Â–… «·⁄„·Ì… ·«‰ﬂ  ⁄œÌ  «·Õœ «·«ﬁ’Ì ··ﬁÌÊœ ﬂ„« Õœœ  ›Ì  —ﬁÌ„ «·”‰œ«  ": Exit Sub
        Else
                       
            If Notes_coding(val(my_branch), XPDtbBill.value) = "" Then
                MsgBox " ·«Ì„ﬂ‰ «‰‘«¡ «·ﬁÌœ ·Â–« «·„” ‰œ ·«‰ﬂ Õœœ   —ﬁÌ„ ﬁÌÊœ ÌœÊÌ  ": Exit Sub
            Else
                TxtNoteSerial.text = Notes_coding(val(my_branch), XPDtbBill.value)
            End If
        End If
    End If
        
    If TxtNoteSerial1.text = "" Then
        If Voucher_coding(val(my_branch), XPDtbBill.value, 6, 150, , 22) = "error" Then
            MsgBox " ·« Ì„ﬂ‰ «÷«›… ”‰œ ›« Ê—… „‘ —Ì«  ÃœÌœ ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": Exit Sub
        Else
                       
            If Voucher_coding(val(my_branch), XPDtbBill.value, 6, 150, , 22) = "" Then
                MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ ›« Ê—… „‘ —Ì«   ÌœÊÌ« ﬂ„« Õœœ   ": Exit Sub
            Else
                TxtNoteSerial1.text = Voucher_coding(val(my_branch), XPDtbBill.value, 6, 150, , 22)
            End If
        End If
    End If
     
xll:
    Set RsNotesGeneral = New ADODB.Recordset
'    RsNotesGeneral.Open "[Notes]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
     strSQL = "SELECT     dbo.Notes.* from dbo.Notes Where (NoteID = -1)"
   RsNotesGeneral.Open strSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText

    If Me.TxtModFlg.text = "N" Then
        Me.oldtxtNoteSerial1.text = Trim$(Me.TxtNoteSerial1.text)
    Else
        StrSqlDel = "delete From Transaction_Details where Transaction_ID=" & val(rs("Transaction_ID").value)
        Cn.Execute StrSqlDel, , adExecuteNoRecords
        StrSqlDel = "delete From Notes where Transaction_ID=" & val(rs("Transaction_ID").value)
        Cn.Execute StrSqlDel, , adExecuteNoRecords
        strSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Transaction_ID=" & val(Me.XPTxtBillID.text)
        Cn.Execute strSQL
        
        StrSqlDel = "delete From Notes where noteid=" & val(TXTNoteID.text)
        Cn.Execute StrSqlDel
        
        general_noteid = val(TXTNoteID.text)
    End If

    RsNotesGeneral.AddNew
    RsNotesGeneral("NoteID").value = CStr(new_id("Notes", "NoteID", "", True))
    general_noteid = RsNotesGeneral("NoteID").value
    TXTNoteID.text = general_noteid
    ' RsNotesGeneral("Transaction_ID").value = Val(XPTxtBillID.text)
    RsNotesGeneral("NoteDate").value = XPDtbBill.value
    RsNotesGeneral("NoteType").value = 150
    RsNotesGeneral("Note_Value").value = val(LblTotal.Caption)
    RsNotesGeneral("NoteSerial").value = IIf(Trim(Me.TxtNoteSerial.text) = "", Null, Trim(Me.TxtNoteSerial.text))
    RsNotesGeneral("NoteSerial1").value = IIf(Trim(Me.TxtNoteSerial1.text) = "", Null, Trim(Me.TxtNoteSerial1.text))
    RsNotesGeneral("OldNoteSerial1").value = Trim$(Me.oldtxtNoteSerial1.text) '
        
    RsNotesGeneral("numbering_type").value = sand_numbering_type(0) '”‰œ «·ﬁÌœ
    RsNotesGeneral("numbering_type1").value = sand_numbering_type(6) '  ›« Ê—… ‘ƒ«¡
    RsNotesGeneral("sanad_year").value = year(XPDtbBill.value)
    RsNotesGeneral("sanad_month").value = Month(XPDtbBill.value)
    'RsNotes("note_value_by_characters").value = Trim$(Me.lbl(18).Caption)
    RsNotesGeneral("branch_no").value = val(Me.dcBranch.BoundText)
    RsNotesGeneral.update

    '---------Start Saving------------------------------------------------

    Set RSTransDetails = New ADODB.Recordset
    Set RsNotes = New ADODB.Recordset
 '   RSTransDetails.Open "[Transaction_Details]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
 '   RsNotes.Open "[Notes]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    
       strSQL = "SELECT     dbo.Notes.* from dbo.Notes Where (NoteID = -1)"
   RsNotes.Open strSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText


 strSQL = "SELECT     dbo.Transaction_Details.* from dbo.Transaction_Details Where (Transaction_ID = -1)"
   RSTransDetails.Open strSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText

    
    Screen.MousePointer = vbArrowHourglass
    Cn.BeginTrans
    BeginTrans = True

    If Me.TxtModFlg.text = "N" Then
        XPTxtBillID.text = CStr(new_id("Transactions", "Transaction_ID", "", True))
        rs.AddNew
        rs("Transaction_ID").value = val(XPTxtBillID.text)
    ElseIf Me.TxtModFlg.text = "E" Then

        If rs("Transaction_ID").value <> val(XPTxtBillID.text) Then
            rs.find "Transaction_ID=" & val(XPTxtBillID.text), , adSearchForward, 1
        End If

        StrSqlDel = "delete From Transaction_Details where Transaction_ID=" & val(rs("Transaction_ID").value)
        Cn.Execute StrSqlDel, , adExecuteNoRecords
        '   StrSqlDel = "delete From Notes where Transaction_ID=" & Val(rs("Transaction_ID").value)
        '   Cn.Execute StrSqlDel, , adExecuteNoRecords
        '   StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Transaction_ID=" & Val(Me.XPTxtBillID.text)
        '   Cn.Execute StrSQL, , adExecuteNoRecords
        
        '   StrSqlDel = "delete From Notes where noteid=" & Val(TxtNoteID.text)
        ' Cn.Execute StrSqlDel, , adExecuteNoRecords
        
    End If

    rs("ManualNo1").value = IIf(TxtManualNo1.text = "", Null, val(TxtManualNo1.text))
    rs("BranchId").value = IIf(Me.dcBranch.BoundText = "", 0, val(dcBranch.BoundText))
    rs("NoteSerial").value = IIf(Trim(Me.TxtNoteSerial.text) = "", Null, Trim(Me.TxtNoteSerial.text))
    rs("NoteSerial1").value = IIf(Trim(Me.TxtNoteSerial1.text) = "", Null, Trim(Me.TxtNoteSerial1.text))
    rs("OldNoteSerial1").value = Trim$(Me.oldtxtNoteSerial1.text) '
    rs("NoteId").value = val(TXTNoteID.text)
    rs("order_no").value = IIf((Txt_order_no.text) = "", Null, Txt_order_no.text)
    rs("Currency_rate").value = IIf(Not IsNumeric(txt_Currency_rate.text), 1, txt_Currency_rate.text)
    rs("Transaction_Serial").value = IIf(Trim(Me.TxtTransSerial.text) = "", Null, Trim(Me.TxtTransSerial.text))
    rs("Transaction_Date").value = XPDtbBill.value
    rs("ArrivalDate").value = DTArrivalDate.value
    rs("Doctype").value = IIf(Me.DCDocTypes.BoundText = "", Null, val(DCDocTypes.BoundText))
    rs("Transaction_Type").value = BillType
    rs("UserID").value = user_id
    rs("nots").value = Text1.text
    rs("TransactionComment").value = IIf(Trim$(TxtBillComment.text) = "", Null, Trim$(TxtBillComment.text))
    
    If XPCboDiscountType.ListIndex = -1 Then
        rs("Trans_DiscountType").value = 0
    Else
        rs("Trans_DiscountType").value = val(XPCboDiscountType.ListIndex)
    End If

    If XPCboDiscountType.ListIndex = -1 Or XPCboDiscountType.ListIndex = 0 Then
        rs("Trans_Discount").value = Null
 
    Else
        rs("Trans_Discount").value = IIf(XPTxtDiscountVal.text = "", Null, (XPTxtDiscountVal.text))
    End If

    If CboPayMentType.ListIndex = -1 Then
        rs("PaymentType").value = 0
    Else
        rs("PaymentType").value = val(CboPayMentType.ListIndex)
    End If

    rs("project_id").value = IIf(DCproject.BoundText = "", Null, (DCproject.BoundText))
    rs("CusID").value = IIf(DBCboClientName.BoundText = "", Null, (DBCboClientName.BoundText))
    rs("StoreID").value = IIf(DCboStoreName.BoundText = "", Null, (DCboStoreName.BoundText))
    rs("TaxFound").value = IIf(XPChkTAX.value = Checked, True, False)
    rs("TaxValue").value = IIf(XPTxtTaxValue.text = "", Null, val(XPTxtTaxValue.text))
    rs("ToTAlELSHahn").value = IIf(Not IsNumeric(TXTToTAlELSHahn.text), 0, Me.TXTToTAlELSHahn.text)
    
    rs("total_expenses").value = IIf(Not IsNumeric(Txt_EXport.text), 0, Txt_EXport.text)
    rs("total_payments").value = IIf(Not IsNumeric(txt_total_bill.text), 0, txt_total_bill.text)
    rs("LcNo").value = IIf(TxtLcNo.text = "", Null, (TxtLcNo.text))

    '÷—»Ì… Œ’„ Ê≈÷«›…
    If ChkTaxAdd.value = vbChecked And val(Me.TxtTaxAddValue.text) > 0 Then
        rs("TaxAddValue").value = val(Me.TxtTaxAddValue.text)
    Else
        rs("TaxAddValue").value = 0
    End If

    '÷—»Ì… œ„€…
    If ChkTaxStamp.value = vbChecked And val(Me.TxtTaxStampValue.text) > 0 Then
        rs("TaxStampValue").value = val(Me.TxtTaxStampValue.text)
    Else
        rs("TaxStampValue").value = 0
    End If

    '÷—»Ì… Œœ„…
    If ChkTaxSerivce.value = vbChecked And val(Me.TxtTaxServiceValue.text) > 0 Then
        rs("TaxServiceValue").value = val(Me.TxtTaxServiceValue.text)
    Else
        rs("TaxServiceValue").value = 0
    End If

    'rs("Shipment_no").value = IIf(txt_Shipment_no.text = "", Null, (txt_Shipment_no.text))
    rs("order_no").value = IIf(Txt_order_no.text = "", Null, (Txt_order_no.text))
    rs("Currency_id").value = IIf(Dccurrency.BoundText = "", Null, val(Dccurrency.BoundText))
    rs("BoxID").value = IIf(Me.DcboBox.BoundText = "", Null, val(DcboBox.BoundText))
    rs("ManualNO").value = IIf(txtManualNO.text = "", Null, (txtManualNO.text))

    If XPCboDiscountType.ListIndex = -1 Then
        rs("CBoBasedON").value = 0
    Else
        rs("CBoBasedON").value = val(CBoBasedON.ListIndex)
    End If
    
    If BillBasedOn(0).value = True Then
        rs("BillBasedOn").value = 0
    ElseIf BillBasedOn(1).value = True Then
        rs("BillBasedOn").value = 1
    ElseIf BillBasedOn(2).value = True Then
        rs("BillBasedOn").value = 2
    End If
    
    rs.update

    If Me.TxtModFlg.text = "E" Then
        StrSqlDel = "delete From Transaction_Details where Transaction_ID=" & val(rs("Transaction_ID").value)
        Cn.Execute StrSqlDel, , adExecuteNoRecords
        '   StrSqlDel = "delete From Notes where Transaction_ID=" & Val(rs("Transaction_ID").value)
        '   Cn.Execute StrSqlDel, , adExecuteNoRecords
    End If

    For RowNum = 1 To FG.Rows - 1

        'Check Repeat Serial
        If FG.TextMatrix(RowNum, FG.ColIndex("Serial")) <> "" Then
            strSQL = "select * From Transaction_Details where ItemSerial='" & FG.TextMatrix(RowNum, FG.ColIndex("Serial")) & "'"
            strSQL = strSQL + " and Transaction_ID =" & XPTxtBillID.text
            Set RsTemp = New ADODB.Recordset
            RsTemp.Open strSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

            If Not (RsTemp.EOF Or RsTemp.BOF) Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "«·”Ì—Ì«· «·Œ«’ »«·’‰›" & Chr(13)
                    Msg = Msg + FG.Cell(flexcpTextDisplay, RowNum, FG.ColIndex("name")) & Chr(13)
                    Msg = Msg + " „ √œŒ«·Â ·ﬁÿ⁄… √Œ—Ï ›Ì Â–Â «·›« Ê—…"
                Else
                    Msg = "Item Serial" & Chr(13)
                    Msg = Msg + FG.Cell(flexcpTextDisplay, RowNum, FG.ColIndex("name")) & Chr(13)
                    Msg = Msg + "Already Exist in this bill"
            
                End If

                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                RsTemp.Close
                XPTab301.CurrTab = 0
                FG.Row = RowNum
                FG.Col = FG.ColIndex("name")
                FG.ShowCell RowNum, FG.ColIndex("name")
                FG.SetFocus
                Screen.MousePointer = vbDefault
                BeginTrans = False
                Cn.RollbackTrans
                Exit Sub
            End If

            RsTemp.Close
        End If

        If FG.TextMatrix(RowNum, FG.ColIndex("Code")) <> "" Then
            RSTransDetails.AddNew
            RSTransDetails("BranchId").value = Me.dcBranch.BoundText
            RSTransDetails("OrderArrivalDate").value = IIf(Not IsDate(FG.TextMatrix(RowNum, FG.ColIndex("OrderArrivalDate"))), Null, FG.TextMatrix(RowNum, FG.ColIndex("OrderArrivalDate")))
            RSTransDetails("Transaction_ID").value = val(XPTxtBillID.text)
            RSTransDetails("Item_ID").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Code")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Code"))))
            RSTransDetails("Quantity").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Count")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Count"))))

            '            RSTransDetails("ItemName").Value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Name")) _
            '            = ""), "", Val(FG.TextMatrix(RowNum, FG.ColIndex("Name"))))
            If Not FG.TextMatrix(RowNum, FG.ColIndex("Name")) = "" Then
                strSQL = "select * From TblItems where ItemID=" & FG.TextMatrix(RowNum, FG.ColIndex("Name"))
                RsTemp.Open strSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If Not (RsTemp.EOF Or RsTemp.BOF) Then
                    If RsTemp("HaveSerial").value = True Then
                        RSTransDetails("ItemSerial").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Serial")) = ""), Null, (FG.TextMatrix(RowNum, FG.ColIndex("Serial"))))
                    End If
                End If

                RsTemp.Close
            End If

            RSTransDetails("ItemCase").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemCase")) = ""), Null, (FG.TextMatrix(RowNum, FG.ColIndex("ItemCase"))))
            RSTransDetails("showPrice").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Price")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Price"))))
            RSTransDetails("ItemDiscountType").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("DiscountType")) = ""), Null, (FG.TextMatrix(RowNum, FG.ColIndex("DiscountType"))))
            RSTransDetails("ItemDiscount").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("DiscountVal")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("DiscountVal"))))
            RSTransDetails("guaranteeTime").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("guaranteeTime")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("guaranteeTime"))))
        
            RSTransDetails("Remarks").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Remarks")) = ""), Null, Trim$(FG.TextMatrix(RowNum, FG.ColIndex("Remarks"))))
         
            '.TextMatrix(LngRow, .ColIndex("ColorID")) = 1
            '.TextMatrix(LngRow, .ColIndex("ItemSize")) = 0
        
            RSTransDetails("ColorID").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ColorID")) = ""), 1, val(FG.TextMatrix(RowNum, FG.ColIndex("ColorID"))))
            RSTransDetails("ItemSize").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemSize")) = ""), 1, Trim$(FG.TextMatrix(RowNum, FG.ColIndex("ItemSize"))))
            RSTransDetails("ClassId").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ClassId")) = ""), 1, val(FG.TextMatrix(RowNum, FG.ColIndex("ClassId"))))
            
            RSTransDetails("UnitID").value = IIf(FG.Cell(flexcpData, RowNum, FG.ColIndex("UnitID")) = "", Null, (FG.Cell(flexcpData, RowNum, FG.ColIndex("UnitID"))))
         
            RSTransDetails("ShowQty").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Count")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Count"))))
            RSTransDetails("order_no").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("order_no")) = ""), Null, FG.TextMatrix(RowNum, FG.ColIndex("order_no")))
            RSTransDetails("OrderArrivalDate").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("OrderArrivalDate")) = ""), Null, FG.TextMatrix(RowNum, FG.ColIndex("OrderArrivalDate")))
             
            RSTransDetails("LineShahn").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("LineShahn")) = ""), 0, FG.TextMatrix(RowNum, FG.ColIndex("LineShahn")))
             
            Dim RsUnitData As ADODB.Recordset
            Dim LngCurItemID As Long
            Dim LngUnitID As Long
            Dim DblQty As Double
        
            LngCurItemID = val(FG.TextMatrix(RowNum, FG.ColIndex("Code")))
            LngUnitID = val(FG.Cell(flexcpData, RowNum, FG.ColIndex("UnitID")))
            DblQty = val(FG.TextMatrix(RowNum, FG.ColIndex("Count")))

            strSQL = "Select * From TblItemsUnits Where ItemID=" & LngCurItemID
            strSQL = strSQL + " AND UnitID=" & LngUnitID
            Set RsUnitData = New ADODB.Recordset
            RsUnitData.Open strSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

            If Not (rs.BOF Or rs.EOF) Then
                RSTransDetails("QtyBySmalltUnit").value = RsUnitData("UnitFactor").value
                RSTransDetails("Quantity").value = RSTransDetails("QtyBySmalltUnit").value * RSTransDetails("showqty").value
            End If

            '          RSTransDetails("Price").value = IIf((Fg.TextMatrix(RowNum, Fg.ColIndex("ShowPrice")) = ""), Null, Val(Fg.TextMatrix(RowNum, Fg.ColIndex("ShowPrice"))))
            '     RSTransDetails("price").value = Round(FG.TextMatrix(RowNum, FG.ColIndex("Valu")) / RSTransDetails("Quantity").value, 2)
            RSTransDetails("price").value = Round(FG.TextMatrix(RowNum, FG.ColIndex("Price")) / RSTransDetails("QtyBySmalltUnit").value, 15)
    
            RSTransDetails("OpeningBurcahseQty").value = RSTransDetails("Quantity").value
            RSTransDetails("OpeningBurcahseValue").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Valu")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Valu"))))
            RSTransDetails("discountvalue").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("discountvalue")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("discountvalue")))) / RSTransDetails("Quantity").value
      
            RSTransDetails("rate").value = val(txt_Currency_rate.text)

            If val(LblTotal.Caption) = 0 Then LblTotal.Caption = 1
            ' RSTransDetails("ToTAlELSHahn") = Round((((RSTransDetails("showPrice") * _
            ' RSTransDetails("ShowQty")) / Val(LblTotal.Caption)) * _
            ' Val(TXTToTAlELSHahn.text)) / RSTransDetails("ShowQty"), 2)   ' / RSTransDetails("ShowQty")
            Dim TotalShahnPerLine As Double
            TotalShahnPerLine = ((((RSTransDetails("price") * RSTransDetails("Quantity") / (LblTotal.Caption))) * val(TXTToTAlELSHahn.text)) / RSTransDetails("Quantity"))
            TotalShahnPerLine = Round(TotalShahnPerLine, 15) 'Val(Format(TotalShahnPerLine, "." & String(Abs(18), "#")))
            RSTransDetails("ToTAlELSHahn") = TotalShahnPerLine
         
            If Me.XPCboDiscountType.ListIndex = 1 Then
                TotalBillDiscount = IIf(XPTxtDiscountVal.text = "", Null, (XPTxtDiscountVal.text))
                     
            ElseIf XPCboDiscountType.ListIndex = 2 Then

                If XPTxtDiscountVal.text <> "" Then
                    TotalBillDiscount = IIf(XPTxtDiscountVal.text = "", Null, (XPTxtDiscountVal.text)) * val(LblTotalAll.Caption) / 100
                             
                Else
                    TotalBillDiscount = 0
                End If
            End If
           
            TotalDiscountPerLine = ((((RSTransDetails("price") * RSTransDetails("Quantity") / (LblTotalAll.Caption))) * val(TotalBillDiscount)) / RSTransDetails("Quantity"))
            RSTransDetails("TotalDiscountPerLine") = TotalDiscountPerLine
            
            ' RSTransDetails.update
        End If

        RSTransDetails("ColorID").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ColorID")) = ""), 1, val(FG.TextMatrix(RowNum, FG.ColIndex("ColorID"))))
            
        RSTransDetails("ItemSize").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ItemSize")) = ""), 1, val(FG.TextMatrix(RowNum, FG.ColIndex("ItemSize"))))
            
        RSTransDetails("ClassId").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ClassId")) = ""), 1, val(FG.TextMatrix(RowNum, FG.ColIndex("ClassId"))))
            
        RSTransDetails("BranchId").value = IIf(Me.dcBranch.BoundText = "", 0, val(dcBranch.BoundText))
        RSTransDetails("OrderArrivalDate").value = IIf(Not IsDate(FG.TextMatrix(RowNum, FG.ColIndex("OrderArrivalDate"))), Null, FG.TextMatrix(RowNum, FG.ColIndex("OrderArrivalDate")))
        RSTransDetails("order_no").value = IIf(FG.TextMatrix(RowNum, FG.ColIndex("order_no")) = "", Null, FG.TextMatrix(RowNum, FG.ColIndex("order_no")))
        RSTransDetails("FoxyNo").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("FoxyNo")) = ""), Null, FG.TextMatrix(RowNum, FG.ColIndex("FoxyNo")))
        RSTransDetails("ProductionDate").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ProductionDate")) = ""), Null, Format((FG.TextMatrix(RowNum, FG.ColIndex("ProductionDate"))), "DD/mm/YYYY"))
        RSTransDetails("ExpiryDate").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ExpiryDate")) = ""), Null, Format((FG.TextMatrix(RowNum, FG.ColIndex("ExpiryDate"))), "DD/mm/YYYY"))
        RSTransDetails("LotNO").value = IIf(FG.TextMatrix(RowNum, FG.ColIndex("LotNO")) = "", Null, FG.TextMatrix(RowNum, FG.ColIndex("LotNO")))


'******************************
            RSTransDetails("ExtraType").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ExtraType")) = ""), Null, (FG.TextMatrix(RowNum, FG.ColIndex("ExtraType"))))
            RSTransDetails("ExtraVal").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("ExtraVal")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("ExtraVal"))))
RSTransDetails("Commisionvalue").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("Commisionvalue")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("Commisionvalue"))))

    RSTransDetails("SBillNO").value = IIf(FG.TextMatrix(RowNum, FG.ColIndex("SBillNO")) = "", Null, FG.TextMatrix(RowNum, FG.ColIndex("SBillNO")))

RSTransDetails("SupplierID").value = IIf((FG.TextMatrix(RowNum, FG.ColIndex("SupplierID")) = ""), Null, val(FG.TextMatrix(RowNum, FG.ColIndex("SupplierID"))))

'******************************
        RSTransDetails.update
    Next RowNum

    '------------------------------------------------------------------------------
     
    '------------------------------------------------------------------------------
    If Me.XPChkPayType(0).value = Checked Then
        'RsNotes.AddNew
        'RsNotes("NoteID").value = CStr(new_id("Notes", "NoteID", "", True))
        'note_id = RsNotes("NoteID").value

        If Me.TxtModFlg.text = "N" Then
            '    RsNotes("NoteSerial").value = CStr(new_id("Notes", "NoteSerial", "", True))
            '    XPTxtSerial(0).text = RsNotes("NoteSerial").value
        ElseIf Trim(XPTxtSerial(0).text) <> "" Then
            '    RsNotes("NoteSerial").value = Trim(XPTxtSerial(0).text)
        Else
            '    RsNotes("NoteSerial").value = CStr(new_id("Notes", "NoteSerial", "", True))
            '    XPTxtSerial(0).text = RsNotes("NoteSerial").value
        End If

        '--------------------------------------------------------------------------
    End If

    If Me.XPChkPayType(1).value = Checked Then
        RsNotes.AddNew
        RsNotes("NoteID").value = CStr(new_id("Notes", "NoteID", "", True))
        note_id = RsNotes("NoteID").value
        RsNotes("NoteDate").value = XPDtbBill.value
 
        RsNotes("NoteSerial1").value = Me.TxtNoteSerial1.text
        RsNotes("NoteSerial").value = Null

        RsNotes("Transaction_ID").value = val(XPTxtBillID.text)
        RsNotes("NoteType").value = 1
        RsNotes("Note_Value").value = IIf(XPTxtValue(1).text = "", Null, val(XPTxtValue(1).text))
        RsNotes("Member_ID").value = IIf(DBCboClientName.BoundText = "", Null, val(DBCboClientName.BoundText))
        RsNotes("BankID").value = Null
        RsNotes("CusID").value = IIf(DBCboClientName.BoundText = "", Null, val(DBCboClientName.BoundText)) 'Null SALIM MY BE ERROR
        RsNotes("DueDate").value = DtpDelayDate.value
        RsNotes.update
 
    End If

    If Me.XPChkPayType(2).value = Checked Then

        With Me.FgCheques

            For i = .FixedRows To .Rows - 1
                '--------------------------------------------------------------------------
            Next i

        End With

    End If

    'Õ›Ÿ «·√›”«ÿ
    If Me.XPChkPayType(1).value = Checked Then
        If ChkInstall.value = vbChecked Then
            'Save installment Data
            Set RsTemp = New ADODB.Recordset
'            RsTemp.Open "InstallMent", Cn, adOpenStatic, adLockOptimistic, adCmdTable
            
                       
 strSQL = " SELECT       * FROM  dbo.InstallMent WHERE     (PartID = - 1)"
   RsTemp.Open strSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText

            RsTemp.AddNew
            RsTemp("PartID").value = CStr(new_id("InstallMent", "PartID", "", True))
            RsTemp("NoteID").value = note_id
            RsTemp("BasicAmmount").value = IIf(XPTxtValue(1).text = "", 0, val(XPTxtValue(1).text))
            RsTemp("InterestType").value = val(Me.LblPrecenType.Tag)
            RsTemp("InterestVal").value = val(LblPrecenValue.Caption)
            RsTemp("Total").value = val(LblInstallTotal.Caption)
            RsTemp("InstallCount").value = val(LblInstallCount.Caption)
            RsTemp("FirstInstallDate").value = CDate(Me.LblFirstInstallDate.Caption)

            If val(LblInstallmentType.Tag) = 0 Then
                RsTemp("InstallmentType").value = 0
            ElseIf val(LblInstallmentType.Tag) = 1 Then
                RsTemp("InstallmentType").value = 1
            ElseIf val(LblInstallmentType.Tag) = 2 Then
                RsTemp("InstallmentType").value = 2
            End If

            RsTemp("InstallSeprator").value = val(Me.LblInstallSeprator.Caption)
            RsTemp("StartValue").value = IIf(val(Me.LblStartValue.Caption) = 0, Null, val(Me.LblStartValue.Caption))
            RsTemp("CustID").value = IIf(DBCboClientName.BoundText = "", Null, val(DBCboClientName.BoundText))
            RsTemp("Type").value = 1
            RsTemp.update
            'save installment Details
            Set RsDetalis = New ADODB.Recordset
'            RsDetalis.Open "InstallMentDetails", Cn, adOpenStatic, adLockOptimistic, adCmdTable
           
 strSQL = " SELECT       * FROM  dbo.InstallMentDetails WHERE     (PartID = - 1)"
   RsDetalis.Open strSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText

            With Me.FgInstallments

                For RowNum = 1 To .Rows - 1
                    RsDetalis.AddNew
                    RsDetalis("QestID").value = CStr(new_id("InstallMentDetails", "QestID", "", True))
                    RsDetalis("PartID").value = RsTemp("PartID").value
                    RsDetalis("QeqtNum").value = IIf(.TextMatrix(RowNum, .ColIndex("Serial")) = "", "", .TextMatrix(RowNum, .ColIndex("Serial")))
                    RsDetalis("Value").value = IIf(.TextMatrix(RowNum, .ColIndex("Value")) = "", "", val(.TextMatrix(RowNum, .ColIndex("Value"))))
                    RsDetalis("DueDate").value = IIf(.TextMatrix(RowNum, .ColIndex("Due_Date")) = "", "", .TextMatrix(RowNum, .ColIndex("Due_Date")))
                    RsDetalis("Receipt").value = False
                    RsDetalis.update
                Next RowNum

            End With

        End If
    End If

    Dim LngDevID As Long
    Dim LngDevNO  As Integer
    Dim StrTempAccountCode As String
    Dim StrTempDes As String
    Dim SngTemp  As Variant
    Dim Account_Code_dynamic As String

    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
    '«·ÿ—› «·„œÌ‰
    SngTemp = (NewGrid.GetItemsTotal(ItemsGoodType) - val(LblDiscountsTotal.Caption)) * val(txt_Currency_rate.text)
    '    SngTemp = (Val(Me.lblTotal.Caption) * Val(txt_Currency_rate.text))
      
    'SngTemp =  (SngTemp, SystemOptions.SysDefCurrencyForamt)
    If SngTemp > 0 Then
        If detect_inventory_work_type = 1 Or detect_inventory_work_type = 2 Then

            Account_Code_dynamic = get_account_code_branch(4, my_branch)
        
            If Account_Code_dynamic = "NO branch" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                Else
                    MsgBox "Branch Not Created", vbCritical
                End If

                GoTo ErrTrap
            Else

                If Account_Code_dynamic = "NO account" Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox "·„ Ì „  ÕœÌœ Õ”«»  «·„‘ —Ì«  ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                    Else
                        MsgBox "Purchase  Account   Not Defined in this Branch", vbCritical
                    End If

                    GoTo ErrTrap
         
                End If
            End If

            If val(DCDocTypes.BoundText) > 0 Then
                getDocAccounts val(DCDocTypes.BoundText), StrTempAccountCode, , , , , usedaccount

                If StrTempAccountCode = "" And usedaccount = 1 Then
                    MsgBox "ÌÊÃœ Œÿ√ ›Ì «·Õ”«» «·„œÌ‰ ··›« Ê—…", vbCritical
                    GoTo ErrTrap
                ElseIf StrTempAccountCode <> "" And usedaccount = 1 Then
        
                ElseIf usedaccount = 0 Then
        
                    StrTempAccountCode = Account_Code_dynamic '«·„‘ —Ì« 
                End If

            Else
                StrTempAccountCode = Account_Code_dynamic '«·„‘ —Ì« 
            End If
            
            ' StrTempAccountCode = "a1a2a5" '„Œ“Ê‰ «·»÷«⁄…
            If SystemOptions.UserInterface = ArabicInterface Then
                StrTempDes = "›« Ê—… ‘—«¡ —ﬁ„ " & Me.TxtNoteSerial1.text & " " & TxtBillComment.text
            Else
                StrTempDes = "Purchase Invoice No: " & Me.TxtNoteSerial1.text & " " & TxtBillComment.text
            End If

            LngDevNO = LngDevNO + 1
    
            If txtManualNO.text <> "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrTempDes = StrTempDes & " ›« Ê—… „Ê—œ —ﬁ„  " & txtManualNO
                Else
                    StrTempDes = StrTempDes & " Supp Bill# " & txtManualNO
                End If
            
            End If

            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, 0, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, val(Me.XPTxtBillID.text), , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                GoTo ErrTrap
            End If
   
        ElseIf detect_inventory_work_type = 3 Then
            Dim groupAccount As String
             
            Dim line_value As Single

            With FG

                For i = 1 To FG.Rows - 1

                    If FG.TextMatrix(i, FG.ColIndex("Code")) <> "" Then
    
                        groupAccount = get_item_group_account_in_branch(FG.TextMatrix(i, FG.ColIndex("Code")), val(my_branch), 4)

                        'groupAccount = get_item_group_account_inventory(FG.TextMatrix(I, FG.ColIndex("Code")), DCboStoreName.BoundText, 4)
                        If groupAccount = "Error" Then
                            If SystemOptions.UserInterface = ArabicInterface Then
                                MsgBox "«·’‰› ›Ì «·”ÿ— —ﬁ„" & i & " €Ì— „Õœœ —ﬁ„ Õ”«»     «·„‘ —Ì«  ··„Œ“‰ «·„Õœœ   ·„Ã„Ê⁄ …"
                            Else
                                MsgBox "Item in line no " & i & "Group Purchase  Account Not Defined"
                            End If

                            GoTo ErrTrap
                        End If

                        line_value = FG.TextMatrix(i, FG.ColIndex("Price")) * val(txt_Currency_rate.text) * FG.TextMatrix(i, FG.ColIndex("Count"))

                        If SystemOptions.UserInterface = ArabicInterface Then
                            StrTempDes = "›« Ê—… ‘—«¡ —ﬁ„ " & Me.TxtNoteSerial1.text & " " & TxtBillComment.text
                        Else
                            StrTempDes = "Purchase Invoice No: " & Me.TxtNoteSerial1.text & " " & TxtBillComment.text
                        End If
                            
                        If txtManualNO.text <> "" Then
                            If SystemOptions.UserInterface = ArabicInterface Then
                                StrTempDes = StrTempDes & " ›« Ê—… „Ê—œ —ﬁ„  " & txtManualNO
                            Else
                                StrTempDes = StrTempDes & " Supp Bill# " & txtManualNO
                            End If
            
                        End If

                        LngDevNO = LngDevNO + 1

                        If ModAccounts.AddNewDev(LngDevID, LngDevNO, groupAccount, line_value, 0, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, val(Me.XPTxtBillID.text), , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                            GoTo ErrTrap
                        End If
    
                    End If

                Next i

            End With

        End If
    
        'Œ’„ ⁄·Ï „” ÊÏ «·”ÿ—
        If detect_inventory_work_type = 3 Then

            With FG

                For i = 1 To FG.Rows - 1
 
                    If FG.TextMatrix(i, FG.ColIndex("Code")) <> "" And val(FG.TextMatrix(i, FG.ColIndex("DiscountType"))) <> 1 Then
    
                        groupAccount = get_item_group_account_in_branch(FG.TextMatrix(i, FG.ColIndex("Code")), val(my_branch), 13)

                        If groupAccount = "Error" Then
                            If SystemOptions.UserInterface = ArabicInterface Then
                                MsgBox "«·’‰› ›Ì «·”ÿ— —ﬁ„" & i & " €Ì— „Õœœ —ﬁ„ Õ”«»  Œ’„ „ﬂ ”»  ·„Ã„Ê⁄ …"
                            Else
                                MsgBox "Item in line no " & i & "Group Name Discount Account Not Defined"
                            End If

                            GoTo ErrTrap
                        End If

                        line_value = (FG.TextMatrix(i, FG.ColIndex("Price")) * val(txt_Currency_rate.text) * FG.TextMatrix(i, FG.ColIndex("Count"))) - FG.TextMatrix(i, FG.ColIndex("Valu"))

                        If SystemOptions.UserInterface = ArabicInterface Then
                            StrTempDes = "›« Ê—… ‘—«¡ —ﬁ„ " & Me.TxtNoteSerial1.text & " " & TxtBillComment.text
                        Else
                            StrTempDes = "Purchase Invoice No: " & Me.TxtNoteSerial1.text & " " & TxtBillComment.text
                        End If

                        LngDevNO = LngDevNO + 1

                        If ModAccounts.AddNewDev(LngDevID, LngDevNO, groupAccount, line_value, 1, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, val(Me.XPTxtBillID.text), , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
                            GoTo ErrTrap
                        End If
    
                    End If

                Next i

            End With
    
        End If

    End If

    If XPChkTAX.value = vbChecked Then
        '   StrTempAccountCode = "a1a3a5" '÷—»Ì… „»Ì⁄«  „œÌ‰…
        '   SngTemp = Val(Me.lbl(25).Caption)
        '   StrTempDes = "›« Ê—… ‘—«¡ —ﬁ„ " & Me.TxtTransSerial.text
        '   LngDevNO = LngDevNO + 1
        '   If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, _
        '      0, StrTempDes, , , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Val(Me.XPTxtBillID.text)) = False Then
        '       GoTo ErrTrap
        '   End If
    End If

    If Me.ChkTaxAdd.value = vbChecked Then
        '   StrTempAccountCode = "a2a5a4" '÷—»Ì… √—»«Õ  Ã«—Ì… (Œ’„ Ê≈÷«›…
        '   StrTempDes = "›« Ê—… ‘—«¡ —ﬁ„ " & Me.TxtTransSerial.text
        '   SngTemp = Val(Me.lbl(32).Caption)
        '   LngDevNO = LngDevNO + 1
        '   If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, _
        '       0, StrTempDes, , , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Val(Me.XPTxtBillID.text)) = False Then
        '       GoTo ErrTrap
        '   End If
    End If

    '«·œ«∆‰
    If Me.XPChkPayType(0).value = vbChecked Then

        '«·Œ“Ì‰…
        If val(DCDocTypes.BoundText) > 0 Then
            getDocAccounts val(DCDocTypes.BoundText), , StrTempAccountCode, , , , , usedaccount

            If StrTempAccountCode = "" And usedaccount = 1 Then
                MsgBox "ÌÊÃœ Œÿ√ ›Ì «·Õ”«» «·œ«∆‰ ··›« Ê—…", vbCritical
                GoTo ErrTrap
            ElseIf StrTempAccountCode <> "" And usedaccount = 1 Then
        
            ElseIf usedaccount = 0 Then
        
                StrTempAccountCode = GetMyAccountCode("TblBoxesData", "BoxID", val(Me.DcboBox.BoundText))
            End If

        Else
            StrTempAccountCode = GetMyAccountCode("TblBoxesData", "BoxID", val(Me.DcboBox.BoundText))
        End If

        If SystemOptions.UserInterface = ArabicInterface Then
            StrTempDes = "›« Ê—… ‘—«¡ —ﬁ„ " & Me.TxtNoteSerial1.text & " " & TxtBillComment.text
        Else
            StrTempDes = "Purchase Invoice No: " & Me.TxtNoteSerial1.text & " " & TxtBillComment.text
        End If
    
        If txtManualNO.text <> "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                StrTempDes = StrTempDes & " ›« Ê—… „Ê—œ —ﬁ„  " & txtManualNO
            Else
                StrTempDes = StrTempDes & " Supp Bill# " & txtManualNO
            End If
            
        End If

        '  SngTemp = (Val(Me.XPTxtValue(0).text) * Val(txt_Currency_rate.text))
        '  SngTemp = (Val(Me.lblTotal.Caption) * Val(txt_Currency_rate.text))
        ' SngTemp = NewGrid.GetItemsTotal(ItemsGoodType) * Val(txt_Currency_rate.text)
        SngTemp = (NewGrid.GetItemsTotal(ItemsGoodType) - val(LblDiscountsTotal.Caption)) * val(txt_Currency_rate.text)
        '   SngTemp = Round(SngTemp, SystemOptions.SysDefCurrencyForamt)
    
        LngDevNO = LngDevNO + 1

        If Trim(TxtLcNo) <> "" Then
            StrTempAccountCode = GetMyAccountCode2("TblLC", "LCNO", TxtLcNo.text)
        End If
        
        If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, 1, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, val(Me.XPTxtBillID.text), , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
            GoTo ErrTrap
        End If
    End If

    If Me.XPChkPayType(1).value = vbChecked Then
    
        '«·√Ã·
        SngTemp = (NewGrid.GetItemsTotal(ItemsGoodType) - val(LblDiscountsTotal.Caption)) * val(txt_Currency_rate.text)

        If val(DCDocTypes.BoundText) > 0 Then
            getDocAccounts val(DCDocTypes.BoundText), , StrTempAccountCode, , , , , usedaccount

            If StrTempAccountCode = "" And usedaccount = 1 Then
                MsgBox "ÌÊÃœ Œÿ√ ›Ì «·Õ”«» «·œ«∆‰ ··›« Ê—…", vbCritical
                GoTo ErrTrap
            ElseIf StrTempAccountCode <> "" And usedaccount = 1 Then
        
            ElseIf usedaccount = 0 Then
        
                StrTempAccountCode = GetMyAccountCode("TblCustemers", "CusID", val(Me.DBCboClientName.BoundText))
            End If

        Else
            StrTempAccountCode = GetMyAccountCode("TblCustemers", "CusID", val(Me.DBCboClientName.BoundText))
        End If

        If SystemOptions.UserInterface = ArabicInterface Then
            StrTempDes = "›« Ê—… ‘—«¡ —ﬁ„ " & Me.TxtNoteSerial1.text & " " & TxtBillComment.text
        Else
            StrTempDes = "Purchase Invoice NO: " & Me.TxtNoteSerial1.text & " " & TxtBillComment.text
        End If

        LngDevNO = LngDevNO + 1
    
        If txtManualNO.text <> "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                StrTempDes = StrTempDes & " ›« Ê—… „Ê—œ —ﬁ„  " & txtManualNO
            Else
                StrTempDes = StrTempDes & " Supp Bill# " & txtManualNO
            End If
            
        End If

        If Trim(TxtLcNo) <> "" Then
            StrTempAccountCode = GetMyAccountCode2("TblLC", "LCNO", TxtLcNo.text)
        End If
        
        If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, 1, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, val(Me.XPTxtBillID.text), , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
            GoTo ErrTrap
        End If
    End If

    If Me.XPChkPayType(2).value = vbChecked Then
        SngTemp = NewGrid.GetItemsTotal(ItemsGoodType) * val(txt_Currency_rate.text)
  
        StrTempAccountCode = bank_account  '‘Ìﬂ«  „ƒÃ·…

        '    StrTempAccountCode = "a2a3a2" '√Ê—«ﬁ «·œ›⁄
        If SystemOptions.UserInterface = ArabicInterface Then
            StrTempDes = "⁄œœ " & Me.lbl(19).Caption & "  ‘Ìﬂ«  " & Chr(13)
            StrTempDes = StrTempDes & "›« Ê—… ‘—«¡ —ﬁ„ " & Me.TxtNoteSerial1.text
        Else
            StrTempDes = "Count " & Me.lbl(19).Caption & "  Cheque " & Chr(13)
            StrTempDes = StrTempDes & "Purchase Invoice No:" & Me.TxtNoteSerial1.text
    
        End If

        LngDevNO = LngDevNO + 1

        If Trim(TxtLcNo) <> "" Then
            StrTempAccountCode = GetMyAccountCode2("TblLC", "LCNO", TxtLcNo.text)
        End If
        
        If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, 1, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, val(Me.XPTxtBillID.text), , , , , , , , , , , , , , , , , val(Me.dcBranch.BoundText)) = False Then
            GoTo ErrTrap
        End If
    End If

    If val(Me.LblDiscountsTotal.Caption) > 0 Then
        Account_Code_dynamic = get_account_code_branch(13, my_branch)
        
        If Account_Code_dynamic = "NO branch" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
            Else
                MsgBox "Branch Not Created", vbCritical
    
            End If

            GoTo ErrTrap
        Else

            If Account_Code_dynamic = "NO account" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·„ Ì „  ÕœÌœ Õ”«»  «·Œ’„ «·„ﬂ ”» ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                Else
                    MsgBox " Discount Earned Account Not Defined in this Branch", vbCritical
                End If

                GoTo ErrTrap
         
            End If
        End If

        If val(DCDocTypes.BoundText) > 0 Then
            getDocAccounts val(DCDocTypes.BoundText), , , StrTempAccountCode, , , , , usedaccount

            If StrTempAccountCode = "" And usedaccount = 1 Then
                MsgBox "ÌÊÃœ Œÿ√ ›Ì «·Õ”«»  Œ’„ „ﬂ ”»", vbCritical
                GoTo ErrTrap
            ElseIf StrTempAccountCode <> "" And usedaccount = 1 Then
        
            ElseIf usedaccount = 0 Then
        
                StrTempAccountCode = Account_Code_dynamic '«·Œ’„ «·„ﬂ ”»
            End If

        Else
            StrTempAccountCode = Account_Code_dynamic '«·Œ’„ «·„ﬂ ”»
        End If
    
        'StrTempAccountCode = "a4a4" '«·Œ’„ «·„ﬂ ”»
        If SystemOptions.UserInterface = ArabicInterface Then
            StrTempDes = "›« Ê—… ‘—«¡ —ﬁ„ " & Me.TxtNoteSerial1.text & " " & TxtBillComment.text
        Else
            StrTempDes = "Purchase Invoice NO: " & Me.TxtNoteSerial1.text & " " & TxtBillComment.text
    
        End If

        ' LngDevNO = LngDevNO + 1

        ' If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Val(Me.LblDiscountsTotal.Caption) * Val(txt_Currency_rate.text), 1, StrTempDes, general_noteid, , , , Me.XPDtbBill.value, Me.DCboUserName.BoundText, Val(Me.XPTxtBillID.text), , , , , , , , , , , , , , , , , Val(Me.DcBranch.BoundText)) = False Then
        '     GoTo ErrTrap
        ' End If
    End If

    If Text1.text <> "" Then
        Cn.Execute "update Transactions set nots =' " & TxtTransSerial.text & "' where Transaction_Type= 20 and Transaction_Serial=" & Text1.text & ""
    End If

    Cn.Execute "update Transactions set NoteSerial =' " & Trim(Me.TxtNoteSerial.text) & "' where Transaction_ID=" & val(Me.XPTxtBillID.text)

    'Õ›Ÿ «·„’—Êﬁ«  «· ﬁœÌ—Ì…
    Dim FactoryExpenses As New ADODB.Recordset

    If Me.TxtModFlg.text = "E" Then
        strSQL = "Delete TblProductOrderFactoryexpenses where Transaction_ID=" & val(XPTxtBillID.text)
        Cn.Execute strSQL
    End If

    strSQL = "Select * from TblProductOrderFactoryexpenses where Transaction_ID=" & val(XPTxtBillID.text)
    FactoryExpenses.Open strSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    For RowNum = 1 To Fg_Journal.Rows - 2

        If Fg_Journal.TextMatrix(RowNum, Fg_Journal.ColIndex("AccountName")) <> "" Then
            FactoryExpenses.AddNew
            FactoryExpenses("Transaction_ID").value = val(XPTxtBillID.text)
         
            FactoryExpenses("Accountcode").value = Fg_Journal.TextMatrix(RowNum, Fg_Journal.ColIndex("Accountcode"))
            FactoryExpenses("AccountName").value = Fg_Journal.TextMatrix(RowNum, Fg_Journal.ColIndex("AccountName"))
            FactoryExpenses("value").value = val(Fg_Journal.TextMatrix(RowNum, Fg_Journal.ColIndex("value")))
            FactoryExpenses("des").value = Fg_Journal.TextMatrix(RowNum, Fg_Journal.ColIndex("des"))
            FactoryExpenses.update
        End If
         
    Next RowNum

    Cn.CommitTrans
    BeginTrans = False
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    CloseIssueVoucher

    If SystemOptions.autoReseiveVoucher = True Then
        CreateRecieveVouchers
    End If
 
    '----------------------------------------------------------------
    '·√‰‰« ﬁ„‰« »≈÷«›… Õ—ﬂ… „‰ ‰Ê⁄ „Œ ·›…
    strSQL = "SELECT * FROM Transactions WHERE Transaction_Type=" & BillType
         
    If SystemOptions.usertype <> UserAdminAll Then
        'StrSQL = StrSQL & " and Transaction_ID=0  AND   BranchId=" & branch_id
    End If

    Set rs = New ADODB.Recordset
    rs.Open strSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    Me.Retrive val(Me.XPTxtBillID.text)
    '----------------------------------------------------------------

    CuurentLogdata

    Select Case Me.TxtModFlg.text

        Case "N"

            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì…" & Chr(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
            Else
                Msg = " Data Was Saved do you want Another Entry" & Chr(13)
    
            End If
    
            If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                Cmd_Click (0)
                Screen.MousePointer = vbDefault
                Exit Sub
            End If
        
        Case "E"

            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Else
                MsgBox "Changes was Saved", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            End If

            lbl(67).Caption = showLabel(TxtNoteSerial1, oldtxtNoteSerial1)
    
    End Select

    close_order2 Me.Txt_order_no
    'Closeorders
    TxtModFlg.text = "R"
    'UpdateTransCost val(Me.XPTxtBillID.text)

    If SystemOptions.SysMainStockCostMethod = ModernWeightAverage Then
        '›Ï Õ«·… «‰  ﬂÊ‰ ÿ—Ìﬁ… Õ”«» „ Ê”ÿ «· ﬂ·›…
        'ÂÊ
        '    ModernWeightAverage
        '·«»œ «‰ ÌﬁÊ„ «·»—‰«„Ã » ⁄œÌ· ﬁÌ„… „ Ê”ÿ «· ﬂ·›… ··√’‰«›
        '«·„ÊÃÊœ… ›Ï «·›« Ê—…
    End If

    Screen.MousePointer = vbDefault
    Command2.Enabled = True
    Txt_EXport.Enabled = True
    'Grid.Visible = False
    Exit Sub
ErrTrap:

    'Stop
    'Resume
    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If

    Screen.MousePointer = vbDefault

    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
    Else
        Msg = "Sorry....Error During Saving" & Chr(13)
    End If

    Msg = Msg & Err.description & Chr(13)
    Msg = Msg & Err.Number & Chr(13)
    Msg = Msg & Err.Source
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
End Sub

Private Sub XPBtnNewClients_Click()

    With FrmAddNewCustemer
        '    .Tag = "x"
        .DealingForm = PurchaseTransaction
        Set .DcboCustomers = DBCboClientName
        .Caption = "≈÷«›… „Ê—œ ÃœÌœ"
        .lbl(1).Caption = "ﬂÊœ «·„Ê—œ"
        .lbl(0).Caption = "«”„ «·„Ê—œ"
        .AddType = 2
        .show vbModal
    End With

End Sub

Private Sub XPCboDiscountType_Change()
    XPCboDiscountType_Click
 
End Sub

Private Sub XPCboDiscountType_Click()
    On Error GoTo ErrTrap

    If XPCboDiscountType.ListIndex = 0 Or XPCboDiscountType.ListIndex = 3 Or XPCboDiscountType.ListIndex = -1 Then
        lbl(11).Enabled = False
        XPTxtDiscountVal.Enabled = True
        XPTxtDiscountVal.text = ""
    Else
        lbl(11).Enabled = True
        XPTxtDiscountVal.Enabled = True
        XPTxtDiscountVal.text = ""
    End If

    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
        If FG.TextMatrix(1, FG.ColIndex("Code")) <> "" Then
            NewGrid.Calculate 1
        End If
    End If

    Me.lbl(55).Visible = (Me.XPCboDiscountType.ListIndex = 2)

    Me.lbl(21).Visible = (Me.XPCboDiscountType.ListIndex = 2)

    If XPCboDiscountType.ListIndex = 0 Then
        ' lbl(8).Visible = False
        XPTxtDiscountVal.Visible = False
        lbl(11).Visible = False
    Else
        lbl(11).Visible = True
        XPTxtDiscountVal.Visible = True
        lbl(11).Visible = True
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub XPChkPayType_Click(Index As Integer)
    On Error GoTo ErrTrap
    Exit Sub

    Select Case Index

        Case 0

            If XPChkPayType(0).value = Checked Then
                If Me.TxtModFlg.text = "N" Then
                    XPTxtValue(0).text = ""
                    XPTxtSerial(0).text = ""
                End If

                If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
                    XPTxtValue(0).Enabled = True
                    '                XPTxtSerial(0).Enabled = True
                    XPTxtValue(0).locked = False
                    '                XPTxtSerial(0).Locked = False
                End If

            Else
                XPTxtValue(0).Enabled = False
                XPTxtValue(0).text = ""
                '            XPTxtSerial(0).Enabled = False
            End If

        Case 1

            If XPChkPayType(1).value = Checked Then
                If Me.TxtModFlg.text = "N" Then
                    XPTxtValue(1).text = ""
                    DtpDelayDate.value = Date
                End If

                If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
                    XPTxtValue(1).Enabled = True
                    XPTxtValue(1).locked = False
                    DtpDelayDate.Enabled = True
                Else
                    DtpDelayDate.Enabled = False
                End If

                Me.ChkInstall.Enabled = True
            Else
                XPTxtValue(1).Enabled = False
                XPTxtValue(1).text = ""
                Me.ChkInstall.Enabled = False
            End If

        Case 2

            If XPChkPayType(2).value = Checked And Me.TxtModFlg.text <> "R" Then
                Me.CmdCheque.Enabled = True
            Else
                Me.CmdCheque.Enabled = False
                Me.lbl(18).Caption = 0
                Me.lbl(19).Caption = 0
                Me.FgCheques.Rows = Me.FgCheques.FixedRows
            End If

    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub XPChkTAX_Click()
    On Error GoTo ErrTrap

    If XPChkTAX.value = Checked Then
        XPTxtTaxValue.Enabled = True
        lbl(22).Enabled = True
        lbl(45).Enabled = True
    Else
        XPTxtTaxValue.text = ""
        XPTxtTaxValue.Enabled = False
        lbl(22).Enabled = False
        lbl(45).Enabled = False
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub XPDtbBill_Change()

    If Trim(TxtNoteSerial1.text) <> "" Then
        oldtxtNoteSerial1.text = TxtNoteSerial1.text
    End If

    CurrentVoucherNo = ""
    TxtNoteSerial.text = ""
    TxtNoteSerial1.text = ""
    DateChanged = True
End Sub

Private Sub XPTab301_Click()
    On Error GoTo ErrTrap

    If Me.TxtModFlg.text <> "R" Then
    
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub printing()
    On Error GoTo ErrTrap

    Dim ShowType As Boolean
    ShowType = GetSetting(StrAppRegPath, "View_Type", "ReportType", True)

    If ShowType = True Then
        If Not XPTxtBillID.text Then
            Set BuyReport = New ClsBuyReport
            BuyReport.ShowBuyData XPTxtBillID.text, 1, True, LblTotal.Caption, txtManualNO.text
        End If

    Else

        If Not XPTxtBillID.text Then
            Set BuyReport = New ClsBuyReport
            BuyReport.ShowBuyDataShort XPTxtBillID.text
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Private Function AvailableDeal() As Boolean
    On Error GoTo ErrTrap
    Dim RowNum As Integer
    Dim Msg As String
    Dim strSQL As String
    Dim RsTemp As ADODB.Recordset
    Dim RsSalle As ADODB.Recordset
    Dim LngItemID As Long

    For RowNum = 1 To FG.Rows - 1

        If FG.TextMatrix(RowNum, FG.ColIndex("Code")) <> "" Then
            strSQL = "select * From QryDelPurchase where Transaction_Date >=" & SQLDate(XPDtbBill.value, True) & ""
            strSQL = strSQL + " and Item_ID=" & FG.TextMatrix(RowNum, FG.ColIndex("Code"))

            '        If FG.TextMatrix(RowNum, FG.ColIndex("HaveSerial")) <> "" Then
            '            If FG.TextMatrix(RowNum, FG.ColIndex("HaveSerial")) = True Then
            If FG.Cell(flexcpChecked, RowNum, FG.ColIndex("HaveSerial")) = flexChecked Then
                If FG.TextMatrix(RowNum, FG.ColIndex("Serial")) <> "" Then
                    strSQL = strSQL + " and ItemSerial='" & FG.TextMatrix(RowNum, FG.ColIndex("Serial")) & "'"
                End If

                '            End If
            End If

            Set RsSalle = New ADODB.Recordset
            RsSalle.Open strSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

            If Not (RsSalle.EOF Or RsSalle.BOF) Then
                If FG.Cell(flexcpChecked, RowNum, FG.ColIndex("HaveSerial")) = flexChecked Then

                    With FrmAlarm
                        .Tag = "x"
                        .DealingForm = PurchaseTransaction
                        .show vbModal
                    End With

                    AvailableDeal = False
                    Exit Function
                    '                End If
                    RsTemp.Close
                Else
                    Set RsTemp = New ADODB.Recordset
                    LngItemID = val(FG.TextMatrix(RowNum, FG.ColIndex("Code")))
                    Set RsTemp = GetItemQuantityStock(LngItemID, Me.DCboStoreName.BoundText, Me.XPDtbBill.value, val(Me.XPTxtBillID.text))

                    If Not (RsTemp.EOF Or RsTemp.BOF) Then
                        If val(RsTemp("QTY").value) < val(FG.TextMatrix(RowNum, FG.ColIndex("Count"))) Then

                            With FrmAlarm
                                .DealingForm = PurchaseTransaction
                                .show vbModal
                            End With

                            AvailableDeal = False
                            Exit Function
                        End If
                    End If

                    RsTemp.Close
                End If
            End If

            RsSalle.Close
        End If

    Next RowNum

    AvailableDeal = True
    Exit Function
ErrTrap:
End Function

Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)
    Dim IntResult As String
    Dim StrMSG As String

    On Error GoTo ErrTrap

    If Me.TxtModFlg.text = "" Then Exit Sub
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

Private Sub CboPayMentType_Change()
    On Error GoTo ErrTrap

    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
        If CboPayMentType.ListIndex = 0 Then
            XPChkPayType(0).Enabled = False
            XPChkPayType(1).Enabled = False
            XPChkPayType(2).Enabled = False
            XPChkPayType(0).value = Checked
            XPChkPayType(1).value = Unchecked
            XPChkPayType(2).value = Unchecked
            XPTxtValue(0).text = XPTxtSum.text
            XPTxtValue(1).text = 0
            '        DBCboClientName.Enabled = False
            DBCboClientName.text = ""
            DcboBox.Enabled = True
        Else
            XPChkPayType(0).Enabled = True
            XPChkPayType(1).Enabled = True
            XPChkPayType(2).Enabled = True
            XPChkPayType(0).value = Unchecked
            XPChkPayType(1).value = Unchecked
            XPChkPayType(2).value = Unchecked
            XPTxtValue(0).text = 0
            XPTxtValue(1).text = XPTxtSum.text
            '         DBCboClientName.Enabled = True
            DcboBox.Enabled = False
            DcboBox.text = ""
        End If
    End If

    Exit Sub
ErrTrap:

End Sub

Private Sub CboPayMentType_Click()
    CboPayMentType_Change
End Sub

Private Sub XPTxtDiscountVal_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.XPTxtDiscountVal.text, 0)
End Sub

Private Sub XPTxtSum_Change()
    On Error GoTo ErrTrap

    If CboPayMentType.ListIndex = 0 Then
        XPChkPayType(0).value = Checked
        XPTxtValue(0).text = XPTxtSum.text
    End If

    Me.LblTotal.Caption = XPTxtSum.text
    Exit Sub
ErrTrap:
End Sub

Public Function RepeatSerial(StrSerial As String, _
                             IntTransType As Integer, _
                             Optional IntTransID As Long = 0, _
                             Optional LngCusID As Long = 0) As Boolean

    Dim rs As ADODB.Recordset
    Dim strSQL As String
    Dim Msg As String
    RepeatSerial = False

    If SystemOptions.SysDataBaseType = SQLServerDataBase Then
        strSQL = "SELECT QryTransactionsTotal.Transaction_ID," & "QryTransactionsTotal.TransNet, QryTransactionsTotal.Transaction_Serial, " & "QryTransactionsTotal.Transaction_Date , QryTransactionsTotal.Transaction_Type," & "dbo.TblCustemers.CusName"
        strSQL = strSQL + " FROM dbo.QryTransactionsTotal() QryTransactionsTotal INNER JOIN " & "dbo.TblCustemers ON QryTransactionsTotal.CusID = dbo.TblCustemers.CusID"
        strSQL = strSQL + " Where QryTransactionsTotal.Transaction_Serial ='" & StrSerial & "'"
        strSQL = strSQL + " AND QryTransactionsTotal.Transaction_Type=" & BillType & ""

        If LngCusID <> 0 Then
            strSQL = strSQL + " AND dbo.TblCustemers.CusID=" & LngCusID & ""
        End If

        If IntTransID <> 0 Then
            strSQL = strSQL + " AND QryTransactionsTotal.Transaction_ID <> " & IntTransID & ""
        End If

        Set rs = New ADODB.Recordset
        rs.Open strSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (rs.BOF Or rs.EOF) Then
            Msg = "—ﬁ„ «·›« Ê—… „ÊÃÊœ „”»ﬁ« ›Ï «·»—‰«„Ã øø" & Chr(13)
            Msg = Msg + "„⁄·Ê„«  ⁄‰ «·›« Ê—… «·„”Ã·…:-" & Chr(13)
        
            Msg = Msg + "—ﬁ„ «·›« Ê—… ›Ï «·»—‰«„Ã:" & rs("Transaction_ID").value & Chr(13)
            Msg = Msg + "„”·”· «·›« Ê—…:" & rs("Transaction_Serial").value & Chr(13)
            Msg = Msg + " «—ÌŒ  ”ÃÌ· «·›« Ê—…:" & rs("Transaction_Date").value & Chr(13)
            Msg = Msg + "«”„ «·⁄„Ì· «Ê «·„Ê—œ:" & rs("CusName").value & Chr(13)
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            RepeatSerial = True
        End If

        rs.Close
        Set rs = Nothing

    End If

End Function

Private Sub SetDefaults()
    Dim StrTemp As String
    Dim RsTemp As ADODB.Recordset

    If SystemOptions.SysPurDateTakeType = InvDateFromLocalCompuer Then
        XPDtbBill.value = Date
    ElseIf SystemOptions.SysPurDateTakeType = InvDateFromServerComputer Then
        StrTemp = "select Getdate() as ServerDate"
        Set RsTemp = New ADODB.Recordset
        RsTemp.Open StrTemp, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (RsTemp.BOF Or RsTemp.EOF) Then
            If Not IsNull(RsTemp("ServerDate").value) Then
                XPDtbBill.value = Format(RsTemp("ServerDate").value, "yyyy/M/d")
            End If

            'XPDtbBill.Value = IIf(IsNull(RsTemp("ServerDate").Value), Date, (RsTemp("ServerDate").Value))
        End If

        RsTemp.Close
        Set RsTemp = Nothing
    End If

    If Not (rs.BOF Or rs.EOF) Then
        rs.MoveLast

        If SystemOptions.SysPurDateTakeType = InvDateFromLastInvDate Then
            XPDtbBill.value = IIf(IsNull(rs("Transaction_Date").value), Date, (rs("Transaction_Date").value))
        End If
    End If

    Me.DcboBox.BoundText = 1
End Sub
