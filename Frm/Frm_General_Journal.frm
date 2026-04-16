VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{85FD608E-54A8-11D4-8ED4-00E07D815373}#1.0#0"; "MBClrPkr.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Frm_General_Journal 
   Caption         =   "„—«Ã⁄… ﬁÌÊœ «·ÌÊ„Ì…"
   ClientHeight    =   7950
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12735
   HelpContextID   =   460
   Icon            =   "Frm_General_Journal.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   7950
   ScaleWidth      =   12735
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin C1SizerLibCtl.C1Elastic ElcMainContainer 
      Height          =   7950
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   12735
      _cx             =   22463
      _cy             =   14023
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
      BackColor       =   14737632
      ForeColor       =   -2147483630
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   5
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
      _GridInfo       =   $"Frm_General_Journal.frx":08CA
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin MSComctlLib.ImageList ImgLst 
         Left            =   4560
         Top             =   5730
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   12
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_General_Journal.frx":094D
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_General_Journal.frx":0CE7
               Key             =   "Doc"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_General_Journal.frx":1081
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_General_Journal.frx":111F
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_General_Journal.frx":11BC
               Key             =   "NoteDoc"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_General_Journal.frx":1556
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_General_Journal.frx":18F0
               Key             =   "IssuedUser"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_General_Journal.frx":1C8A
               Key             =   "PostedUser"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_General_Journal.frx":2024
               Key             =   "PostDate"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_General_Journal.frx":23BE
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_General_Journal.frx":2958
               Key             =   "Plus"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Frm_General_Journal.frx":2EF2
               Key             =   "Min"
            EndProperty
         EndProperty
      End
      Begin C1SizerLibCtl.C1Tab TabMain 
         Height          =   6465
         HelpContextID   =   460
         Left            =   30
         TabIndex        =   6
         Top             =   750
         Width           =   12675
         _cx             =   22357
         _cy             =   11404
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
         FrontTabColor   =   14737632
         BackTabColor    =   12648447
         TabOutlineColor =   -2147483632
         FrontTabForeColor=   -2147483630
         Caption         =   "«·ﬁÌÊœ «·„Õ——…|„·Œ’ ﬁÌ„ √—’œ… «·Õ”«»« "
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
         TabPicturePos   =   1
         CaptionEmpty    =   ""
         Separators      =   0   'False
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   37
         Picture(0)      =   "Frm_General_Journal.frx":348C
         Picture(1)      =   "Frm_General_Journal.frx":3826
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   6000
            Index           =   1
            Left            =   13320
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   45
            Width           =   12585
            _cx             =   22199
            _cy             =   10583
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
            BackColor       =   14737632
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
            Begin VSFlex8Ctl.VSFlexGrid FgAccountsValue 
               Height          =   5385
               HelpContextID   =   460
               Left            =   30
               TabIndex        =   10
               Top             =   45
               Width           =   12525
               _cx             =   22093
               _cy             =   9499
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
               SelectionMode   =   1
               GridLines       =   1
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   3
               Cols            =   7
               FixedRows       =   2
               FixedCols       =   0
               RowHeightMin    =   280
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"Frm_General_Journal.frx":3BC0
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
            Begin MSComCtl2.DTPicker DTPDev_From 
               Height          =   375
               Left            =   60
               TabIndex        =   20
               Top             =   5520
               Visible         =   0   'False
               Width           =   1980
               _ExtentX        =   3493
               _ExtentY        =   661
               _Version        =   393216
               CalendarBackColor=   12648447
               CalendarTitleBackColor=   10383715
               CheckBox        =   -1  'True
               CustomFormat    =   "yyyy/M/d"
               Format          =   75169795
               CurrentDate     =   37773
            End
            Begin MSComCtl2.DTPicker DTPDEV_TO 
               Height          =   345
               Left            =   2070
               TabIndex        =   21
               Top             =   5520
               Visible         =   0   'False
               Width           =   2430
               _ExtentX        =   4286
               _ExtentY        =   609
               _Version        =   393216
               CalendarBackColor=   12648447
               CalendarTitleBackColor=   10383715
               CheckBox        =   -1  'True
               CustomFormat    =   "yyyy/M/d"
               Format          =   75169795
               CurrentDate     =   37773
            End
            Begin ImpulseButton.ISButton CmdShowReport 
               Height          =   480
               Left            =   10035
               TabIndex        =   22
               Top             =   5460
               Width           =   2490
               _ExtentX        =   4392
               _ExtentY        =   847
               ButtonPositionImage=   1
               Caption         =   " ﬁ—Ì— «” «– „”«⁄œ"
               BackColor       =   14737632
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ButtonImage     =   "Frm_General_Journal.frx":3D04
               ColorButton     =   14737632
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   6000
            Index           =   0
            Left            =   45
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   45
            Width           =   12585
            _cx             =   22199
            _cy             =   10583
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
            BackColor       =   14737632
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
            Begin VSFlex8Ctl.VSFlexGrid grd_Journal 
               Height          =   5925
               HelpContextID   =   460
               Left            =   30
               TabIndex        =   9
               Top             =   30
               Width           =   12525
               _cx             =   22093
               _cy             =   10451
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
               AllowSelection  =   0   'False
               AllowBigSelection=   -1  'True
               AllowUserResizing=   0
               SelectionMode   =   0
               GridLines       =   1
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   2
               Cols            =   11
               FixedRows       =   2
               FixedCols       =   0
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"Frm_General_Journal.frx":409E
               ScrollTrack     =   0   'False
               ScrollBars      =   3
               ScrollTips      =   0   'False
               MergeCells      =   5
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
               CellButtonPicture=   "Frm_General_Journal.frx":4209
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
      Begin C1SizerLibCtl.C1Elastic EleFooter 
         Height          =   690
         Left            =   6375
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   7230
         Width           =   6330
         _cx             =   11165
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
         Appearance      =   4
         MousePointer    =   0
         Version         =   801
         BackColor       =   14737632
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
         Begin ImpulseButton.ISButton CmdSearch 
            Height          =   405
            Left            =   4440
            TabIndex        =   24
            Top             =   180
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   714
            Caption         =   "»ÕÀ ⁄‰ «·ﬁÌÊœ"
            BackColor       =   14737632
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColorButton     =   14737632
         End
         Begin MSDataListLib.DataCombo DcboUsers 
            Height          =   315
            Left            =   150
            TabIndex        =   4
            Top             =   225
            Width           =   2835
            _ExtentX        =   5001
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            MatchEntry      =   -1  'True
            Style           =   2
            BackColor       =   -2147483624
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            Caption         =   "«·„” Œœ„ «·ﬁ«∆„ Ì«· —ÕÌ·"
            Height          =   420
            Index           =   3
            Left            =   3045
            RightToLeft     =   -1  'True
            TabIndex        =   5
            Top             =   165
            Width           =   1365
         End
      End
      Begin C1SizerLibCtl.C1Elastic EleHeader 
         Height          =   705
         Left            =   0
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   0
         Width           =   12735
         _cx             =   22463
         _cy             =   1244
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial (Arabic)"
            Size            =   24
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
         Picture         =   "Frm_General_Journal.frx":45A3
         Caption         =   "„—«Ã⁄… ﬁÌÊœ «·ÌÊ„Ì…"
         Align           =   1
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
         PicturePos      =   1
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
            BackColor       =   &H000000FF&
            Height          =   285
            Left            =   810
            RightToLeft     =   -1  'True
            TabIndex        =   2
            Top             =   90
            Visible         =   0   'False
            Width           =   555
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   690
         Index           =   2
         Left            =   3210
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   7230
         Visible         =   0   'False
         Width           =   3150
         _cx             =   5556
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
         Appearance      =   4
         MousePointer    =   0
         Version         =   801
         BackColor       =   14871017
         ForeColor       =   -2147483630
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Picture         =   "Frm_General_Journal.frx":4E7D
         Caption         =   "ÿ—Ìﬁ… «·⁄—÷"
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
         Begin VB.OptionButton OptViewStyle 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ÿ«„ 2"
            Height          =   315
            Index           =   1
            Left            =   60
            RightToLeft     =   -1  'True
            TabIndex        =   13
            Top             =   330
            Width           =   1125
         End
         Begin VB.OptionButton OptViewStyle 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ÿ«„ 1"
            Height          =   315
            Index           =   0
            Left            =   1245
            RightToLeft     =   -1  'True
            TabIndex        =   12
            Top             =   330
            Value           =   -1  'True
            Width           =   1260
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   690
         Index           =   3
         Left            =   30
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   7230
         Width           =   6330
         _cx             =   11165
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
         Begin VB.CheckBox ChkPlanPaper 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " ŒÿÌÿ «·’›Õ…"
            Height          =   195
            Left            =   975
            RightToLeft     =   -1  'True
            TabIndex        =   23
            Top             =   420
            Width           =   1920
         End
         Begin VB.CheckBox ChkColor 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "·Ê‰ „„Ì“"
            Height          =   270
            Index           =   0
            Left            =   4050
            RightToLeft     =   -1  'True
            TabIndex        =   17
            Top             =   45
            Width           =   2220
         End
         Begin VB.CheckBox ChkColor 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "·Ê‰ «· ⁄·Ìﬁ"
            Height          =   270
            Index           =   1
            Left            =   4050
            RightToLeft     =   -1  'True
            TabIndex        =   16
            Top             =   375
            Value           =   1  'Checked
            Width           =   2220
         End
         Begin VB.CheckBox ChkExpand 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   345
            Left            =   30
            RightToLeft     =   -1  'True
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   30
            Width           =   825
         End
         Begin MBColorPicker.ColorPicker CPic 
            Height          =   315
            Index           =   0
            Left            =   2910
            TabIndex        =   18
            ToolTipText     =   "·Ê‰ Œ·›Ì… «·ﬁÌÊœ"
            Top             =   45
            Width           =   1125
            _ExtentX        =   1773
            _ExtentY        =   556
            CustomButtonText=   " Œ’Ì’"
            Color           =   12648447
            BackColor       =   14737632
            Style           =   2
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
            NumColors       =   64
            Color1          =   0
            Color2          =   128
            Color3          =   32768
            Color4          =   32896
            Color5          =   8388608
            Color6          =   8388736
            Color7          =   8421376
            Color8          =   12632256
            Color9          =   8421504
            Color10         =   255
            Color11         =   65280
            Color12         =   65535
            Color13         =   16711680
            Color14         =   16711935
            Color15         =   16776960
            Color18         =   12632319
            Color19         =   12640511
            Color20         =   12648447
            Color21         =   12648384
            Color22         =   16777152
            Color23         =   16761024
            Color24         =   16761087
            Color25         =   14737632
            Color26         =   8421631
            Color27         =   8438015
            Color28         =   8454143
            Color29         =   8454016
            Color30         =   16777088
            Color31         =   16744576
            Color32         =   16744703
            Color33         =   12632256
            Color34         =   255
            Color35         =   33023
            Color36         =   65535
            Color37         =   65280
            Color38         =   16776960
            Color39         =   16711680
            Color40         =   16711935
            Color41         =   8421504
            Color42         =   192
            Color43         =   16576
            Color44         =   49344
            Color45         =   49152
            Color46         =   12632064
            Color47         =   12582912
            Color48         =   12583104
            Color49         =   4210752
            Color50         =   128
            Color51         =   16512
            Color52         =   32896
            Color53         =   32768
            Color54         =   8421376
            Color55         =   8388608
            Color56         =   8388736
            Color57         =   0
            Color58         =   64
            Color59         =   4210816
            Color60         =   16448
            Color61         =   16384
            Color62         =   4210688
            Color63         =   4194304
            Color64         =   4194368
         End
         Begin MBColorPicker.ColorPicker CPic 
            Height          =   315
            Index           =   1
            Left            =   2910
            TabIndex        =   19
            ToolTipText     =   "·Ê‰ Œ·›Ì… «·ﬁÌÊœ"
            Top             =   345
            Width           =   1125
            _ExtentX        =   1773
            _ExtentY        =   556
            CustomButtonText=   " Œ’Ì’"
            Color           =   14737632
            BackColor       =   14737632
            Style           =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumColors       =   64
            Color1          =   0
            Color2          =   128
            Color3          =   32768
            Color4          =   32896
            Color5          =   8388608
            Color6          =   8388736
            Color7          =   8421376
            Color8          =   12632256
            Color9          =   8421504
            Color10         =   255
            Color11         =   65280
            Color12         =   65535
            Color13         =   16711680
            Color14         =   16711935
            Color15         =   16776960
            Color18         =   12632319
            Color19         =   12640511
            Color20         =   12648447
            Color21         =   12648384
            Color22         =   16777152
            Color23         =   16761024
            Color24         =   16761087
            Color25         =   14737632
            Color26         =   8421631
            Color27         =   8438015
            Color28         =   8454143
            Color29         =   8454016
            Color30         =   16777088
            Color31         =   16744576
            Color32         =   16744703
            Color33         =   12632256
            Color34         =   255
            Color35         =   33023
            Color36         =   65535
            Color37         =   65280
            Color38         =   16776960
            Color39         =   16711680
            Color40         =   16711935
            Color41         =   8421504
            Color42         =   192
            Color43         =   16576
            Color44         =   49344
            Color45         =   49152
            Color46         =   12632064
            Color47         =   12582912
            Color48         =   12583104
            Color49         =   4210752
            Color50         =   128
            Color51         =   16512
            Color52         =   32896
            Color53         =   32768
            Color54         =   8421376
            Color55         =   8388608
            Color56         =   8388736
            Color57         =   0
            Color58         =   64
            Color59         =   4210816
            Color60         =   16448
            Color61         =   16384
            Color62         =   4210688
            Color63         =   4194304
            Color64         =   4194368
         End
         Begin VB.Image ImgNote 
            Height          =   240
            Left            =   1875
            Picture         =   "Frm_General_Journal.frx":5217
            Top             =   150
            Visible         =   0   'False
            Width           =   240
         End
      End
   End
End
Attribute VB_Name = "Frm_General_Journal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim nt As New frmFlexNote
Dim m_Cmd_New As Boolean, m_Cmd_Edit As Boolean
Dim m_Cmd_Save As Boolean, m_Cmd_Delete As Boolean
Dim m_Cmd_Search As Boolean, m_Cmd_Undo As Boolean

Private FrmComment As frmFlexNote
Dim TTP As New clstooltip

Private WithEvents m_oText    As TextBox
Attribute m_oText.VB_VarHelpID = -1

Private WithEvents m_InsertMenu As Menu
Attribute m_InsertMenu.VB_VarHelpID = -1

Private WithEvents m_EditMenu As Menu
Attribute m_EditMenu.VB_VarHelpID = -1

Private WithEvents m_DeleteMenu As Menu
Attribute m_DeleteMenu.VB_VarHelpID = -1
Dim RsLoadJournal   As New ADODB.Recordset
Dim IntCounter      As Long
Dim LngMouseRow As Long
Dim BolEditing As Boolean

Private Sub ChkColor_Click(Index As Integer)
    Me.CPic(Index).Enabled = CBool(ChkColor(Index).value)
    SetGrgColors Index, Not CBool(ChkColor(Index).value)
End Sub

Private Sub ChkExpand_Click()
    Dim VarTemp As Variant
    Dim BolHidden As Boolean
    Dim StrNewData As String
    Dim StrOldData As String
    Dim i As Long, j As Long

    BolHidden = Not CBool(ChkExpand.value)

    With grd_Journal

        For j = .FixedRows To .Rows - 1

            If left(CStr(.Cell(flexcpData, j, .ColIndex("Descrip"))), 3) = "Btn" Then
                .ComboList = ""
                StrOldData = CStr(.Cell(flexcpData, j, .ColIndex("Descrip")))
                VarTemp = Split(.Cell(flexcpData, j, .ColIndex("Descrip")), ";", , vbTextCompare)

                If UBound(VarTemp) > 0 Then
                    If BolHidden = False Then
                        .Cell(flexcpPicture, j, .ColIndex("Descrip")) = Me.ImgLst.ListImages("Min").ExtractIcon
                        StrNewData = Replace(StrOldData, "Plus", "Min", 1, -1, vbTextCompare)
                    Else
                        .Cell(flexcpPicture, j, .ColIndex("Descrip")) = Me.ImgLst.ListImages("Plus").ExtractIcon
                        StrNewData = Replace(StrOldData, "Min", "Plus", 1, -1, vbTextCompare)
                    End If

                    For i = val(VarTemp(2)) To val(VarTemp(3))
                        .RowHidden(i) = BolHidden
                    Next i

                    .Cell(flexcpData, j, .ColIndex("Descrip")) = StrNewData
                End If
            End If

        Next j

        '    If SystemOptions.UserPlaySound = True Then
        '        PlaySoundEffect CollapseNode
        '    End If
        grd_Journal.SetFocus
    End With

End Sub

Private Sub ChkPlanPaper_Click()

    If ChkPlanPaper.value = 1 Then
        grd_Journal.GridLines = 1
    Else
        grd_Journal.GridLines = flexGridNone
    End If

End Sub

Private Sub CmdSearch_Click()
    FgAccountsValue.Clear flexClearScrollable, flexClearEverything
    FgAccountsValue.Rows = FgAccountsValue.FixedRows + 1
    grd_Journal.Clear flexClearScrollable, flexClearEverything
    grd_Journal.Rows = grd_Journal.FixedRows + 1
    Frm_JournalSearch.show vbModal
End Sub

Private Sub CmdShowReport_Click()
    Dim cAccountReport As ClsAccReports

    With Me.FgAccountsValue

        If .Row < .FixedRows Then
            GetMsgs 192, vbExclamation
            Exit Sub
        End If

        '    If .IsSelected(.Row) = False Then
        '        GetMsgs 192, vbExclamation
        '        Exit Sub
        '    End If
        If .TextMatrix(.Row, .ColIndex("Account_code")) = "" Then
            'GetMsgs 192, vbExclamation
            'Exit Sub
        End If

        Set cAccountReport = New ClsAccReports
        cAccountReport.BegineDate = Me.DTPDev_From.value
    
        cAccountReport.EndDate = Me.DTPDEV_TO.value
        cAccountReport.ShowLedger "a1a2a1a2", "«·Õ“Ì‰…"
        'cAccountReport.ShowLedger .TextMatrix(.Row, .ColIndex("Account_code")), _
         .TextMatrix(.Row, .ColIndex("Account_Name"))
    End With

    Set cAccountReport = Nothing
End Sub

Private Sub CPic_Change(Index As Integer, _
                        ByVal NewColor As stdole.OLE_COLOR)
    SetGrgColors Index, Not CBool(ChkColor(Index).value)
End Sub

Private Sub ElcMainContainer_MouseDown(Button As Integer, _
                                       Shift As Integer, _
                                       X As Single, _
                                       Y As Single)
    ReleaseCapture
    SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0
End Sub

Private Sub EleHeader_DblClick()
    Me.WindowState = IIf(Me.WindowState = vbMaximized, vbNormal, vbMaximized)
End Sub

Private Sub Form_Load()
    Dim My_SQL As String
    Dim BolRtl As Boolean
    Dim Dcmbos As New ClsDataCombos
    Dim ClsGrdBck As New ClsBackGroundPic
    Dim i As Integer

    If SystemOptions.UserInterface = EnglishInterface Then
        BolRtl = False
        SetInterface Me
        ChangeLang
    Else
        BolRtl = True
    End If

    'Set m_InsertMenu = MDIFrmamin.MnuInsertCom
    'Set m_DeleteMenu = MDIFrmamin.MnuDelCom
    'Set m_EditMenu = MDIFrmamin.MnuEditCom
    With Me.grd_Journal
        .RowHeightMin = 280
        .MergeCells = flexMergeFixedOnly
        .Highlight = flexHighlightWithFocus
        .GridLines = flexGridNone
        .Editable = flexEDKbdMouse
        .OwnerDraw = flexODOver
        .MergeCol(.ColIndex("Post")) = True
        .ColHidden(.ColIndex("DevPost")) = True
        .ColHidden(.ColIndex("AccountCode")) = True
        .ColHidden(.ColIndex("NoteID")) = True
        .ColHidden(.ColIndex("DEV_ID")) = True
        .ColHidden(.ColIndex("DEV_LineNo")) = True
    
        .MergeCol(.ColIndex("DEV_NO")) = True

        If BolRtl = True Then
            .Cell(flexcpText, 0, .ColIndex("DEV_NO"), 1, .ColIndex("Debit")) = "—ﬁ„ «·ﬁÌœ"
        Else
            .Cell(flexcpText, 0, .ColIndex("DEV_NO"), 1, .ColIndex("Debit")) = "DEV NO."
        End If

        .MergeCol(.ColIndex("Debit")) = True

        If BolRtl = True Then
            .Cell(flexcpText, 0, .ColIndex("Debit"), 1, .ColIndex("Debit")) = "„‹‹œÌ‹‹‰"
        Else
            .Cell(flexcpText, 0, .ColIndex("Debit"), 1, .ColIndex("Debit")) = "Debit"
        End If

        .ColWidth(.ColIndex("Debit")) = 1000
    
        .MergeCol(.ColIndex("Credit")) = True

        If BolRtl = True Then
            .Cell(flexcpText, 0, .ColIndex("Credit"), 1, .ColIndex("Credit")) = "œ«∆‹‹‹‹‹‰"
        Else
            .Cell(flexcpText, 0, .ColIndex("Credit"), 1, .ColIndex("Credit")) = "Credit"
        End If

        .ColWidth(.ColIndex("Credit")) = 1000
    
        .MergeCol(.ColIndex("Descrip")) = True

        If BolRtl = True Then
            .Cell(flexcpText, 0, .ColIndex("Descrip"), 1, .ColIndex("Descrip")) = "»‹‹Ì‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹‹«‰"
        Else
            .Cell(flexcpText, 0, .ColIndex("Descrip"), 1, .ColIndex("Descrip")) = "Description"
        End If

        .ColWidth(.ColIndex("Descrip")) = 6000
        .MergeCol(.ColIndex("Note_Date")) = True

        If BolRtl = True Then
            .Cell(flexcpText, 0, .ColIndex("Note_Date"), 1, .ColIndex("Note_Date")) = " ‹«—ÌŒ «·‹ﬁ‹Ì‹œ "
            .Cell(flexcpText, 0, .ColIndex("Post"), 1, .ColIndex("Post")) = "Õ‹‹«·‹‹… «·‹ﬁ‹‹Ì‹œ"
        Else
            .Cell(flexcpText, 0, .ColIndex("Note_Date"), 1, .ColIndex("Note_Date")) = "Issued Date"
            .Cell(flexcpText, 0, .ColIndex("Post"), 1, .ColIndex("Post")) = "Posting Stat"
        End If

        Set .WallPaper = ClsGrdBck.Picture
        '    For I = 0 To .Cols - 1
        '        .ColHidden(I) = False
        '    Next I
    End With

    'Set the Accounts grid
    With FgAccountsValue
        .MergeCells = flexMergeFixedOnly
        .Highlight = flexHighlightWithFocus
        .Editable = flexEDNone
        .MergeCol(.ColIndex("Account_Serial")) = True
        .MergeCol(.ColIndex("Account_Name")) = True
        .MergeRow(0) = True

        If BolRtl = True Then
            .Cell(flexcpText, 0, .ColIndex("Account_Serial"), 1, .ColIndex("Account_Serial")) = "ﬂÊœ «·Õ”«»"
            .Cell(flexcpText, 0, .ColIndex("Account_Name"), 1, .ColIndex("Account_Name")) = "«”„ «·Õ”«»"
            .Cell(flexcpText, 0, .ColIndex("Debit_Posted"), 0, .ColIndex("Credit_Posted")) = "„—Õ·"
            .Cell(flexcpText, 1, .ColIndex("Debit_Posted"), 1, .ColIndex("Debit_Posted")) = "„œÌ‰"
            .Cell(flexcpText, 1, .ColIndex("Credit_Posted"), 1, .ColIndex("Credit_Posted")) = "œ«∆‰"
        
            .Cell(flexcpText, 0, .ColIndex("Debit_NotPosted"), 0, .ColIndex("Credit_NotPosted")) = "€Ì— „—Õ·"
            .Cell(flexcpText, 1, .ColIndex("Debit_NotPosted"), 1, .ColIndex("Debit_NotPosted")) = "„œÌ‰"
            .Cell(flexcpText, 1, .ColIndex("Credit_NotPosted"), 1, .ColIndex("Credit_NotPosted")) = "œ«∆‰"
        Else
            .Cell(flexcpText, 0, .ColIndex("Account_Serial"), 1, .ColIndex("Account_Serial")) = "Account Serial"
            .Cell(flexcpText, 0, .ColIndex("Account_Name"), 1, .ColIndex("Account_Name")) = "Account Name"
        
            .Cell(flexcpText, 0, .ColIndex("Debit_Posted"), 0, .ColIndex("Credit_Posted")) = "Posted"
            .Cell(flexcpText, 1, .ColIndex("Debit_Posted"), 1, .ColIndex("Debit_Posted")) = "Debit"
            .Cell(flexcpText, 1, .ColIndex("Credit_Posted"), 1, .ColIndex("Credit_Posted")) = "Credit"
        
            .Cell(flexcpText, 0, .ColIndex("Debit_NotPosted"), 0, .ColIndex("Credit_NotPosted")) = "Not Posted"
            .Cell(flexcpText, 1, .ColIndex("Debit_NotPosted"), 1, .ColIndex("Debit_NotPosted")) = "Debit"
            .Cell(flexcpText, 1, .ColIndex("Credit_NotPosted"), 1, .ColIndex("Credit_NotPosted")) = "Credit"
        End If

        Set .WallPaper = ClsGrdBck.Picture

        '    .AutoSize 0, .Cols - 1, False
        For i = 0 To .Cols - 1
            '.ColHidden(I) = False
        Next i

    End With

    Dcmbos.GetUsers Me.DcboUsers
    Me.DcboUsers.BoundText = user_id

    With Me.ChkExpand
        Set .Picture = ImgLst.ListImages("Plus").ExtractIcon
        Set .DownPicture = ImgLst.ListImages("Min").ExtractIcon
    End With

    SetDtpickerDate Me.DTPDev_From
    SetDtpickerDate Me.DTPDEV_TO
    Me.TabMain.CurrTab = 0
    Resize_Form Me, ReportSize
    Me.TxtModFlg.text = "R"
    Set Dcmbos = Nothing
    Set ClsGrdBck = Nothing
    AddTip
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)
    Dim IntRes As Integer
    'If Me.TxtModFlg.text = "R" Then
    '    Exit Sub
    'End If
    'IntRes = QueryCloseMsg(Me.TxtModFlg.text, "ﬁÌÊœ «·ÌÊ„Ì…")
    'Select Case IntRes
    '    Case vbYes
    '        Cancel = True
    '        Do_Action Do_save
    '    Case vbNo
    '        Cancel = False
    '        Application_Mode "R"
    '    Case vbCancel
    '        Cancel = True
    'End Select
End Sub

Private Sub Form_Resize()
    Dim SngWith As Single
    Dim i As Integer
    Dim SngTemp As Single

    With FgAccountsValue
        .AutoSize 0, .Cols - 1, False
        SngTemp = .ColWidth(.ColIndex("Account_Serial")) + .ColWidth(.ColIndex("Account_Name"))
        SngWith = Fix((.ClientWidth - SngTemp) / 4)
        .ColWidth(.ColIndex("Debit_Posted")) = SngWith
        .ColWidth(.ColIndex("Credit_Posted")) = SngWith
        .ColWidth(.ColIndex("Debit_NotPosted")) = SngWith
        .ColWidth(.ColIndex("Credit_NotPosted")) = SngWith
        '    For i = 0 To .Cols - 1
        '        If .ColHidden(i) = False Then
        '            .ColWidth(i) = SngWith
        '        End If
        '    Next i
    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)

    'If (frmFlexNote Is Nothing) Then
    'Else
    '    Unload nt
    '    Set frmFlexNote = Nothing
    'End If
    If Loaded("frmFlexNote") Then
        'Unload frmFlexNote
    End If

    Set m_oText = Nothing
    'If Not FrmComment Is Nothing Then
    '    'Unload FrmComment
    'End If
    'Set FrmComment = Nothing
    Set m_InsertMenu = Nothing
    Set m_EditMenu = Nothing
    Set m_DeleteMenu = Nothing
    Set TTP = Nothing
End Sub

Private Sub grd_Journal_AfterEdit(ByVal Row As Long, _
                                  ByVal Col As Long)
    Dim IntLoop As Integer
    Dim BolStop As Boolean
    Dim StrNoteID As String
    Dim StrPosted As String
    Dim StrNotPosted As String

    If SystemOptions.UserInterface = ArabicInterface Then
        StrPosted = "„‹—Õ‹·"
        StrNotPosted = " ‹—Õ‹Ì‹·"
    Else
        StrPosted = "Posted"
        StrNotPosted = "Post"
    End If

    With grd_Journal

        Select Case Col

            Case .ColIndex("Post")

                If .Cell(flexcpChecked, Row, Col) = flexChecked Then
                    .Cell(flexcpForeColor, Row, Col) = vbRed
                    .Cell(flexcpText, Row, Col) = StrPosted
                    IntLoop = Row - 1
                    StrNoteID = .TextMatrix(IntLoop, .ColIndex("NoteID"))

                    If left(StrNoteID, 2) <> "M-" Then
                        'Put Flag
                        .TextMatrix(Row, .ColIndex("NoteID")) = "M-" & StrNoteID
                    End If

                    Do

                        If .TextMatrix(IntLoop, .ColIndex("Post")) = "" Then
                            .Cell(flexcpChecked, IntLoop, .ColIndex("DevPost")) = flexChecked
                            IntLoop = IntLoop - 1
                        Else
                            BolStop = True
                        End If

                    Loop While BolStop <> True

                Else
                    .Cell(flexcpForeColor, Row, Col) = vbBlack
                    .Cell(flexcpText, Row, Col) = StrNotPosted
                    IntLoop = Row - 1
                    StrNoteID = .TextMatrix(IntLoop, .ColIndex("NoteID"))

                    If left(StrNoteID, 2) <> "M-" Then
                        'Put Flag
                        .TextMatrix(Row, .ColIndex("NoteID")) = "M-" & StrNoteID
                    End If

                    Do

                        If .TextMatrix(IntLoop, .ColIndex("Post")) = "" Then
                            .Cell(flexcpChecked, IntLoop, .ColIndex("DevPost")) = flexUnchecked
                            IntLoop = IntLoop - 1
                        Else
                            BolStop = True
                        End If

                    Loop While BolStop <> True

                End If

        End Select

    End With

    'MsgBox grd_Journal.Cell(flexcpChecked, Row, Col)

End Sub

Private Sub grd_Journal_AfterRowColChange(ByVal OldRow As Long, _
                                          ByVal OldCol As Long, _
                                          ByVal NewRow As Long, _
                                          ByVal NewCol As Long)
    Dim VarTemp As Variant

    If SystemOptions.UserInterface = ArabicInterface Then

        With grd_Journal

            If OldCol = .ColIndex("Descrip") Then
                On Error Resume Next

                If .Cell(flexcpData, OldRow, OldCol) <> Empty Then
                    VarTemp = Split(.Cell(flexcpData, OldRow, OldCol), ";", , vbTextCompare)

                    If VarTemp(0) = "Btn" Then
                        If VarTemp(1) = "Min" Then
                            .Cell(flexcpPicture, OldRow, OldCol) = Me.ImgLst.ListImages("Min").ExtractIcon
                        ElseIf VarTemp(1) = "Plus" Then
                            .Cell(flexcpPicture, OldRow, OldCol) = Me.ImgLst.ListImages("Plus").ExtractIcon
                        End If
                    End If
                End If
            End If

        End With

    End If

End Sub

Private Sub grd_Journal_BeforeEdit(ByVal Row As Long, _
                                   ByVal Col As Long, _
                                   Cancel As Boolean)
    Dim VarTemp As Variant

    With grd_Journal

        Select Case Col

            Case .ColIndex("Post")
                .ComboList = ""

                If Me.TxtModFlg.text = "E" Then
                    If .TextMatrix(Row, Col) = "" Then
                        Cancel = True
                    Else
                        Cancel = False
                    End If

                Else
                    Cancel = True
                End If

            Case .ColIndex("Descrip")

                If .Cell(flexcpData, Row, Col) <> Empty Then
                    VarTemp = Split(.Cell(flexcpData, Row, Col), ";", , vbTextCompare)

                    If UBound(VarTemp) > 0 Then
                        If VarTemp(0) = "Btn" Then
                            .ComboList = "..."
                            Cancel = False
                        Else
                            .ComboList = ""
                            Cancel = True
                        End If

                        If SystemOptions.UserInterface = ArabicInterface Then
                            .Cell(flexcpPicture, Row, Col) = Nothing
                        End If

                        If VarTemp(1) = "Min" Then
                            .CellButtonPicture = Me.ImgLst.ListImages("Min").ExtractIcon
                        ElseIf VarTemp(1) = "Plus" Then
                            .CellButtonPicture = Me.ImgLst.ListImages("Plus").ExtractIcon
                        End If
                    End If

                Else
                    .ComboList = ""
                    Cancel = True
                End If

            Case Else
                .ComboList = ""
                Cancel = True
            
        End Select

    End With

End Sub

Private Sub grd_Journal_CellButtonClick(ByVal Row As Long, _
                                        ByVal Col As Long)

    '"Btn;Plus;LngBegRow;LngEndRow"
    Dim VarTemp As Variant
    Dim BolHidden As Boolean
    Dim StrNewData As String
    Dim StrOldData As String
    Dim i As Integer

    With grd_Journal

        If left(CStr(.Cell(flexcpData, Row, Col)), 3) = "Btn" Then
            .ComboList = ""
            StrOldData = CStr(.Cell(flexcpData, Row, Col))
            VarTemp = Split(.Cell(flexcpData, Row, Col), ";", , vbTextCompare)

            If VarTemp(1) = "Plus" Then
                BolHidden = False
                .Cell(flexcpPicture, Row, Col) = Me.ImgLst.ListImages("Min").ExtractIcon
                StrNewData = Replace(StrOldData, "Plus", "Min", 1, -1, vbTextCompare)
            Else
                BolHidden = True
                .Cell(flexcpPicture, Row, Col) = Me.ImgLst.ListImages("Plus").ExtractIcon
                StrNewData = Replace(StrOldData, "Min", "Plus", 1, -1, vbTextCompare)
            End If

            For i = val(VarTemp(2)) To val(VarTemp(3))
                .RowHidden(i) = BolHidden
            Next i

            '        If SystemOptions.UserPlaySound = True Then
            '            PlaySoundEffect CollapseNode
            '        End If
            .Cell(flexcpData, Row, Col) = StrNewData
            grd_Journal.SetFocus
            .ShowCell val(VarTemp(3)), Col
            grd_Journal.Row = Row
            grd_Journal.Col = Col
        End If

    End With

End Sub

Private Sub grd_Journal_CellChanged(ByVal Row As Long, _
                                    ByVal Col As Long)
    Dim VarTemp As Variant

    With grd_Journal

        If .Cell(flexcpData, Row, Col) <> Empty Then
            VarTemp = Split(.Cell(flexcpData, Row, Col), ";", , vbTextCompare)

            If VarTemp(0) = "Btn" Then
                .ComboList = "..."
            Else
                .ComboList = ""
            End If
        End If

    End With

End Sub

Private Sub grd_Journal_Click()

    'Static lNoteRow&, lNoteCol&, r&, C&
    '
    '    ' clicking? no work
    '    'If Button <> 0 Then Exit Sub
    '
    '    ' get mouse coordinates
    '    r = grd_Journal.Row
    '    C = grd_Journal.Col
    '    If grd_Journal.ColKey(C) <> "Descrip" Then
    '        Exit Sub
    '    End If
    '    If grd_Journal.TextMatrix(r, C) = "" Then
    '        Exit Sub
    '    End If
    '    ' same cell or neighbour? no work
    '    If r = lNoteRow And C = lNoteCol Then Exit Sub
    '    If r = lNoteRow And C = lNoteCol + 1 Then Exit Sub
    '
    '    ' other cell, hide current note, if any
    '    If lNoteRow >= 0 And lNoteCol >= 0 Then
    '        grd_Journal.SetFocus
    '        lNoteRow = -1
    '        lNoteCol = -1
    '    End If
    '
    '    ' no note to show? then bail out
    '    If r <= 0 Or C <= 0 Then Exit Sub
    '    Load nt
    '    If TypeName(grd_Journal.Cell(flexcpData, r, C)) <> "String" Then
    '        nt.txtNote.Text = "Ì„ﬂ‰ﬂ ﬂ «»…  ⁄·Ìﬁ Â‰«...∫" & vbCrLf & "=================="
    '    Else
    '        ' show new note
    '        nt.txtNote = grd_Journal.Cell(flexcpData, r, C)
    '    End If
    '    ' show new note
    '    nt.ShowNote grd_Journal, r, C
    '    ' save coordinates for next time
    '    lNoteRow = r
    '    lNoteCol = C
    'MsgBox grd_Journal.ColKey(grd_Journal.Col)
End Sub

Private Sub grd_Journal_DblClick()
    'fg.Cell(flexcpData, r, c) = "** New Note **" & vbCrLf
    'fg.Cell(flexcpPicture, r, c) = imgNote
    'fg.Cell(flexcpPictureAlignment, r, c) = flexPicAlignRightTop

End Sub

Public Property Get Cmd_Search() As Boolean
    Cmd_Search = m_Cmd_Search

End Property

Public Property Let Cmd_Search(ByVal vNewValue As Boolean)
    m_Cmd_Search = vNewValue
End Property

Public Property Get Cmd_save() As Boolean
    Dim IntLoop                     As Integer
    Dim BolOpenTransaction          As Boolean
    Dim RsNotes                     As New ADODB.Recordset
    Dim RsDev                       As New ADODB.Recordset
    Dim BolMode                     As Boolean
    Dim StrNoteID As String
    Dim StrDEV_ID As String
    Cmd_save = m_Cmd_Save
    BolOpenTransaction = False
    On Error GoTo Cmd_save_ErrTrap

    With grd_Journal

        For IntLoop = .FixedRows To .Rows - 1

            If left(.TextMatrix(IntLoop, .ColIndex("NoteID")), 1) = "M" Then
                BolMode = True
                Exit For
            End If

        Next IntLoop

        If BolMode = False Then

            For IntLoop = .FixedRows To .Rows - 1

                If left(.TextMatrix(IntLoop, .ColIndex("DEV_ID")), 1) = "M" Then
                    BolMode = True
                    Exit For
                End If

            Next IntLoop

        End If

    End With

    If BolMode = False Then
        GetMsgs 84, vbExclamation
        Cmd_save = False
        Exit Sub
    End If

    'Open Recordsets
    RsNotes.Open "NOTES", Cn, adOpenStatic, adLockOptimistic, adCmdTableDirect
    RsDev.Open "DOUBLE_ENTREY_VOUCHERS", Cn, adOpenStatic, adLockOptimistic, adCmdTableDirect

    RsNotes.Index = "PrimaryKey"
    RsDev.Index = "PrimaryKey"
    Cn.BeginTrans
    BolOpenTransaction = True

    With grd_Journal

        For IntLoop = .FixedRows To .Rows - 1

            If Trim(.TextMatrix(IntLoop, .ColIndex("NoteID"))) <> "" Then
                If left(.TextMatrix(IntLoop, .ColIndex("NoteID")), 1) = "M" Then
                    StrNoteID = Mid(.TextMatrix(IntLoop, .ColIndex("NoteID")), 3)
                    RsNotes.Seek Trim(StrNoteID)
                    RsNotes("NotePosted").value = IIf(.Cell(flexcpChecked, IntLoop, .ColIndex("Post")) = flexChecked, True, False)
                    RsNotes("PostedBy").value = Me.DcboUsers.BoundText
                    RsNotes("PostDate").value = Format(Date, "yyyy/M/d")
                    RsNotes.update
                End If
            End If

        Next IntLoop

        For IntLoop = .FixedRows To .Rows - 1

            If Trim(.TextMatrix(IntLoop, .ColIndex("DEV_ID"))) <> "" Then
                If left(.TextMatrix(IntLoop, .ColIndex("DEV_ID")), 1) = "M" Then
                    StrDEV_ID = Mid(.TextMatrix(IntLoop, .ColIndex("DEV_ID")), 3)
                    RsDev.Seek Array(Trim(StrDEV_ID), Trim(.TextMatrix(IntLoop, .ColIndex("DEV_LineNo"))))
                    RsDev("Double_Entry_Vouchers_Description").value = CStr(.Cell(flexcpData, IntLoop, .ColIndex("Descrip")))
                    RsDev.update
                End If
            End If

        Next IntLoop

    End With

    Cn.CommitTrans
    MsgBox " „ Õ›Ÿ «·»Ì«‰«  ..", vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    RsNotes.Close
    'RsDEV.Close
    Set RsNotes = Nothing
    'Set RsDEV = Nothing
    Cmd_save = True
    Exit Property
Cmd_save_ErrTrap:

    If BolOpenTransaction Then
        Cn.RollbackTrans
    End If

    MsgBox " ⁄–— Õ›Ÿ ﬁÌœ «·ÌÊ„Ì… ..!!" & Chr(13) & "Õ«Ê· €·ﬁ «·‘«‘… Ê√⁄œ › ÕÂ« „‰ ÃœÌœ", vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Cmd_save = False
End Property

Public Property Let Cmd_save(ByVal vNewValue As Boolean)
    m_Cmd_Save = vNewValue
End Property

Private Sub grd_Journal_MouseMove(Button As Integer, _
                                  Shift As Integer, _
                                  X As Single, _
                                  Y As Single)
    'On Error GoTo ErrTrap
    'Dim LngOnRow As Long
    'Dim LngOnCol As Long
    'With grd_Journal
    '    .MousePointer = flexArrow
    '    If .MouseRow <= -1 Then Exit Sub
    '    If .MouseCol <= -1 Then Exit Sub
    '    If .MouseCol = .ColIndex("Descrip") Then
    '        LngOnRow = .MouseRow
    '        If .TextMatrix(.MouseRow, .ColIndex("AccountCode")) <> "" Then
    '            .Cell(flexcpFontUnderline, .FixedRows, .ColIndex("Descrip"), .Rows - 1, .ColIndex("Descrip")) = False
    '            .Cell(flexcpForeColor, .FixedRows, .ColIndex("Descrip"), .Rows - 1, .ColIndex("Descrip")) = vbBlack
    '            .Cell(flexcpFontBold, .FixedRows, .ColIndex("Descrip"), .Rows - 1, .ColIndex("Descrip")) = False
    '
    '            .Cell(flexcpFontUnderline, LngOnRow, .ColIndex("Descrip")) = True
    '            .Cell(flexcpForeColor, LngOnRow, .ColIndex("Descrip")) = vbBlue
    '            .Cell(flexcpFontBold, LngOnRow, .ColIndex("Descrip")) = True
    '            .MousePointer = flexHand
    '            If SystemOptions.UserInterface = ArabicInterface Then
    '                .ToolTipText = "≈÷€ÿ Â‰« ·„‘«Âœ… „·Œ’ ·‹Õ”«» " & .TextMatrix(LngOnRow, .ColIndex("Descrip"))
    '            Else
    '                .ToolTipText = "Click here to view a summery for " & .TextMatrix(LngOnRow, .ColIndex("Descrip"))
    '            End If
    '        ElseIf CStr(.Cell(flexcpData, .MouseRow, .MouseCol)) Like "Btn" & "*" Then
    '            .MousePointer = flexHand
    '            If SystemOptions.UserInterface = ArabicInterface Then
    '                .ToolTipText = "„⁄·Ê„«  ⁄‰ «·ﬁÌœ «·„”Ã·"
    '            Else
    '                .ToolTipText = "Information about this journal."
    '            End If
    '        End If
    '
    '    Else
    '        .MousePointer = flexArrow
    '        .ToolTipText = ""
    '        If .FixedRows = .Rows Then
    '            Exit Sub
    '        Else
    '            .Cell(flexcpFontUnderline, .FixedRows, .ColIndex("Descrip"), .Rows - 1, .ColIndex("Descrip")) = False
    '            .Cell(flexcpForeColor, .FixedRows, .ColIndex("Descrip"), .Rows - 1, .ColIndex("Descrip")) = vbBlack
    '            .Cell(flexcpFontBold, .FixedRows, .ColIndex("Descrip"), .Rows - 1, .ColIndex("Descrip")) = False
    '        End If
    '    End If
    '    If BolEditing = False Then
    '        ShowComment .MouseRow, .ColIndex("Descrip"), False
    '    End If
    'End With
    '
    'Exit Sub
    'ErrTrap:
End Sub

Private Sub grd_Journal_MouseUp(Button As Integer, _
                                Shift As Integer, _
                                X As Single, _
                                Y As Single)
    Dim LngFindRow As Long
    Dim StrSerachCode As String

    With grd_Journal

        If .MouseRow <= -1 Then Exit Sub
        If .MouseCol <= -1 Then Exit Sub
        If .MouseCol = .ColIndex("Descrip") Then
            If Button = vbLeftButton Then
                .ComboList = ""
                HideComment

                If .TextMatrix(.MouseRow, .ColIndex("AccountCode")) <> "" Then
                    StrSerachCode = .TextMatrix(.MouseRow, .ColIndex("AccountCode"))

                    With FgAccountsValue
                        LngFindRow = .FindRow(StrSerachCode, .FixedRows, .ColIndex("Account_code"), False, True)

                        If LngFindRow <> -1 Then
                            TabMain.CurrTab = 1
                            .SetFocus
                            .Row = LngFindRow
                            .ShowCell LngFindRow, .ColIndex("Account_code")
                            'PlaySoundEffect SnapItem
                        End If

                    End With

                End If

            Else

                If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
                    If .TextMatrix(.MouseRow, .ColIndex("AccountCode")) <> "" Then
                        If .Cell(flexcpData, .MouseRow, .ColIndex("Descrip")) = Empty Or .Cell(flexcpData, .MouseRow, .ColIndex("Descrip")) = "" Then
                            'No Comment So ' Enable the user to insert a new comment
                            '                    MDIFrmamin.MnuInsertCom.Enabled = True
                            '                    MDIFrmamin.MnuEditCom.Enabled = False
                            '                    MDIFrmamin.MnuDelCom.Enabled = False
                        Else
                            'allow user  to edit or delete the exist comment
                            '                    MDIFrmamin.MnuInsertCom.Enabled = False
                            '                    MDIFrmamin.MnuEditCom.Enabled = True
                            '                    MDIFrmamin.MnuDelCom.Enabled = True
                        End If

                        LngMouseRow = .MouseRow
                        'Me.PopupMenu MDIFrmamin.MnuAutiJournal
                    End If
                End If
            End If
        End If

    End With

End Sub

Private Sub m_DeleteMenu_Click()
    'On Error Resume Next
    Dim StrDEVID As String

    If LngMouseRow <= -1 Then Exit Sub
    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then

        With grd_Journal

            If .TextMatrix(LngMouseRow, .ColIndex("AccountCode")) <> "" Then
                If .Cell(flexcpData, LngMouseRow, .ColIndex("Descrip")) <> Empty Or .Cell(flexcpData, LngMouseRow, .ColIndex("Descrip")) <> "" Then
                    .Cell(flexcpData, LngMouseRow, .ColIndex("Descrip")) = ""
                    .Cell(flexcpPicture, LngMouseRow, .ColIndex("Descrip")) = Nothing
                    StrDEVID = .TextMatrix(LngMouseRow, .ColIndex("DEV_ID"))

                    If left(StrDEVID, 2) <> "M-" Then
                        'Put Flag
                        .TextMatrix(LngMouseRow, .ColIndex("DEV_ID")) = "M-" & StrDEVID
                    End If
                End If
            End If

            'MsgBox .TextMatrix(LngMouseRow, .ColIndex("Descrip"))
        End With

    End If

End Sub

Private Sub m_EditMenu_Click()
    ShowComment LngMouseRow, grd_Journal.ColIndex("Descrip"), True
    BolEditing = True
End Sub

Private Sub m_InsertMenu_Click()

    With grd_Journal
        .Cell(flexcpPicture, LngMouseRow, .ColIndex("Descrip")) = ImgNote.Picture

        'Make the pictrue Right alignment
        If val(.TextMatrix(LngMouseRow, .ColIndex("Debit"))) > 0 Then
            .Cell(flexcpPictureAlignment, LngMouseRow, .ColIndex("Descrip")) = flexPicAlignRightTop
        Else
            .Cell(flexcpPictureAlignment, LngMouseRow, .ColIndex("Descrip")) = flexPicAlignLeftTop
        End If

    End With

    ShowComment LngMouseRow, grd_Journal.ColIndex("Descrip"), True
    BolEditing = True
End Sub

Private Sub m_oText_Change()
    Dim StrDEVID As String

    With grd_Journal

        If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
            .Cell(flexcpData, LngMouseRow, .ColIndex("Descrip")) = m_oText.text
            StrDEVID = .TextMatrix(LngMouseRow, .ColIndex("DEV_ID"))

            If left(StrDEVID, 2) <> "M-" Then
                'Put Flag
                .TextMatrix(LngMouseRow, .ColIndex("DEV_ID")) = "M-" & StrDEVID
            End If
        End If

    End With

End Sub

Private Sub m_oText_KeyPress(KeyAscii As Integer)

    ' quit without saving when user hits Escape
    If KeyAscii = vbKeyEscape Then
        HideComment
        BolEditing = False
    End If

End Sub

Private Sub m_oText_LostFocus()
    BolEditing = False
End Sub

Private Sub m_oText_MouseMove(Button As Integer, _
                              Shift As Integer, _
                              X As Single, _
                              Y As Single)
    'If Me.TxtModFlg.Text = "R" Then
    '    HideComment
    'End If
End Sub

Private Sub OptViewStyle_Click(Index As Integer)
    Dim BolRtl As Boolean

    If SystemOptions.UserInterface = ArabicInterface Then
        BolRtl = True
    Else
        BolRtl = False
    End If

    With FgAccountsValue

        Select Case Index

            Case 0
                .ColPosition(.ColIndex("Debit_NotPosted")) = 5

                If BolRtl = True Then
                    .Cell(flexcpText, 0, .ColIndex("Account_Serial"), 1, .ColIndex("Account_Serial")) = "ﬂÊœ «·Õ”«»"
                    .Cell(flexcpText, 0, .ColIndex("Account_Name"), 1, .ColIndex("Account_Name")) = "«”„ «·Õ”«»"
            
                    .Cell(flexcpText, 0, .ColIndex("Debit_Posted"), 0, .ColIndex("Credit_Posted")) = "„—Õ·"
                    .Cell(flexcpText, 1, .ColIndex("Debit_Posted"), 1, .ColIndex("Debit_Posted")) = "„œÌ‰"
                    .Cell(flexcpText, 1, .ColIndex("Credit_Posted"), 1, .ColIndex("Credit_Posted")) = "œ«∆‰"
            
                    .Cell(flexcpText, 0, .ColIndex("Debit_NotPosted"), 0, .ColIndex("Credit_NotPosted")) = "€Ì— „—Õ·"
                    .Cell(flexcpText, 1, .ColIndex("Debit_NotPosted"), 1, .ColIndex("Debit_NotPosted")) = "„œÌ‰"
                    .Cell(flexcpText, 1, .ColIndex("Credit_NotPosted"), 1, .ColIndex("Credit_NotPosted")) = "œ«∆‰"
                Else
                    .Cell(flexcpText, 0, .ColIndex("Account_Serial"), 1, .ColIndex("Account_Serial")) = "Account Serial"
                    .Cell(flexcpText, 0, .ColIndex("Account_Name"), 1, .ColIndex("Account_Name")) = "Account Name"
             
                    .Cell(flexcpText, 0, .ColIndex("Debit_Posted"), 0, .ColIndex("Credit_Posted")) = "Posted"
                    .Cell(flexcpText, 1, .ColIndex("Debit_Posted"), 1, .ColIndex("Debit_Posted")) = "Debit"
                    .Cell(flexcpText, 1, .ColIndex("Credit_Posted"), 1, .ColIndex("Credit_Posted")) = "Credit"

                    .Cell(flexcpText, 0, .ColIndex("Debit_NotPosted"), 0, .ColIndex("Credit_NotPosted")) = "Not Posted"
                    .Cell(flexcpText, 1, .ColIndex("Debit_NotPosted"), 1, .ColIndex("Debit_NotPosted")) = "Debit"
                    .Cell(flexcpText, 1, .ColIndex("Credit_NotPosted"), 1, .ColIndex("Credit_NotPosted")) = "Credit"
                End If

            Case 1
                .ColPosition(.ColIndex("Debit_NotPosted")) = 4

                If BolRtl = True Then
                    .Cell(flexcpText, 0, .ColIndex("Account_Serial"), 1, .ColIndex("Account_Serial")) = "ﬂÊœ «·Õ”«»"
                    .Cell(flexcpText, 0, .ColIndex("Account_Name"), 1, .ColIndex("Account_Name")) = "«”„ «·Õ”«»"
            
                    .Cell(flexcpText, 0, .ColIndex("Debit_Posted"), 0, .ColIndex("Debit_NotPosted")) = "„œÌ‰"
                    .Cell(flexcpText, 1, .ColIndex("Debit_Posted"), 1, .ColIndex("Debit_Posted")) = "„—Õ·"
                    .Cell(flexcpText, 1, .ColIndex("Debit_NotPosted"), 1, .ColIndex("Debit_NotPosted")) = "€Ì— „—Õ·"
             
                    .Cell(flexcpText, 0, .ColIndex("Credit_Posted"), 0, .ColIndex("Credit_NotPosted")) = "œ«∆‰"
                    .Cell(flexcpText, 1, .ColIndex("Credit_Posted"), 1, .ColIndex("Credit_Posted")) = "„—Õ·"
                    .Cell(flexcpText, 1, .ColIndex("Credit_NotPosted"), 1, .ColIndex("Credit_NotPosted")) = "€Ì— „—Õ·"
                Else
                    .Cell(flexcpText, 0, .ColIndex("Account_Serial"), 1, .ColIndex("Account_Serial")) = "Account Serial"
                    .Cell(flexcpText, 0, .ColIndex("Account_Name"), 1, .ColIndex("Account_Name")) = "Account Name"
                    .Cell(flexcpText, 0, .ColIndex("Debit_Posted"), 0, .ColIndex("Debit_NotPosted")) = "Debit"
                    .Cell(flexcpText, 1, .ColIndex("Debit_Posted"), 1, .ColIndex("Debit_Posted")) = "Posted"
                    .Cell(flexcpText, 1, .ColIndex("Debit_NotPosted"), 1, .ColIndex("Debit_NotPosted")) = "Not Posted"
             
                    .Cell(flexcpText, 0, .ColIndex("Credit_Posted"), 0, .ColIndex("Credit_NotPosted")) = "Credit"
                    .Cell(flexcpText, 1, .ColIndex("Credit_Posted"), 1, .ColIndex("Credit_Posted")) = "Posted"
                    .Cell(flexcpText, 1, .ColIndex("Credit_NotPosted"), 1, .ColIndex("Credit_NotPosted")) = "Not Posted"

                End If

        End Select

    End With

End Sub

Private Sub TabMain_Switch(OldTab As Integer, _
                           NewTab As Integer, _
                           Cancel As Integer)

    If NewTab = 1 Then
        Ele(2).Visible = True
        Ele(3).Visible = False
        HideComment
    Else
        Ele(2).Visible = False
        Ele(3).Visible = True
    End If

End Sub

Private Sub TxtModFlg_Change()

    Select Case Me.TxtModFlg.text

        Case "N"

            'Me.grd_Journal.Editable = flexEDKbdMouse
        Case "E"

            'Me.grd_Journal.Editable = flexEDKbdMouse
        Case "R"
            'Me.grd_Journal.Editable = flexEDNone
    End Select

End Sub

Public Property Get Cmd_Edit() As Boolean
    Dim Msg As String

    If Frm_General_Journal.grd_Journal.Rows = Frm_General_Journal.grd_Journal.FixedRows Then
        GetMsgs 85, vbExclamation
        Cmd_Edit = False
        Exit Property
    End If

    Cmd_Edit = True
End Property

Public Property Let Cmd_Edit(ByVal vNewValue As Boolean)
    m_Cmd_Edit = vNewValue
End Property

Public Property Get Cmd_Undo() As Boolean
    Dim Msg As String
    Dim IntRes As Integer

    'IntRes = QueryUndoMsg(Me.TxtModFlg.text, Me.Caption)
    '
    'If IntRes = vbNo Then
    '    Cmd_Undo = False
    '    Exit Property
    'Else
    '    Me.Retrive Me.Tag
    'End If
    'Cmd_Undo = True
End Property

Public Property Let Cmd_Undo(ByVal vNewValue As Boolean)
    m_Cmd_Undo = vNewValue
End Property

Private Sub ChangeLang()
    CmdSearch.Caption = "Gl Search"
    Me.Caption = "Auditing Journal"
    Me.EleHeader.Caption = Me.Caption
    ChkPlanPaper.Caption = "Plan Paper"
    lbl(3).Caption = "Current User"
    TabMain.TabCaption(0) = "Issued Journal"
    TabMain.TabCaption(1) = "Accounts Banlance"
    Ele(2).Caption = "View Style"
    OptViewStyle(0).Caption = "Style 1"
    OptViewStyle(1).Caption = "Style 2"
    ChkColor(0).Caption = "Back Color"
    ChkColor(1).Caption = "Comment Color"
    CmdShowReport.Caption = "Ledger Report"

    With Me.FgAccountsValue
        '    .TextMatrix(0, .ColIndex("Account_Serial")) = "Account Serial"
        '    .TextMatrix(0, .ColIndex("Account_Name")) = "Account Name"
        '    .TextMatrix(0, .ColIndex("Debit_Posted")) = "Debit Posted"
        '    .TextMatrix(0, .ColIndex("Credit_Posted")) = "Credit_Posted"
    End With

End Sub

Public Sub SetGrgColors(IntType As Integer, _
                        BolRemove As Boolean)
    Dim i As Integer
    Dim j As Integer
    'IntType=0 ------->>> Journal
    'IntType=1 ------->>> Comment
    Dim VarTemp As Variant

    With Me.grd_Journal

        If IntType = 0 Then

            For i = .FixedRows To .Rows - 1

                If .TextMatrix(i, .ColIndex("AccountCode")) <> "" Then
                    If BolRemove = True Then
                        .Cell(flexcpBackColor, i, .ColIndex("Descrip"), i, .ColIndex("Descrip")) = 0
                    Else
                        .Cell(flexcpBackColor, i, .ColIndex("Descrip"), i, .ColIndex("Descrip")) = CPic(0).Color
                    End If
                End If

            Next i

        ElseIf IntType = 1 Then

            For i = .FixedRows To .Rows - 1

                If .Cell(flexcpData, i, .ColIndex("Descrip")) <> "" Then
                    VarTemp = Split(.Cell(flexcpData, i, .ColIndex("Descrip")), ";", , vbTextCompare)

                    If UBound(VarTemp) > 0 And (VarTemp(0) = "Btn" Or VarTemp(0) = "Plus") Then
                        If BolRemove = True Then

                            For j = val(VarTemp(2)) To val(VarTemp(3))
                                .Cell(flexcpBackColor, j, .ColIndex("Descrip"), j, .ColIndex("Descrip")) = 0
                            Next j

                        Else

                            For j = val(VarTemp(2)) To val(VarTemp(3))
                                .Cell(flexcpBackColor, j, .ColIndex("Descrip"), j, .ColIndex("Descrip")) = CPic(1).Color
                            Next j

                        End If
                    End If
                End If

            Next i

        End If

    End With

End Sub

Private Sub AddTip()
    Dim Msg As String

    Dim Wrap As String
    Dim i As Integer
    Wrap = Chr(13) + Chr(10)

    If SystemOptions.UserInterface = ArabicInterface Then

        With TTP
            .Create Me.hwnd, "≈ŸÂ«— «· ⁄·Ìﬁ", 1, 15204351, -2147483630, True
            .MaxWidth = 4000
            .VisibleTime = 15000
            .DelayTime = 300
            Msg = "» ›⁄Ì· Â–« «·ŒÌ«— Ì „ ≈ŸÂ«— Ã„Ì⁄" & Wrap & "«· ·„ÌÕ«  ⁄·Ï ﬂ· «·ﬁÌÊœ „—… " & Wrap & "Ê«Õœ… Ê»⁄œ„  ›⁄Ì·Â Ì „ ≈Œ›«¡ " & Wrap & "Ã„Ì⁄ «· ·„ÌÕ«  ."
            .AddControl ChkExpand, Msg, True
        End With

        With TTP
            .Create Me.hwnd, "·Ê‰ „„Ì“ ··ﬁÌÊœ", 1, 15204351, -2147483630, True
            .MaxWidth = 4000
            .VisibleTime = 15000
            .DelayTime = 300
            Msg = "» ›⁄Ì· Â–« «·ŒÌ«— Ì „ ⁄„· ·Ê‰ „⁄Ì‰" & Wrap & " ··ﬁÌÊœ Õ Ï Ì”Â· ⁄„·Ì… „—«Ã⁄… Â–Â" & Wrap & "«·ﬁÌÊœ ."
            .AddControl ChkColor(0), Msg, True
        End With

        With TTP
            .Create Me.hwnd, "·Ê‰ „„Ì“ ·· ·„ÌÕ« ", 1, 15204351, -2147483630, True
            .MaxWidth = 4000
            .VisibleTime = 15000
            .DelayTime = 300
            Msg = "» ›⁄Ì· Â–« «·ŒÌ«— Ì „ ⁄„· ·Ê‰ „⁄Ì‰" & Wrap & " ·· ·„ÌÕ«  Õ Ï Ì”Â· ⁄„·Ì… ﬁ—«¡… " & Wrap & "   ·„ÌÕ«  «·ﬁÌÊœ ."
            .AddControl ChkColor(1), Msg, True
        End With

        '
        With TTP
            .Create Me.hwnd, " ﬁ—Ì— «” «– „”«⁄œ", 1, 15204351, -2147483630, True
            .MaxWidth = 4000
            .VisibleTime = 15000
            .DelayTime = 300
            Msg = "≈÷€ÿ Â‰« Õ Ï Ì „ ⁄—÷  ﬁ—Ì—" & Wrap & "«” «– „”«⁄œ ··Õ”«» «·„Õœœ ."
            .AddControl CmdShowReport, Msg, True
        End With
    
        CPic(0).ToolTipText = " ÕœÌœ ·Ê‰ „⁄Ì‰ ·Œ·›Ì… «·ﬁÌÊœ"
        CPic(1).ToolTipText = " ÕœÌœ ·Ê‰ „⁄Ì‰ ·Œ·›Ì… «· ·„ÌÕ« "
    Else

        With TTP
            .Create Me.hwnd, "Show Comments", 1, 15204351, -2147483630, False
            .MaxWidth = 4000
            .VisibleTime = 15000
            .DelayTime = 300
            Msg = "When enable this option it will show " & Wrap & "all comments for all journals at once," & Wrap & "and when disabled it will hide it at " & Wrap & "once ."
            .AddControl ChkExpand, Msg, False
        End With

        With TTP
            .Create Me.hwnd, "Journal Background Color", 1, 15204351, -2147483630, False
            .MaxWidth = 4000
            .VisibleTime = 15000
            .DelayTime = 300
            Msg = "When enable this option it will " & Wrap & "apply a custome color to the " & Wrap & "background of journals."
            .AddControl ChkColor(0), Msg, False
        End With

        With TTP
            .Create Me.hwnd, "Comments Background Color", 1, 15204351, -2147483630, False
            .MaxWidth = 4000
            .VisibleTime = 15000
            .DelayTime = 300
            Msg = "When enable this option it will " & Wrap & " apply a custome color to the " & Wrap & "background of comments ."
            .AddControl ChkColor(1), Msg, False
        End With

        '
        With TTP
            .Create Me.hwnd, "Show Ledger Report", 1, 15204351, -2147483630, False
            .MaxWidth = 4000
            .VisibleTime = 15000
            .DelayTime = 300
            Msg = "Click here to show Ledger Report" & Wrap & "for the selected account ."
            .AddControl CmdShowReport, Msg, False
        End With
    
        CPic(0).ToolTipText = "Choose a background custome color for journals"
        CPic(1).ToolTipText = "Choose a background custome color for comments"
    End If

End Sub

Private Sub ShowComment(LngShowRow As Long, _
                        LngShowCol As Long, _
                        Optional bSetFocus As Boolean = False)
    Dim uPoint As POINTAPI
    Static lNoteRow As Long, lNoteCol As Long, r As Long, c As Long
    Dim lLeft  As Single
    Dim LTop As Single

    With grd_Journal

        If LngShowRow <= -1 Then Exit Sub
        If LngShowCol <= -1 Then Exit Sub
        If lNoteRow = LngShowRow And lNoteCol = LngShowCol Then
            If bSetFocus = False Then Exit Sub
        End If

        HideComment

        If LngShowCol = .ColIndex("Descrip") And .TextMatrix(LngShowRow, .ColIndex("AccountCode")) <> "" Then
            If (.Cell(flexcpData, LngShowRow, .ColIndex("Descrip")) <> Empty And .Cell(flexcpData, LngShowRow, .ColIndex("Descrip")) <> "") Or bSetFocus = True Then
                ClientToScreen grd_Journal.hwnd, uPoint
                lLeft = (uPoint.X * Screen.TwipsPerPixelX) + .ColPos(LngShowCol) + .ColWidth(LngShowCol)
                LTop = (uPoint.Y * Screen.TwipsPerPixelY) + .RowPos(LngShowRow)
                'lLeft = lLeft + 100
                Set FrmComment = New frmFlexNote

                With FrmComment
                    .left = lLeft
                    .top = LTop
                    '.Show
                    .txtNote.text = CStr(grd_Journal.Cell(flexcpData, LngShowRow, grd_Journal.ColIndex("Descrip")))
                    Set m_oText = .txtNote

                    If bSetFocus = False Then
                        ShowWindow .hwnd, SW_SHOWNA
                        .txtNote.Visible = False
                        .lblNote.Visible = True
                    Else
                        .txtNote.Visible = True
                        .lblNote.Visible = False
                        FrmComment.show
                        FrmComment.SetFocus
                        FrmComment.txtNote.SetFocus
                    End If

                End With

            End If

            lNoteRow = LngShowRow
            lNoteCol = LngShowCol
        End If

    End With

End Sub

Private Sub HideComment()

    If Not FrmComment Is Nothing Then
        Set m_oText = Nothing
        Unload FrmComment
        Set FrmComment = Nothing
    End If

End Sub

Public Sub Retrive(StrSQL As String)
    On Error Resume Next
    Dim StrDEVID        As String
    Dim BolDEVPostState As Boolean
    Dim BolNewDEV       As Boolean
    Dim BolRtl          As Boolean
    Dim LngLoop         As Long
    Dim StrWhere        As String
    Dim BolAccountRTL As Boolean

    If SystemOptions.UserInterface = ArabicInterface Then
        BolRtl = True 'salim erro
    Else
        'BolRtl = False
        BolRtl = True
    End If

    'If SystemOptions.UserShowDataAccounts = ShowArabicData Then
    '    BolAccountRTL = True
    'ElseIf SystemOptions.UserShowDataAccounts = ShowEnglishData Then
    '    BolAccountRTL = False
    'Else
    '    BolAccountRTL = IIf(SystemOptions.UserInterface = ArabicInterface, True, False)
    'End If
    If StrSQL = "" Then
        Exit Sub
    End If

    If RsLoadJournal.State = adStateOpen Then
        RsLoadJournal.Close
    End If

    Me.Tag = StrSQL
    RsLoadJournal.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsLoadJournal.EOF Or RsLoadJournal.BOF) Then
        Screen.MousePointer = vbArrowHourglass
        RsLoadJournal.MoveFirst
        IntCounter = 0

        With Frm_General_Journal.grd_Journal
            .Redraw = flexRDNone
            .Clear flexClearScrollable, flexClearEverything
            .Rows = .FixedRows
            IntCounter = .FixedRows - 1
            BolNewDEV = True

            Do While Not RsLoadJournal.EOF
                StrDEVID = RsLoadJournal("Double_Entry_Vouchers_ID").value
                BolDEVPostState = IIf(IsNull(RsLoadJournal("Posted").value), False, RsLoadJournal("Posted").value)
                IntCounter = IntCounter + 1
                .Rows = .Rows + 1

                If BolNewDEV = True Then
                    IntCounter = IntCounter + 1
                    .Rows = .Rows + 1
                    .TextMatrix(IntCounter, .ColIndex("DEV_NO")) = StrDEVID
                Else
                    .TextMatrix(IntCounter, .ColIndex("DEV_NO")) = ""
                End If

                'begine to write to Debit Side
                If RsLoadJournal("Credit_Or_Debit") = 0 Then
                    .TextMatrix(IntCounter, .ColIndex("Debit")) = IIf(RsLoadJournal("Credit_Or_Debit") = 0, Format(RsLoadJournal("Value"), SystemOptions.SysDefCurrencyForamt), Format(RsLoadJournal("Value"), SystemOptions.SysDefCurrencyForamt))

                    '     If BolRtl = True Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        .TextMatrix(IntCounter, .ColIndex("Descrip")) = "/" & RsLoadJournal("Account_Name").value ' SALIMERROR
                    Else
                        .TextMatrix(IntCounter, .ColIndex("Descrip")) = "" & RsLoadJournal("Account_NameEng").value
                    End If 'salim error

                    .TextMatrix(IntCounter, .ColIndex("AccountCode")) = RsLoadJournal("Account_Code").value
                    .TextMatrix(IntCounter, .ColIndex("NoteID")) = IIf(IsNull(RsLoadJournal("NoteID").value), "", RsLoadJournal("NoteID").value)
                    .TextMatrix(IntCounter, .ColIndex("DEV_ID")) = RsLoadJournal("Double_Entry_Vouchers_ID").value
                    .TextMatrix(IntCounter, .ColIndex("DEV_LineNo")) = RsLoadJournal("DEV_ID_Line_No").value
                    'debit side must be right align
                    .Cell(flexcpAlignment, IntCounter, .ColIndex("Descrip")) = flexAlignRightCenter

                    'the Description Put in the CellData Property  of this cell
                    If Len(CStr(IIf(IsNull(RsLoadJournal("Double_Entry_Vouchers_Description").value), "", RsLoadJournal("Double_Entry_Vouchers_Description").value))) Then
                        .Cell(flexcpData, IntCounter, .ColIndex("Descrip")) = RsLoadJournal("Double_Entry_Vouchers_Description").value
                        'Draw the Image on this Cell to inducate the user we have comment here
                        .Cell(flexcpPicture, IntCounter, .ColIndex("Descrip")) = Frm_General_Journal.ImgNote.Picture
                        'Make the pictrue Right alignment
                        .Cell(flexcpPictureAlignment, IntCounter, .ColIndex("Descrip")) = flexPicAlignRightTop
                    End If

                    .Cell(flexcpChecked, IntCounter, .ColIndex("DevPost")) = IIf(IsNull(RsLoadJournal("Posted").value), False, RsLoadJournal("Posted").value)
                    '=========================================
                ElseIf RsLoadJournal("Credit_Or_Debit") = 1 Then
                    'move next to get the Credit side
                    'IntCounter = IntCounter + 1
                    'begine to write the Credit side
                    .TextMatrix(IntCounter, .ColIndex("Credit")) = IIf(RsLoadJournal("Credit_Or_Debit").value = 1, Format(RsLoadJournal("Value"), SystemOptions.SysDefCurrencyForamt), Format(RsLoadJournal("Value"), SystemOptions.SysDefCurrencyForamt))

                    'If BolRtl = True Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        .TextMatrix(IntCounter, .ColIndex("Descrip")) = " /" & RsLoadJournal("Account_Name").value 'SALIM ERROR
                    Else
                        .TextMatrix(IntCounter, .ColIndex("Descrip")) = "" & RsLoadJournal("Account_NameEng").value
                    End If

                    .TextMatrix(IntCounter, .ColIndex("AccountCode")) = RsLoadJournal("Account_Code").value
                    .TextMatrix(IntCounter, .ColIndex("NoteID")) = IIf(IsNull(RsLoadJournal("NoteId").value), "", RsLoadJournal("NoteId").value)
                    .TextMatrix(IntCounter, .ColIndex("DEV_ID")) = RsLoadJournal("Double_Entry_Vouchers_ID").value
                    .TextMatrix(IntCounter, .ColIndex("DEV_LineNo")) = RsLoadJournal("DEV_ID_Line_No").value
                    'credit side must be left align
                    .Cell(flexcpAlignment, IntCounter, .ColIndex("Descrip")) = flexAlignLeftCenter

                    'the Description Put in the CellData Property  of this cell
                    If Len(CStr(IIf(IsNull(RsLoadJournal("Double_Entry_Vouchers_Description").value), "", RsLoadJournal("Double_Entry_Vouchers_Description").value))) Then
                        .Cell(flexcpData, IntCounter, .ColIndex("Descrip")) = RsLoadJournal("Double_Entry_Vouchers_Description").value
                        'Draw the Image on this Cell to inducate the user we have comment here
                        .Cell(flexcpPicture, IntCounter, .ColIndex("Descrip")) = Frm_General_Journal.ImgNote.Picture
                        'Make the pictrue Right alignment
                        .Cell(flexcpPictureAlignment, IntCounter, .ColIndex("Descrip")) = flexPicAlignLeftTop
                    End If

                    .Cell(flexcpChecked, IntCounter, .ColIndex("DevPost")) = IIf(IsNull(RsLoadJournal("Posted").value), False, RsLoadJournal("Posted").value)
                End If

                RsLoadJournal.MoveNext

                If Not RsLoadJournal.EOF Then
                    If StrDEVID <> RsLoadJournal("Double_Entry_Vouchers_ID").value Then
                        BolNewDEV = True
                        .Select IntCounter, .ColIndex("Descrip")
                        .CellBorder vbBlack - 1, -1, -1, -1, 2, -1, -1
                        .Rows = .Rows + 1
                        IntCounter = IntCounter + 1

                        If SystemOptions.UserInterface = ArabicInterface Then
                            .Cell(flexcpText, IntCounter, .ColIndex("Post")) = IIf(BolDEVPostState = True, "„‹—Õ‹·", " ‹—Õ‹Ì‹·")
                            .Cell(flexcpPictureAlignment, IntCounter, .ColIndex("Post")) = flexPicAlignRightCenter
                        Else
                            .Cell(flexcpText, IntCounter, .ColIndex("Post")) = IIf(BolDEVPostState = True, "Posted", "Post")
                            .Cell(flexcpPictureAlignment, IntCounter, .ColIndex("Post")) = flexPicAlignLeftCenter
                        End If

                        .Cell(flexcpText, IntCounter, .ColIndex("Note_Date")) = Format(RsLoadJournal("RecordDate").value, "yyyy/M/d")
                        .Cell(flexcpChecked, IntCounter, .ColIndex("Post")) = IIf(BolDEVPostState = True, flexChecked, flexUnchecked)
                        .Cell(flexcpForeColor, IntCounter, .ColIndex("Post")) = IIf(BolDEVPostState = True, vbRed, vbBlack)
                        .Cell(flexcpPicture, IntCounter, .ColIndex("Descrip")) = Frm_General_Journal.ImgLst.ListImages("Plus").Picture
                        .Cell(flexcpPictureAlignment, IntCounter, .ColIndex("Descrip")) = flexPicAlignRightCenter
                        '------Add Information
                        AddInfo
                        '-----------------------
                    Else
                        BolNewDEV = False
                    End If

                Else
                    .Select IntCounter, .ColIndex("Descrip")
                    .CellBorder vbBlack - 1, -1, -1, -1, 2, -1, -1
                    .Rows = .Rows + 1
                    IntCounter = IntCounter + 1

                    If SystemOptions.UserInterface = ArabicInterface Then
                        .Cell(flexcpText, IntCounter, .ColIndex("Post")) = IIf(BolDEVPostState = True, "„‹—Õ‹·", " ‹—Õ‹Ì‹·")
                        .Cell(flexcpPictureAlignment, IntCounter, .ColIndex("Post")) = flexPicAlignRightCenter
                    Else
                        .Cell(flexcpText, IntCounter, .ColIndex("Post")) = IIf(BolDEVPostState = True, "Posted", "Post")
                        .Cell(flexcpPictureAlignment, IntCounter, .ColIndex("Post")) = flexPicAlignLeftCenter
                    End If

                    RsLoadJournal.MoveLast
                    .Cell(flexcpText, IntCounter, .ColIndex("Note_Date")) = DisplayDate(RsLoadJournal("RecordDate").value)
                    .Cell(flexcpChecked, IntCounter, .ColIndex("Post")) = IIf(BolDEVPostState = True, flexChecked, flexUnchecked)
                    .Cell(flexcpForeColor, IntCounter, .ColIndex("Post")) = IIf(BolDEVPostState = True, vbRed, vbBlack)
                    .Cell(flexcpPicture, IntCounter, .ColIndex("Descrip")) = Frm_General_Journal.ImgLst.ListImages("Plus").Picture
                    .Cell(flexcpPictureAlignment, IntCounter, .ColIndex("Descrip")) = flexPicAlignRightCenter
                    RsLoadJournal.MoveNext
                    AddInfo
                End If

            Loop

            .Redraw = flexRDDirect
            Screen.MousePointer = vbDefault
        End With

        Frm_General_Journal.DTPDev_From.value = Me.DTPDev_From.value
        Frm_General_Journal.DTPDEV_TO.value = Me.DTPDEV_TO.value
    Else
        GetMsgs 77, vbInformation
        Frm_General_Journal.grd_Journal.Clear flexClearScrollable, flexClearEverything
        Frm_General_Journal.grd_Journal.Rows = Frm_General_Journal.grd_Journal.FixedRows
        Screen.MousePointer = vbDefault
    End If

End Sub

Private Sub AddInfo(Optional BolMoveLast As Boolean = False)
    Dim BolRtl          As Boolean
    Dim LngBegRow       As Long
    Dim LngEndRow       As Long
    Dim LngLoop         As Long

    If SystemOptions.UserInterface = ArabicInterface Then
        BolRtl = True
    Else
        BolRtl = False
    End If

    Exit Sub

    With Frm_General_Journal.grd_Journal
        .Rows = .Rows + 1
        IntCounter = IntCounter + 1
        LngBegRow = IntCounter
        RsLoadJournal.MovePrevious

        If Not (IsNull(RsLoadJournal("Transaction_Type").value)) Then
            .Cell(flexcpPicture, IntCounter, .ColIndex("Descrip")) = Frm_General_Journal.ImgLst.ListImages("Doc").Picture

            If BolRtl = True Then
                .Cell(flexcpText, IntCounter, .ColIndex("Descrip")) = RsLoadJournal("TransactionArabicName").value & " —ﬁ„ " & RsLoadJournal("Transaction_Serial_No").value
            Else
                .Cell(flexcpText, IntCounter, .ColIndex("Descrip")) = RsLoadJournal("TransactionEnglishName").value & " NO. " & RsLoadJournal("Transaction_Serial_No").value
            End If

        ElseIf Not (IsNull(RsLoadJournal("ReNoteType").value)) Then
            .Cell(flexcpPicture, IntCounter, .ColIndex("Descrip")) = Frm_General_Journal.ImgLst.ListImages("NoteDoc").Picture

            If BolRtl = True Then
                .Cell(flexcpText, IntCounter, .ColIndex("Descrip")) = RsLoadJournal("ReNoteArabic").value & " —ﬁ„ " & RsLoadJournal("ReNoteSer").value
            Else
                .Cell(flexcpText, IntCounter, .ColIndex("Descrip")) = RsLoadJournal("ReNoteEng").value & " NO. " & RsLoadJournal("ReNoteSer").value
            End If

        ElseIf Not (IsNull(RsLoadJournal("NoteType").value)) Then
            .Cell(flexcpPicture, IntCounter, .ColIndex("Descrip")) = Frm_General_Journal.ImgLst.ListImages("NoteDoc").Picture

            If BolRtl = True Then
                .Cell(flexcpText, IntCounter, .ColIndex("Descrip")) = RsLoadJournal("NoteNameArabic").value & " —ﬁ„ " & RsLoadJournal("Chique_Serial_No").value
            Else
                .Cell(flexcpText, IntCounter, .ColIndex("Descrip")) = RsLoadJournal("NoteNameEnglish").value & " NO. " & RsLoadJournal("Chique_Serial_No").value
            End If
        End If

        .Rows = .Rows + 1
        IntCounter = IntCounter + 1
        .Cell(flexcpPicture, IntCounter, .ColIndex("Descrip")) = Frm_General_Journal.ImgLst.ListImages("IssuedUser").Picture

        If BolRtl = True Then
            .Cell(flexcpText, IntCounter, .ColIndex("Descrip")) = "«·„” Œœ„ «·„Õ—— :" & RsLoadJournal("IssuedUser").value
        Else
            .Cell(flexcpText, IntCounter, .ColIndex("Descrip")) = "Issued By : " & RsLoadJournal("IssuedUser").value
        End If

        If Not (IsNull(RsLoadJournal("PostedUser"))) Then
            .Rows = .Rows + 1
            IntCounter = IntCounter + 1
            .Cell(flexcpPicture, IntCounter, .ColIndex("Descrip")) = Frm_General_Journal.ImgLst.ListImages("PostedUser").ExtractIcon

            If BolRtl = True Then
                .Cell(flexcpText, IntCounter, .ColIndex("Descrip")) = "«·„” Œœ„ «·„—Õ· :" & RsLoadJournal("PostedUser").value
            Else
                .Cell(flexcpText, IntCounter, .ColIndex("Descrip")) = "Posted By :" & RsLoadJournal("PostedUser").value
            End If

            If Not (IsNull(RsLoadJournal("PostDate").value)) Then
                .Rows = .Rows + 1
                IntCounter = IntCounter + 1
                .Cell(flexcpPicture, IntCounter, .ColIndex("Descrip")) = Frm_General_Journal.ImgLst.ListImages("PostDate").ExtractIcon

                If BolRtl = True Then
                    .Cell(flexcpText, IntCounter, .ColIndex("Descrip")) = " «—ÌŒ «· —ÕÌ· :" & Format(RsLoadJournal("PostDate").value, "yyyy/M/d")
                Else
                    .Cell(flexcpText, IntCounter, .ColIndex("Descrip")) = "Post Date :" & Format(RsLoadJournal("PostDate").value, "yyyy/M/d")
                End If
            End If
        End If

        LngEndRow = IntCounter
        .Cell(flexcpData, LngBegRow - 1, .ColIndex("Descrip")) = "Btn;Plus;" & LngBegRow & ";" & LngEndRow & ""
        .Select LngBegRow, .ColIndex("Descrip"), LngEndRow, .ColIndex("Descrip")

        If BolRtl = True Then
            .Cell(flexcpAlignment, LngBegRow, .ColIndex("Descrip"), LngEndRow, .ColIndex("Descrip")) = flexAlignRightCenter
            .Cell(flexcpPictureAlignment, LngBegRow, .ColIndex("Descrip"), LngEndRow, .ColIndex("Descrip")) = flexPicAlignRightCenter
        Else
            .Cell(flexcpAlignment, LngBegRow, .ColIndex("Descrip"), LngEndRow, .ColIndex("Descrip")) = flexAlignLeftCenter
            .Cell(flexcpPictureAlignment, LngBegRow, .ColIndex("Descrip"), LngEndRow, .ColIndex("Descrip")) = flexPicAlignLeftCenter
        End If

        If Frm_General_Journal.ChkColor(1).value = vbChecked Then
            .Cell(flexcpBackColor, LngBegRow, .ColIndex("Descrip"), LngEndRow, .ColIndex("Descrip")) = Frm_General_Journal.CPic(1).Color
        Else
            .Cell(flexcpBackColor, LngBegRow, .ColIndex("Descrip"), LngEndRow, .ColIndex("Descrip")) = 0
        End If

        .Redraw = flexRDDirect
        .Refresh

        For LngLoop = LngBegRow To LngEndRow
            .RowHidden(LngLoop) = True
        Next LngLoop

        RsLoadJournal.MoveNext
    End With

End Sub

