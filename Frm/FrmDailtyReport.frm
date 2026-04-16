VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmDailtyReport 
   BackColor       =   &H00C0FFFF&
   Caption         =   "«· ﬁ—Ì— «·ÌÊ„Ì"
   ClientHeight    =   7740
   ClientLeft      =   195
   ClientTop       =   450
   ClientWidth     =   15645
   HelpContextID   =   250
   Icon            =   "FrmDailtyReport.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   7740
   ScaleWidth      =   15645
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
      Height          =   7740
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   15645
      _cx             =   27596
      _cy             =   13653
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
      _GridInfo       =   $"FrmDailtyReport.frx":038A
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   555
         Left            =   30
         TabIndex        =   158
         TabStop         =   0   'False
         Top             =   7155
         Width           =   15585
         _cx             =   27490
         _cy             =   979
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
         Begin MSDataListLib.DataCombo DCBranches2 
            Height          =   315
            Left            =   13560
            TabIndex        =   159
            Top             =   0
            Width           =   5145
            _ExtentX        =   9075
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcboStoresx 
            Height          =   315
            Index           =   2
            Left            =   4080
            TabIndex        =   161
            Top             =   0
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "Õœœ „Œ“‰"
            Height          =   195
            Index           =   169
            Left            =   9480
            RightToLeft     =   -1  'True
            TabIndex        =   162
            Top             =   0
            Width           =   915
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "Õœœ «·›—⁄"
            Height          =   195
            Index           =   144
            Left            =   18960
            RightToLeft     =   -1  'True
            TabIndex        =   160
            Top             =   0
            Width           =   915
         End
      End
      Begin C1SizerLibCtl.C1Tab TabMain 
         Height          =   7110
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   15585
         _cx             =   27490
         _cy             =   12541
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
         ForeColor       =   -2147483630
         FrontTabColor   =   14737632
         BackTabColor    =   12648447
         TabOutlineColor =   -2147483632
         FrontTabForeColor=   16711680
         Caption         =   "«·„»Ì⁄« |«·„‘ —Ì« |„— Ã⁄ «·„»Ì⁄« |„— Ã⁄ «·„‘ —Ì« |«·’Ì«‰…|«·„’—Ê›« |«·„œ›Ê⁄« |«·„ﬁ»Ê÷« |«·„·Œ’|«·√’‰«›|«·⁄„·«¡| Õ—ﬂ…  «·Œ“‰"
         Align           =   0
         CurrTab         =   0
         FirstTab        =   0
         Style           =   3
         Position        =   6
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
         Picture(0)      =   "FrmDailtyReport.frx":0410
         Picture(1)      =   "FrmDailtyReport.frx":07AA
         Picture(2)      =   "FrmDailtyReport.frx":0B44
         Picture(3)      =   "FrmDailtyReport.frx":10DE
         Picture(4)      =   "FrmDailtyReport.frx":1678
         Flags(4)        =   2
         Picture(5)      =   "FrmDailtyReport.frx":1A12
         Picture(6)      =   "FrmDailtyReport.frx":1DAC
         Picture(7)      =   "FrmDailtyReport.frx":2146
         Picture(8)      =   "FrmDailtyReport.frx":24E0
         Picture(9)      =   "FrmDailtyReport.frx":287A
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   7020
            Index           =   20
            Left            =   18930
            TabIndex        =   135
            TabStop         =   0   'False
            Top             =   45
            Width           =   13980
            _cx             =   24659
            _cy             =   12383
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
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   900
               Index           =   21
               Left            =   0
               TabIndex        =   140
               TabStop         =   0   'False
               Top             =   6105
               Width           =   13980
               _cx             =   24659
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
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   450
                  Index           =   2
                  Left            =   0
                  TabIndex        =   141
                  Top             =   120
                  Visible         =   0   'False
                  Width           =   240
                  _ExtentX        =   423
                  _ExtentY        =   794
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   "⁄—÷ «· ﬁ—Ì—"
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
                  ButtonImage     =   "FrmDailtyReport.frx":2C14
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
            End
            Begin ImpulseButton.ISButton CmdShow 
               Height          =   315
               Index           =   9
               Left            =   3135
               TabIndex        =   139
               Top             =   15
               Width           =   1995
               _ExtentX        =   3519
               _ExtentY        =   556
               ButtonStyle     =   1
               Caption         =   "⁄—÷ «·»Ì«‰« "
               BackColor       =   12634304
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ColorButton     =   12634304
               DrawFocusRectangle=   0   'False
            End
            Begin MSComCtl2.DTPicker DtpCustomers 
               Height          =   315
               Left            =   5130
               TabIndex        =   138
               Top             =   15
               Width           =   2280
               _ExtentX        =   4022
               _ExtentY        =   556
               _Version        =   393216
               Format          =   98959361
               CurrentDate     =   39578
               MaxDate         =   768106
            End
            Begin VSFlex8UCtl.VSFlexGrid FgCustomers 
               Height          =   5745
               Left            =   0
               TabIndex        =   136
               Top             =   345
               Width           =   13980
               _cx             =   24659
               _cy             =   10134
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
               FormatString    =   $"FrmDailtyReport.frx":2FAE
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
            Begin MSComCtl2.DTPicker DTPicker1 
               Height          =   315
               Left            =   180
               TabIndex        =   157
               Top             =   0
               Visible         =   0   'False
               Width           =   2280
               _ExtentX        =   4022
               _ExtentY        =   556
               _Version        =   393216
               Format          =   98959361
               CurrentDate     =   39578
               MaxDate         =   768106
            End
            Begin VB.Label LblTitle 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·⁄„·«¡ Ê«·„Ê—œÌ‰ «·–Ì‰  „ «· ⁄«„· „⁄Â„ «·ÌÊ„"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   315
               Index           =   3
               Left            =   7410
               RightToLeft     =   -1  'True
               TabIndex        =   137
               Top             =   15
               Width           =   6570
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   7020
            Index           =   18
            Left            =   18630
            TabIndex        =   123
            TabStop         =   0   'False
            Top             =   45
            Width           =   13980
            _cx             =   24659
            _cy             =   12383
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
            Begin MSDataListLib.DataCombo DcboStores 
               Height          =   315
               Left            =   4290
               TabIndex        =   130
               Top             =   30
               Width           =   3705
               _ExtentX        =   6535
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VSFlex8UCtl.VSFlexGrid FgItems 
               Height          =   5640
               Left            =   0
               TabIndex        =   128
               Top             =   345
               Width           =   13980
               _cx             =   24659
               _cy             =   9948
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
               SelectionMode   =   0
               GridLines       =   1
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   50
               Cols            =   8
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmDailtyReport.frx":31E1
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
               AllowUserFreezing=   1
               BackColorFrozen =   0
               ForeColorFrozen =   0
               WallPaperAlignment=   9
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   24
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   990
               Index           =   19
               Left            =   0
               TabIndex        =   127
               TabStop         =   0   'False
               Top             =   6000
               Width           =   13980
               _cx             =   24659
               _cy             =   1746
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
                  Height          =   330
                  Index           =   0
                  Left            =   15
                  TabIndex        =   133
                  Top             =   45
                  Width           =   120
                  _ExtentX        =   212
                  _ExtentY        =   582
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   "⁄—÷ «· ﬁ—Ì—"
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
                  ButtonImage     =   "FrmDailtyReport.frx":3343
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   345
                  Index           =   1
                  Left            =   15
                  TabIndex        =   134
                  Top             =   390
                  Visible         =   0   'False
                  Width           =   120
                  _ExtentX        =   212
                  _ExtentY        =   609
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
                  ButtonImage     =   "FrmDailtyReport.frx":36DD
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Height          =   285
                  Index           =   23
                  Left            =   540
                  RightToLeft     =   -1  'True
                  TabIndex        =   132
                  Top             =   120
                  Width           =   60
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "≈Ã„«·Ï ⁄œœ «·√’‰«›:"
                  Height          =   285
                  Index           =   22
                  Left            =   600
                  RightToLeft     =   -1  'True
                  TabIndex        =   131
                  Top             =   120
                  Width           =   135
               End
            End
            Begin ImpulseButton.ISButton CmdShow 
               Height          =   300
               Index           =   8
               Left            =   0
               TabIndex        =   125
               Top             =   30
               Width           =   2280
               _ExtentX        =   4022
               _ExtentY        =   529
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   "⁄—÷ «·»Ì«‰« "
               BackColor       =   12634304
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ColorButton     =   12634304
               DrawFocusRectangle=   0   'False
            End
            Begin MSComCtl2.DTPicker DtpTransDate 
               Height          =   300
               Index           =   8
               Left            =   2280
               TabIndex        =   126
               Top             =   30
               Width           =   2010
               _ExtentX        =   3545
               _ExtentY        =   529
               _Version        =   393216
               CalendarBackColor=   12648447
               Format          =   98959361
               CurrentDate     =   38718
               MaxDate         =   768106
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«Œ — «”„ «·„Œ“‰"
               ForeColor       =   &H00000080&
               Height          =   300
               Index           =   21
               Left            =   7995
               RightToLeft     =   -1  'True
               TabIndex        =   129
               Top             =   30
               Width           =   1995
            End
            Begin VB.Label LblTitle 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„·Œ’ Õ—ﬂ… «·„Œ“Ê‰ «·ÌÊ„"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   300
               Index           =   2
               Left            =   9990
               RightToLeft     =   -1  'True
               TabIndex        =   124
               Top             =   30
               Width           =   3990
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   7020
            Index           =   1
            Left            =   16230
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   45
            Width           =   13980
            _cx             =   24659
            _cy             =   12383
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
            CaptionPos      =   4
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
               Height          =   990
               Index           =   7
               Left            =   0
               TabIndex        =   12
               TabStop         =   0   'False
               Top             =   6000
               Width           =   13980
               _cx             =   24659
               _cy             =   1746
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
               Begin ImpulseButton.ISButton CmdHelp 
                  Height          =   195
                  Index           =   1
                  Left            =   0
                  TabIndex        =   13
                  Top             =   510
                  Width           =   285
                  _ExtentX        =   503
                  _ExtentY        =   344
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
                  ButtonImage     =   "FrmDailtyReport.frx":3A77
                  ColorButton     =   14871017
                  ColorHighlight  =   16777215
                  ColorHoverText  =   16711680
                  ColorShadow     =   -2147483637
                  ColorOutline    =   0
                  DrawFocusRectangle=   0   'False
                  ColorToggledHoverText=   16711680
                  ColorTextShadow =   -2147483637
               End
               Begin ImpulseButton.ISButton CmdPrintBuy 
                  Height          =   225
                  Index           =   0
                  Left            =   420
                  TabIndex        =   14
                  Top             =   30
                  Width           =   1335
                  _ExtentX        =   2355
                  _ExtentY        =   397
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   "⁄—÷ «· ﬁ—Ì—"
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
                  ButtonImage     =   "FrmDailtyReport.frx":3E11
                  ColorButton     =   14871017
                  ColorHighlight  =   16777215
                  ColorHoverText  =   16711680
                  ColorOutline    =   0
                  DrawFocusRectangle=   0   'False
                  ColorToggledHoverText=   16711680
                  LowerToggledContent=   0   'False
                  ColorTextShadow =   -2147483637
               End
               Begin ImpulseButton.ISButton CmdPrintBuy 
                  Height          =   195
                  Index           =   1
                  Left            =   0
                  TabIndex        =   15
                  Top             =   285
                  Visible         =   0   'False
                  Width           =   285
                  _ExtentX        =   503
                  _ExtentY        =   344
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
                  BackStyle       =   0
                  ButtonImage     =   "FrmDailtyReport.frx":41AB
                  ColorButton     =   14871017
                  ColorHighlight  =   16777215
                  ColorHoverText  =   16711680
                  ColorOutline    =   0
                  DrawFocusRectangle=   0   'False
                  ColorToggledHoverText=   16711680
                  LowerToggledContent=   0   'False
                  ColorTextShadow =   -2147483637
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " ≈Ã„«·Ì ﬁÌ„… «·„‘ —Ì«  :"
                  Height          =   210
                  Index           =   4
                  Left            =   2925
                  RightToLeft     =   -1  'True
                  TabIndex        =   23
                  Top             =   75
                  Width           =   1275
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " ≈Ã„«·Ì «·Œ’Ê„«   :"
                  Height          =   210
                  Index           =   5
                  Left            =   5730
                  RightToLeft     =   -1  'True
                  TabIndex        =   22
                  Top             =   300
                  Width           =   990
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " ’«›Ì ﬁÌ„… «·„‘ —Ì«  :"
                  Height          =   210
                  Index           =   13
                  Left            =   5445
                  RightToLeft     =   -1  'True
                  TabIndex        =   21
                  Top             =   75
                  Width           =   1275
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " ≈Ã„«·Ì ⁄œœ «·›Ê« Ì—  :"
                  Height          =   210
                  Index           =   14
                  Left            =   2940
                  RightToLeft     =   -1  'True
                  TabIndex        =   20
                  Top             =   300
                  Width           =   1110
               End
               Begin VB.Label LblVal 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Height          =   225
                  Index           =   1
                  Left            =   4050
                  RightToLeft     =   -1  'True
                  TabIndex        =   19
                  Top             =   75
                  Width           =   1395
               End
               Begin VB.Label LblDiscount 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Height          =   210
                  Index           =   1
                  Left            =   4200
                  RightToLeft     =   -1  'True
                  TabIndex        =   18
                  Top             =   300
                  Width           =   1245
               End
               Begin VB.Label LblTotal 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Height          =   210
                  Index           =   1
                  Left            =   1530
                  RightToLeft     =   -1  'True
                  TabIndex        =   17
                  Top             =   75
                  Width           =   1275
               End
               Begin VB.Label LblBillCount 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Height          =   210
                  Index           =   1
                  Left            =   1530
                  RightToLeft     =   -1  'True
                  TabIndex        =   16
                  Top             =   300
                  Width           =   1275
               End
            End
            Begin MSComCtl2.DTPicker DtpTransDate 
               Height          =   300
               Index           =   1
               Left            =   2280
               TabIndex        =   3
               Top             =   30
               Width           =   2850
               _ExtentX        =   5027
               _ExtentY        =   529
               _Version        =   393216
               CalendarBackColor=   12648447
               CalendarForeColor=   4210752
               CalendarTitleForeColor=   0
               Format          =   98959361
               CurrentDate     =   38718
               MaxDate         =   768106
            End
            Begin VSFlex8UCtl.VSFlexGrid FgPurchase 
               Height          =   5640
               Left            =   0
               TabIndex        =   6
               Top             =   345
               Width           =   13980
               _cx             =   24659
               _cy             =   9948
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
               SelectionMode   =   0
               GridLines       =   1
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   10
               Cols            =   10
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmDailtyReport.frx":4545
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
            Begin ImpulseButton.ISButton CmdShow 
               Height          =   300
               Index           =   1
               Left            =   0
               TabIndex        =   7
               Top             =   30
               Width           =   2280
               _ExtentX        =   4022
               _ExtentY        =   529
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   "⁄—÷ «·»Ì«‰« "
               BackColor       =   12634304
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ButtonImage     =   "FrmDailtyReport.frx":46DB
               ColorButton     =   12634304
               DrawFocusRectangle=   0   'False
            End
            Begin VB.Label LblTitle 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "›Ê« Ì— «·„‘ —Ì«  «· Ï  „  «·ÌÊ„"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   8
               Left            =   5130
               RightToLeft     =   -1  'True
               TabIndex        =   4
               Top             =   30
               Width           =   8850
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   7020
            Index           =   0
            Left            =   45
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   45
            Width           =   13980
            _cx             =   24659
            _cy             =   12383
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
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   990
               Index           =   6
               Left            =   0
               TabIndex        =   24
               TabStop         =   0   'False
               Top             =   6000
               Width           =   13980
               _cx             =   24659
               _cy             =   1746
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
               Begin ImpulseButton.ISButton CmdPrint 
                  Height          =   270
                  Index           =   0
                  Left            =   270
                  TabIndex        =   25
                  Top             =   45
                  Width           =   2595
                  _ExtentX        =   4577
                  _ExtentY        =   476
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   "⁄—÷ «· ﬁ—Ì—"
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
                  ButtonImage     =   "FrmDailtyReport.frx":4A75
                  ColorButton     =   14871017
                  ColorHighlight  =   16777215
                  ColorHoverText  =   16711680
                  ColorOutline    =   0
                  DrawFocusRectangle=   0   'False
                  ColorToggledHoverText=   16711680
                  LowerToggledContent=   0   'False
                  ColorTextShadow =   -2147483637
               End
               Begin ImpulseButton.ISButton CmdPrint 
                  Height          =   285
                  Index           =   1
                  Left            =   0
                  TabIndex        =   26
                  Top             =   315
                  Visible         =   0   'False
                  Width           =   585
                  _ExtentX        =   1032
                  _ExtentY        =   503
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
                  ButtonImage     =   "FrmDailtyReport.frx":4E0F
                  ColorButton     =   14871017
                  ColorHighlight  =   16777215
                  ColorHoverText  =   16711680
                  ColorOutline    =   0
                  DrawFocusRectangle=   0   'False
                  ColorToggledHoverText=   16711680
                  LowerToggledContent=   0   'False
                  ColorTextShadow =   -2147483637
               End
               Begin ImpulseButton.ISButton CmdHelp 
                  Height          =   285
                  Index           =   0
                  Left            =   0
                  TabIndex        =   27
                  Top             =   600
                  Width           =   585
                  _ExtentX        =   1032
                  _ExtentY        =   503
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
                  ButtonImage     =   "FrmDailtyReport.frx":51A9
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
                  BackStyle       =   0  'Transparent
                  Caption         =   " ≈Ã„«·Ì ﬁÌ„… «·„»Ì⁄«  :"
                  Height          =   225
                  Index           =   0
                  Left            =   10860
                  RightToLeft     =   -1  'True
                  TabIndex        =   35
                  Top             =   90
                  Width           =   2850
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   " ≈Ã„«·Ì «·Œ’Ê„«  :"
                  Height          =   225
                  Index           =   1
                  Left            =   11415
                  RightToLeft     =   -1  'True
                  TabIndex        =   34
                  Top             =   600
                  Width           =   2295
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   " ’«›Ì ﬁÌ„… «·„»Ì⁄«  :"
                  Height          =   255
                  Index           =   2
                  Left            =   4275
                  RightToLeft     =   -1  'True
                  TabIndex        =   33
                  Top             =   120
                  Width           =   3435
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   " ≈Ã„«·Ì ⁄œœ «·›Ê« Ì— :"
                  Height          =   255
                  Index           =   3
                  Left            =   4275
                  RightToLeft     =   -1  'True
                  TabIndex        =   32
                  Top             =   600
                  Width           =   3435
               End
               Begin VB.Label LblTotal 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Height          =   360
                  Index           =   0
                  Left            =   2565
                  RightToLeft     =   -1  'True
                  TabIndex        =   31
                  Top             =   90
                  Width           =   2295
               End
               Begin VB.Label LblBillCount 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Height          =   360
                  Index           =   0
                  Left            =   1425
                  RightToLeft     =   -1  'True
                  TabIndex        =   30
                  Top             =   525
                  Width           =   3435
               End
               Begin VB.Label LblVal 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Height          =   360
                  Index           =   0
                  Left            =   8580
                  RightToLeft     =   -1  'True
                  TabIndex        =   29
                  Top             =   90
                  Width           =   2535
               End
               Begin VB.Label LblDiscount 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Height          =   360
                  Index           =   0
                  Left            =   8580
                  RightToLeft     =   -1  'True
                  TabIndex        =   28
                  Top             =   525
                  Width           =   2535
               End
            End
            Begin MSComCtl2.DTPicker DtpTransDate 
               Height          =   300
               Index           =   0
               Left            =   2280
               TabIndex        =   8
               Top             =   30
               Width           =   2835
               _ExtentX        =   5001
               _ExtentY        =   529
               _Version        =   393216
               CalendarBackColor=   12648447
               Format          =   98959361
               CurrentDate     =   38718
               MaxDate         =   768106
            End
            Begin VSFlex8UCtl.VSFlexGrid FGSall 
               Height          =   5640
               Left            =   0
               TabIndex        =   9
               Top             =   345
               Width           =   13980
               _cx             =   24659
               _cy             =   9948
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
               SelectionMode   =   0
               GridLines       =   1
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   10
               Cols            =   12
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmDailtyReport.frx":5543
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
            Begin ImpulseButton.ISButton CmdShow 
               Height          =   300
               Index           =   0
               Left            =   0
               TabIndex        =   10
               Top             =   30
               Width           =   2280
               _ExtentX        =   4022
               _ExtentY        =   529
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   "⁄—÷ «·»Ì«‰« "
               BackColor       =   12634304
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ColorButton     =   12634304
               DrawFocusRectangle=   0   'False
            End
            Begin VB.Label LblTitle 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "›Ê« Ì— «·„»Ì⁄«  «· Ï  „  «·ÌÊ„"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   7
               Left            =   5115
               RightToLeft     =   -1  'True
               TabIndex        =   11
               Top             =   30
               Width           =   8865
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   7020
            Index           =   4
            Left            =   16530
            TabIndex        =   36
            TabStop         =   0   'False
            Top             =   45
            Width           =   13980
            _cx             =   24659
            _cy             =   12383
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
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   990
               Index           =   17
               Left            =   0
               TabIndex        =   111
               TabStop         =   0   'False
               Top             =   6000
               Width           =   13980
               _cx             =   24659
               _cy             =   1746
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
               Begin ImpulseButton.ISButton CmdHelp 
                  Height          =   195
                  Index           =   8
                  Left            =   0
                  TabIndex        =   112
                  Top             =   420
                  Width           =   435
                  _ExtentX        =   767
                  _ExtentY        =   344
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
                  ButtonImage     =   "FrmDailtyReport.frx":572E
                  ColorButton     =   14871017
                  ColorHighlight  =   16777215
                  ColorHoverText  =   16711680
                  ColorShadow     =   -2147483637
                  ColorOutline    =   0
                  DrawFocusRectangle=   0   'False
                  ColorToggledHoverText=   16711680
                  ColorTextShadow =   -2147483637
               End
               Begin ImpulseButton.ISButton CmdReturnSalling 
                  Height          =   210
                  Index           =   1
                  Left            =   0
                  TabIndex        =   113
                  Top             =   210
                  Visible         =   0   'False
                  Width           =   435
                  _ExtentX        =   767
                  _ExtentY        =   370
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
                  BackStyle       =   0
                  ButtonImage     =   "FrmDailtyReport.frx":5AC8
                  ColorButton     =   14871017
                  ColorHighlight  =   16777215
                  ColorHoverText  =   16711680
                  ColorOutline    =   0
                  DrawFocusRectangle=   0   'False
                  ColorToggledHoverText=   16711680
                  LowerToggledContent=   0   'False
                  ColorTextShadow =   -2147483637
               End
               Begin ImpulseButton.ISButton CmdReturnSalling 
                  Height          =   180
                  Index           =   0
                  Left            =   210
                  TabIndex        =   114
                  Top             =   30
                  Width           =   1725
                  _ExtentX        =   3043
                  _ExtentY        =   318
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   "⁄—÷ «· ﬁ—Ì—"
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
                  ButtonImage     =   "FrmDailtyReport.frx":5E62
                  ColorButton     =   14871017
                  ColorHighlight  =   16777215
                  ColorHoverText  =   16711680
                  ColorOutline    =   0
                  DrawFocusRectangle=   0   'False
                  ColorToggledHoverText=   16711680
                  LowerToggledContent=   0   'False
                  ColorTextShadow =   -2147483637
               End
               Begin VB.Label LblVal 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Height          =   255
                  Index           =   8
                  Left            =   6210
                  RightToLeft     =   -1  'True
                  TabIndex        =   118
                  Top             =   60
                  Width           =   1950
               End
               Begin VB.Label LblBillCount 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Height          =   270
                  Index           =   7
                  Left            =   6450
                  RightToLeft     =   -1  'True
                  TabIndex        =   117
                  Top             =   345
                  Width           =   1710
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   " ≈Ã„«·Ì ⁄œœ «·›Ê« Ì— :"
                  Height          =   180
                  Index           =   18
                  Left            =   8370
                  RightToLeft     =   -1  'True
                  TabIndex        =   116
                  Top             =   420
                  Width           =   1710
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   " ≈Ã„«·Ì ﬁÌ„… «·„— Ã⁄«  :"
                  Height          =   195
                  Index           =   15
                  Left            =   7935
                  RightToLeft     =   -1  'True
                  TabIndex        =   115
                  Top             =   60
                  Width           =   2385
               End
            End
            Begin ImpulseButton.ISButton CmdShow 
               Height          =   300
               Index           =   7
               Left            =   0
               TabIndex        =   119
               Top             =   30
               Width           =   2280
               _ExtentX        =   4022
               _ExtentY        =   529
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   "⁄—÷ «·»Ì«‰« "
               BackColor       =   12634304
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ColorButton     =   12634304
               DrawFocusRectangle=   0   'False
            End
            Begin MSComCtl2.DTPicker DtpTransDate 
               Height          =   300
               Index           =   7
               Left            =   2280
               TabIndex        =   120
               Top             =   30
               Width           =   2850
               _ExtentX        =   5027
               _ExtentY        =   529
               _Version        =   393216
               CalendarBackColor=   12648447
               Format          =   98959361
               CurrentDate     =   38718
               MaxDate         =   768106
            End
            Begin VSFlex8UCtl.VSFlexGrid FgReturnSalling 
               Height          =   5640
               Left            =   0
               TabIndex        =   121
               Top             =   345
               Width           =   13980
               _cx             =   24659
               _cy             =   9948
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
               SelectionMode   =   0
               GridLines       =   1
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   10
               Cols            =   7
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmDailtyReport.frx":61FC
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
            Begin VB.Label LblTitle 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "›Ê« Ì— „— Ã⁄ «·„»Ì⁄«  «· Ì  „  «·ÌÊ„"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   300
               Index           =   1
               Left            =   5130
               RightToLeft     =   -1  'True
               TabIndex        =   122
               Top             =   30
               Width           =   8850
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   7020
            Index           =   2
            Left            =   16830
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   45
            Width           =   13980
            _cx             =   24659
            _cy             =   12383
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
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   990
               Index           =   8
               Left            =   0
               TabIndex        =   38
               TabStop         =   0   'False
               Top             =   6000
               Width           =   13980
               _cx             =   24659
               _cy             =   1746
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
               Begin ImpulseButton.ISButton CmdHelp 
                  Height          =   195
                  Index           =   2
                  Left            =   210
                  TabIndex        =   39
                  Top             =   435
                  Width           =   1305
                  _ExtentX        =   2302
                  _ExtentY        =   344
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
                  ButtonImage     =   "FrmDailtyReport.frx":631F
                  ColorButton     =   14871017
                  ColorHighlight  =   16777215
                  ColorHoverText  =   16711680
                  ColorShadow     =   -2147483637
                  ColorOutline    =   0
                  DrawFocusRectangle=   0   'False
                  ColorToggledHoverText=   16711680
                  ColorTextShadow =   -2147483637
               End
               Begin ImpulseButton.ISButton CmdReturn 
                  Height          =   225
                  Index           =   1
                  Left            =   210
                  TabIndex        =   40
                  Top             =   210
                  Visible         =   0   'False
                  Width           =   1305
                  _ExtentX        =   2302
                  _ExtentY        =   397
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
                  BackStyle       =   0
                  ButtonImage     =   "FrmDailtyReport.frx":66B9
                  ColorButton     =   14871017
                  ColorHighlight  =   16777215
                  ColorHoverText  =   16711680
                  ColorOutline    =   0
                  DrawFocusRectangle=   0   'False
                  ColorToggledHoverText=   16711680
                  LowerToggledContent=   0   'False
                  ColorTextShadow =   -2147483637
               End
               Begin ImpulseButton.ISButton CmdReturn 
                  Height          =   180
                  Index           =   0
                  Left            =   210
                  TabIndex        =   41
                  Top             =   30
                  Width           =   1725
                  _ExtentX        =   3043
                  _ExtentY        =   318
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   "⁄—÷ «· ﬁ—Ì—"
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
                  ButtonImage     =   "FrmDailtyReport.frx":6A53
                  ColorButton     =   14871017
                  ColorHighlight  =   16777215
                  ColorHoverText  =   16711680
                  ColorOutline    =   0
                  DrawFocusRectangle=   0   'False
                  ColorToggledHoverText=   16711680
                  LowerToggledContent=   0   'False
                  ColorTextShadow =   -2147483637
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   " ≈Ã„«·Ì ﬁÌ„… «·„— Ã⁄«  :"
                  Height          =   195
                  Index           =   10
                  Left            =   7080
                  RightToLeft     =   -1  'True
                  TabIndex        =   45
                  Top             =   60
                  Width           =   2355
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   " ≈Ã„«·Ì ⁄œœ «·›Ê« Ì— :"
                  Height          =   195
                  Index           =   9
                  Left            =   7080
                  RightToLeft     =   -1  'True
                  TabIndex        =   44
                  Top             =   345
                  Width           =   2355
               End
               Begin VB.Label LblBillCount 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Height          =   285
                  Index           =   2
                  Left            =   5790
                  RightToLeft     =   -1  'True
                  TabIndex        =   43
                  Top             =   345
                  Width           =   1065
               End
               Begin VB.Label LblVal 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Height          =   255
                  Index           =   2
                  Left            =   5790
                  RightToLeft     =   -1  'True
                  TabIndex        =   42
                  Top             =   60
                  Width           =   1065
               End
            End
            Begin ImpulseButton.ISButton CmdShow 
               Height          =   300
               Index           =   2
               Left            =   0
               TabIndex        =   46
               Top             =   30
               Width           =   2280
               _ExtentX        =   4022
               _ExtentY        =   529
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   "⁄—÷ «·»Ì«‰« "
               BackColor       =   12634304
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ColorButton     =   12634304
               DrawFocusRectangle=   0   'False
            End
            Begin MSComCtl2.DTPicker DtpTransDate 
               Height          =   300
               Index           =   2
               Left            =   2280
               TabIndex        =   47
               Top             =   30
               Width           =   3705
               _ExtentX        =   6535
               _ExtentY        =   529
               _Version        =   393216
               CalendarBackColor=   12648447
               Format          =   98959361
               CurrentDate     =   38718
               MaxDate         =   768106
            End
            Begin VSFlex8UCtl.VSFlexGrid FGReturn 
               Height          =   5640
               Left            =   0
               TabIndex        =   48
               Top             =   345
               Width           =   13980
               _cx             =   24659
               _cy             =   9948
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
               SelectionMode   =   0
               GridLines       =   1
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   10
               Cols            =   7
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmDailtyReport.frx":6DED
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
            Begin VB.Label LblTitle 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "›Ê« Ì— „— Ã⁄ «·„‘ —Ì«  «· Ì  „  «·ÌÊ„"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   300
               Index           =   6
               Left            =   5985
               RightToLeft     =   -1  'True
               TabIndex        =   49
               Top             =   30
               Width           =   7995
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   7020
            Index           =   3
            Left            =   17130
            TabIndex        =   50
            TabStop         =   0   'False
            Top             =   45
            Width           =   13980
            _cx             =   24659
            _cy             =   12383
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
            _GridInfo       =   $"FrmDailtyReport.frx":6F10
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   9
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   1080
               Index           =   5
               Left            =   30
               TabIndex        =   51
               TabStop         =   0   'False
               Top             =   5910
               Width           =   13920
               _cx             =   24553
               _cy             =   1905
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
               Begin ImpulseButton.ISButton CmdHelp 
                  Height          =   300
                  Index           =   3
                  Left            =   240
                  TabIndex        =   52
                  Top             =   675
                  Width           =   2430
                  _ExtentX        =   4286
                  _ExtentY        =   529
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
                  ButtonImage     =   "FrmDailtyReport.frx":6F91
                  ColorButton     =   14871017
                  ColorHighlight  =   16777215
                  ColorHoverText  =   16711680
                  ColorShadow     =   -2147483637
                  ColorOutline    =   0
                  DrawFocusRectangle=   0   'False
                  ColorToggledHoverText=   16711680
                  ColorTextShadow =   -2147483637
               End
               Begin ImpulseButton.ISButton XPBtn_M_Print 
                  Height          =   300
                  Index           =   0
                  Left            =   240
                  TabIndex        =   53
                  Top             =   45
                  Width           =   2430
                  _ExtentX        =   4286
                  _ExtentY        =   529
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   "⁄—÷ «· ﬁ—Ì—"
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
                  ButtonImage     =   "FrmDailtyReport.frx":732B
                  ColorButton     =   14871017
                  ColorHighlight  =   16777215
                  ColorHoverText  =   16711680
                  ColorOutline    =   0
                  DrawFocusRectangle=   0   'False
                  ColorToggledHoverText=   16711680
                  LowerToggledContent=   0   'False
                  ColorTextShadow =   -2147483637
               End
               Begin ImpulseButton.ISButton XPBtn_M_Print 
                  Height          =   330
                  Index           =   1
                  Left            =   240
                  TabIndex        =   54
                  Top             =   345
                  Width           =   2430
                  _ExtentX        =   4286
                  _ExtentY        =   582
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
                  BackStyle       =   0
                  ButtonImage     =   "FrmDailtyReport.frx":76C5
                  ColorButton     =   14871017
                  ColorHighlight  =   16777215
                  ColorHoverText  =   16711680
                  ColorOutline    =   0
                  DrawFocusRectangle=   0   'False
                  ColorToggledHoverText=   16711680
                  LowerToggledContent=   0   'False
                  ColorTextShadow =   -2147483637
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   " ≈Ã„«·Ì ⁄„·Ì«  «·’Ì«‰… :"
                  Height          =   300
                  Index           =   11
                  Left            =   13080
                  RightToLeft     =   -1  'True
                  TabIndex        =   58
                  Top             =   105
                  Width           =   4395
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   " ≈Ã„«·Ì ⁄œœ «·›Ê« Ì— :"
                  Height          =   285
                  Index           =   12
                  Left            =   13335
                  RightToLeft     =   -1  'True
                  TabIndex        =   57
                  Top             =   540
                  Width           =   4140
               End
               Begin VB.Label LblVal 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Height          =   390
                  Index           =   3
                  Left            =   10635
                  RightToLeft     =   -1  'True
                  TabIndex        =   56
                  Top             =   105
                  Width           =   2055
               End
               Begin VB.Label LblBillCount 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Height          =   435
                  Index           =   3
                  Left            =   10635
                  RightToLeft     =   -1  'True
                  TabIndex        =   55
                  Top             =   540
                  Width           =   2055
               End
            End
            Begin ImpulseButton.ISButton CmdShow 
               Height          =   330
               Index           =   3
               Left            =   30
               TabIndex        =   59
               Top             =   30
               Width           =   1395
               _ExtentX        =   2461
               _ExtentY        =   582
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   "⁄—÷ «·»Ì«‰« "
               BackColor       =   12634304
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ColorButton     =   12634304
               DrawFocusRectangle=   0   'False
            End
            Begin MSComCtl2.DTPicker DtpTransDate 
               Height          =   330
               Index           =   3
               Left            =   1440
               TabIndex        =   60
               Top             =   30
               Width           =   4155
               _ExtentX        =   7329
               _ExtentY        =   582
               _Version        =   393216
               CalendarBackColor=   12648447
               Format          =   98959361
               CurrentDate     =   38718
            End
            Begin VSFlex8UCtl.VSFlexGrid FGMaintence 
               Height          =   5520
               Left            =   30
               TabIndex        =   61
               Top             =   375
               Width           =   13920
               _cx             =   24553
               _cy             =   9737
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
               SelectionMode   =   0
               GridLines       =   1
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   10
               Cols            =   5
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmDailtyReport.frx":7A5F
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
            Begin VB.Label LblTitle 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "›Ê« Ì— «·’Ì«‰… «· Ï  „  «·ÌÊ„"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   330
               Index           =   15
               Left            =   5610
               RightToLeft     =   -1  'True
               TabIndex        =   62
               Top             =   30
               Width           =   8340
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   7020
            Index           =   9
            Left            =   17430
            TabIndex        =   63
            TabStop         =   0   'False
            Top             =   45
            Width           =   13980
            _cx             =   24659
            _cy             =   12383
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
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   990
               Index           =   10
               Left            =   0
               TabIndex        =   64
               TabStop         =   0   'False
               Top             =   6000
               Width           =   13980
               _cx             =   24659
               _cy             =   1746
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
               Begin ImpulseButton.ISButton CmdHelp 
                  Height          =   405
                  Index           =   4
                  Left            =   360
                  TabIndex        =   65
                  Top             =   915
                  Width           =   2310
                  _ExtentX        =   4075
                  _ExtentY        =   714
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
                  ButtonImage     =   "FrmDailtyReport.frx":7B31
                  ColorButton     =   14871017
                  ColorHighlight  =   16777215
                  ColorHoverText  =   16711680
                  ColorShadow     =   -2147483637
                  ColorOutline    =   0
                  DrawFocusRectangle=   0   'False
                  ColorToggledHoverText=   16711680
                  ColorTextShadow =   -2147483637
               End
               Begin ImpulseButton.ISButton CmdExpenses 
                  Height          =   405
                  Index           =   0
                  Left            =   360
                  TabIndex        =   66
                  Top             =   60
                  Width           =   2670
                  _ExtentX        =   4710
                  _ExtentY        =   714
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   "⁄—÷ «· ﬁ—Ì—"
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
                  ButtonImage     =   "FrmDailtyReport.frx":7ECB
                  ColorButton     =   14871017
                  ColorHighlight  =   16777215
                  ColorHoverText  =   16711680
                  ColorOutline    =   0
                  DrawFocusRectangle=   0   'False
                  ColorToggledHoverText=   16711680
                  LowerToggledContent=   0   'False
                  ColorTextShadow =   -2147483637
               End
               Begin ImpulseButton.ISButton CmdExpenses 
                  Height          =   450
                  Index           =   1
                  Left            =   360
                  TabIndex        =   67
                  Top             =   465
                  Visible         =   0   'False
                  Width           =   2310
                  _ExtentX        =   4075
                  _ExtentY        =   794
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
                  ButtonImage     =   "FrmDailtyReport.frx":8265
                  ColorButton     =   14871017
                  ColorHighlight  =   16777215
                  ColorHoverText  =   16711680
                  ColorOutline    =   0
                  DrawFocusRectangle=   0   'False
                  ColorToggledHoverText=   16711680
                  LowerToggledContent=   0   'False
                  ColorTextShadow =   -2147483637
               End
               Begin VB.Label LblVal 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Height          =   525
                  Index           =   4
                  Left            =   10995
                  RightToLeft     =   -1  'True
                  TabIndex        =   71
                  Top             =   135
                  Width           =   2295
               End
               Begin VB.Label LblBillCount 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Height          =   585
                  Index           =   4
                  Left            =   10995
                  RightToLeft     =   -1  'True
                  TabIndex        =   70
                  Top             =   735
                  Width           =   2295
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   " ≈Ã„«·Ì «·„’—Ê›«  :"
                  Height          =   405
                  Index           =   16
                  Left            =   13290
                  RightToLeft     =   -1  'True
                  TabIndex        =   69
                  Top             =   135
                  Width           =   3765
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   " ⁄œœ «·⁄„·Ì«  :"
                  Height          =   390
                  Index           =   17
                  Left            =   14025
                  RightToLeft     =   -1  'True
                  TabIndex        =   68
                  Top             =   735
                  Width           =   3030
               End
            End
            Begin ImpulseButton.ISButton CmdShow 
               Height          =   300
               Index           =   4
               Left            =   0
               TabIndex        =   72
               Top             =   30
               Width           =   2280
               _ExtentX        =   4022
               _ExtentY        =   529
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   "⁄—÷ «·»Ì«‰« "
               BackColor       =   12634304
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ColorButton     =   12634304
               DrawFocusRectangle=   0   'False
            End
            Begin MSComCtl2.DTPicker DtpTransDate 
               Height          =   300
               Index           =   4
               Left            =   2280
               TabIndex        =   73
               Top             =   30
               Width           =   3705
               _ExtentX        =   6535
               _ExtentY        =   529
               _Version        =   393216
               CalendarBackColor=   12648447
               Format          =   98959361
               CurrentDate     =   38718
               MaxDate         =   768106
            End
            Begin VSFlex8UCtl.VSFlexGrid FGExpenses 
               Height          =   5640
               Left            =   0
               TabIndex        =   74
               Top             =   345
               Width           =   13980
               _cx             =   24659
               _cy             =   9948
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
               SelectionMode   =   0
               GridLines       =   1
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   10
               Cols            =   6
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmDailtyReport.frx":85FF
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
            Begin VB.Label LblTitle 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„’—Ê›«  «· Ì  „  «·ÌÊ„"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   300
               Index           =   18
               Left            =   5985
               RightToLeft     =   -1  'True
               TabIndex        =   75
               Top             =   30
               Width           =   7995
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   7020
            Index           =   11
            Left            =   17730
            TabIndex        =   76
            TabStop         =   0   'False
            Top             =   45
            Width           =   13980
            _cx             =   24659
            _cy             =   12383
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
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   990
               Index           =   15
               Left            =   0
               TabIndex        =   77
               TabStop         =   0   'False
               Top             =   6000
               Width           =   13980
               _cx             =   24659
               _cy             =   1746
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
               Begin ImpulseButton.ISButton CmdHelp 
                  Height          =   405
                  Index           =   7
                  Left            =   30
                  TabIndex        =   78
                  Top             =   900
                  Width           =   180
                  _ExtentX        =   318
                  _ExtentY        =   714
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
                  ButtonImage     =   "FrmDailtyReport.frx":86F7
                  ColorButton     =   14871017
                  ColorHighlight  =   16777215
                  ColorHoverText  =   16711680
                  ColorShadow     =   -2147483637
                  ColorOutline    =   0
                  DrawFocusRectangle=   0   'False
                  ColorToggledHoverText=   16711680
                  ColorTextShadow =   -2147483637
               End
               Begin ImpulseButton.ISButton CmdPayment 
                  Height          =   405
                  Index           =   0
                  Left            =   30
                  TabIndex        =   79
                  Top             =   60
                  Width           =   2640
                  _ExtentX        =   4657
                  _ExtentY        =   714
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   "⁄—÷ «· ﬁ—Ì—"
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
                  ButtonImage     =   "FrmDailtyReport.frx":8A91
                  ColorButton     =   14871017
                  ColorHighlight  =   16777215
                  ColorHoverText  =   16711680
                  ColorOutline    =   0
                  DrawFocusRectangle=   0   'False
                  ColorToggledHoverText=   16711680
                  LowerToggledContent=   0   'False
                  ColorTextShadow =   -2147483637
               End
               Begin ImpulseButton.ISButton CmdPayment 
                  Height          =   435
                  Index           =   1
                  Left            =   30
                  TabIndex        =   80
                  Top             =   465
                  Visible         =   0   'False
                  Width           =   2640
                  _ExtentX        =   4657
                  _ExtentY        =   767
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
                  ButtonImage     =   "FrmDailtyReport.frx":8E2B
                  ColorButton     =   14871017
                  ColorHighlight  =   16777215
                  ColorHoverText  =   16711680
                  ColorOutline    =   0
                  DrawFocusRectangle=   0   'False
                  ColorToggledHoverText=   16711680
                  LowerToggledContent=   0   'False
                  ColorTextShadow =   -2147483637
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "⁄œœ «·⁄„·Ì«  :"
                  Height          =   375
                  Index           =   6
                  Left            =   15645
                  RightToLeft     =   -1  'True
                  TabIndex        =   84
                  Top             =   720
                  Width           =   2235
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   " ≈Ã„«·Ì «·„œ›Ê⁄«  :"
                  Height          =   405
                  Index           =   7
                  Left            =   15090
                  RightToLeft     =   -1  'True
                  TabIndex        =   83
                  Top             =   300
                  Width           =   2790
               End
               Begin VB.Label LblBillCount 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Height          =   405
                  Index           =   6
                  Left            =   12000
                  RightToLeft     =   -1  'True
                  TabIndex        =   82
                  Top             =   795
                  Width           =   2640
               End
               Begin VB.Label LblVal 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Height          =   405
                  Index           =   6
                  Left            =   11505
                  RightToLeft     =   -1  'True
                  TabIndex        =   81
                  Top             =   315
                  Width           =   3615
               End
            End
            Begin ImpulseButton.ISButton CmdShow 
               Height          =   300
               Index           =   6
               Left            =   0
               TabIndex        =   85
               Top             =   30
               Width           =   2280
               _ExtentX        =   4022
               _ExtentY        =   529
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   "⁄—÷ «·»Ì«‰« "
               BackColor       =   12634304
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ColorButton     =   12634304
               DrawFocusRectangle=   0   'False
            End
            Begin MSComCtl2.DTPicker DtpTransDate 
               Height          =   300
               Index           =   6
               Left            =   2280
               TabIndex        =   86
               Top             =   30
               Width           =   3705
               _ExtentX        =   6535
               _ExtentY        =   529
               _Version        =   393216
               CalendarBackColor=   12648447
               Format          =   98959361
               CurrentDate     =   38718
               MaxDate         =   768106
            End
            Begin VSFlex8UCtl.VSFlexGrid FgPayments 
               Height          =   5640
               Left            =   0
               TabIndex        =   87
               Top             =   330
               Width           =   13980
               _cx             =   24659
               _cy             =   9948
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
               SelectionMode   =   0
               GridLines       =   1
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   10
               Cols            =   10
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmDailtyReport.frx":91C5
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
            Begin VB.Label LblTitle 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„œ›Ê⁄«  «· Ì  „  «·ÌÊ„"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   300
               Index           =   0
               Left            =   5985
               RightToLeft     =   -1  'True
               TabIndex        =   88
               Top             =   30
               Width           =   7995
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   7020
            Index           =   12
            Left            =   18030
            TabIndex        =   89
            TabStop         =   0   'False
            Top             =   45
            Width           =   13980
            _cx             =   24659
            _cy             =   12383
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
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   990
               Index           =   13
               Left            =   0
               TabIndex        =   90
               TabStop         =   0   'False
               Top             =   6000
               Width           =   13980
               _cx             =   24659
               _cy             =   1746
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
               Begin ImpulseButton.ISButton CmdHelp 
                  Height          =   405
                  Index           =   5
                  Left            =   30
                  TabIndex        =   91
                  Top             =   900
                  Width           =   210
                  _ExtentX        =   370
                  _ExtentY        =   714
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
                  ButtonImage     =   "FrmDailtyReport.frx":9387
                  ColorButton     =   14871017
                  ColorHighlight  =   16777215
                  ColorHoverText  =   16711680
                  ColorShadow     =   -2147483637
                  ColorOutline    =   0
                  DrawFocusRectangle=   0   'False
                  ColorToggledHoverText=   16711680
                  ColorTextShadow =   -2147483637
               End
               Begin ImpulseButton.ISButton CmdCaching 
                  Height          =   405
                  Index           =   0
                  Left            =   30
                  TabIndex        =   92
                  Top             =   60
                  Width           =   2670
                  _ExtentX        =   4710
                  _ExtentY        =   714
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   "⁄—÷ «· ﬁ—Ì—"
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
                  ButtonImage     =   "FrmDailtyReport.frx":9721
                  ColorButton     =   14871017
                  ColorHighlight  =   16777215
                  ColorHoverText  =   16711680
                  ColorOutline    =   0
                  DrawFocusRectangle=   0   'False
                  ColorToggledHoverText=   16711680
                  LowerToggledContent=   0   'False
                  ColorTextShadow =   -2147483637
               End
               Begin ImpulseButton.ISButton CmdCaching 
                  Height          =   390
                  Index           =   1
                  Left            =   30
                  TabIndex        =   93
                  Top             =   465
                  Visible         =   0   'False
                  Width           =   2670
                  _ExtentX        =   4710
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
                  ButtonImage     =   "FrmDailtyReport.frx":9ABB
                  ColorButton     =   14871017
                  ColorHighlight  =   16777215
                  ColorHoverText  =   16711680
                  ColorOutline    =   0
                  DrawFocusRectangle=   0   'False
                  ColorToggledHoverText=   16711680
                  LowerToggledContent=   0   'False
                  ColorTextShadow =   -2147483637
               End
               Begin VB.Label LblVal 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Height          =   405
                  Index           =   5
                  Left            =   12180
                  RightToLeft     =   -1  'True
                  TabIndex        =   97
                  Top             =   195
                  Width           =   2670
               End
               Begin VB.Label LblBillCount 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Height          =   405
                  Index           =   5
                  Left            =   12180
                  RightToLeft     =   -1  'True
                  TabIndex        =   96
                  Top             =   780
                  Width           =   2670
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   " ≈Ã„«·Ì «·„ﬁ»Ê÷«  :"
                  Height          =   405
                  Index           =   19
                  Left            =   14850
                  RightToLeft     =   -1  'True
                  TabIndex        =   95
                  Top             =   135
                  Width           =   3270
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "⁄œœ «·⁄„·Ì«  :"
                  Height          =   405
                  Index           =   20
                  Left            =   14850
                  RightToLeft     =   -1  'True
                  TabIndex        =   94
                  Top             =   720
                  Width           =   3270
               End
            End
            Begin ImpulseButton.ISButton CmdShow 
               Height          =   300
               Index           =   5
               Left            =   0
               TabIndex        =   98
               Top             =   30
               Width           =   2280
               _ExtentX        =   4022
               _ExtentY        =   529
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   "⁄—÷ «·»Ì«‰« "
               BackColor       =   12634304
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ColorButton     =   12634304
               DrawFocusRectangle=   0   'False
            End
            Begin MSComCtl2.DTPicker DtpTransDate 
               Height          =   300
               Index           =   5
               Left            =   2280
               TabIndex        =   99
               Top             =   30
               Width           =   3705
               _ExtentX        =   6535
               _ExtentY        =   529
               _Version        =   393216
               CalendarBackColor=   12648447
               Format          =   98959361
               CurrentDate     =   38718
               MaxDate         =   768106
            End
            Begin VSFlex8UCtl.VSFlexGrid FGCashing 
               Height          =   5640
               Left            =   0
               TabIndex        =   100
               Top             =   345
               Width           =   13980
               _cx             =   24659
               _cy             =   9948
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
               SelectionMode   =   0
               GridLines       =   1
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   10
               Cols            =   10
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmDailtyReport.frx":9E55
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
            Begin VB.Label LblTitle 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„ﬁ»Ê÷«  «· Ì  „  «·ÌÊ„"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   300
               Index           =   21
               Left            =   5985
               RightToLeft     =   -1  'True
               TabIndex        =   101
               Top             =   30
               Width           =   7995
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   7020
            Index           =   14
            Left            =   18330
            TabIndex        =   102
            TabStop         =   0   'False
            Top             =   45
            Width           =   13980
            _cx             =   24659
            _cy             =   12383
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
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   990
               Index           =   16
               Left            =   0
               TabIndex        =   103
               TabStop         =   0   'False
               Top             =   6000
               Width           =   13980
               _cx             =   24659
               _cy             =   1746
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
               Begin ImpulseButton.ISButton CmdHelp 
                  Height          =   405
                  Index           =   6
                  Left            =   30
                  TabIndex        =   104
                  Top             =   900
                  Width           =   240
                  _ExtentX        =   423
                  _ExtentY        =   714
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
                  ButtonImage     =   "FrmDailtyReport.frx":A021
                  ColorButton     =   14871017
                  ColorHighlight  =   16777215
                  ColorHoverText  =   16711680
                  ColorShadow     =   4210752
                  ColorOutline    =   0
                  DrawFocusRectangle=   0   'False
                  ColorToggledHoverText=   16711680
                  ColorTextShadow =   4210752
               End
               Begin ImpulseButton.ISButton CmdSummery 
                  Height          =   405
                  Index           =   0
                  Left            =   765
                  TabIndex        =   105
                  Top             =   60
                  Width           =   2445
                  _ExtentX        =   4313
                  _ExtentY        =   714
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   "⁄—÷ «· ﬁ—Ì—"
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
                  ButtonImage     =   "FrmDailtyReport.frx":A3BB
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
               Begin ImpulseButton.ISButton CmdSummery 
                  Height          =   435
                  Index           =   1
                  Left            =   765
                  TabIndex        =   106
                  Top             =   465
                  Visible         =   0   'False
                  Width           =   2445
                  _ExtentX        =   4313
                  _ExtentY        =   767
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
                  ButtonImage     =   "FrmDailtyReport.frx":A755
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
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   " «·≈Ã„«·Ì :"
                  Height          =   405
                  Index           =   8
                  Left            =   15645
                  RightToLeft     =   -1  'True
                  TabIndex        =   108
                  Top             =   135
                  Width           =   2505
               End
               Begin VB.Label LblVal 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   12
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   405
                  Index           =   7
                  Left            =   12735
                  RightToLeft     =   -1  'True
                  TabIndex        =   107
                  Top             =   195
                  Width           =   2910
               End
            End
            Begin VSFlex8UCtl.VSFlexGrid FGSummery 
               Height          =   5640
               Left            =   0
               TabIndex        =   109
               Top             =   345
               Width           =   13980
               _cx             =   24659
               _cy             =   9948
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
               SelectionMode   =   0
               GridLines       =   0
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   50
               Cols            =   3
               FixedRows       =   1
               FixedCols       =   0
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmDailtyReport.frx":AAEF
               ScrollTrack     =   0   'False
               ScrollBars      =   0
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
               ExplorerBar     =   2
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
            Begin VB.Label LblTitle 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„·Œ’ Õ—ﬂ«  «·ÌÊ„"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   300
               Index           =   24
               Left            =   5985
               RightToLeft     =   -1  'True
               TabIndex        =   110
               Top             =   30
               Width           =   7995
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   7020
            Index           =   22
            Left            =   19230
            TabIndex        =   142
            TabStop         =   0   'False
            Top             =   45
            Width           =   13980
            _cx             =   24659
            _cy             =   12383
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
               Height          =   300
               Index           =   3
               Left            =   0
               TabIndex        =   143
               Top             =   6525
               Width           =   2280
               _ExtentX        =   4022
               _ExtentY        =   529
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   "⁄—÷ «· ﬁ—Ì—"
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
               ButtonImage     =   "FrmDailtyReport.frx":AB59
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
               Height          =   4965
               Left            =   0
               TabIndex        =   144
               Top             =   450
               Width           =   13980
               _cx             =   24659
               _cy             =   8758
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
               AllowUserResizing=   1
               SelectionMode   =   0
               GridLines       =   1
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   50
               Cols            =   9
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmDailtyReport.frx":AEF3
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
            Begin MSDataListLib.DataCombo DcbBox 
               Height          =   315
               Left            =   4290
               TabIndex        =   145
               Top             =   0
               Width           =   3705
               _ExtentX        =   6535
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin ImpulseButton.ISButton CmdShow 
               Height          =   300
               Index           =   10
               Left            =   0
               TabIndex        =   146
               Top             =   0
               Width           =   2280
               _ExtentX        =   4022
               _ExtentY        =   529
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   "⁄—÷ «·»Ì«‰« "
               BackColor       =   12634304
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ColorButton     =   12634304
               DrawFocusRectangle=   0   'False
            End
            Begin MSComCtl2.DTPicker DtpTransDate 
               Height          =   300
               Index           =   9
               Left            =   2280
               TabIndex        =   147
               Top             =   0
               Width           =   2010
               _ExtentX        =   3545
               _ExtentY        =   529
               _Version        =   393216
               CalendarBackColor=   12648447
               Format          =   98959361
               CurrentDate     =   38718
               MaxDate         =   768106
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "0"
               ForeColor       =   &H00000080&
               Height          =   300
               Index           =   31
               Left            =   5985
               RightToLeft     =   -1  'True
               TabIndex        =   156
               Top             =   6105
               Width           =   1425
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "0"
               ForeColor       =   &H00000080&
               Height          =   300
               Index           =   30
               Left            =   7695
               RightToLeft     =   -1  'True
               TabIndex        =   155
               Top             =   6105
               Width           =   1425
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "0"
               ForeColor       =   &H00000080&
               Height          =   300
               Index           =   29
               Left            =   9690
               RightToLeft     =   -1  'True
               TabIndex        =   154
               Top             =   6105
               Width           =   1440
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·«Ã„«·Ì"
               ForeColor       =   &H00000080&
               Height          =   315
               Index           =   28
               Left            =   5985
               RightToLeft     =   -1  'True
               TabIndex        =   153
               Top             =   5655
               Width           =   1425
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "œ«∆‰"
               ForeColor       =   &H00000080&
               Height          =   315
               Index           =   27
               Left            =   7695
               RightToLeft     =   -1  'True
               TabIndex        =   152
               Top             =   5655
               Width           =   1425
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„œÌ‰"
               ForeColor       =   &H00000080&
               Height          =   315
               Index           =   26
               Left            =   9690
               RightToLeft     =   -1  'True
               TabIndex        =   151
               Top             =   5655
               Width           =   1440
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·Õ—ﬂ… ··ÌÊ„ ›ﬁÿ"
               ForeColor       =   &H00000080&
               Height          =   315
               Index           =   25
               Left            =   11700
               RightToLeft     =   -1  'True
               TabIndex        =   150
               Top             =   5880
               Width           =   2010
            End
            Begin VB.Label LblTitle 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„·Œ’ Õ—ﬂ… «·Œ“‰… «·ÌÊ„"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   300
               Index           =   4
               Left            =   9690
               RightToLeft     =   -1  'True
               TabIndex        =   149
               Top             =   0
               Width           =   4290
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«Œ — «”„ «·Œ“‰…"
               ForeColor       =   &H00000080&
               Height          =   300
               Index           =   24
               Left            =   7995
               RightToLeft     =   -1  'True
               TabIndex        =   148
               Top             =   0
               Width           =   1695
            End
         End
      End
   End
End
Attribute VB_Name = "FrmDailtyReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim StaticStrSql As String
Dim salesSql As String
Dim REsalesSql As String
Dim PurcahseSQL As String
Dim REPurcahseSQL As String



Private Sub Cmd_Click(Index As Integer)
    Select Case Index
        Case 3
            print_reportBoxws
    End Select

End Sub

Private Sub CmdCaching_Click(Index As Integer)
    'On Error GoTo ErrTrap
    Dim StrSQL As String
    Dim Reports As ClsRepoerts
 '   StrSQL = "select * From CahingReport where NoteDate=#" & SQLDate(DtpTransDate(5).value) & "#"
    
    
       '    StrSQL = "select * From CahingReport where NoteDate=" & SQLDate(CDate(StrDate), True) & ""
          
       StrSQL = "SELECT       dbo.Notes.NoteID, dbo.Notes.NoteDate,Notes.PayDes ,dbo.Notes.Note_Value, dbo.Notes.CusID, dbo.TblCustemers.CusName, dbo.Notes.UserID, "
StrSQL = StrSQL + "                      dbo.TblUsers.UserName, dbo.Notes.CashingType, dbo.Notes.Remark, dbo.Transactions.Transaction_ID, dbo.Transactions.Transaction_Serial,"
StrSQL = StrSQL + "                      "
StrSQL = StrSQL + "                        CASE "
StrSQL = StrSQL + "                                  WHEN CashingType = 0 THEN '„‰ ⁄„Ì·'"
StrSQL = StrSQL + "                                  WHEN CashingType = 1 THEN '„‰ „Ê—œ'"
StrSQL = StrSQL + "                                  WHEN CashingType = 2 THEN '„ﬁ«Ê· »«ÿ‰'"
StrSQL = StrSQL + "                                  WHEN CashingType = 3 THEN '≈Ì—«œ«  ≈Œ—Ï'"
StrSQL = StrSQL + "                                  WHEN CashingType = 4 THEN '„œ›Ê⁄«  „ﬁœ„Â'"
StrSQL = StrSQL + "                                  WHEN CashingType = 5 THEN '„‘—Ê⁄'"
StrSQL = StrSQL + "                                  WHEN CashingType = 6 THEN '„‰ „ÊŸ›'"
StrSQL = StrSQL + "                                  WHEN CashingType = 7 THEN '„‰ Õ”«»'"
StrSQL = StrSQL + "                                  WHEN CashingType = 8 THEN '„‰ ›Ê« Ì— «·‰ﬁ·Ì« '"
StrSQL = StrSQL + "                                  WHEN CashingType = 9 THEN '„‰ ›« Ê—… «·’Ì«‰…'"
StrSQL = StrSQL + "                                  WHEN CashingType = 10 THEN '»‰«¡« ⁄·Ï ﬂ«—  ’Ì«‰…'"
StrSQL = StrSQL + "                                  WHEN CashingType = 11 THEN '„‰ ⁄œ… „” Œ·’« '"
StrSQL = StrSQL + "                                  WHEN CashingType = 12 THEN '»‰«¡« ⁄·Ï ⁄ﬁœ Õ«ÊÌ« '"
StrSQL = StrSQL + "                             END As TransactionTypeName"
       
StrSQL = StrSQL + "                              , dbo.TblBoxesData.BoxID, dbo.TblBoxesData.BoxName, dbo.transactions.Transaction_Type, dbo.Notes.RevenuesID, "
StrSQL = StrSQL + "                      dbo.TblRevenuesTypes.RevenuesName, dbo.Notes.NoteSerial, dbo.Notes.NoteSerial1, dbo.Notes.AccountsCode, dbo.ACCOUNTS.Account_Code,"
StrSQL = StrSQL + "                      dbo.ACCOUNTS.Account_Name, dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode, dbo.TblEmployee.Emp_Name,"
StrSQL = StrSQL + "                      dbo.TblEmployee.Fullcode AS EmployeeFullcode, dbo.TblEmployee.Emp_Namee, dbo.Notes.EmpId , dbo.Notes.PayDes ManulaNO, dbo.Notes.ManualNo"
StrSQL = StrSQL + " FROM         dbo.TblRevenuesTypes RIGHT OUTER JOIN"
StrSQL = StrSQL + "                      dbo.ACCOUNTS RIGHT OUTER JOIN"
StrSQL = StrSQL + " dbo.TblEmployee RIGHT OUTER JOIN"
StrSQL = StrSQL + "                      dbo.TblUsers INNER JOIN"
StrSQL = StrSQL + "                      dbo.Notes ON dbo.TblUsers.UserID = dbo.Notes.UserID ON dbo.TblEmployee.Emp_ID = dbo.Notes.EmpId ON"
StrSQL = StrSQL + "                      dbo.ACCOUNTS.Account_Code = dbo.Notes.AccountsCode ON dbo.TblRevenuesTypes.RevenuesID = dbo.Notes.RevenuesID LEFT OUTER JOIN"
StrSQL = StrSQL + "                      dbo.TblBoxesData ON dbo.Notes.BoxID = dbo.TblBoxesData.BoxID LEFT OUTER JOIN"
StrSQL = StrSQL + "                      dbo.TblCustemers ON dbo.Notes.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
 StrSQL = StrSQL + "                     dbo.Transactions LEFT OUTER JOIN"
StrSQL = StrSQL + "                      dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type ON"
 StrSQL = StrSQL + "                     dbo.Notes.Transaction_ID = dbo.Transactions.Transaction_ID"
 
                      
    StrSQL = StrSQL + " WHERE     (dbo.Notes.NoteType = 4)"
  StrSQL = StrSQL + " AND Notes.NoteDate=" & SQLDate(CDate(DtpTransDate(5).value), True)
     StrSQL = StrSQL + " Order by Notes.NoteSerial1"

 Set Reports = New ClsRepoerts
    Select Case Index

        Case 0
            Reports.CashingReports StrSQL, WindowTarget, " ﬁ«—Ì— «·„ﬁ»Ê÷« "

        Case 1
            Reports.CashingReports StrSQL, PrinterTarget, " ﬁ«—Ì— «·„ﬁ»Ê÷« "
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub CmdExpenses_Click(Index As Integer)
    On Error GoTo ErrTrap
    Dim StrSQL As String
    Dim Reports As ClsRepoerts
   ' StrSQL = "select * From ExpensesReport where NoteDate=#" & SQLDate(DtpTransDate(4).value) & "#"
    
    '      StrSQL = "  SELECT     TOP 100 PERCENT dbo.Notes.NoteID, dbo.Notes.NoteDate, dbo.Notes.NoteType, dbo.Notes.Note_Value, dbo.Notes.ExpensesID, dbo.ExpensesType.Name,"
    '    StrSQL = StrSQL + "dbo.Notes.NoteSerial1,   dbo.Notes.NoteSerial, dbo.Notes.Remark, dbo.TblUsers.UserName, dbo.Notes.BoxID, dbo.Notes.UserID, dbo.TblBoxesData.BoxName, dbo.BanksData.BankName,"
    '    StrSQL = StrSQL + "  dbo.Notes.BankID , dbo.Notes.ChqueNum, dbo.Notes.DueDate, dbo.Notes.NoteSerial1"
    '    StrSQL = StrSQL + "  FROM         dbo.TblUsers INNER JOIN"
    '    StrSQL = StrSQL + "  dbo.ExpensesType INNER JOIN"
    '    StrSQL = StrSQL + "   dbo.Notes ON dbo.ExpensesType.ID = dbo.Notes.ExpensesID ON dbo.TblUsers.UserID = dbo.Notes.UserID LEFT OUTER JOIN"
    '    StrSQL = StrSQL + "   dbo.TblBoxesData ON dbo.Notes.BoxID = dbo.TblBoxesData.BoxID LEFT OUTER JOIN"
    '    StrSQL = StrSQL + "   dbo.BanksData ON dbo.Notes.BankID = dbo.BanksData.BankID"
    '    StrSQL = StrSQL + "  Where (dbo.Notes.NoteType = 3)"
    '    StrSQL = StrSQL + "  and dbo.Notes.NoteDate=" & SQLDate(CDate(DtpTransDate(4).value), True) & ""
    '    StrSQL = StrSQL + "  ORDER BY dbo.Notes.NoteID"
 
 
    
    Set Reports = New ClsRepoerts

    Select Case Index

        Case 0
        print_reportExpenss , DtpTransDate(4).value
            'Reports.ExpensesReports StrSQL, WindowTarget

        Case 1
        print_reportExpenss , DtpTransDate(4).value
            'Reports.ExpensesReports StrSQL, PrinterTarget
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub CmdHelp_Click(Index As Integer)
    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
    SystemOptions.SysHelp.HHDisplayTopicID Me.hwnd
End Sub

Private Sub CmdPayment_Click(Index As Integer)
    On Error GoTo ErrTrap
    Dim StrSQL As String
    Dim Reports As ClsRepoerts
 '   StrSQL = "select * From PaymentsReport where NoteDate=#" & SQLDate(DtpTransDate(6).value) & "#"
         StrSQL = " SELECT     TOP 100 PERCENT dbo.Notes.NoteID, dbo.Notes.NoteDate, dbo.Notes.Note_Value, dbo.Notes.CusID, dbo.TblCustemers.CusName, dbo.Notes.UserID,"
        StrSQL = StrSQL + " dbo.TblUsers.UserName, dbo.Notes.CashingType, dbo.Notes.NoteSerial, dbo.Notes.Remark, dbo.Transactions.Transaction_ID, dbo.Transactions.Transaction_Serial,"
        StrSQL = StrSQL + " dbo.TransactionTypes.TransactionTypeName, dbo.TblBoxesData.BoxID, dbo.TblBoxesData.BoxName, dbo.Transactions.Transaction_Type, dbo.Notes.NoteSerial1,"
        StrSQL = StrSQL + " dbo.BanksData.BankName , dbo.Notes.ChqueNum, dbo.Notes.DueDate"
        StrSQL = StrSQL + " FROM         dbo.Transactions RIGHT OUTER JOIN"
        StrSQL = StrSQL + "  dbo.TblBoxesData RIGHT OUTER JOIN"
        StrSQL = StrSQL + "  dbo.TblUsers RIGHT OUTER JOIN"
        StrSQL = StrSQL + "   dbo.TblCustemers RIGHT OUTER JOIN"
        StrSQL = StrSQL + "   dbo.BanksData RIGHT OUTER JOIN"
        StrSQL = StrSQL + "   dbo.Notes ON dbo.BanksData.BankID = dbo.Notes.BankID ON dbo.TblCustemers.CusID = dbo.Notes.CusID ON dbo.TblUsers.UserID = dbo.Notes.UserID ON"
        StrSQL = StrSQL + "  dbo.TblBoxesData.BoxID = dbo.Notes.BoxID ON dbo.Transactions.Transaction_ID = dbo.Notes.Transaction_ID LEFT OUTER JOIN"
        StrSQL = StrSQL + "  dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type"
        StrSQL = StrSQL + "  where dbo.Notes.NoteDate =" & SQLDate(CDate(DtpTransDate(6).value), True) & ""
        StrSQL = StrSQL + "  AND (dbo.Notes.NoteType = 5)"
        StrSQL = StrSQL + "  ORDER BY dbo.Notes.NoteDate"
        
    
    
    
    
    
    Set Reports = New ClsRepoerts

    Select Case Index

        Case 0
            Reports.PaymentsDailyReport StrSQL, WindowTarget

        Case 1
            Reports.PaymentsDailyReport StrSQL, PrinterTarget
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub CmdPrint_Click(Index As Integer)
    On Error GoTo ErrTrap
  '  Dim StrSQL As String
  '  Dim Reports As ClsRepoerts
  '  StrSQL = "Select * From ReportSallingTime where Transaction_Date=" & SQLDate(DtpTransDate(0).value, True) & ""
  '  Set Reports = New ClsRepoerts

    Select Case Index

        Case 0
            print_reportSaling , DtpTransDate(0).value

        Case 1
            print_reportSaling , DtpTransDate(0).value
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub CmdPrintBuy_Click(Index As Integer)
    On Error GoTo ErrTrap
    Dim StrSQL As String
    Dim Reports As ClsRepoerts
    'StrSQL = "Select * From ReportBuyTime_Client where Transaction_Date=#" & SQLDate(DtpTransDate(1).value) & "#"
    'Set Reports = New ClsRepoerts

    Select Case Index

        Case 0

            print_reportPurchases , DtpTransDate(1).value

        Case 1
             print_reportPurchases , DtpTransDate(1).value

    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub CmdReturn_Click(Index As Integer)
    On Error GoTo ErrTrap
    Dim StrSQL As String
    Dim Reports As ClsRepoerts
   ' StrSQL = "SELECT * From QryReturn"
   ' StrSQL = StrSQL + " where Transaction_Date=#" & SQLDate(DtpTransDate(2).value) & "#"
   ' Set Reports = New ClsRepoerts

    Select Case Index

        Case 0
   '         Reports.DailyReturn StrSQL, WindowTarget
   print_reportReturnPurches , DtpTransDate(2).value

        Case 1
   '         Reports.DailyReturn StrSQL, PrinterTarget
   print_reportReturnPurches , DtpTransDate(2).value
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub CmdReturnSalling_Click(Index As Integer)
    'On Error GoTo ErrTrap
    Dim StrSQL As String
    Dim Reports As ClsRepoerts
    'StrSQL = "SELECT * From QryReturnSalling"
    'StrSQL = StrSQL + " where Transaction_Date=#" & SQLDate(DtpTransDate(7).value) & "#"
    'Set Reports = New ClsRepoerts

    Select Case Index

        Case 0
        print_reportReturnSaling , DtpTransDate(7).value
           ' Reports.DailyReturnSalling StrSQL, WindowTarget

        Case 1
        print_reportReturnSaling , DtpTransDate(7).value
           ' Reports.DailyReturnSalling StrSQL, PrinterTarget
            'DailyReturnSalling.rpt
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub CmdShow_Click(Index As Integer)
    On Error GoTo ErrTrap

    Select Case Index

        Case 0
            LoadSalling DtpTransDate(0).value

        Case 1
            LoadPurchase DtpTransDate(1).value

        Case 2
            LoadReturn DtpTransDate(2).value

        Case 3
            LoadMaintenence DtpTransDate(3).value

        Case 4
            LoadExpenses DtpTransDate(4).value

        Case 5
            LoadCashing DtpTransDate(5).value

        Case 6
            LoadPayments DtpTransDate(6).value

        Case 7
            LoadReturnSalling DtpTransDate(7).value

        Case 8
            LoadItemsTransactions

        Case 9
            'Load Customer_ID
            LoadCustomers
        Case 10
        loadBoxses DtpTransDate(9).value

    End Select

    LoadSummary
    Exit Sub
ErrTrap:
End Sub
Sub loadBoxses(StrDate As String)
Dim Rs2 As ADODB.Recordset
Dim i As Integer
Dim StrAccountCode As String
Set Rs2 = New ADODB.Recordset
Dim sql As String
Dim str As String
StrAccountCode = GetMyAccountCode("TblBoxesData", "BoxID", val(Me.DcbBox.BoundText), "Account_Code")
sql = " SELECT     dbo.Notes.ChqueNum, dbo.Notes.ManualNo, dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_ID, dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit,"
sql = sql & "                       dbo.DOUBLE_ENTREY_VOUCHERS.[Value] AS DEV_Value, dbo.DOUBLE_ENTREY_VOUCHERS.RecordDateH, dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code,"
sql = sql & "                      dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_Description AS DEV_DES,"
sql = sql & "                      dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_Descriptione AS DevDESE, dbo.ACCOUNTS.Account_Name,"
sql = sql & "                      dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No, dbo.TblNotesTypes.NotesTypeName, dbo.DOUBLE_ENTREY_VOUCHERS.UserID, dbo.TblUsers.UserName,"
sql = sql & "                      dbo.DOUBLE_ENTREY_VOUCHERS.RecordDate, dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID, dbo.DOUBLE_ENTREY_VOUCHERS.ReceiptID,"
sql = sql & "                      dbo.DOUBLE_ENTREY_VOUCHERS.OperaID, dbo.DOUBLE_ENTREY_VOUCHERS.Transaction_ID, dbo.Transactions.Transaction_Serial,"
sql = sql & "                      dbo.Transactions.Transaction_Date, dbo.TransactionTypes.TransactionTypeName, dbo.DOUBLE_ENTREY_VOUCHERS.Posted,"
sql = sql & "                      dbo.DOUBLE_ENTREY_VOUCHERS.PostedDate, dbo.DOUBLE_ENTREY_VOUCHERS.PostedUserID, dbo.DOUBLE_ENTREY_VOUCHERS.Account_Interval_ID,"
sql = sql & "                      dbo.Notes.NoteDate, dbo.Notes.NoteType, dbo.Notes.NoteSerial, dbo.Notes.Note_Value, dbo.ACCOUNTS.Account_Serial, dbo.ACCOUNTS.Account_NameEng,"
sql = sql & "                      dbo.ACCOUNTS.Parent_Account_Code, dbo.ACCOUNTS.opening_balance, dbo.ACCOUNTS.opening_balance_type, dbo.ACCOUNTS.Branch,"
sql = sql & "                      dbo.ACCOUNTS.Sum_account, dbo.ACCOUNTS.cost_center, dbo.ACCOUNTS.currenct_code, dbo.Notes.Remark, dbo.Notes.note_value_by_characters,"
sql = sql & "                      dbo.Notes.foxy_no, dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No1, dbo.DOUBLE_ENTREY_VOUCHERS.project_id, dbo.TblNotesTypes.NotesTypeNamee,"
sql = sql & "                      dbo.TransactionTypes.TransactionEnglishName, dbo.DOUBLE_ENTREY_VOUCHERS.opr_fullcode, dbo.Notes.NoteSerial1,"
sql = sql & "                      dbo.DOUBLE_ENTREY_VOUCHERS.branch_id, dbo.TblBranchesData.ActivityTypeId, dbo.DOUBLE_ENTREY_VOUCHERS.notes_all,"
sql = sql & "                      dbo.TblBranchesData.branch_name , dbo.TblBranchesData.branch_namee"
sql = sql & " FROM         dbo.TblBranchesData INNER JOIN"
sql = sql & "                      dbo.TblUsers INNER JOIN"
sql = sql & "                      dbo.ACCOUNTS INNER JOIN"
sql = sql & "                      dbo.DOUBLE_ENTREY_VOUCHERS ON dbo.ACCOUNTS.Account_Code = dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code ON"
sql = sql & "                      dbo.TblUsers.UserID = dbo.DOUBLE_ENTREY_VOUCHERS.UserID ON"
sql = sql & "                      dbo.TblBranchesData.branch_id = dbo.DOUBLE_ENTREY_VOUCHERS.branch_id LEFT OUTER JOIN"
sql = sql & "                      dbo.Notes LEFT OUTER JOIN"
sql = sql & "                     dbo.TblNotesTypes ON dbo.Notes.NoteType = dbo.TblNotesTypes.NotesType ON dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID = dbo.Notes.NoteID LEFT OUTER JOIN"
sql = sql & "                      dbo.Transactions ON dbo.DOUBLE_ENTREY_VOUCHERS.Transaction_ID = dbo.Transactions.Transaction_ID LEFT OUTER JOIN"
sql = sql & "                      dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type"
sql = sql + " WHERE     (dbo.ACCOUNTS.Account_Code = '" & StrAccountCode & "') "
sql = sql + " and RecordDate =" & SQLDate(DtpTransDate(9).value, True) & ""
Rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs2.RecordCount > 0 Then
With VSFlexGrid1
.Rows = Rs2.RecordCount + 2
Rs2.MoveFirst
For i = 1 To .Rows - 1
If i = 1 Then
If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(i, .ColIndex("Account_Name")) = IIf(IsNull(Rs2("Account_Name").value), "", Rs2("Account_Name").value)
.TextMatrix(i, .ColIndex("NotesTypeName")) = "—’Ìœ ”«»ﬁ"
.TextMatrix(i, .ColIndex("DEV_DES")) = "—’Ìœ ”«»ﬁ"
Else
.TextMatrix(i, .ColIndex("Account_Name")) = IIf(IsNull(Rs2("Account_NameEng").value), "", Rs2("Account_NameEng").value)
.TetMatrix(i, .ColIndex("NotesTypeName")) = "Previous Balance"
.TetMatrix(i, .ColIndex("DEV_DES")) = "Previous Balance"
End If

.TextMatrix(i, .ColIndex("RecordDate")) = DateAdd("d", -1, DtpTransDate(9).value)
.TextMatrix(i, .ColIndex("opening_balance")) = IIf(IsNull(Rs2("opening_balance").value), 0, Rs2("opening_balance").value)
If Rs2("opening_balance").value > 0 Then
.TextMatrix(i, .ColIndex("Dept")) = IIf(IsNull(Rs2("opening_balance").value), "", Rs2("opening_balance").value)
ElseIf Rs2("opening_balance").value < 0 Then
.TextMatrix(i, .ColIndex("Credit")) = IIf(IsNull(Rs2("opening_balance").value), "", Rs2("opening_balance").value)
End If
Else

If val(.TextMatrix(i, .ColIndex("Dept"))) - val(.TextMatrix(i, .ColIndex("Credit"))) > 0 Then
.TextMatrix(i, .ColIndex("opening_balance")) = (val(.TextMatrix(i, .ColIndex("Dept"))) - val(.TextMatrix(i, .ColIndex("Credit")))) + val(.TextMatrix(i, .ColIndex("opening_balance")))
Else
.TextMatrix(i, .ColIndex("opening_balance")) = (-1 * (val(.TextMatrix(i, .ColIndex("Credit"))))) + val(.TextMatrix(i, .ColIndex("opening_balance")))
End If
If Not IsNull(Rs2("ChqueNum").value) Then
If (Rs2("ChqueNum").value) <> "" Then
str = Rs2("DEV_DES").value & " —ﬁ„ «·„” ‰œ " & Rs2("NoteSerial1").value & " —ﬁ„ ÌœÊÌ  " & Rs2("ManualNo").value & "  ‘Ìﬂ /—ﬁ„ «·ÕÊ«·…   " & Rs2("ChqueNum").value
Else
str = Rs2("DEV_DES").value & " —ﬁ„ «·„” ‰œ " & Rs2("NoteSerial1").value & " —ﬁ„ ÌœÊÌ  " & Rs2("ManualNo").value
End If
Else
str = Rs2("DEV_DES").value & " —ﬁ„ «·„” ‰œ " & Rs2("NoteSerial1").value & " —ﬁ„ ÌœÊÌ  " & Rs2("ManualNo").value
End If
.TextMatrix(i, .ColIndex("RecordDate")) = IIf(IsNull(Rs2("RecordDate").value), "", Rs2("RecordDate").value)
.TextMatrix(i, .ColIndex("Serial")) = i
.TextMatrix(i, .ColIndex("NoteSerial")) = IIf(IsNull(Rs2("NoteSerial").value), "", Rs2("NoteSerial").value)
If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(i, .ColIndex("Account_Name")) = IIf(IsNull(Rs2("Account_Name").value), "", Rs2("Account_Name").value)
.TextMatrix(i, .ColIndex("NotesTypeName")) = IIf(IsNull(Rs2("NotesTypeName").value), "", Rs2("NotesTypeName").value)
Else
.TextMatrix(i, .ColIndex("Account_Name")) = IIf(IsNull(Rs2("Account_NameEng").value), "", Rs2("Account_NameEng").value)
.TetMatrix(i, .ColIndex("NotesTypeName")) = IIf(IsNull(Rs2("NotesTypeNamee").value), "", Rs2("NotesTypeNamee").value)
End If
.TextMatrix(i, .ColIndex("DEV_DES")) = str
If Not IsNull(Rs2("opening_balance").value) Then
If Rs2("Credit_Or_Debit").value = 0 Then
.TextMatrix(i, .ColIndex("Dept")) = IIf(IsNull(Rs2("DEV_Value").value), "", Rs2("DEV_Value").value)
Else
.TextMatrix(i, .ColIndex("Credit")) = IIf(IsNull(Rs2("DEV_Value").value), "", Rs2("DEV_Value").value)
End If
'.TextMatrix(i, .ColIndex("opening_balance")) = IIf(IsNull(Rs2("DEV_Value").value), "", Rs2("DEV_Value").value)
If val(.TextMatrix(i, .ColIndex("Dept"))) - val(.TextMatrix(i, .ColIndex("Credit"))) > 0 Then
.TextMatrix(i, .ColIndex("opening_balance")) = (val(.TextMatrix(i, .ColIndex("Dept"))) - val(.TextMatrix(i, .ColIndex("Credit")))) + val(.TextMatrix(i - 1, .ColIndex("opening_balance")))
Else
.TextMatrix(i, .ColIndex("opening_balance")) = (-1 * (val(.TextMatrix(i, .ColIndex("Credit"))))) + val(.TextMatrix(i - 1, .ColIndex("opening_balance")))
End If
End If
Rs2.MoveNext
End If
Next i
End With
End If
Relin
End Sub
Private Sub CmdSummery_Click(Index As Integer)
    On Error GoTo ErrTrap
    Dim Reports As ClsRepoerts
    Set Reports = New ClsRepoerts

    Select Case Index

        Case 0
            Reports.DailSummery WindowTarget

        Case 1
            Reports.DailSummery PrinterTarget
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub DtpTransDate_Change(Index As Integer)
    'On Error GoTo ErrTrap
    '
    'Select Case Index
    '   Case 0
    '        LoadSalling DtpTransDate(0).Value
    '    Case 1
    '        LoadPurchase DtpTransDate(1).Value
    '    Case 2
    '        LoadReturn DtpTransDate(2).Value
    '    Case 3
    '        LoadMaintenence DtpTransDate(3).Value
    '    Case 4
    '        LoadExpenses DtpTransDate(4).Value
    '    Case 5
    '        LoadCashing DtpTransDate(5).Value
    '    Case 6
    '        LoadPayments DtpTransDate(6).Value
    '    Case 7
    '        LoadReturnSalling DtpTransDate(7).Value
    'End Select
    'Exit Sub
    'ErrTrap:

End Sub

Private Sub FgItems_BeforeUserResize(ByVal Row As Long, _
                                     ByVal Col As Long, _
                                     Cancel As Boolean)

    If Col = FgItems.ColIndex("Serial") Then
        Cancel = True
    End If

End Sub

Private Sub FgItems_MouseUp(Button As Integer, _
                            Shift As Integer, _
                            X As Single, _
                            Y As Single)
    Dim LngMouseRow As Long
    Dim LngCurrentItemID As Long

    If Button = vbRightButton Then

        With Me.FgItems
            LngMouseRow = .MouseRow

            If LngMouseRow = -1 Then Exit Sub
            If .Col = -1 Then Exit Sub
            mdifrmmain.MnuItemTools_ItemSerial.Tag = ""
            mdifrmmain.MnuItemTools_ItemCart.Tag = ""
            mdifrmmain.MnuItemTools_ItemData.Tag = ""
            mdifrmmain.MnuItemTools_ItemQty.Tag = ""
        
            If val(.TextMatrix(LngMouseRow, .ColIndex("ItemID"))) <> 0 Then
                LngCurrentItemID = val(.TextMatrix(LngMouseRow, .ColIndex("ItemID")))
                mdifrmmain.MnuItemTools_ItemSerial.Enabled = False
            
                mdifrmmain.MnuItemTools_ItemCart.Tag = LngCurrentItemID & "-" & Me.DcboStores.BoundText
                mdifrmmain.MnuItemTools_ItemQty.Tag = LngCurrentItemID
                mdifrmmain.MnuItemTools_ItemData.Tag = LngCurrentItemID
                Me.PopupMenu mdifrmmain.MnuItemTools
            End If

        End With

    End If

End Sub

Private Sub FGSall_DblClick()

    With Me.FGSall

        If .Row = -1 Then Exit Sub
        If .Col = -1 Then Exit Sub
        If val(.TextMatrix(.Row, .ColIndex("BillNum"))) = 0 Then
            If .Col = .ColIndex("Serial") Or .Col = .ColIndex("Date") Or .Col = .ColIndex("BillVal") Or .Col = .ColIndex("DiscType") Or .Col = .ColIndex("DiscVal") Or .Col = .ColIndex("Tax") Or .Col = .ColIndex("Total") Then
                            
            End If
        End If

    End With

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If Shift = 2 Then
        TabMain.SetFocus

        If KeyCode = vbKeyTab Then
            If TabMain.CurrTab < TabMain.NumTabs - 1 Then
                TabMain.CurrTab = TabMain.CurrTab + 1
            Else
                TabMain.CurrTab = 0
            End If
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Load()
    On Error GoTo ErrTrap

    'StaticStrSql=StaticStrSql & " HAVING      (dbo.Transactions.Transaction_Type = 5)"

    Dim ShowTax As Boolean
    Dim StrSQL As String
    Dim RecordNum As Integer
    Dim BackGround As ClsBackGroundPic
    Dim Dcbombos As ClsDataCombos
    Dim i As Integer
    
    Dim Dcombos As ClsDataCombos
Set Dcombos = New ClsDataCombos

Dcombos.GetStores DcboStoresx(2)
Dcombos.GetBranches Me.DCBranches2


    Resize_Form Me, ReportSize

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    TabMain.CurrTab = 0
    FGSall.ColComboList(FGSall.ColIndex("DiscType")) = "#0;·«ÌÊÃœ Œ’„|#1;Œ’„ »ﬁÌ„…|#2;Œ’„ »‰”»…"
    FgPurchase.ColComboList(FGSall.ColIndex("DiscType")) = "#0;·«ÌÊÃœ Œ’„|#1;Œ’„ »ﬁÌ„…|#2;Œ’„ »‰”»…"

    LoadIcons
    Set BackGround = New ClsBackGroundPic
    Set FGSall.WallPaper = BackGround.Picture
    Set FgPurchase.WallPaper = BackGround.Picture
    Set FGReturn.WallPaper = BackGround.Picture
    Set FGMaintence.WallPaper = BackGround.Picture
    Set FGExpenses.WallPaper = BackGround.Picture
    Set FGCashing.WallPaper = BackGround.Picture
    Set FGSummery.WallPaper = BackGround.Picture
    Set FgPayments.WallPaper = BackGround.Picture
    Set FgReturnSalling.WallPaper = BackGround.Picture
    Set FgItems.WallPaper = BackGround.Picture
    Set FgCustomers.WallPaper = BackGround.Picture

    With Me.FgItems
        .ColComboList(.ColIndex("ItemCase")) = "#1;ÃœÌœ|#2;„” ⁄„·"
        .AutoSize 0, .Cols - 1, False
    End With

    Set Dcbombos = New ClsDataCombos

    Dcbombos.GetStores Me.DcboStores
    Dcbombos.GetBoxes Me.DcbBox
'    LoadSalling Date
'    LoadPurchase Date
'    LoadReturn Date
'    LoadReturnSalling Date
'    LoadMaintenence Date
'    LoadExpenses Date
'    LoadPayments Date
'    LoadCashing Date
    LoadSummary

    For i = DtpTransDate.LBound To DtpTransDate.UBound
        SetDtpickerDate DtpTransDate(i)
    Next i

    SetDtpickerDate Me.DtpCustomers

    '≈Œ›«¡ «·Ã“¡ «·Œ«’ »÷—«∆» «·„»Ì⁄« 
    ShowTax = GetSetting(StrAppRegPath, "SallBill", "HaveTaxOnSalles", False)
    FGSall.ColHidden(FGSall.ColIndex("Tax")) = Not (ShowTax)
    Exit Sub
ErrTrap:
End Sub

Private Sub LblTitle_DblClick(Index As Integer)
    On Error GoTo ErrTrap

    If Me.WindowState = vbNormal Then
        Me.WindowState = vbMaximized
    Else
        Me.WindowState = vbNormal
    End If

    Exit Sub
ErrTrap:
End Sub

Sub Relin()
Dim SmDept As Double
Dim SmCredt As Double
Dim i As Integer
SmDept = 0
SmCredt = 0
With VSFlexGrid1
For i = 2 To .Rows - 1
SmDept = SmDept + val(.TextMatrix(i, .ColIndex("Dept")))
SmCredt = SmCredt + val(.TextMatrix(i, .ColIndex("Credit")))
Next i
End With
lbl(29).Caption = SmDept
lbl(30).Caption = SmCredt
lbl(31).Caption = val(lbl(29).Caption) - val(lbl(30).Caption)
End Sub

Private Sub XPBtn_M_Print_Click(Index As Integer)
    On Error GoTo ErrTrap
    Dim StrSQL As String
    Dim Reports As ClsRepoerts
    StrSQL = "select * From ReportMaintence where DateGoIN=#" & SQLDate(DtpTransDate(3).value) & "#"
    Set Reports = New ClsRepoerts

    Select Case Index

        Case 0
            Reports.MaintenceDailyReport StrSQL, WindowTarget

        Case 1
            Reports.MaintenceDailyReport StrSQL, PrinterTarget
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub LoadIcons()
    On Error GoTo ErrTrap

    '«·„»Ì⁄« 
    With FGSall
        .Cell(flexcpPicture, 0, .ColIndex("BillNum")) = mdifrmmain.ImgLstTree.ListImages("number").Picture
        .Cell(flexcpPicture, 0, .ColIndex("Date")) = mdifrmmain.ImgLstTree.ListImages("Date").Picture
        .Cell(flexcpPicture, 0, .ColIndex("CusName")) = mdifrmmain.ImgLstTree.ListImages("User").Picture
        .Cell(flexcpPicture, 0, .ColIndex("StoreName")) = mdifrmmain.ImgLstTree.ListImages("Open_Node").Picture
        .Cell(flexcpPicture, 0, .ColIndex("BillVal")) = mdifrmmain.ImgLstTree.ListImages("Price").Picture
        .Cell(flexcpPicture, 0, .ColIndex("DiscType")) = mdifrmmain.ImgLstTree.ListImages("DiscountType").Picture
        .Cell(flexcpPicture, 0, .ColIndex("DiscVal")) = mdifrmmain.ImgLstTree.ListImages("Discount").Picture
        .Cell(flexcpPicture, 0, .ColIndex("Tax")) = mdifrmmain.ImgLstTree.ListImages("Tax").Picture
        .Cell(flexcpPicture, 0, .ColIndex("Total")) = mdifrmmain.ImgLstTree.ListImages("Currency").Picture
        .Cell(flexcpPictureAlignment, 0, 0, 0, .Cols - 1) = flexPicAlignRightCenter
    End With

    '«·„‘ —Ì« 
    With FgPurchase
        .Cell(flexcpPicture, 0, .ColIndex("BillNum")) = mdifrmmain.ImgLstTree.ListImages("number").Picture
        .Cell(flexcpPicture, 0, .ColIndex("Date")) = mdifrmmain.ImgLstTree.ListImages("Date").Picture
        .Cell(flexcpPicture, 0, .ColIndex("CusName")) = mdifrmmain.ImgLstTree.ListImages("User").Picture
        .Cell(flexcpPicture, 0, .ColIndex("StoreName")) = mdifrmmain.ImgLstTree.ListImages("Open_Node").Picture
        .Cell(flexcpPicture, 0, .ColIndex("BillVal")) = mdifrmmain.ImgLstTree.ListImages("Price").Picture
        .Cell(flexcpPicture, 0, .ColIndex("DiscType")) = mdifrmmain.ImgLstTree.ListImages("DiscountType").Picture
        .Cell(flexcpPicture, 0, .ColIndex("DiscVal")) = mdifrmmain.ImgLstTree.ListImages("Discount").Picture
        .Cell(flexcpPicture, 0, .ColIndex("Total")) = mdifrmmain.ImgLstTree.ListImages("Currency").Picture
        .Cell(flexcpPictureAlignment, 0, 0, 0, .Cols - 1) = flexPicAlignRightCenter
    End With

    '„— Ã⁄ «·„‘ —Ì« 
    With FGReturn
        .Cell(flexcpPicture, 0, .ColIndex("BillNum")) = mdifrmmain.ImgLstTree.ListImages("number").Picture
        .Cell(flexcpPicture, 0, .ColIndex("Date")) = mdifrmmain.ImgLstTree.ListImages("Date").Picture
        .Cell(flexcpPicture, 0, .ColIndex("CusName")) = mdifrmmain.ImgLstTree.ListImages("User").Picture
        .Cell(flexcpPicture, 0, .ColIndex("StoreName")) = mdifrmmain.ImgLstTree.ListImages("Open_Node").Picture
        .Cell(flexcpPicture, 0, .ColIndex("BillVal")) = mdifrmmain.ImgLstTree.ListImages("Currency").Picture
        .Cell(flexcpPictureAlignment, 0, 0, 0, .Cols - 1) = flexPicAlignRightCenter
    End With

    '„— Ã⁄ «·„»Ì⁄« 
    With FgReturnSalling
        .Cell(flexcpPicture, 0, .ColIndex("BillNum")) = mdifrmmain.ImgLstTree.ListImages("number").Picture
        .Cell(flexcpPicture, 0, .ColIndex("Date")) = mdifrmmain.ImgLstTree.ListImages("Date").Picture
        .Cell(flexcpPicture, 0, .ColIndex("CusName")) = mdifrmmain.ImgLstTree.ListImages("User").Picture
        .Cell(flexcpPicture, 0, .ColIndex("StoreName")) = mdifrmmain.ImgLstTree.ListImages("Open_Node").Picture
        .Cell(flexcpPicture, 0, .ColIndex("BillVal")) = mdifrmmain.ImgLstTree.ListImages("Currency").Picture
        .Cell(flexcpPictureAlignment, 0, 0, 0, .Cols - 1) = flexPicAlignRightCenter
    End With

    '«·’Ì«‰…
    With FGMaintence
        .Cell(flexcpPicture, 0, .ColIndex("BillNum")) = mdifrmmain.ImgLstTree.ListImages("number").Picture
        .Cell(flexcpPicture, 0, .ColIndex("Date")) = mdifrmmain.ImgLstTree.ListImages("Date").Picture
        .Cell(flexcpPicture, 0, .ColIndex("CusName")) = mdifrmmain.ImgLstTree.ListImages("User").Picture
        .Cell(flexcpPicture, 0, .ColIndex("BillVal")) = mdifrmmain.ImgLstTree.ListImages("Currency").Picture
        .Cell(flexcpPictureAlignment, 0, 0, 0, .Cols - 1) = flexPicAlignRightCenter

    End With

    '«·„’—Ê›« 
    With FGExpenses
        .Cell(flexcpPicture, 0, .ColIndex("BillNum")) = mdifrmmain.ImgLstTree.ListImages("number").Picture
        .Cell(flexcpPicture, 0, .ColIndex("Date")) = mdifrmmain.ImgLstTree.ListImages("Date").Picture
        .Cell(flexcpPicture, 0, .ColIndex("CusName")) = mdifrmmain.ImgLstTree.ListImages("code").Picture
        .Cell(flexcpPicture, 0, .ColIndex("BillVal")) = mdifrmmain.ImgLstTree.ListImages("Currency").Picture
        .Cell(flexcpPictureAlignment, 0, 0, 0, .Cols - 1) = flexPicAlignRightCenter
    End With

    '«·„ﬁ»Ê÷« 
    With FGCashing
        .Cell(flexcpPicture, 0, .ColIndex("BillNum")) = mdifrmmain.ImgLstTree.ListImages("number").Picture
        .Cell(flexcpPicture, 0, .ColIndex("Date")) = mdifrmmain.ImgLstTree.ListImages("Date").Picture
        .Cell(flexcpPicture, 0, .ColIndex("CashingType")) = mdifrmmain.ImgLstTree.ListImages("code").Picture
        .Cell(flexcpPicture, 0, .ColIndex("CusName")) = mdifrmmain.ImgLstTree.ListImages("User").Picture
        .Cell(flexcpPicture, 0, .ColIndex("BillVal")) = mdifrmmain.ImgLstTree.ListImages("Currency").Picture
        .Cell(flexcpPictureAlignment, 0, 0, 0, .Cols - 1) = flexPicAlignRightCenter
    End With

    '«·„œ›Ê⁄« 
    With FgPayments
        .Cell(flexcpPicture, 0, .ColIndex("BillNum")) = mdifrmmain.ImgLstTree.ListImages("number").Picture
        .Cell(flexcpPicture, 0, .ColIndex("Date")) = mdifrmmain.ImgLstTree.ListImages("Date").Picture
        .Cell(flexcpPicture, 0, .ColIndex("CashingType")) = mdifrmmain.ImgLstTree.ListImages("code").Picture
        .Cell(flexcpPicture, 0, .ColIndex("CusName")) = mdifrmmain.ImgLstTree.ListImages("User").Picture
        .Cell(flexcpPicture, 0, .ColIndex("BillVal")) = mdifrmmain.ImgLstTree.ListImages("Currency").Picture
        .Cell(flexcpPictureAlignment, 0, 0, 0, .Cols - 1) = flexPicAlignRightCenter
    End With

    With FGSummery
        .Cell(flexcpPicture, 0, .ColIndex("Type")) = mdifrmmain.ImgLstTree.ListImages("code").Picture
        .Cell(flexcpPicture, 0, .ColIndex("Value")) = mdifrmmain.ImgLstTree.ListImages("Currency").Picture
        .Cell(flexcpPictureAlignment, 0, 0, 0, .Cols - 1) = flexPicAlignRightCenter
    End With

    Exit Sub
ErrTrap:
End Sub

Private Sub LoadSalling(StrDate As String)
    On Error GoTo ErrTrap

    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    Dim RecordNum As Integer
    '«·„»Ì⁄« 
    LblVal(0).Caption = "0"
    LblDiscount(0).Caption = "0"
    LblTotal(0).Caption = "0"
    LblBillCount(0).Caption = "0"

    With Me.FGSall
        .Clear flexClearScrollable, flexClearEverything
        .Rows = .FixedRows + 1
        Set rs = New ADODB.Recordset
    
        '    StrSQL = "Select * From ReportSallingTime where Transaction_Date=" & SQLDate(CDate(StrDate), True) & ""
    
        StaticStrSql = "  SELECT     TOP 100 PERCENT QryTransactionsTotal.Transaction_ID, QryTransactionsTotal.Transaction_Date, dbo.TblCustemers.CusName, dbo.TblStore.StoreName,"
        StaticStrSql = StaticStrSql + " QryTransactionsTotal.Trans_DiscountType, QryTransactionsTotal.Trans_Discount, QryTransactionsTotal.Transaction_Serial, QryTransactionsTotal.PaymentType,"
        StaticStrSql = StaticStrSql + " QryTransactionsTotal.CusID, QryTransactionsTotal.TaxFound, QryTransactionsTotal.TaxValue, QryTransactionsTotal.TransSum, QryTransactionsTotal.TotalAfterTax,"
        StaticStrSql = StaticStrSql + " QryTransactionsTotal.TransNet, QryTransactionsTotal.Transaction_Type, dbo.TblBoxesData.BoxID, dbo.TblBoxesData.BoxName, dbo.Transactions.SaleType,"
        StaticStrSql = StaticStrSql + "  QryTransactionsTotal.Storeid , QryTransactionsTotal.Emp_id, dbo.TblEmployee.emp_name, dbo.Transactions.NoteSerial, dbo.Transactions.NoteSerial1"
        StaticStrSql = StaticStrSql + " FROM         dbo.Transactions INNER JOIN"
        StaticStrSql = StaticStrSql + " dbo.TblStore INNER JOIN"
        StaticStrSql = StaticStrSql + " dbo.QryTransactionsTotal() QryTransactionsTotal ON dbo.TblStore.StoreID = QryTransactionsTotal.StoreID ON"
        StaticStrSql = StaticStrSql + " dbo.Transactions.Transaction_ID = QryTransactionsTotal.Transaction_ID INNER JOIN"
        StaticStrSql = StaticStrSql + " dbo.TblEmployee ON QryTransactionsTotal.Emp_ID = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
        StaticStrSql = StaticStrSql + " dbo.QryNotesBoxes ON QryTransactionsTotal.Transaction_ID = dbo.QryNotesBoxes.Transaction_ID LEFT OUTER JOIN"
        StaticStrSql = StaticStrSql + " dbo.TblCustemers ON QryTransactionsTotal.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
        StaticStrSql = StaticStrSql + " dbo.TblBoxesData ON dbo.QryNotesBoxes.BoxID = dbo.TblBoxesData.BoxID"
        StaticStrSql = StaticStrSql + " where Transactions.Transaction_Date=" & SQLDate(CDate(StrDate), True) & ""


  If DCBranches2.BoundText <> "" Then
  StaticStrSql = StaticStrSql & "   and dbo.Transactions.BranchId= " & val(DCBranches2.BoundText)
       
  End If
    If val(DcboStoresx(2).BoundText) <> 0 And (DcboStoresx(2).Text) <> "" Then
                  StaticStrSql = StaticStrSql + " and  dbo.Transactions.StoreId=" & DcboStoresx(2).BoundText
     End If
        


  
        StaticStrSql = StaticStrSql + " and (     (QryTransactionsTotal.Transaction_Type = 2) OR"
        StaticStrSql = StaticStrSql + " (QryTransactionsTotal.Transaction_Type = 21))"
        StaticStrSql = StaticStrSql + " ORDER BY QryTransactionsTotal.Transaction_Date"

        rs.Open StaticStrSql, Cn, adOpenStatic, adLockOptimistic, adCmdText
salesSql = StaticStrSql

        If Not (rs.BOF Or rs.EOF) Then
            .Rows = rs.RecordCount + 1

            For RecordNum = 1 To rs.RecordCount
                .TextMatrix(RecordNum, .ColIndex("Index")) = RecordNum
                .TextMatrix(RecordNum, .ColIndex("BillNum")) = IIf(IsNull(rs("Transaction_ID").value), "", rs("Transaction_ID").value)
                .TextMatrix(RecordNum, .ColIndex("Serial")) = IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value)
                .TextMatrix(RecordNum, .ColIndex("Date")) = IIf(IsNull(rs("Transaction_Date").value), "", DisplayDate(rs("Transaction_Date").value))
                .TextMatrix(RecordNum, .ColIndex("CusName")) = IIf(IsNull(rs("CusName").value), "", rs("CusName").value)
                .TextMatrix(RecordNum, .ColIndex("StoreName")) = IIf(IsNull(rs("StoreName").value), "", rs("StoreName").value)
                .TextMatrix(RecordNum, .ColIndex("BillVal")) = IIf(IsNull(rs("TransSum").value), 0, rs("TransSum").value)
                .TextMatrix(RecordNum, .ColIndex("DiscType")) = IIf(IsNull(rs("Trans_DiscountType").value), 0, rs("Trans_DiscountType").value)
                .TextMatrix(RecordNum, .ColIndex("DiscVal")) = IIf(IsNull(rs("Trans_Discount").value), 0, rs("Trans_Discount").value)

                If rs("TaxFound").value = True Then
                    .TextMatrix(RecordNum, .ColIndex("Tax")) = IIf(IsNull(rs("TaxValue").value), "", rs("TaxValue").value)
                    .TextMatrix(RecordNum, .ColIndex("Total")) = IIf(IsNull(rs("TotalAfterTax").value), "", rs("TotalAfterTax").value)
                Else
                    .TextMatrix(RecordNum, .ColIndex("Tax")) = "·«ÌÊÃœ"
                    .TextMatrix(RecordNum, .ColIndex("Total")) = IIf(IsNull(rs("TransNet").value), "", rs("TransNet").value)
                End If

                rs.MoveNext
            Next RecordNum

            .AutoSize 0, .Cols - 1, False
            LblVal(0).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("BillVal"), .Rows - 1, .ColIndex("BillVal"))
            LblTotal(0).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Total"), .Rows - 1, .ColIndex("Total"))
            LblDiscount(0).Caption = val(LblVal(0).Caption) - val(LblTotal(0).Caption)
            LblBillCount(0).Caption = rs.RecordCount
        End If

        .AutoSize 0, .Cols - 1, False
    End With

    rs.Close
    Exit Sub
ErrTrap:
End Sub
Function print_reportPurchases(Optional NoteSerial As String, Optional StrDate As String)
    
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
      MySQL = " SELECT     TOP 100 PERCENT QryTransactionsTotal.Transaction_ID, QryTransactionsTotal.Transaction_Date, dbo.TblCustemers.CusName, dbo.TblStore.StoreName, "
      MySQL = MySQL & "                QryTransactionsTotal.Trans_DiscountType, QryTransactionsTotal.Trans_Discount, QryTransactionsTotal.Transaction_Serial, QryTransactionsTotal.PaymentType,"
      MySQL = MySQL & "                QryTransactionsTotal.CusID, QryTransactionsTotal.TaxFound, QryTransactionsTotal.TaxValue, QryTransactionsTotal.TransSum, QryTransactionsTotal.TotalAfterTax,"
      MySQL = MySQL & "                 QryTransactionsTotal.TransNet, QryTransactionsTotal.Transaction_Type, dbo.TblBoxesData.BoxID, dbo.TblBoxesData.BoxName, QryTransactionsTotal.StoreID,"
      MySQL = MySQL & "                QryTransactionsTotal.Emp_id , dbo.TblEmployee.emp_name, dbo.Transactions.NoteSerial, dbo.Transactions.NoteSerial1"
      MySQL = MySQL & "   FROM         dbo.TblEmployee RIGHT OUTER JOIN"
      MySQL = MySQL & "                dbo.Transactions INNER JOIN"
      MySQL = MySQL & "                dbo.QryTransactionsTotal() QryTransactionsTotal ON dbo.Transactions.Transaction_ID = QryTransactionsTotal.Transaction_ID ON"
      MySQL = MySQL & "                dbo.TblEmployee.Emp_ID = QryTransactionsTotal.Emp_ID LEFT OUTER JOIN"
      MySQL = MySQL & "                dbo.TblStore ON QryTransactionsTotal.StoreID = dbo.TblStore.StoreID LEFT OUTER JOIN"
      MySQL = MySQL & "                dbo.TblCustemers ON QryTransactionsTotal.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
      MySQL = MySQL & "                dbo.QryNotesBoxes ON QryTransactionsTotal.Transaction_ID = dbo.QryNotesBoxes.Transaction_ID LEFT OUTER JOIN"
      MySQL = MySQL & "                dbo.TblBoxesData ON dbo.QryNotesBoxes.BoxID = dbo.TblBoxesData.BoxID"
      MySQL = MySQL & "    WHERE     ((QryTransactionsTotal.Transaction_Type = 1) OR"
      MySQL = MySQL & "                   (QryTransactionsTotal.Transaction_Type = 22))"
      MySQL = MySQL & "  and QryTransactionsTotal.Transaction_Date = " & SQLDate(CDate(StrDate), True) & ""
      MySQL = MySQL & "   ORDER BY QryTransactionsTotal.Transaction_Date"

      '  MySQL = "Select * From ReportBuyTime_Client where "
        If SystemOptions.UserInterface = ArabicInterface Then
          StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\ReportDayPrushes.rpt"
        Else
        StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\ReportDayPrushesE.rpt"
        End If
    If Dir(StrFileName) = "" Then
        Screen.MousePointer = vbDefault
        Exit Function
    End If
 MySQL = PurcahseSQL
 
    Set RsData = New ADODB.Recordset
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        'GetMsgs 138, vbExclamation
        If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        Else
        MsgBox "No Data"
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
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
     

        StrReportTitle = ""
 
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
 
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
Function print_reportBoxws(Optional NoteSerial As String, Optional StrDate As String)
    
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
 Dim StrAccountCode As String
       StrAccountCode = GetMyAccountCode("TblBoxesData", "BoxID", val(Me.DcbBox.BoundText), "Account_Code")
MySQL = " SELECT     dbo.Notes.ChqueNum, dbo.Notes.ManualNo, dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_ID, dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit,"
MySQL = MySQL & "                       dbo.DOUBLE_ENTREY_VOUCHERS.[Value] AS DEV_Value, dbo.DOUBLE_ENTREY_VOUCHERS.RecordDateH, dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code,"
MySQL = MySQL & "                      dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_Description AS DEV_DES,"
MySQL = MySQL & "                      dbo.DOUBLE_ENTREY_VOUCHERS.Double_Entry_Vouchers_Descriptione AS DevDESE, dbo.ACCOUNTS.Account_Name,"
MySQL = MySQL & "                      dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No, dbo.TblNotesTypes.NotesTypeName, dbo.DOUBLE_ENTREY_VOUCHERS.UserID, dbo.TblUsers.UserName,"
MySQL = MySQL & "                      dbo.DOUBLE_ENTREY_VOUCHERS.RecordDate, dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID, dbo.DOUBLE_ENTREY_VOUCHERS.ReceiptID,"
MySQL = MySQL & "                      dbo.DOUBLE_ENTREY_VOUCHERS.OperaID, dbo.DOUBLE_ENTREY_VOUCHERS.Transaction_ID, dbo.Transactions.Transaction_Serial,"
MySQL = MySQL & "                      dbo.Transactions.Transaction_Date, dbo.TransactionTypes.TransactionTypeName, dbo.DOUBLE_ENTREY_VOUCHERS.Posted,"
MySQL = MySQL & "                      dbo.DOUBLE_ENTREY_VOUCHERS.PostedDate, dbo.DOUBLE_ENTREY_VOUCHERS.PostedUserID, dbo.DOUBLE_ENTREY_VOUCHERS.Account_Interval_ID,"
MySQL = MySQL & "                      dbo.Notes.NoteDate, dbo.Notes.NoteType, dbo.Notes.NoteSerial, dbo.Notes.Note_Value, dbo.ACCOUNTS.Account_Serial, dbo.ACCOUNTS.Account_NameEng,"
MySQL = MySQL & "                      dbo.ACCOUNTS.Parent_Account_Code, dbo.ACCOUNTS.opening_balance, dbo.ACCOUNTS.opening_balance_type, dbo.ACCOUNTS.Branch,"
MySQL = MySQL & "                      dbo.ACCOUNTS.Sum_account, dbo.ACCOUNTS.cost_center, dbo.ACCOUNTS.currenct_code, dbo.Notes.Remark, dbo.Notes.note_value_by_characters,"
MySQL = MySQL & "                      dbo.Notes.foxy_no, dbo.DOUBLE_ENTREY_VOUCHERS.DEV_ID_Line_No1, dbo.DOUBLE_ENTREY_VOUCHERS.project_id, dbo.TblNotesTypes.NotesTypeNamee,"
MySQL = MySQL & "                      dbo.TransactionTypes.TransactionEnglishName, dbo.DOUBLE_ENTREY_VOUCHERS.opr_fullcode, dbo.Notes.NoteSerial1,"
MySQL = MySQL & "                      dbo.DOUBLE_ENTREY_VOUCHERS.branch_id, dbo.TblBranchesData.ActivityTypeId, dbo.DOUBLE_ENTREY_VOUCHERS.notes_all,"
MySQL = MySQL & "                      dbo.TblBranchesData.branch_name , dbo.TblBranchesData.branch_namee"
MySQL = MySQL & " FROM         dbo.TblBranchesData INNER JOIN"
MySQL = MySQL & "                      dbo.TblUsers INNER JOIN"
MySQL = MySQL & "                      dbo.ACCOUNTS INNER JOIN"
MySQL = MySQL & "                      dbo.DOUBLE_ENTREY_VOUCHERS ON dbo.ACCOUNTS.Account_Code = dbo.DOUBLE_ENTREY_VOUCHERS.Account_Code ON"
MySQL = MySQL & "                      dbo.TblUsers.UserID = dbo.DOUBLE_ENTREY_VOUCHERS.UserID ON"
MySQL = MySQL & "                      dbo.TblBranchesData.branch_id = dbo.DOUBLE_ENTREY_VOUCHERS.branch_id LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.Notes LEFT OUTER JOIN"
MySQL = MySQL & "                     dbo.TblNotesTypes ON dbo.Notes.NoteType = dbo.TblNotesTypes.NotesType ON dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID = dbo.Notes.NoteID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.Transactions ON dbo.DOUBLE_ENTREY_VOUCHERS.Transaction_ID = dbo.Transactions.Transaction_ID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type"
MySQL = MySQL + " WHERE     (dbo.ACCOUNTS.Account_Code = '" & StrAccountCode & "') "
MySQL = MySQL + " and RecordDate =" & SQLDate(DtpTransDate(9).value, True) & ""
 
        If SystemOptions.UserInterface = ArabicInterface Then
          StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\RptDayBoxes.rpt"
        Else
        StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\RptDayBoxes.rpt"
        End If
    If Dir(StrFileName) = "" Then
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        'GetMsgs 138, vbExclamation
        If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        Else
        MsgBox "No Data"
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
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
     

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
Function print_reportExpenss(Optional NoteSerial As String, Optional StrDate As String)
    
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
 
        MySQL = " SELECT     TOP 100 PERCENT dbo.Notes.NoteID, dbo.Notes.NoteDate, dbo.Notes.NoteType, dbo.Notes.Note_Value, dbo.Notes.ExpensesID, dbo.ExpensesType.Name,"
        MySQL = MySQL + "                      dbo.Notes.NoteSerial, dbo.Notes.Remark, dbo.TblUsers.UserName, dbo.Notes.BoxID, dbo.Notes.UserID, dbo.TblBoxesData.BoxName, dbo.BanksData.BankName,"
        MySQL = MySQL + "                      dbo.Notes.BankID , dbo.Notes.ChqueNum, dbo.Notes.DueDate, dbo.Notes.NoteSerial1"
        MySQL = MySQL + " FROM         dbo.TblUsers RIGHT OUTER JOIN"
        MySQL = MySQL + "                      dbo.ExpensesType RIGHT OUTER JOIN"
        MySQL = MySQL + "                      dbo.Notes ON dbo.ExpensesType.ID = dbo.Notes.ExpensesID ON dbo.TblUsers.UserID = dbo.Notes.UserID LEFT OUTER JOIN"
        MySQL = MySQL + "                      dbo.TblBoxesData ON dbo.Notes.BoxID = dbo.TblBoxesData.BoxID LEFT OUTER JOIN"
        MySQL = MySQL + "                      dbo.BanksData ON dbo.Notes.BankID = dbo.BanksData.BankID"
        MySQL = MySQL + " WHERE    (dbo.Notes.NoteType = 53 or dbo.Notes.NoteType = 3)"

        MySQL = MySQL + "  and dbo.Notes.NoteDate=" & SQLDate(CDate(StrDate), True) & ""
        MySQL = MySQL + "  ORDER BY dbo.Notes.NoteID"
 
        If SystemOptions.UserInterface = ArabicInterface Then
          StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\ReportDayExpensses.rpt"
        Else
        StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\ReportDayExpenssesE.rpt"
        End If
    If Dir(StrFileName) = "" Then
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        'GetMsgs 138, vbExclamation
        If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        Else
        MsgBox "No Data"
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
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
     

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
Function print_reportReturnPurches(Optional NoteSerial As String, Optional StrDate As String)
    
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
 
        MySQL = " SELECT     dbo.Transactions.Transaction_ID, dbo.Transactions.Transaction_Date, dbo.Transactions.Transaction_Type, "
        MySQL = MySQL & "  SUM(dbo.Transaction_Details.showPrice * dbo.Transaction_Details.ShowQty) AS Total, dbo.TblStore.StoreName, dbo.TblCustemers.CusName,"
        MySQL = MySQL & "   dbo.Transactions.Transaction_serial, dbo.Transactions.NoteSerial, dbo.Transactions.NoteSerial1"
        MySQL = MySQL & " FROM         dbo.TblStore INNER JOIN"
        MySQL = MySQL & " dbo.TblCustemers INNER JOIN"
        MySQL = MySQL & "  dbo.Transactions ON dbo.TblCustemers.CusID = dbo.Transactions.CusID ON dbo.TblStore.StoreID = dbo.Transactions.StoreID INNER JOIN"
        MySQL = MySQL & "  dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID"
        MySQL = MySQL + " where Transactions.Transaction_Date=" & SQLDate(CDate(StrDate), True) & ""
        MySQL = MySQL & " GROUP BY dbo.Transactions.Transaction_ID, dbo.Transactions.Transaction_Date, dbo.Transactions.Transaction_Type, dbo.TblStore.StoreName, dbo.TblCustemers.CusName,"
        MySQL = MySQL & "   dbo.Transactions.Transaction_serial, dbo.Transactions.NoteSerial, dbo.Transactions.NoteSerial1"
        MySQL = MySQL & " HAVING      (dbo.Transactions.Transaction_Type = 5)"
 
        If SystemOptions.UserInterface = ArabicInterface Then
          StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\ReportDayReturnSailling.rpt"
        Else
        StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\ReportDayReturnSaillingE.rpt"
        End If
    If Dir(StrFileName) = "" Then
        Screen.MousePointer = vbDefault
        Exit Function
    End If
   
MySQL = REPurcahseSQL

    Set RsData = New ADODB.Recordset
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        'GetMsgs 138, vbExclamation
        If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        Else
        MsgBox "No Data"
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
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
     

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
Function print_reportReturnSaling(Optional NoteSerial As String, Optional StrDate As String)
    
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
 
        MySQL = " SELECT     dbo.Transactions.Transaction_ID, dbo.Transactions.Transaction_Date, dbo.Transactions.Transaction_Type, "
        MySQL = MySQL + "  SUM(dbo.Transaction_Details.showPrice * dbo.Transaction_Details.ShowQty) AS Total, dbo.TblStore.StoreName, dbo.TblCustemers.CusName,"
        MySQL = MySQL + "  dbo.Transactions.Transaction_serial l, dbo.Transactions.NoteSerial, dbo.Transactions.NoteSerial1 "
        MySQL = MySQL + " FROM         dbo.TblStore INNER JOIN"
        MySQL = MySQL + " dbo.TblCustemers INNER JOIN"
        MySQL = MySQL + " dbo.Transactions ON dbo.TblCustemers.CusID = dbo.Transactions.CusID ON dbo.TblStore.StoreID = dbo.Transactions.StoreID INNER JOIN"
        MySQL = MySQL + "  dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID"
        MySQL = MySQL + " where Transactions.Transaction_Date=" & SQLDate(CDate(StrDate), True) & ""
        MySQL = MySQL + " GROUP BY dbo.Transactions.Transaction_ID, dbo.Transactions.Transaction_Date, dbo.Transactions.Transaction_Type, dbo.TblStore.StoreName, dbo.TblCustemers.CusName,"
        MySQL = MySQL + "  dbo.Transactions.Transaction_serial , dbo.Transactions.NoteSerial, dbo.Transactions.NoteSerial1"
        MySQL = MySQL + " HAVING      (dbo.Transactions.Transaction_Type = 9)"
  
        If SystemOptions.UserInterface = ArabicInterface Then
          StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\ReportDayReturnPurches.rpt"
        Else
        StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\ReportDayReturnPurchesE.rpt"
        End If
    If Dir(StrFileName) = "" Then
        Screen.MousePointer = vbDefault
        Exit Function
    End If
MySQL = REsalesSql
    Set RsData = New ADODB.Recordset
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        'GetMsgs 138, vbExclamation
        If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        Else
        MsgBox "No Data"
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
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
     

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
Function print_reportSaling(Optional NoteSerial As String, Optional StrDate As String)
    
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
 
        MySQL = "  SELECT     TOP 100 PERCENT QryTransactionsTotal.Transaction_ID, QryTransactionsTotal.Transaction_Date, dbo.TblCustemers.CusName, dbo.TblStore.StoreName,"
        MySQL = MySQL + " QryTransactionsTotal.Trans_DiscountType, QryTransactionsTotal.Trans_Discount, QryTransactionsTotal.Transaction_Serial, QryTransactionsTotal.PaymentType,"
        MySQL = MySQL + " QryTransactionsTotal.CusID, QryTransactionsTotal.TaxFound, QryTransactionsTotal.TaxValue, QryTransactionsTotal.TransSum, QryTransactionsTotal.TotalAfterTax,"
        MySQL = MySQL + " QryTransactionsTotal.TransNet, QryTransactionsTotal.Transaction_Type, dbo.TblBoxesData.BoxID, dbo.TblBoxesData.BoxName, dbo.Transactions.SaleType,"
        MySQL = MySQL + "  QryTransactionsTotal.Storeid , QryTransactionsTotal.Emp_id, dbo.TblEmployee.emp_name, dbo.Transactions.NoteSerial, dbo.Transactions.NoteSerial1"
        MySQL = MySQL + " FROM         dbo.Transactions INNER JOIN"
        MySQL = MySQL + " dbo.TblStore INNER JOIN"
        MySQL = MySQL + " dbo.QryTransactionsTotal() QryTransactionsTotal ON dbo.TblStore.StoreID = QryTransactionsTotal.StoreID ON"
        MySQL = MySQL + " dbo.Transactions.Transaction_ID = QryTransactionsTotal.Transaction_ID INNER JOIN"
        MySQL = MySQL + " dbo.TblEmployee ON QryTransactionsTotal.Emp_ID = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
        MySQL = MySQL + " dbo.QryNotesBoxes ON QryTransactionsTotal.Transaction_ID = dbo.QryNotesBoxes.Transaction_ID LEFT OUTER JOIN"
        MySQL = MySQL + " dbo.TblCustemers ON QryTransactionsTotal.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
        MySQL = MySQL + " dbo.TblBoxesData ON dbo.QryNotesBoxes.BoxID = dbo.TblBoxesData.BoxID"
        MySQL = MySQL + " where Transactions.Transaction_Date=" & SQLDate(CDate(StrDate), True) & ""

        MySQL = MySQL + " and (     (QryTransactionsTotal.Transaction_Type = 2) OR"
        MySQL = MySQL + " (QryTransactionsTotal.Transaction_Type = 21))"
        MySQL = MySQL + " ORDER BY QryTransactionsTotal.Transaction_Date"
 
MySQL = salesSql
        If SystemOptions.UserInterface = ArabicInterface Then
          StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\ReportDaySailling.rpt"
        Else
        StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\ReportDaySaillingE.rpt"
        End If
    If Dir(StrFileName) = "" Then
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        'GetMsgs 138, vbExclamation
        If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        Else
        MsgBox "No Data"
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
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
     

        StrReportTitle = ""
 
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
 
'    xReport.ParameterFields(5).AddCurrentValue ToHijriDate(RsData("notedate").value)
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.title
    xReport.ReportAuthor = App.title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName, , StaticStrSql

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
End Function
Private Sub LoadPurchase(StrDate As String)
   ' On Error GoTo ErrTrap

    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    Dim RecordNum As Integer
    LblVal(1).Caption = "0"
    LblDiscount(1).Caption = "0"
    LblTotal(1).Caption = "0"
    LblBillCount(1).Caption = "0"

    With Me.FgPurchase
        .Clear flexClearScrollable, flexClearEverything
        .Rows = .FixedRows + 1
      StrSQL = " SELECT     TOP 100 PERCENT QryTransactionsTotal.Transaction_ID, QryTransactionsTotal.Transaction_Date, dbo.TblCustemers.CusName, dbo.TblStore.StoreName, "
      StrSQL = StrSQL & "                QryTransactionsTotal.Trans_DiscountType, QryTransactionsTotal.Trans_Discount, QryTransactionsTotal.Transaction_Serial, QryTransactionsTotal.PaymentType,"
      StrSQL = StrSQL & "                QryTransactionsTotal.CusID, QryTransactionsTotal.TaxFound, QryTransactionsTotal.TaxValue, QryTransactionsTotal.TransSum, QryTransactionsTotal.TotalAfterTax,"
      StrSQL = StrSQL & "                 QryTransactionsTotal.TransNet, QryTransactionsTotal.Transaction_Type, dbo.TblBoxesData.BoxID, dbo.TblBoxesData.BoxName, QryTransactionsTotal.StoreID,"
      StrSQL = StrSQL & "                QryTransactionsTotal.Emp_id , dbo.TblEmployee.emp_name, dbo.Transactions.NoteSerial, dbo.Transactions.NoteSerial1"
      StrSQL = StrSQL & "   FROM         dbo.TblEmployee RIGHT OUTER JOIN"
      StrSQL = StrSQL & "                dbo.Transactions INNER JOIN"
      StrSQL = StrSQL & "                dbo.QryTransactionsTotal() QryTransactionsTotal ON dbo.Transactions.Transaction_ID = QryTransactionsTotal.Transaction_ID ON"
      StrSQL = StrSQL & "                dbo.TblEmployee.Emp_ID = QryTransactionsTotal.Emp_ID LEFT OUTER JOIN"
      StrSQL = StrSQL & "                dbo.TblStore ON QryTransactionsTotal.StoreID = dbo.TblStore.StoreID LEFT OUTER JOIN"
      StrSQL = StrSQL & "                dbo.TblCustemers ON QryTransactionsTotal.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
      StrSQL = StrSQL & "                dbo.QryNotesBoxes ON QryTransactionsTotal.Transaction_ID = dbo.QryNotesBoxes.Transaction_ID LEFT OUTER JOIN"
      StrSQL = StrSQL & "                dbo.TblBoxesData ON dbo.QryNotesBoxes.BoxID = dbo.TblBoxesData.BoxID"
      StrSQL = StrSQL & "    WHERE     ((QryTransactionsTotal.Transaction_Type = 1) OR"
      StrSQL = StrSQL & "                   (QryTransactionsTotal.Transaction_Type = 22))"
      StrSQL = StrSQL & "  and QryTransactionsTotal.Transaction_Date = " & SQLDate(CDate(StrDate), True) & ""
      
      
      
  If DCBranches2.BoundText <> "" Then
'  StrSQL = StrSQL & "   and dbo.Transactions.BranchId= " & val(DCBranches2.BoundText)
       
  End If
    If val(DcboStoresx(2).BoundText) <> 0 And (DcboStoresx(2).Text) <> "" Then
                  StrSQL = StrSQL + " and  dbo.TblStore.StoreId=" & DcboStoresx(2).BoundText
     End If
        
        
      StrSQL = StrSQL & "   ORDER BY QryTransactionsTotal.Transaction_Date"
       ' StrSQL = "Select * From ReportBuyTime_Client where Transaction_Date=" & SQLDate(CDate(StrDate), True) & ""
    
        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
PurcahseSQL = StrSQL
        If Not (rs.BOF Or rs.EOF) Then
            .Rows = rs.RecordCount + 1

            For RecordNum = 1 To rs.RecordCount
                .TextMatrix(RecordNum, .ColIndex("Index")) = RecordNum
                .TextMatrix(RecordNum, .ColIndex("BillNum")) = IIf(IsNull(rs("Transaction_ID").value), "", rs("Transaction_ID").value)
               '.TextMatrix(RecordNum, .ColIndex("Serial")) = get_transaction_NoteSerial1ByiD(val(.TextMatrix(RecordNum, .ColIndex("BillNum"))), "")
                  .TextMatrix(RecordNum, .ColIndex("Serial")) = get_transaction_NoteSerial1ByiDTemp(val(.TextMatrix(RecordNum, .ColIndex("BillNum"))))
                .TextMatrix(RecordNum, .ColIndex("Date")) = IIf(IsNull(rs("Transaction_Date").value), "", Format(rs("Transaction_Date").value, "YYYY/MM/DD"))
                .TextMatrix(RecordNum, .ColIndex("CusName")) = IIf(IsNull(rs("CusName").value), "", rs("CusName").value)
                .TextMatrix(RecordNum, .ColIndex("StoreName")) = IIf(IsNull(rs("StoreName").value), "", rs("StoreName").value)
                .TextMatrix(RecordNum, .ColIndex("BillVal")) = IIf(IsNull(rs("TransSum").value), 0, rs("TransSum").value)
                .TextMatrix(RecordNum, .ColIndex("DiscType")) = IIf(IsNull(rs("Trans_DiscountType").value), 0, rs("Trans_DiscountType").value)
                .TextMatrix(RecordNum, .ColIndex("DiscVal")) = IIf(IsNull(rs("Trans_Discount").value), 0, rs("Trans_Discount").value)

                If rs("TaxFound").value = True Then
                
                    .TextMatrix(RecordNum, .ColIndex("Total")) = IIf(IsNull(rs("TotalAfterTax").value), "", rs("TotalAfterTax").value)
                Else

                    .TextMatrix(RecordNum, .ColIndex("Total")) = IIf(IsNull(rs("TransNet").value), "", rs("TransNet").value)
                End If

                rs.MoveNext
           
            Next RecordNum

            .AutoSize 0, .Cols - 1, False
            LblVal(1).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("BillVal"), .Rows - 1, .ColIndex("BillVal"))
            LblTotal(1).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Total"), .Rows - 1, .ColIndex("Total"))
            LblDiscount(1).Caption = val(LblVal(1).Caption) - val(LblTotal(1).Caption)
            LblBillCount(1).Caption = rs.RecordCount
        End If

        .AutoSize 0, .Cols - 1, False
    End With

    rs.Close
    Exit Sub
ErrTrap:
End Sub

Private Sub LoadReturn(StrDate As String)
    On Error GoTo ErrTrap

    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    Dim RecordNum As Integer
    LblVal(2).Caption = "0"
    LblBillCount(2).Caption = "0"

    '„— Ã⁄ «·„‘ —Ì« 
    With Me.FGReturn
        .Clear flexClearScrollable, flexClearEverything
        .Rows = .FixedRows + 1
        '    StrSQL = "SELECT * From QryReturn"
        '    StrSQL = StrSQL + " where Transaction_Date=" & SQLDate(CDate(StrDate), True) & ""
        StaticStrSql = "SELECT     dbo.Transactions.Transaction_ID, dbo.Transactions.Transaction_Date, dbo.Transactions.Transaction_Type, "
        StaticStrSql = StaticStrSql & "  SUM(dbo.Transaction_Details.showPrice * dbo.Transaction_Details.ShowQty) AS Total, dbo.TblStore.StoreName, dbo.TblCustemers.CusName,"
        StaticStrSql = StaticStrSql & "   dbo.Transactions.Transaction_serial , dbo.Transactions.NoteSerial, dbo.Transactions.NoteSerial1"
        StaticStrSql = StaticStrSql & " FROM         dbo.TblStore INNER JOIN"
        StaticStrSql = StaticStrSql & " dbo.TblCustemers INNER JOIN"
        StaticStrSql = StaticStrSql & "  dbo.Transactions ON dbo.TblCustemers.CusID = dbo.Transactions.CusID ON dbo.TblStore.StoreID = dbo.Transactions.StoreID INNER JOIN"
        StaticStrSql = StaticStrSql & "  dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID"
        '
    
        StaticStrSql = StaticStrSql + " where Transactions.Transaction_Date=" & SQLDate(CDate(StrDate), True) & ""
  
    If DCBranches2.BoundText <> "" Then
  StaticStrSql = StaticStrSql & "   and dbo.Transactions.BranchId= " & val(DCBranches2.BoundText)
       
  End If
    If val(DcboStoresx(2).BoundText) <> 0 And (DcboStoresx(2).Text) <> "" Then
                  StaticStrSql = StaticStrSql + " and  dbo.Transactions.StoreId=" & DcboStoresx(2).BoundText
     End If


        StaticStrSql = StaticStrSql & " GROUP BY dbo.Transactions.Transaction_ID, dbo.Transactions.Transaction_Date, dbo.Transactions.Transaction_Type, dbo.TblStore.StoreName, dbo.TblCustemers.CusName,"
        StaticStrSql = StaticStrSql & "   dbo.Transactions.Transaction_serial , dbo.Transactions.NoteSerial, dbo.Transactions.NoteSerial1"
        StaticStrSql = StaticStrSql & " HAVING      (dbo.Transactions.Transaction_Type = 5)"
  REPurcahseSQL = StaticStrSql
        Set rs = New ADODB.Recordset
        rs.Open StaticStrSql, Cn, adOpenStatic, adLockOptimistic, adCmdText

        If Not (rs.BOF Or rs.EOF) Then
            FGReturn.Rows = rs.RecordCount + 1

            For RecordNum = 1 To rs.RecordCount
                .TextMatrix(RecordNum, .ColIndex("Index")) = RecordNum
                .TextMatrix(RecordNum, .ColIndex("BillNum")) = IIf(IsNull(rs("Transaction_ID").value), "", rs("Transaction_ID").value)
                '     .TextMatrix(RecordNum, .ColIndex("Serial")) = IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value)
                .TextMatrix(RecordNum, .ColIndex("Serial")) = IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value)

                .TextMatrix(RecordNum, .ColIndex("Date")) = IIf(IsNull(rs("Transaction_Date").value), "", Format(rs("Transaction_Date").value, "YYYY/MM/DD"))
                .TextMatrix(RecordNum, .ColIndex("CusName")) = IIf(IsNull(rs("CusName").value), "", rs("CusName").value)
                .TextMatrix(RecordNum, .ColIndex("StoreName")) = IIf(IsNull(rs("StoreName").value), "", rs("StoreName").value)
                .TextMatrix(RecordNum, .ColIndex("BillVal")) = IIf(IsNull(rs("Total").value), "", rs("Total").value)
                rs.MoveNext
            Next RecordNum

            .AutoSize 0, .Cols - 1, False
            LblVal(2).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("BillVal"), .Rows - 1, .ColIndex("BillVal"))
            LblBillCount(2).Caption = rs.RecordCount
        End If

        .AutoSize 0, .Cols - 1, False
    End With

    rs.Close
    Exit Sub
ErrTrap:

End Sub

Private Sub LoadMaintenence(StrDate As String)
    On Error GoTo ErrTrap

    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    Dim RecordNum As Integer
    LblVal(3).Caption = "0"
    LblBillCount(3).Caption = "0"

    With Me.FGMaintence
        .Clear flexClearScrollable, flexClearEverything
        .Rows = .FixedRows + 1
        '«·’Ì«‰…
        StrSQL = "select * From ReportMaintence where DateGoIN=" & SQLDate(CDate(StrDate), True) & ""
        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

        If Not (rs.BOF Or rs.EOF) Then
            FGMaintence.Rows = rs.RecordCount + 1

            For RecordNum = 1 To rs.RecordCount
                .TextMatrix(RecordNum, .ColIndex("Index")) = RecordNum
                .TextMatrix(RecordNum, .ColIndex("BillNum")) = IIf(IsNull(rs("MaintananceID").value), "", rs("MaintananceID").value)
                .TextMatrix(RecordNum, .ColIndex("Date")) = IIf(IsNull(rs("DateGoIN").value), "", Format(rs("DateGoIN").value, "YYYY/MM/DD"))
                .TextMatrix(RecordNum, .ColIndex("CusName")) = IIf(IsNull(rs("CusName").value), "", rs("CusName").value)
                .TextMatrix(RecordNum, .ColIndex("BillVal")) = IIf(IsNull(rs("Summtion").value), "", rs("Summtion").value)
                rs.MoveNext
            Next RecordNum

            .AutoSize 0, .Cols - 1, False
            LblVal(3).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("BillVal"), .Rows - 1, .ColIndex("BillVal"))
            LblBillCount(3).Caption = rs.RecordCount
        End If

        .AutoSize 0, .Cols - 1, False
    End With

    rs.Close
    Exit Sub
ErrTrap:

End Sub

Private Sub LoadExpenses(StrDate As String)
    On Error GoTo ErrTrap

    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    Dim RecordNum As Integer
    LblVal(4).Caption = "0"
    LblBillCount(4).Caption = "0"

    '«·„’—Ê›« 
    With Me.FGExpenses
        .Clear flexClearScrollable, flexClearEverything
        .Rows = .FixedRows + 1
    
StaticStrSql = " SELECT     TOP 100 PERCENT dbo.Notes.NoteID, dbo.Notes.NoteDate, dbo.Notes.NoteType, dbo.Notes.Note_Value, dbo.Notes.ExpensesID, dbo.ExpensesType.Name,"
StaticStrSql = StaticStrSql + "                      dbo.Notes.NoteSerial, dbo.Notes.Remark, dbo.TblUsers.UserName, dbo.Notes.BoxID, dbo.Notes.UserID, dbo.TblBoxesData.BoxName, dbo.BanksData.BankName,"
StaticStrSql = StaticStrSql + "                      dbo.Notes.BankID , dbo.Notes.ChqueNum, dbo.Notes.DueDate, dbo.Notes.NoteSerial1"
StaticStrSql = StaticStrSql + " FROM         dbo.TblUsers RIGHT OUTER JOIN"
StaticStrSql = StaticStrSql + "                      dbo.ExpensesType RIGHT OUTER JOIN"
StaticStrSql = StaticStrSql + "                      dbo.Notes ON dbo.ExpensesType.ID = dbo.Notes.ExpensesID ON dbo.TblUsers.UserID = dbo.Notes.UserID LEFT OUTER JOIN"
StaticStrSql = StaticStrSql + "                      dbo.TblBoxesData ON dbo.Notes.BoxID = dbo.TblBoxesData.BoxID LEFT OUTER JOIN"
StaticStrSql = StaticStrSql + "                      dbo.BanksData ON dbo.Notes.BankID = dbo.BanksData.BankID"
StaticStrSql = StaticStrSql + " WHERE    (dbo.Notes.NoteType = 53 or dbo.Notes.NoteType = 3)  AND (NOT (dbo.Notes.ExpensesID IS NULL)) "

        StaticStrSql = StaticStrSql + "  and dbo.Notes.NoteDate=" & SQLDate(CDate(StrDate), True) & ""
        StaticStrSql = StaticStrSql + "  ORDER BY dbo.Notes.NoteID"

        '    StrSQL = "select * From ExpensesReport where NoteDate=" & SQLDate(CDate(StrDate), True) & ""
    
        Set rs = New ADODB.Recordset
        rs.Open StaticStrSql, Cn, adOpenStatic, adLockOptimistic, adCmdText

        If Not (rs.BOF Or rs.EOF) Then
            FGExpenses.Rows = rs.RecordCount + 1

            For RecordNum = 1 To rs.RecordCount
                .TextMatrix(RecordNum, .ColIndex("Index")) = RecordNum
                .TextMatrix(RecordNum, .ColIndex("BillNum")) = IIf(IsNull(rs("NoteID").value), "", rs("NoteID").value)
                .TextMatrix(RecordNum, .ColIndex("NoteSerial")) = IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value)
                .TextMatrix(RecordNum, .ColIndex("Date")) = IIf(IsNull(rs("NoteDate").value), "", Format(rs("NoteDate").value, "YYYY/MM/DD"))
                .TextMatrix(RecordNum, .ColIndex("CusName")) = IIf(IsNull(rs("Name").value), "", rs("Name").value)
                .TextMatrix(RecordNum, .ColIndex("BillVal")) = IIf(IsNull(rs("Note_Value").value), "", rs("Note_Value").value)
                rs.MoveNext
            Next RecordNum

            .AutoSize 0, .Cols - 1, False
            LblVal(4).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("BillVal"), .Rows - 1, .ColIndex("BillVal"))
            LblBillCount(4).Caption = rs.RecordCount
        End If

        .AutoSize 0, .Cols - 1, False
    End With

    rs.Close
    Exit Sub
ErrTrap:
End Sub

Private Sub LoadCashing(StrDate As String)
    On Error GoTo ErrTrap

    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    Dim RecordNum As Integer
    LblVal(5).Caption = "0"
    LblBillCount(5).Caption = "0"

    '«·„ﬁ»Ê÷« 
    With Me.FGCashing
        .Clear flexClearScrollable, flexClearEverything
        .Rows = .FixedRows + 1
        StrSQL = "select * From CahingReport where NoteDate=" & SQLDate(CDate(StrDate), True) & ""
    
        StaticStrSql = " SELECT     TOP 100 PERCENT dbo.Notes.NoteID, dbo.Notes.NoteDate, dbo.Notes.Note_Value, dbo.Notes.CusID, dbo.TblCustemers.CusName, dbo.Notes.UserID,"
        StaticStrSql = StaticStrSql + " dbo.TblUsers.UserName, dbo.Notes.CashingType, dbo.Notes.Remark, dbo.Transactions.Transaction_ID, dbo.Transactions.Transaction_Serial,"
        StaticStrSql = StaticStrSql + " dbo.TransactionTypes.TransactionTypeName, dbo.TblBoxesData.BoxID, dbo.TblBoxesData.BoxName, dbo.Transactions.Transaction_Type, dbo.Notes.RevenuesID,"
        StaticStrSql = StaticStrSql + " dbo.TblRevenuesTypes.RevenuesName , dbo.Notes.NoteSerial, dbo.Notes.NoteSerial1, dbo.BanksData.BankName, dbo.Notes.ChqueNum, dbo.Notes.DueDate"
        StaticStrSql = StaticStrSql + " FROM         dbo.TblBoxesData RIGHT OUTER JOIN"
        StaticStrSql = StaticStrSql + " dbo.TblRevenuesTypes RIGHT OUTER JOIN"
        StaticStrSql = StaticStrSql + " dbo.BanksData RIGHT OUTER JOIN"
        StaticStrSql = StaticStrSql + " dbo.TblUsers INNER JOIN"
        StaticStrSql = StaticStrSql + " dbo.Notes ON dbo.TblUsers.UserID = dbo.Notes.UserID ON dbo.BanksData.BankID = dbo.Notes.BankID ON"
        StaticStrSql = StaticStrSql + " dbo.TblRevenuesTypes.RevenuesID = dbo.Notes.RevenuesID ON dbo.TblBoxesData.BoxID = dbo.Notes.BoxID LEFT OUTER JOIN"
        StaticStrSql = StaticStrSql + " dbo.TblCustemers ON dbo.Notes.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
        StaticStrSql = StaticStrSql + " dbo.Transactions LEFT OUTER JOIN"
        StaticStrSql = StaticStrSql + " dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type ON"
        StaticStrSql = StaticStrSql + " dbo.Notes.Transaction_ID = dbo.Transactions.Transaction_ID"
        StaticStrSql = StaticStrSql + " Where (dbo.Notes.NoteType = 4)"
        StaticStrSql = StaticStrSql + "  and dbo.Notes.NoteDate=" & SQLDate(CDate(StrDate), True) & ""
        StaticStrSql = StaticStrSql + "  ORDER BY dbo.Notes.NoteDate"

        Set rs = New ADODB.Recordset
        rs.Open StaticStrSql, Cn, adOpenStatic, adLockOptimistic, adCmdText

        If Not (rs.BOF Or rs.EOF) Then
            FGCashing.Rows = rs.RecordCount + 1

            For RecordNum = 1 To rs.RecordCount
                .TextMatrix(RecordNum, .ColIndex("Index")) = RecordNum
                .TextMatrix(RecordNum, .ColIndex("BillNum")) = IIf(IsNull(rs("NoteID").value), "", rs("NoteID").value)
                .TextMatrix(RecordNum, .ColIndex("NoteSerial")) = IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value)
                .TextMatrix(RecordNum, .ColIndex("Date")) = IIf(IsNull(rs("NoteDate").value), "", DisplayDate(rs("NoteDate").value))

                If Not IsNull(IsNull(rs("CashingType").value)) Then
                    If rs("CashingType").value = 0 Then
                        .TextMatrix(RecordNum, .ColIndex("CashingType")) = "„‰ ⁄„Ì·"
                        .TextMatrix(RecordNum, .ColIndex("CusName")) = IIf(IsNull(rs("CusName").value), "", rs("CusName").value)
                    ElseIf rs("CashingType").value = 1 Then
                        .TextMatrix(RecordNum, .ColIndex("CashingType")) = "„‰ „Ê—œ"
                        .TextMatrix(RecordNum, .ColIndex("CusName")) = IIf(IsNull(rs("CusName").value), "", rs("CusName").value)
                    ElseIf rs("CashingType").value = 2 Then
                        .TextMatrix(RecordNum, .ColIndex("CashingType")) = "„ ⁄·ﬁ« "
                        .TextMatrix(RecordNum, .ColIndex("CusName")) = IIf(IsNull(rs("CusName").value), "", rs("CusName").value)
                    ElseIf rs("CashingType").value = 3 Then
                        .TextMatrix(RecordNum, .ColIndex("CashingType")) = "≈Ì—«œ«  √Œ—Ï"
                        .TextMatrix(RecordNum, .ColIndex("CusName")) = IIf(IsNull(rs("RevenuesName").value), "", rs("RevenuesName").value)
                    End If
                End If

                If Not (IsNull(rs("Transaction_ID").value)) Then
                    .Cell(flexcpChecked, RecordNum, .ColIndex("BolTransaction")) = flexChecked
                    .TextMatrix(RecordNum, .ColIndex("TransactionSerial")) = IIf(IsNull(rs("Transaction_Serial").value), "", rs("Transaction_Serial").value)
                    .TextMatrix(RecordNum, .ColIndex("TranactionTypeName")) = IIf(IsNull(rs("TransactionTypeName").value), "", rs("TransactionTypeName").value)
                End If
            
                .TextMatrix(RecordNum, .ColIndex("BillVal")) = IIf(IsNull(rs("Note_Value").value), "", rs("Note_Value").value)
                rs.MoveNext
            Next RecordNum

            .AutoSize 0, .Cols - 1, False
            LblVal(5).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("BillVal"), .Rows - 1, .ColIndex("BillVal"))
            LblBillCount(5).Caption = rs.RecordCount
        End If

        .AutoSize 0, .Cols - 1, False
    End With

    rs.Close
    Exit Sub
ErrTrap:
End Sub

Private Sub LoadSummary()
    On Error GoTo ErrTrap

    With FGSummery
        .TextMatrix(1, .ColIndex("Type")) = "≈Ã„«·Ì ﬁÌ„… «·„»Ì⁄« :"
        .TextMatrix(1, .ColIndex("Value")) = LblVal(0).Caption
        .TextMatrix(2, .ColIndex("Type")) = "≈Ã„«·Ì Œ’Ê„«  «·„»Ì⁄« :"
        .TextMatrix(2, .ColIndex("Value")) = LblDiscount(0).Caption
        .TextMatrix(3, .ColIndex("Type")) = "’«›Ì ﬁÌ„… «·„»Ì⁄« :"
        .TextMatrix(3, .ColIndex("Value")) = LblTotal(0).Caption
        .TextMatrix(4, .ColIndex("Type")) = "≈Ã„«·Ì ﬁÌ„… «·„‘ —Ì« :"
        .TextMatrix(4, .ColIndex("Value")) = LblVal(1).Caption
        .TextMatrix(5, .ColIndex("Type")) = "≈Ã„«·Ì Œ’Ê„«  «·„‘ —Ì« :"
        .TextMatrix(5, .ColIndex("Value")) = LblDiscount(1).Caption
        .TextMatrix(6, .ColIndex("Type")) = "’«›Ì ﬁÌ„… «·„‘ —Ì« :"
        .TextMatrix(6, .ColIndex("Value")) = LblTotal(1).Caption
        .TextMatrix(7, .ColIndex("Type")) = "≈Ã„«·Ì ﬁÌ„… „— Ã⁄ «·„»Ì⁄« :"
        .TextMatrix(7, .ColIndex("Value")) = LblVal(8).Caption
        .TextMatrix(8, .ColIndex("Type")) = "≈Ã„«·Ì ﬁÌ„… „— Ã⁄ «·„‘ —Ì« :"
        .TextMatrix(8, .ColIndex("Value")) = LblVal(2).Caption
        .TextMatrix(9, .ColIndex("Type")) = "≈Ã„«·Ì ⁄„·Ì«  «·’Ì«‰…:"
        .TextMatrix(9, .ColIndex("Value")) = LblVal(3).Caption
        .TextMatrix(10, .ColIndex("Type")) = "≈Ã„«·Ì «·„’—Ê›« :"
        .TextMatrix(10, .ColIndex("Value")) = LblVal(4).Caption
    
        .TextMatrix(11, .ColIndex("Type")) = "≈Ã„«·Ì «·„œ›Ê⁄« :"
        .TextMatrix(11, .ColIndex("Value")) = LblVal(6).Caption
    
        .TextMatrix(12, .ColIndex("Type")) = "≈Ã„«·Ì «·„ﬁ»Ê÷« :"
        .TextMatrix(12, .ColIndex("Value")) = LblVal(5).Caption
    
        LblVal(7).Caption = val(.TextMatrix(3, .ColIndex("Value"))) - val(.TextMatrix(6, .ColIndex("Value"))) - val(.TextMatrix(7, .ColIndex("Value"))) + val(.TextMatrix(8, .ColIndex("Value"))) + val(.TextMatrix(9, .ColIndex("Value"))) - val(.TextMatrix(10, .ColIndex("Value"))) - val(.TextMatrix(11, .ColIndex("Value"))) + val(.TextMatrix(12, .ColIndex("Value")))

        If LblVal(7).Caption <= 0 Then
            LblVal(7).ForeColor = vbRed
        Else
            LblVal(7).ForeColor = vbBlack
        End If

        '    .Cell(flexcpAlignment, 1, .ColIndex("Type"), .Rows - 1, .ColIndex("Type")) = flexAlignCenterCenter
        .AutoSize .ColIndex("Type"), .Cols - 1, False
    End With

    Exit Sub
ErrTrap:
End Sub

Private Sub LoadPayments(StrDate As String)
    On Error GoTo ErrTrap

    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    Dim RecordNum As Integer
    LblVal(6).Caption = "0"
    LblBillCount(6).Caption = "0"

    '«·„œ›Ê⁄« 
    With Me.FgPayments
        .Clear flexClearScrollable, flexClearEverything
        .Rows = .FixedRows + 1
        ' StrSQL = "select * From PaymentsReport where NoteDate=" & SQLDate(CDate(StrDate), True) & ""
 
        StaticStrSql = " SELECT     TOP 100 PERCENT dbo.Notes.NoteID, dbo.Notes.NoteDate, dbo.Notes.Note_Value, dbo.Notes.CusID, dbo.TblCustemers.CusName, dbo.Notes.UserID,"
        StaticStrSql = StaticStrSql + " dbo.TblUsers.UserName, dbo.Notes.CashingType, dbo.Notes.NoteSerial, dbo.Notes.Remark, dbo.Transactions.Transaction_ID, dbo.Transactions.Transaction_Serial,"
        StaticStrSql = StaticStrSql + " dbo.TransactionTypes.TransactionTypeName, dbo.TblBoxesData.BoxID, dbo.TblBoxesData.BoxName, dbo.Transactions.Transaction_Type, dbo.Notes.NoteSerial1,"
        StaticStrSql = StaticStrSql + " dbo.BanksData.BankName , dbo.Notes.ChqueNum, dbo.Notes.DueDate"
        StaticStrSql = StaticStrSql + " FROM         dbo.Transactions RIGHT OUTER JOIN"
        StaticStrSql = StaticStrSql + "  dbo.TblBoxesData RIGHT OUTER JOIN"
        StaticStrSql = StaticStrSql + "  dbo.TblUsers RIGHT OUTER JOIN"
        StaticStrSql = StaticStrSql + "   dbo.TblCustemers RIGHT OUTER JOIN"
        StaticStrSql = StaticStrSql + "   dbo.BanksData RIGHT OUTER JOIN"
        StaticStrSql = StaticStrSql + "   dbo.Notes ON dbo.BanksData.BankID = dbo.Notes.BankID ON dbo.TblCustemers.CusID = dbo.Notes.CusID ON dbo.TblUsers.UserID = dbo.Notes.UserID ON"
        StaticStrSql = StaticStrSql + "  dbo.TblBoxesData.BoxID = dbo.Notes.BoxID ON dbo.Transactions.Transaction_ID = dbo.Notes.Transaction_ID LEFT OUTER JOIN"
        StaticStrSql = StaticStrSql + "  dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type"
        StaticStrSql = StaticStrSql + "  where dbo.Notes.NoteDate =" & SQLDate(CDate(StrDate), True) & ""
        StaticStrSql = StaticStrSql + "  AND (dbo.Notes.NoteType = 5)"
        StaticStrSql = StaticStrSql + "  ORDER BY dbo.Notes.NoteDate"

        Set rs = New ADODB.Recordset
        rs.Open StaticStrSql, Cn, adOpenStatic, adLockOptimistic, adCmdText

        If Not (rs.BOF Or rs.EOF) Then
            FgPayments.Rows = rs.RecordCount + 1

            For RecordNum = 1 To rs.RecordCount
                .TextMatrix(RecordNum, .ColIndex("Index")) = RecordNum
                .TextMatrix(RecordNum, .ColIndex("BillNum")) = IIf(IsNull(rs("NoteID").value), "", rs("NoteID").value)
                .TextMatrix(RecordNum, .ColIndex("NoteSerial")) = IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value)
                .TextMatrix(RecordNum, .ColIndex("Date")) = IIf(IsNull(rs("NoteDate").value), "", DisplayDate(rs("NoteDate").value))

                If rs("CashingType").value = 0 Then
                    .TextMatrix(RecordNum, .ColIndex("CashingType")) = "„‰ ⁄„Ì·"
                    .TextMatrix(RecordNum, .ColIndex("CusName")) = IIf(IsNull(rs("CusName").value), "", rs("CusName").value)
                ElseIf rs("CashingType").value = 1 Then
                    .TextMatrix(RecordNum, .ColIndex("CashingType")) = "„‰ „Ê—œ"
                    .TextMatrix(RecordNum, .ColIndex("CusName")) = IIf(IsNull(rs("CusName").value), "", rs("CusName").value)
                ElseIf rs("CashingType").value = 2 Then
                    .TextMatrix(RecordNum, .ColIndex("CashingType")) = "„ ⁄·ﬁ« "
                    .TextMatrix(RecordNum, .ColIndex("CusName")) = IIf(IsNull(rs("CusName").value), "", rs("CusName").value)
                End If

                .TextMatrix(RecordNum, .ColIndex("CusName")) = IIf(IsNull(rs("CusName").value), "", rs("CusName").value)
                .TextMatrix(RecordNum, .ColIndex("BillVal")) = IIf(IsNull(rs("Note_Value").value), "", rs("Note_Value").value)

                If Not (IsNull(rs("Transaction_ID").value)) Then
                    .Cell(flexcpChecked, RecordNum, .ColIndex("BolTransaction")) = flexChecked
                    .TextMatrix(RecordNum, .ColIndex("TransactionSerial")) = IIf(IsNull(rs("Transaction_Serial").value), "", rs("Transaction_Serial").value)
                    .TextMatrix(RecordNum, .ColIndex("TranactionTypeName")) = IIf(IsNull(rs("TransactionTypeName").value), "", rs("TransactionTypeName").value)
                End If

                rs.MoveNext
            Next RecordNum

            .AutoSize 0, .Cols - 1, False
            LblVal(6).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("BillVal"), .Rows - 1, .ColIndex("BillVal"))
            LblBillCount(6).Caption = rs.RecordCount
        End If

        .AutoSize 0, .Cols - 1, False
    End With

    rs.Close
    Exit Sub
ErrTrap:

End Sub

Private Sub LoadReturnSalling(StrDate As String)
    'On Error GoTo ErrTrap

    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    Dim RecordNum As Integer
    LblVal(8).Caption = "0"
    LblBillCount(7).Caption = "0"

    '„— Ã⁄ «·„»Ì⁄« 
    With Me.FgReturnSalling
        .Clear flexClearScrollable, flexClearEverything
        .Rows = .FixedRows + 1
        StaticStrSql = "SELECT     dbo.Transactions.Transaction_ID, dbo.Transactions.Transaction_Date, dbo.Transactions.Transaction_Type, "
        StaticStrSql = StaticStrSql & "  SUM(dbo.Transaction_Details.showPrice * dbo.Transaction_Details.ShowQty) AS Total, dbo.TblStore.StoreName, dbo.TblCustemers.CusName,"
        StaticStrSql = StaticStrSql & "   dbo.Transactions.Transaction_serial l, dbo.Transactions.NoteSerial, dbo.Transactions.NoteSerial1 "
        StaticStrSql = StaticStrSql & " FROM         dbo.TblStore INNER JOIN"
        StaticStrSql = StaticStrSql & " dbo.TblCustemers INNER JOIN"
        StaticStrSql = StaticStrSql & "  dbo.Transactions ON dbo.TblCustemers.CusID = dbo.Transactions.CusID ON dbo.TblStore.StoreID = dbo.Transactions.StoreID INNER JOIN"
        StaticStrSql = StaticStrSql & "  dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID"
        '
    
        StaticStrSql = StaticStrSql + " where Transactions.Transaction_Date=" & SQLDate(CDate(StrDate), True) & ""
  
  If DCBranches2.BoundText <> "" Then
  StaticStrSql = StaticStrSql & "   and dbo.Transactions.BranchId= " & val(DCBranches2.BoundText)
       
  End If
    If val(DcboStoresx(2).BoundText) <> 0 And (DcboStoresx(2).Text) <> "" Then
                  StaticStrSql = StaticStrSql + " and  dbo.Transactions.StoreId=" & DcboStoresx(2).BoundText
     End If


        StaticStrSql = StaticStrSql & " GROUP BY dbo.Transactions.Transaction_ID, dbo.Transactions.Transaction_Date, dbo.Transactions.Transaction_Type, dbo.TblStore.StoreName, dbo.TblCustemers.CusName,"
        StaticStrSql = StaticStrSql & "   dbo.Transactions.Transaction_serial , dbo.Transactions.NoteSerial, dbo.Transactions.NoteSerial1"
        StaticStrSql = StaticStrSql & " HAVING      (dbo.Transactions.Transaction_Type = 9)"

        'StaticStrSql=StaticStrSql & " HAVING      (dbo.Transactions.Transaction_Type = 5)"
        ' StrSQL = "SELECT * From QryReturnSalling"
        ' StrSQL = StrSQL + " where Transaction_Date=" & SQLDate(CDate(StrDate), True) & ""
  REsalesSql = StaticStrSql
        Set rs = New ADODB.Recordset
        rs.Open StaticStrSql, Cn, adOpenStatic, adLockOptimistic, adCmdText

        If Not (rs.BOF Or rs.EOF) Then
            .Rows = rs.RecordCount + 1

            For RecordNum = 1 To rs.RecordCount
                .TextMatrix(RecordNum, .ColIndex("Index")) = RecordNum
                .TextMatrix(RecordNum, .ColIndex("BillNum")) = IIf(IsNull(rs("Transaction_ID").value), "", rs("Transaction_ID").value)
                ' .TextMatrix(RecordNum, .ColIndex("Serial")) = IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value)
                .TextMatrix(RecordNum, .ColIndex("Serial")) = IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value)
            
                .TextMatrix(RecordNum, .ColIndex("Date")) = IIf(IsNull(rs("Transaction_Date").value), "", Format(rs("Transaction_Date").value, "YYYY/MM/DD"))
                .TextMatrix(RecordNum, .ColIndex("CusName")) = IIf(IsNull(rs("CusName").value), "", rs("CusName").value)
                .TextMatrix(RecordNum, .ColIndex("StoreName")) = IIf(IsNull(rs("StoreName").value), "", rs("StoreName").value)
                .TextMatrix(RecordNum, .ColIndex("BillVal")) = IIf(IsNull(rs("Total").value), "", rs("Total").value)
                rs.MoveNext
            Next RecordNum

            .AutoSize 0, .Cols - 1, False
            LblVal(8).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("BillVal"), .Rows - 1, .ColIndex("BillVal"))
            LblBillCount(7).Caption = rs.RecordCount
        End If

        .AutoSize 0, .Cols - 1, False
    End With

    rs.Close
    Exit Sub
ErrTrap:

End Sub

Private Sub LoadItemsTransactions()
    Dim rs As ADODB.Recordset
    Dim i As Integer

    If Me.DcboStores.BoundText = "" Then
        Exit Sub
    End If

    With Me.FgItems
        .Clear flexClearScrollable, flexClearEverything
        .Rows = .FixedRows
    End With

    Set rs = GetItemsRsTransactions(val(Me.DcboStores.BoundText), Me.DtpTransDate(8).value)

    If rs Is Nothing Then
        Exit Sub
    End If

    If rs.BOF Or rs.EOF Then
        Exit Sub
    End If

    With Me.FgItems
        .Rows = .FixedRows + rs.RecordCount

        For i = .FixedRows To rs.RecordCount
            .TextMatrix(i, .ColIndex("Serial")) = i
            .TextMatrix(i, .ColIndex("ItemID")) = IIf(IsNull(rs("ItemID").value), "", rs("ItemID").value)
            .TextMatrix(i, .ColIndex("ItemCode")) = IIf(IsNull(rs("ItemCode").value), "", rs("ItemCode").value)
            .TextMatrix(i, .ColIndex("GroupName")) = IIf(IsNull(rs("GroupName").value), "", rs("GroupName").value)
            .TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(rs("ItemName").value), "", rs("ItemName").value)
            .TextMatrix(i, .ColIndex("ItemCase")) = IIf(IsNull(rs("ItemCase").value), "", rs("ItemCase").value)
            .TextMatrix(i, .ColIndex("SumQuantity_1")) = IIf(IsNull(rs("SumQuantity_1").value), 0, rs("SumQuantity_1").value)
            .TextMatrix(i, .ColIndex("SumQuantity1")) = IIf(IsNull(rs("SumQuantity1").value), 0, rs("SumQuantity1").value)
            rs.MoveNext
        Next i

        Me.lbl(23).Caption = .Aggregate(flexSTCount, .FixedRows, .ColIndex("ItemID"), .Rows - 1, .ColIndex("ItemId")) & " ’‰› "
        .AutoSize 0, .Cols - 1, False
    End With

    rs.Close
    Set rs = Nothing
End Sub

Private Sub ChangeLang()
    Me.Caption = "Daily Report"
    Me.TabMain.TabCaption(0) = "Sales"
    Me.TabMain.TabCaption(1) = "Purchase"
    Me.TabMain.TabCaption(2) = "Sales Return"
    Me.TabMain.TabCaption(3) = "Purchase Return"
    Me.TabMain.TabCaption(4) = "Maintaince"
    Me.TabMain.TabCaption(5) = "Expenses"
    Me.TabMain.TabCaption(6) = "Notes Payable"
    Me.TabMain.TabCaption(7) = "Notes Receivable"
    Me.TabMain.TabCaption(8) = "Summary"
    Me.TabMain.TabCaption(9) = "Stock"
    Me.TabMain.TabCaption(11) = "Boxes"
    Me.TabMain.TabCaption(10) = "Customers"
    cmdPrint(0).Caption = "Print"
    CmdPrintBuy(0).Caption = "Print"
    CmdReturnSalling(0).Caption = "Print"
    CmdReturn(0).Caption = "Print"
    CmdExpenses(0).Caption = "Print"
    CmdPayment(0).Caption = "Print"
    CmdCaching(0).Caption = "Print"
    CmdSummery(0).Caption = "Print"
  
Cmd(0).Caption = "Print"
Cmd(2).Caption = "Print"
Cmd(3).Caption = "Print"


CMDShow(0).Caption = "Display Data"
CMDShow(1).Caption = "Display Data"
CMDShow(2).Caption = "Display Data"
CMDShow(3).Caption = "Display Data"
CMDShow(4).Caption = "Display Data"
CMDShow(5).Caption = "Display Data"
CMDShow(6).Caption = "Display Data"
CMDShow(7).Caption = "Display Data"
CMDShow(8).Caption = "Display Data"
CMDShow(9).Caption = "Display Data"
CMDShow(10).Caption = "Display Data"

End Sub

Private Sub LoadCustomers()
    Dim RsAllCustomers As ADODB.Recordset
    Dim RsCus As ADODB.Recordset
    Dim StrSQL As String
    Dim RecordNum As Integer
    Dim i As Integer, j As Integer
    Dim SngBeforeAccount  As Single
    Dim SngAfterAccount As Single
    Dim SngDayAccount As Single
    Dim SngCurrentAccount As Single
    Dim rs As ADODB.Recordset
    Dim LngFindRow As Long
Dim EmpCode  As String
    Dim StrText As String
                Dim FirstPeriod As Date
                Dim CUstAccount_Code As String
 getFirstPeriodDateInthisYear FirstPeriod
    On Error GoTo ErrTrap

    With Me.FgCustomers
        .Clear flexClearScrollable, flexClearEverything
        .Rows = .FixedRows + 1

        If SystemOptions.SysDataBaseType = SQLServerDataBase Then
            StrSQL = "SELECT DISTINCT  Transactions.CusID, TblCustemers.CusName FROM TblCustemers INNER JOIN " & "Transactions ON TblCustemers.CusID = Transactions.CusID Where Transaction_date=" & SQLDate(Me.DtpCustomers.value, True) & ""
            StrSQL = StrSQL + " UNION "
            StrSQL = StrSQL + " SELECT DISTINCT  NOTES.CusID, TblCustemers.CusName FROM " & " TblCustemers INNER JOIN NOTES  ON TblCustemers.CusID =NOTES.CusID " & "Where NOTEDATE=" & SQLDate(Me.DtpCustomers.value, True) & ""
        ElseIf SystemOptions.SysDataBaseType = AccessDataBase Then
            StrSQL = "SELECT DISTINCT  Transactions.CusID, TblCustemers.CusName FROM TblCustemers INNER JOIN " & "Transactions ON TblCustemers.CusID = Transactions.CusID Where Transaction_date=" & SQLDate(Me.DtpCustomers.value, True) & ""
            StrSQL = StrSQL + " UNION "
            StrSQL = StrSQL + " SELECT DISTINCT  NOTES.CusID, TblCustemers.CusName FROM " & " TblCustemers INNER JOIN NOTES  ON TblCustemers.CusID =NOTES.CusID " & "Where NOTEDATE=" & SQLDate(Me.DtpCustomers.value, True) & ""
        End If

        Set RsAllCustomers = New ADODB.Recordset
        RsAllCustomers.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If RsAllCustomers.RecordCount > 0 Then
            .Rows = RsAllCustomers.RecordCount + 1

            For i = 1 To RsAllCustomers.RecordCount
                .TextMatrix(i, .ColIndex("CusID")) = RsAllCustomers("CusID").value
    GetTblCustemersCode , , .TextMatrix(i, .ColIndex("CusID")), EmpCode
    .TextMatrix(i, .ColIndex("FullCode")) = EmpCode
                .TextMatrix(i, .ColIndex("CusName")) = RsAllCustomers("CusName").value
                '«·—’Ìœ ﬁ»· «·ÌÊ„
                DTPicker1.value = DateAdd("d", -1, DtpCustomers.value)
             '   SngBeforeAccount = GetCustomerAccount(RsAllCustomers("CusID").Value, True, Me.DTPicker1.Value)

             '   If SngBeforeAccount < 0 Then
             '       StrText = Abs(SngBeforeAccount) & "    „œÌ‰ "
             '   ElseIf SngBeforeAccount > 0 Then
             '       StrText = Abs(SngBeforeAccount) & "    œ«∆‰ "
             '   ElseIf SngBeforeAccount = 0 Then
             '       StrText = "    Œ«·’ "
             '   End If
             '
CUstAccount_Code = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(RsAllCustomers("CusID").value), "Account_code")

 
                StrText = GetActualAccountBalance(CUstAccount_Code, 0, FirstPeriod, DTPicker1.value)
            

                .TextMatrix(i, .ColIndex("DayBeforeAccount")) = StrText
                '----------------------------------------------------------------------
                '«·—’Ìœ »⁄œ «·ÌÊ„
'                SngAfterAccount = GetCustomerAccount(RsAllCustomers("CusID").value, True, Me.DtpCustomers.value + 1)
'
'                If SngAfterAccount < 0 Then
'                    StrText = Abs(SngAfterAccount) & "    „œÌ‰ "
'                ElseIf SngAfterAccount > 0 Then
'                    StrText = Abs(SngAfterAccount) & "    œ«∆‰ "
''                ElseIf SngAfterAccount = 0 Then
 '                   StrText = "    Œ«·’ "
 '               End If

 

 
                StrText = GetActualAccountBalance(CUstAccount_Code, 0, FirstPeriod, DtpCustomers.value)
            
            
                .TextMatrix(i, .ColIndex("DayAfterAccount")) = StrText
                '-----------------------------------------------------------------------
                
                                '-----------------------------------------------------------------------
                '«·—’Ìœ «·Õ«·Ï
               ' SngCurrentAccount = GetCustomerAccount(RsAllCustomers("CusID").value, True)

               ' If SngCurrentAccount < 0 Then
               '     StrText = Abs(SngCurrentAccount) & "    „œÌ‰ "
               ' ElseIf SngCurrentAccount > 0 Then
               '     StrText = Abs(SngCurrentAccount) & "     œ«∆‰ "
               ' ElseIf SngCurrentAccount = 0 Then
               '     StrText = "     Œ«·’ "
               ' End If

                .TextMatrix(i, .ColIndex("CurrentAccount")) = StrText
                '-----------------------------------------------------------------------
                
                'ÕÃ„  ⁄«„· «·ÌÊ„
                SngDayAccount = SngAfterAccount - SngBeforeAccount

                If SngDayAccount < 0 Then
                    StrText = Format(Abs(SngDayAccount), SystemOptions.SysDefCurrencyForamt) & "    „œÌ‰ "
                ElseIf SngDayAccount > 0 Then
                    StrText = Format(Abs(SngDayAccount), SystemOptions.SysDefCurrencyForamt) & "     œ«∆‰ "
                ElseIf SngDayAccount = 0 Then
                    StrText = ""
                End If

                .TextMatrix(i, .ColIndex("DayAccount")) = StrText
                '-----------------------------------------------------------------------
                StrSQL = ""

                RsAllCustomers.MoveNext
            Next i

            '---------------------------------------------------------------------------
            '«·„»Ì⁄« 
            StrSQL = "SELECT Sum(TOTALAfterTAx)As SumX,CusID,CusName"
            StrSQL = StrSQL + " From "
            StrSQL = StrSQL + " ( "
            StrSQL = StrSQL + " Select * From ReportSallingTime where " & "Transaction_Date = " & SQLDate(Me.DtpCustomers.value, True) & ""
            StrSQL = StrSQL + ")XTable Group BY CusID,CusName"
            Set rs = New ADODB.Recordset
            rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

            If Not (rs.BOF Or rs.EOF) Then
                rs.MoveFirst

                For i = 1 To rs.RecordCount
                    LngFindRow = .FindRow(rs("CusID").value, .FixedRows, .ColIndex("CusID"), False, True)

                    If LngFindRow <> -1 Then
                        .TextMatrix(LngFindRow, .ColIndex("CustomerSales")) = IIf(IsNull(rs("SumX").value), "", rs("SumX").value)
                    End If

                    rs.MoveNext
                Next i

            End If

            '----------------------------------------------------------------------------
            '«·„‘ —Ì« 
           ' StrSQL = "SELECT Sum(TOTALAfterTAx)As SumX,CusID,CusName"
           ' StrSQL = StrSQL + " From "
           ' StrSQL = StrSQL + " ( "
           ' StrSQL = StrSQL + " Select * From ReportBuyTime_Client where " & "Transaction_Date = " & SQLDate(Me.DtpCustomers.value, True) & ""
           ' StrSQL = StrSQL + ")XTable Group BY CusID,CusName"
         StrSQL = "   SELECT     SUM(dbo.Transactions.Transaction_NetValue) AS SumX, dbo.Transactions.CusID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee"
         StrSQL = StrSQL + "    FROM         dbo.Transactions INNER JOIN"
         StrSQL = StrSQL + "             dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID"
         StrSQL = StrSQL + "      WHERE     (dbo.Transactions.Transaction_Type = 1) AND (dbo.Transactions.Transaction_Date = " & SQLDate(Me.DtpCustomers.value, True) & ") OR"
         StrSQL = StrSQL + "             (dbo.Transactions.Transaction_Type = 22)"
         StrSQL = StrSQL + "   GROUP BY dbo.Transactions.CusID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee"
            Set rs = New ADODB.Recordset
            rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

            If Not (rs.BOF Or rs.EOF) Then
                rs.MoveFirst

                For i = 1 To rs.RecordCount
                    LngFindRow = .FindRow(rs("CusID").value, .FixedRows, .ColIndex("CusID"), False, True)

                    If LngFindRow <> -1 Then
                        .TextMatrix(LngFindRow, .ColIndex("CustomerPur")) = IIf(IsNull(rs("SumX").value), "", rs("SumX").value)
                    End If

                    rs.MoveNext
                Next i

            End If

            '----------------------------------------------------------------------------------
            '«·„ﬁ»Ê÷« 
            StrSQL = "Select Sum(Note_Value) as SumX ,CusID,CusName " & "From ( "
            StrSQL = StrSQL + " Select * From CahingReport " & "Where NoteDate=" & SQLDate(Me.DtpCustomers.value, True) & " )XTable Group By CusID,CusName "
            Set rs = New ADODB.Recordset
            rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

            If Not (rs.BOF Or rs.EOF) Then
                rs.MoveFirst

                For i = 1 To rs.RecordCount
                    LngFindRow = .FindRow(rs("CusID").value, .FixedRows, .ColIndex("CusID"), False, True)

                    If LngFindRow <> -1 Then
                        .TextMatrix(LngFindRow, .ColIndex("CustomerCash")) = IIf(IsNull(rs("SumX").value), "", rs("SumX").value)
                    End If

                    rs.MoveNext
                Next i

            End If

            '----------------------------------------------------------------------------------
            '«·„œ›Ê⁄« 
            StrSQL = "Select Sum(Note_Value) as SumX ,CusID,CusName " & "From ( "
            StrSQL = StrSQL + " Select * From PaymentsReport " & "Where NoteDate=" & SQLDate(Me.DtpCustomers.value, True) & " )XTable Group By CusID,CusName "
            Set rs = New ADODB.Recordset
            rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

            If Not (rs.BOF Or rs.EOF) Then
                rs.MoveFirst

                For i = 1 To rs.RecordCount
                    LngFindRow = .FindRow(rs("CusID").value, .FixedRows, .ColIndex("CusID"), False, True)

                    If LngFindRow <> -1 Then
                        .TextMatrix(LngFindRow, .ColIndex("CustomerPayment")) = IIf(IsNull(rs("SumX").value), "", rs("SumX").value)
                    End If

                    rs.MoveNext
                Next i

            End If

            '--------------------------------------------------------------------------------
            For i = .FixedRows To .Rows - 1

                If val(.TextMatrix(i, .ColIndex("CusID"))) <> 0 Then
                    StrSQL = "SELECT ReportSallingTime.TransactionComment "
                    StrSQL = StrSQL + " From ReportSallingTime "
                    StrSQL = StrSQL + " WHERE (ReportSallingTime.Transaction_Date)=" & SQLDate(Me.DtpCustomers.value, True) & ""
                    StrSQL = StrSQL + " AND ReportSallingTime.CusID=" & val(.TextMatrix(i, .ColIndex("CusID")))
                    Set rs = New ADODB.Recordset
                    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                    If Not (rs.BOF Or rs.EOF) Then
                        StrText = ""

                        For j = 1 To rs.RecordCount
                            StrText = StrText & IIf(IsNull(rs("TransactionComment").value), "", rs("TransactionComment").value)
                            rs.MoveNext
                        Next j

                        .TextMatrix(i, .ColIndex("Comments")) = Trim$(StrText)
                    End If

                    rs.Close
                    Set rs = Nothing
                End If

            Next i

        End If

        .AutoSize 0, .Cols - 1, False
    End With

    Exit Sub
ErrTrap:
End Sub
 
Private Sub PrintDayCustomersAccounts()
    Dim xApp As New CRAXDRT.Application
    Dim xReport As New CRAXDRT.Report
    Dim StrFileName As String
    Dim XViewr As ClsReportViewer
    Dim StrSQL As String
    Dim i  As Integer
    Dim CViewer As ClsReportViewer
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim StrRS As ADODB.Recordset
    On Error GoTo ErrTrap

    DB_CreateTable "TempPrintCustomersDayAccount"
    DB_CreateField "TempPrintCustomersDayAccount", "CusID", adInteger, adColNullable, , , , False, False
    DB_CreateField "TempPrintCustomersDayAccount", "CusName", adVarWChar, adColNullable, 100, , , False, False
    DB_CreateField "TempPrintCustomersDayAccount", "BeforeDayAccount", adVarWChar, adColNullable, 100, , , False, False
    DB_CreateField "TempPrintCustomersDayAccount", "DayAccount", adVarWChar, adColNullable, 100, , , False, False
    DB_CreateField "TempPrintCustomersDayAccount", "AfterDayAccount", adVarWChar, adColNullable, 100, , , False, False
    DB_CreateField "TempPrintCustomersDayAccount", "CurrentAccount", adVarWChar, adColNullable, 100, , , False, False

    DB_CreateField "TempPrintCustomersDayAccount", "CustomerSales", adVarWChar, adColNullable, 100, , , False, False
    DB_CreateField "TempPrintCustomersDayAccount", "CustomerPur", adVarWChar, adColNullable, 100, , , False, False
    DB_CreateField "TempPrintCustomersDayAccount", "CustomerCash", adVarWChar, adColNullable, 100, , , False, False
    DB_CreateField "TempPrintCustomersDayAccount", "CustomerPayment", adVarWChar, adColNullable, 100, , , False, False
    DB_CreateField "TempPrintCustomersDayAccount", "Comments", adVarWChar, adColNullable, 255, , , False, False

    StrSQL = "Delete From TempPrintCustomersDayAccount"
    Cn.Execute StrSQL, , adExecuteNoRecords

    Set rs = New ADODB.Recordset
    rs.Open "TempPrintCustomersDayAccount", Cn, adOpenStatic, adLockOptimistic, adCmdTable

    With FgCustomers

        For i = 1 To Me.FgCustomers.Rows - 1
            rs.AddNew
            rs("TempCol").value = 1
            rs("CusID").value = val(.TextMatrix(i, .ColIndex("CusID")))
            rs("CusName").value = Trim$(.TextMatrix(i, .ColIndex("CusName")))
            rs("BeforeDayAccount").value = Trim$(.TextMatrix(i, .ColIndex("DayBeforeAccount")))
            rs("DayAccount").value = IIf(Trim$(.TextMatrix(i, .ColIndex("DayAccount"))) = "", Null, Trim$(.TextMatrix(i, .ColIndex("DayAccount"))))
            rs("AfterDayAccount").value = Trim$(.TextMatrix(i, .ColIndex("DayAfterAccount")))
            rs("CurrentAccount").value = Trim$(.TextMatrix(i, .ColIndex("CurrentAccount")))
            rs("CustomerSales").value = IIf(Trim$(.TextMatrix(i, .ColIndex("CustomerSales"))) = "", Null, Trim$(.TextMatrix(i, .ColIndex("CustomerSales"))))
            rs("CustomerPur").value = IIf(Trim$(.TextMatrix(i, .ColIndex("CustomerPur"))) = "", Null, Trim$(.TextMatrix(i, .ColIndex("CustomerPur"))))
            rs("CustomerCash").value = IIf(Trim$(.TextMatrix(i, .ColIndex("CustomerCash"))) = "", Null, Trim$(.TextMatrix(i, .ColIndex("CustomerCash"))))
            rs("CustomerPayment").value = IIf(Trim$(.TextMatrix(i, .ColIndex("CustomerPayment"))) = "", Null, Trim$(.TextMatrix(i, .ColIndex("CustomerPayment"))))
            rs("Comments").value = IIf(Trim$(.TextMatrix(i, .ColIndex("Comments"))) = "", Null, Trim$(.TextMatrix(i, .ColIndex("Comments"))))
            rs.update
        Next i

    End With

    rs.Close
    Set rs = Nothing
   If SystemOptions.UserInterface = ArabicInterface Then
          StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\RptCustomersDayAccounts.rpt"
        Else
        StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\RptCustomersDayAccountsE.rpt"
        End If

    If Dir(StrFileName) = "" Then
        Msg = "„·› «· ﬁ—Ì— €Ì— „ÊÃÊœ..!!" & CHR(13)
        Msg = Msg & "»—Ã«¡ «· √ﬂœ „‰ ÊÃÊœ Â–« «·„·› ›Ï „”«— «·»—‰«„Ã" & CHR(13)
        Msg = Msg & StrFileName
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

    Set rs = New ADODB.Recordset
    StrSQL = "SELECT     dbo.TblCustemers.Fullcode, dbo.TempPrintCustomersDayAccount.*"
    StrSQL = StrSQL & "       FROM         dbo.TempPrintCustomersDayAccount INNER JOIN"
    StrSQL = StrSQL & "                  dbo.TblCustemers ON dbo.TempPrintCustomersDayAccount.CusID = dbo.TblCustemers.CusID"
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
'Screen.MousePointer = vbDefault
    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)
    xReport.Database.SetDataSource rs
    Set StrRS = New ADODB.Recordset
    StrRS.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    xReport.ParameterFields(1).AddCurrentValue IIf(IsNull(StrRS("Company_Arabic_Name").value), "", StrRS("Company_Arabic_Name").value)
    xReport.ParameterFields(2).AddCurrentValue IIf(IsNull(StrRS("Company_Comment").value), "", StrRS("Company_Comment").value)
    xReport.ParameterFields(3).AddCurrentValue user_name
    Screen.MousePointer = vbDefault
    Set XViewr = New ClsReportViewer
    'XViewr.FireReport xReport, WindowTarget

    xReport.ReportAuthor = App.title
     Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName
    Exit Sub
ErrTrap:
    Msg = "⁄›Ê« " & CHR(13) & "·«Ì„ﬂ‰ ÿ»«⁄… «· ﬁ—Ì—" & CHR(13)
    Msg = Msg & CHR(13) & Err.Number
    Msg = Msg & CHR(13) & Err.description
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Screen.MousePointer = vbDefault

End Sub
