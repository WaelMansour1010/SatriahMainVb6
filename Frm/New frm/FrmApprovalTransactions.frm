VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SuiteCtrls.ocx"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form FrmApprovalTransactions 
   BackColor       =   &H00C0FFC0&
   Caption         =   "«·„” ‰œ«   ﬁÌœ «·«⁄ „«œ"
   ClientHeight    =   8415
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11760
   HelpContextID   =   440
   Icon            =   "FrmApprovalTransactions.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   8415
   ScaleWidth      =   11760
   WindowState     =   2  'Maximized
   Begin C1SizerLibCtl.C1Elastic C1Elastic2 
      Height          =   8415
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   11760
      _cx             =   20743
      _cy             =   14843
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
      BackColor       =   -2147483635
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
      Begin C1SizerLibCtl.C1Tab C1Tab1 
         Height          =   8385
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   11730
         _cx             =   20690
         _cy             =   14790
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
         FrontTabColor   =   -2147483633
         BackTabColor    =   14871017
         TabOutlineColor =   -2147483632
         FrontTabForeColor=   -2147483630
         Caption         =   "”‰œ«  ﬁÌœ «·«⁄ „«œ|Õ«·… «·”‰œ« |›Ê« Ì— «·„»Ì⁄« |ﬁÌÊœ Õ—ﬂ«  €Ì— „⁄ „œ…"
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
         Begin C1SizerLibCtl.C1Elastic C1Elastic10 
            Height          =   8010
            Left            =   12975
            TabIndex        =   56
            TabStop         =   0   'False
            Top             =   45
            Width           =   11640
            _cx             =   20532
            _cy             =   14129
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
            Begin VSFlex8UCtl.VSFlexGrid GRIDGE 
               Height          =   8550
               Left            =   0
               TabIndex        =   57
               Tag             =   "1"
               Top             =   1560
               Width           =   11565
               _cx             =   20399
               _cy             =   15081
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
               Cols            =   16
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmApprovalTransactions.frx":038A
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
               WordWrap        =   -1  'True
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
            Begin MSComCtl2.DTPicker DTPicker3 
               Height          =   390
               Left            =   17370
               TabIndex        =   58
               Top             =   735
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   688
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   242876417
               CurrentDate     =   41640
            End
            Begin MSComCtl2.DTPicker DTPicker4 
               Height          =   390
               Left            =   14535
               TabIndex        =   59
               Top             =   720
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   688
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   242876417
               CurrentDate     =   41640
            End
            Begin VB.Label Label16 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "≈ŸÂ«— «·„” ‰œ« "
               Height          =   390
               Left            =   13215
               RightToLeft     =   -1  'True
               TabIndex        =   62
               Top             =   855
               Width           =   1200
            End
            Begin VB.Label Label11 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "«·Ï  «—ÌŒ"
               Height          =   270
               Left            =   16530
               RightToLeft     =   -1  'True
               TabIndex        =   61
               Top             =   735
               Width           =   720
            End
            Begin VB.Label Label10 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "„‰  «—ÌŒ"
               Height          =   315
               Left            =   19185
               RightToLeft     =   -1  'True
               TabIndex        =   60
               Top             =   735
               Width           =   720
            End
         End
         Begin C1SizerLibCtl.C1Elastic EleMain 
            Height          =   8010
            Left            =   45
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   45
            Width           =   11640
            _cx             =   20532
            _cy             =   14129
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
            Begin VB.Timer Timer1 
               Interval        =   500
               Left            =   0
               Top             =   0
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   1140
               Left            =   0
               TabIndex        =   3
               TabStop         =   0   'False
               Top             =   7305
               Width           =   20205
               _cx             =   35639
               _cy             =   2011
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
               Begin VB.CheckBox ChkShow 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "·«  ŸÂ— Â–Â «·‰«›–… ⁄‰œ  ‘€Ì· «·»—‰«„Ã"
                  ForeColor       =   &H000000FF&
                  Height          =   345
                  Left            =   6870
                  RightToLeft     =   -1  'True
                  TabIndex        =   6
                  Top             =   1770
                  Visible         =   0   'False
                  Width           =   3930
               End
               Begin VB.Frame Frame1 
                  Caption         =   "œ·«·«  «·«·Ê«‰"
                  Height          =   750
                  Left            =   1695
                  RightToLeft     =   -1  'True
                  TabIndex        =   4
                  Top             =   75
                  Width           =   3135
                  Begin VB.Shape Shape1 
                     BorderColor     =   &H000000FF&
                     FillColor       =   &H000000FF&
                     FillStyle       =   0  'Solid
                     Height          =   255
                     Left            =   1200
                     Top             =   240
                     Width           =   375
                  End
                  Begin VB.Label Label2 
                     Alignment       =   1  'Right Justify
                     Caption         =   "„ √Œ—"
                     Height          =   255
                     Left            =   120
                     RightToLeft     =   -1  'True
                     TabIndex        =   5
                     Top             =   240
                     Width           =   735
                  End
               End
               Begin ImpulseButton.ISButton CmdExit 
                  Height          =   255
                  Left            =   240
                  TabIndex        =   7
                  Top             =   330
                  Width           =   1335
                  _ExtentX        =   2355
                  _ExtentY        =   450
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
                  ButtonImage     =   "FrmApprovalTransactions.frx":0611
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
               Begin ImpulseButton.ISButton CmdPrint 
                  Height          =   255
                  Left            =   4770
                  TabIndex        =   8
                  Top             =   330
                  Width           =   1500
                  _ExtentX        =   2646
                  _ExtentY        =   450
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
                  ButtonImage     =   "FrmApprovalTransactions.frx":09AB
                  ColorButton     =   14871017
                  ColorHighlight  =   16777215
                  ColorHoverText  =   16711680
                  ColorShadow     =   -2147483637
                  ColorOutline    =   0
                  DrawFocusRectangle=   0   'False
                  ColorToggledHoverText=   16711680
                  ColorTextShadow =   -2147483637
               End
               Begin ALLButtonS.ALLButton cmdAdd 
                  Height          =   450
                  Left            =   6750
                  TabIndex        =   9
                  Tag             =   "Delete Row"
                  Top             =   315
                  Width           =   5370
                  _ExtentX        =   9472
                  _ExtentY        =   794
                  BTYPE           =   3
                  TX              =   " ÕœÌÀ «·»Ì«‰« "
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
                  BCOL            =   16744576
                  BCOLO           =   16744576
                  FCOL            =   0
                  FCOLO           =   0
                  MCOL            =   12632256
                  MPTR            =   1
                  MICON           =   "FrmApprovalTransactions.frx":0D45
                  UMCOL           =   -1  'True
                  SOFT            =   0   'False
                  PICPOS          =   0
                  NGREY           =   0   'False
                  FX              =   0
                  HAND            =   0   'False
                  CHECK           =   0   'False
                  VALUE           =   0   'False
               End
               Begin VB.Label Label1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Ì „  ÕœÌœ Â–Â «·»Ì«‰«  »‰«¡« ⁄·Ï «· «—ÌŒ «·Õ«·Ì ›Ì «·ÃÂ«“"
                  ForeColor       =   &H000000FF&
                  Height          =   345
                  Left            =   11520
                  RightToLeft     =   -1  'True
                  TabIndex        =   10
                  Top             =   420
                  Width           =   8145
               End
            End
            Begin VSFlex8UCtl.VSFlexGrid Fg 
               Height          =   6720
               Left            =   180
               TabIndex        =   11
               Top             =   585
               Width           =   11460
               _cx             =   20214
               _cy             =   11853
               Appearance      =   1
               BorderStyle     =   1
               Enabled         =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
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
               BackColorSel    =   49344
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
               SelectionMode   =   1
               GridLines       =   1
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   50
               Cols            =   24
               FixedRows       =   1
               FixedCols       =   0
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmApprovalTransactions.frx":0D61
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
               ExplorerBar     =   1
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
            Begin VB.Label LblCaption 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Caption         =   "«·„” ‰œ«   ﬁÌœ «·«⁄ „«œ"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   24
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404000&
               Height          =   690
               Left            =   -120
               RightToLeft     =   -1  'True
               TabIndex        =   12
               Top             =   -120
               Width           =   20145
            End
            Begin VB.Image Image1 
               Height          =   480
               Left            =   2055
               Picture         =   "FrmApprovalTransactions.frx":111C
               Top             =   165
               Visible         =   0   'False
               Width           =   480
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic3 
            Height          =   8010
            Left            =   12375
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   45
            Width           =   11640
            _cx             =   20532
            _cy             =   14129
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
            Begin VSFlex8UCtl.VSFlexGrid GRID2 
               Height          =   5550
               Left            =   60
               TabIndex        =   14
               Tag             =   "1"
               Top             =   1710
               Width           =   11280
               _cx             =   19897
               _cy             =   9790
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
               Cols            =   15
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmApprovalTransactions.frx":1DE6
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
               WordWrap        =   -1  'True
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
               Begin VB.Frame Frame3 
                  Height          =   3615
                  Left            =   0
                  RightToLeft     =   -1  'True
                  TabIndex        =   34
                  Top             =   2400
                  Visible         =   0   'False
                  Width           =   7695
                  Begin VB.CommandButton Command7 
                     BackColor       =   &H000000FF&
                     Caption         =   "X"
                     Height          =   255
                     Left            =   7320
                     Style           =   1  'Graphical
                     TabIndex        =   35
                     Top             =   0
                     Width           =   375
                  End
                  Begin VB.Shape Shape5 
                     BorderWidth     =   2
                     Height          =   3375
                     Left            =   120
                     Top             =   240
                     Width           =   7575
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00C0FFFF&
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
                     Height          =   3420
                     Index           =   25
                     Left            =   120
                     RightToLeft     =   -1  'True
                     TabIndex        =   36
                     Top             =   240
                     Width           =   7575
                  End
               End
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic4 
               Height          =   690
               Left            =   0
               TabIndex        =   16
               TabStop         =   0   'False
               Top             =   7320
               Width           =   11640
               _cx             =   20532
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
               Begin VB.Frame Frame2 
                  Caption         =   "œ·«·«  «·«·Ê«‰"
                  Height          =   825
                  Left            =   8115
                  RightToLeft     =   -1  'True
                  TabIndex        =   30
                  Top             =   195
                  Width           =   6660
                  Begin VB.Label Label9 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H0000FF00&
                     BackStyle       =   0  'Transparent
                     Caption         =   "„⁄ „œ"
                     Height          =   255
                     Left            =   0
                     RightToLeft     =   -1  'True
                     TabIndex        =   33
                     Top             =   120
                     Width           =   735
                  End
                  Begin VB.Shape Shape4 
                     BorderColor     =   &H0000FF00&
                     FillColor       =   &H0000FF00&
                     FillStyle       =   0  'Solid
                     Height          =   255
                     Left            =   840
                     Top             =   120
                     Width           =   375
                  End
                  Begin VB.Label Label8 
                     Alignment       =   1  'Right Justify
                     Caption         =   "„‰ Ÿ—"
                     Height          =   255
                     Left            =   1320
                     RightToLeft     =   -1  'True
                     TabIndex        =   32
                     Top             =   120
                     Width           =   735
                  End
                  Begin VB.Shape Shape3 
                     BorderColor     =   &H0000FFFF&
                     FillColor       =   &H0000FFFF&
                     FillStyle       =   0  'Solid
                     Height          =   255
                     Left            =   2160
                     Top             =   120
                     Width           =   375
                  End
                  Begin VB.Label Label7 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H0080FFFF&
                     BackStyle       =   0  'Transparent
                     Caption         =   "„—›Ê÷"
                     Height          =   255
                     Left            =   2880
                     RightToLeft     =   -1  'True
                     TabIndex        =   31
                     Top             =   120
                     Width           =   735
                  End
                  Begin VB.Shape Shape2 
                     BorderColor     =   &H000000FF&
                     FillColor       =   &H000000FF&
                     FillStyle       =   0  'Solid
                     Height          =   255
                     Left            =   3720
                     Top             =   120
                     Width           =   375
                  End
               End
               Begin ImpulseButton.ISButton ISButton1 
                  Height          =   945
                  Left            =   180
                  TabIndex        =   17
                  Top             =   405
                  Width           =   2340
                  _ExtentX        =   4128
                  _ExtentY        =   1667
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
                  ButtonImage     =   "FrmApprovalTransactions.frx":2038
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
               Begin ALLButtonS.ALLButton ALLButton1 
                  Height          =   945
                  Left            =   3390
                  TabIndex        =   18
                  Tag             =   "Delete Row"
                  Top             =   330
                  Width           =   4545
                  _ExtentX        =   8017
                  _ExtentY        =   1667
                  BTYPE           =   3
                  TX              =   " ÕœÌÀ «·»Ì«‰« "
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
                  BCOL            =   16744576
                  BCOLO           =   16744576
                  FCOL            =   0
                  FCOLO           =   0
                  MCOL            =   12632256
                  MPTR            =   1
                  MICON           =   "FrmApprovalTransactions.frx":23D2
                  UMCOL           =   -1  'True
                  SOFT            =   0   'False
                  PICPOS          =   0
                  NGREY           =   0   'False
                  FX              =   0
                  HAND            =   0   'False
                  CHECK           =   0   'False
                  VALUE           =   0   'False
               End
               Begin MSDataListLib.DataCombo DCboUserName 
                  Height          =   315
                  Left            =   8115
                  TabIndex        =   29
                  Top             =   1695
                  Visible         =   0   'False
                  Width           =   5595
                  _ExtentX        =   9869
                  _ExtentY        =   556
                  _Version        =   393216
                  Enabled         =   0   'False
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic5 
               Height          =   960
               Left            =   0
               TabIndex        =   19
               TabStop         =   0   'False
               Top             =   720
               Width           =   11640
               _cx             =   20532
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
               Begin MSComCtl2.DTPicker FrmDate 
                  Height          =   270
                  Left            =   7515
                  TabIndex        =   20
                  Top             =   120
                  Width           =   960
                  _ExtentX        =   1693
                  _ExtentY        =   476
                  _Version        =   393216
                  CheckBox        =   -1  'True
                  Format          =   202047489
                  CurrentDate     =   41640
               End
               Begin MSComCtl2.DTPicker ToDate 
                  Height          =   255
                  Left            =   7515
                  TabIndex        =   21
                  Top             =   525
                  Width           =   915
                  _ExtentX        =   1614
                  _ExtentY        =   450
                  _Version        =   393216
                  CheckBox        =   -1  'True
                  Format          =   202047489
                  CurrentDate     =   41640
               End
               Begin C1SizerLibCtl.C1Elastic C1Elastic6 
                  Height          =   855
                  Left            =   105
                  TabIndex        =   24
                  TabStop         =   0   'False
                  Top             =   120
                  Width           =   5835
                  _cx             =   10292
                  _cy             =   1508
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
                  Begin XtremeSuiteControls.RadioButton Rd 
                     Height          =   480
                     Index           =   0
                     Left            =   3555
                     TabIndex        =   25
                     Top             =   165
                     Width           =   2070
                     _Version        =   786432
                     _ExtentX        =   3651
                     _ExtentY        =   847
                     _StockProps     =   79
                     Caption         =   "«·„—›Ê÷Â"
                     BackColor       =   14871017
                     UseVisualStyle  =   -1  'True
                     TextAlignment   =   1
                     RightToLeft     =   -1  'True
                  End
                  Begin XtremeSuiteControls.RadioButton Rd 
                     Height          =   480
                     Index           =   1
                     Left            =   1965
                     TabIndex        =   26
                     Top             =   165
                     Width           =   1590
                     _Version        =   786432
                     _ExtentX        =   2805
                     _ExtentY        =   847
                     _StockProps     =   79
                     Caption         =   "«·„⁄ „œÂ"
                     BackColor       =   14871017
                     UseVisualStyle  =   -1  'True
                     TextAlignment   =   1
                     RightToLeft     =   -1  'True
                  End
                  Begin XtremeSuiteControls.RadioButton Rd 
                     Height          =   480
                     Index           =   2
                     Left            =   105
                     TabIndex        =   27
                     Top             =   165
                     Width           =   1245
                     _Version        =   786432
                     _ExtentX        =   2196
                     _ExtentY        =   847
                     _StockProps     =   79
                     Caption         =   "«·ﬂ·"
                     BackColor       =   14871017
                     UseVisualStyle  =   -1  'True
                     TextAlignment   =   1
                     RightToLeft     =   -1  'True
                  End
               End
               Begin VB.Label Label18 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·Ì"
                  Height          =   210
                  Left            =   8400
                  RightToLeft     =   -1  'True
                  TabIndex        =   64
                  Top             =   495
                  Width           =   570
               End
               Begin VB.Label Label17 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "„‰ "
                  Height          =   225
                  Left            =   8400
                  RightToLeft     =   -1  'True
                  TabIndex        =   63
                  Top             =   210
                  Width           =   570
               End
               Begin VB.Label Label6 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "≈ŸÂ«— «·„” ‰œ« "
                  Height          =   570
                  Left            =   5940
                  RightToLeft     =   -1  'True
                  TabIndex        =   28
                  Top             =   120
                  Width           =   2100
               End
               Begin VB.Label Label5 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·Ï  «—ÌŒ"
                  Height          =   390
                  Left            =   12180
                  RightToLeft     =   -1  'True
                  TabIndex        =   23
                  Top             =   120
                  Width           =   1380
               End
               Begin VB.Label Label4 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "„‰  «—ÌŒ"
                  Height          =   465
                  Left            =   18345
                  RightToLeft     =   -1  'True
                  TabIndex        =   22
                  Top             =   120
                  Width           =   1245
               End
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Caption         =   " «·„” ‰œ«  «·„⁄ „œÂ/«·„—›Ê÷Â"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   24
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404000&
               Height          =   900
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   15
               Top             =   0
               Width           =   20265
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic1 
            Height          =   8010
            Left            =   12675
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   45
            Width           =   11640
            _cx             =   20532
            _cy             =   14129
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
            Begin C1SizerLibCtl.C1Elastic C1Elastic7 
               Height          =   1140
               Left            =   0
               TabIndex        =   38
               TabStop         =   0   'False
               Top             =   7305
               Width           =   20205
               _cx             =   35639
               _cy             =   2011
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
               Begin VB.CheckBox Check1 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "·«  ŸÂ— Â–Â «·‰«›–… ⁄‰œ  ‘€Ì· «·»—‰«„Ã"
                  ForeColor       =   &H000000FF&
                  Height          =   615
                  Left            =   20430
                  RightToLeft     =   -1  'True
                  TabIndex        =   39
                  Top             =   3060
                  Visible         =   0   'False
                  Width           =   11865
               End
               Begin ImpulseButton.ISButton ISButton2 
                  Height          =   435
                  Left            =   720
                  TabIndex        =   40
                  Top             =   570
                  Width           =   3975
                  _ExtentX        =   7011
                  _ExtentY        =   767
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
                  ButtonImage     =   "FrmApprovalTransactions.frx":23EE
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
               Begin ImpulseButton.ISButton ISButton3 
                  Height          =   435
                  Left            =   9180
                  TabIndex        =   41
                  Top             =   570
                  Visible         =   0   'False
                  Width           =   4485
                  _ExtentX        =   7911
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
                  ButtonImage     =   "FrmApprovalTransactions.frx":2788
                  ColorButton     =   14871017
                  ColorHighlight  =   16777215
                  ColorHoverText  =   16711680
                  ColorShadow     =   -2147483637
                  ColorOutline    =   0
                  DrawFocusRectangle=   0   'False
                  ColorToggledHoverText=   16711680
                  ColorTextShadow =   -2147483637
               End
               Begin ALLButtonS.ALLButton ALLButton2 
                  Height          =   765
                  Left            =   20115
                  TabIndex        =   42
                  Tag             =   "Delete Row"
                  Top             =   555
                  Width           =   16155
                  _ExtentX        =   28496
                  _ExtentY        =   1349
                  BTYPE           =   3
                  TX              =   " ÕœÌÀ «·»Ì«‰« "
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
                  BCOL            =   16744576
                  BCOLO           =   16744576
                  FCOL            =   0
                  FCOLO           =   0
                  MCOL            =   12632256
                  MPTR            =   1
                  MICON           =   "FrmApprovalTransactions.frx":2B22
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
            Begin VSFlex8UCtl.VSFlexGrid VSFlexGrid1 
               Height          =   5880
               Left            =   0
               TabIndex        =   43
               Top             =   1425
               Width           =   11580
               _cx             =   20426
               _cy             =   10372
               Appearance      =   1
               BorderStyle     =   1
               Enabled         =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
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
               BackColorSel    =   49344
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
               SelectionMode   =   1
               GridLines       =   1
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   50
               Cols            =   22
               FixedRows       =   1
               FixedCols       =   0
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmApprovalTransactions.frx":2B3E
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
               ExplorerBar     =   1
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
            Begin C1SizerLibCtl.C1Elastic C1Elastic8 
               Height          =   870
               Left            =   0
               TabIndex        =   45
               TabStop         =   0   'False
               Top             =   600
               Width           =   11640
               _cx             =   20532
               _cy             =   1535
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
               Begin MSComCtl2.DTPicker DTPicker1 
                  Height          =   885
                  Left            =   45540
                  TabIndex        =   46
                  Top             =   195
                  Width           =   8760
                  _ExtentX        =   15452
                  _ExtentY        =   1561
                  _Version        =   393216
                  CheckBox        =   -1  'True
                  Format          =   202113025
                  CurrentDate     =   41640
               End
               Begin MSComCtl2.DTPicker DTPicker2 
                  Height          =   885
                  Left            =   30945
                  TabIndex        =   47
                  Top             =   180
                  Width           =   9285
                  _ExtentX        =   16378
                  _ExtentY        =   1561
                  _Version        =   393216
                  CheckBox        =   -1  'True
                  Format          =   202113025
                  CurrentDate     =   41640
               End
               Begin C1SizerLibCtl.C1Elastic C1Elastic9 
                  Height          =   1335
                  Left            =   210
                  TabIndex        =   48
                  TabStop         =   0   'False
                  Top             =   180
                  Width           =   22395
                  _cx             =   39502
                  _cy             =   2355
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
                  Begin XtremeSuiteControls.RadioButton Rd 
                     Height          =   570
                     Index           =   3
                     Left            =   5625
                     TabIndex        =   49
                     Top             =   195
                     Width           =   3030
                     _Version        =   786432
                     _ExtentX        =   5345
                     _ExtentY        =   1005
                     _StockProps     =   79
                     Caption         =   "«·„—›Ê÷Â"
                     BackColor       =   14871017
                     UseVisualStyle  =   -1  'True
                     TextAlignment   =   1
                     RightToLeft     =   -1  'True
                  End
                  Begin XtremeSuiteControls.RadioButton Rd 
                     Height          =   570
                     Index           =   4
                     Left            =   2820
                     TabIndex        =   50
                     Top             =   195
                     Width           =   2490
                     _Version        =   786432
                     _ExtentX        =   4392
                     _ExtentY        =   1005
                     _StockProps     =   79
                     Caption         =   "«·„⁄ „œÂ"
                     BackColor       =   14871017
                     UseVisualStyle  =   -1  'True
                     TextAlignment   =   1
                     RightToLeft     =   -1  'True
                  End
                  Begin XtremeSuiteControls.RadioButton Rd 
                     Height          =   570
                     Index           =   5
                     Left            =   105
                     TabIndex        =   51
                     Top             =   195
                     Width           =   1860
                     _Version        =   786432
                     _ExtentX        =   3281
                     _ExtentY        =   1005
                     _StockProps     =   79
                     Caption         =   "«·ﬂ·"
                     BackColor       =   14871017
                     UseVisualStyle  =   -1  'True
                     TextAlignment   =   1
                     RightToLeft     =   -1  'True
                  End
                  Begin XtremeSuiteControls.RadioButton Rd 
                     Height          =   570
                     Index           =   6
                     Left            =   9285
                     TabIndex        =   55
                     Top             =   195
                     Width           =   2790
                     _Version        =   786432
                     _ExtentX        =   4921
                     _ExtentY        =   1005
                     _StockProps     =   79
                     Caption         =   "«·ÃœÌœ…"
                     BackColor       =   14871017
                     UseVisualStyle  =   -1  'True
                     TextAlignment   =   1
                     RightToLeft     =   -1  'True
                  End
               End
               Begin VB.Label Label15 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "„‰  «—ÌŒ"
                  Height          =   705
                  Left            =   54720
                  RightToLeft     =   -1  'True
                  TabIndex        =   54
                  Top             =   195
                  Width           =   3765
               End
               Begin VB.Label Label14 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "«·Ï  «—ÌŒ"
                  Height          =   615
                  Left            =   41160
                  RightToLeft     =   -1  'True
                  TabIndex        =   53
                  Top             =   195
                  Width           =   3870
               End
               Begin VB.Label Label13 
                  Alignment       =   2  'Center
                  BackStyle       =   0  'Transparent
                  Caption         =   "≈ŸÂ«— «·„” ‰œ« "
                  Height          =   870
                  Left            =   24075
                  RightToLeft     =   -1  'True
                  TabIndex        =   52
                  Top             =   465
                  Width           =   6150
               End
            End
            Begin VB.Image Image2 
               Height          =   480
               Left            =   2055
               Picture         =   "FrmApprovalTransactions.frx":2EBE
               Top             =   165
               Visible         =   0   'False
               Width           =   480
            End
            Begin VB.Label Label12 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               Caption         =   "«·„” ‰œ«   ﬁÌœ «·«⁄ „«œ"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   24
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404000&
               Height          =   690
               Left            =   -120
               RightToLeft     =   -1  'True
               TabIndex        =   44
               Top             =   -120
               Width           =   20145
            End
         End
      End
   End
End
Attribute VB_Name = "FrmApprovalTransactions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Askinterval As String
Dim Askcount As Integer
Public ScreenName As String

Function GetHobStatus() As Integer
Dim sql As String
GetHobStatus = 0
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
sql = " SELECT     id, Vacation"
sql = sql & " From dbo.jopstatus"
sql = sql & " Where (Vacation = 1)"
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
GetHobStatus = IIf(IsNull(Rs3("id").value), 0, Rs3("id").value)
Else
GetHobStatus = 0
End If
End Function

Private Sub ALLButton1_Click()
fillapprovData
End Sub

Private Sub ALLButton2_Click()
FillInvoice
End Sub

Private Sub cmdAdd_Click()
loadFlexGrid
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CmdPrint_Click()
    On Error GoTo ErrTrap
    Dim Reports As ClsRepoerts
    Dim StrSQL As String

    Askinterval = GetSetting(StrAppRegPath, "Setting", "INTERVAL_InstallmentMustPayed", True)
    Askcount = GetSetting(StrAppRegPath, "Setting", "count_InstallmentMustPayed", True)
    
    'StrSQL = "select * From QestNotReceipted where  DueDate<='" & SQLDate(DateAdd(Askinterval, Askcount, Date)) & "'"
    ' StrSQL = StrSQL + " order by CusName,Transaction_ID,QeqtNum"

    StrSQL = "SELECT     dbo.Transactions.Transaction_ID, dbo.Transactions.Transaction_Date, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.Transactions.Posted, "
    StrSQL = StrSQL + "  dbo.TblUsers.UserName , dbo.Transactions.order_no,  dbo.Transactions.PostedDate"
    StrSQL = StrSQL + " FROM         dbo.Transactions INNER JOIN"
    StrSQL = StrSQL + " dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
    StrSQL = StrSQL + " dbo.TblUsers ON dbo.Transactions.Posted = dbo.TblUsers.UserID"
    StrSQL = StrSQL + " WHERE     (NOT (dbo.Transactions.Posted IS NULL)) AND (dbo.Transactions.order_no NOT IN"
    StrSQL = StrSQL + " (SELECT     order_no"
    StrSQL = StrSQL + " From Transactions"
    StrSQL = StrSQL + " WHERE     Transaction_Type = 21 AND NOT (order_no IS NULL))) AND (dbo.Transactions.Transaction_Type = 17)"
    StrSQL = StrSQL + " ORDER BY dbo.Transactions.PostedDate"
   
    Set Reports = New ClsRepoerts
    Reports.AccreditOrders StrSQL, , LblCaption.Caption
    Exit Sub
ErrTrap:
End Sub

 

Private Sub Command7_Click()
Frame3.Visible = False
End Sub

   Private Sub FG_CellButtonClick(ByVal row As Long, _
                               ByVal Col As Long)
Dim currrentScreenName As String
Dim newapprovalno As Double
            Dim sql As String
                Dim X As Integer

    With Me.FG
 currrentScreenName = (.TextMatrix(row, .ColIndex("ScreenName")))
        Select Case .ColKey(Col)

            Case "Show"

           

        If currrentScreenName = "FrmPO" Then
            Unload FrmPO
            FrmPO.show
            FrmPO.Retrive val(.TextMatrix(row, .ColIndex("Transaction_ID")))
        ElseIf currrentScreenName = "FrmProductionOrder" Then
            Unload FrmProductionOrder
            FrmProductionOrder.show
            FrmProductionOrder.Retrive val(.TextMatrix(row, .ColIndex("Transaction_ID"))), True
        ElseIf currrentScreenName = "FrmTransacRegistr" Then
            Unload FrmTransacRegistr
            FrmTransacRegistr.show
            FrmTransacRegistr.FindRec val(.TextMatrix(row, .ColIndex("Transaction_ID")))
            
            
        ElseIf currrentScreenName = "FrmVocationEntitlements" Then
                       FrmVocationEntitlements.show
                        FrmVocationEntitlements.Retrive val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                        
                        ElseIf currrentScreenName = "FrmPO1" Then
Unload FrmPO1
                         FrmPO1.show
                         FrmPO1.Retrive val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                        
                      ElseIf currrentScreenName = "FrmPO2" Then
                 Unload FrmPO2
                 FrmPO2.show
                               FrmPO2.Retrive val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                   ElseIf currrentScreenName = "FrmPO3" Then
                       Unload FrmPO3
                       FrmPO3.show
                         FrmPO3.Retrive val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                        
                      ElseIf currrentScreenName = "FrmPO4" Then
                               Unload FrmPO4

                              FrmPO4.show
                               FrmPO4.Retrive val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                        
                             ElseIf currrentScreenName = "FrmPO5" Then
                               Unload FrmPO5
                               FrmPO5.show
                               FrmPO5.Retrive val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                        
                        
                                ElseIf currrentScreenName = "FrmPO6" Then
                              Unload FrmPO6
                              FrmPO6.show
                              FrmPO6.Retrive val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                        
                                      ElseIf currrentScreenName = "FrmPO7" Then
                                      Unload FrmPO7
                         FrmPO7.show
                         FrmPO7.Retrive val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                                               ElseIf currrentScreenName = "FrmPO8" Then
                                               Unload FrmPO8
                         FrmPO8.show
                         FrmPO8.Retrive val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                                                         ElseIf currrentScreenName = "FrmPO10" Then
                                                         Unload FrmPO10
                         FrmPO10.show
                         FrmPO10.Retrive val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                        
                            
                                                              ElseIf currrentScreenName = "FrmEmpSalary3" Then
                                                              Unload FrmEmpSalary3
                         FrmEmpSalary3.show
                         FrmEmpSalary3.Retrive (val(Me.FG.TextMatrix(Me.FG.row, Me.FG.ColIndex("NoteSerial"))))
                        
                            
                        
                     ElseIf currrentScreenName = "FrmEmpsAdvanceRequest" Then
                       Unload FrmEmpsAdvanceRequest
                        FrmEmpsAdvanceRequest.show
                              FrmEmpsAdvanceRequest.Retrive val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                        
              ElseIf currrentScreenName = "Frmpassover" Then
                   '    FrmPassover.show
                   '     FrmPassover.Retrive val(.TextMatrix(Row, .ColIndex("Transaction_ID")))
                ElseIf currrentScreenName = "FrmBusinessJob" Then
                   Unload FrmBusinessJob
                   FrmBusinessJob.show
                   FrmBusinessJob.Retrive val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                 ElseIf currrentScreenName = "FrmEmbarkation" Then
                 Unload FrmEmbarkation
                        FrmEmbarkation.show
                         FrmEmbarkation.Retrive val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                        
                            ElseIf currrentScreenName = "formvocatinl" Then
                            Unload formvocatinl
                       formvocatinl.show
                        formvocatinl.Retrive val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                        
                              
                                 
                            ElseIf currrentScreenName = "FormEmpMoveDepartment" Then
                                     Unload FormEmpMoveDepartment
                         FormEmpMoveDepartment.show
                           FormEmpMoveDepartment.Retrive val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                        
                        
                                   ElseIf currrentScreenName = "FrmEmpsAdvanceRequest" Then
                                     Unload FrmEmpsAdvanceRequest
                              FrmEmpsAdvanceRequest.show
                               FrmEmpsAdvanceRequest.Retrive val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                         
                         
                                        ElseIf currrentScreenName = "FrmTypeExchange" Then
                                        Unload FrmTypeExchange
                       FrmTypeExchange.show
                         FrmTypeExchange.Retrive val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                         
                                            ElseIf currrentScreenName = "FrmShipmentOrder" Then
                                              Unload FrmShipmentOrder
                              FrmShipmentOrder.show
                              FrmShipmentOrder.Retrive val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                                            
                             ElseIf currrentScreenName = "FrmCreditFacicity" Then
                                 Unload FrmCreditFacicity
                                 FrmCreditFacicity.show
                                 FrmCreditFacicity.Retrive val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                              ElseIf currrentScreenName = "FrmAccEditJournal" Then
                              Unload FrmAccEditJournal
                                   FrmAccEditJournal.show
                       
                        FrmAccEditJournal.Retrive (.TextMatrix(row, .ColIndex("NoteSerial")))
                              ElseIf currrentScreenName = "FrmAccEditJournal4" Then
                              Unload FrmAccEditJournal4
                                   FrmAccEditJournal4.show
                        FrmAccEditJournal4.Retrive val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                                ElseIf currrentScreenName = "FrmAccEditJournal1" Then
                                Unload FrmAccEditJournal1
                                   FrmAccEditJournal1.show
                        FrmAccEditJournal1.Retrive val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                        
                           ElseIf currrentScreenName = "FrmExpenses5" Then
                           Unload FrmExpenses5
                                   FrmExpenses5.show
                        FrmExpenses5.Retrive val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                           ElseIf currrentScreenName = "FrmAccEditJournal3" Then
                           Unload FrmAccEditJournal3
                                   FrmAccEditJournal3.show
                        FrmAccEditJournal3.Retrive (.TextMatrix(row, .ColIndex("NoteSerial")))
                           ElseIf currrentScreenName = "FrmCashing" Then
                           Unload FrmCashing
                                   FrmCashing.show
                        FrmCashing.Retrive val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                           ElseIf currrentScreenName = "FrmDestruction" Then
                           Unload FrmDestruction
                                   FrmDestruction.show
                        FrmDestruction.Retrive val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                           ElseIf currrentScreenName = "FrmDestructionRet" Then
                           Unload FrmDestructionRet
                                   FrmDestructionRet.show
                        FrmDestructionRet.Retrive val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                          ElseIf currrentScreenName = "FrmExpenses30" Then
                          Unload FrmExpenses30
                                   FrmExpenses30.show
                        FrmExpenses30.Retrive val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                        
                           ElseIf currrentScreenName = "FrmExpenses301" Then
                          Unload FrmExpenses301
                                   FrmExpenses301.show
                        FrmExpenses301.Retrive val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                        
                            ElseIf currrentScreenName = "FrmExpenses3" Then
                            Unload FrmExpenses3
                                   FrmExpenses3.show
                        FrmExpenses3.Retrive val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                        
                   ElseIf currrentScreenName = "FrmExpenses4" Then
                   Unload FrmExpenses4
                                   FrmExpenses4.show
                        FrmExpenses4.Retrive val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                   ElseIf currrentScreenName = "FrmBoxDrawing" Then
                   Unload FrmBoxDrawing
                                   FrmBoxDrawing.show
                        FrmBoxDrawing.Retrive val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                   ElseIf currrentScreenName = "FrmPayments1" Then
                   Unload FrmPayments1
                                   FrmPayments1.show
                        FrmPayments1.Retrive val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                             ElseIf currrentScreenName = "FrmBankPledge1" Then
                             Unload FrmBankPledge1
                                   FrmBankPledge1.show
                        FrmBankPledge1.Retrive val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                              ElseIf currrentScreenName = "FrmBankPledge2" Then
                                Unload FrmBankPledge2
                               FrmBankPledge2.show
                          FrmBankPledge2.Retrive val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                              ElseIf currrentScreenName = "FrmBankPledge3" Then
                            Unload FrmBankPledge3
                            FrmBankPledge3.show
                            FrmBankPledge3.Retrive val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                         
                         
                         
                             ElseIf currrentScreenName = "FrmBankPledge4" Then
                               Unload FrmBankPledge4
                           FrmBankPledge4.show
                         FrmBankPledge4.Retrive val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                         
                         
                       ElseIf currrentScreenName = "FrmAdvancedHousingpayments" Then
                          Unload FrmAdvancedHousingpayments
                          FrmAdvancedHousingpayments.show
                          FrmAdvancedHousingpayments.Retrive val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                          ElseIf currrentScreenName = "FrmMovingEmp" Then
                          Unload FrmMovingEmp
                                   FrmMovingEmp.show
                        FrmMovingEmp.Retrive val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                     ElseIf currrentScreenName = "FrmQUesEmp" Then
                           Unload FrmQUesEmp
                          FrmQUesEmp.show
                          FrmQUesEmp.Retrive val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                     ElseIf currrentScreenName = "End_oF_service" Then
                          Unload End_oF_service
                          End_oF_service.show
                          End_oF_service.Retrive val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                          
                    ElseIf currrentScreenName = "FrmMoving" Then
                          Unload FrmMoving
                          FrmMoving.show
                          FrmMoving.Retrive val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                   ElseIf currrentScreenName = "projectsbill" Then
                                projectsbill.Retrive val(.TextMatrix(.row, .ColIndex("Transaction_ID")))
    
   
                   End If
                   

Case "CancelApprove"

      If SystemOptions.UserInterface = ArabicInterface Then
                        X = MsgBox("—›÷  «·«⁄ „«œ", vbExclamation + vbYesNoCancel)
                Else
                        X = MsgBox(" Cancel Approval", vbExclamation + vbYesNoCancel)
                End If
                If X = vbYes Then
                sql = "update ApprovalData set  Currcursor=null, Remarks= '   „ —›÷ «·«⁄ „«œ  »”»»  " & (.TextMatrix(row, .ColIndex("Remarks"))) & "',CancelApprove=getdate()  where id=" & val(.TextMatrix(row, .ColIndex("id")))
                Cn.Execute sql
                
                
                If SystemOptions.cancellAllApprove = True Then
                
                  sql = "update ApprovalData set  Currcursor=null, Remarks=' „ —›÷ «·«⁄ „«œ  »”»»  " & (.TextMatrix(row, .ColIndex("Remarks"))) & "'  where id=" & val(.TextMatrix(row, .ColIndex("id")))
                Cn.Execute sql
                
                
                
                          
                
                Else
                 newapprovalno = GetCurrentApprovalForTransactions(val(.TextMatrix(row, .ColIndex("Transaction_ID"))), currrentScreenName)
                If newapprovalno > 0 Then
              sql = "update ApprovalData set   SendTime=getdate() , Currcursor=1 , FromUser='" & user_name & "'    where id=" & newapprovalno
                Cn.Execute sql
                End If
                
                
                
                End If
                
    End If

            Case "Approve"
            
            
      
                If SystemOptions.UserInterface = ArabicInterface Then
                        X = MsgBox(" √ﬂÌœ «·«⁄ „«œ", vbExclamation + vbYesNoCancel)
                Else
                        X = MsgBox(" Confirm Approval", vbExclamation + vbYesNoCancel)
                End If
                If X = vbYes Then
                 sql = "update ApprovalData set  Currcursor=null, Remarks='" & (.TextMatrix(row, .ColIndex("Remarks"))) & "',ApprovDate=getdate()  where id=" & val(.TextMatrix(row, .ColIndex("id")))
                Cn.Execute sql
                newapprovalno = GetCurrentApprovalForTransactions(val(.TextMatrix(row, .ColIndex("Transaction_ID"))), currrentScreenName)
                
                
                If newapprovalno > 0 Then
              sql = "update ApprovalData set   SendTime=getdate() , Currcursor=1 , FromUser='" & user_name & "'    where id=" & newapprovalno
                Cn.Execute sql
                End If
                
              If CheckLastApprovLevel(currrentScreenName, val(.TextMatrix(row, .ColIndex("Transaction_ID")))) = 0 Then
                          If currrentScreenName = "FrmEmpsAdvanceRequest" Then
                                    
                                      sql = "update TblEmpAdvanceRequest set   Approved=1      where AdvanceID=" & val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                                      Cn.Execute sql
                                      
                           ElseIf currrentScreenName = "FrmVocationEntitlements" Then
                                    
                                     sql = "update dbo.TblVocationEntitlements set   Approved=1      where ID=" & val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                                      Cn.Execute sql
                            ElseIf currrentScreenName = "Frmpassover" Then
                                    
                                      sql = "update TblEmpPassOver set   Approved=1      where AdvanceID=" & val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                                      Cn.Execute sql
                                   
                                  ElseIf currrentScreenName = "FrmMovingEmp" Then
                                    
                                      sql = "update TblEmpPassOver set   Approved=1      where AdvanceID=" & val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                                      Cn.Execute sql
                                      DoEvents
                                      sql = "Update TblEmployee Set  jopstatusid=" & Me.GetHobStatus() & ",workstate=0 Where Emp_ID=" & GetEmIDUnpaidVacation(val(.TextMatrix(row, .ColIndex("Transaction_ID")))) & ""
                                      Cn.Execute sql, , adExecuteNoRecords
                            ElseIf currrentScreenName = "FrmBusinessJob" Then
                                    
                                      sql = "update TblEmpJobOrder set   Approved=1      where AdvanceID=" & val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                                      Cn.Execute sql
                           
                           
                     ElseIf currrentScreenName = "FrmEmbarkation" Then
                                    
                                      sql = "update TblEmbarkation set   Approved=1      where id=" & val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                                      Cn.Execute sql
                         ElseIf currrentScreenName = "End_oF_service" Then
                                    
                                      sql = "update End_of_service set   Approved=1      where id=" & val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                                       Cn.Execute sql
                                      
                      ElseIf currrentScreenName = "FrmTransacRegistr" Then
                                    
                                      sql = "update TblTransacRegistr set   Approved=1      where id=" & val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                                      Cn.Execute sql
                            '          updateemployeeEmbarkation (val(.TextMatrix(Row, .ColIndex("Transaction_ID"))))
                           
           ElseIf currrentScreenName = "formvocatinl" Then
                                      sql = "update TblVocation set   Approved=1      where id=" & val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                                      Cn.Execute sql
                ElseIf currrentScreenName = "FrmAccEditJournal1" Then
                                      sql = "update DOUBLE_ENTREY_VOUCHERS1  set   Posted=null      where Notes_ID=" & val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                                      Cn.Execute sql
               ElseIf currrentScreenName = "FrmExpenses3" Then
                                      sql = "update DOUBLE_ENTREY_VOUCHERS  set   Posted=null      where notes_all=" & val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                                      Cn.Execute sql
               ElseIf currrentScreenName = "FrmExpenses4" Then
                                      sql = "update DOUBLE_ENTREY_VOUCHERS  set   Posted=null      where notes_all=" & val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                                      Cn.Execute sql
               ElseIf currrentScreenName = "FrmExpenses30" Then
                                      sql = "update DOUBLE_ENTREY_VOUCHERS  set   Posted=null      where notes_all=" & val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                                      Cn.Execute sql
                                      
                 ElseIf currrentScreenName = "FrmExpenses301" Then
                                      sql = "update DOUBLE_ENTREY_VOUCHERS  set   Posted=null      where notes_all=" & val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                                      Cn.Execute sql
                                      
             ElseIf currrentScreenName = "FrmAccEditJournal4" Then
                                      sql = "update DOUBLE_ENTREY_VOUCHERS  set   Posted=null      where Notes_ID=" & val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                                      Cn.Execute sql
           ElseIf currrentScreenName = "FrmAccEditJournal" Or currrentScreenName = "FrmAccEditJournal3" Or currrentScreenName = "FrmCashing" Then
                                      sql = "update DOUBLE_ENTREY_VOUCHERS  set   Posted=null      where Notes_ID=" & val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                                      Cn.Execute sql
           ElseIf currrentScreenName = "FrmExpenses5" Then
                                      sql = "update DOUBLE_ENTREY_VOUCHERS  set   Posted=null      where notes_all=" & val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                                      Cn.Execute sql
           ElseIf currrentScreenName = "FrmDestruction" Then
                                      sql = "update DOUBLE_ENTREY_VOUCHERS  set   Posted=null      where Transaction_ID=" & val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                                      Cn.Execute sql
                                      sql = "update Transactions set  Approved=1,  Transaction_Type=18      where Transaction_ID=" & val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                                      Cn.Execute sql
            ElseIf currrentScreenName = "FrmDestructionRet" Then
                                      sql = "update DOUBLE_ENTREY_VOUCHERS  set   Posted=null      where Transaction_ID=" & val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                                      Cn.Execute sql
                                      sql = "update Transactions set  Approved=1,   Transaction_Type=66      where Transaction_ID=" & val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                                      Cn.Execute sql
            ElseIf currrentScreenName = "FrmBoxDrawing" Then
                                      sql = "update DOUBLE_ENTREY_VOUCHERS  set   Posted=null      where Notes_ID=" & val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                                      Cn.Execute sql
                                         sql = "update DOUBLE_ENTREY_VOUCHERS  set   Posted=null      where Notes_ID=" & val(.TextMatrix(row, .ColIndex("Transaction_ID"))) + 1
                                      Cn.Execute sql
            ElseIf currrentScreenName = "FrmPayments1" Then
                                      sql = "update DOUBLE_ENTREY_VOUCHERS  set   Posted=null      where Notes_ID=" & val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                                      Cn.Execute sql
                                '      updateemployeeEmbarkation1 (val(.TextMatrix(Row, .ColIndex("Transaction_ID"))))
            ElseIf currrentScreenName = "FrmBankPledge1" Then
                                      sql = " update TblBankPledge set   Approved=1      where id=" & val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                                      Cn.Execute sql
            ElseIf currrentScreenName = "FrmBankPledge2" Then
                                      sql = " update TblBankPledge2 set   Approved=1      where id=" & val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                                      Cn.Execute sql
            ElseIf currrentScreenName = "FrmBankPledge3" Then
                                      sql = " update TblBankPledge3 set   Approved=1      where id=" & val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                                      Cn.Execute sql
                                      
                                                 
                       ElseIf currrentScreenName = "FormEmpMoveDepartment" Then
                                    
                                      sql = "update TblMoveEmp1 set   Approved=1      where id=" & val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                                      Cn.Execute sql
                              
                                      
                                '      updateemployeeEmbarkation2 (val(.TextMatrix(Row, .ColIndex("Transaction_ID"))))
                                                                           
                                                                           
                                   ElseIf currrentScreenName = "FrmEmpsAdvanceRequest" Then
                                    
                                      sql = "update TblEmpAdvanceRequest set   Approved=1      where id=" & val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                                      Cn.Execute sql
                                      
                                     ElseIf currrentScreenName = "FrmTypeExchange" Then
                                    
                                      sql = "update TblExchange set   Approved=1      where id=" & val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                                      Cn.Execute sql
                             
                                                            
                                                ElseIf currrentScreenName = "FrmEmpSalary3" Then
                                    
                                      sql = "update opr_Employee set   Approved=1      where id=" & val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                                      Cn.Execute sql
                                                            
                                    ElseIf currrentScreenName = "FrmMoving" Then
                                      sql = "update DOUBLE_ENTREY_VOUCHERS  set   Posted=null      where Transaction_ID=" & val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                                      Cn.Execute sql
                                     '  sql = "update DOUBLE_ENTREY_VOUCHERS  set   Posted=null      where Transaction_ID=" & val(.TextMatrix(Row, .ColIndex("Transaction_ID"))) + 1
                                     ' Cn.Execute sql
                                      sql = "update Transactions set  Approved=1,  Transaction_Type=10      where Transaction_ID=" & val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                                      Cn.Execute sql
                                       sql = "update Transactions set  Approved=1,  Transaction_Type=11      where Transaction_ID=" & val(.TextMatrix(row, .ColIndex("Transaction_ID"))) + 1
                                      Cn.Execute sql
                                      ElseIf currrentScreenName = "projectsbill" Then
                                      sql = "update project_billl set   Approved=1      where id=" & val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                                      Cn.Execute sql
                                       sql = "update DOUBLE_ENTREY_VOUCHERS  set   Posted=null      where Notes_ID=" & val(.TextMatrix(row, .ColIndex("noteid")))
                                      Cn.Execute sql
                                                            
                            Else
                                        sql = "update Transactions set   Approved=1      where Transaction_ID=" & val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                                          Cn.Execute sql
            
                            End If
                
              End If
              
              
              
                loadFlexGrid .row
                End If
                
        End Select

    End With
loadFlexGrid Me.FG.row
End Sub
Function fillapprovData()
Dim Num As Integer
 Dim RsDetails As New ADODB.Recordset
 Dim StrSQL As String
  Dim Price As Double
 Dim ToPerson As String
 StrSQL = " SELECT                 dbo.ApprovalData.Currcursor, dbo.ApprovalData.ScreenName, dbo.ApprovalData.levelo, dbo.ApprovalData.EmpID, dbo.ApprovalData.levelorder, "
 StrSQL = StrSQL & "               dbo.ApprovalData.currorder, dbo.ApprovalData.Transaction_ID, dbo.ApprovalData.NoteID, dbo.ApprovalData.ApprovDate, dbo.ApprovalData.Remarks,"
 StrSQL = StrSQL & "               dbo.TbLLevels.Name, dbo.TbLLevels.Namee, dbo.TblUsers.UserID, dbo.TblUsers.UserName, dbo.Screens.ScreenCaption, dbo.Screens.ScreenTitleEng,"
 StrSQL = StrSQL & "               dbo.ApprovalData.NoteSerial, dbo.ApprovalData.Transaction_Date, dbo.ApprovalData.FromUser, dbo.ApprovalData.SendTime, dbo.ApprovalData.ExpectedtimeTime,"
 StrSQL = StrSQL & "               dbo.ApprovalData.CancelApprove"
 StrSQL = StrSQL & " FROM          dbo.ApprovalData LEFT OUTER JOIN"
 StrSQL = StrSQL & "               dbo.TbLLevels ON dbo.ApprovalData.levelo = dbo.TbLLevels.LevelID INNER JOIN"
 StrSQL = StrSQL & "               dbo.TblUsers ON dbo.ApprovalData.EmpID = dbo.TblUsers.UserID INNER JOIN"
 StrSQL = StrSQL & "               dbo.Screens ON dbo.ApprovalData.ScreenName = dbo.Screens.ScreenName"
' StrSQL = StrSQL & " WHERE     ((NOT (dbo.ApprovalData.CancelApprove IS NULL)) OR"
 
' StrSQL = StrSQL & "                     (NOT (dbo.ApprovalData.ApprovDate IS NULL))) and dbo.ApprovalData.FromUser='" & DCboUserName.Text & "'"




' If Not IsNull(Me.FrmDate.value) Then
' StrSQL = StrSQL & " and dbo.ApprovalData.approvdate >=" & SQLDate(Me.FrmDate.value, True) & ""
' End If
'  If Not IsNull(Me.ToDate.value) Then
' StrSQL = StrSQL & " and dbo.ApprovalData.approvdate <=" & SQLDate(Me.ToDate.value, True) & ""
' End If
 '  DATEADD(dd, 0, DATEDIFF(dd, 0, ApprovDate ))
 'DATEADD(dd, 0, DATEDIFF(dd, 0,  approvdate)
    StrSQL = StrSQL & " WHERE    ApprovalData.Transaction_ID in  ("
StrSQL = StrSQL & "  SELECT     Transaction_ID"
StrSQL = StrSQL & "  From dbo.ApprovalData"
StrSQL = StrSQL & "  WHERE   1 = 1 "

 If (SystemOptions.usertype <> UserAdminAll Or SystemOptions.usertype <> UserAdmin) And Not Rd(2).value Then
    StrSQL = StrSQL & " and (  (FromUser = '" & DCboUserName.text & "' and not (FromUser is null))"
    
    StrSQL = StrSQL & " or  dbo.ApprovalData.EmpID = " & user_id & ") "
End If
If Not Rd(2).value Then
    StrSQL = StrSQL & " and (  (FromUser = '" & DCboUserName.text & "' and not (FromUser is null))"
    
    StrSQL = StrSQL & " or  dbo.ApprovalData.EmpID = " & user_id & ") "

End If
 If Not IsNull(Me.FrmDate.value) Then
 StrSQL = StrSQL & " and DATEADD(dd, 0, DATEDIFF(dd, 0,  Transaction_Date)) >=" & SQLDate(Me.FrmDate.value, True) & ""
 End If
  If Not IsNull(Me.ToDate.value) Then
 StrSQL = StrSQL & " and DATEADD(dd, 0, DATEDIFF(dd, 0,  Transaction_Date) )<=" & SQLDate(Me.ToDate.value, True) & ""
 End If
 StrSQL = StrSQL & " )"

 If Not IsNull(Me.FrmDate.value) Then
 StrSQL = StrSQL & " and DATEADD(dd, 0, DATEDIFF(dd, 0,  Transaction_Date)) >=" & SQLDate(Me.FrmDate.value, True) & ""
 End If
  If Not IsNull(Me.ToDate.value) Then
 StrSQL = StrSQL & " and DATEADD(dd, 0, DATEDIFF(dd, 0,  Transaction_Date) )<=" & SQLDate(Me.ToDate.value, True) & ""
 End If
 
 If Rd(0).value = True Then
 StrSQL = StrSQL & " and (NOT (dbo.ApprovalData.CancelApprove IS NULL))"
 End If
  If Rd(1).value = True Then
 StrSQL = StrSQL & " and (NOT (dbo.ApprovalData.ApprovDate IS NULL))"
 End If
 StrSQL = StrSQL & " ORDER BY dbo.ApprovalData.levelorder"

    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

 If Not (RsDetails.EOF Or RsDetails.BOF) Then
        GRID2.rows = RsDetails.RecordCount + 1
 

        For Num = 1 To RsDetails.RecordCount
        
       GRID2.TextMatrix(Num, GRID2.ColIndex("Currcursor")) = IIf(IsNull(RsDetails("Currcursor")), "", RsDetails("Currcursor"))
       GRID2.TextMatrix(Num, GRID2.ColIndex("NoteSerial")) = IIf(IsNull(RsDetails("NoteSerial")), "", RsDetails("NoteSerial"))
       GRID2.TextMatrix(Num, GRID2.ColIndex("Transaction_ID")) = IIf(IsNull(RsDetails("Transaction_ID")), 0, RsDetails("Transaction_ID"))
      GRID2.TextMatrix(Num, GRID2.ColIndex("ScreenCaption")) = IIf(IsNull(RsDetails("ScreenCaption")), "", RsDetails("ScreenCaption"))
      GRID2.TextMatrix(Num, GRID2.ColIndex("ScreenName")) = IIf(IsNull(RsDetails("ScreenName")), "", RsDetails("ScreenName"))
      
            GRID2.TextMatrix(Num, GRID2.ColIndex("Approved")) = IIf(IsNull(RsDetails("ApprovDate")), "", flexChecked)
            
           If SystemOptions.UserInterface = ArabicInterface Then
            GRID2.TextMatrix(Num, GRID2.ColIndex("levelName")) = IIf(IsNull(RsDetails("Name")), "", Trim(RsDetails("Name").value))
          Else
             GRID2.TextMatrix(Num, GRID2.ColIndex("levelName")) = IIf(IsNull(RsDetails("Namee")), "", Trim(RsDetails("Namee").value))
          End If
          
          If GRID2.TextMatrix(Num, GRID2.ColIndex("levelName")) = "" Then
           GRID2.TextMatrix(Num, GRID2.ColIndex("levelName")) = "«·„œÌ— «·„»«‘—"
          End If
          
            If SystemOptions.UserInterface = ArabicInterface Then
            GRID2.TextMatrix(Num, GRID2.ColIndex("EmpName")) = IIf(IsNull(RsDetails("UserName")), "", (RsDetails("UserName").value))
            Else
            GRID2.TextMatrix(Num, GRID2.ColIndex("EmpName")) = IIf(IsNull(RsDetails("UserName")), "", (RsDetails("UserName").value))
            End If
            If SystemOptions.UserInterface = ArabicInterface Then
             GRID2.TextMatrix(Num, GRID2.ColIndex("show")) = "«÷€ÿ ··⁄—÷ "
             Else
             GRID2.TextMatrix(Num, GRID2.ColIndex("show")) = "Show"
             End If
             
            GRID2.TextMatrix(Num, GRID2.ColIndex("ApprovDate")) = IIf(IsNull(RsDetails("ApprovDate")), IIf(IsNull(RsDetails("CancelApprove")), "", (RsDetails("CancelApprove").value)), (RsDetails("ApprovDate").value))
          GRID2.TextMatrix(Num, GRID2.ColIndex("REMARKS")) = IIf(IsNull(RsDetails("REMARKS")), "", (RsDetails("REMARKS").value))
          If Not IsNull(RsDetails("CancelApprove").value) Then
                        If SystemOptions.UserInterface = ArabicInterface Then
                        GRID2.TextMatrix(Num, GRID2.ColIndex("Satus")) = "„—›Ê÷"
                        Else
                        GRID2.TextMatrix(Num, GRID2.ColIndex("Satus")) = "Unapproved"
                        End If
                        GRID2.cell(flexcpBackColor, Num, 1, Num, 14) = &HC0C0FF
                        
          Else
          
          If IsNull(RsDetails("ApprovDate").value) Then
          
                        If SystemOptions.UserInterface = ArabicInterface Then
                        GRID2.TextMatrix(Num, GRID2.ColIndex("Satus")) = "„‰ Ÿ—"
                        Else
                        GRID2.TextMatrix(Num, GRID2.ColIndex("Satus")) = "pending"
                        End If
                    GRID2.cell(flexcpBackColor, Num, 1, Num, 14) = &HC0FFFF
                    
               Else
               
                         If SystemOptions.UserInterface = ArabicInterface Then
                        GRID2.TextMatrix(Num, GRID2.ColIndex("Satus")) = "„⁄ „œ"
                        Else
                        GRID2.TextMatrix(Num, GRID2.ColIndex("Satus")) = "Approved"
                        End If
                        GRID2.cell(flexcpBackColor, Num, 1, Num, 14) = &HC0FFC0
                        
                     GRID2.ColComboList(GRID2.ColIndex("Remarks")) = "..."
            End If
                        
          End If
           If GRID2.TextMatrix(Num, GRID2.ColIndex("ScreenName")) = "FrmTypeExchange" Then
            
             GetValNameExch val(GRID2.TextMatrix(Num, GRID2.ColIndex("Transaction_ID"))), ToPerson, Price
             GRID2.TextMatrix(Num, GRID2.ColIndex("Price")) = Price
             GRID2.TextMatrix(Num, GRID2.ColIndex("ToPerson")) = ToPerson
             End If
 
RsDetails.MoveNext
If Num = RsDetails.RecordCount Then



End If

        Next Num
Else
 GRID2.rows = 1
    End If
RsDetails.Close

End Function
Public Function loadFlexGrid(Optional k As Integer = 0)
    Dim My_SQL As String
    Dim RowNum As Integer
    Dim ReCount As Integer
    Dim RsTemp As New ADODB.Recordset
Dim screenName1 As String
Dim timecatlog As Double
Dim hours As Double
Dim minutes As Double
Dim LateType As String
 Dim Price As Double
 Dim ToPerson As String
    If SystemOptions.SysDataBaseType = AccessDataBase Then
    
    ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
     
        Dim StrSQL As String

        StrSQL = "  SELECT  ApprovalData.OverProject    ,  dbo.ApprovalData.ExpectedtimeTime, dbo.ApprovalData.SendTime, dbo.ApprovalData.ScreenName, dbo.ApprovalData.levelo, "
 StrSQL = StrSQL + "      dbo.ApprovalData.EmpID, dbo.ApprovalData.levelorder, dbo.ApprovalData.currorder, dbo.ApprovalData.FromUser, dbo.ApprovalData.Transaction_ID,"
 StrSQL = StrSQL + "      dbo.ApprovalData.NoteID, dbo.ApprovalData.ApprovDate, dbo.ApprovalData.Remarks, dbo.TbLLevels.Name, dbo.TbLLevels.Namee, dbo.Screens.ScreenCaption,"
 StrSQL = StrSQL + "      dbo.Screens.ScreenTitleEng, dbo.ApprovalData.Currcursor, dbo.ApprovalData.id AS searchid, dbo.ApprovalData.NoteSerial, dbo.ApprovalData.Transaction_Date"
 StrSQL = StrSQL + "      FROM         dbo.ApprovalData left JOIN"
 StrSQL = StrSQL + "      dbo.TbLLevels ON dbo.ApprovalData.levelo = dbo.TbLLevels.LevelID INNER JOIN"
 StrSQL = StrSQL + "      dbo.Screens ON dbo.ApprovalData.ScreenName = dbo.Screens.ScreenName"
 
        StrSQL = StrSQL + "   Where (dbo.ApprovalData.Currcursor = 1) And (dbo.ApprovalData.EmpID = " & user_id & ")"
      If ScreenName <> "" Then
      StrSQL = StrSQL & "  AND (ApprovalData.ScreenName = N'" & ScreenName & "') "
      End If
      
        StrSQL = StrSQL + "   ORDER BY dbo.ApprovalData.currorder"
         

    End If

    RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Not (RsTemp.EOF Or RsTemp.BOF) Then

        With FG
            .rows = .FixedRows

            For ReCount = 1 To RsTemp.RecordCount
                .rows = .rows + 1
                RowNum = .rows - 1

                .TextMatrix(RowNum, .ColIndex("Ser")) = ReCount
                'Transaction_ID
                
                screenName1 = IIf(IsNull(RsTemp("ScreenName").value), "", RsTemp("ScreenName").value)
               .TextMatrix(RowNum, .ColIndex("id")) = IIf(IsNull(RsTemp("searchid").value), "", RsTemp("searchid").value)
               .TextMatrix(RowNum, .ColIndex("Transaction_ID")) = IIf(IsNull(RsTemp("Transaction_ID").value), "", RsTemp("Transaction_ID").value)
               .TextMatrix(RowNum, .ColIndex("ScreenName")) = screenName1 ' IIf(IsNull(RsTemp("ScreenName").value), "", RsTemp("ScreenName").value)
               .TextMatrix(RowNum, .ColIndex("NoteSerial")) = IIf(IsNull(RsTemp("NoteSerial").value), "", RsTemp("NoteSerial").value)
               .TextMatrix(RowNum, .ColIndex("Transaction_Date")) = IIf(IsNull(RsTemp("Transaction_Date").value), "", Format(RsTemp("Transaction_Date").value, "YYYY/MM/DD"))
                        
        Dim OverProject As Double
        .TextMatrix(RowNum, .ColIndex("OverProject")) = IIf(IsNull(RsTemp("OverProject").value), 0, RsTemp("OverProject").value)
         



.TextMatrix(RowNum, .ColIndex("noteid")) = IIf(IsNull(RsTemp("noteid").value), "", RsTemp("noteid").value)

             .TextMatrix(RowNum, .ColIndex("FromUser")) = IIf(IsNull(RsTemp("FromUser").value), "", RsTemp("FromUser").value)
             .TextMatrix(RowNum, .ColIndex("SendTime")) = IIf(IsNull(RsTemp("SendTime").value), "", Format(RsTemp("SendTime").value, "YYYY/MM/DD  HH:MM AM/PM"))
             .TextMatrix(RowNum, .ColIndex("ExpectedtimeTime")) = IIf(IsNull(RsTemp("ExpectedtimeTime").value), "", RsTemp("ExpectedtimeTime").value)
        '     .TextMatrix(RowNum, .ColIndex("LateType")) = GetTimeforTransaction(screenName1, timecatlog)
             If SystemOptions.UserInterface = ArabicInterface Then
             .TextMatrix(RowNum, .ColIndex("show")) = "«÷€ÿ ··⁄—÷ "
             .TextMatrix(RowNum, .ColIndex("Approve")) = "«÷€ÿ ··≈⁄ „«œ "
             .TextMatrix(RowNum, .ColIndex("CancelApprove")) = "—›÷ ··≈⁄ „«œ "
             
             Else
             .TextMatrix(RowNum, .ColIndex("show")) = "Show"
             .TextMatrix(RowNum, .ColIndex("Approve")) = "Approve"
             .TextMatrix(RowNum, .ColIndex("CancelApprove")) = "Cancel Approve"
             End If
             
             If screenName1 = "FrmTypeExchange" Then
            
             GetValNameExch val(.TextMatrix(RowNum, .ColIndex("Transaction_ID"))), ToPerson, Price
             .TextMatrix(RowNum, .ColIndex("Price")) = Price
             .TextMatrix(RowNum, .ColIndex("ToPerson")) = ToPerson
             End If



             Dim timediff As String
             timediff = DateDiff("N", .TextMatrix(RowNum, .ColIndex("ExpectedtimeTime")), Now)
             
             '"YYYY/MM/DD  HH:MM AM/PM"
             .TextMatrix(RowNum, .ColIndex("ExpectedtimeTime")) = Format(.TextMatrix(RowNum, .ColIndex("ExpectedtimeTime")), "dd/MM/yyyy  HH:MM AM/PM")
If timediff > 0 Then


hours = timediff \ 60
minutes = timediff - (hours * 60)
LateType = hours & ":" & minutes
  .TextMatrix(RowNum, .ColIndex("LateType")) = LateType
      .cell(flexcpBackColor, RowNum, 0, RowNum, 21) = &HFF&
 Else
 LateType = ""
 End If
 If k = RowNum Then
 .cell(flexcpBackColor, RowNum, 0, RowNum, 21) = &HC0C0&
 End If
 ' If timecatlog = 0 Then
 '           If SystemOptions.UserInterface = ArabicInterface Then
 '           .TextMatrix(RowNum, .ColIndex("LateType")) = "œﬁÌﬁ…"
 '           Else
 '           .TextMatrix(RowNum, .ColIndex("LateType")) = "Minute"
 '           End If
 '
 '
 ' ElseIf timecatlog = 1 Then
 '             If SystemOptions.UserInterface = ArabicInterface Then
 '           .TextMatrix(RowNum, .ColIndex("LateType")) = "”«⁄Â"
 '           Else
 '           .TextMatrix(RowNum, .ColIndex("LateType")) = "Hour"
 '           End If
 ' ElseIf timecatlog = 2 Then
 '             If SystemOptions.UserInterface = ArabicInterface Then
 '           .TextMatrix(RowNum, .ColIndex("LateType")) = "ÌÊ„"
 '           Else
 '           .TextMatrix(RowNum, .ColIndex("LateType")) = "Day/s"
 '           End If
 ' End If
   
             
            
                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(RowNum, .ColIndex("ScreenCaption")) = IIf(IsNull(RsTemp("ScreenCaption").value), "", RsTemp("ScreenCaption").value)
                    .TextMatrix(RowNum, .ColIndex("LevelName")) = IIf(IsNull(RsTemp("name").value), "", RsTemp("name").value)
                Else
                    .TextMatrix(RowNum, .ColIndex("ScreenCaption")) = IIf(IsNull(RsTemp("ScreenTitleEng").value), "", RsTemp("ScreenTitleEng").value)
                   .TextMatrix(RowNum, .ColIndex("LevelName")) = IIf(IsNull(RsTemp("namee").value), "", RsTemp("namee").value)
                End If
             
 
            
                .ColComboList(.ColIndex("Show")) = "..."
                
                .ColComboList(.ColIndex("Approve")) = "..."
             .ColComboList(.ColIndex("CancelApprove")) = "..."
                RsTemp.MoveNext
            Next ReCount

            .AutoSize 0, .Cols - 1, False
        End With
Else
   FG.rows = 1
    End If

End Function
Sub GetValNameExch(Optional ID As Double, Optional ByRef ToPerson As String, Optional ByRef Price As Double)
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
sql = "select * from TblExchange where ID =" & ID & ""
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
ToPerson = IIf(IsNull(rs2("ToPerson").value), "", rs2("ToPerson").value)
Price = IIf(IsNull(rs2("Price").value), "", rs2("Price").value)
Else
Price = 0
ToPerson = ""
End If
End Sub
Private Sub fg_Click()
Dim i As Integer
With FG
i = .row
If .ColKey(FG.Col) <> "Show" And .ColKey(FG.Col) <> "Approve" And .ColKey(FG.Col) <> "CancelApprove" And .ColKey(FG.Col) <> "Remarks" Then
loadFlexGrid i
ElseIf .ColKey(FG.Col) <> "Remarks" Then
 .cell(flexcpBackColor, i, 0, i, 16) = &HC0C0&
End If

End With

End Sub

Private Sub Fg_DblClick()
Exit Sub
End Sub

Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim BGround As New ClsBackGroundPic
    Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos
   Rd(6).value = True
    Dim BolShowRequest As Boolean
    Dcombos.GetUsers Me.DCboUserName
    DCboUserName.BoundText = user_id
    DTPicker1.value = Date
    DTPicker2.value = Date
     DTPicker3.value = Date
      DTPicker4.value = Date
      
    Me.Height = 10000
    Me.Width = 17600
    ToDate.value = Date
    FrmDate.value = Date
    FrmDate.value = DateAdd("d", -1, FrmDate.value)
    If SystemOptions.UserInterface = ArabicInterface Then
        VSFlexGrid1.ColComboList(VSFlexGrid1.ColIndex("Status")) = "#1; ÃœÌœ|#2; „⁄ „œ|#3; „—›Ê÷"
    ElseIf SystemOptions.UserInterface = EnglishInterface Then
        VSFlexGrid1.ColComboList(VSFlexGrid1.ColIndex("Status")) = "#1;New  |#2;Approve |#3;Rejection "
    End If
    FillInvoice
    'ToDate.value = DateAdd("d", 1, ToDate.value)
    fillapprovData
C1Tab1.CurrTab = 0

        Me.left = (mdifrmmain.Width - Me.Width) / 2 - 1200
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500
    
            Me.left = 0 '(mdifrmmain.Width - Me.Width) / 2
    Me.top = -100 '(mdifrmmain.Height - Me.Height) / 2 - 500

    Me.Width = (mdifrmmain.Width) - 500
    Me.Height = (mdifrmmain.Height) - 600

    'FormPostion Me, GetPostion
    LoadIcons
    FG.WallPaper = BGround.Picture
 '   BolShowRequest = GetSetting(StrAppRegPath, "View_Type", "InstallmentMustPayed", True)

    If BolShowRequest = True Then
        ChkShow.value = Unchecked
    Else
        ChkShow.value = Checked
    End If

    loadFlexGrid
'    Resize_Form Me, ReportSize

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub ChangeLang()
Label15.Caption = "From Date"
Label14.Caption = "To Date"
Label13.Caption = "Show"
Rd(3).RightToLeft = False
Rd(4).RightToLeft = False
Rd(5).RightToLeft = False
Rd(6).RightToLeft = False
ALLButton2.Caption = "Update"
ISButton2.Caption = "Exit"
    Me.Caption = "Doc. To  Approve"
    Rd(6).Caption = "New"
    Rd(5).Caption = "All"
    Rd(4).Caption = "Approved"
    Rd(3).Caption = "Rejection"
    LblCaption.Caption = Me.Caption
    ChkShow.Caption = "Dont Show at Start"
    Label1.Caption = "Data Based in your System Date"
    Me.CmdExit.Caption = "Exit"
    Me.CmdPrint.Caption = "Print"
C1Tab1.Caption = "Doc. To  Approve| Doc. Status|Appove Invoice"
Me.Label3.Caption = "  Doc. Approved/Not Approved"
CmdAdd.Caption = "Refresh"
Frame1.Caption = "Color Map"
Label2.Caption = "Late"
Label4.Caption = "From Date"
Label5.Caption = "To Date"
Label6.Caption = "Show Doc."
Label7.Caption = "Rejected"
Label8.Caption = "Waiting"
Label9.Caption = "Certified"
Rd(0).RightToLeft = False
Rd(1).RightToLeft = False
Rd(2).RightToLeft = False
Rd(0).Caption = "Unapproved"
Rd(1).Caption = "Approved"
Rd(2).Caption = "All"
 Me.ISButton1.Caption = "Exit"
ALLButton1.Caption = "Update"
Frame2.Caption = Frame1.Caption
With VSFlexGrid1
.TextMatrix(0, .ColIndex("Emp_Name")) = "Employee"
.TextMatrix(0, .ColIndex("Fullcode")) = "Customer Code"
.TextMatrix(0, .ColIndex("CusName")) = "Customer Name"
.TextMatrix(0, .ColIndex("TransDate")) = "Date Invoice"
.TextMatrix(0, .ColIndex("IssueDate")) = "Due Date"
.TextMatrix(0, .ColIndex("BillValue")) = "Value"
.TextMatrix(0, .ColIndex("SkipValue")) = "Credit Limit"
.TextMatrix(0, .ColIndex("Value")) = "Balance"
.TextMatrix(0, .ColIndex("LimitValue")) = "Value Skip"
.TextMatrix(0, .ColIndex("SkipNoDay")) = "Credit Period"
.TextMatrix(0, .ColIndex("LimitDay")) = "Period Skip"
.TextMatrix(0, .ColIndex("Approve")) = "Approve "
.TextMatrix(0, .ColIndex("CancelApprove")) = "Cancel Approve "
.TextMatrix(0, .ColIndex("Status")) = "Status "
.TextMatrix(0, .ColIndex("Remarks")) = "Remarks "
End With
    With GRID2
    .TextMatrix(0, .ColIndex("ToPerson")) = "To Person"
    .TextMatrix(0, .ColIndex("Price")) = "Price"
    .TextMatrix(0, .ColIndex("Approved")) = "Approved"
        .TextMatrix(0, .ColIndex("Approved")) = "Approved"
        .TextMatrix(0, .ColIndex("levelName")) = "Level"
        .TextMatrix(0, .ColIndex("EmpName")) = "Employee"
        .TextMatrix(0, .ColIndex("ApprovDate")) = "Approve Date"
        .TextMatrix(0, .ColIndex("Remarks")) = "Notes"
        .TextMatrix(0, .ColIndex("NoteSerial")) = "Serial"
        .TextMatrix(0, .ColIndex("ScreenCaption")) = "Doc"
        .TextMatrix(0, .ColIndex("show")) = "Show"
        .TextMatrix(0, .ColIndex("Remarks")) = "Remarks"
        .TextMatrix(0, .ColIndex("Satus")) = "Status"
        
    End With


    With Me.FG
    
    .TextMatrix(0, .ColIndex("Ser")) = "Serial"
    .TextMatrix(0, .ColIndex("ToPerson")) = "To Person"
    .TextMatrix(0, .ColIndex("Price")) = "Price"
        .TextMatrix(0, .ColIndex("FromUser")) = "From User"
        .TextMatrix(0, .ColIndex("SendTime")) = "Send Time"
        .TextMatrix(0, .ColIndex("ExpectedtimeTime")) = "Exp. Time"
        .TextMatrix(0, .ColIndex("NoteSerial")) = "Serial"
        .TextMatrix(0, .ColIndex("ScreenCaption")) = "Doc"
        .TextMatrix(0, .ColIndex("LateType")) = "Late Time"
        .TextMatrix(0, .ColIndex("Transaction_Date")) = "Date"
        .TextMatrix(0, .ColIndex("LevelName")) = "Level"
        .TextMatrix(0, .ColIndex("show")) = "Show"
        .TextMatrix(0, .ColIndex("Approve")) = "Approve"
        .TextMatrix(0, .ColIndex("CancelApprove")) = "Cancel"
        .TextMatrix(0, .ColIndex("Remarks")) = "Remarks"
        

    End With



End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap

    If ChkShow.value = Checked Then
        SaveSetting StrAppRegPath, "View_Type", "InstallmentMustPayed", False
    Else
        SaveSetting StrAppRegPath, "View_Type", "InstallmentMustPayed", True
    End If

    FormPostion Me, SavePostion
    Exit Sub
ErrTrap:
End Sub

Private Sub LoadIcons()
    On Error GoTo ErrTrap

    With FG
        .cell(flexcpPicture, 0, .ColIndex("Name")) = mdifrmmain.ImgLstTree.ListImages("User").Picture
        .cell(flexcpPicture, 0, .ColIndex("BillIID")) = mdifrmmain.ImgLstTree.ListImages("number").Picture
        .cell(flexcpPicture, 0, .ColIndex("TransDate")) = mdifrmmain.ImgLstTree.ListImages("qty").Picture
        .cell(flexcpPicture, 0, .ColIndex("QestNum")) = mdifrmmain.ImgLstTree.ListImages("number").Picture
        .cell(flexcpPicture, 0, .ColIndex("DueDate")) = mdifrmmain.ImgLstTree.ListImages("Date").Picture
        .cell(flexcpPicture, 0, .ColIndex("Value")) = mdifrmmain.ImgLstTree.ListImages("Price").Picture
        .cell(flexcpPictureAlignment, 0, 0, 0, .Cols - 1) = flexPicAlignRightCenter
    End With

    Exit Sub
ErrTrap:
Resume Next
End Sub

Private Sub FrmDate_Change()
fillapprovData
End Sub

Private Sub GRID2_BeforeEdit(ByVal row As Long, ByVal Col As Long, Cancel As Boolean)
'With GRID2
'If .ColKey(Col) <> "Show" Then
'Cancel = True
'End If
'End With
End Sub

Private Sub GRID2_CellButtonClick(ByVal row As Long, ByVal Col As Long)
Dim currrentScreenName As String
Dim newapprovalno As Double
            Dim sql As String
                Dim X As Integer
 
With GRID2
currrentScreenName = (.TextMatrix(row, .ColIndex("ScreenName")))
Select Case .ColKey(Col)
Case "Remarks"
Frame3.Visible = False
lbl(25).Caption = ""
lbl(25).Caption = .TextMatrix(row, .ColIndex("Remarks"))
Frame3.Visible = True
           Case "Show"

           

                    If currrentScreenName = "FrmPO" Then
                      Unload FrmPO
                          FrmPO.show
                              FrmPO.Retrive val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                    ElseIf currrentScreenName = "FrmTransacRegistr" Then
Unload FrmTransacRegistr
                         FrmTransacRegistr.show
                         FrmTransacRegistr.FindRec val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                        ElseIf currrentScreenName = "FrmPO1" Then
Unload FrmPO1
                         FrmPO1.show
                         FrmPO1.Retrive val(.TextMatrix(row, .ColIndex("Transaction_ID")))
           ElseIf currrentScreenName = "FrmBankPledge4" Then
                               Unload FrmBankPledge4
                           FrmBankPledge4.show
                         FrmBankPledge4.Retrive val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                         
                      ElseIf currrentScreenName = "FrmPO2" Then
    Unload FrmPO2
                           FrmPO2.show
                               FrmPO2.Retrive val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                   ElseIf currrentScreenName = "FrmPO3" Then
Unload FrmPO3
                       FrmPO3.show
                         FrmPO3.Retrive val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                        
                      ElseIf currrentScreenName = "FrmPO4" Then
   Unload FrmPO4

                               FrmPO4.show
                               FrmPO4.Retrive val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                        
                             ElseIf currrentScreenName = "FrmPO5" Then
   Unload FrmPO5
                               FrmPO5.show
                               FrmPO5.Retrive val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                        
                        
                                ElseIf currrentScreenName = "FrmPO6" Then
                                  Unload FrmPO6
                               FrmPO6.show
                               FrmPO6.Retrive val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                        
                                      ElseIf currrentScreenName = "FrmPO7" Then
                                      Unload FrmPO7
                         FrmPO7.show
                         FrmPO7.Retrive val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                                               ElseIf currrentScreenName = "FrmPO8" Then
                                               Unload FrmPO8
                         FrmPO8.show
                         FrmPO8.Retrive val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                                                         ElseIf currrentScreenName = "FrmPO10" Then
                                                         Unload FrmPO10
                         FrmPO10.show
                         FrmPO10.Retrive val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                        
                            
                                                              ElseIf currrentScreenName = "FrmEmpSalary3" Then
                                                              Unload FrmEmpSalary3
                         FrmEmpSalary3.show
                         FrmEmpSalary3.Retrive (val(Me.FG.TextMatrix(Me.FG.row, Me.FG.ColIndex("NoteSerial"))))
                        
                            
                        
                     ElseIf currrentScreenName = "FrmEmpsAdvanceRequest" Then
                       Unload FrmEmpsAdvanceRequest
                              FrmEmpsAdvanceRequest.show
                              FrmEmpsAdvanceRequest.Retrive val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                     
              ElseIf currrentScreenName = "Frmpassover" Then
                   '    FrmPassover.show
                   '     FrmPassover.Retrive val(.TextMatrix(Row, .ColIndex("Transaction_ID")))
               ElseIf currrentScreenName = "FrmVocationEntitlements" Then
                       FrmVocationEntitlements.show
                        FrmVocationEntitlements.Retrive val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                ElseIf currrentScreenName = "FrmBusinessJob" Then
                  Unload FrmBusinessJob
                              FrmBusinessJob.show
                               FrmBusinessJob.Retrive val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                 ElseIf currrentScreenName = "FrmEmbarkation" Then
                 Unload FrmEmbarkation
                        FrmEmbarkation.show
                         FrmEmbarkation.Retrive val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                        
                            ElseIf currrentScreenName = "formvocatinl" Then
                            Unload formvocatinl
                       formvocatinl.show
                        formvocatinl.Retrive val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                        
                              
                                 
                            ElseIf currrentScreenName = "FormEmpMoveDepartment" Then
                                     Unload FormEmpMoveDepartment
                          FormEmpMoveDepartment.show
                           FormEmpMoveDepartment.Retrive val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                        
                        
                                   ElseIf currrentScreenName = "FrmEmpsAdvanceRequest" Then
                                     Unload FrmEmpsAdvanceRequest
                               FrmEmpsAdvanceRequest.show
                                FrmEmpsAdvanceRequest.Retrive val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                         
                         
                                        ElseIf currrentScreenName = "FrmTypeExchange" Then
                                        Unload FrmTypeExchange
                       FrmTypeExchange.show
                         FrmTypeExchange.Retrive val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                         
                                            ElseIf currrentScreenName = "FrmShipmentOrder" Then
                                              Unload FrmShipmentOrder
                              FrmShipmentOrder.show
                               FrmShipmentOrder.Retrive val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                                            
                             ElseIf currrentScreenName = "FrmCreditFacicity" Then
                               Unload FrmCreditFacicity
                                         FrmCreditFacicity.show
                                 FrmCreditFacicity.Retrive val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                              ElseIf currrentScreenName = "FrmAccEditJournal" Then
                              Unload FrmAccEditJournal
                                   FrmAccEditJournal.show
                       
                        FrmAccEditJournal.Retrive (.TextMatrix(row, .ColIndex("NoteSerial")))
                              ElseIf currrentScreenName = "FrmAccEditJournal4" Then
                              Unload FrmAccEditJournal4
                                   FrmAccEditJournal4.show
                        FrmAccEditJournal4.Retrive (.TextMatrix(row, .ColIndex("NoteSerial")))
                                ElseIf currrentScreenName = "FrmAccEditJournal1" Then
                                Unload FrmAccEditJournal1
                                   FrmAccEditJournal1.show
                        FrmAccEditJournal1.Retrive val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                        
                           ElseIf currrentScreenName = "FrmExpenses5" Then
                           Unload FrmExpenses5
                                   FrmExpenses5.show
                        FrmExpenses5.Retrive val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                           ElseIf currrentScreenName = "FrmAccEditJournal3" Then
                           Unload FrmAccEditJournal3
                                   FrmAccEditJournal3.show
                        FrmAccEditJournal3.Retrive (.TextMatrix(row, .ColIndex("NoteSerial")))
                           ElseIf currrentScreenName = "FrmCashing" Then
                           Unload FrmCashing
                                   FrmCashing.show
                        FrmCashing.Retrive val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                           ElseIf currrentScreenName = "FrmDestruction" Then
                           Unload FrmDestruction
                                   FrmDestruction.show
                        FrmDestruction.Retrive val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                           ElseIf currrentScreenName = "FrmDestructionRet" Then
                           Unload FrmDestructionRet
                                   FrmDestructionRet.show
                        FrmDestructionRet.Retrive val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                          ElseIf currrentScreenName = "FrmExpenses30" Then
                          Unload FrmExpenses30
                                   FrmExpenses30.show
                        FrmExpenses30.Retrive val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                        
                           ElseIf currrentScreenName = "FrmExpenses301" Then
                          Unload FrmExpenses301
                                   FrmExpenses301.show
                        FrmExpenses301.Retrive val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                        
                        
                            ElseIf currrentScreenName = "FrmExpenses3" Then
                            Unload FrmExpenses3
                                   FrmExpenses3.show
                        FrmExpenses3.Retrive val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                        
                   ElseIf currrentScreenName = "FrmExpenses4" Then
                   Unload FrmExpenses4
                                   FrmExpenses4.show
                        FrmExpenses4.Retrive val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                   ElseIf currrentScreenName = "FrmBoxDrawing" Then
                   Unload FrmBoxDrawing
                                   FrmBoxDrawing.show
                        FrmBoxDrawing.Retrive val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                   ElseIf currrentScreenName = "FrmPayments1" Then
                   Unload FrmPayments1
                                   FrmPayments1.show
                        FrmPayments1.Retrive val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                             ElseIf currrentScreenName = "FrmBankPledge1" Then
                             Unload FrmBankPledge1
                                   FrmBankPledge1.show
                        FrmBankPledge1.Retrive val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                              ElseIf currrentScreenName = "FrmBankPledge2" Then
                                Unload FrmBankPledge2
                                     FrmBankPledge2.show
                          FrmBankPledge2.Retrive val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                              ElseIf currrentScreenName = "FrmBankPledge3" Then
                                Unload FrmBankPledge3
                                     FrmBankPledge3.show
                           FrmBankPledge3.Retrive val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                       ElseIf currrentScreenName = "FrmAdvancedHousingpayments" Then
                         Unload FrmAdvancedHousingpayments
                                    FrmAdvancedHousingpayments.show
                          FrmAdvancedHousingpayments.Retrive val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                          ElseIf currrentScreenName = "FrmMovingEmp" Then
                          Unload FrmMovingEmp
                                   FrmMovingEmp.show
                        FrmMovingEmp.Retrive val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                                 ElseIf currrentScreenName = "FrmQUesEmp" Then
                                   Unload FrmQUesEmp
                                     FrmQUesEmp.show
                          FrmQUesEmp.Retrive val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                           ElseIf currrentScreenName = "End_oF_service" Then
                                     Unload End_oF_service
                                     End_oF_service.show
                                     End_oF_service.Retrive val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                           ElseIf currrentScreenName = "FrmMoving" Then
                                     Unload FrmMoving
                                     FrmMoving.show
                                     FrmMoving.Retrive val(.TextMatrix(row, .ColIndex("Transaction_ID")))
                        
                   End If
End Select
End With
End Sub

 

Private Sub grid2_StartEdit(ByVal row As Long, ByVal Col As Long, Cancel As Boolean)
With GRID2
Select Case .ColKey(Col)
           Case "Show"
           .ColComboList(.ColIndex("Show")) = "..."
 End Select
 End With

End Sub
Sub FillInvoice()
Dim i As Integer
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
    VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
    VSFlexGrid1.rows = 1
sql = " SELECT     dbo.TblAproveInvoice.ID, dbo.TblAproveInvoice.CusID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode,"
sql = sql & "                       dbo.TblAproveInvoice.EmpID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode AS EmpCode, dbo.TblEmployee.Emp_Namee,"
sql = sql & "                        dbo.TblAproveInvoice.IssueDate, dbo.TblAproveInvoice.TransDate, dbo.TblAproveInvoice.BillValue, dbo.TblAproveInvoice.NoDay, dbo.TblAproveInvoice.SkipNoDay,"
sql = sql & "                        dbo.TblAproveInvoice.FlagApproved, dbo.TblAproveInvoice.[Value], dbo.TblAproveInvoice.SkipValue, dbo.TblAproveInvoice.UserID,"
sql = sql & "                        dbo.TblAproveInvoice.Remarks"
sql = sql & "   FROM         dbo.TblAproveInvoice LEFT OUTER JOIN"
sql = sql & "                        dbo.TblEmployee ON dbo.TblAproveInvoice.EmpID = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
sql = sql & "                        dbo.TblCustemers ON dbo.TblAproveInvoice.CusID = dbo.TblCustemers.CusID"
sql = sql & " where dbo.TblAproveInvoice.UserID =" & user_id & " "
If Rd(6).value = True Then
sql = sql & " and   dbo.TblAproveInvoice.FlagApproved is null"
End If
If Rd(4).value = True Then
sql = sql & " and   dbo.TblAproveInvoice.FlagApproved =1"
End If
If Rd(3).value = True Then
sql = sql & " and   dbo.TblAproveInvoice.FlagApproved =2"
End If
If Not IsNull(DTPicker1.value) Then
sql = sql & " and   dbo.TblAproveInvoice.TransDate >=" & SQLDate(DTPicker1.value, True) & ""
End If
If Not IsNull(DTPicker2.value) Then
sql = sql & " and   dbo.TblAproveInvoice.TransDate <=" & SQLDate(DTPicker2.value, True) & ""
End If

rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
rs2.MoveFirst
With VSFlexGrid1
.rows = .rows + rs2.RecordCount
For i = 1 To .rows - 1
.TextMatrix(i, .ColIndex("Ser")) = 1
.TextMatrix(i, .ColIndex("Status")) = IIf(IsNull(rs2("FlagApproved").value), 0, rs2("FlagApproved").value) + 1
.TextMatrix(i, .ColIndex("CusID")) = IIf(IsNull(rs2("CusID").value), 0, rs2("CusID").value)
.TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs2("ID").value), 0, rs2("ID").value)
.TextMatrix(i, .ColIndex("Fullcode")) = IIf(IsNull(rs2("Fullcode").value), "", rs2("Fullcode").value)
.TextMatrix(i, .ColIndex("IssueDate")) = IIf(IsNull(rs2("IssueDate").value), "", rs2("IssueDate").value)
.TextMatrix(i, .ColIndex("TransDate")) = IIf(IsNull(rs2("TransDate").value), "", rs2("TransDate").value)
.TextMatrix(i, .ColIndex("BillValue")) = IIf(IsNull(rs2("BillValue").value), "", rs2("BillValue").value)
.TextMatrix(i, .ColIndex("SkipNoDay")) = IIf(IsNull(rs2("SkipNoDay").value), "", rs2("SkipNoDay").value)
.TextMatrix(i, .ColIndex("Value")) = IIf(IsNull(rs2("Value").value), "", rs2("Value").value)
.TextMatrix(i, .ColIndex("SkipValue")) = IIf(IsNull(rs2("SkipValue").value), "", rs2("SkipValue").value)
.TextMatrix(i, .ColIndex("UserID")) = IIf(IsNull(rs2("UserID").value), "", rs2("UserID").value)
.TextMatrix(i, .ColIndex("Remarks")) = IIf(IsNull(rs2("Remarks").value), "", rs2("Remarks").value)
.TextMatrix(i, .ColIndex("LimitValue")) = val(.TextMatrix(i, .ColIndex("BillValue"))) - (val(.TextMatrix(i, .ColIndex("SkipValue"))) - val(.TextMatrix(i, .ColIndex("Value"))))
.TextMatrix(i, .ColIndex("LimitDay")) = IIf(IsNull(rs2("NoDay").value), 0, Abs(rs2("NoDay").value)) - val(.TextMatrix(i, .ColIndex("SkipNoDay")))
If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(rs2("CusName").value), "", rs2("CusName").value)
.TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(rs2("Emp_Name").value), "", rs2("Emp_Name").value)
Else
.TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(rs2("CusName").value), "", rs2("CusName").value)
.TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(rs2("Emp_Namee").value), "", rs2("Emp_Namee").value)
End If
             If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(i, .ColIndex("Approve")) = "«÷€ÿ ··≈⁄ „«œ "
             .TextMatrix(i, .ColIndex("CancelApprove")) = "—›÷ ··≈⁄ „«œ "
             
             Else
             .TextMatrix(i, .ColIndex("Approve")) = "Approve"
             .TextMatrix(i, .ColIndex("CancelApprove")) = "Cancel Approve"
             End If

.ColComboList(.ColIndex("Approve")) = "..."
.ColComboList(.ColIndex("CancelApprove")) = "..."
rs2.MoveNext
Next i
End With
End If
End Sub


Sub FillINotapprovedGL()
Dim i As Integer
Dim sql As String
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
    VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
    VSFlexGrid1.rows = 1
sql = " SELECT     dbo.Notes.NoteDate, dbo.Notes.NoteType, dbo.Notes.NoteSerial, dbo.Notes.NoteSerial1, dbo.TblNotesTypes.NotesTypeName, "
sql = sql & "                      dbo.TblNotesTypes.NotesTypeNameE"
sql = sql & "  FROM         dbo.Notes INNER JOIN"
sql = sql & "                      dbo.DOUBLE_ENTREY_VOUCHERS ON dbo.Notes.NoteID = dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID INNER JOIN"
sql = sql & "                      dbo.TblNotesTypes ON dbo.Notes.NoteType = dbo.TblNotesTypes.NotesType"
sql = sql & "  Where (dbo.DOUBLE_ENTREY_VOUCHERS.Posted = 1)"
If Not IsNull(DTPicker1.value) Then
sql = sql & " and    dbo.Notes.NoteDate >=" & SQLDate(DTPicker3.value, True) & ""
End If
If Not IsNull(DTPicker2.value) Then
sql = sql & " and    dbo.Notes.NoteDate  <=" & SQLDate(DTPicker4.value, True) & ""
End If

sql = sql & "  GROUP BY dbo.Notes.NoteDate, dbo.Notes.NoteType, dbo.Notes.NoteSerial, dbo.Notes.NoteSerial1, dbo.TblNotesTypes.NotesTypeName, dbo.TblNotesTypes.NotesTypeNamee"



rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
rs2.MoveFirst
With GridGE
.rows = rs2.RecordCount + 1

For i = 1 To .rows - 1
.TextMatrix(i, .ColIndex("Ser")) = i
 .TextMatrix(i, .ColIndex("NoteDate")) = IIf(IsNull(rs2("NoteDate").value), "", rs2("NoteDate").value)
.TextMatrix(i, .ColIndex("NoteSerial")) = IIf(IsNull(rs2("NoteSerial").value), "", rs2("NoteSerial").value)
.TextMatrix(i, .ColIndex("NoteSerial1")) = IIf(IsNull(rs2("NoteSerial1").value), "", rs2("NoteSerial1").value)
  
 If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(i, .ColIndex("NotesTypeName")) = IIf(IsNull(rs2("NotesTypeName").value), "", rs2("NotesTypeName").value)
 
Else
.TextMatrix(i, .ColIndex("NotesTypeName")) = IIf(IsNull(rs2("NotesTypeNamee").value), "", rs2("NotesTypeNamee").value)
 
End If
    .ColComboList(.ColIndex("Show")) = "..."
 rs2.MoveNext
Next i
End With
End If
End Sub


Private Sub GRIDGE_CellButtonClick(ByVal row As Long, ByVal Col As Long)
With GridGE
Select Case .ColKey(Col)
Case "Show"
            If checkApility("FrmAccEditJournal") = False Then
                Exit Sub
            End If

            FrmAccEditJournal.show
            
 FrmAccEditJournal.show
 FrmAccEditJournal.Retrive (.TextMatrix(row, .ColIndex("NoteSerial")))
 
 End Select
End With

End Sub

Private Sub ISButton1_Click()
 Unload Me
End Sub

Private Sub ISButton2_Click()
Unload Me
End Sub

Private Sub Label13_Click()
FillInvoice
End Sub

Private Sub Label16_Click()
FillINotapprovedGL
End Sub

Private Sub LblCaption_Click()
    On Error GoTo ErrTrap

    If Me.WindowState = vbNormal Then
        Me.WindowState = vbMaximized
    Else
        Me.WindowState = vbNormal
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Rd_Click(index As Integer)
fillapprovData
End Sub

Private Sub Timer1_Timer()
Dim RowNum As Double
        With FG
            
            For RowNum = 1 To .rows - 1
          
                   If val(.TextMatrix(RowNum, .ColIndex("OverProject"))) = 1 Then
                                      If .cell(flexcpBackColor, RowNum, 10, RowNum, 10) = &H8080FF Then
                                               .cell(flexcpBackColor, RowNum, 10, RowNum, 10) = vbWhite
                                    Else
                                              .cell(flexcpBackColor, RowNum, 10, RowNum, 10) = &H8080FF
                                    End If
            
            
            End If
            
                
                Next RowNum
    
    End With
End Sub

Private Sub ToDate_Change()
fillapprovData
End Sub

Private Sub VSFlexGrid1_CellButtonClick(ByVal row As Long, ByVal Col As Long)
With VSFlexGrid1
Select Case .ColKey(Col)
Case "Approve"
Cn.Execute "Update TblAproveInvoice set FlagApproved=1 where id=" & val(.TextMatrix(row, .ColIndex("id"))) & " "
Case "CancelApprove"
Cn.Execute "Update TblAproveInvoice set FlagApproved=2 where id=" & val(.TextMatrix(row, .ColIndex("id"))) & " "
End Select
End With
FillInvoice
End Sub

