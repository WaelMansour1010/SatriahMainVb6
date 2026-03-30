VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmContractExam 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   10290
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13905
   Icon            =   "FrmContractExam.frx":0000
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   ScaleHeight     =   10290
   ScaleMode       =   0  'User
   ScaleTop        =   60
   ScaleWidth      =   13905
   WindowState     =   2  'Maximized
   Begin VB.TextBox TxtModFlg 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      Enabled         =   0   'False
      Height          =   285
      Left            =   15600
      TabIndex        =   4
      Text            =   "modflag"
      Top             =   4200
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.ComboBox CmbType 
      BackColor       =   &H80000018&
      Height          =   315
      ItemData        =   "FrmContractExam.frx":6852
      Left            =   15480
      List            =   "FrmContractExam.frx":6862
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   3120
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.TextBox TxtVac_ID 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      Height          =   240
      Left            =   15600
      TabIndex        =   2
      Top             =   1200
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Frame Frmo2 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   375
      Left            =   15600
      TabIndex        =   1
      Top             =   1680
      Visible         =   0   'False
      Width           =   3105
   End
   Begin VB.TextBox Emp_id 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   15240
      TabIndex        =   0
      Top             =   5640
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSDataListLib.DataCombo DCUser 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   15840
      TabIndex        =   5
      Tag             =   "„‰ ›÷·ﬂ √œŒ· —ﬁ„ «·ﬁ÷Ì…"
      Top             =   960
      Visible         =   0   'False
      Width           =   2100
      _ExtentX        =   3704
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      BackColor       =   -2147483624
      Text            =   ""
      RightToLeft     =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo DCPreFix 
      Height          =   315
      Left            =   15480
      TabIndex        =   6
      Top             =   2280
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSComctlLib.ImageList GrdImageList 
      Left            =   15600
      Top             =   3720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmContractExam.frx":687B
            Key             =   "CompanyName"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmContractExam.frx":6C15
            Key             =   "Ser"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmContractExam.frx":6FAF
            Key             =   "Vac_Name"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmContractExam.frx":7349
            Key             =   "ShareCount"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmContractExam.frx":76E3
            Key             =   "Dis_Count"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmContractExam.frx":7A7D
            Key             =   "Bouns"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmContractExam.frx":7E17
            Key             =   "SharesValue"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmContractExam.frx":83B1
            Key             =   "BuyValue"
         EndProperty
      EndProperty
   End
   Begin ImpulseButton.ISButton BtnUpdate 
      Height          =   330
      Left            =   15600
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   " ÕœÌÀ ﬁ«⁄œ… «·»Ì«‰« "
      Top             =   5040
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
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
      ButtonImage     =   "FrmContractExam.frx":874B
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin ImpulseButton.ISButton ISButton1 
      Height          =   405
      Left            =   18840
      TabIndex        =   9
      TabStop         =   0   'False
      ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
      Top             =   120
      Visible         =   0   'False
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   714
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "ÿ»«⁄… "
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
      ButtonImage     =   "FrmContractExam.frx":EFAD
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
      DisabledImageStyle=   1
   End
   Begin ImpulseButton.ISButton btnQuery 
      Height          =   330
      Left            =   16920
      TabIndex        =   10
      TabStop         =   0   'False
      ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
      Top             =   120
      Visible         =   0   'False
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   582
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "»ÕÀ"
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
      ButtonImage     =   "FrmContractExam.frx":1580F
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin C1SizerLibCtl.C1Elastic C1Elastic12 
      Height          =   10290
      Left            =   0
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   0
      Width           =   13905
      _cx             =   24527
      _cy             =   18150
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
      Begin C1SizerLibCtl.C1Tab TabMain 
         Height          =   10155
         Left            =   90
         TabIndex        =   12
         Top             =   75
         Width           =   13770
         _cx             =   24289
         _cy             =   17912
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
         Caption         =   "„œ… «·«Œ »«— ›Ì ⁄ﬁœ «·„ÊŸ›Ì‰| ‰»ÌÂ«  «·⁄ﬁÊœ «· Ì ” ‰ ÂÌ"
         Align           =   0
         CurrTab         =   1
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
            Height          =   9780
            Index           =   1
            Left            =   -14325
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   45
            Width           =   13680
            _cx             =   24130
            _cy             =   17251
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
               Height          =   9660
               Index           =   0
               Left            =   19065
               TabIndex        =   14
               Top             =   765
               Width           =   13470
               _cx             =   23760
               _cy             =   17039
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
               FormatString    =   $"FrmContractExam.frx":15BA9
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
            Begin C1SizerLibCtl.C1Elastic frm_Main 
               Height          =   9630
               Left            =   0
               TabIndex        =   15
               TabStop         =   0   'False
               Top             =   0
               Width           =   13395
               _cx             =   23627
               _cy             =   16986
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
               Frame           =   0
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
               Begin VB.Frame FraHeader 
                  BackColor       =   &H00FFFFFF&
                  BorderStyle     =   0  'None
                  Height          =   855
                  Left            =   13455
                  TabIndex        =   16
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   11745
                  Begin VB.TextBox tXTRootAccount 
                     Alignment       =   1  'Right Justify
                     Height          =   285
                     Left            =   3240
                     TabIndex        =   18
                     Top             =   360
                     Visible         =   0   'False
                     Width           =   2055
                  End
                  Begin VB.TextBox TxtName 
                     Alignment       =   1  'Right Justify
                     Height          =   285
                     Index           =   0
                     Left            =   6000
                     TabIndex        =   17
                     Top             =   240
                     Visible         =   0   'False
                     Width           =   2055
                  End
               End
               Begin C1SizerLibCtl.C1Elastic C1Elastic1 
                  Height          =   1200
                  Left            =   0
                  TabIndex        =   19
                  TabStop         =   0   'False
                  Top             =   9825
                  Width           =   13365
                  _cx             =   23574
                  _cy             =   2117
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
                  Begin ImpulseButton.ISButton btnNew 
                     Height          =   330
                     Left            =   12090
                     TabIndex        =   20
                     ToolTipText     =   "· ”ÃÌ· »Ì«‰«  ÃœÌœ…"
                     Top             =   600
                     Width           =   1110
                     _ExtentX        =   1958
                     _ExtentY        =   582
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
                     ButtonImage     =   "FrmContractExam.frx":15C69
                     ColorButton     =   14871017
                     DrawFocusRectangle=   0   'False
                  End
                  Begin ImpulseButton.ISButton btnModify 
                     Height          =   330
                     Left            =   10365
                     TabIndex        =   21
                     ToolTipText     =   "· ⁄œÌ· «·»Ì«‰«  «·Õ«·Ì…"
                     Top             =   600
                     Width           =   1365
                     _ExtentX        =   2408
                     _ExtentY        =   582
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
                     ButtonImage     =   "FrmContractExam.frx":1C4CB
                     ColorButton     =   14871017
                     DrawFocusRectangle=   0   'False
                  End
                  Begin ImpulseButton.ISButton btnSave 
                     Height          =   330
                     Left            =   8940
                     TabIndex        =   22
                     ToolTipText     =   "Õ›Ÿ «·»Ì«‰«  œ«Œ· ﬁ«⁄œ… «·»Ì«‰« "
                     Top             =   600
                     Width           =   1125
                     _ExtentX        =   1984
                     _ExtentY        =   582
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
                     ButtonImage     =   "FrmContractExam.frx":22D2D
                     ColorButton     =   14871017
                     DrawFocusRectangle=   0   'False
                  End
                  Begin ImpulseButton.ISButton BtnUndo 
                     Height          =   330
                     Left            =   7140
                     TabIndex        =   23
                     ToolTipText     =   "·· —«Ã⁄ ⁄‰ «·ÕœÀ Ê«·—ÃÊ⁄ «·Ï «·Ê÷⁄ «·ÿ»Ì⁄Ì"
                     Top             =   600
                     Width           =   1515
                     _ExtentX        =   2672
                     _ExtentY        =   582
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
                     ButtonImage     =   "FrmContractExam.frx":230C7
                     ColorButton     =   14871017
                     DrawFocusRectangle=   0   'False
                  End
                  Begin ImpulseButton.ISButton btnDelete 
                     Height          =   330
                     Left            =   5790
                     TabIndex        =   24
                     ToolTipText     =   "Õ–› «·»Ì«‰«  «·„Õœœ…"
                     Top             =   600
                     Width           =   1215
                     _ExtentX        =   2143
                     _ExtentY        =   582
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
                     ButtonImage     =   "FrmContractExam.frx":23461
                     ColorButton     =   14871017
                     DrawFocusRectangle=   0   'False
                  End
                  Begin ImpulseButton.ISButton ISButton5 
                     Height          =   420
                     Left            =   3870
                     TabIndex        =   25
                     TabStop         =   0   'False
                     ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
                     Top             =   600
                     Width           =   1185
                     _ExtentX        =   2090
                     _ExtentY        =   741
                     ButtonStyle     =   1
                     ButtonPositionImage=   1
                     Caption         =   "ÿ»«⁄… "
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
                     ButtonImage     =   "FrmContractExam.frx":239FB
                     ColorButton     =   14871017
                     DrawFocusRectangle=   0   'False
                     DisabledImageStyle=   1
                  End
                  Begin ImpulseButton.ISButton ISButton8 
                     Height          =   330
                     Left            =   120
                     TabIndex        =   26
                     TabStop         =   0   'False
                     ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
                     Top             =   600
                     Visible         =   0   'False
                     Width           =   1095
                     _ExtentX        =   1931
                     _ExtentY        =   582
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
                     ButtonImage     =   "FrmContractExam.frx":2A25D
                     ColorButton     =   14871017
                     DrawFocusRectangle=   0   'False
                  End
                  Begin ImpulseButton.ISButton btnCancel 
                     Height          =   330
                     Left            =   1200
                     TabIndex        =   27
                     ToolTipText     =   "«·Œ—ÊÃ «·Ï  «·‰«›–… «·—∆Ì”Ì…"
                     Top             =   600
                     Width           =   915
                     _ExtentX        =   1614
                     _ExtentY        =   582
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
                     ButtonImage     =   "FrmContractExam.frx":2A5F7
                     ColorButton     =   14871017
                     DrawFocusRectangle=   0   'False
                     DisabledImageStyle=   1
                  End
                  Begin MSDataListLib.DataCombo DCboUserName 
                     Height          =   315
                     Left            =   8400
                     TabIndex        =   28
                     Top             =   90
                     Width           =   3465
                     _ExtentX        =   6112
                     _ExtentY        =   556
                     _Version        =   393216
                     Enabled         =   0   'False
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin VB.Label LabCountRec 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     ForeColor       =   &H00C00000&
                     Height          =   195
                     Left            =   315
                     TabIndex        =   33
                     Top             =   240
                     Width           =   630
                  End
                  Begin VB.Label LabCurrRec 
                     Alignment       =   2  'Center
                     BackColor       =   &H00E2E9E9&
                     ForeColor       =   &H00800000&
                     Height          =   195
                     Left            =   2370
                     TabIndex        =   32
                     Top             =   240
                     Width           =   780
                  End
                  Begin VB.Label Label2 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "⁄œœ «·”Ã·« :"
                     Height          =   195
                     Index           =   1
                     Left            =   1080
                     TabIndex        =   31
                     Top             =   240
                     Width           =   1155
                  End
                  Begin VB.Label Label2 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·”Ã· «·Õ«·Ì:"
                     Height          =   195
                     Index           =   0
                     Left            =   3255
                     TabIndex        =   30
                     Top             =   240
                     Width           =   1245
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackStyle       =   0  'Transparent
                     Caption         =   "Õ—— »Ê«”ÿ…  "
                     Height          =   345
                     Index           =   14
                     Left            =   12270
                     TabIndex        =   29
                     Top             =   90
                     Width           =   1140
                  End
               End
               Begin C1SizerLibCtl.C1Elastic ELe 
                  Height          =   855
                  Index           =   18
                  Left            =   0
                  TabIndex        =   34
                  TabStop         =   0   'False
                  Top             =   0
                  Width           =   13455
                  _cx             =   23733
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
                  BackColor       =   16777215
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
                  Begin ImpulseButton.ISButton btnLast 
                     Height          =   360
                     Left            =   150
                     TabIndex        =   35
                     Top             =   255
                     Visible         =   0   'False
                     Width           =   450
                     _ExtentX        =   794
                     _ExtentY        =   635
                     ButtonStyle     =   1
                     ButtonPositionImage=   1
                     Caption         =   ""
                     BackColor       =   16777215
                     FontSize        =   12
                     FontName        =   "Arial"
                     FontBold        =   -1  'True
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Arial"
                        Size            =   12
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ButtonImage     =   "FrmContractExam.frx":2A991
                     ColorButton     =   16777215
                     AcclimateGrayTones=   -1  'True
                     DrawFocusRectangle=   0   'False
                     DisabledImageExtraction=   0
                     DisabledImageStyle=   1
                  End
                  Begin ImpulseButton.ISButton btnNext 
                     Height          =   360
                     Left            =   690
                     TabIndex        =   36
                     Top             =   255
                     Visible         =   0   'False
                     Width           =   450
                     _ExtentX        =   794
                     _ExtentY        =   635
                     ButtonStyle     =   1
                     ButtonPositionImage=   1
                     Caption         =   ""
                     BackColor       =   16777215
                     FontSize        =   12
                     FontName        =   "Arial"
                     FontBold        =   -1  'True
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Arial"
                        Size            =   12
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ButtonImage     =   "FrmContractExam.frx":2AD2B
                     ColorButton     =   16777215
                     DrawFocusRectangle=   0   'False
                     DisabledImageStyle=   1
                  End
                  Begin ImpulseButton.ISButton btnPrevious 
                     Height          =   360
                     Left            =   1350
                     TabIndex        =   37
                     Top             =   255
                     Visible         =   0   'False
                     Width           =   480
                     _ExtentX        =   847
                     _ExtentY        =   635
                     ButtonStyle     =   1
                     ButtonPositionImage=   1
                     Caption         =   ""
                     BackColor       =   16777215
                     FontSize        =   12
                     FontName        =   "Arial"
                     FontBold        =   -1  'True
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Arial"
                        Size            =   12
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ButtonImage     =   "FrmContractExam.frx":2B0C5
                     ColorButton     =   16777215
                     DrawFocusRectangle=   0   'False
                     DisabledImageStyle=   1
                  End
                  Begin ImpulseButton.ISButton btnFirst 
                     Height          =   360
                     Left            =   1950
                     TabIndex        =   38
                     Top             =   255
                     Visible         =   0   'False
                     Width           =   480
                     _ExtentX        =   847
                     _ExtentY        =   635
                     ButtonStyle     =   1
                     ButtonPositionImage=   1
                     Caption         =   ""
                     BackColor       =   16777215
                     FontSize        =   12
                     FontName        =   "Arial"
                     FontBold        =   -1  'True
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Arial"
                        Size            =   12
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ButtonImage     =   "FrmContractExam.frx":2B45F
                     ColorButton     =   16777215
                     DrawFocusRectangle=   0   'False
                     DisabledImageStyle=   1
                  End
                  Begin MSComCtl2.DTPicker DTPicker1 
                     Height          =   375
                     Left            =   3480
                     TabIndex        =   39
                     Top             =   120
                     Visible         =   0   'False
                     Width           =   1350
                     _ExtentX        =   2381
                     _ExtentY        =   661
                     _Version        =   393216
                     Enabled         =   0   'False
                     Format          =   113967105
                     CurrentDate     =   38784
                  End
                  Begin VB.Label Label1 
                     Alignment       =   1  'Right Justify
                     BackStyle       =   0  'Transparent
                     Caption         =   "„œ… «·«Œ »«— ›Ì ⁄ﬁœ «·„ÊŸ›Ì‰"
                     BeginProperty Font 
                        Name            =   "Arial"
                        Size            =   14.25
                        Charset         =   178
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00000000&
                     Height          =   450
                     Index           =   0
                     Left            =   7620
                     TabIndex        =   73
                     Top             =   150
                     Width           =   4650
                  End
                  Begin VB.Image Image1 
                     Height          =   690
                     Left            =   12375
                     Picture         =   "FrmContractExam.frx":2B7F9
                     Stretch         =   -1  'True
                     Top             =   120
                     Width           =   840
                  End
               End
               Begin C1SizerLibCtl.C1Elastic C1Elastic4 
                  Height          =   8490
                  Left            =   -210
                  TabIndex        =   40
                  TabStop         =   0   'False
                  Top             =   1080
                  Width           =   13455
                  _cx             =   23733
                  _cy             =   14975
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
                  Frame           =   0
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
                  Begin ImpulseButton.ISButton ISButton3 
                     Height          =   330
                     Left            =   240
                     TabIndex        =   41
                     ToolTipText     =   "«÷«›… «·»Ì«‰«  «·Ï «·œ« «"
                     Top             =   8040
                     Width           =   2940
                     _ExtentX        =   5186
                     _ExtentY        =   582
                     Caption         =   " ‰›Ì–"
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
                     ButtonImage     =   "FrmContractExam.frx":2CBFE
                     ColorButton     =   14871017
                     DrawFocusRectangle=   0   'False
                     DisabledImageExtraction=   0
                     LowerToggledContent=   0   'False
                  End
                  Begin ImpulseButton.ISButton ISButton2 
                     Height          =   330
                     Left            =   3960
                     TabIndex        =   42
                     ToolTipText     =   "«÷«›… «·»Ì«‰«  «·Ï «·œ« «"
                     Top             =   8040
                     Width           =   4125
                     _ExtentX        =   7276
                     _ExtentY        =   582
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
                     ButtonImage     =   "FrmContractExam.frx":33460
                     ColorButton     =   14871017
                     DrawFocusRectangle=   0   'False
                     DisabledImageExtraction=   0
                     LowerToggledContent=   0   'False
                  End
                  Begin VSFlex8Ctl.VSFlexGrid Fg1 
                     Height          =   7890
                     Left            =   120
                     TabIndex        =   43
                     Top             =   -90
                     Width           =   13425
                     _cx             =   23680
                     _cy             =   13917
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
                     BackColor       =   16777215
                     ForeColor       =   -2147483640
                     BackColorFixed  =   14871017
                     ForeColorFixed  =   -2147483630
                     BackColorSel    =   -2147483635
                     ForeColorSel    =   -2147483634
                     BackColorBkg    =   -2147483636
                     BackColorAlternate=   16776960
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
                     Cols            =   12
                     FixedRows       =   1
                     FixedCols       =   1
                     RowHeightMin    =   320
                     RowHeightMax    =   0
                     ColWidthMin     =   0
                     ColWidthMax     =   0
                     ExtendLastCol   =   0   'False
                     FormatString    =   $"FrmContractExam.frx":39CC2
                     ScrollTrack     =   0   'False
                     ScrollBars      =   3
                     ScrollTips      =   0   'False
                     MergeCells      =   0
                     MergeCompare    =   0
                     AutoResize      =   0   'False
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
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic ELe 
            Height          =   9780
            Index           =   0
            Left            =   45
            TabIndex        =   44
            TabStop         =   0   'False
            Top             =   45
            Width           =   13680
            _cx             =   24130
            _cy             =   17251
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
               Height          =   9645
               Index           =   1
               Left            =   19065
               TabIndex        =   45
               Top             =   840
               Width           =   13455
               _cx             =   23733
               _cy             =   17013
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
               FormatString    =   $"FrmContractExam.frx":39EC4
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
            Begin C1SizerLibCtl.C1Elastic ELe 
               Height          =   9780
               Index           =   2
               Left            =   765
               TabIndex        =   46
               TabStop         =   0   'False
               Top             =   210
               Width           =   13680
               _cx             =   24130
               _cy             =   17251
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
                  Height          =   9660
                  Index           =   2
                  Left            =   19065
                  TabIndex        =   47
                  Top             =   765
                  Width           =   13455
                  _cx             =   23733
                  _cy             =   17039
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
                  FormatString    =   $"FrmContractExam.frx":39F84
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
               Begin C1SizerLibCtl.C1Elastic C1Elastic2 
                  Height          =   9630
                  Left            =   -510
                  TabIndex        =   48
                  TabStop         =   0   'False
                  Top             =   -60
                  Width           =   13365
                  _cx             =   23574
                  _cy             =   16986
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
                  Frame           =   0
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
                     BackColor       =   &H00FFFFFF&
                     BorderStyle     =   0  'None
                     Height          =   840
                     Left            =   13410
                     TabIndex        =   49
                     Top             =   0
                     Visible         =   0   'False
                     Width           =   11745
                     Begin VB.TextBox TxtName 
                        Alignment       =   1  'Right Justify
                        Height          =   285
                        Index           =   1
                        Left            =   6000
                        TabIndex        =   51
                        Top             =   240
                        Visible         =   0   'False
                        Width           =   2055
                     End
                     Begin VB.TextBox Text1 
                        Alignment       =   1  'Right Justify
                        Height          =   285
                        Left            =   3240
                        TabIndex        =   50
                        Top             =   360
                        Visible         =   0   'False
                        Width           =   2055
                     End
                  End
                  Begin C1SizerLibCtl.C1Elastic C1Elastic3 
                     Height          =   1200
                     Left            =   0
                     TabIndex        =   52
                     TabStop         =   0   'False
                     Top             =   9825
                     Width           =   13350
                     _cx             =   23548
                     _cy             =   2117
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
                     Begin ImpulseButton.ISButton ISButton4 
                        Height          =   330
                        Left            =   12090
                        TabIndex        =   53
                        ToolTipText     =   "· ”ÃÌ· »Ì«‰«  ÃœÌœ…"
                        Top             =   600
                        Width           =   1110
                        _ExtentX        =   1958
                        _ExtentY        =   582
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
                        ButtonImage     =   "FrmContractExam.frx":3A044
                        ColorButton     =   14871017
                        DrawFocusRectangle=   0   'False
                     End
                     Begin ImpulseButton.ISButton ISButton6 
                        Height          =   330
                        Left            =   10365
                        TabIndex        =   54
                        ToolTipText     =   "· ⁄œÌ· «·»Ì«‰«  «·Õ«·Ì…"
                        Top             =   600
                        Width           =   1365
                        _ExtentX        =   2408
                        _ExtentY        =   582
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
                        ButtonImage     =   "FrmContractExam.frx":408A6
                        ColorButton     =   14871017
                        DrawFocusRectangle=   0   'False
                     End
                     Begin ImpulseButton.ISButton ISButton7 
                        Height          =   330
                        Left            =   8940
                        TabIndex        =   55
                        ToolTipText     =   "Õ›Ÿ «·»Ì«‰«  œ«Œ· ﬁ«⁄œ… «·»Ì«‰« "
                        Top             =   600
                        Width           =   1125
                        _ExtentX        =   1984
                        _ExtentY        =   582
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
                        ButtonImage     =   "FrmContractExam.frx":47108
                        ColorButton     =   14871017
                        DrawFocusRectangle=   0   'False
                     End
                     Begin ImpulseButton.ISButton ISButton9 
                        Height          =   330
                        Left            =   7140
                        TabIndex        =   56
                        ToolTipText     =   "·· —«Ã⁄ ⁄‰ «·ÕœÀ Ê«·—ÃÊ⁄ «·Ï «·Ê÷⁄ «·ÿ»Ì⁄Ì"
                        Top             =   600
                        Width           =   1515
                        _ExtentX        =   2672
                        _ExtentY        =   582
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
                        ButtonImage     =   "FrmContractExam.frx":474A2
                        ColorButton     =   14871017
                        DrawFocusRectangle=   0   'False
                     End
                     Begin ImpulseButton.ISButton ISButton10 
                        Height          =   330
                        Left            =   5790
                        TabIndex        =   57
                        ToolTipText     =   "Õ–› «·»Ì«‰«  «·„Õœœ…"
                        Top             =   600
                        Width           =   1215
                        _ExtentX        =   2143
                        _ExtentY        =   582
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
                        ButtonImage     =   "FrmContractExam.frx":4783C
                        ColorButton     =   14871017
                        DrawFocusRectangle=   0   'False
                     End
                     Begin ImpulseButton.ISButton ISButton11 
                        Height          =   420
                        Left            =   3870
                        TabIndex        =   58
                        TabStop         =   0   'False
                        ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
                        Top             =   600
                        Width           =   1185
                        _ExtentX        =   2090
                        _ExtentY        =   741
                        ButtonStyle     =   1
                        ButtonPositionImage=   1
                        Caption         =   "ÿ»«⁄… "
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
                        ButtonImage     =   "FrmContractExam.frx":47DD6
                        ColorButton     =   14871017
                        DrawFocusRectangle=   0   'False
                        DisabledImageStyle=   1
                     End
                     Begin ImpulseButton.ISButton ISButton12 
                        Height          =   330
                        Left            =   120
                        TabIndex        =   59
                        TabStop         =   0   'False
                        ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
                        Top             =   600
                        Visible         =   0   'False
                        Width           =   1095
                        _ExtentX        =   1931
                        _ExtentY        =   582
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
                        ButtonImage     =   "FrmContractExam.frx":4E638
                        ColorButton     =   14871017
                        DrawFocusRectangle=   0   'False
                     End
                     Begin ImpulseButton.ISButton ISButton13 
                        Height          =   330
                        Left            =   1200
                        TabIndex        =   60
                        ToolTipText     =   "«·Œ—ÊÃ «·Ï  «·‰«›–… «·—∆Ì”Ì…"
                        Top             =   600
                        Width           =   915
                        _ExtentX        =   1614
                        _ExtentY        =   582
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
                        ButtonImage     =   "FrmContractExam.frx":4E9D2
                        ColorButton     =   14871017
                        DrawFocusRectangle=   0   'False
                        DisabledImageStyle=   1
                     End
                     Begin MSDataListLib.DataCombo DataCombo1 
                        Height          =   315
                        Left            =   8400
                        TabIndex        =   61
                        Top             =   90
                        Width           =   3465
                        _ExtentX        =   6112
                        _ExtentY        =   556
                        _Version        =   393216
                        Enabled         =   0   'False
                        Text            =   ""
                        RightToLeft     =   -1  'True
                     End
                     Begin VB.Label lbl 
                        Alignment       =   1  'Right Justify
                        BackStyle       =   0  'Transparent
                        Caption         =   "Õ—— »Ê«”ÿ…  "
                        Height          =   345
                        Index           =   0
                        Left            =   12270
                        TabIndex        =   66
                        Top             =   90
                        Width           =   1140
                     End
                     Begin VB.Label Label2 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00E2E9E9&
                        Caption         =   "«·”Ã· «·Õ«·Ì:"
                        Height          =   195
                        Index           =   2
                        Left            =   3255
                        TabIndex        =   65
                        Top             =   240
                        Width           =   1245
                     End
                     Begin VB.Label Label2 
                        Alignment       =   1  'Right Justify
                        BackColor       =   &H00E2E9E9&
                        Caption         =   "⁄œœ «·”Ã·« :"
                        Height          =   195
                        Index           =   3
                        Left            =   1080
                        TabIndex        =   64
                        Top             =   240
                        Width           =   1155
                     End
                     Begin VB.Label Label3 
                        Alignment       =   2  'Center
                        BackColor       =   &H00E2E9E9&
                        ForeColor       =   &H00800000&
                        Height          =   195
                        Left            =   2370
                        TabIndex        =   63
                        Top             =   240
                        Width           =   780
                     End
                     Begin VB.Label Label4 
                        Alignment       =   2  'Center
                        BackColor       =   &H00E2E9E9&
                        ForeColor       =   &H00C00000&
                        Height          =   195
                        Left            =   315
                        TabIndex        =   62
                        Top             =   240
                        Width           =   630
                     End
                  End
                  Begin C1SizerLibCtl.C1Elastic ELe 
                     Height          =   855
                     Index           =   3
                     Left            =   0
                     TabIndex        =   67
                     TabStop         =   0   'False
                     Top             =   -180
                     Width           =   13410
                     _cx             =   23654
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
                     BackColor       =   16777215
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
                     Begin VB.Label Label1 
                        Alignment       =   1  'Right Justify
                        BackStyle       =   0  'Transparent
                        Caption         =   " ‰»ÌÂ«  «·⁄ﬁÊœ «· Ì ” ‰ ÂÌ"
                        BeginProperty Font 
                           Name            =   "Arial"
                           Size            =   14.25
                           Charset         =   178
                           Weight          =   700
                           Underline       =   0   'False
                           Italic          =   0   'False
                           Strikethrough   =   0   'False
                        EndProperty
                        ForeColor       =   &H00000000&
                        Height          =   435
                        Index           =   2
                        Left            =   7080
                        TabIndex        =   74
                        Top             =   270
                        Width           =   4665
                     End
                     Begin VB.Image Image2 
                        Height          =   705
                        Left            =   12330
                        Picture         =   "FrmContractExam.frx":4ED6C
                        Stretch         =   -1  'True
                        Top             =   0
                        Width           =   840
                     End
                  End
                  Begin C1SizerLibCtl.C1Elastic C1Elastic5 
                     Height          =   8490
                     Left            =   570
                     TabIndex        =   68
                     TabStop         =   0   'False
                     Top             =   870
                     Width           =   13425
                     _cx             =   23680
                     _cy             =   14975
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
                     Frame           =   0
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
                     Begin ImpulseButton.ISButton ISButton18 
                        Height          =   315
                        Left            =   240
                        TabIndex        =   69
                        ToolTipText     =   "«÷«›… «·»Ì«‰«  «·Ï «·œ« «"
                        Top             =   8055
                        Width           =   2925
                        _ExtentX        =   5159
                        _ExtentY        =   556
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
                        ButtonImage     =   "FrmContractExam.frx":50171
                        ColorButton     =   14871017
                        DrawFocusRectangle=   0   'False
                        DisabledImageExtraction=   0
                        LowerToggledContent=   0   'False
                     End
                     Begin ImpulseButton.ISButton ISButton19 
                        Height          =   315
                        Left            =   3960
                        TabIndex        =   70
                        ToolTipText     =   "«÷«›… «·»Ì«‰«  «·Ï «·œ« «"
                        Top             =   8055
                        Width           =   4125
                        _ExtentX        =   7276
                        _ExtentY        =   556
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
                        ButtonImage     =   "FrmContractExam.frx":569D3
                        ColorButton     =   14871017
                        DrawFocusRectangle=   0   'False
                        DisabledImageExtraction=   0
                        LowerToggledContent=   0   'False
                     End
                     Begin VSFlex8Ctl.VSFlexGrid Fg2 
                        Height          =   7500
                        Left            =   -240
                        TabIndex        =   71
                        Top             =   360
                        Width           =   13410
                        _cx             =   23654
                        _cy             =   13229
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
                        BackColor       =   16777215
                        ForeColor       =   -2147483640
                        BackColorFixed  =   14871017
                        ForeColorFixed  =   -2147483630
                        BackColorSel    =   -2147483635
                        ForeColorSel    =   -2147483634
                        BackColorBkg    =   -2147483636
                        BackColorAlternate=   16776960
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
                        Cols            =   13
                        FixedRows       =   1
                        FixedCols       =   1
                        RowHeightMin    =   320
                        RowHeightMax    =   0
                        ColWidthMin     =   0
                        ColWidthMax     =   0
                        ExtendLastCol   =   0   'False
                        FormatString    =   $"FrmContractExam.frx":5D235
                        ScrollTrack     =   0   'False
                        ScrollBars      =   3
                        ScrollTips      =   0   'False
                        MergeCells      =   0
                        MergeCompare    =   0
                        AutoResize      =   0   'False
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
                     Begin MSComCtl2.DTPicker txtOrderDate 
                        Height          =   330
                        Index           =   0
                        Left            =   4170
                        TabIndex        =   76
                        Top             =   30
                        Width           =   1455
                        _ExtentX        =   2566
                        _ExtentY        =   582
                        _Version        =   393216
                        CheckBox        =   -1  'True
                        Format          =   113967107
                        CurrentDate     =   38887
                     End
                     Begin MSComCtl2.DTPicker txtOrderDate 
                        Height          =   330
                        Index           =   1
                        Left            =   960
                        TabIndex        =   77
                        Top             =   30
                        Width           =   1455
                        _ExtentX        =   2566
                        _ExtentY        =   582
                        _Version        =   393216
                        CheckBox        =   -1  'True
                        Format          =   113967107
                        CurrentDate     =   38887
                     End
                     Begin VB.Label Label5 
                        Caption         =   "«·Ï  «—ÌŒ"
                        Height          =   285
                        Index           =   1
                        Left            =   2610
                        TabIndex        =   75
                        Top             =   60
                        Width           =   1290
                     End
                     Begin VB.Label Label5 
                        Caption         =   "„‰  «—ÌŒ"
                        Height          =   285
                        Index           =   0
                        Left            =   5700
                        TabIndex        =   72
                        Top             =   60
                        Width           =   1290
                     End
                  End
               End
            End
         End
      End
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "«·„” Œœ„"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   13
      Left            =   15480
      TabIndex        =   8
      Top             =   360
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "FrmContractExam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit
 Dim RsSavRec As ADODB.Recordset
 Dim StrSQL As String
 Dim RsDevsub As ADODB.Recordset
 Dim BKGrndPic As ClsBackGroundPic
 Dim RecID As String
 Dim II As Long
Public mIndex As Integer
Sub filgrid1()
Dim Sql As String
Dim Rs2 As ADODB.Recordset
Set Rs2 = New ADODB.Recordset
Dim Period As Integer
Dim StrPeriod As String
Dim i As Integer
Dim k As Integer
Sql = " SELECT     dbo.Contract.Contract_ID, dbo.Contract.Emp_id, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee, dbo.TblEmployee.BranchId, "
Sql = Sql & "                       dbo.Contract.Contract_date, dbo.Contract.DateH, dbo.Contract.Contract_Enddate, dbo.Contract.DateH1, dbo.Contract.StutsID, dbo.Contract.test_period_no,"
Sql = Sql & "                      dbo.Contract.test_period"
Sql = Sql & " FROM         dbo.Contract LEFT OUTER JOIN"
Sql = Sql & "                      dbo.TblEmployee ON dbo.Contract.Emp_id = dbo.TblEmployee.Emp_ID"
Sql = Sql & " WHERE   (  (dbo.Contract.StutsID IS NULL) OR"
Sql = Sql & "                      (dbo.Contract.StutsID = 0) )"
Period = -1
   Fg1.Clear flexClearScrollable, flexClearEverything
   Fg1.Rows = 1
   StrPeriod = ""
Rs2.Open Sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs2.RecordCount > 0 Then
With Fg1
k = .Rows
.Rows = .Rows + Rs2.RecordCount
Rs2.MoveFirst
For i = k To .Rows - 1
.TextMatrix(i, .ColIndex("DateH")) = IIf(IsNull(Rs2("DateH").value), "", Rs2("DateH").value)
.TextMatrix(i, .ColIndex("test_period_no")) = IIf(IsNull(Rs2("test_period_no").value), 0, Rs2("test_period_no").value)
.TextMatrix(i, .ColIndex("Contract_date")) = IIf(IsNull(Rs2("Contract_date").value), "", Rs2("Contract_date").value)
If Not IsNull(Rs2("test_period").value) Then
If (Rs2("test_period").value) = 1 Then
Period = 1
If SystemOptions.UserInterface = ArabicInterface Then
StrPeriod = "”‰…"
Else
StrPeriod = "Year"
End If
DTPicker1.value = DateAdd("YYYY", val(.TextMatrix(i, .ColIndex("test_period_no"))), .TextMatrix(i, .ColIndex("Contract_date")))
Else
Period = 0
If SystemOptions.UserInterface = ArabicInterface Then
StrPeriod = "‘Â—"
Else
StrPeriod = "Month"
End If
DTPicker1.value = DateAdd("m", val(.TextMatrix(i, .ColIndex("test_period_no"))), .TextMatrix(i, .ColIndex("Contract_date")))
End If
Else
Period = -1
End If
.TextMatrix(i, .ColIndex("Serial")) = i
.TextMatrix(i, .ColIndex("test_period_no")) = .TextMatrix(i, .ColIndex("test_period_no")) & " " & StrPeriod
.TextMatrix(i, .ColIndex("Contract_ID")) = IIf(IsNull(Rs2("Contract_ID").value), 0, Rs2("Contract_ID").value)
.TextMatrix(i, .ColIndex("Emp_id")) = IIf(IsNull(Rs2("Emp_id").value), 0, Rs2("Emp_id").value)
.TextMatrix(i, .ColIndex("Fullcode")) = IIf(IsNull(Rs2("Fullcode").value), "", Rs2("Fullcode").value)
.TextMatrix(i, .ColIndex("Contract_Enddate")) = DTPicker1.value
.TextMatrix(i, .ColIndex("DateH1")) = ToHijriDate(DTPicker1.value)
If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(Rs2("Emp_Name").value), "", Rs2("Emp_Name").value)
Else
.TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(Rs2("Emp_Namee").value), "", Rs2("Emp_Namee").value)
End If
Rs2.MoveNext
Next i
End With
End If
End Sub



Sub filgrid2()
Dim Sql As String
Dim Rs2 As ADODB.Recordset
Set Rs2 = New ADODB.Recordset
Dim Period As Integer
Dim StrPeriod As String
Dim i As Integer
Dim k As Integer

Sql = " SELECT *,DateRemain = DATEDIFF(D,GETDATE(),Enddate)"
Sql = Sql & " FROM   ("
Sql = Sql & "            SELECT dbo.Contract.Contract_ID,"

Sql = Sql & "                 Contract_periodName ="
Sql = Sql & "                              Case Contract_period"
Sql = Sql & "                                             WHEN 1 THEN '”‰…'"
Sql = Sql & "                                             WHEN 0 THEN '‘Â—'"
Sql = Sql & "                                             END,"
Sql = Sql & "                   Enddate = CASE Contract_period"
Sql = Sql & "                                  WHEN 1 THEN DATEADD(YY, Contract_period_no, Contract_date)"
Sql = Sql & "                                  WHEN 0 THEN DATEADD(M, Contract_period_no, Contract_date)"
Sql = Sql & "                             END,"
Sql = Sql & "                   dbo.Contract.Contract_date,"
Sql = Sql & "                   dbo.Contract.Contract_period_no,"
Sql = Sql & "                   CONTRACT.Contract_period,"
Sql = Sql & "                   dbo.Contract.Emp_id,"
Sql = Sql & "                   test_period,"
Sql = Sql & "                   dbo.TblEmployee.Emp_Name,"
Sql = Sql & "                   dbo.TblEmployee.Fullcode,"
Sql = Sql & "                   dbo.TblEmployee.Emp_Namee,"
Sql = Sql & "                   dbo.TblEmployee.BranchId,"
Sql = Sql & "                   dbo.Contract.DateH,"
Sql = Sql & "                   dbo.Contract.Contract_Enddate,"
Sql = Sql & "                   dbo.Contract.DateH1,"
Sql = Sql & "                   dbo.Contract.StutsID,"
Sql = Sql & "                   dbo.Contract.test_period_no"
Sql = Sql & "            From dbo.Contract"
Sql = Sql & "                   LEFT OUTER JOIN dbo.TblEmployee"
Sql = Sql & "                        ON  dbo.Contract.Emp_id = dbo.TblEmployee.Emp_ID"
Sql = Sql & "        ) AS TT"


Sql = Sql & " Where  1 = 1 "
Dim mDate As Boolean
If Not IsNull(Me.txtOrderDate(0).value) Then
    mDate = True
    Sql = Sql & " and TT.EndDate   >=" & SQLDate(Me.txtOrderDate(0).value, True) & ""
End If
If Not IsNull(Me.txtOrderDate(1).value) Then
    mDate = True
    Sql = Sql & " and  TT.EndDate   <=" & SQLDate(Me.txtOrderDate(1).value, True) & ""
End If
If Not mDate Then
    Sql = Sql & " and month(TT.EndDate) = " & month(Date) & " And year(TT.EndDate) = " & year(Date)
End If


'(  (dbo.Contract.StutsID IS NULL) OR"
'Sql = Sql & "                      (dbo.Contract.StutsID = 0) )"

'txtOrderDate
Period = -1
   Fg2.Clear flexClearScrollable, flexClearEverything
   Fg2.Rows = 1
   StrPeriod = ""
Rs2.Open Sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs2.RecordCount > 0 Then
With Fg2
k = .Rows
.Rows = .Rows + Rs2.RecordCount
Rs2.MoveFirst
For i = k To .Rows - 1
.TextMatrix(i, .ColIndex("DateH")) = IIf(IsNull(Rs2("DateH").value), "", Rs2("DateH").value)
.TextMatrix(i, .ColIndex("Contract_period_no")) = IIf(IsNull(Rs2("Contract_period_no").value), 0, Rs2("Contract_period_no").value)
.TextMatrix(i, .ColIndex("Contract_date")) = IIf(IsNull(Rs2("Contract_date").value), "", Rs2("Contract_date").value)
If Not IsNull(Rs2("test_period").value) Then
If (Rs2("Contract_period").value) = 1 Then
Period = 1
If SystemOptions.UserInterface = ArabicInterface Then
StrPeriod = "”‰…"
Else
StrPeriod = "Year"
End If
'DTPicker1.value = DateAdd("YYYY", val(.TextMatrix(i, .ColIndex("test_period_no"))), .TextMatrix(i, .ColIndex("Contract_date")))
Else
Period = 0
If SystemOptions.UserInterface = ArabicInterface Then
StrPeriod = "‘Â—"
Else
StrPeriod = "Month"
End If
'DTPicker1.value = DateAdd("m", val(.TextMatrix(i, .ColIndex("test_period_no"))), .TextMatrix(i, .ColIndex("Contract_date")))
End If
Else
Period = -1
End If
.TextMatrix(i, .ColIndex("Serial")) = i
'.TextMatrix(i, .ColIndex("test_period_no")) = .TextMatrix(i, .ColIndex("test_period_no")) & " " & StrPeriod
.TextMatrix(i, .ColIndex("test_period")) = StrPeriod '.TextMatrix(i, .ColIndex("test_period")) & " " & StrPeriod
.TextMatrix(i, .ColIndex("Contract_ID")) = IIf(IsNull(Rs2("Contract_ID").value), 0, Rs2("Contract_ID").value)
.TextMatrix(i, .ColIndex("DateRemain")) = IIf(IsNull(Rs2("DateRemain").value), 0, Rs2("DateRemain").value)

.TextMatrix(i, .ColIndex("Emp_id")) = IIf(IsNull(Rs2("Emp_id").value), 0, Rs2("Emp_id").value)
.TextMatrix(i, .ColIndex("Fullcode")) = IIf(IsNull(Rs2("Fullcode").value), "", Rs2("Fullcode").value)
.TextMatrix(i, .ColIndex("Contract_Enddate")) = IIf(IsNull(Rs2("Enddate").value), "", Rs2("Enddate").value)
.TextMatrix(i, .ColIndex("DateH1")) = ToHijriDate(.TextMatrix(i, .ColIndex("Contract_Enddate")))
If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(Rs2("Emp_Name").value), "", Rs2("Emp_Name").value)
Else
.TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(Rs2("Emp_Namee").value), "", Rs2("Emp_Namee").value)
End If
Rs2.MoveNext
Next i
End With
End If
End Sub
Function print_report2(Optional NoteSerial As String)
    
     
    Dim Sql As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String

Sql = " SELECT *,DateRemain = DATEDIFF(D,GETDATE(),Enddate)"
Sql = Sql & " FROM   ("
Sql = Sql & "            SELECT dbo.Contract.Contract_ID,"

Sql = Sql & "                 Contract_periodName ="
Sql = Sql & "                              Case Contract_period"
Sql = Sql & "                                             WHEN 1 THEN '”‰…'"
Sql = Sql & "                                             WHEN 0 THEN '‘Â—'"
Sql = Sql & "                                             END,"
Sql = Sql & "                   Enddate = CASE Contract_period"
Sql = Sql & "                                  WHEN 1 THEN DATEADD(YY, Contract_period_no, Contract_date)"
Sql = Sql & "                                  WHEN 0 THEN DATEADD(M, Contract_period_no, Contract_date)"
Sql = Sql & "                             END,"
Sql = Sql & "                   dbo.Contract.Contract_date,"
Sql = Sql & "                   dbo.Contract.Contract_period_no,"
Sql = Sql & "                   CONTRACT.Contract_period,"
Sql = Sql & "                   dbo.Contract.Emp_id,"
Sql = Sql & "                   test_period,"
Sql = Sql & "                   dbo.TblEmployee.Emp_Name,"
Sql = Sql & "                   dbo.TblEmployee.Fullcode,"
Sql = Sql & "                   dbo.TblEmployee.Emp_Namee,"
Sql = Sql & "                   dbo.TblEmployee.BranchId,"
Sql = Sql & "                   dbo.Contract.DateH,"
Sql = Sql & "                   dbo.Contract.Contract_Enddate,"
Sql = Sql & "                   dbo.Contract.DateH1,"
Sql = Sql & "                   dbo.Contract.StutsID,"
Sql = Sql & "                   dbo.Contract.test_period_no"
Sql = Sql & "            From dbo.Contract"
Sql = Sql & "                   LEFT OUTER JOIN dbo.TblEmployee"
Sql = Sql & "                        ON  dbo.Contract.Emp_id = dbo.TblEmployee.Emp_ID"
Sql = Sql & "        ) AS TT"

Sql = Sql & " Where  1 = 1 "
Dim mDate As Boolean
If Not IsNull(Me.txtOrderDate(0).value) Then
    mDate = True
    Sql = Sql & " and TT.EndDate   >=" & SQLDate(Me.txtOrderDate(0).value, True) & ""
End If
If Not IsNull(Me.txtOrderDate(1).value) Then
    mDate = True
    Sql = Sql & "  and TT.EndDate   <=" & SQLDate(Me.txtOrderDate(1).value, True) & ""
End If
If Not mDate Then
    Sql = Sql & " and month(TT.EndDate) = " & month(Date) & " And year(TT.EndDate) = " & year(Date)
End If

 If SystemOptions.UserInterface = ArabicInterface Then
          StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "EndContractEmp.rpt"
     Else
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "EndContractEmp.rpt"
       End If
    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open Sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        'GetMsgs 138, vbExclamation
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        Else
            Msg = "There's no record to show"
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
      '  xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""
        'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        '    StrReportTitle = StrReportTitle + " From Date " & (Me.DTPickerAccFrom.value) & ""
        'End If
        'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        '    StrReportTitle = StrReportTitle + " To Date :  " & (Me.DTPickerAccTo.value) & ""
        'End If
    End If
    If Not mDate Then
        xReport.ParameterFields(10).AddCurrentValue CStr(Date)
        xReport.ParameterFields(11).AddCurrentValue CStr(Date)
    Else
        xReport.ParameterFields(10).AddCurrentValue CStr(txtOrderDate(0).value)
        xReport.ParameterFields(11).AddCurrentValue CStr(txtOrderDate(1).value)
    
    End If

   ' xReport.ParameterFields(3).AddCurrentValue user_name
       ' xReport.ParameterFields(4).AddCurrentValue WriteNo(Format(val(TxtAdvanceValue.text), "0.00"), 0, True, ".")
'        xReport.ParameterFields(6).AddCurrentValue val(lbl(23).Caption)
      '   xReport.ParameterFields(7).AddCurrentValue DBIssueDate.value
   ' xReport.ParameterFields(8).AddCurrentValue IIf(IsNumeric(Fg.TextMatrix(Me.Fg.FixedRows, Fg.ColIndex("PartValue"))), val(Fg.TextMatrix(Me.Fg.FixedRows, Fg.ColIndex("PartValue"))), 0)
' xReport.ParameterFields(9).AddCurrentValue val(lbl(22).Caption)
'  xReport.ParameterFields(10).AddCurrentValue val(TxtDiscount.text)
 '  xReport.ParameterFields(11).AddCurrentValue txtDiscountDES.text
   
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


Private Sub Fg1_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With Fg1
Select Case .ColKey(Col)
Case "StatusID"
If .Cell(flexcpChecked, Row, .ColIndex("Slect")) = flexChecked Then
.ComboList = ""
Else
Cancel = True
End If
End Select
End With
End Sub

 Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim conection As String
    Dim My_SQL As String
  
    With Fg1
     If SystemOptions.UserInterface = ArabicInterface Then
        .ColComboList(.ColIndex("StatusID")) = "#1; ⁄ÌÌ‰ |#2;«‰Â«¡ "
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
           .ColComboList(.ColIndex("StatusID")) = "#1;Hired |#2;No Hired "
            End If
    End With
    txtOrderDate(0).value = "01-" & month(Date) & "-" & year(Date)
    txtOrderDate(1).value = DateAdd("m", 1, txtOrderDate(0).value) - 1
    
    Resize_Form Me
    Dim Dcombos As New ClsDataCombos
   Dcombos.GetUsers Me.DCboUserName
   
If mIndex = 0 Then
TabMain.TabVisible(1) = False
TabMain.TabVisible(0) = True
TabMain.CurrTab = 0
ElseIf mIndex = 1 Then
TabMain.TabVisible(1) = True
TabMain.TabVisible(0) = False
TabMain.CurrTab = 1
filgrid2
End If
   filgrid1
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
        SwitchKeyboardLang LANG_ENGLISH
        Else
        SwitchKeyboardLang LANG_ARABIC
    End If
   Me.Refresh
ErrTrap:
End Sub

Private Sub ISButton18_Click()
print_report2
End Sub

Private Sub ISButton19_Click()
filgrid2
End Sub

Private Sub ISButton2_Click()
filgrid1
End Sub


  ' cancel camnd sub
  '+++++++++++++++++++++++++++++++
  Private Sub BtnCancel_Click()
    Unload Me
End Sub

Private Sub Form_Terminate()
     ' Set FrmVacancy = Nothing
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap
    If RsSavRec.State = adStateOpen Then
        If Not (RsSavRec.EOF Or RsSavRec.BOF) Then
            If RsSavRec.EditMode <> adEditNone Then
                RsSavRec.CancelUpdate
            End If
        End If
        RsSavRec.Close
        Set RsSavRec = Nothing
    End If
ErrTrap:
End Sub
Private Sub Form_Activate()
    Me.ZOrder 0
End Sub

Private Sub ChangeLang()
On Error GoTo ErrTrap
   ' form name
  Label1(2).Caption = "Alarms Of Duration Of The Test  "
 ISButton2.Caption = "Update"
 ISButton3.Caption = "Excute"

  With Fg1
  .TextMatrix(0, .ColIndex("Serial")) = "Serial"
  .TextMatrix(0, .ColIndex("Slect")) = "Select"
  .TextMatrix(0, .ColIndex("Fullcode")) = "Employee Code"
  .TextMatrix(0, .ColIndex("Emp_Name")) = " Employee Name"
  .TextMatrix(0, .ColIndex("Contract_date")) = "Contract Date"
  .TextMatrix(0, .ColIndex("DateH")) = "Contract Date"
  .TextMatrix(0, .ColIndex("test_period_no")) = "Period"
  .TextMatrix(0, .ColIndex("Contract_Enddate")) = "End Period Date"
  .TextMatrix(0, .ColIndex("DateH1")) = "End Period Date"
  .TextMatrix(0, .ColIndex("StatusID")) = "Procedure"

  End With
ErrTrap:
End Sub

Private Sub ISButton3_Click()
Dim i As Integer
With Fg1
For i = 1 To .Rows - 1
If .Cell(flexcpChecked, i, .ColIndex("Slect")) = flexChecked Then
Cn.Execute "Update Contract set StutsID=" & val(.TextMatrix(i, .ColIndex("StatusID"))) & " where Contract_ID=" & val(.TextMatrix(i, .ColIndex("Contract_ID"))) & " "
End If
Next i
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox " „ «· ÕœÌÀ »‰Ã«Õ"
Else
MsgBox "Update Successfully"
End If
End With
filgrid1
End Sub
