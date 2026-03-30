VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmTransacRegAlarm 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8865
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   17580
   Icon            =   "FrmTransacRegAlarm.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   17580
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   8865
      Left            =   0
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   0
      Width           =   17595
      _cx             =   31036
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
      Begin C1SizerLibCtl.C1Tab C1Tab1 
         Height          =   7575
         Left            =   0
         TabIndex        =   14
         Top             =   840
         Width           =   17655
         _cx             =   31141
         _cy             =   13361
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
         Caption         =   "„” ‰œ«  „—”·…|„” ‰œ«  „ √Œ—…"
         Align           =   0
         CurrTab         =   1
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
         Begin C1SizerLibCtl.C1Elastic C1Elastic2 
            Height          =   7200
            Left            =   -18210
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   330
            Width           =   17565
            _cx             =   30983
            _cy             =   12700
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
            Begin VSFlex8Ctl.VSFlexGrid Fg 
               Height          =   7155
               Left            =   0
               TabIndex        =   16
               Top             =   0
               Width           =   17475
               _cx             =   30824
               _cy             =   12621
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
               BackColorAlternate=   16777088
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
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmTransacRegAlarm.frx":6852
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
               Begin MSComctlLib.ProgressBar ProgressBar1 
                  Height          =   615
                  Left            =   1200
                  TabIndex        =   17
                  Top             =   960
                  Visible         =   0   'False
                  Width           =   11295
                  _ExtentX        =   19923
                  _ExtentY        =   1085
                  _Version        =   393216
                  Appearance      =   0
               End
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic3 
            Height          =   7200
            Left            =   45
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   330
            Width           =   17565
            _cx             =   30983
            _cy             =   12700
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
            Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid1 
               Height          =   7155
               Left            =   0
               TabIndex        =   19
               Top             =   0
               Width           =   17475
               _cx             =   30824
               _cy             =   12621
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
               BackColorAlternate=   16777088
               GridColor       =   128
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
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmTransacRegAlarm.frx":6A0A
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
               Begin MSComctlLib.ProgressBar ProgressBar2 
                  Height          =   615
                  Left            =   1200
                  TabIndex        =   20
                  Top             =   960
                  Visible         =   0   'False
                  Width           =   11295
                  _ExtentX        =   19923
                  _ExtentY        =   1085
                  _Version        =   393216
                  Appearance      =   0
               End
            End
         End
      End
      Begin VB.Frame FraHeader 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   780
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   17625
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "⁄—÷ «·„” ‰œ« "
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
            Height          =   375
            Index           =   2
            Left            =   8880
            TabIndex        =   13
            Top             =   240
            Width           =   4080
         End
      End
      Begin ImpulseButton.ISButton CmdExit 
         Cancel          =   -1  'True
         Height          =   360
         Left            =   120
         TabIndex        =   21
         Top             =   8400
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   635
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
         ButtonImage     =   "FrmTransacRegAlarm.frx":6BE5
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
   End
   Begin VB.TextBox TxtModFlg 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      Enabled         =   0   'False
      Height          =   285
      Left            =   20760
      TabIndex        =   4
      Text            =   "modflag"
      Top             =   4200
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.ComboBox CmbType 
      BackColor       =   &H80000018&
      Height          =   315
      ItemData        =   "FrmTransacRegAlarm.frx":6F7F
      Left            =   20640
      List            =   "FrmTransacRegAlarm.frx":6F8F
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
      Left            =   20760
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
      Left            =   20760
      TabIndex        =   1
      Top             =   1680
      Visible         =   0   'False
      Width           =   3105
   End
   Begin VB.TextBox Emp_id 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   20400
      TabIndex        =   0
      Top             =   5640
      Visible         =   0   'False
      Width           =   1575
   End
   Begin MSDataListLib.DataCombo DCUser 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   21000
      TabIndex        =   5
      Tag             =   "„‰ ›÷·ﬂ √œŒ· —ﬁ„ «·ﬁ÷Ì…"
      Top             =   960
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
      Left            =   20640
      TabIndex        =   6
      Top             =   2280
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSComctlLib.ImageList GrdImageList 
      Left            =   20760
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
            Picture         =   "FrmTransacRegAlarm.frx":6FA8
            Key             =   "CompanyName"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransacRegAlarm.frx":7342
            Key             =   "Ser"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransacRegAlarm.frx":76DC
            Key             =   "Vac_Name"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransacRegAlarm.frx":7A76
            Key             =   "ShareCount"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransacRegAlarm.frx":7E10
            Key             =   "Dis_Count"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransacRegAlarm.frx":81AA
            Key             =   "Bouns"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransacRegAlarm.frx":8544
            Key             =   "SharesValue"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmTransacRegAlarm.frx":8ADE
            Key             =   "BuyValue"
         EndProperty
      EndProperty
   End
   Begin ImpulseButton.ISButton BtnUpdate 
      Height          =   330
      Left            =   20760
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
      ButtonImage     =   "FrmTransacRegAlarm.frx":8E78
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
      ButtonImage     =   "FrmTransacRegAlarm.frx":F6DA
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
      DisabledImageStyle=   1
   End
   Begin ImpulseButton.ISButton btnQuery 
      Height          =   330
      Left            =   22080
      TabIndex        =   10
      TabStop         =   0   'False
      ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
      Top             =   120
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
      ButtonImage     =   "FrmTransacRegAlarm.frx":15F3C
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
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
      Left            =   20640
      TabIndex        =   8
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "FrmTransacRegAlarm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub Fg_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
With Fg
Select Case .ColKey(Col)
Case "show"
If val(.TextMatrix(Row, .ColIndex("ID"))) <> 0 Then
FrmTransacRegistr.Ind = 1
FrmTransacRegistr.show
FrmTransacRegistr.Ind = 1
FrmTransacRegistr.Retrive2 val(.TextMatrix(Row, .ColIndex("ID")))
End If
End Select
End With
End Sub

Private Sub Fg_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With Me.Fg
Select Case .ColKey(Col)
 Case "show"
            .ColComboList(.ColIndex("show")) = "..."
     End Select
    End With
End Sub

Private Sub Form_Load()
    On Error GoTo ErrTrap
        If SystemOptions.UserInterface = ArabicInterface Then
                Fg.ColComboList(Fg.ColIndex("ImportExport")) = "#1;  ’«œ—|#2; Ê«—œ"
                VSFlexGrid1.ColComboList(VSFlexGrid1.ColIndex("ImportExport")) = "#1;  ’«œ—|#2; Ê«—œ"
                ElseIf SystemOptions.UserInterface = EnglishInterface Then
               Fg.ColComboList(Fg.ColIndex("ImportExport")) = "#1;Import  |#2;Export "
               VSFlexGrid1.ColComboList(VSFlexGrid1.ColIndex("ImportExport")) = "#1;Import  |#2;Export "
           End If
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
    FillGrid
    FillGrid2
   Me.Refresh
ErrTrap:
End Sub

Private Sub Fg_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With Me.Fg
If .ColKey(Col) <> "show" Then
Cancel = True
End If
End With
End Sub
Private Sub ChangeLang()
On Error GoTo ErrTrap
   ' form name
   ''''''''''''''''''''///
      Label1(2).Caption = "View Documents"
      CmdExit.Caption = "Exit"
      
    'Me.TabMain.TabCaption(2) = "Approval Status"
  ' Label11.Caption = "Approval Requested By"
    C1Tab1.Caption = "Sent Documents|Delayed Documents"

    With VSFlexGrid1
        .TextMatrix(0, .ColIndex("Ser")) = "Serial"
        .TextMatrix(0, .ColIndex("TransRegID")) = "No#"
        .TextMatrix(0, .ColIndex("RecordDate")) = "Date"
        .TextMatrix(0, .ColIndex("RecordDateH")) = "Date Hegira "
        .TextMatrix(0, .ColIndex("RecDate")) = "Sending Date"
        .TextMatrix(0, .ColIndex("ImportExport")) = "Type"
        .TextMatrix(0, .ColIndex("Name")) = "Process Type"
        .TextMatrix(0, .ColIndex("UserName")) = "Sent From"
        .TextMatrix(0, .ColIndex("Summary")) = "Summary"
        .TextMatrix(0, .ColIndex("ExitDate")) = "Date of Termination"
        .TextMatrix(0, .ColIndex("late")) = "Delay"
        .TextMatrix(0, .ColIndex("show")) = "Show"
    End With
        With Fg
        .TextMatrix(0, .ColIndex("Ser")) = "Serial"
        .TextMatrix(0, .ColIndex("TransRegID")) = "No#"
        .TextMatrix(0, .ColIndex("RecordDate")) = "Date"
        .TextMatrix(0, .ColIndex("RecordDateH")) = "Date Hegira "
        .TextMatrix(0, .ColIndex("RecDate")) = "Sending Date"
        .TextMatrix(0, .ColIndex("ImportExport")) = "Type"
        .TextMatrix(0, .ColIndex("Name")) = "Process Type"
        .TextMatrix(0, .ColIndex("UserName")) = "Sent From"
        .TextMatrix(0, .ColIndex("Summary")) = "Summary"
        .TextMatrix(0, .ColIndex("ExitDate")) = "Date of Termination"
       ' .TextMatrix(0, .ColIndex("late")) = "Delay"
        .TextMatrix(0, .ColIndex("show")) = "Show"
    End With
ErrTrap:
End Sub
Sub FillGrid()

Dim Sql As String
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
Dim i As Integer
Sql = " SELECT     dbo.TblTransacRegistrDet.ID, dbo.TblTransacRegistrDet.TransRegID, dbo.TblTransacRegistr.RecordDate, dbo.TblTransacRegistr.RecordDateH, "
Sql = Sql & "                       dbo.TblTransacRegistr.RecordTime, dbo.TblTransacRegistr.BrnchID, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee,"
Sql = Sql & "                       dbo.TblTransacRegistr.barcode, dbo.TblTransacRegistr.ImportExport, dbo.TblTransacRegistr.NoImpExp, dbo.TblTransacRegistr.ImpExpDate,"
Sql = Sql & "                       dbo.TblTransacRegistr.ImpExpDateH, dbo.TblTransacRegistr.Summary, dbo.TblTransacRegistr.EnterDate, dbo.TblTransacRegistr.Remarks,"
Sql = Sql & "                       dbo.TblTransacRegistr.MHD, dbo.TblTransacRegistr.MHDID, dbo.TblTransacRegistr.ExitDate, dbo.TblTransacRegistr.TypTrans, dbo.TblXXArchDocType.Name,"
Sql = Sql & "                       dbo.TblXXArchDocType.Namee, dbo.TblXXArchDocType.Code, dbo.TblTransacRegistrDet.FromUser, dbo.TblUsers.UserName, dbo.TblTransacRegistrDet.ToUser,"
Sql = Sql & "                       TblUsers_1.UserName AS ToUserName, dbo.TblTransacRegistrDet.FlgTrans, dbo.TblTransacRegistrDet.RecDate, dbo.TblTransacRegistrDet.ProcedureReq"
Sql = Sql & "  FROM         dbo.TblUsers TblUsers_1 RIGHT OUTER JOIN"
Sql = Sql & "                       dbo.TblTransacRegistrDet ON TblUsers_1.UserID = dbo.TblTransacRegistrDet.ToUser LEFT OUTER JOIN"
Sql = Sql & "                       dbo.TblUsers ON dbo.TblTransacRegistrDet.FromUser = dbo.TblUsers.UserID LEFT OUTER JOIN"
Sql = Sql & "                       dbo.TblXXArchDocType RIGHT OUTER JOIN"
Sql = Sql & "                       dbo.TblTransacRegistr ON dbo.TblXXArchDocType.ID = dbo.TblTransacRegistr.TypTrans LEFT OUTER JOIN"
Sql = Sql & "                       dbo.TblBranchesData ON dbo.TblTransacRegistr.BrnchID = dbo.TblBranchesData.branch_id ON"
Sql = Sql & "                       dbo.TblTransacRegistrDet.TransRegID = dbo.TblTransacRegistr.ID"
Sql = Sql & "  where  dbo.TblTransacRegistrDet.FlgTrans is null"
Sql = Sql & "  and DATEDIFF(n,GETDATE(),dbo.TblTransacRegistr.ExitDate )>=0"
Sql = Sql & "  AND      dbo.TblTransacRegistr.BrnchID in(" & Current_branchSql & ")"

Rs3.Open Sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
 Fg.Clear flexClearScrollable, flexClearEverything
           Fg.Rows = 1
If Rs3.RecordCount > 0 Then
Rs3.MoveFirst
With Fg
.Rows = .Rows + Rs3.RecordCount
For i = 1 To .Rows - 1
.TextMatrix(i, .ColIndex("Ser")) = i
.TextMatrix(i, .ColIndex("ExitDate")) = IIf(IsNull(Rs3("ExitDate").value), "", Rs3("ExitDate").value)
.TextMatrix(i, .ColIndex("ID")) = IIf(IsNull(Rs3("ID").value), 0, Rs3("ID").value)
.TextMatrix(i, .ColIndex("TransRegID")) = IIf(IsNull(Rs3("TransRegID").value), "", Rs3("TransRegID").value)
.TextMatrix(i, .ColIndex("RecordDate")) = IIf(IsNull(Rs3("RecordDate").value), "", Rs3("RecordDate").value)
.TextMatrix(i, .ColIndex("RecordDateH")) = IIf(IsNull(Rs3("RecordDateH").value), "", Rs3("RecordDateH").value)
.TextMatrix(i, .ColIndex("ImportExport")) = IIf(IsNull(Rs3("ImportExport").value), -1, Rs3("ImportExport").value) + 1
.TextMatrix(i, .ColIndex("Summary")) = IIf(IsNull(Rs3("Summary").value), "", Rs3("Summary").value)
.TextMatrix(i, .ColIndex("UserName")) = IIf(IsNull(Rs3("UserName").value), "", Rs3("UserName").value)
.TextMatrix(i, .ColIndex("RecDate")) = IIf(IsNull(Rs3("RecDate").value), "", Rs3("RecDate").value)
If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(Rs3("Name").value), "", Rs3("Name").value)
Else
.TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(Rs3("Namee").value), "", Rs3("Namee").value)
End If
Rs3.MoveNext
Next i
End With
End If
End Sub
Sub FillGrid2()

Dim Sql As String
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
Dim i As Integer
Sql = " SELECT     dbo.TblTransacRegistrDet.ID, dbo.TblTransacRegistrDet.TransRegID, dbo.TblTransacRegistr.RecordDate, dbo.TblTransacRegistr.RecordDateH, "
Sql = Sql & "                       dbo.TblTransacRegistr.RecordTime, dbo.TblTransacRegistr.BrnchID, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee,"
Sql = Sql & "                       dbo.TblTransacRegistr.barcode, dbo.TblTransacRegistr.ImportExport, dbo.TblTransacRegistr.NoImpExp, dbo.TblTransacRegistr.ImpExpDate,"
Sql = Sql & "                       dbo.TblTransacRegistr.ImpExpDateH, dbo.TblTransacRegistr.Summary, dbo.TblTransacRegistr.EnterDate, dbo.TblTransacRegistr.Remarks,"
Sql = Sql & "                       dbo.TblTransacRegistr.MHD, dbo.TblTransacRegistr.MHDID, dbo.TblTransacRegistr.ExitDate, dbo.TblTransacRegistr.TypTrans, dbo.TblXXArchDocType.Name,"
Sql = Sql & "                       dbo.TblXXArchDocType.Namee, dbo.TblXXArchDocType.Code, dbo.TblTransacRegistrDet.FromUser, dbo.TblUsers.UserName, dbo.TblTransacRegistrDet.ToUser,"
Sql = Sql & "                       TblUsers_1.UserName AS ToUserName, dbo.TblTransacRegistrDet.FlgTrans, dbo.TblTransacRegistrDet.RecDate, dbo.TblTransacRegistrDet.ProcedureReq ,DATEDIFF(n,dbo.TblTransacRegistr.ExitDate ," & SQLDate(Now, True) & ") as diff"
Sql = Sql & "  FROM         dbo.TblUsers TblUsers_1 RIGHT OUTER JOIN"
Sql = Sql & "                       dbo.TblTransacRegistrDet ON TblUsers_1.UserID = dbo.TblTransacRegistrDet.ToUser LEFT OUTER JOIN"
Sql = Sql & "                       dbo.TblUsers ON dbo.TblTransacRegistrDet.FromUser = dbo.TblUsers.UserID LEFT OUTER JOIN"
Sql = Sql & "                       dbo.TblXXArchDocType RIGHT OUTER JOIN"
Sql = Sql & "                       dbo.TblTransacRegistr ON dbo.TblXXArchDocType.ID = dbo.TblTransacRegistr.TypTrans LEFT OUTER JOIN"
Sql = Sql & "                       dbo.TblBranchesData ON dbo.TblTransacRegistr.BrnchID = dbo.TblBranchesData.branch_id ON"
Sql = Sql & "                       dbo.TblTransacRegistrDet.TransRegID = dbo.TblTransacRegistr.ID "
Sql = Sql & "    where  dbo.TblTransacRegistrDet.FlgTrans is null "
Sql = Sql & "  and DATEDIFF(n,GETDATE(),dbo.TblTransacRegistr.ExitDate )<0"
Sql = Sql & "  AND      dbo.TblTransacRegistr.BrnchID in(" & Current_branchSql & ")"
 Dim MHDID As Integer
Rs3.Open Sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
 VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
           VSFlexGrid1.Rows = 1
If Rs3.RecordCount > 0 Then
Rs3.MoveFirst
With VSFlexGrid1
.Rows = .Rows + Rs3.RecordCount
For i = 1 To .Rows - 1
.TextMatrix(i, .ColIndex("Ser")) = i
.TextMatrix(i, .ColIndex("ExitDate")) = IIf(IsNull(Rs3("ExitDate").value), "", Rs3("ExitDate").value)
MHDID = IIf(IsNull(Rs3("MHDID").value), 0, Rs3("MHDID").value)
Select Case MHDID
Case 0
If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(i, .ColIndex("late")) = DateDiff("n", .TextMatrix(i, .ColIndex("ExitDate")), Now) & " " & "œﬁÌﬁ…"
Else
.TextMatrix(i, .ColIndex("late")) = " Minute" & " " & DateDiff("n", .TextMatrix(i, .ColIndex("ExitDate")), Now)
End If
Case 1
If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(i, .ColIndex("late")) = DateDiff("h", .TextMatrix(i, .ColIndex("ExitDate")), Now) & " " & "”«⁄…"
Else
.TextMatrix(i, .ColIndex("late")) = "Hour" & " " & DateDiff("h", .TextMatrix(i, .ColIndex("ExitDate")), Now)
End If
Case 2
If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(i, .ColIndex("late")) = DateDiff("d", .TextMatrix(i, .ColIndex("ExitDate")), Now) & " " & "ÌÊ„"
Else
.TextMatrix(i, .ColIndex("late")) = "Day" & " " & DateDiff("d", .TextMatrix(i, .ColIndex("ExitDate")), Now)
End If
Case 3
If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(i, .ColIndex("late")) = DateDiff("m", .TextMatrix(i, .ColIndex("ExitDate")), Now) & " " & "‘Â—"
Else
.TextMatrix(i, .ColIndex("late")) = "Month" & " " & DateDiff("m", .TextMatrix(i, .ColIndex("ExitDate")), Now)
End If
End Select

.TextMatrix(i, .ColIndex("ID")) = IIf(IsNull(Rs3("ID").value), 0, Rs3("ID").value)
.TextMatrix(i, .ColIndex("TransRegID")) = IIf(IsNull(Rs3("TransRegID").value), "", Rs3("TransRegID").value)
.TextMatrix(i, .ColIndex("RecordDate")) = IIf(IsNull(Rs3("RecordDate").value), "", Rs3("RecordDate").value)
.TextMatrix(i, .ColIndex("RecordDateH")) = IIf(IsNull(Rs3("RecordDateH").value), "", Rs3("RecordDateH").value)
.TextMatrix(i, .ColIndex("ImportExport")) = IIf(IsNull(Rs3("ImportExport").value), -1, Rs3("ImportExport").value) + 1
.TextMatrix(i, .ColIndex("Summary")) = IIf(IsNull(Rs3("Summary").value), "", Rs3("Summary").value)
.TextMatrix(i, .ColIndex("UserName")) = IIf(IsNull(Rs3("UserName").value), "", Rs3("UserName").value)
.TextMatrix(i, .ColIndex("RecDate")) = IIf(IsNull(Rs3("RecDate").value), "", Rs3("RecDate").value)
If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(Rs3("Name").value), "", Rs3("Name").value)
Else
.TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(Rs3("Namee").value), "", Rs3("Namee").value)
End If
Rs3.MoveNext
Next i
End With
End If
End Sub

Private Sub VSFlexGrid1_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With Me.VSFlexGrid1
If .ColKey(Col) <> "show" Then
Cancel = True
End If
End With
End Sub

Private Sub VSFlexGrid1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
With VSFlexGrid1
Select Case .ColKey(Col)
Case "show"
If val(.TextMatrix(Row, .ColIndex("ID"))) <> 0 Then
FrmTransacRegistr.Ind = 1
FrmTransacRegistr.show
FrmTransacRegistr.Ind = 1
FrmTransacRegistr.Retrive2 val(.TextMatrix(Row, .ColIndex("ID")))
End If
End Select
End With
End Sub

Private Sub VSFlexGrid1_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With Me.VSFlexGrid1
Select Case .ColKey(Col)
 Case "show"
            .ColComboList(.ColIndex("show")) = "..."
     End Select
    End With
End Sub
