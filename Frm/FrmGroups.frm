VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{E1BFA30F-D929-4F80-AEDD-76FC2BDF5E23}#1.0#0"; "ciaXPPopUp30.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{D155F1AE-D9A4-458C-8CEE-498CB717DB7B}#1.0#0"; "DBPix20.ocx"
Begin VB.Form FrmGroups 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "»Ì«‰«   „Ã„Ê⁄«  ··«’‰«›"
   ClientHeight    =   9420
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10950
   HelpContextID   =   200
   Icon            =   "FrmGroups.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   9420
   ScaleWidth      =   10950
   Begin C1SizerLibCtl.C1Elastic C1Elastic3 
      Height          =   9435
      Left            =   0
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   30
      Width           =   10965
      _cx             =   19341
      _cy             =   16642
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
      Begin VB.TextBox XPTxtID 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   5280
         Locked          =   -1  'True
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   121
         Top             =   120
         Width           =   1035
      End
      Begin VB.TextBox TxtMenuState 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3210
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Text            =   "N"
         Top             =   30
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox TxtModFlg 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2400
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   60
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox TxtCutKey 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   4080
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   0
         Visible         =   0   'False
         Width           =   615
      End
      Begin ImpulseButton.ISButton XPBtnMove 
         Height          =   330
         Index           =   0
         Left            =   1065
         TabIndex        =   6
         Top             =   120
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   582
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
         ButtonImage     =   "FrmGroups.frx":038A
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
         Height          =   330
         Index           =   2
         Left            =   0
         TabIndex        =   7
         Top             =   120
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   582
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
         ButtonImage     =   "FrmGroups.frx":0724
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
         Height          =   330
         Index           =   1
         Left            =   1590
         TabIndex        =   8
         Top             =   120
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   582
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
         ButtonImage     =   "FrmGroups.frx":0ABE
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
         Height          =   330
         Index           =   3
         Left            =   525
         TabIndex        =   9
         Top             =   120
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   582
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
         ButtonImage     =   "FrmGroups.frx":0E58
         ColorHighlight  =   4194304
         ColorHoverText  =   16777215
         ColorShadow     =   -2147483631
         ColorOutline    =   -2147483631
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
         ColorToggledHoverText=   16777215
         ColorTextShadow =   16777215
      End
      Begin C1SizerLibCtl.C1Tab C1Tab1 
         Height          =   7500
         Left            =   0
         TabIndex        =   13
         Top             =   585
         Width           =   10935
         _cx             =   19288
         _cy             =   13229
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
         Caption         =   "»Ì«‰«  «·„Ã„Ê⁄« |⁄‰«’— „—«ﬁ»… «·ÃÊœ…|‰ﬁ«ÿ «·⁄„·«¡|«· ’‰Ì›« |ŒÿÊÿ «·«‰ «Ã"
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
         Begin C1SizerLibCtl.C1Elastic C1Elastic2 
            Height          =   7125
            Left            =   11580
            TabIndex        =   61
            TabStop         =   0   'False
            Top             =   45
            Width           =   10845
            _cx             =   19129
            _cy             =   12568
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
            Begin VSFlex8UCtl.VSFlexGrid Grid 
               Height          =   3105
               Left            =   120
               TabIndex        =   62
               Top             =   120
               Width           =   8865
               _cx             =   15637
               _cy             =   5477
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
               AllowBigSelection=   0   'False
               AllowUserResizing=   0
               SelectionMode   =   1
               GridLines       =   1
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   12
               Cols            =   7
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   320
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmGroups.frx":11F2
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
               WallPaperAlignment=   9
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   24
            End
            Begin ImpulseButton.ISButton DeleteRow 
               Height          =   390
               Left            =   7560
               TabIndex        =   63
               Top             =   3240
               Width           =   1410
               _ExtentX        =   2487
               _ExtentY        =   688
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   "Õ–› ”ÿ—"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ButtonImage     =   "FrmGroups.frx":12ED
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton DeleteAll 
               Height          =   390
               Left            =   6000
               TabIndex        =   64
               Top             =   3240
               Width           =   1410
               _ExtentX        =   2487
               _ExtentY        =   688
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   "Õ–› «·ﬂ·"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ButtonImage     =   "FrmGroups.frx":1887
               DrawFocusRectangle=   0   'False
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic1 
            Height          =   7125
            Left            =   45
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   45
            Width           =   10845
            _cx             =   19129
            _cy             =   12568
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
            Begin VB.CheckBox chkTaxExempt 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„⁄›«Â „‰ «·÷—Ì»…"
               Height          =   255
               Left            =   -360
               RightToLeft     =   -1  'True
               TabIndex        =   129
               Top             =   2430
               Width           =   1935
            End
            Begin VB.CheckBox chkIsTransfere 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " ÕÊÌ·"
               Height          =   285
               Left            =   5730
               TabIndex        =   126
               Top             =   3690
               Width           =   765
            End
            Begin VB.ComboBox ServersName 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   6930
               RightToLeft     =   -1  'True
               TabIndex        =   125
               Top             =   6810
               Visible         =   0   'False
               Width           =   3345
            End
            Begin VB.CheckBox chkIsShowCover 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ÌŸÂ— ›Ï «·ﬂ›—« "
               Height          =   255
               Left            =   30
               RightToLeft     =   -1  'True
               TabIndex        =   122
               Top             =   2820
               Width           =   1935
            End
            Begin VB.TextBox txtMaxPercentDiscount 
               Alignment       =   2  'Center
               Height          =   315
               Left            =   4845
               TabIndex        =   96
               Top             =   3390
               Width           =   615
            End
            Begin VB.CheckBox chkIsAdditions 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«÷«›« "
               Height          =   255
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   93
               Top             =   3630
               Width           =   1935
            End
            Begin VB.CheckBox chkIsProducer 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„‰ Ã"
               Height          =   255
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   92
               Top             =   3390
               Width           =   1935
            End
            Begin VB.CheckBox CritiCalStock 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " ŸÂ— ›Ì «·„Œ“‰ «·Õ—Ã"
               Height          =   255
               Left            =   2280
               RightToLeft     =   -1  'True
               TabIndex        =   70
               Top             =   2790
               Width           =   1935
            End
            Begin VB.CheckBox posgroup 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " ŸÂ— ›Ì ‰ﬁÿ… «·»Ì⁄"
               Height          =   255
               Left            =   2280
               RightToLeft     =   -1  'True
               TabIndex        =   69
               Top             =   3120
               Width           =   1935
            End
            Begin VB.ComboBox CBOprintrs 
               Height          =   315
               Left            =   75
               RightToLeft     =   -1  'True
               TabIndex        =   67
               Text            =   "Combo1"
               Top             =   600
               Width           =   1935
            End
            Begin VB.CheckBox ChSeparate 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„‰›’·…"
               Height          =   255
               Left            =   4080
               RightToLeft     =   -1  'True
               TabIndex        =   66
               Top             =   3000
               Width           =   1335
            End
            Begin VB.CheckBox ChkHoldingMaterials 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„Ê«œ „Õ„·Â"
               Height          =   255
               Left            =   5160
               RightToLeft     =   -1  'True
               TabIndex        =   65
               Top             =   3030
               Width           =   1335
            End
            Begin VB.CheckBox IsMaterial 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " ŸÂ— ›Ì «·„ƒ‘—«  «·ÕÌ…"
               Height          =   255
               Left            =   4440
               RightToLeft     =   -1  'True
               TabIndex        =   22
               Top             =   2790
               Width           =   2055
            End
            Begin VB.CheckBox Chklast 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„Ã„Ê⁄Â ‰Â«∆Ì…"
               Height          =   255
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   21
               Top             =   1200
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.TextBox TxtEXpireValue 
               Alignment       =   2  'Center
               Height          =   315
               Left            =   4515
               TabIndex        =   20
               Top             =   2430
               Width           =   615
            End
            Begin VB.ComboBox CboEXpirType 
               Height          =   315
               ItemData        =   "FrmGroups.frx":1E21
               Left            =   3480
               List            =   "FrmGroups.frx":1E2E
               RightToLeft     =   -1  'True
               TabIndex        =   19
               Top             =   2400
               Width           =   855
            End
            Begin VB.TextBox TxtOverHead 
               Alignment       =   2  'Center
               Height          =   315
               Left            =   1755
               TabIndex        =   18
               Top             =   2430
               Width           =   615
            End
            Begin VB.TextBox XPTxtName 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   75
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   17
               Top             =   1005
               Width           =   5055
            End
            Begin VB.TextBox txtid 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   4125
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   16
               Top             =   600
               Width           =   1005
            End
            Begin VB.TextBox XPTxtNameE 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   75
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   15
               Top             =   1320
               Width           =   5055
            End
            Begin MSComctlLib.TreeView TreeGroups 
               Height          =   6525
               Left            =   6600
               TabIndex        =   23
               Top             =   0
               Width           =   4275
               _ExtentX        =   7541
               _ExtentY        =   11509
               _Version        =   393217
               HideSelection   =   0   'False
               Indentation     =   441
               LabelEdit       =   1
               LineStyle       =   1
               Style           =   7
               Appearance      =   1
               Enabled         =   0   'False
            End
            Begin MSDataListLib.DataCombo XPCboGroup 
               Height          =   315
               Left            =   75
               TabIndex        =   24
               Top             =   1725
               Width           =   5055
               _ExtentX        =   8916
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DCPreFix 
               Height          =   315
               Left            =   3120
               TabIndex        =   25
               Top             =   600
               Visible         =   0   'False
               Width           =   915
               _ExtentX        =   1614
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   16777215
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin DBPIXLib.DBPix20 DBPix201 
               Height          =   555
               Left            =   0
               TabIndex        =   26
               Top             =   2640
               Width           =   1575
               _Version        =   131072
               _ExtentX        =   2778
               _ExtentY        =   979
               _StockProps     =   1
               BackColor       =   12632256
               _Image          =   "FrmGroups.frx":1E41
               ImageResampleWidth=   100
               ImageResampleHeight=   100
               ImageResampleMode=   1
               ImageSaveFormat =   0
               JPEGQuality     =   75
               JPEGEncoding    =   0
               JPEGColorMode   =   0
               JPEGNoRecompress=   -1  'True
               JPEGRotateWarning=   0
               PNGColorDepth   =   0
               PNGCompression  =   0
               PNGFilter       =   0
               PNGInterlace    =   1
               ImageDitherMethod=   3
               ImagePaletteMethod=   4
               ImagePreviewMode=   0   'False
               ImageKeepMetaData=   -1  'True
               UseAmbientBackcolor=   -1  'True
               ViewAsyncDecoding=   -1  'True
               ViewEnableMouseZoom=   -1  'True
               ViewInitialZoom =   0
               ViewHAlign      =   1
               ViewVAlign      =   1
               ViewMenuMode    =   0
            End
            Begin ImpulseButton.ISButton ISButton1 
               Height          =   255
               Left            =   2280
               TabIndex        =   27
               Top             =   3420
               Width           =   1545
               _ExtentX        =   2725
               _ExtentY        =   450
               ButtonPositionImage=   1
               Caption         =   "«œ—«Ã ’Ê—… «·„Ã„Ê⁄Â"
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
               DisabledImageExtraction=   0
               ColorToggledHoverText=   16711680
               ColorTextShadow =   -2147483637
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   1605
               Index           =   0
               Left            =   120
               TabIndex        =   28
               TabStop         =   0   'False
               Top             =   5460
               Width           =   6345
               _cx             =   11192
               _cy             =   2831
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
               ForeColor       =   4210752
               FloodColor      =   6553600
               ForeColorDisabled=   -2147483631
               Picture         =   "FrmGroups.frx":1E59
               Caption         =   "„⁄·Ê„«  ≈÷«›Ì… ⁄‰ «·„Ã„Ê⁄…"
               Align           =   0
               AutoSizeChildren=   0
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
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "0"
                  Height          =   225
                  Index           =   10
                  Left            =   30
                  RightToLeft     =   -1  'True
                  TabIndex        =   38
                  Top             =   330
                  Width           =   2775
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "0"
                  Height          =   225
                  Index           =   11
                  Left            =   30
                  RightToLeft     =   -1  'True
                  TabIndex        =   37
                  Top             =   570
                  Width           =   2775
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "0"
                  Height          =   225
                  Index           =   12
                  Left            =   2700
                  RightToLeft     =   -1  'True
                  TabIndex        =   36
                  Top             =   990
                  Width           =   1215
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "0"
                  Height          =   225
                  Index           =   13
                  Left            =   30
                  RightToLeft     =   -1  'True
                  TabIndex        =   35
                  Top             =   1350
                  Width           =   2775
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«Œ— ’‰› „÷«› ≈·Ï «·„Ã„Ê⁄…"
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
                  Height          =   225
                  Index           =   9
                  Left            =   3990
                  RightToLeft     =   -1  'True
                  TabIndex        =   34
                  Top             =   1230
                  Width           =   2175
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«Ê· ’‰› „÷«› ≈·Ï «·„Ã„Ê⁄…"
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
                  Height          =   225
                  Index           =   8
                  Left            =   3750
                  RightToLeft     =   -1  'True
                  TabIndex        =   33
                  Top             =   870
                  Width           =   2415
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "⁄œœ «·√’‰«› «· Ï  Õ ÊÌÂ« «·„Ã„Ê⁄…"
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
                  Height          =   225
                  Index           =   7
                  Left            =   2790
                  RightToLeft     =   -1  'True
                  TabIndex        =   32
                  Top             =   570
                  Width           =   3375
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "⁄œœ «·„Ã„Ê⁄«  «·›—⁄Ì… «· Ï  Õ ÊÌÂ« «·„Ã„Ê⁄…"
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
                  Height          =   225
                  Index           =   4
                  Left            =   2790
                  RightToLeft     =   -1  'True
                  TabIndex        =   31
                  Top             =   330
                  Width           =   3375
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "0"
                  Height          =   225
                  Index           =   14
                  Left            =   30
                  RightToLeft     =   -1  'True
                  TabIndex        =   30
                  Top             =   990
                  Width           =   2775
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "0"
                  Height          =   225
                  Index           =   15
                  Left            =   2700
                  RightToLeft     =   -1  'True
                  TabIndex        =   29
                  Top             =   1350
                  Width           =   1215
               End
            End
            Begin MSDataListLib.DataCombo Dcbranch 
               Bindings        =   "FrmGroups.frx":21F3
               Height          =   315
               Left            =   75
               TabIndex        =   71
               Top             =   120
               Width           =   2415
               _ExtentX        =   4260
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               BackColor       =   16777215
               ListField       =   "account_name"
               BoundColumn     =   "code"
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
            Begin C1SizerLibCtl.C1Elastic C1Elastic5 
               Height          =   1095
               Left            =   120
               TabIndex        =   73
               TabStop         =   0   'False
               Top             =   4320
               Width           =   6345
               _cx             =   11192
               _cy             =   1931
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
               ForeColor       =   4210752
               FloodColor      =   6553600
               ForeColorDisabled=   -2147483631
               Picture         =   "FrmGroups.frx":2208
               Caption         =   "≈Õ ”«» ‰ﬁ«ÿ «·⁄„·«¡"
               Align           =   0
               AutoSizeChildren=   0
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
               Begin VB.TextBox TxtEquation2 
                  Alignment       =   2  'Center
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   120
                  TabIndex        =   88
                  Top             =   720
                  Width           =   855
               End
               Begin VB.TextBox TxtEquation1 
                  Alignment       =   2  'Center
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   120
                  TabIndex        =   87
                  Top             =   360
                  Width           =   855
               End
               Begin VB.TextBox TxtEqualReal 
                  Alignment       =   2  'Center
                  Height          =   315
                  Left            =   2760
                  TabIndex        =   85
                  Top             =   720
                  Width           =   855
               End
               Begin VB.TextBox TxtEqualPiont 
                  Alignment       =   2  'Center
                  Height          =   315
                  Left            =   2760
                  TabIndex        =   82
                  Top             =   360
                  Width           =   855
               End
               Begin VB.TextBox TxtEveryPoints 
                  Alignment       =   2  'Center
                  Height          =   315
                  Left            =   4800
                  TabIndex        =   80
                  Top             =   720
                  Width           =   855
               End
               Begin VB.TextBox TxtEveryAmount 
                  Alignment       =   2  'Center
                  Height          =   315
                  Left            =   4800
                  TabIndex        =   76
                  Top             =   360
                  Width           =   855
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " ”«ÊÌ"
                  ForeColor       =   &H00000080&
                  Height          =   195
                  Index           =   36
                  Left            =   3720
                  RightToLeft     =   -1  'True
                  TabIndex        =   91
                  Top             =   720
                  Width           =   495
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·„⁄«œ·…"
                  Height          =   315
                  Index           =   35
                  Left            =   1080
                  RightToLeft     =   -1  'True
                  TabIndex        =   90
                  Top             =   720
                  Width           =   615
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "«·„⁄«œ·…"
                  Height          =   315
                  Index           =   34
                  Left            =   1080
                  RightToLeft     =   -1  'True
                  TabIndex        =   89
                  Top             =   360
                  Width           =   615
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "—Ì«·"
                  ForeColor       =   &H00800000&
                  Height          =   315
                  Index           =   32
                  Left            =   2280
                  RightToLeft     =   -1  'True
                  TabIndex        =   86
                  Top             =   720
                  Width           =   375
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   " ”«ÊÌ"
                  ForeColor       =   &H00000080&
                  Height          =   195
                  Index           =   31
                  Left            =   3720
                  RightToLeft     =   -1  'True
                  TabIndex        =   84
                  Top             =   360
                  Width           =   495
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "‰ﬁÿ…"
                  ForeColor       =   &H00800000&
                  Height          =   315
                  Index           =   30
                  Left            =   2280
                  RightToLeft     =   -1  'True
                  TabIndex        =   83
                  Top             =   360
                  Width           =   375
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "‰ﬁÿ…"
                  ForeColor       =   &H00800000&
                  Height          =   315
                  Index           =   29
                  Left            =   4320
                  RightToLeft     =   -1  'True
                  TabIndex        =   81
                  Top             =   720
                  Width           =   375
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "—Ì«·"
                  ForeColor       =   &H00800000&
                  Height          =   315
                  Index           =   27
                  Left            =   4320
                  RightToLeft     =   -1  'True
                  TabIndex        =   79
                  Top             =   360
                  Width           =   375
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "ﬂ· "
                  Height          =   315
                  Index           =   26
                  Left            =   5025
                  RightToLeft     =   -1  'True
                  TabIndex        =   78
                  Top             =   720
                  Width           =   1215
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "⁄‰ ﬂ· "
                  Height          =   315
                  Index           =   25
                  Left            =   5280
                  RightToLeft     =   -1  'True
                  TabIndex        =   77
                  Top             =   360
                  Width           =   975
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "0"
                  Height          =   225
                  Index           =   33
                  Left            =   30
                  RightToLeft     =   -1  'True
                  TabIndex        =   75
                  Top             =   1470
                  Width           =   3135
               End
               Begin VB.Label lbl 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "0"
                  Height          =   225
                  Index           =   28
                  Left            =   30
                  RightToLeft     =   -1  'True
                  TabIndex        =   74
                  Top             =   1950
                  Width           =   3375
               End
            End
            Begin MSDataListLib.DataCombo cmbSpecification 
               Bindings        =   "FrmGroups.frx":25A2
               Height          =   315
               Left            =   60
               TabIndex        =   94
               Top             =   2040
               Width           =   5025
               _ExtentX        =   8864
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   16777215
               ListField       =   ""
               BoundColumn     =   "code"
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
            Begin MSDataListLib.DataCombo cmbAqrCompenet 
               Bindings        =   "FrmGroups.frx":25B7
               Height          =   315
               Left            =   0
               TabIndex        =   123
               Top             =   3990
               Width           =   5025
               _ExtentX        =   8864
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   16777215
               ListField       =   ""
               BoundColumn     =   "code"
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
            Begin MSDataListLib.DataCombo DcActivityType 
               Bindings        =   "FrmGroups.frx":25CC
               Height          =   315
               Left            =   3120
               TabIndex        =   127
               Top             =   120
               Width           =   2415
               _ExtentX        =   4260
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               BackColor       =   16777215
               ListField       =   "account_name"
               BoundColumn     =   "code"
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
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·‰‘«ÿ"
               Height          =   315
               Index           =   41
               Left            =   5505
               RightToLeft     =   -1  'True
               TabIndex        =   128
               Top             =   150
               Width           =   735
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„Ã„Ê⁄«  «·„’—Ê›"
               Height          =   225
               Index           =   24
               Left            =   5130
               RightToLeft     =   -1  'True
               TabIndex        =   124
               Top             =   4020
               Width           =   1305
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·›—⁄"
               Height          =   315
               Index           =   40
               Left            =   2460
               RightToLeft     =   -1  'True
               TabIndex        =   112
               Top             =   150
               Width           =   495
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "√ﬁ’Ï Œ’„"
               Height          =   315
               Index           =   39
               Left            =   5190
               RightToLeft     =   -1  'True
               TabIndex        =   98
               Top             =   3390
               Width           =   1215
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "%"
               Height          =   315
               Index           =   38
               Left            =   4530
               RightToLeft     =   -1  'True
               TabIndex        =   97
               Top             =   3390
               Width           =   255
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·›—“"
               Height          =   315
               Index           =   37
               Left            =   5160
               RightToLeft     =   -1  'True
               TabIndex        =   95
               Top             =   2070
               Width           =   1215
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ÿ«»⁄Â"
               Height          =   315
               Index           =   23
               Left            =   2040
               RightToLeft     =   -1  'True
               TabIndex        =   68
               Top             =   600
               Width           =   735
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·’Ê—…"
               Height          =   195
               Index           =   19
               Left            =   3840
               RightToLeft     =   -1  'True
               TabIndex        =   47
               Top             =   3420
               Width           =   615
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "%"
               Height          =   315
               Index           =   22
               Left            =   1440
               RightToLeft     =   -1  'True
               TabIndex        =   46
               Top             =   2430
               Width           =   255
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·’·«ÕÌ…"
               Height          =   315
               Index           =   18
               Left            =   5220
               RightToLeft     =   -1  'True
               TabIndex        =   45
               Top             =   2430
               Width           =   1215
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "‰”»… «· Õ„Ì·"
               Height          =   315
               Index           =   21
               Left            =   2100
               RightToLeft     =   -1  'True
               TabIndex        =   44
               Top             =   2430
               Width           =   1215
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·«”„ ⁄—»Ì"
               Height          =   315
               Index           =   5
               Left            =   5220
               RightToLeft     =   -1  'True
               TabIndex        =   43
               Top             =   1005
               Width           =   1215
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„Ã„Ê⁄… «·—∆Ì”Ì…"
               Height          =   315
               Index           =   0
               Left            =   5220
               RightToLeft     =   -1  'True
               TabIndex        =   42
               Top             =   1725
               Width           =   1215
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„Ã„Ê⁄Â ‰Â«∆Ì…"
               Height          =   315
               Index           =   16
               Left            =   1440
               RightToLeft     =   -1  'True
               TabIndex        =   41
               Top             =   -360
               Visible         =   0   'False
               Width           =   1215
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬂÊœ «·„Ã„Ê⁄…"
               Height          =   315
               Index           =   17
               Left            =   5220
               RightToLeft     =   -1  'True
               TabIndex        =   40
               Top             =   600
               Width           =   1215
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·«”„ «‰Ã·Ì“Ì"
               Height          =   315
               Index           =   20
               Left            =   5220
               RightToLeft     =   -1  'True
               TabIndex        =   39
               Top             =   1320
               Width           =   1215
            End
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic4 
            Height          =   7125
            Left            =   11880
            TabIndex        =   72
            TabStop         =   0   'False
            Top             =   45
            Width           =   10845
            _cx             =   19129
            _cy             =   12568
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
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic6 
            Height          =   7125
            Left            =   12180
            TabIndex        =   99
            TabStop         =   0   'False
            Top             =   45
            Width           =   10845
            _cx             =   19129
            _cy             =   12568
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
            Begin MSDataListLib.DataCombo dcClass 
               Height          =   315
               Index           =   1
               Left            =   5820
               TabIndex        =   102
               Top             =   270
               Width           =   2595
               _ExtentX        =   4577
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   16777215
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo dcClass 
               Height          =   315
               Index           =   2
               Left            =   5820
               TabIndex        =   104
               Top             =   750
               Width           =   2595
               _ExtentX        =   4577
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   16777215
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo dcClass 
               Height          =   315
               Index           =   3
               Left            =   5820
               TabIndex        =   106
               Top             =   1320
               Width           =   2595
               _ExtentX        =   4577
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   16777215
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo dcClass 
               Height          =   315
               Index           =   4
               Left            =   5820
               TabIndex        =   108
               Top             =   1860
               Width           =   2595
               _ExtentX        =   4577
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   16777215
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo dcClass 
               Height          =   315
               Index           =   5
               Left            =   5820
               TabIndex        =   110
               Top             =   2400
               Width           =   2595
               _ExtentX        =   4577
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   16777215
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo dcClass 
               Height          =   315
               Index           =   6
               Left            =   5820
               TabIndex        =   111
               Top             =   2940
               Width           =   2595
               _ExtentX        =   4577
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   16777215
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«· ’‰Ì› 6"
               ForeColor       =   &H00000000&
               Height          =   375
               Index           =   5
               Left            =   9090
               TabIndex        =   109
               Top             =   3000
               Width           =   690
            End
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«· ’‰Ì› 5"
               ForeColor       =   &H00000000&
               Height          =   375
               Index           =   4
               Left            =   9090
               TabIndex        =   107
               Top             =   2490
               Width           =   690
            End
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«· ’‰Ì› 4"
               ForeColor       =   &H00000000&
               Height          =   375
               Index           =   3
               Left            =   9090
               TabIndex        =   105
               Top             =   1950
               Width           =   690
            End
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«· ’‰Ì› 3"
               ForeColor       =   &H00000000&
               Height          =   375
               Index           =   2
               Left            =   9090
               TabIndex        =   103
               Top             =   1380
               Width           =   690
            End
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«· ’‰Ì› 2"
               ForeColor       =   &H00000000&
               Height          =   375
               Index           =   0
               Left            =   9090
               TabIndex        =   101
               Top             =   870
               Width           =   690
            End
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«· ’‰Ì› 1"
               ForeColor       =   &H00000000&
               Height          =   375
               Index           =   1
               Left            =   9090
               TabIndex        =   100
               Top             =   330
               Width           =   690
            End
         End
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   7125
            Index           =   22
            Left            =   12480
            TabIndex        =   113
            TabStop         =   0   'False
            Top             =   45
            Width           =   10845
            _cx             =   19129
            _cy             =   12568
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
            Begin VSFlex8UCtl.VSFlexGrid grdProductLine 
               Height          =   4680
               Left            =   150
               TabIndex        =   115
               Top             =   540
               Width           =   10470
               _cx             =   18468
               _cy             =   8255
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
               Rows            =   1
               Cols            =   5
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"FrmGroups.frx":25E1
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
               WallPaperAlignment=   0
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   24
            End
            Begin C1SizerLibCtl.C1Elastic Ele 
               Height          =   1440
               Index           =   23
               Left            =   225
               TabIndex        =   116
               TabStop         =   0   'False
               Top             =   5430
               Width           =   10845
               _cx             =   19129
               _cy             =   2540
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
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   420
                  Index           =   40
                  Left            =   2100
                  TabIndex        =   117
                  Top             =   765
                  Width           =   1260
                  _ExtentX        =   2223
                  _ExtentY        =   741
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
                  ColorButton     =   14871017
               End
               Begin ImpulseButton.ISButton Cmd 
                  Height          =   420
                  Index           =   41
                  Left            =   735
                  TabIndex        =   118
                  Top             =   765
                  Width           =   1260
                  _ExtentX        =   2223
                  _ExtentY        =   741
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
               End
               Begin MSDataListLib.DataCombo cmbProductLine 
                  Height          =   315
                  Left            =   3570
                  TabIndex        =   119
                  Top             =   660
                  Width           =   4230
                  _ExtentX        =   7461
                  _ExtentY        =   556
                  _Version        =   393216
                  Text            =   ""
                  RightToLeft     =   -1  'True
               End
               Begin VB.Label lbl«”„«·ÊÕœ… 
                  Alignment       =   2  'Center
                  BackColor       =   &H00E2E9E9&
                  Caption         =   "Œÿ «·«‰ «Ã"
                  Height          =   450
                  Index           =   8
                  Left            =   8295
                  TabIndex        =   120
                  Top             =   675
                  Width           =   1740
               End
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Label1"
               Height          =   105
               Left            =   4215
               RightToLeft     =   -1  'True
               TabIndex        =   114
               Top             =   570
               Width           =   225
            End
         End
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   390
         Index           =   0
         Left            =   8085
         TabIndex        =   52
         Top             =   8895
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   688
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
         ColorShadow     =   -2147483637
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         DisabledImageExtraction=   0
         ColorToggledHoverText=   16711680
         ColorTextShadow =   -2147483637
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   390
         Index           =   1
         Left            =   7335
         TabIndex        =   53
         Top             =   8895
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   688
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
         ColorShadow     =   -2147483637
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         ColorToggledHoverText=   16711680
         ColorTextShadow =   -2147483637
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   390
         Index           =   2
         Left            =   6600
         TabIndex        =   54
         Top             =   8895
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   688
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
         ColorShadow     =   -2147483637
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         ColorToggledHoverText=   16711680
         ColorTextShadow =   -2147483637
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   390
         Index           =   3
         Left            =   5895
         TabIndex        =   55
         Top             =   8895
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   688
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
         ColorShadow     =   -2147483637
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         ColorToggledHoverText=   16711680
         ColorTextShadow =   -2147483637
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   390
         Index           =   4
         Left            =   4965
         TabIndex        =   56
         Top             =   8895
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   688
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
         ColorShadow     =   -2147483637
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         ColorToggledHoverText=   16711680
         ColorTextShadow =   -2147483637
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   390
         Index           =   5
         Left            =   4110
         TabIndex        =   57
         Top             =   8895
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   688
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
         ColorShadow     =   -2147483637
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         ColorToggledHoverText=   16711680
         ColorTextShadow =   -2147483637
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   390
         Index           =   6
         Left            =   1560
         TabIndex        =   58
         Top             =   8895
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   688
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
         ColorShadow     =   -2147483637
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         ColorToggledHoverText=   16711680
         ColorTextShadow =   -2147483637
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   390
         Index           =   7
         Left            =   3270
         TabIndex        =   59
         Top             =   8895
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   688
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
         ColorShadow     =   -2147483637
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         ColorToggledHoverText=   16711680
         ColorTextShadow =   -2147483637
      End
      Begin ImpulseButton.ISButton CmdHelp 
         Height          =   390
         Left            =   2280
         TabIndex        =   60
         Top             =   8895
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   688
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
         ColorShadow     =   -2147483637
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         ColorToggledHoverText=   16711680
         ColorTextShadow =   -2147483637
      End
      Begin VB.Label XPTxtCurrent 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Height          =   300
         Left            =   2220
         RightToLeft     =   -1  'True
         TabIndex        =   51
         Top             =   8100
         Width           =   765
      End
      Begin VB.Label XPTxtCount 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Height          =   300
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   50
         Top             =   8070
         Width           =   615
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "⁄œœ «·”Ã·« :"
         Height          =   300
         Index           =   2
         Left            =   870
         RightToLeft     =   -1  'True
         TabIndex        =   49
         Top             =   8070
         Width           =   1095
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«·”Ã· «·Õ«·Ì:"
         Height          =   300
         Index           =   1
         Left            =   2880
         RightToLeft     =   -1  'True
         TabIndex        =   48
         Top             =   8070
         Width           =   1185
      End
      Begin VB.Label LblHeader 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   " »Ì«‰«   „Ã„Ê⁄«  ··«’‰«›  "
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
         Height          =   555
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   0
         Width           =   10995
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "«‰‘«¡ «·Õ”«»« "
      Height          =   255
      Left            =   3480
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   9600
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox TxtGroupCode 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   2700
      Locked          =   -1  'True
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   9540
      Width           =   1875
   End
   Begin ciaXPPopMenu30.XPPopUp30 XPPopUp 
      Left            =   4110
      Top             =   5190
      _ExtentX        =   900
      _ExtentY        =   873
      VisualStyle     =   0
      BeginProperty DefaultMenuItemFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty ToolTipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MenuItemSpacing =   0
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "ﬂÊœ «·„Ã„Ê⁄…"
      Height          =   315
      Index           =   3
      Left            =   4920
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   9480
      Width           =   1335
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·„Ã„Ê⁄…"
      Height          =   315
      Index           =   6
      Left            =   1260
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   9540
      Width           =   1095
   End
End
Attribute VB_Name = "FrmGroups"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim TTP As clstooltip
Dim GroupReport As ClsGroupReport
Dim cSearchDcbo As clsDCboSearch


Private Sub RemoveGridRow()
If Me.TxtModFlg.text <> "R" Then
    With Me.Grid
        If .row <= 0 Then Exit Sub
        .RemoveItem .row
    End With
    End If
End Sub

Private Sub Cmd_Click(Index As Integer)
    Dim currentgroup As Integer
    On Error GoTo ErrTrap

    Select Case Index

        Case 0

            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If
        
            currentgroup = val(XPCboGroup.BoundText)
            TxtModFlg.text = "N"
        
            clear_all Me

            '        XPTxtID.text = CStr(new_id("Groups", "GroupID", "", True))
            If currentgroup = 0 Then
                XPCboGroup.BoundText = 1
            Else
                XPCboGroup.BoundText = currentgroup

            End If
            

            If Me.XPCboGroup.BoundText <> "" Then
            
            End If

            XPTxtName.SetFocus

            Grid.Clear flexClearScrollable, flexClearEverything
            Grid.Rows = 2
        Case 1

            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If
           Grid.Rows = Grid.Rows + 1
            TxtModFlg.text = "E"
            CuurentLogdata

        Case 2

            Dim currentcode As String
 
 If SystemOptions.WorkWithGroupCode = False Then
            If txtid = "" Then
                MsgBox "«œŒ· ﬂÊœ «·„Ã„Ê⁄Â     "
                Exit Sub
                'Else
                'txtid = CurrentCode
            End If
End If

            'End If

            SaveData
Case 40
    AddNewRowProductLine

Case 41
    RemoveProductLineRow

        Case 3
            Undo

        Case 4

            If DoPremis(Do_Delete, Me.Name, True) = False Then
                Exit Sub
            End If

            Del_Group

        Case 5

            If DoPremis(Do_Search, Me.Name, True) = False Then
                Exit Sub
            End If

        
      'WaelComment    FrmGroupSearch.show vbModal

        Case 6
            Unload Me

        Case 7

            If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If
            print_report2

            'saa PrintReport
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub CmdHelp_Click()
    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
    SystemOptions.SysHelp.HHDisplayTopicID Me.hWnd
End Sub

Private Sub Command1_Click()
    'Â–… «·œ«·…  ﬁÊ„ »Õ–› ﬂ· Õ”«»«  «·„Ã„Ê⁄«  »‘—ÿ ⁄œ„ «‘ —«ﬂÂ« ›Ì ﬁÌÊœ Ê«‰‘«¡ Õ”«»«  ÃœÌœ… ·ﬂ· «·„Ã„Ê⁄«  ÿ»ﬁ« ·Õ”«»«  «·›—⁄ «·ÃœÌœ…
    Dim Msg As String
    Dim StrSQL As String
    Dim RsTemp As New ADODB.Recordset
    Dim RsTemp2 As New ADODB.Recordset
    Dim i As Integer
 
    StrSQL = "select * From Groups  "
    RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Not (RsTemp.EOF Or RsTemp.BOF) Then

        For i = 1 To RsTemp.RecordCount
            delete_group_account RsTemp("GroupID").value
            RsTemp.MoveNext
        Next i

    End If

    RsTemp.Close
    StrSQL = "select * From Groups  where LastGroup=1 "
    RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    
    If Not (RsTemp.EOF Or RsTemp.BOF) Then

        For i = 1 To RsTemp.RecordCount

            If create_accounts(RsTemp("GroupID").value, RsTemp("GroupName").value) Then
                
            End If
         
            RsTemp.MoveNext
        Next i

    End If

    MsgBox "Done"

End Sub

Private Sub DcActivityType_Click(Area As Integer)
If Me.TxtModFlg <> "R" Then

  
Dim My_SQL As String

If SystemOptions.UserInterface = ArabicInterface Then
         My_SQL = "  select branch_id,branch_name from TblBranchesData  where ActivityTypeId=" & val(DcActivityType.BoundText) & "    order by branch_id  "
 Else
      My_SQL = "  select branch_id,branch_namee from TblBranchesData    where ActivityTypeId=" & val(DcActivityType.BoundText) & "   order by branch_id  "
 End If
 
    fill_combo dcBranch, My_SQL
    
    End If
    
End Sub

Private Sub DeleteAll_Click()
If Me.TxtModFlg.text <> "R" Then
 Grid.Clear flexClearScrollable, flexClearEverything
 Grid.Rows = 2
End If
End Sub

Private Sub DeleteRow_Click()
RemoveGridRow
End Sub

Private Sub Form_Activate()
    XPTxtID.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        If Me.TxtModFlg.text = "R" Then
            Cmd_Click (0)
        Else
            Sendkeys "{TAB}"
        End If
    End If

    If Me.TxtModFlg.text = "R" Then
        If KeyCode = vbKeyDown Or KeyCode = vbKeyEnd Then
            XPBtnMove_Click (2)
        ElseIf KeyCode = vbKeyUp Or KeyCode = vbKeyHome Then
            XPBtnMove_Click (1)
        ElseIf KeyCode = vbKeyRight Or KeyCode = vbKeyPageDown Then
            XPBtnMove_Click (3)
        ElseIf KeyCode = vbKeyLeft Or KeyCode = vbKeyPageUp Then
            XPBtnMove_Click (0)
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

    If Shift = 2 Then
        If KeyCode = vbKeyX Then
            If Cmd(6).Enabled = False Then Exit Sub
            Cmd_Click (6)
        End If
    End If

    Exit Sub
ErrTrap:
End Sub
Function OpenPrinters()
Dim a As String
Dim VarSet As String
    If Dir(App.path & "\printersGroup.txt", vbNormal) <> "" Then

        
           
         
        
            Open App.path & "\printersGroup.txt" For Input As #1
    CBOprintrs.Clear

    Do Until EOF(1)
        Line Input #1, a
        'subsequent lines
 
        If a <> "" Then
        
                CBOprintrs.AddItem (a)
         
           
        End If
    
    Loop

    Close #1
  End If
End Function

Private Sub Form_Load()
    Dim Num    As Integer
    Dim StrSQL As String
    '    On Error GoTo ErrTrap

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    OpenPrinters
    ScreenNameArabic = " »Ì«‰«   „Ã„Ê⁄«  ··«’‰«› "
    ScreenNameEnglish = "ItemS Groups  "
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish, "1"

    If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "  SELECT id ,name FROM tblActivitesType order by name"
    Else
        StrSQL = "  SELECT id ,namee FROM tblActivitesType order by name"
    End If

    fill_combo Me.DcActivityType, StrSQL
  
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("New").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Edit").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("save").Picture
    Set Cmd(3).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Undo").Picture
    Set Cmd(4).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Del").Picture
    Set Cmd(5).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(6).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    Set Cmd(7).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Print").Picture
    Set CmdHelp.ButtonImage = mdifrmmain.ImgLstTree.ListImages("Help").Picture
    Resize_Form Me
    TreeGroups.ImageList = mdifrmmain.ImgLstTree
    LoadTreeGroups Me.TreeGroups
    Set rs = New ADODB.Recordset
    'StrSQL = "select * From groups where GroupID <> 1"
    StrSQL = "select * From groups where GroupID > 1"
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    FillGroupCombo
    'Set cSearchDcbo = New clsDCboSearch
    'Set cSearchDcbo.Client = XPCboGroup
    Dim Dcombos As New ClsDataCombos
    Dcombos.GetBranches Me.dcBranch
    Dcombos.GetPrefix Me.DCPreFix, 2, val(branch_id)
    Dcombos.GetItemSpecifications Me.cmbSpecification
    
    Dim My_SQL As String

    My_SQL = "SELECT      id,name FROM         dbo.TblProductLine Where 1 = 1 "
   
    My_SQL = My_SQL & " And IsNull(IsBasicLine,0) = 0"
    My_SQL = My_SQL & " Order By name ASC"
    fill_combo cmbProductLine, My_SQL
        
    My_SQL = "SELECT      id,name FROM         dbo.TblAqrCompenet Where 1 = 1 "
        
    My_SQL = My_SQL & " Order By name ASC"
    fill_combo cmbAqrCompenet, My_SQL
    
    If SystemOptions.UserInterface = ArabicInterface Then
        My_SQL = "  SELECT SizeID,SizeName FROM TblItemsclasses   "
    Else
        My_SQL = "  SELECT SizeID,SizeNamee FROM TblItemsclasses "
    End If

    Dim i As Integer

    For i = 1 To dcClass.count
        fill_combo dcClass(i), My_SQL
    Next

    XPBtnMove_Click 2
    LoadMenus
    Me.TxtModFlg.text = "R"

    AddTip

    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish

    If rs.State = adStateOpen Then
        If Not (rs.EOF Or rs.BOF) Then
            If rs.EditMode <> adEditNone Then
                rs.CancelUpdate
            End If
        End If

        rs.Close
    End If

    Set rs = Nothing
    Set TTP = Nothing
    Set GroupReport = Nothing
    Exit Sub
ErrTrap:
End Sub

Function CuurentLogdata(Optional Currentmode As String)
    LogTextA = "    ‘«‘… " & ScreenNameArabic & CHR(13) & " ﬂÊœ «·„Ã„Ê⁄… " & txtid.text & CHR(13) & "   «”„ «·„Ã„Ê⁄Â" & XPTxtName.text & CHR(13) & " «·„Ã„Ê⁄Â «·—∆Ì”Ì…   " & XPCboGroup.text

    If Chklast.value = vbChecked Then
        LogTextA = LogTextA & " „Ã„Ê⁄Â ‰Â«∆Ì…"
    End If
                     
    LogTexte = "    Screen  " & ScreenNameEnglish & CHR(13) & "Code " & txtid.text & CHR(13) & "   Name  " & XPTxtName.text & CHR(13) & " Parent Group" & XPCboGroup.text

    If Chklast.value = vbChecked Then
        LogTextA = LogTextA & "  Final Group"
    End If
                     
    If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, Me.TxtModFlg
    Else
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, "D"
    End If
    
End Function

Private Sub Grid_AfterEdit(ByVal row As Long, ByVal Col As Long)
    On Error GoTo ErrTrap
    Dim StrAccountCode As String
    Dim Msg As String
    Dim rs2 As New ADODB.Recordset
    Dim StrSQL As String

    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long
    With Grid
        Select Case .ColKey(Col)
            Case "Name"
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("GQCID"), False, True)
                .TextMatrix(row, .ColIndex("GQCID")) = StrAccountCode
                StrSQL = " SELECT     Grade,Comment From dbo.TblQCItems Where (qcid = " & val(.TextMatrix(row, .ColIndex("GQCID"))) & ")"
                Set rs2 = New ADODB.Recordset
                rs2.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                If rs2.RecordCount > 0 Then
                .TextMatrix(row, .ColIndex("Comment")) = IIf(IsNull(rs2("Comment").value), "", rs2("Comment").value)
                .TextMatrix(row, .ColIndex("Standrd")) = IIf(IsNull(rs2("Grade").value), "", rs2("Grade").value)
                Else
                .TextMatrix(row, .ColIndex("Comment")) = ""
                .TextMatrix(row, .ColIndex("Standrd")) = ""
                End If
         End Select
        If row = .Rows - 1 Then
            .Rows = .Rows + 1
        End If
    End With

ErrTrap:
End Sub

Private Sub Grid_BeforeEdit(ByVal row As Long, ByVal Col As Long, Cancel As Boolean)
With Grid
Select Case .ColKey(Col)
Case "Comment"
.ComboList = ""
Case "Standrd"
.ComboList = ""
Case "EXTENDER"
.ComboList = ""
Case "GRINDER"
.ComboList = ""
End Select
End With
End Sub

Private Sub ISButton1_Click()
Dim X As Integer
    If txtid.text = "" Then Exit Sub
    X = MsgBox("Â·  —Ìœ ’Ê—… „‰ „·›", vbExclamation + vbYesNoCancel)

    If X = vbYes Then
        DBPix201.ImageLoad

        DoEvents
        MsgBox " „  Õ„Ì· «·’Ê—…"
    Else

        If X = vbNo Then
            DBPix201.TWAINAcquire
            MsgBox " „ „”Õ ÷Ê∆Ì  ··’Ê—…"

            DoEvents
        Else

            Exit Sub
        End If
    End If

    DBPix201.ImageSaveFile (App.path & "\images\pos\" & XPTxtID.text & ".JPG")
End Sub

Sub Calaulte()
If val(TxtEveryAmount.text) <> 0 Then
TxtEquation1.text = val(TxtEqualPiont.text) / val(TxtEveryAmount.text)
Else
TxtEquation1.text = 0
End If
If val(TxtEveryPoints.text) <> 0 Then
TxtEquation2.text = val(TxtEqualReal.text) / val(TxtEveryPoints.text)
Else
TxtEquation2.text = 0
End If
End Sub

Private Sub TreeGroups_MouseUp(Button As Integer, _
                               Shift As Integer, _
                               X As Single, _
                               Y As Single)
    On Error GoTo ErrTrap
    Dim tp            As POINTAPI
    Dim lX            As Single
    Dim lY            As Single
    Dim tr            As RECT
    Dim XNodeSeelcted As MSComctlLib.Node

    If Me.TreeGroups.SelectedItem Is Nothing Then
        Exit Sub
    End If

    'TxtMenuState_Change
    If Button = vbRightButton Then
'        GetCursorPos tp
'        lX = (tp.X) * Screen.TwipsPerPixelX
'        lY = tp.Y * Screen.TwipsPerPixelY
'        XPPopUp.PopupMenu "mnuDropMenu1", lX, lY
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub TreeGroups_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim NodeKey As String
    NodeKey = left(Node.Key, Len(Node.Key) - 1)

    If Me.TxtModFlg.text = "N" Or Me.TxtModFlg.text = "E" Then
        If right(Node.Key, 1) = "G" Then
        
            XPCboGroup.BoundText = val(Node.Key)
        
        End If

        Exit Sub
    End If

    If Node <> "" And NodeKey <> "1" Then
        Retrive (NodeKey)
    End If

End Sub

Private Sub TxtEqualPiont_Change()
Calaulte
End Sub

Private Sub TxtEqualReal_Change()
Calaulte
End Sub

Private Sub TxtEveryAmount_Change()
Calaulte
End Sub

Private Sub TxtEveryPoints_Change()
Calaulte
End Sub

Private Sub TxtMenuState_Change()
    'Select Case TxtMenuState.Text
    '    Case "C"
    '        Me.XPPopUp.Menus(1).MenuItems(7).Enabled = True
    '        Me.XPBtnMove(0).Enabled = False
    '        Me.XPBtnMove(1).Enabled = False
    '        Me.XPBtnMove(2).Enabled = False
    '        Me.XPBtnMove(3).Enabled = False
    '    Case "N"
    '        Me.XPPopUp.Menus(1).MenuItems(7).Enabled = False
    '        Me.XPBtnMove(0).Enabled = True
    '        Me.XPBtnMove(1).Enabled = True
    '        Me.XPBtnMove(2).Enabled = True
    '        Me.XPBtnMove(3).Enabled = True
    'End Select
End Sub

Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap

    Select Case Me.TxtModFlg.text

        Case "R"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«  «·„Ã„Ê⁄« "
            Else
                Me.Caption = "Items Groups"
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
        
            Me.XPTxtID.locked = True
            Me.XPTxtName.locked = True
            Me.XPCboGroup.locked = True

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

            TreeGroups.Enabled = True
        
            XPCboGroup.Enabled = False

        Case "N"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«  «·„Ã„Ê⁄« ( ÃœÌœ )"
            Else
                Me.Caption = "Items Groups(New Record)"
            End If
        
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
        
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(5).Enabled = False
            Me.Cmd(7).Enabled = False
            XPCboGroup.Enabled = True
            '     Me.XPBtnMove(0).Enabled = False
            '     Me.XPBtnMove(1).Enabled = False
            '     Me.XPBtnMove(2).Enabled = False
            '     Me.XPBtnMove(3).Enabled = False
            Me.XPTxtID.locked = True
            Me.XPTxtName.locked = False
            Me.XPCboGroup.locked = False
            '   TreeGroups.Enabled = False
            Me.lbl(10).Caption = ""
            Me.lbl(11).Caption = ""
            Me.lbl(12).Caption = ""
            Me.lbl(13).Caption = ""
            Me.lbl(14).Caption = ""
            Me.lbl(15).Caption = ""

        Case "E"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«  «·„Ã„Ê⁄« (  ⁄œÌ·)"
            Else
                Me.Caption = "Items Groups(Edit Record)"
            End If
        
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
        
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(5).Enabled = False
            Me.Cmd(7).Enabled = False
            XPCboGroup.Enabled = False
           TreeGroups.Enabled = False
            Me.XPBtnMove(0).Enabled = False
            Me.XPBtnMove(1).Enabled = False
            Me.XPBtnMove(2).Enabled = False
            Me.XPBtnMove(3).Enabled = False
            Me.XPTxtID.locked = True
            Me.XPTxtName.locked = False
            Me.XPCboGroup.locked = False
            '   TreeGroups.Enabled = False
            '        TxtMenuState.Text = "C"
    End Select

    Exit Sub
ErrTrap:
End Sub



Private Sub Grid_StartEdit(ByVal row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim rs As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim Msg As String
    With Grid
        Select Case .ColKey(Col)

            Case "Name"

                StrSQL = " select * from TblQCItems "
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = Grid.BuildComboList(rs, "name", "qcid")
                Else
                    StrComboList = Grid.BuildComboList(rs, "namee", "qcid")
                End If
                Debug.Print StrSQL
 
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
        End Select

    End With
End Sub

Private Sub XPBtnMove_Click(Index As Integer)
    On Error GoTo ErrTrap

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
    Exit Sub
ErrTrap:
End Sub

Public Sub Retrive(Optional Lngid As Long = 0)
    On Error GoTo ErrTrap

    If rs.RecordCount < 1 Then
        XPTxtCurrent.Caption = 0
        XPTxtCount.Caption = 0
        Exit Sub
    End If

    If rs.EOF Or rs.BOF Then
        Exit Sub
    End If

    If Lngid <> 0 Then
        rs.find "GroupID=" & Lngid, , adSearchForward, adBookmarkFirst

        If rs.EOF Or rs.BOF Then
            Exit Sub
        End If
    End If

    XPTxtID.text = IIf(IsNull(rs("GroupID").value), "", (rs("GroupID").value))
    XPTxtName.text = IIf(IsNull(rs("GroupName").value), "", Trim(rs("GroupName").value))
    XPTxtNameE.text = IIf(IsNull(rs("GroupNamee").value), "", Trim(rs("GroupNamee").value))
    Me.dcBranch.BoundText = IIf(IsNull(rs("BranchID").value), "", Trim(rs("BranchID").value))
    
    Me.DcActivityType.BoundText = IIf(IsNull(rs("ActivityTypeId").value), "", Trim(rs("ActivityTypeId").value))
        
        
    Me.cmbSpecification.BoundText = IIf(IsNull(rs("SpecificationID").value), "", Trim(rs("SpecificationID").value))
    
    Me.cmbAqrCompenet.BoundText = IIf(IsNull(rs("AqrCompenetID").value), "", Trim(rs("AqrCompenetID").value))
            
    ''//////////////////
    TxtEveryAmount.text = IIf(IsNull(rs("EveryAmount").value), 0, rs("EveryAmount").value)
    TxtEqualPiont.text = IIf(IsNull(rs("EqualPiont").value), 0, rs("EqualPiont").value)
    TxtEquation1.text = IIf(IsNull(rs("Equation1").value), 0, rs("Equation1").value)
    TxtEveryPoints.text = IIf(IsNull(rs("EveryPoints").value), 0, rs("EveryPoints").value)
    TxtEqualReal.text = IIf(IsNull(rs("EqualReal").value), 0, rs("EqualReal").value)
    TxtEquation2.text = IIf(IsNull(rs("Equation2").value), 0, rs("Equation2").value)
    If Not IsNull(rs("ParentID")) Then
        XPCboGroup.BoundText = rs("ParentID")
    Else
        XPCboGroup.text = ""
    End If

    If IsNull(rs("IsTransfere").value) Then
        Me.chkIsTransfere.value = vbUnchecked
    Else
        Me.chkIsTransfere.value = IIf(rs("IsTransfere").value = 0, vbUnchecked, vbChecked)
    End If



    Me.TxtGroupCode.text = IIf(IsNull(rs("GroupCode").value), "", Trim(rs("GroupCode").value))
    DCPreFix.text = IIf(IsNull(rs("prifix").value), "", rs("prifix").value)
    Me.txtid.text = IIf(IsNull(rs("code").value), "", rs("code").value)

    If XPTxtID.text <> "" Then
        DBPix201.ImageClear
 

        If Dir(App.path & "\images\pos\" & XPTxtID.text & ".JPG") <> "" Then
            DBPix201.ImageLoadFile (App.path & "\images\pos\" & XPTxtID.text & ".JPG")
        End If

 
 
    End If
        Dim i As Integer
        Dim mFieldName As String
        For i = 1 To dcClass.count
            mFieldName = "ClassId" & i
            dcClass(i).BoundText = val(rs(mFieldName) & "")
            
        Next
            

    CboEXpirType.ListIndex = IIf(IsNull(rs("EXpirType").value), -1, rs("EXpirType").value)

    If CboEXpirType.ListIndex = -1 Then
        TxtEXpireValue.text = ""
    Else
        Me.TxtEXpireValue.text = IIf(IsNull(rs("EXpireValue").value), "", rs("EXpireValue").value)
    End If
            Me.TxtOverHead.text = IIf(IsNull(rs("OverHead").value), 0, rs("OverHead").value)
            Me.CBOprintrs.text = IIf(IsNull(rs("GroupPrinterName").value), "", rs("GroupPrinterName").value)
            Me.txtMaxPercentDiscount.text = IIf(IsNull(rs("MaxPercentDiscount").value), 0, rs("MaxPercentDiscount").value)

    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount

    If rs("LastGroup").value = True Then
        Chklast.value = vbChecked
    Else
        Chklast.value = Unchecked
    End If
    
    If Not IsNull(rs("CritiCalStock").value) Then
        If rs("CritiCalStock").value = 1 Then
            CritiCalStock.value = vbChecked
        Else
            CritiCalStock.value = vbUnchecked
        End If
    Else
        CritiCalStock.value = vbUnchecked
    End If
        
        
    
    If Not IsNull(rs("IsShowCover").value) Then
        If CBool(rs("IsShowCover") & "") Then
            chkIsShowCover.value = vbChecked
        Else
            chkIsShowCover.value = vbUnchecked
        End If
    Else
        chkIsShowCover.value = vbUnchecked
    End If
        
        
        
            If rs("posgroup").value = True Then
        posgroup.value = vbChecked
    Else
        posgroup.value = Unchecked
    End If
    
 
        
     
            
        
    If rs("IsMaterial").value = 1 Then
        IsMaterial.value = vbChecked
    Else
        IsMaterial.value = Unchecked
    End If
    
    If rs("IsProducer").value Then
        chkIsProducer.value = vbChecked
    Else
        chkIsProducer.value = Unchecked
    End If
    
    If rs("IsAdditions").value Then
        chkIsAdditions.value = vbChecked
    Else
        chkIsAdditions.value = Unchecked
    End If
    
    If Not IsNull(rs("Separate").value) Then
    If rs("Separate").value = 1 Then
        ChSeparate.value = vbChecked
    Else
        ChSeparate.value = Unchecked
    End If
       Else
        ChSeparate.value = Unchecked
    End If
    
    If rs("HoldingMaterials").value = True Then
        ChkHoldingMaterials.value = vbChecked
    Else
        ChkHoldingMaterials.value = Unchecked
    End If
    
    
    If rs("chkTaxExempt").value = True Then
        chkTaxExempt.value = vbChecked
    Else
        chkTaxExempt.value = Unchecked
    End If
    
    
    GetGroupInfo val(XPTxtID.text)
    FillGrid
    
        Dim s As String
    
   s = " SELECT TblProductLine.ID,TblProductLine.name ProductLineName,TblGroupItemProductLine.Remarks from TblGroupItemProductLine"
   s = s & "  LEFT OUTER JOIN TblProductLine ON TblProductLine.Id = TblGroupItemProductLine.ProductLineId"
   s = s & " WHERE TblGroupItemProductLine.GroupID =" & val(XPTxtID.text)
   loadgrid s, grdProductLine, True, False
  
  
 
    Exit Sub
ErrTrap:
End Sub

Private Sub Undo()
    On Error GoTo ErrTrap

    Select Case TxtModFlg.text

        Case "N"
            clear_all Me
            Me.TxtModFlg.text = "R"
            XPBtnMove_Click (1)

        Case "E"
            rs.find "GroupID='" & val(XPTxtID.text) & "'", , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Me.TxtModFlg.text = "R"
                Exit Sub
            End If

            Retrive
            Me.TxtModFlg.text = "R"
    End Select

    FillGroupCombo

    If val(XPTxtID.text) <> 0 Then
        Me.Retrive (XPTxtID.text)
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub delete_group_account(group_id As Integer)
    Dim Msg As String
    Dim StrSQL As String
    Dim RsTemp As New ADODB.Recordset
    Dim RsTemp2 As New ADODB.Recordset
    Dim i As Integer
    'On Error GoTo ErrTrap
    'groups_account_in_inventory

    '  StrSQL = "select * From Notes where BoxID=" & Trim(XPTxtBoxID.text)
    StrSQL = "select * From groups_account_in_inventory where group_id=" & group_id
    RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Not (RsTemp.EOF Or RsTemp.BOF) Then

        For i = 1 To RsTemp.RecordCount
            StrSQL = "select * From DOUBLE_ENTREY_VOUCHERS where Account_Code='" & RsTemp("Account_Code").value & "'"
            RsTemp2.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

            If RsTemp2.RecordCount = 0 Then
                If ModAccounts.DeleteAccount(RsTemp("Account_Code").value) = True Then
       
                End If
                
            End If

            RsTemp2.Close
            RsTemp.MoveNext
        Next i
        
    End If

    StrSQL = "Delete From groups_account_in_inventory where group_id=" & group_id
    Cn.Execute StrSQL, , adExecuteNoRecords

End Sub

Private Sub Del_Group()
    Dim Msg As String
    Dim StrSQL As String
    Dim RsTemp As New ADODB.Recordset
    On Error GoTo ErrTrap

    If XPTxtID.text <> "" Then
    
        StrSQL = "select * From TblItems where GroupID=" & Trim(XPTxtID.text)
        RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

        If Not (RsTemp.EOF Or RsTemp.BOF) Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "·« Ì„ﬂ‰ Õ–› Â–Â «·„Ã„Ê⁄… " & CHR(13)
                Msg = Msg + "Â‰«ﬂ √’‰«›  ‰œ—Ã  Õ  Â–Â «·„Ã„Ê⁄…"
            Else
                Msg = "Can't Delete this Group because it have Items"
            End If

            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Exit Sub
        End If
       Cn.Execute "Delete from TblGroupQCItems where GroupID =" & val(XPTxtID.text) & ""
        RsTemp.Close
    
        StrSQL = "select * From Groups where ParentID=" & Trim(XPTxtID.text)
        RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

        If Not (RsTemp.EOF Or RsTemp.BOF) Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "·« Ì„ﬂ‰ Õ–› Â–Â «·„Ã„Ê⁄… " & CHR(13)
                Msg = Msg + "Â‰«ﬂ „Ã„Ê⁄«    ‰œ—Ã  Õ  Â–Â «·„Ã„Ê⁄…"
            Else
                Msg = "Can't Delete this Group because it have Chilrd Goup "
            End If

            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Exit Sub
        End If
    
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "”Ì „ Õ–› »Ì«‰«  «·„Ã„Ê⁄… —ﬁ„ " & CHR(13)
            Msg = Msg + (XPTxtID.text) & CHR(13)
            Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"
        Else
            Msg = "Confirm Delete Group " & CHR(13)
    
        End If
    
        If MsgBox(Msg, vbYesNo + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbNo Then
            Exit Sub
        End If

        If Not rs.RecordCount < 1 Then
            delete_group_account val(XPTxtID.text)
            CuurentLogdata ("D")
            rs.delete
            TreeGroups.Nodes.Remove (Trim(XPTxtID.text) & "G")
            rs.MoveFirst

            If rs.RecordCount < 1 Then
                FillGroupCombo
                Me.Retrive (XPTxtID.text)
                clear_all Me
                TxtModFlg_Change
                XPTxtCurrent.Caption = 0
                XPTxtCount.Caption = 0
            Else
                FillGroupCombo
                Retrive
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
        Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & CHR(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–Â «·„Ã„Ê⁄… "
        MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
        rs.CancelUpdate
    End If

End Sub

Private Sub AddTip()
    Dim Wrap As String
    Dim BolRtl As Boolean

    On Error GoTo ErrTrap
    Wrap = CHR(13) + CHR(10)
    Set TTP = New clstooltip

    If SystemOptions.UserInterface = ArabicInterface Then
        BolRtl = True

        With TTP
            .Create Me.hWnd, "»Ì«‰«  «·„Ã„Ê⁄« ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«  „Ã„Ê⁄… ÃœÌœ…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  «·„Ã„Ê⁄« ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(7), "ÿ»«⁄… ..." & Wrap & "·⁄—÷ «·»Ì«‰«  «·Õ«·Ì… ›Ì  ﬁ—Ì— " & Wrap & " Ì„ﬂ‰ ÿ»«⁄ Â ⁄‰ ÿ—Ìﬁ «·ÿ«»⁄…", True
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  «·„Ã„Ê⁄« ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  Â–Â «·„Ã„Ê⁄…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  «·„Ã„Ê⁄« ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  «·„Ã„Ê⁄… «·ÃœÌœ…" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  «·„Ã„Ê⁄« ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  «·„Ã„Ê⁄« ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  «·„Ã„Ê⁄…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  «·„Ã„Ê⁄« ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(5), "»ÕÀ ..." & Wrap & "···»ÕÀ ⁄‰ „Ã„Ê⁄…" & Wrap & "Ì‰ÿ»ﬁ ⁄·ÌÂ« ‘—Êÿ „⁄Ì‰…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  «·„Ã„Ê⁄« ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  «·„Ã„Ê⁄« ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  «·„Ã„Ê⁄« ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  «·„Ã„Ê⁄« ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  «·„Ã„Ê⁄« ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
        End With

        With TTP
            .Create Me.hWnd, "»Ì«‰«  «·„Ã„Ê⁄« ", 1, 15204351, -2147483630
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl CmdHelp, "„”«⁄œ… ..." & Wrap & "·· ⁄—› ⁄·Ï ÊŸÌ›… Â–Â «·‰«›–…" & Wrap & "ÊﬂÌ›Ì… «· ⁄«„· „⁄Â«" & Wrap & "≈÷€ÿ Â‰«" & Wrap, True
        End With

    ElseIf SystemOptions.UserInterface = EnglishInterface Then
        BolRtl = False

        With TTP
            .Create Me.hWnd, "Items Groups Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(0), "New Record" & Wrap & "Enter New Group..." & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Items Groups Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(7), "Print" & Wrap & "Print the current record data" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Items Groups Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(1), "Edit..." & Wrap & "Edit this record data" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Items Groups Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(2), "Save..." & Wrap & "Save the new record or save the edit" & Wrap & "in the current record" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Items Groups Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(3), "Undo" & Wrap & "Undo in the adding new record" & Wrap & "Or undo in the current editing" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Items Groups Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(4), "Delete ...." & Wrap & "Delete the current record." & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Items Groups Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(5), "Search..." & Wrap & "Search for an Items Group" & Wrap & "" & Wrap, BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Items Groups Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl Cmd(6), "Exit" & Wrap & "Close this window" & Wrap, BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Items Groups Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(1), "Frist Record" & Wrap & "Move to Frist Record" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Items Groups Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(0), "Previous" & Wrap & "Move to Previous Record" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Items Groups Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(3), "Next" & Wrap & "Move to Next Record" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Items Groups Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl XPBtnMove(2), "Last" & Wrap & "Move to Last Record" & Wrap & "", BolRtl
        End With

        With TTP
            .Create Me.hWnd, "Items Groups Data", 1, 15204351, -2147483630, BolRtl
            .MaxWidth = 4000
            .VisibleTime = 9000
            .DelayTime = 600
            .AddControl CmdHelp, "Help" & Wrap & "Show the Help File" & Wrap & "" & Wrap, BolRtl
        End With

    End If

    Exit Sub
ErrTrap:
End Sub

Function create_accounts(group_id As Integer, group_name As String) As Boolean
    Dim rsOut As New ADODB.Recordset
    Dim Current_case As Integer
    Set rsOut = New ADODB.Recordset
    rsOut.Open "[TblOptions]", Cn, adOpenStatic, adLockOptimistic, adCmdTable

    If Not (rsOut.EOF Or rsOut.BOF) Then
        If rsOut!opt_group = False Then
            Current_case = -1
        ElseIf rsOut!opt_group = True And rsOut!Opt_Inventory_create_account = 1 Then
            Current_case = 0 '„Œ«“‰ ›ﬁÿ
        ElseIf rsOut!opt_group = True And rsOut!opt_inv_and_branch_create_account = 1 Then
            Current_case = 1 '„Œ«“‰ Ê›—⁄
        End If
    End If

    If Current_case = -1 Then Exit Function
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String
    Dim i As Integer
    sql = "Select * from TblStore "
 
    Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If Rs3.RecordCount = 0 Then Exit Function

    For i = 1 To Rs3.RecordCount

        If create_inventory_group(Rs3("StoreID").value, group_id, Rs3("StoreName").value & " " & group_name) = True Then
        End If

        Rs3.MoveNext
    Next i

    Rs3.Close

    Select Case Current_case

        Case 1:
            sql = "Select * from branches "
 
            Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
            If Rs3.RecordCount = 0 Then Exit Function

            For i = 1 To Rs3.RecordCount

                If create_Branch_group(Rs3("branch_id").value, group_id, group_name) = True Then
                End If

                Rs3.MoveNext
            Next i

            Rs3.Close

    End Select

End Function

Private Sub SaveData()
    Dim RsTemp As New ADODB.Recordset
    Dim StrSQL As String
    Dim Msg As String
    Dim BeginTrans As Boolean
    Dim XNode As MSComctlLib.Node
    On Error GoTo ErrTrap

    If Me.TxtModFlg.text <> "R" Then
        If XPTxtName.text = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "„‰ ›÷·ﬂ √œŒ· «”„ «·„Ã„Ê⁄…"
            Else
                Msg = "plz enter group name firstly"
            End If

            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            XPTxtName.SetFocus
            Exit Sub
        End If

        If XPCboGroup.text = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = " ÌÃ»  ÕœÌœ «·„Ã„Ê⁄… «·—∆Ì”Ì…" & CHR(13)
            Else
                Msg = "Select primary group First" & CHR(13)
            End If

            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            XPCboGroup.SetFocus
            Sendkeys "{F4}"
            Exit Sub
        End If
   
        Me.TxtGroupCode.text = GetNewGroupCode(Me.XPCboGroup.BoundText)

        Select Case TxtModFlg.text

            Case "N"
                StrSQL = "select * From Groups where GroupName='" & Trim(XPTxtName.text) & "'"
                RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                If RsTemp.RecordCount > 0 Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        Msg = " ÊÃœ „Ã„Ê⁄… „”Ã·… „”»ﬁ« »Â–« «·«”„" & CHR(13)
                        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·«”„ «·’ÕÌÕ " & CHR(13)
                        Msg = Msg + "√Ê  €ÌÌ— √Ê  „ÌÌ“ «”„ «·„Ã„Ê⁄…"

                    Else
                        Msg = "This group Name Already Exisi" & CHR(13)

                    End If

                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                    Exit Sub
                End If
            
                XPTxtID.text = CStr(new_id("Groups", "GroupID", "", True))

            Case "E"
                StrSQL = "select * From Groups where GroupName='" & Trim(XPTxtName.text) & "'"
                RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                If RsTemp.RecordCount > 0 Then
                    If RsTemp("GroupID").value <> val(XPTxtID.text) Then
                        If SystemOptions.UserInterface = ArabicInterface Then
                            Msg = " ÊÃœ „Ã„Ê⁄… „”Ã·… „”»ﬁ« »Â–« «·«”„" & CHR(13)
                            Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·«”„ «·’ÕÌÕ " & CHR(13)
                            Msg = Msg + "√Ê  €ÌÌ— √Ê  „ÌÌ“ «”„ «·„Ã„Ê⁄…"
                        Else
                            Msg = "This group Name Already Exisi" & CHR(13)
                        End If

                        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                        Exit Sub
                    End If
                End If


If val(XPCboGroup.BoundText) = val(XPTxtID.text) Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                            Msg = "·« Ì„ﬂ‰ —»ÿ «·„Ã„Ê⁄Â »‰›”Â« " & CHR(13)
                            Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·„Ã„Ê⁄Â «·—∆Ì”Ì…   " & CHR(13)
                             
                        Else
                            Msg = "Can't link  Group with it self" & CHR(13)
                        End If

                        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                        Exit Sub

End If



        End Select
     If Me.TxtModFlg.text = "E" Then
     Cn.Execute "Delete from TblGroupQCItems where GroupID =" & val(XPTxtID.text) & ""
     End If
     
        Select Case TxtModFlg.text

            Case "N"
                Cn.BeginTrans
                BeginTrans = True
            
                rs.AddNew
                rs("GroupID").value = IIf(XPTxtID.text = "", "", val(XPTxtID.text))

                If Chklast.value = vbChecked Then
                    If create_accounts(XPTxtID.text, XPTxtName.text) Then
                
                    End If
                End If
            
            Case "E"

                If XPTxtName.text = XPCboGroup.text Then
                    Msg = "·«Ì„ﬂ‰ √‰  ﬂÊ‰ «·„Ã„Ê⁄… «·—∆Ì”Ì… " & CHR(13)
                    Msg = Msg + "ÂÌ ‰›” «·„Ã„Ê⁄… «·›—⁄Ì…"
                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                    Exit Sub
                End If

                Cn.BeginTrans
                BeginTrans = True
        End Select

        rs("GroupName").value = IIf(XPTxtName.text = "", "", Trim(XPTxtName.text))
        rs("GroupNamee").value = IIf(XPTxtNameE.text = "", "", Trim(XPTxtNameE.text))
        rs("ParentID").value = XPCboGroup.BoundText
        rs("GroupCode").value = Me.TxtGroupCode.text
        rs("Branch_NO").value = val(branch_id)
        rs("BranchID").value = val(Me.dcBranch.BoundText)
        rs("ActivityTypeId").value = val(Me.DcActivityType.BoundText)
        
        rs("SpecificationID").value = val(Me.cmbSpecification.BoundText)
        rs("AqrCompenetID").value = val(Me.cmbAqrCompenet.BoundText)
        
        ''/////////
        rs("EveryAmount").value = val(Me.TxtEveryAmount.text)
        rs("EqualPiont").value = val(TxtEqualPiont.text)
        rs("Equation1").value = val(TxtEquation1.text)
        rs("EveryPoints").value = val(TxtEveryPoints.text)
        rs("EqualReal").value = val(TxtEqualReal.text)
        rs("Equation2").value = val(TxtEquation2.text)
        ''////////
        If CritiCalStock.value = vbChecked Then
            rs("CritiCalStock").value = 1
        Else
            rs("CritiCalStock").value = 0
        End If
        
        
        If chkIsTransfere.value = vbChecked Then
            rs("IsTransfere").value = 1
        Else
            rs("IsTransfere").value = 0
        End If
    

        If chkIsShowCover.value = vbChecked Then
            rs("IsShowCover").value = 1
        Else
            rs("IsShowCover").value = 0
        End If
        
        
        
        
        rs("code").value = txtid.text
        rs("Fullcode").value = IIf(DCPreFix.BoundText = "", Null, DCPreFix.text) & IIf(Trim(txtid.text) = "", Null, txtid.text)
        rs("prifix").value = IIf(DCPreFix.text = "", Null, DCPreFix.text)
  
        If CboEXpirType.ListIndex = -1 Then
            rs("EXpirType").value = Null
            rs("EXpireValue").value = Null
        Else
            rs("EXpirType").value = (CboEXpirType.ListIndex)
            rs("EXpireValue").value = val(TxtEXpireValue.text)
        End If
  
        If ChSeparate.value = vbChecked Then
            rs("Separate").value = 1
        Else
            rs("Separate").value = 0
        End If
        
            If posgroup.value = vbChecked Then
            rs("posgroup").value = 1
        Else
            rs("posgroup").value = 0
        End If
        
          rs("GroupPrinterName").value = (CBOprintrs.text)
          
       If Chklast.value = vbChecked Then
            rs("LastGroup").value = 1
        Else
            rs("LastGroup").value = 0
        End If
                If IsMaterial.value = vbChecked Then
            rs("IsMaterial").value = 1
        Else
            rs("IsMaterial").value = 0
        End If
        
 
        If chkIsAdditions.value = vbChecked Then
            rs("IsAdditions").value = 1
        Else
            rs("IsMaterial").value = 0
        End If
        
 
        If chkIsProducer.value = vbChecked Then
            rs("IsProducer").value = 1
        Else
            rs("IsProducer").value = 0
        End If
          
 
 
                If ChkHoldingMaterials.value = vbChecked Then
            rs("HoldingMaterials").value = 1
        Else
            rs("HoldingMaterials").value = 0
        End If
         
        
 
        If chkTaxExempt.value = vbChecked Then
            rs("chkTaxExempt").value = 1
        Else
            rs("chkTaxExempt").value = 0
        End If
          
         
    
            rs("OverHead").value = val(TxtOverHead.text)
            rs("MaxPercentDiscount").value = val(txtMaxPercentDiscount.text)
            
            
        Dim i As Integer
        Dim mFieldName As String
        For i = 1 To dcClass.count
            mFieldName = "ClassId" & i
            rs(mFieldName) = val(dcClass(i).BoundText)
        Next
            
            
        rs.update
        Dim rs2 As ADODB.Recordset
        Set rs2 = New ADODB.Recordset
        Dim sql As String
       ' Dim i As Integer
        sql = "select * from TblGroupQCItems where 1=-1"
        rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
 With Grid
        For i = .FixedRows To .Rows - 2
            If val(.TextMatrix(i, .ColIndex("GQCID"))) <> 0 Then
                rs2.AddNew
                rs2("GroupID").value = val(XPTxtID.text)
                rs2("GQCID").value = val(.TextMatrix(i, .ColIndex("GQCID")))
                rs2("Standrd").value = .TextMatrix(i, .ColIndex("Standrd"))
                rs2("Comment").value = .TextMatrix(i, .ColIndex("Comment"))
                If .Cell(flexcpChecked, i, .ColIndex("GRINDER")) = flexChecked Then
                rs2("GRINDER").value = 1
                Else
                rs2("GRINDER").value = 0
                End If
                If .Cell(flexcpChecked, i, .ColIndex("EXTENDER")) = flexChecked Then
                rs2("EXTENDER").value = 1
                Else
                rs2("EXTENDER").value = 0
                End If
                rs2.update
            End If
        Next i
    End With
    
        Dim s As String
        s = "Delete TblGroupItemProductLine Where GroupID = " & val(Me.XPTxtID.text)
        Cn.Execute s
        s = "Select * from TblGroupItemProductLine Where GroupID = " & val(Me.XPTxtID.text)
        saveGrid s, grdProductLine, "ProductLineID", "", "GroupID", val(Me.XPTxtID.text)
        
        
        
        Cn.CommitTrans
        BeginTrans = False
        XPTxtCurrent.Caption = rs.AbsolutePosition
        XPTxtCount.Caption = rs.RecordCount
        FillGroupCombo

        If TxtModFlg.text = "E" Then
            'TreeGroups.Nodes.Remove (Trim(Rs("GroupID").Value) & "G")
            TreeGroups.Nodes(Trim(rs("GroupID").value) & "G").text = rs("GroupName").value
        ElseIf TxtModFlg.text = "N" Then
            Set XNode = TreeGroups.Nodes.Add(Trim(rs("ParentID").value) & "G", tvwChild, Trim(rs("GroupID").value) & "G", Trim(rs("GroupName").value), "Closed_Node", "Open_Node")
            TreeGroups.Nodes(Trim(rs("GroupID").value) & "G").Selected = True
        End If
        Dim strsq As String
If Me.chkTaxExempt.value = vbChecked Then

strsq = "update TblItems set ItemWithOutVAT=2 where GroupID=" & val(XPTxtID.text) & "  and isnull(ItemWithOutVAT,0)<>1  "
Cn.Execute strsq
Else
strsq = "update TblItems set ItemWithOutVAT=0 where GroupID=" & val(XPTxtID.text) & "  and isnull(ItemWithOutVAT,0)<>1  "
Cn.Execute strsq
End If
        Me.Retrive (XPTxtID.text)
        CuurentLogdata

        Select Case Me.TxtModFlg.text

            Case "N"

                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·„Ã„Ê⁄…" & CHR(13)
                    Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
                Else
                    Msg = " Data was Saved , do you want to enter another data y/n" & CHR(13)
                End If

                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                    Cmd_Click (0)
                    Exit Sub
                End If
            
            Case "E"
    
                TreeGroups.ImageList = mdifrmmain.ImgLstTree
                LoadTreeGroups Me.TreeGroups

                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Else
                    MsgBox "Changes Saved", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        
                End If

        End Select

        TxtModFlg.text = "R"
        TreeGroups.SetFocus
    End If

    Exit Sub
ErrTrap:

    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If

    If Err.Number = -2147217900 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
            Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & CHR(13)
            Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
        Else
            Msg = "Can't Save error in entered data " & CHR(13)
        End If

        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

    If Err.Number = -2147217887 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·⁄„·Ì… " & CHR(13)
            Msg = Msg + "· ﬂ«„· «·»Ì«‰« " & CHR(13)
        Else
            Msg = "Can't save Data , Reasons: Data integrity " & CHR(13)
 
        End If

        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        rs.CancelUpdate
        Exit Sub
    End If

    Select Case Me.TxtModFlg.text

        Case "N"

            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
            Else
                Msg = "Sorry...... Error During Saving Data " & CHR(13)
            End If

            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title

        Case "E"

            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «· ⁄œÌ·«  " & CHR(13)
            Else
                Msg = "Sorry...... Error During Saving cahanges" & CHR(13)
            End If

            MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    End Select

End Sub
Sub FillGrid()
Dim sql As String
Dim i As Integer
Dim rs2 As ADODB.Recordset
Set rs2 = New ADODB.Recordset
Grid.Clear flexClearScrollable, flexClearEverything
Grid.Rows = 1
sql = " SELECT     dbo.TblGroupQCItems.Standrd, dbo.TblGroupQCItems.GroupID, dbo.TblGroupQCItems.GQCID, dbo.TblQCItems.name, dbo.TblQCItems.namee ,dbo.TblGroupQCItems.Comment ,dbo.TblGroupQCItems.GRINDER,dbo.TblGroupQCItems.EXTENDER"
sql = sql & " FROM         dbo.TblGroupQCItems LEFT OUTER JOIN"
sql = sql & "                      dbo.TblQCItems ON dbo.TblGroupQCItems.GQCID = dbo.TblQCItems.qcid"
sql = sql & " Where (dbo.TblGroupQCItems.GroupID =" & val(XPTxtID.text) & " )"
rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If rs2.RecordCount > 0 Then
With Grid
rs2.MoveFirst
.Rows = .Rows + rs2.RecordCount
For i = 1 To .Rows - 1
.TextMatrix(i, .ColIndex("Ser")) = i
.TextMatrix(i, .ColIndex("GRINDER")) = IIf(IsNull(rs2("GRINDER").value), 0, rs2("GRINDER").value)
.TextMatrix(i, .ColIndex("EXTENDER")) = IIf(IsNull(rs2("EXTENDER").value), 0, rs2("EXTENDER").value)
.TextMatrix(i, .ColIndex("GQCID")) = IIf(IsNull(rs2("GQCID").value), "", rs2("GQCID").value)
.TextMatrix(i, .ColIndex("Standrd")) = IIf(IsNull(rs2("Standrd").value), "", rs2("Standrd").value)
.TextMatrix(i, .ColIndex("Comment")) = IIf(IsNull(rs2("Comment").value), "", rs2("Comment").value)
If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs2("name").value), "", rs2("name").value)
Else
.TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs2("namee").value), "", rs2("namee").value)
End If
rs2.MoveNext
Next i
End With
End If
End Sub

Function print_report2(Optional NoteSerial As String)
    
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
MySQL = " SELECT     dbo.Groups.GroupID, dbo.Groups.GroupName, dbo.Groups.GroupCode, dbo.Groups.Fullcode, dbo.Groups.EXpirType, dbo.Groups.prifix, dbo.Groups.EXpireValue, "
MySQL = MySQL & "                      dbo.Groups.GroupNamee, dbo.Groups.OverHead, dbo.Groups.LastGroup, dbo.Groups.code, dbo.Groups.Branch_NO, dbo.Groups.ParentID,"
MySQL = MySQL & "                       Groups_1.GroupName AS ParGroupName, Groups_1.GroupNamee AS ParGroupNameE, Groups_1.code AS Parcode, Groups_1.Fullcode AS ParFullcode"
MySQL = MySQL & "  FROM         dbo.Groups LEFT OUTER JOIN"
MySQL = MySQL & "                       dbo.Groups Groups_1 ON dbo.Groups.ParentID = Groups_1.GroupID"
MySQL = MySQL & "  Where (dbo.Groups.GroupID =" & val(XPTxtID.text) & ")"


 If SystemOptions.UserInterface = ArabicInterface Then
          StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepGroup.rpt"
     Else
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepGroup.rpt"
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

    xReport.ParameterFields(3).AddCurrentValue user_name
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
'Private Sub PrintReport()
'    On Error GoTo ErrTrap
'
'    If XPTxtID.text <> "" Then
'        Set GroupReport = New ClsGroupReport
'        GroupReport.GroupData XPTxtID.text
'    End If

'    Exit Sub
'ErrTrap:
'End Sub

Private Sub FillGroupCombo()
    On Error GoTo ErrTrap
    Dim Num As Integer
    Dim StrSQL As String
    Dim RsTemp As New ADODB.Recordset
  If SystemOptions.UserInterface = ArabicInterface Then
    StrSQL = "SELECT GroupID,   isnull(fullcode,' ') + ' ' +  isnull(GroupName,'')    as grupname FROM groups Order By GroupName"
   Else
   StrSQL = "SELECT GroupID,  isnull(fullcode,' ') + ' ' +  isnull(GroupNamee,'')  as grupnamee FROM groups Order By GroupNamee"
   End If
   
    RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    fill_combo XPCboGroup, StrSQL
    Set cSearchDcbo = New clsDCboSearch
    Set cSearchDcbo.Client = XPCboGroup
    cSearchDcbo.Refresh
    'XPCboGroup.Clear
    'If Not (RsTemp.EOF Or RsTemp.BOF) Then
    '    RsTemp.MoveFirst
    '    'XPCboGroup
    '    For Num = 0 To RsTemp.RecordCount - 1
    '        XPCboGroup.AddItem RsTemp("GroupName").Value, Num
    '        XPCboGroup.ItemData(Num) = RsTemp("GroupID").Value
    '        RsTemp.MoveNext
    '    Next Num
    '    RsTemp.MoveFirst
    'End If
    Exit Sub
ErrTrap:
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)

    Dim IntResult As String
    Dim StrMSG As String
    On Error GoTo ErrTrap

    If Me.TxtModFlg.text <> "R" Then

        Select Case Me.TxtModFlg.text

            Case "N"
    
                If SystemOptions.UserInterface = EnglishInterface Then
                    StrMSG = "You will close this screen before save " & CHR(13)
                    StrMSG = StrMSG & " the new data  " & CHR(13)
                    StrMSG = StrMSG & " do you want save before exit" & CHR(13)
                    StrMSG = StrMSG & "yes" & "-" & "save the new data" & CHR(13)
                    StrMSG = StrMSG & "no" & "-" & "Don't save" & CHR(13)
                    StrMSG = StrMSG & "cancel" & "-" & "Cancel Exit" & CHR(13)
    
                Else
                    StrMSG = "”Ê› Ì „ €·ﬁ «·‘«‘… Ê·„  ‰ Â „‰  ”ÃÌ·" & CHR(13)
                    StrMSG = StrMSG & " «·»Ì«‰«  «·ÃœÌœ… «·Õ«·Ì… " & CHR(13)
                    StrMSG = StrMSG & " Â·  —Ìœ «·Õ›Ÿ ﬁ»· «·Œ—ÊÃ" & CHR(13)
                    StrMSG = StrMSG & "‰⁄„" & "-" & "Ì „ Õ›Ÿ «·»Ì«‰«  «·ÃœÌœ…" & CHR(13)
                    StrMSG = StrMSG & "·«" & "-" & "·‰ Ì „ «·Õ›Ÿ" & CHR(13)
                    StrMSG = StrMSG & "≈·€«¡ «·√„—" & "-" & "≈·€«¡ ⁄„·Ì… «·Œ—ÊÃ" & CHR(13)
        
                End If
        
            Case "E"

                If SystemOptions.UserInterface = EnglishInterface Then
                    StrMSG = "You will close this screen before save  " & CHR(13)
                    StrMSG = StrMSG & " the Modifications  " & CHR(13)
                    StrMSG = StrMSG & " do you want save before exit" & CHR(13)
                    StrMSG = StrMSG & "yes" & "-" & "save the new data" & CHR(13)
                    StrMSG = StrMSG & "no" & "-" & "Don't save" & CHR(13)
                    StrMSG = StrMSG & "cancel" & "-" & "Cancel Exit" & CHR(13)
    
                Else
                    StrMSG = "”Ê› Ì „ €·ﬁ «·‘«‘… Ê·„  ‰ Â „‰  ”ÃÌ·" & CHR(13)
                    StrMSG = StrMSG & " «· ⁄œÌ·«  «·ÃœÌœ… ⁄·Ï «·”Ã· «·Õ«·Ï " & CHR(13)
                    StrMSG = StrMSG & " Â·  —Ìœ «·Õ›Ÿ ﬁ»· «·Œ—ÊÃ" & CHR(13)
                    StrMSG = StrMSG & "‰⁄„" & "-" & "Ì „ Õ›Ÿ «· ⁄œÌ·«   «·ÃœÌœ…" & CHR(13)
                    StrMSG = StrMSG & "·«" & "-" & "·‰ Ì „ «·Õ›Ÿ" & CHR(13)
                    StrMSG = StrMSG & "≈·€«¡ «·√„—" & "-" & "≈·€«¡ ⁄„·Ì… «·Œ—ÊÃ" & CHR(13)
                
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

Private Sub LoadTreeGroups(ItemsTree As MSComctlLib.TreeView)
    Dim Rs_items As ADODB.Recordset
    Dim My_SQL As String
    Dim nodX As Node
    Dim nodz As Node
    Dim RsOptions As ADODB.Recordset
    Dim my_ch_rs As ADODB.Recordset
    Dim BolDisplayArabic As Boolean
    Dim LngLoop As Long
    On Error GoTo ErrTrap
    Dim lngCount As Long
    Dim kRS As New ADODB.Recordset
    Dim k_SQL As String
    Dim test As Integer
    With ItemsTree.Nodes

        For lngCount = .count To 1 Step -1
            .Remove lngCount
        Next

    End With
    
    k_SQL = "select min(Groups.GroupID) as minGroupId from groups "
    kRS.Open k_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    test = kRS("minGroupId")
    If kRS("minGroupId") = 0 Then
      If SystemOptions.UserInterface = ArabicInterface Then
          BolDisplayArabic = True
          ItemsTree.Tag = "A"
          Make_RightToLeft ItemsTree
          '''''''''''''''''''''''''''add root
          Set nodX = ItemsTree.Nodes.Add(, , "0G", "„Ã„Ê⁄… «·√’‰«›", "Root")
          ItemsTree.Nodes("0G").Expanded = True
      Else
          BolDisplayArabic = False
          '''''''''''''''''''''''''''add root
          ItemsTree.Tag = "E"
          Set nodX = ItemsTree.Nodes.Add(, , "0G", "Groups of items", "Root")
          ItemsTree.Nodes("0G").Expanded = True
      End If

      Me.TreeGroups.Sorted = False
      '''''''''''''''''''''''''''' add group
      My_SQL = " SELECT Groups.* "
      My_SQL = My_SQL + "  From Groups "
      My_SQL = My_SQL + " where (ParentID = 0) "
      Set my_ch_rs = New ADODB.Recordset
      my_ch_rs.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
      'BolDisplayArabic = True

      If BolDisplayArabic = True Then
          Call fill_my_Children("0G", my_ch_rs, ItemsTree, "GROUPS", "ParentID")
      Else
          Call fill_my_Children("0G", my_ch_rs, ItemsTree, "GROUPS", "ParentID", , 11)
      End If
      Else
        If SystemOptions.UserInterface = ArabicInterface Then
        BolDisplayArabic = True
        ItemsTree.Tag = "A"
        Make_RightToLeft ItemsTree
        '''''''''''''''''''''''''''add root
        Set nodX = ItemsTree.Nodes.Add(, , "1G", "„Ã„Ê⁄… «·√’‰«›", "Root")
        ItemsTree.Nodes("1G").Expanded = True
      Else
        BolDisplayArabic = False
        '''''''''''''''''''''''''''add root
        ItemsTree.Tag = "E"
        Set nodX = ItemsTree.Nodes.Add(, , "1G", "Groups of items", "Root")
        ItemsTree.Nodes("1G").Expanded = True
      End If
      Me.TreeGroups.Sorted = False
      '''''''''''''''''''''''''''' add group
      My_SQL = " SELECT Groups.* "
      My_SQL = My_SQL + "  From Groups "
      My_SQL = My_SQL + " where (ParentID = 1) "
      Set my_ch_rs = New ADODB.Recordset
      my_ch_rs.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
      'BolDisplayArabic = True

      If BolDisplayArabic = True Then
        Call fill_my_Children("1G", my_ch_rs, ItemsTree, "GROUPS", "ParentID")
      Else
        Call fill_my_Children("1G", my_ch_rs, ItemsTree, "GROUPS", "ParentID", , 11)
      End If
    End If
    ItemsTree.Refresh
    Exit Sub
ErrTrap:
End Sub

Private Sub LoadMenus()
    On Error GoTo ErrTrap

    With Me.XPPopUp
        '~~Clear the Menu and ToolBars
        .ClearAll

        If SystemOptions.UserInterface = ArabicInterface Then
            .RightToLeft = True
            .SetImageList mdifrmmain.img16

            With .Menus.Add("mnuDropMenu1", tsSecondaryMenu, True)
                .MenuItems.Add tsMenuCaption, "...≈÷«›… „Ã„Ê⁄…", False, True, 2, , 2, , , "AddGroup", , , , "≈÷«›… „Ã„Ê⁄…"
                .MenuItems.Add tsMenuCaption, " ⁄œÌ·", False, True, 3, , , , , "EditGroup", , , , " ⁄œÌ·"
                .MenuItems.Add tsMenuCaption, "Õ–›", False, True, 4, , , , , "DelGroup", , , , "Õ–› «·„Ã„Ê⁄…"
                .MenuItems.Add tsMenuCaption, "„”Õ «·«Œ Ì«—", False, False, 5, , , , , "ClearGroup", , , , "„”Õ «·«Œ Ì«—"
                .MenuItems.Add tsMenuSeperator
                .MenuItems.Add tsMenuCaption, "ﬁ’", False, False, 6, , , True, , "CutGroup", , , , "ﬁ’"
                .MenuItems.Add tsMenuCaption, "·’ﬁ", False, False, 7, , , , , "PasteGroup", , , , "·’ﬁ"
                .MenuItems.Add tsMenuSeperator
                .MenuItems.Add tsMenuCaption, "‰ﬁ· √’‰«› «·„Ã„Ê⁄… ≈·Ï ", False, False, 8, , , True, , "RemoveGroup", , , , "‰ﬁ· √’‰«› «·„Ã„Ê⁄…"
                .MenuItems.Add tsMenuCaption, "Œ’«∆’", False, False, 9, , , True, , "GroupProperties", , , , "Œ’«∆’"
                .MenuItems.Add tsMenuSeperator
                .MenuItems.Add tsMenuCaption, "ÿ»«⁄… ‘Ã—… «·√’‰«›", False, False, 10, , , True, , "PrintGroup", , , , "ÿ»«⁄…"
            End With

        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            .RightToLeft = False
            .SetImageList mdifrmmain.img16

            With .Menus.Add("mnuDropMenu1", tsSecondaryMenu, True)
                .MenuItems.Add tsMenuCaption, "Add New Group...", False, True, 2, , 2, , , "AddGroup", , , , "Add New Group"
                .MenuItems.Add tsMenuCaption, "Edit", False, True, 3, , , , , "EditGroup", , , , "Edit this Group"
                .MenuItems.Add tsMenuCaption, "Delete", False, True, 4, , , , , "DelGroup", , , , "Delete This Group"
                .MenuItems.Add tsMenuCaption, "Clear Checked", False, False, 5, , , , , "ClearGroup", , , , "Clear This Group"
                .MenuItems.Add tsMenuSeperator
                .MenuItems.Add tsMenuCaption, "Cut", False, False, 6, , , True, , "CutGroup", , , , "Cut"
                .MenuItems.Add tsMenuCaption, "Paste", False, False, 7, , , , , "PasteGroup", , , , "Paste"
                .MenuItems.Add tsMenuSeperator
                .MenuItems.Add tsMenuCaption, "Move this items group to...", False, False, 8, , , True, , "RemoveGroup", , , , "Move the items of this group to another group"
                .MenuItems.Add tsMenuCaption, "Properties", False, False, 9, , , True, , "GroupProperties", , , , "Properties"
                .MenuItems.Add tsMenuSeperator
                .MenuItems.Add tsMenuCaption, "Print Items Tree", False, False, 10, , , True, , "PrintGroup", , , , "Print all of the tree of items"
            End With

        End If

    End With

    Exit Sub
ErrTrap:
End Sub

Private Sub XPPopUp_MenuItemClick(ByVal MenuIndex As Integer, _
                                  ByVal MenuID As String, _
                                  ByVal MenuItemIndex As Integer, _
                                  ByVal MenuItemID As String)
    On Error GoTo ErrTrap
    Dim XNode As MSComctlLib.Node
    Dim StrSQL As String

    Select Case MenuItemID

        Case "AddGroup"
            Cmd_Click (0)
            Set XNode = TreeGroups.SelectedItem
            XPCboGroup.BoundText = left(XNode.Key, Len(XNode.Key) - 1)

        Case "EditGroup"
            Cmd_Click (1)

        Case "DelGroup"
            Cmd_Click (4)

        Case "ClearGroup"

        Case "CutGroup"
            TreeGroups.SelectedItem.backcolor = vbGreen
            TxtCutKey.text = (TreeGroups.SelectedItem.Key)

            '        Me.TxtMenuState.Text = "C"
        Case "PasteGroup"
            TreeGroups.Nodes.Remove (TxtCutKey.text)
            Set XNode = TreeGroups.Nodes.Add(Trim(TreeGroups.SelectedItem.Key), tvwChild, rs("GroupID") & "G", rs("GroupName"), "Closed_Node", "Open_Node")
            StrSQL = "update Groups set ParentID=" & val(left(TreeGroups.SelectedItem.Key, Len(TreeGroups.SelectedItem.Key) - 1)) & " where GroupID=" & val(rs("GroupID").value)
            Cn.Execute StrSQL
            Retrive (val(rs("GroupID").value))

            '        Me.TxtMenuState.Text = "N"
        Case "GroupProperties"

        Case "PrintGroup"
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub ChangeLang()
    Dim XPic As IPictureDisp
ChSeparate.Caption = "Separate"
ChSeparate.RightToLeft = False
posgroup.Caption = "POS Group"
CritiCalStock.RightToLeft = False
CritiCalStock.Caption = "Critical Stock"
'lbl(24).Caption = "Branch"
    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic
IsMaterial.Caption = "Dash Board"
    Set XPic = Me.XPBtnMove(0).ButtonImage
    Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    Set Me.XPBtnMove(3).ButtonImage = XPic
    lbl(18).Caption = "Validitation"
    Me.Caption = "Items Groups"
    Me.LblHeader.Caption = Me.Caption
    Chklast.Caption = "Last Group"
    Me.lbl(6).Caption = "Group ID"
    Me.lbl(17).Caption = "Group Code"
    Me.lbl(5).Caption = "Arabic Name"
    lbl(23).Caption = "Printer"
    lbl(39).Caption = "Max Disc"
    lbl(37).Caption = "Filter"
    C1Elastic5.Caption = "Calculate customer points"
    lbl(31).Caption = "Equal"
    lbl(36).Caption = "Equal"
    lbl(30).Caption = "Point"
     lbl(29).Caption = "Point"
    lbl(32).Caption = "SR"
    lbl(27).Caption = "SR"
    'lbl(37).Caption = "Filter"
     With grdProductLine
        .TextMatrix(0, .ColIndex("ProductLineName")) = "Product Line"
        .TextMatrix(0, .ColIndex("Remarks")) = "Remarks"
        
    End With
    chkIsProducer.Caption = "Producer"
    chkIsAdditions.Caption = "Additions"
        Me.lbl(20).Caption = "English Name"
        ChkHoldingMaterials.Caption = "Carried Items"
        chkTaxExempt.Caption = "Tax Exempt"
    Me.lbl(0).Caption = "Parent Group"
    Me.lbl(1).Caption = "Current Record"
    lbl(25).Caption = "For all"
    lbl(26).Caption = "for"
    lbl(34).Caption = "equation"
    lbl(35).Caption = "equation"
    
    Me.lbl(2).Caption = "NO. Recordes"
lbl(19).Caption = "Image"
ISButton1.Caption = "Add Image"
lbl(21).Caption = "OverHead"

    'Ele.Caption = "More Information"
    'ELe.Font.Bold = True
    Me.lbl(4).Caption = "Sub Groups Count"
    Me.lbl(7).Caption = "Items Group Count"
    Me.lbl(8).Caption = "Frist item added to the group"
    Me.lbl(9).Caption = "Last item added to the group"
    
    Me.Cmd(0).Caption = "New"
    Me.Cmd(1).Caption = "Edit"
    Me.Cmd(2).Caption = "Save"
    Me.Cmd(3).Caption = "Undo"
    Me.Cmd(4).Caption = "Delete"
    Me.Cmd(5).Caption = "Search"
    Me.Cmd(6).Caption = "Exit"
    Me.Cmd(7).Caption = "Print"
    Me.CmdHelp.Caption = "Help"
    DeleteRow.Caption = "Delete"
    DeleteAll.Caption = "Delete All"
 With Me.Grid
  .TextMatrix(0, .ColIndex("Ser")) = "Serial"
  .TextMatrix(0, .ColIndex("Name")) = "Description"
  .TextMatrix(0, .ColIndex("Standrd")) = "Standrd"
  .TextMatrix(0, .ColIndex("Comment")) = "Remarks"
  End With
C1Tab1.Caption = "Group Data| Quality Control Items"
End Sub

Private Function GetNewGroupCode(LngParentGroupID As Long) As String
    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    Dim StrParentCode  As String
    Dim StrNewGroupCode As String
    Dim StrLastGroupCode As String
    Dim IntTemp As String

    On Error GoTo ErrTrap
    StrSQL = "Select GroupCode From Groups Where GroupID=" & LngParentGroupID & ""
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.EOF Or rs.BOF) Then
        StrParentCode = IIf(IsNull(rs("GroupCode").value), "", rs("GroupCode").value)
    End If

    rs.Close
    Set rs = New ADODB.Recordset
    StrSQL = "Select * From Groups Where ParentID=" & LngParentGroupID & " Order By GroupID"
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        StrNewGroupCode = StrParentCode & "1"
    Else
        rs.MoveLast
        StrLastGroupCode = IIf(IsNull(rs("GroupCode").value), "", rs("GroupCode").value)
        IntTemp = val(mId(StrLastGroupCode, Len(StrParentCode) + 1))
        StrNewGroupCode = StrParentCode & CStr(IntTemp + 1)
    End If

    rs.Close
    Set rs = Nothing
    GetNewGroupCode = StrNewGroupCode
    Exit Function
ErrTrap:
End Function

Private Sub GetGroupInfo(Optional Lngid As Long = 0)
    Dim rs As ADODB.Recordset
    Dim StrSQL As String

    If Lngid = 0 Then
        Exit Sub
    End If

    Set rs = New ADODB.Recordset
    StrSQL = "SELECT Count(Groups.GroupID) AS CountGroupID"
    StrSQL = StrSQL + " From Groups WHERE (Groups.ParentID=" & Lngid & ")"
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        Me.lbl(10).Caption = "0"
    Else
        Me.lbl(10).Caption = IIf(IsNull(rs("CountGroupID").value), 0, rs("CountGroupID").value)
    End If

    rs.Close
    Set rs = New ADODB.Recordset
    StrSQL = "SELECT Count(Groups.GroupID) AS CountGroupID "
    StrSQL = StrSQL + " FROM Groups INNER JOIN TblItems ON Groups.GroupID = TblItems.GroupID "
    StrSQL = StrSQL + " Where (((TblItems.GroupID) =" & Lngid & "))"
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        Me.lbl(11).Caption = "0"
    Else
        Me.lbl(11).Caption = IIf(IsNull(rs("CountGroupID").value), 0, rs("CountGroupID").value)
    End If

    rs.Close
    Set rs = New ADODB.Recordset
    StrSQL = "SELECT TblItems.ItemCode,TblItems.ItemName,TblItems.ItemName2 "
    StrSQL = StrSQL + " FROM Groups INNER JOIN TblItems ON Groups.GroupID = TblItems.GroupID"
    StrSQL = StrSQL + " WHERE (((TblItems.GroupID)=" & Lngid & ")) Order By TblItems.ItemID ASC;"
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.BOF Or rs.EOF) Then
        rs.MoveFirst
        Me.lbl(12).Caption = IIf(IsNull(rs("ItemCode").value), "", rs("ItemCode").value)
 If SystemOptions.UserInterface = ArabicInterface Then
        Me.lbl(14).Caption = IIf(IsNull(rs("ItemName").value), "", rs("ItemName").value)
  Else
  Me.lbl(14).Caption = IIf(IsNull(rs("ItemNamee").value), "", rs("ItemNamee").value)
  End If
        rs.MoveLast
        Me.lbl(15).Caption = IIf(IsNull(rs("ItemCode").value), "", rs("ItemCode").value)
If SystemOptions.UserInterface = ArabicInterface Then
        Me.lbl(13).Caption = IIf(IsNull(rs("ItemName").value), "", rs("ItemName").value)
  Else
  Me.lbl(13).Caption = IIf(IsNull(rs("ItemNamee").value), "", rs("ItemNamee").value)
  End If
    
    
    
    Else
        Me.lbl(12).Caption = ""
        Me.lbl(13).Caption = ""
        Me.lbl(14).Caption = ""
        Me.lbl(15).Caption = ""
    End If

End Sub

Private Sub XPTxtName_GotFocus()
     SwitchKeyboardLang LANG_ARABIC
End Sub

Private Sub XPTxtNameE_GotFocus()
SwitchKeyboardLang LANG_ENGLISH
End Sub


Private Sub AddNewRowProductLine()

    Dim Msg As String
    Dim LngRow As Long
    Dim LngFindRow As Long

    If cmbProductLine.text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "ÌÃ»    «”„ Œÿ «·«‰ «Ã ...!!!"
        Else
            Msg = "must Specify Product Line...!!!"
        End If

        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If
  
  
 
    'If FgPrices1.Rows = 1 Then
    '    LngRow = Val(Me.TxtRowNumber.text)
    'Else
    Me.grdProductLine.Rows = Me.grdProductLine.Rows + 1
    LngRow = Me.grdProductLine.Rows - 1
    'End If
  
    On Error Resume Next

    With Me.grdProductLine
    
        .TextMatrix(LngRow, .ColIndex("ProductLineName")) = Me.cmbProductLine.text
        .TextMatrix(LngRow, .ColIndex("ProductLineId")) = val(Me.cmbProductLine.BoundText)
        
       
        
   
        '    .AutoSize 0, .Cols - 1, False
    End With

    
 
    cmbProductLine.text = ""
    
    
 
End Sub

Private Sub RemoveProductLineRow()

    With Me.grdProductLine

        If .row <= 0 Then Exit Sub
        .RemoveItem .row
    End With

End Sub


