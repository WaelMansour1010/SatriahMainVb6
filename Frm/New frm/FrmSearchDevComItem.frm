VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form FrmSearchDevComItem 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6330
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12705
   Icon            =   "FrmSearchDevComItem.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6330
   ScaleWidth      =   12705
   Begin C1SizerLibCtl.C1Elastic AdvElastic 
      Height          =   5895
      Left            =   120
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   -60
      Visible         =   0   'False
      Width           =   12735
      _cx             =   22463
      _cy             =   10398
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
      Begin VSFlex8UCtl.VSFlexGrid fg2 
         Height          =   4425
         Left            =   -1950
         TabIndex        =   37
         Top             =   780
         Width           =   12675
         _cx             =   22357
         _cy             =   7805
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
         Rows            =   50
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FrmSearchDevComItem.frx":000C
         ScrollTrack     =   -1  'True
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
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   660
         Left            =   -600
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   150
         Width           =   12705
         Begin MSComctlLib.ImageList ImageList1 
            Left            =   3120
            Top             =   0
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
                  Picture         =   "FrmSearchDevComItem.frx":00E6
                  Key             =   "CompanyName"
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmSearchDevComItem.frx":0480
                  Key             =   "Ser"
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmSearchDevComItem.frx":081A
                  Key             =   "Vac_Name"
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmSearchDevComItem.frx":0BB4
                  Key             =   "ShareCount"
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmSearchDevComItem.frx":0F4E
                  Key             =   "Dis_Count"
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmSearchDevComItem.frx":12E8
                  Key             =   "Bouns"
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmSearchDevComItem.frx":1682
                  Key             =   "SharesValue"
               EndProperty
               BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmSearchDevComItem.frx":1C1C
                  Key             =   "BuyValue"
               EndProperty
            EndProperty
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·œ›⁄«  «·„ﬁœ„… ··⁄„Ì·"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004000&
            Height          =   375
            Index           =   0
            Left            =   7095
            RightToLeft     =   -1  'True
            TabIndex        =   40
            Top             =   90
            Width           =   5280
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "0"
            Height          =   135
            Left            =   360
            RightToLeft     =   -1  'True
            TabIndex        =   39
            Top             =   240
            Visible         =   0   'False
            Width           =   975
         End
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   3
         Left            =   2040
         TabIndex        =   41
         Top             =   4800
         Width           =   945
         _ExtentX        =   1667
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
         BackStyle       =   0
         ColorButton     =   14871017
         ColorHighlight  =   16777215
         ColorHoverText  =   16711680
         ColorShadow     =   4210752
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         DisabledImageExtraction=   0
         ColorToggledHoverText=   16711680
         LowerToggledContent=   0   'False
         ColorTextShadow =   4210752
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   375
         Index           =   4
         Left            =   1020
         TabIndex        =   42
         Top             =   4800
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   "„”Õ"
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
      Begin ImpulseButton.ISButton Cmd 
         Cancel          =   -1  'True
         Height          =   375
         Index           =   5
         Left            =   120
         TabIndex        =   43
         Top             =   5400
         Width           =   855
         _ExtentX        =   1508
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
         BackStyle       =   0
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
   Begin VB.Frame Frame3 
      Caption         =   "«· «—ÌŒ"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3960
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   3480
      Width           =   5775
      Begin MSComCtl2.DTPicker dbFromDate 
         Height          =   270
         Left            =   3840
         TabIndex        =   17
         Top             =   240
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   476
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   107479041
         CurrentDate     =   38784
      End
      Begin MSComCtl2.DTPicker DBTo 
         Height          =   270
         Left            =   1800
         TabIndex        =   18
         Top             =   240
         Width           =   1440
         _ExtentX        =   2540
         _ExtentY        =   476
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   107479041
         CurrentDate     =   38784
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "„‰"
         Height          =   270
         Index           =   5
         Left            =   5190
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   240
         Width           =   465
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·Ï"
         Height          =   270
         Index           =   2
         Left            =   3120
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   240
         Width           =   600
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   " »‰«¡ ⁄·Ï"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1785
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   4080
      Width           =   12495
      Begin VB.TextBox TxtAttachedItemCode 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   7230
         TabIndex        =   55
         Top             =   1410
         Width           =   1500
      End
      Begin VB.TextBox TxtNoteSerial15 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   3660
         RightToLeft     =   -1  'True
         TabIndex        =   50
         Top             =   1020
         Width           =   1500
      End
      Begin VB.TextBox TxtNoteSerial13 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   7230
         RightToLeft     =   -1  'True
         TabIndex        =   48
         Top             =   990
         Width           =   1500
      End
      Begin VB.TextBox TxtNoteSerial12 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   9720
         RightToLeft     =   -1  'True
         TabIndex        =   45
         Top             =   1230
         Width           =   1500
      End
      Begin VB.TextBox TxtNoteSerial11 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   9720
         RightToLeft     =   -1  'True
         TabIndex        =   44
         Top             =   900
         Width           =   1500
      End
      Begin VB.CheckBox Selct 
         Alignment       =   1  'Right Justify
         Caption         =   "Ì „  Œ’Ì’ «·„ﬂÊ‰«  »‘ﬂ· ›⁄·Ì"
         Height          =   315
         Index           =   0
         Left            =   -480
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   240
         Width           =   3255
      End
      Begin VB.CheckBox Selct 
         Alignment       =   1  'Right Justify
         Caption         =   "Ì „ ⁄„· ’—› «·Ì"
         Height          =   315
         Index           =   1
         Left            =   1200
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   600
         Width           =   1575
      End
      Begin VB.CheckBox Selct 
         Alignment       =   1  'Right Justify
         Caption         =   "Ì „ ⁄„· «” ·«„ «·Ì"
         Height          =   195
         Index           =   2
         Left            =   1200
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox TxtSearchCode 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   5160
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   240
         Width           =   945
      End
      Begin VB.TextBox TxtMaxNo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   9720
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   600
         Width           =   2040
      End
      Begin VB.TextBox TxtMaxName 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   600
         Width           =   5760
      End
      Begin MSDataListLib.DataCombo Dcbranch 
         Height          =   315
         Left            =   9720
         TabIndex        =   21
         Top             =   240
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DCboStoreName 
         Height          =   315
         Left            =   6720
         TabIndex        =   22
         Top             =   240
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DBCboClientName 
         Height          =   315
         Left            =   3000
         TabIndex        =   28
         Top             =   240
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "6"
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcboItemID1 
         Height          =   315
         Left            =   4440
         TabIndex        =   53
         Top             =   1410
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " «·’‰›"
         Height          =   225
         Index           =   26
         Left            =   8820
         RightToLeft     =   -1  'True
         TabIndex        =   54
         Top             =   1455
         Width           =   765
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "—ﬁ„ ”‰œ ’—› «·„»Ì⁄« "
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   4
         Left            =   5250
         RightToLeft     =   -1  'True
         TabIndex        =   51
         Top             =   1050
         Width           =   1755
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "—ﬁ„ «·„»Ì⁄« "
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   3
         Left            =   8580
         RightToLeft     =   -1  'True
         TabIndex        =   49
         Top             =   1050
         Width           =   1065
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "—ﬁ„ ”‰œ «·«” ·«„"
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   1
         Left            =   11310
         RightToLeft     =   -1  'True
         TabIndex        =   47
         Top             =   1230
         Width           =   1155
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "—ﬁ„ ”‰œ «·’—›"
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   0
         Left            =   11370
         RightToLeft     =   -1  'True
         TabIndex        =   46
         Top             =   930
         Width           =   1065
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·⁄„Ì·"
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   42
         Left            =   5610
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   240
         Width           =   1050
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "—ﬁ„ «·„ﬂ”"
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   29
         Left            =   11760
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   600
         Width           =   705
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "«”„ «·„ﬂ”"
         Height          =   255
         Index           =   30
         Left            =   8640
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "«·›—⁄"
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   36
         Left            =   11760
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   240
         Width           =   585
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "«·„Œ“‰"
         Height          =   270
         Index           =   50
         Left            =   9015
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   240
         Width           =   570
      End
   End
   Begin ALLButtonS.ALLButton ALLButton1 
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   8400
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "»ÕÀ"
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
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmSearchDevComItem.frx":1FB6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtremark 
      Alignment       =   1  'Right Justify
      Height          =   1020
      Left            =   14160
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   4680
      Width           =   7830
   End
   Begin VB.TextBox txtorder_no 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   9840
      TabIndex        =   3
      Top             =   3600
      Width           =   1935
   End
   Begin VB.Frame FraHeader 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   660
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   12705
      Begin VB.PictureBox GrdImageList 
         BackColor       =   &H000000FF&
         Height          =   1000
         Left            =   0
         ScaleHeight     =   945
         ScaleWidth      =   945
         TabIndex        =   52
         Top             =   0
         Width           =   1000
      End
      Begin VB.Label lbltype 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   135
         Left            =   360
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   240
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "»ÕÀ ⁄‰ „ﬂÊ‰«  «·«’‰«›"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   375
         Index           =   2
         Left            =   7095
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   90
         Width           =   5280
      End
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Height          =   315
      Left            =   12840
      TabIndex        =   4
      Top             =   6120
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "6"
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VSFlex8UCtl.VSFlexGrid Fg 
      Height          =   2745
      Left            =   0
      TabIndex        =   8
      Top             =   720
      Width           =   12675
      _cx             =   22357
      _cy             =   4842
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
      Rows            =   50
      Cols            =   20
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmSearchDevComItem.frx":1FD2
      ScrollTrack     =   -1  'True
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
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   0
      Left            =   2040
      TabIndex        =   9
      Top             =   5940
      Width           =   945
      _ExtentX        =   1667
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
      BackStyle       =   0
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      DisabledImageExtraction=   0
      ColorToggledHoverText=   16711680
      LowerToggledContent=   0   'False
      ColorTextShadow =   4210752
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   1
      Left            =   1020
      TabIndex        =   10
      Top             =   5940
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "„”Õ"
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
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   11
      Top             =   5940
      Width           =   855
      _ExtentX        =   1508
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
      BackStyle       =   0
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
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   556
      _Version        =   393216
      Format          =   107479041
      CurrentDate     =   38784
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "«· «—ÌŒ"
      Height          =   375
      Left            =   1800
      TabIndex        =   14
      Top             =   0
      Width           =   495
   End
   Begin VB.Label lblitemid 
      Alignment       =   1  'Right Justify
      Height          =   495
      Left            =   960
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "„·ÕÊŸ…"
      Height          =   375
      Left            =   14160
      TabIndex        =   5
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "—ﬁ„ «·”‰œ"
      Height          =   375
      Left            =   11280
      TabIndex        =   2
      Top             =   3600
      Width           =   1215
   End
End
Attribute VB_Name = "FrmSearchDevComItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim rs As ADODB.Recordset
Dim cSearchDcbo As clsDCboSearch

'Private m_DcboItems As DataCombo

'Private m_RetrunType As Integer
'Public WithEvents FG1 As VSFlex8UCtl.vsFlexGrid

'Public WithEvents NewGrid As VSFlex8UCtl.vsFlexGrid
'Public NewGrid As New ClsGrid
 
Public LngRow As Long

Public LngCol As Long
Public CusID As Long




Public Sub Cmd_Click(Index As Integer)
    On Error GoTo ErrTrap

    Select Case Index

        Case 0

            If rs.State = adStateOpen Then
                rs.Close
            End If

            rs.Open Build_Sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
        
            If SystemOptions.UserInterface = ArabicInterface Then
                '   LblRes.Caption = "‰ ÌÃ… «·»ÕÀ = " & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
                '   LblRes.Caption = "Search Result=" & rs.RecordCount
            End If
    
            If rs.RecordCount < 1 Then
                Fg.Clear flexClearScrollable, flexClearEverything
                Fg.Rows = 2

                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Else
                    Msg = "NO Search Results Found...!!!"
                    MsgBox Msg, vbOKOnly + vbExclamation, App.title
                End If

                Exit Sub
            End If

            Retrive
            Fg.SetFocus

        Case 1
            clear_all Me
            Fg.Clear flexClearScrollable, flexClearEverything
            dbFromDate.value = ""
          
DBTo.value = ""
        Case 2
            Unload Me
    End Select

    Exit Sub
ErrTrap:

    If Err.Number = -2147217900 Then
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & CHR(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… „⁄«ÌÌ— «·»ÕÀ Ê√⁄œ «·„Õ«Ê·…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

End Sub







Private Sub DBCboClientName_Click(Area As Integer)
If val(DBCboClientName.BoundText) = 0 Then Exit Sub

    Dim EmpCode  As String
 
    GetTblCustemersCode , , DBCboClientName.BoundText, EmpCode
    Me.TxtSearchCode.Text = EmpCode
End Sub

Private Sub Fg_Click()
  '  On Error GoTo ErrTrap
       
    If Me.lbltype = 0 Then
        FrmDefinCompItem.Retrive val(Fg.TextMatrix(Fg.Row, Fg.ColIndex("id")))
     ElseIf Me.lbltype = 1 Then
     FrmPO9.txtMixID.Text = val(Fg.TextMatrix(Fg.Row, Fg.ColIndex("id")))
     FrmPO9.txtMIxCode.Text = (Fg.TextMatrix(Fg.Row, Fg.ColIndex("MaxNo")))
     
     ElseIf Me.lbltype = 2 Then
     FrmProductionOrder.txtMixID.Text = val(Fg.TextMatrix(Fg.Row, Fg.ColIndex("id")))
     FrmProductionOrder.txtMIxCode.Text = (Fg.TextMatrix(Fg.Row, Fg.ColIndex("MaxNo")))
     ElseIf Me.lbltype = 3 Then
     frmsalebill.Fg.TextMatrix(frmsalebill.Fg.Row, frmsalebill.Fg.ColIndex("MixNo")) = (Fg.TextMatrix(Fg.Row, Fg.ColIndex("MaxNo")))
     frmsalebill.Fg.TextMatrix(frmsalebill.Fg.Row, frmsalebill.Fg.ColIndex("StoreID2")) = (Fg.TextMatrix(Fg.Row, Fg.ColIndex("StoreID2")))
     'FrmProductionOrder
     End If
     
   
    
        
        
'    Exit Sub
'ErrTrap:
End Sub

Private Sub Retrive()
    Dim Num As Integer
    On Error GoTo ErrTrap
    Fg.Clear flexClearScrollable, flexClearEverything
 fg2.Clear flexClearScrollable, flexClearEverything
 
    If Not (rs.EOF Or rs.BOF) Then
        Fg.Rows = rs.RecordCount + 1
fg2.Rows = rs.RecordCount + 1

If Me.AdvElastic.Visible = False Then
        For Num = 1 To rs.RecordCount
            With Fg
                .TextMatrix(Num, .ColIndex("StoreID2")) = IIf(IsNull(rs("StoreID2").value), "", rs("StoreID2").value)
                .TextMatrix(Num, .ColIndex("id")) = IIf(IsNull(rs("IDMain").value), "", rs("IDMain").value)
                .TextMatrix(Num, .ColIndex("RecordDate")) = IIf(IsNull(rs("RecordDate").value), "", rs("RecordDate").value)
            If SystemOptions.UserInterface = EnglishInterface Then
                .TextMatrix(Num, .ColIndex("branch_name")) = IIf(IsNull(rs("branch_namee").value), "", Trim(rs("branch_namee").value))
                .TextMatrix(Num, .ColIndex("StoreNam2")) = IIf(IsNull(rs("StoreNamee3").value), "", Trim(rs("StoreNamee3").value))
                .TextMatrix(Num, .ColIndex("CusName")) = IIf(IsNull(rs("CusNamee").value), "", Trim(rs("CusNamee").value))
                                
               Else
               .TextMatrix(Num, .ColIndex("branch_name")) = IIf(IsNull(rs("branch_name").value), "", Trim(rs("branch_name").value))
               .TextMatrix(Num, .ColIndex("StoreNam2")) = IIf(IsNull(rs("StoreNam2").value), "", Trim(rs("StoreNam2").value))
               .TextMatrix(Num, .ColIndex("CusName")) = IIf(IsNull(rs("CusName").value), "", Trim(rs("CusName").value))

                End If
.TextMatrix(Num, .ColIndex("NoteSerial11")) = IIf(IsNull(rs("NoteSerial11").value), "", rs("NoteSerial11").value)
.TextMatrix(Num, .ColIndex("NoteSerial12")) = IIf(IsNull(rs("NoteSerial12").value), "", rs("NoteSerial12").value)
.TextMatrix(Num, .ColIndex("NoteSerial13")) = IIf(IsNull(rs("NoteSerial13").value), "", rs("NoteSerial13").value)
.TextMatrix(Num, .ColIndex("NoteSerial15")) = IIf(IsNull(rs("NoteSerial15").value), "", rs("NoteSerial15").value)

.TextMatrix(Num, .ColIndex("ItemName")) = IIf(IsNull(rs("ItemName").value), "", rs("ItemName").value)
          
          

                       .TextMatrix(Num, .ColIndex("MaxName")) = IIf(IsNull(rs("MaxName").value), "", Trim(rs("MaxName").value))
                .TextMatrix(Num, .ColIndex("MaxNo")) = IIf(IsNull(rs("MaxNo").value), "", Trim(rs("MaxNo").value))


               
               If (rs("Allocated").value) = True Then
.TextMatrix(Num, .ColIndex("a1")) = -1
Else
.TextMatrix(Num, .ColIndex("a1")) = 0
    
            End If
                          If rs("AlloPay").value = True Then
.TextMatrix(Num, .ColIndex("a2")) = -1
Else
.TextMatrix(Num, .ColIndex("a2")) = 0
    
            End If
                          If (rs("AlloRecep").value) = True Then
.TextMatrix(Num, .ColIndex("a3")) = -1
Else
.TextMatrix(Num, .ColIndex("a3")) = 0
    
            End If
    
            End With
 
            rs.MoveNext
        Next Num

        ' Fg.AutoSize 0, Fg.Cols - 1, False
   

Else 'adv

        For Num = 1 To rs.RecordCount


                   With fg2
        .TextMatrix(Num, .ColIndex("NoteID")) = IIf(IsNull(rs("NoteID").value), "", rs("NoteID").value)
             .TextMatrix(Num, .ColIndex("NoteSerial1")) = IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value)
              .TextMatrix(Num, .ColIndex("NoteDate")) = IIf(IsNull(rs("NoteDate").value), "", rs("NoteDate").value)
              .TextMatrix(Num, .ColIndex("Value")) = IIf(IsNull(rs("Value").value), "", rs("Value").value)
                   End With
 
            rs.MoveNext
        Next Num

        ' Fg.AutoSize 0, Fg.Cols - 1, False
    End If



 End If
    Exit Sub
ErrTrap:
End Sub

Private Sub Fg_DblClick()
    Fg_Click
    Unload Me
End Sub

Private Sub fg2_Click()
FrmPO3.txtAdvPay = val(fg2.TextMatrix(fg2.Row, fg2.ColIndex("Value")))
FrmPO3.tXTaDid = val(fg2.TextMatrix(fg2.Row, fg2.ColIndex("NoteID")))
End Sub

Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim StrSQL As String
    Dim BG As New ClsBackGroundPic
    Dim Dcombos As ClsDataCombos
       Set Dcombos = New ClsDataCombos
       If CusID <> 0 Then
       AdvElastic.Visible = True
       End If
       
       Dcombos.GetCustomersSuppliers 0, Me.DBCboClientName, True
    Dcombos.GetBranches Me.Dcbranch
Dcombos.GetStores Me.DCboStoreName
    'Dcombos.GetItemsNames Me.DcboItemID1, -1, -1, 1
        Dcombos.GetItemsNames Me.DcboItemID1
            dbFromDate.value = ""
         
DBTo.value = ""

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Clear").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
 
    Dim My_SQL As String
    Set Dcombos = New ClsDataCombos

 
    CenterForm Me

    FormPostion Me, GetPostion
    Fg.WallPaper = BG.SearchWallpaper
    Set rs = New ADODB.Recordset
  
    Exit Sub
ErrTrap:

End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap

    If rs.State = adStateOpen Then
        rs.Close
        Set rs = Nothing
    End If

    Set cSearchDcbo = Nothing

    FormPostion Me, SavePostion
    Set m_DcboItems = Nothing
    Exit Sub
ErrTrap:
End Sub

Private Function Build_Sql()
    Dim StrSQL As String
    Dim Begin As Boolean
    Dim StrWhere As String
    Dim BolHaveSerial As Boolean
    Dim IntHaveSerial As Integer
 
    On Error GoTo ErrTrap
    '
If AdvElastic.Visible = False Then
StrSQL = "SELECT DISTINCT    dbo.TblDefComItem.RecordDate, dbo.TblDefComItem.StoreID, dbo.TblDefComItem.StoreID2, TblStore_1.StoreName AS StoreNam2, "
StrSQL = StrSQL + " TblDefComItem.NoteSerial11,TblDefComItem.NoteSerial12,TblDefComItem.NoteSerial13,TblDefComItem.NoteSerial15,"
 StrSQL = StrSQL + "                     TblStore_1.StoreNamee AS StoreNamee3, dbo.TblDefComItem.StoreID3, dbo.TblDefComItem.CusID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee,"
 StrSQL = StrSQL + "                     dbo.TblDefComItem.MaxNo, dbo.TblDefComItem.MaxName, dbo.TblDefComItem.Allocated, dbo.TblDefComItem.AlloPay, dbo.TblDefComItem.AlloRecep,"
StrSQL = StrSQL + "                      dbo.TblDefComItem.ID AS IDMain, dbo.TblDefComItem.BranchID, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee,"
StrSQL = StrSQL + "                      dbo.TblDefComItem.ItemCode,TblItems.ItemName"
StrSQL = StrSQL + " FROM         dbo.TblBranchesData RIGHT OUTER JOIN"
StrSQL = StrSQL + "                      dbo.TblDefComItem ON dbo.TblBranchesData.branch_id = dbo.TblDefComItem.BranchID LEFT OUTER JOIN"
StrSQL = StrSQL + "                      dbo.TblCustemers ON dbo.TblDefComItem.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
StrSQL = StrSQL + "                      dbo.TblStore TblStore_1 ON dbo.TblDefComItem.StoreID2 = TblStore_1.StoreID"
StrSQL = StrSQL + "                Left Outer join      dbo.TblDefComItemData ON dbo.TblDefComItemData.IDDefCIT = TblDefComItem.ID"
StrSQL = StrSQL + "                Left Outer join      dbo.TblItems ON dbo.TblDefComItemData.ItemID= TblItems.ItemID"
    
    StrSQL = StrSQL + " WHERE     (1 = 1)"
 
    If Me.txtorder_no.Text <> "" Then
        '     FrmProductionOrder1.Retrive (Val(FG.TextMatrix(FG.Row, 3)))
     
        StrWhere = StrWhere + " and     (dbo.TblDefComItem.id = " & val(txtorder_no.Text) & ")"
    
    End If
    


    If Me.DcboItemID1.Text <> "" And val(Me.DcboItemID1.BoundText) <> 0 Then
 
        StrWhere = StrWhere + " and dbo.TblItems.ItemID =" & val(Me.DcboItemID1.BoundText) & ""
 
    End If
    If Me.Dcbranch.Text <> "" And val(Me.Dcbranch.BoundText) <> 0 Then
 
        StrWhere = StrWhere + " and dbo.TblDefComItem.BranchID =" & val(Me.Dcbranch.BoundText) & ""
 
    End If
      If Me.DCboStoreName.Text <> "" And val(Me.DCboStoreName.BoundText) <> 0 Then
 
        StrWhere = StrWhere + " and dbo.TblDefComItem.StoreID =" & val(Me.DCboStoreName.BoundText) & ""
 
    End If
    
      If Me.TxtNoteSerial11.Text <> "" Then
 
        StrWhere = StrWhere + " and dbo.TblDefComItem.NoteSerial11 ='" & Trim(Me.TxtNoteSerial11.Text) & "'"
 
    End If
    
    
      If Me.TxtNoteSerial15.Text <> "" Then
 
        StrWhere = StrWhere + " and dbo.TblDefComItem.NoteSerial15 ='" & Trim(Me.TxtNoteSerial15.Text) & "'"
 
    End If
    
      If Me.TxtNoteSerial12.Text <> "" Then
 
        StrWhere = StrWhere + " and dbo.TblDefComItem.NoteSerial12 ='" & Trim(Me.TxtNoteSerial12.Text) & "'"
 
    End If
    
      If Me.TxtNoteSerial13.Text <> "" Then
 
        StrWhere = StrWhere + " and dbo.TblDefComItem.NoteSerial13 ='" & Trim(Me.TxtNoteSerial13.Text) & "'"
 
    End If
    
    
    
        If Me.DBCboClientName.Text <> "" And val(Me.DBCboClientName.BoundText) <> 0 Then
 
        StrWhere = StrWhere + " and dbo.TblDefComItem.CusID =" & val(Me.DBCboClientName.BoundText) & ""
 
    End If
  If Selct(0).value = vbChecked Then
         StrWhere = StrWhere + " and dbo.TblDefComItem.Allocated = 1"
     End If
   If Selct(1).value = vbChecked Then
         StrWhere = StrWhere + " and dbo.TblDefComItem.AlloPay = 1"
     End If
       If Selct(2).value = vbChecked Then
         StrWhere = StrWhere + " and dbo.TblDefComItem.AlloRecep = 1"
     End If

      If Me.TxtMaxNo.Text <> "" Then
 
        StrWhere = StrWhere + " and dbo.TblDefComItem.MaxNo ='" & Me.TxtMaxNo.Text & "'"
 
    End If
    
     If Me.TxtMaxName.Text <> "" Then
 
        StrWhere = StrWhere + " and dbo.TblDefComItem.MaxName like '%" & Me.TxtMaxName.Text & "%'"
 
    End If
    
If Me.lbltype = 3 Then
StrWhere = StrWhere + " and dbo.TblDefComItem.ItemNameID =" & val(frmsalebill.Fg.TextMatrix(frmsalebill.Fg.Row, frmsalebill.Fg.ColIndex("Code"))) & ""
End If
     If Not IsNull(Me.dbFromDate.value) Then
        
            StrWhere = StrWhere & " AND dbo.TblDefComItem.RecordDate >=" & SQLDate(Me.dbFromDate.value, True) & ""
      End If
 
       If Not IsNull(Me.DBTo.value) Then
        
            StrWhere = StrWhere & " AND dbo.TblDefComItem.RecordDate <=" & SQLDate(Me.DBTo.value, True) & ""
      End If
  StrWhere = StrWhere + " order by dbo.TblDefComItem.MaxName"

    Build_Sql = StrSQL + StrWhere
    
    Else
    
    StrSQL = "SELECT  dbo.Notes. NoteID  ,dbo.Notes.NoteSerial1, dbo.Notes.NoteDate, dbo.DOUBLE_ENTREY_VOUCHERS.[Value], dbo.Notes.CusID"
StrSQL = StrSQL & " FROM         dbo.Notes INNER JOIN"
StrSQL = StrSQL & "                       dbo.DOUBLE_ENTREY_VOUCHERS ON dbo.Notes.NoteID = dbo.DOUBLE_ENTREY_VOUCHERS.Notes_ID"
StrSQL = StrSQL & "  WHERE     (dbo.Notes.NCashingType = 3) AND (dbo.DOUBLE_ENTREY_VOUCHERS.Credit_Or_Debit = 1) AND (dbo.Notes.CusID = " & CusID & ")"
Build_Sql = StrSQL
    End If
    Exit Function
ErrTrap:
End Function

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        If Me.ActiveControl Is Fg Then
            If Not Fg.TextMatrix(Fg.Row, 1) = "" Then
                Fg_Click
                Unload Me
            End If

        Else
            Cmd_Click (0)
        End If
    End If

    On Error GoTo ErrTrap

    If Shift = 2 Then
        If KeyCode = vbKeyX Then
            Cmd_Click (2)
        End If
    End If

    Exit Sub
ErrTrap:
End Sub





'Public Property Get RetrunType() As Integer
'    RetrunType = m_RetrunType
'End Property

'Public Property Let RetrunType(ByVal vNewValue As Integer)
'    m_RetrunType = vNewValue
    ' 0 = Retrun in the Items Screen
    ' 1 = Retrun in the Data Combo
'End Property

Private Sub ChangeLang()
    Me.Caption = "Search Definition of varieties / assemble components"
    Label1(2).Caption = Me.Caption
  Selct(0).Caption = "Customize components"
        Selct(1).Caption = "Exchange into action"
        Selct(2).Caption = "Work to receive"
        Label2.Caption = "No"
        Frame3.Caption = "Date"
         lbl(5).Caption = "From"
lbl(2).Caption = "To"
lbl(36).Caption = "Branch"
lbl(50).Caption = "Store"
lbl(42).Caption = "Customer"
lbl(29).Caption = "MixNo"
lbl(30).Caption = "MixName"
    Frame2.Caption = "By"




    Cmd(0).Caption = "Search"
    Cmd(1).Caption = "Clear"
    Cmd(2).Caption = "Exit"

    'OptType(0).Caption = "Start of the name"
    'OptType(1).Caption = "any part of the name"
    With Me.Fg
        .TextMatrix(0, .ColIndex("NumIndex")) = "Serial"
         .TextMatrix(0, .ColIndex("id")) = "No  "
        .TextMatrix(0, .ColIndex("RecordDate")) = "Date"
        .TextMatrix(0, .ColIndex("branch_name")) = " Branch"
             .TextMatrix(0, .ColIndex("StoreNam2")) = "Store "
  .TextMatrix(0, .ColIndex("MaxNo")) = "Max No"
  .TextMatrix(0, .ColIndex("MaxName")) = " Max Name"
  .TextMatrix(0, .ColIndex("CusName")) = " Customer"
  .TextMatrix(0, .ColIndex("a1")) = " Customize component"
.TextMatrix(0, .ColIndex("a2")) = " Exchange into action"
  .TextMatrix(0, .ColIndex("a3")) = " Work to receive"
  
  .TextMatrix(0, .ColIndex("NoteSerial11")) = "Out OrderNo"
  .TextMatrix(0, .ColIndex("NoteSerial12")) = " Receipt receipt"
  .TextMatrix(0, .ColIndex("NoteSerial13")) = "Inv No"
  .TextMatrix(0, .ColIndex("NoteSerial15")) = "Inv No out"
  
  
        '  .AutoSize 0, .Cols - 1, False
    End With

End Sub



Private Sub TxtAttachedItemCode_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
        If TxtAttachedItemCode.Text = "" Then
            Me.DcboItemID1.BoundText = ""
        Else
            Me.DcboItemID1.BoundText = GetItemID(Trim$(Me.TxtAttachedItemCode.Text))
        End If
    End If
End Sub

'Private Sub TxtItemCode_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'          If KeyCode = vbKeyReturn Then
'                If Trim(Me.TxtItemCode(1).text) = "" Then Exit Sub
'                StrSQL = "Select ItemID From TblItems Where ItemCode='" & Trim(Me.TxtItemCode(Index).text) & "'"
'                Set rs = New ADODB.Recordset
'                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
''
 '               If Not (rs.BOF Or rs.EOF) Then
 '                   DCboItem.BoundText = rs("ItemID").value
 '               Else
 '                   Msg = "·«ÌÊÃœ ’‰› „”Ã· »Â–« «·ﬂÊœ..!"
 '                   MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
 '               End If
 '           End If
 '
'End Sub


Private Sub TxtSearchCode_KeyPress(KeyAscii As Integer)
 Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetTblCustemersCode TxtSearchCode.Text, EmpID
        DBCboClientName.BoundText = EmpID
    End If
End Sub
