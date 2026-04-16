VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FRMGranteeData 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " ”ÃÌ· »Ì«‰«  «·÷„«‰"
   ClientHeight    =   9270
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9795
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   9270
   ScaleWidth      =   9795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.Frame Frame2 
      BackColor       =   &H00E2E9E9&
      Height          =   2655
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   -120
      Width           =   4575
      Begin VB.TextBox TxtNoOFVisits 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Text            =   "12"
         Top             =   720
         Width           =   855
      End
      Begin MSComCtl2.DTPicker DTRegMaintDate 
         Height          =   330
         Left            =   840
         TabIndex        =   25
         Top             =   240
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   582
         _Version        =   393216
         Format          =   119209985
         CurrentDate     =   38784
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   390
         Index           =   20
         Left            =   120
         TabIndex        =   26
         Top             =   2160
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
         ButtonImage     =   "FrmGaranteeData.frx":0000
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin C1SizerLibCtl.C1Elastic ELe 
         Height          =   975
         Index           =   3
         Left            =   1080
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   1140
         Width           =   3330
         _cx             =   5874
         _cy             =   1720
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
         Begin VB.OptionButton OptInt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÌÊ„"
            Height          =   210
            Index           =   0
            Left            =   2535
            RightToLeft     =   -1  'True
            TabIndex        =   31
            Top             =   345
            Width           =   630
         End
         Begin VB.OptionButton OptInt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‘Â—"
            Height          =   225
            Index           =   1
            Left            =   1650
            RightToLeft     =   -1  'True
            TabIndex        =   30
            Top             =   345
            Value           =   -1  'True
            Width           =   840
         End
         Begin VB.TextBox Txt 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Height          =   270
            Index           =   7
            Left            =   30
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   29
            Text            =   "1"
            Top             =   585
            Width           =   915
         End
         Begin VB.OptionButton OptInt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "”‰…"
            Height          =   225
            Index           =   2
            Left            =   870
            RightToLeft     =   -1  'True
            TabIndex        =   28
            Top             =   345
            Width           =   675
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„œ… «·› —…"
            Height          =   195
            Index           =   17
            Left            =   45
            RightToLeft     =   -1  'True
            TabIndex        =   33
            Top             =   345
            Width           =   825
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·› —… »Ì‰ «·“Ì«—« "
            ForeColor       =   &H00FF0000&
            Height          =   210
            Index           =   18
            Left            =   1050
            RightToLeft     =   -1  'True
            TabIndex        =   32
            Top             =   0
            Width           =   1980
         End
      End
      Begin MSDataListLib.DataCombo DCVisits 
         Height          =   315
         Left            =   960
         TabIndex        =   37
         Top             =   2160
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰Ê⁄ «·“Ì«—…"
         Height          =   255
         Index           =   15
         Left            =   3600
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   2160
         Width           =   795
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " «—ÌŒ «·’Ì«‰… Ì»œ√ „‰"
         Height          =   375
         Index           =   13
         Left            =   2280
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   360
         Width           =   2115
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "⁄œœ «·“Ì«—« "
         Height          =   375
         Index           =   14
         Left            =   2280
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   720
         Width           =   2115
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2E9E9&
      Height          =   2535
      Left            =   4680
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   0
      Width           =   5175
      Begin VB.OptionButton GranteeTypeopt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„⁄ «·ﬁÿ⁄"
         Height          =   195
         Index           =   1
         Left            =   960
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   1080
         Width           =   975
      End
      Begin VB.OptionButton GranteeTypeopt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "»œÊ‰ «·ﬁÿ⁄"
         Height          =   195
         Index           =   0
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   1080
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.TextBox txtvlaue 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Text            =   "12"
         Top             =   1800
         Width           =   855
      End
      Begin MSComCtl2.DTPicker GranteeStartDate 
         Height          =   330
         Left            =   1320
         TabIndex        =   9
         Top             =   1440
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   582
         _Version        =   393216
         Format          =   119209985
         CurrentDate     =   38784
      End
      Begin MSComCtl2.DTPicker GranteeEndDate 
         Height          =   330
         Left            =   1320
         TabIndex        =   10
         Top             =   2160
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   582
         _Version        =   393216
         Format          =   119209985
         CurrentDate     =   38784
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·”ÿ—: "
         Height          =   255
         Index           =   0
         Left            =   3690
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   180
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Height          =   255
         Index           =   3
         Left            =   2100
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   180
         Visible         =   0   'False
         Width           =   1545
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ﬂÊœ «·’‰›: "
         Height          =   255
         Index           =   1
         Left            =   3690
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   480
         Width           =   1155
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Height          =   255
         Index           =   4
         Left            =   2100
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   480
         Width           =   1545
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«”„ «·’‰›:"
         Height          =   255
         Index           =   2
         Left            =   3810
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   780
         Width           =   1035
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Height          =   255
         Index           =   5
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   780
         Width           =   3765
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " «—ÌŒ »œ«Ì… «·÷„«‰"
         Height          =   255
         Index           =   6
         Left            =   2880
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   1440
         Width           =   1995
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Height          =   255
         Index           =   7
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   120
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " «—ÌŒ ‰Â«Ì… «·÷„«‰"
         Height          =   255
         Index           =   9
         Left            =   2760
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   2280
         Width           =   2115
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰Ê⁄ «·÷„«‰"
         Height          =   255
         Index           =   10
         Left            =   3720
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   1080
         Width           =   1155
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "› —… «·÷„«‰"
         Height          =   255
         Index           =   11
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   1800
         Width           =   2235
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‘Â—"
         Height          =   255
         Index           =   12
         Left            =   1200
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   1800
         Width           =   435
      End
   End
   Begin ImpulseButton.ISButton CmdOk 
      Height          =   405
      Left            =   1020
      TabIndex        =   1
      Top             =   8850
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   714
      Caption         =   "Õ›Ÿ"
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
   Begin VB.TextBox TxtComment 
      Alignment       =   1  'Right Justify
      Height          =   975
      Left            =   30
      MaxLength       =   255
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   10110
      Width           =   4425
   End
   Begin ImpulseButton.ISButton CmdCancel 
      Height          =   405
      Left            =   60
      TabIndex        =   2
      Top             =   8850
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   714
      Caption         =   "«·€«¡"
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
      Index           =   21
      Left            =   8280
      TabIndex        =   4
      Top             =   8760
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   688
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   " Õ–› ”ÿ—"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonImage     =   "FrmGaranteeData.frx":039A
      DrawFocusRectangle=   0   'False
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   390
      Index           =   0
      Left            =   6840
      TabIndex        =   38
      Top             =   8760
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   688
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   " Õ–› «·ﬂ·"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonImage     =   "FrmGaranteeData.frx":0934
      DrawFocusRectangle=   0   'False
   End
   Begin C1SizerLibCtl.C1Tab TabMain 
      Height          =   6135
      Left            =   0
      TabIndex        =   39
      Top             =   2550
      Width           =   9750
      _cx             =   17198
      _cy             =   10821
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
      Caption         =   "»Ì«‰«  «·’Ì«‰…|»Ì«‰«  «·÷„«‰"
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
         Height          =   5760
         Index           =   1
         Left            =   -10305
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   45
         Width           =   9660
         _cx             =   17039
         _cy             =   10160
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
            Height          =   5685
            Index           =   0
            Left            =   13440
            TabIndex        =   41
            Top             =   495
            Width           =   9540
            _cx             =   16828
            _cy             =   10028
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
            FormatString    =   $"FrmGaranteeData.frx":0ECE
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
         Begin VSFlex8Ctl.VSFlexGrid Grid 
            Height          =   5955
            Left            =   0
            TabIndex        =   44
            Top             =   0
            Width           =   9750
            _cx             =   17198
            _cy             =   10504
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
            Cols            =   23
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmGaranteeData.frx":0F8E
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
      End
      Begin C1SizerLibCtl.C1Elastic ELe 
         Height          =   5760
         Index           =   0
         Left            =   45
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   45
         Width           =   9660
         _cx             =   17039
         _cy             =   10160
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
            Height          =   5685
            Index           =   1
            Left            =   13440
            TabIndex        =   43
            Top             =   495
            Width           =   9540
            _cx             =   16828
            _cy             =   10028
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
            FormatString    =   $"FrmGaranteeData.frx":1268
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
         Begin VSFlex8UCtl.VSFlexGrid GranteeTypeGrd 
            Height          =   5640
            Left            =   120
            TabIndex        =   45
            Top             =   60
            Width           =   9435
            _cx             =   16642
            _cy             =   9948
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
            Cols            =   6
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmGaranteeData.frx":1328
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
      End
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   255
      Index           =   8
      Left            =   5880
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   1800
      Visible         =   0   'False
      Width           =   1245
   End
End
Attribute VB_Name = "FRMGranteeData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public inde As Integer
Public FG As VSFlex8UCtl.vsFlexGrid

Public LngRow As Long

Public LngCol As Long

Public AllDate As String
Public AllIDS As String

Private Sub Cmd_Click(Index As Integer)

    Select Case Index
Case 0
     cleargrid

        Case 20
          '  addrow
calcrows
        Case 21
            RemoveGridRow
    End Select

End Sub
Function cleargrid()
    With Me.Grid
 '      .Clear flexClearScrollable, flexClearEverything
           .Clear flexClearScrollable

        .Rows = 1
     End With
End Function
Function calcrows()
  Dim intervalstr As String
Dim name As String
Dim NameE As String
Dim Remarks As String
 If Grid.Rows = 0 Then
Grid.Rows = 1
End If
 
If val(TxtNoOFVisits) = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
          MsgBox "Õœœ  ⁄œœ «·“Ì«—« ", vbCritical
        Else
        MsgBox "Enter No oF visits", vbCritical
        End If
        Exit Function
End If
   


If val(Txt(7)) = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
          MsgBox "Õœœ    «·› —… »Ì‰ «·“Ì«—« ", vbCritical
        Else
        MsgBox "Define Visits Period", vbCritical
        End If
        Exit Function
        
End If
   
   
  Dim NewDate As Date
  Dim PreDate As Date
    Dim DateInterval As String
    Dim DateNumber As Integer
Dim i As Integer
 If OptInt(0).value = True Then
        DateInterval = "d"
    ElseIf OptInt(1).value = True Then
        DateInterval = "M"
    ElseIf OptInt(2).value = True Then
        DateInterval = "yyyy"
    End If


    With Me.Grid
'       .Clear flexClearScrollable, flexClearEverything
          .Clear flexClearScrollable

        .Rows = 1 + val(TxtNoOFVisits.Text)

        For i = 1 To val(TxtNoOFVisits)

            DoEvents
    
            If i = 1 Then
                NewDate = DTRegMaintDate.value
     
            ElseIf i > 1 Then
                PreDate = CDate(Trim(.TextMatrix(i - 1, .ColIndex("MaDate"))))
                NewDate = DateAdd(DateInterval, Txt(7), PreDate)
            End If

             .TextMatrix(i, .ColIndex("MaDate")) = Format(NewDate, "dd/mm/yyyy")
            Due_Date = Format(NewDate, "yyyy/M/d")
           .TextMatrix(i, .ColIndex("Ser")) = i
        


If i = 3 Or i = 9 Then 'qurater
DCVisits.BoundText = 2
ElseIf i = 6 Then 'hA
DCVisits.BoundText = 3
ElseIf i = 12 Then 'Annual
DCVisits.BoundText = 4
Else
DCVisits.BoundText = 1
End If




         .TextMatrix(i, .ColIndex("MainID")) = val(DCVisits.BoundText)
     
          

 
  getMaintentTypesData val(.TextMatrix(i, .ColIndex("MainID"))), name, NameE, Remarks, intervalstr
 If SystemOptions.UserInterface = ArabicInterface Then
       .TextMatrix(i, .ColIndex("MainName")) = name
  Else
  .TextMatrix(i, .ColIndex("MainName")) = NameE
  End If
                    .TextMatrix(i, .ColIndex("REMARKS")) = Remarks
 
       
                    .TextMatrix(i, .ColIndex("Interval")) = intervalstr
                    
                    
          
        Next i

        .AutoSize 1, .Cols - 1, False
       
    End With

ReLineGrid
End Function
Public Sub FillGridWithData()

    On Error GoTo ErrTrap
    Dim strInputString As String
    Dim strFilterText As String
    Dim astrSplitItems() As String
    Dim astrFilteredItems() As String
    Dim strFilteredString As String
    Dim intX As Integer
    Dim intervalstr As String
Dim name As String
Dim NameE As String
Dim Remarks As String

    'strInputString = seriallist
    Dim Item_ID As Long
       Item_ID = val(frmsalebill.FG.TextMatrix(LngRow, frmsalebill.FG.ColIndex("Code")))
      Dim s As String
       s = " SELECT GranteeType.ID GranteeTypeID,GranteeType.name GranteeTypeName, ItemsGranteeType.Remarks,ItemsGranteeType.Period from ItemsGranteeType"
   s = s & "  LEFT OUTER JOIN GranteeType ON GranteeType.Id = ItemsGranteeType.GranteeTypeID"
   s = s & " WHERE ItemsGranteeType.ItemID =" & val(Item_ID)
   LoadGrid s, GranteeTypeGrd, True, False
     Dim astrSplitItems1() As String
     
     
    strFilterText = ","
 
    astrSplitItems = Split(Me.AllDate, strFilterText)
    If (UBound(astrSplitItems) + 1) > 0 Then
    Grid.Rows = UBound(astrSplitItems) + 1
End If
 astrSplitItems1 = Split(Me.AllIDS, strFilterText)
 If (UBound(astrSplitItems) + 1) > 0 Then
    Grid.Rows = UBound(astrSplitItems) + 1
End If


    For intX = 0 To UBound(astrSplitItems)
      
        Grid.TextMatrix(intX + 1, Grid.ColIndex("MaDate")) = Format$(astrSplitItems(intX), "dd/mm/yyyy")
    Grid.TextMatrix(intX + 1, Grid.ColIndex("MainID")) = astrSplitItems1(intX)
     
  getMaintentTypesData val(astrSplitItems1(intX)), name, NameE, Remarks, intervalstr
        Grid.TextMatrix(intX + 1, Grid.ColIndex("REMARKS")) = Remarks
 
       If SystemOptions.UserInterface = ArabicInterface Then
        Grid.TextMatrix(intX + 1, Grid.ColIndex("MainName")) = name
       Else
       Grid.TextMatrix(intX + 1, Grid.ColIndex("MainName")) = NameE
       End If
                   Grid.TextMatrix(intX + 1, Grid.ColIndex("Interval")) = intervalstr
                    
 
      Grid.TextMatrix(intX + 1, Grid.ColIndex("Ser")) = intX + 1
 
    Next

  
   
ErrTrap:
End Sub

Private Sub RemoveGridRow()

    With Me.Grid

        If .Row <= 0 Then Exit Sub
        .RemoveItem .Row
    End With

    ReLineGrid
End Sub

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdOk_Click()
    Dim Msg As String
    Dim ExpiryDate As Date
    Dim Askinterval As String
ReLineGrid
If inde = 1 Then
 If GranteeTypeopt(0).value = True Then
            FrmWarrantyOffer.GranteeTypeopt(0).value = True
 Else
            FrmWarrantyOffer.GranteeTypeopt(0).value = True
 End If
FrmWarrantyOffer.GranteeStartDate.value = GranteeStartDate.value
FrmWarrantyOffer.GranteeEndDate.value = GranteeEndDate.value
FrmWarrantyOffer.TxtAllDate.Text = AllDate
FrmWarrantyOffer.TxtAllIDS.Text = AllIDS
FrmWarrantyOffer.txtvlaue.Text = txtvlaue.Text
        Unload Me
  ElseIf inde = 2 Then
  '''//////////////
      If Not FrmWarrantyOffer.FG Is Nothing Then
 With FrmWarrantyOffer.FG
  .TextMatrix(LngRow, .ColIndex("GranteeType")) = IIf(GranteeTypeopt(0).value = True, 0, 1)
  .TextMatrix(LngRow, .ColIndex("GranteeStartDate")) = GranteeStartDate.value
  .TextMatrix(LngRow, .ColIndex("GranteeEndDate")) = GranteeEndDate.value
  .TextMatrix(LngRow, .ColIndex("RegularMaintenancedates")) = AllDate
  .TextMatrix(LngRow, .ColIndex("RegularMaintenanceIDS")) = AllIDS
  .TextMatrix(LngRow, .ColIndex("Period")) = val(txtvlaue.Text)
  
 End With
            
        Unload Me
    End If
  '''///////////
  Else
    If Not FG Is Nothing Then
 
        If Me.FG.ColIndex("GranteeType") <> -1 Then
 
            FG.TextMatrix(LngRow, FG.ColIndex("GranteeType")) = IIf(GranteeTypeopt(0).value = True, 0, 1)
        End If

        If Me.FG.ColIndex("GranteeStartDate") <> -1 Then
 
            FG.TextMatrix(LngRow, FG.ColIndex("GranteeStartDate")) = GranteeStartDate.value
        End If

        If Me.FG.ColIndex("GranteeEndDate") <> -1 Then
 
            FG.TextMatrix(LngRow, FG.ColIndex("GranteeEndDate")) = GranteeEndDate.value
        End If

        If Me.FG.ColIndex("RegularMaintenancedates") <> -1 Then
 
            FG.TextMatrix(LngRow, FG.ColIndex("RegularMaintenancedates")) = AllDate
        End If



        If Me.FG.ColIndex("RegularMaintenanceIDS") <> -1 Then
 
            FG.TextMatrix(LngRow, FG.ColIndex("RegularMaintenanceIDS")) = AllIDS
        End If
        
        
        If Me.FG.ColIndex("guaranteeTime") <> -1 Then
 
            FG.TextMatrix(LngRow, FG.ColIndex("guaranteeTime")) = val(txtvlaue.Text)
  
        End If

        Unload Me
    End If
  End If

End Sub

Sub addrow()
    Dim Msg As String
    Dim LngRow As Long
    Dim LngFindRow As Long
    Dim des As String
 
    Me.Grid.Rows = Me.Grid.Rows + 1
    LngRow = Me.Grid.Rows - 1

    With Me.Grid
 
        .TextMatrix(LngRow, .ColIndex("MaDate")) = (DTRegMaintDate.value)
        .AutoSize 0, .Cols - 1, False
    End With
  
    ReLineGrid
 
End Sub



Private Sub ReLineGrid()
    Dim IntCounter As Integer
    IntCounter = 0
    Dim i As Integer
    AllDate = ""
AllIDS = ""
    With Me.Grid

        For i = .FixedRows To .Rows - 1
    
            If .TextMatrix(i, .ColIndex("MaDate")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("Ser")) = IntCounter
                AllDate = AllDate & .TextMatrix(i, .ColIndex("MaDate")) & ","
         AllIDS = AllIDS & .TextMatrix(i, .ColIndex("MainID")) & ","
            End If

        Next i
   
    End With

End Sub

Private Sub Form_Activate()
If SystemOptions.UserInterface = ArabicInterface Then
If inde = 1 Then
lbl(0).Visible = False
lbl(1).Caption = "ﬂÊœ «·„‘—Ê⁄"
lbl(2).Caption = "«·„‘—Ê⁄"
Else
lbl(1).Caption = "ﬂÊœ «·’‰›"
lbl(2).Caption = "«”„ «·’‰›"
lbl(0).Visible = True
End If
Else
If inde = 1 Then
lbl(0).Visible = False
lbl(1).Caption = "Code"
lbl(2).Caption = "Project"
Else
lbl(0).Visible = True
lbl(1).Caption = "Code"
lbl(2).Caption = "Item Name"
End If
End If
End Sub

Private Sub Form_Load()
    CenterForm Me
 
    FormPostion Me, GetPostion

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        cahngelang
    End If

    Me.CmdOk.ButtonStyle = impActive
    Set CmdOk.ButtonImage = mdifrmmain.ImgLstMenuIcons.ListImages("Save").Picture
    CmdOk.ButtonPositionImage = impRightOfText

    Me.CmdCancel.ButtonStyle = impActive
    Set CmdCancel.ButtonImage = mdifrmmain.ImgLstMenuIcons.ListImages("Hide").Picture
    CmdCancel.ButtonPositionImage = impRightOfText
    'GranteeStartDate.value = Date
    GranteeEndDate.value = Date
    DTRegMaintDate.value = Date
 Dim Dcombos As New ClsDataCombos
 Dcombos.GetmaintennceType Me.DCVisits
End Sub

Function cahngelang()
    Me.Caption = "Guarantee Data"

    lbl(1).Caption = "ItemCode"
    lbl(2).Caption = "Item Name"
    lbl(10).Caption = "G. Type"
    GranteeTypeopt(0).Caption = "WithOut Part"
    GranteeTypeopt(1).Caption = "With Part"
    lbl(6).Caption = "Guarantee  Start Date"
    lbl(9).Caption = "Guarantee  Emd Date"
    lbl(11).Caption = "Guarantee Period"
    lbl(12).Caption = "Month"
    lbl(13).Caption = "preventive maintenance Dates"
    Cmd(20).Caption = "ADD"
    Cmd(21).Caption = "Delete Row"
    CmdOk.Caption = "Save"
    lbl(18).Caption = "Period"
    lbl(14).Caption = "Visit No."
    CmdCancel.Caption = "Cancel"
OptInt(0).RightToLeft = False
OptInt(1).RightToLeft = False
OptInt(2).RightToLeft = False
OptInt(0).Caption = "Day"
OptInt(1).Caption = "Month"
OptInt(2).Caption = "Year"
lbl(17).Caption = "Period"
lbl(15).Caption = "Type Visit"

    With Me.Grid
        .TextMatrix(0, .ColIndex("Ser")) = "Ser"
        .TextMatrix(0, .ColIndex("MaDate")) = "Preventive Maintenance Dates"
        .TextMatrix(0, .ColIndex("MainName")) = "Maintenance Name"
        .TextMatrix(0, .ColIndex("Interval")) = "Period"
        .TextMatrix(0, .ColIndex("Remarks")) = "Remarks"
    End With

End Function

Private Sub Form_Unload(Cancel As Integer)

    FormPostion Me, SavePostion
End Sub
 Public Function getMaintentTypesData(StrAccountCode As String, Optional ByRef name As String _
 , Optional ByRef NameE As String, Optional ByRef Remarks As String, Optional ByRef intervalstr As String)
  
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim intervaltype As Integer
  Dim interval As Double
               StrSQL = " select * from TblMaintenanceType   where id=" & StrAccountCode
                Set rs = Nothing
              rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
 

                If Not (rs.BOF Or rs.EOF) Then
                   name = IIf(IsNull(rs("name").value), "", rs("name").value)
                   NameE = IIf(IsNull(rs("namee").value), "", rs("namee").value)
                   
                    Remarks = IIf(IsNull(rs("REMARKS").value), "", rs("REMARKS").value)
                    intervaltype = IIf(IsNull(rs("intervaltype").value), 0, rs("intervaltype").value)
                    interval = IIf(IsNull(rs("interval").value), 0, rs("interval").value)
                    
                    intervalstr = ""
                      If SystemOptions.UserInterface = ArabicInterface Then
                            If intervaltype = 0 Then
                            intervalstr = "œﬁÌﬁ…"
                            ElseIf intervaltype = 1 Then
                            intervalstr = "”«⁄Â"
                             ElseIf intervaltype = 2 Then
                            intervalstr = "ÌÊ„"
                             ElseIf intervaltype = 3 Then
                            intervalstr = "«”»Ê⁄"
                             ElseIf intervaltype = 4 Then
                            intervalstr = "‘Â—"
                            ElseIf intervaltype = 5 Then
                            intervalstr = "”‰…"
                            Else
                            intervalstr = ""
                            End If
                    Else
                    
                            If intervaltype = 0 Then
                            intervalstr = "Minute"
                            ElseIf intervaltype = 1 Then
                            intervalstr = "hour"
                            ElseIf intervaltype = 2 Then
                            intervalstr = "day"
                            ElseIf intervaltype = 3 Then
                            intervalstr = "week"
                            ElseIf intervaltype = 4 Then
                            intervalstr = "Month"
                            ElseIf intervaltype = 5 Then
                            intervalstr = "Year"
                            Else
                            intervalstr = ""
                            End If
                            
                    End If



                  intervalstr = interval & "   " & intervalstr
                 End If

 End Function

Private Sub Grid_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    On Error GoTo ErrTrap
    Dim StrAccountCode As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String

    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long

    With Grid

        Select Case .ColKey(Col)
 
            Case "MainName"
                '  .TextMatrix(Row, .ColIndex("userid")) = user_id
                        
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("MainID"), False, True)
                .TextMatrix(Row, .ColIndex("MainID")) = StrAccountCode
             
    
 
Dim intervalstr As String
Dim name As String
Dim NameE As String
Dim Remarks As String
 
 
  getMaintentTypesData StrAccountCode, name, name, Remarks, intervalstr
                    .TextMatrix(Row, .ColIndex("REMARKS")) = Remarks
 
       
                    .TextMatrix(Row, .ColIndex("Interval")) = intervalstr
        End Select
    End With
 
ErrTrap:

End Sub

Private Sub Grid_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
                                  

    With Grid

        Select Case .ColKey(Col)

            Case "MainID"
                .ComboList = ""

            Case "interval"
                .ComboList = ""
        
            Case "REMARKS"
                .ComboList = ""
                  Case "MaDate"
                .ComboList = ""
                 Case "Ser"
                .ComboList = ""
                 Cancel = True
            
        End Select

    End With
End Sub

 
Private Sub Grid_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim rs As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim Msg As String

    With Grid

        Select Case .ColKey(Col)

            Case "MainName"
                'Full Path Display
                 
                StrSQL = " select * from TblMaintenanceType "
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = Grid.BuildComboList(rs, "*name", "id")
                Else
                    StrComboList = Grid.BuildComboList(rs, "*name", "id")
                End If
                
          
 
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
        End Select

    End With

End Sub

Public Sub txtvlaue_Change()
    Me.GranteeEndDate.value = DateAdd("M", val(Me.txtvlaue), Me.GranteeStartDate.value)
End Sub

Private Sub txtvlaue_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, txtvlaue.Text, 0)
End Sub
