VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Begin VB.Form FrmTimeSetting 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "≈⁄œ«œ „Ê«⁄Ìœ «·Õ÷Ê— Ê«·«‰’—«› ..."
   ClientHeight    =   7395
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9240
   HelpContextID   =   520
   Icon            =   "FrmTimeSetting.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   7395
   ScaleWidth      =   9240
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
      Height          =   732
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   31
      Top             =   6600
      Width           =   9012
      Begin ImpulseButton.ISButton Cmd 
         Height          =   492
         Index           =   0
         Left            =   8220
         TabIndex        =   32
         Top             =   120
         Width           =   732
         _ExtentX        =   1296
         _ExtentY        =   873
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
         Height          =   492
         Index           =   1
         Left            =   7560
         TabIndex        =   33
         Top             =   120
         Width           =   852
         _ExtentX        =   1508
         _ExtentY        =   873
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
         Height          =   492
         Index           =   2
         Left            =   6960
         TabIndex        =   34
         Top             =   120
         Width           =   768
         _ExtentX        =   1349
         _ExtentY        =   873
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
         CausesValidation=   0   'False
         Height          =   492
         Index           =   3
         Left            =   6192
         TabIndex        =   35
         Top             =   120
         Width           =   768
         _ExtentX        =   1349
         _ExtentY        =   873
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
         Height          =   492
         Index           =   4
         Left            =   5400
         TabIndex        =   36
         Top             =   120
         Width           =   768
         _ExtentX        =   1349
         _ExtentY        =   873
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
         CausesValidation=   0   'False
         Height          =   492
         Index           =   6
         Left            =   3960
         TabIndex        =   37
         Top             =   120
         Width           =   768
         _ExtentX        =   1349
         _ExtentY        =   873
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
         Height          =   492
         Index           =   5
         Left            =   4716
         TabIndex        =   38
         Top             =   120
         Visible         =   0   'False
         Width           =   768
         _ExtentX        =   1349
         _ExtentY        =   873
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
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " ⁄œœ «·”Ã·« :"
         Height          =   312
         Index           =   4
         Left            =   720
         TabIndex        =   43
         Top             =   240
         Width           =   1092
      End
      Begin VB.Label lblCount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Height          =   252
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   42
         Top             =   240
         Width           =   492
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " «·”Ã· «·Õ«·Ì:"
         Height          =   312
         Index           =   3
         Left            =   2580
         TabIndex        =   41
         Top             =   240
         Width           =   1308
      End
      Begin VB.Label lblcurrent 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Height          =   252
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   240
         Width           =   612
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2E9E9&
      Height          =   1212
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   600
      Width           =   9012
      Begin VB.TextBox rec_id 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   288
         Left            =   5844
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   360
         Width           =   2412
      End
      Begin MSDataListLib.DataCombo dcBranch 
         Height          =   288
         Left            =   5844
         TabIndex        =   24
         Top             =   720
         Width           =   2412
         _ExtentX        =   4233
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo dcEmps 
         Height          =   288
         Left            =   2484
         TabIndex        =   28
         Top             =   360
         Width           =   2052
         _ExtentX        =   3598
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   348
         Index           =   10
         Left            =   120
         TabIndex        =   30
         Top             =   360
         Visible         =   0   'False
         Width           =   408
         _ExtentX        =   714
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
         ButtonImage     =   "FrmTimeSetting.frx":038A
         DrawFocusRectangle=   0   'False
      End
      Begin MSDataListLib.DataCombo DcboEmps 
         Height          =   288
         Left            =   2520
         TabIndex        =   39
         Top             =   720
         Width           =   2052
         _ExtentX        =   3625
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·ﬁ”„"
         Height          =   312
         Left            =   4464
         LinkTimeout     =   0
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   360
         Width           =   1008
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«”„ «·„ÊŸ›"
         Height          =   312
         Index           =   0
         Left            =   4500
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   720
         Width           =   1008
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·›—⁄"
         Height          =   312
         Left            =   7824
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   720
         Width           =   1008
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·—ﬁ„"
         Height          =   312
         Left            =   7860
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   360
         Width           =   1008
      End
   End
   Begin VB.TextBox TxtModFlg 
      Alignment       =   1  'Right Justify
      Height          =   288
      Left            =   -360
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   480
      Visible         =   0   'False
      Width           =   372
   End
   Begin ImpulseButton.ISButton CmdSave 
      Height          =   336
      Left            =   1056
      TabIndex        =   5
      Top             =   7800
      Visible         =   0   'False
      Width           =   768
      _ExtentX        =   1349
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
      ButtonImage     =   "FrmTimeSetting.frx":0724
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin ImpulseButton.ISButton CmdExit 
      Height          =   336
      Left            =   120
      TabIndex        =   6
      Top             =   7800
      Visible         =   0   'False
      Width           =   768
      _ExtentX        =   1349
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
      ButtonImage     =   "FrmTimeSetting.frx":0ABE
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
      DisabledImageStyle=   1
   End
   Begin C1SizerLibCtl.C1Tab CTab 
      Height          =   4500
      Left            =   156
      TabIndex        =   0
      Top             =   1836
      Width           =   8940
      _cx             =   15769
      _cy             =   7937
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
      FrontTabColor   =   14871017
      BackTabColor    =   12648447
      TabOutlineColor =   -2147483632
      FrontTabForeColor=   16711680
      Caption         =   "„Ê«⁄Ìœ «·Õ÷Ê—|‘—«∆Õ «·Œ’„"
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
      Picture(0)      =   "FrmTimeSetting.frx":0E58
      Picture(1)      =   "FrmTimeSetting.frx":11F2
      Begin C1SizerLibCtl.C1Elastic Elast1 
         Height          =   4095
         Left            =   15
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   15
         Width           =   8910
         _cx             =   15716
         _cy             =   7223
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
         AutoSizeChildren=   0
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
         Begin VSFlex8Ctl.VSFlexGrid Grid 
            Height          =   3390
            Left            =   75
            TabIndex        =   3
            Top             =   405
            Width           =   8910
            _cx             =   15716
            _cy             =   5980
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
            GridColor       =   14737632
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
            Cols            =   16
            FixedRows       =   2
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmTimeSetting.frx":158C
            ScrollTrack     =   0   'False
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   1
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
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "«·‘Ì›  «·«Ê·"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   372
            Left            =   5400
            RightToLeft     =   -1  'True
            TabIndex        =   11
            Top             =   0
            Width           =   4092
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "«·‘Ì›  «·À«‰Ì"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   480
            RightToLeft     =   -1  'True
            TabIndex        =   10
            Top             =   0
            Width           =   4095
         End
      End
      Begin C1SizerLibCtl.C1Elastic Elast2 
         Height          =   4095
         Left            =   9555
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   15
         Width           =   8910
         _cx             =   15716
         _cy             =   7223
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
         AutoSizeChildren=   0
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
         Begin ImpulseButton.ISButton CmdDelAll 
            Height          =   312
            Left            =   7380
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   2976
            Width           =   396
            _ExtentX        =   688
            _ExtentY        =   556
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
            ButtonImage     =   "FrmTimeSetting.frx":181E
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorShadow     =   -2147483631
            ColorOutline    =   -2147483631
            DrawFocusRectangle=   0   'False
            LowerToggledContent=   0   'False
         End
         Begin ImpulseButton.ISButton CmdAddRow 
            Height          =   312
            Left            =   8700
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   2976
            Width           =   396
            _ExtentX        =   688
            _ExtentY        =   556
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
            ButtonImage     =   "FrmTimeSetting.frx":1DB8
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorShadow     =   -2147483631
            ColorOutline    =   -2147483631
            DrawFocusRectangle=   0   'False
            LowerToggledContent=   0   'False
         End
         Begin ImpulseButton.ISButton CmdSubRow 
            Height          =   312
            Left            =   8220
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   2976
            Width           =   396
            _ExtentX        =   688
            _ExtentY        =   556
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
            ButtonImage     =   "FrmTimeSetting.frx":2152
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorShadow     =   -2147483631
            ColorOutline    =   -2147483631
            DrawFocusRectangle=   0   'False
            LowerToggledContent=   0   'False
         End
         Begin VSFlex8Ctl.VSFlexGrid Flex 
            Height          =   2820
            Left            =   72
            TabIndex        =   4
            Top             =   72
            Width           =   9012
            _cx             =   15896
            _cy             =   4974
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
            HighLight       =   2
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   50
            Cols            =   6
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmTimeSetting.frx":24EC
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
   Begin ImpulseButton.ISButton btnModify 
      Height          =   336
      Left            =   1920
      TabIndex        =   12
      Top             =   7800
      Visible         =   0   'False
      Width           =   756
      _ExtentX        =   1323
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
      ButtonImage     =   "FrmTimeSetting.frx":25E8
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin ImpulseButton.ISButton btnNew 
      Height          =   336
      Left            =   2760
      TabIndex        =   13
      Top             =   7800
      Visible         =   0   'False
      Width           =   756
      _ExtentX        =   1323
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
      ButtonImage     =   "FrmTimeSetting.frx":2982
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin C1SizerLibCtl.C1Elastic Ele 
      Height          =   528
      Index           =   5
      Left            =   0
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   0
      Width           =   9312
      _cx             =   16431
      _cy             =   926
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
      Picture         =   "FrmTimeSetting.frx":2D1C
      Caption         =   "     «⁄œ«œ „Ê«⁄Ìœ «·Õ÷Ê— Ê«·«‰’—«›     "
      Align           =   0
      AutoSizeChildren=   0
      BorderWidth     =   0
      ChildSpacing    =   0
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
      Begin ImpulseButton.ISButton XPBtnMove 
         Height          =   372
         Index           =   0
         Left            =   1572
         TabIndex        =   18
         Top             =   96
         Width           =   492
         _ExtentX        =   873
         _ExtentY        =   661
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
         ButtonImage     =   "FrmTimeSetting.frx":39F6
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
         Height          =   372
         Index           =   2
         Left            =   516
         TabIndex        =   19
         Top             =   96
         Width           =   492
         _ExtentX        =   873
         _ExtentY        =   661
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
         ButtonImage     =   "FrmTimeSetting.frx":3D90
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
         Height          =   372
         Index           =   1
         Left            =   2100
         TabIndex        =   20
         Top             =   96
         Width           =   492
         _ExtentX        =   873
         _ExtentY        =   661
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
         ButtonImage     =   "FrmTimeSetting.frx":412A
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
         Height          =   372
         Index           =   3
         Left            =   1032
         TabIndex        =   21
         Top             =   96
         Width           =   492
         _ExtentX        =   873
         _ExtentY        =   661
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
         ButtonImage     =   "FrmTimeSetting.frx":44C4
         ColorHighlight  =   4194304
         ColorHoverText  =   16777215
         ColorShadow     =   -2147483631
         ColorOutline    =   -2147483631
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
         ColorToggledHoverText=   16777215
         ColorTextShadow =   16777215
      End
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   312
      Left            =   120
      TabIndex        =   16
      Top             =   396
      Width           =   492
   End
   Begin VB.Label XPTxtCount 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   0
      TabIndex        =   15
      Top             =   30
      Width           =   495
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   -36
      X2              =   6489
      Y1              =   6456
      Y2              =   6471
   End
End
Attribute VB_Name = "FrmTimeSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim BKGrndPic As ClsBackGroundPic
Dim StrText As String
Dim m_WorkType As Integer
   Dim rs As ADODB.Recordset
  '   rs = New ADODB.Recordset
    
    
Private Sub btnModify_Click()
    Dim Msg As String

'    If DoPremis(Do_Edit, Me.name, True) = False Then
'        Exit Sub
'    End If

    On Error GoTo ErrTrap

 
    Exit Sub
ErrTrap:

    Select Case Err.Number

        Case -2147467259
            'Could not update; currently locked.
            Msg = "⁄›Ê«" & Chr(13)
            Msg = Msg & " ·«Ì„ﬂ‰  ⁄œÌ· Â–« «·”Ã· ›Ï «·Êﬁ  «·Õ«·Ï" & Chr(13)
            Msg = Msg & "ÕÌÀ «‰Â ﬁÌœ «· ⁄œÌ· „‰ ﬁ»· „” Œœ„ «Œ— ⁄·Ï «·‘»ﬂ…"
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
        
            If rs.EditMode <> adEditNone Then
                rs.CancelUpdate
                'RsSavRec.Requery
            End If

    End Select

End Sub

Public Sub AddNewRec()
    On Error GoTo ErrTrap
    Dim StrRecID As String
    StrRecID = new_id("tblTimeSetting", "rec_id", "")
    rs.AddNew
    rs.Fields("rec_id").value = IIf(StrRecID <> "", StrRecID, Null)
    SaveData
ErrTrap:
End Sub


Public Sub FiLLRec()
 

End Sub




Private Sub Cmd_Click(Index As Integer)
    
    ' On Error GoTo ErrTrap

    Select Case Index

        Case 0

            If DoPremis(Do_New, Me.name, True) = False Then
                Exit Sub
            End If
        
            TxtModFlg.text = "N"
            'clear_all Me
            Me.rec_id.text = CStr(new_id("tblTimeSettingH", "ID", "", True))
            'Grid.Clear flexClearScrollable, flexClearEverything
            'Grid.Rows = 1
            Grid.Enabled = True
            Dim i As Integer
          
          For i = 1 To 7
            Grid.TextMatrix(i + 1, 0) = GetWeekdayName(i)
          Next

dcEmps.BoundText = ""
Dcbranch.BoundText = ""
DcboEmps.BoundText = ""

           Clear_Grid

        Case 1

            If DoPremis(Do_Edit, Me.name, True) = False Then
                Exit Sub
            End If

           ' If ChKauto.value = vbChecked Then
           '     If SystemOptions.UserInterface = ArabicInterface Then
           '         MsgBox " ·« Ì„ﬂ‰  ⁄œÌ·  Œ’Ì’ «·Ì ", vbCritical
           '     Else
           '         MsgBox " Can't Delete Auto Employee Allocation ", vbCritical
           '     End If
'
'                Exit Sub
'            End If

            TxtModFlg.text = "E"
            'GRID2.Rows = GRID2.Rows + 1
            'GRID2.Enabled = True

        Case 2
    
            SaveDataT
           
        Case 3
            Undo

        Case 4

            If DoPremis(Do_Delete, Me.name, True) = False Then
                Exit Sub
            End If

             Del_Trans
        Case 5

            If DoPremis(Do_Search, Me.name, True) = False Then
                Exit Sub
            End If

General_Search.send_form = "visa"
            Load General_Search
'            FrmNotesSearch.SearchType = 3
          General_Search.show

        Case 6
            Unload Me

        Case 7
                 If DoPremis(Do_Print, Me.name, True) = False Then
                Exit Sub
            End If
            print_report
            
            Case 10
            
    FillGridWithData True
    
         End Select

    Exit Sub
ErrTrap:
    
    
    
End Sub


Private Sub Clear_Grid()
Dim i, j  As Integer
i = 2

For i = 2 To Grid.Rows - 1
 For j = 1 To Grid.Cols - 1
    Grid.TextMatrix(i, j) = ""
    Next
Next


End Sub


Private Sub Del_Trans()
    Dim Msg As String
    Dim StrSQL As String

    'On Error GoTo ErrTrap

    If Dcbranch.BoundText <> "" Then
        Msg = "”Ì „ Õ–› »Ì«‰«  «·⁄„·Ì… —ﬁ„ " & Chr(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            If Not rs.RecordCount < 1 Then
                rs.delete
                StrSQL = "Delete From tblTimeSettingH Where ID=" & val(rec_id.text)
                Cn.Execute StrSQL, , adExecuteNoRecords
                rs.MoveFirst
                Cn.Execute "  Delete from tblTimeSetting where Hid =  " & val(rec_id.text)
                FillGrid_2
                
                 Cn.Execute "delete tblTimeSettingEmp where rec_id=" & val(Me.Dcbranch.BoundText)

                If rs.RecordCount < 1 Then
              '  Grid.Clear flexClearScrollable, flexClearEverything
              Clear_Grid
                        ' Grid.Rows = 2
                    'GRID2.Clear flexClearScrollable, flexClearEverything
                    ' GRID2.Rows = 2
                    clear_all Me
                    TxtModFlg_Change
                   ' XPTxtCurrent.Caption = 0
                   ' XPTxtCount.Caption = 0
                Else
                    'FillGridWithData
                    XPBtnMove_Click (2)
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
    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ–› Â–Â «·»Ì«‰«  " & Chr(13)
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
    rs.CancelUpdate
End Sub

Private Sub Undo()
    On Error GoTo ErrTrap

    Select Case TxtModFlg.text

        Case "N"
            clear_all Me
            Me.TxtModFlg.text = "R"
            XPBtnMove_Click (1)

        Case "E"
 
            FillGridWithData
            Me.TxtModFlg.text = "R"
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub Dcbranch_Click(Area As Integer)
'    FillGridWithData
End Sub

Private Sub TxtModFlg_Change()

    If Me.TxtModFlg.text = "N" Then
       ' CmdRemove.Enabled = True
       ' Ele(1).Enabled = True
        Cmd(0).Enabled = False
        Cmd(1).Enabled = False
        Cmd(4).Enabled = False
        Cmd(5).Enabled = False

        Cmd(2).Enabled = True
        Cmd(3).Enabled = True

    ElseIf Me.TxtModFlg.text = "E" Then
       ' CmdRemove.Enabled = True
       ' Ele(1).Enabled = True
        Cmd(2).Enabled = True
        Cmd(3).Enabled = True

        Cmd(0).Enabled = False
        Cmd(1).Enabled = False
        Cmd(4).Enabled = False

        Cmd(5).Enabled = False

    Else
       ' Ele(1).Enabled = False

       ' CmdRemove.Enabled = False
        Cmd(2).Enabled = False
        Cmd(3).Enabled = False
        Cmd(0).Enabled = True
        Cmd(1).Enabled = True
        Cmd(4).Enabled = True

        Cmd(5).Enabled = True

    End If

End Sub



Private Sub CmdAddRow_Click()
    On Error GoTo ErrTrap

    Flex.Rows = Flex.Rows + 1
    Flex.TextMatrix(Flex.Rows - 1, Flex.ColIndex("Slice_Discount")) = Flex.Rows - 1
ErrTrap:
End Sub

Private Sub CmdDelAll_Click()
    Flex.Clear flexClearScrollable
    Flex.Rows = 1
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub SaveData()
    On Error GoTo ErrTrap
   
    Dim RsLat As ADODB.Recordset
    Dim II As Integer
    Dim My_SQL As String
    Dim Msg As String
    Dim BeginTrans As Boolean
    Cn.BeginTrans
    
   Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
    
    BeginTrans = True
  
    Set RsLat = New ADODB.Recordset

    If Me.WorkType = 0 Then
    
           With Grid
            For II = 2 To .Rows - 1
            
                           ' rs.find "Rec_ID=" & .TextMatrix(II, .ColIndex("Rec_ID")), 0, adSearchForward, 1

'                If rs.RecordCount > 0 Then
                rs.Fields("branchID").value = IIf(Dcbranch.BoundText = "", 0, val(Dcbranch.BoundText))
                    rs.Fields("Is_WorkDay").value = IIf(.TextMatrix(II, .ColIndex("Is_WorkDay")) = "", Null, .TextMatrix(II, .ColIndex("Is_WorkDay")))
                    rs.Fields("Bring_HourTime").value = IIf(.TextMatrix(II, .ColIndex("Bring_HourTime")) = "", Null, .TextMatrix(II, .ColIndex("Bring_HourTime")))
                    rs.Fields("Bring_MinuteTime").value = IIf(.TextMatrix(II, .ColIndex("Bring_MinuteTime")) = "", Null, .TextMatrix(II, .ColIndex("Bring_MinuteTime")))
                    rs.Fields("Go_HourTime").value = IIf(.TextMatrix(II, .ColIndex("Go_HourTime")) = "", Null, .TextMatrix(II, .ColIndex("Go_HourTime")))
                    rs.Fields("Go_MinuteTime").value = IIf(.TextMatrix(II, .ColIndex("Go_MinuteTime")) = "", Null, .TextMatrix(II, .ColIndex("Go_MinuteTime")))
                    rs.Fields("Bring_Time").value = IIf(.TextMatrix(II, .ColIndex("Bring_Time")) = "", Null, .TextMatrix(II, .ColIndex("Bring_Time")))
                    rs.Fields("Go_Time").value = IIf(.TextMatrix(II, .ColIndex("Go_Time")) = "", Null, .TextMatrix(II, .ColIndex("Go_Time")))
                
                    rs.Fields("Bring_HourTime1").value = IIf(.TextMatrix(II, .ColIndex("Bring_HourTime1")) = "", Null, .TextMatrix(II, .ColIndex("Bring_HourTime1")))
                    rs.Fields("Bring_MinuteTime1").value = IIf(.TextMatrix(II, .ColIndex("Bring_MinuteTime1")) = "", Null, .TextMatrix(II, .ColIndex("Bring_MinuteTime1")))
                    rs.Fields("Go_HourTime1").value = IIf(.TextMatrix(II, .ColIndex("Go_HourTime1")) = "", Null, .TextMatrix(II, .ColIndex("Go_HourTime1")))
                    rs.Fields("Go_MinuteTime1").value = IIf(.TextMatrix(II, .ColIndex("Go_MinuteTime1")) = "", Null, .TextMatrix(II, .ColIndex("Go_MinuteTime1")))
                    rs.Fields("Bring_Time1").value = IIf(.TextMatrix(II, .ColIndex("Bring_Time1")) = "", Null, .TextMatrix(II, .ColIndex("Bring_Time1")))
                    rs.Fields("Go_Time1").value = IIf(.TextMatrix(II, .ColIndex("Go_Time1")) = "", Null, .TextMatrix(II, .ColIndex("Go_Time1")))
                    rs.update
'                End If

            Next

        End With

    ElseIf Me.WorkType = 1 Then
        My_SQL = "Delete From tblTimeSettingEmp "
        My_SQL = My_SQL + " Where Emp_ID=" & Me.DcboEmps.BoundText
        Cn.Execute My_SQL, , adExecuteNoRecords
        rs.Open "tblTimeSettingEmp", Cn, adOpenKeyset, adLockOptimistic, adCmdTable

        With Grid

            For II = 2 To .Rows - 1
                rs.AddNew
                rs("Emp_ID").value = Me.DcboEmps.BoundText
                rs.Fields("DayNo").value = II - 1
                rs.Fields("Is_WorkDay").value = IIf(.TextMatrix(II, .ColIndex("Is_WorkDay")) = "", Null, .TextMatrix(II, .ColIndex("Is_WorkDay")))
                rs.Fields("Bring_HourTime").value = IIf(.TextMatrix(II, .ColIndex("Bring_HourTime")) = "", Null, .TextMatrix(II, .ColIndex("Bring_HourTime")))
                rs.Fields("Bring_MinuteTime").value = IIf(.TextMatrix(II, .ColIndex("Bring_MinuteTime")) = "", Null, .TextMatrix(II, .ColIndex("Bring_MinuteTime")))
                rs.Fields("Go_HourTime").value = IIf(.TextMatrix(II, .ColIndex("Go_HourTime")) = "", Null, .TextMatrix(II, .ColIndex("Go_HourTime")))
                rs.Fields("Go_MinuteTime").value = IIf(.TextMatrix(II, .ColIndex("Go_MinuteTime")) = "", Null, .TextMatrix(II, .ColIndex("Go_MinuteTime")))
                rs.Fields("Bring_Time").value = IIf(.TextMatrix(II, .ColIndex("Bring_Time")) = "", Null, .TextMatrix(II, .ColIndex("Bring_Time")))
                rs.Fields("Go_Time").value = IIf(.TextMatrix(II, .ColIndex("Go_Time")) = "", Null, .TextMatrix(II, .ColIndex("Go_Time")))
                rs.update
            Next II

        End With

    End If

    'Save Data to  tblSliceDiscount --------------------------------------------------------------------------------
    My_SQL = "delete From tblSliceDiscount"
    Cn.Execute My_SQL

    My_SQL = "select * From tblSliceDiscount order by Slice_ID"
    RsLat.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText

    With Flex

        For II = 1 To .Rows - 1
            RsLat.AddNew
            'Slice_ID
            RsLat("Slice_ID").value = new_id("tblSliceDiscount", "Slice_ID", "")
            RsLat.Fields("Late_Time").value = IIf(.TextMatrix(II, .ColIndex("Late_Time")) = "", Null, .TextMatrix(II, .ColIndex("Late_Time")))
            
            RsLat.Fields("Late_Type").value = IIf(.TextMatrix(II, .ColIndex("Late_Type")) = "", Null, .TextMatrix(II, .ColIndex("Late_Type")))
                                           
            RsLat.Fields("Discount").value = IIf(.TextMatrix(II, .ColIndex("Discount")) = "", Null, .TextMatrix(II, .ColIndex("Discount")))

            RsLat.Fields("Dis_Type").value = IIf(.TextMatrix(II, .ColIndex("Dis_Type")) = "", Null, .TextMatrix(II, .ColIndex("Dis_Type")))
        
            RsLat.update
        Next

    End With

    Cn.CommitTrans
    BeginTrans = False
    Msg = "·ﬁœ  „  ⁄„·Ì… «·Õ›Ÿ »‰Ã«Õ"
    MsgBox Msg, vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading

    Exit Sub
ErrTrap:

    If rs.EditMode <> adEditNone Then
        rs.CancelUpdate
    End If

    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If

    Msg = "·ﬁœ ›‘·  ⁄„·Ì… «·Õ›Ÿ "
    MsgBox Msg, vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading

End Sub

'/////////////////////////

Private Sub SaveDataT()
    Dim Msg As String
    Dim RsTemp As New ADODB.Recordset
    Dim StrSQL As String
    Dim BeginTrans As Boolean
    Dim RsDetails As ADODB.Recordset
    Dim i As Integer
    Dim LngDevID As Long
    Dim LngDevLineNo As Long
    Dim StrAccountCode As String


  Dim RsLat As ADODB.Recordset
    Dim II As Integer
    Dim My_SQL As String


Msg = "⁄›Ê« Ì—ÃÏ ≈” ﬂ„«· »«ﬁÏ «·»Ì«‰« "

    If ChkFlexData = False Then
        CTab.CurrTab = 0
        MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
        Exit Sub
    End If

 If ChkData = False Then
        CTab.CurrTab = 0
        MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
        Exit Sub
    End If


 If Me.dcEmps.BoundText = "" Then
  If SystemOptions.UserInterface = EnglishInterface Then
       Msg = "Select Department First ... !! "
        Else
            Msg = "ÌÃ» «ŒÌ«— «·ﬁ”„ «Ê·«..!! "
      End If
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            dcEmps.SetFocus
         '   SendKeys "{F4}"
            Exit Sub
        End If

 If Me.Dcbranch.BoundText = "" Then
           If SystemOptions.UserInterface = EnglishInterface Then
       Msg = "Select Branch First ... !! "
        Else
            Msg = "ÌÃ» «ŒÌ«— «·›—⁄ «Ê·«..!! "
            End If
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            dcEmps.SetFocus
         '   SendKeys "{F4}"
            Exit Sub
        End If


Set RsLat = New ADODB.Recordset
    If Me.TxtModFlg.text <> "R" Then
        Cn.BeginTrans
        BeginTrans = True
        
        If TxtModFlg.text = "N" Then
           rec_id.text = CStr(new_id("TblTimeSettingH", "ID", "", True))
           rs.AddNew
        ElseIf Me.TxtModFlg.text = "E" Then
            StrSQL = "Delete From tblTimeSetting  Where HID  = " & val(Me.rec_id.text)
            Cn.Execute StrSQL, , adExecuteNoRecords
        End If
        rs("ID").value = val(rec_id.text)
        rs("BranchID").value = IIf(Me.Dcbranch.BoundText = "", Null, Me.Dcbranch.BoundText)
        rs("DepID").value = IIf(Me.dcEmps.BoundText = "", Null, Me.dcEmps.BoundText)
        rs("EmpID").value = IIf(Me.DcboEmps.BoundText = "", Null, Me.DcboEmps.BoundText)
        rs.update
        
        Cn.CommitTrans
        BeginTrans = False
          Set RsDetails = New ADODB.Recordset
        RsDetails.Open "tblTimeSetting", Cn, adOpenKeyset, adLockOptimistic, adCmdTable
 
 
 
 
 
         '////////////////////
       
       
 If Me.WorkType = 0 Then
   With Grid
            For II = 2 To .Rows - 1
                        
              If .TextMatrix(II, .ColIndex("Day")) <> "" Then
                RsDetails.AddNew
                RsDetails.Fields("HID").value = IIf(rec_id.text = "", 0, val(rec_id.text))
                    RsDetails.Fields("rec_id").value = CStr(new_id("tblTimeSetting", "rec_ID", "", True))
                    RsDetails.Fields("DayNo").value = II - 1
                    RsDetails.Fields("branchID").value = IIf(Dcbranch.BoundText = "", 0, val(Dcbranch.BoundText))
                    RsDetails.Fields("Is_WorkDay").value = IIf(.TextMatrix(II, .ColIndex("Is_WorkDay")) = "", Null, .TextMatrix(II, .ColIndex("Is_WorkDay")))
                    RsDetails.Fields("Bring_HourTime").value = IIf(.TextMatrix(II, .ColIndex("Bring_HourTime")) = "", Null, .TextMatrix(II, .ColIndex("Bring_HourTime")))
                    RsDetails.Fields("Bring_MinuteTime").value = IIf(.TextMatrix(II, .ColIndex("Bring_MinuteTime")) = "", Null, .TextMatrix(II, .ColIndex("Bring_MinuteTime")))
                    RsDetails.Fields("Go_HourTime").value = IIf(.TextMatrix(II, .ColIndex("Go_HourTime")) = "", Null, .TextMatrix(II, .ColIndex("Go_HourTime")))
                    RsDetails.Fields("Go_MinuteTime").value = IIf(.TextMatrix(II, .ColIndex("Go_MinuteTime")) = "", Null, .TextMatrix(II, .ColIndex("Go_MinuteTime")))
                    RsDetails.Fields("Bring_Time").value = IIf(.TextMatrix(II, .ColIndex("Bring_Time")) = "", Null, .TextMatrix(II, .ColIndex("Bring_Time")))
                    RsDetails.Fields("Go_Time").value = IIf(.TextMatrix(II, .ColIndex("Go_Time")) = "", Null, .TextMatrix(II, .ColIndex("Go_Time")))
                    RsDetails.Fields("Bring_HourTime1").value = IIf(.TextMatrix(II, .ColIndex("Bring_HourTime1")) = "", Null, .TextMatrix(II, .ColIndex("Bring_HourTime1")))
                    RsDetails.Fields("Bring_MinuteTime1").value = IIf(.TextMatrix(II, .ColIndex("Bring_MinuteTime1")) = "", Null, .TextMatrix(II, .ColIndex("Bring_MinuteTime1")))
                    RsDetails.Fields("Go_HourTime1").value = IIf(.TextMatrix(II, .ColIndex("Go_HourTime1")) = "", Null, .TextMatrix(II, .ColIndex("Go_HourTime1")))
                    RsDetails.Fields("Go_MinuteTime1").value = IIf(.TextMatrix(II, .ColIndex("Go_MinuteTime1")) = "", Null, .TextMatrix(II, .ColIndex("Go_MinuteTime1")))
                    RsDetails.Fields("Bring_Time1").value = IIf(.TextMatrix(II, .ColIndex("Bring_Time1")) = "", Null, .TextMatrix(II, .ColIndex("Bring_Time1")))
                    RsDetails.Fields("Go_Time1").value = IIf(.TextMatrix(II, .ColIndex("Go_Time1")) = "", Null, .TextMatrix(II, .ColIndex("Go_Time1")))
                    RsDetails.update
                    RsDetails.update
                  'updatedata
                End If

            Next

        End With
  
    
    
    
    
    
    
    
           With Grid
            For II = 2 To .Rows - 1
            
                           ' rs.find "Rec_ID=" & .TextMatrix(II, .ColIndex("Rec_ID")), 0, adSearchForward, 1

'                If rs.RecordCount > 0 Then
                    RsDetails.Fields("HID").value = IIf(rec_id.text = "", 0, val(rec_id.text))
                RsDetails.Fields("branchID").value = IIf(Dcbranch.BoundText = "", 0, val(Dcbranch.BoundText))
                    RsDetails.Fields("Is_WorkDay").value = IIf(.TextMatrix(II, .ColIndex("Is_WorkDay")) = "", Null, .TextMatrix(II, .ColIndex("Is_WorkDay")))
                    RsDetails.Fields("Bring_HourTime").value = IIf(.TextMatrix(II, .ColIndex("Bring_HourTime")) = "", Null, .TextMatrix(II, .ColIndex("Bring_HourTime")))
                    RsDetails.Fields("Bring_MinuteTime").value = IIf(.TextMatrix(II, .ColIndex("Bring_MinuteTime")) = "", Null, .TextMatrix(II, .ColIndex("Bring_MinuteTime")))
                    RsDetails.Fields("Go_HourTime").value = IIf(.TextMatrix(II, .ColIndex("Go_HourTime")) = "", Null, .TextMatrix(II, .ColIndex("Go_HourTime")))
                    RsDetails.Fields("Go_MinuteTime").value = IIf(.TextMatrix(II, .ColIndex("Go_MinuteTime")) = "", Null, .TextMatrix(II, .ColIndex("Go_MinuteTime")))
                    RsDetails.Fields("Bring_Time").value = IIf(.TextMatrix(II, .ColIndex("Bring_Time")) = "", Null, .TextMatrix(II, .ColIndex("Bring_Time")))
                    RsDetails.Fields("Go_Time").value = IIf(.TextMatrix(II, .ColIndex("Go_Time")) = "", Null, .TextMatrix(II, .ColIndex("Go_Time")))
                
                    RsDetails.Fields("Bring_HourTime1").value = IIf(.TextMatrix(II, .ColIndex("Bring_HourTime1")) = "", Null, .TextMatrix(II, .ColIndex("Bring_HourTime1")))
                    RsDetails.Fields("Bring_MinuteTime1").value = IIf(.TextMatrix(II, .ColIndex("Bring_MinuteTime1")) = "", Null, .TextMatrix(II, .ColIndex("Bring_MinuteTime1")))
                    RsDetails.Fields("Go_HourTime1").value = IIf(.TextMatrix(II, .ColIndex("Go_HourTime1")) = "", Null, .TextMatrix(II, .ColIndex("Go_HourTime1")))
                    RsDetails.Fields("Go_MinuteTime1").value = IIf(.TextMatrix(II, .ColIndex("Go_MinuteTime1")) = "", Null, .TextMatrix(II, .ColIndex("Go_MinuteTime1")))
                    RsDetails.Fields("Bring_Time1").value = IIf(.TextMatrix(II, .ColIndex("Bring_Time1")) = "", Null, .TextMatrix(II, .ColIndex("Bring_Time1")))
                    RsDetails.Fields("Go_Time1").value = IIf(.TextMatrix(II, .ColIndex("Go_Time1")) = "", Null, .TextMatrix(II, .ColIndex("Go_Time1")))
                    RsDetails.update
'                End If

            Next

        End With

    ElseIf Me.WorkType = 1 Then
        My_SQL = "Delete From tblTimeSettingEmp "
        My_SQL = My_SQL + " Where Emp_ID=" & Me.DcboEmps.BoundText
        Cn.Execute My_SQL, , adExecuteNoRecords
        RsDetails.Open "tblTimeSettingEmp", Cn, adOpenKeyset, adLockOptimistic, adCmdTable

        With Grid
            For II = 2 To .Rows - 1
                RsDetails.AddNew
                RsDetails("Emp_ID").value = Me.DcboEmps.BoundText
                RsDetails.Fields("DayNo").value = II - 1
                RsDetails.Fields("Is_WorkDay").value = IIf(.TextMatrix(II, .ColIndex("Is_WorkDay")) = "", Null, .TextMatrix(II, .ColIndex("Is_WorkDay")))
                RsDetails.Fields("Bring_HourTime").value = IIf(.TextMatrix(II, .ColIndex("Bring_HourTime")) = "", Null, .TextMatrix(II, .ColIndex("Bring_HourTime")))
                RsDetails.Fields("Bring_MinuteTime").value = IIf(.TextMatrix(II, .ColIndex("Bring_MinuteTime")) = "", Null, .TextMatrix(II, .ColIndex("Bring_MinuteTime")))
                RsDetails.Fields("Go_HourTime").value = IIf(.TextMatrix(II, .ColIndex("Go_HourTime")) = "", Null, .TextMatrix(II, .ColIndex("Go_HourTime")))
                RsDetails.Fields("Go_MinuteTime").value = IIf(.TextMatrix(II, .ColIndex("Go_MinuteTime")) = "", Null, .TextMatrix(II, .ColIndex("Go_MinuteTime")))
                RsDetails.Fields("Bring_Time").value = IIf(.TextMatrix(II, .ColIndex("Bring_Time")) = "", Null, .TextMatrix(II, .ColIndex("Bring_Time")))
                RsDetails.Fields("Go_Time").value = IIf(.TextMatrix(II, .ColIndex("Go_Time")) = "", Null, .TextMatrix(II, .ColIndex("Go_Time")))
                RsDetails.update
            Next II
        End With
    End If

    'Save Data to  tblSliceDiscount --------------------------------------------------------------------------------
    My_SQL = "delete From tblSliceDiscount"
    Cn.Execute My_SQL

    My_SQL = "select * From tblSliceDiscount order by Slice_ID"
    RsLat.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText

    With Flex

        For II = 1 To .Rows - 1
            RsLat.AddNew
            'Slice_ID
            RsLat("Slice_ID").value = new_id("tblSliceDiscount", "Slice_ID", "")
            RsLat.Fields("Late_Time").value = IIf(.TextMatrix(II, .ColIndex("Late_Time")) = "", Null, .TextMatrix(II, .ColIndex("Late_Time")))
            
            RsLat.Fields("Late_Type").value = IIf(.TextMatrix(II, .ColIndex("Late_Type")) = "", Null, .TextMatrix(II, .ColIndex("Late_Type")))
                                           
            RsLat.Fields("Discount").value = IIf(.TextMatrix(II, .ColIndex("Discount")) = "", Null, .TextMatrix(II, .ColIndex("Discount")))

            RsLat.Fields("Dis_Type").value = IIf(.TextMatrix(II, .ColIndex("Dis_Type")) = "", Null, .TextMatrix(II, .ColIndex("Dis_Type")))
        
            RsLat.update
        Next

    End With
       
       
       
          '/////////////////
        
        
        
        
        
       'curr.Caption = rs.AbsolutePosition
       ' XPTxtCount.Caption = rs.RecordCount
    
        Select Case Me.TxtModFlg.text

            Case "N"
                Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & Chr(13)
                Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"

                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                    Cmd_Click (0)
                    Exit Sub
                End If

            Case "E"
                MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        End Select

        TxtModFlg.text = "R"
    End If

    Exit Sub
ErrTrap:

    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If

    If Err.Number = -2147217900 Then
        Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & Chr(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
End Sub



'////////////////////////////




Private Sub CmdSave_Click()
    On Error GoTo ErrTrap
    Dim Msg As String

    If DoPremis(Do_New, Me.name, True) = False Then
        Exit Sub
    End If

    Msg = "⁄›Ê« Ì—ÃÏ ≈” ﬂ„«· »«ﬁÏ «·»Ì«‰« "

    If ChkData = False Then
        CTab.CurrTab = 0
        MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
        Exit Sub
    End If

    If ChkFlexData = False Then
        CTab.CurrTab = 1
        MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
        Exit Sub
    End If

    Select Case Me.TxtModFlg.text

            '------------------------------ new record ----------------------------
        Case "N"
      
            '------------------------- save record -----------------------------
            AddNewRec
            'BtnLast_Click

        Case "E"

            '----------------------------- save edit -------------------------------
            FiLLRec
    End Select




    SaveData
ErrTrap:
End Sub


Private Sub SaveData1()
'On Error GoTo ErrTrap
   
'   On Error GoTo ErrTrap
   
    Dim RsLat As ADODB.Recordset
    Dim II As Integer
    Dim My_SQL As String
    Dim Msg As String
    Dim BeginTrans As Boolean
    Cn.BeginTrans
    BeginTrans = True
  
    Set RsLat = New ADODB.Recordset
 
 Dim Rs1 As ADODB.Recordset
   Set Rs1 = New ADODB.Recordset


    'Dim Msg As String
    Dim RsTemp As New ADODB.Recordset
    Dim StrSQL As String
     Dim RsDev As ADODB.Recordset
    Dim LngDevID As Long

'Set rs = New Recordset

   On Error GoTo ErrTrap
    If Me.TxtModFlg.text <> "R" Then
 
    End If

    '-------------------------------------------------------------------------------------------
   
  '  Cn.BeginTrans
   ' BeginTrans = True

Set Rs1 = New Recordset
    My_SQL = "select * From tblTimeSetting  where  branchID = " & val(Dcbranch.BoundText)
    Rs1.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
 Rs1.find " BranchID = " & val(Dcbranch.BoundText), 0, adSearchForward, 1
       If Rs1.RecordCount > 0 Then
       Rs1.MoveFirst
        Rs1.delete
                StrSQL = "Delete From tblTimeSetting Where branchID=" & val(Dcbranch.BoundText)
                Cn.Execute StrSQL, , adExecuteNoRecords
                Rs1.MoveFirst
       End If
       
    
    If TxtModFlg.text = "N" Then
        'rs.AddNew
    ElseIf Me.TxtModFlg.text = "E" Then
        Cn.Execute "delete tblTimeSettingEmp where rec_id=" & val(Me.rec_id.text)
   End If
    
    
    If Me.WorkType = 0 Then
    
           With Grid
            For II = 2 To .Rows - 1
                        
              If .TextMatrix(II, .ColIndex("Day")) <> "" Then
                Rs1.AddNew
                    Rs1.Fields("rec_id").value = CStr(new_id("tblTimeSetting", "rec_ID", "", True))
                    Rs1.Fields("DayNo").value = II - 1
                    Rs1.Fields("branchID").value = IIf(Dcbranch.BoundText = "", 0, val(Dcbranch.BoundText))
                    Rs1.Fields("Is_WorkDay").value = IIf(.TextMatrix(II, .ColIndex("Is_WorkDay")) = "", Null, .TextMatrix(II, .ColIndex("Is_WorkDay")))
                    Rs1.Fields("Bring_HourTime").value = IIf(.TextMatrix(II, .ColIndex("Bring_HourTime")) = "", Null, .TextMatrix(II, .ColIndex("Bring_HourTime")))
                    Rs1.Fields("Bring_MinuteTime").value = IIf(.TextMatrix(II, .ColIndex("Bring_MinuteTime")) = "", Null, .TextMatrix(II, .ColIndex("Bring_MinuteTime")))
                    Rs1.Fields("Go_HourTime").value = IIf(.TextMatrix(II, .ColIndex("Go_HourTime")) = "", Null, .TextMatrix(II, .ColIndex("Go_HourTime")))
                    Rs1.Fields("Go_MinuteTime").value = IIf(.TextMatrix(II, .ColIndex("Go_MinuteTime")) = "", Null, .TextMatrix(II, .ColIndex("Go_MinuteTime")))
                    Rs1.Fields("Bring_Time").value = IIf(.TextMatrix(II, .ColIndex("Bring_Time")) = "", Null, .TextMatrix(II, .ColIndex("Bring_Time")))
                    Rs1.Fields("Go_Time").value = IIf(.TextMatrix(II, .ColIndex("Go_Time")) = "", Null, .TextMatrix(II, .ColIndex("Go_Time")))
                    Rs1.Fields("Bring_HourTime1").value = IIf(.TextMatrix(II, .ColIndex("Bring_HourTime1")) = "", Null, .TextMatrix(II, .ColIndex("Bring_HourTime1")))
                    Rs1.Fields("Bring_MinuteTime1").value = IIf(.TextMatrix(II, .ColIndex("Bring_MinuteTime1")) = "", Null, .TextMatrix(II, .ColIndex("Bring_MinuteTime1")))
                    Rs1.Fields("Go_HourTime1").value = IIf(.TextMatrix(II, .ColIndex("Go_HourTime1")) = "", Null, .TextMatrix(II, .ColIndex("Go_HourTime1")))
                    Rs1.Fields("Go_MinuteTime1").value = IIf(.TextMatrix(II, .ColIndex("Go_MinuteTime1")) = "", Null, .TextMatrix(II, .ColIndex("Go_MinuteTime1")))
                    Rs1.Fields("Bring_Time1").value = IIf(.TextMatrix(II, .ColIndex("Bring_Time1")) = "", Null, .TextMatrix(II, .ColIndex("Bring_Time1")))
                    Rs1.Fields("Go_Time1").value = IIf(.TextMatrix(II, .ColIndex("Go_Time1")) = "", Null, .TextMatrix(II, .ColIndex("Go_Time1")))
                    Rs1.update
                    rs.update
                  'updatedata
                End If

            Next

        End With

    ElseIf Me.WorkType = 1 Then
        My_SQL = "Delete From tblTimeSettingEmp "
        My_SQL = My_SQL + " Where Emp_ID=" & Me.DcboEmps.BoundText
        Cn.Execute My_SQL, , adExecuteNoRecords
        Rs1.Open "tblTimeSettingEmp", Cn, adOpenKeyset, adLockOptimistic, adCmdTable

        With Grid

            For II = 2 To .Rows - 1
                Rs1.AddNew
                Rs1("Emp_ID").value = Me.DcboEmps.BoundText
                Rs1.Fields("DayNo").value = II - 1
                Rs1.Fields("Is_WorkDay").value = IIf(.TextMatrix(II, .ColIndex("Is_WorkDay")) = "", Null, .TextMatrix(II, .ColIndex("Is_WorkDay")))
                Rs1.Fields("Bring_HourTime").value = IIf(.TextMatrix(II, .ColIndex("Bring_HourTime")) = "", Null, .TextMatrix(II, .ColIndex("Bring_HourTime")))
                Rs1.Fields("Bring_MinuteTime").value = IIf(.TextMatrix(II, .ColIndex("Bring_MinuteTime")) = "", Null, .TextMatrix(II, .ColIndex("Bring_MinuteTime")))
                Rs1.Fields("Go_HourTime").value = IIf(.TextMatrix(II, .ColIndex("Go_HourTime")) = "", Null, .TextMatrix(II, .ColIndex("Go_HourTime")))
                Rs1.Fields("Go_MinuteTime").value = IIf(.TextMatrix(II, .ColIndex("Go_MinuteTime")) = "", Null, .TextMatrix(II, .ColIndex("Go_MinuteTime")))
                Rs1.Fields("Bring_Time").value = IIf(.TextMatrix(II, .ColIndex("Bring_Time")) = "", Null, .TextMatrix(II, .ColIndex("Bring_Time")))
                Rs1.Fields("Go_Time").value = IIf(.TextMatrix(II, .ColIndex("Go_Time")) = "", Null, .TextMatrix(II, .ColIndex("Go_Time")))
                Rs1.update
            Next II

        End With

    End If
    Rs1.update
    
  

    Cn.CommitTrans
    BeginTrans = False
 
    Select Case Me.TxtModFlg.text

        Case "N"
            Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·⁄„·Ì… " & Chr(13)
            Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"

            '    Fg_Journal.Enabled = False
            If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                Cmd_Click (0)
                Exit Sub
            End If

        Case "E"
            MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            '  Fg_Journal.Enabled = False
            XPBtnMove_Click (2)
    End Select

    TxtModFlg.text = "R"
    'End If

    Exit Sub
ErrTrap:

    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If

    If Err.Number = -2147217900 Then
        Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & Chr(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title


End Sub




Private Sub CmdSubRow_Click()
    On Error GoTo ErrTrap

    If Flex.Rows = 1 Then Exit Sub
    Flex.RemoveItem Flex.Row
ErrTrap:
End Sub

Private Sub Command1_Click()

End Sub

Private Sub DcboEmps_Change()
'    FillGridWithData
End Sub

Private Sub DcboEmps_Click(Area As Integer)
    'FillGridWithData
End Sub

Private Sub Flex_AfterEdit(ByVal Row As Long, _
                           ByVal Col As Long)

    With Flex

        Select Case .Col

            Case .ColIndex("Late_Time")
            
                If val(.TextMatrix(.Row, .Col)) < 1 Then
                    .TextMatrix(.Row, .Col) = val(.TextMatrix(.Row, .Col))
                End If
        
        End Select

    End With

End Sub

Private Sub Flex_EnterCell()
    StrText = ""
End Sub

Private Sub Flex_KeyPressEdit(ByVal Row As Long, _
                              ByVal Col As Long, _
                              KeyAscii As Integer)

    With Flex

        Select Case .Col

            Case .ColIndex("Late_Time")
                KeyAscii = DataFormat(CurOnly, KeyAscii)

                If KeyAscii = 46 Then
                    If InStr(1, StrText, ".", vbTextCompare) > 0 Then
                        KeyAscii = 0
                    End If
                End If

        End Select

    End With

    StrText = StrText & Chr(KeyAscii)
End Sub

Private Sub ChangeLang()
    
    
Label6.Caption = "Department"
lbl(0).Caption = "Employee"
lbl(3).Caption = "Current Record"
lbl(4).Caption = "Records Count"

    Me.Caption = "Time Setting"
    'lbl.Caption = "Employee"
    CMDSave.Caption = "Save"
    CmdExit.Caption = "Exit"
    Me.CTab.TabCaption(0) = "Attendance Time"
    Me.CTab.TabCaption(1) = "Discount Slice"

    Label3.Caption = "Shift 1"
    Label2.Caption = "Shift 2"

    Cmd(0).Caption = "New"
    Cmd(1).Caption = "Update"
    Cmd(2).Caption = "Save"
    Cmd(3).Caption = "Undo"
    Cmd(4).Caption = "Delete"
    Cmd(6).Caption = "Exit"
    Label1.Caption = "Branch"
    'lbl.Caption = "Employee"
    Label4.Caption = "No."
    Ele(5).Caption = "Attendance Settings"
    
    With Me.Flex
        .TextMatrix(0, .ColIndex("Slice_Discount")) = "Slice Discount"
        .TextMatrix(0, .ColIndex("Late_Time")) = "Late Time"
        .TextMatrix(0, .ColIndex("Late_Type")) = "Late Type"
        .TextMatrix(0, .ColIndex("Dis_Type")) = "Dis Type"
    End With

    With Grid
        .TextMatrix(0, .ColIndex("Day")) = "Day"
        .TextMatrix(0, .ColIndex("Is_WorkDay")) = "WorkDay"
        .TextMatrix(0, .ColIndex("Bring_HourTime")) = "Hour"
        .TextMatrix(0, .ColIndex("Bring_MinuteTime")) = "Minute "
        .TextMatrix(0, .ColIndex("Bring_Time")) = ""
        .TextMatrix(0, .ColIndex("Go_HourTime")) = "Hour"
        .TextMatrix(0, .ColIndex("Go_MinuteTime")) = "Minute "
        .TextMatrix(0, .ColIndex("Go_Time")) = ""

        .TextMatrix(0, .ColIndex("Bring_HourTime1")) = "Hour"
        .TextMatrix(0, .ColIndex("Bring_MinuteTime1")) = "Minute "
        .TextMatrix(0, .ColIndex("Bring_Time1")) = ""
        .TextMatrix(0, .ColIndex("Go_HourTime1")) = "Hour"
        .TextMatrix(0, .ColIndex("Go_MinuteTime1")) = "Minute "
        .TextMatrix(0, .ColIndex("Go_Time1")) = ""

    End With

End Sub

Private Sub Form_Load()
    Dim Dcombos As ClsDataCombos
    
        
        Set Dcombos = New ClsDataCombos
  Dcombos.GetBranches Dcbranch
    Dcombos.GetEmpDepartments Me.dcEmps
   
    
    Set BKGrndPic = New ClsBackGroundPic
    Resize_Form Me

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    Set Dcombos = New ClsDataCombos
    Dcombos.GetEmployees Me.DcboEmps, True

    With Grid
        .MergeCells = flexMergeFixedOnly
        .MergeCompare = flexMCExact
        .MergeCol(0) = True
        .MergeCol(1) = True
        .MergeRow(0) = True
        .Editable = flexEDKbdMouse
        .RowHeight(-1) = 350
        .ExtendLastCol = True
        .WallPaper = BKGrndPic.Picture
    End With

    With Flex
        .ExtendLastCol = True
        .WallPaper = BKGrndPic.Picture
    End With

    CTab.CurrTab = 0
    FillGridCombo
    FillFlexWithData
    FillGridWithData
    ShowTip
    
    Dim StrSQL As String
    Set rs = New Recordset
   ' StrSQL = "select * From tblTimeSetting  "
  ' StrSQL = " select branchID  from tblTimeSetting group by branchID "
   StrSQL = " select * from TblTimeSettingH "
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    
   XPBtnMove_Click 2
    Me.TxtModFlg.text = "R"
    
    
End Sub

Public Sub FillGridWithData(Optional ID As Integer = 0, Optional LoadDef As Boolean = False)
   On Error GoTo ErrTrap
    Dim My_SQL As String
    Dim II As Integer
    Dim rs As ADODB.Recordset

    Set rs = New ADODB.Recordset

   If Me.WorkType = 0 Or LoadDef = True Then
     'My_SQL = "select * From tblTimeSetting  where  HID = " & val(dcBranch.BoundText) & "    order by Rec_ID "
     My_SQL = "select * From tblTimeSetting  where  HID = " & ID & "    order by Rec_ID "
   ElseIf Me.WorkType = 1 Then

        If Me.DcboEmps.BoundText = "" Then Exit Sub
       My_SQL = "select * From tblTimeSettingEmp "
       My_SQL = My_SQL + " Where Emp_ID=" & Me.DcboEmps.BoundText & " Order by Rec_ID "
   End If
   
   rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText

    With Grid

        If rs.RecordCount > 0 Then
            .Rows = rs.RecordCount + 2
            rs.MoveFirst

 If Not (rs.EOF Or rs.BOF) Then
            For II = 2 To .Rows - 1
                 If Not IsNull(rs("DayNO").value) Then
                 .TextMatrix(II, .ColIndex("Day")) = GetWeekdayName(rs("DayNO").value) ' IIf(IsNull(Rs.Fields("DayNO").Value), "", Rs.Fields("DayNO").Value)
                 End If
                 
                .TextMatrix(II, .ColIndex("Rec_ID")) = IIf(IsNull(rs.Fields("Rec_ID").value), "", rs.Fields("Rec_ID").value)
                .TextMatrix(II, .ColIndex("Is_WorkDay")) = IIf(IsNull(rs.Fields("Is_WorkDay").value), "", rs.Fields("Is_WorkDay").value)
                .TextMatrix(II, .ColIndex("Bring_HourTime")) = IIf(IsNull(rs.Fields("Bring_HourTime").value), "", rs.Fields("Bring_HourTime").value)
                .TextMatrix(II, .ColIndex("Bring_MinuteTime")) = IIf(IsNull(rs.Fields("Bring_MinuteTime").value), "", rs.Fields("Bring_MinuteTime").value)
                .TextMatrix(II, .ColIndex("Bring_Time")) = IIf(IsNull(rs.Fields("Bring_Time").value), "", rs.Fields("Bring_Time").value)
                .TextMatrix(II, .ColIndex("Go_HourTime")) = IIf(IsNull(rs.Fields("Go_HourTime").value), "", rs.Fields("Go_HourTime").value)
                .TextMatrix(II, .ColIndex("Go_MinuteTime")) = IIf(IsNull(rs.Fields("Go_MinuteTime").value), "", rs.Fields("Go_MinuteTime").value)
                .TextMatrix(II, .ColIndex("Go_Time")) = IIf(IsNull(rs.Fields("Go_Time").value), "", rs.Fields("Go_Time").value)
            
                .TextMatrix(II, .ColIndex("Bring_HourTime1")) = IIf(IsNull(rs.Fields("Bring_HourTime1").value), "", rs.Fields("Bring_HourTime1").value)
                .TextMatrix(II, .ColIndex("Bring_MinuteTime1")) = IIf(IsNull(rs.Fields("Bring_MinuteTime1").value), "", rs.Fields("Bring_MinuteTime1").value)
                .TextMatrix(II, .ColIndex("Bring_Time1")) = IIf(IsNull(rs.Fields("Bring_Time1").value), "", rs.Fields("Bring_Time1").value)
                .TextMatrix(II, .ColIndex("Go_HourTime1")) = IIf(IsNull(rs.Fields("Go_HourTime1").value), "", rs.Fields("Go_HourTime1").value)
                .TextMatrix(II, .ColIndex("Go_MinuteTime1")) = IIf(IsNull(rs.Fields("Go_MinuteTime1").value), "", rs.Fields("Go_MinuteTime1").value)
                .TextMatrix(II, .ColIndex("Go_Time1")) = IIf(IsNull(rs.Fields("Go_Time1").value), "", rs.Fields("Go_Time1").value)
            
                rs.MoveNext
            Next II
End If

        Else
            .Clear flexClearScrollable, flexClearText
        End If

    End With

ErrTrap:
End Sub



'------------------------------------------------------------------------------------------------------------

Private Sub FillGrid_2()
 ' On Error GoTo ErrTrap
    Dim My_SQL As String
    Dim II As Integer
      Dim rs2 As ADODB.Recordset

    Set rs2 = New ADODB.Recordset

If rs.EOF Or rs.BOF Then
Exit Sub
End If


   If Me.WorkType = 0 And Not (IsNull(rs("ID").value)) Then
    ' My_SQL = "select * From tblTimeSetting  where  branchid = " & val(rs("branchID").value)
     My_SQL = "select * From tblTimeSetting  where  HID = " & val(rs("ID").value)
     
     rec_id.text = rs("ID").value
    Dcbranch.BoundText = IIf(IsNull(rs("BranchID").value), "", rs("BranchID").value)
        dcEmps.BoundText = IIf(IsNull(rs("DepID").value), "", rs("DepID").value)
           DcboEmps.BoundText = IIf(IsNull(rs("EmpID").value), "", rs("EmpID").value)
           
           
   ElseIf Me.WorkType = 1 Then

        If Me.DcboEmps.BoundText = "" Then Exit Sub
       My_SQL = "select * From tblTimeSettingEmp "
       My_SQL = My_SQL + " Where Emp_ID=" & Me.DcboEmps.BoundText & " Order by Rec_ID "
   End If
   
   rs2.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    
    
    
    
With Grid

        If rs2.RecordCount > 0 Then
          '  .Rows = rs.RecordCount + 2
          .Rows = 10
        '    rs.MoveFirst

Dcbranch.BoundText = IIf(IsNull(rs2("branchID").value), "", rs2("branchID").value)
 If Not (rs2.EOF Or rs2.BOF) Then
            For II = 2 To rs2.RecordCount + 1
                 If Not IsNull(rs2("DayNO").value) Then
                 .TextMatrix(II, .ColIndex("Day")) = GetWeekdayName(rs2("DayNO").value) ' IIf(IsNull(Rs2.Fields("DayNO").Value), "", Rs2.Fields("DayNO").Value)
                 End If
                 
                .TextMatrix(II, .ColIndex("Rec_ID")) = IIf(IsNull(rs2.Fields("Rec_ID").value), "", rs2.Fields("Rec_ID").value)
                .TextMatrix(II, .ColIndex("Is_WorkDay")) = IIf(IsNull(rs2.Fields("Is_WorkDay").value), "", rs2.Fields("Is_WorkDay").value)
                .TextMatrix(II, .ColIndex("Bring_HourTime")) = IIf(IsNull(rs2.Fields("Bring_HourTime").value), "", rs2.Fields("Bring_HourTime").value)
                .TextMatrix(II, .ColIndex("Bring_MinuteTime")) = IIf(IsNull(rs2.Fields("Bring_MinuteTime").value), "", rs2.Fields("Bring_MinuteTime").value)
                .TextMatrix(II, .ColIndex("Bring_Time")) = IIf(IsNull(rs2.Fields("Bring_Time").value), "", rs2.Fields("Bring_Time").value)
                .TextMatrix(II, .ColIndex("Go_HourTime")) = IIf(IsNull(rs2.Fields("Go_HourTime").value), "", rs2.Fields("Go_HourTime").value)
                .TextMatrix(II, .ColIndex("Go_MinuteTime")) = IIf(IsNull(rs2.Fields("Go_MinuteTime").value), "", rs2.Fields("Go_MinuteTime").value)
                .TextMatrix(II, .ColIndex("Go_Time")) = IIf(IsNull(rs2.Fields("Go_Time").value), "", rs2.Fields("Go_Time").value)
            
                .TextMatrix(II, .ColIndex("Bring_HourTime1")) = IIf(IsNull(rs2.Fields("Bring_HourTime1").value), "", rs2.Fields("Bring_HourTime1").value)
                .TextMatrix(II, .ColIndex("Bring_MinuteTime1")) = IIf(IsNull(rs2.Fields("Bring_MinuteTime1").value), "", rs2.Fields("Bring_MinuteTime1").value)
                .TextMatrix(II, .ColIndex("Bring_Time1")) = IIf(IsNull(rs2.Fields("Bring_Time1").value), "", rs2.Fields("Bring_Time1").value)
                .TextMatrix(II, .ColIndex("Go_HourTime1")) = IIf(IsNull(rs2.Fields("Go_HourTime1").value), "", rs2.Fields("Go_HourTime1").value)
                .TextMatrix(II, .ColIndex("Go_MinuteTime1")) = IIf(IsNull(rs2.Fields("Go_MinuteTime1").value), "", rs2.Fields("Go_MinuteTime1").value)
                .TextMatrix(II, .ColIndex("Go_Time1")) = IIf(IsNull(rs2.Fields("Go_Time1").value), "", rs2.Fields("Go_Time1").value)
            
                rs2.MoveNext
            Next II
End If

        Else
            .Clear flexClearScrollable, flexClearText
        End If

    End With



End Sub



Public Sub FillGridCombo()
    On Error GoTo ErrTrap

    Dim II As Integer
    Dim StrHour As String
    Dim StrMinute As String

    For II = 1 To 24
        StrHour = StrHour & "#" & II & ";" & II & "|"
    Next II

    For II = 0 To 59

        StrMinute = StrMinute & "#" & II & ";" & II & "|"
    Next II

    With Grid

If SystemOptions.UserInterface = ArabicInterface Then
        .ColComboList(1) = "#0;" & "⁄„·" & "|#1;" & "⁄ÿ·…"
        Else
        .ColComboList(1) = "#0;" & "Work" & "|#1;" & "Holiday"
        End If
        .ColComboList(2) = StrMinute
        .ColComboList(3) = StrHour
        .ColComboList(9) = StrMinute
        .ColComboList(10) = StrHour
    
        .ColComboList(4) = "#0;" & "’" & "|#1;" & "„"
        .ColComboList(11) = "#0;" & "’" & "|#1;" & "„"
    
        .ColComboList(5) = StrMinute
        .ColComboList(6) = StrHour
        .ColComboList(12) = StrMinute
        .ColComboList(13) = StrHour
    
        .ColComboList(7) = "#0;" & "’" & "|#1;" & "„"
        .ColComboList(14) = "#0;" & "’" & "|#1;" & "„"
    End With

    With Flex
    If SystemOptions.UserInterface = ArabicInterface Then
        .ColComboList(.ColIndex("Late_Type")) = "#0;" & "œﬁÌﬁ…" & "|#.1;" & "”«⁄…" & "|#1;"
    Else
      .ColComboList(.ColIndex("Late_Type")) = "#0;" & "Munit" & "|#.1;" & "Hour" & "|#1;"
    End If
    If SystemOptions.UserInterface = ArabicInterface Then
        .ColComboList(.ColIndex("Dis_Type")) = "#.25;" & "—»⁄ ÌÊ„" & "|#.5;" & "‰’› ÌÊ„" & "|#1;" & "ÌÊ„"
        Else
         .ColComboList(.ColIndex("Dis_Type")) = "#.25;" & "Quarter Day" & "|#.5;" & "Half Day" & "|#1;" & "Day"
        End If
        
        .Editable = flexEDKbdMouse
    End With

ErrTrap:
End Sub

'--------------------------------------------------------------------------------------------------------------------

Private Sub Grid_AfterEdit(ByVal Row As Long, _
                           ByVal Col As Long)
    On Error GoTo ErrTrap

    With Grid

        If Col = .ColIndex("Is_WorkDay") Then
            If .TextMatrix(Row, .ColIndex("Is_WorkDay")) = CStr(1) Then
                .TextMatrix(Row, .ColIndex("Bring_HourTime")) = ""
                .TextMatrix(Row, .ColIndex("Bring_MinuteTime")) = ""
                .TextMatrix(Row, .ColIndex("Go_HourTime")) = ""
                .TextMatrix(Row, .ColIndex("Go_MinuteTime")) = ""
                .TextMatrix(Row, .ColIndex("Bring_Time")) = ""
                .TextMatrix(Row, .ColIndex("Go_Time")) = ""
        
            End If
        End If

    End With

ErrTrap:
End Sub

'------------------------------------------------------------------------------------------------------------------

Private Sub Grid_EnterCell()
    On Error GoTo ErrTrap

    With Grid
        .Editable = flexEDKbdMouse

        If .Col > 1 Then
            If .TextMatrix(.Row, .ColIndex("Is_WorkDay")) = CStr(1) Then
                .Editable = flexEDNone
            Else
                .Editable = flexEDKbdMouse
            End If
        End If

    End With

ErrTrap:
End Sub

Private Function ChkData() As Boolean
    On Error GoTo ErrTrap

    Dim II As Integer

    With Grid

        For II = 2 To 8

            If .TextMatrix(II, .ColIndex("Is_WorkDay")) = "" Then
                ChkData = False
                Exit Function
            ElseIf .TextMatrix(II, .ColIndex("Is_WorkDay")) = 0 Then

                If .TextMatrix(II, .ColIndex("Bring_HourTime")) = "" Or .TextMatrix(II, .ColIndex("Bring_MinuteTime")) = "" Or .TextMatrix(II, .ColIndex("Go_HourTime")) = "" Or .TextMatrix(II, .ColIndex("Go_MinuteTime")) = "" Then
                    '      .TextMatrix(II, .ColIndex("Bring_Time")) = "" Or  then
                    '   .TextMatrix(II, .ColIndex("Go_Time")) = "" Then
                    ChkData = False
                    Exit Function
                Else
                End If
            End If

        Next

    End With

    ChkData = True
ErrTrap:
End Function
'-----------------------------------------------------------------------------------------

Private Function ChkFlexData() As Boolean
    On Error GoTo ErrTrap
    '.TextMatrix(II, .ColIndex("Discount")) = "" Or_

    Dim II As Integer

    With Flex

        For II = 1 To .Rows - 1

            If .TextMatrix(II, .ColIndex("Slice_Discount")) <> "" Then
                If .TextMatrix(II, .ColIndex("Late_Time")) = "" Or .TextMatrix(II, .ColIndex("Late_Type")) = "" Or .TextMatrix(II, .ColIndex("Dis_Type")) = "" Then
                
                    ChkFlexData = False
                    Exit Function
                End If
            End If

        Next

    End With

    ChkFlexData = True
ErrTrap:
End Function

Private Sub FillFlexWithData()
    On Error GoTo ErrTrap
    Dim II As Integer
    Dim My_SQL As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset

    My_SQL = "select * From tblSliceDiscount order by Slice_ID"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText

    With Flex

        If rs.RecordCount > 0 Then
            .Rows = rs.RecordCount + 1
          '.Rows = 10
            rs.MoveFirst
    
            For II = 1 To .Rows - 1
                .TextMatrix(II, .ColIndex("Slice_Discount")) = II
        
                .TextMatrix(II, .ColIndex("Slice_ID")) = IIf(IsNull(rs.Fields("Slice_ID").value), "", rs.Fields("Slice_ID").value)
                                                
                .TextMatrix(II, .ColIndex("Late_Time")) = IIf(IsNull(rs.Fields("Late_Time").value), "", rs.Fields("Late_Time").value)
                                                
                .TextMatrix(II, .ColIndex("Late_Type")) = IIf(IsNull(rs.Fields("Late_Type").value), "", rs.Fields("Late_Type").value)
                                                
                .TextMatrix(II, .ColIndex("Discount")) = IIf(IsNull(rs.Fields("Discount").value), "", rs.Fields("Discount").value)
            
                .TextMatrix(II, .ColIndex("Dis_Type")) = IIf(IsNull(rs.Fields("Dis_Type").value), "", rs.Fields("Dis_Type").value)
     
                rs.MoveNext
            Next II

        Else
            .Rows = 1
        End If

    End With

ErrTrap:
End Sub

'-------------------------------------------------------------
Private Sub ShowTip()
    On Error GoTo ErrTrap
    Dim TTP As New clstooltip
    Dim Wrap As String
    Dim Msg As String
    Wrap = Chr(13) + Chr(10)

    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Õ›Ÿ" & Wrap & "· ”ÃÌ· «·»Ì«‰«  œ«Œ· ﬁ«⁄œ… " & Wrap & " «·»Ì«‰«  ≈÷€ÿ Â–« «·„› «Õ"
        .AddControl CMDSave, Msg, True
    End With

    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Œ—ÊÃ" & Wrap & "·≈€·«ﬁ Â–Â «·‰«›–…" & Wrap & "≈÷€ÿ Â–« «·„› «Õ"
        .AddControl CmdExit, Msg, True
    End With

    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "”ÿ— ÃœÌœ" & Wrap & "≈÷«›… ”ÿ— ÃœÌœ „‰ √Ã· «÷«›… ‘—ÌÕ… Œ’„" & Wrap & "≈÷€ÿ Â–« «·„› «Õ"
        .AddControl CmdAddRow, Msg, True
    End With

    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Õ–› ”ÿ— " & Wrap & "·Õ–› «·”ÿ— «·Õ«·Ï „‰ «·ÃœÊ·" & Wrap & "≈÷€ÿ Â–« «·„› «Õ"
        .AddControl CmdSubRow, Msg, True
    End With

    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Õ–› «·ﬂ· " & Wrap & "·Õ–› ﬂ· «·«”ÿ— „‰ «·ÃœÊ·" & Wrap & "≈÷€ÿ Â–« «·„› «Õ"
        .AddControl CmdDelAll, Msg, True
    End With

ErrTrap:
End Sub

Public Property Get WorkType() As Integer
    WorkType = m_WorkType
End Property

Public Property Let WorkType(ByVal vNewValue As Integer)
 

End Property

Private Sub ISButton1_Click()

End Sub

Private Sub ISButton2_Click()
End Sub

Sub updatedata()
Dim StrSQL As String
   'StrSQL = " select branchID  from tblTimeSetting group by branchID "
   'rs.Clone
   'rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

End Sub

Private Sub XPBtnMove_Click(Index As Integer)

    If Me.TxtModFlg.text = "N" Then
        clear_all Me
        Me.TxtModFlg.text = "R"
        XPBtnMove_Click (1)
    End If

   ' On Error GoTo ErrTrap

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

    FillGrid_2
    
    lblCurrent.Caption = rs.AbsolutePosition
    LBLCOUNT.Caption = rs.RecordCount
    
    Exit Sub

End Sub
