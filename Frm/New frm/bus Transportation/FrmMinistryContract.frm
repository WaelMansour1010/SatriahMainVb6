VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmMinistryContract 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«· ⁄«ﬁœ „⁄ «·Ê“«—…"
   ClientHeight    =   9165
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13410
   Icon            =   "FrmMinistryContract.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   9165
   ScaleWidth      =   13410
   WindowState     =   2  'Maximized
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin C1SizerLibCtl.C1Elastic C1Elastic2 
      Height          =   9165
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   13410
      _cx             =   23654
      _cy             =   16166
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
      Begin C1SizerLibCtl.C1Elastic C1Elastic6 
         Height          =   495
         Left            =   120
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   7815
         Width           =   5610
         _cx             =   9895
         _cy             =   873
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
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " ⁄œœ «·”Ã·« :"
            Height          =   288
            Index           =   4
            Left            =   912
            RightToLeft     =   -1  'True
            TabIndex        =   57
            Top             =   120
            Width           =   1212
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «·”Ã· «·Õ«·Ì:"
            Height          =   288
            Index           =   2
            Left            =   3948
            RightToLeft     =   -1  'True
            TabIndex        =   56
            Top             =   120
            Width           =   1224
         End
         Begin VB.Label XPTxtCount 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   288
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   55
            Top             =   120
            Width           =   492
         End
         Begin VB.Label XPTxtCurrent 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   288
            Left            =   2496
            RightToLeft     =   -1  'True
            TabIndex        =   54
            Top             =   120
            Width           =   876
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic5 
         Height          =   732
         Left            =   120
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   2640
         Width           =   13140
         _cx             =   23178
         _cy             =   1296
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
         Begin VB.TextBox txtYearCustom 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   6360
            Locked          =   -1  'True
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   91
            Top             =   240
            Width           =   888
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   312
            Left            =   3744
            RightToLeft     =   -1  'True
            TabIndex        =   12
            Top             =   240
            Width           =   372
         End
         Begin VB.CommandButton cmdSub 
            Caption         =   "-"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   312
            Left            =   3180
            RightToLeft     =   -1  'True
            TabIndex        =   13
            Top             =   240
            Width           =   324
         End
         Begin VB.TextBox txtDiscount 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Left            =   2208
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   14
            Top             =   240
            Width           =   960
         End
         Begin VB.TextBox txtNet 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Left            =   120
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   15
            Top             =   240
            Width           =   1104
         End
         Begin VB.TextBox txtStudentCount 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   10920
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   9
            Top             =   240
            Width           =   960
         End
         Begin VB.TextBox txtStudentCustom 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   8700
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   10
            Top             =   240
            Width           =   912
         End
         Begin VB.TextBox txtTotal 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   4356
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   11
            Top             =   240
            Width           =   888
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„Œ’’ «·ÿ«·» «·”‰ÊÏ"
            Height          =   408
            Index           =   19
            Left            =   7260
            RightToLeft     =   -1  'True
            TabIndex        =   92
            Top             =   240
            Width           =   984
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "/"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   18
            Left            =   3492
            RightToLeft     =   -1  'True
            TabIndex        =   52
            Top             =   240
            Width           =   180
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·’«›Ï"
            Height          =   288
            Index           =   14
            Left            =   1356
            RightToLeft     =   -1  'True
            TabIndex        =   51
            Top             =   240
            Width           =   708
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„Œ’’ «·ÿ«·» «·ÌÊ„Ï"
            Height          =   408
            Index           =   11
            Left            =   9624
            RightToLeft     =   -1  'True
            TabIndex        =   50
            Top             =   240
            Width           =   984
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «·ÿ·«»"
            Height          =   288
            Index           =   12
            Left            =   11772
            RightToLeft     =   -1  'True
            TabIndex        =   49
            Top             =   240
            Width           =   1044
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·ﬁÌ„… «·«Ã„«·Ì…"
            Height          =   288
            Index           =   13
            Left            =   5256
            RightToLeft     =   -1  'True
            TabIndex        =   48
            Top             =   240
            Width           =   888
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic4 
         Height          =   4245
         Left            =   120
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   3480
         Width           =   13170
         _cx             =   23230
         _cy             =   7488
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
         Begin VB.CommandButton Command2 
            Appearance      =   0  'Flat
            BackColor       =   &H00E2E9E9&
            Caption         =   "Õ–› «·ﬂ·"
            Height          =   372
            Left            =   600
            RightToLeft     =   -1  'True
            TabIndex        =   94
            Top             =   120
            Width           =   1092
         End
         Begin VB.CommandButton Command1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E2E9E9&
            Caption         =   "Õ–› ”ÿ—"
            Height          =   372
            Left            =   1800
            RightToLeft     =   -1  'True
            TabIndex        =   93
            Top             =   120
            Width           =   1092
         End
         Begin VB.TextBox txtValue 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   5280
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   85
            Top             =   120
            Width           =   1836
         End
         Begin VB.ComboBox DcbPeriodsID 
            Height          =   315
            ItemData        =   "FrmMinistryContract.frx":038A
            Left            =   -324
            List            =   "FrmMinistryContract.frx":0397
            RightToLeft     =   -1  'True
            TabIndex        =   20
            Top             =   4935
            Visible         =   0   'False
            Width           =   1476
         End
         Begin VB.TextBox TxtPeriods 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   -60
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   19
            Top             =   4695
            Visible         =   0   'False
            Width           =   1164
         End
         Begin VB.TextBox TxtPaymentCount 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   48
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   16
            Top             =   4335
            Visible         =   0   'False
            Width           =   1548
         End
         Begin VSFlex8UCtl.VSFlexGrid FgInstallments 
            Height          =   3465
            Left            =   120
            TabIndex        =   22
            Top             =   630
            Width           =   12990
            _cx             =   22913
            _cy             =   6112
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
            Rows            =   1
            Cols            =   9
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmMinistryContract.frx":03AA
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
         Begin ImpulseButton.ISButton Cmd 
            Height          =   372
            Index           =   5
            Left            =   3000
            TabIndex        =   21
            Top             =   120
            Width           =   1092
            _ExtentX        =   1931
            _ExtentY        =   661
            ButtonPositionImage=   1
            Caption         =   "«÷«›…"
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
            ButtonImage     =   "FrmMinistryContract.frx":0510
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
         Begin Dynamic_Byte.NourHijriCal FirstPaymentDateH 
            Height          =   324
            Left            =   8592
            TabIndex        =   18
            Top             =   120
            Width           =   1584
            _ExtentX        =   2805
            _ExtentY        =   582
         End
         Begin MSComCtl2.DTPicker FirstPaymentDate 
            Height          =   312
            Left            =   10164
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   120
            Width           =   1752
            _ExtentX        =   3096
            _ExtentY        =   556
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CustomFormat    =   "yyyy/MM/dd"
            Format          =   98304003
            CurrentDate     =   41640
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬁÌ„… «·œ›⁄…"
            Height          =   288
            Index           =   6
            Left            =   7248
            RightToLeft     =   -1  'True
            TabIndex        =   86
            Top             =   120
            Width           =   876
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·› —Â »Ì‰ «·œ›⁄« "
            Height          =   300
            Index           =   11
            Left            =   -210
            RightToLeft     =   -1  'True
            TabIndex        =   46
            Top             =   4695
            Visible         =   0   'False
            Width           =   1350
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   " «—ÌŒ «·œ›⁄…"
            Height          =   300
            Index           =   9
            Left            =   11916
            RightToLeft     =   -1  'True
            TabIndex        =   45
            Top             =   120
            Width           =   1032
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "⁄œœ «·œ›⁄« "
            Height          =   300
            Index           =   8
            Left            =   285
            RightToLeft     =   -1  'True
            TabIndex        =   44
            Top             =   4695
            Visible         =   0   'False
            Width           =   1095
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic3 
         Height          =   1812
         Left            =   120
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   720
         Width           =   13140
         _cx             =   23178
         _cy             =   3201
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
         Begin VB.TextBox txtYearDays 
            Alignment       =   1  'Right Justify
            Height          =   288
            Left            =   7416
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   89
            Top             =   1440
            Width           =   1704
         End
         Begin VB.TextBox txtContractYears 
            Alignment       =   1  'Right Justify
            Height          =   288
            Left            =   10332
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   87
            Top             =   1440
            Width           =   1584
         End
         Begin VB.TextBox txtIDMC 
            Alignment       =   1  'Right Justify
            Height          =   288
            Left            =   10332
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   77
            Top             =   216
            Width           =   1584
         End
         Begin VB.TextBox txtProcessNo 
            Alignment       =   1  'Right Justify
            Height          =   252
            Left            =   240
            Locked          =   -1  'True
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   1
            Top             =   1680
            Visible         =   0   'False
            Width           =   372
         End
         Begin VB.TextBox XPTxtBoxName 
            Alignment       =   1  'Right Justify
            Height          =   288
            Left            =   7416
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   7
            Top             =   600
            Width           =   4500
         End
         Begin VB.TextBox txtMinistryContractNo 
            Alignment       =   1  'Right Justify
            Height          =   288
            Left            =   7416
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   2
            Top             =   216
            Width           =   1704
         End
         Begin MSComCtl2.DTPicker dtpToDate 
            Height          =   156
            Left            =   72
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   1056
            Visible         =   0   'False
            Width           =   372
            _ExtentX        =   661
            _ExtentY        =   265
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CustomFormat    =   "yyyy/M/d"
            Format          =   98304003
            CurrentDate     =   37140
         End
         Begin MSComCtl2.DTPicker dtpFromDate 
            Height          =   252
            Left            =   1560
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   240
            Width           =   1788
            _ExtentX        =   3149
            _ExtentY        =   450
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CustomFormat    =   "yyyy/M/d"
            Format          =   98304003
            CurrentDate     =   37140
         End
         Begin Dynamic_Byte.NourHijriCal dtpFromDateH 
            Height          =   252
            Left            =   1560
            TabIndex        =   4
            Top             =   600
            Width           =   1800
            _ExtentX        =   3175
            _ExtentY        =   450
         End
         Begin Dynamic_Byte.NourHijriCal dtpToDateH 
            Height          =   288
            Left            =   840
            TabIndex        =   6
            Top             =   1440
            Visible         =   0   'False
            Width           =   372
            _ExtentX        =   661
            _ExtentY        =   503
         End
         Begin MSDataListLib.DataCombo dcVendor 
            Height          =   288
            Left            =   7416
            TabIndex        =   30
            Top             =   1056
            Width           =   1704
            _ExtentX        =   3016
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo dcCity 
            Height          =   288
            Left            =   10332
            TabIndex        =   31
            Top             =   1056
            Width           =   1584
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo dcClient 
            Height          =   288
            Left            =   4680
            TabIndex        =   8
            Top             =   636
            Width           =   1764
            _ExtentX        =   3122
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin C1SizerLibCtl.C1Elastic C1Elastic8 
            Height          =   744
            Left            =   1548
            TabIndex        =   78
            TabStop         =   0   'False
            Top             =   960
            Width           =   4896
            _cx             =   8652
            _cy             =   1323
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
            Caption         =   "„œ… «·⁄ﬁœ"
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
            Begin Dynamic_Byte.NourHijriCal dtpSContractDateH 
               Height          =   264
               Left            =   120
               TabIndex        =   79
               Top             =   108
               Width           =   1248
               _ExtentX        =   2196
               _ExtentY        =   476
            End
            Begin MSComCtl2.DTPicker dtpSContractDate 
               Height          =   264
               Left            =   1416
               TabIndex        =   80
               TabStop         =   0   'False
               Top             =   108
               Width           =   1572
               _ExtentX        =   2778
               _ExtentY        =   476
               _Version        =   393216
               CalendarBackColor=   12648447
               CalendarTitleBackColor=   10383715
               CustomFormat    =   "yyyy/M/d"
               Format          =   98304003
               CurrentDate     =   37140
            End
            Begin MSComCtl2.DTPicker dtpEContractDate 
               Height          =   276
               Left            =   1416
               TabIndex        =   82
               TabStop         =   0   'False
               Top             =   408
               Width           =   1572
               _ExtentX        =   2778
               _ExtentY        =   476
               _Version        =   393216
               CalendarBackColor=   12648447
               CalendarTitleBackColor=   10383715
               CustomFormat    =   "yyyy/M/d"
               Format          =   98304003
               CurrentDate     =   37140
            End
            Begin Dynamic_Byte.NourHijriCal dtpEContractDateH 
               Height          =   276
               Left            =   120
               TabIndex        =   83
               Top             =   408
               Width           =   1248
               _ExtentX        =   2196
               _ExtentY        =   476
            End
            Begin VB.Label Lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "Ì‰ ÂÏ ›Ï "
               Height          =   276
               Index           =   8
               Left            =   3036
               RightToLeft     =   -1  'True
               TabIndex        =   84
               Top             =   408
               Width           =   828
            End
            Begin VB.Label Lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "Ì»œ√ „‰"
               Height          =   264
               Index           =   5
               Left            =   2856
               RightToLeft     =   -1  'True
               TabIndex        =   81
               Top             =   108
               Width           =   1056
            End
         End
         Begin MSDataListLib.DataCombo dcBranch 
            Height          =   288
            Left            =   4680
            TabIndex        =   95
            Top             =   240
            Width           =   1764
            _ExtentX        =   3122
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   " «—ÌŒ «·⁄ﬁœ Â‹ "
            ForeColor       =   &H00000000&
            Height          =   252
            Left            =   3480
            TabIndex        =   97
            Top             =   600
            Width           =   972
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·›—⁄"
            Height          =   312
            Index           =   24
            Left            =   6708
            RightToLeft     =   -1  'True
            TabIndex        =   96
            Top             =   240
            Width           =   468
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ √Ì«„ «·”‰…"
            Height          =   288
            Index           =   17
            Left            =   9060
            RightToLeft     =   -1  'True
            TabIndex        =   90
            Top             =   1440
            Width           =   1152
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ ”‰Ê«  «·⁄ﬁœ"
            Height          =   288
            Index           =   7
            Left            =   11820
            RightToLeft     =   -1  'True
            TabIndex        =   88
            Top             =   1440
            Width           =   1152
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·⁄„Ì·"
            Height          =   288
            Index           =   1
            Left            =   6432
            RightToLeft     =   -1  'True
            TabIndex        =   75
            Top             =   636
            Width           =   720
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·«œ«—… «· ⁄·Ì„Ì…"
            Height          =   396
            Index           =   9
            Left            =   9264
            RightToLeft     =   -1  'True
            TabIndex        =   33
            Top             =   1056
            Width           =   996
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„Õ«›Ÿ…"
            Height          =   276
            Index           =   10
            Left            =   11964
            RightToLeft     =   -1  'True
            TabIndex        =   32
            Top             =   1056
            Width           =   924
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   " «—ÌŒ «·⁄ﬁœ „"
            ForeColor       =   &H00000000&
            Height          =   252
            Left            =   3480
            TabIndex        =   29
            Top             =   240
            Width           =   972
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·Ï  «—ÌŒ"
            ForeColor       =   &H00000000&
            Height          =   120
            Index           =   0
            Left            =   72
            TabIndex        =   28
            Top             =   1260
            Visible         =   0   'False
            Width           =   372
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„”„Ï «· ⁄«ﬁœ"
            Height          =   288
            Index           =   3
            Left            =   11736
            RightToLeft     =   -1  'True
            TabIndex        =   27
            Top             =   636
            Width           =   1152
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ ⁄ﬁœ «·Ê“«—…"
            Height          =   252
            Index           =   15
            Left            =   9264
            RightToLeft     =   -1  'True
            TabIndex        =   26
            Top             =   216
            Width           =   996
         End
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·⁄ﬁœ"
            Height          =   252
            Index           =   0
            Left            =   12012
            RightToLeft     =   -1  'True
            TabIndex        =   25
            Top             =   216
            Width           =   876
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic1 
         Height          =   3132
         Left            =   13860
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   720
         Visible         =   0   'False
         Width           =   4728
         _cx             =   8334
         _cy             =   5530
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
         Begin VB.TextBox txtYear 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1080
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   74
            Top             =   240
            Width           =   2388
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   312
            Left            =   2280
            TabIndex        =   35
            TabStop         =   0   'False
            Top             =   600
            Width           =   1212
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CustomFormat    =   "yyyy/M/d"
            Format          =   98304003
            CurrentDate     =   37140
         End
         Begin Dynamic_Byte.NourHijriCal NourHijriCal1 
            Height          =   312
            Left            =   1080
            TabIndex        =   36
            Top             =   600
            Width           =   1092
            _ExtentX        =   1931
            _ExtentY        =   556
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   312
            Left            =   2280
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   960
            Width           =   1212
            _ExtentX        =   2143
            _ExtentY        =   556
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CustomFormat    =   "yyyy/M/d"
            Format          =   98304003
            CurrentDate     =   37140
         End
         Begin Dynamic_Byte.NourHijriCal NourHijriCal2 
            Height          =   312
            Left            =   1080
            TabIndex        =   38
            Top             =   960
            Width           =   1092
            _ExtentX        =   1931
            _ExtentY        =   556
         End
         Begin VSFlex8UCtl.VSFlexGrid fg_Year 
            Height          =   1692
            Left            =   120
            TabIndex        =   39
            Top             =   1320
            Width           =   4380
            _cx             =   7726
            _cy             =   2984
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
            Rows            =   1
            Cols            =   6
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"FrmMinistryContract.frx":6D72
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
         Begin MSComCtl2.DTPicker DtatAdd 
            Height          =   276
            Left            =   120
            TabIndex        =   66
            TabStop         =   0   'False
            Top             =   -120
            Visible         =   0   'False
            Width           =   612
            _ExtentX        =   1085
            _ExtentY        =   476
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CustomFormat    =   "yyyy/MM/dd"
            Format          =   98304003
            CurrentDate     =   41640
         End
         Begin ImpulseButton.ISButton Cmd 
            Height          =   456
            Index           =   8
            Left            =   120
            TabIndex        =   67
            Top             =   240
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   794
            ButtonPositionImage=   1
            Caption         =   "«÷«›…"
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
            ButtonImage     =   "FrmMinistryContract.frx":6E63
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
         Begin VB.Label Lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·⁄«„ «·œ—«”Ï"
            Height          =   312
            Index           =   16
            Left            =   3480
            RightToLeft     =   -1  'True
            TabIndex        =   42
            Top             =   240
            Width           =   972
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Ì»œ√ „‰ "
            ForeColor       =   &H00000000&
            Height          =   252
            Left            =   3600
            TabIndex        =   41
            Top             =   600
            Width           =   816
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Ì‰ ÂÏ ›Ï"
            ForeColor       =   &H00000000&
            Height          =   252
            Left            =   3600
            TabIndex        =   40
            Top             =   960
            Width           =   816
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic7 
         Height          =   690
         Left            =   120
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   8415
         Width           =   13170
         _cx             =   23230
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
         Begin ImpulseButton.ISButton Cmd 
            Height          =   456
            Index           =   0
            Left            =   11508
            TabIndex        =   58
            Top             =   120
            Width           =   1404
            _ExtentX        =   2487
            _ExtentY        =   794
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
            ButtonImage     =   "FrmMinistryContract.frx":D6C5
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
            Height          =   456
            Index           =   1
            Left            =   10020
            TabIndex        =   59
            Top             =   120
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   794
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
            ButtonImage     =   "FrmMinistryContract.frx":13F27
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
            Height          =   456
            Index           =   2
            Left            =   8496
            TabIndex        =   60
            Top             =   120
            Width           =   1524
            _ExtentX        =   2699
            _ExtentY        =   794
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
            ButtonImage     =   "FrmMinistryContract.frx":1A789
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
            Height          =   456
            Index           =   3
            Left            =   7044
            TabIndex        =   61
            Top             =   120
            Width           =   1368
            _ExtentX        =   2408
            _ExtentY        =   794
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
            ButtonImage     =   "FrmMinistryContract.frx":20FEB
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
            Height          =   456
            Index           =   4
            Left            =   5616
            TabIndex        =   62
            Top             =   120
            Width           =   1356
            _ExtentX        =   2381
            _ExtentY        =   794
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
            ButtonImage     =   "FrmMinistryContract.frx":2784D
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
            Height          =   450
            Index           =   6
            Left            =   2745
            TabIndex        =   63
            Top             =   120
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   794
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
            ButtonImage     =   "FrmMinistryContract.frx":2E0AF
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   -2147483637
         End
         Begin ImpulseButton.ISButton CmdAttach 
            Height          =   450
            Left            =   1320
            TabIndex        =   64
            Top             =   120
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   794
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
            ButtonImage     =   "FrmMinistryContract.frx":57CD1
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
            Height          =   90
            Index           =   7
            Left            =   4305
            TabIndex        =   65
            Top             =   120
            Visible         =   0   'False
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   159
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
            ButtonImage     =   "FrmMinistryContract.frx":5E533
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
            Height          =   450
            Index           =   9
            Left            =   4080
            TabIndex        =   76
            Top             =   120
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   794
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
            ButtonImage     =   "FrmMinistryContract.frx":64D95
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   -2147483637
         End
      End
      Begin C1SizerLibCtl.C1Elastic EleHeader 
         Height          =   588
         Left            =   0
         TabIndex        =   68
         TabStop         =   0   'False
         Top             =   0
         Width           =   13560
         _cx             =   23918
         _cy             =   1032
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
         Appearance      =   0
         MousePointer    =   0
         Version         =   801
         BackColor       =   16777215
         ForeColor       =   4210688
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   "       «· ⁄«ﬁœ „⁄ «·Ê“«—…     "
         Align           =   0
         AutoSizeChildren=   0
         BorderWidth     =   2
         ChildSpacing    =   1
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
         CaptionStyle    =   1
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
            Height          =   345
            Left            =   12960
            RightToLeft     =   -1  'True
            TabIndex        =   69
            Top             =   1320
            Visible         =   0   'False
            Width           =   855
         End
         Begin ImpulseButton.ISButton XPBtnMove 
            Height          =   345
            Index           =   0
            Left            =   1155
            TabIndex        =   70
            Top             =   120
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   609
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
            ButtonImage     =   "FrmMinistryContract.frx":6B5F7
            ColorButton     =   16777215
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
            Left            =   90
            TabIndex        =   71
            Top             =   120
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   609
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
            ButtonImage     =   "FrmMinistryContract.frx":6B991
            ColorButton     =   16777215
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
            Left            =   1680
            TabIndex        =   72
            Top             =   120
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   609
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
            ButtonImage     =   "FrmMinistryContract.frx":6BD2B
            ColorButton     =   16777215
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
            Left            =   615
            TabIndex        =   73
            Top             =   120
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   609
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
            ButtonImage     =   "FrmMinistryContract.frx":6C0C5
            ColorButton     =   16777215
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
   End
End
Attribute VB_Name = "FrmMinistryContract"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit
Dim rs As ADODB.Recordset
Dim Rs_Temp5 As ADODB.Recordset
Dim TTP As clstooltip
Dim rsInst As ADODB.Recordset
Dim Rs_Temp As ADODB.Recordset
Dim Operation As String
 Dim rsYrs As ADODB.Recordset

Private Sub Cmd_Click(Index As Integer)
 '    On Error GoTo ErrTrap

    Select Case Index
        Case 0
            If DoPremis(Do_New, Me.Name, True) = False Then
                Exit Sub
            End If
            TxtModFlg.Text = "N"
            clear_all Me
          
            XPTxtBoxName.SetFocus
           FgInstallments.Rows = 1
            txtIDMC.Text = CStr(new_id("TblMinistryContract", "IDMC", "", True))
        Case 1

            If DoPremis(Do_Edit, Me.Name, True) = False Then
                Exit Sub
            End If
            
         
                       TxtModFlg.Text = "E"
        
            
        Case 2
            SaveData
        Case 3
            Undo
        Case 4
            If DoPremis(Do_Delete, Me.Name, True) = False Then
                Exit Sub
            End If
            If ISAllowDeleteUpdateContract(val(txtIDMC.Text)) = False Then
                        MsgBox ("·« Ì„ﬂ‰ Õ–› «·⁄ﬁœ »”»» «Ã—«¡ ⁄„·Ì«  ⁄·Ï «·⁄ﬁœ ")
            Else
                        Del_Company
            End If
        Case 5
                AddRowToGrid
                'Calculations
                numbering
        Case 6
                Unload Me
         Case 7
'        print_report2
   
   Case 8
             AddYear
             Case 9
             Unload FrmSearch_MinistryContract
             FrmSearch_MinistryContract.SendForm = "MC"
             FrmSearch_MinistryContract.show
             'FrmSearch_ClosedRequerMainten.show
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub numbering()
Dim i As Integer

With FgInstallments
    For i = .FixedRows To .Rows - 1
            .TextMatrix(i, .ColIndex("QestID")) = i
    Next
End With

End Sub


Private Sub AddRowToGrid()

If val(txtValue) <= 0 Then   '
MsgBox ("«œŒ· ﬁÌ„… «·œ›⁄…")
Exit Sub
End If

Dim i As Integer
With FgInstallments
i = .Rows
.Rows = i + 1
.TextMatrix(i, .ColIndex("QestID")) = i
.TextMatrix(i, .ColIndex("Value")) = val(txtValue.Text)
.TextMatrix(i, .ColIndex("Due_DateH")) = Format(FirstPaymentDateH.value, "yyyy/MM/dd")
.TextMatrix(i, .ColIndex("Due_Date")) = (FirstPaymentDate.value)
End With
txtValue.Text = ""
FirstPaymentDateH.value = ToHijriDate(Date)
FirstPaymentDate.value = Date
End Sub


Private Sub AddYear()



Dim i As Integer
fg_Year.Rows = fg_Year.Rows + 1
i = fg_Year.Rows
i = i - 1
With fg_Year
  .TextMatrix(i, .ColIndex("Serial")) = i - 1
  .TextMatrix(i, .ColIndex("Year")) = txtYear.Text
  .TextMatrix(i, .ColIndex("FromDate")) = DTPicker1.value
  .TextMatrix(i, .ColIndex("ToDate")) = DTPicker2.value
  .TextMatrix(i, .ColIndex("FromDateH")) = NourHijriCal1.value
  .TextMatrix(i, .ColIndex("ToDateH")) = NourHijriCal2.value
       
End With
txtYear.Text = ""

End Sub


Private Sub CmdHelp_Click()
    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
    SystemOptions.SysHelp.HHDisplayTopicID Me.hwnd
End Sub


Private Sub CmdAttach_Click()
       If DoPremis(Do_Attach, Me.Name, True) = False Then
                Exit Sub
            End If
            
ShowAttachments txtIDMC, "10062020003"

End Sub

Private Sub Command1_Click()
If FgInstallments.Row < FgInstallments.FixedRows Then Exit Sub:

If FgInstallments.TextMatrix(FgInstallments.Row, FgInstallments.ColIndex("Status")) = "1" Then
            MsgBox ("·«Ì„ﬂ‰ Õ–› «·œ›⁄Â »”»»  ”ÃÌ· ”œ«œ ·Â« ")
            Exit Sub
End If
FgInstallments.RemoveItem (FgInstallments.Row)
numbering
End Sub

Private Sub Command2_Click()
Dim i As Integer, cnt As Integer
If FgInstallments.Rows <= FgInstallments.FixedRows Then Exit Sub:
cnt = FgInstallments.Rows - 1
 
For i = FgInstallments.FixedRows To cnt
        If FgInstallments.TextMatrix(i, FgInstallments.ColIndex("Status")) = "1" Then
                 '   MsgBox ("·«Ì„ﬂ‰ Õ–› «·œ›⁄Â »”»»  ”ÃÌ· ”œ«œ ·Â« ")
                '    Exit Sub
        Else
                    FgInstallments.RemoveItem (i)
                    cnt = cnt - 1
                    Command2_Click
                    Exit Sub
        End If
Next

numbering
End Sub

Private Sub Command3_Click()
txtNet.Text = val(txtTotal.Text) - val(txtDiscount.Text)
End Sub

Private Sub cmdAdd_Click()
Operation = "add"
txtDiscount.Enabled = True
End Sub

Private Sub cmdSub_Click()
Operation = "sub"
txtDiscount.Enabled = True
End Sub

Private Sub dcCity_Change()
Dim str As String
'If DcCity.BoundText = "" Then Exit Sub

' Set dcVendor.RowSource = rs
Set Rs_Temp = New ADODB.Recordset
Set dcVendor.RowSource = Rs_Temp

If SystemOptions.UserInterface = ArabicInterface Then
    str = " Select ID , Name   from TblManagerialArea  where cityid = " & val(dcCity.BoundText)
Else
    str = " Select ID , NameE   from TblManagerialArea  where cityid = " & val(dcCity.BoundText)
End If
    fill_combo dcVendor, str
dcVendor.Refresh
End Sub


Private Sub dcVendor_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
            Unload FrmSearch_BasicData
            FrmSearch_BasicData.SendForm = "MCMA"
            FrmSearch_BasicData.show
End If
End Sub

Private Sub dtpEContractDate_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)
   
        dtpEContractDateH.value = ToHijriDate(dtpEContractDate.value)
        
End Sub

Private Sub dtpEContractDateH_GotFocus()
  VBA.Calendar = vbCalGreg
        dtpEContractDate.value = ToGregorianDate(dtpEContractDateH.value)
End Sub

Private Sub dtpFromDate_Change()
  '   If Me.TxtModFlg.text <> "R" Then
        dtpFromDateH.value = ToHijriDate(dtpFromDate.value)
  '   End If
End Sub



Private Sub dtpFromDateH_LostFocus()
 'If Me.TxtModFlg.text <> "R" Then
              VBA.Calendar = vbCalGreg
            dtpFromDate.value = ToGregorianDate(dtpFromDateH.value)
 '       End If
End Sub

Private Sub DTPicker1_Change()
 '   If Me.TxtModFlg.text <> "R" Then
        NourHijriCal1.value = ToHijriDate(DTPicker1.value)
 '    End If
End Sub
Private Sub DTPicker2_Change()
 '    If Me.TxtModFlg.text <> "R" Then
        NourHijriCal2.value = ToHijriDate(DTPicker2.value)
 '    End If
End Sub

Private Sub dtpSContractDate_Change()
 '    If Me.TxtModFlg.text <> "R" Then
        dtpSContractDateH.value = ToHijriDate(dtpSContractDate.value)
 '    End If
End Sub

Private Sub dtpSContractDateH_LostFocus()
On Error Resume Next
If Me.TxtModFlg.Text <> "R" Then
        VBA.Calendar = vbCalGreg
        dtpSContractDate.value = ToGregorianDate(dtpSContractDateH.value)
End If
        
End Sub

Private Sub dtpToDate_Change()
    If Me.TxtModFlg.Text <> "R" Then
        dtpToDateH.value = ToHijriDate(dtpToDate.value)
     End If
End Sub


Private Sub dtpToDateH_LostFocus()
If Me.TxtModFlg.Text <> "R" Then
              VBA.Calendar = vbCalGreg
            dtpToDate.value = ToGregorianDate(dtpToDateH.value)
        End If
End Sub



Private Sub FgInstallments_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

With FgInstallments
    If .TextMatrix(Row, .ColIndex("Status")) = "1" Then
            Cancel = True
    End If

End With

End Sub

Private Sub FirstPaymentDate_Change()
 If Me.TxtModFlg.Text <> "R" Then
        FirstPaymentDateH.value = ToHijriDate(FirstPaymentDate.value)
     End If
End Sub



Private Sub FirstPaymentDateH_LostFocus()

 If Me.TxtModFlg.Text <> "R" Then
              VBA.Calendar = vbCalGreg
            FirstPaymentDate.value = ToGregorianDate(FirstPaymentDateH.value)
        End If

End Sub

Private Sub Form_Activate()
'    XPTxtBoxID.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        If Me.TxtModFlg.Text = "R" Then
            Cmd_Click (0)
        Else
            SendKeys "{TAB}"
        End If
    End If

    If Me.TxtModFlg.Text = "R" Then
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
 '   On Error GoTo ErrTrap
 
    Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos
    Dcombos.getCountriesGovernments dcCity
    Dcombos.GetBranches dcBranch
    Dcombos.GetCustomersSuppliers 1, dcClient
    
    
    Dim str As String
    
    
    
    
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    LogTextA = "   «·œŒÊ· «·Ì ‘«‘… " & " «· ⁄«ﬁœ „⁄ Ê“«—…  "
    LogTexte = " Open Window " & "  Ministry Contract"
    AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, "O", "", ""

    Dim My_SQL As String
    
    AddTip
    Set rs = New ADODB.Recordset
   Dim StrSQL As String
   StrSQL = "SELECT  *  From TblMinistryContract"
   rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
      
        
    Me.TxtModFlg.Text = "R"
    XPBtnMove_Click 2
    If OPEN_NEW_SCREEN = True Then
        Cmd_Click (0)
    End If

   ' Exit Sub

dtpFromDate.value = Date
dtpSContractDate.value = Date
dtpToDate.value = Date
dtpEContractDate.value = Date
DTPicker1.value = Date
DTPicker2.value = Date

ErrTrap:
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)
    Dim IntResult As String
    Dim StrMSG As String
    On Error GoTo ErrTrap

    If Me.TxtModFlg.Text <> "R" Then

        Select Case Me.TxtModFlg.Text

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

Private Sub ChangeLang()
    Dim XPic As IPictureDisp

    Set XPic = Me.XPBtnMove(1).ButtonImage
    Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    Set Me.XPBtnMove(2).ButtonImage = XPic

 
   Lbl(0).Caption = "No."
   Lbl(3).Caption = " Name Ar"
   Lbl(7).Caption = " Name En"
   Label3.Caption = "City"
   
  Lbl(2).Caption = "Current Record"
  Lbl(4).Caption = "Recors Count"
   
    Me.Caption = "Managerial Area"
    EleHeader.Caption = Me.Caption
   
    Me.Cmd(0).Caption = "New"
    Me.Cmd(1).Caption = "Edit"
    Me.Cmd(2).Caption = "Save"
    Me.Cmd(3).Caption = "Undo"
    Me.Cmd(4).Caption = "Delete"
 
    Me.Cmd(6).Caption = "Exit"
    Me.Cmd(7).Caption = "Print"


End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap
    LogTextA = "     «·Œ—ÊÃ „‰ ‘«‘… " & "  »Ì«‰«  «· ⁄«ﬁœ „⁄ Ê“«—…  "
    LogTexte = " Exit Window " & "  Boxes Data "
    AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, "O", "", ""

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
    Exit Sub
ErrTrap:
End Sub


Private Sub NourHijriCal3_LostFocus()
If Me.TxtModFlg.Text <> "R" Then
              VBA.Calendar = vbCalGreg
            'FirstPaymentDate.value = ToGregorianDate(NourHijriCal3.value)
        End If
End Sub

Private Sub txtContractYears_Change()
txtTotal.Text = val(txtYearCustom.Text) * val(txtContractYears.Text) * val(txtStudentCount.Text)
End Sub

Private Sub txtDiscount_Change()

If Operation = "add" Then
    txtNet.Text = val(txtTotal.Text) + val(txtDiscount.Text)
ElseIf Operation = "sub" Then
    txtNet.Text = val(txtTotal.Text) - val(txtDiscount.Text)
End If

End Sub

Private Sub TxtModFlg_Change()
    On Error GoTo ErrTrap

    Select Case Me.TxtModFlg.Text

        Case "R"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "«· ⁄«ﬁœ „⁄ «·Ê“«—…"
            Else
                Me.Caption = "Ministry Contract"
            End If

            Me.Cmd(2).Enabled = False
            Me.Cmd(3).Enabled = False
        
            Me.Cmd(0).Enabled = True
            Me.Cmd(1).Enabled = True
            Me.Cmd(4).Enabled = True
         Me.Cmd(9).Enabled = True
         
            Me.XPBtnMove(0).Enabled = True
            Me.XPBtnMove(1).Enabled = True
            Me.XPBtnMove(2).Enabled = True
            Me.XPBtnMove(3).Enabled = True
            If rs.RecordCount < 1 Then
                Me.XPBtnMove(0).Enabled = False
                Me.XPBtnMove(1).Enabled = False
                Me.XPBtnMove(2).Enabled = False
                Me.XPBtnMove(3).Enabled = False
                Me.Cmd(1).Enabled = False
                Me.Cmd(4).Enabled = False
            End If
            
               C1Elastic3.Enabled = False
               C1Elastic4.Enabled = False
               C1Elastic5.Enabled = False
        Case "N"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«  «· ⁄«ﬁœ „⁄ «·Ê“«—… ( ÃœÌœ )"
            Else
                Me.Caption = "Boxes Data(New)"
            End If
        
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "«· ⁄«ﬁœ „⁄ «·Ê“«—…"
            Else
                Me.Caption = "Ministry Contarct"
            End If
        
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
            Me.Cmd(9).Enabled = False
            
            txtProcessNo.locked = False
            Me.XPTxtBoxName.locked = False
            
               C1Elastic3.Enabled = True
               C1Elastic4.Enabled = True
               C1Elastic5.Enabled = True
        Case "E"

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.Caption = "»Ì«‰«  «· ⁄«ﬁœ „⁄ «·Ê“«—…(  ⁄œÌ· )"
            Else
                Me.Caption = "Boxes Data(Edit)"
            End If
        
            Me.Cmd(2).Enabled = True
            Me.Cmd(3).Enabled = True
        
            Me.Cmd(0).Enabled = False
            Me.Cmd(1).Enabled = False
            Me.Cmd(4).Enabled = False
        Me.Cmd(9).Enabled = False
            
            Me.XPBtnMove(0).Enabled = False
            Me.XPBtnMove(1).Enabled = False
            Me.XPBtnMove(2).Enabled = False
            Me.XPBtnMove(3).Enabled = False
        
            txtProcessNo.locked = True
            Me.XPTxtBoxName.locked = False
       '     Me.XPMTxtRemark.locked = False
       
          C1Elastic3.Enabled = True
               C1Elastic4.Enabled = True
               C1Elastic5.Enabled = True
    End Select

    Exit Sub
ErrTrap:
End Sub
Public Sub Retrive(Optional Lngid As Long = 0)
    'On Error GoTo ErrTrap
    If rs.RecordCount < 1 Then
        XPTxtCurrent.Caption = 0
        XPTxtCount.Caption = 0
        Exit Sub
    End If

    If rs.EOF Or rs.BOF Then
        Exit Sub
    Else
        If Lngid <> 0 Then
            rs.find "IDMC =" & Lngid, , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Exit Sub
            End If
        End If
    End If

    If IsNull(rs("AdditionalType").value) Or rs("AdditionalType").value = "" Then
        txtDiscount.Enabled = False
    Else
        Operation = rs("AdditionalType").value
        txtDiscount.Enabled = True
    End If
    
    FgInstallments.Rows = 1
    
    txtIDMC.Text = IIf(IsNull(rs("IDMC").value), "", rs("IDMC").value)
    txtProcessNo.Text = txtIDMC.Text
    XPTxtBoxName.Text = IIf(IsNull(rs("Name").value), "", Trim(rs("Name").value))
     
    
    dcCity.BoundText = IIf(IsNull(rs("CityID").value), "", rs("CityID").value)
    dcVendor.BoundText = IIf(IsNull(rs("VendorID").value), "", rs("VendorID").value)
    dtpFromDate.value = IIf(IsNull(rs("FromDate").value), Date, rs("FromDate").value)
    dtpToDate.value = IIf(IsNull(rs("ToDate").value), Date, rs("ToDate").value)
    dtpSContractDate.value = IIf(IsNull(rs("StartContractDate").value), Date, rs("StartContractDate").value)
    dtpEContractDate.value = IIf(IsNull(rs("EndContractDate").value), Date, rs("EndContractDate").value)
    dtpFromDateH.value = IIf(IsNull(rs("FromDateh").value), Date, rs("FromDateh").value)
    dtpToDateH.value = IIf(IsNull(rs("ToDateh").value), Date, rs("ToDateh").value)
    dtpSContractDateH.value = IIf(IsNull(rs("StartContractDateh").value), Date, rs("StartContractDateh").value)
    dtpEContractDateH.value = IIf(IsNull(rs("EndContractDateh").value), Date, rs("EndContractDateh").value)
     txtStudentCount.Text = IIf(IsNull(rs("StudentCount").value), "", Trim(rs("StudentCount").value))
     txtStudentCustom.Text = IIf(IsNull(rs("StudentCustom").value), "", Trim(rs("StudentCustom").value))
     txtTotal.Text = val(txtStudentCustom.Text) * val(txtStudentCount.Text)
     txtDiscount.Text = IIf(IsNull(rs("Discount").value), "", Trim(rs("Discount").value))
     txtContractYears.Text = IIf(IsNull(rs("ContractYears").value), "", rs("ContractYears").value)
     txtYearDays.Text = IIf(IsNull(rs("YearDays").value), "", rs("YearDays").value)
     dcBranch.BoundText = IIf(IsNull(rs("BranchID").value), "", rs("BranchID").value)
     
   If Operation = "add" Then
      txtNet.Text = val(txtTotal.Text) + val(txtDiscount.Text)
   ElseIf Operation = "sub" Then
     txtNet.Text = val(txtTotal.Text) - val(txtDiscount.Text)
   End If
     TxtPaymentCount.Text = IIf(IsNull(rs("PaymentCount").value), "", Trim(rs("PaymentCount").value))
     txtMinistryContractNo.Text = IIf(IsNull(rs("MinistryContractNo").value), "", rs("MinistryContractNo").value)
    
    txtYearCustom.Text = IIf(IsNull(rs("YearCustom").value), "", rs("YearCustom").value)
    txtTotal.Text = IIf(IsNull(rs("Total").value), "", rs("Total").value)
    txtNet.Text = IIf(IsNull(rs("Net").value), "", rs("Net").value)
    dcClient.BoundText = IIf(IsNull(rs("ClientID").value), "", rs("ClientID").value)
    
    txtStudentCustom_Change
    
    Set rsInst = New ADODB.Recordset
    Dim StrSQL As String, paid As Boolean
    StrSQL = " SELECT * from TblMinistryContract_Installment where type = 1 and  idmc =  " & val(txtIDMC.Text)
    rsInst.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If rsInst.RecordCount > 0 Then
    rsInst.MoveFirst
     With FgInstallments
        FgInstallments.Rows = rsInst.RecordCount + 1
        Dim j As Integer
        For j = 1 To FgInstallments.Rows - 1
        .TextMatrix(j, .ColIndex("Serial")) = j
        .TextMatrix(j, .ColIndex("ID")) = IIf(IsNull(rsInst("ID").value), "", rsInst("ID").value)
        .TextMatrix(j, .ColIndex("QestID")) = IIf(IsNull(rsInst("InstallmentNo").value), "", rsInst("InstallmentNo").value)
        .TextMatrix(j, .ColIndex("value")) = IIf(IsNull(rsInst("Value").value), 0, rsInst("Value").value)
        .TextMatrix(j, .ColIndex("Due_Date")) = IIf(IsNull(rsInst("Due_Date").value), Date, rsInst("Due_Date").value)
        .TextMatrix(j, .ColIndex("Due_DateH")) = IIf(IsNull(rsInst("Due_DateH").value), "", rsInst("Due_DateH").value)
        paid = IIf(IsNull(rsInst("Paid").value), False, rsInst("Paid").value)
        
        If paid = True Then
                    .TextMatrix(j, .ColIndex("Paid")) = " „ «·«” Õﬁ«ﬁ"
                    .TextMatrix(j, .ColIndex("Status")) = 1
        Else
                    .TextMatrix(j, .ColIndex("Status")) = 0
        End If
        
        rsInst.MoveNext
         Next
        End With
    End If
     
  
    XPTxtCurrent.Caption = rs.AbsolutePosition
    XPTxtCount.Caption = rs.RecordCount
    Exit Sub
ErrTrap:
End Sub

Private Sub txtStudentCount_Change()
txtTotal.Text = val(txtYearCustom.Text) * val(txtContractYears.Text) * val(txtStudentCount.Text)
End Sub

Private Sub txtStudentCustom_Change()
txtYearCustom.Text = val(txtStudentCustom.Text) * val(txtYearDays.Text)
End Sub

Private Sub TxtTotal_Change()
txtNet.Text = val(txtTotal.Text)
End Sub

Private Sub txtYearCustom_Change()
txtTotal.Text = val(txtYearCustom.Text) * val(txtContractYears.Text) * val(txtStudentCount.Text)
End Sub

Private Sub txtYearDays_Change()
txtYearCustom.Text = val(txtStudentCustom.Text) * val(txtYearDays.Text)
End Sub

Private Sub XPBtnMove_Click(Index As Integer)

    'On Error GoTo ErrTrap
    If Me.TxtModFlg.Text = "N" Then
        clear_all Me
        Me.TxtModFlg.Text = "R"
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

Function CuurentLogdata(Optional Currentmode As String)
   

End Function
 
Private Sub SaveData()
    Dim Msg As String
    Dim StrSQL As String
    Dim RsTemp As New ADODB.Recordset
    Dim RsTempM As New ADODB.Recordset
    Dim BeginTrans As Boolean
   ' On Error GoTo ErrTrap

    If Me.TxtModFlg.Text <> "R" Then
    
        If XPTxtBoxName.Text = "" Then
            If SystemOptions.UserInterface = ArabicInterface Then
                 MsgBox "„‰ ›÷·ﬂ √œŒ· „”„Ï «· ⁄«ﬁœ  ", vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Else
                 MsgBox "Please Entrer Contract Name  ", vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            End If
            XPTxtBoxName.SetFocus
            Exit Sub
        End If
        
       If dcCity.BoundText = "" Then
                 If SystemOptions.UserInterface = ArabicInterface Then
                           MsgBox "«Œ — «·„Õ«›Ÿ…  ", vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                 Else
                           MsgBox "select city  ", vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                 End If
                dcCity.SetFocus
                
                Exit Sub
               
       End If
        
        
       If dcVendor.BoundText = "" Then
                 If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "«Œ — «·«œ«—… «· ⁄·Ì„Ì…  ", vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                 Else
                    MsgBox "select Managerial area   ", vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                 End If
                 dcVendor.SetFocus
                 'SendKeys ("{F4}")
                  Exit Sub
        End If
        
       
         
         If txtMinistryContractNo.Text = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox " «œŒ· —ﬁ„ ⁄ﬁœ «·Ê“«—…  ", vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Else
                    MsgBox "Enter Ministry Contract No.   ", vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                End If
                txtMinistryContractNo.SetFocus
                Exit Sub
        End If
        
        
        If dcBranch.BoundText = "" Then
             MsgBox ("„‰ ›÷··ﬂ «Œ — «·›—⁄ «Ê·« ")
             Exit Sub
        End If
         
          If dcClient.BoundText = "" Then
             MsgBox ("„‰ ›÷··ﬂ «Œ — «·⁄„Ì·  «Ê·« ")
             Exit Sub
        End If
         
         If val(txtContractYears.Text) <= 0 Then
            MsgBox "«œŒ· ⁄œœ ”‰Ê«  «·⁄ﬁœ"
            Exit Sub
       End If
       
      If val(txtYearDays.Text) <= 0 Then
            MsgBox "«œŒ· ⁄œœ «Ì«„ «·”‰…"
            Exit Sub
       End If
    
         If val(txtStudentCount.Text) <= 0 Then
            MsgBox "«œŒ· ⁄œœ «·ÿ·«»"
            Exit Sub
       End If
    
       If val(txtStudentCustom.Text) <= 0 Then
            MsgBox "«œŒ· „Œ’’ ﬂ· ÿ«·»"
            Exit Sub
       End If
         
        If FgInstallments.Rows <= 1 Then
                 MsgBox " ·„ Ì „ «œŒ«· «Ï œ›⁄… "
            Exit Sub
        End If
        
        
        Select Case Me.TxtModFlg.Text
            Case "N"
                StrSQL = "select * From  TblMinistryContract  where Name ='" & Trim(XPTxtBoxName.Text) & "'"
                Set RsTemp = New ADODB.Recordset
                RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                If RsTemp.RecordCount > 0 Then
                    Msg = "Â‰«ﬂ   ⁄«ﬁœ  „”Ã· „”»ﬁ« »Â–« «·„”„Ï" & CHR(13)
                    Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·«”„ «·’ÕÌÕ " & CHR(13)
                    Msg = Msg + "√Ê  €ÌÌ— √Ê  „ÌÌ“ «”„ «·„‰ÿﬁ…"
                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                    XPTxtBoxName.SetFocus
                    Exit Sub
                End If
                
                
                StrSQL = "select * From  TblMinistryContract where MinistryContractNo ='" & Trim(txtMinistryContractNo.Text) & "'"
                Set RsTemp = New ADODB.Recordset
                RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                If RsTemp.RecordCount > 0 Then
                        Msg = " Â–« «·—ﬁ„ «·Ê“«—Ï „”Ã· „‰ ﬁ»· " & CHR(13)
                        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·—ﬁ„ «·’ÕÌÕ " & CHR(13)
                        Msg = Msg + "√Ê  €ÌÌ— √Ê  „ÌÌ“ «·—ﬁ„ "
                        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                        XPTxtBoxName.SetFocus
                        Exit Sub
                End If
                 
          
               ' Calculations
                       
             rs.AddNew
             txtIDMC.Text = CStr(new_id("TblMinistryContract", "IDMC", "", True))
             
            Case "E"
                StrSQL = "select * From  TblMinistryContract where Name='" & Trim(XPTxtBoxName.Text) & "'"
                RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                If RsTemp.RecordCount > 0 Then
                    If RsTemp("IDMC").value <> val(txtIDMC.Text) Then
                        Msg = "Â‰«ﬂ  ⁄«ﬁœ  „”Ã·Â „”»ﬁ« »Â–« «·„”„Ï" & CHR(13)
                        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·„”„Ï «·’ÕÌÕ " & CHR(13)
                        Msg = Msg + "√Ê  €ÌÌ— √Ê  „ÌÌ“ «·„”„Ï "
                        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                        XPTxtBoxName.SetFocus
                        Exit Sub
                    End If
                End If
                
                StrSQL = "select * From  TblMinistryContract where MinistryContractNo ='" & Trim(txtMinistryContractNo.Text) & "'"
                Set RsTemp = New ADODB.Recordset
                RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

                If RsTemp.RecordCount > 0 Then
                    If RsTemp("IDMC").value <> val(txtIDMC.Text) Then
                        Msg = " Â–« «·—ﬁ„ «·Ê“«—Ï „”Ã· „‰ ﬁ»· " & CHR(13)
                        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·—ﬁ„ «·’ÕÌÕ " & CHR(13)
                        Msg = Msg + "√Ê  €ÌÌ— √Ê  „ÌÌ“ «·—ﬁ„ "
                        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                        XPTxtBoxName.SetFocus
                        Exit Sub
                    End If
                End If
                
              
        End Select

        Cn.BeginTrans
        BeginTrans = True
        If TxtModFlg.Text = "E" Then
           '      strSQL = "delete from TblMinistryContract_Installment where type=1 and idmc =  " & val(txtIDMC.text)
           '      Cn.Execute strSQL, , adExecuteNoRecords
        End If
    
        rs("IDMC").value = txtIDMC.Text
        rs("ProcessNo").value = txtIDMC.Text
        rs("Name").value = IIf(Trim(XPTxtBoxName.Text) = "", Null, XPTxtBoxName.Text)
      
        rs("FromDate").value = dtpFromDate.value
        rs("ToDate").value = dtpToDate.value
        rs("FromDateH").value = dtpFromDateH.value
        rs("ToDateH").value = dtpToDateH.value
          
        rs("StartContractDate").value = dtpSContractDate.value
        rs("EndContractDate").value = dtpEContractDate.value
        rs("StartContractDateH").value = dtpSContractDateH.value
        rs("EndContractDateH").value = dtpEContractDateH.value
        rs("CityID").value = IIf(dcCity.BoundText = "", Null, val(dcCity.BoundText))
        rs("VendorID").value = IIf(dcVendor.BoundText = "", Null, dcVendor.BoundText)
       
        rs("StudentCount").value = IIf(txtStudentCount.Text = "", 0, val(txtStudentCount.Text))
        rs("StudentCustom").value = IIf(txtStudentCustom.Text = "", 0, val(txtStudentCustom.Text))
        rs("DisCount").value = IIf(txtDiscount.Text = "", 0, val(txtDiscount.Text))
        rs("PaymentCount").value = IIf(TxtPaymentCount.Text = "", 0, val(TxtPaymentCount.Text))
        
        rs("FirstPaymentDate").value = FirstPaymentDate.value
        rs("AdditionalType").value = Operation
        rs("MinistryContractNo").value = txtMinistryContractNo.Text
    
        rs("ContractYears").value = IIf(txtContractYears = "", "", val(txtContractYears.Text))
        rs("YearDays").value = IIf(txtYearDays = "", "", val(txtYearDays.Text))
        rs("BranchID").value = IIf(dcBranch.BoundText = "", Null, dcBranch.BoundText)
        
         rs("Total").value = val(txtTotal.Text)
         rs("Net").value = val(txtNet.Text)
         rs("YearCustom").value = val(txtYearCustom.Text)
         rs("ClientID").value = IIf(IsNull(dcClient.BoundText), Null, dcClient.BoundText)

         rs.update
                
                
        Dim rsIns As ADODB.Recordset
        Set rsIns = New ADODB.Recordset
        StrSQL = "select * from TblMinistryContract_Installment where type = 1 order by id "
        rsIns.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
        With FgInstallments
        Dim j As Integer
       ' FgInstallments.Rows = val(TxtPaymentCount.text) + 1
        Dim AllID As String
        
       ' rsIns.MoveFirst
       
        For j = FgInstallments.FixedRows To FgInstallments.Rows - 1
           If .TextMatrix(j, .ColIndex("QestID")) <> "" Then
           
                    If .TextMatrix(j, .ColIndex("ID")) = "" Then
                            rsIns.AddNew
                            rsIns("ID") = CStr(new_id("TblMinistryContract_Installment", "ID", "", True))
                    Else
                            rsIns.find " ID ='" & val(.TextMatrix(j, .ColIndex("ID"))) & "'", , adSearchForward, adBookmarkFirst
                            
                            If rsIns.EOF Or rsIns.BOF Then
                                    Exit Sub
                            End If
                    End If
                    
                    
                    rsIns("IDMC") = val(txtIDMC.Text)
                    rsIns("InstallmentNo") = .TextMatrix(j, .ColIndex("QestID"))
                    rsIns("Value") = .TextMatrix(j, .ColIndex("value"))
                    rsIns("Due_Date") = .TextMatrix(j, .ColIndex("Due_Date"))
                    rsIns("Due_DateH") = .TextMatrix(j, .ColIndex("Due_DateH"))
                    rsIns("Type") = 1
                    rsIns.update
                    
                    If j = FgInstallments.FixedRows Then
                                AllID = rsIns("ID").value
                    Else
                                AllID = AllID & "  ,  " & CStr(rsIns("ID").value)
                    End If
                    
                 End If
           Next
        End With
        
        
         'Dim strSQL As String
         If AllID <> "" Then
                StrSQL = "delete from TblMinistryContract_Installment where idmc = " & txtIDMC.Text & " and  id not in  ( " & AllID & "  ) "
                 Cn.Execute StrSQL, , adExecuteNoRecords
         End If
                     
                     
        
        Cn.CommitTrans
        BeginTrans = False
        XPTxtCurrent.Caption = rs.AbsolutePosition
        XPTxtCount.Caption = rs.RecordCount
       'CuurentLogdata

        Select Case Me.TxtModFlg.Text

            Case "N"
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "  „ Õ›Ÿ »Ì«‰«  Â–Â «·„‰ÿﬁ… " & CHR(13)
                    Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"
                Else
                    Msg = "Saved" & CHR(13)
                    Msg = Msg + "Do you want enter another One"
                End If
                If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.title) = vbYes Then
                    Cmd_Click (0)
                    Exit Sub
                End If
            
            Case "E"
        
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox " „ Õ›Ÿ Â–Â «· ⁄œÌ·« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Else
                    MsgBox "Saved", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                End If

        End Select

        TxtModFlg.Text = "R"
    End If

    Exit Sub
ErrTrap:

    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If

    If Err.Number = -2147217900 Then
        Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & CHR(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & CHR(13)
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
End Sub

Private Sub Undo()
    On Error GoTo ErrTrap

    Select Case TxtModFlg.Text

        Case "N"
            clear_all Me
            Me.TxtModFlg.Text = "R"
            XPBtnMove_Click (1)

        Case "E"
            rs.find "IDMC ='" & val(txtIDMC.Text) & "'", , adSearchForward, adBookmarkFirst

            If rs.EOF Or rs.BOF Then
                Me.TxtModFlg.Text = "R"
                Exit Sub
            End If

            Retrive
            Me.TxtModFlg.Text = "R"
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub Del_Company()
    Dim Msg As String
    Dim StrSQL As String
    Dim RsTemp As New ADODB.Recordset
    Dim StrAccountCode As String
    Dim StrAccountCode1 As String
    Dim ParentAccount As String
    '
 
    On Error GoTo ErrTrap
            
        If txtIDMC.Text <> "" Then

    
        Msg = "”Ì „ Õ–› »Ì«‰«  «·„‰ÿﬁ… —ﬁ„ " & CHR(13)
        Msg = Msg + (txtIDMC.Text) & CHR(13)
        Msg = Msg + " Â·  —€» ›Ì Õ–› Â–Â «·»Ì«‰« ø"

        If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title) = vbYes Then
            
    
            If Not rs.RecordCount < 1 Then
                StrSQL = "delete From TblMinistryContract where  IDMC =" & val(txtIDMC.Text)
                Cn.Execute StrSQL, , adExecuteNoRecords
                    
                     StrSQL = "delete from TblMinistryContract_Installment where type = 1 and idmc =  " & val(txtIDMC.Text)
                     Cn.Execute StrSQL, , adExecuteNoRecords
                    
                
                   StrSQL = "SELECT  *  From TblMinistryContract"
                   rs.Close
                   rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
                

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
    'If Err.Number = -2147217887 Then
    Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· · ﬂ«„· «·»Ì«‰«  " & CHR(13) & " ÊÃœ »Ì«‰«  „— »ÿ… »Â–« «·Œ“‰… "
    Msg = Msg & CHR(13) & Err.description
    MsgBox Msg, vbMsgBoxRight + vbMsgBoxRtlReading + vbExclamation, App.title
    rs.CancelUpdate
    'End If
End Sub


Public Function ISAllowDeleteUpdateContract(ID As Integer) As Boolean
Dim str As String

str = " select * from TblVehicleAllocation where IDMC =    " & ID
Set Rs_Temp5 = New ADODB.Recordset
Rs_Temp5.Open str, Cn, adOpenStatic, adLockOptimistic
If Rs_Temp5.RecordCount > 0 Then
        ISAllowDeleteUpdateContract = False
        Exit Function
End If


str = " select * from TblAttributionContract where IDMC   =  " & ID
Set Rs_Temp5 = New ADODB.Recordset
Rs_Temp5.Open str, Cn, adOpenStatic, adLockOptimistic
If Rs_Temp5.RecordCount > 0 Then
        ISAllowDeleteUpdateContract = False
        Exit Function
End If


str = " SELECT  *  from TblRequest_MinistryContract where MinstryID    =  " & ID
Set Rs_Temp5 = New ADODB.Recordset
Rs_Temp5.Open str, Cn, adOpenStatic, adLockOptimistic
If Rs_Temp5.RecordCount > 0 Then
        ISAllowDeleteUpdateContract = False
        Exit Function
End If


ISAllowDeleteUpdateContract = True

End Function


Private Sub AddTip()
    Dim Wrap As String
    On Error GoTo ErrTrap
    Set TTP = New clstooltip
    Wrap = CHR(13) + CHR(10)

    With TTP
        .Create Me.hwnd, "»Ì«‰«  «· ⁄«ﬁœ „⁄ Ê“«—…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(0), "ÃœÌœ ..." & Wrap & "·«÷«›… »Ì«‰«  ⁄ﬁœ Ê“«—… ÃœÌœ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  ⁄ﬁœ Ê“«—…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(1), " ⁄œÌ· ..." & Wrap & "· ⁄œÌ· »Ì«‰«  ⁄ﬁœ Ê“«—…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  ⁄ﬁœ Ê“«—…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(2), "Õ›Ÿ ..." & Wrap & "·Õ›Ÿ »Ì«‰«  ⁄ﬁœ Ê“«—… «·ÃœÌœ" & Wrap & "·Õ›Ÿ «· ⁄œÌ·« " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  ⁄ﬁœ Ê“«—…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(3), " —«Ã⁄ ..." & Wrap & "·· —«Ã⁄ ⁄‰ ⁄„·Ì… «·«÷«›…" & Wrap & "··· —«Ã⁄ ⁄‰ ⁄„·Ì… «· ⁄œÌ·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  ⁄ﬁœ Ê“«—…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(4), "Õ–› ..." & Wrap & "·Õ–› »Ì«‰«  Â–« ⁄ﬁœ Ê“«—… " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  ⁄ﬁœ Ê“«—… ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(5), "»ÕÀ ..." & Wrap & "···»ÕÀ ⁄‰ ⁄ﬁœ Ê“«—… " & Wrap & "Ì‰ÿ»ﬁ ⁄·ÌÂ ‘—Êÿ „⁄Ì‰…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  ⁄ﬁœ Ê“«—… ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl Cmd(6), "Œ—ÊÃ ..." & Wrap & "·«€·«ﬁ Â–Â «·‰«›–…" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  ⁄ﬁœ Ê“«—… ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(1), "«·√Ê· ..." & Wrap & "··«‰ ﬁ«· «·Ï √Ê· ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  ⁄ﬁœ Ê“«—… ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(0), "«·”«»ﬁ ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  ⁄ﬁœ Ê“«—…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(3), "«· «·Ì ..." & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ì" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  ⁄ﬁœ Ê“«—…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl XPBtnMove(2), "«·√ŒÌ— ..." & Wrap & "··«‰ ﬁ«· «·Ï ¬Œ— ”Ã·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

    With TTP
        .Create Me.hwnd, "»Ì«‰«  ⁄ﬁœ Ê“«—… ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
    '    .AddControl CmdHelp, "„”«⁄œ… ..." & Wrap & "·· ⁄—› ⁄·Ï ÊŸÌ›… Â–Â «·‰«›–…" & Wrap & "ÊﬂÌ›Ì… «· ⁄«„· „⁄Â«" & Wrap & "≈÷€ÿ Â‰«" & Wrap, True
    End With

    Exit Sub
ErrTrap:
End Sub

Private Sub XPTxtBoxName_GotFocus()

    SwitchKeyboardLang LANG_ARABIC
End Sub

Private Sub XPTxtBoxNamee_GotFocus()
    SwitchKeyboardLang LANG_ENGLISH
End Sub

Private Sub Calculations(Optional WithMsg As Boolean = True)
'    On Error GoTo ErrTrap
    Dim SngAllValue As Single
    Dim i  As Integer
    Dim DateInterval, Msg As String
    Dim NewDateH As String
    Dim NewDate As String
    Dim PreDateH As String
 
 If IsNumeric(TxtPaymentCount.Text) Then
    If Not (val(TxtPaymentCount.Text) > 0) Then
            MsgBox ("««œŒ· ⁄œœ «·œ›⁄«  «Ê·« ")
            TxtPaymentCount.SetFocus
            Exit Sub
    End If
 Else '
    Exit Sub
 End If
 
 If DcbPeriodsID.ListIndex = -1 Then
 MsgBox (" «œŒ· «·› —… »Ì‰ «·œ›⁄« ")
 Exit Sub
 End If
 
 
   If DcbPeriodsID.ListIndex = 0 Then
        DateInterval = "d"
    ElseIf DcbPeriodsID.ListIndex = 1 Then
        DateInterval = "M"
    ElseIf DcbPeriodsID.ListIndex = 2 Then
        DateInterval = "yyyy"
        Else
        DateInterval = "D"
        
    End If
    
    DtatAdd.value = FirstPaymentDate.value
    With Me.FgInstallments
        .Clear flexClearScrollable, flexClearEverything
        .Rows = .FixedRows + val(TxtPaymentCount.Text)

        For i = 1 To .Rows - 1

            .TextMatrix(i, .ColIndex("QestID")) = i
            .TextMatrix(i, .ColIndex("Value")) = Round(val(txtNet.Text) / val(TxtPaymentCount.Text), 2)
            
          If i = 1 Then
            .TextMatrix(i, .ColIndex("Due_DateH")) = Format(FirstPaymentDateH.value, "yyyy/MM/dd")
             .TextMatrix(i, .ColIndex("Due_Date")) = ToGregorianDate(FirstPaymentDateH.value)
             Else
             PreDateH = (Trim(.TextMatrix(i - 1, .ColIndex("Due_DateH"))))
             NewDateH = (DateAdd(DateInterval, val(TxtPeriods.Text), PreDateH))
             NewDate = ToGregorianDate(NewDateH)
             'DtatAdd.value = DateAdd((DateInterval), val(TxtPeriods.text), DtatAdd.value)
             .TextMatrix(i, .ColIndex("Due_Date")) = NewDate
             .TextMatrix(i, .ColIndex("Due_DateH")) = Format(NewDateH, "yyyy/MM/dd")
             End If
         Next i
         
         .AutoSize 1, .Cols - 1, False
         End With
    Exit Sub
ErrTrap:
End Sub


