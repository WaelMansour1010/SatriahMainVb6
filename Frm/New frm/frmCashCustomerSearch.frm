VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmCashCustomerSearch 
   Appearance      =   0  'Flat
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "‘«‘… »ÕÀ «·⁄„·«¡/«·„Ê—œÌ‰  «·‰ﬁœÌ"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9315
   Icon            =   "frmCashCustomerSearch.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5895
   ScaleWidth      =   9315
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.Frame Frame4 
      BackColor       =   &H00E2E9E9&
      Height          =   5175
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   55
      Top             =   -120
      Width           =   9255
      Begin VB.TextBox TxtPlatNo 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   5280
         TabIndex        =   65
         Top             =   4710
         Width           =   2985
      End
      Begin VB.TextBox TxtItemName 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   360
         RightToLeft     =   -1  'True
         TabIndex        =   61
         Top             =   3900
         Width           =   7920
      End
      Begin VB.TextBox TxtItemID 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   360
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   60
         Top             =   4350
         Width           =   7920
      End
      Begin VB.CheckBox Check17 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " ÕœÌœ «·ﬂ·"
         Height          =   255
         Left            =   7680
         RightToLeft     =   -1  'True
         TabIndex        =   56
         Top             =   120
         Width           =   1425
      End
      Begin VSFlex8UCtl.VSFlexGrid Fg 
         Height          =   3375
         Left            =   120
         TabIndex        =   59
         Top             =   480
         Width           =   9075
         _cx             =   16007
         _cy             =   5953
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
         Cols            =   17
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmCashCustomerSearch.frx":030A
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
      Begin VB.Label LblPla 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "—ﬁ„ «··ÊÕ…"
         Height          =   210
         Left            =   7890
         TabIndex        =   66
         Top             =   4740
         Width           =   1215
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·«”„ ⁄—»Ì"
         Height          =   345
         Index           =   17
         Left            =   8190
         RightToLeft     =   -1  'True
         TabIndex        =   63
         Top             =   3900
         Width           =   915
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«· ·Ì›Ê‰"
         Height          =   345
         Index           =   16
         Left            =   8190
         RightToLeft     =   -1  'True
         TabIndex        =   62
         Top             =   4350
         Width           =   915
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E2E9E9&
      Height          =   5175
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   31
      Top             =   0
      Width           =   9255
      Begin VB.ComboBox CBOCartTYpe2 
         Height          =   315
         ItemData        =   "frmCashCustomerSearch.frx":05B4
         Left            =   5640
         List            =   "frmCashCustomerSearch.frx":05BE
         RightToLeft     =   -1  'True
         TabIndex        =   64
         Top             =   4920
         Visible         =   0   'False
         Width           =   3480
      End
      Begin VB.TextBox Txtdiscount 
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
         Height          =   360
         Left            =   120
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   53
         Top             =   4680
         Width           =   3480
      End
      Begin VB.Frame Frame14 
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·ﬂÊœ"
         Height          =   645
         Left            =   5040
         RightToLeft     =   -1  'True
         TabIndex        =   48
         Top             =   2880
         Width           =   4155
         Begin VB.TextBox TxtFromID 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   1920
            RightToLeft     =   -1  'True
            TabIndex        =   50
            Top             =   240
            Width           =   915
         End
         Begin VB.TextBox TxtToID 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   49
            Top             =   240
            Width           =   915
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„‰"
            Height          =   195
            Index           =   67
            Left            =   2775
            RightToLeft     =   -1  'True
            TabIndex        =   52
            Top             =   240
            Width           =   540
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈·Ï"
            Height          =   195
            Index           =   66
            Left            =   1260
            RightToLeft     =   -1  'True
            TabIndex        =   51
            Top             =   240
            Width           =   525
         End
      End
      Begin VB.Frame Frame13 
         BackColor       =   &H00E2E9E9&
         Caption         =   " «—ÌŒ «·’·«ÕÌ…"
         Height          =   915
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   43
         Top             =   2880
         Width           =   4455
         Begin MSComCtl2.DTPicker ToRecordDate 
            Height          =   330
            Left            =   120
            TabIndex        =   44
            Top             =   360
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   235732995
            CurrentDate     =   38887
         End
         Begin MSComCtl2.DTPicker FrmRecordDate 
            Height          =   330
            Left            =   2280
            TabIndex        =   45
            TabStop         =   0   'False
            Top             =   360
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   582
            _Version        =   393216
            CalendarBackColor=   12648447
            CalendarTitleBackColor=   10383715
            CheckBox        =   -1  'True
            Format          =   235732995
            CurrentDate     =   37140
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„‰"
            Height          =   195
            Index           =   64
            Left            =   3720
            RightToLeft     =   -1  'True
            TabIndex        =   47
            Top             =   330
            Width           =   540
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈·Ï"
            Height          =   195
            Index           =   65
            Left            =   1575
            RightToLeft     =   -1  'True
            TabIndex        =   46
            Top             =   420
            Width           =   480
         End
      End
      Begin VB.TextBox TxtVacNamee 
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
         Height          =   360
         Left            =   120
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   41
         Top             =   3840
         Width           =   3480
      End
      Begin VB.ComboBox CBOCartTYpe 
         Height          =   315
         ItemData        =   "frmCashCustomerSearch.frx":05CD
         Left            =   4800
         List            =   "frmCashCustomerSearch.frx":05D7
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   4680
         Width           =   3480
      End
      Begin VB.TextBox TxtVacName 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   4800
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   3840
         Width           =   3480
      End
      Begin VB.TextBox TxtTelepone 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   4800
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   4290
         Width           =   3480
      End
      Begin VB.TextBox Txtcard 
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
         Height          =   360
         Left            =   120
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   4260
         Width           =   3480
      End
      Begin VSFlex8UCtl.VSFlexGrid Fg1 
         Height          =   2745
         Left            =   120
         TabIndex        =   32
         Top             =   120
         Width           =   9075
         _cx             =   16007
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
         Cols            =   18
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmCashCustomerSearch.frx":05E6
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
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ﬁÌ„… «·Œ’„"
         Height          =   285
         Index           =   4
         Left            =   2715
         RightToLeft     =   -1  'True
         TabIndex        =   54
         Top             =   4755
         Width           =   1890
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·«”„ «‰Ã·Ì“Ì"
         Height          =   285
         Index           =   1
         Left            =   2715
         RightToLeft     =   -1  'True
         TabIndex        =   42
         Top             =   3840
         Width           =   1890
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰Ê⁄ «·ﬂ«— "
         Height          =   285
         Index           =   7
         Left            =   8220
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   4680
         Width           =   930
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·«”„ ⁄—»Ì"
         Height          =   345
         Index           =   1
         Left            =   8190
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   3840
         Width           =   915
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«· ·Ì›Ê‰"
         Height          =   345
         Index           =   4
         Left            =   8190
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   4290
         Width           =   915
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·ﬂ— "
         Height          =   285
         Index           =   6
         Left            =   3540
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   4290
         Width           =   930
      End
   End
   Begin VB.TextBox TxtCopun 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   2520
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   29
      Top             =   6240
      Width           =   3000
   End
   Begin VB.TextBox txtMessage 
      Alignment       =   1  'Right Justify
      Height          =   495
      Left            =   1200
      RightToLeft     =   -1  'True
      TabIndex        =   27
      Top             =   5880
      Width           =   4335
   End
   Begin VB.ComboBox CboItemCodeSearch 
      Height          =   315
      Left            =   3630
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   5790
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·’‰› «·„—«œ «·»ÕÀ ⁄‰Â ÌÕ ÊÏ ⁄·Ï Â–« «·’‰› ﬂ«Õœ „·Õﬁ« Â"
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
      Height          =   885
      Left            =   30
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   7470
      Width           =   6495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·’‰› «·„—«œ «·»ÕÀ ⁄‰Â ÌÕ ÊÏ ⁄·Ï Â–« «·’‰› ﬂ«Õœ „ﬂÊ‰« Â"
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
      Height          =   885
      Left            =   30
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   7740
      Width           =   6495
   End
   Begin VB.ComboBox CboArchive 
      Height          =   315
      Left            =   30
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   6300
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ComboBox CboGuar 
      Height          =   315
      Left            =   2160
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   7410
      Width           =   1305
   End
   Begin VB.ComboBox CboNameSearch 
      Height          =   315
      Left            =   3630
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   5940
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.ComboBox CboAttachedItem 
      Height          =   315
      Left            =   4470
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   7740
      Width           =   1095
   End
   Begin VB.ComboBox CboAssbliedItem 
      Height          =   315
      Left            =   2160
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   8940
      Width           =   1305
   End
   Begin VB.ComboBox CboItemType 
      Height          =   315
      Left            =   4380
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   7530
      Width           =   1215
   End
   Begin VB.ComboBox CboSerial 
      Height          =   315
      ItemData        =   "frmCashCustomerSearch.frx":08A9
      Left            =   30
      List            =   "frmCashCustomerSearch.frx":08AB
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   8820
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.TextBox XPTxtItemCode 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   2610
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   6615
      Width           =   1395
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   0
      Left            =   3390
      TabIndex        =   10
      Top             =   5235
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
      Left            =   2370
      TabIndex        =   11
      Top             =   5235
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
      Index           =   2
      Left            =   1470
      TabIndex        =   12
      Top             =   5235
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
   Begin MSDataListLib.DataCombo DCboGroupName 
      Height          =   315
      Left            =   2610
      TabIndex        =   3
      Top             =   8490
      Width           =   2985
      _ExtentX        =   5265
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   3
      Left            =   0
      TabIndex        =   26
      Top             =   5760
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   661
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "«—”«·"
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
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·«”„"
      Height          =   345
      Index           =   15
      Left            =   6960
      RightToLeft     =   -1  'True
      TabIndex        =   58
      Top             =   3660
      Width           =   915
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«· ·Ì›Ê‰"
      Height          =   345
      Index           =   14
      Left            =   6960
      RightToLeft     =   -1  'True
      TabIndex        =   57
      Top             =   4110
      Width           =   915
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "ﬂÊœ «·⁄—÷"
      Height          =   345
      Index           =   13
      Left            =   5640
      RightToLeft     =   -1  'True
      TabIndex        =   30
      Top             =   6360
      Width           =   915
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·—”«·Â"
      Height          =   345
      Index           =   12
      Left            =   5640
      RightToLeft     =   -1  'True
      TabIndex        =   28
      Top             =   5760
      Width           =   915
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "„Ã«· «·»ÕÀ"
      Height          =   345
      Index           =   11
      Left            =   5190
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   5790
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·√—‘Ì›"
      Height          =   285
      Index           =   10
      Left            =   1410
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   8970
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·÷„«‰"
      Height          =   285
      Index           =   9
      Left            =   3480
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   7410
      Width           =   885
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "„Ã«· «·»ÕÀ"
      Height          =   345
      Index           =   8
      Left            =   5190
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   5940
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«Ìﬁ«› «· ⁄«„·"
      Height          =   315
      Index           =   7
      Left            =   5490
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   7140
      Width           =   1065
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " Ã„Ì⁄"
      Height          =   285
      Index           =   6
      Left            =   3480
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   8940
      Width           =   915
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "‰Ê⁄ «·’‰›"
      Height          =   285
      Index           =   5
      Left            =   5640
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   9000
      Width           =   915
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«Ìﬁ«› «· ⁄«„·"
      Height          =   315
      Index           =   2
      Left            =   1470
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   7530
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label LblRes 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      ForeColor       =   &H000000C0&
      Height          =   315
      Left            =   5040
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   5280
      Width           =   1905
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "ﬂÊœ «·„—ﬂ“"
      Height          =   345
      Index           =   0
      Left            =   4020
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   7470
      Width           =   795
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·›∆…"
      Height          =   285
      Index           =   3
      Left            =   5640
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   8730
      Width           =   915
   End
End
Attribute VB_Name = "frmCashCustomerSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim cSearchDcbo As clsDCboSearch

Private m_DcboItems As DataCombo

Private m_RetrunType As Integer

Public Function doit()
Cmd_Click (0)
Me.Height = 6210
End Function

Private Sub Check17_Click()
  Dim i As Integer

    If Check17.value = vbChecked Then

        With Me.FG
 
            For i = 1 To .rows - 2
        
                .TextMatrix(i, .ColIndex("Send")) = True
            Next i

        End With

    Else

        With Me.FG

            For i = 1 To .rows - 2
        
                .TextMatrix(i, .ColIndex("Send")) = False
            Next i

        End With

    End If
End Sub

Private Sub Cmd_Click(index As Integer)
    On Error GoTo ErrTrap

    Select Case index

        Case 0
If RetrunType = 4 Then
GetDate
Else
            If rs.State = adStateOpen Then
                rs.Close
            End If

            rs.Open Build_Sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

            If SystemOptions.UserInterface = ArabicInterface Then
                LblRes.Caption = "‰ ÌÃ… «·»ÕÀ = " & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
                LblRes.Caption = "Search Result=" & rs.RecordCount
            End If
    
            If rs.RecordCount < 1 Then
                FG.Clear flexClearScrollable, flexClearEverything
                FG.rows = 2

                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Else
                    Msg = "NO Search Results Found...!!!"
                    MsgBox Msg, vbOKOnly + vbExclamation, App.Title
                End If

                Exit Sub
            End If

            Retrive
            FG.SetFocus
End If
        Case 1
            clear_all Me
            FG.Clear flexClearScrollable, flexClearEverything
            FG1.Clear flexClearScrollable, flexClearEverything
FrmRecordDate.value = ""
ToRecordDate.value = ""
        Case 2
            Unload Me
    End Select

    Exit Sub
ErrTrap:

    If Err.Number = -2147217900 Then
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & CHR(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… „⁄«ÌÌ— «·»ÕÀ Ê√⁄œ «·„Õ«Ê·…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

End Sub

Private Sub fg_Click()
    On Error GoTo ErrTrap

    If Not FG.TextMatrix(FG.row, 4) = "" Then
        If Me.RetrunType = 0 Then
            frmsalebill.TxtCashCustomerName.text = (FG.TextMatrix(FG.row, 4))
            frmsalebill.TxtPhone.text = (FG.TextMatrix(FG.row, 5))
            
        ElseIf Me.RetrunType = 1 Then
          FrmBillBuy.TxtCashCustomerName.text = (FG.TextMatrix(FG.row, 4))
            FrmBillBuy.TxtPhone.text = (FG.TextMatrix(FG.row, 5))
            
   
        ElseIf Me.RetrunType = 2 Then
         frmsalebill2.CashCustomerName.text = (FG.TextMatrix(FG.row, 4))
            frmsalebill2.TxtPhone(0).text = (FG.TextMatrix(FG.row, 5))
            
       ElseIf Me.RetrunType = 205 Then
         frmsalebill5.TxtCashCustomerName.text = (FG.TextMatrix(FG.row, 4))
            frmsalebill5.TxtPhone.text = (FG.TextMatrix(FG.row, 5))
         ElseIf Me.RetrunType = 206 Then
         frmsalebill6.TxtCashCustomerName.text = (FG.TextMatrix(FG.row, 4))
            frmsalebill6.TxtPhone.text = (FG.TextMatrix(FG.row, 5))
            frmsalebill6.TxtPlatNo.text = (FG.TextMatrix(FG.row, FG.ColIndex("PlateNo")))
             
             
  ElseIf Me.RetrunType = 3 Then
         FrmReturnSalling.TxtCashCustomerName.text = (FG.TextMatrix(FG.row, 4))
            FrmReturnSalling.TxtPhone.text = (FG.TextMatrix(FG.row, 5))
            
  
           
 
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Retrive()
    Dim Num As Integer
    On Error GoTo ErrTrap
    FG.Clear flexClearScrollable, flexClearEverything

    If Not (rs.EOF Or rs.BOF) Then
        FG.rows = rs.RecordCount + 1

        For Num = 1 To rs.RecordCount

            With FG
                .TextMatrix(Num, .ColIndex("NumIndex")) = Num
                .TextMatrix(Num, .ColIndex("Parent")) = IIf(IsNull(rs("CashCustomerPhone").value), "", (rs("CashCustomerPhone").value))
 
                .TextMatrix(Num, .ColIndex("KindNme")) = IIf(IsNull(rs("CashCustomerName").value), "", Trim(rs("CashCustomerName").value))
                .TextMatrix(Num, .ColIndex("PlateNo")) = IIf(IsNull(rs("PlateNo").value), "", Trim(rs("PlateNo").value))
 
    
            End With

            rs.MoveNext
        Next Num

        FG.AutoSize 0, FG.Cols - 1, False
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Fg_DblClick()
    fg_Click
    Unload Me
End Sub

Private Sub Fg1_Click()
 If Me.RetrunType = 4 Then
         FrmCustCash.FindRec val(FG1.TextMatrix(FG1.row, FG1.ColIndex("id")))
   End If
End Sub

Private Sub Form_Load()
  On Error GoTo ErrTrap
    Dim StrSQL As String
    Dim BG As New ClsBackGroundPic
    Dim Dcombos As ClsDataCombos
    Frame4.Visible = False
    Frame3.Visible = False
If RetrunType = 4 Then
Frame3.Visible = True
FrmRecordDate.value = ""
ToRecordDate.value = ""
Else
Frame4.Visible = True
End If
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Clear").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture

    Set Dcombos = New ClsDataCombos
    Dcombos.GetExpensesGroups Me.DCboGroupName

    Set cSearchDcbo = New clsDCboSearch
    'cSearchDcbo.AllowWriting = False
    Set cSearchDcbo.Client = Me.DCboGroupName

    If SystemOptions.UserInterface = ArabicInterface Then

        With Me.CboItemCodeSearch
            .Clear
            .AddItem "»ÕÀ „ÿ«»ﬁ"
            .AddItem "»ÕÀ „‰ «·»œ«Ì…"
            .AddItem "»ÕÀ „‰ «·‰Â«Ì…"
            .AddItem "»ÕÀ ›Ï «Ï „ﬂ«‰"
        End With

        With Me.CboSerial
            .Clear
            .AddItem "«·ﬂ·"
            .ItemData(0) = 0
            .AddItem "·Â ”Ì—Ì«·"
            .ItemData(1) = 1
            .AddItem "·Ì” ·Â ”Ì—Ì«·"
            .ItemData(2) = 2
        End With

        With Me.CboNameSearch
            .Clear
            .AddItem "„‰ «Ê· «·√”„"
            .AddItem "›Ï «Ï Ã“¡ „‰ «·√”„"
        End With

        With Me.CboItemType
            .Clear
            .AddItem "”·⁄…"
            .AddItem "Œœ„…"
            .AddItem "«·ﬂ·"
        End With

        With Me.CboGuar
            .Clear
            .AddItem "·Â ÷„«‰"
            .AddItem "·Ì” ·Â ÷„«‰"
            .AddItem "«·ﬂ·"
        End With

        With Me.CboArchive
            .Clear
            .AddItem "›Ï «·√—‘Ì›"
            .AddItem "·Ì” ›Ï «·√—‘Ì›"
            .AddItem "«·ﬂ·"
        End With

        With Me.CboAssbliedItem
            .Clear
            .AddItem "’‰› „Ã„⁄"
            .AddItem "’‰› ⁄«œÏ"
            .AddItem "«·ﬂ·"
        End With

        With Me.CboAttachedItem
            .Clear
            .AddItem "‰⁄„"
            .AddItem "·«"
         
        End With

    ElseIf SystemOptions.UserInterface = EnglishInterface Then

        With Me.CboItemCodeSearch
            .Clear
            .AddItem "Typical Search"
            .AddItem "From The Start"
            .AddItem "From The End"
            .AddItem "Any Where"
        End With

        With Me.CboSerial
            .Clear
            .AddItem "All"
            .ItemData(0) = 0
            .AddItem "Has Serial"
            .ItemData(1) = 1
            .AddItem "NO Serial"
            .ItemData(2) = 2
        End With

        With Me.CboNameSearch
            .Clear
            .AddItem "Start Name"
            .AddItem "Any Part of Name"
        End With

        With Me.CboItemType
            .Clear
            .AddItem "Goods"
            .AddItem "Services"
            .AddItem "All"
        End With

        With Me.CboGuar
            .Clear
            .AddItem "YES"
            .AddItem "NO"
            .AddItem "ALL"
        End With

        With Me.CboArchive
            .Clear
            .AddItem "YES"
            .AddItem "NO"
            .AddItem "ALL"
        End With

        With Me.CboAssbliedItem
            .Clear
            .AddItem "YES"
            .AddItem "NO"
            .AddItem "ALL"
        End With

        With Me.CboAttachedItem
            .Clear
            .AddItem "YES"
            .AddItem "NO"
 
        End With

    End If

    CenterForm Me

    FormPostion Me, GetPostion
    FG.WallPaper = BG.SearchWallpaper
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

    StrSQL = "SELECT DISTINCT CashCustomerName,CashCustomerPhone,PlateNo  from dbo.Transactions"
 ' StrSQL = StrSQL + " Where 1=1 "
 If RetrunType = 0 Or RetrunType = 2 Or RetrunType = 3 Or RetrunType = 206 Then
  StrSQL = StrSQL + " Where Transaction_Type=21 "
 Else
 StrSQL = StrSQL + " Where Transaction_Type=22 "
 End If
 
    If (Me.TxtItemID.text) <> "" Then
        StrSQL = StrSQL + " AND CashCustomerPhone like'%" & val(Me.TxtItemID.text) & "%'"
    End If
 
    If Trim(Me.txtItemName.text) <> "" Then
            StrSQL = StrSQL + " and CashCustomerName Like '%" & Trim(Me.txtItemName.text) & "%'"
    End If

    If TxtPlatNo <> "" Then
            
        StrSQL = StrSQL + " and PlateNo Like '%" & Trim(Trim(Me.TxtPlatNo.text)) & "'   "
           
    End If
 
    Build_Sql = StrSQL
    Exit Function
ErrTrap:
End Function
Sub GetDate()
Dim i As Integer
    Dim StrSQL As String
    Dim StrWhere As String
    Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    On Error GoTo ErrTrap

    StrSQL = "SELECT  * from dbo.TblCusCsh"
   StrSQL = StrSQL + " Where 1=1 "
 If val(TxtFromID.text) <> 0 Then
  StrSQL = StrSQL + " AND Id >= " & val(Me.TxtFromID.text) & ""
 End If
  If val(TxtToID.text) <> 0 Then
  StrSQL = StrSQL + " AND Id <= " & val(Me.TxtToID.text) & ""
 End If
    If (Me.TxtTelepone.text) <> "" Then
        StrSQL = StrSQL + " AND tel like'%" & (Me.TxtTelepone.text) & "%'"
    End If
 
    If Trim(Me.TxtVacName.text) <> "" Then
            StrSQL = StrSQL + " and name Like '%" & Trim(Me.TxtVacName.text) & "%'"
    End If
    If Trim(Me.TxtVacNamee.text) <> "" Then
            StrSQL = StrSQL + " and namee Like '%" & Trim(Me.TxtVacNamee.text) & "%'"
    End If
     If Trim(Me.Txtcard.text) <> "" Then
            StrSQL = StrSQL + " and card Like '%" & Trim(Me.Txtcard.text) & "%'"
    End If
 If val(txtDiscount.text) <> 0 Then
  StrSQL = StrSQL + " AND discount = " & val(Me.txtDiscount.text) & ""
 End If
 If CBOCartTYpe.text <> "" And val(CBOCartTYpe.ListIndex) <> -1 Then
 StrSQL = StrSQL + " AND CartTYpe = '" & (Me.CBOCartTYpe) & "'"
 End If
 If Not IsNull(Me.FrmRecordDate.value) Then
 StrSQL = StrSQL + " AND RecordDate >= " & SQLDate(FrmRecordDate.value, True) & ""
 End If
  If Not IsNull(Me.ToRecordDate.value) Then
 StrSQL = StrSQL + " AND RecordDate <= " & SQLDate(ToRecordDate.value, True) & ""
 End If
  Rs3.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    FG1.Clear flexClearScrollable, flexClearEverything

 If Rs3.RecordCount > 0 Then
        FG1.rows = Rs3.RecordCount + 1
Rs3.MoveFirst
        For i = 1 To Rs3.RecordCount

            With FG1
                .TextMatrix(i, .ColIndex("NumIndex")) = i
                .TextMatrix(i, .ColIndex("Id")) = IIf(IsNull(Rs3("Id").value), 0, (Rs3("Id").value))
                .TextMatrix(i, .ColIndex("tel")) = IIf(IsNull(Rs3("tel").value), "", Trim(Rs3("tel").value))
                .TextMatrix(i, .ColIndex("discount")) = IIf(IsNull(Rs3("discount").value), "", Trim(Rs3("discount").value))
                .TextMatrix(i, .ColIndex("RecordDate")) = IIf(IsNull(Rs3("RecordDate").value), "", Trim(Rs3("RecordDate").value))
                .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(Rs3("name").value), "", Trim(Rs3("name").value))
                .TextMatrix(i, .ColIndex("namee")) = IIf(IsNull(Rs3("namee").value), "", Trim(Rs3("namee").value))
                .TextMatrix(i, .ColIndex("card")) = IIf(IsNull(Rs3("card").value), "", (Rs3("card").value))
                 CBOCartTYpe2.text = IIf(IsNull(Rs3("CartTYpe").value), "", (Rs3("CartTYpe").value))
                .TextMatrix(i, .ColIndex("CartTYpe")) = CBOCartTYpe2.text
            End With

            Rs3.MoveNext
        Next i

        FG1.AutoSize 0, FG.Cols - 1, False
    End If
    Exit Sub
ErrTrap:
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        If Me.ActiveControl Is FG Then
            If Not FG.TextMatrix(FG.row, 1) = "" Then
                fg_Click
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

Public Property Get DcboItems() As DataCombo
    Set DcboItems = m_DcboItems
End Property

Public Property Set DcboItems(ByVal vNewValue As DataCombo)
    Set m_DcboItems = vNewValue
End Property

Public Property Get RetrunType() As Integer
    RetrunType = m_RetrunType
End Property

Public Property Let RetrunType(ByVal vNewValue As Integer)
    m_RetrunType = vNewValue
    ' 0 = Retrun in the Items Screen
    ' 1 = Retrun in the Data Combo
End Property

Private Sub ChangeLang()
    Me.Caption = "Search For Cash Customers"
Check17.Caption = "Select All"

    lbl(1).Caption = " Name"
 
     lbl(4).Caption = " Mobile"

lbl(12).Caption = " Message"
lbl(13).Caption = " Over Code"

    Cmd(0).Caption = "Search"
    Cmd(1).Caption = "Clear"
    Cmd(2).Caption = "Exit"
    Cmd(3).Caption = "Send"
    With Me.FG
  
         .TextMatrix(0, .ColIndex("Send")) = " Select"
        .TextMatrix(0, .ColIndex("Parent")) = " Mobile"
        
        .TextMatrix(0, .ColIndex("KindNme")) = " Name"
    
    
        .AutoSize 0, .Cols - 1, False
    End With

End Sub

Private Sub TxtItemName_Change()

    If Trim$(Me.txtItemName.text) = "" Then
        Me.lbl(8).Enabled = False
        Me.CboNameSearch.Enabled = False
    Else
        Me.lbl(8).Enabled = True
        Me.CboNameSearch.Enabled = True
    End If

End Sub

Private Sub XPTxtItemCode_Change()

    If Trim$(Me.XPTxtItemCode.text) = "" Then
        Me.lbl(11).Enabled = False
        Me.CboItemCodeSearch.Enabled = False
    Else
        Me.lbl(11).Enabled = True
        Me.CboItemCodeSearch.Enabled = True
    End If

End Sub

