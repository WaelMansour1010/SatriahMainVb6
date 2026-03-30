VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTestWael 
   Caption         =   "Form1"
   ClientHeight    =   8280
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   13260
   LinkTopic       =   "Form1"
   ScaleHeight     =   8280
   ScaleWidth      =   13260
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   1185
      Left            =   840
      TabIndex        =   2
      Top             =   2460
      Width           =   2505
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   405
      Left            =   510
      TabIndex        =   1
      Top             =   1320
      Width           =   1575
   End
   Begin C1SizerLibCtl.C1Elastic C1Elastic12 
      Height          =   5790
      Left            =   3030
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1800
      Width           =   9525
      _cx             =   16801
      _cy             =   10213
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
      Begin VB.Frame Fra_Header 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   675
         Index           =   7
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   0
         Width           =   7515
         Begin VB.TextBox TxtModFlg2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H0000FF00&
            Enabled         =   0   'False
            Height          =   285
            Index           =   2
            Left            =   2580
            RightToLeft     =   -1  'True
            TabIndex        =   17
            Text            =   "modflag"
            Top             =   90
            Visible         =   0   'False
            Width           =   465
         End
         Begin MSComctlLib.ImageList GrdImageList2 
            Index           =   8
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
                  Picture         =   "frmTestWael.frx":0000
                  Key             =   "CompanyName"
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTestWael.frx":039A
                  Key             =   "Ser"
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTestWael.frx":0734
                  Key             =   "Vac_Name"
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTestWael.frx":0ACE
                  Key             =   "ShareCount"
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTestWael.frx":0E68
                  Key             =   "Dis_Count"
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTestWael.frx":1202
                  Key             =   "Bouns"
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTestWael.frx":159C
                  Key             =   "SharesValue"
               EndProperty
               BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmTestWael.frx":1B36
                  Key             =   "BuyValue"
               EndProperty
            EndProperty
         End
         Begin ImpulseButton.ISButton btn_Last 
            Height          =   315
            Index           =   2
            Left            =   90
            TabIndex        =   18
            Top             =   30
            Width           =   405
            _ExtentX        =   714
            _ExtentY        =   556
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   ""
            BackColor       =   14871017
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
            ButtonImage     =   "frmTestWael.frx":1ED0
            ColorButton     =   14871017
            AcclimateGrayTones=   -1  'True
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btn_Next 
            Height          =   315
            Index           =   2
            Left            =   555
            TabIndex        =   19
            Top             =   30
            Width           =   405
            _ExtentX        =   714
            _ExtentY        =   556
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   ""
            BackColor       =   14871017
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
            ButtonImage     =   "frmTestWael.frx":226A
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btn_Previous 
            Height          =   315
            Index           =   2
            Left            =   1155
            TabIndex        =   20
            Top             =   30
            Width           =   405
            _ExtentX        =   714
            _ExtentY        =   556
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   ""
            BackColor       =   14871017
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
            ButtonImage     =   "frmTestWael.frx":2604
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btn_First 
            Height          =   315
            Index           =   2
            Left            =   1620
            TabIndex        =   21
            Top             =   30
            Width           =   405
            _ExtentX        =   714
            _ExtentY        =   556
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   ""
            BackColor       =   14871017
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
            ButtonImage     =   "frmTestWael.frx":299E
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "›∆… «·”œ«œ"
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
            Index           =   22
            Left            =   4650
            RightToLeft     =   -1  'True
            TabIndex        =   22
            Top             =   60
            Width           =   2640
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E2E9E9&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   1800
         Index           =   2
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   4215
         Width           =   6300
         Begin VB.ComboBox Combo4 
            BackColor       =   &H80000018&
            Height          =   315
            ItemData        =   "frmTestWael.frx":2D38
            Left            =   2280
            List            =   "frmTestWael.frx":2D48
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   3150
            Visible         =   0   'False
            Width           =   1005
         End
         Begin VB.TextBox TxtSerial1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   2
            Left            =   3030
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   11
            Top             =   330
            Width           =   1065
         End
         Begin VB.TextBox TxtName 
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
            Height          =   315
            Index           =   3
            Left            =   1395
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   10
            Top             =   705
            Width           =   2760
         End
         Begin VB.TextBox txtNamee 
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
            Height          =   315
            Index           =   2
            Left            =   1395
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   9
            Top             =   1020
            Width           =   2760
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·ﬂÊœ "
            Height          =   195
            Index           =   3
            Left            =   4695
            RightToLeft     =   -1  'True
            TabIndex        =   15
            Top             =   450
            Width           =   990
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ ⁄—»Ì"
            Height          =   285
            Index           =   8
            Left            =   4350
            RightToLeft     =   -1  'True
            TabIndex        =   14
            Top             =   780
            Width           =   1350
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «‰Ã·Ì“Ì"
            Height          =   285
            Index           =   4
            Left            =   4200
            RightToLeft     =   -1  'True
            TabIndex        =   13
            Top             =   1140
            Width           =   1500
         End
      End
      Begin ImpulseButton.ISButton btn_New 
         Height          =   420
         Index           =   2
         Left            =   6660
         TabIndex        =   23
         Top             =   6855
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   741
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
         ButtonImage     =   "frmTestWael.frx":2D61
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton btn_Save 
         Height          =   420
         Index           =   2
         Left            =   4845
         TabIndex        =   24
         Top             =   6855
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   741
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
         ButtonImage     =   "frmTestWael.frx":30FB
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton btn_Modify 
         Height          =   420
         Index           =   2
         Left            =   5700
         TabIndex        =   25
         Top             =   6855
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   741
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
         ButtonImage     =   "frmTestWael.frx":3495
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton Btn_Undo 
         Height          =   420
         Index           =   2
         Left            =   3870
         TabIndex        =   26
         Top             =   6855
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   741
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
         ButtonImage     =   "frmTestWael.frx":382F
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton btn_Delete 
         Height          =   420
         Index           =   2
         Left            =   3030
         TabIndex        =   27
         Top             =   6855
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   741
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
         ButtonImage     =   "frmTestWael.frx":3BC9
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton Btn_Update 
         Height          =   240
         Index           =   2
         Left            =   5325
         TabIndex        =   28
         TabStop         =   0   'False
         ToolTipText     =   " ÕœÌÀ ﬁ«⁄œ… «·»Ì«‰« "
         Top             =   6015
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   423
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   " ÕœÌÀ"
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
         ButtonImage     =   "frmTestWael.frx":4163
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton btn_Cancel 
         Height          =   420
         Index           =   2
         Left            =   0
         TabIndex        =   29
         Top             =   6855
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   741
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
         ButtonImage     =   "frmTestWael.frx":44FD
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton Btn_Print 
         Height          =   510
         Index           =   2
         Left            =   2055
         TabIndex        =   30
         TabStop         =   0   'False
         ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
         Top             =   6765
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   900
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
         ButtonImage     =   "frmTestWael.frx":4897
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btn_Query 
         Height          =   570
         Index           =   2
         Left            =   720
         TabIndex        =   31
         TabStop         =   0   'False
         ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
         Top             =   6705
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   1005
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
         ButtonImage     =   "frmTestWael.frx":B0F9
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin VSFlex8Ctl.VSFlexGrid Grid3 
         Height          =   3495
         Left            =   0
         TabIndex        =   32
         Top             =   750
         Width           =   7635
         _cx             =   13467
         _cy             =   6165
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
         Rows            =   50
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   320
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmTestWael.frx":B493
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
      Begin VB.Label LabCount_Rec 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Height          =   255
         Index           =   2
         Left            =   2790
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   6390
         Width           =   480
      End
      Begin VB.Label LabCurr_Rec 
         BackColor       =   &H00E2E9E9&
         Height          =   255
         Index           =   2
         Left            =   4485
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   6390
         Width           =   840
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "⁄œœ «·”Ã·« :"
         Height          =   240
         Index           =   16
         Left            =   3510
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   6375
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·”Ã· «·Õ«·Ì:"
         Height          =   240
         Index           =   17
         Left            =   5325
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   6375
         Width           =   1215
      End
   End
   Begin C1SizerLibCtl.C1Tab TabMain 
      Height          =   7755
      Left            =   2730
      TabIndex        =   3
      Top             =   0
      Width           =   4530
      _cx             =   7990
      _cy             =   13679
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
      Caption         =   "⁄—÷ ‘Ã—Ï|⁄—÷ ÃœÊ·Ï"
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
      Begin C1SizerLibCtl.C1Elastic ELe 
         Height          =   7380
         Index           =   1
         Left            =   45
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   45
         Width           =   4440
         _cx             =   7832
         _cy             =   13018
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
            Height          =   7275
            Index           =   0
            Left            =   6180
            TabIndex        =   5
            Top             =   645
            Width           =   4380
            _cx             =   7726
            _cy             =   12832
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
            FormatString    =   $"frmTestWael.frx":B522
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
      End
      Begin C1SizerLibCtl.C1Elastic ELe 
         Height          =   7380
         Index           =   0
         Left            =   5175
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   45
         Width           =   4440
         _cx             =   7832
         _cy             =   13018
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
            Height          =   7275
            Index           =   1
            Left            =   6180
            TabIndex        =   7
            Top             =   645
            Width           =   4380
            _cx             =   7726
            _cy             =   12832
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
            FormatString    =   $"frmTestWael.frx":B5E2
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
      End
   End
End
Attribute VB_Name = "frmTestWael"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
 Dim MySQL As String
MySQL = " SELECT     dbo.TblOrderUpload.ID,dbo.TblOrderUpload.CountOrders, dbo.TblOrderUpload.RecordDate NoteDate, dbo.TblOrderUpload.BranchID, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, "
MySQL = MySQL & "                       CONVERT(char(10), dbo.TblOrderUpload.TimeOrder, 108) as TimeOrder,"
MySQL = MySQL & "                      dbo.TblOrderUpload.DrievType, dbo.TblOrderUpload.EmpID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee,"
MySQL = MySQL & "                      dbo.TblOrderUpload.IDNo, dbo.TblOrderUpload.LeaderName, dbo.TblOrderUpload.Nationality, dbo.TblOrderUpload.CarType, dbo.TblOrderUpload.CusID,"
MySQL = MySQL & "                      dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode AS CusFullcode, dbo.TblOrderUpload.TypGoods,"
MySQL = MySQL & "                      dbo.TblOrderUpload.OrderNo, dbo.TblOrderUpload.Remarks, dbo.TblOrderUpload.PartPrice, dbo.TblOrderUpload.Price, dbo.TblOrderUpload.Total,"
MySQL = MySQL & "                      dbo.TblOrderUpload.CityID, TblCountriesGovernments_2.GovernmentName AS FromCity, dbo.TblOrderUpload.CityID2,"
MySQL = MySQL & "                      TblCountriesGovernments_1.GovernmentName AS ToCity, dbo.TblOrderUpload.CarID, dbo.TblCarsData.BoardNO, dbo.TblOrderUpload.CarID2,"
MySQL = MySQL & "                      TblVendorCars_2.BoardNo AS BoardNo2, dbo.TblOrderUpload.SupplemID, dbo.FixedAssets.Name AS SupplemName, dbo.FixedAssets.namee AS SupplemNameE,"
MySQL = MySQL & "                      dbo.TblOrderUpload.SupplemID2, TblVendorCars_1.accessory, dbo.TblOrderUpload.CustId1, TblCustemers_1.CusName AS CusName2,"
MySQL = MySQL & "                      TblCustemers_1.CusNamee AS CusName2E, TblCustemers_1.Fullcode AS CusFullcode2, dbo.TravKItemDet1.[Count], dbo.TravKItemDet1.ItemID,"
MySQL = MySQL & "                      dbo.TblItems.itemcode , dbo.TblItems.itemname, dbo.TblItems.ItemNamee, dbo.TravKItemDet1.UnitID, dbo.TblUnites.UnitName, dbo.TblUnites.UnitNamee"
MySQL = MySQL & " FROM         dbo.TblUnites RIGHT OUTER JOIN "
MySQL = MySQL & "                      dbo.TravKItemDet1 ON dbo.TblUnites.UnitID = dbo.TravKItemDet1.UnitID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblItems ON dbo.TravKItemDet1.ItemID = dbo.TblItems.ItemID RIGHT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblOrderUpload ON dbo.TravKItemDet1.MasterID = dbo.TblOrderUpload.ID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblCustemers TblCustemers_1 ON dbo.TblOrderUpload.CustId1 = TblCustemers_1.CusID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblVendorCars TblVendorCars_1 ON dbo.TblOrderUpload.SupplemID2 = TblVendorCars_1.ID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.FixedAssets ON dbo.TblOrderUpload.SupplemID = dbo.FixedAssets.id LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblVendorCars TblVendorCars_2 ON dbo.TblOrderUpload.CarID2 = TblVendorCars_2.ID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblCarsData ON dbo.TblOrderUpload.CarID = dbo.TblCarsData.id LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblCountriesGovernments TblCountriesGovernments_1 ON dbo.TblOrderUpload.CityID2 = TblCountriesGovernments_1.GovernmentID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblCountriesGovernments TblCountriesGovernments_2 ON dbo.TblOrderUpload.CityID = TblCountriesGovernments_2.GovernmentID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblCustemers ON dbo.TblOrderUpload.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblEmployee ON dbo.TblOrderUpload.EmpID = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
MySQL = MySQL & "                      dbo.TblBranchesData ON dbo.TblOrderUpload.BranchID = dbo.TblBranchesData.branch_id"
   
db_createOrUpdateviewSQL "View_RepOrderUplaod", MySQL

DB_CreateField "TblOrderUpload", "TimeOrder", adDBTimeStamp, adColNullable, , , , False, True
DB_CreateField "TblOrderUpload", "CountOrders", adDouble, adColNullable, , , , False, True


DB_CreateField "TblBranchesData", "StoreId", adInteger, adColNullable, , , " ???    ", False, True

DB_CreateField "Tbl_BusinessDialyDet", "TConID", adInteger, adColNullable, , , , False, True

DB_CreateField "Notes", "TotalNotesValue", adCurrency, adColNullable, , , " ???    ", False, True



DB_CreateField "TblDefComItemDet", "IsAdd", adBoolean, adColNullable, , , "                ", False, True
DB_CreateField "TblDefComItem", "Price", adCurrency, adColNullable, , , , False, True
DB_CreateField "TblDefComItem", "TotalAdd", adCurrency, adColNullable, , , , False, True
DB_CreateField "TblDefComItem", "GroupID", adInteger, adColNullable, , , , False, True


DB_CreateField "TblOptions", "NotAllowedCalcVata", adBoolean, adColNullable, , , "                ", False, True


DB_CreateField "TblTravDueK", "RecNo", adInteger, adColNullable, , , , False, True
DB_CreateField "TblTravDueK", "Weight", adCurrency, adColNullable, , , , False, True
DB_CreateField "TblTravDueK", "CarNo", adInteger, adColNullable, , , , False, True

RecNo Weight

DB_CreateField "notes_all", "RecNo", adInteger, adColNullable, , , , False, True
DB_CreateField "notes_all", "Weight", adCurrency, adColNullable, , , , False, True

DB_CreateField "TblTravDueKDet", "RecNo", adInteger, adColNullable, , , , False, True
DB_CreateField "TblTravDueKDet", "Weight", adCurrency, adColNullable, , , , False, True
DB_CreateField "TblTravDueKDet", "Price", adCurrency, adColNullable, , , , False, True


DB_CreateField "TblEmployee", "chkStop", adBoolean, adColNullable, , "0", "        ", False, True

chkStop.Enabled = False

DB_CreateField "TblTravDueK", "chkTypeTransport", adBoolean, adColNullable, , "0", "        ", False, True

End Sub

Private Sub Command2_Click()
   For i = 1 To xReport.FormulaFields.count
        Select Case xReport.FormulaFields.Item(i).Name
        Case "{@ArabicInterface}"
            Rpt.FormulaFields.Item(i).Text = ArabicInterface
        End Select
    Next i

    xReport.EnableParameterPrompting = False
    For i = 1 To xReport.ParameterFields.count
        Select Case xReport.ParameterFields.Item(i).ParameterFieldName
        Case "@FrmDate"
            xReport.ParameterFields.Item(i).AddCurrentValue CurrentBranch
        Case "@FrmDate"
            xReport.ParameterFields.Item(i).AddCurrentValue CurrentBranch
        Case "@FrmDate"
            xReport.ParameterFields.Item(i).AddCurrentValue CurrentBranch
   
        End Select
    Next i
    
    
    
End Sub




Dim StrSQL  As String

    
    StrSQL = " SELECT     dbo.TblTravDueKDet.ID, dbo.TblTravDueKDet.TravID,dbo.TblTravDueK.RdQty ,dbo.TblTravDueKDet.TripNo, dbo.TblTravDueKDet.TripDate, dbo.TblTravDueKDet.BranchID, "
StrSQL = StrSQL & "                          TblTravDueK.RecordDate ,dbo.TblTravDueK.TotalValue , dbo.TblTravDueK.Vat,dbo.TblTravDueK.TotalValue + TblTravDueK.Vat as NetValue,"
StrSQL = StrSQL & "                      dbo.TblTravDueKDet.Price,TblTravDueKDet.RecNo,TblTravDueKDet.Weight,"
StrSQL = StrSQL & "                      dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblTravDueKDet.Typed, dbo.TblTravDueKDet.[Value], dbo.TblTravDueKDet.Remarks,"
StrSQL = StrSQL & "                      dbo.TblTravDueKDet.NoteID, dbo.TblTravDueKDet.QtyDownload, dbo.TblTravDueKDet.QtyDischarge, dbo.TblTravDueKDet.CardNO, dbo.TblTravDueKDet.CardNO2,"
StrSQL = StrSQL & "                      dbo.TblTravDueKDet.CarType1, dbo.TblTravDueKDet.CarID, dbo.TblCarsData.BoardNO, dbo.TblVendorCars.BoardNo AS BoardNo2, dbo.TblTravDueKDet.FromID,"
StrSQL = StrSQL & "                      TblCountriesGovernments_2.GovernmentName, dbo.TblTravDueKDet.ToID, TblCountriesGovernments_1.GovernmentName AS ToGovernmentName,"
StrSQL = StrSQL & "                      dbo.TblTravDueKDet.CarTypeID, dbo.TBLCarTypes.name, dbo.TBLCarTypes.namee, dbo.TblTravDueKDet.TypeTrans, dbo.TblTravDueKDet.ShipID,"
StrSQL = StrSQL & "                      dbo.TblShipsData.Name AS ShipName, dbo.TblShipsData.NameE AS ShipNameE, dbo.TblTravDueKDet.LeaderName,"
StrSQL = StrSQL & "                      tc.CusName , tc.VATNO, tc.Address,TblTravDueK.noteserial1"
StrSQL = StrSQL & " FROM         dbo.TblTravDueKDet LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblShipsData ON dbo.TblTravDueKDet.ShipID = dbo.TblShipsData.id LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TBLCarTypes ON dbo.TblTravDueKDet.CarTypeID = dbo.TBLCarTypes.id LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblCountriesGovernments TblCountriesGovernments_1 ON dbo.TblTravDueKDet.ToID = TblCountriesGovernments_1.GovernmentID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblCountriesGovernments TblCountriesGovernments_2 ON dbo.TblTravDueKDet.FromID = TblCountriesGovernments_2.GovernmentID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblVendorCars ON dbo.TblTravDueKDet.CarID = dbo.TblVendorCars.ID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblCarsData ON dbo.TblTravDueKDet.CarID = dbo.TblCarsData.id LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblBranchesData ON dbo.TblTravDueKDet.BranchID = dbo.TblBranchesData.branch_id"
StrSQL = StrSQL & "                                  LEFT OUTER JOIN dbo.TblTravDueK"
StrSQL = StrSQL & "                                              ON  dbo.TblTravDueK.ID = dbo.TblTravDueKDet.TravID"
StrSQL = StrSQL & "                                              LEFT OUTER JOIN dbo.TblCustemers AS tc"
StrSQL = StrSQL & "                                              ON  tc.CusId = dbo.TblTravDueK.CusId"
StrSQL = StrSQL & "   Where 1= 1 and (dbo.TblTravDueKDet.TypeTrans is null or dbo.TblTravDueKDet.TypeTrans=0)  "
db_createOrUpdateviewSQL "View_TblTravDueKDet", StrSQL

Qty1 , ItemNameID, UnitID, UserID, Qty, GroupID, DateStart, DateEnd, PrintStiker

    If DB_CreateTable("TblProductLineDistribution", True, "ID ", False) = True Then
        DB_CreateField "TblProductLineDistribution", "IDDefCIT", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblProductLineDistribution", "ProductLineID", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblProductLineDistribution", "ItemNameID", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblProductLineDistribution", "UnitID", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblProductLineDistribution", "UserID", adInteger, adColNullable, , , "  ", False, True
        DB_CreateField "TblProductLineDistribution", "GroupID", adInteger, adColNullable, , , "  ", False, True
        
        DB_CreateField "TblProductLineDistribution", "Qty1", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "TblProductLineDistribution", "Qty", adDouble, adColNullable, , , "    ", False, True
        DB_CreateField "TblProductLineDistribution", "RecordeDate", adDBTimeStamp, adColNullable, , , "      ", False, True
        DB_CreateField "TblProductLineDistribution", "DateStart", adDBTimeStamp, adColNullable, , , "      ", False, True
        DB_CreateField "TblProductLineDistribution", "DateEnd", adDBTimeStamp, adColNullable, , , "      ", False, True
        DB_CreateField "TblProductLineDistribution", "PrintStiker", adVarWChar, adColNullable, 4000, , "C?C??   ", False, True, , True
        
        
        DB_CreateField "transactions", "ProductLineID", adInteger, adColNullable, , , "  ", False, True
        
    End If
    
    
