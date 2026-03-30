VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{D155F1AE-D9A4-458C-8CEE-498CB717DB7B}#1.0#0"; "DBPix20.ocx"
Object = "{D95CB779-00CB-4B49-97B9-9F0B61CAB3C1}#4.0#0"; "biokey.ocx"
Begin VB.Form dean 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   10125
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   19410
   Icon            =   "dean.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   10125
   ScaleWidth      =   19410
   Begin C1SizerLibCtl.C1Tab TabMain 
      Height          =   10125
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   19410
      _cx             =   34237
      _cy             =   17859
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
      Caption         =   $"dean.frx":058A
      Align           =   5
      CurrTab         =   2
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
         Height          =   9750
         Index           =   1
         Left            =   -20265
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   45
         Width           =   19320
         _cx             =   34078
         _cy             =   17198
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
            BackColor       =   &H00E2E9E9&
            Height          =   1440
            Index           =   1
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   6
            Top             =   0
            Width           =   19515
            Begin VB.TextBox TxtModFlg2 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H0000FF00&
               Enabled         =   0   'False
               Height          =   285
               Index           =   0
               Left            =   2580
               RightToLeft     =   -1  'True
               TabIndex        =   7
               Text            =   "modflag"
               Top             =   90
               Visible         =   0   'False
               Width           =   345
            End
            Begin MSComctlLib.ImageList GrdImageList2 
               Index           =   1
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
                     Picture         =   "dean.frx":066F
                     Key             =   "CompanyName"
                  EndProperty
                  BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "dean.frx":0A09
                     Key             =   "Ser"
                  EndProperty
                  BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "dean.frx":0DA3
                     Key             =   "Vac_Name"
                  EndProperty
                  BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "dean.frx":113D
                     Key             =   "ShareCount"
                  EndProperty
                  BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "dean.frx":14D7
                     Key             =   "Dis_Count"
                  EndProperty
                  BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "dean.frx":1871
                     Key             =   "Bouns"
                  EndProperty
                  BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "dean.frx":1C0B
                     Key             =   "SharesValue"
                  EndProperty
                  BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "dean.frx":21A5
                     Key             =   "BuyValue"
                  EndProperty
               EndProperty
            End
            Begin ImpulseButton.ISButton btn_Last 
               Height          =   315
               Index           =   0
               Left            =   450
               TabIndex        =   8
               Top             =   510
               Width           =   765
               _ExtentX        =   1349
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
               ButtonImage     =   "dean.frx":253F
               ColorButton     =   14871017
               AcclimateGrayTones=   -1  'True
               DrawFocusRectangle=   0   'False
               DisabledImageExtraction=   0
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btn_Next 
               Height          =   315
               Index           =   0
               Left            =   315
               TabIndex        =   9
               Top             =   510
               Width           =   1605
               _ExtentX        =   2831
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
               ButtonImage     =   "dean.frx":28D9
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btn_Previous 
               Height          =   315
               Index           =   0
               Left            =   1875
               TabIndex        =   10
               Top             =   510
               Width           =   525
               _ExtentX        =   926
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
               ButtonImage     =   "dean.frx":2C73
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btn_First 
               Height          =   435
               Index           =   0
               Left            =   2580
               TabIndex        =   11
               Top             =   450
               Width           =   525
               _ExtentX        =   926
               _ExtentY        =   767
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
               ButtonImage     =   "dean.frx":300D
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«‰Ê«⁄ «·„Â«„"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   36
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   1215
               Index           =   8
               Left            =   9480
               RightToLeft     =   -1  'True
               TabIndex        =   12
               Top             =   270
               Width           =   5760
            End
         End
         Begin VB.TextBox txtNamee 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFF80&
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
            Index           =   0
            Left            =   5910
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   5
            Top             =   6630
            Width           =   2565
         End
         Begin VB.TextBox TxtName 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFF80&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   0
            Left            =   5910
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   4
            Top             =   6225
            Width           =   2370
         End
         Begin VB.TextBox TxtSerial1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFF80&
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
            Index           =   0
            Left            =   6705
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   3
            Top             =   5775
            Width           =   1575
         End
         Begin VB.TextBox txtPercentV 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFF80&
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
            Left            =   6315
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   2
            Top             =   7020
            Width           =   1965
         End
         Begin ImpulseButton.ISButton btn_Save 
            Height          =   435
            Index           =   0
            Left            =   6315
            TabIndex        =   13
            Top             =   7440
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   767
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "Õ›Ÿ"
            BackColor       =   14871017
            FontSize        =   9.75
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "dean.frx":33A7
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton Btn_Undo 
            Height          =   435
            Index           =   0
            Left            =   5130
            TabIndex        =   14
            Top             =   7440
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   767
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   " —«Ã⁄"
            BackColor       =   14871017
            FontSize        =   9.75
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "dean.frx":3741
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Delete 
            Height          =   435
            Index           =   0
            Left            =   4140
            TabIndex        =   15
            Top             =   7440
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   767
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "Õ–›"
            BackColor       =   14871017
            FontSize        =   9.75
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "dean.frx":3ADB
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Cancel 
            Height          =   435
            Index           =   0
            Left            =   1185
            TabIndex        =   16
            Top             =   7440
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   767
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "Œ—ÊÃ"
            BackColor       =   14871017
            FontSize        =   9.75
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "dean.frx":4075
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton Btn_Print 
            Height          =   525
            Index           =   0
            Left            =   3150
            TabIndex        =   17
            TabStop         =   0   'False
            ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
            Top             =   7395
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   926
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "ÿ»«⁄… "
            BackColor       =   14871017
            FontSize        =   9.75
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "dean.frx":440F
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btn_Query 
            Height          =   585
            Index           =   0
            Left            =   2175
            TabIndex        =   18
            TabStop         =   0   'False
            ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
            Top             =   7365
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   1032
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "»ÕÀ"
            BackColor       =   14871017
            FontSize        =   9.75
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "dean.frx":AC71
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin VSFlex8Ctl.VSFlexGrid Grid1 
            Height          =   4230
            Left            =   0
            TabIndex        =   19
            Top             =   1440
            Width           =   9465
            _cx             =   16695
            _cy             =   7461
            Appearance      =   1
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   0
            BackColor       =   16777215
            ForeColor       =   -2147483640
            BackColorFixed  =   16777215
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
            Cols            =   6
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   600
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"dean.frx":B00B
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
         Begin ImpulseButton.ISButton btn_New 
            Height          =   435
            Index           =   0
            Left            =   8280
            TabIndex        =   495
            Top             =   7440
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   767
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "ÃœÌœ"
            BackColor       =   14871017
            FontSize        =   9.75
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "dean.frx":B0D0
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Modify 
            Height          =   435
            Index           =   0
            Left            =   7095
            TabIndex        =   496
            Top             =   7440
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   767
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   " ⁄œÌ·"
            BackColor       =   14871017
            FontSize        =   9.75
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "dean.frx":B46A
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin VB.Label LabCount_Rec 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   240
            Index           =   0
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   27
            Top             =   9420
            Width           =   390
         End
         Begin VB.Label LabCurr_Rec 
            BackColor       =   &H00E2E9E9&
            Height          =   240
            Index           =   0
            Left            =   1575
            RightToLeft     =   -1  'True
            TabIndex        =   26
            Top             =   9375
            Width           =   990
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «·”Ã·« :"
            Height          =   210
            Index           =   5
            Left            =   390
            RightToLeft     =   -1  'True
            TabIndex        =   25
            Top             =   9375
            Width           =   1185
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”Ã· «·Õ«·Ì:"
            Height          =   210
            Index           =   4
            Left            =   2565
            RightToLeft     =   -1  'True
            TabIndex        =   24
            Top             =   9375
            Width           =   780
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «‰Ã·Ì“Ì"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   8670
            RightToLeft     =   -1  'True
            TabIndex        =   23
            Top             =   6600
            Width           =   1185
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ ⁄—»Ì"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   0
            Left            =   7890
            RightToLeft     =   -1  'True
            TabIndex        =   22
            Top             =   6225
            Width           =   1575
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·ﬂÊœ "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   7890
            RightToLeft     =   -1  'True
            TabIndex        =   21
            Top             =   5730
            Width           =   1380
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·‰”»…"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   6
            Left            =   8280
            RightToLeft     =   -1  'True
            TabIndex        =   20
            Top             =   6990
            Width           =   1185
         End
      End
      Begin C1SizerLibCtl.C1Elastic ELe 
         Height          =   9750
         Index           =   2
         Left            =   -19965
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   45
         Width           =   19320
         _cx             =   34078
         _cy             =   17198
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
            BackColor       =   &H00E2E9E9&
            BorderStyle     =   0  'None
            Height          =   690
            Index           =   0
            Left            =   195
            RightToLeft     =   -1  'True
            TabIndex        =   36
            Top             =   60
            Width           =   19125
            Begin VB.TextBox TxtModFlg2 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H0000FF00&
               Enabled         =   0   'False
               Height          =   285
               Index           =   1
               Left            =   2580
               RightToLeft     =   -1  'True
               TabIndex        =   37
               Text            =   "modflag"
               Top             =   90
               Visible         =   0   'False
               Width           =   465
            End
            Begin MSComctlLib.ImageList GrdImageList2 
               Index           =   0
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
                     Picture         =   "dean.frx":B804
                     Key             =   "CompanyName"
                  EndProperty
                  BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "dean.frx":BB9E
                     Key             =   "Ser"
                  EndProperty
                  BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "dean.frx":BF38
                     Key             =   "Vac_Name"
                  EndProperty
                  BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "dean.frx":C2D2
                     Key             =   "ShareCount"
                  EndProperty
                  BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "dean.frx":C66C
                     Key             =   "Dis_Count"
                  EndProperty
                  BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "dean.frx":CA06
                     Key             =   "Bouns"
                  EndProperty
                  BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "dean.frx":CDA0
                     Key             =   "SharesValue"
                  EndProperty
                  BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "dean.frx":D33A
                     Key             =   "BuyValue"
                  EndProperty
               EndProperty
            End
            Begin ImpulseButton.ISButton btn_Last 
               Height          =   315
               Index           =   1
               Left            =   90
               TabIndex        =   38
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
               ButtonImage     =   "dean.frx":D6D4
               ColorButton     =   14871017
               AcclimateGrayTones=   -1  'True
               DrawFocusRectangle=   0   'False
               DisabledImageExtraction=   0
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btn_Next 
               Height          =   315
               Index           =   1
               Left            =   555
               TabIndex        =   39
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
               ButtonImage     =   "dean.frx":DA6E
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btn_Previous 
               Height          =   315
               Index           =   1
               Left            =   1155
               TabIndex        =   40
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
               ButtonImage     =   "dean.frx":DE08
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btn_First 
               Height          =   315
               Index           =   1
               Left            =   1620
               TabIndex        =   41
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
               ButtonImage     =   "dean.frx":E1A2
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«‰Ê«⁄ «·„ﬁ«”« "
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   24
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   1095
               Index           =   5
               Left            =   7800
               RightToLeft     =   -1  'True
               TabIndex        =   42
               Top             =   0
               Width           =   4920
            End
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00E2E9E9&
            BorderStyle     =   0  'None
            Height          =   1410
            Index           =   1
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   29
            Top             =   4620
            Width           =   12810
            Begin VB.TextBox TxtSerial1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFF80&
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
               Height          =   675
               Index           =   1
               Left            =   10995
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   32
               Top             =   300
               Width           =   2760
            End
            Begin VB.TextBox TxtName 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFF80&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   675
               Index           =   1
               Left            =   4035
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   31
               Top             =   420
               Width           =   4920
            End
            Begin VB.TextBox txtNamee 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFF80&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   675
               Index           =   1
               Left            =   4035
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   30
               Top             =   1380
               Width           =   4920
            End
            Begin VB.Label Label6 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ﬂÊœ "
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   18
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   555
               Index           =   2
               Left            =   12735
               RightToLeft     =   -1  'True
               TabIndex        =   35
               Top             =   330
               Width           =   2190
            End
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ ⁄—»Ì"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   18
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   645
               Index           =   2
               Left            =   6870
               RightToLeft     =   -1  'True
               TabIndex        =   34
               Top             =   420
               Width           =   3390
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «‰Ã·Ì“Ì"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   18
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   645
               Index           =   2
               Left            =   6960
               RightToLeft     =   -1  'True
               TabIndex        =   33
               Top             =   1380
               Width           =   3540
            End
         End
         Begin ImpulseButton.ISButton btn_New 
            Height          =   435
            Index           =   1
            Left            =   9075
            TabIndex        =   43
            Top             =   6060
            Width           =   585
            _ExtentX        =   1032
            _ExtentY        =   767
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "ÃœÌœ"
            BackColor       =   14871017
            FontSize        =   9.75
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "dean.frx":E53C
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Save 
            Height          =   435
            Index           =   1
            Left            =   6900
            TabIndex        =   44
            Top             =   6060
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   767
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "Õ›Ÿ"
            BackColor       =   14871017
            FontSize        =   9.75
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "dean.frx":E8D6
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Modify 
            Height          =   435
            Index           =   1
            Left            =   7890
            TabIndex        =   45
            Top             =   6060
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   767
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   " ⁄œÌ·"
            BackColor       =   14871017
            FontSize        =   9.75
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "dean.frx":EC70
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton Btn_Undo 
            Height          =   435
            Index           =   1
            Left            =   5520
            TabIndex        =   46
            Top             =   6060
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   767
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   " —«Ã⁄"
            BackColor       =   14871017
            FontSize        =   9.75
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "dean.frx":F00A
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Delete 
            Height          =   435
            Index           =   1
            Left            =   4335
            TabIndex        =   47
            Top             =   6060
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   767
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "Õ–›"
            BackColor       =   14871017
            FontSize        =   9.75
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "dean.frx":F3A4
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Cancel 
            Height          =   435
            Index           =   1
            Left            =   195
            TabIndex        =   48
            Top             =   6060
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   767
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "Œ—ÊÃ"
            BackColor       =   14871017
            FontSize        =   9.75
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "dean.frx":F93E
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton Btn_Print 
            Height          =   525
            Index           =   1
            Left            =   2760
            TabIndex        =   49
            TabStop         =   0   'False
            ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
            Top             =   6015
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   926
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "ÿ»«⁄… "
            BackColor       =   14871017
            FontSize        =   9.75
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "dean.frx":FCD8
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btn_Query 
            Height          =   585
            Index           =   1
            Left            =   1575
            TabIndex        =   50
            TabStop         =   0   'False
            ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
            Top             =   5985
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   1032
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "»ÕÀ"
            BackColor       =   14871017
            FontSize        =   9.75
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "dean.frx":1653A
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin VSFlex8Ctl.VSFlexGrid Grid2 
            Height          =   3660
            Left            =   0
            TabIndex        =   51
            Top             =   885
            Width           =   10245
            _cx             =   18071
            _cy             =   6456
            Appearance      =   1
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            BackColorFixed  =   16777215
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
            Cols            =   4
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   600
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"dean.frx":168D4
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
            Height          =   240
            Index           =   1
            Left            =   2370
            RightToLeft     =   -1  'True
            TabIndex        =   55
            Top             =   6570
            Width           =   390
         End
         Begin VB.Label LabCurr_Rec 
            BackColor       =   &H00E2E9E9&
            Height          =   240
            Index           =   1
            Left            =   4725
            RightToLeft     =   -1  'True
            TabIndex        =   54
            Top             =   6570
            Width           =   795
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «·”Ã·« :"
            Height          =   210
            Index           =   2
            Left            =   2955
            RightToLeft     =   -1  'True
            TabIndex        =   53
            Top             =   6570
            Width           =   1185
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”Ã· «·Õ«·Ì:"
            Height          =   210
            Index           =   3
            Left            =   5520
            RightToLeft     =   -1  'True
            TabIndex        =   52
            Top             =   6525
            Width           =   1185
         End
      End
      Begin C1SizerLibCtl.C1Elastic ELe 
         Height          =   9750
         Index           =   0
         Left            =   45
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   45
         Width           =   19320
         _cx             =   34078
         _cy             =   17198
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
         Begin VB.Frame Frm2 
            BackColor       =   &H00E2E9E9&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   7290
            Left            =   195
            RightToLeft     =   -1  'True
            TabIndex        =   559
            Top             =   600
            Width           =   18735
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
               Height          =   315
               Left            =   6225
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   562
               Tag             =   "enter English Name"
               Top             =   5055
               Width           =   5040
            End
            Begin VB.TextBox TxtVacName 
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
               Left            =   6225
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   561
               Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·œÌ«‰Â"
               Top             =   4695
               Width           =   5040
            End
            Begin VB.TextBox TxtSerial 
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
               Left            =   10230
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   560
               Top             =   4335
               Width           =   1065
            End
            Begin VSFlex8Ctl.VSFlexGrid Grid 
               Height          =   3450
               Left            =   60
               TabIndex        =   563
               Top             =   270
               Width           =   15270
               _cx             =   26935
               _cy             =   6085
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
               ExtendLastCol   =   0   'False
               FormatString    =   $"dean.frx":16964
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
            Begin ImpulseButton.ISButton btnNew 
               Height          =   330
               Left            =   14370
               TabIndex        =   567
               Top             =   6270
               Width           =   750
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
               ButtonImage     =   "dean.frx":169EC
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnSave 
               Height          =   330
               Left            =   10800
               TabIndex        =   568
               Top             =   6270
               Width           =   750
               _ExtentX        =   1323
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
               ButtonImage     =   "dean.frx":16D86
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnModify 
               Height          =   330
               Left            =   12600
               TabIndex        =   569
               Top             =   6270
               Width           =   750
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
               ButtonImage     =   "dean.frx":17120
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton BtnUndo 
               Height          =   330
               Left            =   8040
               TabIndex        =   570
               Top             =   6270
               Width           =   750
               _ExtentX        =   1323
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
               ButtonImage     =   "dean.frx":174BA
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnDelete 
               Height          =   330
               Left            =   5670
               TabIndex        =   571
               Top             =   6270
               Width           =   750
               _ExtentX        =   1323
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
               ButtonImage     =   "dean.frx":17854
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton BtnPrint 
               Height          =   285
               Left            =   4485
               TabIndex        =   572
               TabStop         =   0   'False
               Top             =   5760
               Visible         =   0   'False
               Width           =   285
               _ExtentX        =   503
               _ExtentY        =   503
               ButtonStyle     =   1
               ButtonPositionImage=   2
               Caption         =   ""
               BackColor       =   14871017
               FontSize        =   14.25
               FontName        =   "Arial"
               FontBold        =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   14.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ButtonImage     =   "dean.frx":17DEE
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnCancel 
               Height          =   330
               Left            =   2490
               TabIndex        =   573
               Top             =   6270
               Width           =   750
               _ExtentX        =   1323
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
               ButtonImage     =   "dean.frx":18188
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·”Ã· «·Õ«·Ì:"
               Height          =   210
               Index           =   0
               Left            =   2265
               RightToLeft     =   -1  'True
               TabIndex        =   577
               Top             =   5835
               Width           =   975
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄œœ «·”Ã·« :"
               Height          =   210
               Index           =   1
               Left            =   570
               RightToLeft     =   -1  'True
               TabIndex        =   576
               Top             =   5835
               Width           =   975
            End
            Begin VB.Label LabCurrRec 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Height          =   210
               Left            =   1560
               RightToLeft     =   -1  'True
               TabIndex        =   575
               Top             =   5850
               Width           =   675
            End
            Begin VB.Label LabCountRec 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Height          =   210
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   574
               Top             =   5835
               Width           =   540
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«‰Ã·Ì“Ì"
               Height          =   285
               Index           =   1
               Left            =   10935
               RightToLeft     =   -1  'True
               TabIndex        =   566
               Top             =   5130
               Width           =   1890
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄—»Ì"
               Height          =   285
               Index           =   3
               Left            =   10890
               RightToLeft     =   -1  'True
               TabIndex        =   565
               Top             =   4770
               Width           =   1890
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬂÊœ  "
               Height          =   195
               Index           =   6
               Left            =   11775
               RightToLeft     =   -1  'True
               TabIndex        =   564
               Top             =   4320
               Width           =   990
            End
         End
         Begin VB.Frame FraHeader 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   540
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   57
            Top             =   0
            Width           =   19320
            Begin VB.TextBox TxtModFlg 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H0000FF00&
               Enabled         =   0   'False
               Height          =   285
               Left            =   6090
               RightToLeft     =   -1  'True
               TabIndex        =   58
               Text            =   "modflag"
               Top             =   -30
               Visible         =   0   'False
               Width           =   465
            End
            Begin ImpulseButton.ISButton btnLast 
               Height          =   315
               Left            =   450
               TabIndex        =   59
               Top             =   240
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   556
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
               ButtonImage     =   "dean.frx":18522
               ColorButton     =   16777215
               AcclimateGrayTones=   -1  'True
               DrawFocusRectangle=   0   'False
               DisabledImageExtraction=   0
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnNext 
               Height          =   315
               Left            =   915
               TabIndex        =   60
               Top             =   240
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   556
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
               ButtonImage     =   "dean.frx":188BC
               ColorButton     =   16777215
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnPrevious 
               Height          =   315
               Left            =   1515
               TabIndex        =   61
               Top             =   240
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   556
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
               ButtonImage     =   "dean.frx":18C56
               ColorButton     =   16777215
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnFirst 
               Height          =   315
               Left            =   2040
               TabIndex        =   62
               Top             =   240
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   556
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
               ButtonImage     =   "dean.frx":18FF0
               ColorButton     =   16777215
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin MSComctlLib.ImageList GrdImageList2 
               Index           =   6
               Left            =   8520
               Top             =   -120
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
                     Picture         =   "dean.frx":1938A
                     Key             =   "CompanyName"
                  EndProperty
                  BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "dean.frx":19724
                     Key             =   "Ser"
                  EndProperty
                  BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "dean.frx":19ABE
                     Key             =   "Vac_Name"
                  EndProperty
                  BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "dean.frx":19E58
                     Key             =   "ShareCount"
                  EndProperty
                  BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "dean.frx":1A1F2
                     Key             =   "Dis_Count"
                  EndProperty
                  BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "dean.frx":1A58C
                     Key             =   "Bouns"
                  EndProperty
                  BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "dean.frx":1A926
                     Key             =   "SharesValue"
                  EndProperty
                  BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "dean.frx":1AEC0
                     Key             =   "BuyValue"
                  EndProperty
               EndProperty
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«·œÌ«‰« "
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
               Left            =   8490
               RightToLeft     =   -1  'True
               TabIndex        =   63
               Top             =   60
               Width           =   3720
            End
            Begin VB.Image GrdImageList 
               Height          =   612
               Left            =   12960
               Picture         =   "dean.frx":1B25A
               Stretch         =   -1  'True
               Top             =   120
               Visible         =   0   'False
               Width           =   732
            End
         End
         Begin C1SizerLibCtl.C1Elastic EltCont 
            Height          =   1920
            Left            =   0
            TabIndex        =   64
            TabStop         =   0   'False
            Top             =   6930
            Width           =   18735
            _cx             =   33046
            _cy             =   3387
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
            BorderWidth     =   1
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
         End
         Begin MSDataListLib.DataCombo DCUser 
            Height          =   315
            Left            =   6510
            TabIndex        =   65
            Top             =   6960
            Width           =   3150
            _ExtentX        =   5556
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
      End
      Begin C1SizerLibCtl.C1Elastic ELe 
         Height          =   9750
         Index           =   3
         Left            =   20055
         TabIndex        =   66
         TabStop         =   0   'False
         Top             =   45
         Width           =   19320
         _cx             =   34078
         _cy             =   17198
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
         Begin VB.Frame Frame1 
            Height          =   4095
            Index           =   0
            Left            =   8040
            TabIndex        =   596
            Top             =   1800
            Visible         =   0   'False
            Width           =   6510
            Begin VB.TextBox txtCustomerName2 
               Alignment       =   1  'Right Justify
               Height          =   345
               Left            =   2400
               RightToLeft     =   -1  'True
               TabIndex        =   599
               Top             =   3240
               Width           =   2955
            End
            Begin VB.TextBox XPTxtCusID 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   2400
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   598
               Top             =   2880
               Width           =   2955
            End
            Begin VB.TextBox XPTxtPhone 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   2400
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   597
               Top             =   3720
               Width           =   2955
            End
            Begin VSFlex8UCtl.VSFlexGrid FGS 
               Height          =   2265
               Left            =   240
               TabIndex        =   600
               Top             =   480
               Width           =   6195
               _cx             =   10927
               _cy             =   3995
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
               Cols            =   7
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   -1  'True
               FormatString    =   $"dean.frx":1C65F
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
               Index           =   20
               Left            =   720
               TabIndex        =   601
               Top             =   3600
               Width           =   915
               _ExtentX        =   1614
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
            Begin VB.Label XPLbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·«”„"
               Height          =   315
               Index           =   16
               Left            =   5160
               RightToLeft     =   -1  'True
               TabIndex        =   605
               Top             =   3270
               Width           =   1215
            End
            Begin VB.Label XPLbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ﬂÊœ"
               Height          =   315
               Index           =   2
               Left            =   5160
               RightToLeft     =   -1  'True
               TabIndex        =   604
               Top             =   2880
               Width           =   1215
            End
            Begin VB.Label XPLbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·Â« ›"
               Height          =   315
               Index           =   5
               Left            =   5160
               RightToLeft     =   -1  'True
               TabIndex        =   603
               Top             =   3720
               Width           =   1215
            End
            Begin VB.Label lblexit 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "X"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   270
               Index           =   0
               Left            =   5280
               TabIndex        =   602
               Top             =   120
               Width           =   570
            End
         End
         Begin VB.TextBox txtTotalNet 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   255
            Left            =   195
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   547
            Top             =   5625
            Width           =   990
         End
         Begin VB.TextBox txtTotalDisc 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Height          =   270
            Left            =   195
            RightToLeft     =   -1  'True
            TabIndex        =   546
            Top             =   2595
            Width           =   990
         End
         Begin VB.TextBox txtTotalPay 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Height          =   255
            Left            =   195
            RightToLeft     =   -1  'True
            TabIndex        =   545
            Top             =   1770
            Width           =   990
         End
         Begin VB.TextBox txtTotalAdd 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Height          =   255
            Left            =   195
            RightToLeft     =   -1  'True
            TabIndex        =   544
            Top             =   1335
            Width           =   990
         End
         Begin VB.TextBox txtGeneralTotal 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Height          =   255
            Left            =   195
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   543
            Top             =   915
            Width           =   990
         End
         Begin VB.TextBox txtTotalDiscPerc 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Height          =   255
            Left            =   195
            RightToLeft     =   -1  'True
            TabIndex        =   542
            Top             =   2160
            Width           =   990
         End
         Begin VB.TextBox txtRequiredAmount 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Height          =   255
            Left            =   195
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   541
            Top             =   3075
            Width           =   990
         End
         Begin VB.TextBox txtVatYou 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Enabled         =   0   'False
            Height          =   270
            Left            =   195
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   540
            Text            =   "15"
            Top             =   3525
            Width           =   990
         End
         Begin VB.TextBox txtTotalAfterVat 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Height          =   255
            Left            =   195
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   539
            Top             =   4545
            Width           =   990
         End
         Begin VB.TextBox txtVat 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Height          =   255
            Left            =   195
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   538
            Top             =   4050
            Width           =   990
         End
         Begin VB.TextBox txtPaymedValue 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Height          =   255
            Left            =   195
            RightToLeft     =   -1  'True
            TabIndex        =   537
            Top             =   5115
            Width           =   990
         End
         Begin VB.PictureBox Picture1 
            Height          =   1770
            Left            =   3165
            ScaleHeight     =   1710
            ScaleWidth      =   1710
            TabIndex        =   536
            Top             =   4170
            Width           =   1770
         End
         Begin VB.TextBox TxtNoteSerial11 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   7290
            RightToLeft     =   -1  'True
            TabIndex        =   92
            Top             =   870
            Width           =   1380
         End
         Begin VB.CommandButton Command8 
            Caption         =   "ﬂ‘› Õ”«»"
            Height          =   270
            Left            =   4140
            RightToLeft     =   -1  'True
            TabIndex        =   91
            Top             =   1815
            Width           =   1575
         End
         Begin VB.TextBox TxtSearchCode 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   315
            Left            =   14385
            RightToLeft     =   -1  'True
            TabIndex        =   90
            Top             =   1470
            Width           =   1380
         End
         Begin VB.TextBox TxtPhone 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   9465
            TabIndex        =   89
            Top             =   1440
            Width           =   1575
         End
         Begin VB.CommandButton cmdAddCustomer 
            Caption         =   "«÷«›… ⁄„Ì· ÃœÌœ"
            Height          =   345
            Left            =   4140
            RightToLeft     =   -1  'True
            TabIndex        =   88
            Top             =   1440
            Width           =   1575
         End
         Begin VB.TextBox txtCustomerName 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   5715
            TabIndex        =   87
            Top             =   1440
            Width           =   2955
         End
         Begin VB.CommandButton Command1 
            Caption         =   "ÿ»«⁄… «·›« Ê—…"
            Height          =   435
            Left            =   3240
            RightToLeft     =   -1  'True
            TabIndex        =   86
            Top             =   2865
            Width           =   1380
         End
         Begin VB.CommandButton cmdPrintCash 
            Caption         =   "ÿ»«⁄… ”‰œ «·ﬁ»÷"
            Height          =   435
            Left            =   3240
            RightToLeft     =   -1  'True
            TabIndex        =   85
            Top             =   3360
            Width           =   1380
         End
         Begin VB.TextBox txtNoteSerialCash 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   1
            Left            =   3285
            RightToLeft     =   -1  'True
            TabIndex        =   84
            Top             =   3870
            Visible         =   0   'False
            Width           =   1380
         End
         Begin VB.TextBox txtNoteSerialCash 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   0
            Left            =   4140
            RightToLeft     =   -1  'True
            TabIndex        =   83
            Top             =   2130
            Visible         =   0   'False
            Width           =   1380
         End
         Begin VB.TextBox TXTTransactionID3 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   5715
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   82
            Top             =   120
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.TextBox TxtNoteSerial13 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   9660
            RightToLeft     =   -1  'True
            TabIndex        =   81
            Top             =   915
            Width           =   1575
         End
         Begin VB.Frame Fra_Header 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   630
            Index           =   2
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   70
            Top             =   60
            Width           =   21285
            Begin VB.TextBox TXTTransactionID1 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   14370
               RightToLeft     =   -1  'True
               TabIndex        =   75
               Top             =   360
               Visible         =   0   'False
               Width           =   825
            End
            Begin VB.TextBox TxtVac_ID 
               Alignment       =   1  'Right Justify
               Height          =   375
               Left            =   11280
               RightToLeft     =   -1  'True
               TabIndex        =   74
               Top             =   240
               Visible         =   0   'False
               Width           =   855
            End
            Begin VB.TextBox Text1 
               Alignment       =   1  'Right Justify
               Height          =   285
               Left            =   3480
               RightToLeft     =   -1  'True
               TabIndex        =   73
               Text            =   "Text1"
               Top             =   180
               Visible         =   0   'False
               Width           =   150
            End
            Begin VB.TextBox txtNoteid3 
               Alignment       =   1  'Right Justify
               Height          =   375
               Left            =   5070
               RightToLeft     =   -1  'True
               TabIndex        =   72
               Top             =   300
               Visible         =   0   'False
               Width           =   975
            End
            Begin VB.TextBox TxtModFlg2 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H0000FF00&
               Enabled         =   0   'False
               Height          =   285
               Index           =   3
               Left            =   2580
               RightToLeft     =   -1  'True
               TabIndex        =   71
               Text            =   "modflag"
               Top             =   90
               Visible         =   0   'False
               Width           =   465
            End
            Begin MSComctlLib.ImageList GrdImageList2 
               Index           =   2
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
                     Picture         =   "dean.frx":1C772
                     Key             =   "CompanyName"
                  EndProperty
                  BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "dean.frx":1CB0C
                     Key             =   "Ser"
                  EndProperty
                  BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "dean.frx":1CEA6
                     Key             =   "Vac_Name"
                  EndProperty
                  BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "dean.frx":1D240
                     Key             =   "ShareCount"
                  EndProperty
                  BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "dean.frx":1D5DA
                     Key             =   "Dis_Count"
                  EndProperty
                  BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "dean.frx":1D974
                     Key             =   "Bouns"
                  EndProperty
                  BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "dean.frx":1DD0E
                     Key             =   "SharesValue"
                  EndProperty
                  BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "dean.frx":1E2A8
                     Key             =   "BuyValue"
                  EndProperty
               EndProperty
            End
            Begin ImpulseButton.ISButton btn_Last 
               Height          =   315
               Index           =   3
               Left            =   90
               TabIndex        =   76
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
               ButtonImage     =   "dean.frx":1E642
               ColorButton     =   14871017
               AcclimateGrayTones=   -1  'True
               DrawFocusRectangle=   0   'False
               DisabledImageExtraction=   0
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btn_Next 
               Height          =   315
               Index           =   3
               Left            =   555
               TabIndex        =   77
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
               ButtonImage     =   "dean.frx":1E9DC
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btn_Previous 
               Height          =   315
               Index           =   3
               Left            =   1155
               TabIndex        =   78
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
               ButtonImage     =   "dean.frx":1ED76
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btn_First 
               Height          =   315
               Index           =   3
               Left            =   1620
               TabIndex        =   79
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
               ButtonImage     =   "dean.frx":1F110
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "√Ê«„— «·‘€·"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   36
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   735
               Index           =   10
               Left            =   11640
               RightToLeft     =   -1  'True
               TabIndex        =   80
               Top             =   -90
               Width           =   5040
            End
            Begin VB.Image ImgFavorites 
               Height          =   390
               Index           =   0
               Left            =   7560
               Picture         =   "dean.frx":1F4AA
               Stretch         =   -1  'True
               Top             =   0
               Width           =   525
            End
         End
         Begin VB.ComboBox DCOPrType 
            Height          =   315
            Left            =   14790
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   69
            Top             =   9750
            Visible         =   0   'False
            Width           =   2565
         End
         Begin VB.ComboBox DcbType 
            Height          =   315
            Left            =   1380
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   68
            Top             =   8970
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.TextBox TxtSerial1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   345
            Index           =   3
            Left            =   14385
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   67
            Top             =   915
            Width           =   1380
         End
         Begin VSFlex8UCtl.VSFlexGrid FgItems 
            Height          =   8925
            Index           =   1
            Left            =   22275
            TabIndex        =   93
            Top             =   780
            Width           =   15780
            _cx             =   27834
            _cy             =   15743
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
            FormatString    =   $"dean.frx":23112
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
         Begin MSComCtl2.DTPicker XPDtbTrans 
            Height          =   315
            Index           =   3
            Left            =   12030
            TabIndex        =   94
            Top             =   900
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   556
            _Version        =   393216
            CalendarBackColor=   16777152
            Format          =   243859457
            CurrentDate     =   38784
         End
         Begin MSDataListLib.DataCombo Dcbranch 
            Bindings        =   "dean.frx":231D2
            Height          =   315
            Index           =   3
            Left            =   4140
            TabIndex        =   95
            Top             =   945
            Width           =   2565
            _ExtentX        =   4524
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777152
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
         Begin MSDataListLib.DataCombo DcCustmer 
            Height          =   315
            Left            =   12225
            TabIndex        =   96
            Top             =   1470
            Width           =   2160
            _ExtentX        =   3810
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777152
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
         Begin MSDataListLib.DataCombo DCboUserName 
            Height          =   315
            Index           =   3
            Left            =   13995
            TabIndex        =   97
            Top             =   9465
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin ImpulseButton.ISButton btn_New 
            Height          =   360
            Index           =   3
            Left            =   13215
            TabIndex        =   98
            Top             =   8670
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   635
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "ÃœÌœ"
            BackColor       =   14871017
            FontSize        =   9.75
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "dean.frx":231E7
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Save 
            Height          =   345
            Index           =   3
            Left            =   11040
            TabIndex        =   99
            Top             =   8685
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   609
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "Õ›Ÿ"
            BackColor       =   14871017
            FontSize        =   9.75
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "dean.frx":23581
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Modify 
            Height          =   225
            Index           =   3
            Left            =   12030
            TabIndex        =   100
            Top             =   8745
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   397
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   " ⁄œÌ·"
            BackColor       =   14871017
            FontSize        =   9.75
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "dean.frx":2391B
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton Btn_Undo 
            Height          =   225
            Index           =   3
            Left            =   10050
            TabIndex        =   101
            Top             =   8745
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   397
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   " —«Ã⁄"
            BackColor       =   14871017
            FontSize        =   9.75
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "dean.frx":23CB5
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Delete 
            Height          =   345
            Index           =   3
            Left            =   9075
            TabIndex        =   102
            Top             =   8685
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   609
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "Õ–›"
            BackColor       =   14871017
            FontSize        =   9.75
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "dean.frx":2404F
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Cancel 
            Height          =   330
            Index           =   3
            Left            =   5325
            TabIndex        =   103
            Top             =   8655
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   582
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "Œ—ÊÃ"
            BackColor       =   14871017
            FontSize        =   9.75
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "dean.frx":245E9
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton Btn_Print 
            Height          =   390
            Index           =   3
            Left            =   7695
            TabIndex        =   104
            TabStop         =   0   'False
            ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
            Top             =   8655
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   688
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "ÿ»«⁄… "
            BackColor       =   14871017
            FontSize        =   9.75
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "dean.frx":24983
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btn_Query 
            Height          =   405
            Index           =   3
            Left            =   6510
            TabIndex        =   105
            TabStop         =   0   'False
            ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
            Top             =   8640
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   714
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "»ÕÀ"
            BackColor       =   14871017
            FontSize        =   9.75
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "dean.frx":2B1E5
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin C1SizerLibCtl.C1Tab TabMain2 
            Height          =   5100
            Left            =   5010
            TabIndex        =   106
            Top             =   3225
            Width           =   12420
            _cx             =   21907
            _cy             =   8996
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
            Caption         =   "»Ì«‰« "
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
               Height          =   4725
               Index           =   4
               Left            =   45
               TabIndex        =   107
               TabStop         =   0   'False
               Top             =   45
               Width           =   12330
               _cx             =   21749
               _cy             =   8334
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
               Begin VB.TextBox txtPercentTotal 
                  Alignment       =   1  'Right Justify
                  Height          =   360
                  Left            =   1725
                  RightToLeft     =   -1  'True
                  TabIndex        =   527
                  Top             =   4305
                  Width           =   870
               End
               Begin VB.TextBox TXTTransactionID5 
                  Alignment       =   1  'Right Justify
                  Height          =   480
                  Left            =   645
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   109
                  Top             =   -1125
                  Visible         =   0   'False
                  Width           =   870
               End
               Begin VB.TextBox txtNet 
                  Alignment       =   1  'Right Justify
                  Height          =   360
                  Left            =   2160
                  Locked          =   -1  'True
                  RightToLeft     =   -1  'True
                  TabIndex        =   108
                  Top             =   4755
                  Visible         =   0   'False
                  Width           =   1515
               End
               Begin VSFlex8UCtl.VSFlexGrid fg 
                  Height          =   3915
                  Left            =   0
                  TabIndex        =   110
                  Top             =   360
                  Width           =   12120
                  _cx             =   21378
                  _cy             =   6906
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
                  BackColorFixed  =   16777215
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
                  AllowBigSelection=   -1  'True
                  AllowUserResizing=   1
                  SelectionMode   =   0
                  GridLines       =   1
                  GridLinesFixed  =   2
                  GridLineWidth   =   1
                  Rows            =   12
                  Cols            =   18
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   400
                  RowHeightMax    =   740
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   -1  'True
                  FormatString    =   $"dean.frx":2B57F
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
                  WallPaperAlignment=   0
                  AccessibleName  =   ""
                  AccessibleDescription=   ""
                  AccessibleValue =   ""
                  AccessibleRole  =   24
               End
               Begin ImpulseButton.ISButton Cmd_DeleteRow 
                  Height          =   390
                  Index           =   3
                  Left            =   1515
                  TabIndex        =   528
                  Top             =   0
                  Width           =   1725
                  _ExtentX        =   3043
                  _ExtentY        =   688
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   " Õ–› ”ÿ—"
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
                  ButtonImage     =   "dean.frx":2B813
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
               Begin ImpulseButton.ISButton Cmd_DeleteAll 
                  Height          =   390
                  Index           =   3
                  Left            =   0
                  TabIndex        =   529
                  Top             =   0
                  Width           =   1515
                  _ExtentX        =   2672
                  _ExtentY        =   688
                  ButtonStyle     =   1
                  ButtonPositionImage=   1
                  Caption         =   " Õ–› «·ﬂ·"
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
                  ButtonImage     =   "dean.frx":2BDAD
                  ColorButton     =   14871017
                  DrawFocusRectangle=   0   'False
               End
            End
            Begin C1SizerLibCtl.C1Elastic ELe 
               Height          =   4725
               Index           =   5
               Left            =   13065
               TabIndex        =   111
               TabStop         =   0   'False
               Top             =   45
               Width           =   12330
               _cx             =   21749
               _cy             =   8334
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
                  Height          =   3435
                  Index           =   3
                  Left            =   8430
                  TabIndex        =   112
                  Top             =   885
                  Width           =   6285
                  _cx             =   11086
                  _cy             =   6059
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
                  FormatString    =   $"dean.frx":2C347
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
         Begin MSComCtl2.DTPicker txtDateRec 
            Height          =   315
            Left            =   12030
            TabIndex        =   113
            Top             =   2565
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   556
            _Version        =   393216
            CalendarBackColor=   16777152
            CalendarTitleBackColor=   16777152
            Format          =   244383745
            CurrentDate     =   38784
         End
         Begin MSComCtl2.DTPicker txtDateRehearsal 
            Height          =   315
            Left            =   10845
            TabIndex        =   114
            Top             =   2565
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   556
            _Version        =   393216
            CalendarBackColor=   16777152
            Format          =   244383745
            CurrentDate     =   38784
         End
         Begin MSComCtl2.DTPicker txtRehearsalDateFInish 
            Height          =   315
            Left            =   9660
            TabIndex        =   115
            Top             =   2565
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   556
            _Version        =   393216
            CalendarBackColor=   16777152
            Format          =   244383745
            CurrentDate     =   38784
         End
         Begin MSComCtl2.DTPicker txtDateDelivery 
            Height          =   315
            Left            =   8280
            TabIndex        =   116
            Top             =   2565
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   556
            _Version        =   393216
            CalendarBackColor=   16777152
            Format          =   244383745
            CurrentDate     =   38784
         End
         Begin MSComCtl2.DTPicker txtDeliveryDateFinish 
            Height          =   315
            Left            =   7095
            TabIndex        =   117
            Top             =   2565
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   556
            _Version        =   393216
            CalendarBackColor=   16777152
            CalendarTitleBackColor=   16777152
            Format          =   244383745
            CurrentDate     =   38784
         End
         Begin MSComCtl2.DTPicker txtDateDeliveryAct 
            Height          =   315
            Left            =   5910
            TabIndex        =   118
            Top             =   2565
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   556
            _Version        =   393216
            CalendarBackColor=   16777152
            Format          =   244383745
            CurrentDate     =   38784
         End
         Begin MSComCtl2.DTPicker ToDate 
            Height          =   285
            Left            =   5715
            TabIndex        =   119
            TabStop         =   0   'False
            Top             =   1815
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   503
            _Version        =   393216
            CalendarBackColor=   16777152
            CalendarTitleBackColor=   10383715
            Format          =   244449283
            CurrentDate     =   41640
         End
         Begin MSComCtl2.DTPicker FrmDate 
            Height          =   285
            Left            =   7290
            TabIndex        =   120
            TabStop         =   0   'False
            Top             =   1815
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   503
            _Version        =   393216
            CalendarBackColor=   16777152
            CalendarTitleBackColor=   10383715
            Format          =   244449283
            CurrentDate     =   41640
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·’«›Ï "
            Height          =   345
            Index           =   19
            Left            =   1935
            TabIndex        =   558
            Top             =   5655
            Width           =   795
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "„»·€ «·Œ’„"
            Height          =   300
            Index           =   18
            Left            =   1935
            RightToLeft     =   -1  'True
            TabIndex        =   557
            Top             =   2655
            Width           =   795
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·«Ã—…"
            Height          =   390
            Index           =   16
            Left            =   1740
            TabIndex        =   556
            Top             =   930
            Width           =   990
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "œ›⁄… „ﬁœ„…"
            Height          =   270
            Index           =   17
            Left            =   1740
            TabIndex        =   555
            Top             =   1800
            Width           =   990
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·«÷«›« "
            Height          =   375
            Index           =   23
            Left            =   1545
            TabIndex        =   554
            Top             =   1410
            Width           =   1185
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "‰”»… «·Œ’„"
            Height          =   270
            Index           =   9
            Left            =   1740
            TabIndex        =   553
            Top             =   2220
            Width           =   990
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·«Ã„«·Ï ﬁ»· «·÷—Ì»…"
            Height          =   345
            Index           =   10
            Left            =   945
            RightToLeft     =   -1  'True
            TabIndex        =   552
            Top             =   3195
            Width           =   1785
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·„”œœ"
            Height          =   375
            Index           =   11
            Left            =   1545
            RightToLeft     =   -1  'True
            TabIndex        =   551
            Top             =   5205
            Width           =   1185
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "„»·€ «·÷—Ì»…"
            Height          =   300
            Index           =   34
            Left            =   1815
            RightToLeft     =   -1  'True
            TabIndex        =   550
            Top             =   4080
            Width           =   915
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "‰”»… «·÷—Ì»…"
            Height          =   300
            Index           =   35
            Left            =   1575
            RightToLeft     =   -1  'True
            TabIndex        =   549
            Top             =   3585
            Width           =   1155
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·«Ã„«·Ï »⁄œ «·÷—Ì»…"
            Height          =   300
            Index           =   36
            Left            =   1155
            RightToLeft     =   -1  'True
            TabIndex        =   548
            Top             =   4650
            Width           =   1575
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "”‰œ «·’—›"
            Height          =   255
            Index           =   41
            Left            =   8670
            RightToLeft     =   -1  'True
            TabIndex        =   143
            Top             =   930
            Width           =   795
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "„‰"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   69
            Left            =   8475
            RightToLeft     =   -1  'True
            TabIndex        =   142
            Top             =   1770
            Width           =   600
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·Ï"
            Height          =   210
            Index           =   70
            Left            =   7095
            RightToLeft     =   -1  'True
            TabIndex        =   141
            Top             =   1815
            Width           =   195
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " ·Ì›Ê‰"
            Height          =   255
            Index           =   84
            Left            =   10845
            TabIndex        =   140
            Top             =   1500
            Width           =   990
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄„Ì· ÃœÌœ"
            Height          =   510
            Index           =   76
            Left            =   8670
            RightToLeft     =   -1  'True
            TabIndex        =   139
            Top             =   1500
            Width           =   795
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "›« Ê—… «·„»Ì⁄« "
            Height          =   480
            Index           =   32
            Left            =   11040
            TabIndex        =   138
            Top             =   930
            Width           =   990
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000C&
            BackStyle       =   0  'Transparent
            Caption         =   "«· ”·Ì„ «·›⁄·Ì"
            Height          =   255
            Index           =   7
            Left            =   5715
            RightToLeft     =   -1  'True
            TabIndex        =   137
            Top             =   2355
            Width           =   1185
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000C&
            BackStyle       =   0  'Transparent
            Caption         =   "Ã«Â“ ·· ”·Ì„"
            Height          =   255
            Index           =   6
            Left            =   6900
            RightToLeft     =   -1  'True
            TabIndex        =   136
            Top             =   2355
            Width           =   1185
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000C&
            BackStyle       =   0  'Transparent
            Caption         =   " «—ÌŒ «· ”·Ì„"
            Height          =   255
            Index           =   5
            Left            =   8475
            RightToLeft     =   -1  'True
            TabIndex        =   135
            Top             =   2355
            Width           =   795
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000C&
            BackStyle       =   0  'Transparent
            Caption         =   "Ã«Â“ ··»—Ê›…"
            Height          =   255
            Index           =   2
            Left            =   9660
            RightToLeft     =   -1  'True
            TabIndex        =   134
            Top             =   2355
            Width           =   990
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000C&
            BackStyle       =   0  'Transparent
            Caption         =   " «—ÌŒ «·»—Ê›…"
            Height          =   255
            Index           =   1
            Left            =   10845
            RightToLeft     =   -1  'True
            TabIndex        =   133
            Top             =   2355
            Width           =   990
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000C&
            BackStyle       =   0  'Transparent
            Caption         =   " «—ÌŒ «·«” ·«„"
            Height          =   255
            Index           =   0
            Left            =   12030
            RightToLeft     =   -1  'True
            TabIndex        =   132
            Top             =   2355
            Width           =   975
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·⁄„Ì· «·Õ«·Ì"
            Height          =   510
            Index           =   24
            Left            =   15975
            TabIndex        =   131
            Top             =   1500
            Width           =   975
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·›—⁄"
            Height          =   360
            Index           =   15
            Left            =   6315
            TabIndex        =   130
            Top             =   930
            Width           =   780
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «·”‰œ"
            Height          =   360
            Index           =   4
            Left            =   13605
            TabIndex        =   129
            Top             =   930
            Width           =   585
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”Ã· «·Õ«·Ì:"
            Height          =   270
            Index           =   7
            Left            =   5325
            RightToLeft     =   -1  'True
            TabIndex        =   128
            Top             =   9270
            Width           =   990
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «·”Ã·« :"
            Height          =   270
            Index           =   6
            Left            =   3345
            RightToLeft     =   -1  'True
            TabIndex        =   127
            Top             =   9270
            Width           =   1185
         End
         Begin VB.Label LabCurr_Rec 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   225
            Index           =   3
            Left            =   4530
            RightToLeft     =   -1  'True
            TabIndex        =   126
            Top             =   9270
            Width           =   600
         End
         Begin VB.Label LabCount_Rec 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   225
            Index           =   3
            Left            =   2760
            RightToLeft     =   -1  'True
            TabIndex        =   125
            Top             =   9345
            Width           =   585
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Õ—— »Ê«”ÿ…  "
            Height          =   285
            Index           =   8
            Left            =   17355
            TabIndex        =   124
            Top             =   9420
            Width           =   975
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "%"
            Height          =   225
            Index           =   3
            Left            =   3555
            TabIndex        =   123
            Top             =   5340
            Width           =   780
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "—ﬁ„ «·ﬁÌœ"
            Height          =   405
            Index           =   14
            Left            =   7890
            RightToLeft     =   -1  'True
            TabIndex        =   122
            Top             =   8805
            Width           =   780
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "—ﬁ„ «·ﬁÿ⁄…"
            Height          =   360
            Index           =   1
            Left            =   15975
            TabIndex        =   121
            Top             =   930
            Width           =   975
         End
      End
      Begin C1SizerLibCtl.C1Elastic ELe 
         Height          =   9750
         Index           =   6
         Left            =   20355
         TabIndex        =   144
         TabStop         =   0   'False
         Top             =   45
         Width           =   19320
         _cx             =   34078
         _cy             =   17198
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
         Begin VB.TextBox TxtSerial1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   345
            Index           =   4
            Left            =   8865
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   152
            Top             =   1365
            Width           =   1380
         End
         Begin VB.Frame Fra_Header 
            BackColor       =   &H00E2E9E9&
            BorderStyle     =   0  'None
            Height          =   630
            Index           =   3
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   145
            Top             =   0
            Width           =   19320
            Begin VB.TextBox TxtModFlg2 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H0000FF00&
               Enabled         =   0   'False
               Height          =   285
               Index           =   4
               Left            =   2580
               RightToLeft     =   -1  'True
               TabIndex        =   146
               Text            =   "modflag"
               Top             =   90
               Visible         =   0   'False
               Width           =   465
            End
            Begin MSComctlLib.ImageList GrdImageList2 
               Index           =   4
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
                     Picture         =   "dean.frx":2C407
                     Key             =   "CompanyName"
                  EndProperty
                  BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "dean.frx":2C7A1
                     Key             =   "Ser"
                  EndProperty
                  BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "dean.frx":2CB3B
                     Key             =   "Vac_Name"
                  EndProperty
                  BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "dean.frx":2CED5
                     Key             =   "ShareCount"
                  EndProperty
                  BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "dean.frx":2D26F
                     Key             =   "Dis_Count"
                  EndProperty
                  BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "dean.frx":2D609
                     Key             =   "Bouns"
                  EndProperty
                  BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "dean.frx":2D9A3
                     Key             =   "SharesValue"
                  EndProperty
                  BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "dean.frx":2DF3D
                     Key             =   "BuyValue"
                  EndProperty
               EndProperty
            End
            Begin ImpulseButton.ISButton btn_Last 
               Height          =   315
               Index           =   4
               Left            =   90
               TabIndex        =   147
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
               ButtonImage     =   "dean.frx":2E2D7
               ColorButton     =   14871017
               AcclimateGrayTones=   -1  'True
               DrawFocusRectangle=   0   'False
               DisabledImageExtraction=   0
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btn_Next 
               Height          =   315
               Index           =   4
               Left            =   555
               TabIndex        =   148
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
               ButtonImage     =   "dean.frx":2E671
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btn_Previous 
               Height          =   315
               Index           =   4
               Left            =   1155
               TabIndex        =   149
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
               ButtonImage     =   "dean.frx":2EA0B
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btn_First 
               Height          =   315
               Index           =   4
               Left            =   1620
               TabIndex        =   150
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
               ButtonImage     =   "dean.frx":2EDA5
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin VB.Image ImgFavorites 
               Height          =   390
               Index           =   1
               Left            =   7560
               Picture         =   "dean.frx":2F13F
               Stretch         =   -1  'True
               Top             =   0
               Width           =   525
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   " ”ÃÌ· «·«‰ «ÃÌ… «·ÌÊ„Ì… ··„ÊŸ›« "
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   24
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   735
               Index           =   15
               Left            =   3030
               RightToLeft     =   -1  'True
               TabIndex        =   151
               Top             =   180
               Width           =   8880
            End
         End
         Begin VSFlex8UCtl.VSFlexGrid Fg4 
            Height          =   3870
            Left            =   195
            TabIndex        =   153
            Top             =   1770
            Width           =   11430
            _cx             =   20161
            _cy             =   6826
            Appearance      =   2
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   0
            BackColor       =   16777088
            ForeColor       =   -2147483640
            BackColorFixed  =   16777215
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   12632256
            BackColorAlternate=   16777215
            GridColor       =   16777215
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
            Cols            =   18
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   600
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"dean.frx":32DA7
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
         Begin MSDataListLib.DataCombo DCboUserName 
            Height          =   315
            Index           =   4
            Left            =   15765
            TabIndex        =   154
            Top             =   9480
            Width           =   2565
            _ExtentX        =   4524
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin ImpulseButton.ISButton btn_New 
            Height          =   405
            Index           =   4
            Left            =   9660
            TabIndex        =   155
            Top             =   6585
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   714
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "ÃœÌœ"
            BackColor       =   14871017
            FontSize        =   13.5
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "dean.frx":33029
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Save 
            Height          =   405
            Index           =   4
            Left            =   7095
            TabIndex        =   156
            Top             =   6585
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   714
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "Õ›Ÿ"
            BackColor       =   14871017
            FontSize        =   13.5
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "dean.frx":333C3
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Modify 
            Height          =   405
            Index           =   4
            Left            =   8280
            TabIndex        =   157
            Top             =   6585
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   714
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   " ⁄œÌ·"
            BackColor       =   14871017
            FontSize        =   13.5
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "dean.frx":3375D
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton Btn_Undo 
            Height          =   405
            Index           =   4
            Left            =   5910
            TabIndex        =   158
            Top             =   6585
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   714
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   " —«Ã⁄"
            BackColor       =   14871017
            FontSize        =   13.5
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "dean.frx":33AF7
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Delete 
            Height          =   405
            Index           =   4
            Left            =   4935
            TabIndex        =   159
            Top             =   6585
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   714
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "Õ–›"
            BackColor       =   14871017
            FontSize        =   13.5
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "dean.frx":33E91
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Cancel 
            Height          =   570
            Index           =   4
            Left            =   390
            TabIndex        =   160
            Top             =   6540
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   1005
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "Œ—ÊÃ"
            BackColor       =   14871017
            FontSize        =   13.5
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "dean.frx":3442B
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton Btn_Print 
            Height          =   405
            Index           =   4
            Left            =   3345
            TabIndex        =   161
            TabStop         =   0   'False
            ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
            Top             =   6585
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   714
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "ÿ»«⁄… "
            BackColor       =   14871017
            FontSize        =   13.5
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "dean.frx":347C5
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btn_Query 
            Height          =   525
            Index           =   4
            Left            =   1770
            TabIndex        =   162
            TabStop         =   0   'False
            ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
            Top             =   6585
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   926
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "»ÕÀ"
            BackColor       =   14871017
            FontSize        =   13.5
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "dean.frx":3B027
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton Cmd_DeleteRow 
            Height          =   285
            Index           =   4
            Left            =   1575
            TabIndex        =   163
            Top             =   5715
            Width           =   1980
            _ExtentX        =   3493
            _ExtentY        =   503
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   " Õ–› ”ÿ—"
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
            ButtonImage     =   "dean.frx":3B3C1
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton Cmd_DeleteAll 
            Height          =   300
            Index           =   4
            Left            =   0
            TabIndex        =   164
            Top             =   5700
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   529
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   " Õ–› «·ﬂ·"
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
            ButtonImage     =   "dean.frx":3B95B
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin MSComCtl2.DTPicker XPDtbTrans 
            Height          =   345
            Index           =   4
            Left            =   6105
            TabIndex        =   165
            Top             =   1380
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   609
            _Version        =   393216
            CalendarBackColor=   16777152
            Format          =   241303553
            CurrentDate     =   38784
         End
         Begin ImpulseButton.ISButton ISButton1 
            Height          =   345
            Left            =   4530
            TabIndex        =   166
            ToolTipText     =   "«÷«›… «·»Ì«‰«  «·Ï «·œ« «"
            Top             =   1365
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   609
            Caption         =   "«÷«›…  ”ÿ—"
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
            ButtonImage     =   "dean.frx":3BEF5
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            LowerToggledContent=   0   'False
         End
         Begin MSDataListLib.DataCombo Dcbranch 
            Bindings        =   "dean.frx":42757
            Height          =   315
            Index           =   4
            Left            =   195
            TabIndex        =   167
            Top             =   1455
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            BackColor       =   16777152
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
            Caption         =   "«·›—⁄"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   33
            Left            =   3345
            TabIndex        =   175
            Top             =   1485
            Width           =   795
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «·”‰œ"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   14
            Left            =   7485
            TabIndex        =   174
            Top             =   1440
            Width           =   1380
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "—ﬁ„ «·”‰œ"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   3
            Left            =   10050
            TabIndex        =   173
            Top             =   1440
            Width           =   1575
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”Ã· «·Õ«·Ì:"
            Height          =   225
            Index           =   9
            Left            =   2760
            RightToLeft     =   -1  'True
            TabIndex        =   172
            Top             =   9510
            Width           =   795
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «·”Ã·« :"
            Height          =   225
            Index           =   8
            Left            =   795
            RightToLeft     =   -1  'True
            TabIndex        =   171
            Top             =   9510
            Width           =   1170
         End
         Begin VB.Label LabCurr_Rec 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   210
            Index           =   4
            Left            =   1965
            RightToLeft     =   -1  'True
            TabIndex        =   170
            Top             =   9525
            Width           =   795
         End
         Begin VB.Label LabCount_Rec 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   210
            Index           =   4
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   169
            Top             =   9525
            Width           =   795
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Õ—— »Ê«”ÿ…  "
            Height          =   300
            Index           =   13
            Left            =   18525
            TabIndex        =   168
            Top             =   9510
            Width           =   600
         End
      End
      Begin C1SizerLibCtl.C1Elastic ELe 
         Height          =   9750
         Index           =   7
         Left            =   20655
         TabIndex        =   176
         TabStop         =   0   'False
         Top             =   45
         Width           =   19320
         _cx             =   34078
         _cy             =   17198
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
         Begin VB.Frame Frame1 
            BackColor       =   &H00E2E9E9&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   1800
            Index           =   5
            Left            =   390
            RightToLeft     =   -1  'True
            TabIndex        =   184
            Top             =   4305
            Width           =   17355
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
               Index           =   5
               Left            =   3195
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   188
               Top             =   960
               Width           =   2760
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
               Index           =   5
               Left            =   3195
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   187
               Top             =   645
               Width           =   2760
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
               Index           =   5
               Left            =   4830
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   186
               Top             =   270
               Width           =   1065
            End
            Begin VB.ComboBox Combo3 
               BackColor       =   &H80000018&
               Height          =   315
               ItemData        =   "dean.frx":4276C
               Left            =   2280
               List            =   "dean.frx":4277C
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   185
               Top             =   3150
               Visible         =   0   'False
               Width           =   1005
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «‰Ã·Ì“Ì"
               Height          =   285
               Index           =   3
               Left            =   6000
               RightToLeft     =   -1  'True
               TabIndex        =   191
               Top             =   1080
               Width           =   1500
            End
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ ⁄—»Ì"
               Height          =   285
               Index           =   4
               Left            =   6150
               RightToLeft     =   -1  'True
               TabIndex        =   190
               Top             =   720
               Width           =   1350
            End
            Begin VB.Label Label6 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ﬂÊœ "
               Height          =   195
               Index           =   1
               Left            =   6495
               RightToLeft     =   -1  'True
               TabIndex        =   189
               Top             =   390
               Width           =   990
            End
         End
         Begin VB.Frame Fra_Header 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   675
            Index           =   5
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   177
            Top             =   0
            Width           =   19320
            Begin VB.TextBox TxtModFlg2 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H0000FF00&
               Enabled         =   0   'False
               Height          =   285
               Index           =   5
               Left            =   2580
               RightToLeft     =   -1  'True
               TabIndex        =   178
               Text            =   "modflag"
               Top             =   90
               Visible         =   0   'False
               Width           =   465
            End
            Begin MSComctlLib.ImageList GrdImageList2 
               Index           =   5
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
                     Picture         =   "dean.frx":42795
                     Key             =   "CompanyName"
                  EndProperty
                  BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "dean.frx":42B2F
                     Key             =   "Ser"
                  EndProperty
                  BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "dean.frx":42EC9
                     Key             =   "Vac_Name"
                  EndProperty
                  BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "dean.frx":43263
                     Key             =   "ShareCount"
                  EndProperty
                  BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "dean.frx":435FD
                     Key             =   "Dis_Count"
                  EndProperty
                  BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "dean.frx":43997
                     Key             =   "Bouns"
                  EndProperty
                  BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "dean.frx":43D31
                     Key             =   "SharesValue"
                  EndProperty
                  BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "dean.frx":442CB
                     Key             =   "BuyValue"
                  EndProperty
               EndProperty
            End
            Begin ImpulseButton.ISButton btn_Last 
               Height          =   315
               Index           =   5
               Left            =   90
               TabIndex        =   179
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
               ButtonImage     =   "dean.frx":44665
               ColorButton     =   14871017
               AcclimateGrayTones=   -1  'True
               DrawFocusRectangle=   0   'False
               DisabledImageExtraction=   0
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btn_Next 
               Height          =   315
               Index           =   5
               Left            =   555
               TabIndex        =   180
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
               ButtonImage     =   "dean.frx":449FF
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btn_Previous 
               Height          =   315
               Index           =   5
               Left            =   1155
               TabIndex        =   181
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
               ButtonImage     =   "dean.frx":44D99
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btn_First 
               Height          =   315
               Index           =   5
               Left            =   1620
               TabIndex        =   182
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
               ButtonImage     =   "dean.frx":45133
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«‰Ê«⁄ «·ÕÃ“"
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
               Index           =   16
               Left            =   6690
               RightToLeft     =   -1  'True
               TabIndex        =   183
               Top             =   60
               Width           =   2640
            End
         End
         Begin ImpulseButton.ISButton btn_New 
            Height          =   435
            Index           =   5
            Left            =   7290
            TabIndex        =   192
            Top             =   6480
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   767
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
            ButtonImage     =   "dean.frx":454CD
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Save 
            Height          =   435
            Index           =   5
            Left            =   5520
            TabIndex        =   193
            Top             =   6480
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   767
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
            ButtonImage     =   "dean.frx":45867
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Modify 
            Height          =   435
            Index           =   5
            Left            =   6705
            TabIndex        =   194
            Top             =   6480
            Width           =   585
            _ExtentX        =   1032
            _ExtentY        =   767
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
            ButtonImage     =   "dean.frx":45C01
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton Btn_Undo 
            Height          =   435
            Index           =   5
            Left            =   4530
            TabIndex        =   195
            Top             =   6480
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   767
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
            ButtonImage     =   "dean.frx":45F9B
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Delete 
            Height          =   435
            Index           =   5
            Left            =   3945
            TabIndex        =   196
            Top             =   6480
            Width           =   585
            _ExtentX        =   1032
            _ExtentY        =   767
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
            ButtonImage     =   "dean.frx":46335
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Cancel 
            Height          =   420
            Index           =   5
            Left            =   585
            TabIndex        =   197
            Top             =   6480
            Width           =   795
            _ExtentX        =   1402
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
            ButtonImage     =   "dean.frx":468CF
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton Btn_Print 
            Height          =   375
            Index           =   5
            Left            =   2565
            TabIndex        =   198
            TabStop         =   0   'False
            ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
            Top             =   6510
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   661
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
            ButtonImage     =   "dean.frx":46C69
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btn_Query 
            Height          =   435
            Index           =   5
            Left            =   1770
            TabIndex        =   199
            TabStop         =   0   'False
            ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
            Top             =   6480
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   767
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
            ButtonImage     =   "dean.frx":4D4CB
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin VSFlex8Ctl.VSFlexGrid Grid5 
            Height          =   3510
            Left            =   390
            TabIndex        =   200
            Top             =   720
            Width           =   7890
            _cx             =   13917
            _cy             =   6191
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
            FormatString    =   $"dean.frx":4D865
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
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”Ã· «·Õ«·Ì:"
            Height          =   195
            Index           =   11
            Left            =   5325
            RightToLeft     =   -1  'True
            TabIndex        =   204
            Top             =   6225
            Width           =   1185
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «·”Ã·« :"
            Height          =   210
            Index           =   10
            Left            =   1380
            RightToLeft     =   -1  'True
            TabIndex        =   203
            Top             =   6270
            Width           =   795
         End
         Begin VB.Label LabCurr_Rec 
            BackColor       =   &H00E2E9E9&
            Height          =   225
            Index           =   5
            Left            =   3750
            RightToLeft     =   -1  'True
            TabIndex        =   202
            Top             =   6195
            Width           =   1185
         End
         Begin VB.Label LabCount_Rec 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   225
            Index           =   5
            Left            =   585
            RightToLeft     =   -1  'True
            TabIndex        =   201
            Top             =   6225
            Width           =   405
         End
      End
      Begin C1SizerLibCtl.C1Elastic ELe 
         Height          =   9750
         Index           =   8
         Left            =   20955
         TabIndex        =   205
         TabStop         =   0   'False
         Top             =   45
         Width           =   19320
         _cx             =   34078
         _cy             =   17198
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
         Begin VB.Timer Timer3 
            Interval        =   10000
            Left            =   4440
            Top             =   780
         End
         Begin VB.Frame Fra_Header 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   630
            Index           =   4
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   206
            Top             =   0
            Width           =   19320
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "⁄—÷ «·ÕÃÊ“« "
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
               Index           =   18
               Left            =   7470
               RightToLeft     =   -1  'True
               TabIndex        =   207
               Top             =   0
               Width           =   2640
            End
            Begin VB.Image ImgFavorites 
               Height          =   390
               Index           =   2
               Left            =   7560
               Picture         =   "dean.frx":4D8F4
               Stretch         =   -1  'True
               Top             =   0
               Width           =   525
            End
         End
         Begin VSFlex8UCtl.VSFlexGrid FG6 
            Height          =   7365
            Left            =   0
            TabIndex        =   208
            Top             =   2175
            Width           =   19125
            _cx             =   33734
            _cy             =   12991
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
            Cols            =   19
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"dean.frx":5155C
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
         Begin ImpulseButton.ISButton ISButton2 
            Height          =   585
            Left            =   795
            TabIndex        =   209
            ToolTipText     =   "«÷«›… «·»Ì«‰«  «·Ï «·œ« «"
            Top             =   1320
            Width           =   2160
            _ExtentX        =   3810
            _ExtentY        =   1032
            Caption         =   "«œ—«Ã ÕÃÊ“«  «·ÌÊ„"
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
            ButtonImage     =   "dean.frx":51832
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageExtraction=   0
            LowerToggledContent=   0   'False
         End
         Begin MSComCtl2.DTPicker XPDtbBill 
            Height          =   360
            Index           =   0
            Left            =   16170
            TabIndex        =   210
            Top             =   1395
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   635
            _Version        =   393216
            Format          =   244514817
            CurrentDate     =   38784
         End
         Begin MSDataListLib.DataCombo CmbEmp 
            Height          =   315
            Left            =   9855
            TabIndex        =   530
            Top             =   1440
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777152
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
         Begin MSComCtl2.DTPicker XPDtbBill 
            Height          =   345
            Index           =   1
            Left            =   13800
            TabIndex        =   532
            Top             =   1365
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   609
            _Version        =   393216
            Format          =   244514817
            CurrentDate     =   38784
         End
         Begin MSDataListLib.DataCombo cmbCustomer 
            Height          =   315
            Left            =   6900
            TabIndex        =   534
            Top             =   1440
            Width           =   1965
            _ExtentX        =   3466
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777152
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
         Begin MSDataListLib.DataCombo cmbItems 
            Height          =   315
            Left            =   3555
            TabIndex        =   606
            Top             =   1440
            Width           =   1965
            _ExtentX        =   3466
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777152
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
            Caption         =   "«·Œœ„…"
            Height          =   240
            Index           =   28
            Left            =   5715
            RightToLeft     =   -1  'True
            TabIndex        =   607
            Top             =   1440
            Width           =   600
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·⁄„Ì·…"
            Height          =   240
            Index           =   27
            Left            =   8865
            RightToLeft     =   -1  'True
            TabIndex        =   535
            Top             =   1440
            Width           =   795
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·Ï  «—ÌŒ"
            Height          =   240
            Index           =   26
            Left            =   14790
            RightToLeft     =   -1  'True
            TabIndex        =   533
            Top             =   1485
            Width           =   1185
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„ÊŸ›…"
            Height          =   360
            Index           =   12
            Left            =   12030
            RightToLeft     =   -1  'True
            TabIndex        =   531
            Top             =   1440
            Width           =   780
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„‰  «—ÌŒ"
            Height          =   390
            Index           =   40
            Left            =   17355
            RightToLeft     =   -1  'True
            TabIndex        =   211
            Top             =   1410
            Width           =   780
         End
      End
      Begin C1SizerLibCtl.C1Elastic ELe 
         Height          =   9750
         Index           =   9
         Left            =   21255
         TabIndex        =   212
         TabStop         =   0   'False
         Top             =   45
         Width           =   19320
         _cx             =   34078
         _cy             =   17198
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
         Begin VB.CommandButton cmdReloadList 
            Caption         =   "«·€«¡ «·„Õœœ"
            Height          =   225
            Index           =   1
            Left            =   8475
            RightToLeft     =   -1  'True
            TabIndex        =   500
            Top             =   3930
            Width           =   1980
         End
         Begin VB.Frame Frame4 
            Height          =   2550
            Index           =   1
            Left            =   390
            RightToLeft     =   -1  'True
            TabIndex        =   221
            Top             =   1245
            Width           =   12030
            Begin VB.CommandButton cmdInsertEmpItems 
               Caption         =   "«œ—«Ã"
               Height          =   945
               Left            =   7980
               RightToLeft     =   -1  'True
               TabIndex        =   226
               Top             =   2760
               Width           =   585
            End
            Begin VB.ListBox ListProductLineSelected 
               BackColor       =   &H0080FFFF&
               Height          =   3375
               ItemData        =   "dean.frx":58094
               Left            =   8550
               List            =   "dean.frx":5809B
               MultiSelect     =   1  'Simple
               RightToLeft     =   -1  'True
               TabIndex        =   225
               Top             =   360
               Width           =   3765
            End
            Begin VB.ListBox ListProductLineAll 
               Height          =   3375
               ItemData        =   "dean.frx":580B2
               Left            =   13770
               List            =   "dean.frx":580B9
               MultiSelect     =   1  'Simple
               RightToLeft     =   -1  'True
               TabIndex        =   224
               Top             =   330
               Width           =   3825
            End
            Begin VB.ListBox ListGroupSelected 
               BackColor       =   &H0080FFFF&
               Height          =   3375
               ItemData        =   "dean.frx":580CB
               Left            =   240
               List            =   "dean.frx":580D2
               RightToLeft     =   -1  'True
               TabIndex        =   223
               Top             =   390
               Width           =   3675
            End
            Begin VB.ListBox ListGroupAll 
               Height          =   3375
               ItemData        =   "dean.frx":580E9
               Left            =   4440
               List            =   "dean.frx":580F0
               MultiSelect     =   1  'Simple
               RightToLeft     =   -1  'True
               TabIndex        =   222
               Top             =   390
               Width           =   3225
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               Caption         =   "ﬂ· «·„ÊŸ›Ì‰"
               Height          =   255
               Index           =   59
               Left            =   15510
               RightToLeft     =   -1  'True
               TabIndex        =   238
               Top             =   30
               Width           =   1335
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               Caption         =   "«·„ÊŸ›Ì‰ «·„Õœœ…"
               Height          =   255
               Index           =   58
               Left            =   9270
               RightToLeft     =   -1  'True
               TabIndex        =   237
               Top             =   60
               Width           =   1335
            End
            Begin VB.Label Label31 
               Alignment       =   2  'Center
               Caption         =   "<"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   255
               Left            =   12300
               RightToLeft     =   -1  'True
               TabIndex        =   236
               Top             =   1470
               Width           =   495
            End
            Begin VB.Label Label30 
               Alignment       =   2  'Center
               Caption         =   "<<"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   255
               Left            =   12300
               RightToLeft     =   -1  'True
               TabIndex        =   235
               Top             =   1230
               Width           =   495
            End
            Begin VB.Label Label29 
               Alignment       =   2  'Center
               Caption         =   ">>"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Left            =   12300
               RightToLeft     =   -1  'True
               TabIndex        =   234
               Top             =   870
               Width           =   495
            End
            Begin VB.Label Label28 
               Alignment       =   2  'Center
               Caption         =   ">"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Left            =   12300
               RightToLeft     =   -1  'True
               TabIndex        =   233
               Top             =   630
               Width           =   495
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               Caption         =   "«·«’‰«› «·„Õœœ…"
               Height          =   255
               Index           =   53
               Left            =   510
               RightToLeft     =   -1  'True
               TabIndex        =   232
               Top             =   0
               Width           =   1335
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               Caption         =   "ﬂ· «·«’‰«›"
               Height          =   255
               Index           =   57
               Left            =   5850
               RightToLeft     =   -1  'True
               TabIndex        =   231
               Top             =   60
               Width           =   1335
            End
            Begin VB.Label Label53 
               Alignment       =   2  'Center
               Caption         =   "<"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   255
               Left            =   3900
               RightToLeft     =   -1  'True
               TabIndex        =   230
               Top             =   1470
               Width           =   495
            End
            Begin VB.Label Label63 
               Alignment       =   2  'Center
               Caption         =   "<<"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   255
               Left            =   3900
               RightToLeft     =   -1  'True
               TabIndex        =   229
               Top             =   1230
               Width           =   495
            End
            Begin VB.Label Label7 
               Alignment       =   2  'Center
               Caption         =   ">>"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Left            =   3900
               RightToLeft     =   -1  'True
               TabIndex        =   228
               Top             =   870
               Width           =   495
            End
            Begin VB.Label Label8 
               Alignment       =   2  'Center
               Caption         =   ">"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Left            =   3900
               RightToLeft     =   -1  'True
               TabIndex        =   227
               Top             =   630
               Width           =   495
            End
         End
         Begin VB.TextBox TxtSerial1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   360
            Index           =   7
            Left            =   8865
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   220
            Top             =   855
            Width           =   1980
         End
         Begin VB.Frame Fra_Header 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   630
            Index           =   6
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   213
            Top             =   0
            Width           =   19320
            Begin VB.TextBox TxtModFlg2 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H0000FF00&
               Enabled         =   0   'False
               Height          =   285
               Index           =   7
               Left            =   2580
               RightToLeft     =   -1  'True
               TabIndex        =   214
               Text            =   "modflag"
               Top             =   90
               Visible         =   0   'False
               Width           =   465
            End
            Begin MSComctlLib.ImageList GrdImageList2 
               Index           =   7
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
                     Picture         =   "dean.frx":58102
                     Key             =   "CompanyName"
                  EndProperty
                  BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "dean.frx":5849C
                     Key             =   "Ser"
                  EndProperty
                  BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "dean.frx":58836
                     Key             =   "Vac_Name"
                  EndProperty
                  BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "dean.frx":58BD0
                     Key             =   "ShareCount"
                  EndProperty
                  BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "dean.frx":58F6A
                     Key             =   "Dis_Count"
                  EndProperty
                  BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "dean.frx":59304
                     Key             =   "Bouns"
                  EndProperty
                  BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "dean.frx":5969E
                     Key             =   "SharesValue"
                  EndProperty
                  BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "dean.frx":59C38
                     Key             =   "BuyValue"
                  EndProperty
               EndProperty
            End
            Begin ImpulseButton.ISButton btn_Last 
               Height          =   315
               Index           =   7
               Left            =   90
               TabIndex        =   215
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
               ButtonImage     =   "dean.frx":59FD2
               ColorButton     =   14871017
               AcclimateGrayTones=   -1  'True
               DrawFocusRectangle=   0   'False
               DisabledImageExtraction=   0
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btn_Next 
               Height          =   315
               Index           =   7
               Left            =   555
               TabIndex        =   216
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
               ButtonImage     =   "dean.frx":5A36C
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btn_Previous 
               Height          =   315
               Index           =   7
               Left            =   1155
               TabIndex        =   217
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
               ButtonImage     =   "dean.frx":5A706
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btn_First 
               Height          =   315
               Index           =   7
               Left            =   1620
               TabIndex        =   218
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
               ButtonImage     =   "dean.frx":5AAA0
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin VB.Image ImgFavorites 
               Height          =   390
               Index           =   3
               Left            =   11100
               Picture         =   "dean.frx":5AE3A
               Stretch         =   -1  'True
               Top             =   30
               Width           =   525
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "—»ÿ «·„ÊŸ›Ì‰ »«·Œœ„«  Ê«·«’‰«›"
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
               Index           =   20
               Left            =   6000
               RightToLeft     =   -1  'True
               TabIndex        =   219
               Top             =   90
               Width           =   4320
            End
         End
         Begin VSFlex8UCtl.VSFlexGrid FG7 
            Height          =   1755
            Left            =   195
            TabIndex        =   239
            Top             =   4170
            Width           =   12225
            _cx             =   21564
            _cy             =   3096
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
            FormatString    =   $"dean.frx":5EAA2
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
         Begin MSDataListLib.DataCombo DCboUserName 
            Height          =   315
            Index           =   7
            Left            =   8085
            TabIndex        =   240
            Top             =   6060
            Width           =   2760
            _ExtentX        =   4868
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSComCtl2.DTPicker XPDtbTrans 
            Height          =   345
            Index           =   7
            Left            =   5715
            TabIndex        =   241
            Top             =   855
            Width           =   1770
            _ExtentX        =   3122
            _ExtentY        =   609
            _Version        =   393216
            Format          =   235208705
            CurrentDate     =   38784
         End
         Begin MSDataListLib.DataCombo Dcbranch 
            Bindings        =   "dean.frx":5EB5B
            Height          =   315
            Index           =   7
            Left            =   195
            TabIndex        =   242
            Top             =   855
            Width           =   2955
            _ExtentX        =   5212
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
         Begin ImpulseButton.ISButton btn_New 
            Height          =   330
            Index           =   7
            Left            =   11430
            TabIndex        =   578
            Top             =   7005
            Width           =   990
            _ExtentX        =   1746
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
            ButtonImage     =   "dean.frx":5EB70
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Save 
            Height          =   345
            Index           =   7
            Left            =   9465
            TabIndex        =   579
            Top             =   6990
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   609
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
            ButtonImage     =   "dean.frx":5EF0A
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Modify 
            Height          =   225
            Index           =   7
            Left            =   10455
            TabIndex        =   580
            Top             =   7005
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   397
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
            ButtonImage     =   "dean.frx":5F2A4
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton Btn_Undo 
            Height          =   225
            Index           =   7
            Left            =   8865
            TabIndex        =   581
            Top             =   7005
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   397
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
            ButtonImage     =   "dean.frx":5F63E
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Delete 
            Height          =   345
            Index           =   7
            Left            =   8085
            TabIndex        =   582
            Top             =   6990
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   609
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
            ButtonImage     =   "dean.frx":5F9D8
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Cancel 
            Height          =   345
            Index           =   7
            Left            =   4530
            TabIndex        =   583
            Top             =   6945
            Width           =   600
            _ExtentX        =   1058
            _ExtentY        =   609
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
            ButtonImage     =   "dean.frx":5FF72
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton Btn_Print 
            Height          =   390
            Index           =   7
            Left            =   6705
            TabIndex        =   584
            TabStop         =   0   'False
            ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
            Top             =   6930
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   688
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
            ButtonImage     =   "dean.frx":6030C
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btn_Query 
            Height          =   420
            Index           =   7
            Left            =   5130
            TabIndex        =   585
            TabStop         =   0   'False
            ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
            Top             =   6900
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   741
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
            ButtonImage     =   "dean.frx":66B6E
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton Cmd_DeleteRow 
            Height          =   285
            Index           =   7
            Left            =   1185
            TabIndex        =   586
            Top             =   6165
            Width           =   1965
            _ExtentX        =   3466
            _ExtentY        =   503
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   " Õ–› ”ÿ—"
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
            ButtonImage     =   "dean.frx":66F08
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton Cmd_DeleteAll 
            Height          =   300
            Index           =   7
            Left            =   -390
            TabIndex        =   587
            Top             =   6150
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   529
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   " Õ–› «·ﬂ·"
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
            ButtonImage     =   "dean.frx":674A2
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin VB.Label LabCount_Rec 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   225
            Index           =   7
            Left            =   1770
            RightToLeft     =   -1  'True
            TabIndex        =   591
            Top             =   6630
            Width           =   795
         End
         Begin VB.Label LabCurr_Rec 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   225
            Index           =   7
            Left            =   3750
            RightToLeft     =   -1  'True
            TabIndex        =   590
            Top             =   6630
            Width           =   780
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «·”Ã·« :"
            Height          =   225
            Index           =   14
            Left            =   2565
            RightToLeft     =   -1  'True
            TabIndex        =   589
            Top             =   6615
            Width           =   1185
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”Ã· «·Õ«·Ì:"
            Height          =   225
            Index           =   15
            Left            =   4530
            RightToLeft     =   -1  'True
            TabIndex        =   588
            Top             =   6615
            Width           =   795
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·›—⁄"
            Height          =   270
            Index           =   39
            Left            =   3345
            TabIndex        =   246
            Top             =   885
            Width           =   795
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «·”‰œ"
            Height          =   300
            Index           =   38
            Left            =   7485
            TabIndex        =   245
            Top             =   900
            Width           =   1185
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "—ﬁ„ «·”‰œ"
            Height          =   300
            Index           =   7
            Left            =   10845
            TabIndex        =   244
            Top             =   900
            Width           =   1380
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Õ—— »Ê«”ÿ…  "
            Height          =   315
            Index           =   37
            Left            =   11625
            TabIndex        =   243
            Top             =   6015
            Width           =   990
         End
      End
      Begin C1SizerLibCtl.C1Elastic ELe 
         Height          =   9750
         Index           =   10
         Left            =   21555
         TabIndex        =   247
         TabStop         =   0   'False
         Top             =   45
         Width           =   19320
         _cx             =   34078
         _cy             =   17198
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
         Begin VB.Frame Frame1 
            BackColor       =   &H00E2E9E9&
            BorderStyle     =   0  'None
            Height          =   2460
            Index           =   8
            Left            =   990
            RightToLeft     =   -1  'True
            TabIndex        =   255
            Top             =   4245
            Width           =   11820
            Begin VB.CheckBox chkIsBoardNo 
               Alignment       =   1  'Right Justify
               Caption         =   "«Œ›«¡ "
               Height          =   225
               Index           =   1
               Left            =   2760
               RightToLeft     =   -1  'True
               TabIndex        =   524
               Top             =   1440
               Width           =   1665
            End
            Begin VB.CheckBox chkIsBoardNo 
               Alignment       =   1  'Right Justify
               Caption         =   "—ﬁ„ «··ÊÕ… «·“«„Ï"
               Height          =   225
               Index           =   0
               Left            =   5340
               RightToLeft     =   -1  'True
               TabIndex        =   260
               Top             =   1500
               Width           =   1665
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
               Index           =   8
               Left            =   4380
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   259
               Top             =   420
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
               Index           =   8
               Left            =   2745
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   258
               Top             =   795
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
               Index           =   8
               Left            =   2745
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   257
               Top             =   1110
               Width           =   2760
            End
            Begin VB.CommandButton Command7 
               Caption         =   "«Œ — «··Ê‰"
               Height          =   465
               Left            =   2490
               TabIndex        =   256
               Top             =   1860
               Width           =   1005
            End
            Begin VB.Label LabCurr_Rec 
               BackColor       =   &H00E2E9E9&
               Height          =   435
               Index           =   8
               Left            =   2385
               RightToLeft     =   -1  'True
               TabIndex        =   595
               Top             =   120
               Width           =   1455
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄œœ «·”Ã·« :"
               Height          =   435
               Index           =   18
               Left            =   345
               RightToLeft     =   -1  'True
               TabIndex        =   594
               Top             =   120
               Width           =   1740
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·”Ã· «·Õ«·Ì:"
               Height          =   435
               Index           =   19
               Left            =   3840
               RightToLeft     =   -1  'True
               TabIndex        =   593
               Top             =   120
               Width           =   2055
            End
            Begin VB.Label LabCount_Rec 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Height          =   435
               Index           =   8
               Left            =   -240
               RightToLeft     =   -1  'True
               TabIndex        =   592
               Top             =   120
               Width           =   285
            End
            Begin VB.Label Label6 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ﬂÊœ "
               Height          =   195
               Index           =   4
               Left            =   5805
               RightToLeft     =   -1  'True
               TabIndex        =   265
               Top             =   540
               Width           =   990
            End
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ ⁄—»Ì"
               Height          =   285
               Index           =   10
               Left            =   5460
               RightToLeft     =   -1  'True
               TabIndex        =   264
               Top             =   870
               Width           =   1350
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «‰Ã·Ì“Ì"
               Height          =   285
               Index           =   5
               Left            =   5550
               RightToLeft     =   -1  'True
               TabIndex        =   263
               Top             =   1230
               Width           =   1500
            End
            Begin VB.Label lblServiceColor 
               Caption         =   " "
               Height          =   375
               Left            =   3690
               TabIndex        =   262
               Top             =   1860
               Width           =   1905
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "·Ê‰ «·›∆…"
               Height          =   315
               Index           =   91
               Left            =   5550
               RightToLeft     =   -1  'True
               TabIndex        =   261
               Top             =   1950
               Width           =   1470
            End
         End
         Begin VB.Frame Fra_Header 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   675
            Index           =   8
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   248
            Top             =   0
            Width           =   19320
            Begin VB.TextBox TxtModFlg2 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H0000FF00&
               Enabled         =   0   'False
               Height          =   285
               Index           =   8
               Left            =   2580
               RightToLeft     =   -1  'True
               TabIndex        =   249
               Text            =   "modflag"
               Top             =   90
               Visible         =   0   'False
               Width           =   465
            End
            Begin MSComctlLib.ImageList GrdImageList2 
               Index           =   9
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
                     Picture         =   "dean.frx":67A3C
                     Key             =   "CompanyName"
                  EndProperty
                  BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "dean.frx":67DD6
                     Key             =   "Ser"
                  EndProperty
                  BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "dean.frx":68170
                     Key             =   "Vac_Name"
                  EndProperty
                  BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "dean.frx":6850A
                     Key             =   "ShareCount"
                  EndProperty
                  BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "dean.frx":688A4
                     Key             =   "Dis_Count"
                  EndProperty
                  BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "dean.frx":68C3E
                     Key             =   "Bouns"
                  EndProperty
                  BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "dean.frx":68FD8
                     Key             =   "SharesValue"
                  EndProperty
                  BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "dean.frx":69572
                     Key             =   "BuyValue"
                  EndProperty
               EndProperty
            End
            Begin ImpulseButton.ISButton btn_Last 
               Height          =   315
               Index           =   8
               Left            =   90
               TabIndex        =   250
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
               ButtonImage     =   "dean.frx":6990C
               ColorButton     =   14871017
               AcclimateGrayTones=   -1  'True
               DrawFocusRectangle=   0   'False
               DisabledImageExtraction=   0
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btn_Next 
               Height          =   315
               Index           =   8
               Left            =   555
               TabIndex        =   251
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
               ButtonImage     =   "dean.frx":69CA6
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btn_Previous 
               Height          =   315
               Index           =   8
               Left            =   1155
               TabIndex        =   252
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
               ButtonImage     =   "dean.frx":6A040
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btn_First 
               Height          =   315
               Index           =   8
               Left            =   1620
               TabIndex        =   253
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
               ButtonImage     =   "dean.frx":6A3DA
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
               Index           =   4
               Left            =   7950
               RightToLeft     =   -1  'True
               TabIndex        =   254
               Top             =   180
               Width           =   2640
            End
         End
         Begin ImpulseButton.ISButton btn_New 
            Height          =   390
            Index           =   8
            Left            =   6705
            TabIndex        =   266
            Top             =   6705
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   688
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "ÃœÌœ"
            BackColor       =   14871017
            FontSize        =   12
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "dean.frx":6A774
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Save 
            Height          =   390
            Index           =   8
            Left            =   4935
            TabIndex        =   267
            Top             =   6705
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   688
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "Õ›Ÿ"
            BackColor       =   14871017
            FontSize        =   12
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "dean.frx":6AB0E
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Modify 
            Height          =   390
            Index           =   8
            Left            =   5715
            TabIndex        =   268
            Top             =   6705
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   688
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   " ⁄œÌ·"
            BackColor       =   14871017
            FontSize        =   12
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "dean.frx":6AEA8
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton Btn_Undo 
            Height          =   390
            Index           =   8
            Left            =   4140
            TabIndex        =   269
            Top             =   6705
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   688
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   " —«Ã⁄"
            BackColor       =   14871017
            FontSize        =   12
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "dean.frx":6B242
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Delete 
            Height          =   390
            Index           =   8
            Left            =   3150
            TabIndex        =   270
            Top             =   6705
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   688
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "Õ–›"
            BackColor       =   14871017
            FontSize        =   12
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "dean.frx":6B5DC
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Cancel 
            Height          =   390
            Index           =   8
            Left            =   390
            TabIndex        =   271
            Top             =   6705
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   688
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "Œ—ÊÃ"
            BackColor       =   14871017
            FontSize        =   12
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "dean.frx":6BB76
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton Btn_Print 
            Height          =   390
            Index           =   8
            Left            =   2370
            TabIndex        =   272
            TabStop         =   0   'False
            ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
            Top             =   6705
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   688
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "ÿ»«⁄… "
            BackColor       =   14871017
            FontSize        =   12
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "dean.frx":6BF10
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btn_Query 
            Height          =   390
            Index           =   8
            Left            =   1380
            TabIndex        =   273
            TabStop         =   0   'False
            ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
            Top             =   6705
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   688
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "»ÕÀ"
            BackColor       =   14871017
            FontSize        =   12
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "dean.frx":72772
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin VSFlex8Ctl.VSFlexGrid Grid8 
            Height          =   3510
            Left            =   195
            TabIndex        =   274
            Top             =   720
            Width           =   8475
            _cx             =   14949
            _cy             =   6191
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
            FormatString    =   $"dean.frx":72B0C
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
         Begin MSComDlg.CommonDialog CommonDialog1 
            Left            =   0
            Top             =   0
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
      End
      Begin C1SizerLibCtl.C1Elastic ELe 
         Height          =   9750
         Index           =   11
         Left            =   21855
         TabIndex        =   275
         TabStop         =   0   'False
         Top             =   45
         Width           =   19320
         _cx             =   34078
         _cy             =   17198
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   700
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
         Begin VB.OptionButton optCash 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰ﬁœÌ"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   13995
            RightToLeft     =   -1  'True
            TabIndex        =   522
            Top             =   1200
            Value           =   -1  'True
            Width           =   1980
         End
         Begin VB.OptionButton optLater 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "¬Ã·"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   10845
            RightToLeft     =   -1  'True
            TabIndex        =   521
            Top             =   1200
            Width           =   1770
         End
         Begin VB.PictureBox Picture2 
            BackColor       =   &H00E2E9E9&
            Height          =   360
            Left            =   4935
            ScaleHeight     =   300
            ScaleWidth      =   14130
            TabIndex        =   298
            Top             =   9855
            Visible         =   0   'False
            Width           =   14190
            Begin VB.Frame Frame7 
               BackColor       =   &H00E2E9E9&
               Height          =   1965
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   299
               Top             =   -120
               Width           =   13935
               Begin VB.TextBox txtLetter4 
                  Alignment       =   2  'Center
                  Height          =   555
                  Left            =   2280
                  MaxLength       =   1
                  RightToLeft     =   -1  'True
                  TabIndex        =   301
                  Top             =   240
                  Visible         =   0   'False
                  Width           =   555
               End
               Begin VB.TextBox ntxtLetter4 
                  Alignment       =   2  'Center
                  Height          =   555
                  Left            =   6450
                  MaxLength       =   1
                  RightToLeft     =   -1  'True
                  TabIndex        =   300
                  Top             =   -1110
                  Visible         =   0   'False
                  Width           =   555
               End
            End
            Begin VB.Label XPLbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "‰ » Ã  1 2 3"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   195
               Index           =   15
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   303
               Top             =   -240
               Width           =   1185
            End
            Begin VB.Label XPLbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„À«· "
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000C0&
               Height          =   315
               Index           =   14
               Left            =   1290
               RightToLeft     =   -1  'True
               TabIndex        =   302
               Top             =   -240
               Width           =   465
            End
         End
         Begin VB.TextBox txtAmountVisa 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   195
            MaxLength       =   8
            RightToLeft     =   -1  'True
            TabIndex        =   297
            Top             =   7410
            Width           =   1770
         End
         Begin VB.TextBox txtAmountCash 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   3345
            MaxLength       =   8
            RightToLeft     =   -1  'True
            TabIndex        =   296
            Top             =   7410
            Width           =   2175
         End
         Begin VB.TextBox TxtSearchCode2 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   0
            RightToLeft     =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   295
            Top             =   1740
            Width           =   2955
         End
         Begin VB.TextBox txtCodeSend 
            Alignment       =   1  'Right Justify
            Height          =   435
            Left            =   5130
            RightToLeft     =   -1  'True
            TabIndex        =   294
            Text            =   "+966"
            Top             =   6540
            Visible         =   0   'False
            Width           =   1185
         End
         Begin VB.TextBox txtNoteSerial1 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFFF&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   13995
            Locked          =   -1  'True
            TabIndex        =   293
            Top             =   600
            Width           =   1980
         End
         Begin VB.TextBox TxtVAt2 
            Alignment       =   1  'Right Justify
            Height          =   330
            Left            =   195
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   292
            Top             =   8025
            Visible         =   0   'False
            Width           =   795
         End
         Begin VB.TextBox TxtVAt22 
            Alignment       =   2  'Center
            BackColor       =   &H80000010&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   12420
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   291
            Top             =   7425
            Width           =   1575
         End
         Begin VB.TextBox txtTotalWithVat2 
            Alignment       =   2  'Center
            BackColor       =   &H80000010&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   6900
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   290
            Top             =   7410
            Width           =   1965
         End
         Begin VB.Frame Fra_Header 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   570
            Index           =   7
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   283
            Top             =   -30
            Width           =   20700
            Begin VB.TextBox TxtModFlg2 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H0000FF00&
               Enabled         =   0   'False
               Height          =   285
               Index           =   9
               Left            =   2580
               RightToLeft     =   -1  'True
               TabIndex        =   284
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
                     Picture         =   "dean.frx":72B9B
                     Key             =   "CompanyName"
                  EndProperty
                  BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "dean.frx":72F35
                     Key             =   "Ser"
                  EndProperty
                  BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "dean.frx":732CF
                     Key             =   "Vac_Name"
                  EndProperty
                  BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "dean.frx":73669
                     Key             =   "ShareCount"
                  EndProperty
                  BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "dean.frx":73A03
                     Key             =   "Dis_Count"
                  EndProperty
                  BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "dean.frx":73D9D
                     Key             =   "Bouns"
                  EndProperty
                  BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "dean.frx":74137
                     Key             =   "SharesValue"
                  EndProperty
                  BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "dean.frx":746D1
                     Key             =   "BuyValue"
                  EndProperty
               EndProperty
            End
            Begin ImpulseButton.ISButton btn_Last 
               Height          =   315
               Index           =   9
               Left            =   90
               TabIndex        =   285
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
               ButtonImage     =   "dean.frx":74A6B
               ColorButton     =   14871017
               AcclimateGrayTones=   -1  'True
               DrawFocusRectangle=   0   'False
               DisabledImageExtraction=   0
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btn_Next 
               Height          =   315
               Index           =   9
               Left            =   555
               TabIndex        =   286
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
               ButtonImage     =   "dean.frx":74E05
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btn_Previous 
               Height          =   315
               Index           =   9
               Left            =   1155
               TabIndex        =   287
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
               ButtonImage     =   "dean.frx":7519F
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btn_First 
               Height          =   315
               Index           =   9
               Left            =   1620
               TabIndex        =   288
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
               ButtonImage     =   "dean.frx":75539
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   " ”ÃÌ· œŒÊ· «·„⁄œ« /«·”Ì«—« "
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   21.75
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   495
               Index           =   7
               Left            =   10140
               RightToLeft     =   -1  'True
               TabIndex        =   289
               Top             =   -30
               Width           =   3720
            End
         End
         Begin VB.TextBox TxtSerial1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   360
            Index           =   9
            Left            =   17550
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   282
            Top             =   330
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.TextBox txtPhoneCust 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   7695
            MaxLength       =   9
            TabIndex        =   281
            Top             =   2520
            Width           =   8280
         End
         Begin VB.ComboBox CboPayMentType 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   13005
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   280
            Top             =   10905
            Visible         =   0   'False
            Width           =   2970
         End
         Begin VB.CommandButton cmdPay 
            Caption         =   "«⁄«œ… «—”«· «·—”«·…"
            Height          =   450
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   279
            Top             =   2520
            Width           =   3150
         End
         Begin VB.TextBox XPTxtVal 
            Alignment       =   2  'Center
            BackColor       =   &H8000000C&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   16365
            MaxLength       =   8
            RightToLeft     =   -1  'True
            TabIndex        =   278
            Top             =   7410
            Width           =   1965
         End
         Begin VB.TextBox txtCustName 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   14985
            TabIndex        =   277
            Top             =   2865
            Visible         =   0   'False
            Width           =   2955
         End
         Begin VB.TextBox txtAmountLater 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000010&
            Enabled         =   0   'False
            Height          =   600
            Left            =   5325
            MaxLength       =   8
            RightToLeft     =   -1  'True
            TabIndex        =   276
            Top             =   -1200
            Width           =   1185
         End
         Begin ImpulseButton.ISButton btn_New 
            Height          =   540
            Index           =   9
            Left            =   14790
            TabIndex        =   304
            Top             =   9075
            Width           =   1770
            _ExtentX        =   3122
            _ExtentY        =   953
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "ÃœÌœ"
            BackColor       =   14871017
            FontSize        =   18
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "dean.frx":758D3
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Save 
            Height          =   540
            Index           =   9
            Left            =   11040
            TabIndex        =   305
            Top             =   9075
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   953
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "Õ›Ÿ"
            BackColor       =   14871017
            FontSize        =   18
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "dean.frx":75C6D
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Modify 
            Height          =   540
            Index           =   9
            Left            =   13005
            TabIndex        =   306
            Top             =   9090
            Width           =   1590
            _ExtentX        =   2805
            _ExtentY        =   953
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   " ⁄œÌ·"
            BackColor       =   14871017
            FontSize        =   18
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "dean.frx":76007
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton Btn_Undo 
            Height          =   540
            Index           =   9
            Left            =   9075
            TabIndex        =   307
            Top             =   9075
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   953
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   " —«Ã⁄"
            BackColor       =   14871017
            FontSize        =   18
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "dean.frx":763A1
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Delete 
            Height          =   540
            Index           =   9
            Left            =   7695
            TabIndex        =   308
            Top             =   9075
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   953
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "Õ–›"
            BackColor       =   14871017
            FontSize        =   18
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "dean.frx":7673B
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Cancel 
            Height          =   540
            Index           =   9
            Left            =   1575
            TabIndex        =   309
            Top             =   9075
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   953
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "Œ—ÊÃ"
            BackColor       =   14871017
            FontSize        =   18
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "dean.frx":76CD5
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton Btn_Print 
            Height          =   540
            Index           =   9
            Left            =   5715
            TabIndex        =   310
            TabStop         =   0   'False
            ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
            Top             =   9075
            Visible         =   0   'False
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   953
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "ÿ»«⁄… "
            BackColor       =   14871017
            FontSize        =   18
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "dean.frx":7706F
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btn_Query 
            Height          =   540
            Index           =   9
            Left            =   3750
            TabIndex        =   311
            TabStop         =   0   'False
            ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
            Top             =   9075
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   953
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "»ÕÀ"
            BackColor       =   14871017
            FontSize        =   18
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "dean.frx":7D8D1
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin MSComCtl2.DTPicker XPDtbTrans 
            Height          =   420
            Index           =   9
            Left            =   9465
            TabIndex        =   312
            Top             =   600
            Width           =   2565
            _ExtentX        =   4524
            _ExtentY        =   741
            _Version        =   393216
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   198377473
            CurrentDate     =   38784
         End
         Begin MSDataListLib.DataCombo Dcbranch 
            Bindings        =   "dean.frx":7DC6B
            Height          =   480
            Index           =   9
            Left            =   0
            TabIndex        =   313
            Top             =   600
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   847
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            ListField       =   "account_name"
            BoundColumn     =   "code"
            Text            =   ""
            RightToLeft     =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSComCtl2.DTPicker StartTime 
            Height          =   420
            Left            =   4935
            TabIndex        =   314
            Top             =   600
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   741
            _Version        =   393216
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "'Time: 'hh:mm tt"
            Format          =   198377475
            UpDown          =   -1  'True
            CurrentDate     =   40909
         End
         Begin MSDataListLib.DataCombo cmbPaymentClass 
            Height          =   480
            Left            =   9075
            TabIndex        =   315
            Top             =   8295
            Visible         =   0   'False
            Width           =   2760
            _ExtentX        =   4868
            _ExtentY        =   847
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataListLib.DataCombo cmbCustName 
            Height          =   315
            Left            =   9855
            TabIndex        =   316
            Top             =   10650
            Visible         =   0   'False
            Width           =   6120
            _ExtentX        =   10795
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DCboUserName 
            Height          =   480
            Index           =   9
            Left            =   12615
            TabIndex        =   317
            Top             =   8355
            Width           =   3360
            _ExtentX        =   5927
            _ExtentY        =   847
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataListLib.DataCombo DBCboClientName 
            Height          =   675
            Left            =   7695
            TabIndex        =   318
            Top             =   1740
            Width           =   8280
            _ExtentX        =   14605
            _ExtentY        =   1191
            _Version        =   393216
            ListField       =   "6"
            Text            =   ""
            RightToLeft     =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VSFlex8UCtl.VSFlexGrid grd 
            Height          =   750
            Left            =   195
            TabIndex        =   319
            Top             =   5100
            Width           =   15780
            _cx             =   27834
            _cy             =   1323
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
            GridLines       =   13
            GridLinesFixed  =   2
            GridLineWidth   =   40
            Rows            =   1
            Cols            =   2
            FixedRows       =   0
            FixedCols       =   0
            RowHeightMin    =   800
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"dean.frx":7DC80
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
            Ellipsis        =   1
            ExplorerBar     =   7
            PicturesOver    =   0   'False
            FillStyle       =   0
            RightToLeft     =   -1  'True
            PictureType     =   0
            TabBehavior     =   0
            OwnerDraw       =   0
            Editable        =   2
            ShowComboButton =   1
            WordWrap        =   -1  'True
            TextStyle       =   3
            TextStyleFixed  =   4
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
         Begin C1SizerLibCtl.C1Elastic C1Elastic1 
            Height          =   1815
            Left            =   195
            TabIndex        =   502
            TabStop         =   0   'False
            Top             =   3240
            Width           =   15780
            _cx             =   27834
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
            Begin VB.TextBox txtnBoardNo 
               Alignment       =   2  'Center
               BackColor       =   &H80000000&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1095
               Left            =   810
               RightToLeft     =   -1  'True
               TabIndex        =   518
               Top             =   540
               Width           =   225
            End
            Begin VB.TextBox ntxtLetter1 
               Alignment       =   2  'Center
               BackColor       =   &H80000003&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1095
               Left            =   1440
               MaxLength       =   1
               RightToLeft     =   -1  'True
               TabIndex        =   517
               Top             =   540
               Width           =   75
            End
            Begin VB.TextBox ntxtLetter2 
               Alignment       =   2  'Center
               BackColor       =   &H80000003&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1095
               Left            =   1380
               MaxLength       =   1
               RightToLeft     =   -1  'True
               TabIndex        =   516
               Top             =   540
               Width           =   60
            End
            Begin VB.TextBox ntxtLetter3 
               Alignment       =   2  'Center
               BackColor       =   &H80000003&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1095
               Left            =   1335
               MaxLength       =   1
               RightToLeft     =   -1  'True
               TabIndex        =   515
               Top             =   540
               Width           =   45
            End
            Begin VB.TextBox ntxtNum1 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1095
               Left            =   1230
               MaxLength       =   1
               RightToLeft     =   -1  'True
               TabIndex        =   514
               Top             =   540
               Width           =   75
            End
            Begin VB.TextBox ntxtNum2 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1095
               Left            =   1170
               MaxLength       =   1
               RightToLeft     =   -1  'True
               TabIndex        =   513
               Top             =   540
               Width           =   60
            End
            Begin VB.TextBox ntxtNum3 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1095
               Left            =   1125
               MaxLength       =   1
               RightToLeft     =   -1  'True
               TabIndex        =   512
               Top             =   540
               Width           =   60
            End
            Begin VB.TextBox ntxtNum4 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1095
               Left            =   1050
               MaxLength       =   1
               RightToLeft     =   -1  'True
               TabIndex        =   511
               Top             =   540
               Width           =   75
            End
            Begin VB.TextBox txtBoardNo 
               Alignment       =   2  'Center
               BackColor       =   &H80000000&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1095
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   510
               Top             =   540
               Width           =   225
            End
            Begin VB.TextBox txtNum4 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1095
               Left            =   255
               MaxLength       =   1
               RightToLeft     =   -1  'True
               TabIndex        =   509
               Top             =   540
               Width           =   75
            End
            Begin VB.TextBox txtNum3 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1095
               Left            =   330
               MaxLength       =   1
               RightToLeft     =   -1  'True
               TabIndex        =   508
               Top             =   540
               Width           =   60
            End
            Begin VB.TextBox txtNum2 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1095
               Left            =   390
               MaxLength       =   1
               RightToLeft     =   -1  'True
               TabIndex        =   507
               Top             =   540
               Width           =   45
            End
            Begin VB.TextBox txtNum1 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1095
               Left            =   435
               MaxLength       =   1
               RightToLeft     =   -1  'True
               TabIndex        =   506
               Top             =   540
               Width           =   75
            End
            Begin VB.TextBox txtLetter3 
               Alignment       =   2  'Center
               BackColor       =   &H80000003&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1095
               Left            =   525
               MaxLength       =   1
               RightToLeft     =   -1  'True
               TabIndex        =   505
               Top             =   540
               Width           =   60
            End
            Begin VB.TextBox txtLetter2 
               Alignment       =   2  'Center
               BackColor       =   &H80000003&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1095
               Left            =   585
               MaxLength       =   1
               RightToLeft     =   -1  'True
               TabIndex        =   504
               Top             =   540
               Width           =   60
            End
            Begin VB.TextBox txtLetter1 
               Alignment       =   2  'Center
               BackColor       =   &H80000003&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1095
               Left            =   645
               MaxLength       =   1
               RightToLeft     =   -1  'True
               TabIndex        =   503
               Top             =   540
               Width           =   75
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "⁄—»Ì"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   18
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   525
               Index           =   80
               Left            =   990
               RightToLeft     =   -1  'True
               TabIndex        =   520
               Top             =   0
               Width           =   285
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«‰Ã·Ì“Ì"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   18
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   525
               Index           =   79
               Left            =   210
               RightToLeft     =   -1  'True
               TabIndex        =   519
               Top             =   0
               Width           =   225
            End
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "„·«ÕŸ« "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Index           =   22
            Left            =   3150
            RightToLeft     =   -1  'True
            TabIndex        =   523
            Top             =   1920
            Width           =   2175
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "541793243 „À«·  9 Œ«‰«  »œÊ‰ ’›— "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   510
            Index           =   20
            Left            =   3345
            TabIndex        =   346
            Top             =   2595
            Width           =   4350
         End
         Begin VB.Label XPLbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·œ›⁄"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   13
            Left            =   15975
            TabIndex        =   345
            Top             =   1230
            Width           =   2160
         End
         Begin VB.Label XPLbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‘»ﬂ…"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Index           =   10
            Left            =   1965
            TabIndex        =   344
            Top             =   7500
            Width           =   795
         End
         Begin VB.Label XPLbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰ﬁœÌ"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   9
            Left            =   5520
            TabIndex        =   343
            Top             =   7500
            Width           =   795
         End
         Begin VB.Label XPLbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·⁄„Ì·"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   8
            Left            =   15975
            TabIndex        =   342
            Top             =   1755
            Width           =   2160
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   315
            Index           =   56
            Left            =   5715
            TabIndex        =   341
            Top             =   1440
            Visible         =   0   'False
            Width           =   1980
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„” Œœ„"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Index           =   20
            Left            =   16365
            RightToLeft     =   -1  'True
            TabIndex        =   340
            Top             =   8325
            Width           =   1770
         End
         Begin VB.Label XPLbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·ﬁÌ„…"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   4
            Left            =   18330
            TabIndex        =   339
            Top             =   7455
            Width           =   795
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·ﬁÌ„… «·„÷«›…"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   690
            Index           =   49
            Left            =   14190
            RightToLeft     =   -1  'True
            TabIndex        =   338
            Top             =   7440
            Width           =   1575
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·«Ã„«·Ì »⁄œ «·÷—Ì»…"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   690
            Index           =   48
            Left            =   8670
            RightToLeft     =   -1  'True
            TabIndex        =   337
            Top             =   7440
            Width           =   3165
         End
         Begin VB.Label LabCount_Rec 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   9
            Left            =   2175
            RightToLeft     =   -1  'True
            TabIndex        =   336
            Top             =   8580
            Width           =   975
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”Ã· «·Õ«·Ì:"
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
            Index           =   16
            Left            =   6900
            RightToLeft     =   -1  'True
            TabIndex        =   335
            Top             =   8475
            Width           =   1575
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «·”Ã·« :"
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
            Index           =   17
            Left            =   3750
            RightToLeft     =   -1  'True
            TabIndex        =   334
            Top             =   8475
            Width           =   1575
         End
         Begin VB.Label LabCurr_Rec 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   9
            Left            =   5910
            RightToLeft     =   -1  'True
            TabIndex        =   333
            Top             =   8595
            Width           =   795
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«· «—ÌŒ"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   43
            Left            =   12225
            TabIndex        =   332
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "—ﬁ„ «·”‰œ"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   8
            Left            =   15975
            TabIndex        =   331
            Top             =   600
            Width           =   2160
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·›—⁄"
            Height          =   270
            Index           =   44
            Left            =   9855
            TabIndex        =   330
            Top             =   -1275
            Visible         =   0   'False
            Width           =   795
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "Êﬁ  "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   45
            Left            =   7890
            RightToLeft     =   -1  'True
            TabIndex        =   329
            Top             =   615
            Width           =   1185
         End
         Begin VB.Label XPLbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„Êﬁ⁄"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   25
            Left            =   3150
            TabIndex        =   328
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label XPLbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «··ÊÕ…  "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Index           =   0
            Left            =   15975
            TabIndex        =   327
            Top             =   3285
            Width           =   2550
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÃÊ«·"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   510
            Index           =   46
            Left            =   15975
            TabIndex        =   326
            Top             =   2595
            Width           =   2160
         End
         Begin VB.Label XPLbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÿ—Ìﬁ… «·”œ«œ"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Index           =   1
            Left            =   16365
            TabIndex        =   325
            Top             =   10920
            Visible         =   0   'False
            Width           =   1770
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "›∆… «·”œ«œ"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Index           =   100
            Left            =   16365
            TabIndex        =   324
            Top             =   5160
            Width           =   1575
         End
         Begin VB.Label XPLbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«”„ «·⁄„Ì· «·‰ﬁœÌ"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   3
            Left            =   18135
            TabIndex        =   323
            Top             =   3435
            Visible         =   0   'False
            Width           =   2370
         End
         Begin VB.Label XPLbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "¬Ã·"
            Height          =   405
            Index           =   11
            Left            =   6900
            TabIndex        =   322
            Top             =   120
            Width           =   390
         End
         Begin VB.Label lblClassCat 
            Alignment       =   2  'Center
            Caption         =   " "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   12810
            TabIndex        =   321
            Top             =   6000
            Width           =   3165
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·›∆Â «·„Õœœ…"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   101
            Left            =   16560
            TabIndex        =   320
            Top             =   6000
            Width           =   1575
         End
      End
      Begin C1SizerLibCtl.C1Elastic ELe 
         Height          =   9750
         Index           =   12
         Left            =   22155
         TabIndex        =   347
         TabStop         =   0   'False
         Top             =   45
         Width           =   19320
         _cx             =   34078
         _cy             =   17198
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
         Begin VB.Frame Frame3 
            Caption         =   "«Êﬁ«  «·œÊ«„"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   900
            Left            =   11040
            RightToLeft     =   -1  'True
            TabIndex        =   368
            Top             =   3870
            Width           =   8085
            Begin MSComCtl2.DTPicker TimeIn 
               Height          =   495
               Left            =   3780
               TabIndex        =   369
               Top             =   330
               Width           =   2565
               _ExtentX        =   4524
               _ExtentY        =   873
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               CustomFormat    =   "'Time: 'hh:mm tt"
               Format          =   244645891
               UpDown          =   -1  'True
               CurrentDate     =   40909
            End
            Begin MSComCtl2.DTPicker TimeOut 
               Height          =   435
               Left            =   90
               TabIndex        =   370
               Top             =   360
               Width           =   2895
               _ExtentX        =   5106
               _ExtentY        =   767
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               CustomFormat    =   "'Time: 'hh:mm tt"
               Format          =   244645891
               UpDown          =   -1  'True
               CurrentDate     =   40909
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«·Ï"
               Height          =   285
               Index           =   51
               Left            =   2820
               RightToLeft     =   -1  'True
               TabIndex        =   372
               Top             =   420
               Width           =   645
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "„‰"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   465
               Index           =   50
               Left            =   6210
               RightToLeft     =   -1  'True
               TabIndex        =   371
               Top             =   360
               Width           =   915
            End
         End
         Begin VB.Frame Frame2 
            Height          =   795
            Left            =   6900
            RightToLeft     =   -1  'True
            TabIndex        =   365
            Top             =   3990
            Width           =   3750
            Begin VB.OptionButton optIsEmp 
               Alignment       =   1  'Right Justify
               Caption         =   "„ÊŸ›"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   150
               RightToLeft     =   -1  'True
               TabIndex        =   367
               Top             =   240
               Width           =   1155
            End
            Begin VB.OptionButton optIsResponsible 
               Alignment       =   1  'Right Justify
               Caption         =   "„‘—›"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   2070
               RightToLeft     =   -1  'True
               TabIndex        =   366
               Top             =   270
               Width           =   1155
            End
         End
         Begin VB.TextBox txtSalary 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   5520
            TabIndex        =   364
            Top             =   1485
            Width           =   2955
         End
         Begin VB.TextBox txtEmpName 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   11235
            TabIndex        =   363
            Top             =   2460
            Width           =   5910
         End
         Begin VB.TextBox txtHafizaNo 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   11235
            MaxLength       =   10
            TabIndex        =   362
            Top             =   1635
            Width           =   5910
         End
         Begin VB.TextBox TxtSerial1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Index           =   10
            Left            =   11235
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   361
            Top             =   915
            Width           =   5910
         End
         Begin VB.Frame Fra_Header 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   675
            Index           =   10
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   354
            Top             =   0
            Width           =   19320
            Begin VB.TextBox TxtModFlg2 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H0000FF00&
               Enabled         =   0   'False
               Height          =   285
               Index           =   10
               Left            =   2580
               RightToLeft     =   -1  'True
               TabIndex        =   355
               Text            =   "modflag"
               Top             =   90
               Visible         =   0   'False
               Width           =   465
            End
            Begin MSComctlLib.ImageList GrdImageList2 
               Index           =   10
               Left            =   3150
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
                     Picture         =   "dean.frx":7DC96
                     Key             =   "CompanyName"
                  EndProperty
                  BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "dean.frx":7E030
                     Key             =   "Ser"
                  EndProperty
                  BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "dean.frx":7E3CA
                     Key             =   "Vac_Name"
                  EndProperty
                  BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "dean.frx":7E764
                     Key             =   "ShareCount"
                  EndProperty
                  BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "dean.frx":7EAFE
                     Key             =   "Dis_Count"
                  EndProperty
                  BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "dean.frx":7EE98
                     Key             =   "Bouns"
                  EndProperty
                  BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "dean.frx":7F232
                     Key             =   "SharesValue"
                  EndProperty
                  BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "dean.frx":7F7CC
                     Key             =   "BuyValue"
                  EndProperty
               EndProperty
            End
            Begin ImpulseButton.ISButton btn_Last 
               Height          =   315
               Index           =   10
               Left            =   90
               TabIndex        =   356
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
               ButtonImage     =   "dean.frx":7FB66
               ColorButton     =   14871017
               AcclimateGrayTones=   -1  'True
               DrawFocusRectangle=   0   'False
               DisabledImageExtraction=   0
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btn_Next 
               Height          =   315
               Index           =   10
               Left            =   540
               TabIndex        =   357
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
               ButtonImage     =   "dean.frx":7FF00
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btn_Previous 
               Height          =   315
               Index           =   10
               Left            =   1155
               TabIndex        =   358
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
               ButtonImage     =   "dean.frx":8029A
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btn_First 
               Height          =   315
               Index           =   10
               Left            =   1620
               TabIndex        =   359
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
               ButtonImage     =   "dean.frx":80634
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   " ”ÃÌ· »Ì«‰«  «·„ÊŸ›Ì‰"
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
               Index           =   9
               Left            =   14610
               RightToLeft     =   -1  'True
               TabIndex        =   360
               Top             =   60
               Width           =   2640
            End
         End
         Begin VB.TextBox txtRemark 
            Alignment       =   2  'Center
            Height          =   600
            Left            =   7890
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   353
            Top             =   4920
            Width           =   10245
         End
         Begin VB.TextBox Text2 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   0
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   352
            Top             =   1065
            Visible         =   0   'False
            Width           =   795
         End
         Begin VB.TextBox txtFingerPrint 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   9855
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   351
            Top             =   4575
            Visible         =   0   'False
            Width           =   9075
         End
         Begin VB.Frame Frame4 
            Height          =   1590
            Index           =   3
            Left            =   8280
            RightToLeft     =   -1  'True
            TabIndex        =   349
            Top             =   7050
            Visible         =   0   'False
            Width           =   1380
            Begin ZKFPEngXControl.ZKFPEngX ZKFPEngX1 
               Left            =   0
               Top             =   0
               EnrollCount     =   3
               SensorIndex     =   0
               Threshold       =   10
               VerTplFileName  =   ""
               RegTplFileName  =   ""
               OneToOneThreshold=   10
               Active          =   0   'False
               IsRegister      =   0   'False
               EnrollIndex     =   0
               SensorSN        =   ""
               FPEngineVersion =   "9"
               ImageWidth      =   0
               ImageHeight     =   0
               SensorCount     =   0
               TemplateLen     =   1152
               EngineValid     =   0   'False
               ForceSecondMatch=   0   'False
               IsReturnNoLic   =   -1  'True
               LowestQuality   =   30
               FakeFunOn       =   1
            End
            Begin VB.PictureBox ZKFPEngX1222 
               Height          =   480
               Left            =   2520
               ScaleHeight     =   420
               ScaleWidth      =   1140
               TabIndex        =   350
               Top             =   120
               Width           =   1200
            End
         End
         Begin VB.TextBox TxtMobileNO 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   5520
            MaxLength       =   9
            TabIndex        =   348
            Top             =   2520
            Width           =   2955
         End
         Begin ImpulseButton.ISButton btn_New 
            Height          =   435
            Index           =   10
            Left            =   13800
            TabIndex        =   373
            Top             =   9090
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   767
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "ÃœÌœ"
            BackColor       =   14871017
            FontSize        =   18
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "dean.frx":809CE
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Save 
            Height          =   435
            Index           =   10
            Left            =   10245
            TabIndex        =   374
            Top             =   9105
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   767
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "Õ›Ÿ"
            BackColor       =   14871017
            FontSize        =   18
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "dean.frx":80D68
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Modify 
            Height          =   435
            Index           =   10
            Left            =   12225
            TabIndex        =   375
            Top             =   9105
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   767
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   " ⁄œÌ·"
            BackColor       =   14871017
            FontSize        =   18
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "dean.frx":81102
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton Btn_Undo 
            Height          =   435
            Index           =   10
            Left            =   8670
            TabIndex        =   376
            Top             =   9105
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   767
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   " —«Ã⁄"
            BackColor       =   14871017
            FontSize        =   18
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "dean.frx":8149C
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Delete 
            Height          =   435
            Index           =   10
            Left            =   6315
            TabIndex        =   377
            Top             =   9105
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   767
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "Õ–›"
            BackColor       =   14871017
            FontSize        =   18
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "dean.frx":81836
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton Btn_Update 
            Height          =   240
            Index           =   10
            Left            =   1770
            TabIndex        =   378
            TabStop         =   0   'False
            ToolTipText     =   " ÕœÌÀ ﬁ«⁄œ… «·»Ì«‰« "
            Top             =   8565
            Visible         =   0   'False
            Width           =   1185
            _ExtentX        =   2090
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
            ButtonImage     =   "dean.frx":81DD0
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Cancel 
            Height          =   435
            Index           =   10
            Left            =   0
            TabIndex        =   379
            Top             =   9075
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   767
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "Œ—ÊÃ"
            BackColor       =   14871017
            FontSize        =   18
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "dean.frx":8216A
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton Btn_Print 
            Height          =   510
            Index           =   10
            Left            =   4335
            TabIndex        =   380
            TabStop         =   0   'False
            ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
            Top             =   9060
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   900
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "ÿ»«⁄… "
            BackColor       =   14871017
            FontSize        =   18
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "dean.frx":82504
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btn_Query 
            Height          =   570
            Index           =   10
            Left            =   2175
            TabIndex        =   381
            TabStop         =   0   'False
            ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
            Top             =   9000
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   1005
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "»ÕÀ"
            BackColor       =   14871017
            FontSize        =   18
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "dean.frx":88D66
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin MSComCtl2.DTPicker XPDtbTrans 
            Height          =   540
            Index           =   10
            Left            =   5520
            TabIndex        =   382
            Top             =   915
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   953
            _Version        =   393216
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   244645889
            CurrentDate     =   38784
         End
         Begin MSDataListLib.DataCombo Dcbranch 
            Bindings        =   "dean.frx":89100
            Height          =   480
            Index           =   10
            Left            =   11235
            TabIndex        =   383
            Top             =   3180
            Width           =   5910
            _ExtentX        =   10425
            _ExtentY        =   847
            _Version        =   393216
            Style           =   2
            BackColor       =   16777215
            ListField       =   "account_name"
            BoundColumn     =   "code"
            Text            =   ""
            RightToLeft     =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataListLib.DataCombo DCboUserName 
            Height          =   480
            Index           =   10
            Left            =   13800
            TabIndex        =   384
            Top             =   8655
            Width           =   4140
            _ExtentX        =   7303
            _ExtentY        =   847
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSComCtl2.DTPicker txtStartDate 
            Height          =   375
            Left            =   5520
            TabIndex        =   385
            Top             =   3240
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   661
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   244645889
            CurrentDate     =   38784
         End
         Begin DBPIXLib.DBPix20 DBPix201 
            Height          =   2685
            Left            =   195
            TabIndex        =   386
            Top             =   4680
            Width           =   4530
            _Version        =   131072
            _ExtentX        =   7990
            _ExtentY        =   4736
            _StockProps     =   1
            BackColor       =   16777152
            _Image          =   "dean.frx":89115
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
         Begin ImpulseButton.ISButton ISButton4 
            Height          =   735
            Left            =   195
            TabIndex        =   387
            Top             =   7515
            Width           =   4530
            _ExtentX        =   7990
            _ExtentY        =   1296
            ButtonPositionImage=   1
            Caption         =   "«œ—«Ã  ÂÊÌ… «·„ÊŸ›"
            BackColor       =   14871017
            FontSize        =   13.5
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
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
         Begin VSFlex8UCtl.VSFlexGrid GrdFinger 
            Height          =   2880
            Left            =   0
            TabIndex        =   388
            Top             =   5475
            Visible         =   0   'False
            Width           =   4140
            _cx             =   7302
            _cy             =   5080
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
            GridLines       =   13
            GridLinesFixed  =   2
            GridLineWidth   =   40
            Rows            =   11
            Cols            =   6
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   100
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"dean.frx":8912D
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
            Ellipsis        =   1
            ExplorerBar     =   7
            PicturesOver    =   0   'False
            FillStyle       =   0
            RightToLeft     =   -1  'True
            PictureType     =   0
            TabBehavior     =   0
            OwnerDraw       =   0
            Editable        =   2
            ShowComboButton =   1
            WordWrap        =   -1  'True
            TextStyle       =   3
            TextStyleFixed  =   4
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
         Begin ImpulseButton.ISButton ISButton5 
            Height          =   750
            Left            =   195
            TabIndex        =   389
            Top             =   2400
            Width           =   4530
            _ExtentX        =   7990
            _ExtentY        =   1323
            ButtonPositionImage=   1
            Caption         =   "«œ—«Ã »’„Â «·„ÊŸ›"
            BackColor       =   14871017
            FontSize        =   13.5
            FontBold        =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
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
         Begin VSFlex8UCtl.VSFlexGrid GrdFinger2 
            Height          =   3015
            Left            =   7890
            TabIndex        =   501
            Top             =   5610
            Width           =   10050
            _cx             =   17727
            _cy             =   5318
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
            GridLines       =   13
            GridLinesFixed  =   2
            GridLineWidth   =   40
            Rows            =   6
            Cols            =   5
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   500
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"dean.frx":89235
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
            Ellipsis        =   1
            ExplorerBar     =   7
            PicturesOver    =   0   'False
            FillStyle       =   0
            RightToLeft     =   -1  'True
            PictureType     =   0
            TabBehavior     =   0
            OwnerDraw       =   0
            Editable        =   2
            ShowComboButton =   1
            WordWrap        =   -1  'True
            TextStyle       =   3
            TextStyleFixed  =   4
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
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "»œ«Ì… «·⁄„·"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Index           =   55
            Left            =   8670
            RightToLeft     =   -1  'True
            TabIndex        =   409
            Top             =   3270
            Width           =   2370
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·«Ã— «·ÌÊ„Ï"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Index           =   54
            Left            =   8670
            TabIndex        =   408
            Top             =   1485
            Width           =   2370
         End
         Begin VB.Label XPLbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·«”„"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Index           =   12
            Left            =   17355
            TabIndex        =   407
            Top             =   2460
            Width           =   1770
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "—ﬁ„ «·ÂÊÌ…"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   9
            Left            =   17355
            TabIndex        =   406
            Top             =   1635
            Width           =   1770
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «· ”ÃÌ·"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   52
            Left            =   8670
            TabIndex        =   405
            Top             =   945
            Width           =   2370
         End
         Begin VB.Label LabCurr_Rec 
            BackColor       =   &H00E2E9E9&
            Height          =   255
            Index           =   10
            Left            =   5130
            RightToLeft     =   -1  'True
            TabIndex        =   404
            Top             =   8610
            Width           =   390
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «·”Ã·« :"
            Height          =   255
            Index           =   23
            Left            =   3945
            RightToLeft     =   -1  'True
            TabIndex        =   403
            Top             =   8595
            Width           =   1185
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”Ã· «·Õ«·Ì:"
            Height          =   255
            Index           =   22
            Left            =   5520
            RightToLeft     =   -1  'True
            TabIndex        =   402
            Top             =   8595
            Width           =   1575
         End
         Begin VB.Label LabCount_Rec 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   270
            Index           =   10
            Left            =   3150
            RightToLeft     =   -1  'True
            TabIndex        =   401
            Top             =   8580
            Width           =   405
         End
         Begin VB.Label XPLbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„Êﬁ⁄ «·⁄„·"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   690
            Index           =   7
            Left            =   17355
            TabIndex        =   400
            Top             =   3240
            Width           =   1770
         End
         Begin VB.Label XPLbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„·«ÕŸ« "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Index           =   6
            Left            =   17745
            TabIndex        =   399
            Top             =   4920
            Width           =   1380
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„” Œœ„"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   510
            Index           =   21
            Left            =   17940
            RightToLeft     =   -1  'True
            TabIndex        =   398
            Top             =   8655
            Width           =   1185
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·„”·”·"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   510
            Index           =   11
            Left            =   17355
            TabIndex        =   397
            Top             =   945
            Width           =   1770
         End
         Begin VB.Label StatusBar 
            BeginProperty Font 
               Name            =   "SimSun"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   0
            TabIndex        =   396
            Top             =   0
            Width           =   5520
         End
         Begin VB.Label lblProgressFinger 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   495
            Left            =   4935
            RightToLeft     =   -1  'True
            TabIndex        =   395
            Top             =   6720
            Width           =   1170
         End
         Begin VB.Label lblFingerStatus 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00FF0000&
            Height          =   525
            Left            =   6105
            RightToLeft     =   -1  'True
            TabIndex        =   394
            Top             =   6720
            Width           =   1185
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Õ«·Â ÃÂ«“ «·»’„Â"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   495
            Left            =   2760
            RightToLeft     =   -1  'True
            TabIndex        =   393
            Top             =   1680
            Width           =   2175
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "„ ’·/€Ì— „ ’·"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   495
            Left            =   195
            RightToLeft     =   -1  'True
            TabIndex        =   392
            Top             =   1695
            Width           =   1980
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÃÊ«·"
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
            Index           =   42
            Left            =   8670
            TabIndex        =   391
            Top             =   2640
            Width           =   2370
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "541793243 „À«·  9 Œ«‰«  »œÊ‰ ’›— "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   390
            Index           =   60
            Left            =   5910
            TabIndex        =   390
            Top             =   2175
            Width           =   3555
         End
      End
      Begin C1SizerLibCtl.C1Elastic ELe 
         Height          =   9750
         Index           =   13
         Left            =   22455
         TabIndex        =   410
         TabStop         =   0   'False
         Top             =   45
         Width           =   19320
         _cx             =   34078
         _cy             =   17198
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
         Begin VB.CommandButton Command4 
            Caption         =   "«€·«ﬁ «·ÌÊ„"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   990
            RightToLeft     =   -1  'True
            TabIndex        =   526
            Top             =   720
            Width           =   2565
         End
         Begin VB.Frame Fra_Header 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   675
            Index           =   9
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   414
            Top             =   30
            Width           =   19320
            Begin MSComctlLib.ImageList GrdImageList2 
               Index           =   11
               Left            =   3150
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
                     Picture         =   "dean.frx":89343
                     Key             =   "CompanyName"
                  EndProperty
                  BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "dean.frx":896DD
                     Key             =   "Ser"
                  EndProperty
                  BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "dean.frx":89A77
                     Key             =   "Vac_Name"
                  EndProperty
                  BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "dean.frx":89E11
                     Key             =   "ShareCount"
                  EndProperty
                  BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "dean.frx":8A1AB
                     Key             =   "Dis_Count"
                  EndProperty
                  BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "dean.frx":8A545
                     Key             =   "Bouns"
                  EndProperty
                  BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "dean.frx":8A8DF
                     Key             =   "SharesValue"
                  EndProperty
                  BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "dean.frx":8AE79
                     Key             =   "BuyValue"
                  EndProperty
               EndProperty
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "‘«‘…  ”ÃÌ· «·Õ÷Ê— Ê«·«‰’—«›"
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
               Index           =   12
               Left            =   8370
               RightToLeft     =   -1  'True
               TabIndex        =   415
               Top             =   60
               Width           =   4440
            End
         End
         Begin VB.CommandButton cmdRefresh 
            Caption         =   " ÕœÌÀ"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   3555
            RightToLeft     =   -1  'True
            TabIndex        =   413
            Top             =   735
            Width           =   2550
         End
         Begin VB.Timer Timer1 
            Left            =   7440
            Top             =   900
         End
         Begin VB.Timer Timer2 
            Interval        =   60000
            Left            =   8220
            Top             =   990
         End
         Begin VB.TextBox txtFingerPrint2 
            Alignment       =   1  'Right Justify
            Height          =   600
            Left            =   -990
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   412
            Top             =   1905
            Visible         =   0   'False
            Width           =   8880
         End
         Begin VB.Frame Frame4 
            Height          =   2190
            Index           =   2
            Left            =   585
            RightToLeft     =   -1  'True
            TabIndex        =   411
            Top             =   4515
            Visible         =   0   'False
            Width           =   1785
            Begin ZKFPEngXControl.ZKFPEngX ZKFPEngX2 
               Left            =   0
               Top             =   0
               EnrollCount     =   3
               SensorIndex     =   0
               Threshold       =   10
               VerTplFileName  =   ""
               RegTplFileName  =   ""
               OneToOneThreshold=   10
               Active          =   0   'False
               IsRegister      =   0   'False
               EnrollIndex     =   0
               SensorSN        =   ""
               FPEngineVersion =   "9"
               ImageWidth      =   0
               ImageHeight     =   0
               SensorCount     =   0
               TemplateLen     =   1152
               EngineValid     =   0   'False
               ForceSecondMatch=   0   'False
               IsReturnNoLic   =   -1  'True
               LowestQuality   =   30
               FakeFunOn       =   1
            End
         End
         Begin MSDataListLib.DataCombo cmbEmpName 
            Height          =   675
            Left            =   8865
            TabIndex        =   416
            Top             =   1470
            Width           =   8685
            _ExtentX        =   15319
            _ExtentY        =   1191
            _Version        =   393216
            ListField       =   "6"
            Text            =   ""
            RightToLeft     =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSComCtl2.DTPicker XPDtbTrans 
            Height          =   405
            Index           =   11
            Left            =   14595
            TabIndex        =   417
            Top             =   840
            Width           =   2760
            _ExtentX        =   4868
            _ExtentY        =   714
            _Version        =   393216
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   245039105
            CurrentDate     =   38784
         End
         Begin MSComCtl2.DTPicker txtTimeIn 
            Height          =   405
            Left            =   8865
            TabIndex        =   418
            Top             =   885
            Visible         =   0   'False
            Width           =   2370
            _ExtentX        =   4180
            _ExtentY        =   714
            _Version        =   393216
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "'Time: 'hh:mm tt"
            Format          =   245039107
            UpDown          =   -1  'True
            CurrentDate     =   40909
         End
         Begin VSFlex8UCtl.VSFlexGrid GrdEmp 
            Height          =   7335
            Left            =   990
            TabIndex        =   419
            Top             =   2550
            Width           =   17535
            _cx             =   30930
            _cy             =   12938
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
            GridLines       =   13
            GridLinesFixed  =   2
            GridLineWidth   =   40
            Rows            =   2
            Cols            =   9
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   800
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"dean.frx":8B213
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
            Ellipsis        =   1
            ExplorerBar     =   7
            PicturesOver    =   0   'False
            FillStyle       =   0
            RightToLeft     =   -1  'True
            PictureType     =   0
            TabBehavior     =   0
            OwnerDraw       =   0
            Editable        =   2
            ShowComboButton =   1
            WordWrap        =   -1  'True
            TextStyle       =   3
            TextStyleFixed  =   4
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
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "»’„… «·«’»⁄ €Ì— „⁄—›… ÷⁄ «’»⁄ﬂ „—… «Œ—Ï"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   510
            Index           =   25
            Left            =   8085
            RightToLeft     =   -1  'True
            TabIndex        =   525
            Top             =   1320
            Visible         =   0   'False
            Width           =   5910
         End
         Begin VB.Label XPLbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„ÊŸ›"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   18
            Left            =   17745
            TabIndex        =   423
            Top             =   1470
            Width           =   1380
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "«·”«⁄…"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   62
            Left            =   11835
            RightToLeft     =   -1  'True
            TabIndex        =   422
            Top             =   915
            Visible         =   0   'False
            Width           =   1380
         End
         Begin VB.Label XPLbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·ÌÊ„"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   19
            Left            =   17550
            TabIndex        =   421
            Top             =   765
            Width           =   1380
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "»’„… «·«’»⁄ €Ì— „⁄—›… ÷⁄ «’»⁄ﬂ „—… «Œ—Ï"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   510
            Index           =   61
            Left            =   2565
            RightToLeft     =   -1  'True
            TabIndex        =   420
            Top             =   1560
            Visible         =   0   'False
            Width           =   5910
         End
      End
      Begin C1SizerLibCtl.C1Elastic ELe 
         Height          =   9750
         Index           =   15
         Left            =   22755
         TabIndex        =   424
         TabStop         =   0   'False
         Top             =   45
         Width           =   19320
         _cx             =   34078
         _cy             =   17198
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
            Height          =   630
            Index           =   11
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   461
            Top             =   0
            Width           =   19320
            Begin VB.TextBox TxtModFlg2 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H0000FF00&
               Enabled         =   0   'False
               Height          =   285
               Index           =   12
               Left            =   2580
               RightToLeft     =   -1  'True
               TabIndex        =   462
               Text            =   "modflag"
               Top             =   90
               Visible         =   0   'False
               Width           =   465
            End
            Begin MSComctlLib.ImageList GrdImageList2 
               Index           =   12
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
                     Picture         =   "dean.frx":8B351
                     Key             =   "CompanyName"
                  EndProperty
                  BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "dean.frx":8B6EB
                     Key             =   "Ser"
                  EndProperty
                  BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "dean.frx":8BA85
                     Key             =   "Vac_Name"
                  EndProperty
                  BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "dean.frx":8BE1F
                     Key             =   "ShareCount"
                  EndProperty
                  BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "dean.frx":8C1B9
                     Key             =   "Dis_Count"
                  EndProperty
                  BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "dean.frx":8C553
                     Key             =   "Bouns"
                  EndProperty
                  BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "dean.frx":8C8ED
                     Key             =   "SharesValue"
                  EndProperty
                  BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "dean.frx":8CE87
                     Key             =   "BuyValue"
                  EndProperty
               EndProperty
            End
            Begin ImpulseButton.ISButton btn_Last 
               Height          =   315
               Index           =   12
               Left            =   90
               TabIndex        =   463
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
               ButtonImage     =   "dean.frx":8D221
               ColorButton     =   14871017
               AcclimateGrayTones=   -1  'True
               DrawFocusRectangle=   0   'False
               DisabledImageExtraction=   0
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btn_Next 
               Height          =   315
               Index           =   12
               Left            =   555
               TabIndex        =   464
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
               ButtonImage     =   "dean.frx":8D5BB
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btn_Previous 
               Height          =   315
               Index           =   12
               Left            =   1155
               TabIndex        =   465
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
               ButtonImage     =   "dean.frx":8D955
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btn_First 
               Height          =   315
               Index           =   12
               Left            =   1620
               TabIndex        =   466
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
               ButtonImage     =   "dean.frx":8DCEF
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Œ’Ê„«  «·⁄ﬁ«—« "
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
               Index           =   0
               Left            =   14250
               RightToLeft     =   -1  'True
               TabIndex        =   467
               Top             =   90
               Width           =   4320
            End
            Begin VB.Image ImgFavorites 
               Height          =   390
               Index           =   4
               Left            =   11100
               Picture         =   "dean.frx":8E089
               Stretch         =   -1  'True
               Top             =   30
               Width           =   525
            End
         End
         Begin VB.TextBox TxtSerial1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   375
            Index           =   12
            Left            =   15975
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   460
            Top             =   975
            Width           =   1965
         End
         Begin VB.Frame Frame4 
            Height          =   3990
            Index           =   4
            Left            =   195
            RightToLeft     =   -1  'True
            TabIndex        =   427
            Top             =   1590
            Width           =   18735
            Begin VB.CommandButton cmdReloadList 
               Caption         =   "«·€«¡ «·„Õœœ"
               Height          =   225
               Index           =   0
               Left            =   9810
               RightToLeft     =   -1  'True
               TabIndex        =   499
               Top             =   3720
               Width           =   1995
            End
            Begin VB.ListBox ListUnitNoAll2 
               Height          =   1035
               ItemData        =   "dean.frx":91CF1
               Left            =   2550
               List            =   "dean.frx":91CF8
               MultiSelect     =   1  'Simple
               RightToLeft     =   -1  'True
               TabIndex        =   498
               Top             =   1980
               Visible         =   0   'False
               Width           =   1755
            End
            Begin VB.ListBox ListUnitNoSelected2 
               BackColor       =   &H0080FFFF&
               Height          =   1230
               ItemData        =   "dean.frx":91D0A
               Left            =   240
               List            =   "dean.frx":91D11
               RightToLeft     =   -1  'True
               TabIndex        =   497
               Top             =   2130
               Visible         =   0   'False
               Width           =   1575
            End
            Begin VB.ListBox ListAqarAll 
               Height          =   3375
               ItemData        =   "dean.frx":91D28
               Left            =   11280
               List            =   "dean.frx":91D2F
               MultiSelect     =   1  'Simple
               RightToLeft     =   -1  'True
               TabIndex        =   435
               Top             =   330
               Width           =   1995
            End
            Begin VB.ListBox ListAqarSelected 
               BackColor       =   &H0080FFFF&
               Height          =   3375
               ItemData        =   "dean.frx":91D41
               Left            =   9030
               List            =   "dean.frx":91D48
               RightToLeft     =   -1  'True
               TabIndex        =   434
               Top             =   360
               Width           =   1845
            End
            Begin VB.ListBox ListBranchAll 
               Height          =   3375
               ItemData        =   "dean.frx":91D5F
               Left            =   16020
               List            =   "dean.frx":91D66
               MultiSelect     =   1  'Simple
               RightToLeft     =   -1  'True
               TabIndex        =   433
               Top             =   330
               Width           =   2025
            End
            Begin VB.ListBox ListBranchSelected 
               BackColor       =   &H0080FFFF&
               Height          =   3375
               ItemData        =   "dean.frx":91D78
               Left            =   13500
               List            =   "dean.frx":91D7F
               MultiSelect     =   1  'Simple
               RightToLeft     =   -1  'True
               TabIndex        =   432
               Top             =   330
               Width           =   2055
            End
            Begin VB.ListBox ListUnitTypeAll 
               Height          =   3375
               ItemData        =   "dean.frx":91D96
               Left            =   7110
               List            =   "dean.frx":91D9D
               MultiSelect     =   1  'Simple
               RightToLeft     =   -1  'True
               TabIndex        =   431
               Top             =   360
               Width           =   1755
            End
            Begin VB.ListBox ListUnitTypeSelected 
               BackColor       =   &H0080FFFF&
               Height          =   3375
               ItemData        =   "dean.frx":91DAF
               Left            =   4590
               List            =   "dean.frx":91DB6
               RightToLeft     =   -1  'True
               TabIndex        =   430
               Top             =   390
               Width           =   2055
            End
            Begin VB.ListBox ListUnitNoSelected 
               BackColor       =   &H0080FFFF&
               Height          =   3375
               ItemData        =   "dean.frx":91DCD
               Left            =   150
               List            =   "dean.frx":91DD4
               RightToLeft     =   -1  'True
               TabIndex        =   429
               Top             =   390
               Width           =   1875
            End
            Begin VB.ListBox ListUnitNoAll 
               Height          =   3375
               ItemData        =   "dean.frx":91DEB
               Left            =   2490
               List            =   "dean.frx":91DF2
               MultiSelect     =   1  'Simple
               RightToLeft     =   -1  'True
               TabIndex        =   428
               Top             =   360
               Width           =   1755
            End
            Begin VB.Label Label4 
               Alignment       =   2  'Center
               Caption         =   ">"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Left            =   10860
               RightToLeft     =   -1  'True
               TabIndex        =   459
               Top             =   570
               Width           =   495
            End
            Begin VB.Label Label9 
               Alignment       =   2  'Center
               Caption         =   ">>"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Left            =   10860
               RightToLeft     =   -1  'True
               TabIndex        =   458
               Top             =   810
               Width           =   495
            End
            Begin VB.Label Label11 
               Alignment       =   2  'Center
               Caption         =   "<<"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   255
               Left            =   10860
               RightToLeft     =   -1  'True
               TabIndex        =   457
               Top             =   1170
               Width           =   495
            End
            Begin VB.Label Label12 
               Alignment       =   2  'Center
               Caption         =   "<"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   255
               Left            =   10860
               RightToLeft     =   -1  'True
               TabIndex        =   456
               Top             =   1410
               Width           =   495
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               Caption         =   "ﬂ· «·⁄ﬁ«—« "
               Height          =   255
               Index           =   63
               Left            =   11730
               RightToLeft     =   -1  'True
               TabIndex        =   455
               Top             =   0
               Width           =   1335
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               Caption         =   "«·⁄ﬁ«—«  «·„Õœœ…"
               Height          =   255
               Index           =   64
               Left            =   9210
               RightToLeft     =   -1  'True
               TabIndex        =   454
               Top             =   -30
               Width           =   1335
            End
            Begin VB.Label Label14 
               Alignment       =   2  'Center
               Caption         =   ">"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Left            =   15510
               RightToLeft     =   -1  'True
               TabIndex        =   453
               Top             =   480
               Width           =   495
            End
            Begin VB.Label Label15 
               Alignment       =   2  'Center
               Caption         =   ">>"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Left            =   15510
               RightToLeft     =   -1  'True
               TabIndex        =   452
               Top             =   720
               Width           =   495
            End
            Begin VB.Label Label16 
               Alignment       =   2  'Center
               Caption         =   "<<"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   255
               Left            =   15510
               RightToLeft     =   -1  'True
               TabIndex        =   451
               Top             =   1080
               Width           =   495
            End
            Begin VB.Label Label17 
               Alignment       =   2  'Center
               Caption         =   "<"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   255
               Left            =   15510
               RightToLeft     =   -1  'True
               TabIndex        =   450
               Top             =   1320
               Width           =   495
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               Caption         =   "«·›—Ê⁄ «·„Õœœ…"
               Height          =   255
               Index           =   65
               Left            =   13830
               RightToLeft     =   -1  'True
               TabIndex        =   449
               Top             =   30
               Width           =   1335
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               Caption         =   "ﬂ· «·›—Ê⁄"
               Height          =   255
               Index           =   66
               Left            =   16320
               RightToLeft     =   -1  'True
               TabIndex        =   448
               Top             =   30
               Width           =   1335
            End
            Begin VB.Label Label18 
               Alignment       =   2  'Center
               Caption         =   ">"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Left            =   6630
               RightToLeft     =   -1  'True
               TabIndex        =   447
               Top             =   600
               Width           =   495
            End
            Begin VB.Label Label19 
               Alignment       =   2  'Center
               Caption         =   ">>"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Left            =   6630
               RightToLeft     =   -1  'True
               TabIndex        =   446
               Top             =   840
               Width           =   495
            End
            Begin VB.Label Label20 
               Alignment       =   2  'Center
               Caption         =   "<<"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   255
               Left            =   6630
               RightToLeft     =   -1  'True
               TabIndex        =   445
               Top             =   1200
               Width           =   495
            End
            Begin VB.Label Label21 
               Alignment       =   2  'Center
               Caption         =   "<"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   255
               Left            =   6630
               RightToLeft     =   -1  'True
               TabIndex        =   444
               Top             =   1440
               Width           =   495
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               Caption         =   "ﬂ· «·ÊÕœ« "
               Height          =   255
               Index           =   70
               Left            =   7410
               RightToLeft     =   -1  'True
               TabIndex        =   443
               Top             =   30
               Width           =   1335
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               Caption         =   "«·ÊÕœ«  «·„Õœœ…"
               Height          =   255
               Index           =   71
               Left            =   4980
               RightToLeft     =   -1  'True
               TabIndex        =   442
               Top             =   0
               Width           =   1335
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               Caption         =   "«·ÊÕœ«  «·„Õœœ…"
               Height          =   255
               Index           =   73
               Left            =   390
               RightToLeft     =   -1  'True
               TabIndex        =   441
               Top             =   0
               Width           =   1335
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               Caption         =   "ﬂ· «—ﬁ«„ «·ÊÕœ« "
               Height          =   255
               Index           =   74
               Left            =   2820
               RightToLeft     =   -1  'True
               TabIndex        =   440
               Top             =   30
               Width           =   1335
            End
            Begin VB.Label Label22 
               Alignment       =   2  'Center
               Caption         =   "<"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   255
               Left            =   2040
               RightToLeft     =   -1  'True
               TabIndex        =   439
               Top             =   1440
               Width           =   495
            End
            Begin VB.Label Label23 
               Alignment       =   2  'Center
               Caption         =   "<<"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   255
               Left            =   2040
               RightToLeft     =   -1  'True
               TabIndex        =   438
               Top             =   1200
               Width           =   495
            End
            Begin VB.Label Label24 
               Alignment       =   2  'Center
               Caption         =   ">>"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Left            =   2040
               RightToLeft     =   -1  'True
               TabIndex        =   437
               Top             =   840
               Width           =   495
            End
            Begin VB.Label Label25 
               Alignment       =   2  'Center
               Caption         =   ">"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   255
               Left            =   2040
               RightToLeft     =   -1  'True
               TabIndex        =   436
               Top             =   600
               Width           =   495
            End
         End
         Begin VB.TextBox txtDiscountPercent 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1575
            RightToLeft     =   -1  'True
            TabIndex        =   426
            Top             =   1035
            Width           =   2175
         End
         Begin VB.CommandButton Command3 
            Caption         =   "«œ—«Ã"
            Height          =   780
            Left            =   195
            RightToLeft     =   -1  'True
            TabIndex        =   425
            Top             =   690
            Width           =   1185
         End
         Begin VSFlex8UCtl.VSFlexGrid GrdIqar 
            Height          =   2505
            Left            =   195
            TabIndex        =   468
            Top             =   5670
            Width           =   18930
            _cx             =   33390
            _cy             =   4419
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
            Cols            =   13
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"dean.frx":91E04
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
         Begin MSDataListLib.DataCombo DCboUserName 
            Height          =   315
            Index           =   12
            Left            =   13800
            TabIndex        =   469
            Top             =   8955
            Width           =   2565
            _ExtentX        =   4524
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin ImpulseButton.ISButton btn_New 
            Height          =   330
            Index           =   12
            Left            =   11835
            TabIndex        =   470
            Top             =   9285
            Width           =   975
            _ExtentX        =   1720
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
            ButtonImage     =   "dean.frx":91FE1
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Save 
            Height          =   345
            Index           =   12
            Left            =   9855
            TabIndex        =   471
            Top             =   9270
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   609
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
            ButtonImage     =   "dean.frx":9237B
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Modify 
            Height          =   225
            Index           =   12
            Left            =   10845
            TabIndex        =   472
            Top             =   9285
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   397
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
            ButtonImage     =   "dean.frx":92715
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton Btn_Undo 
            Height          =   225
            Index           =   12
            Left            =   9270
            TabIndex        =   473
            Top             =   9285
            Width           =   585
            _ExtentX        =   1032
            _ExtentY        =   397
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
            ButtonImage     =   "dean.frx":92AAF
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Delete 
            Height          =   345
            Index           =   12
            Left            =   8475
            TabIndex        =   474
            Top             =   9270
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   609
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
            ButtonImage     =   "dean.frx":92E49
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Cancel 
            Height          =   345
            Index           =   12
            Left            =   4935
            TabIndex        =   475
            Top             =   9225
            Width           =   585
            _ExtentX        =   1032
            _ExtentY        =   609
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
            ButtonImage     =   "dean.frx":933E3
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton Btn_Print 
            Height          =   390
            Index           =   12
            Left            =   7095
            TabIndex        =   476
            TabStop         =   0   'False
            ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
            Top             =   9210
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   688
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
            ButtonImage     =   "dean.frx":9377D
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btn_Query 
            Height          =   420
            Index           =   12
            Left            =   5520
            TabIndex        =   477
            TabStop         =   0   'False
            ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
            Top             =   9180
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   741
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
            ButtonImage     =   "dean.frx":99FDF
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton Cmd_DeleteRow 
            Height          =   285
            Index           =   12
            Left            =   1575
            TabIndex        =   478
            Top             =   8445
            Width           =   1980
            _ExtentX        =   3493
            _ExtentY        =   503
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   " Õ–› ”ÿ—"
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
            ButtonImage     =   "dean.frx":9A379
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton Cmd_DeleteAll 
            Height          =   300
            Index           =   12
            Left            =   0
            TabIndex        =   479
            Top             =   8430
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   529
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   " Õ–› «·ﬂ·"
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
            ButtonImage     =   "dean.frx":9A913
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin MSComCtl2.DTPicker XPDtbTrans 
            Height          =   345
            Index           =   12
            Left            =   12810
            TabIndex        =   480
            Top             =   945
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   609
            _Version        =   393216
            Format          =   245039105
            CurrentDate     =   38784
         End
         Begin MSDataListLib.DataCombo Dcbranch 
            Bindings        =   "dean.frx":9AEAD
            Height          =   315
            Index           =   12
            Left            =   14985
            TabIndex        =   481
            Top             =   510
            Visible         =   0   'False
            Width           =   2955
            _ExtentX        =   5212
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
         Begin MSComCtl2.DTPicker txtDateC 
            Height          =   345
            Index           =   0
            Left            =   8475
            TabIndex        =   482
            Top             =   885
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   609
            _Version        =   393216
            Format          =   236847105
            CurrentDate     =   38784
         End
         Begin MSComCtl2.DTPicker txtDateC 
            Height          =   345
            Index           =   1
            Left            =   5520
            TabIndex        =   483
            Top             =   885
            Width           =   1770
            _ExtentX        =   3122
            _ExtentY        =   609
            _Version        =   393216
            Format          =   236847105
            CurrentDate     =   38784
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Õ—— »Ê«”ÿ…  "
            Height          =   315
            Index           =   67
            Left            =   17145
            TabIndex        =   494
            Top             =   8895
            Width           =   990
         End
         Begin VB.Label LabCount_Rec 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   225
            Index           =   12
            Left            =   2175
            RightToLeft     =   -1  'True
            TabIndex        =   493
            Top             =   8910
            Width           =   780
         End
         Begin VB.Label LabCurr_Rec 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   225
            Index           =   12
            Left            =   4140
            RightToLeft     =   -1  'True
            TabIndex        =   492
            Top             =   8910
            Width           =   795
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «·”Ã·« :"
            Height          =   225
            Index           =   24
            Left            =   2955
            RightToLeft     =   -1  'True
            TabIndex        =   491
            Top             =   8895
            Width           =   1185
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”Ã· «·Õ«·Ì:"
            Height          =   225
            Index           =   25
            Left            =   4935
            RightToLeft     =   -1  'True
            TabIndex        =   490
            Top             =   8895
            Width           =   780
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "—ﬁ„ «·”‰œ"
            Height          =   315
            Index           =   12
            Left            =   18135
            TabIndex        =   489
            Top             =   795
            Width           =   990
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «·”‰œ"
            Height          =   315
            Index           =   68
            Left            =   14385
            TabIndex        =   488
            Top             =   975
            Width           =   1380
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·›—⁄"
            Height          =   270
            Index           =   69
            Left            =   18330
            TabIndex        =   487
            Top             =   1185
            Width           =   600
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰”»… «·Œ’„"
            Height          =   300
            Index           =   72
            Left            =   3750
            TabIndex        =   486
            Top             =   1065
            Width           =   1185
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "„‰  «—ÌŒ"
            Height          =   300
            Index           =   21
            Left            =   10245
            TabIndex        =   485
            Top             =   915
            Width           =   1380
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·Ï  «—ÌŒ "
            Height          =   300
            Index           =   75
            Left            =   7290
            TabIndex        =   484
            Top             =   945
            Width           =   1185
         End
      End
   End
End
Attribute VB_Name = "dean"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsSavRec As ADODB.Recordset
Dim mPath As String
Dim BKGrndPic As ClsBackGroundPic
Dim RecId As String
Dim ii As Long
Dim s As String
Dim Dcombos As ClsDataCombos
Dim mTableName As String
Dim rs As ADODB.Recordset
Dim TTP As clstooltip
Dim GroupReport As ClsGroupReport
Dim cSearchDcbo As clsDCboSearch
Dim rsDummy As ADODB.Recordset
Public mIndex As Integer
Dim mEmpId As Long
Dim FTempLen As Integer
Dim FRegTemplate As String
Dim FRegTemp As Variant
Dim FingerCount As Long
Dim fpcHandle As Long
Dim FFingerNames() As String
Dim FMatchType As Integer
Dim mSenesor As Boolean
Dim mFinger As Long
Dim StrSQL  As String
Dim rsDD As ADODB.Recordset
Private Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hWnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long

Dim mBranchID As Long


Private Sub InsertEmp(ByVal mEmp As Long)
Dim rsDummy2 As New ADODB.Recordset
Dim rsDummyMaxID As New ADODB.Recordset
s = " Select * from TblEmpData Where Id = " & val(mEmp)
rsDummy2.Open s, Cn, adOpenKeyset, adLockOptimistic
If rsDummy2.EOF Then
    Exit Sub
End If
s = " Select * from TblEmpDataInOut Where EmpID = " & val(mEmp) & " And RecordDate =" & SQLDate(XPDtbTrans(mIndex).value, True)
Set rsDummy = New ADODB.Recordset
rsDummy.Open s, Cn, adOpenKeyset, adLockOptimistic
Dim mMaxId As Long
mMaxId = 1
If rsDummy.EOF Then
    s = " Select Max(ID) mMaxID from TblEmpDataInOut "
    Set rsDummyMaxID = New ADODB.Recordset
    rsDummyMaxID.Open s, Cn, adOpenKeyset, adLockOptimistic
    If Not rsDummyMaxID.EOF Then
        mMaxId = val(rsDummyMaxID!mMaxId & "") + 1
    End If


    rsDummy.AddNew
    rsDummy!ID = mMaxId
    rsDummy!TimeIn = Time
    rsDummy!RecordDate = XPDtbTrans(mIndex).value
    rsDummy!BranchID = rsDummy2!BranchID
    rsDummy!HafizaNo = rsDummy2!HafizaNo
    rsDummy!EmpID = rsDummy2!ID
    rsDummy!EmpName = rsDummy2!EmpName
    rsDummy!UserID = user_id
    
Else
    rsDummy!TimeOut = Time
    rsDummy!hours = GetTimeDiff(rsDummy!TimeIn, rsDummy!TimeOut, 1)
    
End If
rsDummy.update

End Sub


Private Sub cmbEmpName_Change()
    If val(cmbEmpName.BoundText) <> 0 And cmbEmpName.text <> "" Then
        InsertEmp cmbEmpName.BoundText
    End If
End Sub

Private Sub Cmd_Click(Index As Integer)
If Index = 20 Then
    GetDataCus
End If
End Sub

Private Sub CmdRefresh_Click()
Dim ss As String
If mIndex = 11 Then
'ss = "Select * from TblEmpDataInOut Where RecordDate =" & SQLDate(XPDtbTrans(mIndex).value, True)
   ss = "Select * from TblEmpDataInOut Where RecordDate"
       ss = ss & "  in ( select max(RecordDate) as RecordDate from TblEmpDataInOut where timeout is null )"
loadgrid ss, GrdEmp, True, False
End If

End Sub
Private Sub FGS_Click()
    On Error GoTo ErrTrap

    Dim mCode As String, mFullCode As String

    If Not FGS.TextMatrix(FGS.Row, 1) = "" Then
        mCode = val(FGS.TextMatrix(FGS.Row, 1))
        mFullCode = Trim(FGS.TextMatrix(FGS.Row, FGS.ColIndex("Fullcode")))
       TxtSearchCode = mFullCode
               DcCustmer.BoundText = val(FGS.TextMatrix(FGS.Row, 1))
                If Trim(FGS.TextMatrix(FGS.Row, FGS.ColIndex("Cus_mobile"))) <> "" Then
                        TxtPhone.text = Trim(FGS.TextMatrix(FGS.Row, FGS.ColIndex("Cus_mobile")))
                Else
                        TxtPhone.text = Trim(FGS.TextMatrix(FGS.Row, FGS.ColIndex("Phone")))
                        'frmsalebill3.TxtPhone
                End If

    End If
    Frame1(0).Visible = False
ErrTrap:
    Exit Sub
End Sub

Private Sub cmdReloadList_Click(Index As Integer)
    If Index = 1 Then
        'ListProductLineSelected.Clear
        FillMylist
    Else
            
        FillMylist2 True, True, True, True
    End If
End Sub

Private Sub Command3_Click()

If val(txtDiscountPercent) = 0 Then MsgBox "Õœœ ‰”»Â Œ’„ ’ÕÌÕÂ"", vbCritical : Exit Sub"
'FG7.Rows = 1
Dim ii As Long
Dim j As Long
Dim jj As Long
Dim jjj As Long
    If GrdIqar.rows <= 2 Then
        If Trim(GrdIqar.TextMatrix(GrdIqar.rows - 1, GrdIqar.ColIndex("BranchName"))) = "" Then
            GrdIqar.rows = GrdIqar.rows - 1
        End If
    End If
    For ii = 0 To ListBranchSelected.ListCount - 1
        For j = 0 To ListAqarSelected.ListCount - 1
            For jj = 0 To ListUnitTypeSelected.ListCount - 1
                For jjj = 0 To ListUnitNoSelected.ListCount - 1
                    If chkEmpItem2(val(ListBranchSelected.ItemData(ii)), val(ListAqarSelected.ItemData(j)), val(ListUnitTypeSelected.ItemData(jj)), val(ListUnitNoSelected.ItemData(jjj)), val(ListUnitNoSelected2.ItemData(jjj))) Then
                        GrdIqar.rows = GrdIqar.rows + 1
                        'salim here****************************************
                        GrdIqar.TextMatrix(GrdIqar.rows - 1, 0) = GrdIqar.rows - 1
                        GrdIqar.TextMatrix(GrdIqar.rows - 1, GrdIqar.ColIndex("BranchID")) = ListBranchSelected.ItemData(ii)
                        GrdIqar.TextMatrix(GrdIqar.rows - 1, GrdIqar.ColIndex("BranchName")) = ListBranchSelected.List(ii)
                        GrdIqar.TextMatrix(GrdIqar.rows - 1, GrdIqar.ColIndex("Iqar")) = ListAqarSelected.ItemData(j)
                        GrdIqar.TextMatrix(GrdIqar.rows - 1, GrdIqar.ColIndex("IqarName")) = ListAqarSelected.List(j)
                        
                        GrdIqar.TextMatrix(GrdIqar.rows - 1, GrdIqar.ColIndex("unittype")) = ListUnitTypeSelected.ItemData(jj)
                        GrdIqar.TextMatrix(GrdIqar.rows - 1, GrdIqar.ColIndex("unittypeName")) = ListUnitTypeSelected.List(jj)
                        'salim here****************************************
                        
                        GrdIqar.TextMatrix(GrdIqar.rows - 1, GrdIqar.ColIndex("UnitNo")) = ListUnitNoSelected.ItemData(jjj)
                        GrdIqar.TextMatrix(GrdIqar.rows - 1, GrdIqar.ColIndex("UnitNoName")) = ListUnitNoSelected.List(jjj)
                        
                        GrdIqar.TextMatrix(GrdIqar.rows - 1, GrdIqar.ColIndex("DiscountPercent")) = txtDiscountPercent
                        GrdIqar.TextMatrix(GrdIqar.rows - 1, GrdIqar.ColIndex("FromDate")) = txtDateC(0)
                        GrdIqar.TextMatrix(GrdIqar.rows - 1, GrdIqar.ColIndex("ToDate")) = txtDateC(1)
                        
                        
                    End If
                Next jjj
            Next jj
        Next j
    Next ii

End Sub


Private Sub Command5_Click()

Dim ss0 As String, ss1 As String, ss2 As String, ss3 As String, ss4 As String, ss6 As String, ss7 As String, ss8 As String, ss9 As String, ss5 As String
Dim ss10 As String, ss11 As String, ss12 As String, ss13 As String, ss14 As String, ss15 As String, ss16 As String, ss17 As String, ss18 As String, ss19 As String, ss20 As String
ss0 = "mspcV3KFhqJ1QQgCpXjBBwYmW8EK1ihYAQdPqFHBCs6pTcEMTCpSQQtLqgyBF0ssV0EU0i1gwRNLrlSBDUcvUQEHxLAQwRLEMQ5BDz40XEEKvbRZwQo8t3TBPCA6cIEatDtqgRWqPhtBCrTAKgEDOsAYQQ5GwGxBDp1AdkEHl8E1gQc4QncBCRpCbkEJHURNwRAszSOBCzrNKoENMU8+AQeh02RBBZLUJoEXRBeErRwAEEVGRv//////Sk1TV15qcwMIABBGR0ZF/0ZFRUZIS01TYXMFDgEQQf////////9QVVxjb3UDBwAQRkVCPT1BQ0JCQ0VHSlZ2ChILEGBocXUBBAAQRkNBOjs/Pz8/QD8/PjsgFBYMEGtvcnQBABBGQ0E5Ozs9Ozs7OTg1MCMaFw0PcHJzABBFQj86Ojs5ODg4NzQvKSEbGAAA/wEQPzs9Ozo3NTU0MCkkIRwaGQAA/wEQQD0+Ojk1MjAtKCMgHRsZGQAA/wIQPz46NjEtLCgkIR8cGxkYAAD/AhA/PTcyLSonJSEfHBsZGBYAAP8DDz42LickIiEdHBoZGBcAAP8DDkM3KyEfHBwbGhkZFwQNPicdGxoaGhoaGQ=="
ss1 = "mspcV3KFhqJ1QQgCpXjBBwYmW8EK1ihYAQdPqFHBCs6pTcEMTCpSQQtLqgyBF0ssV0EU0i1gwRNLrlSBDUcvUQEHxLAQwRLEMQ5BDz40XEEKvbRZwQo8t3TBPCA6cIEatDtqgRWqPhtBCrTAKgEDOsAYQQ5GwGxBDp1AdkEHl8E1gQc4QncBCRpCbkEJHURNwRAszSOBCzrNKoENMU8+AQeh02RBBZLUJoEXRBeErRwAEEVGRv//////Sk1TV15qcwMIABBGR0ZF/0ZFRUZIS01TYXMFDgEQQf////////9QVVxjb3UDBwAQRkVCPT1BQ0JCQ0VHSlZ2ChILEGBocXUBBAAQRkNBOjs/Pz8/QD8/PjsgFBYMEGtvcnQBABBGQ0E5Ozs9Ozs7OTg1MCMaFw0PcHJzABBFQj86Ojs5ODg4NzQvKSEbGAAA/wEQPzs9Ozo3NTU0MCkkIRwaGQAA/wEQQD0+Ojk1MjAtKCMgHRsZGQAA/wIQPz46NjEtLCgkIR8cGxkYAAD/AhA/PTcyLSonJSEfHBsZGBYAAP8DDz42LickIiEdHBoZGBcAAP8DDkM3KyEfHBwbGhkZFwQNPicdGxoaGhoaGQ=="
ss2 = "mspc11FQno5ZwQYZk3ECCJuUXYEDlZpAgQ6DGm+CBZgbYkEIkxxsgQMioDTBBn4hbAEFJqFzwgeZo4KCBJ8kS8EIICduAgSfrlQCB5yvboEGLbAvwRFws0qCC5i5VUIELjs7ARuDu19BBKcEEHB1BQ8YHyMmJygpKSgEEG91BREaICYpKysrKioEEHF1BAwUGyAjJCYnKCgEEG50BBIcIyktLi0tLCsFEHYDCRAYHCAiJCYoKAQQbHMDEyAnLC4vLi0rKwUQAQQIDxUZHCAiJScoBRByBBckKy4wMC4sKykFEAMFCg8VGBsdIiUoKQYPDCErMDExLy0qKAYQBw4SFhgbHyIlKSoGDicvMjMzMS8tLAcPERMYGBsdIiUpCA0zNDMxLiwJDhcXGhwhJQkLNDMvAAD/AAD/AAD/AAD/AAD/AAD/AAD/"
ss3 = "mspcF4WNiRVagQiBGDWBCwaaSAEFgJ5rwQYQoFABBYAiQIEDBKMmQQhtoz7BBHonCwEO4CchwQnnqCaBCWspcgEJhy1ZwQYNsFfBCYEzMUENZThbwgyKulDBCxS8doEGIT0pQQ1SvU3BDhW/KUEQ0j9fwQePQjWBE1JFYEEFl8krQQxCyl+BBZpOMkEHtNFRgQad02lBBZpTcUEJmtRTAQYf1FkBBR5UOkETq9UlARi01XBBCBgfJLMQABBgYWNlaW1ydQIGCAoNDxETFAAQXF1fYmZrcXYDBggMDxMVFxgAEGVlZmhtcXR2AgUHCg0PEhQVABBYWVxfYmdudAMGCg8UFxgaGwAPamhoanBzdQEDBQgLDQ8PEQAQVldYW19janMCBwsRGBocHBwADm5ra25ydXcCBAYICg0ODwAQUlNVV1tfZW8CCQ8VGh0fHx8BDm5vcHN1dwMFBwkLDA0NABBOT1FSVVtibAELExgdISIhIAINcnF0dXcEBwgJCwwNABBKS0tNUFRbZXcNFxsgJCQjIQcLBAgKCw0AEEhISEhKTVJeAQ8aHyMkJCMiAAD/ABBEQ0NEREVHTw8YHyIkJSMiIQAA/wAQQUBAQUA/Pj0sISAiIiMhISAAAP8AED0+Pj8+Ojg0KiMgICAgHx8cAAD/AA85Ojo5NjMxLiYiICAfHRwfAQ42NjQuLCwnISAhISAdHQ=="
ss4 = "mspc14eNhRVwwQaMGVZBCIucQUEGhx93gQuIInuBFoOiHMEKcKJIwQOBI2SBBxMkNkECCyU2wQWApxcBCAAoG0EIbypyAQONrmYBCIuyTAEGEzJ9gQYfNCOBC2s1SgEHiD1LQgiQvhgBC1Y+PUEMFcAXgQ/WQWaBBiHBPMEMGUJKwQiRxCLBFFXKTQEFmcpGwQeVyxcBDkLRLQJELtF8AguW0i/BJC9VJIESstZGQQaYV16BBpgh5JMLABBpaW1xdQIFCQsODxIUFhcZGgAQZGdqcHQBBQoNDxASFBcaGxwAEGxscHQBAwYICw4QEhMTFBYYABBhY2htcwEGCw4QEhQWGRwcHQAQb29ydwMFBwkLDQ8SExITFBYAEF1fZGlwdwcMDxMWGBkbHR0dARBydQEEBggKDQ8PEBERExQWABBbXGBkbHUGDREWGhwcHR8fHwMQAgQGCQ0PDxASExQVFhYAEFZXXGJqdAcPFRodICAgHx8fBQ8IDA8PERITFBUWFwAQUVJWXWRyBxIYHSEjIiEgICAGDgwODxEUFRUWFgAQS0xRVl5uBxUbICQlJCMhIiIIDA8SFBUVABBGR0lNVGkLGB8jJSUkIiIiIwAA/wAQQUJCREhDGxwiJCYlIyEhIiMAAP8AED8/Pj0+NSUgIiMjISAfHyAiAAD/ABA7Ozk3NS4nIiIhIB8dHBwfIwEPNzUxLionIyAfHRsaGhwd"
ss5 = "mspcV5C1iZZpgQWAmDWBB3AdUYEECKFSAQV8piyBCl2oWEILhKhzwQccq0tBCwkuS0ENDS9gQQmKsSgBDEszKQEQyzRhAQePNDbBFEs1W8EKi7wvQQo6vkaCVyY/f4EIj8B0AQqWQDeBB6zASYEjJEBmgQWTQG2BBpbAVsEGlMBewQWUwX7BBRBCdQEFE0JgwQMbwmfBBBnCb8EEGUNZAQUbQz+BEaxECMEKuMYJgQs6SA3BBDbIQsENnMhJwQibSS4BEazKI0EQsEowgRQmS0RBDB/LTEEKHssjgRMtTCcBFCnWFoEGoRlkjxECEFdYXF9lb3YFChAWGhsaGgEQUFJSVVhfaXUGDRMYHR0dGwIQWFxfZGtzdwMHDRMXGBgYARBJS0tOUVdhcwYPFRodHx0cAxBeYmhtdHcDBgoPExUWFwAQRUVFRkdKTlVwCBEZHR8fHRsDEF9jaW5zdgIEBwoPExUWABBCQUFBQkJER14NFRwfHx0bGQQPaG5xdHYCAwUHCw8RABBBQD8+Pj09OzMbGBsbGxoYFwUOcXJ0dgEDBQYICgAQPz07Ojo4NjMsIRsaGhkYFhQGDXR2dwIEBQYHABA7Ojk4OTczLiYfGxkZGBYUEgcMdHcDBAYHABA5NzY1NDAqJSEdHBwbGRYUEQAA/wAQNzUyMC0oIyEfHBsbGRcUEhAAAP8AEDMwLSsoJSIgHxwbGRcVEhAOAAD/ARAqKSclIiAdHBsZFhMRDw8OAg8nJSMhHRwbGhcUEQ8ODA=="
ss6 = "mspcV3eBghs3AQQEHkcBBH4gHUEDcqNCQQV6pFkBBhCrLcEEAi8wAQd0NFKBCBW1N8ILfjgMgQtWvTFBCQi+b4EEjz4/wQ2BQjsBDoDDDEELRMUOgQ7BRhwBF0TId0EKiklmwQeLSlyBDo/LZ4EFD0xWAQiOTF/BBA9NSIEMi00wwmAgTVjBBhLPSoEIE1A0QR4b0ylBEKdTdcEJh9V1wQ2B1z0BCY8SVPMLAA5naGxwc3YBBAcLERISERAADmJlam9ydQIFCg4REhMUEwENa21xc3V3AgUJDxAQEAAPYWNobnJ1AgcNERMUFRYVFQIMcHJ1dwEDBQcMDw8AEF1eY2lwdAIJEBUWFhcXFxYUAwt0dncCAwUHCQwAEFdYXWRtdAMLEhgZGBcYFxcWBAl3AQMDBAYAEFJUV11ocgMLExkZGRcXFhYUBwgBAwAQTE5RVmBuAwsUGRkYFhUVFRMAAP8AEEVHSExUZwMNFRkYFxQSEhMSAAD/ABBAQUFERlgJDxUXFhUREA8QEAAA/wAQOzs6OjkxGhIVFRQSDw4NEBEAAP8AEDc3NTMwKR8VFBMTEA8NDg8SAAD/ABA0NDIuKiMbFhQUExEODQwSGgEPMjAtJh0XFBUUExEPDQoO"
ss7 = "mspcl4yxh5hNAQkEmErBCW8eKMEJU6ZVgRF+KSoBDEKrKwENwas4ARVCLXjBEIqvcoELijBsgQiHMQdCBEExdYEID7JsgQYPNFBBKxq0V0EJiLRgwQkQNVrBBxA3R8ETpbtaAQiOu1PBC5M8TcEMlLxcwQoSPT7BD6E9VsEPFb5HgQyYv08BDBc/QMENG8BJQQUZwDSBFKfBEEEJMcI2gRIcwxgBBy1DNIETIEQnQQ2pxSnBDSVPJwEHmU97AR0ezxNBDCfPdEEVfs9PggmL0nGBOXFVdcETENdSgQqDWEkBCoMNZAwTABBGRkZGRkhLT1txBxEVFRQRDgAQQkJBQUBCQkVLcAkRFBIPDQsAEElJSktMUFRZZ3MFDxUWFRQRABA+Pj09Ozs6OjkmEhESEA8MCgAPTE1OUFJVWWFtdQUOFBYVFQAQOzs6OTg3NTIuIxYREBAPDQsBD1BRVVZbYWhwdgQLEhUVFQAPOjk4NjQzMCwnIBYSEhMRDQMOWVtgZmxwdQMIDxITAA85ODUyMTEuJyAbFhQSEhAPBQ1kam5xdQIGCg4ADzg2MzAuLSghGxgWFBIPDA0KCwIFAA83NDEtKiciHRoYFhQRDgsPAAD/ABA1Mi4pJSMgHBgWExEPDAkQGQAA/wAQMi8qJSIgHRkWExAODAoLGR0AAP8BEC0nIh0bGRUSDwwLCgoQGyAAAP8CECMdGBUUEA4KCQoMDxMaIAMPGBQREA4MCAcJDhMVFQ=="

ss8 = "mspcF4KpgyE/QQUAojtBBHYjPwEEAqNOQQR4JyYBBXAoYoELCylMAQR4qhfBBuWsF4EH564agghrsjhBBO6zN8EDdjQ7QQUCtTpBBXK5XUEIEDtDwgl5PxfBDFXAO0EJAsBFgQd6QX/BCYrDPYEJBMRMgQuAR0nBC37KGIEMQMt+ARyDTHVBBYXMKYEVRE0bQQy9zWqBDIhPZkELh89vgQoJ0F+BDYNRaQEHC9JYAQ2FUmNBBw3TUYEJh1NcwQYNVChCDC1ULEENtFVUAQgLVUNBLRdYQEEUqQvUtA8BDmttcHFydHV3AgQICgwKAQ9oaW1vcXN1dwMGCgwNDQ4CDW9xcnN1dncCBAcJCgAQY2Vnam5xc3UBBQgLDA4PEBADDHJzdHV3dwEEBQYAEF9gYmdscHJ1AgcKDQ4PERISBAp0dXV2dwEEABBcXWBkam5xdQMIDQ8QERMTEwYJdnZ3AQAQV1hbX2VrcHUECQ8SExMTExIAAP8AEFJUVVleZW50BAkQExQTEhITAAD/ABBNTk5SV15ocgMKEBMTEg8QEQAA/wAQSEhIS09UX28CCxASEQ8NDhEAAP8AEENCQkRFSFFnAgoPDw4MCw4TAAD/ABA9PT09PT49PgkLDQ0LCgoNFAAA/wAQODg3NjQ0MScWDw0NDAsLDBIBDzY0MzAvLCUbEg4PDgwKCQ=="
ss9 = "mspcV5XVhJhkwQoPoCBBDFKjR4EJBqRRQQ1+J09BDHqqIsEMQKwywRZCrSQBC7+tQEEm5a1yQQyHLj9BJ0SubMEKha90gQwIsGZBDoMxbwEJDbFewQ2HMmhBCAuzYYEIDTQywQqwNFTBDoM1LwEKLTVLQSgXtlZBBg03R8EXqblBARKkO1UBCYo8T4EPkL1YAQwPPUdBE5M+UoEPDz5BwQ2XPxEBEas/OQEQnz9LwRASwAwBDazAEkEQKUA0gRGiQETBCRZBPEEKFsI3QQ8awg/BDTHEE4EFLEQvgRQeRSLBDqfHJAEMIs5yQRF6z0vCConQI0EJmFJ3wSl6VG2BLnZUcYEeD1VDgQyH12kBMBcJxKwRABBGRkdHSk1TXm4CCxESEA4MCAAQQkJBQUNDR1FnAgoPDw0KCgcAD0lKTU5RVV1ncQEKERQTERAAED4+PT09PT49OAsMDQwMCQkHAQ9OUVNXXGNtcwMLERMTEhEAEDk5ODg3NTQwJxUNDAwNDAsIAg5WV11jaW90AgoQEhMRABA4Nzc1My8tKCIWDQ0PDw0LCQMOXWJobHB0AggNDxAQAA82NDIvLy0pIhoUEBAPDgwLBg1tcHUCBgkMDgAPNDMwLS0qIxsXExEPDgwKCgkLAQQGAA8wLy0rKSUdGBYUEQ8MCQcJAAD/ARAvKyYjIBsXFBIPDAkHBg4gAAD/AhAlIh8dGRUSDw0KBwYGFiIAAP8DECAcGhYRDw0KCAUGChwiAAD/AxAbGBUSDQoICAoKDw8ZIQQPFRIPCgcGCQ4PFhMW"
ss10 = "mspc14Otgx8dAQVtn0LBCHajDwEE4yYQwQloqTCBBeysMoEGcC9WgQcQMjyCC3c3NEEKALc/QQd4OBIBC1E6N4EKADtGQQyAPkJBDXpCb4EOgUMVAQs+RGUBFIfEF8ELvMQkwRU8RDTBJePFcAEeCMUxwSVARmBBDYjGZ4EPCEdagQeFyGIBCAvIUsEPhclcgQcLTEgBDoHMJYELrkw/wSsWzSMBCi1Pc4IYHU97gScbTzzBHqXRNkESolF5gRMe0nKBGIdUUMIMhdU9ARyLVXjCEpVWasIKENdBwRMTCJSSDgANY2NmaW5wcnYECAsMDhAADl5fYmdsb3J3BQoNDg8SFAAMZmZpa25xc3YCBgoMDgAPW1xfZWltcXYGDRAREhMTEAALaWptbnBxc3UBBAkMAA9WV1tfZWtxdwcPExQSEhERAQlsb29xc3R2dwIAEFJSVFhfZm93Bg8TFBIREBIOAwhwcnN1dncAD0tMTVJXYGx1Bg8UFBEPDhMFB3N1dgAPRkZGSU1UYXQHDxIRDQsLEwAA/wAPQUBBQUJFU3UIDg8NCwkJEQAA/wAQPTs7Ojo6OhMLDAwLCgkJEh8AAP8AEDg3NjUzMC0dDwwLDA0MDBchAAD/ABA2NDIvLSkmHA8MDg8NDA8cIAAA/wEQMC8tKiUgGBEPDQ4PDxIcJQIPLCsoIxsWEA8MDg4PExk="

ss11 = "mspcV4i1hRs8wQfom1rBBXSdPAELbZ1agQVwoy4BCuUkMIEHaCdQQQbsKFCBBXKpUgEGACoYAQlYKlJBB3CtdUEJDy0WgQnTLhXBB1MuHcEHWC8YAQfVsFpCCni0U8EKArVeAQd6tTDBC1G4VQEKArhkwQ5+vGEBDXzAM4EKPkFDwRU+QjYBC7xECkEHRcV5QQaDxnGBDIPHegEICUhrAQuFyHPBCAjJEcIIPUlmwQ2DSlwBLxZKbwEHC8tnAQkNTX8BEodOVEEQodBpgQaIUWLBCo5SbEEGDVJcwRSN02WBEhLVXgESEwr08RUAEFZZW15fYGRpbXB0AQUICwwNABBVVVZZW15hZ2tvcwEGCw4PDwEPXV9hY2Roam5xc3cCBQgLABBTUlJUV1ldYmhucgEIDhEREAIOYmNmZ2lrbnFzdgEECAAQUVBPUFNUV1xia3ICCQ8TExEEDGlqbG1wcXN1dwAQTEtLTE5PUVVcZG93CRATExIGC2xtb3FzdQAQSUhHR0hISk5TXWx3CRASEhEAAP8AEEZFRERDQ0RER01fdQgPDw4NAAD/ABBEQkA/Pz4+Pj8/QgkJDQsLCwAA/wAQQT89Ozo5ODc2NC4aDQwLCw0AAP8AED47OTg3NjQyLyokGg8MDA8RAAD/ARA5ODY0MjAtKyYhGBAODxIUAAD/AhA1NDEuLSwoIRsXERAOERIDDzIuLS0sJx8ZGBQTDxA="
ss12 = "mspcF4KRiJcuQQZ+n06BBhohJYEICSMygQmFJScBCQ2oN4EGiqoxwQyHqw9BFEsraoEElDILwg03sw4BDr00IcJIJbVaAQOPtVBBBpK1ZEEDjzVugQaPtjqBA442QoEIj7ZbwQQWtzIBBZC3UUEHE7dwQQsQNyWBFSC3Z4EFEzhDAQMXODyBBRc5NkEIGbkZwRGoPi3BBJO/JQEJmMBmAQiLQC+BBRrAJ8ELHFBiAQWBURkBBZHXU8IGfxrE7AcADlNVW2sDDBUaHB0cGxwcHwAPS0tQYgUOFxscHBsaGRoaGgAOW11kcQMLExgcHBwcHB0gAA9AQUJFFBMYGRoaGhkZGBkYAA5cYGlzBAoPFhobGxscHSAADzk4NzEhGRkYGRgYFxcWFhYBDWZvdQMHDhMXGRkZGhsADzQzMCkhGhkXGBYWFRYUEhECDXN2AwcMEBQXFxgZGgAPLy0pJSAbGRgYFxUVExEPDwMMdgIHCw4QFBYYGQAPJyYjIRwbGhkYFhQSEA8NDwULBwkLDRAUFwAOIiIgHxwaGRcVFBMPDAsLAAD/AA8hIB0cGhgWFBIREA0KCQkYAAD/AA8fHRwaGBYTEQ8ODQsIBwgfAAD/ABAcGxkYFRMQDgwMCgkHCBAfIwAA/wAQGRgXFRMQDgwKCQgHBwoZHyQBDxcVFBIPDQoJCAcGCA0fHQ=="
ss13 = "mspc13uVjZo4gRZUHVWBC4ugLgEORCAxgQ/GpEJBehSlQsFcISc5AQy/J0HBMTipREEiLCxYAQSfLWABBJ2tH8EQNK1nwQidLllBAySuM4ENqy5vgQeeL2GBAyKwMwEILbBAQQWnsUcBBKAxagECIjFxgQogskGBBSkzSIEEJbQLgQmsPimBBp/DHYELm0QSAQ6ZxA1BEBvFDIEMkscoAQqYyxnBCI5LNcIKkk0vgQmLzjGBChBVTAEKiNZJwhMiIzRpEAAOOjs+Pj47OTQqJCMkJSUlAA44ODk4NjQzLyglIyQkIyMBDj9AQEA/Pz4qICIlJygoAA81NTUzMC4tKiclJCMiICAgAg5BQkRFSlIGFx0kJyorAA8xMjIwLSwqKCYlJCMiIB8fAg5FRkpOVF52DxkhJSksAA8rLCwrKikoJiUkIyIgHx8fBA1MVFtldgsVHSMmAA4mJycmJiUlIyIhICAdHRwFDFxibAEKExkfAA4fISEiIiIiIR8dHRwcGxsGC2ZvAQgPFAANGBobHB0fHx0cGxoZGRkAAP8BDBcXGBgaGhoaGRgYGAAA/wIMFRQUFRcXFxYWFhYAAP8DEBAQERITFRUVFBQZ/yEiAAD/AxANDhESExQVEhQXGhgdIgQPERISExYWExMYGRgb"
ss14 = "mspcl3yRi5dHAQYQGEUBBogZSIEGFZpFAQiFmh+BC2giSAIJjKQVQQtVpmEBByMmFkEO1ac6AQwZKiEBE1UrRsEJlK9DgQeWMBbBDEI1HIEItrkhAQ64ukKBBZq7SkEHmLtYwQmdvFEBB508CkETtrxCwQQgvGLBCJm9SkEFID0KgRIvvgZBDzG+U4EHIL5ZgQcdviLBDqW/YkEFHT8qwQekwDGBA57BI4EIJ0IrAQYlwjLBBSTOEoEFmx/k7AoADk5PVFhhcwoWHCEkJCQjIwAOSUpNUllzDhkgIyYmJSUlAA5VVltgaXYJEhkdISIiISIADkVFRUdMMhkfIyUmJSQkJAANWVteY2x3CA4VGh0fHyAADkFBPz89MSMhIiQkJCMjIwANXl9iZ3ABBwsRFxocHSAADT8+Ozk1LSUiIiIhISAgAQxgZGlyAggLDxQXGBoADTc2MzEtKCQiISEgICAgAgtnbXQDBwoOERQWAA0vLSwqKCUjIiEgHR0cHQYIBwoNAA0tLCooJiQjIiEgHx0dHwAA/wAMJycmJiQjIiEfHRwcGwAA/wAMJCQjIyEgHx0cHBsbGwAA/wAQISEhIB8dHBsaGhkaGP//ISMAAP8BECAdHBsbGRgYGBcYGBgYHCMCDxsbGhoYFxYXFRgYGBgb"
ss15 = "mspcF3qNiBtUAQYPHlMBB4OfLEEMaCJPAggZJlUCCY+pIwELUylvAQciqyNBDtMrSEEMF64uwRJUsFRBCZOzUEEIlDUkgQxCOSrBCLi6OYJBLzw7QSUtPjBBD7S+QoEHlL5IgQadv1ABBpo/WEEHmb9fgQuav2fBCprAQwEFIkBKwQshwVFBCCDCWQEGIEJhgQceQmiBCRzDFEETNEQ4gQejRT9BBZ7HOUEHJkc/QQQh1CECB50hJG4OAA5RUVJVW2FsAQoTGB0iIyMADktMTVFVXGV2DRccISQkJQAOU1VXW19kbwEIDhQZHSAgAA5KSUpLT1ReAREaICMlJSUBDVlbX2NpcgMHDBAWGRsADkNERUZHS1ILGB8jJSUkJAINXmFnbXQDBwsPExUWAA5CQkJBQEE/KyAgIiQjIyMDDGJqcHYEBwoOEBIADj9AQD87OTQpIyEhISAhIQYKdgMHCQ0ADj07Ojc1My4mIiEhIB8fHwAA/wANNjUzLy0rJyQjIiEfHBwAAP8ADTMyMC0qKCUjIiIhHx0dAAD/AA0uLSspJyYkIiEgHx0cGwAA/wAQKCclJCMiIR8dHRwcGhgaISMAAP8AECUkIiEgIB0dGxsZGhkYGB0kAQ8iIR8dHRwcGhkXGBgYGRw="
ss16 = "mspcF36hipc7gQcQGjmBC4WcEkEKZqI8QgqKJDBBDBMmVgEHIaYJgQxSJzoBCYsnL4ENFqgKQQ3TKxaBElIsPEEJkS85wQiTMgyBDEK2EgEIuDciAkQuuCPBIyq6GIEPtDoxgQacOysBCJQ7QcEGmLxIQQWZPFBBB5k8LIEHID0zAQcgPVoBBpa9OoEFIL5jAQSVvgtBDbC+T0EFG75KwQUfP1pBAxu/DIEOLUAaQQ6jwSFBB6JBKYEFmsIawQolwyJBBiNDKcEEH1ALwQWaIVRtCAANVFddZ3YMFxwhIyQjIyMADU5QVV4BEBofIyUkJCQkAA1YXWNtAQoTGB0hIiEhIgANR0hLVAYXHSIkJSQjIyQADF5gZnACBw4UGh0fICAADUFBQUApHyAiIyQjIiEiAAtiZGpzAwcMERcaHB8ADT07OTQoIiAgICEgIB8gAQppbnUEBwsPFBcZAA03NTItJCEgIB8fHx0dHQMJdgQHCw8RFAANLS0qJyMhISAfHBwbGxsAAP8ADCgoJiMiISEgHx0cGxsAAP8ADCYmJSMiISAdHBsaGhkAAP8ADCQjIyEfHRwbGxoYFxcAAP8AECIhIB8cGxkYFxcVFRIU/yEjAAD/ABAfHRwcGhgWFhUUExMXGBgcIwEPGxsaGBcUFBMSEhQYGBgb"
ss17 = "mspcV4WphxxYgQYNHVaBBoefWAEHEh9XgQeDoS8BDGanWEIKiypzQQchqyYBDVMsJ8EP1SxLwQwWMDKBElSxWMELkrlbAQSauzpBPj68PQJLLj4/ASQqPzTBD7RARwEIlEBNAQedQDpBFbZAVUEHmUBdgQiXQGTBCplBbAEJmUFIgQUiQk9BCyBCFsEQusIbwRO4QlbBCB/DXkEGH8NsgQgdQ2bBBh5ELoEPqEQVwQ80xRgBEjHFHAESL8Y8AQehRkNBBJ7HLgEJKsg9QQclSUUBByFVJ8IHnR+kTg8ADlFRVFVcYGh0Bg4UGh8hIQAPTExPUVZcYnEGDxgdIiQkIwAOVVZYWV9ja3UFCQ8VGRwcAA9JSUpLUFVcbggSGiAkJSUkAQ5bXF5jaXB3BQgNERUYGgAPRkZGSElMUWgOFh8iJCQkJAINXmFnbHIBBQcLDhIUAA9BQkJDQ0NFORwdIiIjIyQjBAxqb3QBBAgLDQ8ADj8/QUA+PTkwJCAhICEgIQYLdHcEBgkKAA47PT07OTczKyQhICAfICAAAP8ADjY3NjQwLSsnIyEfHx0cHAAA/wAPNDQxLSooJiMiISEgHx0dJQAA/wAPLy8tKignJSMhISAfHRwbIgAA/wAQKionJiQkIiEfHR0dHRsaICMAAP8AECYmJCMiISAdHBscHR0aGB0jAQ8kIiEfIB0cGhobHB0ZFxs="
ss18 = "mspcl4uxgBkzgQYCG0KBBHggGgEHcCBXwQcLIUBBBXahU4EOgKJXwQQNpwxBB2spcgEIC6otwQTsLS/BBHCxUgEHELIvQgcCMzjCCno4LsEK7rg8wQd6uQwBC1K5cgEGjLsxwQkCvELBC4BAQUELgMN2AQSBxA8BCz5EbUEIg8R4QQUIxR+BFz5FY4ELhUUSwQy9xW3BBQjHXYEIhcdkgQgIyFbBB4NJYUEEC8pPQQiFyldBBgtLSMEGh8tRwQYLTUtBBwvNOsEkFlIyAQ2h1EcBBohVQEEMjtVLgQkQ10NBDg8K5FINABBkZmptcHN1AQUJCw0NDgwLCQAQYGNnbG9ydgMHCw0ODw8ODQsAD2lpbG5xc3V3BAgLDAwLCgkAEF5gZWpucnYECQ0PEBAREA8NAQ5sb3BydHV3AgYKCwoJCQAQW1xfZWtwdgQKDxISEREPDw0CDXBxc3V2AQIFCQkICAAQVVZZX2dvdgQLEBMSERAODgwDDXJ0dncBAgUHBwUGABBPUFNXYGp0AwoQEhIPDgwLCwQIdHZ2dwEAEEhJS09WY3IDCxAREA4MCgkIBgd2dQAQQkJDREdTbAMKDg8NCwkICAcAAP8AED09PT0+Pg4JCw0MCwkHBwcHAAD/ABA4Nzc1MzAkFA0MCwsJCAcICQAA/wAQNTMyLysmIRcPDQwMCQcGBwoAAP8BEC4tKigjHRUQDw0LCQcGBwoCDyooJiEaEg8PDQsJBwYD"
ss19 = "mspcF6Ddhw9gAQiHk0jBCIOWIcEKBJg2AQR+mGwBCIWcWkEHDx0/wQJ+IC+BAwSgZ8EDhyIUQQtooi5BBnolYYEHgyYRwQzlJxXBCmmpdwEHGipHwQQLrkhBBoCyIUENYbVMwgyHOGgBCB64QYEKD7tNAQuIO0CBDRA8HIEKTj4dwQ3Ov0+BCIvBKUETT8JMwQiNSSFBCjjJOIJ37Es2gSY4TSmBB65OPAEgJs5nAQiVznFBCZNOSYEFlU5ZgQWTTmGBCJROUsEFk1A3gRmwUGmBDBdRTAEJG9FSAQMZ0VsBBBvRYgEFGVFzAQ4NUTFBD6zRRUEFHFQlwQ6r1TKBEZ9VLMELpdY6QQuf1kJBB5jWKsEKH9hDQQcWG2RyDgAQYGJmam9zdwMGCAsPExYYGBgAEF1fYmdtcncDBgoPExYYGRkZABBhZWpucXUBAwYICgwPExYXFwAQWVxfY2lwdgMGDBMXGBkaGhoAEGNqb3J0dgEDBggKDA8SFRcXABBWV1teZG11AwcPFRobHBscGwEPbnFzdQEDBAYICgwMDhEWABBSU1ZZX2hzAwsSGBwdHx0dHQIOcnR2AgQFBggKDAwODwAQTU5QVFhgbwINFRofICAfHyADDnQBBAYHCAkKCwsNDwAQR0hKTlFXZwMPGB0gICAfHx8EDgIFBwgJCwsODxASABBDQ0RFSExcBhMbHyEgHx8dHQUNBQcKDA4NDw8RABBAQEA/QEFAIBkbHR8dHBwbGgAA/wEQOzs6OTg1Jh0bGxwbGxocHQAA/wEQOTk4NTQvJh8bGxoZGBgbHwAA/wIPOTYvKiUgHBsaGxkXFRoDDjEqJSEdGxwaGxkWFQ=="
ss20 = "mspcF395hRQ6AQ2FGyuBC3afWYEJhaAswgYCI0WBBYAoacEGEClNwQN+rDuBBAatIUEIay54gQaKrzxCBnsyGwEH6LIiwQZrM29BCIe4UwEEDbtUgQiBvisBDWXDVkIMisVIQQwQxnKBByJIR0EOE0gjgQtRyiQBDdJLXIEJk00wgRNU0FtBBJbQRYEXF1QmwQtCVFHBCZNVW8EFmx30NQ8AD2xrbG9zdgEDBQYJCw0PDg4AEGppaW1xdHYBBAYJCw4QERITAA9ubm9ydXcCBAYHCQoNDg0NABBoZmdpbnF1AQQGCQoNDxIVGAAPcXFydHV3AgUICQoLDA0NDQAQYmJjZ2twc3cEBwoMDhEUGBoADnV0dHV2AQMGCQsLDQ4PDwAQXl9gY2htcwEFBwoOERQWGhwBDXR1dgECBAYJDA4PDw8AEFtcXWBkanF3BQkNEhYYGhwfAgx2dwMFBwkLDA4PDwAQV1hZXWBmbnYFCg8VGhscHR0DCgIFBwkMDg0OABBUVFZYXWJqdQYPFRkdHyAfHwAA/wAQUFBSVFheZnMHERgcISIiISAAAP8AEEtLTE5SVl9xCBQbICMkJCMiAAD/ABBGR0dIS05WbgwXHyMkJCQlJQAA/wAQQkJDRERGSkMYHSIkJSQjJCgBD0BBQkFCQzcgISQmJSMiIw=="


Static xx As Long
xx = xx + 1
Select Case xx
Case 1
    txtFingerPrint2 = ss0
Case 2
    txtFingerPrint2 = ss1
Case 3
    txtFingerPrint2 = ss2
Case 4
    txtFingerPrint2 = ss3
Case 5
    txtFingerPrint2 = ss4
Case 6
    txtFingerPrint2 = ss5
Case 7
    txtFingerPrint2 = ss6
Case 8
    txtFingerPrint2 = ss7
Case 9
    txtFingerPrint2 = ss8
Case 10
    txtFingerPrint2 = ss9
Case 11
    txtFingerPrint2 = ss10
Case 12
    txtFingerPrint2 = ss11
Case 13
    txtFingerPrint2 = ss12
Case 14
    txtFingerPrint2 = ss13
Case 15
    txtFingerPrint2 = ss14
Case 16
    txtFingerPrint2 = ss15
Case 17
    txtFingerPrint2 = ss16
Case 18
    txtFingerPrint2 = ss17
Case 19
    txtFingerPrint2 = ss18
Case 20
    txtFingerPrint2 = ss19

End Select
If xx = 20 Then xx = 0

End Sub

Private Sub Command4_Click()
Dim X As Integer
Dim str As String


X = MsgBox(" √ﬂÌœ «€·«ﬁ «·ÌÊ„", vbInformation + vbYesNo)

If X = vbYes Then
str = "update  TblEmpDataInOut  set timeout = timein  where timeout is null "
 Cn.Execute str

   s = "Select * from TblEmpDataInOut Where RecordDate"
       s = s & "  in ( select max(RecordDate) as RecordDate from TblEmpDataInOut where timeout is null )"
    '   My_SQL = s
        loadgrid s, GrdEmp, True, False
        
End If



End Sub

Private Sub Command7_Click()
CommonDialog1.CancelError = True
  On Error GoTo errHandler
  'Set the Flags property
  CommonDialog1.Flags = cdlCCRGBInit
  ' Display the Color Dialog box
  CommonDialog1.ShowColor
  ' Set the form's background color to selected color
  lblServiceColor.backcolor = CommonDialog1.Color
  Exit Sub
errHandler:
End Sub

 




Private Sub cmbPaymentClass_Change()
If Me.TxtModFlg2(mIndex).text = "N" Or Me.TxtModFlg2(mIndex).text = "E" Then
    If val(cmbPaymentClass.text) <> 0 Then
        XPTxtVal = cmbPaymentClass.text
      Else
      XPTxtVal = 0
      
    End If
End If
End Sub

Private Sub Frame11_DragDrop(Source As Control, X As Single, Y As Single)

End Sub



Private Sub DcCustmer_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
'         FrmCustemerSearch.show vbModal
'        FrmCustemerSearch.SearchType = 6549
        Frame1(0).Visible = True
    End If

End Sub

Private Sub grd_Click()
    cmbPaymentClass.BoundText = val(grd.ColKey(grd.Col))
    txtAmountCash = ""
    txtAmountVisa = ""
    Dim i As Long
    Dim mCol As Long
    mCol = grd.Col
    
    For i = 0 To grd.Cols - 1
       
        grd.Col = i
        'grd.CellBackColor = vbWhite
          grd.CellBackColor = val(grd.ColEditMask(i))
    Next
    grd.Col = mCol
'    grd.CellBackColor = vbBlue
    lblClassCat.backcolor = IIf(grd.CellBackColor = 0, vbWhite, grd.CellBackColor)
    lblClassCat.Caption = cmbPaymentClass.text
End Sub

Private Sub grd_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
 ' cmbPaymentClass.BoundText = val(grd.TextMatrix(1, grd.Col))
 Cancel = True
End Sub

Private Sub GrdFinger2_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    With Me.GrdFinger2

        Select Case .ColKey(Col)

        Case "Finger"
            mFinger = Row
            ISButton5_Click
        End Select
        
       End With
End Sub

Private Sub ISButton4_Click()
    Dim X As String
    If TxtSerial1(mIndex).text = "" Then Exit Sub
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

    DBPix201.ImageSaveFile (system_path & "\" & SystemOptions.ImagesPath & "\" & TxtSerial1(mIndex).text & ".JPG")
End Sub

Private Sub ISButton5_Click()
If Me.TxtModFlg2(mIndex).text = "N" Or Me.TxtModFlg2(mIndex).text = "E" Then
        
        If ZKFPEngX1.InitEngine <> 0 Then
            txtFingerPrint.Enabled = False
        
        
        End If
        
        s = "Select * from TblEmpData Where HafizaNo = N'" & Trim(txtHafizaNo) & "' and Id <> " & val(TxtSerial1(mIndex))
        Set rsDummy = New ADODB.Recordset
        rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly
        If Not rsDummy.EOF Then
            ZKFPEngX1.EndEngine
            Label13.Caption = "€Ì— „ ’·"
            MsgBox "·«Ì„ﬂ‰ «·Õ›Ÿ , »—Ã«¡ «œŒ«· —ﬁ„ ÂÊÌ… ¬Œ— ·«‰ Â–« «·—ﬁ„ „ﬂ—— „⁄ «·„ÊŸ› " & rsDummy!EmpName & "", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            Exit Sub
        
        End If

        mSenesor = True
        FingerCount = 0
        fpcHandle = ZKFPEngX1.CreateFPCacheDB
        
        ZKFPEngX1.BeginEnroll
        If ZKFPEngX1.IsRegister Then
            ZKFPEngX1.CancelEnroll
        End If
        If ZKFPEngX1.InitEngine = 0 Then
        
        End If
        txtFingerPrint.Enabled = True
        ZKFPEngX1.SensorIndex = 0
        lblProgressFinger.Caption = ""
        lblFingerStatus.Tag = ""
        Label13.Caption = "„ ’·"
Else
    Label13.Caption = "€Ì— „ ’·"
    mSenesor = False
End If
End Sub

Private Sub Label11_Click()
    ListAqarSelected.Clear
End Sub

Private Sub Label12_Click()
On Error Resume Next
If ListAqarSelected.ListIndex > -1 Then
        ListAqarSelected.RemoveItem ListAqarSelected.ListIndex
    End If

End Sub

Private Sub Label14_Click()
On Error Resume Next
If ListBranchAll.ListIndex = -1 Then Exit Sub
    
Dim i As Long

For i = 0 To ListBranchAll.ListCount - 1
    If ListBranchAll.Selected(i) Then
        ListBranchSelected.AddItem ListBranchAll.List(i)
        ListBranchSelected.ItemData(ListBranchSelected.NewIndex) = ListBranchAll.ItemData(i)
        
    End If
Next
    
    


FillMylist2 False, True, False, False
End Sub

Private Sub Label15_Click()
On Error Resume Next
    Dim i As Integer
    ListBranchSelected.Clear

    For i = 0 To ListBranchAll.ListCount - 1
        ListBranchSelected.AddItem ListBranchAll.List(i)
        ListBranchSelected.ItemData(i) = ListBranchAll.ItemData(i)
    Next i
    
    
FillMylist2 False, True, False, False
End Sub

Private Sub Label16_Click()
On Error Resume Next
    ListBranchSelected.Clear
End Sub

Private Sub Label17_Click()
On Error Resume Next
If ListBranchSelected.ListIndex > -1 Then
        ListBranchSelected.RemoveItem ListBranchSelected.ListIndex
    End If
End Sub

Private Sub Label18_Click()
On Error Resume Next

If ListUnitTypeAll.ListIndex = -1 Then Exit Sub
    
Dim i As Long

For i = 0 To ListUnitTypeAll.ListCount - 1
    If ListUnitTypeAll.Selected(i) Then
        ListUnitTypeSelected.AddItem ListUnitTypeAll.List(i)
        ListUnitTypeSelected.ItemData(ListUnitTypeSelected.NewIndex) = ListUnitTypeAll.ItemData(i)
        
    End If
Next
    
    
FillMylist2 False, False, False, True
End Sub

Private Sub Label19_Click()
On Error Resume Next
    Dim i As Integer
    ListUnitTypeSelected.Clear

    For i = 0 To ListUnitTypeAll.ListCount - 1
        ListUnitTypeSelected.AddItem ListUnitTypeAll.List(i)
        ListUnitTypeSelected.ItemData(i) = ListUnitTypeAll.ItemData(i)
    Next i
    FillMylist2 False, False, False, True

End Sub

Private Sub Label20_Click()
  On Error Resume Next
    ListUnitTypeSelected.Clear
End Sub

Private Sub Label21_Click()
On Error Resume Next
If ListUnitTypeSelected.ListIndex > -1 Then
        ListUnitTypeSelected.RemoveItem ListUnitTypeSelected.ListIndex
    End If

End Sub

Private Sub Label22_Click()
On Error Resume Next
If ListUnitNoSelected.ListIndex > -1 Then
        ListUnitNoSelected.RemoveItem ListUnitNoSelected.ListIndex
        ListUnitNoSelected2.RemoveItem ListUnitNoSelected.ListIndex
    End If

End Sub

Private Sub Label23_Click()
    ListUnitNoSelected.Clear
    ListUnitNoSelected2.Clear
End Sub

Private Sub Label24_Click()
On Error Resume Next
    Dim i As Integer
    ListUnitNoSelected.Clear
    ListUnitNoSelected2.Clear
    For i = 0 To ListUnitNoAll.ListCount - 1
        ListUnitNoSelected.AddItem ListUnitNoAll.List(i)
        ListUnitNoSelected.ItemData(i) = ListUnitNoAll.ItemData(i)
        
        ListUnitNoSelected2.AddItem ListUnitNoAll2.List(i)
        ListUnitNoSelected2.ItemData(i) = ListUnitNoAll2.ItemData(i)
        
    Next i


End Sub

Private Sub Label25_Click()
On Error Resume Next
'    If ListUnitNoAll.ListIndex = -1 Then Exit Sub
'    ListUnitNoSelected.AddItem ListUnitNoAll.List(ListUnitNoAll.ListIndex)
'    ListUnitNoSelected.ItemData(ListUnitNoSelected.NewIndex) = ListUnitNoAll.ItemData(ListUnitNoAll.ListIndex)
'
'
'    ListUnitNoSelected2.AddItem ListUnitNoAll2.List(ListUnitNoAll2.ListIndex)
'    ListUnitNoSelected2.ItemData(ListUnitNoSelected2.NewIndex) = ListUnitNoAll2.ItemData(ListUnitNoAll2.ListIndex)
'

If ListUnitNoAll.ListIndex = -1 Then Exit Sub
    
Dim i As Long

For i = 0 To ListUnitNoAll.ListCount - 1
    If ListUnitNoAll.Selected(i) Then
        ListUnitNoSelected.AddItem ListUnitNoAll.List(i)
        ListUnitNoSelected.ItemData(ListUnitNoSelected.NewIndex) = ListUnitNoAll.ItemData(i)
        
        ListUnitNoSelected2.AddItem ListUnitNoAll2.List(i)
        ListUnitNoSelected2.ItemData(ListUnitNoSelected2.NewIndex) = ListUnitNoAll2.ItemData(i)
    End If
Next
    

End Sub

Private Sub Label4_Click()
On Error Resume Next

If ListAqarAll.ListIndex = -1 Then Exit Sub
    
Dim i As Long

For i = 0 To ListAqarAll.ListCount - 1
    If ListAqarAll.Selected(i) Then
        ListAqarSelected.AddItem ListAqarAll.List(i)
        ListAqarSelected.ItemData(ListAqarSelected.NewIndex) = ListAqarAll.ItemData(i)
        
    End If
Next
    
End Sub

Private Sub Label9_Click()
   On Error Resume Next
    Dim i As Integer
    ListAqarSelected.Clear

    For i = 0 To ListAqarAll.ListCount - 1
        ListAqarSelected.AddItem ListAqarAll.List(i)
        ListAqarSelected.ItemData(i) = ListAqarAll.ItemData(i)
    Next i

End Sub

Private Sub lblexit_Click(Index As Integer)
Frame1(0).Visible = False
End Sub

Private Sub ntxtLetter1_Change()
FilltxtBord
txtLetter1 = GerNoCarEn(ntxtLetter1)
End Sub

Private Sub ntxtLetter1_GotFocus()
SwitchKeyboardLang LANG_ARABIC
End Sub

Private Sub ntxtLetter1_KeyPress(KeyAscii As Integer)

ntxtLetter1.text = ""
If Len(ntxtLetter1.text) > 0 Then
KeyAscii = 0
End If
Select Case CHR(KeyAscii)
    Case 0 To 9
        KeyAscii = 0
    Case 8
        Exit Sub
    Case Else
        ntxtLetter2.SetFocus
End Select
End Sub

Private Sub ntxtLetter2_GotFocus()
SwitchKeyboardLang LANG_ARABIC
End Sub

Private Sub ntxtLetter3_GotFocus()
SwitchKeyboardLang LANG_ARABIC
End Sub

Private Sub ntxtNum1_KeyPress(KeyAscii As Integer)
ntxtNum1.text = ""
If Len(ntxtNum1.text) > 0 Then
KeyAscii = 0
End If
If Not (CHR(KeyAscii) >= 0 And CHR(KeyAscii) <= 9) Then
KeyAscii = 0
Else
        ntxtNum2.SetFocus
End If
End Sub

Private Sub ntxtNum2_KeyPress(KeyAscii As Integer)
ntxtNum2.text = ""
If Len(ntxtNum2.text) > 0 Then
KeyAscii = 0
End If
If Not (CHR(KeyAscii) >= 0 And CHR(KeyAscii) <= 9) Then
KeyAscii = 0
Else
        ntxtNum3.SetFocus
End If
End Sub

Private Sub ntxtNum3_KeyPress(KeyAscii As Integer)
ntxtNum3.text = ""
If Len(ntxtNum3.text) > 0 Then
KeyAscii = 0
End If
If Not (CHR(KeyAscii) >= 0 And CHR(KeyAscii) <= 9) Then
KeyAscii = 0
Else
        ntxtNum4.SetFocus

End If
End Sub

Private Sub ntxtNum4_KeyPress(KeyAscii As Integer)
ntxtNum4.text = ""
If Len(ntxtNum4.text) > 0 Then
KeyAscii = 0
End If
If Not (CHR(KeyAscii) >= 0 And CHR(KeyAscii) <= 9) Then
KeyAscii = 0
End If
End Sub

Private Sub ntxtLetter2_Change()
FilltxtBord
txtLetter2 = GerNoCarEn(ntxtLetter2)
End Sub

Private Sub ntxtLetter2_KeyPress(KeyAscii As Integer)

ntxtLetter2.text = ""
If Len(ntxtLetter2.text) > 0 Then
KeyAscii = 0
End If
Select Case CHR(KeyAscii)
    Case 0 To 9
        KeyAscii = 0
    Case Else
        ntxtLetter3.SetFocus
End Select

End Sub

Private Sub ntxtLetter3_Change()
FilltxtBord
txtLetter3 = GerNoCarEn(ntxtLetter3)
End Sub

Private Sub ntxtLetter3_KeyPress(KeyAscii As Integer)

ntxtLetter3.text = ""
If Len(ntxtLetter3.text) > 0 Then
KeyAscii = 0
End If
Select Case CHR(KeyAscii)
    Case 0 To 9
        KeyAscii = 0
    Case Else
        ntxtNum1.SetFocus
End Select

End Sub

Private Sub ntxtLetter4_Change()
FilltxtBord
txtLetter1 = GerNoCarEn(ntxtLetter1)
End Sub

Private Sub ntxtNum1_Change()
FilltxtBord
txtNum1 = ntxtNum1
End Sub


Private Sub ntxtNum2_Change()
FilltxtBord

txtNum2 = ntxtNum2
End Sub


Private Sub ntxtNum3_Change()
FilltxtBord
txtNum3 = ntxtNum3
End Sub


Private Sub ntxtNum4_Change()
FilltxtBord
txtNum4 = ntxtNum4
End Sub


Private Sub Text6_Change()

End Sub

Private Sub Text3_Change()

End Sub

Private Sub Timer1_Timer()
txtTimeIn.value = Time
End Sub

Private Sub Timer2_Timer()
CmdRefresh_Click
End Sub

Private Sub Timer3_Timer()
If mIndex = 8 Then
ISButton2_Click
End If
End Sub

Private Sub txtAmountCash_GotFocus()
If Me.TxtModFlg2(mIndex).text = "N" Or Me.TxtModFlg2(mIndex).text = "E" Then
    'txtAmountVisa = val(XPTxtVal) - val(txtAmountCash)
    txtAmountCash = val(txtTotalWithVat2) - val(txtAmountVisa)
End If

End Sub

Private Sub txtAmountLater_GotFocus()
If Me.TxtModFlg2(mIndex).text = "N" Or Me.TxtModFlg2(mIndex).text = "E" Then
    'txtAmountVisa = val(XPTxtVal) - val(txtAmountCash)
    txtAmountLater = txtTotalWithVat2
End If

End Sub

Private Sub txtAmountVisa_GotFocus()
If Me.TxtModFlg2(mIndex).text = "N" Or Me.TxtModFlg2(mIndex).text = "E" Then
    txtAmountVisa = val(txtTotalWithVat2) - val(txtAmountCash)
End If
End Sub

Private Sub TxtBoardNO_KeyPress(KeyAscii As Integer)
SwitchKeyboardLang LANG_ENGLISH
End Sub

Private Sub txtFingerPrint_Change()

If Me.TxtModFlg2(mIndex).text = "N" Or Me.TxtModFlg2(mIndex).text = "E" Then
    
    Static X As Long
'    If val(lblFingerStatus.Tag) = 100 Then
'        Dim MSGType As Integer
'
'        MSGType = MsgBox("Â–« «·„ÊŸ›  „ «œ—«Ã »’„« Â Â·  Êœ «⁄«œ…  ⁄ÌÌ‰ »’„« Â", vbQuestion + vbMsgBoxRtlReading + vbYesNo + vbMsgBoxRight, App.title)
'
'
'        If MSGType = vbYes Then
'            X = 0
'            lblProgressFinger.Caption = ""
'            lblFingerStatus.Tag = ""
'        End If
'    End If
     If val(lblProgressFinger.Caption) = 100 Then
        ZKFPEngX1.EndEngine
        MsgBox " „ «œ—«Ã «·»’„… »‰Ã«Õ"
        lblFingerStatus.Tag = "100"
        GrdFinger2.TextMatrix(1, GrdFinger2.ColIndex("Percent")) = 100
        btn_Save_Click mIndex
        
    
        Exit Sub
    End If
    If Trim(txtFingerPrint) <> "" Then
        If X > 10 Then X = 0
        X = X + 1
        If X > 10 Then X = 0
        Select Case mFinger
        Case 0, 1
            GrdFinger.TextMatrix(X, GrdFinger.ColIndex("FingerPrint")) = txtFingerPrint
            GrdFinger2.TextMatrix(1, GrdFinger2.ColIndex("Percent")) = val(GrdFinger2.TextMatrix(mFinger, GrdFinger2.ColIndex("Percent"))) + 10
            
        Case 2
            GrdFinger.TextMatrix(X, GrdFinger.ColIndex("FingerPrint2")) = txtFingerPrint
            GrdFinger2.TextMatrix(mFinger, GrdFinger2.ColIndex("Percent")) = val(GrdFinger2.TextMatrix(mFinger, GrdFinger2.ColIndex("Percent"))) + 10
        Case 3
            GrdFinger.TextMatrix(X, GrdFinger.ColIndex("FingerPrint3")) = txtFingerPrint
            GrdFinger2.TextMatrix(mFinger, GrdFinger2.ColIndex("Percent")) = val(GrdFinger2.TextMatrix(mFinger, GrdFinger2.ColIndex("Percent"))) + 10
        Case 4
            GrdFinger.TextMatrix(X, GrdFinger.ColIndex("FingerPrint4")) = txtFingerPrint
            GrdFinger2.TextMatrix(mFinger, GrdFinger2.ColIndex("Percent")) = val(GrdFinger2.TextMatrix(mFinger, GrdFinger2.ColIndex("Percent"))) + 10
        Case 5
            GrdFinger.TextMatrix(X, GrdFinger.ColIndex("FingerPrint5")) = txtFingerPrint
            GrdFinger2.TextMatrix(mFinger, GrdFinger2.ColIndex("Percent")) = val(GrdFinger2.TextMatrix(mFinger, GrdFinger2.ColIndex("Percent"))) + 10
        Case Else
            GrdFinger.TextMatrix(X, GrdFinger.ColIndex("FingerPrint")) = txtFingerPrint
GrdFinger2.TextMatrix(mFinger, GrdFinger2.ColIndex("Percent")) = 100
        End Select
        
        
        lblProgressFinger.Caption = val(lblProgressFinger.Caption) + 10 & "%"
        If X = 10 Then lblFingerStatus.Tag = "100"
        
        If X > 10 Then X = 0
    End If
    
Else
   X = 0
End If
End Sub

Private Sub txtFingerPrint2_Change()
'cmbEmpName.BoundText = 0
If Trim(txtFingerPrint2) = "" Then Exit Sub
s = "Select * from TblEmpDataFingerPrint "
Set rsDummy = New ADODB.Recordset
rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly
If rsDummy.RecordCount = 0 Then Exit Sub
Label14.Visible = False
Do While Not rsDummy.EOF
    If ZKFPEngX2.VerFingerFromStr(Trim(txtFingerPrint2), Trim(rsDummy!FingerPrint & ""), False, True) Then
        cmbEmpName.BoundText = val(rsDummy!EmpID & "")
        Exit Sub
    
    End If
    rsDummy.MoveNext
    
Loop

rsDummy.MoveFirst
Do While Not rsDummy.EOF
    If ZKFPEngX2.VerFingerFromStr(Trim(txtFingerPrint2), Trim(rsDummy!FingerPrint2 & ""), False, True) Then
        cmbEmpName.BoundText = val(rsDummy!EmpID & "")
        Exit Sub
    
    End If
    rsDummy.MoveNext
    
Loop


rsDummy.MoveFirst
Do While Not rsDummy.EOF
    If ZKFPEngX2.VerFingerFromStr(Trim(txtFingerPrint2), Trim(rsDummy!FingerPrint3 & ""), False, True) Then
        cmbEmpName.BoundText = val(rsDummy!EmpID & "")
        Exit Sub
    
    End If
    rsDummy.MoveNext
    
Loop


rsDummy.MoveFirst
Do While Not rsDummy.EOF
    If ZKFPEngX2.VerFingerFromStr(Trim(txtFingerPrint2), Trim(rsDummy!FingerPrint4 & ""), False, True) Then
        cmbEmpName.BoundText = val(rsDummy!EmpID & "")
        Exit Sub
    
    End If
    rsDummy.MoveNext
    
Loop


rsDummy.MoveFirst
Do While Not rsDummy.EOF
    If ZKFPEngX2.VerFingerFromStr(Trim(txtFingerPrint2), Trim(rsDummy!FingerPrint5 & ""), False, True) Then
        cmbEmpName.BoundText = val(rsDummy!EmpID & "")
        Exit Sub
    
    End If
    rsDummy.MoveNext
    
Loop

Label14.Visible = True



'txtFingerPrint2.Text = ""
'
' If ZKFPEngX2.VerFingerFromStr(txtFingerPrint2, ss11, False, True) Then
'          MessageBox 0, "Verify success", "information", 0
'       Else
'       '   MessageBox 0, "Verify Failed", "information", 0
'       End If
'
'If ZKFPEngX2.VerFingerFromStr(txtFingerPrint2, ss1, False, True) Then
'          MessageBox 0, "Verify success", "information", 0
'       Else
'     '     MessageBox 0, "Verify Failed", "information", 0
'       End If
'
'
'       If ZKFPEngX2.VerFingerFromStr(txtFingerPrint2, ss2, False, True) Then
'          MessageBox 0, "Verify success", "information", 0
'       Else
'     '     MessageBox 0, "Verify Failed", "information", 0
'       End If
'
'
'If ZKFPEngX2.VerFingerFromStr(txtFingerPrint2, ss3, False, True) Then
'          MessageBox 0, "Verify success", "information", 0
'       Else
'     '     MessageBox 0, "Verify Failed", "information", 0
'       End If
'If ZKFPEngX2.VerFingerFromStr(txtFingerPrint2, ss4, False, True) Then
'          MessageBox 0, "Verify success", "information", 0
'       Else
'     '     MessageBox 0, "Verify Failed", "information", 0
'       End If
       
End Sub

Private Sub txtHafizaNo_KeyPress(KeyAscii As Integer)
 KeyAscii = KeyAscii_Num(KeyAscii, Me.txtHafizaNo.text, 1)
End Sub

Private Sub txtLetter1_Change()
txtLetter1.text = UCase(txtLetter1.text)
FilltxtBord
ntxtLetter1 = GerNoCarAR(txtLetter1)
End Sub

Private Sub txtLetter2_Change()
txtLetter2.text = UCase(txtLetter2.text)
FilltxtBord
End Sub

Private Sub txtLetter2_GotFocus()
SwitchKeyboardLang LANG_ENGLISH
End Sub

Private Sub txtLetter3_Change()
txtLetter3.text = UCase(txtLetter3.text)
FilltxtBord
End Sub


Private Sub txtLetter3_GotFocus()
SwitchKeyboardLang LANG_ENGLISH
End Sub

Private Sub txtLetter4_Change()
txtLetter4.text = UCase(txtLetter4.text)
FilltxtBord
End Sub


 


Private Sub txtLetter1_GotFocus()
SwitchKeyboardLang LANG_ENGLISH
End Sub

Private Sub txtLetter1_KeyPress(KeyAscii As Integer)

txtLetter1.text = ""
If Len(txtLetter1.text) > 0 Then
KeyAscii = 0
End If
Select Case CHR(KeyAscii)
    Case 0 To 9
        KeyAscii = 0
    Case 8
        Exit Sub
    Case Else
        txtLetter2.SetFocus
End Select
End Sub
Private Sub txtLetter2_KeyPress(KeyAscii As Integer)

txtLetter2.text = ""
If Len(txtLetter2.text) > 0 Then
KeyAscii = 0
End If
Select Case CHR(KeyAscii)
    Case 0 To 9
        KeyAscii = 0
    Case Else
        txtLetter3.SetFocus
End Select
End Sub
Private Sub txtLetter3_KeyPress(KeyAscii As Integer)

txtLetter3.text = ""
If Len(txtLetter3.text) > 0 Then
KeyAscii = 0
End If
Select Case CHR(KeyAscii)
    Case 0 To 9
        KeyAscii = 0
    Case Else
        txtNum1.SetFocus
End Select
End Sub
Private Sub txtLetter4_KeyPress(KeyAscii As Integer)
txtLetter4.text = ""
If Len(txtLetter4.text) > 0 Then
KeyAscii = 0
End If
Select Case CHR(KeyAscii)
    Case 0 To 9
        KeyAscii = 0
    Case Else
        txtNum1.SetFocus
End Select
End Sub

Private Function GerNoCarEn(ByVal mTxt As String) As String
    Select Case mTxt
    Case "√", "«", "¬"
        GerNoCarEn = "A"
    Case "»"
        GerNoCarEn = "B"
    Case "Õ"
        GerNoCarEn = "J"
    Case "œ"
        GerNoCarEn = "D"
    Case "—"
        GerNoCarEn = "R"
    Case "”"
        GerNoCarEn = "S"
    Case "’"
        GerNoCarEn = "X"
    Case "ÿ"
        GerNoCarEn = "T"
    Case "⁄"
        GerNoCarEn = "E"
    Case "ﬁ"
        GerNoCarEn = "G"
    Case "ﬂ"
        GerNoCarEn = "K"
    Case "·"
        GerNoCarEn = "L"
    Case "„"
        GerNoCarEn = "Z"
    Case "‰"
        GerNoCarEn = "N"
    Case "Â"
        GerNoCarEn = "H"
    Case "Ê"
        GerNoCarEn = "U"
    Case "Ï", "Ì"
        GerNoCarEn = "V"
    Case Else
        GerNoCarEn = ""
    End Select
    
End Function
Private Function GerNoCarAR(ByVal mTxt As String) As String
    mTxt = UCase(mTxt)
    Select Case mTxt
    Case "A"
        GerNoCarAR = "«"
    Case "B"
        GerNoCarAR = "»"
    Case "J"
        GerNoCarAR = "Õ"
    Case "D"
        GerNoCarAR = "œ"
    Case "R"
        GerNoCarAR = "—"
    Case "S"
        GerNoCarAR = "”"
    Case "X"
        GerNoCarAR = "’"
    Case "T"
        GerNoCarAR = "ÿ"
    Case "E"
        GerNoCarAR = "⁄"
    Case "G"
        GerNoCarAR = "ﬁ"
    Case "K"
        GerNoCarAR = "ﬂ"
    Case "L"
        GerNoCarAR = "·"
    Case "Z"
        GerNoCarAR = "„"
    Case "N"
        GerNoCarAR = "‰"
    Case "H"
        GerNoCarAR = "Â"
    Case "U"
        GerNoCarAR = "Ê"
    Case "V"
        GerNoCarAR = "Ï"
    Case Else
        GerNoCarAR = ""
    End Select
    
End Function

Private Sub TxtMobileNO_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtMobileNO.text, 1)
End Sub

Private Sub txtnBoardNo_KeyPress(KeyAscii As Integer)
    SwitchKeyboardLang LANG_ARABIC
End Sub

Private Sub txtNum1_Change()
FilltxtBord
End Sub

Private Sub txtNum1_KeyPress(KeyAscii As Integer)
txtNum1.text = ""
If Len(txtNum1.text) > 0 Then
KeyAscii = 0
End If
If Not (CHR(KeyAscii) >= 0 And CHR(KeyAscii) <= 9) Then
KeyAscii = 0
Else
        txtNum2.SetFocus
End If
End Sub

Private Sub txtNum2_KeyPress(KeyAscii As Integer)
txtNum2.text = ""
If Len(txtNum2.text) > 0 Then
KeyAscii = 0
End If
If Not (CHR(KeyAscii) >= 0 And CHR(KeyAscii) <= 9) Then
KeyAscii = 0
Else
        txtNum3.SetFocus
End If
End Sub

Private Sub txtNum3_Change()
FilltxtBord
End Sub

Private Sub txtNum3_KeyPress(KeyAscii As Integer)
txtNum3.text = ""
If Len(txtNum3.text) > 0 Then
KeyAscii = 0
End If
If Not (CHR(KeyAscii) >= 0 And CHR(KeyAscii) <= 9) Then
KeyAscii = 0
Else
        txtNum4.SetFocus

End If
End Sub

Private Sub txtNum4_Change()
FilltxtBord
End Sub

Private Sub txtNum4_KeyPress(KeyAscii As Integer)
txtNum4.text = ""
If Len(txtNum4.text) > 0 Then
KeyAscii = 0
End If
If Not (CHR(KeyAscii) >= 0 And CHR(KeyAscii) <= 9) Then
KeyAscii = 0
End If
End Sub


Private Sub DBCboClientName_Change()
    If Me.TxtModFlg2(mIndex).text = "N" Or Me.TxtModFlg2(mIndex).text = "E" Then
            Dim DefaultSalesPersonId As Integer
         '    Me.DcboEmp.BoundText = ""
            Dim mFull As String
            GetCustomersDetail val(DBCboClientName.BoundText), DefaultSalesPersonId, mFull
            
        '    TxtSearchCode2.Text = mFull
            If Not DefaultSalesPersonId = 0 Then

 '               Me.DcboEmp.BoundText = DefaultSalesPersonId
            End If
            GetCustomerNamebyPhone , , DBCboClientName.BoundText
            
        End If
End Sub
Private Sub DBCboClientName_Click(Area As Integer)
 If val(DBCboClientName.BoundText) = 0 Then Exit Sub

    Dim EmpCode  As String
 
    GetTblCustemersCode , , DBCboClientName.BoundText, EmpCode
    Me.TxtSearchCode.text = EmpCode
 '   Me.TxtSearchCode2.Text = EmpCode
End Sub

Private Sub DBCboClientName_KeyUp(KeyCode As Integer, _
                                  Shift As Integer)

    If KeyCode = vbKeyF3 Then
        FrmCustemerSearch.SearchType = 2020
        FrmCustemerSearch.show vbModal
    End If
    
    If KeyCode = vbKeyF5 Then
        Dim Dcombos As ClsDataCombos
        
        Set Dcombos = New ClsDataCombos
        Dcombos.GetCustomersSuppliers 0, Me.DBCboClientName, True
 
    End If
        
End Sub

Private Sub optCash_Click()
If optCash Then
    txtCustName.Enabled = True
    DBCboClientName.Enabled = False
    DBCboClientName.BoundText = 2
    txtAmountCash.Visible = True
    XPLbl(9).Visible = True
    txtAmountVisa.Visible = True
    XPLbl(10).Visible = True
    XPLbl(11).Visible = False
    txtAmountLater.Visible = False
    txtAmountLater = ""
    cmbPaymentClass.BoundText = 1
'    If Me.TxtModFlg2(mIndex).Text = "N" Or Me.TxtModFlg2(mIndex).Text = "E" Then
'        CalCulteVAT 3
'    End If
Else
    DBCboClientName.Enabled = True
    txtCustName.Enabled = False
    txtCustName = ""
     XPLbl(11).Visible = True
    txtAmountLater.Visible = True
    txtAmountCash = ""
    txtAmountVisa = ""
    txtAmountCash.Visible = False
    XPLbl(9).Visible = False
    txtAmountVisa.Visible = False
    XPLbl(10).Visible = False
    txtAmountLater = txtTotalWithVat2
    
End If
End Sub

Private Sub optLater_Click()
optCash_Click
End Sub

Private Sub txtAmountVisa_Change()
If Me.TxtModFlg2(mIndex).text = "N" Or Me.TxtModFlg2(mIndex).text = "E" Then
  '  txtAmountCash = val(XPTxtVal) - val(txtAmountVisa)
End If
End Sub

Private Sub txtRemarks2_Change()

End Sub

Private Sub txtPhoneCust_GotFocus()
txtPhoneCust.text = ""
End Sub

Private Sub TxtSalary_KeyPress(KeyAscii As Integer)
 KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtSalary.text, 1)
End Sub

Private Sub TxtSearchCode_KeyPress(KeyAscii As Integer)
 Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        'GetTblCustemersCode TxtSearchCode.Text, EmpID
        'DBCboClientName.BoundText = EmpID
        GetCustomerNamebyPhone , , , TxtSearchCode.text
    End If
End Sub


Private Sub Btn_Print_Click(Index As Integer)
        If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

            print_report "", mIndex
End Sub

Private Sub btn_Query_Click(Index As Integer)
Dim mfrm As New FemSearchDevelopment
If Index = 9 Then
   
        mfrm.mIndex = 1
        Load mfrm
    mfrm.Caption = "»ÕÀ ⁄‰ œŒÊ· «·„⁄œ« /«·”Ì«—« "

    mfrm.show vbModal
ElseIf Index = 3 Then
    
        mfrm.mIndex = 3
        Load mfrm
    mfrm.Caption = "»ÕÀ ⁄‰ √Ê«„— «·‘€·"

    mfrm.show vbModal
End If
End Sub

Private Sub CMDPAy_Click()
            Dim subject As String
            Dim Msg As String
            Dim msgstatus As Boolean
           Dim CompanyName As String
           Dim cOptions As ClsCompanyInfo
           Set cOptions = New ClsCompanyInfo
           Dim companyphone As String
           Dim optIsResponsible As Integer
            Dim CurrentMessage As String
            Dim t As String
    CurrentMessage = ComposMessage(Me.Name, 0, "", "", optIsResponsible)
    
    If Trim(txtPhoneCust) = "" Then Exit Sub
     
      
    'If Not SystemOptions.UserInterface = EnglishInterface Then
        Msg = "‰‘ﬂ—ﬂ„ ·“Ì«— ‰«" & CHR(13)
        Msg = Msg & "  «—ÌŒ " & XPDtbTrans(mIndex) & CHR(13)
        Msg = Msg & " Êﬁ  " & startTime.value & CHR(13)
        Msg = Msg & " «·ﬁÌ„… " & XPTxtVal & "—Ì«· " & CHR(13)
        Msg = Msg & " ﬁ.„ " & TxtVAt22 & "—Ì«· " & CHR(13)
        Msg = Msg & " —ﬁ„ «·›« Ê—… " & TxtNoteSerial1 & CHR(13)
        Msg = Msg & " —ﬁ„ «·„⁄œÂ/«·”Ì«—… " & txtnBoardNo.text & CHR(13)
    'Else
   
    
    'End If
    Dim isFound As Boolean
    isFound = True
    txtCodeSend = "+966"
    
    If Not FindString(txtPhoneCust, "+966", 1) Then
        If Not FindString(txtPhoneCust, "00966", 1) Then
            isFound = False
        End If
        isFound = False
    End If
    If Not isFound Then
        txtCodeSend = "+966"
    Else
        txtCodeSend = ""
        'txtPhoneCust = "+966" & val(txtPhoneCust)
    End If
    Dim mTxt As String
    mTxt = txtCodeSend & val(txtPhoneCust)
    lbl(56).Caption = mTxt
    t = sendMessageM("user", "password", Msg, "", mTxt)
    
         Msg = "We thank you for visiting us" & CHR(13)
        Msg = Msg & " Date " & XPDtbTrans(mIndex) & CHR(13)
        Msg = Msg & " Time " & startTime.value & CHR(13)
        Msg = Msg & " Value " & XPTxtVal & "SAR " & CHR(13)
        Msg = Msg & " Vat " & TxtVAt22 & "SAR " & CHR(13)
        Msg = Msg & " Invoice No " & TxtNoteSerial1 & CHR(13)
        Msg = Msg & " CarNo " & txtBoardNo.text & CHR(13)
        t = sendMessageM("user", "password", Msg, "", mTxt)
    DoEvents

End Sub
Private Function FindString(Control As Control, FindStr As String, Optional StartPos As Integer = 1) As Boolean
Dim a As Integer
a = InStr(StartPos, LCase$(Control.text), LCase$(FindStr))
If a = 0 Then
FindString = False
Else
FindString = True
Control.SetFocus
Control.SelStart = a - 1
Control.SelLength = Len(FindStr)
End If
End Function

Private Sub CMDSHOWISSUE_Click()
 'FrmOut.Retrive val(TXTTransactionID1.Text)
  FrmOut.Retrive val(TXTTransactionID1.text)
 
End Sub


Private Sub cmdReturnSales_Click()
Load FrmReturnSalling
FrmReturnSalling.Cmd_Click (0)
FrmReturnSalling.CboRetrunType.ListIndex = 1
FrmReturnSalling.TxtInvSerial = TxtNoteSerial13
FrmReturnSalling.CmdOpenTrans_Click
FrmReturnSalling.show
End Sub

Private Sub Command8_Click()
Dim StrTempAccountCode As String
                   StrTempAccountCode = GetMyAccountCode("TblCustemers", "CusID", val(Me.DcCustmer.BoundText))
 
            ShowReport StrTempAccountCode, DcCustmer.text, FrmDate.value, ToDate.value

End Sub


Private Sub DcCustmer_Validate(Cancel As Boolean)
 If Me.TxtModFlg2(mIndex).text = "N" Or Me.TxtModFlg2(mIndex).text = "E" Then
            Dim DefaultSalesPersonId As Integer
         '    Me.DcboEmp.BoundText = ""
            Dim mFull As String
            GetCustomersDetail val(DcCustmer.BoundText), DefaultSalesPersonId, mFull
            TxtSearchCode.text = mFull
            
            If Not DefaultSalesPersonId = 0 Then

 '               Me.DcboEmp.BoundText = DefaultSalesPersonId
            End If
            GetCustomerNamebyPhone , , DcCustmer.BoundText
            
        End If
End Sub

Private Sub TxtPhone_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    GetCustomerNamebyPhone (TxtPhone.text)
End If
If Trim(TxtPhone) = "" Then

    Dim Dcombos As New ClsDataCombos
    Dcombos.GetCustomersSuppliers 1, Me.DcCustmer, True
Else
    Dim sql  As String
    sql = "SELECT     Cus_mobile , CusName,CusID,Fullcode,cPaymentType,EmpId From dbo.TblCustemers  WHERE     (Cus_mobile = '" & TxtPhone & "')"
    fill_combo DcCustmer, sql
End If
End Sub

Private Sub cmdAddCustomer_Click()
    Dim Dcombos As New ClsDataCombos
If SystemOptions.DontShowMoreDetailsCompItem Then
    
    FrmCustemers.show
    FrmCustemers.Retrive val(DcCustmer.BoundText), Me.Name
    FrmCustemers.FormNamee = Me.Name
    
   ' Dcombos.GetCustomersSuppliers 1, Me.DcCustmer, True
    If DcCustmer.text = "" Then
   '     DcCustmer.BoundText = mCustId
    End If
    Exit Sub
End If
           
Dim CUSTID As Double
Dim mCode As String

If SystemOptions.UserInterface = ArabicInterface Then
    If Trim(txtCustomerName) = "" Then MsgBox "«œŒ· «”„ «·⁄„Ì·": Exit Sub
    If Trim(TxtPhone) = "" Then MsgBox "«œŒ· —ﬁ„ «·Â« ›/«·ÃÊ«·  ": Exit Sub
Else
    If Trim(txtCustomerName) = "" Then MsgBox "Enter the customer name": Exit Sub
    If Trim(TxtPhone) = "" Then MsgBox "Enter your phone / mobile number  ": Exit Sub

End If

Dim s As String
Dim rsDummy As New ADODB.Recordset

s = "Select * from TblCustemers WHere 1=1 "
If Trim(TxtPhone) <> "" Then
    s = s & " And Cus_mobile = N'" & Trim(TxtPhone) & "' "
End If
If Trim(txtCustomerName) <> "" Then
    'If Trim(TxtPhone) <> "" Then
    '    s = s & " Or CusName = '" & Trim(txtCustomerName.Text) & "'"
    'Else
    '    s = s & " and CusName = '" & Trim(txtCustomerName.Text) & "'"
    'End If
End If
rsDummy.Open s, Cn, adOpenStatic
If Not rsDummy.EOF Then
    TxtSearchCode.text = rsDummy!fullcode & ""
    
    DcCustmer.BoundText = val(rsDummy!CusID & "")
   
    txtCustomerName.backcolor = vbGreen
    TxtPhone.backcolor = vbGreen
    Exit Sub
Else
    txtCustomerName.backcolor = vbWhite
    TxtPhone.backcolor = vbWhite
End If

    createCustomer txtCustomerName.text, txtCustomerName.text, val(dcBranch(mIndex).BoundText), CUSTID, TxtPhone.text, mCode
    TxtSearchCode.text = mCode
    
    Dcombos.GetCustomersSuppliers 1, Me.DcCustmer, True
    DcCustmer.BoundText = CUSTID
    If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox " „ «÷«›… «·⁄„Ì·"
    Else
        MsgBox "Customer added"
    End If
    'txtCustomerName = ""

End Sub

Private Sub lbl_MouseMove(Index As Integer, _
                          Button As Integer, _
                          Shift As Integer, _
                          X As Single, _
                          Y As Single)

    If val(lbl(Index).Caption) <> 0 Then
        lbl(Index).ToolTipText = WriteNo(lbl(Index).Caption, 0, True)
    End If
    'ff

End Sub


Private Sub cmdInsertEmpItems_Click()

'FG7.Rows = 1
Dim ii As Long
Dim j As Long
    If fg7.rows <= 2 Then
        If Trim(fg7.TextMatrix(fg7.rows - 1, fg7.ColIndex("EmpID"))) = "" Then
            fg7.rows = fg7.rows - 1
        End If
    End If
    For ii = 0 To ListProductLineSelected.ListCount - 1
        For j = 0 To ListGroupSelected.ListCount - 1
            
            If chkEmpItem(val(ListProductLineSelected.ItemData(ii)), val(ListGroupSelected.ItemData(j))) Then
                fg7.rows = fg7.rows + 1
                fg7.TextMatrix(fg7.rows - 1, 0) = fg7.rows - 1
                fg7.TextMatrix(fg7.rows - 1, fg7.ColIndex("EmpID")) = ListProductLineSelected.ItemData(ii)
                fg7.TextMatrix(fg7.rows - 1, fg7.ColIndex("EmpName")) = ListProductLineSelected.List(ii)
                fg7.TextMatrix(fg7.rows - 1, fg7.ColIndex("ItemID")) = ListGroupSelected.ItemData(j)
                fg7.TextMatrix(fg7.rows - 1, fg7.ColIndex("ItemName")) = ListGroupSelected.List(j)
            End If
        Next j
    Next ii
End Sub
Private Function chkEmpItem(ByVal mEmpId As Long, ByVal mItemId As Long) As Boolean
    Dim i As Long
    Dim j As Long
    Dim mEmpID2 As Long
    Dim mItemId2 As Long
    For i = 1 To fg7.rows - 1
        mEmpID2 = val(fg7.TextMatrix(i, fg7.ColIndex("EmpID")))
        mItemId2 = val(fg7.TextMatrix(i, fg7.ColIndex("ItemID")))
        If mEmpId = mEmpID2 And mItemId2 = mItemId And mEmpID2 <> 0 Then chkEmpItem = False: Exit Function
        

    Next
     chkEmpItem = True
End Function


Private Function chkEmpItem2(ByVal mBranchID As Long, ByVal mAqarID As Long, ByVal mUnitTypeID As Long, ByVal mUnitNo As Long, ByVal mAqarUnitID As Long) As Boolean
    Dim i As Long
    Dim j As Long
    Dim mBranchId2 As Long
    Dim mAqarID2 As Long
    Dim mUnitTypeID2 As Long
    Dim mUnitNo2 As Long
    Dim mAqarUnitID2 As Long
    For i = 1 To GrdIqar.rows - 1
        mBranchId2 = val(GrdIqar.TextMatrix(i, GrdIqar.ColIndex("BranchId")))
        mAqarID2 = val(GrdIqar.TextMatrix(i, GrdIqar.ColIndex("Iqar")))
        mUnitTypeID2 = val(GrdIqar.TextMatrix(i, GrdIqar.ColIndex("unittype")))
        mUnitNo2 = val(GrdIqar.TextMatrix(i, GrdIqar.ColIndex("UnitNo")))
        If mBranchID = mBranchId2 And mAqarID2 = mAqarID And mUnitTypeID2 = mUnitTypeID And mAqarID2 = mAqarUnitID And mUnitNo2 = mUnitNo And mBranchID <> 0 Then chkEmpItem2 = False: Exit Function
        

    Next
     chkEmpItem2 = True
End Function
Private Sub Label28_Click()
    If ListProductLineAll.ListIndex = -1 Then Exit Sub
'    ListProductLineSelected.AddItem ListProductLineAll.List(ListProductLineAll.ListIndex)
'    ListProductLineSelected.ItemData(ListProductLineSelected.NewIndex) = ListProductLineAll.ItemData(ListProductLineAll.ListIndex)
'
    
Dim i As Long

For i = 0 To ListProductLineAll.ListCount - 1
    If ListProductLineAll.Selected(i) Then
        ListProductLineSelected.AddItem ListProductLineAll.List(i)
        ListProductLineSelected.ItemData(ListProductLineSelected.NewIndex) = ListProductLineAll.ItemData(i)
        
    End If
Next
    
'    FG.Rows = ListProductLineSelected.ListCount + 1
'    FG.TextMatrix(FG.Rows - 1, FG.ColIndex("Name")) = ListProductLineAll.List(ListProductLineAll.ListIndex)
'    FG.TextMatrix(FG.Rows - 1, FG.ColIndex("ProductLineID")) = ListProductLineAll.ItemData(ListProductLineAll.ListIndex)
End Sub

Private Sub Label29_Click()
    Dim i As Integer
    ListProductLineSelected.Clear
'    FG.Rows = 1
'    FG.Rows = ListProductLineSelected.ListCount + 1
    For i = 0 To ListProductLineAll.ListCount - 1
        ListProductLineSelected.AddItem ListProductLineAll.List(i)
        ListProductLineSelected.ItemData(i) = ListProductLineAll.ItemData(i)
'        FG.TextMatrix(i + 1, FG.ColIndex("Name")) = ListProductLineAll.List(ListProductLineAll.ListIndex)
'        FG.TextMatrix(i + 1, FG.ColIndex("ProductLineID")) = ListProductLineAll.ItemData(ListProductLineAll.ListIndex)
        
    Next i

End Sub

Private Sub Label30_Click()
 ListProductLineSelected.Clear
' FG.Rows = 1
End Sub

Private Sub Label31_Click()
'    If ListProductLineSelected.ListIndex > -1 Then
'      ListProductLineSelected.RemoveItem ListProductLineSelected.ListIndex
'        'FG.RemoveItem
'    End If


Dim i As Long

For i = 0 To ListProductLineSelected.ListCount - 1
    If i > ListProductLineSelected.ListCount - 1 Then Exit For
    If ListProductLineSelected.Selected(i) Then
        ListProductLineSelected.RemoveItem i
        'ListProductLineSelected.ListIndex
        i = i - 1
    End If
Next
    
End Sub




Private Sub ISButton2_Click()
Dim s As String



s = " SELECT TblStudCalling2.*,TblStudCalling.ID as ReservNo,TblStudCalling.EnterDate,HoursT as TimeR,HoursT as Hours,TblCustemers.CusName ,TblCustemers.CusID,"
s = s & "       TblEmployee.Emp_Name        EmpName,"
s = s & "       TblItems.ItemName,"
s = s & "       tblReservationType.Name  AS ReservationTypeName"
s = s & " From TblStudCalling2"
s = s & "       Left Outer JOIN tblReservationType"
s = s & "            ON  tblReservationType.ID = TblStudCalling2.ReservationTypeCode"
s = s & "       INNER JOIN TblEmployee"
s = s & "            ON  TblEmployee.Emp_ID = TblStudCalling2.EmpID"
s = s & "       INNER JOIN TblStudCalling "
s = s & "            ON  TblStudCalling.ID = TblStudCalling2.MasterID"
s = s & "       INNER JOIN TblCustemers"
s = s & "            ON  TblCustemers.CusID = TblStudCalling.CompID"


s = s & "       INNER JOIN TblItems"
s = s & "            ON  TblItems.ItemID = TblStudCalling2.ItemID"


s = s & " Where TblStudCalling.EnterDate>= " & SQLDate(XPDtbBill(0), True) & "  and TblStudCalling.EnterDate <= " & SQLDate(XPDtbBill(1), True)
If CmbEmp.text <> "" And val(CmbEmp.BoundText) <> 0 Then
    s = s & " and TblEmployee.Emp_ID  = " & val(CmbEmp.BoundText)
End If
If cmbCustomer.text <> "" And val(cmbCustomer.BoundText) <> 0 Then
    s = s & " and TblCustemers.CusID= " & val(cmbCustomer.BoundText)
End If

If cmbItems.text <> "" And val(cmbItems.BoundText) <> 0 Then
    s = s & " and TblItems.ItemID= " & val(cmbItems.BoundText)
End If


loadgrid s, Fg6, True, True


Dim i As Long
For i = 1 To Fg6.rows - 1
    If Trim(Fg6.TextMatrix(i, Fg6.ColIndex("Hours"))) <> "" Then
     Dim mStartDate As String
     
     Dim mStartDate2 As Date
     Dim mDateNow As String
     
     Dim mDateNow2 As Date
     'mStartDate = Format$(FG6.TextMatrix(i, FG6.ColIndex("EnterDate")), FG6.TextMatrix(i, FG6.ColIndex("Hours")))
     On Error Resume Next
     mStartDate = Format$(Fg6.TextMatrix(i, Fg6.ColIndex("EnterDate")) + " " + Fg6.TextMatrix(i, Fg6.ColIndex("Hours")))
     mStartDate2 = CDate(mStartDate)
      
     'mDateNow = Now
     mStartDate2 = CDate(mStartDate)
     
        '    FG6.TextMatrix(i, FG6.ColIndex("ItemName")) = GetTimeDiff(Now, mStartDate2, 1, 1)
               Fg6.TextMatrix(i, Fg6.ColIndex("StillPeriod")) = GetTimeDiff(Now, mStartDate2, 1, 1)
             If val(Fg6.TextMatrix(i, Fg6.ColIndex("StillPeriod"))) > 10 Then
                Fg6.TextMatrix(i, Fg6.ColIndex("StillPeriod")) = Fg6.TextMatrix(i, Fg6.ColIndex("StillPeriod")) & "  ”«⁄… "
             Else
                Fg6.TextMatrix(i, Fg6.ColIndex("StillPeriod")) = Fg6.TextMatrix(i, Fg6.ColIndex("StillPeriod")) & "  ”«⁄«  "
             End If

                Fg6.TextMatrix(i, Fg6.ColIndex("minutes")) = GetTimeDiff(Now, mStartDate2, 1, 2)
                If val(Fg6.TextMatrix(i, Fg6.ColIndex("minutes"))) > 10 Then
                    Fg6.TextMatrix(i, Fg6.ColIndex("minutes")) = Fg6.TextMatrix(i, Fg6.ColIndex("minutes")) & " œﬁÌﬁ…"
                Else
                    Fg6.TextMatrix(i, Fg6.ColIndex("minutes")) = Fg6.TextMatrix(i, Fg6.ColIndex("minutes")) & " œﬁ«∆ﬁ"
                End If
    End If

Next


End Sub


Private Sub Label53_Click()
'   If ListGroupSelected.ListIndex > -1 Then
'        ListGroupSelected.RemoveItem ListGroupSelected.ListIndex
'    End If
    
    

               
Dim i As Long

For i = 0 To ListGroupSelected.ListCount - 1
    If i > ListGroupSelected.ListCount - 1 Then Exit For
    If ListGroupSelected.Selected(i) Then
        ListGroupSelected.RemoveItem i
        'ListProductLineSelected.ListIndex
        i = i - 1
    End If
Next
End Sub

Private Sub Label63_Click()
ListGroupSelected.Clear
End Sub

Private Sub Label7_Click()
    Dim i As Integer
    ListGroupSelected.Clear

    For i = 0 To ListGroupAll.ListCount - 1
        ListGroupSelected.AddItem ListGroupAll.List(i)
        ListGroupSelected.ItemData(i) = ListGroupAll.ItemData(i)
    Next i
End Sub
Private Sub Label8_Click()
    If ListGroupAll.ListIndex = -1 Then Exit Sub
'    ListGroupSelected.AddItem ListGroupAll.List(ListGroupAll.ListIndex)
'    ListGroupSelected.ItemData(ListGroupSelected.NewIndex) = ListGroupAll.ItemData(ListGroupAll.ListIndex)
'
    Dim i As Long
    
    For i = 0 To ListGroupAll.ListCount - 1
        If ListGroupAll.Selected(i) Then
            ListGroupSelected.AddItem ListGroupAll.List(i)
            ListGroupSelected.ItemData(ListGroupSelected.NewIndex) = ListGroupAll.ItemData(i)
            
        End If
    Next
            
End Sub

Private Sub Cmd_DeleteAll_Click(Index As Integer)
If Me.TxtModFlg2(mIndex).text <> "R" Then

    RemoveGridRowAll Index

End If
End Sub

Private Sub Cmd_DeleteRow_Click(Index As Integer)
If Me.TxtModFlg2(mIndex).text <> "R" Then

    RemoveGridRow Index

End If
CalcAmount
End Sub

Private Sub RemoveGridRowAll(ByVal mInx As Long)
    
    If mIndex = 3 Then
        FG.rows = 1
CalcAmount
    ElseIf mIndex = 4 Then
        FG4.rows = 1
    ElseIf mIndex = 6 Then
        Fg6.rows = 1
    ElseIf mIndex = 7 Then
        fg7.rows = 1
    ElseIf mIndex = 8 Then
    
        'FG7.Rows = 1
    ElseIf mIndex = 12 Then
        GrdIqar.rows = 1
        
    End If
    
End Sub


Private Sub RemoveGridRow(ByVal mInx As Long)
    
    If mIndex = 3 Then
        With Me.FG
    'MsgBox .Row
            If .Row <= 0 Then
                    .rows = 2
            Exit Sub
            Else
            .RemoveItem .Row
            End If
        End With
    ElseIf mIndex = 4 Then
        With Me.FG4
    'MsgBox .Row
            If .Row <= 0 Then
                .rows = 2
                Exit Sub
            Else
                .RemoveItem .Row
            End If
        End With
    ElseIf mIndex = 6 Then
        With Me.Fg6
    'MsgBox .Row
            If .Row <= 0 Then
                .rows = 2
                Exit Sub
            Else
                .RemoveItem .Row
            End If
        End With
    ElseIf mIndex = 7 Then
        With Me.fg7
    'MsgBox .Row
            If .Row <= 0 Then
                .rows = 2
                Exit Sub
            Else
                .RemoveItem .Row
            End If
        End With
    ElseIf mIndex = 12 Then
        With Me.GrdIqar
    'MsgBox .Row
            If .Row <= 0 Then
                .rows = 2
                Exit Sub
            Else
                .RemoveItem .Row
            End If
        End With
                
        
    End If
    
End Sub

Private Sub btn_Cancel_Click(Index As Integer)

   Unload Me
End Sub

Private Sub btn_Delete_Click(Index As Integer)
   Dim MSGType As Integer
   Dim StrSQL As String
    Dim BlnRecExist As Boolean
    Dim StrMSG  As String
    Dim s As String
    Dim rsDummy As New ADODB.Recordset
    Dim Msg As String
    On Error GoTo ErrTrap
    'Index = TabMain.CurrTab
    
    If mIndex = 10 Then
        s = "Select * from TblEmpDataInOut Where EmpId = " & val(TxtSerial1(mIndex))
        Set rsDummy = New ADODB.Recordset
        rsDummy.Open s, Cn, adOpenStatic
        If Not rsDummy.EOF Then
            MsgBox "·« Ì„ﬂ‰ Õ–› Â–« «·„ÊŸ› ·«‰ ·œÌÂ ”Ã· Õ÷Ê— Ê«‰’—«›"
            Exit Sub
        End If
    End If
    Dim i As Long
    If DoPremis(Do_Delete, Me.Name, True) = False Then
        Exit Sub
    End If
    If TxtSerial1(mIndex).text <> "" Then
        '    If CheckDelCountry(Val(Me.TxtVac_ID.text)) = False Then
        '        Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã·...!!!"
        '        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        '        Exit Sub
        '    End If
        If SystemOptions.UserInterface = ArabicInterface Then
        MSGType = MsgBox("Â·  —€» ›Ì Õ–› Â–« «·”Ã·", vbQuestion + vbMsgBoxRtlReading + vbYesNo + vbMsgBoxRight, App.Title)
        Else
        MSGType = MsgBox("Confirm Delete", vbQuestion + vbMsgBoxRtlReading + vbYesNo + vbMsgBoxRight, App.Title)
        End If

        If MSGType = vbYes Then
            
         '   CuurentLogdata ("D")
            RsSavRec.delete
            If mIndex = 3 Then
            '    End If
                StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(txtNoteSerialCash(1).text)
                Cn.Execute StrSQL, , adExecuteNoRecords
                StrSQL = "Delete From TblMultuPayment Where NoteID=" & val(txtNoteSerialCash(1).text)
                Cn.Execute StrSQL, , adExecuteNoRecords
                 StrSQL = "Delete From Notes Where NoteID=" & val(txtNoteSerialCash(1).text)
                Cn.Execute StrSQL, , adExecuteNoRecords
    '            StrSQL = "Delete  marakes_taklefa_temp  where kedno =" & val(Text1.Text)
    '            Cn.Execute StrSQL, , adExecuteNoRecords
    
    
                StrSQL = " delete   notes where NoteType= 2000   and  NoteSerial='" & txtNoteSerialCash(0).text & "'"
                DelSales
                
            ElseIf mIndex = 4 Then
                  StrSQL = "Delete From TblJobOrdersTasks2 Where MasterID=" & val(TxtSerial1(mIndex).text)
                Cn.Execute StrSQL, , adExecuteNoRecords
            

            ElseIf mIndex = 12 Then
                StrSQL = " Delete From TblIqarDiscountTrans2 Where MasterID = " & val(TxtSerial1(mIndex).text)
    
        
        

                Cn.Execute StrSQL, , adExecuteNoRecords
            ElseIf mIndex = 6 Then
                  StrSQL = "Delete From TblAppointmentlist2 Where MasterID=" & val(TxtSerial1(mIndex).text)
                Cn.Execute StrSQL, , adExecuteNoRecords
           

            ElseIf mIndex = 7 Then
                  StrSQL = "Delete From TblEmpItemsTrans2 Where MasterID=" & val(TxtSerial1(mIndex).text)
                Cn.Execute StrSQL, , adExecuteNoRecords
            End If

            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox " „  ⁄„·Ì… «·Õ–› »‰Ã«Õ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
            Else
                MsgBox "Delete Successfully", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
            End If
            '------------------------------ Move Next ---------------------------.
            RsSavRec.Requery
            
            If mIndex = 0 Then
                FiLLTXT1
                FillGridWithData1
            ElseIf mIndex = 1 Then
                FiLLTXT2
                FillGridWithData2
            ElseIf mIndex = 3 Then
                FiLLTXT3
            ElseIf mIndex = 4 Then
                FiLLTXT4
            ElseIf mIndex = 5 Then
                FiLLTXT5
                FillGridWithData5
            ElseIf mIndex = 6 Then
                FiLLTXT6
            ElseIf mIndex = 7 Then
                FiLLTXT7
            ElseIf mIndex = 8 Then
                FiLLTXT8
                FillGridWithData8
            ElseIf mIndex = 9 Then
                FiLLTXT9
            ElseIf mIndex = 10 Then
                FiLLTXT10
            ElseIf mIndex = 11 Then
                FiLLTXT11
            ElseIf mIndex = 12 Then
                FiLLTXT12
                                
                
            Else
                FillGridWithData2
            End If
            btn_Next_Click mIndex
        End If
    End If
        Dim mCol As Long
   
      For i = 0 To grd.Cols - 1
        
            grd.Col = i
            grd.CellBackColor = vbWhite
       
      
    Next
   
    For i = 0 To grd.Cols - 1
        If val(cmbPaymentClass.BoundText) = val(grd.ColKey(i)) Then
            grd.Col = i
            grd.CellBackColor = vbBlue
            Exit For
        End If
    Next
    Exit Sub
ErrTrap:
 
    Select Case Err.Number

        Case -2147217873, -2147467259
        If SystemOptions.UserInterface = ArabicInterface Then
            StrMSG = "⁄›Ê« ·« ÌÃÊ“ Õ–› «·”Ã· ·«— »«ÿÂ »»Ì«‰«  √Œ—Ì"
            Else
            StrMSG = "Sorry you can not delete the record of its connection with other data"
            End If
            RsSavRec.CancelUpdate
            MsgBox StrMSG, vbOKOnly + vbMsgBoxRight, App.Title
            'clear the ConnectiOn Errors
            Cn.Errors.Clear
    End Select

End Sub
Private Sub DelSales()
Dim StrSQL As String
  Cn.Execute "delete from Transaction_Details where Transaction_ID =  " & val(TXTTransactionID3.text)
  Cn.Execute "delete from Transactions where Transaction_ID =  " & val(TXTTransactionID3.text)
  Cn.Execute "delete from DOUBLE_ENTREY_VOUCHERS where Transaction_ID =  " & val(TXTTransactionID3.text)
  
  Cn.Execute "delete From TblSalesPayment where TransID=" & val(Me.TXTTransactionID3.text)   'Val(rs("Transaction_ID").value)
  Cn.Execute "delete From TblSalesMixItems where TransectionID=" & val(Me.TXTTransactionID3.text) 'Val(rs("Transaction_ID").value)
   StrSQL = "Delete From TblPayPrePayed Where TypeTrans=1 and  NoteID1=" & val(Me.TXTTransactionID3.text)
   Cn.Execute StrSQL, , adExecuteNoRecords
   StrSQL = "Delete From TblProjePayPrePayed Where TypeTrans=1 and  NoteID=" & val(Me.TXTTransactionID3.text)
   Cn.Execute StrSQL, , adExecuteNoRecords
         ' DeleteBillBuy
          
        
            Cn.Execute "Delete from TransactionValueAdded where Transaction_ID=" & val(Me.TXTTransactionID3.text) & ""
            StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS  " & "Where DOUBLE_ENTREY_VOUCHERS.Transaction_ID=" & val(TXTTransactionID3)
            Cn.Execute StrSQL, , adExecuteNoRecords
            StrSQL = "Delete From TblGridguranteeSales  " & "Where Transec_ID=" & val(TXTTransactionID3)
            Cn.Execute StrSQL, , adExecuteNoRecords
        
                  StrSQL = "Delete From TblTransactionPayments Where Transaction_ID=" & val(Me.TXTTransactionID3.text)
        Cn.Execute StrSQL, , adExecuteNoRecords
        
        
        Cn.Execute " delete TBLRegularMaint where Transaction_ID=" & val(TXTTransactionID3.text)
        
            '                StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS  " & _
            '         "Where DOUBLE_ENTREY_VOUCHERS.Transaction_ID=" & get_transaction_id(rs("nots").value, 19)
            '         Cn.Execute StrSQL, , adExecuteNoRecords
                
            '         StrSQL = "Delete From Transactions  " & _
            '         "Where Transaction_ID=" & get_transaction_id(rs("nots").value, 19)
            '         Cn.Execute StrSQL, , adExecuteNoRecords
            
            StrSQL = "delete From Notes where noteid=" & val(txtNoteid3.text)
    
            Cn.Execute StrSQL, , adExecuteNoRecords
 
            StrSQL = "delete From Notes where noteid=" & val(txtNoteSerialCash(1).text)
            Cn.Execute StrSQL, , adExecuteNoRecords
  

End Sub
Private Sub btn_First_Click(Index As Integer)
  On Error GoTo ErrTrap

    Dim Msg As String

  
    If Me.TxtModFlg2(mIndex).text = "N" Then
        FindRec val(TxtSerial1(mIndex).text)
        TxtModFlg2(mIndex).text = "R"
    End If

    TxtModFlg2(mIndex) = "R"

    If RsSavRec.RecordCount = 0 Then
        clear_all Me
        Exit Sub
    End If

BegnieWork:
    RsSavRec.MoveFirst
    If mIndex = 0 Then
        FiLLTXT1
    ElseIf mIndex = 1 Then
        FiLLTXT2
    ElseIf mIndex = 3 Then
        FiLLTXT3
    ElseIf mIndex = 4 Then
        FiLLTXT4
   ElseIf mIndex = 5 Then
        FiLLTXT5
   ElseIf mIndex = 6 Then
        FiLLTXT6
   ElseIf mIndex = 7 Then
        FiLLTXT7
   ElseIf mIndex = 8 Then
        FiLLTXT8
   ElseIf mIndex = 9 Then
        FiLLTXT9
   ElseIf mIndex = 10 Then
        FiLLTXT10
   ElseIf mIndex = 11 Then
        FiLLTXT11
   ElseIf mIndex = 12 Then
        FiLLTXT12
        
        
        
    End If

    Exit Sub

ErrTrap:

    Select Case Err.Number

        Case -2147217885
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· " & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
          Else
            Msg = "Sorry I have been deleted the record" & CHR(13)
            Msg = Msg & "By another user on the network " & CHR(13)
            Msg = Msg & "Data will be updated"
        End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.Title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub

Function loadLogo()
    Dim rs As ADODB.Recordset
    Dim BolShowLogo As Boolean
    Dim xLogo As CRAXDRT.OLEObject
    Dim StrFileName As String
    Dim MsgErr As String

     

    Set rs = New ADODB.Recordset
    rs.Open "TblOptions", Cn, adOpenStatic, adLockReadOnly, adCmdTable

    If rs.BOF Or rs.EOF Then
       
        Exit Function
    End If

   

   If Not (IsNull(rs("CompanyLogo").value)) Then
        'LoadPictureFromDB ImgPic, rs, "CompanyLogo"
        'LoadPictureFromDB Image1, rs, "CompanyLogo"
        
    End If
    
End Function

Private Sub btn_Modify_Click(Index As Integer)
    Dim Msg As String

    If DoPremis(Do_Edit, Me.Name, True) = False Then
        Exit Sub
    End If

    On Error GoTo ErrTrap
    If mIndex = 10 Then
      '  ISButton5_Click
    End If
    
    If TxtSerial1(mIndex).text <> "" Then
        TxtModFlg2(mIndex) = "E"
    Frame1(1).Enabled = True
    
        'Frm2.Enabled = True
      '  TreeGroups.Enabled = True
    End If

    Exit Sub
ErrTrap:

    Select Case Err.Number

        Case -2147467259
            'Could not update; currently locked.
           If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê«" & CHR(13)
            Msg = Msg & " ·«Ì„ﬂ‰  ⁄œÌ· Â–« «·”Ã· ›Ï «·Êﬁ  «·Õ«·Ï" & CHR(13)
            Msg = Msg & "ÕÌÀ «‰Â ﬁÌœ «· ⁄œÌ· „‰ ﬁ»· „” Œœ„ «Œ— ⁄·Ï «·‘»ﬂ…"
           Else
            Msg = "Sorry" & CHR(13)
            Msg = Msg & " You can not edit this record now" & CHR(13)
            Msg = Msg & "Where it was being edited by another user on the network"
           End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.Title
        
            If RsSavRec.EditMode <> adEditNone Then
                RsSavRec.CancelUpdate
                'RsSavRec.Requery
            End If

    End Select

End Sub

Public Sub btn_New_Click(Index As Integer)
    Dim My_SQL As String
    Dim rs As ADODB.Recordset
 Dim currentgroup As Integer
    If DoPremis(Do_New, Me.Name, True) = False Then
        Exit Sub
    End If

    On Error GoTo ErrTrap
    Set rs = New ADODB.Recordset
   ' Frame1(mIndex).Enabled = True
    
    clear_all Me
    TxtModFlg2(mIndex).text = "N"

    If mIndex = 0 Then
        My_SQL = "select * from  TblTasks"
        'DCboUserName(mIndex) = user_id
             
         clear_all Me
          
            

   
    'rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable
rs.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If rs.RecordCount > 0 Then
        TxtSerial1(mIndex).text = rs.RecordCount + 1
    Else
        TxtSerial1(mIndex).text = 1
    End If

    rs.Close
    'CmbType.ListIndex = 0
    TxtName(mIndex).SetFocus
        
    
    ElseIf mIndex = 1 Then
       
      '  DCboUserName(mIndex) = user_id
        My_SQL = "select *  from TblSizesNames"
        'DCboUserName(mIndex) = user_id
     
         clear_all Me
    

   rs.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    'rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable

    If rs.RecordCount > 0 Then
        TxtSerial1(mIndex).text = rs.RecordCount + 1
    Else
        TxtSerial1(mIndex).text = 1
    End If

    rs.Close
        
    ElseIf mIndex = 3 Then

                My_SQL = "select * from TblJobOrders"
        'DCboUserName(mIndex) = user_id
     
         clear_all Me
        dcBranch(mIndex).BoundText = branch_id
         DCboUserName(mIndex).BoundText = user_id

   rs.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    'rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable

    If rs.RecordCount > 0 Then
        TxtSerial1(mIndex).text = rs.RecordCount + 1
    Else
        TxtSerial1(mIndex).text = 1
    End If

    rs.Close
    FG.rows = 50
   ElseIf mIndex = 4 Then

                My_SQL = "select * from TblJobOrdersTasks"
       
     
         clear_all Me
          DCboUserName(mIndex).BoundText = user_id
        dcBranch(mIndex).BoundText = branch_id

   rs.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable

    If rs.RecordCount > 0 Then
        TxtSerial1(mIndex).text = rs.RecordCount + 1
    Else
        TxtSerial1(mIndex).text = 1
    End If

    rs.Close
    
   ElseIf mIndex = 5 Then
        
        My_SQL = "select * from tblReservationType"
        
        
        clear_all Me
        DCboUserName(mIndex).BoundText = user_id
        dcBranch(mIndex).BoundText = branch_id
        
        
    '    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable
    rs.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
        If rs.RecordCount > 0 Then
            TxtSerial1(mIndex).text = rs.RecordCount + 1
        Else
            TxtSerial1(mIndex).text = 1
        End If
        
        rs.Close
   ElseIf mIndex = 6 Then

        My_SQL = " select * from TblAppointmentlist"
        
        
        clear_all Me
        DCboUserName(mIndex).BoundText = user_id
        dcBranch(mIndex).BoundText = branch_id
        
        
 '       rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable
 rs.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
 
        If rs.RecordCount > 0 Then
            TxtSerial1(mIndex).text = rs.RecordCount + 1
        Else
            TxtSerial1(mIndex).text = 1
        End If
        
        rs.Close
   ElseIf mIndex = 7 Then

        My_SQL = "select * from TblEmpItemsTrans"
        
  
             
    
   ElseIf mIndex = 8 Then
        'My_SQL = "tblPaymentClass"
         My_SQL = "select * from tblPaymentClass where  1=1"
         
   ElseIf mIndex = 9 Then
   
   
        My_SQL = "select * from TblTripReg where 1=1"
             DCboUserName(mIndex).BoundText = user_id
        dcBranch(mIndex).BoundText = branch_id

   ElseIf mIndex = 10 Then
        My_SQL = "select * from TblEmpData"
    '    ISButton5_Click
    lblProgressFinger.Caption = ""
    lblFingerStatus.Tag = ""
   ElseIf mIndex = 11 Then
        My_SQL = "select * from   TblEmpItemsTrans"
    ElseIf mIndex = 12 Then

        My_SQL = "select * from  TblIqarDiscountTrans"
    
    End If
        clear_all Me

        If mIndex = 12 Then
            ListBranchSelected.Clear
            ListAqarSelected.Clear
            ListUnitTypeSelected.Clear
            ListUnitNoSelected.Clear
            ListUnitNoSelected2.Clear
            ListBranchAll.Clear
            ListAqarAll.Clear
            ListUnitTypeAll.Clear
            ListUnitTypeSelected.Clear
            ListUnitNoAll2.Clear
            ListUnitNoAll.Clear
            ListUnitNoSelected.Clear
            FillMylist2 True, False, True, False
        End If
        On Error GoTo 11
        DCboUserName(mIndex).BoundText = user_id
        dcBranch(mIndex).BoundText = branch_id
       SetGridFinger
11:
  '     rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable
rs.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
        If rs.RecordCount > 0 Then
            TxtSerial1(mIndex).text = rs.RecordCount + 1
        Else
            TxtSerial1(mIndex).text = 1
        End If
        
        rs.Close
If mIndex = 9 Then
    optCash.value = True
End If
optIsEmp.value = True
    FG.rows = 50
         startTime = Time
         ReloadCompo
         GrdFinger.rows = 11
         txtPhoneCust.text = "123456789"
         lblClassCat.Caption = "0"
'    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable

'
''    If rs.RecordCount > 0 Then
''        TxtSerial1(mIndex).Text = rs.RecordCount + 1
''    Else
''        TxtSerial1(mIndex).Text = 1
''    End If
'
'    rs.Close
    'CmbType.ListIndex = 0
    'TxtVacName.SetFocus
ErrTrap:

End Sub


Private Sub btn_Next_Click(Index As Integer)
On Error GoTo ErrTrap
    Dim Msg As String

    If Me.TxtModFlg2(mIndex).text = "N" Then
        FindRec val(TxtSerial1(mIndex).text)
        TxtModFlg2(mIndex).text = "R"
    End If

    TxtModFlg2(mIndex) = "R"

    If RsSavRec.RecordCount = 0 Then
        clear_all Me
        Exit Sub
    End If

BegnieWork:

    If RsSavRec.EOF Then
        RsSavRec.MoveLast
    Else
        RsSavRec.MoveNext

        If RsSavRec.EOF Then
            RsSavRec.MoveLast
        End If
    End If

    If mIndex = 0 Then
        FiLLTXT1
    ElseIf mIndex = 1 Then
        FiLLTXT2
    ElseIf mIndex = 3 Then
        FiLLTXT3
    ElseIf mIndex = 4 Then
        FiLLTXT4
    ElseIf mIndex = 5 Then
        FiLLTXT5
    ElseIf mIndex = 6 Then
        FiLLTXT6
    ElseIf mIndex = 7 Then
        FiLLTXT7
    ElseIf mIndex = 8 Then
        FiLLTXT8
    ElseIf mIndex = 9 Then
        FiLLTXT9
    ElseIf mIndex = 10 Then
        FiLLTXT10
    ElseIf mIndex = 11 Then
        FiLLTXT11
    ElseIf mIndex = 12 Then
        FiLLTXT12

    End If
    Exit Sub
ErrTrap:

    Select Case Err.Number

        Case -2147217885
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· " & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
        Else
            Msg = "Sorry I have been deleted the  record" & CHR(13)
            Msg = Msg & "By another user on the network " & CHR(13)
            Msg = Msg & "Data will be updated"
        End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.Title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub

Private Sub btn_Previous_Click(Index As Integer)
  On Error GoTo ErrTrap
    Dim Msg As String

    If TxtModFlg2(mIndex).text = "N" Then
        FindRec val(TxtSerial1(mIndex).text)
        TxtModFlg2(mIndex).text = "R"
    End If

    Me.TxtModFlg2(mIndex) = "R"

    If RsSavRec.RecordCount = 0 Then
        clear_all Me
        Exit Sub
    End If

BegnieWork:

    RsSavRec.MovePrevious

    If RsSavRec.BOF Then
        RsSavRec.MoveFirst
    End If

    If mIndex = 0 Then
        FiLLTXT1
    ElseIf mIndex = 1 Then
        FiLLTXT2
    ElseIf mIndex = 4 Then
        FiLLTXT4
    ElseIf mIndex = 5 Then
        FiLLTXT5
    ElseIf mIndex = 6 Then
        FiLLTXT6
    ElseIf mIndex = 7 Then
        FiLLTXT7
    ElseIf mIndex = 8 Then
        FiLLTXT8
    ElseIf mIndex = 9 Then
        FiLLTXT9
    ElseIf mIndex = 10 Then
        FiLLTXT10
    ElseIf mIndex = 11 Then
        FiLLTXT11
    ElseIf mIndex = 12 Then
        FiLLTXT12
             
        

    End If
    Exit Sub
ErrTrap:

    Select Case Err.Number

        Case -2147217885
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· " & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
             Else
            Msg = "Sorry I have been deleted the  record" & CHR(13)
            Msg = Msg & "By another user on the network " & CHR(13)
            Msg = Msg & "Data will be updated"
        End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.Title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub

Private Sub btn_Save_Click(Index As Integer)
    On Error GoTo ErrTrap
    Dim Msg As String
    Dim StrVacCode As String
    Dim StrVacName As String
    Dim CtrlTxt As Control
    '---------------------- check if data Vaclete -----------------------

    For Each CtrlTxt In Me.Controls

        If TypeOf CtrlTxt Is TextBox Or TypeOf CtrlTxt Is ComboBox Then
            If CtrlTxt.text = "" And CtrlTxt.Tag <> "" And CtrlTxt.Enabled = True Then
                MsgBox CtrlTxt.Tag, vbOKOnly + vbMsgBoxRight, App.Title
                CtrlTxt.SetFocus
                Exit Sub
            End If
        End If

    Next


If mIndex = 2 Then
'     If Dcbranch(mIndex).Text = "" Then
'        MsgBox "Please Enter Branch"
'        Dcbranch(mIndex).SetFocus
'        Exit Sub
'    End If
End If
    
    '------------------------------ check if Empcode exist ----------------------

   

    ' -------------------------------------- txtmodflg type -------------------
    Select Case TxtModFlg2(mIndex).text

            '------------------------------ new record ----------------------------
        Case "N"
      
            '------------------------- save record -----------------------------
            If mIndex = 0 Then
                AddNewRec
                FiLLRec1
                
                FiLLTXT1
            ElseIf mIndex = 1 Then
                AddNewRec
                FiLLRec2
                FiLLTXT2
            ElseIf mIndex = 3 Then
                AddNewRec
                
                FiLLRec3
                
            ElseIf mIndex = 4 Then
                AddNewRec
                FiLLRec4
                FiLLTXT4
            ElseIf mIndex = 5 Then
                AddNewRec
                FiLLRec5
                FiLLTXT5
            ElseIf mIndex = 6 Then
                AddNewRec
                FiLLRec6
                FiLLTXT6
            ElseIf mIndex = 7 Then
                AddNewRec
                FiLLRec7
                FiLLTXT7
            ElseIf mIndex = 8 Then
                AddNewRec
                FiLLRec8
                FiLLTXT8
            ElseIf mIndex = 9 Then
                
                FiLLRec9
                
            ElseIf mIndex = 10 Then
                
                FiLLRec10
               
            ElseIf mIndex = 11 Then
                AddNewRec
                FiLLRec11
                FiLLTXT11
                
            ElseIf mIndex = 12 Then
                AddNewRec
                FiLLRec12
                FiLLTXT12
                
                
            End If
            

        Case "E"

            '----------------------------- save edit -------------------------------
            
            If mIndex = 0 Then
                FiLLRec1
            ElseIf mIndex = 1 Then
                FiLLRec2
            ElseIf mIndex = 3 Then
                DelSales
                FiLLRec3
            ElseIf mIndex = 4 Then
                FiLLRec4
            ElseIf mIndex = 5 Then
                FiLLRec5
            ElseIf mIndex = 6 Then
                FiLLRec6
            ElseIf mIndex = 7 Then
                FiLLRec7
            ElseIf mIndex = 8 Then
                FiLLRec8
            ElseIf mIndex = 9 Then
                FiLLRec9
            ElseIf mIndex = 10 Then
                FiLLRec10
            
                
                
            End If
    End Select

    Exit Sub
ErrTrap:
If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Â‰«ﬂ Œÿ√ „« ›Ì ≈œŒ«· «·»Ì«‰« ", vbOKOnly + vbMsgBoxRight, App.Title
 Else
  MsgBox "Sorry...error douring insert data", vbOKOnly + vbMsgBoxRight, App.Title
End If
 
End Sub

Private Sub Btn_Undo_Click(Index As Integer)
    Undo
End Sub
Private Sub Undo()
    On Error GoTo ErrTrap
    If mIndex = 2 Then
        Select Case TxtModFlg.text
        
        Case "N"
            clear_all Me
            Me.TxtModFlg.text = "R"
            XPBtnMove_Click (1)
        
        Case "E"
            rs.Find "Id='" & val(TxtSerial.text) & "'", , adSearchForward, adBookmarkFirst
        
        If rs.EOF Or rs.BOF Then
            Me.TxtModFlg.text = "R"
            Exit Sub
        End If
        
            'Retrive
            Me.TxtModFlg.text = "R"
        End Select
    
    Else
    
    Select Case TxtModFlg2(mIndex).text

        Case "N"
            clear_all Me
            TxtModFlg2(mIndex).text = "R"
           
            btn_First_Click (mIndex)
        Case "E"
            If mIndex = 0 Then
            
                RsSavRec.Find "ID='" & val(TxtSerial1(mIndex).text) & "'", , adSearchForward, adBookmarkFirst
            Else
                RsSavRec.Find "ID='" & val(TxtSerial1(mIndex).text) & "'", , adSearchForward, adBookmarkFirst
            End If

            If RsSavRec.EOF Or RsSavRec.BOF Then
                TxtModFlg2(mIndex).text = "R"
                Exit Sub
            End If

            If mIndex = 0 Then
                FiLLTXT1
            ElseIf mIndex = 1 Then
                FiLLTXT2
            ElseIf mIndex = 3 Then
                FiLLTXT3
            ElseIf mIndex = 4 Then
                FiLLTXT4
            ElseIf mIndex = 5 Then
                FiLLTXT5
            ElseIf mIndex = 6 Then
                FiLLTXT6
            ElseIf mIndex = 7 Then
                FiLLTXT7
           ElseIf mIndex = 8 Then
                FiLLTXT8
            ElseIf mIndex = 9 Then
                FiLLTXT9
            ElseIf mIndex = 10 Then
                FiLLTXT10
            ElseIf mIndex = 11 Then
                FiLLTXT11
            ElseIf mIndex = 12 Then
                FiLLTXT12
     
    
            End If
            TxtModFlg2(mIndex).text = "R"
    End Select
    End If
    Exit Sub
ErrTrap:
End Sub
Private Sub btn_Last_Click(Index As Integer)
  On Error GoTo ErrTrap

    Dim Msg As String

    If Me.TxtModFlg.text = "N" Then
        FindRec val(TxtSerial1(mIndex).text)
        Me.TxtModFlg2(mIndex).text = "R"
    End If

    Me.TxtModFlg2(mIndex) = "R"

    If RsSavRec.RecordCount = 0 Then
        clear_all Me
        Exit Sub
    End If

BegnieWork:

    RsSavRec.MoveLast
    If mIndex = 0 Then
        FiLLTXT1
    
    ElseIf mIndex = 1 Then
        FiLLTXT2
    ElseIf mIndex = 3 Then
        FiLLTXT3
    ElseIf mIndex = 4 Then
        FiLLTXT4
    ElseIf mIndex = 5 Then
        FiLLTXT5
    ElseIf mIndex = 6 Then
        FiLLTXT6
    ElseIf mIndex = 7 Then
        FiLLTXT7
    ElseIf mIndex = 8 Then
        FiLLTXT8
    ElseIf mIndex = 9 Then
        FiLLTXT9
    ElseIf mIndex = 10 Then
        FiLLTXT10
    ElseIf mIndex = 11 Then
        FiLLTXT11
    ElseIf mIndex = 12 Then
        FiLLTXT12
        
    
    End If
    
    Exit Sub

ErrTrap:

    Select Case Err.Number

        Case -2147217885
       If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· " & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
        Else
            Msg = "Sorry I have been deleted the record" & CHR(13)
            Msg = Msg & "By another user on the network " & CHR(13)
            Msg = Msg & "Data will be updated"
        End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.Title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub


Private Sub cmdCreateSales_Click()

If TxtNoteSerial13 <> "" Then
    
    
    frmsalebill5.show
   'frmsalebill5.XPBtnMove_Click (2)
    frmsalebill5.Retrive val(TXTTransactionID3.text)
    frmsalebill5.DBCboClientName.Visible = True
    frmsalebill5.lbl(7).Visible = True
    
    
End If


End Sub

Private Sub cmdPrintCash_Click()
  
 
 If txtNoteSerialCash(0) <> "" Then
                print_reportCash txtNoteSerialCash(0), txtNoteSerialCash(0), "", "", "", DcCustmer.text
    End If
End Sub


Public Function print_reportCash(Optional NoteSerial As String, Optional NoteSerial1 As String, Optional BankName As String, Optional PaymentType As String, Optional Box As String, Optional Custcode As String)
    
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String

   ' MySQL = "Select * From payment_voucher  where NoteID=" & val(XPTxtID.Text)
MySQL = "SELECT BillMaintNo, Notes.paydes,    dbo.Notes.Note_Value, dbo.Notes.BankID, dbo.Notes.ChqueNum, dbo.BanksData.BankName, dbo.Notes.NoteType, dbo.Notes.BoxID, dbo.TblBoxesData.BoxName, "
MySQL = MySQL & "                        dbo.Notes.CusID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.Notes.Remark, dbo.Notes.NoteSerial, dbo.Notes.NoteDate,"
MySQL = MySQL & "                                 dbo.Notes.note_value_by_characters, dbo.Notes.NoteID, dbo.Notes.general_des_notes, dbo.Notes.person, dbo.TblCustemers.Fullcode, dbo.Notes.PreVAT,"
MySQL = MySQL & "                                 dbo.Notes.Vat , dbo.Notes.NoteSerial1, dbo.Notes.ManulaNO, dbo.Notes.ManualNO"
MySQL = MySQL & "           FROM         dbo.Notes LEFT OUTER JOIN"
MySQL = MySQL & "                                 dbo.TblBoxesData ON dbo.Notes.BoxID = dbo.TblBoxesData.BoxID LEFT OUTER JOIN"
MySQL = MySQL & "                                 dbo.TblCustemers ON dbo.Notes.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
MySQL = MySQL & "                                 dbo.BanksData ON dbo.Notes.BankID = dbo.BanksData.BankID"
MySQL = MySQL & "           Where (dbo.Notes.NoteType = 4)"
MySQL = MySQL & "           and NoteID=" & val(txtNoteSerialCash(1).text)

MySQL = "SELECT BillMaintNo, Notes.paydes,    dbo.Notes.Note_Value, dbo.Notes.BankID, dbo.Notes.ChqueNum, dbo.BanksData.BankName, dbo.Notes.NoteType, dbo.Notes.BoxID, dbo.TblBoxesData.BoxName, "
MySQL = MySQL & "                        dbo.Notes.CusID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.Notes.Remark, dbo.Notes.NoteSerial, dbo.Notes.NoteDate,"
MySQL = MySQL & "                                 dbo.Notes.note_value_by_characters,Notes.NoteCashingType, dbo.Notes.NoteID, dbo.Notes.general_des_notes, dbo.Notes.person, dbo.TblCustemers.Fullcode, dbo.Notes.PreVAT,"
MySQL = MySQL & "                                 dbo.Notes.Vat , dbo.Notes.NoteSerial1, dbo.Notes.ManulaNO, dbo.Notes.ManualNO"
MySQL = MySQL & "           FROM         dbo.Notes LEFT OUTER JOIN"
MySQL = MySQL & "                                 dbo.TblBoxesData ON dbo.Notes.BoxID = dbo.TblBoxesData.BoxID LEFT OUTER JOIN"
MySQL = MySQL & "                                 dbo.TblCustemers ON dbo.Notes.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
MySQL = MySQL & "                                 dbo.BanksData ON dbo.Notes.BankID = dbo.BanksData.BankID"
MySQL = MySQL & "           Where (dbo.Notes.NoteType = 4)"
MySQL = MySQL & "           and NoteID=" & val(txtNoteSerialCash(1).text)

    If SystemOptions.UserInterface = ArabicInterface Then
    '    StrFileName = App.path & "\Reports\" & "Payment_voucher.rpt"
        StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\Payment_voucher.rpt"
    Else
     '   StrFileName = App.path & "\Reports\" & "Payment_voucherE.rpt"
        StrFileName = App.path & "\Special\" & SystemOptions.Reportpath & "\Payment_voucher.rpt"
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
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
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
        xReport.ParameterFields(5).AddCurrentValue "" '''DcboCreditSide.Text
         xReport.ParameterFields(5).AddCurrentValue DcCustmer.text
   
   
        StrReportTitle = "" '& StrAccountName
 
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
    
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(mBranchID))
        xReport.ParameterFields(5).AddCurrentValue "" 'DcboCreditSide.Text
        StrReportTitle = ""
 
    End If
Dim i As Integer
Dim str As String
'With Grid5
'str = ""
'For i = 1 To .Rows - 1
'If (.TextMatrix(i, .ColIndex("NoteSerial1"))) <> "" And .Cell(flexcpChecked, i, .ColIndex("payed")) = flexChecked Then
'str = str & .TextMatrix(i, .ColIndex("NoteSerial1"))
'If i <> (.Rows - 1) Then
'str = str & ","
'End If
'End If
'Next i

    xReport.ParameterFields(3).AddCurrentValue user_name
    '
    xReport.ParameterFields(6).AddCurrentValue NoteSerial1

    xReport.ParameterFields(7).AddCurrentValue BankName
    xReport.ParameterFields(8).AddCurrentValue PaymentType
    xReport.ParameterFields(9).AddCurrentValue Box
    xReport.ParameterFields(10).AddCurrentValue Custcode
    xReport.ParameterFields(11).AddCurrentValue str
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName, , MySQL

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault

End Function
 


Private Sub CalcAmount()
   Dim i As Long
   txtPercentTotal = 0
   txtGeneralTotal = 0
   For i = 1 To FG.rows - 1
        FG.TextMatrix(i, FG.ColIndex("Amount")) = val(txtRequiredAmount) * val(FG.TextMatrix(i, FG.ColIndex("PercentV"))) / 100
        txtPercentTotal = val(txtPercentTotal) + val(FG.TextMatrix(i, FG.ColIndex("PercentV")))
        txtGeneralTotal = val(txtGeneralTotal) + val(FG.TextMatrix(i, FG.ColIndex("Amount0")))
        If val(txtPercentTotal) > 100 Then
            MsgBox "·« Ì„ﬂ‰ «‰   ⁄œÏ «·‰”»… 100% »—Ã«¡  ⁄œÌ· «·‰”»… «·«ŒÌ—…"
            txtPercentTotal = val(txtPercentTotal) - val(FG.TextMatrix(i, FG.ColIndex("PercentV")))
            FG.TextMatrix(i, FG.ColIndex("PercentV")) = 100 - val(txtPercentTotal)
            FG.TextMatrix(i, FG.ColIndex("Amount")) = val(txtRequiredAmount) * val(FG.TextMatrix(i, FG.ColIndex("PercentV"))) / 100
            txtPercentTotal = 100
            'fg.TextMatrix(i, fg.ColIndex("Amount")) = 0
            Exit Sub
        End If
   Next
   Calc
End Sub

Private Sub Command1_Click()

 Dim Msg As String, AskOption As String
 Dim SaleReport As ClsSaleReport


     If DoPremis(Do_Print, Me.Name, True) = False Then
                Exit Sub
            End If

            If Me.TXTTransactionID3.text = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "·« ÊÃœ ›Ê« Ì— ·Ì „ ÿ»«⁄ Â«"
                Else
                    Msg = "There are no invoices to print"
                End If
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Exit Sub
            End If

            AskOption = GetSetting(StrAppRegPath, "View_Type", "ShowMe", False)

          If AskOption = False Then
'             FrmSallReportOptions.show vbModal
'
'              If FrmSallReportOptions.UserCanceled = True Then
'                   Unload FrmSallReportOptions
'
'             Exit Sub
'               End If
'
'            Unload FrmSallReportOptions
            End If
        updateCopyNo "Transactions", "CopyNO", "Transaction_ID", val(Me.TXTTransactionID3.text)
        
        If TXTTransactionID3.text <> "" Then
            Set SaleReport = New ClsSaleReport
           ' SaleReport.ShowSallingDataDetailed TXTTransactionID3.Text, 18, , , Round(val(txtTotalAfterVat), SystemOptions.Count_ACCOUNT_digit), DcCustmer.BoundText, , , , , , XPDtbTrans(mIndex).value, , , , , , , , , , , , val(Me.Dcbranch(mIndex).BoundText)
        
        
            Set SaleReport = New ClsSaleReport
            SaleReport.ShowSallingDataDetailed TXTTransactionID3.text, , , , Round(val(Me.txtTotalAfterVat.text), SystemOptions.Count_ACCOUNT_digit), DcCustmer.BoundText, , , , , , XPDtbTrans(mIndex).value, , , , , , , , , , , , val(Me.dcBranch(mIndex).BoundText)

            '  If MDIFrmMain.MnuInvPrintReceipt.Checked = True Then
            '      SaleReport.PrintInvoiceReceipt Val(XPTxtBillID.text), P_Target
            '  End If
        End If
        RsSavRec.Resync adAffectCurrent
       
End Sub

Private Sub Command2_Click()
    FrmCashing.show
    FrmCashing.Retrive val(txtNoteSerialCash(1).text)
 
End Sub

Private Sub FG_AfterEdit(ByVal Row As Long, ByVal Col As Long)
'****
Dim StrAccountCode As String, LngRow As Long
Dim rsDummy As New ADODB.Recordset
Dim s As String
With FG
'CalcAmount
   Select Case .ColKey(Col)
    Case "TasksName"
             StrAccountCode = .ComboData
                 LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("TasksID"), False, True)
                .TextMatrix(Row, .ColIndex("TasksID")) = StrAccountCode
                s = "Select PercentV from TblTasks Where Id = " & val(StrAccountCode)
                rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly
                If Not rsDummy.EOF Then
                    .TextMatrix(Row, .ColIndex("PercentV")) = rsDummy!PercentV & ""
                End If
                
    Case "TasksID"
                
                 Dim mTaskId As Long
                mTaskId = val(.TextMatrix(Row, .ColIndex("TasksID")))
                s = "Select * from TblTasks Where Id = " & val(mTaskId)
                rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly
                If Not rsDummy.EOF Then
                    .TextMatrix(Row, .ColIndex("PercentV")) = rsDummy!PercentV & ""
                    .TextMatrix(Row, .ColIndex("TasksName")) = rsDummy!Name & ""
                End If

                
   Case "Emp_ID", "Emp_Name"
             StrAccountCode = .ComboData
                 LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("Emp_ID"), False, True)
                .TextMatrix(Row, .ColIndex("Emp_ID")) = StrAccountCode
   Case "ItemID", "ItemName"
             StrAccountCode = .ComboData
                 LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("ItemID"), False, True)
                .TextMatrix(Row, .ColIndex("ItemID")) = StrAccountCode
                 
                 
                '.TextMatrix(Row, .ColIndex("TasksName")) = StrAccountCode
    End Select
    
End With
CalcAmount
End Sub
Private Sub FG_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

With FG

   Select Case .ColKey(Col)
        Case "Amount0", "Amount2", "Amount3", "PercentV", "Amount", "DateStart", "DateEnd", "RemarkItem"
        .ComboList = ""
           Case "NoteNo"
        .ComboList = ""
        Case "DayMeter"
        .ComboList = ""
        Case "Name"
       ' Cancel = True
        End Select
        
    End With
 
End Sub

Private Sub fg_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
  
    Dim rs As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim Msg As String
    Dim StrAccountCode As String
    Dim StrAccountCode1 As String

    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long

    With FG

        Select Case .ColKey(Col)
 
            Case "TasksName"
             .TextMatrix(Row, .ColIndex("TasksName")) = ""
                StrSQL = "select * from TblTasks "
                Set rs = New ADODB.Recordset
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = FG.BuildComboList(rs, "Name", "ID")
                Else
                    StrComboList = FG.BuildComboList(rs, "Namee", "ID")
                End If
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If
                 .ComboList = StrComboList
            Case "ItemName"
             .TextMatrix(Row, .ColIndex("ItemName")) = ""
                StrSQL = "select * from tblItems where groupId In (Select Groups.GroupId from groups where groups.BranchID =  " & mBranchID & " and dbo.Groups.Separate = 1) "
                
                Set rs = New ADODB.Recordset
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = FG.BuildComboList(rs, "ItemName", "ItemID")
                Else
                    StrComboList = FG.BuildComboList(rs, "ItemNamee", "ItemID")
                End If
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If
                 .ComboList = StrComboList
                 
            Case "Emp_Name"
             .TextMatrix(Row, .ColIndex("Emp_Name")) = ""
                StrSQL = "select * from TblEmployee where IsNull(chkShowTasks,0) = 1 and BranchId = " & mBranchID
                Set rs = New ADODB.Recordset
                rs.Open StrSQL, Cn, adOpenForwardOnly, adLockPessimistic

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = FG.BuildComboList(rs, "Emp_Name", "Emp_ID")
                Else
                    StrComboList = FG.BuildComboList(rs, "Emp_Namee", "Emp_ID")
                End If
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If
                 .ComboList = StrComboList
                 
                 
            End Select
        End With
End Sub







Private Sub fg4_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim StrAccountCode As String, LngRow As Long
Dim rsDummy As New ADODB.Recordset
Dim s As String
With FG4

   Select Case .ColKey(Col)
    Case "TasksID", "TasksName"
             StrAccountCode = .ComboData
                 LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("TasksID"), False, True)
                .TextMatrix(Row, .ColIndex("TasksID")) = StrAccountCode
                s = "Select PercentV from TblTasks Where Id = " & val(StrAccountCode)
                s = "select TblTasks.ID,TblTasks.PercentV,TblJobOrders2.AMount from TblTasks Inner join TblJobOrders2 On TblJobOrders2.TasksID =TblTasks.Id Where MasterID =   " & val(.TextMatrix(Row, .ColIndex("JobOrdersNo"))) & " And TblTasks.Id  = " & val(.TextMatrix(Row, .ColIndex("TasksID")))
                rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly
                If Not rsDummy.EOF Then
                    .TextMatrix(Row, .ColIndex("TasksID")) = rsDummy!ID & ""
                    .TextMatrix(Row, .ColIndex("PercentV")) = rsDummy!PercentV & ""
                    .TextMatrix(Row, .ColIndex("Amount")) = rsDummy!Amount & ""
              
                    
                End If
     Case "JobOrdersNo"
        s = "select TblCustemers.CusID,TblCustemers.CusName CustName,TblEmployee.Emp_ID EmpID,TblEmployee.Emp_Name EmpName,tblItems.ItemName, TblJobOrders2.*,      TblJobOrders.TotalAfterVat, TblTasks.PercentV "
        s = s & " from  TblJobOrders "
        s = s & " INNER JOIN TblCustemers"
        s = s & "             ON  TblCustemers.CusID = TblJobOrders.CusID"
        
        s = s & "             INNER JOIN TblJobOrders2"
        s = s & "             ON TblJobOrders.ID = TblJobOrders2.MasterID"
                s = s & "        Left Outer join tblItems"
        s = s & "             ON  tblItems.ItemID = TblJobOrders2.ItemID"
        s = s & "        Left Outer JOIN TblEmployee"
        s = s & "             ON  TblEmployee.Emp_Id = TblJobOrders2.Emp_Id"
        s = s & "             Left Outer JOIN  TblTasks"
        s = s & "             ON TblJobOrders2.TasksID = TblTasks.ID"
        
        '
        s = s & " Where TblJobOrders.ID =   " & val(.TextMatrix(Row, .ColIndex("JobOrdersNo"))) & " "
        'And TblTasks.Id  = " & val(.TextMatrix(Row, .ColIndex("TasksID")))
        rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly
        If Not rsDummy.EOF Then
            .TextMatrix(Row, .ColIndex("EmpName")) = rsDummy!EmpName & ""
            '.TextMatrix(Row, .ColIndex("EmpID")) = rsDummy!EmpID & ""
            .TextMatrix(Row, .ColIndex("CusID")) = rsDummy!CusID & ""
            .TextMatrix(Row, .ColIndex("CustName")) = rsDummy!CustName & ""
            .TextMatrix(Row, .ColIndex("EmpID")) = rsDummy!EmpID & ""
            .TextMatrix(Row, .ColIndex("ItemName")) = rsDummy!ItemName & ""
             .TextMatrix(Row, .ColIndex("PercentV")) = rsDummy!PercentV & ""
             .TextMatrix(Row, .ColIndex("Amount")) = rsDummy!Amount & ""
             .TextMatrix(Row, .ColIndex("DateStart")) = rsDummy!DateStart & ""
             
             '.TextMatrix(Row, .ColIndex("Total")) = val(rsDummy!Amount & "") * val(.TextMatrix(Row, .ColIndex("Hours")))
                   .TextMatrix(Row, .ColIndex("Total")) = rsDummy!TotalAfterVat & ""
            
            
        End If
    Case "PercentV"
        If val(.TextMatrix(Row, .ColIndex("PercentV"))) <> 0 Then
            .TextMatrix(Row, .ColIndex("Amount")) = val(.TextMatrix(Row, .ColIndex("Total"))) * val(.TextMatrix(Row, .ColIndex("PercentV"))) / 100
        End If
    Case "Hours"
       ' .TextMatrix(Row, .ColIndex("Total")) = val(.TextMatrix(Row, .ColIndex("Amount"))) * val(.TextMatrix(Row, .ColIndex("Hours")))
                '.TextMatrix(Row, .ColIndex("TasksName")) = StrAccountCode
    End Select
  '  CalcAmount
        If Row = .rows - 1 Then
            .rows = .rows + 1
        End If
End With
End Sub
Private Sub fg4_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

With FG4

   Select Case .ColKey(Col)
        Case "Amount0", "Amount2", "Amount3", "PercentV", "Amount", "DateStart", "DateEnd", "JobOrdersNo", "Hours"
        .ComboList = ""
           Case "NoteNo"
        .ComboList = ""
        Case "DayMeter"
        .ComboList = ""
        Case "ItemName", "CustName", "PercentV"
        Cancel = True
        End Select
        
    End With
 
End Sub

Private Sub fg4_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
  
    Dim rs As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim Msg As String
    Dim StrAccountCode As String
    Dim StrAccountCode1 As String

    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long

    With FG4

        Select Case .ColKey(Col)
 
            Case "TasksName"
             .TextMatrix(Row, .ColIndex("TasksName")) = ""
                StrSQL = "select TblTasks.ID,TblTasks.Name,Namee from TblTasks Inner join TblJobOrders2 On TblJobOrders2.TasksID =TblTasks.Id Where MasterID =   " & val(.TextMatrix(Row, .ColIndex("JobOrdersNo")))
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = FG4.BuildComboList(rs, "Name", "ID")
                Else
                    StrComboList = FG4.BuildComboList(rs, "Namee", "ID")
                End If
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If
                 .ComboList = StrComboList
            Case "EmpName"
             .TextMatrix(Row, .ColIndex("EmpName")) = ""
                StrSQL = "SELECT Emp_Id,Emp_Name,Emp_Namee FROM TblEmployee "
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = Grid.BuildComboList(rs, "Emp_Name", "Emp_Id")
                Else
                    StrComboList = Grid.BuildComboList(rs, "Emp_Namee", "Emp_Id")
                End If
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If
                 .ComboList = StrComboList
            End Select
        End With
End Sub


Private Sub fg6_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim StrAccountCode As String, LngRow As Long
Dim rsDummy As New ADODB.Recordset
Dim s As String
With Fg6

   Select Case .ColKey(Col)
    Case "CusID", "CustName"
             StrAccountCode = .ComboData
                 LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("CusID"), False, True)
                .TextMatrix(Row, .ColIndex("CusID")) = StrAccountCode
'                s = "Select PercentV from TblTasks Where Id = " & val(StrAccountCode)
'                s = "select TblTasks.ID,TblTasks.PercentV from TblTasks Where TblTasks.Id  = " & val(.TextMatrix(Row, .ColIndex("TasksID")))
'                rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly
'                If Not rsDummy.EOF Then
'                    .TextMatrix(Row, .ColIndex("TasksID")) = rsDummy!ID & ""
'                    '.TextMatrix(Row, .ColIndex("PercentV")) = rsDummy!PercentV & ""
'                End If
  
    Case "EmpID", "EmpName"
             StrAccountCode = .ComboData
                 LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("EmpName"), False, True)
                .TextMatrix(Row, .ColIndex("EmpID")) = StrAccountCode
'                s = "Select PercentV from TblTasks Where Id = " & val(StrAccountCode)
'                s = "select TblTasks.ID,TblTasks.PercentV from TblTasks Where TblTasks.Id  = " & val(.TextMatrix(Row, .ColIndex("TasksID")))
'                rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly
'                If Not rsDummy.EOF Then
'                    .TextMatrix(Row, .ColIndex("TasksID")) = rsDummy!ID & ""
'                    '.TextMatrix(Row, .ColIndex("PercentV")) = rsDummy!PercentV & ""
'                End If

    Case "ReservationTypeCode", "ReservationTypeName"
             StrAccountCode = .ComboData
                 LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("ReservationTypeName"), False, True)
                .TextMatrix(Row, .ColIndex("ReservationTypeCode")) = StrAccountCode
'                s = "Select PercentV from TblTasks Where Id = " & val(StrAccountCode)
'                s = "select TblTasks.ID,TblTasks.PercentV from TblTasks Where TblTasks.Id  = " & val(.TextMatrix(Row, .ColIndex("TasksID")))
'                rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly
'                If Not rsDummy.EOF Then
'                    .TextMatrix(Row, .ColIndex("TasksID")) = rsDummy!ID & ""
'                    '.TextMatrix(Row, .ColIndex("PercentV")) = rsDummy!PercentV & ""
'                End If
    Case "Hours"
        '.TextMatrix(Row, .ColIndex("Total")) = val(.TextMatrix(Row, .ColIndex("Amount"))) * val(.TextMatrix(Row, .ColIndex("Hours")))
                '.TextMatrix(Row, .ColIndex("TasksName")) = StrAccountCode
         Fg6.TextMatrix(Row, Fg6.ColIndex("StillPeriod")) = GetTimeDiff(Fg6.TextMatrix(Row, Fg6.ColIndex("Hours")), Time, 1, 1) & "  ”«⁄«  "
            Fg6.TextMatrix(Row, Fg6.ColIndex("minutes")) = GetTimeDiff(Fg6.TextMatrix(Row, Fg6.ColIndex("Hours")), Time, 1, 2) & " œﬁÌﬁ… "

    End Select
  '  CalcAmount
        If Row = .rows - 1 Then
            .rows = .rows + 1
        End If
End With
End Sub
Private Sub fg6_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

With Fg6

   Select Case .ColKey(Col)
        Case "Amount0", "Amount2", "Amount3", "PercentV", "Amount", "DateStart", "DateEnd", "JobOrdersNo", "Hours", "minutes", "ReservNo", "TimeR", "ServiceNo"
            .ComboList = ""
        Case "NoteNo"
            .ComboList = ""
        Case "DayMeter"
            .ComboList = ""
        Case "ItemName", "PercentV"
            Cancel = True
        End Select
        
    End With
 
End Sub

Private Sub fg6_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
  
    Dim rs As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim Msg As String
    Dim StrAccountCode As String
    Dim StrAccountCode1 As String

    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long

    With Fg6

        Select Case .ColKey(Col)
 
            Case "CustName"
             .TextMatrix(Row, .ColIndex("CustName")) = ""
                StrSQL = "select CusID,CusName,CusNamee from TblCustemers "
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = Fg6.BuildComboList(rs, "CusName", "CusID")
                Else
                    StrComboList = Fg6.BuildComboList(rs, "CusNamee", "CusID")
                End If
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If
                 .ComboList = StrComboList
            Case "EmpName"
             .TextMatrix(Row, .ColIndex("EmpName")) = ""
                StrSQL = "SELECT Emp_Id,Emp_Name,Emp_Namee FROM TblEmployee "
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = Fg6.BuildComboList(rs, "Emp_Name", "Emp_Id")
                Else
                    StrComboList = Fg6.BuildComboList(rs, "Emp_Namee", "Emp_Id")
                End If
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If
                 .ComboList = StrComboList
            Case "ReservationTypeName"
                .TextMatrix(Row, .ColIndex("ReservationTypeName")) = ""
                StrSQL = "SELECT Id,Name,Namee FROM tblReservationType "
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = Fg6.BuildComboList(rs, "Name", "Id")
                Else
                    StrComboList = Fg6.BuildComboList(rs, "Namee", "Id")
                End If
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If
                 .ComboList = StrComboList
                 
                 
            End Select
        End With
End Sub



 

Private Sub Grid1_EnterCell()
  On Error GoTo ErrTrap
    FindRec val(Me.GRID1.TextMatrix(Me.GRID1.Row, Me.GRID1.ColIndex("id")))
ErrTrap:
End Sub

Private Sub Grid8_EnterCell()
  On Error GoTo ErrTrap
    FindRec val(Me.Grid8.TextMatrix(Me.Grid8.Row, Me.Grid8.ColIndex("id")))
ErrTrap:
End Sub


Private Sub Grid2_EnterCell()
    On Error GoTo ErrTrap
    FindRec val(Me.Grid2.TextMatrix(Me.Grid2.Row, Me.Grid2.ColIndex("id")))
ErrTrap:
End Sub



Private Sub Grid5_EnterCell()
  On Error GoTo ErrTrap
    FindRec val(Me.Grid5.TextMatrix(Me.Grid5.Row, Me.Grid5.ColIndex("id")))
ErrTrap:
End Sub




Private Sub ISButton1_Click()
  Dim k As Long, LngNewRow As Long
  If Trim(FG4.TextMatrix(FG4.rows - 1, FG4.ColIndex("JobOrdersNo"))) = "" Then
        FG4.rows = FG4.rows - 1
    End If
    If FG4.rows = 1 Then FG4.rows = 2 Else FG4.rows = FG4.rows + 1
    
    
    k = FG4.rows
   
    If FG4.rows <= 1 Then
        FG4.rows = FG4.rows + 1
    End If
    LngNewRow = FG4.rows - 1
     'mNewId = LngNewRow

End Sub

'Private Sub ISButton3_Click()
'  Dim k As Long, LngNewRow As Long
'  If Trim(FG.TextMatrix(FG.Rows - 1, FG.ColIndex("TasksName"))) = "" Then
'        FG.Rows = FG.Rows - 1
'    End If
'    If FG.Rows = 1 Then FG.Rows = 2 Else FG.Rows = FG.Rows + 1
'
'
'    k = FG.Rows
'
'    If FG.Rows <= 1 Then
'        FG.Rows = FG.Rows + 1
'    End If
'    LngNewRow = FG.Rows - 1
'     'mNewId = LngNewRow
'
'
'
'
'
'
'
'    'fg.TextMatrix(LngNewRow, fg.ColIndex("Amount0")) = txtAmount
'    'fg.TextMatrix(LngNewRow, fg.ColIndex("Amount")) = txtAmount
'    FG.TextMatrix(LngNewRow, FG.ColIndex("DateStart")) = txtDateStart.value
'    'fg.TextMatrix(LngNewRow, fg.ColIndex("DateEnd")) = txtDateEnd.value
'
''
''    Dim rsDummy As New ADODB.Recordset
''    Dim s As String
''    s = "Select PercentV from TblTasks Where Id = " & val(cmbTasks.BoundText)
''    rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly
''    If Not rsDummy.EOF Then
''        FG.TextMatrix(LngNewRow, FG.ColIndex("PercentV")) = rsDummy!PercentV & ""
''    End If
'    CalcAmount
'
''    Fg_AfterEdit LngNewRow, fg.ColIndex("TasksName")
'End Sub
'
'
Private Sub Calc()
If val(txtVatYou) = 0 Then txtVatYou = 15
txtRequiredAmount = Round(val(txtGeneralTotal) + val(txtTotalAdd), 2)
'txtTotalNet = val(txtGeneralTotal) + val(txtTotalAdd) - val(txtTotalDisc)

txtTotalDisc = Round(val(txtRequiredAmount) * val(txtTotalDiscPerc) / 100, 2)
txtRequiredAmount = Round(val(txtGeneralTotal) + val(txtTotalAdd) - val(txtTotalDisc), 2)
TxtVAT = val(txtRequiredAmount) * val(txtVatYou) / 100
txtTotalAfterVat = val(txtRequiredAmount) + val(TxtVAT)
txtTotalNet = Round(val(txtTotalAfterVat) - val(txtPaymedValue), 2)

End Sub

Private Sub txtGeneralTotal_Validate(Cancel As Boolean)

Calc
End Sub

Private Sub TxtModFlg2_Change(Index As Integer)
 On Error GoTo ErrTrap

    Select Case Me.TxtModFlg2(mIndex).text

        Case "R"
            '        Me.Caption = " ’—ÌÕ Œ—ÊÃ „ƒﬁ "
            Me.btn_Save(Index).Enabled = False
            Me.Btn_Undo(Index).Enabled = False
            Me.btn_New(Index).Enabled = True
            Me.btn_Modify(Index).Enabled = True
            Me.btn_Delete(Index).Enabled = True
            Me.btn_Query(Index).Enabled = True
            btn_Previous(Index).Enabled = True
            btn_First(Index).Enabled = True
            btn_Last(Index).Enabled = True
            btn_Next(Index).Enabled = True
            'Frame1(mIndex).Enabled = False

       
        
         
       
'            If rs.RecordCount < 1 Then
'                btn_Previous(Index).Enabled = False
'                btn_First(Index).Enabled = False
'                btn_Last(Index).Enabled = False
'                btn_Next(Index).Enabled = False
'                Me.btn_Modify(Index).Enabled = False
'                Me.btn_Delete(Index).Enabled = False
'            End If

        Case "N"
            '        Me.Caption = " ’—ÌÕ Œ—ÊÃ „ƒﬁ ( ÃœÌœ )"
            Me.btn_Save(Index).Enabled = True
            Me.Btn_Undo(Index).Enabled = True
            Me.btn_New(Index).Enabled = False
            Me.btn_Modify(Index).Enabled = False
            Me.btn_Delete(Index).Enabled = False
            Me.btn_Query(Index).Enabled = False
            '      btn_Previous(Index).Enabled = False
            '      btn_First(Index).Enabled = False
            '      btn_Last(Index).Enabled = False
            '      btn_Next(Index).Enabled = False
           
'            XPDtbTrans.Enabled = True
'            XPDtbTrans.value = Date
       ' Frame1(mIndex).Enabled = True

        Case "E"
            '        Me.Caption = " ’—ÌÕ Œ—ÊÃ „ƒﬁ (  ⁄œÌ· )"
            Me.btn_Save(Index).Enabled = True
            Me.Btn_Undo(Index).Enabled = True
            Me.btn_New(Index).Enabled = False
            Me.btn_Modify(Index).Enabled = False
            Me.btn_Delete(Index).Enabled = False
            Me.btn_Query(Index).Enabled = False
            
            btn_Previous(Index).Enabled = False
            btn_First(Index).Enabled = False
            btn_Last(Index).Enabled = False
            btn_Next(Index).Enabled = False
      
     '   Frame1(mIndex).Enabled = True
           ' XPDtbTrans.Enabled = True
    End Select

    Exit Sub
ErrTrap:
End Sub

 

 

Private Sub DboAccount_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF3 Then
        Account_search.show
        Account_search.case_id = 196
    End If
    
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

     
    Exit Sub
ErrTrap:
End Sub

Private Sub BtnCancel_Click()
    Unload Me
End Sub

Private Sub btnDelete_Click()
    Dim MSGType As Integer
    Dim BlnRecExist As Boolean
    Dim StrMSG  As String
    Dim Msg As String
    On Error GoTo ErrTrap

    If DoPremis(Do_Delete, Me.Name, True) = False Then
        Exit Sub
    End If

    'If TxtVac_ID.text <> "" Then
    '    If CheckDelCountry(Val(Me.TxtVac_ID.text)) = False Then
    '        Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã·...!!!"
    '        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    '        Exit Sub
    '    End If
    MSGType = MsgBox("Â·  —€» ›Ì Õ–› Â–« «·”Ã·", vbQuestion + vbMsgBoxRtlReading + vbYesNo + vbMsgBoxRight, App.Title)

    If MSGType = vbYes Then
        RsSavRec.Find "id=" & val(TxtVac_ID.text), , adSearchForward, 1
        RsSavRec.delete
        MsgBox " „  ⁄„·Ì… «·Õ–› »‰Ã«Õ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
        '------------------------------ Move Next ---------------------------.
        FillGridWithData
        BtnNext_Click
    End If

    'End If
    Exit Sub
ErrTrap:
 
    Select Case Err.Number

        Case -2147217873, -2147467259
            StrMSG = "⁄›Ê« ·« ÌÃÊ“ Õ–› «·”Ã· ·«— »«ÿÂ »»Ì«‰«  √Œ—Ì"
            RsSavRec.CancelUpdate
            MsgBox StrMSG, vbOKOnly + vbMsgBoxRight, App.Title
            'clear the ConnectiOn Errors
            Cn.Errors.Clear
    End Select

End Sub

Private Sub BtnFirst_Click()
    On Error GoTo ErrTrap

    Dim Msg As String

    If Me.TxtModFlg.text = "N" Then
        FindRec val(TxtVac_ID.text)
        Me.TxtModFlg.text = "R"
    End If

    TxtModFlg = "R"

    If RsSavRec.RecordCount = 0 Then
        clear_all Me
        Exit Sub
    End If

BegnieWork:
    RsSavRec.MoveFirst
    FiLLTXT

    Exit Sub

ErrTrap:

    Select Case Err.Number

        Case -2147217885
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.Title
            RsSavRec.Requery
            Resume BegnieWork
    End Select

End Sub

Private Sub BtnLast_Click()
    On Error GoTo ErrTrap

    Dim Msg As String

    If Me.TxtModFlg.text = "N" Then
        FindRec val(TxtVac_ID.text)
        Me.TxtModFlg.text = "R"
    End If

    TxtModFlg = "R"

    If RsSavRec.RecordCount = 0 Then
        clear_all Me
        Exit Sub
    End If

BegnieWork:

    RsSavRec.MoveLast
    FiLLTXT
    Exit Sub

ErrTrap:

    Select Case Err.Number

        Case -2147217885
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.Title
            RsSavRec.Requery
            Resume BegnieWork
    End Select

End Sub

Private Sub btnModify_Click()
    Dim Msg As String

    If DoPremis(Do_Edit, Me.Name, True) = False Then
        Exit Sub
    End If

    On Error GoTo ErrTrap

    If TxtVac_ID.text <> "" Then
        TxtModFlg = "E"
        Frm2.Enabled = True
        Me.TxtVacName.SetFocus
    End If

    Exit Sub
ErrTrap:

    Select Case Err.Number

        Case -2147467259
            'Could not update; currently locked.
            Msg = "⁄›Ê«" & CHR(13)
            Msg = Msg & " ·«Ì„ﬂ‰  ⁄œÌ· Â–« «·”Ã· ›Ï «·Êﬁ  «·Õ«·Ï" & CHR(13)
            Msg = Msg & "ÕÌÀ «‰Â ﬁÌœ «· ⁄œÌ· „‰ ﬁ»· „” Œœ„ «Œ— ⁄·Ï «·‘»ﬂ…"
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.Title
        
            If RsSavRec.EditMode <> adEditNone Then
                RsSavRec.CancelUpdate
                'RsSavRec.Requery
            End If

    End Select

End Sub

Private Sub btnNew_Click()
    Dim My_SQL As String
    Dim rs As ADODB.Recordset

    If DoPremis(Do_New, Me.Name, True) = False Then
        Exit Sub
    End If

    On Error GoTo ErrTrap
    Set rs = New ADODB.Recordset
    Frm2.Enabled = True
    clear_all Me
    TxtModFlg.text = "N"

    My_SQL = "dean"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable

    If rs.RecordCount > 0 Then
        TxtSerial.text = rs.RecordCount + 1
    Else
        TxtSerial.text = 1
    End If

    rs.Close
  '  CmbType.ListIndex = 0
    TxtVacName.SetFocus
ErrTrap:
End Sub

Private Sub BtnNext_Click()
    On Error GoTo ErrTrap
    Dim Msg As String

    If Me.TxtModFlg.text = "N" Then
        FindRec val(TxtVac_ID.text)
        Me.TxtModFlg.text = "R"
    End If

    TxtModFlg = "R"

    If RsSavRec.RecordCount = 0 Then
        clear_all Me
        Exit Sub
    End If

BegnieWork:

    If RsSavRec.EOF Then
        RsSavRec.MoveLast
    Else
        RsSavRec.MoveNext

        If RsSavRec.EOF Then
            RsSavRec.MoveLast
        End If
    End If

    FiLLTXT
    Exit Sub
ErrTrap:

    Select Case Err.Number

        Case -2147217885
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.Title
            RsSavRec.Requery
            Resume BegnieWork
    End Select

End Sub

Private Sub BtnPrevious_Click()
    On Error GoTo ErrTrap
    Dim Msg As String

    If Me.TxtModFlg.text = "N" Then
        FindRec val(TxtVac_ID.text)
        Me.TxtModFlg.text = "R"
    End If

    TxtModFlg = "R"

    If RsSavRec.RecordCount = 0 Then
        clear_all Me
        Exit Sub
    End If

BegnieWork:

    RsSavRec.MovePrevious

    If RsSavRec.BOF Then
        RsSavRec.MoveFirst
    End If

    FiLLTXT
    Exit Sub
ErrTrap:

    Select Case Err.Number

        Case -2147217885
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.Title
            RsSavRec.Requery
            Resume BegnieWork
    End Select

End Sub

Private Sub btnSave_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
    Dim StrVacCode As String
    Dim StrVacName As String
    Dim CtrlTxt As Control
    '---------------------- check if data Vaclete -----------------------

    For Each CtrlTxt In Me.Controls

        If TypeOf CtrlTxt Is TextBox Or TypeOf CtrlTxt Is ComboBox Then
            If CtrlTxt.text = "" And CtrlTxt.Tag <> "" And CtrlTxt.Enabled = True Then
                MsgBox CtrlTxt.Tag, vbOKOnly + vbMsgBoxRight, App.Title
                CtrlTxt.SetFocus
                Exit Sub
            End If
        End If

    Next

    '------------------------------ check if Empcode exist ----------------------

    StrVacName = IsRecExist("dean", "name", Trim(TxtVacName.text), "name", "Vac_ID<>'" & Trim(TxtVac_ID.text) & "'")

    If StrVacName <> "" Then
        Msg = "·ﬁœ ”»ﬁ  ”ÃÌ· Â–« «·‰Ê⁄ „‰ ﬁ»·"
         
        MsgBox Msg, vbOKOnly + vbMsgBoxRight, App.Title
        TxtVacName.SetFocus
    
        Exit Sub

    End If

    ' -------------------------------------- txtmodflg type -------------------
    Select Case Me.TxtModFlg.text

            '------------------------------ new record ----------------------------
        Case "N"
      
            '------------------------- save record -----------------------------
            AddNewRec
            BtnLast_Click

        Case "E"

            '----------------------------- save edit -------------------------------
            FiLLRec
    End Select

    Exit Sub
ErrTrap:
    MsgBox "Â‰«ﬂ Œÿ√ „« ›Ì ≈œŒ«· «·»Ì«‰« ", vbOKOnly + vbMsgBoxRight, App.Title

End Sub
 
Private Sub BtnUndo_Click()
    FindRec val(TxtVac_ID.text)
    Me.TxtModFlg.text = "R"
End Sub

Private Sub BtnUpdate_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
    Dim FristCount As Long
    Dim LastCount As Long
    FristCount = RsSavRec.RecordCount
    RsSavRec.Requery
    LastCount = RsSavRec.RecordCount
    BtnUndo_Click

    If FristCount = LastCount Then
        Msg = "·«  ÊÃœ »Ì«‰«  ÃœÌœ…"
    Else
        Msg = "⁄œœ «·”Ã·«  ﬁ»· «· ÕœÌÀ" & vbCrLf & FristCount & vbCrLf & "⁄œœ «·”Ã·«  »⁄œ «· ÕœÌÀ" & vbCrLf & LastCount
        
        If LastCount > FristCount Then
            Msg = Msg + vbCrLf & "⁄œœ «·”Ã·«  «·ÃœÌœ…" & vbCrLf & LastCount - FristCount
        Else
            Msg = Msg + vbCrLf & "⁄œœ «·”Ã·«  «·„Õ–Ê›…" & vbCrLf & FristCount - LastCount
        End If
    End If

    MsgBox Msg, vbOKOnly + vbMsgBoxRight + vbInformation, App.Title
ErrTrap:
End Sub

Private Sub Form_Load()
        'On Error GoTo ErrTrap
        On Error Resume Next
    Dim i As Integer
    Dim My_SQL As String
    Dim s As String
'    loadLogo
            If mIndex <> 11 Then ' «—ÌŒ «Œ— »’„Â ·„  Œ—Ã
        XPDtbTrans(mIndex) = Date
        End If
        
        mBranchID = MainBranchID
    Set Dcombos = New ClsDataCombos
    TabMain.TabVisible(1) = False
    TabMain.TabVisible(2) = False
    TabMain.TabVisible(0) = False
    TabMain.TabVisible(3) = False
    TabMain.TabVisible(4) = False
    TabMain.TabVisible(5) = False
    TabMain.TabVisible(6) = False
    TabMain.TabVisible(7) = False
    TabMain.TabVisible(8) = False
    TabMain.TabVisible(9) = False
    TabMain.TabVisible(10) = False
    TabMain.TabVisible(11) = False
    TabMain.TabVisible(12) = False
    mPath = App.path & "\DataFiles\"
  '  TabMain.TabVisible(8) = False

StrSQL = "select Emp_ID,Emp_Name from TblEmployee where IsNull(chkShowTasks,0) = 1 and BranchId = " & mBranchID
        fill_combo CmbEmp, StrSQL

            Dim LOcalCBO As String
           LOcalCBO = " SELECT     dbo.TblItems.ItemID, dbo.TblItems.ItemName  FROM         dbo.Groups INNER JOIN dbo.TblItems ON dbo.Groups.GroupID = dbo.TblItems.GroupID  Where ( Groups.BranchId = " & mBranchID & " )"
        fill_combo cmbItems, LOcalCBO

Set Dcombos = New ClsDataCombos
 Dcombos.GetCustomersSuppliers 1, Me.cmbCustomer, True

'    Me.left = (mdifrmmain.Width - Me.Width) / 2
'    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500
 '   Me.MaxButton = False
 Set rsDummy = New ADODB.Recordset
    s = "Select EmpID from tblUsers where UserId = " & user_id
    rsDummy.Open s, Cn, adOpenForwardOnly, adLockReadOnly
    If Not rsDummy.EOF Then
        mEmpId = val(rsDummy!EmpID & "")
    End If

    If mIndex = 0 Then
        ScreenNameArabic = "«·„Â«„"
        ScreenNameEnglish = "Tasks"
       
        TabMain.TabVisible(0) = True
        TabMain.CurrTab = 0
        Me.Caption = "«·„Â«„"
        My_SQL = "TblTasks"
        Set BKGrndPic = New ClsBackGroundPic
        Set RsSavRec = New ADODB.Recordset
        RsSavRec.CursorLocation = adUseClient
        RsSavRec.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdTableDirect
      TxtModFlg2(mIndex).text = "R"
        btn_First_Click (mIndex)
       ' Me.Width = Grid2.Width + 400
        'FillGridWithData2
        FillGridWithData1
'        Me.Height = 10545
         Me.BorderStyle = FixedSingle
       'Me.MaxButton = False
        ' Me.C1Elastic12.Align = asNone
        Fra_Header(1).Width = GRID1.Width + 400
        'C1Elastic12.Width = Fra_Header(1).Width + 400
        TabMain.Align = asNone
        
        TabMain.Width = Fra_Header(1).Width
      '  Me.Width = TabMain.Width
      '  Me.ScaleWidth = Me.Width
      '   Grid1.Width = TabMain.Width - 400
'         Me.WindowState = 0
   '      Resize_Form Me
         TabMain.Align = asFill
      '   Me.Max
      ' '  Me.C1Elastic12.AutoSizeChildren = azNone
       ' Me.Width = Fra_Header(1).Width + 800
        
        Me.WindowState = 2
    ElseIf mIndex = 1 Then
        ScreenNameArabic = "„”„Ì«  «·„ﬁ«”« "
        ScreenNameEnglish = "Tasks"
       
         TabMain.TabVisible(mIndex) = True
        TabMain.CurrTab = mIndex
        Me.Caption = "„”„Ì«  «·„ﬁ«”« "
        My_SQL = "TblSizesNames"
        Set BKGrndPic = New ClsBackGroundPic
        Set RsSavRec = New ADODB.Recordset
        RsSavRec.CursorLocation = adUseClient
        RsSavRec.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdTableDirect
      TxtModFlg2(mIndex).text = "R"
        btn_First_Click (mIndex)
        
        Me.Width = Grid2.Width + 400
        FillGridWithData2
'    ElseIf mIndex = 2 Then

        
    ElseIf mIndex = 5 Then
        ScreenNameArabic = "«‰Ê«⁄ «·ÕÃ“"
        ScreenNameEnglish = "Tasks"
       
         TabMain.TabVisible(mIndex) = True
        TabMain.CurrTab = mIndex
        Me.Caption = "«‰Ê«⁄ «·ÕÃ“"
        My_SQL = "tblReservationType"
        Set BKGrndPic = New ClsBackGroundPic
        Set RsSavRec = New ADODB.Recordset
        RsSavRec.CursorLocation = adUseClient
        RsSavRec.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdTableDirect
        Frame1(5).Enabled = True
        TxtModFlg2(mIndex).text = "R"
        btn_First_Click (mIndex)
        
        Me.Width = Grid5.Width + 400
        FillGridWithData5
'    ElseIf mIndex = 2 Then
    ElseIf mIndex = 6 Then
        ScreenNameArabic = "⁄—÷ «·ÕÃÊ“« "
        ScreenNameEnglish = "Tasks"
        
         TabMain.TabVisible(mIndex) = True
        TabMain.CurrTab = mIndex
        Me.Caption = ScreenNameArabic
        My_SQL = "TblAppointmentlist"
        Set BKGrndPic = New ClsBackGroundPic
        Set RsSavRec = New ADODB.Recordset
        RsSavRec.CursorLocation = adUseClient
        RsSavRec.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdTableDirect

        Dcombos.GetBranches dcBranch(mIndex)
        XPDtbBill(0).value = Date
        XPDtbBill(1).value = Date
        Me.Width = Fg6.Width + 400
      
        ISButton2_Click
    ElseIf mIndex = 7 Then
        ScreenNameArabic = "—»ÿ «·„ÊŸ›Ì‰ »«·Œœ„«  Ê«·«’‰«›"
        ScreenNameEnglish = "Tasks"
        Dcombos.GetUsers Me.DCboUserName(mIndex)
         TabMain.TabVisible(mIndex) = True
        TabMain.CurrTab = mIndex
        Me.Caption = ScreenNameArabic
        My_SQL = "TblEmpItemsTrans"
        Set BKGrndPic = New ClsBackGroundPic
        Set RsSavRec = New ADODB.Recordset
        RsSavRec.CursorLocation = adUseClient
        RsSavRec.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdTableDirect
      TxtModFlg2(mIndex).text = "R"
        btn_First_Click (mIndex)
       ' XPDtbBill.value = Date
        Me.Width = fg7.Width + 400
        Dcombos.GetBranches dcBranch(mIndex)
        ListProductLineSelected.Clear
        ListGroupSelected.Clear
        FillMylist
  ElseIf mIndex = 8 Then
        ScreenNameArabic = "›∆«  «·”œ«œ"
        ScreenNameEnglish = "Tasks"
       
         TabMain.TabVisible(mIndex) = True
        TabMain.CurrTab = mIndex
        Me.Caption = "›∆«  «·”œ«œ"
        My_SQL = "select * from tblPaymentClass where 1=1"
        Set BKGrndPic = New ClsBackGroundPic
        Set RsSavRec = New ADODB.Recordset
        RsSavRec.CursorLocation = adUseClient
      '  RsSavRec.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdTableDirect
        RsSavRec.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
        
      TxtModFlg2(mIndex).text = "R"
        btn_First_Click (mIndex)
        
        Me.Width = Grid8.Width + 400
        FillGridWithData8
'    ElseIf mIndex = 2 Then
    ElseIf mIndex = 9 Then
    'Me.WindowState = 0
        My_SQL = "TblTripReg"
       ' Set BKGrndPic = New ClsBackGroundPic
        Set RsSavRec = New ADODB.Recordset
        RsSavRec.CursorLocation = adUseClient
        RsSavRec.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdTableDirect
        
        
           TabMain.TabVisible(mIndex) = True
        TabMain.CurrTab = mIndex
     '   Dcombos.GetItemsNames DcboItemID1, -1, -1
        Dcombos.GetUsers Me.DCboUserName(mIndex)
        Dcombos.GetBranches dcBranch(mIndex)
        Dcombos.GetCustomersSuppliers 1, Me.DBCboClientName, True
        Dim sql As String
         
        sql = "SELECT DISTINCT Id, Name ,Namee ,ServiceColor from tblPaymentClass where IsBoardNoHide is null or IsBoardNoHide=0"
        fill_combo cmbPaymentClass, sql
        ReloadCompo
        
        TxtModFlg2(mIndex).text = "R"
        DCboUserName(mIndex).BoundText = user_id
       

        Me.Caption = " ”ÃÌ· œŒÊ· «·„⁄œ« /«·”Ì«—« "
   Me.WindowState = 2
   
 If SystemOptions.UserInterface = ArabicInterface Then

        With CboPayMentType
             .Clear
             .AddItem "‰ﬁœ«"
             .AddItem "„œÏ"
             .AddItem "›Ì“«"
             .AddItem "„«” — ﬂ«—œ"
         End With
         
    Else
         With CboPayMentType
            .Clear
            'AddItem "Cash"
            
            .AddItem "Cash"
            .AddItem "Cheque"
            .AddItem "Visa"
            .AddItem "Master Card"
        End With
        
    End If
        btn_First_Click (mIndex)
         Me.Width = TabMain2.Width + 400
   ElseIf mIndex = 10 Then
     Me.WindowState = 0
        
        My_SQL = "TblEmpData"
       ' Set BKGrndPic = New ClsBackGroundPic
        Set RsSavRec = New ADODB.Recordset
        RsSavRec.CursorLocation = adUseClient
        RsSavRec.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdTableDirect
        
        
           TabMain.TabVisible(mIndex) = True
        TabMain.CurrTab = mIndex
     '   Dcombos.GetItemsNames DcboItemID1, -1, -1
        Dcombos.GetUsers Me.DCboUserName(mIndex)
        Dcombos.GetBranches dcBranch(mIndex)
        
       SetGridFinger
        ReloadCompo
        
        TxtModFlg2(mIndex).text = "R"
        DCboUserName(mIndex).BoundText = user_id
       

        Me.Caption = " ”ÃÌ· »Ì«‰«  «·„ÊŸ›Ì‰"
   
         btn_First_Click (mIndex)
         Me.Width = TabMain2.Width + 400

   ElseIf mIndex = 11 Then
       FingerCount = 0
Me.WindowState = 0

    fpcHandle = ZKFPEngX2.CreateFPCacheDB
    ZKFPEngX2.SensorIndex = 0
      ZKFPEngX2.BeginEnroll
    If ZKFPEngX2.IsRegister Then
        ZKFPEngX2.CancelEnroll
    End If
     If ZKFPEngX2.InitEngine = 0 Then

     End If
'

        My_SQL = "TblEmpDataInOut"
       ' Set BKGrndPic = New ClsBackGroundPic
       's = "Select * from TblEmpDataInOut Where RecordDate =" & SQLDate(XPDtbTrans(mIndex).value, True)
       s = "Select * from TblEmpDataInOut Where RecordDate"
       s = s & "  in ( select max(RecordDate) as RecordDate from TblEmpDataInOut where timeout is null )"
       My_SQL = s
        loadgrid s, GrdEmp, True, False
        Set RsSavRec = New ADODB.Recordset
        RsSavRec.CursorLocation = adUseClient
       ' RsSavRec.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdTableDirect
        RsSavRec.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
        
        If RsSavRec.RecordCount > 0 Then
        XPDtbTrans(11).value = IIf(IsNull(RsSavRec.Fields("recorddate").value), Date, RsSavRec.Fields("recorddate").value)
        Else
        XPDtbTrans(11).value = Date
        End If
        
        TabMain.TabVisible(mIndex) = True
        TabMain.CurrTab = mIndex
     '   Dcombos.GetItemsNames DcboItemID1, -1, -1
      
        'Dcombos.GetBranches dcBranch(mIndex)
        
       
        s = "SELECT DISTINCT Id EmpID, EmpName"
        s = s & " From dbo.TblEmpData"
        s = s & " WHERE     (NOT (EmpName IS NULL)) "
        fill_combo cmbEmpName, s

        
        txtTimeIn.value = Time
        'ReloadCompo
        
       ' TxtModFlg2(mIndex).Text = "R"
        
       

        Me.Caption = " ”ÃÌ· »Ì«‰«  «·Õ÷Ê— Ê«·«‰’—«›"
   
    '     btn_First_Click (mIndex)
         Me.Width = TabMain2.Width + 400

    ElseIf mIndex = 3 Then
    Me.WindowState = 0
        My_SQL = "TblJobOrders"
       ' Set BKGrndPic = New ClsBackGroundPic
        Set RsSavRec = New ADODB.Recordset
        RsSavRec.CursorLocation = adUseClient
        RsSavRec.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdTableDirect
        
        
           TabMain.TabVisible(mIndex) = True
        TabMain.CurrTab = mIndex
'        ReloadCompo
        
     '   Dcombos.GetItemsNames DcboItemID1, -1, -1
        Dcombos.GetUsers Me.DCboUserName(mIndex)

        Dcombos.GetCustomersSuppliers 1, Me.DcCustmer, True
        Dcombos.GetBranches dcBranch(mIndex)

        'Dcombos.GetEmployees DcboEmp, , , , True
        
        
        
        TxtModFlg2(mIndex).text = "R"
        DCboUserName(mIndex).BoundText = user_id
       

        Me.Caption = "√Ê«„— «·‘€·"
   

        btn_First_Click (mIndex)
         Me.Width = TabMain2.Width + 400
  ElseIf mIndex = 4 Then
        My_SQL = "TblJobOrdersTasks"
       ' Set BKGrndPic = New ClsBackGroundPic
        Set RsSavRec = New ADODB.Recordset
        RsSavRec.CursorLocation = adUseClient
        RsSavRec.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdTableDirect
        
        
           TabMain.TabVisible(mIndex) = True
        TabMain.CurrTab = mIndex
      
     
              Dcombos.GetUsers Me.DCboUserName(mIndex)
           Dcombos.GetBranches dcBranch(mIndex)
        TxtModFlg2(mIndex).text = "R"
        DCboUserName(mIndex).BoundText = user_id
       


        Me.Caption = " ”ÃÌ· «·«‰ «ÃÌ… «·ÌÊ„Ì… ··„ÊŸ›« "

        btn_First_Click (mIndex)
         Me.Width = FG4.Width + 400
         
    ElseIf mIndex = 2 Then

        My_SQL = "dean"
        Set BKGrndPic = New ClsBackGroundPic
        Set RsSavRec = New ADODB.Recordset
        RsSavRec.CursorLocation = adUseClient
        RsSavRec.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdTableDirect
        Me.TxtModFlg.text = "R"
        Resize_Form Me
        'load tblUsers -----------------------------------------------
        My_SQL = "select UserID,UserName From tblUsers "
        fill_combo DCUser, My_SQL
    
        FillGridWithData
        Me.Width = Grid.Width + 200
        With Me.Grid
         '   .Cell(flexcpPicture, 0, .ColIndex("name")) = Me.GrdImageList.ListImage("Vac_Name").ExtractIcon
         '   .Cell(flexcpPicture, 0, .ColIndex("Ser")) = Me.GrdImageList.ListImages("Ser").ExtractIcon
    
            For i = 0 To .Cols - 1
                .cell(flexcpPictureAlignment, 0, i) = flexPicAlignRightCenter
            Next
       
            .ExtendLastCol = True
            .WallPaper = BKGrndPic.Picture
            .RowHeight(-1) = 300
        End With
            Me.Caption = "«·œÌ«‰« "
        BtnFirst_Click
        ShowTip
        Frm2.Enabled = True
       TabMain.TabVisible(mIndex) = True
        TabMain.CurrTab = mIndex
        Me.TxtModFlg.text = "R"
        If OPEN_NEW_SCREEN = True Then
            btnNew_Click
   
        End If
     ElseIf mIndex = 12 Then
        ScreenNameArabic = " Ê“Ì⁄ «·Œ’Ê„«  ⁄·Ï «·⁄ﬁ«—« "
        ScreenNameEnglish = "Tasks"
       
         TabMain.TabVisible(mIndex) = True
        TabMain.CurrTab = mIndex
        Me.Caption = ScreenNameArabic
        My_SQL = "TblIqarDiscountTrans"
        Set BKGrndPic = New ClsBackGroundPic
        Set RsSavRec = New ADODB.Recordset
        RsSavRec.CursorLocation = adUseClient
        RsSavRec.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdTableDirect
      TxtModFlg2(mIndex).text = "R"
        btn_First_Click (mIndex)
       ' XPDtbBill.value = Date
        Me.Width = GrdIqar.Width + 400
        Dcombos.GetBranches dcBranch(mIndex)
        Dcombos.GetUsers Me.DCboUserName(mIndex)
        ListBranchSelected.Clear
        ListAqarSelected.Clear
        ListUnitTypeSelected.Clear
        ListUnitNoSelected.Clear
        ListUnitNoSelected2.Clear
        ListBranchAll.Clear
        ListAqarAll.Clear
        ListUnitTypeAll.Clear
        ListUnitTypeSelected.Clear
        ListUnitNoAll2.Clear
        ListUnitNoAll.Clear
        ListUnitNoSelected.Clear
        FillMylist2 True, False, True, False
   End If
    'Me.Caption = ScreenNameArabic
    Me.WindowState = 2
    If SystemOptions.UserInterface = EnglishInterface Then
        Me.Caption = ScreenNameEnglish
        SetInterface Me
        ChangeLang
    End If


ErrTrap:




End Sub
Sub ReloadCompo()
Dim sql As String
'sql = "SELECT DISTINCT LocationsName, LocationsName AS LocationsName"
'sql = sql & " From dbo.TblTripReg"
'sql = sql & " WHERE     (NOT (LocationsName IS NULL)) "
'fill_combo cmbLocationsName, sql
'

'sql = "SELECT DISTINCT LocationsName, LocationsName AS LocationsName"
'sql = sql & " From dbo.TblEmpData"
'sql = sql & " WHERE     (NOT (LocationsName IS NULL)) "
'fill_combo cmbLocationsName2, sql

'
'sql = "SELECT DISTINCT CarName, CarName AS CarName"
'sql = sql & " From dbo.TblTripReg"
'sql = sql & " WHERE     (NOT (CarName IS NULL)) "
'fill_combo cmbCarName, sql

'sql = "SELECT DISTINCT CustName, CustName AS CustName"
'sql = sql & " From dbo.TblTripReg"
'sql = sql & " WHERE     (NOT (CustName IS NULL)) "
'fill_combo cmbCustName, sql

sql = "SELECT DISTINCT Id, Name ,Namee ,ServiceColor from tblPaymentClass where IsBoardNoHide is null or IsBoardNoHide=0"
Set rsDummy = New ADODB.Recordset
rsDummy.Open sql, Cn, adOpenStatic
Dim i As Long
i = 0
grd.rows = 1
grd.Cols = 0
Do While Not rsDummy.EOF
    grd.Cols = i + 1
    grd.ColWidth(i) = 1665
    grd.FontBold = True
    grd.fontsize = 16
    grd.ColAlignment(i) = flexAlignCenterCenter
    grd.TextMatrix(0, i) = rsDummy!Name & ""
    grd.ColKey(i) = rsDummy!ID & ""
    grd.Col = i
    grd.CellBackColor = val(rsDummy!ServiceColor & "")
    grd.ColEditMask(i) = val(rsDummy!ServiceColor & "")
    i = i + 1
    rsDummy.MoveNext
Loop
'grd.RowHidden(0) = True
grd.Cols = grd.Cols + 1




End Sub
Public Sub FillGridWithData1()

    On Error GoTo ErrTrap

    Dim i As Integer
    Dim rs As ADODB.Recordset
    Dim My_SQL As String

    Set rs = New ADODB.Recordset
    My_SQL = "select * From TblTasks order by id"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText

    With Me.GRID1
        .rows = 2
        .Clear flexClearScrollable

        If rs.RecordCount > 0 Then
            .rows = rs.RecordCount + 1
            rs.MoveFirst

            For i = 1 To .rows - 1
                .TextMatrix(i, .ColIndex("Ser")) = i
             
                .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(rs.Fields("name").value), "", rs.Fields("name").value)
                .TextMatrix(i, .ColIndex("namee")) = IIf(IsNull(rs.Fields("namee").value), "", rs.Fields("namee").value)
                .TextMatrix(i, .ColIndex("Percent")) = IIf(IsNull(rs.Fields("PercentV").value), "", rs.Fields("PercentV").value)
                
               
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs.Fields("id").value), "", rs.Fields("id").value)
                
            
                '    .Cell(flexcpBackColor, .Rows - 1, 1, .Rows - 1, .Cols - 1) = vbYellow
                rs.MoveNext
            Next

            rs.Close
        End If

        .RowHeight(-1) = 300
    End With

ErrTrap:
End Sub

Public Sub FillGridWithData5()

    On Error GoTo ErrTrap

    Dim i As Integer
    Dim rs As ADODB.Recordset
    Dim My_SQL As String

    Set rs = New ADODB.Recordset
    My_SQL = "select * From tblReservationType order by id"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText

    With Me.Grid5
        .rows = 2
        .Clear flexClearScrollable

        If rs.RecordCount > 0 Then
            .rows = rs.RecordCount + 1
            rs.MoveFirst

            For i = 1 To .rows - 1
                .TextMatrix(i, .ColIndex("Ser")) = i
             
                .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(rs.Fields("name").value), "", rs.Fields("name").value)
                .TextMatrix(i, .ColIndex("namee")) = IIf(IsNull(rs.Fields("namee").value), "", rs.Fields("namee").value)
                
               
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs.Fields("id").value), "", rs.Fields("id").value)
                
            
                '    .Cell(flexcpBackColor, .Rows - 1, 1, .Rows - 1, .Cols - 1) = vbYellow
                rs.MoveNext
            Next

            rs.Close
        End If

        .RowHeight(-1) = 300
    End With

ErrTrap:
End Sub


Public Sub FillGridWithData8()

    On Error GoTo ErrTrap

    Dim i As Integer
    Dim rs As ADODB.Recordset
    Dim My_SQL As String

    Set rs = New ADODB.Recordset
    My_SQL = "select * From tblPaymentClass where 1=1"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText

    With Me.Grid8
        .rows = 2
        .Clear flexClearScrollable

        If rs.RecordCount > 0 Then
            .rows = rs.RecordCount + 1
            rs.MoveFirst

            For i = 1 To .rows - 1
                .TextMatrix(i, .ColIndex("Ser")) = i
             
                .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(rs.Fields("name").value), "", rs.Fields("name").value)
                .TextMatrix(i, .ColIndex("namee")) = IIf(IsNull(rs.Fields("namee").value), "", rs.Fields("namee").value)
                
               
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs.Fields("id").value), "", rs.Fields("id").value)
                
            
                '    .Cell(flexcpBackColor, .Rows - 1, 1, .Rows - 1, .Cols - 1) = vbYellow
                rs.MoveNext
            Next

            rs.Close
        End If

        .RowHeight(-1) = 300
    End With

ErrTrap:
End Sub

Public Sub FiLLRec1()
    On Error GoTo ErrTrap

    RsSavRec.Fields("name").value = IIf(TxtName(mIndex).text <> "", Trim(TxtName(mIndex).text), Null)
    RsSavRec.Fields("namee").value = IIf(TxtNameE(mIndex).text <> "", Trim(TxtNameE(mIndex).text), Null)
    
    RsSavRec.Fields("PercentV").value = IIf(txtPercentV.text <> "", Trim(txtPercentV.text), Null)
    
    
    RsSavRec.update
    MsgBox " „  ⁄„·Ì… «·Õ›Ÿ »‰Ã«Õ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    
    FillGridWithData1
    TxtModFlg2(mIndex) = "R"

    Exit Sub
ErrTrap:

    If RsSavRec.EditMode <> adEditNone Then
        RsSavRec.CancelUpdate
    End If

End Sub
Public Sub FiLLRec9()
On Error GoTo ErrTrap
Dim StoreId1 As Integer

Dim j As Long


If Trim(DBCboClientName.text) = "" Then
    MsgBox "·«Ì„ﬂ‰ «·Õ›Ÿ , »—Ã«¡ «œŒ«· «·⁄„Ì·", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    Exit Sub

End If
Dim mNum As Long
mNum = val(txtPhoneCust)
If Len(CStr(mNum)) <> 9 Then
    MsgBox "·«Ì„ﬂ‰ «·Õ›Ÿ , »—Ã«¡ «œŒ«· —ﬁ„  ·Ì›Ê‰ ’ÕÌÕ ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    Exit Sub

End If
If optCash Then
    If val(txtTotalWithVat2) <> val(txtAmountVisa) + val(txtAmountCash) Then
        MsgBox "·«Ì„ﬂ‰ «·Õ›Ÿ , »—Ã«¡ «œŒ«· «·ﬁÌ„… «·’ÕÌÕÌ… ··›« Ê—… ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
        
    End If
End If

'If Trim(txtCarName.Text) = "" Then
'    MsgBox "·«Ì„ﬂ‰ «·Õ›Ÿ , »—Ã«¡ «œŒ«· «·„⁄œÂ/«·”Ì«—…", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
'    Exit Sub
'
'End If




If Trim(dcBranch(mIndex).text) = "" Then
    MsgBox "·«Ì„ﬂ‰ «·Õ›Ÿ , »—Ã«¡ «œŒ«· «·„Êﬁ⁄", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    Exit Sub

End If

If cmbPaymentClass.text <> "" Then
    s = "Select IsBoardNo from tblPaymentClass where Id = " & val(cmbPaymentClass.BoundText)
    Set rsDummy = New ADODB.Recordset
    rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly
    If Not rsDummy.EOF Then
        If Not IsNull(rsDummy!IsBoardNo) Then
            If rsDummy!IsBoardNo Then
                If Trim(txtBoardNo) = "" Or Trim(txtnBoardNo) = "" Then
                    MsgBox "·«Ì„ﬂ‰ «·Õ›Ÿ , »—Ã«¡ «œŒ«· —ﬁ„ «··ÊÕ…", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                    Exit Sub
                End If
            End If
        End If
    End If
End If

If Trim(CboPayMentType.text) = "" Then
CboPayMentType.ListIndex = 0
    'MsgBox "·«Ì„ﬂ‰ «·Õ›Ÿ , »—Ã«¡ «œŒ«· ÿ—Ìﬁ… «·”œ«œ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    'Exit Sub
'
End If
If Trim(txtPhoneCust.text) = "" Then
    MsgBox "·«Ì„ﬂ‰ «·Õ›Ÿ , »—Ã«¡ «œŒ«· «· ·Ì›Ê‰", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    Exit Sub

End If


If Trim(txtTotalWithVat2.text) = "" Then
    MsgBox "·«Ì„ﬂ‰ «·Õ›Ÿ , »—Ã«¡ «œŒ«· «·ﬁÌ„…", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    Exit Sub

End If

    
    If TxtModFlg2(mIndex).text = "N" Then
           ' TxtSerial1(mIndex).Text = CStr(new_id("TblAdditionsAssest", "ID", "", True))
       ' RsSavRec.AddNew
         If TxtNoteSerial1.text = "" Then
                        If Voucher_coding(val(dcBranch(mIndex).BoundText), XPDtbTrans(mIndex).value, 82, 1101, , , , , , , "TblHandWages") = "error" Then
                            If SystemOptions.UserInterface = ArabicInterface Then
                                MsgBox " ·« Ì„ﬂ‰ «÷«›… ”‰œ  ÃœÌœ ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": Exit Sub
                            Else
                                MsgBox " Cant't Create Expenses Voucher to this Process no You exceed the maximum number ": Exit Sub
                            End If
        
                        Else
                 
                            If Voucher_coding(val(dcBranch(mIndex).BoundText), XPDtbTrans(mIndex).value, 82, 1101, , , , , , , "TblTripReg") = "" Then
                                If SystemOptions.UserInterface = ArabicInterface Then
                                    
                                    TxtNoteSerial1.locked = False
                                    MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ «·”‰œ ÌœÊÌ« ﬂ„« Õœœ   ": Exit Sub
                                Else
                                    MsgBox "  Enter Voucher No Manually or Define Coding ": Exit Sub
                                End If
        
                            Else
                                TxtNoteSerial1.text = Voucher_coding(val(dcBranch(mIndex).BoundText), XPDtbTrans(mIndex).value, 82, 1101, , , , , , , "TblTripReg")
                            End If
                        End If
                    End If
       
       AddNewRec
        TxtSerial1(mIndex).text = new_id("TblTripReg", "id", "")
        RsSavRec.Fields("ID").value = val(TxtSerial1(mIndex).text)
    End If
    RsSavRec("NoteSerial1").value = Trim$(Me.TxtNoteSerial1.text)
    RsSavRec.Fields("BranchID").value = IIf(dcBranch(mIndex).text <> "", Trim(dcBranch(mIndex).BoundText), Null)
    
    RsSavRec("RecordDate").value = XPDtbTrans(mIndex).value
    RsSavRec("DateRec").value = txtDateRec.value
    RsSavRec("StartTime").value = startTime.value

    


    RsSavRec("UserID").value = user_id
    
    RsSavRec.Fields("Value").value = val(XPTxtVal.text)
    RsSavRec.Fields("VAt22").value = val(TxtVAt22.text)
    RsSavRec.Fields("TotalWithVat2").value = val(txtTotalWithVat2.text)
    RsSavRec.Fields("PayMentType").value = val(CboPayMentType.ListIndex)
    
    RsSavRec.Fields("PayType").value = IIf(optCash, 0, 1)
    RsSavRec!CusID = val(DBCboClientName.BoundText)
    RsSavRec!AmountCash = val(txtAmountCash)
    RsSavRec!AmountVisa = val(txtAmountVisa)
    RsSavRec!AmountLater = val(txtAmountLater)

     RsSavRec.Fields("Remarks").value = Trim(TxtSearchCode2.text)
    'RsSavRec.Fields("LocationsName").value = Trim(cmbLocationsName.Text)
    RsSavRec.Fields("CarName").value = Trim(txtnBoardNo.text)
    
   RsSavRec.Fields("nBoardNo").value = Trim(txtnBoardNo.text)
   RsSavRec.Fields("BoardNo").value = Trim(txtBoardNo.text)
   RsSavRec.Fields("txtLetter1").value = Trim(txtLetter1.text)
   RsSavRec.Fields("txtLetter2").value = Trim(txtLetter2.text)
   RsSavRec.Fields("txtLetter3").value = Trim(txtLetter3.text)
   RsSavRec.Fields("txtLetter4").value = Trim(txtLetter4.text)
   RsSavRec.Fields("ntxtLetter1").value = Trim(ntxtLetter1.text)
   RsSavRec.Fields("ntxtLetter2").value = Trim(ntxtLetter2.text)
   RsSavRec.Fields("ntxtLetter3").value = Trim(ntxtLetter3.text)
   RsSavRec.Fields("ntxtLetter4").value = Trim(ntxtLetter4.text)
      
    RsSavRec!txtNum1 = IIf(txtNum1.text = "", Null, Trim(txtNum1.text))
    RsSavRec!txtNum2 = IIf(txtNum2.text = "", Null, Trim(txtNum2.text))
    RsSavRec!txtNum3 = IIf(txtNum3.text = "", Null, Trim(txtNum3.text))
    RsSavRec!txtNum4 = IIf(txtNum4.text = "", Null, Trim(txtNum4.text))
    
    RsSavRec!ntxtNum1 = IIf(ntxtNum1.text = "", Null, Trim(ntxtNum1.text))
    RsSavRec!ntxtNum2 = IIf(ntxtNum2.text = "", Null, Trim(ntxtNum2.text))
    RsSavRec!ntxtNum3 = IIf(ntxtNum3.text = "", Null, Trim(ntxtNum3.text))
    RsSavRec!ntxtNum4 = IIf(ntxtNum4.text = "", Null, Trim(ntxtNum4.text))
    
 
    RsSavRec.Fields("CustName").value = Trim(txtCustName.text)
    RsSavRec.Fields("PhoneCust").value = Trim(txtPhoneCust.text)
    RsSavRec!PaymentClassID = val(cmbPaymentClass.BoundText)
    
   'RsSavRec("RecType").value = cmbRecType.ListIndex
    'RsSavRec("ContractNo").value = txtContractNo.Text
    'RsSavRec("RecName").value = txtRecName.Text
    'RsSavRec("RecordTime").value = XPDtbTransTime.Value
    

    
    
    
    
    
    '*********************
     
    
    
       
   

    RsSavRec.update
  
    MsgBox " „  ⁄„·Ì… «·Õ›Ÿ »‰Ã«Õ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    
     If TxtModFlg2(mIndex) = "N" And txtPhoneCust.text <> "123456789" Then
    CMDPAy_Click
    End If
    
    'FillGridWithData1
    TxtModFlg2(mIndex) = "R"
ReloadCompo
    Dim My_SQL As String
     My_SQL = "TblTripReg"
    ' Set BKGrndPic = New ClsBackGroundPic
     Set RsSavRec = New ADODB.Recordset
     RsSavRec.CursorLocation = adUseClient
     RsSavRec.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdTableDirect
     RsSavRec.Find "Id = " & val(TxtSerial1(mIndex))

    FiLLTXT9
    btn_New_Click (9)
    Exit Sub
ErrTrap:

    If RsSavRec.EditMode <> adEditNone Then
        RsSavRec.CancelUpdate
    End If

End Sub

Public Sub FiLLRec3()
    On Error GoTo ErrTrap
Dim StoreId1 As Integer

Dim j As Long
Dim mFound As Boolean
For j = 1 To FG.rows - 1
    If Trim(FG.TextMatrix(j, FG.ColIndex("TasksID"))) <> "" Then
        mFound = True
    End If
    
Next
If Not mFound Then
    MsgBox "·«Ì„ﬂ‰ «·Õ›Ÿ , »—Ã«¡ «œŒ«· „Â«„ ›Ï «·ÃœÊ·", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    Exit Sub
End If

'If Trim(DcboItemID1.Text) = "" Then
'    MsgBox "·«Ì„ﬂ‰ «·Õ›Ÿ , »—Ã«¡ «œŒ«· «·ﬁÿ⁄…", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
'    Exit Sub
'
'End If




If Trim(DcCustmer.text) = "" Then
    MsgBox "·«Ì„ﬂ‰ «·Õ›Ÿ , »—Ã«¡ «œŒ«· «·⁄„Ì·…", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    Exit Sub

End If

Dim s As String
Dim StrTempAccountCode As String
Dim rsDummy As New ADODB.Recordset
s = "Select StoreID,StoreID,StoreID2,StoreID3 from tblUsers Where UserID = " & user_id
Set rsDummy = New ADODB.Recordset

rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly, adCmdText
If Not rsDummy.EOF Then
    StoreId1 = val(rsDummy!StoreID & "")
End If
If val(StoreId1) = 0 Then
  MsgBox "·«Ì„ﬂ‰ «·Õ›Ÿ , »—Ã«¡  ÕœÌœ ›—⁄ Ê„Œ“‰ «·„” Œœ„", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
  Exit Sub
End If
    
Dim rsOut As New ADODB.Recordset
Dim Current_case As Integer, mBoxID As Long
Set rsOut = New ADODB.Recordset

s = "Select BoxID From TblBoxesData Where Empid In (Select tblUsers.EmpId from tblUsers where UserId = " & user_id & " )"
rsOut.Open s, Cn, adOpenStatic, adLockReadOnly
If Not rsOut.EOF Then
    mBoxID = val(rsOut!BoxID & "")
End If
If val(DcCustmer.BoundText) = 2 And val(txtPaymedValue) <> 0 Then
      MsgBox "·«Ì„ﬂ‰ «·Õ›Ÿ , ·⁄œ„ «„ﬂ«‰Ì… «‰‘«¡ ”‰œ ﬁ»÷ ··⁄„Ì· «·‰ﬁœÏ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
  Exit Sub
End If
    If TxtModFlg2(mIndex).text = "N" Then
           ' TxtSerial1(mIndex).Text = CStr(new_id("TblAdditionsAssest", "ID", "", True))
       ' RsSavRec.AddNew
        TxtSerial1(mIndex).text = new_id("TblJobOrders", "id", "")
        RsSavRec.Fields("ID").value = val(TxtSerial1(mIndex).text)
    End If
    
    RsSavRec.Fields("BranchID").value = IIf(dcBranch(mIndex).text <> "", Trim(dcBranch(mIndex).BoundText), Null)
    RsSavRec("RecordDate").value = XPDtbTrans(mIndex).value
    RsSavRec("DateRec").value = txtDateRec.value
    RsSavRec("DateRehearsal").value = txtDateRehearsal.value
    RsSavRec("RehearsalDateFinish").value = txtRehearsalDateFInish.value
    RsSavRec("DateDelivery").value = txtDateDelivery.value
    RsSavRec("DeliveryDateFinish").value = txtDeliveryDateFinish.value
    RsSavRec("DateDeliveryAct").value = txtDateDeliveryAct.value
    
    'RsSavRec("EmpId").value = val(DcboEmp.BoundText)
    RsSavRec("CusId").value = val(DcCustmer.BoundText)
    'RsSavRec("ItemID").value = val(DcboItemID1.BoundText)
    RsSavRec("UserID").value = user_id
    
    RsSavRec.Fields("GeneralTotal").value = val(txtGeneralTotal.text)
    RsSavRec.Fields("TotalAdd").value = val(txtTotalAdd.text)
    RsSavRec.Fields("TotalPay").value = val(txtTotalPay.text)
    
    RsSavRec.Fields("VatYou").value = val(txtVatYou.text)
    RsSavRec.Fields("Vat").value = val(TxtVAT.text)
    RsSavRec.Fields("TotalAfterVat").value = val(txtTotalAfterVat.text)
    
    RsSavRec.Fields("TotalDiscPerc").value = val(txtTotalDiscPerc.text)
    RsSavRec.Fields("TotalDisc").value = val(txtTotalDisc.text)
    RsSavRec.Fields("RequiredAmount").value = val(txtRequiredAmount.text)
    RsSavRec.Fields("PaymedValue").value = val(txtPaymedValue.text)
    RsSavRec.Fields("TotalNet").value = val(txtTotalNet.text)
    
    
   'RsSavRec("RecType").value = cmbRecType.ListIndex
    'RsSavRec("ContractNo").value = txtContractNo.Text
    'RsSavRec("RecName").value = txtRecName.Text
    'RsSavRec("RecordTime").value = XPDtbTransTime.Value
    

    
    
    'RsSavRec("Remarks").value = TxtRemarks.Text
    
    
    '*********************
     
    
    
      
   

    RsSavRec.update
    
    
                
   
    s = " Delete From TblJobOrders2 Where MasterID = " & val(TxtSerial1(mIndex).text)
    Cn.Execute s
    
    s = "Select * from TblJobOrders2 Where Id = -1"
    saveGrid s, FG, "TasksID", "SerID", "MasterID", val(TxtSerial1(mIndex).text)

    
    
   
    s = " Delete From Notes Where NoteID = " & val(txtNoteid3.text)
    Cn.Execute s
    
    

    RsSavRec.update
    CreateSales
    MsgBox " „  ⁄„·Ì… «·Õ›Ÿ »‰Ã«Õ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    
    'FillGridWithData1
    TxtModFlg2(mIndex) = "R"

    Dim My_SQL As String
     My_SQL = "TblJobOrders"
    ' Set BKGrndPic = New ClsBackGroundPic
     Set RsSavRec = New ADODB.Recordset
     RsSavRec.CursorLocation = adUseClient
     RsSavRec.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdTableDirect
     RsSavRec.Find "Id = " & val(TxtSerial1(mIndex))

FiLLTXT3
    Exit Sub
ErrTrap:

    If RsSavRec.EditMode <> adEditNone Then
        RsSavRec.CancelUpdate
    End If

End Sub

Public Sub FiLLRec4()
    On Error GoTo ErrTrap

    
    If TxtModFlg2(mIndex).text = "N" Then
           ' TxtSerial1(mIndex).Text = CStr(new_id("TblAdditionsAssest", "ID", "", True))
       ' RsSavRec.AddNew
        TxtSerial1(mIndex).text = new_id("TblJobOrdersTasks", "id", "")
        RsSavRec.Fields("ID").value = val(TxtSerial1(mIndex).text)
    End If
    
    RsSavRec.Fields("BranchID").value = IIf(dcBranch(mIndex).text <> "", Trim(dcBranch(mIndex).BoundText), Null)
    RsSavRec("RecordDate").value = XPDtbTrans(mIndex).value
'    RsSavRec("Remarks").value = TxtRemarks.Text
    RsSavRec("UserID").value = user_id
    
    '*********************
     
    
    
      
   

    RsSavRec.update
    
    Dim s As String
                
   
    s = " Delete From TblJobOrdersTasks2 Where MasterID = " & val(TxtSerial1(mIndex).text)
    
        
        
    
    Cn.Execute s
    
    s = "Select * from TblJobOrdersTasks2 Where Id = -1"
    saveGrid s, FG4, "JobOrdersNo", "SerID", "MasterID", val(TxtSerial1(mIndex).text)

    

    RsSavRec.update

    MsgBox " „  ⁄„·Ì… «·Õ›Ÿ »‰Ã«Õ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    
    'FillGridWithData1
    TxtModFlg2(mIndex) = "R"

    Exit Sub
ErrTrap:

    If RsSavRec.EditMode <> adEditNone Then
        RsSavRec.CancelUpdate
    End If

End Sub




Public Sub FiLLRec6()
    On Error GoTo ErrTrap

    
    If TxtModFlg2(mIndex).text = "N" Then
           ' TxtSerial1(mIndex).Text = CStr(new_id("TblAdditionsAssest", "ID", "", True))
       ' RsSavRec.AddNew
        TxtSerial1(mIndex).text = new_id("TblAppointmentlist", "id", "")
        RsSavRec.Fields("ID").value = val(TxtSerial1(mIndex).text)
    End If
    
    RsSavRec.Fields("BranchID").value = IIf(dcBranch(mIndex).text <> "", Trim(dcBranch(mIndex).BoundText), Null)
    RsSavRec("RecordDate").value = XPDtbTrans(mIndex).value
'    RsSavRec("Remarks").value = TxtRemarks.Text
    RsSavRec("UserID").value = user_id
    
    '*********************
     
    
    
      
   

    RsSavRec.update
    
    Dim s As String
                
   
    s = " Delete From TblAppointmentlist2 Where MasterID = " & val(TxtSerial1(mIndex).text)
    
        
        
    
    Cn.Execute s
    
    s = "Select * from TblAppointmentlist2 Where Id = -1"
    saveGrid s, Fg6, "ReservNo", "SerID", "MasterID", val(TxtSerial1(mIndex).text)

    

    RsSavRec.update

    MsgBox " „  ⁄„·Ì… «·Õ›Ÿ »‰Ã«Õ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    
    'FillGridWithData1
    TxtModFlg2(mIndex) = "R"

    Exit Sub
ErrTrap:

    If RsSavRec.EditMode <> adEditNone Then
        RsSavRec.CancelUpdate
    End If

End Sub

Public Sub FiLLRec7()
    On Error GoTo ErrTrap

    
    If TxtModFlg2(mIndex).text = "N" Then
           ' TxtSerial1(mIndex).Text = CStr(new_id("TblAdditionsAssest", "ID", "", True))
       ' RsSavRec.AddNew
        TxtSerial1(mIndex).text = new_id("TblEmpItemsTrans", "id", "")
        RsSavRec.Fields("ID").value = val(TxtSerial1(mIndex).text)
    End If
    
    RsSavRec.Fields("BranchID").value = IIf(dcBranch(mIndex).text <> "", Trim(dcBranch(mIndex).BoundText), Null)
    RsSavRec("RecordDate").value = XPDtbTrans(mIndex).value
'    RsSavRec("Remarks").value = TxtRemarks.Text
    RsSavRec("UserID").value = user_id
    
    '*********************
     
    
    
      
   

    RsSavRec.update
    
    Dim s As String
                
   
    s = " Delete From TblEmpItemsTrans2 Where MasterID = " & val(TxtSerial1(mIndex).text)
    
        
        

    Cn.Execute s
    
    s = "Select * from TblEmpItemsTrans2 Where Id = -1"
    saveGrid s, fg7, "ItemID", "SerID", "MasterID", val(TxtSerial1(mIndex).text)

    

    RsSavRec.update

    MsgBox " „  ⁄„·Ì… «·Õ›Ÿ »‰Ã«Õ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    
    'FillGridWithData1
    TxtModFlg2(mIndex) = "R"

    Exit Sub
ErrTrap:

    If RsSavRec.EditMode <> adEditNone Then
        RsSavRec.CancelUpdate
    End If

End Sub


Public Sub FiLLRec12()
    On Error GoTo ErrTrap

    
    If TxtModFlg2(mIndex).text = "N" Then
           ' TxtSerial1(mIndex).Text = CStr(new_id("TblAdditionsAssest", "ID", "", True))
       ' RsSavRec.AddNew
        TxtSerial1(mIndex).text = new_id("TblIqarDiscountTrans", "id", "")
        RsSavRec.Fields("ID").value = val(TxtSerial1(mIndex).text)
    End If
    
    RsSavRec.Fields("BranchID").value = IIf(dcBranch(mIndex).text <> "", Trim(dcBranch(mIndex).BoundText), Null)
    RsSavRec("RecordDate").value = XPDtbTrans(mIndex).value
'    RsSavRec("Remarks").value = TxtRemarks.Text
    RsSavRec("UserID").value = user_id
    RsSavRec("DiscountPercent").value = val(txtDiscountPercent)
    
    '*********************
     
    
    
      
   

    RsSavRec.update
    
    Dim s As String
                
   
    s = " Delete From TblIqarDiscountTrans2 Where MasterID = " & val(TxtSerial1(mIndex).text)
    
        
        

    Cn.Execute s
    
    s = "Select * from TblIqarDiscountTrans2 Where Id = -1"
    saveGrid s, GrdIqar, "BranchID", "SerID", "MasterID", val(TxtSerial1(mIndex).text)

    

    RsSavRec.update

    MsgBox " „  ⁄„·Ì… «·Õ›Ÿ »‰Ã«Õ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    
    'FillGridWithData1
    TxtModFlg2(mIndex) = "R"

    Exit Sub
ErrTrap:

    If RsSavRec.EditMode <> adEditNone Then
        RsSavRec.CancelUpdate
    End If

End Sub




Public Sub FiLLRec11()
    On Error GoTo ErrTrap

    
    If TxtModFlg2(mIndex).text = "N" Then
           ' TxtSerial1(mIndex).Text = CStr(new_id("TblAdditionsAssest", "ID", "", True))
       ' RsSavRec.AddNew
        TxtSerial1(mIndex).text = new_id("TblEmpItemsTrans", "id", "")
        RsSavRec.Fields("ID").value = val(TxtSerial1(mIndex).text)
    End If
    
    RsSavRec.Fields("BranchID").value = IIf(dcBranch(mIndex).text <> "", Trim(dcBranch(mIndex).BoundText), Null)
    RsSavRec("RecordDate").value = XPDtbTrans(mIndex).value
'    RsSavRec("Remarks").value = TxtRemarks.Text
    RsSavRec("UserID").value = user_id
    
    '*********************
     
    
    
      
   

    RsSavRec.update
    
    Dim s As String
                
   
    s = " Delete From TblEmpItemsTrans2 Where MasterID = " & val(TxtSerial1(mIndex).text)
    
        
        

    Cn.Execute s
    
    s = "Select * from TblEmpItemsTrans2 Where Id = -1"
    saveGrid s, fg7, "ItemID", "SerID", "MasterID", val(TxtSerial1(mIndex).text)

    

    RsSavRec.update

    MsgBox " „  ⁄„·Ì… «·Õ›Ÿ »‰Ã«Õ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    
    'FillGridWithData1
    TxtModFlg2(mIndex) = "R"

    Exit Sub
ErrTrap:

    If RsSavRec.EditMode <> adEditNone Then
        RsSavRec.CancelUpdate
    End If

End Sub


Public Sub FiLLRec10()
' On Error GoTo ErrTrap
Dim StoreId1 As Integer

Dim j As Long


If Trim(txtEmpName.text) = "" Then
    MsgBox "·«Ì„ﬂ‰ «·Õ›Ÿ , »—Ã«¡ «œŒ«· «·„ÊŸ›", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    Exit Sub

End If
      
s = "Select * from TblEmpData Where HafizaNo = N'" & Trim(txtHafizaNo) & "' and Id <> " & val(TxtSerial1(mIndex))
Set rsDummy = New ADODB.Recordset
rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly
If Not rsDummy.EOF Then
    ZKFPEngX1.EndEngine
    Label13.Caption = "€Ì— „ ’·"
    MsgBox "·«Ì„ﬂ‰ «·Õ›Ÿ , »—Ã«¡ «œŒ«· —ﬁ„ ÂÊÌ… ¬Œ— ·«‰ Â–« «·—ﬁ„ „ﬂ—— „⁄ «·„ÊŸ› " & rsDummy!EmpName & "", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    Exit Sub

End If
If Trim(txtHafizaNo.text) = "" Then
    MsgBox "·«Ì„ﬂ‰ «·Õ›Ÿ , »—Ã«¡ «œŒ«· —ﬁ„ ÂÊÌ… «·„ÊŸ›", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    Exit Sub

End If

If Len(Trim(txtHafizaNo)) <> 10 Then
    MsgBox "·«Ì„ﬂ‰ «·Õ›Ÿ , »—Ã«¡ «œŒ«· —ﬁ„ «·ÂÊÌ… «·’ÕÌÕ ··„ÊŸ›", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    Exit Sub

End If


If Trim(TxtMobileNO.text) = "" Then
    MsgBox "·«Ì„ﬂ‰ «·Õ›Ÿ , »—Ã«¡ «œŒ«· —ﬁ„ «·ÃÊ«· ··„ÊŸ›", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    Exit Sub

End If

If Len(Trim(TxtMobileNO)) <> 9 Then
    MsgBox "·«Ì„ﬂ‰ «·Õ›Ÿ , »—Ã«¡ «œŒ«· —ﬁ„ «·ÃÊ«· «·’ÕÌÕ ··„ÊŸ› 9 Œ«‰« ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    Exit Sub

End If

'If Trim(txtCarName.Text) = "" Then
'    MsgBox "·«Ì„ﬂ‰ «·Õ›Ÿ , »—Ã«¡ «œŒ«· «·„⁄œÂ/«·”Ì«—…", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
'    Exit Sub
'
'End If




If Trim(dcBranch(mIndex).text) = "" Then
    MsgBox "·«Ì„ﬂ‰ «·Õ›Ÿ , »—Ã«¡ «œŒ«· «·„Êﬁ⁄", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    Exit Sub

End If


'If Trim(cmbLocationsName2.Text) = "" Then
'    MsgBox "·«Ì„ﬂ‰ «·Õ›Ÿ , »—Ã«¡ «œŒ«· „Êﬁ⁄ «·⁄„·", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
'    Exit Sub
'
'End If


    
    If TxtModFlg2(mIndex).text = "N" Then
           ' TxtSerial1(mIndex).Text = CStr(new_id("TblAdditionsAssest", "ID", "", True))
       ' RsSavRec.AddNew
'         If TxtNoteSerial1.Text = "" Then
'                        If Voucher_coding(val(dcBranch(mIndex).BoundText), XPDtbTrans(mIndex).value, 82, 1101, , , , , , , "TblHandWages") = "error" Then
'                            If SystemOptions.UserInterface = ArabicInterface Then
'                                MsgBox " ·« Ì„ﬂ‰ «÷«›… ”‰œ  ÃœÌœ ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": Exit Sub
'                            Else
'                                MsgBox " Cant't Create Expenses Voucher to this Process no You exceed the maximum number ": Exit Sub
'                            End If
'
'                        Else
'
'                            If Voucher_coding(val(dcBranch(mIndex).BoundText), XPDtbTrans(mIndex).value, 82, 1101, , , , , , , "TblTripReg") = "" Then
'                                If SystemOptions.UserInterface = ArabicInterface Then
'
'                                    TxtNoteSerial1.locked = False
'                                    MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ «·”‰œ ÌœÊÌ« ﬂ„« Õœœ   ": Exit Sub
'                                Else
'                                    MsgBox "  Enter Voucher No Manually or Define Coding ": Exit Sub
'                                End If
'
'                            Else
'                                TxtNoteSerial1.Text = Voucher_coding(val(dcBranch(mIndex).BoundText), XPDtbTrans(mIndex).value, 82, 1101, , , , , , , "TblTripReg")
'                            End If
'                        End If
'                    End If
       
       AddNewRec
        TxtSerial1(mIndex).text = new_id("TblEmpData", "id", "")
        RsSavRec.Fields("ID").value = val(TxtSerial1(mIndex).text)
    End If
    'RsSavRec("NoteSerial1").value = Trim$(Me.TxtNoteSerial1.Text)
    RsSavRec.Fields("BranchID").value = IIf(dcBranch(mIndex).text <> "", Trim(dcBranch(mIndex).BoundText), Null)
    
    RsSavRec("RecordDate").value = XPDtbTrans(mIndex).value
    RsSavRec("startDate").value = txtStartDate.value
    RsSavRec("TimeIn").value = TimeIn.value
    RsSavRec("TimeOut").value = TimeOut.value
    
    RsSavRec("FingerPrint").value = Trim(txtFingerPrint)
    RsSavRec("FingerStatus").value = val(lblFingerStatus.Tag)
    
    RsSavRec("Photo2").value = DBPix201.Image
    
    
    


    RsSavRec("UserID").value = user_id
    RsSavRec.Fields("MobileNO").value = Trim(TxtMobileNO.text)
    RsSavRec.Fields("HafizaNo").value = Trim(txtHafizaNo.text)
    RsSavRec.Fields("Salary").value = val(TxtSalary.text)
    
    
    RsSavRec.Fields("IsEmp").value = IIf(optIsEmp, 0, 1)
    
   
    'RsSavRec.Fields("Remarks").value = Trim(TxtRemarks.Text)
 '   RsSavRec.Fields("LocationsName").value = Trim(cmbLocationsName2.Text)
    RsSavRec.Fields("EmpName").value = Trim(txtEmpName.text)
    
    
    
                
   

    RsSavRec.update
  
    s = " Delete FROM  TblEmpDataFingerPrint Where EmpId = " & val(TxtSerial1(mIndex))
    Cn.Execute s
    s = "Select * from TblEmpDataFingerPrint Where Id = -1"
    
    
    saveGrid s, GrdFinger, "FingerPrint", "", "EmpID", val(Me.TxtSerial1(mIndex).text), "HafizaNo", Trim(txtHafizaNo)
    Dim Msg As String
If TxtModFlg2(mIndex).text = "N" Then

                Msg = "  „ Õ›Ÿ «·»Ì«‰«      " & CHR(13)
            Msg = Msg + "Â·  —€» ›Ì ≈÷«›… »Ì«‰«  √Œ—Ì"

            If MsgBox(Msg, vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading + vbDefaultButton2, App.Title) = vbYes Then
                btn_New_Click (mIndex)
                Exit Sub
            End If
    Else
    MsgBox " „  ⁄„·Ì… «·Õ›Ÿ »‰Ã«Õ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    End If
    
  
    
    'FillGridWithData1
    TxtModFlg2(mIndex) = "R"
ReloadCompo
    Dim My_SQL As String
     My_SQL = "TblEmpData"
    ' Set BKGrndPic = New ClsBackGroundPic
     Set RsSavRec = New ADODB.Recordset
     RsSavRec.CursorLocation = adUseClient
     RsSavRec.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdTableDirect
     RsSavRec.Find "Id = " & val(TxtSerial1(mIndex))

    FiLLTXT10
    Exit Sub
ErrTrap:

    If RsSavRec.EditMode <> adEditNone Then
        RsSavRec.CancelUpdate
    End If

End Sub





Public Sub FiLLRec2()
    On Error GoTo ErrTrap

    RsSavRec.Fields("name").value = IIf(TxtName(mIndex).text <> "", Trim(TxtName(mIndex).text), Null)
    RsSavRec.Fields("namee").value = IIf(TxtNameE(mIndex).text <> "", Trim(TxtNameE(mIndex).text), Null)
    

    RsSavRec.update
    MsgBox " „  ⁄„·Ì… «·Õ›Ÿ »‰Ã«Õ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    
    FillGridWithData2
    TxtModFlg2(mIndex) = "R"

    Exit Sub
ErrTrap:

    If RsSavRec.EditMode <> adEditNone Then
        RsSavRec.CancelUpdate
    End If

End Sub

Public Sub FiLLRec5()
    On Error GoTo ErrTrap

    RsSavRec.Fields("name").value = IIf(TxtName(mIndex).text <> "", Trim(TxtName(mIndex).text), Null)
    RsSavRec.Fields("namee").value = IIf(TxtNameE(mIndex).text <> "", Trim(TxtNameE(mIndex).text), Null)
    

    RsSavRec.update
    MsgBox " „  ⁄„·Ì… «·Õ›Ÿ »‰Ã«Õ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    
    FillGridWithData5
    TxtModFlg2(mIndex) = "R"

    Exit Sub
ErrTrap:

    If RsSavRec.EditMode <> adEditNone Then
        RsSavRec.CancelUpdate
    End If

End Sub

Public Sub FiLLRec8()
    On Error GoTo ErrTrap

    RsSavRec.Fields("name").value = IIf(TxtName(mIndex).text <> "", Trim(TxtName(mIndex).text), Null)
    RsSavRec.Fields("namee").value = IIf(TxtNameE(mIndex).text <> "", Trim(TxtNameE(mIndex).text), Null)
    If Me.chkIsBoardNo(0).value = vbChecked Then
        RsSavRec("IsBoardNo").value = 1
    ElseIf Me.chkIsBoardNo(0).value = vbUnchecked Then
        RsSavRec("IsBoardNo").value = 0
    End If
    
        If Me.chkIsBoardNo(1).value = vbChecked Then
        RsSavRec("IsBoardNoHide").value = 1
    ElseIf Me.chkIsBoardNo(1).value = vbUnchecked Then
        RsSavRec("IsBoardNoHide").value = 0
    End If
    
    
    RsSavRec("ServiceColor").value = val(lblServiceColor.backcolor)
    RsSavRec.update
    MsgBox " „  ⁄„·Ì… «·Õ›Ÿ »‰Ã«Õ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    
    FillGridWithData8
    TxtModFlg2(mIndex) = "R"

    Exit Sub
ErrTrap:

    If RsSavRec.EditMode <> adEditNone Then
        RsSavRec.CancelUpdate
    End If

End Sub

Public Sub FiLLTXT1()

    On Error GoTo ErrTrap
    Dim i As Integer
  '  Frame1(mIndex).Enabled = False
    TxtSerial1(mIndex).text = IIf(IsNull(RsSavRec.Fields("id").value), "", RsSavRec.Fields("id").value)
    TxtName(mIndex).text = IIf(IsNull(RsSavRec.Fields("name").value), "", RsSavRec.Fields("name").value)
    TxtNameE(mIndex).text = IIf(IsNull(RsSavRec.Fields("nameE").value), "", RsSavRec.Fields("nameE").value)
    
    txtPercentV.text = IIf(IsNull(RsSavRec.Fields("PercentV").value), "", RsSavRec.Fields("PercentV").value)
    
            
    LabCurr_Rec(mIndex).Caption = RsSavRec.AbsolutePosition
    LabCount_Rec(mIndex).Caption = RsSavRec.RecordCount




    With GRID1

        For i = 1 To .rows - 1

            If Trim(TxtSerial1(mIndex).text) = .TextMatrix(i, .ColIndex("id")) Then
                TxtSerial1(mIndex).text = .TextMatrix(i, .ColIndex("Ser"))
                .Row = i
                Exit Sub
            End If

        Next

    End With

ErrTrap:

End Sub

Public Sub FiLLTXT2()

    On Error GoTo ErrTrap
    Dim i As Integer
   ' Frame1(mIndex).Enabled = False
    TxtSerial1(mIndex).text = IIf(IsNull(RsSavRec.Fields("id").value), "", RsSavRec.Fields("id").value)
    TxtName(mIndex).text = IIf(IsNull(RsSavRec.Fields("name").value), "", RsSavRec.Fields("name").value)
    TxtNameE(mIndex).text = IIf(IsNull(RsSavRec.Fields("nameE").value), "", RsSavRec.Fields("nameE").value)
    LabCurr_Rec(mIndex).Caption = RsSavRec.AbsolutePosition
    LabCount_Rec(mIndex).Caption = RsSavRec.RecordCount
    With Grid2

        For i = 1 To .rows - 1

            If Trim(TxtSerial1(mIndex).text) = .TextMatrix(i, .ColIndex("id")) Then
                TxtSerial1(mIndex).text = .TextMatrix(i, .ColIndex("Ser"))
                .Row = i
                Exit Sub
            End If

        Next

    End With

ErrTrap:

End Sub
Public Sub FiLLTXT11()

    On Error GoTo ErrTrap
    Dim i As Integer
  '  Frame1(mIndex).Enabled = False
    TxtSerial1(mIndex).text = IIf(IsNull(RsSavRec.Fields("id").value), "", RsSavRec.Fields("id").value)
    TxtName(mIndex).text = IIf(IsNull(RsSavRec.Fields("name").value), "", RsSavRec.Fields("name").value)
    TxtNameE(mIndex).text = IIf(IsNull(RsSavRec.Fields("nameE").value), "", RsSavRec.Fields("nameE").value)
    LabCurr_Rec(mIndex).Caption = RsSavRec.AbsolutePosition
    LabCount_Rec(mIndex).Caption = RsSavRec.RecordCount
    With Grid2

        For i = 1 To .rows - 1

            If Trim(TxtSerial1(mIndex).text) = .TextMatrix(i, .ColIndex("id")) Then
                TxtSerial1(mIndex).text = .TextMatrix(i, .ColIndex("Ser"))
                .Row = i
                Exit Sub
            End If

        Next

    End With

ErrTrap:

End Sub

Public Sub FiLLTXT10()

      On Error GoTo ErrTrap
    
    Dim RsDetails As New ADODB.Recordset
    Dim StrSQL As String
    Dim RsNotes As New ADODB.Recordset
    Dim RsTest  As ADODB.Recordset
    Dim RsReplace As ADODB.Recordset
    Dim LngPartID As Long
    Dim RsPartDetails As ADODB.Recordset
    Dim i As Long

'     If Lngid <> 0 Then
'        RsSavRec.Find "ID=" & Lngid, , adSearchForward, adBookmarkFirst
'
'        If RsSavRec.BOF Or RsSavRec.EOF Then
'            Exit Sub
'        End If
'    End If

    On Error GoTo ErrTrap

ReloadCompo
    Dim Dcombos As New ClsDataCombos
    mSenesor = False

    'Frm2.Enabled = False
    TxtSerial1(mIndex).text = IIf(IsNull(RsSavRec.Fields("id").value), "", RsSavRec.Fields("id").value)
    XPDtbTrans(mIndex).value = IIf(IsNull(RsSavRec("RecordDate").value), Date, RsSavRec("RecordDate").value)
    'Me.TxtNoteSerial1.Text = IIf(IsNull(RsSavRec("NoteSerial1").value), "", RsSavRec("NoteSerial1").value)
   txtEmpName.text = IIf(IsNull(RsSavRec.Fields("EmpName").value), "", RsSavRec.Fields("EmpName").value)
'   cmbLocationsName2.Text = IIf(IsNull(RsSavRec.Fields("LocationsName").value), "", RsSavRec.Fields("LocationsName").value)
 
'RsSavRec("Photo").value = DBPix201.Image
DBPix201.Image = IIf(IsNull(RsSavRec.Fields("Photo2").value), "", RsSavRec.Fields("Photo2").value)
   txtHafizaNo = IIf(IsNull(RsSavRec.Fields("HafizaNo").value), "", RsSavRec.Fields("HafizaNo").value)
   TxtMobileNO = IIf(IsNull(RsSavRec.Fields("MobileNO").value), "", RsSavRec.Fields("MobileNO").value)
   
   TxtSalary = IIf(IsNull(RsSavRec.Fields("Salary").value), "", RsSavRec.Fields("Salary").value)
   
   txtFingerPrint = IIf(IsNull(RsSavRec.Fields("FingerPrint").value), "", RsSavRec.Fields("FingerPrint").value)
   
   
    lblFingerStatus.Tag = IIf(IsNull(RsSavRec.Fields("FingerStatus").value), "", RsSavRec.Fields("FingerStatus").value)
    lblProgressFinger.Caption = lblFingerStatus.Tag & "%"
    
    txtStartDate = IIf(IsNull(RsSavRec("StartDate").value), Date, RsSavRec("StartDate").value)
    
 '   Me.TxtNoteSerial1.Text = IIf(IsNull(RsSavRec("NoteSerial1").value), "", RsSavRec("NoteSerial1").value)
    TimeIn.value = IIf(IsNull(RsSavRec("TimeIn").value), "", RsSavRec("TimeIn").value)
    TimeOut.value = IIf(IsNull(RsSavRec("TimeOut").value), "", RsSavRec("TimeOut").value)
    
    optIsEmp = IIf(val(RsSavRec!IsEmp & "") = 0, True, False)
    
    optIsResponsible = Not optIsEmp
    
    
    
'    TxtRemarks = IIf(IsNull(RsSavRec("Remarks").value), "", RsSavRec("Remarks").value)
    
    ' txtTotalDiscPerc = IIf(IsNull(RsSavRec("DiscPercent").value), "", RsSavRec("DiscPercent").value)
    
    dcBranch(mIndex).BoundText = IIf(IsNull(RsSavRec("BranchID").value), "", RsSavRec("BranchID").value)
    ZKFPEngX1.EndEngine
    
    Label13.Caption = "€Ì— „ ’·"
 
    
    Me.DCboUserName(mIndex).BoundText = IIf(IsNull(RsSavRec("UserID").value), "", RsSavRec("UserID").value)

    s = " Select * from TblEmpDataFingerPrint Where EmpId = " & val(TxtSerial1(mIndex))
    loadgrid s, GrdFinger, True, False
    GrdFinger.rows = 11
    
     LabCurr_Rec(mIndex).Caption = RsSavRec.AbsolutePosition
    LabCount_Rec(mIndex).Caption = RsSavRec.RecordCount

    
    
    





            

        
   
ErrTrap:


End Sub
Public Sub FiLLTXT5()

    On Error GoTo ErrTrap
    Dim i As Integer
 '   Frame1(mIndex).Enabled = False
    TxtSerial1(mIndex).text = IIf(IsNull(RsSavRec.Fields("id").value), "", RsSavRec.Fields("id").value)
    TxtName(mIndex).text = IIf(IsNull(RsSavRec.Fields("name").value), "", RsSavRec.Fields("name").value)
    TxtNameE(mIndex).text = IIf(IsNull(RsSavRec.Fields("nameE").value), "", RsSavRec.Fields("nameE").value)
    LabCurr_Rec(mIndex).Caption = RsSavRec.AbsolutePosition
    LabCount_Rec(mIndex).Caption = RsSavRec.RecordCount
    With Grid5

        For i = 1 To .rows - 1

            If Trim(TxtSerial1(mIndex).text) = .TextMatrix(i, .ColIndex("id")) Then
                TxtSerial1(mIndex).text = .TextMatrix(i, .ColIndex("Ser"))
                .Row = i
                Exit Sub
            End If

        Next

    End With

ErrTrap:

End Sub


Public Sub FiLLTXT8()

    On Error GoTo ErrTrap
    Dim i As Integer
'    Frame1(mIndex).Enabled = False
    TxtSerial1(mIndex).text = IIf(IsNull(RsSavRec.Fields("id").value), "", RsSavRec.Fields("id").value)
    TxtName(mIndex).text = IIf(IsNull(RsSavRec.Fields("name").value), "", RsSavRec.Fields("name").value)
    TxtNameE(mIndex).text = IIf(IsNull(RsSavRec.Fields("nameE").value), "", RsSavRec.Fields("nameE").value)
     lblServiceColor.backcolor = val(RsSavRec("ServiceColor").value & "")
    If RsSavRec("IsBoardNo").value = True Then
        Me.chkIsBoardNo(0).value = vbChecked
    Else
        Me.chkIsBoardNo(0).value = Unchecked
    End If
    
        If RsSavRec("IsBoardNoHide").value = True Then
        Me.chkIsBoardNo(1).value = vbChecked
    Else
        Me.chkIsBoardNo(1).value = Unchecked
    End If
    
    
    LabCurr_Rec(mIndex).Caption = RsSavRec.AbsolutePosition
    LabCount_Rec(mIndex).Caption = RsSavRec.RecordCount
    With Grid8

        For i = 1 To .rows - 1

            If Trim(TxtSerial1(mIndex).text) = .TextMatrix(i, .ColIndex("id")) Then
                TxtSerial1(mIndex).text = .TextMatrix(i, .ColIndex("Ser"))
                .Row = i
                Exit Sub
            End If

        Next

    End With

ErrTrap:

End Sub
Private Sub FilltxtBord()
Dim board As String
Dim lettter As String
Dim Num As String
Dim nboard As String
Dim nlettter As String
Dim nNum As String


lettter = txtLetter1.text & " " & txtLetter2.text & " " & txtLetter3.text & " " & txtLetter4.text
Num = txtNum1.text & " " & txtNum2.text & " " & txtNum3.text & " " & txtNum4.text

nlettter = ntxtLetter1.text & " " & ntxtLetter2.text & " " & ntxtLetter3.text & " " & ntxtLetter4.text
nNum = ntxtNum1.text & " " & ntxtNum2.text & " " & ntxtNum3.text & " " & ntxtNum4.text

board = lettter & " " & Num

nboard = nlettter & " " & nNum

txtBoardNo = board
txtnBoardNo = nboard
End Sub

Public Sub FiLLTXT3(Optional Lngid As Long = 0)

    On Error GoTo ErrTrap
    
    Dim RsDetails As New ADODB.Recordset
    Dim StrSQL As String
    Dim RsNotes As New ADODB.Recordset
    Dim RsTest  As ADODB.Recordset
    Dim RsReplace As ADODB.Recordset
    Dim LngPartID As Long
    Dim RsPartDetails As ADODB.Recordset
    Dim i As Long

     If Lngid <> 0 Then
        RsSavRec.Find "ID=" & Lngid, , adSearchForward, adBookmarkFirst

        If RsSavRec.BOF Or RsSavRec.EOF Then
            Exit Sub
        End If
    End If

    On Error GoTo ErrTrap


    Dim Dcombos As New ClsDataCombos
    Dcombos.GetCustomersSuppliers 1, Me.DcCustmer, True

    'Frm2.Enabled = False
    TxtSerial1(mIndex).text = IIf(IsNull(RsSavRec.Fields("id").value), "", RsSavRec.Fields("id").value)
    XPDtbTrans(mIndex).value = IIf(IsNull(RsSavRec("RecordDate").value), Date, RsSavRec("RecordDate").value)
    'Me.TxtNoteSerial1.Text = IIf(IsNull(RsSavRec("NoteSerial1").value), "", RsSavRec("NoteSerial1").value)
    
    DcCustmer.BoundText = val(RsSavRec!CusID & "")
    
'    DcboEmp.BoundText = val(RsSavRec!EmpID & "")
    'DcboItemID1.BoundText = val(RsSavRec!ItemID & "")
    
    txtDateRec.value = IIf(IsNull(RsSavRec("DateRec").value), "", RsSavRec("DateRec").value)
    txtDateRehearsal.value = IIf(IsNull(RsSavRec("DateRehearsal").value), "", RsSavRec("DateRehearsal").value)
    txtRehearsalDateFInish.value = IIf(IsNull(RsSavRec("RehearsalDateFInish").value), "", RsSavRec("RehearsalDateFInish").value)
    txtDateDelivery.value = IIf(IsNull(RsSavRec("DateDelivery").value), "", RsSavRec("DateDelivery").value)
    txtDeliveryDateFinish.value = IIf(IsNull(RsSavRec("DeliveryDateFinish").value), "", RsSavRec("DeliveryDateFinish").value)
    txtDateDeliveryAct.value = IIf(IsNull(RsSavRec("DateDeliveryAct").value), "", RsSavRec("DateDeliveryAct").value)
    
    
    
    txtNoteSerialCash(1) = IIf(IsNull(RsSavRec("NoteIDCash").value), "", (RsSavRec("NoteIDCash").value))
    txtNoteSerialCash(0) = IIf(IsNull(RsSavRec("NoteSerialCash").value), "", (RsSavRec("NoteSerialCash").value))
    
    
'    TxtRemarks = IIf(IsNull(RsSavRec("Remarks").value), "", RsSavRec("Remarks").value)
    
    ' txtTotalDiscPerc = IIf(IsNull(RsSavRec("DiscPercent").value), "", RsSavRec("DiscPercent").value)
    
    TXTTransactionID3.text = IIf(IsNull(RsSavRec("TransactionID3").value), "", RsSavRec("TransactionID3").value)
    
    TxtNoteSerial13.text = IIf(IsNull(RsSavRec("NoteSerial13").value), "", RsSavRec("NoteSerial13").value)
    
    
    
    TXTTransactionID1.text = IIf(IsNull(RsSavRec("TransactionID1").value), "", RsSavRec("TransactionID1").value)
    
    TxtNoteSerial11.text = IIf(IsNull(RsSavRec("NoteSerial11").value), "", RsSavRec("NoteSerial11").value)
    
    
    
    txtNoteid3.text = IIf(IsNull(RsSavRec("Noteid3").value), "", RsSavRec("Noteid3").value)

    txtGeneralTotal = IIf(IsNull(RsSavRec("GeneralTotal").value), "", RsSavRec("GeneralTotal").value)
    txtTotalAdd = IIf(IsNull(RsSavRec("TotalAdd").value), "", RsSavRec("TotalAdd").value)
    txtTotalPay = IIf(IsNull(RsSavRec("TotalPay").value), "", RsSavRec("TotalPay").value)
    txtTotalDiscPerc = IIf(IsNull(RsSavRec("TotalDiscPerc").value), "", RsSavRec("TotalDiscPerc").value)
    txtTotalDisc = IIf(IsNull(RsSavRec("TotalDisc").value), "", RsSavRec("TotalDisc").value)
    txtRequiredAmount = IIf(IsNull(RsSavRec("RequiredAmount").value), "", RsSavRec("RequiredAmount").value)
    txtPaymedValue = IIf(IsNull(RsSavRec("PaymedValue").value), "", RsSavRec("PaymedValue").value)
    
    txtVatYou = IIf(IsNull(RsSavRec("VatYou").value), "", RsSavRec("VatYou").value)
    If val(txtVatYou.text) = 0 Then
        txtVatYou.text = 15
    End If
    TxtVAT = IIf(IsNull(RsSavRec("Vat").value), "", RsSavRec("Vat").value)
    txtTotalAfterVat = IIf(IsNull(RsSavRec("TotalAfterVat").value), "", RsSavRec("TotalAfterVat").value)

  
    
    txtTotalNet = IIf(IsNull(RsSavRec("TotalNet").value), "", RsSavRec("TotalNet").value)
    
    
     

   
    dcBranch(mIndex).BoundText = IIf(IsNull(RsSavRec("BranchID").value), "", RsSavRec("BranchID").value)
    
    
    Me.DCboUserName(mIndex).BoundText = IIf(IsNull(RsSavRec("UserID").value), "", RsSavRec("UserID").value)

    
     
    
 
     
    
    
    
    
'      TxtNoteID = RsSavRec!NoteID & ""
'    TxtNoteSerial = RsSavRec!NoteSerial & ""
   
  '   TxtNoteID = RsSavRec!NoteID & ""
  '  TxtNoteSerial = RsSavRec!NoteSerial & ""
    
'     If val(TxtNoteID) <> 0 Then
'        CmdCreateV2.Enabled = False
'        cmdPrintNote.Enabled = True
'        cmdDelNote.Enabled = True
'
'     Else
'        CmdCreateV2.Enabled = True
'        cmdPrintNote.Enabled = False
'        cmdDelNote.Enabled = False
'
'    End If
    LabCurr_Rec(mIndex).Caption = RsSavRec.AbsolutePosition
    LabCount_Rec(mIndex).Caption = RsSavRec.RecordCount

    Dim s As String
    
    





            
    FG.rows = 1
    
    s = " SELECT tblItems.ItemName,tblItems.ItemNamee,TblEmployee.Emp_Name,TblEmployee.Emp_Namee,TblTasks.Name as TasksName,TblTasks.PercentV, TblJobOrders2.* "
    
    s = s & " from TblJobOrders2 inner join TblTasks On TblTasks.Id = TblJobOrders2.TasksID "
    s = s & " Left outer join tblItems On tblItems.ItemId = TblJobOrders2.ItemID "
    s = s & " Left outer join TblEmployee On TblEmployee.Emp_ID= TblJobOrders2.Emp_ID "
    s = s & " Where MasterID = " & val(TxtSerial1(mIndex))
    
    loadgrid s, FG, True, True
    
    FillGridSales
    Calc
    CalcAmount
    
    FG.rows = FG.rows + 50
'CalcTotal2
ErrTrap:

End Sub




Public Sub FiLLTXT9(Optional Lngid As Long = 0)

    On Error GoTo ErrTrap
    
    Dim RsDetails As New ADODB.Recordset
    Dim StrSQL As String
    Dim RsNotes As New ADODB.Recordset
    Dim RsTest  As ADODB.Recordset
    Dim RsReplace As ADODB.Recordset
    Dim LngPartID As Long
    Dim RsPartDetails As ADODB.Recordset
    Dim i As Long

     If Lngid <> 0 Then
        RsSavRec.Find "ID=" & Lngid, , adSearchForward, adBookmarkFirst

        If RsSavRec.BOF Or RsSavRec.EOF Then
            Exit Sub
        End If
    End If

    On Error GoTo ErrTrap


    Dim Dcombos As New ClsDataCombos
    
TxtSearchCode2 = ""
    'Frm2.Enabled = False
    TxtSerial1(mIndex).text = IIf(IsNull(RsSavRec.Fields("id").value), "", RsSavRec.Fields("id").value)
    XPDtbTrans(mIndex).value = IIf(IsNull(RsSavRec("RecordDate").value), Date, RsSavRec("RecordDate").value)
    'Me.TxtNoteSerial1.Text = IIf(IsNull(RsSavRec("NoteSerial1").value), "", RsSavRec("NoteSerial1").value)
   txtCustName.text = IIf(IsNull(RsSavRec.Fields("CustName").value), "", RsSavRec.Fields("CustName").value)
'   cmbLocationsName.Text = IIf(IsNull(RsSavRec.Fields("LocationsName").value), "", RsSavRec.Fields("LocationsName").value)
'   txtCarName.Text = IIf(IsNull(RsSavRec.Fields("CarName").value), "", RsSavRec.Fields("CarName").value)
   txtPhoneCust = IIf(IsNull(RsSavRec.Fields("PhoneCust").value), "", RsSavRec.Fields("PhoneCust").value)
   CboPayMentType.ListIndex = val(RsSavRec!PaymentType & "")
    cmbPaymentClass.BoundText = val(RsSavRec!PaymentClassID & "")
    
    Me.TxtNoteSerial1.text = IIf(IsNull(RsSavRec("NoteSerial1").value), "", RsSavRec("NoteSerial1").value)
    startTime.value = IIf(IsNull(RsSavRec("StartTime").value), "", RsSavRec("StartTime").value)
    


    optCash = IIf(val(RsSavRec!PayType & "") = 0, True, False)
    
    optLater = Not optCash
    txtAmountCash.text = IIf(IsNull(RsSavRec("AmountCash").value), "", RsSavRec("AmountCash").value)
    txtAmountVisa.text = IIf(IsNull(RsSavRec("AmountVisa").value), "", RsSavRec("AmountVisa").value)
    txtTotalWithVat2.text = IIf(IsNull(RsSavRec("TotalWithVat2").value), "", RsSavRec("TotalWithVat2").value)
    
    
    txtnBoardNo.text = IIf(IsNull(RsSavRec("nBoardNo").value), "", RsSavRec("nBoardNo").value)
    txtBoardNo.text = IIf(IsNull(RsSavRec("BoardNo").value), "", RsSavRec("BoardNo").value)
    
    txtLetter1.text = IIf(IsNull(RsSavRec("txtLetter1").value), "", RsSavRec("txtLetter1").value)
    txtLetter2.text = IIf(IsNull(RsSavRec("txtLetter2").value), "", RsSavRec("txtLetter2").value)
    txtLetter3.text = IIf(IsNull(RsSavRec("txtLetter3").value), "", RsSavRec("txtLetter3").value)
    txtLetter4.text = IIf(IsNull(RsSavRec("txtLetter4").value), "", RsSavRec("txtLetter4").value)
    
    ntxtLetter1.text = IIf(IsNull(RsSavRec("ntxtLetter1").value), "", RsSavRec("ntxtLetter1").value)
    ntxtLetter2.text = IIf(IsNull(RsSavRec("ntxtLetter2").value), "", RsSavRec("ntxtLetter2").value)
    ntxtLetter3.text = IIf(IsNull(RsSavRec("ntxtLetter3").value), "", RsSavRec("ntxtLetter3").value)
    ntxtLetter4.text = IIf(IsNull(RsSavRec("ntxtLetter4").value), "", RsSavRec("ntxtLetter4").value)

    txtNum1.text = IIf(IsNull(RsSavRec("txtNum1").value), "", RsSavRec("txtNum1").value)
    txtNum2.text = IIf(IsNull(RsSavRec("txtNum2").value), "", RsSavRec("txtNum2").value)
    txtNum3.text = IIf(IsNull(RsSavRec("txtNum3").value), "", RsSavRec("txtNum3").value)
    txtNum4.text = IIf(IsNull(RsSavRec("txtNum4").value), "", RsSavRec("txtNum4").value)
    
    ntxtNum1.text = IIf(IsNull(RsSavRec("ntxtNum1").value), "", RsSavRec("ntxtNum1").value)
    ntxtNum2.text = IIf(IsNull(RsSavRec("ntxtNum2").value), "", RsSavRec("ntxtNum2").value)
    ntxtNum3.text = IIf(IsNull(RsSavRec("ntxtNum3").value), "", RsSavRec("ntxtNum3").value)
    ntxtNum4.text = IIf(IsNull(RsSavRec("ntxtNum4").value), "", RsSavRec("ntxtNum4").value)
    
   

    
    Me.DBCboClientName.BoundText = IIf(IsNull(RsSavRec("CusID").value), "", RsSavRec("CusID").value)
    
    
    txtAmountLater.text = IIf(IsNull(RsSavRec("AmountLater").value), "", RsSavRec("AmountLater").value)
    XPTxtVal.text = IIf(IsNull(RsSavRec("Value").value), "", RsSavRec("Value").value)
    TxtVAt22.text = IIf(IsNull(RsSavRec("VAt22").value), "", RsSavRec("VAt22").value)
    txtTotalWithVat2.text = IIf(IsNull(RsSavRec("TotalWithVat2").value), "", RsSavRec("TotalWithVat2").value)
    
    
    TxtSearchCode2 = IIf(IsNull(RsSavRec("Remarks").value), "", RsSavRec("Remarks").value)
    
    ' txtTotalDiscPerc = IIf(IsNull(RsSavRec("DiscPercent").value), "", RsSavRec("DiscPercent").value)
    
    dcBranch(mIndex).BoundText = IIf(IsNull(RsSavRec("BranchID").value), "", RsSavRec("BranchID").value)
    
    
    Me.DCboUserName(mIndex).BoundText = IIf(IsNull(RsSavRec("UserID").value), "", RsSavRec("UserID").value)

    
     
    
 
      txtCodeSend = "+966"
    Dim isFound As Boolean
    If Not FindString(txtPhoneCust, "+966", 1) Then
        If Not FindString(txtPhoneCust, "00966", 1) Then
            isFound = False
        End If
        isFound = False
    End If
    If Not isFound Then
        txtCodeSend = "+966"
    Else
        txtCodeSend = ""
        'txtPhoneCust = "+966" & val(txtPhoneCust)
    End If
    Dim mTxt As String
    mTxt = txtCodeSend & val(txtPhoneCust)
    lbl(56).Caption = mTxt
    
    
    
    
'      TxtNoteID = RsSavRec!NoteID & ""
'    TxtNoteSerial = RsSavRec!NoteSerial & ""
   
  '   TxtNoteID = RsSavRec!NoteID & ""
  '  TxtNoteSerial = RsSavRec!NoteSerial & ""
    
'     If val(TxtNoteID) <> 0 Then
'        CmdCreateV2.Enabled = False
'        cmdPrintNote.Enabled = True
'        cmdDelNote.Enabled = True
'
'     Else
'        CmdCreateV2.Enabled = True
'        cmdPrintNote.Enabled = False
'        cmdDelNote.Enabled = False
'
'    End If
    LabCurr_Rec(mIndex).Caption = RsSavRec.AbsolutePosition
    LabCount_Rec(mIndex).Caption = RsSavRec.RecordCount

    Dim s As String
    
    





            

        
   ReloadCompo
     Dim mCol As Long
    mCol = grd.Col
    
    For i = 0 To grd.Cols - 1
        If val(cmbPaymentClass.BoundText) = val(grd.ColKey(i)) Then
            
            grd.Col = i
        'grd.CellBackColor = vbWhite
          
          lblClassCat.backcolor = IIf(grd.CellBackColor = 0, vbWhite, grd.CellBackColor)
        End If
    Next
    
    lblClassCat.Caption = cmbPaymentClass.text
'    Dim mCol As Long
'
'      For i = 0 To grd.Cols - 1
'
'            grd.Col = i
'            grd.CellBackColor = vbWhite
'
'
'    Next
'
'    For i = 0 To grd.Cols - 1
'        If val(cmbPaymentClass.BoundText) = val(grd.ColKey(i)) Then
'            grd.Col = i
'            grd.CellBackColor = vbBlue
'            Exit For
'        End If
'    Next
    'Grd.Col = mCol
    'Grd.CellBackColor = vbBlue
    
'CalcTotal2
ErrTrap:

End Sub




Public Sub FiLLTXT4(Optional Lngid As Long = 0)

    On Error GoTo ErrTrap
    
    Dim RsDetails As New ADODB.Recordset
    Dim StrSQL As String
    Dim RsNotes As New ADODB.Recordset
    Dim RsTest  As ADODB.Recordset
    Dim RsReplace As ADODB.Recordset
    Dim LngPartID As Long
    Dim RsPartDetails As ADODB.Recordset
    Dim i As Long

     If Lngid <> 0 Then
        RsSavRec.Find "ID=" & Lngid, , adSearchForward, adBookmarkFirst

        If RsSavRec.BOF Or RsSavRec.EOF Then
            Exit Sub
        End If
    End If

    On Error GoTo ErrTrap

    'Frm2.Enabled = False
    TxtSerial1(mIndex).text = IIf(IsNull(RsSavRec.Fields("id").value), "", RsSavRec.Fields("id").value)
    XPDtbTrans(mIndex).value = IIf(IsNull(RsSavRec("RecordDate").value), Date, RsSavRec("RecordDate").value)
    'Me.TxtNoteSerial1.Text = IIf(IsNull(RsSavRec("NoteSerial1").value), "", RsSavRec("NoteSerial1").value)
    
   
    dcBranch(mIndex).BoundText = IIf(IsNull(RsSavRec("BranchID").value), "", RsSavRec("BranchID").value)
    
    
    Me.DCboUserName(mIndex).BoundText = IIf(IsNull(RsSavRec("UserID").value), "", RsSavRec("UserID").value)

    
     
    
 
     
    
    
    
    
'      TxtNoteID = RsSavRec!NoteID & ""
'    TxtNoteSerial = RsSavRec!NoteSerial & ""
   
  '   TxtNoteID = RsSavRec!NoteID & ""
  '  TxtNoteSerial = RsSavRec!NoteSerial & ""
    
'     If val(TxtNoteID) <> 0 Then
'        CmdCreateV2.Enabled = False
'        cmdPrintNote.Enabled = True
'        cmdDelNote.Enabled = True
'
'     Else
'        CmdCreateV2.Enabled = True
'        cmdPrintNote.Enabled = False
'        cmdDelNote.Enabled = False
'
'    End If
    LabCurr_Rec(mIndex).Caption = RsSavRec.AbsolutePosition
    LabCount_Rec(mIndex).Caption = RsSavRec.RecordCount

    Dim s As String
    
    





            
    FG4.rows = 1
    

    
    
s = " SELECT TblTasks.Name          AS TasksName,"
s = s & "        TblCustemers.CusName      CustName,"
s = s & "        tblItems.ItemName,"
s = s & "        TblEmployee.Emp_Name     EmpName,"
s = s & "        TblJobOrdersTasks2.*"
s = s & " From TblJobOrdersTasks2"
s = s & "        INNER JOIN TblTasks"
s = s & "             ON  TblTasks.Id = TblJobOrdersTasks2.TasksID"
s = s & "        INNER JOIN TblEmployee"
s = s & "             ON  TblEmployee.Emp_Id = TblJobOrdersTasks2.EmpId"
s = s & "        INNER JOIN TblJobOrders"
s = s & "             ON  TblJobOrders.Id = TblJobOrdersTasks2.JobOrdersNo"
s = s & "        INNER JOIN TblCustemers"
s = s & "             ON  TblCustemers.CusId = TblJobOrders.cusId"
s = s & "        INNER JOIN tblItems"
s = s & "             ON  tblItems.ItemID = TblJobOrders.ItemID"
s = s & " Where MasterID = " & val(TxtSerial1(mIndex))
loadgrid s, FG4, True, True
    
    
'CalcTotal2
ErrTrap:

End Sub



Public Sub FiLLTXT6(Optional Lngid As Long = 0)

    On Error GoTo ErrTrap
    
    Dim RsDetails As New ADODB.Recordset
    Dim StrSQL As String
    Dim RsNotes As New ADODB.Recordset
    Dim RsTest  As ADODB.Recordset
    Dim RsReplace As ADODB.Recordset
    Dim LngPartID As Long
    Dim RsPartDetails As ADODB.Recordset
    Dim i As Long

     If Lngid <> 0 Then
        RsSavRec.Find "ID=" & Lngid, , adSearchForward, adBookmarkFirst

        If RsSavRec.BOF Or RsSavRec.EOF Then
            Exit Sub
        End If
    End If

    On Error GoTo ErrTrap

    'Frm2.Enabled = False
    TxtSerial1(mIndex).text = IIf(IsNull(RsSavRec.Fields("id").value), "", RsSavRec.Fields("id").value)
    XPDtbTrans(mIndex).value = IIf(IsNull(RsSavRec("RecordDate").value), Date, RsSavRec("RecordDate").value)
    'Me.TxtNoteSerial1.Text = IIf(IsNull(RsSavRec("NoteSerial1").value), "", RsSavRec("NoteSerial1").value)
    
   
    dcBranch(mIndex).BoundText = IIf(IsNull(RsSavRec("BranchID").value), "", RsSavRec("BranchID").value)
    
    
    Me.DCboUserName(mIndex).BoundText = IIf(IsNull(RsSavRec("UserID").value), "", RsSavRec("UserID").value)

    
     
    
 
     
    
    
    
    
'      TxtNoteID = RsSavRec!NoteID & ""
'    TxtNoteSerial = RsSavRec!NoteSerial & ""
   
  '   TxtNoteID = RsSavRec!NoteID & ""
  '  TxtNoteSerial = RsSavRec!NoteSerial & ""
    
'     If val(TxtNoteID) <> 0 Then
'        CmdCreateV2.Enabled = False
'        cmdPrintNote.Enabled = True
'        cmdDelNote.Enabled = True
'
'     Else
'        CmdCreateV2.Enabled = True
'        cmdPrintNote.Enabled = False
'        cmdDelNote.Enabled = False
'
'    End If
    LabCurr_Rec(mIndex).Caption = RsSavRec.AbsolutePosition
    LabCount_Rec(mIndex).Caption = RsSavRec.RecordCount

    Dim s As String
    
    





            
    Fg6.rows = 1
    

    
    
s = " SELECT tblReservationType.Name          AS ReservationTypeName,TblItems.ItemCode,"
s = s & "        TblCustemers.CusName      CusName,"

s = s & "        TblEmployee.Emp_Name     EmpName,"
s = s & "        TblAppointmentlist2.*"
s = s & " From TblAppointmentlist2"
s = s & "        INNER JOIN tblReservationType"
s = s & "             ON  tblReservationType.Id = TblAppointmentlist2.ReservationTypeCode"
s = s & "        Left Outer JOIN TblEmployee"
s = s & "             ON  TblEmployee.Emp_Id = TblAppointmentlist2.EmpId"

s = s & "        Left Outer JOIN TblItems"
s = s & "             ON  TblItems.ItemID= TblAppointmentlist2.ItemID"

s = s & "        Left Outer JOIN TblCustemers"
s = s & "             ON  TblCustemers.CusId = TblAppointmentlist2.cusId"




s = s & " Where TblAppointmentlist2.MasterID = " & val(TxtSerial1(mIndex))

loadgrid s, Fg6, True, True
    

For i = 1 To Fg6.rows - 1

             Fg6.TextMatrix(i, Fg6.ColIndex("StillPeriod")) = GetTimeDiff(Fg6.TextMatrix(i, Fg6.ColIndex("Hours")), Time, 1, 1)
             If val(Fg6.TextMatrix(i, Fg6.ColIndex("StillPeriod"))) > 10 Then
                Fg6.TextMatrix(i, Fg6.ColIndex("StillPeriod")) = Fg6.TextMatrix(i, Fg6.ColIndex("StillPeriod")) & "  ”«⁄… "
             Else
                Fg6.TextMatrix(i, Fg6.ColIndex("StillPeriod")) = Fg6.TextMatrix(i, Fg6.ColIndex("StillPeriod")) & "  ”«⁄«  "
             End If

                Fg6.TextMatrix(i, Fg6.ColIndex("minutes")) = GetTimeDiff(Fg6.TextMatrix(i, Fg6.ColIndex("Hours")), Time, 1, 2)
                If val(Fg6.TextMatrix(i, Fg6.ColIndex("minutes"))) > 10 Then
                    Fg6.TextMatrix(i, Fg6.ColIndex("minutes")) = Fg6.TextMatrix(i, Fg6.ColIndex("minutes")) & " œﬁÌﬁ…"
                Else
                    Fg6.TextMatrix(i, Fg6.ColIndex("minutes")) = Fg6.TextMatrix(i, Fg6.ColIndex("minutes")) & " œﬁ«∆ﬁ"
                End If

Next

'CalcTotal2
ErrTrap:

End Sub

 
Public Function GetTimeDiff(ByRef StartD As Date, _
Endd As Date, DTF As Integer, Optional ByVal PartOfDate As Integer = 0) As String
On Error GoTo ErrEvt 'simple error handling

Dim ThePartH As Long 'derive hours total
Dim ThePartM As Long 'derive minutes total
Dim ThePartS As Long 'derive remaineder of seconds
Dim SecondsTot As Long 'for internal check only


Select Case DTF
Case 0
    'No change neccessary
    'The dates provided are already in Long format
Case 1
    'convert the time to a long date format thus reducing
    'the number of lines for correction of times
    StartD = Date & " " & StartD
    Endd = Date & " " & Endd
Case Else
    'raise error
    'Invalid integer option
    GetTimeDiff = ""
    
    Exit Function

End Select

'This section is for error handling only
'It is added to prevent the user form entering
'values in incorrect order and similar user related errors
If StartD = Endd Then
'this is the result returned by the function
    GetTimeDiff = "0:00:00"
    Exit Function
ElseIf StartD > Endd Then
'this is the result returned by the function
  '  GetTimeDiff = ""

 '   Exit Function
End If

'This is the section doing the trick
'simply derive the sum of all seconds to
'hours, minutes and seconds
   
ThePartH = Int(DateDiff("s", StartD, Endd) / 3600) 'rounded off hours
ThePartM = Int((DateDiff("s", StartD, Endd) - (ThePartH * 3600)) / 60) 'rounded off minutes
ThePartS = Int(DateDiff("s", StartD, Endd) - (ThePartH * 3600) - (ThePartM * 60)) 'rest is seconds

'THIS IS JUST A SECOND CALCULATION FOR INTERNAL DEBUG 'SecondsTot = DateDiff("s", StartD, EndD)

'THIS IS THE RETURN VALUE OF THE FUNCTION

If PartOfDate = 0 Then
    GetTimeDiff = ThePartH & ":" & ThePartM & ":" & ThePartS
ElseIf PartOfDate = 1 Then
    GetTimeDiff = ThePartH
ElseIf PartOfDate = 2 Then
    GetTimeDiff = ThePartM
End If
Exit Function ' Avoid Error Handling
ErrEvt:
    Select Case Err.Number
        Case 60980
    Err.Clear
  '  MsgBox "Something went wrong here!" & vbCrLf & _
    Err.description, vbCritical, "Input Error " & Err.Number
        Case 60981
    Err.Clear
  '  MsgBox "Something went wrong here!" & vbCrLf & _
    Err.description, vbCritical, "Reversed Dates " & Err.Number
        Case Else
    Err.Clear
  '  MsgBox "Something went wrong here!" & vbCrLf & _
    Err.description, vbCritical, "Error " & Err.Number
    End Select
Resume Next
End Function



Public Sub FiLLTXT7(Optional Lngid As Long = 0)

    On Error GoTo ErrTrap
    
    Dim RsDetails As New ADODB.Recordset
    Dim StrSQL As String
    Dim RsNotes As New ADODB.Recordset
    Dim RsTest  As ADODB.Recordset
    Dim RsReplace As ADODB.Recordset
    Dim LngPartID As Long
    Dim RsPartDetails As ADODB.Recordset
    Dim i As Long

     If Lngid <> 0 Then
        RsSavRec.Find "ID=" & Lngid, , adSearchForward, adBookmarkFirst

        If RsSavRec.BOF Or RsSavRec.EOF Then
            Exit Sub
        End If
    End If

    On Error GoTo ErrTrap

    'Frm2.Enabled = False
    TxtSerial1(mIndex).text = IIf(IsNull(RsSavRec.Fields("id").value), "", RsSavRec.Fields("id").value)
    XPDtbTrans(mIndex).value = IIf(IsNull(RsSavRec("RecordDate").value), Date, RsSavRec("RecordDate").value)
    'Me.TxtNoteSerial1.Text = IIf(IsNull(RsSavRec("NoteSerial1").value), "", RsSavRec("NoteSerial1").value)
    
   
    dcBranch(mIndex).BoundText = IIf(IsNull(RsSavRec("BranchID").value), "", RsSavRec("BranchID").value)
    
    
    Me.DCboUserName(mIndex).BoundText = IIf(IsNull(RsSavRec("UserID").value), "", RsSavRec("UserID").value)

    
     
    
 
     
    
    
    
    
'      TxtNoteID = RsSavRec!NoteID & ""
'    TxtNoteSerial = RsSavRec!NoteSerial & ""
   
  '   TxtNoteID = RsSavRec!NoteID & ""
  '  TxtNoteSerial = RsSavRec!NoteSerial & ""
    
'     If val(TxtNoteID) <> 0 Then
'        CmdCreateV2.Enabled = False
'        cmdPrintNote.Enabled = True
'        cmdDelNote.Enabled = True
'
'     Else
'        CmdCreateV2.Enabled = True
'        cmdPrintNote.Enabled = False
'        cmdDelNote.Enabled = False
'
'    End If
    LabCurr_Rec(mIndex).Caption = RsSavRec.AbsolutePosition
    LabCount_Rec(mIndex).Caption = RsSavRec.RecordCount

    Dim s As String
    
    





            
    fg7.rows = 1
    

    
    
s = " SELECT "
s = s & "        TblItems.ItemName,"

s = s & "        TblEmployee.Emp_Namee     EmpName,TblEmployee.fullcode,"
s = s & "        TblEmpItemsTrans2.*"
s = s & " From TblEmpItemsTrans2"
s = s & "        Left Outer JOIN TblEmployee"
s = s & "             ON  TblEmployee.Emp_Id = TblEmpItemsTrans2.EmpId"

s = s & "        Left Outer JOIN TblItems"
s = s & "             ON  TblItems.ItemID= TblEmpItemsTrans2.ItemID"





s = s & " Where MasterID = " & val(TxtSerial1(mIndex))
loadgrid s, fg7, True, True
    
    
'CalcTotal2
ErrTrap:

End Sub

  
 
 

Public Sub FiLLTXT12(Optional Lngid As Long = 0)

    On Error GoTo ErrTrap
    
    Dim RsDetails As New ADODB.Recordset
    Dim StrSQL As String
    Dim RsNotes As New ADODB.Recordset
    Dim RsTest  As ADODB.Recordset
    Dim RsReplace As ADODB.Recordset
    Dim LngPartID As Long
    Dim RsPartDetails As ADODB.Recordset
    Dim i As Long

     If Lngid <> 0 Then
        RsSavRec.Find "ID=" & Lngid, , adSearchForward, adBookmarkFirst

        If RsSavRec.BOF Or RsSavRec.EOF Then
            Exit Sub
        End If
    End If

    On Error GoTo ErrTrap

    'Frm2.Enabled = False
    TxtSerial1(mIndex).text = IIf(IsNull(RsSavRec.Fields("id").value), "", RsSavRec.Fields("id").value)
    XPDtbTrans(mIndex).value = IIf(IsNull(RsSavRec("RecordDate").value), Date, RsSavRec("RecordDate").value)
    'Me.TxtNoteSerial1.Text = IIf(IsNull(RsSavRec("NoteSerial1").value), "", RsSavRec("NoteSerial1").value)
  txtDiscountPercent.text = IIf(IsNull(RsSavRec.Fields("DiscountPercent").value), "", RsSavRec.Fields("DiscountPercent").value)
   
    dcBranch(mIndex).BoundText = IIf(IsNull(RsSavRec("BranchID").value), "", RsSavRec("BranchID").value)
    
    
    Me.DCboUserName(mIndex).BoundText = IIf(IsNull(RsSavRec("UserID").value), "", RsSavRec("UserID").value)

    
   ' XPDtbTrans(mIndex).value = IIf(IsNull(RsSavRec("RecordDate").value), Date, RsSavRec("RecordDate").value)
     
    
 
     
    
    
    
    
'      TxtNoteID = RsSavRec!NoteID & ""
'    TxtNoteSerial = RsSavRec!NoteSerial & ""
   
  '   TxtNoteID = RsSavRec!NoteID & ""
  '  TxtNoteSerial = RsSavRec!NoteSerial & ""
    
'     If val(TxtNoteID) <> 0 Then
'        CmdCreateV2.Enabled = False
'        cmdPrintNote.Enabled = True
'        cmdDelNote.Enabled = True
'
'     Else
'        CmdCreateV2.Enabled = True
'        cmdPrintNote.Enabled = False
'        cmdDelNote.Enabled = False
'
'    End If
    LabCurr_Rec(mIndex).Caption = RsSavRec.AbsolutePosition
    LabCount_Rec(mIndex).Caption = RsSavRec.RecordCount

    Dim s As String
    
    





            
    GrdIqar.rows = 1
    

    
s = " SELECT TblBranchesData.branch_name as BranchName,TblAqar.aqarname IqarName,TblAqarDetai.unitno as UnitNoName,"
s = s & "        TblAkarUnit.name unittypeName,"


s = s & "        TblIqarDiscountTrans2.*"
s = s & " From TblIqarDiscountTrans2 "
s = s & "        Left Outer JOIN TblAqar"
s = s & "             ON  TblAqar.Aqarid = TblIqarDiscountTrans2.Iqar"

s = s & "        Left Outer JOIN TblBranchesData"
s = s & "             ON  TblBranchesData.branch_id= TblIqarDiscountTrans2.branchid"

s = s & "        Left Outer JOIN TblAkarUnit"
s = s & "             ON  TblAkarUnit.id= TblIqarDiscountTrans2.unittype"

s = s & "        Left Outer JOIN TblAqarDetai"
s = s & "             ON  TblAqarDetai.id= TblIqarDiscountTrans2.UnitNo"



s = s & " Where MasterID = " & val(TxtSerial1(mIndex))
loadgrid s, GrdIqar, True, True
    
    
'CalcTotal2
ErrTrap:

End Sub

  
 
 
 
Private Sub FillGridSales()
        Dim StrSQL  As String
        
End Sub



Public Sub FillGridWithData2()

    On Error GoTo ErrTrap

    Dim i As Integer
    Dim rs As ADODB.Recordset
    Dim My_SQL As String

    Set rs = New ADODB.Recordset
    My_SQL = "select * From TblSizesNames order by id"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText

    With Me.Grid2
        .rows = 2
        .Clear flexClearScrollable

        If rs.RecordCount > 0 Then
            .rows = rs.RecordCount + 1
            rs.MoveFirst

            For i = 1 To .rows - 1
                .TextMatrix(i, .ColIndex("Ser")) = i
             
                .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(rs.Fields("name").value), "", rs.Fields("name").value)
                .TextMatrix(i, .ColIndex("namee")) = IIf(IsNull(rs.Fields("namee").value), "", rs.Fields("namee").value)
                
               
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs.Fields("id").value), "", rs.Fields("id").value)
                
            
                '    .Cell(flexcpBackColor, .Rows - 1, 1, .Rows - 1, .Cols - 1) = vbYellow
                rs.MoveNext
            Next

            rs.Close
        End If

        .RowHeight(-1) = 300
    End With

ErrTrap:
End Sub


Private Sub ChangeLang()
    Dim XPic As IPictureDisp
    Set XPic = Me.btnFirst.ButtonImage
    Set Me.btnFirst.ButtonImage = Me.btnLast.ButtonImage
    Set Me.btnLast.ButtonImage = XPic
    Set XPic = Me.btnPrevious.ButtonImage
    Set Me.btnPrevious.ButtonImage = Me.btnNext.ButtonImage
    Set Me.btnNext.ButtonImage = XPic

    Me.Caption = "Religions"
    Me.Label1(2).Caption = Me.Caption
    Label1(3).Caption = "Code"
    Label1(0).Caption = "Name AR"
    Label1(1).Caption = "Name ENG"

    Label2(0).Caption = "Current Record"
    Label2(1).Caption = "NO. Recordes"

    btnNew.Caption = "New"
    btnModify.Caption = "Modify"
    btnSave.Caption = "Save"
    BtnUndo.Caption = "Undo"
    btnDelete.Caption = "Delete"
    btnCancel.Caption = "Exit"

    With Me.Grid
        .TextMatrix(0, .ColIndex("Ser")) = "Ser"
        .TextMatrix(0, .ColIndex("id")) = "Id"
        .TextMatrix(0, .ColIndex("name")) = "Name AR"
        .TextMatrix(0, .ColIndex("namee")) = "Name ENG"
    End With

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)
    Dim IntResult As String
    Dim StrMSG As String
    On Error GoTo ErrTrap
    Dim mSelectModFlg As String

    If mIndex = 0 Then
        mSelectModFlg = Me.TxtModFlg.text
    Else
        mSelectModFlg = Me.TxtModFlg2(mIndex).text
    End If
    
    
   ' If mSelectModFlg <> "R" Then
    

        Select Case mSelectModFlg
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



   IntResult = MsgBox(StrMSG, vbMsgBoxRight + vbYesNoCancel + vbMsgBoxRtlReading + vbQuestion, App.Title)

        Select Case IntResult

            Case vbYes
                Cancel = True
                If mIndex = 0 Then
                    btnSave_Click
                Else
                    btn_Save_Click CInt(mIndex)
                End If

            Case vbCancel
                Cancel = True
        End Select




   ' End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Terminate()
    'Set cSearchDCombo = Nothing
    'Set BKGrndPic = Nothing
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
    If mIndex = 10 Then
       ZKFPEngX1.FreeFPCacheDB (fpcHandle)
    ElseIf mIndex = 10 Then
        ZKFPEngX2.FreeFPCacheDB (fpcHandle)
    End If
ErrTrap:
End Sub

Private Sub Form_Activate()
    Me.ZOrder 0
End Sub



Public Sub AddNewRec()
    On Error GoTo ErrTrap
    Dim StrRecID As String
   

    If mIndex = 0 Then
        StrRecID = new_id("TblTasks", "id", "")
         RsSavRec.AddNew
        RsSavRec.Fields("Id").value = IIf(StrRecID <> "", StrRecID, Null)
        TxtSerial1(mIndex).text = StrRecID
    ElseIf mIndex = 1 Then
    
        
        StrRecID = new_id("TblSizesNames", "id", "")
        RsSavRec.AddNew
        RsSavRec.Fields("id").value = IIf(StrRecID <> "", StrRecID, Null)
       ' FiLLRec1
    ElseIf mIndex = 2 Then
        StrRecID = new_id("dean", "id", "")
        RsSavRec.AddNew
        RsSavRec.Fields("id").value = IIf(StrRecID <> "", StrRecID, Null)
        FiLLRec

    ElseIf mIndex = 3 Then
        StrRecID = new_id("TblJobOrders", "id", "")
        RsSavRec.AddNew
        RsSavRec.Fields("id").value = IIf(StrRecID <> "", StrRecID, Null)
    '    FiLLRec3
    ElseIf mIndex = 4 Then
        StrRecID = new_id("TblJobOrdersTasks", "id", "")
        RsSavRec.AddNew
        RsSavRec.Fields("id").value = IIf(StrRecID <> "", StrRecID, Null)
    ElseIf mIndex = 5 Then
        StrRecID = new_id("tblReservationType", "id", "")
        RsSavRec.AddNew
        RsSavRec.Fields("id").value = IIf(StrRecID <> "", StrRecID, Null)
        
    ElseIf mIndex = 6 Then
        StrRecID = new_id("TblAppointmentlist", "id", "")
        RsSavRec.AddNew
        RsSavRec.Fields("id").value = IIf(StrRecID <> "", StrRecID, Null)
        
    '    FiLLRec3
    ElseIf mIndex = 7 Then
        StrRecID = new_id("TblEmpItemsTrans", "id", "")
        RsSavRec.AddNew
        RsSavRec.Fields("id").value = IIf(StrRecID <> "", StrRecID, Null)
      ElseIf mIndex = 8 Then
        StrRecID = new_id("tblPaymentClass", "id", "")
        RsSavRec.AddNew
        RsSavRec.Fields("id").value = IIf(StrRecID <> "", StrRecID, Null)
    ElseIf mIndex = 9 Then
        StrRecID = new_id("TblTripReg", "id", "")
        RsSavRec.AddNew
        RsSavRec.Fields("id").value = IIf(StrRecID <> "", StrRecID, Null)
    ElseIf mIndex = 10 Then
        StrRecID = new_id("TblEmpData", "id", "")
        RsSavRec.AddNew
        RsSavRec.Fields("id").value = IIf(StrRecID <> "", StrRecID, Null)
        DBPix201.ImageClear
     '   DBPix202.ImageClear
    '    FiLLRec3
ElseIf mIndex = 12 Then
        StrRecID = new_id("TblIqarDiscountTrans", "id", "")
        RsSavRec.AddNew
        RsSavRec.Fields("id").value = IIf(StrRecID <> "", StrRecID, Null)
        
    End If
    
ErrTrap:
   
   
  
    

End Sub



Public Sub FiLLRec()
    On Error GoTo ErrTrap

    RsSavRec.Fields("name").value = IIf(TxtVacName.text <> "", Trim(TxtVacName.text), Null)
    RsSavRec.Fields("namee").value = IIf(TxtVacNamee.text <> "", Trim(TxtVacNamee.text), Null)

    RsSavRec.update
         If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox " „  ⁄„·Ì… «·Õ›Ÿ »‰Ã«Õ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Else
                MsgBox "Saved", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        End If
    FillGridWithData
    TxtModFlg = "R"

    Exit Sub
ErrTrap:

    If RsSavRec.EditMode <> adEditNone Then
        RsSavRec.CancelUpdate
    End If

End Sub

Public Sub FiLLTXT()

    On Error GoTo ErrTrap
    Dim i As Integer
  '  Frm2.Enabled = False
    TxtVac_ID.text = IIf(IsNull(RsSavRec.Fields("id").value), "", RsSavRec.Fields("id").value)
    TxtVacName.text = IIf(IsNull(RsSavRec.Fields("name").value), "", RsSavRec.Fields("name").value)
    TxtVacNamee.text = IIf(IsNull(RsSavRec.Fields("namee").value), "", RsSavRec.Fields("namee").value)

    LabCurrRec.Caption = RsSavRec.AbsolutePosition
    LabCountRec.Caption = RsSavRec.RecordCount

    With Grid

        For i = 1 To .rows - 1

            If Trim(TxtVac_ID.text) = .TextMatrix(i, .ColIndex("id")) Then
                TxtSerial.text = .TextMatrix(i, .ColIndex("Ser"))
                .Row = i
                Exit Sub
            End If

        Next

    End With

ErrTrap:

End Sub

Public Sub EditRec(StrTable As String, _
                   RecId As String)
    'My_SQL = "select * From " & StrTable & " where "
    'RsSavRec.Open My_SQL, cn, adOpenStatic, adLockReadOnly, adCmdTableDirect
    FiLLRec

End Sub

Private Sub Grid_EnterCell()
    On Error GoTo ErrTrap
    FindRec val(Me.Grid.TextMatrix(Me.Grid.Row, Me.Grid.ColIndex("id")))
ErrTrap:
End Sub

Private Sub txtPaymedValue_Validate(Cancel As Boolean)
Calc
End Sub

Private Sub TxtPhone_Validate(Cancel As Boolean)
If Trim(TxtPhone) = "" Then

    Dim Dcombos As New ClsDataCombos
    Dcombos.GetCustomersSuppliers 1, Me.DcCustmer, True
End If
End Sub

Private Sub txtPhoneCust_KeyPress(KeyAscii As Integer)
 KeyAscii = KeyAscii_Num(KeyAscii, Me.txtPhoneCust.text, 1)
End Sub




Public Function KeyAscii_Num(KeyAsc As Integer, _
                             Txt As String, _
                             Optional IntFilterType As Integer = 0) As Integer

    'IntFilterType=0 Readl Number
    'IntFilterType=1 Integer Number

    If KeyAsc = 8 Then
        KeyAscii_Num = KeyAsc
        Exit Function
    End If

    If IntFilterType = 0 Then
        If CBool(InStr(1, ".", CHR(KeyAsc))) And CBool(InStr(1, Txt, CHR(KeyAsc))) Then
            KeyAscii_Num = 0
            Exit Function
        ElseIf InStr(1, "+0123456789.", CHR(KeyAsc)) = 0 Then
            KeyAscii_Num = 0
        Else
            KeyAscii_Num = KeyAsc
        End If

    ElseIf IntFilterType = 1 Then

        If InStr(1, "+0123456789", CHR(KeyAsc)) = 0 Then
            KeyAscii_Num = 0
        Else
            KeyAscii_Num = KeyAsc
        End If
    End If

End Function

Private Sub TxtSearchCode2_KeyPress(KeyAscii As Integer)
 Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        'GetTblCustemersCode TxtSearchCode.Text, EmpID
        'DBCboClientName.BoundText = EmpID
        GetCustomerNamebyPhone2 , , , TxtSearchCode2.text
    End If
End Sub

Private Sub txtTotalAdd_Validate(Cancel As Boolean)
Calc
End Sub

Private Sub txtTotalDisc_Validate(Cancel As Boolean)
If (val(txtGeneralTotal) + val(txtTotalAdd)) <> 0 Then
    txtTotalDiscPerc = val(txtTotalDisc) / (val(txtGeneralTotal) + val(txtTotalAdd)) * 100
End If

Calc
End Sub

Private Sub txtTotalDiscPerc_Validate(Cancel As Boolean)
Calc
End Sub

Private Sub txtTotalNet_Change()
CalcAmount
End Sub

Private Sub TxtVac_ID_Change()
    Dim TxtMod As String
    TxtMod = TxtModFlg.text
    TxtModFlg.text = ""
    TxtModFlg = TxtMod
End Sub

Public Function FindRec(ByVal RecId As Long, Optional ByVal mIndex2 As Integer = 0)
    On Error GoTo ErrTrap
    RsSavRec.Find "id=" & RecId, , adSearchForward, 1
    mIndex2 = mIndex
    If Not (RsSavRec.EOF) Then
        If mIndex2 = 2 Then
            FiLLTXT
        ElseIf mIndex2 = 0 Then
            FiLLTXT1
        ElseIf mIndex2 = 1 Then
            FiLLTXT2
        ElseIf mIndex2 = 3 Then
            FiLLTXT3
        ElseIf mIndex2 = 4 Then
            FiLLTXT4
        ElseIf mIndex2 = 5 Then
            FiLLTXT5
        ElseIf mIndex2 = 6 Then
            FiLLTXT6
            
        ElseIf mIndex2 = 8 Then
            FiLLTXT8
            

        End If
    End If

    Exit Function
ErrTrap:

    If RsSavRec.EditMode <> adEditNone Then
        RsSavRec.CancelUpdate
        If mIndex = 2 Then
            BtnUndo_Click
        Else
            Btn_Undo_Click (mIndex2)
       
        End If
        
        
    End If

    'RsSavRec.Filter = adFilterNone
End Function
'Private Sub TxtVacCode_KeyPress(KeyAscii As Integer)
'KeyAscii = DataFormat(ChrOnly, KeyAscii)
'End Sub

Private Sub TxtModFlg_Change()

    If TxtModFlg.text = "N" Then
        Frm2.Enabled = True
        Me.btnNew.Enabled = False
        btnModify.Enabled = False
        btnDelete.Enabled = False
        'Me.btnQuery.Enabled = False
        Grid.Enabled = False
        BtnUndo.Enabled = True
        Me.btnSave.Enabled = True
     '   BtnUpdate.Enabled = False
        '  btnNext.Enabled = False
        '  btnPrevious.Enabled = False
        '  btnFirst.Enabled = False
        '  btnLast.Enabled = False
    
    ElseIf TxtModFlg.text = "R" Then
        Frm2.Enabled = True
        Grid.Enabled = True
        btnModify.Enabled = False
        btnDelete.Enabled = False

        If TxtVac_ID.text <> "" Then
            btnModify.Enabled = True
            btnDelete.Enabled = True
        End If

       ' BtnUpdate.Enabled = True
        'Me.btnQuery.Enabled = True
        Me.btnNew.Enabled = True
        BtnUndo.Enabled = False
        Me.btnSave.Enabled = False
    
        btnNext.Enabled = True
        btnPrevious.Enabled = True
        btnFirst.Enabled = True
        btnLast.Enabled = True
    
    ElseIf TxtModFlg.text = "E" Then
        Frm2.Enabled = True
        Me.btnNew.Enabled = False
        btnModify.Enabled = False
        btnDelete.Enabled = False
        'Me.btnQuery.Enabled = False
'        BtnUpdate.Enabled = False
        BtnUndo.Enabled = True
        Me.btnSave.Enabled = True
        Grid.Enabled = False
        btnNext.Enabled = False
        btnPrevious.Enabled = False
        btnFirst.Enabled = False
        btnLast.Enabled = False
    
    End If

End Sub

Public Sub FillGridWithData()

    On Error GoTo ErrTrap

    Dim i As Integer
    Dim rs As ADODB.Recordset
    Dim My_SQL As String

    Set rs = New ADODB.Recordset
    My_SQL = "select * From dean order by id"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText

    With Me.Grid
        .rows = 2
        .Clear flexClearScrollable

        If rs.RecordCount > 0 Then
            .rows = rs.RecordCount + 1
            rs.MoveFirst

            For i = 1 To .rows - 1
                .TextMatrix(i, .ColIndex("Ser")) = i
             
                .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(rs.Fields("name").value), "", rs.Fields("name").value)
               
                .TextMatrix(i, .ColIndex("namee")) = IIf(IsNull(rs.Fields("namee").value), "", rs.Fields("namee").value)
               
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs.Fields("id").value), "", rs.Fields("id").value)
            
                rs.MoveNext
            Next

            rs.Close
        End If

        .RowHeight(-1) = 300
    End With

ErrTrap:
End Sub

'-------------------------------------------------------------
Private Sub ShowTip()
    On Error GoTo ErrTrap
    Dim TTP As New clstooltip
    Dim Wrap As String
    Dim Msg As String
    Wrap = CHR(13) + CHR(10)

    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "ÃœÌœ" & Wrap & "·› Õ ”Ã· ÃœÌœ " & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F12 √Ê Enter"
            
        .AddControl btnNew, Msg, True
    End With

    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = " ⁄œÌ·" & Wrap & "· ⁄œÌ·  ”Ã· «·Õ«·Ï " & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F11"
        .AddControl btnModify, Msg, True
    End With

    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Õ›Ÿ" & Wrap & "· ”ÃÌ· «·»Ì«‰«  œ«Œ· ﬁ«⁄œ… " & Wrap & "«·»Ì«‰«  ≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F10"
        .AddControl btnSave, Msg, True
    End With

    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = " —«Ã⁄" & Wrap & "·· —«Ã⁄ ⁄‰ «·⁄„·Ì… «·Õ«·Ì…" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F9"
        .AddControl BtnUndo, Msg, True
    End With

    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Õ–› «·”Ã·" & Wrap & "·Õ–› «·”Ã· «·Õ«·Ï" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F18"
        .AddControl btnDelete, Msg, True
    End With

    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Œ—ÊÃ" & Wrap & "·≈€·«ﬁ Â–Â «·‰«›–…" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " Ctrl+x"
        .AddControl btnCancel, Msg, True
    End With

    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«·«Ê·" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·«Ê·" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " Home √Ê UpArrow"
        .AddControl btnFirst, Msg, True
    End With

    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«·”«»ﬁ" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " PageUp √Ê LeftArrow"
        .AddControl btnPrevious, Msg, True
    End With

    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«· «·Ï" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ï" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " PageDown √Ê RightArrow"
        .AddControl btnNext, Msg, True
    End With

    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«·«ŒÌ—" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·«ŒÌ—" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " End √Ê DownArrow"
        .AddControl btnLast, Msg, True
    End With

ErrTrap:
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        If Me.TxtModFlg.text = "R" Then
        '    btnNew_Click
        Else
            Sendkeys "{TAB}"
        End If
    End If

    'New ---------------------------
    If KeyCode = vbKeyF12 Then
        If btnNew.Enabled = False Then Exit Sub
        btnNew_Click
    End If

    'Edit ------------------------
    If KeyCode = vbKeyF11 Then
        If btnModify.Enabled = False Then Exit Sub
        btnModify_Click
    End If

    'save --------------------------------------------------------------------------------
    If KeyCode = vbKeyF10 Then
        If btnSave.Enabled = False Then Exit Sub
        btnSave_Click
    End If

    'undo ------------------------------------------------------------------------------
    If KeyCode = vbKeyF9 Then
        If BtnUndo.Enabled = False Then Exit Sub
        BtnUndo_Click
    End If

    'Delete ---------------------------------------------------------------------------
    If KeyCode = vbKeyF8 Then
        If btnDelete.Enabled = False Then Exit Sub
        btnDelete_Click
    End If

    'Exit ----------------------------------------------------------------------
    If Shift = 2 Then
        If KeyCode = vbKeyX Then
            If btnCancel.Enabled = False Then Exit Sub
            BtnCancel_Click
        End If
    End If

    'Moveing through Records ---------------------------------------------------------------------------
    'If TxtModFlg.Text = "R" Then
    'Move first --------------------------------------------
    If KeyCode = vbKeyUp Or KeyCode = vbKeyHome Then
        If btnFirst.Enabled = False Then Exit Sub
     '   BtnFirst_Click
    End If

    'Move Previous---------------------------------------------------------
    If KeyCode = vbKeyLeft Or KeyCode = vbKeyPageUp Then
        If btnPrevious.Enabled = False Then Exit Sub
       ' BtnPrevious_Click
    End If

    'Move Next---------------------------------------------------------
    If KeyCode = vbKeyRight Or KeyCode = vbKeyPageDown Then
        If btnNext.Enabled = False Then Exit Sub
   '     BtnNext_Click
    End If

    'Move Last---------------------------------------------------------
    If KeyCode = vbKeyDown Or KeyCode = vbKeyEnd Then
        If btnLast.Enabled = False Then Exit Sub
       ' BtnLast_Click
    End If

    'End If

    Exit Sub
ErrTrap:
End Sub

Private Function CheckDelCountry(Lngid As Long) As Boolean
    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    StrSQL = "Select * From TblEmployee Where id=" & Lngid & ""
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.BOF Or rs.EOF) Then
        CheckDelCountry = False
    Else
        CheckDelCountry = True
    End If

    rs.Close
    Set rs = Nothing
End Function

Private Sub TxtVacName_GotFocus()
    SwitchKeyboardLang LANG_ARABIC
End Sub

Private Sub TxtVacNamee_GotFocus()
    SwitchKeyboardLang LANG_ENGLISH

End Sub
Private Sub CreateSales()
 
Dim s As String, StrSqlDel As String, StrSQL As String
Dim BeginTrans As Boolean
Dim rsDummy As New ADODB.Recordset
Dim StoreId1 As Integer
Dim StrTempAccountCode As String
s = "Select StoreID,StoreID,StoreID2,StoreID3 from tblUsers Where UserID = " & user_id
Set rsDummy = New ADODB.Recordset

rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly, adCmdText
If Not rsDummy.EOF Then
    StoreId1 = val(rsDummy!StoreID & "")
End If

Dim rsOut As New ADODB.Recordset
Dim Current_case As Integer, mBoxID As Long
Set rsOut = New ADODB.Recordset

s = "Select BoxID From TblBoxesData Where Empid In (Select tblUsers.EmpId from tblUsers where UserId = " & user_id & " )"
rsOut.Open s, Cn, adOpenStatic, adLockReadOnly
If Not rsOut.EOF Then
    mBoxID = val(rsOut!BoxID & "")
End If


StrTempAccountCode = GetMyAccountCode("TblBoxesData", "BoxID", mBoxID)


'-----------------------------------
                    
                    
   ' Cn.BeginTrans
    BeginTrans = True
        

    StrSqlDel = "Select Transaction_ID,NoteID from Transactions Where nots = " & val(Me.TXTTransactionID3.text) & " and  Transaction_Type = 19 "
    Set rsOut = New ADODB.Recordset
    rsOut.Open StrSqlDel, Cn, adOpenStatic, adLockReadOnly
    If Not rsOut.EOF Then
        StrSqlDel = "delete From Transactions where Transaction_ID=" & val(rsOut!Transaction_ID & "")  'Val(rs("Transaction_ID").value)
        Cn.Execute StrSqlDel, , adExecuteNoRecords
                
        
        StrSqlDel = "delete From Transaction_Details where Transaction_ID=" & val(rsOut!Transaction_ID & "")  'Val(rs("Transaction_ID").value)
        Cn.Execute StrSqlDel, , adExecuteNoRecords
        
        StrSqlDel = "delete From Notes where NoteID=" & val(rsOut!NoteID & "")   'Val(rs("Transaction_ID").value)
        Cn.Execute StrSqlDel, , adExecuteNoRecords
        StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(rsOut!NoteID & "")
        Cn.Execute StrSQL, , adExecuteNoRecords
        
        Cn.Execute "Delete from TransactionValueAdded where Transaction_ID=" & val(rsOut!Transaction_ID & "") & ""
    End If
    StrSqlDel = "delete From Transactions where Transaction_ID=" & val(Me.TXTTransactionID3.text) 'Val(rs("Transaction_ID").value)
    Cn.Execute StrSqlDel, , adExecuteNoRecords
            
    
    StrSqlDel = "delete From Transaction_Details where Transaction_ID=" & val(Me.TXTTransactionID3.text) 'Val(rs("Transaction_ID").value)
    Cn.Execute StrSqlDel, , adExecuteNoRecords
    
    
    StrSqlDel = "delete From Transactions where Transaction_ID=" & val(Me.TXTTransactionID1.text) 'Val(rs("Transaction_ID").value)
    Cn.Execute StrSqlDel, , adExecuteNoRecords
            
    
    StrSqlDel = "delete From Transaction_Details where Transaction_ID=" & val(Me.TXTTransactionID1.text) 'Val(rs("Transaction_ID").value)
    Cn.Execute StrSqlDel, , adExecuteNoRecords
    
    
    StrSqlDel = "delete From Notes where NoteID=" & val(Me.txtNoteid3.text)  'Val(rs("Transaction_ID").value)
    Cn.Execute StrSqlDel, , adExecuteNoRecords
    StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Transaction_ID=" & val(Me.TXTTransactionID3.text)
    Cn.Execute StrSQL, , adExecuteNoRecords
    
    Cn.Execute "Delete from TransactionValueAdded where Transaction_ID=" & val(Me.TXTTransactionID3.text) & ""
    
      
    
    
    If Trim(TxtNoteSerial13.text) = "" Then
        TxtNoteSerial13.text = Voucher_coding(val(val(dcBranch(mIndex).BoundText)), XPDtbTrans(mIndex).value, 7, 170, , 21, , StoreId1)
    End If
            
            
    'CreateSalesTrans Dcbranch(mIndex).BoundText, 0, XPDtbTrans(mIndex).value, 21, 0, val(user_id), 0, DcCustmer.BoundText, CDbl(StoreId1), 1, val(DcboEmp.BoundText), TxtRemarks & "  "   "›« Ê—… „»Ì⁄«  »‰«¡« ⁄·Ï «„— ‘€· —ﬁ„ " & TxtSerial1(mIndex)
    CreateSalesTrans val(dcBranch(mIndex).BoundText), 0, XPDtbTrans(mIndex).value, 21, 0, val(user_id), 0, DcCustmer.BoundText, CDbl(StoreId1), 1, mEmpId, "›« Ê—… „»Ì⁄«  »‰«¡« ⁄·Ï «„— ‘€· —ﬁ„ " & TxtSerial1(mIndex)
    
'
    StrSQL = "UPDATE TblJobOrders SET  TransactionID3=" & val(TXTTransactionID3) & ",TransactionID1=" & val(TXTTransactionID1) & ",   NoteSerial11='" & TxtNoteSerial11 & "',  Noteid3=" & val(txtNoteid3) & ", NoteSerial13='" & TxtNoteSerial13 & "',NoteIDCash = " & val(Me.txtNoteSerialCash(1).text) & ",NoteSerialCash = '" & Trim(Me.txtNoteSerialCash(0).text) & "' WHERE ID  =" & val(TxtSerial1(mIndex))
    Cn.Execute StrSQL


'
'    StrSQL = "UPDATE TblJobOrders SET  Noteid3=" & val(txtNoteid3) & " , TransactionID3=" & val(TXTTransactionID3) & ",  NoteSerial13='" & TxtNoteSerial13 & "' WHERE ID  =" & val(TxtSerial1(mIndex))
'    Cn.Execute StrSQL
'Cn.CommitTrans


End Sub



Private Sub CreateSalesTrans(BranchID As Double, _
BoxID As Double, _
Transaction_Date As Date, _
Transaction_Type As Double, _
CBoBasedON As Double, _
UserID As Double, _
Trans_DiscountType As Double, _
CusID As Double, _
StoreID As Double, _
PaymentType As Double, _
Emp_id As Long, _
TransactionComment As String)

Dim BolTemp As Boolean
Dim sql As String
Dim Msg As String
Dim NoteID As Long
Dim Transaction_ID As Long
Dim Transaction_ID1 As Long
Dim Transaction_serial As String
Dim NoteSerial As String
Dim NoteSerial1 As String
Dim StrSQL As String
Dim Percetage As Double
Dim AccountVATCreit As String
Dim mPrice As Double
Dim PercetageVat As Double
Dim StoreAccount  As String
' «·”⁄— Â‰« ÂÊ ’«›Ï «·”⁄— »⁄œ Œ’„ «·«÷«›Ï Ê«·Œ’Ê„« 
If BranchID = 0 Then BranchID = mBranchID
PercentgValueAddedAccount_Transec XPDtbTrans(mIndex).value, 21, 0, AccountVATCreit, Percetage
PercetageVat = Percetage

'BillTOTAL = 0
'CostTOTAL = 0
'Check

  Set rsDummy = New ADODB.Recordset
      s = "Select EmpID from tblUsers where UserId = " & user_id
    rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly
    If Not rsDummy.EOF Then
        mEmpId = val(rsDummy!EmpID & "")
    End If
 If val(mEmpId) = 0 Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "ÌÃ»  ÕœÌœ «”„ «·»«∆⁄/«·„‰œÊ» —«Ã⁄ «·„” Œœ„ Ê«·„ÊŸ› «·„—»Êÿ »Â..!!!"
        Else
            Msg = "Must Specify SalesPerson/Saller..!!!"
        End If
        'Cmd(2).Enabled = True
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title

        Sendkeys "{F4}"
        Screen.MousePointer = vbDefault
        btn_Save(mIndex).Enabled = True
        Exit Sub
    End If
    

 If TxtNoteSerial13 = "" Then
 NoteSerial1 = Voucher_coding(val(BranchID), Transaction_Date, 7, 170, , 21)
 TxtNoteSerial13 = NoteSerial1
 End If
Transaction_ID = CStr(new_id("Transactions", "Transaction_ID", "", True))
 
   
    NoteSerial1 = Voucher_coding(val(BranchID), Transaction_Date, 7, 170, , 21)  '„»Ì⁄« 
        If NoteSerial1 = "" Then
                 If NoteSerial1 = "error" Then
                     MsgBox " ·« Ì„ﬂ‰ «÷«›… ”‰œ   „»Ì⁄«   ÃœÌœ ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": Exit Sub
                 ElseIf NoteSerial1 = "" Then
                         MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ «·”‰œ ÌœÊÌ« ﬂ„« Õœœ   ": Exit Sub
        
                 End If
        End If

NoteSerial = Notes_coding(val(BranchID), Transaction_Date)
 If NoteSerial = "" Then
            If NoteSerial = "error" Then
                MsgBox " ·« Ì„ﬂ‰ «÷«›… ﬁÌÊœ ·Â–… «·⁄„·Ì… ·«‰ﬂ  ⁄œÌ  «·Õœ «·«ﬁ’Ì ··ﬁÌÊœ ﬂ„« Õœœ  ›Ì  —ﬁÌ„ «·”‰œ«  ": Exit Sub
            ElseIf NoteSerial = "" Then
                    MsgBox " ·«Ì„ﬂ‰ «‰‘«¡ «·ﬁÌœ ·Â–« «·„” ‰œ ·«‰ﬂ Õœœ   —ﬁÌ„ ﬁÌÊœ ÌœÊÌ  ": Exit Sub
                 
            End If
End If
           
              
  
   '«· √ﬂœ „‰ ⁄œ„  ﬂ—«— —ﬁ„ «·›« Ê—…
    If Voucher_coding(val(BranchID), XPDtbTrans(mIndex).value, 7, 170, , 21) = "" Then
        If Me.TxtModFlg2(mIndex).text = "N" Then
    
            BolTemp = UniqueNoteSerial1(Trim(Me.TxtNoteSerial13.text), 21, , CInt(BranchID))
        ElseIf Me.TxtModFlg.text = "E" Then
        
            BolTemp = UniqueNoteSerial1(Trim(Me.TxtNoteSerial13.text), 21, Transaction_ID, CInt(BranchID))
        End If
 
        If BolTemp = False Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "—ﬁ„ «·›« Ê—… „”Ã· „”»ﬁ« ›Ï «·»—‰«„Ã.." & CHR(13)
                Msg = Msg & "Ê·«Ì„ﬂ‰  ﬂ—«— —ﬁ„ «·›« Ê—…"
            Else
                Msg = "This Bill No Already Exist" & CHR(13)
        
            End If
            btn_Save(mIndex).Enabled = True
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            TxtNoteSerial13.SetFocus
            Screen.MousePointer = vbDefault
          '  Cmd(2).Enabled = True
            Exit Sub
        End If
     
    End If
      
  
'
'           CostAccount = get_account_code_branch(1, CInt(BranchID))
'
'            If CostAccount = "NO branch" Or CostAccount = "NO account" Then
'                If SystemOptions.UserInterface = ArabicInterface Then
'                    MsgBox "·„ Ì „ —»ÿ  ﬂ·›… «·«‰ «Ã „Ê«œ  ", vbCritical
'                Else
'                    MsgBox "Sales Not Created", vbCritical
'                End If
'
'             Exit Sub
'              End If
              
              

'            StoreAccount = get_store_Account(CInt(StoreId), "Account_Code")
'            If StoreAccount = "" Then
'                If SystemOptions.UserInterface = ArabicInterface Then
'                    MsgBox "·„ Ì „  ÕœÌœ Õ”«»  ··„Œ“Ê‰ «·”·⁄Ì ·Â–« «·„Œ“‰ ›Ì Â–« «·›—⁄    ", vbCritical
'                Else
'                    MsgBox "No inventory account for this store has been specified in this section", vbCritical
'                End If
'                Exit Sub
'            End If



 'end Check
 Dim Account_Code_dynamic213  As String
 Account_Code_dynamic213 = get_account_code_branch(213, CStr(mBranchID))
 
        TXTTransactionID3.text = Transaction_ID
        TxtNoteSerial13.text = NoteSerial1
     Dim rsOut As New ADODB.Recordset
            Dim Current_case As Integer, mBoxID As Long
            Set rsOut = New ADODB.Recordset
            
            s = "Select BoxID From TblBoxesData Where Empid In (Select tblUsers.EmpId from tblUsers where UserId = " & user_id & " )"



            rsOut.Open s, Cn, adOpenStatic, adLockReadOnly
            If Not rsOut.EOF Then
                BoxID = val(rsOut!BoxID & "")
            End If
           ' mBoxID = val(DcboBox.BoundText)
 sql = "INSERT INTO  Transactions (  "
sql = sql & " Transaction_ID ,"
sql = sql & " BranchID ,"
sql = sql & " NoteSerial ,"
sql = sql & " NoteSerial1 ,"
sql = sql & " boxId ,"
sql = sql & " Transaction_serial ,"
sql = sql & " Transaction_Date ,"
sql = sql & " Transaction_Type ,"
sql = sql & " BillBasedOn ,"
sql = sql & " UserID ,"
sql = sql & " Trans_DiscountType ,"
sql = sql & " CusID ,"
sql = sql & " StoreId ,"
sql = sql & " PaymentType ,"
sql = sql & " Emp_id ,"
sql = sql & " Transaction_NetValue ,"
sql = sql & " Vat, netvalue, PayedValue, "
sql = sql & " Currency_rate, Currency_id,sumVatLine,DueDate,"
 sql = sql & " TransactionComment ,ExtraAccount,ExtraValue)"
 sql = sql & " VALUES("
sql = sql & " " & Transaction_ID & " ,"
sql = sql & " " & BranchID & " ,"
sql = sql & "'" & NoteSerial & "' ,"
sql = sql & "'" & NoteSerial1 & "' ,"
sql = sql & " " & BoxID & " ,"
sql = sql & "'" & Transaction_serial & "',"
sql = sql & " " & SQLDate(Transaction_Date, True) & " ,"
sql = sql & " " & Transaction_Type & " ,"
sql = sql & " 0 ,"
sql = sql & " " & user_id & " ,"
sql = sql & " 0 ,"
sql = sql & " " & CusID & " ,"
sql = sql & " " & StoreID & " ,"
sql = sql & " " & 1 & " ,"
sql = sql & " " & Emp_id & " ,"
sql = sql & " " & val(txtRequiredAmount) & " ,"
'Vat
sql = sql & " " & val(TxtVAT) & " ,"
sql = sql & " " & val(txtTotalAfterVat) & " ,"
sql = sql & " " & val(txtTotalAfterVat) & " ,"
sql = sql & " " & 1 & " ,"
sql = sql & " " & 1 & " ,0,"
sql = sql & " " & SQLDate(Transaction_Date, True) & " ,"
sql = sql & "'" & TransactionComment & "', '" & Trim(Account_Code_dynamic213) & "'," & val(txtTotalAdd) & " )"
 
Cn.Execute sql
 
 
            

 
Dim RSTransDetails As New ADODB.Recordset
     
StrSQL = "SELECT     dbo.Transaction_Details.* from dbo.Transaction_Details Where (Transaction_ID = -1)"
RSTransDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   
       Dim rs2  As ADODB.Recordset
    
    Dim mItemNo As Long, mUnitNo As Long, mQty As Long, mVAt2 As Double, mTotal As Double
    Dim mwidtj As Double, mhight As Double, mTotalAdd As Double, mTotalDisc As Double, mNet As Double, mTotalWithVat As Double, mLength As Double
    Dim mItemName2 As String
    Dim mCost As Double
    Dim mRemark As String
   Dim UnitID As Integer
    Dim UnitID2 As Long
    Dim UnitName As String
    Dim i As Long
    For i = 1 To FG.rows - 1
        mItemNo = val(FG.TextMatrix(i, FG.ColIndex("ItemID")))
        mRemark = Trim(FG.TextMatrix(i, FG.ColIndex("RemarkItem")))
        GetDefaultItemUnit val(mItemNo), UnitID2, UnitName
    
    
    
        If mItemNo <> 0 Then
        
               
            UnitID = UnitID2
            mUnitNo = UnitID2
            mQty = 1
            mPrice = val(FG.TextMatrix(i, FG.ColIndex("Amount0")))
            'val(txtGeneralTotal)
            'mCost = val(.TextMatrix(i, .ColIndex("Cost")))
            
            mTotal = val(mPrice)
          '  mRemark = ""
            mTotalDisc = val(txtTotalDisc)
            mTotalAdd = val(txtTotalAdd)
            mNet = val(txtTotalAfterVat)
            
            mTotalWithVat = val(val(txtTotalAfterVat))
            
        
                
            RSTransDetails.AddNew
            RSTransDetails("Transaction_ID").value = Transaction_ID
            RSTransDetails!SavedItemType = 0
            RSTransDetails("ColorID").value = 1
            RSTransDetails("ItemSize").value = 1
            RSTransDetails("ClassId").value = 1
            RSTransDetails("Item_ID").value = mItemNo
            RSTransDetails("UnitID").value = mUnitNo
            RSTransDetails("SHOWQTY").value = mQty
            RSTransDetails("showPrice").value = mPrice
            RSTransDetails("Vat").value = val(TxtVAT)
            If SystemOptions.PriceWithVAT = True Then
                Percetage = 0
                RSTransDetails("TypeVAT").value = 0
                
                RSTransDetails("Vatyo").value = 0
            Else
                RSTransDetails("TypeVAT").value = Percetage
                
                RSTransDetails("Vatyo").value = val(Percetage)
            End If
            RSTransDetails("Remarks").value = IIf(mRemark <> "", " " & mRemark, "")
        
        'FG.TextMatrix(Num, FG.ColIndex("Vat")) = IIf(IsNull(RsDetails("Vat")), "", (RsDetails("Vat").value))
                      
                'RSTransDetails("NoCount").value = IIf((Fg.TextMatrix(RowNum, Fg.ColIndex("NoCount")) = ""), Null, val(Fg.TextMatrix(RowNum, Fg.ColIndex("NoCount"))))
                RSTransDetails("ItemDiscountType").value = 2
                RSTransDetails("ItemDiscount").value = val(txtTotalDisc)
                
                  RSTransDetails("CostPrice").value = mCost
                  If mCost = 0 Then
                        If SystemOptions.TypicalProduction = False Then
              
                            RSTransDetails("CostPrice").value = ModItemCostPrice.GetCostItemPrice(mItemNo, 0, , , SystemOptions.SysMainStockCostMethod, , , XPDtbTrans(mIndex).value, val(Me.Text1.text), RSTransDetails("UnitID").value, StoreID)
            
                            If RSTransDetails("CostPrice").value = 0 Then
                                RSTransDetails("CostPrice").value = ModItemCostPrice.GetCostItemPrice(mItemNo, 0, , , LastPurPriceType, , , XPDtbTrans(mIndex).value, val(Me.Text1.text), RSTransDetails("UnitID").value, StoreID)
                                
                            End If
                              
                        Else
                            RSTransDetails("CostPrice").value = 0
                        
                        End If
                    End If
                      
                                  '«·ÊÕœ« 
                   
                    Dim RsUnitData As ADODB.Recordset
                    Dim LngCurItemID As Long
                    Dim LngUnitID As Long
                    Dim DblQty As Double
                
                    LngCurItemID = val(mItemNo)
                    LngUnitID = val(mUnitNo)
                    DblQty = val(mQty)
        
                    StrSQL = "Select * From TblItemsUnits Where ItemID=" & LngCurItemID
                    StrSQL = StrSQL + " AND UnitID=" & LngUnitID
                    Set RsUnitData = New ADODB.Recordset
                    RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        
                    If Not (RsUnitData.BOF Or RsUnitData.EOF) Then
                        RSTransDetails("QtyBySmalltUnit").value = RsUnitData("UnitFactor").value
                        RSTransDetails("Quantity").value = RSTransDetails("QtyBySmalltUnit").value * RSTransDetails("showqty").value
                        RSTransDetails("OpeningSalesQty").value = RSTransDetails("Quantity").value
                        RSTransDetails("OpeningSalesValue").value = RSTransDetails("CostPrice").value
                        RSTransDetails("Price").value = val(IIf((mPrice = 0), 0, val(mPrice))) / RSTransDetails("QtyBySmalltUnit").value
                    
                    End If
        
                
                     UpdateTransactionsCost CStr(Transaction_ID)
                     RSTransDetails.update
    
      '  Dim i As Integer
        'Dim sql As String
        
        Set rs2 = New ADODB.Recordset
        
        sql = "Select * from  TransactionValueAdded where 1=-1"
        rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
        If val(LngCurItemID) <> 0 And SystemOptions.PriceWithVAT = False Then
            rs2.AddNew
            rs2("Transaction_ID").value = val(Transaction_ID)
            rs2("Transaction_Type").value = 21
            rs2("ItemID").value = LngCurItemID
            rs2("Vatyo").value = Percetage
            rs2("Vat").value = val(TxtVAT)
            rs2("Valu").value = val(mTotal) + val(mTotalAdd)
            rs2("selectd").value = 1
        
        End If
        If SystemOptions.PriceWithVAT = False Then
            rs2.update
        End If
    End If
Next
        NoteSerial = Notes_coding(val(BranchID), Transaction_Date)
        
        
        CreateNotes NoteID, Transaction_Date, CInt(BranchID), 170, val(txtTotalAfterVat), NoteSerial, NoteSerial1, "Transactions", "Transaction_ID", Transaction_ID, " »‰«¡« ⁄·Ï «„— ‘€· —ﬁ„ " & TxtSerial1(mIndex), ToHijriDate(Transaction_Date)
        txtNoteid3 = NoteID

'***********************

'***********************
        Dim cnt As Double
        Dim usedaccount As Integer
        Dim ItemsGoodsTotalsnew As Variant
        cnt = 1
        PG IIf(IsNull(RSTransDetails("quantity").value), 0, RSTransDetails("quantity").value), cnt, usedaccount, ItemsGoodsTotalsnew
        
        If val(txtPaymedValue) <> 0 Then
            If Not CreateCash Then GoTo ErrTrap
        End If
 
        'StrSQL = "UPDATE Transactions SET NOTS=" & Transaction_ID & " WHERE Transaction_ID=" & val(Me.XPTxtBillID.text)
        'Cn.Execute StrSQL
  
          If SystemOptions.UserInterface = ArabicInterface Then
              MsgBox " „ «‰‘«¡ ›« Ê—… „»Ì⁄« "
          Else
              MsgBox "Sales Invoice created"
          End If
        
'******************************************************issueVoucher



'Load frmsalebill
'frmsalebill.TxtModFlg.Text = "R"
'frmsalebill.mFormName = Me.Name
'frmsalebill.XPBtnMove_Click 2

        If Transaction_ID <> 0 Then
            createVoucher BranchID, 0, XPDtbTrans(mIndex).value, 19, 0, val(user_id), 0, DcCustmer.BoundText, StoreID, 0, 0, "”‰œ  ’—› »‰«¡ ⁄·Ì «„— ‘€· —ﬁ„ " & TxtSerial1(mIndex)
           
            'frmsalebill.Retrive Transaction_ID
        End If
        
        SaveQRCode "transactions", "Transaction_ID", val(Transaction_ID), NoteSerial1, (Transaction_Date), _
        (txtTotalNet.text), Picture1, 0, (TxtVAT.text), (txtTotalNet.text)


'frmsalebill.CreateIssueVoucher
'Unload frmsalebill



Exit Sub
ErrTrap:
Cn.RollbackTrans
    End Sub
     
     
Private Function CreateCash() As Boolean
CreateCash = False
         Dim rsCash As New ADODB.Recordset
         Dim StrSQL As String
    'StrSQL = "select * From Notes where NoteType=4 and   displayed is null Order By NoteID"
    StrSQL = "select * From Notes where NoteType=-1"
'StrSQL = StrSQL & " and CashingType<=11 and akarid is Null"

    'If SystemOptions.usertype <> UserAdminAll Then
    '    StrSQL = StrSQL & " AND   branch_no=" & Current_branch
    'End If
   On Error GoTo Err

    rsCash.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText


        If TxtModFlg2(mIndex).text = "N" Then
            txtNoteSerialCash(1).text = CStr(new_id("Notes", "NoteID", "", True))
            'Me.TxtNoteSerial.text = CStr(new_id("Notes", "NoteSerial", "", True, "NoteType=4"))
            rsCash.AddNew
       
            rsCash("NoteID").value = val(txtNoteSerialCash(1).text)
            'Me.oldtxtNoteSerial1.Text = Trim$(Me.TxtNoteSerial1.Text)
         
        ElseIf TxtModFlg2(mIndex).text = "E" Then
    
               txtNoteSerialCash(1).text = CStr(new_id("Notes", "NoteID", "", True))
            'Me.TxtNoteSerial.text = CStr(new_id("Notes", "NoteSerial", "", True, "NoteType=4"))
            rsCash.AddNew
       
            rsCash("NoteID").value = val(txtNoteSerialCash(1).text)
            
         End If


            Dim Current_case As Integer, s As String, mBoxID As Long
            Dim rsOut As New ADODB.Recordset
            s = "Select BoxID From TblBoxesData Where Empid = " & mEmpId



            rsOut.Open s, Cn, adOpenStatic, adLockReadOnly
            If Not rsOut.EOF Then
                mBoxID = val(rsOut!BoxID & "")
            End If
            If mBoxID = 0 Then
                rsOut.Close
                
                s = " SELECT tu.BoxID FROM TblUsers AS tu where UserId = " & user_id
                rsOut.Open s, Cn, adOpenStatic, adLockReadOnly
                If Not rsOut.EOF Then
                    mBoxID = val(rsOut!BoxID & "")
                End If
            End If
        If mBoxID = 0 Then
            MsgBox "ÌÃ»  ”ÃÌ· Œ“Ì‰… ··„” Œœ„ «Ê ··»«∆⁄"
            Exit Function
        End If

        rsCash("branch_no").value = val(Me.dcBranch(mIndex).BoundText)
        rsCash("EmpId").value = mEmpId
        'rsCash("foxy_no").value = val(Text1.Text)
        'rsCash("general_cost_center").value = IIf(Me.DcCostCenter.BoundText = "", "", Me.DcCostCenter.BoundText)
        'rsCash("Prefix").value = IIf(DCPreFix.Text = "", Null, DCPreFix.Text)

        'rsCash("CarId").value = IIf(Me.Dccar.BoundText = "", Null, (Me.Dccar.BoundText))
        'rsCash("DriverId").value = IIf(Me.DCDriver.BoundText = "", Null, (Me.DCDriver.BoundText))
    
        If val(txtNoteSerialCash(0).text) = 0 Then
            txtNoteSerialCash(0).text = Voucher_coding(val(dcBranch(mIndex).BoundText), XPDtbTrans(mIndex).value, 2, 4, , , "")
        End If
        Dim mNoteSerial As String
        If dcBranch(mIndex).BoundText = 0 Then dcBranch(mIndex).BoundText = mBranchID
            mNoteSerial = Notes_coding(val(dcBranch(mIndex).BoundText), XPDtbTrans(mIndex).value)
       
        
'        If CboStatus.ListIndex <> 0 Then
'        TxtNoteSerial.Text = ""
'
'        End If
       
    'If Option1.value = True Then
  '     rsCash("NCashingType").value = 1
   'ElseIf optIsEmp.value = True Then
   '     rsCash("NCashingType").value = 2
   'ElseIf optCash.value = True Then
   '     rsCash("NCashingType").value = 3
   '    ElseIf Option7.value = True Then
   '     rsCash("NCashingType").value = 7
        
   ' Else
    
         rsCash("NCashingType").value = 0
  ' End If
       
    
        'rsCash("ContainerNo").value = IIf(Trim(Me.txtContainerNo.Text) = "", Null, Trim(Me.txtContainerNo.Text))
        'rsCash("ManulaNO").value = IIf(Trim(Me.TxtManulaNO.Text) = "", Null, Trim(Me.TxtManulaNO.Text))
        'rsCash("ManualNo").value = IIf(Trim(Me.TxtManulaNO.Text) = "", Null, Trim(Me.TxtManulaNO.Text))
        'rsCash("BookNo").value = IIf(Trim(Me.TxtBookNo.Text) = "", Null, Trim(Me.TxtBookNo.Text))
        
        '
        rsCash("NoteSerial").value = mNoteSerial
        rsCash("NoteSerial1").value = IIf(Trim(Me.txtNoteSerialCash(0).text) = "", Null, Trim(Me.txtNoteSerialCash(0).text))
        'rsCash("OldNoteSerial1").value = Trim$(Me.oldtxtNoteSerial1.Text) '
        rsCash("NCashingType").value = 2
    
        'rsCash("person").value = IIf(TXTperson.Text = "", "", Trim(TXTperson.Text))
        rsCash("Note_Value").value = IIf(txtPaymedValue.text = "", Null, val(txtPaymedValue.text))
        'rsCash("Adv_payment_value").value = IIf(txtAdv_payment_value.Text = "", Null, val(txtAdv_payment_value.Text))
        'rsCash("VAT").value = IIf(TxtVATValue.Text = "", Null, val(TxtVATValue.Text))
    
        '    Rs("Remark").value = IIf(dcproject.BoundText = "", "", Trim(dcproject.BoundText))
        'If lblinvoices.Caption = "" Then
        rsCash("Remark").value = "”‰œ ﬁ»÷ ¬·Ï „‰ ›« Ê—… „»Ì⁄«  —ﬁ„" & TxtNoteSerial13
        'Else
        'rsCash("Remark").value = IIf(XPMTxtRemarks.Text = "", "", Trim(XPMTxtRemarks.Text)) & vbEnter & lblinvoices.Caption
        'End If
        
        'rsCash("BankName").value = IIf(TXTBankName.Text = "", "", Trim(TXTBankName.Text))
        rsCash("NoteType").value = 4
        rsCash("NoteDate").value = XPDtbTrans(mIndex).value
        rsCash("BillTransNo").value = TxtNoteSerial13.text
        rsCash("BillTransID").value = val(TXTTransactionID3.text)
        rsCash("Transaction_ID").value = val(TXTTransactionID3.text)
        
        'rsCash("BillMaintNo").value = TxtBillMaintNo.Text
        'rsCash("BillMaintID").value = val(TxtBillMaintID.Text)
        'rsCash("NoteDate").value = Format$(Date, "dd-mm-yyyy")
        'rsCash("NoteDateH").value = Me.Txt_DateHigri.value


        rsCash("CashingType").value = 0
        
        '
        rsCash("TotalNotesValue").value = 0
        
        rsCash("CurrentBalance").value = val(txtPaymedValue)
        rsCash("PaymentValue").value = val(txtPaymedValue)
        'rsCash("Percentage").value = val(TxtPercentage.Text)
        'rsCash("PercentageValue").value = val(TxtPercentageValue.Text)
        
        
        rsCash("CusID").value = IIf(DcCustmer.text = "", Null, DcCustmer.BoundText)
     
       

        '--------------------------------------------------------------------------
        'ÿ—Ìﬁ… «·œ›⁄ «·‰ﬁœÏ «Ê «·‘Ìﬂ
        
        rsCash("NoteCashingType").value = 0
        rsCash("BoxID").value = mBoxID
        rsCash("BankID").value = Null
        rsCash("ChqueNum").value = Null
        rsCash("DueDate").value = Null
    
       

        '--------------------------------------------------------------------------
        rsCash("UserID").value = user_id
        rsCash("numbering_type").value = sand_numbering_type(0)   '”‰œ «·ﬁÌœ
        rsCash("numbering_type1").value = sand_numbering_type(2) '”‰œ «·ﬁ»÷
    
      
    
     '  If DCboCashType.ListIndex = 8 Then
     '       rsCash("ContractNo").value = IIf(TxtContractNo.Text = "", Null, TxtContractNo.Text)
     '       rsCash("ContNo").value = IIf(TXTContNo.Text = "", Null, TXTContNo.Text)
     '       Else
     '        rsCash("ContractNo").value = Null
     '        rsCash("ContNo").value = Null
     '   End If
        
        
   '  If DCboCashType.ListIndex = 9 Then
   ' rsCash("akarid").value = IIf(val(Me.DcbIqara.BoundText) <> 0, val(DcbIqara.BoundText), Null)
   '  rsCash.Fields("UnitType").value = IIf(Me.DcbUnitType.BoundText <> "", val(DcbUnitType.BoundText), Null)
   '  rsCash.Fields("UnitNo").value = IIf(Me.DcbUnitNo.BoundText <> "", val(DcbUnitNo.BoundText), Null)
  '   rsCash("interval").value = IIf(txtinterval.Text = "", Null, val(txtinterval.Text))
  '   rsCash("intervaltype").value = val(cbointervaltype.ListIndex)
  '   rsCash("renterName").value = IIf(txtrenterName.Text = "", Null, txtrenterName.Text)
  '            If cbointervaltype.ListIndex = 0 Then
  '            rsCash("allowdate").value = DateAdd("d", val(txtinterval), XPDtbTrans.value)
  '            ElseIf cbointervaltype.ListIndex = 1 Then
  '            rsCash("allowdate").value = DateAdd("M", val(txtinterval), XPDtbTrans.value)
  '
  '          ElseIf cbointervaltype.ListIndex = 2 Then
  '            rsCash("allowdate").value = DateAdd("YYYY", val(txtinterval), XPDtbTrans.value)
  '
  '           End If
  '                rsCash("allowdateH").value = ToHijriDate(rsCash("allowdate").value)
  '
  '          Else
  '        rsCash("akarid").value = Null
  '   rsCash.Fields("UnitType").value = Null
  '   rsCash.Fields("UnitNo").value = Null
  '   rsCash("interval").value = Null
  '   rsCash("intervaltype").value = Null
  '   rsCash("renterName").value = Null
          
  '      End If
              
              
              
        
        rsCash("sanad_year").value = year(XPDtbTrans(mIndex).value)
        rsCash("sanad_month").value = Month(XPDtbTrans(mIndex).value)
    
       
        rsCash("note_value_by_characters").value = Trim$(val(txtPaymedValue))
       

        
            rsCash("cus_or_sub").value = 0 '⁄„Ì· ‰Â«∆Ì
       
    
        rsCash.update
saveBillBuy2

CmdCreateV2_Click
s = "Update Transactions Set PayedValue2 =" & val(txtPaymedValue) & " , StillValue =" & val(txtTotalAfterVat) - val(txtPaymedValue) & " , NoteIDCash = " & val(Me.txtNoteSerialCash(1).text) & ",NoteSerialCash = '" & Trim(Me.txtNoteSerialCash(0).text) & "' Where Transaction_ID = " & val(val(TXTTransactionID3.text))
            
    
                    
Cn.Execute s
CreateCash = True
Exit Function
Err:
CreateCash = False
End Function



Function saveBillBuy2()
    Dim StrSQL As String
   ' Dim StrSQL  As String
    Dim i As Integer
    Dim Diff As Double
    Dim Note_Value1 As Double
    Diff = 0
Dim RsDetails As ADODB.Recordset
    
    StrSQL = "Delete From TblNotesBillBuyPayment2 Where NoteID1=" & val(Me.txtNoteSerialCash(1).text) & " and TransType is null"
    Cn.Execute StrSQL, , adExecuteNoRecords
        StrSQL = "Delete From TblBillBuyPayment2 Where TypTrans IS NULL and  NoteID=" & val(Me.txtNoteSerialCash(1).text) & " and TransType is null"
    Cn.Execute StrSQL, , adExecuteNoRecords
   
    Dim mTotal As Double
    mTotal = val(txtTotalAfterVat)
    Set RsDetails = New ADODB.Recordset
   ' RsDetails.Open "TblEmpAdvanceDetails", Cn, adOpenKeyset, adLockOptimistic, adCmdTable

    StrSQL = "SELECT     * from dbo.TblNotesBillBuyPayment2 Where (1 = -1)"
   RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    
  
    'TxtValueTemp.Text = val(XPTxtVal.Text)
    
            RsDetails.AddNew
           ' val (Me.txtNoteSerialCash(1).Text)
            
            RsDetails("NoteID1").value = val(Me.txtNoteSerialCash(1).text)
            RsDetails("NoteID").value = val(TXTTransactionID3.text)
            RsDetails("branch_no").value = val(dcBranch(mIndex).BoundText)
            RsDetails("NoteSerial1").value = val(TxtNoteSerial13)
            RsDetails("Note_Value").value = val(mTotal)
            Note_Value1 = val(val(txtTotalNet) - val(txtPaymedValue))
            Diff = 0
'            If val(TxtValueTemp.Text) > 0 Then
'          If val(TxtValueTemp.Text) <= Note_Value1 Then
'          Diff = val(TxtValueTemp.Text)
'          TxtValueTemp.Text = val(TxtValueTemp.Text) - Note_Value1
'          Else
'          Diff = Note_Value1
'          TxtValueTemp.Text = val(TxtValueTemp.Text) - Note_Value1
'          End If
'            End If
          ' .TextMatrix(i, .ColIndex("RemainingValue")) = val(.TextMatrix(i, .ColIndex("RemainingValue"))) - val(.TextMatrix(i, .ColIndex("RemainingValue")))
            '.TextMatrix(i, .ColIndex("TransPayedValue")) = Diff
            
            'RsDetails("PayedValue").value = val(XPTxtValue(3)) ' val(.TextMatrix(i, .ColIndex("PayedValue")))
            
            'RsDetails("too").value = (.TextMatrix(i, .ColIndex("too")))
            RsDetails("NoteDate").value = XPDtbTrans(mIndex).value
           
            RsDetails("DueDate").value = Null
          
            RsDetails("TransPayedValue").value = val(txtPaymedValue)
           '.TextMatrix(i, .ColIndex("NetValue")) = val(XPTxtValue(3))
            RsDetails("NetValue").value = val(txtTotalNet) - val(txtPaymedValue)
            RsDetails("RemainingValue").value = val(mTotal)
            RsDetails.update
                
            If val(txtTotalNet) - val(txtPaymedValue) = 0 Then
                StrSQL = "Update Transactions Set  TotalPayed=1 Where Transaction_ID=" & val(TXTTransactionID3.text) & ""
                Cn.Execute StrSQL, , adExecuteNoRecords
             Else
                 StrSQL = "Update Transactions Set  TotalPayed=0 Where Transaction_ID=" & val(TXTTransactionID3.text) & ""
                Cn.Execute StrSQL, , adExecuteNoRecords
            End If
      

    Set RsDetails = New ADODB.Recordset
   ' RsDetails.Open "TblEmpAdvanceDetails", Cn, adOpenKeyset, adLockOptimistic, adCmdTable

    StrSQL = "SELECT     * from dbo.TblBillBuyPayment2 Where (1 = -1)"
   RsDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    

            RsDetails.AddNew
            RsDetails("NoteID").value = val(txtNoteSerialCash(1).text)
            RsDetails("RecDate").value = XPDtbTrans(mIndex).value
            RsDetails("Serial").value = txtNoteSerialCash(0).text
            RsDetails("Transaction_ID").value = val(TXTTransactionID3.text)
            RsDetails("Note_Value").value = val(mTotal)
            RsDetails("PayedValue").value = val(txtPaymedValue)
            RsDetails.update



End Function


Private Sub CmdCreateV2_Click()
Dim s As String
'CHECKaCCOUNTS

     


'END CHECK

If Not createVoucher2 Then Exit Sub
       'FindRec val(TXTLCNO.Text)
       
    

'Me.TxtModFlg2(mIndex).Text = "R"
End Sub
Function createVoucher2() As Boolean

'ee
Dim NoteID As Long
Dim NoteDate As Date
Dim NoteSerial As String
Dim Notevalue As Double
Dim notytype As Integer
Dim des As String
des = "    Õ”«» «·" '& TxtNoteSerial.Text


Dim tablename As String
Dim Filedname As String
Dim NoteSerial1 As Long
Dim BranchID As Integer
Dim sql As String
Dim mRate  As Double
tablename = "Notes"

Filedname = "NoteID"
'NoteSerial1 = CInt(val(txtNoteSerialCash(0).Text))

BranchID = val(dcBranch(mIndex).BoundText)
mRate = 1

'



notytype = 4
Notevalue = val(txtPaymedValue)

'mAccNO = val(DboParentAccount.BoundText)
NoteDate = (XPDtbTrans(mIndex).value)
 
If Notevalue > 0 Then
   

    If Not CREATE_VOUCHER_GE2(val(txtNoteSerialCash(1).text), BranchID, val(DCboUserName(mIndex).BoundText), NoteDate) Then createVoucher2 = False Else createVoucher2 = True
    RsSavRec.Resync adAffectCurrent

    updateNotesValueAndNobytext val(txtNoteSerialCash(0).text), Format(txtPaymedValue.text, "###.00")
'
'
'    StrSQL = "update  " & tablename & "   set NoteID=" & NoteID & ",NoteSerial='" & NoteSerial & "'"
'
'    StrSQL = StrSQL & " Where " & Filedname & " = " & NoteSerial1 & ""
'    Cn.Execute StrSQL
     
     
 
End If
End Function


Public Function CREATE_VOUCHER_GE2(general_noteid As Long, BranchID As Integer, UserID As Long _
, NoteDate As Date) As Boolean
Dim StrSQL As String

    Dim Current_case As Integer, s As String, mBoxID As Long
            Dim rsOut As New ADODB.Recordset
            s = "Select BoxID From TblBoxesData Where Empid = " & mEmpId



            rsOut.Open s, Cn, adOpenStatic, adLockReadOnly
            If Not rsOut.EOF Then
                mBoxID = val(rsOut!BoxID & "")
            End If
                        If mBoxID = 0 Then
                rsOut.Close
                
                s = " SELECT tu.BoxID FROM TblUsers AS tu where UserId = " & user_id
                rsOut.Open s, Cn, adOpenStatic, adLockReadOnly
                If Not rsOut.EOF Then
                    mBoxID = val(rsOut!BoxID & "")
                End If
            End If



'Dim StrAccountCodeDebt As String
Dim StrAccountCodeCridet As String
Dim StrAccountCodeDebt As String
StrAccountCodeDebt = GetMyAccountCode("TblBoxesData", "BoxID", mBoxID)   '«·„»Ì⁄« 
StrAccountCodeCridet = GetMyAccountCode("TblCustemers", "CusID", val(Me.DcCustmer.BoundText))

     StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & general_noteid
    Cn.Execute StrSQL, , adExecuteNoRecords
    
    Dim i As Integer
    Dim sql As String
    Dim StoreID6 As Integer
    Dim rs2 As ADODB.Recordset
    Set rs2 = New ADODB.Recordset
    Dim Notevalue As Double
    Dim LngDevID As Long
    Dim Msg As String
    'Dim StrAccountCodeDebt As String
    'Dim StrAccountCodeCridet As String
    Dim X As Integer
   
    Dim rs As New ADODB.Recordset
    Dim notes_serial As String
    Dim notes_id As String
    Msg = "    Õ”«» " & TxtNoteSerial13.text
    notes_id = general_noteid
   ' my_branch = val(Dcbranch(mIndex).BoundText)
    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
    Dim line_no As Integer
    line_no = 1
    
    'Dim s As String
    Dim mRate As Double
    mRate = 1
    ' „‰ Õ”«» «·⁄„Ì·
    
    

   
    Notevalue = val(txtPaymedValue.text)
    If Notevalue > 0 Then
        
       ' StrAccountCodeDebt = Trim(DboParentAccount.BoundText)
        If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCodeDebt, Notevalue, 0, Msg & "    Õ”«»  «·’‰œÊﬁ  ", val(notes_id), , , , NoteDate, val(DCboUserName(mIndex).BoundText), , , , , , CLng(mRate), , , setfoxy_Line, , , , , , , , , _
        val(dcBranch(mIndex).BoundText), , , , , , , , , , , , , , , , , , , , , , , , DcCustmer.BoundText) = False Then
            GoTo ErrTrap
        End If
       ' «·Ï Õ”«» «·ﬁÌ„… «·„÷«›…
      
        
        line_no = line_no + 1

    End If

    
    ' «·«ÿ—«›
    
     ' «·Ï Õ”«» «·⁄„Ì·
         
  '  Notevalue = val(txtTotal.Text)
    If Notevalue > 0 Then
    
              

        
        
 
        
        If ModAccounts.AddNewDev(LngDevID, line_no, StrAccountCodeCridet, Notevalue, 1, Msg & "    Õ”«» «·⁄„Ì·  ", val(notes_id), , , , NoteDate, val(DCboUserName(mIndex).BoundText), , , , , , CLng(mRate), , , setfoxy_Line, , , , , , , , , _
        val(dcBranch(mIndex).BoundText)) = False Then
            GoTo ErrTrap
        End If

        line_no = line_no + 1
    End If
    

    updateNotesValueAndNobytext (val(notes_id))
    CREATE_VOUCHER_GE2 = True
    Exit Function
ErrTrap:
CREATE_VOUCHER_GE2 = False
txtNoteSerialCash(1) = ""
txtNoteSerialCash(0) = ""

 

     
 
    '
 




 

End Function



Sub PG(Optional Qty As Double, Optional cnt As Double, Optional usedaccount As Integer, Optional ItemsGoodsTotalsnew As Variant, Optional ItemsServiceTotalsnew As Variant)
    Dim i As Integer
    Dim LngDevID As Long
    Dim LngDevNO  As Integer
    Dim StrTempAccountCode As String
    Dim StrTempDes As String
    Dim Account_Code_dynamic As String
    Dim SngTemp As Variant
    Dim TotalValue As Double
    On Error GoTo ErrTrap
    Dim TepAccount As String
    Dim OtherInformation As New ClsGLOther
    Dim general_noteid As Long
    Dim mBoxID As Long
    Dim txtAdvPay As Double, PercetageVat As Double
    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
    '----------------
    general_noteid = val(txtNoteid3)
    Dim mBoxAccount As String
    Dim mBoxID22 As Long
                Dim rsOut22 As ADODB.Recordset
            Set rsOut22 = New ADODB.Recordset
            s = "Select BoxID From TblBoxesData Where Empid = " & val(mEmpId)
            rsOut22.Open s, Cn, adOpenStatic, adLockReadOnly
            If Not rsOut22.EOF Then
                mBoxID22 = val(rsOut22!BoxID & "")
            End If
            mBoxAccount = GetMyAccountCode("TblBoxesData", "BoxID", mBoxID22)

    'SngTemp = NewGrid.GetItemsCostTotal * Qty / cnt


    Dim bankCommAccount As String
    Dim commision As Variant
   
    Dim Commisionvalue As Single
    Dim BankID As Long
    BankID = 0 ' GetPaymentTypeBank(val(Me.DCPaymentNet.BoundText))
    ' totalvalue = Val(Me.XPTxtValue(0).text) * Val(txt_Currency_rate.text)
   
    
    
    TotalValue = val(txtTotalAfterVat) '- val(txtTotalDisc)
   'TotalValue = Format((TotalValue), "#,###." & String(Abs(SystemOptions.Count_ACCOUNT_digit), "0"))


   Dim AdvancedAccount As String
   If SystemOptions.CustomerhavethreeAccounts = True Then
   AdvancedAccount = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DcCustmer.BoundText), "Account_code2")
   Else
   AdvancedAccount = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DcCustmer.BoundText), "Account_code")
   End If
   If AdvancedAccount = "" Then txtAdvPay = 0
   TepAccount = AdvancedAccount
  OtherInformation.NextAccount_Code = get_account_code_branch(2, val(dcBranch(mIndex).BoundText))
  'OtherInformation.NextAccount_Code = get_account_code_branch(149, VAL(Dcbranch(mIndex).BoundText ))
   Dim DebitAccountTemp As String
       'Dim AdvancedAccount As String
   If SystemOptions.CustomerhavethreeAccounts = True Then
        AdvancedAccount = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DcCustmer.BoundText), "Account_code2")
   Else
        AdvancedAccount = ModAccounts.GetMyAccountCode("TblCustemers", "CusID", val(Me.DcCustmer.BoundText), "Account_code")
   End If
   
'    If Me.CboPayMentType.ListIndex = 0 Then 'cash
'            mBoxID = val(DcboBox.BoundText)
'
'          '  mBoxID = 2
'            StrTempAccountCode = GetMyAccountCode("TblBoxesData", "BoxID", mBoxID)   '«·„»Ì⁄« 
'     Else
        StrTempAccountCode = AdvancedAccount
        StrTempAccountCode = GetMyAccountCode("TblCustemers", "CusID", val(Me.DcCustmer.BoundText))
    
    ' End If
        Dim maxvalue As Double
       
    
        If SystemOptions.UserInterface = ArabicInterface Then
            StrTempDes = "›« Ê—… »Ì⁄ —ﬁ„ " & Me.TxtNoteSerial13.text & " »‰«¡« ⁄·Ï «„— ‘€· —ﬁ„ " & TxtSerial1(mIndex)
        Else
            StrTempDes = "Sales Invoice NO: " & Me.TxtNoteSerial13.text & " »‰«¡« ⁄·Ï «„— ‘€· —ﬁ„ " & TxtSerial1(mIndex)
        End If

        LngDevNO = LngDevNO + 1
    Dim ValuGird As Double
   Dim StrMSG As String
   OtherInformation.NextAccount_Code = get_account_code_branch(2, val(dcBranch(mIndex).BoundText))
       If val(Me.DcCustmer.BoundText) = 2 Then
            Dim rsOut As ADODB.Recordset
            Set rsOut = New ADODB.Recordset
            s = "Select BoxID From TblBoxesData Where Empid = " & val(mEmpId)
            rsOut.Open s, Cn, adOpenStatic, adLockReadOnly
            If Not rsOut.EOF Then
                mBoxID = val(rsOut!BoxID & "")
            End If
            StrTempAccountCode = GetMyAccountCode("TblBoxesData", "BoxID", mBoxID)
        End If
        
        If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, TotalValue - val(txtAdvPay), 0, StrTempDes, general_noteid, , , , Me.XPDtbTrans(mIndex).value, user_id, val(TXTTransactionID3), , , , , , , , , , , , , , , , , val(Me.dcBranch(mIndex).BoundText), , , , , , , , , , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
            GoTo ErrTrap
        End If
        TepAccount = StrTempAccountCode
        DebitAccountTemp = StrTempAccountCode
            LngDevNO = LngDevNO + 1
            
            
     

       'End If
        DebitAccountTemp = StrTempAccountCode
  






    '«·œ«∆‰ ›Ì Õ«·… «·«’‰«›

    '  ÕœÌœ ÿ—Ìﬁ… —»ÿ «·„Œ«“‰ Ê «·Õ”«»«  ÊÂÌ ⁄·Ï „” ÊÏ «·›—⁄ Ê —»ÿ ⁄·Ï „” ÊÏ «·„Ã„Ê⁄«  Ê«·›—⁄ «Ê «·„Ã„Ê⁄«  Ê «·„Œ«“‰

    '1 work with branch
    '2 work with inventory
    '3 work with groups
    SngTemp = val(txtRequiredAmount)

    SngTemp = Round(SngTemp, SystemOptions.Count_ACCOUNT_digit)
'    TotalValue = Format((TotalValue), "#,###." & String(Abs(SystemOptions.Count_ACCOUNT_digit), "0"))
If SystemOptions.PriceWithVAT = True Then
SngTemp = SngTemp / 1.15
End If
    If SngTemp > 0 Then
        If detect_inventory_work_type = 1 Or detect_inventory_work_type = 2 Then
            Account_Code_dynamic = get_account_code_branch(2, val(dcBranch(mIndex).BoundText))
        
            If Account_Code_dynamic = "NO branch" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                Else
                    MsgBox "Branch Not Created", vbCritical
                End If

                GoTo ErrTrap
            Else

                If Account_Code_dynamic = "NO account" Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox "·„ Ì „  ÕœÌœ Õ”«»  «·„»Ì⁄«   ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                    Else
                        MsgBox "Sales Account Not Defined in this Branch", vbCritical
                    End If

                    GoTo ErrTrap
         
                End If
            End If

    
                StrTempAccountCode = Account_Code_dynamic '«·„»Ì⁄« 
   

OtherInformation.NextAccount_Code = TepAccount
            '           StrTempAccountCode = Account_Code_dynamic '«·„»Ì⁄« 
            'StrTempAccountCode = "a4a1" '«·„»Ì⁄« 
            If SystemOptions.UserInterface = ArabicInterface Then
                StrTempDes = "›« Ê—… »Ì⁄ —ﬁ„ " & Me.TxtNoteSerial13.text & " »‰«¡« ⁄·Ï «„— ‘€· —ﬁ„ " & TxtSerial1(mIndex)
            Else
                StrTempDes = "Sales Invoice NO: " & Me.TxtNoteSerial13.text & " »‰«¡« ⁄·Ï «„— ‘€· —ﬁ„ " & TxtSerial1(mIndex)
            End If

            LngDevNO = LngDevNO + 1

            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, 1, StrTempDes, general_noteid, , , , Me.XPDtbTrans(mIndex).value, user_id, val(TXTTransactionID3), , , , , , , , , , , , , , , , , val(Me.dcBranch(mIndex).BoundText), , , , , , , , , , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
                GoTo ErrTrap
            End If
            
            
  Dim value As Double
'  value = val(Me.txtTotalDisc)
'  If value > 0 Then
'        Account_Code_dynamic = get_account_code_branch(12, VAL(Dcbranch(mIndex).BoundText ))
'
'        If Account_Code_dynamic = "NO branch" Then
'            If SystemOptions.UserInterface = ArabicInterface Then
'                MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
'            Else
'                MsgBox "Branch Not Created ", vbCritical
'            End If
'
'            GoTo ErrTrap
'        Else
'
'            If Account_Code_dynamic = "NO account" Then
'                If SystemOptions.UserInterface = ArabicInterface Then
'                    MsgBox "·„ Ì „  ÕœÌœ Õ”«»    «·Œ’„ «·„”„ÊÕ »Â   ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
'                Else
'                    MsgBox "Allowance Discount Not Deined in this Branch", vbCritical
'                End If
'
'                GoTo ErrTrap
'
'            End If
'        End If
'
'
'        If val(Me.txtTotalDisc) > 0 Then
'         StrTempAccountCode = Account_Code_dynamic
'                If SystemOptions.DiscountSalesCreateVchr = True Then
'                 LngDevNO = LngDevNO + 1
'                       '     If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, val(Me.LblDiscountsTotal.Caption), 0, StrTempDes, , , , , Me.XPDtbTrans(mIndex).value, Me.DCboUserName(mIndex).BoundText, val(Me.XPTxtBillID.text)) = False Then
'                                            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, value, 0, StrTempDes, general_noteid, , , , Me.XPDtbTrans(mIndex).value, Me.DCboUserName(mIndex).BoundText, val(Transaction_ID), , , , , , , , , , , , , , , , , val(Me.Dcbranch(mIndex).BoundText), , , , , , , , , , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
'
'                                                GoTo ErrTrap
'                                            End If
'
'                                End If
'
'                End If
'
   ' End If



'Õ”«» «·«÷«›« 

    
        ElseIf detect_inventory_work_type = 3 Then
'
        End If

    End If
   
 Dim Account_Code_dynamic213  As String
 Account_Code_dynamic213 = get_account_code_branch(213, CStr(mBranchID))

    If Account_Code_dynamic213 <> "" And val(txtTotalAdd.text) > 0 Then
    
    
        LngDevNO = LngDevNO + 1

        If ModAccounts.AddNewDev(LngDevID, LngDevNO, mBoxAccount, val(txtTotalAdd.text), 0, StrTempDes, general_noteid, , , , Me.XPDtbTrans(3).value, user_id, val(TXTTransactionID3), , , , , , , , , , , , , , , , , val(mBranchID), , , , , , , , , , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
            GoTo ErrTrap
        End If
    
    
        If SystemOptions.UserInterface = ArabicInterface Then
            StrTempDes = "›« Ê—… »Ì⁄ —ﬁ„      " & Me.TxtNoteSerial13.text
        Else
            StrTempDes = "Sales Invoice No. " & Me.TxtNoteSerial13.text
        End If

        LngDevNO = LngDevNO + 1

        If ModAccounts.AddNewDev(LngDevID, LngDevNO, Account_Code_dynamic213, val(txtTotalAdd.text), 1, StrTempDes, general_noteid, , , , Me.XPDtbTrans(3).value, user_id, val(TXTTransactionID3), , , , , , , , , , , , , , , , , val(mBranchID), , , , , , , , , , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
            GoTo ErrTrap
        End If
    End If

    '
Dim mVat As Double
If SystemOptions.PriceWithVAT = True Then
    mVat = (TotalValue / 1.15) * 0.15
End If
mVat = val(TxtVAT)
        If val(mVat) > 0 Then
    Dim AccountVATCreit As String
 GetValueAddedAccount XPDtbTrans(mIndex).value, , AccountVATCreit, 1, 21


         If SystemOptions.UserInterface = ArabicInterface Then
                                StrTempDes = "  ﬁÌ„… „÷«›… »‰”»… " & txtVatYou & " %  " & "›« Ê—… »Ì⁄ —ﬁ„ " & Me.TxtNoteSerial13.text & " »‰«¡« ⁄·Ï ›« Ê—… „»Ì⁄«  —ﬁ„ " & TxtNoteSerial13
                            Else
                                StrTempDes = "VAT Sales Invoice NO: " & Me.TxtNoteSerial13.text
        End If
            
                            LngDevNO = LngDevNO + 1
        If ModAccounts.AddNewDev(LngDevID, LngDevNO, AccountVATCreit, val(mVat), 1, StrTempDes, general_noteid, , , , Me.XPDtbTrans(mIndex).value, user_id, val(TXTTransactionID3), , , , , , , , , , , , , , , , , val(Me.dcBranch(mIndex).BoundText), , , , , , , , , , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
            GoTo ErrTrap
        End If
        mVat = 0
     End If
     ''/////////////
     Dim Account_Code_dynamic82 As String

     ''//////////
'     If SystemOptions.DealingWithPrepayAccount = True Then
'      If val(TxtVAt2.Text) > 0 Then
'
'             GetValueAddedAccount XPDtbTrans(mIndex).value, , AccountVATCreit, 1, 21
'         If SystemOptions.UserInterface = ArabicInterface Then
'                                StrTempDes = "  ﬁÌ„… „÷«›… " & "›« Ê—… »Ì⁄ —ﬁ„ " & Me.TxtNoteSerial13.Text
'                            Else
'                                StrTempDes = "VAT ""Sales Invoice NO: " & Me.TxtNoteSerial13.Text
'        End If
'
'                            LngDevNO = LngDevNO + 1
'        If ModAccounts.AddNewDev(LngDevID, LngDevNO, AccountVATCreit, val(TxtVAt2.Text), 0, StrTempDes, general_noteid, , , , Me.XPDtbTrans(mIndex).value, Me.DCboUserName(mIndex).BoundText, val(Transaction_ID), , , , , , , , , , , , , , , , , val(Me.Dcbranch(mIndex).BoundText), , , , , , , , , , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
'            GoTo ErrTrap
'        End If
'                 If SystemOptions.UserInterface = ArabicInterface Then
'                                StrTempDes = "  Õ”«» «·⁄„Ì· " & "›« Ê—… »Ì⁄ —ﬁ„ " & Me.TxtNoteSerial13.Text
'                            Else
'                                StrTempDes = "Customer ""Sales Invoice NO: " & Me.TxtNoteSerial13.Text
'                 End If
'                  LngDevNO = LngDevNO + 1
'        AccountVATCreit = GetMyAccountCode("TblCustemers", "CusID", val(Me.DcCustmer.BoundText))
'             If ModAccounts.AddNewDev(LngDevID, LngDevNO, AccountVATCreit, val(TxtVAt2.Text), 1, StrTempDes, general_noteid, , , , Me.XPDtbTrans(mIndex).value, Me.DCboUserName(mIndex).BoundText, val(Transaction_ID), , , , , , , , , , , , , , , , , val(Me.Dcbranch(mIndex).BoundText), , , , , , , , , , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
'            GoTo ErrTrap
'        End If
'     End If
'     End If
   
xl:

'************************************************************************************


ErrTrap:
End Sub






Function FillMylist()

    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim Balance As Double
    Dim i As Integer


    sql = " SELECT * from  TblItems  Where GroupID In (Select Groups.GroupID from groups where BranchID = " & Current_branch & ")"
 
    If SystemOptions.UserInterface = ArabicInterface Then
        sql = sql & " order by  ItemName"
    Else
        sql = sql & " order by  ItemNamee"
    End If
 
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    ListGroupAll.Clear
    'ListGroupSelected.Clear

    If rs.RecordCount > 0 Then
        For i = 1 To rs.RecordCount
            If SystemOptions.UserInterface = ArabicInterface Then
                ListGroupAll.AddItem IIf(IsNull(rs("ItemName").value), "", rs("ItemName").value)
            Else
                ListGroupAll.AddItem IIf(IsNull(rs("ItemNamee").value), "", rs("ItemNamee").value)
            End If

            ListGroupAll.ItemData(ListGroupAll.NewIndex) = rs("ItemID").value
            rs.MoveNext
        Next i
    End If
    rs.Close
    
  

   
  sql = "select * from TblEmployee "
    ' sql = "select* from TblBoxesData where  "
   
   sql = sql & "   where  IsNull(chkShowTasks,1) = 1 and  branchid in"
sql = sql & "    ("
sql = sql & "    select branch_id from TblBranchesData where Beauty=1"
sql = sql & "    )"

   
    sql = sql & " order by  Emp_Namee"
    
 
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    ListProductLineAll.Clear
    
    If rs.RecordCount > 0 Then
        For i = 1 To rs.RecordCount
            ListProductLineAll.AddItem IIf(IsNull(rs("Fullcode").value), "", rs("Fullcode").value) & "-" & IIf(IsNull(rs("Emp_Namee").value), "", rs("Emp_Namee").value)

            ListProductLineAll.ItemData(ListProductLineAll.NewIndex) = rs("Emp_ID").value
            rs.MoveNext
        Next i
    End If
End Function




Function FillMylist2(Optional ByVal IsBranch As Boolean = True, Optional ByVal IsAqar As Boolean = True, Optional ByVal IsUnitType As Boolean = True, Optional ByVal IsUnitNo As Boolean = True)
  
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim Balance As Double
    Dim i As Integer
    Dim mWhere As String
    Dim ii As Long
    If IsBranch Then
        sql = " SELECT * from  TblBranchesData "
     
        If SystemOptions.UserInterface = ArabicInterface Then
            sql = sql & " order by  branch_name"
        Else
            sql = sql & " order by  branch_namee"
        End If
     
        rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    
        ListBranchAll.Clear
        'ListGroupSelected.Clear
    
        If rs.RecordCount > 0 Then
            For i = 1 To rs.RecordCount
                If SystemOptions.UserInterface = ArabicInterface Then
                    ListBranchAll.AddItem IIf(IsNull(rs("branch_name").value), "", rs("branch_name").value)
                Else
                    ListBranchAll.AddItem IIf(IsNull(rs("branch_namee").value), "", rs("branch_namee").value)
                End If
    
                ListBranchAll.ItemData(ListBranchAll.NewIndex) = rs("branch_id").value
                rs.MoveNext
            Next i
        End If
        rs.Close
    End If
  
    
    
 If IsAqar Then
      sql = " SELECT Aqarid,aqarname From TblAqar  "
        ' sql = "select* from TblBoxesData where  "
        sql = sql & " Where 1 = 1 "
        For ii = 0 To ListBranchSelected.ListCount - 1
             If val(ListBranchSelected.ItemData(ii)) <> 0 Then
                If mWhere = "" Then
                    mWhere = val(ListBranchSelected.ItemData(ii))
                Else
                    mWhere = mWhere & "," & val(ListBranchSelected.ItemData(ii))
                End If
                
             End If
        Next ii
        If mWhere <> "" Then
            sql = sql & " And BranchId In (" & mWhere & ")"
        End If
        sql = sql & " order by  aqarname"
        
     
        rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    
        ListAqarAll.Clear
        
        If rs.RecordCount > 0 Then
            For i = 1 To rs.RecordCount
                ListAqarAll.AddItem IIf(IsNull(rs("aqarname").value), "", rs("aqarname").value)
    
                ListAqarAll.ItemData(ListAqarAll.NewIndex) = rs("Aqarid").value
                rs.MoveNext
            Next i
        End If
        
       rs.Close
    End If
    
    If IsUnitType Then
      sql = " SELECT  id,name  From TblAkarUnit "
        ' sql = "select* from TblBoxesData where  "
       
        sql = sql & " order by  name"
        
     
        rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    
        ListUnitTypeAll.Clear
        
        If rs.RecordCount > 0 Then
            For i = 1 To rs.RecordCount
                ListUnitTypeAll.AddItem IIf(IsNull(rs("name").value), "", rs("name").value)
    
                ListUnitTypeAll.ItemData(ListUnitTypeAll.NewIndex) = rs("id").value
                rs.MoveNext
            Next i
        End If
        
    
        rs.Close
   End If
   mWhere = ""
   Dim mWhere2 As String
   If IsUnitNo Then
        sql = "  SELECT Id,unitno,Aqarid From TblAqarDetai "
        ' sql = "select* from TblBoxesData where  "
       
       
        sql = sql & " Where 1 = 1 "
        For ii = 0 To ListAqarSelected.ListCount - 1
             If val(ListAqarSelected.ItemData(ii)) <> 0 Then
                If mWhere = "" Then
                    mWhere = val(ListAqarSelected.ItemData(ii))
                Else
                    mWhere = mWhere & "," & val(ListAqarSelected.ItemData(ii))
                End If
                
             End If
        Next ii
        If mWhere <> "" Then
            sql = sql & " And Aqarid In (" & mWhere & ")"
        End If
        mWhere2 = ""
        For ii = 0 To ListUnitTypeSelected.ListCount - 1
             If val(ListUnitTypeSelected.ItemData(ii)) <> 0 Then
                If mWhere2 = "" Then
                    mWhere2 = val(ListUnitTypeSelected.ItemData(ii))
                Else
                    mWhere2 = mWhere2 & "," & val(ListUnitTypeSelected.ItemData(ii))
                End If
                
             End If
        Next ii
        If mWhere2 <> "" Then
            sql = sql & " And unittype In (" & mWhere2 & ")"
        End If
        
        
        sql = sql & " Order by  unitno"
        
     
        rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
    
        ListUnitNoAll.Clear
        ListUnitNoAll2.Clear
        If rs.RecordCount > 0 Then
            For i = 1 To rs.RecordCount
                ListUnitNoAll.AddItem IIf(IsNull(rs("unitno").value), "", rs("unitno").value)
    
                ListUnitNoAll.ItemData(ListUnitNoAll.NewIndex) = rs("id").value
                  
                ListUnitNoAll2.AddItem IIf(IsNull(rs("Aqarid").value), "", rs("Aqarid").value)
    
                ListUnitNoAll2.ItemData(ListUnitNoAll2.NewIndex) = rs("Aqarid").value
                  
                rs.MoveNext
            Next i
        End If
    End If
End Function



Public Sub GetCustomerNamebyPhone2(Optional ByVal phone As String = "", Optional ByVal Name As String = "", Optional ByVal CUSTID As String = "", Optional ByVal SearchCode As String = "")
            If phone = "" And Name = "" And CUSTID = "" And SearchCode = "" Then Exit Sub
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim Balance As Double

        If phone <> "" Then
            sql = "SELECT     Cus_mobile , CusName,CusID,Fullcode,cPaymentType,EmpId From dbo.TblCustemers  WHERE     (Cus_mobile = '" & phone & "')"
        ElseIf Name <> "" Then
            sql = "SELECT     Cus_mobile, CusName,CusID,Fullcode,cPaymentType,EmpId From dbo.TblCustemers  WHERE     (CusName = '" & Name & "')"
        ElseIf CUSTID <> "" Then
            sql = "SELECT     Cus_mobile, CusName,CusID,Fullcode,cPaymentType,EmpId From dbo.TblCustemers  WHERE     (CusID = " & val(CUSTID) & ")"
        ElseIf SearchCode <> "" Then
            sql = "SELECT     Cus_mobile, CusName,CusID,Fullcode,cPaymentType,EmpId From dbo.TblCustemers  WHERE     Fullcode ='" & SearchCode & "'"
        Else
        Exit Sub
        End If
  
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then

        txtPhoneCust = rs!Cus_mobile & ""
     '   TxtSearchCode2.Text = rs!Fullcode & ""
        
        DBCboClientName.BoundText = val(rs!CusID & "")
        'DcboEmp.BoundText = val(rs!empid & "")
        txtCustName.text = IIf(IsNull(rs!CusName), "", rs!CusName)
        If SystemOptions.DontShowMoreDetailsCompItem Then
            CboPayMentType.ListIndex = IIf(IsNull(rs("cPaymentType").value), 0, rs("cPaymentType").value)
        End If
    Else
         txtPhoneCust = "123456789"
         TxtSearchCode2 = ""
         DBCboClientName.BoundText = ""
          txtCustName.text = ""
              If Me.TxtModFlg2(mIndex) <> "R" Then

        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "Â–« «·⁄„Ì· €Ì— „ÊÃÊœ", vbCritical
        Else
            MsgBox "This client does not exist", vbCritical
        End If
End If
    End If

    rs.Close

End Sub

Public Sub GetCustomerNamebyPhone(Optional ByVal phone As String = "", Optional ByVal Name As String = "", Optional ByVal CUSTID As String = "", Optional ByVal SearchCode As String = "")
            If phone = "" And Name = "" And CUSTID = "" And SearchCode = "" Then Exit Sub
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim Balance As Double

        If phone <> "" Then
            sql = "SELECT     Cus_mobile , CusName,CusID,Fullcode,cPaymentType,EmpId From dbo.TblCustemers  WHERE     (Cus_mobile = '" & phone & "')"
        ElseIf Name <> "" Then
            sql = "SELECT     Cus_mobile, CusName,CusID,Fullcode,cPaymentType,EmpId From dbo.TblCustemers  WHERE     (CusName = '" & Name & "')"
        ElseIf CUSTID <> "" Then
            sql = "SELECT     Cus_mobile, CusName,CusID,Fullcode,cPaymentType,EmpId From dbo.TblCustemers  WHERE     (CusID = " & val(CUSTID) & ")"
        ElseIf SearchCode <> "" Then
            sql = "SELECT     Cus_mobile, CusName,CusID,Fullcode,cPaymentType,EmpId From dbo.TblCustemers  WHERE     Fullcode ='" & SearchCode & "'"
        Else
        Exit Sub
        End If
          
        
        
        
        
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then

        TxtPhone = rs!Cus_mobile & ""
        
        TxtSearchCode.text = rs!fullcode & ""
        DcCustmer.BoundText = val(rs!CusID & "")
        
        txtCustomerName.text = IIf(IsNull(rs!CusName), "", rs!CusName)
'        If SystemOptions.DontShowMoreDetailsCompItem Then
'            CboPayMentType.ListIndex = IIf(IsNull(rs("cPaymentType").value), 0, rs("cPaymentType").value)
'        End If
    Else
         TxtPhone = ""
         TxtSearchCode = ""
         DcCustmer.BoundText = ""
          txtCustomerName.text = ""
              If Me.TxtModFlg <> "R" Then

        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "Â–« «·⁄„Ì· €Ì— „ÊÃÊœ", vbCritical
        Else
            MsgBox "This client does not exist", vbCritical
        End If
End If
    End If

    rs.Close

End Sub



Private Sub createVoucher(BranchID As Double, _
BoxID As Double, _
Transaction_Date As Date, _
Transaction_Type As Double, _
CBoBasedON As Double, _
UserID As Double, _
Trans_DiscountType As Double, _
CusID As Double, _
StoreID As Double, _
PaymentType As Double, _
Emp_id As Double, _
TransactionComment As String, Optional invoice As Integer = 0)
Dim sql As String
Dim Msg As String
Dim NoteID As Long
Dim Transaction_ID As Long
Dim Transaction_ID1 As Long
Dim Transaction_serial As String
Dim NoteSerial As String
Dim NoteSerial1 As String
Dim CostAccount As String
 Dim CostTOTAL As Double
 Dim StoreAccount As String
 Dim costPrice As Double
'BillTOTAL = 0
CostTOTAL = 0
'Check
  'NoteSerial1 = Voucher_coding(val(BranchID), Transaction_Date, 10, 180, , 27)
    
'    If Transaction_Type = 27 Then
'         NoteSerial1 = Voucher_coding(val(BranchID), Transaction_Date, 18, 240, , CInt(Transaction_Type), , CDbl(StoreId))              '’—› „Ê«œ Œ«„
'    Else
        NoteSerial1 = Voucher_coding(val(BranchID), Transaction_Date, 7, 170, , CInt(Transaction_Type))    '’—› „Ê«œ Œ«„
 '   End If
                
        If NoteSerial1 = "" Then
                 If NoteSerial1 = "error" Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox " ·« Ì„ﬂ‰ «÷«›… ”‰œ   „Ê«œ Œ«„ ··«‰ «Ã  ÃœÌœ ·«‰ﬂ  ⁄œÌ  «·Õœ «·–Ì ﬁ„  » ÕœÌœ… „‰ «·”‰œ«   ": Exit Sub
                    Else
                        MsgBox " You can not add a raw material bond to a new production because you have exceeded the limit on which you have selected the bonds ": Exit Sub
                    End If
            
                 ElseIf NoteSerial1 = "" Then
                         MsgBox " ·«»œ „‰ ﬂ «»… —ﬁ„ «·”‰œ ÌœÊÌ« ﬂ„« Õœœ   ": Exit Sub
        
                 End If
        End If

NoteSerial = Notes_coding(val(BranchID), Transaction_Date)
 If NoteSerial = "" Then
            If NoteSerial = "error" Then
                MsgBox " ·« Ì„ﬂ‰ «÷«›… ﬁÌÊœ ·Â–… «·⁄„·Ì… ·«‰ﬂ  ⁄œÌ  «·Õœ «·«ﬁ’Ì ··ﬁÌÊœ ﬂ„« Õœœ  ›Ì  —ﬁÌ„ «·”‰œ«  ": Exit Sub
            ElseIf NoteSerial = "" Then
                    MsgBox " ·«Ì„ﬂ‰ «‰‘«¡ «·ﬁÌœ ·Â–« «·„” ‰œ ·«‰ﬂ Õœœ   —ﬁÌ„ ﬁÌÊœ ÌœÊÌ  ": Exit Sub
                 
            End If
End If
           
 
   If Trim(StoreID) = 0 Then
         MsgBox "·„ Ì „  ÕœÌœ Õ”«»  ··„Œ“Ê‰ «·”·⁄Ì ·Â–« «·„Œ“‰ ›Ì Â–« «·›—⁄    ", vbCritical
   End If
  
  
 
           'CostAccount = get_account_code_branch(137, CInt(BranchID))
           CostAccount = get_account_code_branch(1, CInt(BranchID))
        
            If CostAccount = "NO branch" Or CostAccount = "NO account" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·„ Ì „ —»ÿ  ﬂ·›…   «·„»Ì⁄«   ", vbCritical
                Else
                    MsgBox "Sales Not Created", vbCritical
                End If

             Exit Sub
              End If
              
              

    StoreAccount = get_store_Account(CInt(StoreID), "Account_Code")
      If StoreAccount = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·„ Ì „  ÕœÌœ Õ”«»  ··„Œ“Ê‰ «·”·⁄Ì ·Â–« «·„Œ“‰ ›Ì Â–« «·›—⁄    ", vbCritical
                Else
                    MsgBox "No inventory account for this store has been specified in this section", vbCritical
                End If
           Exit Sub
            End If
          Dim RsUnitData As ADODB.Recordset
            Dim LngCurItemID As Long
            Dim LngUnitID As Long
            Dim DblQty As Double

 'end Check

        
Transaction_ID = CStr(new_id("Transactions", "Transaction_ID", "", True))
Transaction_serial = NoteSerial1

        
TXTTransactionID1 = Transaction_ID
TxtNoteSerial11 = NoteSerial
        
Dim mCust As Long
Dim rsDummyChkCust As New ADODB.Recordset
sql = "Select * from TblCustemers Where CusId = " & CusID

rsDummyChkCust.Open sql, Cn, adOpenStatic, adLockReadOnly
If rsDummyChkCust.EOF Then
    sql = "Select Top 1 CusId from TblCustemers "
    rsDummyChkCust.Close
    rsDummyChkCust.Open sql, Cn, adOpenStatic, adLockReadOnly
    CusID = val(rsDummyChkCust!CusID & "")
End If
        
 sql = "INSERT INTO  Transactions (  "
sql = sql & " Transaction_ID ,"
sql = sql & " BranchID ,"
sql = sql & " NoteSerial ,"
sql = sql & " NoteSerial1 ,"
sql = sql & " boxId ,"
sql = sql & " Transaction_serial ,"
sql = sql & " Transaction_Date ,"
sql = sql & " Transaction_Type ,"
sql = sql & " BillBasedOn ,"
sql = sql & " UserID ,"
sql = sql & " Trans_DiscountType ,"
sql = sql & " CusID ,"
sql = sql & " StoreId ,"
sql = sql & " PaymentType ,"
sql = sql & " Emp_id ,InvoiceOrderNo,"
 sql = sql & " TransactionComment )"
 
 sql = sql & " VALUES("
sql = sql & " " & Transaction_ID & " ,"
sql = sql & " " & BranchID & " ,"
sql = sql & "'" & NoteSerial & "' ,"
sql = sql & "'" & NoteSerial1 & "' ,"
sql = sql & " " & BoxID & " ,"
sql = sql & "'" & Transaction_serial & "',"
sql = sql & " " & SQLDate(Transaction_Date, True) & " ,"
sql = sql & " " & Transaction_Type & " ,"
sql = sql & " 2 ,"
sql = sql & " " & user_id & " ,"
sql = sql & " 0 ,"
sql = sql & " " & CusID & " ,"
sql = sql & " " & StoreID & " ,"
sql = sql & " 0 ,"
sql = sql & " " & Emp_id & " ," & val(TxtSerial1(mIndex)) & ","
 sql = sql & "'" & TransactionComment & "')"
 

         Cn.Execute sql
 


Dim mTotal As Double
mTotal = 0
 Dim i As Long
 Dim mItemId As Long
        Dim RSTransDetails As New ADODB.Recordset
  Dim StrSQL As String
StrSQL = "SELECT     dbo.Transaction_Details.* from dbo.Transaction_Details Where (Transaction_ID = -1)"
   RSTransDetails.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
   
   
   For i = 1 To FG.rows - 1
        mItemId = val(FG.TextMatrix(i, FG.ColIndex("ItemId")))
    If Transaction_Type = 19 Then
    
                Dim UnitID As Integer
                Dim UnitID2 As Long
                Dim UnitName As String
                
                GetDefaultItemUnit val(mItemId), UnitID2, UnitName
             
             

            If val(mItemId) <> 0 Then
                
                
                
                
                RSTransDetails.AddNew
                RSTransDetails("Transaction_ID").value = Transaction_ID
             
                RSTransDetails("ColorID").value = 1
                RSTransDetails("ItemSize").value = 1
                RSTransDetails("ClassId").value = 1
        RSTransDetails("Item_ID").value = val(mItemId)
                RSTransDetails("UnitID").value = UnitID2
               RSTransDetails("SHOWQTY").value = 1
               RSTransDetails("showPrice").value = val(txtGeneralTotal)
              
              

        
            LngCurItemID = val(mItemId)
            LngUnitID = UnitID2
            DblQty = 1
            costPrice = val(txtGeneralTotal)
       '     costPrice = ModItemCostPrice.GetCostItemPrice(CLng(LngCurItemID), 0, "", , SystemOptions.SysMainStockCostMethod, DblQty, , XPDtbBill, , LngUnitID)
  ' costPrice = ModItemCostPrice.GetCostItemPrice(CLng(LngCurItemID), 0, "", , SystemOptions.SysMainStockCostMethod, DblQty, , XPDtbBill, , LngUnitID)
 'costPrice = 20
  ' CostTOTAL = CostTOTAL + costPrice * DblQty
  
            ' FG2.TextMatrix(RowNum, FG2.ColIndex("cost")) = costPrice
                  
          'RSTransDetails("ShowPrice").value = costPrice
          RSTransDetails("showPrice").value = Round(costPrice)
         RSTransDetails("ShowQty").value = DblQty
                    
          

            StrSQL = "Select * From TblItemsUnits Where ItemID=" & LngCurItemID
            StrSQL = StrSQL + " AND UnitID=" & LngUnitID
            Set RsUnitData = New ADODB.Recordset
            RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        'fg2.TextMatrix(RowNum, fg2.ColIndex("Price")) = 0

            If Not RsUnitData.EOF Then
 
                RSTransDetails("QtyBySmalltUnit").value = IIf(IsNull(RsUnitData("UnitFactor").value), 1, RsUnitData("UnitFactor").value)
                RSTransDetails("Quantity").value = RSTransDetails("QtyBySmalltUnit").value * RSTransDetails("showqty").value
                  RSTransDetails("Price").value = Round(val(costPrice) / RSTransDetails("QtyBySmalltUnit").value, 3)
            
            End If
            RSTransDetails("CostPrice").value = costPrice
            
                   CostTOTAL = CostTOTAL + (val(Round(val(RSTransDetails("showPrice").value) / RSTransDetails("QtyBySmalltUnit").value, 3)) * DblQty)
            
                RSTransDetails.update
            End If
        End If
             UpdateTransactionsCost CStr(Transaction_ID)
    Next
'Exit Sub
 
NoteSerial = Notes_coding(val(BranchID), Transaction_Date)
 'If Transaction_Type = 27 Then
 '   CreateNotes NoteID, Transaction_Date, CInt(BranchID), 240, mTotal, NoteSerial, NoteSerial1, "Transactions", "Transaction_ID", Transaction_ID, " »‰«¡« ⁄·Ï ”‰œ  Ã„Ì⁄ —ﬁ„ " & TxtTransSerial, ToHijriDate(Transaction_Date)
'Else
    CreateNotes NoteID, Transaction_Date, CInt(BranchID), 180, mTotal, NoteSerial, NoteSerial1, "Transactions", "Transaction_ID", Transaction_ID, " »‰«¡« ⁄·Ï ›« Ê—… „»Ì⁄«  —ﬁ„ " & TxtNoteSerial13, ToHijriDate(Transaction_Date)
'End If

'TxtNoteSerial11
'***********************
        
            
    StrSQL = "UPDATE Transactions SET  Nots=" & val(TXTTransactionID3) & ",BillBasedOn =2,nots2 = '" & Trim(TxtNoteSerial13.text) & "',Closed = 1   WHERE Transaction_ID  =" & val(Transaction_ID)
    Cn.Execute StrSQL
        
       
'***********************

  CREATE_VOUCHER_GE1 Transaction_ID, NoteSerial1, "", NoteID, val(dcBranch(mIndex).BoundText), StoreID, Transaction_Date, 0, invoice
       
 
        'StrSQL = "UPDATE Transactions SET NOTS=" & Transaction_ID & " WHERE Transaction_ID=" & val(Me.XPTxtBillID.text)
        'Cn.Execute StrSQL
  'MsgBox " „   «·‰ﬁ·"
  
'******************************************************issueVoucher








     
 
    '
 
ErrTrap:

End Sub

 
 
Function CREATE_VOUCHER_GE1(Transaction_ID As Long, TxtNoteSerialV As String, TxtNoteSerial1V As String, general_noteid As Long, BranchID As Integer, StoreID As Double, Transaction_Date As Date, BoxID As Double, Optional invoice As Integer = 0)
    Dim LngDevID As Long
    Dim LngDevNO  As Integer
    Dim StrTempAccountCode As String
    Dim StrTempDes As String
    Dim Line1 As Double
    Dim Line2 As Double
    Dim OtherInformation As New ClsGLOther
    Dim DebitAccount As String
    Dim rsDummy As New ADODB.Recordset
    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
    '----------------
    Dim Account_Code_dynamic As String
    Dim CostTOTAL As Double, SngTemp As Double
    Dim CreditAccount As String
    'SngTemp = NewGrid.GetItemsCostTotal * RSTransDetails("quantity").value / Cnt
   CostTOTAL = val(txtTotalAfterVat)
    SngTemp = CostTOTAL
 Dim StoreId1 As Integer
Dim s As String
s = "Select StoreID,StoreID,StoreID2,StoreID3 from tblUsers Where UserID = " & user_id
Set rsDummy = New ADODB.Recordset

rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly, adCmdText
If Not rsDummy.EOF Then
    StoreId1 = val(rsDummy!StoreID & "")
End If

    If SngTemp > 0 Then
        '1 work with branch
        '2 work with inventory
        '3 work with groups
OtherInformation.NextAccount_Code = get_store_Account(val(StoreID), "Account_Code")
        If detect_inventory_work_type = 1 Then
            Account_Code_dynamic = get_account_code_branch(1, val(dcBranch(mIndex).BoundText))
        
            If Account_Code_dynamic = "NO branch" Then
                MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                GoTo ErrTrap
            Else

                If Account_Code_dynamic = "NO account" Then
                    MsgBox "·„ Ì „  ÕœÌœ Õ”«»   ﬂ·›… «·„»Ì⁄«  ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                    GoTo ErrTrap
         
                End If
            End If

            Dim UseCustomerAcc As Integer

    
                StrTempAccountCode = Account_Code_dynamic ' ﬂ·›… «·„»Ì⁄«  1
   

            DebitAccount = StrTempAccountCode
    
            'StrTempAccountCode = "a3a2" ' ﬂ·›… «·„»Ì⁄« 
            If SystemOptions.UserInterface = ArabicInterface Then
                StrTempDes = "  √–‰ ’—›  —ﬁ„     " & Me.TxtNoteSerial11.text & "  "
            Else
                StrTempDes = "Issue Voucher No.  " & Me.TxtNoteSerial11.text & "  "
            End If

            Line1 = setfoxy_Line
            LngDevNO = LngDevNO + 1

            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, 0, StrTempDes, general_noteid, , , , Me.XPDtbTrans(mIndex).value, Me.DCboUserName(mIndex).BoundText, Transaction_ID, , , , , , , , Line1, , , , , , , , , val(Me.dcBranch(mIndex).BoundText), , , , , , , , , , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
                GoTo ErrTrap
            End If
    
    
    
            '«·„Œ“Ê‰ ›Ì «·›—⁄
            Account_Code_dynamic = get_account_code_branch(0, val(dcBranch(mIndex).BoundText))
        
            If Account_Code_dynamic = "NO branch" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                Else
                    MsgBox "The branch was not created", vbCritical
                End If
                GoTo ErrTrap
            Else

                If Account_Code_dynamic = "NO account" Then
                     If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·„ Ì „  ÕœÌœ Õ”«»   ﬂ·›… «·„Œ“Ê‰ ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                Else
                                    MsgBox "The inventory cost calculation in the branch is not specified for this process", vbCritical
                End If
                    GoTo ErrTrap
         
                End If
            End If
        
           
                StrTempAccountCode = Account_Code_dynamic '«·„Œ“Ê‰ 0 ›Ì «·›—⁄
          

            CreditAccount = StrTempAccountCode
    
            If SystemOptions.UserInterface = ArabicInterface Then
                StrTempDes = "√–‰ ’—›  —ﬁ„ " & Me.TxtNoteSerial11.text
            Else
                StrTempDes = "Issue Voucher No. " & Me.TxtNoteSerial11.text
            End If
    
            LngDevNO = LngDevNO + 1
            Line2 = setfoxy_Line

            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, 1, StrTempDes, general_noteid, , , , Me.XPDtbTrans(mIndex).value, Me.DCboUserName(mIndex).BoundText, Transaction_ID, , , , , , , , Line2, , , , , , , , , val(Me.dcBranch(mIndex).BoundText), , , , , , , , , , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
                GoTo ErrTrap
            End If
    
        ElseIf detect_inventory_work_type = 2 Then
            
     'salimhere
     If invoice = 0 Then '« «Ã
     Account_Code_dynamic = get_account_code_branch(37, CInt(BranchID))
        Else
        
        Account_Code_dynamic = get_account_code_branch(1, val(dcBranch(mIndex).BoundText))  '„»Ì⁄« 
        End If
            If Account_Code_dynamic = "NO branch" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                Else
                    MsgBox "The branch was not created", vbCritical
                End If
                GoTo ErrTrap
            Else

                If Account_Code_dynamic = "NO account" Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox "·„ Ì „  ÕœÌœ Õ”«»   ﬂ·›… «·«‰ «Ã ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                    Else
                         MsgBox "The production cost calculation is not determined in the section for this process", vbCritical
                    End If
                    GoTo ErrTrap
         
                End If
            End If

           
            StrTempAccountCode = Account_Code_dynamic ' ﬂ·›… «·„»Ì⁄«  1
          
            DebitAccount = StrTempAccountCode
            
            Line1 = setfoxy_Line

            'StrTempAccountCode = "a3a2" ' ﬂ·›… «·„»Ì⁄« 
            If SystemOptions.UserInterface = ArabicInterface Then
                StrTempDes = "√–‰ ’—›  —ﬁ„ " & Me.TxtNoteSerial11.text
            Else
                StrTempDes = "Issue Voucher No. " & Me.TxtNoteSerial11.text
            End If
    
            LngDevNO = LngDevNO + 1
       Dim project_id As Integer
'        project_id = IIf(Me.DcbProject.BoundText = "", 0, Me.DcbProject.BoundText)
             If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, 0, StrTempDes, general_noteid, , , , Me.XPDtbTrans(mIndex).value, Me.DCboUserName(mIndex).BoundText, Transaction_ID, , , , , , , , Line1, , , , , , , , , val(Me.dcBranch(mIndex).BoundText), , , , , , , , , , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
                GoTo ErrTrap
            End If

            '«·„Œ“Ê‰ «·”·⁄Ì ⁄·Ï „” ÊÏ «·„Œ“‰
            SngTemp = CostTOTAL

            
            Account_Code_dynamic = get_store_Account(val(StoreId1), "Account_Code")
            
        
            If Account_Code_dynamic = "" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·„ Ì „  ÕœÌœ Õ”«»  ··„Œ“Ê‰ «·”·⁄Ì ·Â–« «·„Œ“‰ ›Ì Â–« «·›—⁄    ", vbCritical
                Else
                    MsgBox "No inventory account for this store has been specified in this section  ", vbCritical
                End If
                
                GoTo ErrTrap
            End If
    
            StrTempAccountCode = Account_Code_dynamic  '„Õ“Ê‰ «·”·⁄Ì ··„Œ“‰
            CreditAccount = StrTempAccountCode
OtherInformation.NextAccount_Code = DebitAccount
            ' StrTempAccountCode = "a1a2a5" '„Õ“Ê‰ «·»÷«⁄…
            If SystemOptions.UserInterface = ArabicInterface Then
                StrTempDes = "√–‰ ’—›  —ﬁ„ " & Me.TxtNoteSerial11.text
            Else
                StrTempDes = "Issue Voucher No. " & Me.TxtNoteSerial11.text
            End If

            Line2 = setfoxy_Line
         
            LngDevNO = LngDevNO + 1

            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, SngTemp, 1, StrTempDes, general_noteid, , , , Me.XPDtbTrans(mIndex).value, Me.DCboUserName(mIndex).BoundText, Transaction_ID, , , , , , , , Line2, , , , , , , , , val(Me.dcBranch(mIndex).BoundText), , , , , , , , , , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
                GoTo ErrTrap
            End If

        ElseIf detect_inventory_work_type = 3 Then
            Dim groupAccount As String
             
            Dim line_value As Single

           
                        line_value = val(val(txtTotalAfterVat))
    
                        If SystemOptions.UserInterface = ArabicInterface Then
                            StrTempDes = "√–‰ ’—›  —ﬁ„ " & Me.TxtNoteSerial11.text
                        Else
                            StrTempDes = "Issue Voucher No. " & Me.TxtNoteSerial11.text
                        End If
    
                        LngDevNO = LngDevNO + 1

                        If ModAccounts.AddNewDev(LngDevID, LngDevNO, groupAccount, line_value, 0, StrTempDes, general_noteid, , , , Me.XPDtbTrans(mIndex).value, Me.DCboUserName(mIndex).BoundText, Transaction_ID, , , , , , , , , , , , , , , , , val(Me.dcBranch(mIndex).BoundText)) = False Then
                            GoTo ErrTrap
                        End If
    
                    End If
            Dim mItemId As Long
                If val(mItemId) <> 0 Then

                        ' groupAccount = get_item_group_account(FG.TextMatrix(i, FG.ColIndex("Code")), DCboStoreName.BoundText, 2)
                        groupAccount = get_item_group_account_inventory("", StoreId1, 0)

                        If groupAccount = "Error" Then
                            If SystemOptions.UserInterface = ArabicInterface Then
                                MsgBox " €Ì— „Õœœ —ﬁ„ Õ”«»  «·„Œ“Ê‰ «·”·⁄Ì ··„Œ“‰ «·„Õœœ   ·„Ã„Ê⁄ …"
                            Else
                                MsgBox "Group Name Account Not Defined"
                            End If

                            GoTo ErrTrap
                        End If

                        line_value = val(val(txtTotalAfterVat))
    
                        If SystemOptions.UserInterface = ArabicInterface Then
                            StrTempDes = "√–‰ ’—›  —ﬁ„ " & Me.TxtNoteSerial11.text
                        Else
                            StrTempDes = "Issue Voucher No. " & Me.TxtNoteSerial11.text
                        End If

                        LngDevNO = LngDevNO + 1

                        If ModAccounts.AddNewDev(LngDevID, LngDevNO, groupAccount, line_value, 1, StrTempDes, general_noteid, , , , Me.XPDtbTrans(mIndex).value, Me.DCboUserName(mIndex).BoundText, Transaction_ID, , , , , , , , , , , , , , , , , val(Me.dcBranch(mIndex).BoundText), , , , , , , , , , , , , , , , , , , , , , , , , , , OtherInformation) = False Then
                            GoTo ErrTrap
                        End If
    
                    End If

        

        '----------------
        'LngDevID = LngDevID + 1
        'LngDevNO = 0
    End If
   ' ute StrSQL
ErrTrap:
End Function





Function print_report(Optional NoteSerial As String, Optional indexe As Integer)
    
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
        MySQL = "SELECT *,tc.CusName,tc.CusNamee FROM tblJobOrders"
    MySQL = MySQL & " INNER JOIN TblCustemers AS tc"
    MySQL = MySQL & " ON tc.CusID = tblJobOrders.CusId"
    MySQL = MySQL & " Where tblJobOrders.Id =" & val(TxtSerial1(mIndex))


        If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "JobOrders.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "JobOrders.rpt"
        End If

        ''''''


    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        'GetMsgs 138, vbExclamation
        Msg = "?CE??I E?C?CE ?????"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        RsData.Close
        Set RsData = Nothing
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)
    xReport.Database.SetDataSource RsData

    Dim cCompanyInfo As New ClsCompanyInfo
    Dim oorderdate As Date
    Dim CBoBasedON As Integer
    Dim PONo As String

     
    
    If SystemOptions.UserInterface = ArabicInterface Then
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
        ' xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
        StrReportTitle = "" '& StrAccountName
        'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        '    StrReportTitle = StrReportTitle + " EIC?E ?? " & Format(Me.DTPickerAccFrom.value, "yyyy/M/d") & ""
        'End If
        'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        '    StrReportTitle = StrReportTitle + " ??? " & Format(Me.DTPickerAccTo.value, "yyyy/M/d") & " "
        'End If
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        'xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(val(dcBranch(mIndex).BoundText)))
        StrReportTitle = ""
        'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        '    StrReportTitle = StrReportTitle + " From Date " & (Me.DTPickerAccFrom.value) & ""
        'End If
        'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        '    StrReportTitle = StrReportTitle + " To Date :  " & (Me.DTPickerAccTo.value) & ""
        'End If
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
      '  xReport.ParameterFields(4).AddCurrentValue WriteNo(Format(val(TxtAdvanceValue.text), "0.00"), 0, True, ".")
       ' xReport.ParameterFields(6).AddCurrentValue val(lbl(23).Caption)
       '  xReport.ParameterFields(7).AddCurrentValue DBIssueDate.value
'    xReport.ParameterFields(8).AddCurrentValue IIf(IsNumeric(fg.TextMatrix(Me.fg.FixedRows, fg.ColIndex("PartValue"))), val(fg.TextMatrix(Me.fg.FixedRows, fg.ColIndex("PartValue"))), 0)
' xReport.ParameterFields(9).AddCurrentValue val(lbl(22).Caption)
 ' xReport.ParameterFields(10).AddCurrentValue val(TxtDiscount.text)
  ' xReport.ParameterFields(12).AddCurrentValueval (lbTotalMente.Caption)
  
   
'    xReport.ParameterFields(5).AddCurrentValue ToHijriDate(RsData("notedate").value)
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault


 
  
 
End Function





Private Sub XPDtbBill_Change(Index As Integer)

ISButton2_Click


End Sub

Private Sub XPTxtVal_Change()
If Me.TxtModFlg2(mIndex).text = "N" Or Me.TxtModFlg2(mIndex).text = "E" Then
    CalCulteVAT 3
End If
End Sub
Sub CalCulteVAT(Optional Ind As Integer = 0)

If Me.TxtModFlg2(mIndex).text = "N" Or Me.TxtModFlg2(mIndex).text = "E" Then
    Dim AccountVATCreit As String
    Dim Percetage As Double
    
    Dim mVal As Double
        
        If Ind = 3 Then
            PercentgValueAddedAccount_Transec XPDtbTrans(mIndex).value, 21, 0, AccountVATCreit, Percetage
            TxtVAt22.text = val(Format((XPTxtVal.text), "###.00")) * Percetage / 100
             
           '  TxtVATValue.Text = val(Format((XPTxtVal.Text), "###.00")) * Percetage / 100
           '  TxtVAt2.Text = TxtVATValue.Text
             
             
             mVal = val(Format((XPTxtVal.text), "###.00"))
            ' TxtVATValue.Text = val(Format((mVal), "###.00")) * Percetage / 100
             txtTotalWithVat2.text = Round(val(Format((mVal), "###.00")) + val(TxtVAt22.text), 2)
             
             
    '         Exit Sub
        End If
        'XPDtbTrans.value = 100
        'XPTxtVal = 100
        
         txtTotalWithVat2.text = Round(val(Format((mVal), "###.00")) + val(TxtVAt22.text), 2)
'        If optCash Then
'            txtAmountCash = val(txtTotalWithVat2) - val(txtAmountVisa)
'        End If
'
        If optLater Then
            txtAmountLater = val(txtTotalWithVat2)
        End If
        
    '    If SystemOptions.UserInterface = ArabicInterface Then
    '        Me.lblTotalNet.Caption = WriteNo(txtTotalWithVat2.Text, 0, True, ".", , 0)
    '    Else
    '        Me.lblTotalNet.Caption = WriteNo(txtTotalWithVat2.Text, 0, True, ".", , 1)
    '    End If
    'TxtVAt2.Text = TxtVATValue.Text
    'TxtVAt22.Text = TxtVATValue.Text
End If
End Sub


Private Sub XPTxtVal_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_Num(KeyAscii, Me.XPTxtVal.text, 1)
End Sub



Private Sub ZKFPEngX1_OnCapture(ByVal ActionResult As Boolean, ByVal atemplate As Variant)
    
   If Not mSenesor Then Exit Sub
    Dim fi As Long, i As Long
    Dim lentgh  As Long
    Dim Score As Long, ProcessNum As Long
    Dim RegChanged As Boolean
    Dim sTemp1 As String
    Dim sTemp As Variant
    Dim AtempFinger
    cmbEmpName.BoundText = 0
    sTemp1 = ZKFPEngX1.GetTemplateAsString()
    sTemp = ZKFPEngX1.GetTemplate()
  
  AtempFinger = atemplate
    StatusBar.Caption = "Acqired Template"
    
'   If ZKFPEngX1.SaveTemplate("C:\fingerprint.tpl", AtempFinger) Then
'   MsgBox "SaveTemplate C:\fingerprint.tpl success"
'   Else
'   MsgBox "SaveTemplate fail"
'   End If
   txtFingerPrint = sTemp1
    
'    If FMatchType = 1 Then  '1:1
'       If ZKFPEngX1.VerFingerFromStr(FRegTemplate, sTemp1, False, RegChanged) Then
'          MessageBox 0, "Verify success", "information", 0
'       Else
'          MessageBox 0, "Verify Failed", "information", 0
'       End If
'
'    ElseIf FMatchType = 2 Then '1:N
'       Score = 8
'
'        fi = ZKFPEngX1.IdentificationInFPCacheDB(fpcHandle, sTemp, Score, ProcessNum)
'       If fi = -1 Then
'          MessageBox 0, "Identification failed£°", "information", 0
'       Else
'          MessageBox 0, "Identification Success Name=" & FFingerNames(fi) & " Score = " & Score & " Processed Number = " & ProcessNum, "information", 0
'       End If
'    End If
'
   
End Sub

Private Sub ZKFPEngX1_OnEnroll(ByVal ActionResult As Boolean, ByVal atemplate As Variant)
  Dim i As Long
  If Not mSenesor Then Exit Sub
  If Not ActionResult Then
    MessageBox 0, "Register failed", "Warning", 0
  Else
    MessageBox 0, "Regsiter success", "Information", 0
    
  
    FRegTemplate = ZKFPEngX1.GetTemplateAsString()
    FRegTemp = ZKFPEngX1.GetTemplate()
     
               

     ZKFPEngX1.AddRegTemplateToFPCacheDB fpcHandle, FingerCount, FRegTemp
        ReDim Preserve FFingerNames(FingerCount + 1)
'    FFingerNames(FingerCount) = TextFingerName.Text
    FingerCount = FingerCount + 1
  End If
End Sub

Private Sub ZKFPEngX1_OnFeatureInfo(ByVal AQuality As Long)
  Dim sTemp As String
  If Not mSenesor Then Exit Sub
  sTemp = ""
  If ZKFPEngX1.IsRegister Then
     If ZKFPEngX1.EnrollIndex - 1 > 0 Then
     sTemp = "Register status: still press finger " & ZKFPEngX1.EnrollIndex - 1 & " times"
     Else
     sTemp = ""
     End If
  End If
  sTemp = sTemp & " Fingerprint quality"
  If AQuality <> 0 Then
     sTemp = sTemp & " no good " & AQuality
  Else
     sTemp = sTemp & " good"
  End If
  StatusBar.Caption = sTemp
End Sub

Private Sub ZKFPEngX1_OnImageReceived(AImageValid As Boolean)
  ZKFPEngX1.PrintImageAt hDC, Frame4(3).Width + 6, Frame4(3).top, ZKFPEngX1.ImageWidth, ZKFPEngX1.ImageHeight
  End Sub











Private Sub ZKFPEngX2_OnCapture(ByVal ActionResult As Boolean, ByVal atemplate As Variant)
    Dim fi As Long, i As Long
    Dim lentgh  As Long
    Dim Score As Long, ProcessNum As Long
    Dim RegChanged As Boolean
    Dim sTemp1 As String
    Dim sTemp As Variant
    Dim AtempFinger
    sTemp1 = ZKFPEngX2.GetTemplateAsString()
    sTemp = ZKFPEngX2.GetTemplate()
  
  AtempFinger = atemplate
    StatusBar.Caption = "Acqired Template"
    
'   If ZKFPEngX2.SaveTemplate("C:\fingerprint.tpl", AtempFinger) Then
'   MsgBox "SaveTemplate C:\fingerprint.tpl success"
'   Else
'   MsgBox "SaveTemplate fail"
'   End If
   txtFingerPrint2 = sTemp1
    
'    If FMatchType = 1 Then  '1:1
'       If ZKFPEngX2.VerFingerFromStr(FRegTemplate, sTemp1, False, RegChanged) Then
'          MessageBox 0, "Verify success", "information", 0
'       Else
'          MessageBox 0, "Verify Failed", "information", 0
'       End If
'
'    ElseIf FMatchType = 2 Then '1:N
'       Score = 8
'
'        fi = ZKFPEngX2.IdentificationInFPCacheDB(fpcHandle, sTemp, Score, ProcessNum)
'       If fi = -1 Then
'          MessageBox 0, "Identification failed£°", "information", 0
'       Else
'          MessageBox 0, "Identification Success Name=" & FFingerNames(fi) & " Score = " & Score & " Processed Number = " & ProcessNum, "information", 0
'       End If
'    End If
'
   
End Sub

Private Sub ZKFPEngX2_OnEnroll(ByVal ActionResult As Boolean, ByVal atemplate As Variant)
  Dim i As Long
  
  If Not ActionResult Then
    MessageBox 0, "Register failed", "Warning", 0
  Else
    MessageBox 0, "Regsiter success", "Information", 0
    
  
    FRegTemplate = ZKFPEngX2.GetTemplateAsString()
    FRegTemp = ZKFPEngX2.GetTemplate()
     
               

     ZKFPEngX2.AddRegTemplateToFPCacheDB fpcHandle, FingerCount, FRegTemp
        ReDim Preserve FFingerNames(FingerCount + 1)
'    FFingerNames(FingerCount) = TextFingerName.Text
    FingerCount = FingerCount + 1
  End If
End Sub

Private Sub ZKFPEngX2_OnFeatureInfo(ByVal AQuality As Long)
  Dim sTemp As String
  
  sTemp = ""
  If ZKFPEngX2.IsRegister Then
     If ZKFPEngX2.EnrollIndex - 1 > 0 Then
     sTemp = "Register status: still press finger " & ZKFPEngX2.EnrollIndex - 1 & " times"
     Else
     sTemp = ""
     End If
  End If
  sTemp = sTemp & " Fingerprint quality"
  If AQuality <> 0 Then
     sTemp = sTemp & " no good " & AQuality
  Else
     sTemp = sTemp & " good"
  End If
  StatusBar.Caption = sTemp
End Sub

Private Sub ZKFPEngX2_OnImageReceived(AImageValid As Boolean)
  ZKFPEngX2.PrintImageAt hDC, Frame4(2).Width + 2, Frame4(2).top, ZKFPEngX2.ImageWidth, ZKFPEngX2.ImageHeight
  End Sub


Private Sub SetGridFinger()
GrdFinger2.rows = 1
GrdFinger2.rows = 6
GrdFinger2.TextMatrix(1, GrdFinger2.ColIndex("Finger")) = "«·«’»⁄ «·«Ê·"
GrdFinger2.TextMatrix(2, GrdFinger2.ColIndex("Finger")) = "«·«’»⁄ «·À«‰Ì"
GrdFinger2.TextMatrix(3, GrdFinger2.ColIndex("Finger")) = "«·«’»⁄ «·À«·À"
GrdFinger2.TextMatrix(4, GrdFinger2.ColIndex("Finger")) = "«·«’»⁄ «·—«»⁄"
GrdFinger2.TextMatrix(5, GrdFinger2.ColIndex("Finger")) = "«·«’»⁄ «·Œ«„”"

End Sub



Private Sub GetDataCus()
            Set rsDD = New ADODB.Recordset
            rsDD.Open Build_Sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
        Dim Msg As String
            If rsDD.RecordCount < 1 Then
                FGS.Clear flexClearScrollable, flexClearEverything
                FGS.rows = 2
                
                       If SystemOptions.UserInterface = ArabicInterface Then
                Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
                Else
                Msg = "No Avilable Data"
                End If
                
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Exit Sub
            End If
'
'            If SystemOptions.UserInterface = ArabicInterface Then
'                Me.XPLbl(2).Caption = "‰ ÌÃ… «·»ÕÀ : " & rsDD.RecordCount
'            Else
'                Me.XPLbl(2).Caption = "Search Results: " & rsDD.RecordCount
'            End If

            RetriveSE
            FGS.SetFocus

End Sub



Private Sub RetriveSE()
    Dim Num As Integer
    On Error GoTo ErrTrap
    FGS.Clear flexClearScrollable, flexClearEverything
    
    If Not (rsDD.EOF Or rsDD.BOF) Then
        FGS.rows = rsDD.RecordCount + 1

        For Num = 1 To rsDD.RecordCount

            With FGS
                .TextMatrix(Num, .ColIndex("NumIndex")) = Num
                .TextMatrix(Num, .ColIndex("MemCode")) = IIf(IsNull(rsDD("CusID").value), "", val(rsDD("CusID").value))
                'Fullcode
                .TextMatrix(Num, .ColIndex("Fullcode")) = IIf(IsNull(rsDD("Fullcode").value), "", (rsDD("Fullcode").value))
           If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(Num, .ColIndex("MemNme")) = IIf(IsNull(rsDD("CusName").value), "", Trim(rsDD("CusName").value))
            Else
            .TextMatrix(Num, .ColIndex("MemNme")) = IIf(IsNull(rsDD("CusNameE").value), "", Trim(rsDD("CusNameE").value))
            End If
                .TextMatrix(Num, .ColIndex("Phone")) = IIf(IsNull(rsDD("Cus_Phone").value), "", Trim(rsDD("Cus_Phone").value))
                
                .TextMatrix(Num, .ColIndex("Cus_mobile")) = IIf(IsNull(rsDD("Cus_mobile").value), "", Trim(rsDD("Cus_mobile").value))
                .TextMatrix(Num, .ColIndex("CustGID")) = IIf(IsNull(rsDD("CustGID").value), "", Trim(rsDD("CustGID").value))
                
            End With

            rsDD.MoveNext
        Next Num

        FGS.AutoSize 0, FGS.Cols - 1, False
    End If

    Exit Sub
ErrTrap:
End Sub


Private Function Build_Sql()
    Dim StrSQL As String
    Dim Begin As Boolean
    Dim StrWhere As String
    On Error GoTo ErrTrap


    StrSQL = "select * From TblCustemers where 1=1"
      Begin = True

   

        StrSQL = "select * From TblCustemers where type=1 "
        StrSQL = StrSQL & " and ( BranchId=0  or      BranchId in(" & Current_branchSql & "))"
        Begin = True

ll:
    If XPTxtCusID.text <> "" Then
        If Begin = True Then
            StrWhere = StrWhere + " and fullcode LIKE '%" & (XPTxtCusID.text) & "%'"
        Else
            StrWhere = StrWhere + " and fullcode LIKE '%" & (XPTxtCusID.text) & "%'"
            Begin = True
        End If
    End If


    If xptxtphone.text <> "" Then
        If Begin = True Then
            StrWhere = StrWhere + " and Cus_Phone LIKE'%" & (xptxtphone.text) & "%'"
          Else
            StrWhere = StrWhere + " and Cus_Phone LIKE'%" & (xptxtphone.text) & "%"
            
            Begin = True
        End If
    End If

'    If XPTxtmobil.Text <> "" Then
'        If Begin = True Then
'
'                        StrWhere = StrWhere + " and Cus_mobile LIKE'%" & (XPTxtmobile.Text) & "%'"
'
'        Else
'                        StrWhere = StrWhere + " and Cus_mobile LIKE'%" & (XPTxtmobile.Text) & "%'"
'            Begin = True
'        End If
'    End If
'
'
'    If txtCustGID.Text <> "" Then
'        If Begin = True Then
'            StrWhere = StrWhere + " and CustGID=" & (txtCustGID.Text)
'        Else
'            StrWhere = StrWhere + " where CustGID=" & (txtCustGID.Text)
'            Begin = True
'        End If
'    End If

If SystemOptions.UserInterface = ArabicInterface Then
    If Trim(Me.txtCustomerName.text) <> "" Then
'        If XPChkSearchType.value = Checked Then
'            If Begin = True Then
'                StrWhere = StrWhere + " and CusName ='" & Trim(Me.txtCustomerName.Text) & "'"
'            Else
'                StrWhere = StrWhere + " where CusName ='" & Trim(Me.txtCustomerName.Text) & "'"
'                Begin = True
'            End If
'
'        Else

            If Begin = True Then
                StrWhere = StrWhere + " and CusName like '%" & Trim(txtCustomerName2.text) & "%'"
            Else
                StrWhere = StrWhere + " where CusName like '%" & Trim(txtCustomerName2.text) & "%'"
                Begin = True
            End If
        End If
'    End If

Else


  If Trim(Me.txtCustomerName2.text) <> "" Then
'     '   If XPChkSearchType.value = Checked Then
'            If Begin = True Then
'                StrWhere = StrWhere + " and CusNameE ='" & Trim(Me.txtCustomerName.Text) & "'"
'            Else
'                StrWhere = StrWhere + " where CusNameE ='" & Trim(Me.txtCustomerName.Text) & "'"
'                Begin = True
'            End If
'
'        Else

            If Begin = True Then
                StrWhere = StrWhere + " and CusNameE like '%" & Trim(txtCustomerName2.text) & "%'"
            Else
                StrWhere = StrWhere + " where CusNameE like '%" & Trim(txtCustomerName2.text) & "%'"
                Begin = True
            End If
        End If
    'End If
End If
  If Begin = True Then
 '               StrWhere = StrWhere + " and  ( BranchId in(" & Current_branchSql & ") or (BranchId is null)or BranchId=0 )"
          Else
 '              StrWhere = StrWhere + " where    (BranchId in(" & Current_branchSql & ") or (BranchId is null)or BranchId=0) "
               Begin = True
           End If
    Build_Sql = StrSQL + StrWhere
    Exit Function
ErrTrap:
End Function



