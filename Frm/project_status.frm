VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SuiteCtrls.ocx"
Begin VB.Form project_status 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "«‰Ê«⁄ «·Ê—œÌ« "
   ClientHeight    =   9690
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   18450
   Icon            =   "project_status.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   9690
   ScaleWidth      =   18450
   ShowInTaskbar   =   0   'False
   Begin C1SizerLibCtl.C1Tab TabMain 
      Height          =   9690
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   18870
      _cx             =   33285
      _cy             =   17092
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
      Caption         =   "«‰Ê«⁄ «·Ê—œÌ« | ”·Ì„ Ê—œÌ…|Õ«·«  «·„‘«—Ì⁄|«‰Ê«⁄ «·«ÿ»«¡"
      Align           =   0
      CurrTab         =   3
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
         Height          =   9315
         Index           =   0
         Left            =   -20025
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   45
         Width           =   18780
         _cx             =   33126
         _cy             =   16431
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
            Height          =   2190
            Left            =   60
            RightToLeft     =   -1  'True
            TabIndex        =   103
            Top             =   4320
            Width           =   6255
            Begin VB.ComboBox Combo3 
               BackColor       =   &H80000018&
               Height          =   315
               ItemData        =   "project_status.frx":058A
               Left            =   2280
               List            =   "project_status.frx":059A
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   107
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
               Index           =   1
               Left            =   3120
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   106
               Top             =   285
               Width           =   1065
            End
            Begin VB.TextBox txtName 
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
               Left            =   1395
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   105
               Top             =   645
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
               Left            =   1395
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   104
               Top             =   960
               Width           =   2760
            End
            Begin VB.Label Label6 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬂÊœ «·Ê—œÌ…"
               Height          =   195
               Left            =   4695
               RightToLeft     =   -1  'True
               TabIndex        =   110
               Top             =   390
               Width           =   990
            End
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·Ê—œÌ… ⁄—»Ì"
               Height          =   285
               Index           =   0
               Left            =   4350
               RightToLeft     =   -1  'True
               TabIndex        =   109
               Top             =   720
               Width           =   1350
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·Ê—œÌ…«‰Ã·Ì“Ì"
               Height          =   285
               Left            =   4200
               RightToLeft     =   -1  'True
               TabIndex        =   108
               Top             =   1080
               Width           =   1500
            End
         End
         Begin VB.Frame Fra_Header 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   570
            Index           =   0
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   2
            Top             =   0
            Width           =   18840
            Begin VB.Frame Frame2 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   375
               Index           =   0
               Left            =   540
               RightToLeft     =   -1  'True
               TabIndex        =   5
               Top             =   450
               Visible         =   0   'False
               Width           =   3105
               Begin MSDataListLib.DataCombo DataCombo4 
                  CausesValidation=   0   'False
                  Height          =   315
                  Index           =   0
                  Left            =   -255
                  TabIndex        =   6
                  Tag             =   "„‰ ›÷·ﬂ √œŒ· —ﬁ„ «·ﬁ÷Ì…"
                  Top             =   -15
                  Width           =   2340
                  _ExtentX        =   4128
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
                  Index           =   4
                  Left            =   2160
                  RightToLeft     =   -1  'True
                  TabIndex        =   7
                  Top             =   45
                  Width           =   855
               End
            End
            Begin VB.TextBox TxtModFlg2 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H0000FF00&
               Enabled         =   0   'False
               Height          =   285
               Index           =   1
               Left            =   2580
               RightToLeft     =   -1  'True
               TabIndex        =   4
               Text            =   "modflag"
               Top             =   90
               Visible         =   0   'False
               Width           =   465
            End
            Begin VB.TextBox Text2 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H000000FF&
               Height          =   240
               Index           =   0
               Left            =   3030
               RightToLeft     =   -1  'True
               TabIndex        =   3
               Top             =   510
               Visible         =   0   'False
               Width           =   945
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
                     Picture         =   "project_status.frx":05B3
                     Key             =   "CompanyName"
                  EndProperty
                  BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "project_status.frx":094D
                     Key             =   "Ser"
                  EndProperty
                  BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "project_status.frx":0CE7
                     Key             =   "Vac_Name"
                  EndProperty
                  BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "project_status.frx":1081
                     Key             =   "ShareCount"
                  EndProperty
                  BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "project_status.frx":141B
                     Key             =   "Dis_Count"
                  EndProperty
                  BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "project_status.frx":17B5
                     Key             =   "Bouns"
                  EndProperty
                  BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "project_status.frx":1B4F
                     Key             =   "SharesValue"
                  EndProperty
                  BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "project_status.frx":20E9
                     Key             =   "BuyValue"
                  EndProperty
               EndProperty
            End
            Begin ImpulseButton.ISButton btn_Last 
               Height          =   315
               Index           =   1
               Left            =   90
               TabIndex        =   8
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
               ButtonImage     =   "project_status.frx":2483
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
               TabIndex        =   9
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
               ButtonImage     =   "project_status.frx":281D
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btn_Previous 
               Height          =   315
               Index           =   1
               Left            =   1155
               TabIndex        =   10
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
               ButtonImage     =   "project_status.frx":2BB7
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btn_First 
               Height          =   315
               Index           =   1
               Left            =   1620
               TabIndex        =   11
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
               ButtonImage     =   "project_status.frx":2F51
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«‰Ê«⁄ «·Ê—œÌ« "
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
               Index           =   5
               Left            =   11340
               RightToLeft     =   -1  'True
               TabIndex        =   12
               Top             =   90
               Width           =   2640
            End
         End
         Begin VSFlex8UCtl.VSFlexGrid FgItems 
            Height          =   9195
            Index           =   1
            Left            =   25950
            TabIndex        =   13
            Top             =   810
            Width           =   18585
            _cx             =   32782
            _cy             =   16219
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
            FormatString    =   $"project_status.frx":32EB
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
         Begin ImpulseButton.ISButton btn_New 
            Height          =   345
            Index           =   1
            Left            =   6675
            TabIndex        =   14
            Top             =   8610
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   609
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
            ButtonImage     =   "project_status.frx":33AB
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Save 
            Height          =   345
            Index           =   1
            Left            =   4875
            TabIndex        =   15
            Top             =   8610
            Width           =   900
            _ExtentX        =   1588
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
            ButtonImage     =   "project_status.frx":3745
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Modify 
            Height          =   345
            Index           =   1
            Left            =   5775
            TabIndex        =   16
            Top             =   8610
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   609
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
            ButtonImage     =   "project_status.frx":3ADF
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton Btn_Undo 
            Height          =   345
            Index           =   1
            Left            =   3975
            TabIndex        =   17
            Top             =   8610
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   609
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
            ButtonImage     =   "project_status.frx":3E79
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Delete 
            Height          =   345
            Index           =   1
            Left            =   3135
            TabIndex        =   18
            Top             =   8610
            Width           =   840
            _ExtentX        =   1482
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
            ButtonImage     =   "project_status.frx":4213
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton Btn_Update 
            Height          =   345
            Index           =   1
            Left            =   5475
            TabIndex        =   19
            TabStop         =   0   'False
            ToolTipText     =   " ÕœÌÀ ﬁ«⁄œ… «·»Ì«‰« "
            Top             =   7830
            Visible         =   0   'False
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   609
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
            ButtonImage     =   "project_status.frx":47AD
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Cancel 
            Height          =   345
            Index           =   1
            Left            =   60
            TabIndex        =   20
            Top             =   8610
            Width           =   900
            _ExtentX        =   1588
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
            ButtonImage     =   "project_status.frx":4B47
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton Btn_Print 
            Height          =   540
            Index           =   1
            Left            =   2040
            TabIndex        =   21
            TabStop         =   0   'False
            ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
            Top             =   8520
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   953
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
            ButtonImage     =   "project_status.frx":4EE1
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btn_Query 
            Height          =   600
            Index           =   1
            Left            =   780
            TabIndex        =   22
            TabStop         =   0   'False
            ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
            Top             =   8475
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   1058
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
            ButtonImage     =   "project_status.frx":B743
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton Cmd_DeleteRow 
            Height          =   300
            Index           =   1
            Left            =   2220
            TabIndex        =   23
            Top             =   7785
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   529
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
            ButtonImage     =   "project_status.frx":BADD
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton Cmd_DeleteAll 
            Height          =   300
            Index           =   1
            Left            =   360
            TabIndex        =   24
            Top             =   7785
            Width           =   1740
            _ExtentX        =   3069
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
            ButtonImage     =   "project_status.frx":C077
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin VSFlex8Ctl.VSFlexGrid Grid2 
            Height          =   3600
            Left            =   60
            TabIndex        =   111
            Top             =   690
            Width           =   7530
            _cx             =   13282
            _cy             =   6350
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
            FormatString    =   $"project_status.frx":C611
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
            Height          =   225
            Index           =   1
            Left            =   2775
            RightToLeft     =   -1  'True
            TabIndex        =   28
            Top             =   8280
            Width           =   540
         End
         Begin VB.Label LabCurr_Rec 
            BackColor       =   &H00E2E9E9&
            Height          =   225
            Index           =   1
            Left            =   4575
            RightToLeft     =   -1  'True
            TabIndex        =   27
            Top             =   8280
            Width           =   780
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «·”Ã·« :"
            Height          =   210
            Index           =   2
            Left            =   3495
            RightToLeft     =   -1  'True
            TabIndex        =   26
            Top             =   8265
            Width           =   1080
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”Ã· «·Õ«·Ì:"
            Height          =   210
            Index           =   3
            Left            =   5415
            RightToLeft     =   -1  'True
            TabIndex        =   25
            Top             =   8265
            Width           =   1080
         End
      End
      Begin C1SizerLibCtl.C1Elastic ELe 
         Height          =   9315
         Index           =   2
         Left            =   -19725
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   45
         Width           =   18780
         _cx             =   33126
         _cy             =   16431
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
         Begin VB.ComboBox cmbCarStatus 
            Height          =   315
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   174
            Top             =   1680
            Width           =   1800
         End
         Begin VB.ComboBox DcbType 
            Height          =   315
            Left            =   12225
            RightToLeft     =   -1  'True
            TabIndex        =   168
            Top             =   780
            Width           =   1320
         End
         Begin VB.TextBox txtNoteLate 
            Alignment       =   1  'Right Justify
            Height          =   645
            Left            =   240
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   161
            Top             =   7770
            Width           =   8730
         End
         Begin VB.TextBox txtRemarks 
            Alignment       =   1  'Right Justify
            Height          =   645
            Left            =   9690
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   159
            Top             =   7770
            Width           =   8310
         End
         Begin VB.TextBox txtNoteStill 
            Alignment       =   1  'Right Justify
            Height          =   3270
            Left            =   240
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   157
            Top             =   4170
            Width           =   8730
         End
         Begin VB.TextBox txtNoteDone 
            Alignment       =   1  'Right Justify
            Height          =   3240
            Left            =   9690
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   156
            Top             =   4200
            Width           =   8310
         End
         Begin VB.ComboBox DcbStutsMaint 
            Height          =   315
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   146
            Top             =   1305
            Width           =   1800
         End
         Begin VB.Frame Frame6 
            BackColor       =   &H00E2E9E9&
            Caption         =   "ﬁ«∆œ «·„⁄œ…"
            Enabled         =   0   'False
            Height          =   1035
            Left            =   10290
            RightToLeft     =   -1  'True
            TabIndex        =   140
            Top             =   2280
            Width           =   7170
            Begin VB.TextBox Text5 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   4830
               RightToLeft     =   -1  'True
               TabIndex        =   142
               Top             =   240
               Width           =   1065
            End
            Begin VB.TextBox TxtLeaderName 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   141
               Top             =   600
               Width           =   5775
            End
            Begin XtremeSuiteControls.RadioButton ChLeaderType 
               Height          =   255
               Index           =   0
               Left            =   5640
               TabIndex        =   143
               Top             =   240
               Width           =   1455
               _Version        =   786432
               _ExtentX        =   2566
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "„ÊŸ›"
               BackColor       =   14871017
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcbLeaderID 
               Bindings        =   "project_status.frx":C6A7
               Height          =   315
               Left            =   120
               TabIndex        =   144
               Top             =   240
               Width           =   4575
               _ExtentX        =   8070
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
            Begin XtremeSuiteControls.RadioButton ChLeaderType 
               Height          =   255
               Index           =   1
               Left            =   5640
               TabIndex        =   145
               Top             =   600
               Width           =   1455
               _Version        =   786432
               _ExtentX        =   2566
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "€Ì— „ÊŸ›"
               BackColor       =   14871017
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00E2E9E9&
            Caption         =   "„”·„ «·„⁄œ…"
            Enabled         =   0   'False
            Height          =   1035
            Left            =   3495
            RightToLeft     =   -1  'True
            TabIndex        =   134
            Top             =   2280
            Width           =   6795
            Begin VB.TextBox TxtDrievName 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   630
               RightToLeft     =   -1  'True
               TabIndex        =   136
               Top             =   600
               Width           =   5115
            End
            Begin VB.TextBox Text6 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   4710
               RightToLeft     =   -1  'True
               TabIndex        =   135
               Top             =   240
               Width           =   1065
            End
            Begin XtremeSuiteControls.RadioButton ChDrievType 
               Height          =   255
               Index           =   0
               Left            =   5280
               TabIndex        =   137
               Top             =   240
               Width           =   1455
               _Version        =   786432
               _ExtentX        =   2566
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "„ÊŸ›"
               BackColor       =   14871017
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcbDrievID 
               Bindings        =   "project_status.frx":C6BC
               Height          =   315
               Left            =   630
               TabIndex        =   138
               Top             =   240
               Width           =   4065
               _ExtentX        =   7170
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
            Begin XtremeSuiteControls.RadioButton ChDrievType 
               Height          =   255
               Index           =   1
               Left            =   5280
               TabIndex        =   139
               Top             =   600
               Width           =   1455
               _Version        =   786432
               _ExtentX        =   2566
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "€Ì— „ÊŸ›"
               BackColor       =   14871017
               UseVisualStyle  =   -1  'True
               TextAlignment   =   1
               RightToLeft     =   -1  'True
            End
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00E2E9E9&
            Enabled         =   0   'False
            Height          =   1140
            Left            =   8250
            RightToLeft     =   -1  'True
            TabIndex        =   116
            Top             =   1155
            Width           =   9210
            Begin VB.TextBox TxtOperatorN 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   2520
               TabIndex        =   118
               Top             =   240
               Width           =   1635
            End
            Begin VB.TextBox TxtBoardNO 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   2520
               TabIndex        =   117
               Top             =   600
               Width           =   1635
            End
            Begin MSDataListLib.DataCombo DcbEquepment 
               Height          =   315
               Left            =   5280
               TabIndex        =   119
               Top             =   240
               Width           =   3135
               _ExtentX        =   5530
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcbBranchFrom 
               Height          =   315
               Left            =   5280
               TabIndex        =   120
               Top             =   600
               Width           =   3135
               _ExtentX        =   5530
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic7 
               Height          =   435
               Left            =   120
               TabIndex        =   121
               TabStop         =   0   'False
               Top             =   600
               Width           =   2325
               _cx             =   4101
               _cy             =   767
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
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
               Begin VB.TextBox txtLetter1 
                  Alignment       =   2  'Center
                  Height          =   315
                  Left            =   1935
                  MaxLength       =   1
                  RightToLeft     =   -1  'True
                  TabIndex        =   129
                  Top             =   0
                  Width           =   285
               End
               Begin VB.TextBox txtLetter2 
                  Alignment       =   2  'Center
                  Height          =   315
                  Left            =   1710
                  MaxLength       =   1
                  RightToLeft     =   -1  'True
                  TabIndex        =   128
                  Top             =   0
                  Width           =   240
               End
               Begin VB.TextBox txtLetter3 
                  Alignment       =   2  'Center
                  Height          =   315
                  Left            =   1440
                  MaxLength       =   1
                  RightToLeft     =   -1  'True
                  TabIndex        =   127
                  Top             =   0
                  Width           =   315
               End
               Begin VB.TextBox txtNum1 
                  Alignment       =   2  'Center
                  Height          =   315
                  Left            =   795
                  MaxLength       =   1
                  RightToLeft     =   -1  'True
                  TabIndex        =   126
                  Top             =   0
                  Width           =   360
               End
               Begin VB.TextBox txtNum2 
                  Alignment       =   2  'Center
                  Height          =   315
                  Left            =   480
                  MaxLength       =   1
                  RightToLeft     =   -1  'True
                  TabIndex        =   125
                  Top             =   0
                  Width           =   330
               End
               Begin VB.TextBox txtNum3 
                  Alignment       =   2  'Center
                  Height          =   315
                  Left            =   270
                  MaxLength       =   1
                  RightToLeft     =   -1  'True
                  TabIndex        =   124
                  Top             =   0
                  Width           =   300
               End
               Begin VB.TextBox txtLetter4 
                  Alignment       =   2  'Center
                  Height          =   315
                  Left            =   1155
                  MaxLength       =   1
                  RightToLeft     =   -1  'True
                  TabIndex        =   123
                  Top             =   0
                  Width           =   360
               End
               Begin VB.TextBox txtNum4 
                  Alignment       =   2  'Center
                  Height          =   315
                  Left            =   0
                  MaxLength       =   1
                  RightToLeft     =   -1  'True
                  TabIndex        =   122
                  Top             =   0
                  Width           =   300
               End
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ «··ÊÕ…"
               Height          =   285
               Index           =   67
               Left            =   4080
               TabIndex        =   133
               Top             =   600
               Width           =   1005
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·—ﬁ„ «· ‘€Ì·Ì"
               Height          =   285
               Index           =   66
               Left            =   4200
               TabIndex        =   132
               Top             =   240
               Width           =   1005
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ÃÂ…"
               Height          =   285
               Index           =   2
               Left            =   7800
               TabIndex        =   131
               Top             =   600
               Width           =   1125
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„⁄œÂ"
               Height          =   285
               Index           =   29
               Left            =   7800
               TabIndex        =   130
               Top             =   240
               Width           =   1125
            End
         End
         Begin VB.TextBox txtDes 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   600
            Left            =   3555
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   115
            Top             =   1665
            Width           =   3360
         End
         Begin VB.TextBox TxtSearchCode 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   330
            Left            =   6075
            RightToLeft     =   -1  'True
            TabIndex        =   114
            Top             =   1290
            Width           =   840
         End
         Begin VB.TextBox TxtDeptNotes 
            Alignment       =   1  'Right Justify
            Height          =   480
            Left            =   10470
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   113
            Top             =   3390
            Width           =   6990
         End
         Begin VB.TextBox TxtInitialNotes 
            Alignment       =   1  'Right Justify
            Height          =   480
            Left            =   2595
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   112
            Top             =   3390
            Width           =   6375
         End
         Begin VB.TextBox TxtSerial1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Enabled         =   0   'False
            Height          =   405
            Index           =   2
            Left            =   15945
            RightToLeft     =   -1  'True
            TabIndex        =   42
            Top             =   720
            Width           =   1455
         End
         Begin VB.TextBox txtOrderMaintinNo 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   4155
            RightToLeft     =   -1  'True
            TabIndex        =   41
            Top             =   780
            Width           =   1380
         End
         Begin VB.Frame Fra_Header 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   570
            Index           =   1
            Left            =   60
            RightToLeft     =   -1  'True
            TabIndex        =   30
            Top             =   0
            Width           =   18540
            Begin VB.TextBox Text2 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H000000FF&
               Height          =   240
               Index           =   1
               Left            =   3030
               RightToLeft     =   -1  'True
               TabIndex        =   35
               Top             =   510
               Visible         =   0   'False
               Width           =   945
            End
            Begin VB.TextBox TxtModFlg2 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H0000FF00&
               Enabled         =   0   'False
               Height          =   285
               Index           =   2
               Left            =   2580
               RightToLeft     =   -1  'True
               TabIndex        =   34
               Text            =   "modflag"
               Top             =   90
               Visible         =   0   'False
               Width           =   465
            End
            Begin VB.Frame Frame2 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   375
               Index           =   1
               Left            =   540
               RightToLeft     =   -1  'True
               TabIndex        =   31
               Top             =   450
               Visible         =   0   'False
               Width           =   3105
               Begin MSDataListLib.DataCombo DataCombo4 
                  CausesValidation=   0   'False
                  Height          =   315
                  Index           =   1
                  Left            =   -255
                  TabIndex        =   32
                  Tag             =   "„‰ ›÷·ﬂ √œŒ· —ﬁ„ «·ﬁ÷Ì…"
                  Top             =   -15
                  Width           =   2340
                  _ExtentX        =   4128
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
                  Index           =   7
                  Left            =   2160
                  RightToLeft     =   -1  'True
                  TabIndex        =   33
                  Top             =   45
                  Width           =   855
               End
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
                     Picture         =   "project_status.frx":C6D1
                     Key             =   "CompanyName"
                  EndProperty
                  BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "project_status.frx":CA6B
                     Key             =   "Ser"
                  EndProperty
                  BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "project_status.frx":CE05
                     Key             =   "Vac_Name"
                  EndProperty
                  BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "project_status.frx":D19F
                     Key             =   "ShareCount"
                  EndProperty
                  BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "project_status.frx":D539
                     Key             =   "Dis_Count"
                  EndProperty
                  BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "project_status.frx":D8D3
                     Key             =   "Bouns"
                  EndProperty
                  BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "project_status.frx":DC6D
                     Key             =   "SharesValue"
                  EndProperty
                  BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "project_status.frx":E207
                     Key             =   "BuyValue"
                  EndProperty
               EndProperty
            End
            Begin ImpulseButton.ISButton btn_Last 
               Height          =   315
               Index           =   2
               Left            =   90
               TabIndex        =   36
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
               ButtonImage     =   "project_status.frx":E5A1
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
               TabIndex        =   37
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
               ButtonImage     =   "project_status.frx":E93B
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btn_Previous 
               Height          =   315
               Index           =   2
               Left            =   1155
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
               ButtonImage     =   "project_status.frx":ECD5
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btn_First 
               Height          =   315
               Index           =   2
               Left            =   1620
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
               ButtonImage     =   "project_status.frx":F06F
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   " ”·Ì„ Ê—œÌ…"
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
               Index           =   8
               Left            =   11970
               RightToLeft     =   -1  'True
               TabIndex        =   40
               Top             =   120
               Width           =   2640
            End
         End
         Begin ImpulseButton.ISButton btn_New 
            Height          =   345
            Index           =   2
            Left            =   9450
            TabIndex        =   43
            Top             =   8730
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   609
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
            ButtonImage     =   "project_status.frx":F409
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Save 
            Height          =   345
            Index           =   2
            Left            =   7590
            TabIndex        =   44
            Top             =   8730
            Width           =   1020
            _ExtentX        =   1799
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
            ButtonImage     =   "project_status.frx":F7A3
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Modify 
            Height          =   345
            Index           =   2
            Left            =   8610
            TabIndex        =   45
            Top             =   8730
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   609
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
            ButtonImage     =   "project_status.frx":FB3D
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton Btn_Undo 
            Height          =   345
            Index           =   2
            Left            =   6555
            TabIndex        =   46
            Top             =   8730
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   609
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
            ButtonImage     =   "project_status.frx":FED7
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Delete 
            Height          =   345
            Index           =   2
            Left            =   5595
            TabIndex        =   47
            Top             =   8730
            Width           =   960
            _ExtentX        =   1693
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
            ButtonImage     =   "project_status.frx":10271
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton Btn_Update 
            Height          =   345
            Index           =   2
            Left            =   7590
            TabIndex        =   48
            TabStop         =   0   'False
            ToolTipText     =   " ÕœÌÀ ﬁ«⁄œ… «·»Ì«‰« "
            Top             =   8340
            Visible         =   0   'False
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   609
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
            ButtonImage     =   "project_status.frx":1080B
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Cancel 
            Height          =   345
            Index           =   2
            Left            =   420
            TabIndex        =   49
            Top             =   8730
            Width           =   900
            _ExtentX        =   1588
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
            ButtonImage     =   "project_status.frx":10BA5
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin MSDataListLib.DataCombo DCboUserName 
            Height          =   315
            Index           =   2
            Left            =   14265
            TabIndex        =   50
            Top             =   8610
            Width           =   2595
            _ExtentX        =   4577
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo Dcbranch 
            Bindings        =   "project_status.frx":10F3F
            Height          =   315
            Index           =   2
            Left            =   7170
            TabIndex        =   51
            Top             =   750
            Width           =   2220
            _ExtentX        =   3916
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
         Begin MSComCtl2.DTPicker XPDtbRecordDate 
            Height          =   420
            Left            =   9930
            TabIndex        =   52
            Top             =   750
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   741
            _Version        =   393216
            Format          =   145817601
            CurrentDate     =   38784
         End
         Begin ImpulseButton.ISButton Btn_Print 
            Height          =   555
            Index           =   2
            Left            =   4275
            TabIndex        =   53
            TabStop         =   0   'False
            ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
            Top             =   8670
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   979
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
            ButtonImage     =   "project_status.frx":10F54
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btn_Query 
            Height          =   600
            Index           =   2
            Left            =   3135
            TabIndex        =   54
            TabStop         =   0   'False
            ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
            Top             =   8640
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   1058
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
            ButtonImage     =   "project_status.frx":177B6
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin MSDataListLib.DataCombo DcboEmpName 
            Height          =   315
            Left            =   3495
            TabIndex        =   147
            Top             =   1290
            Width           =   2520
            _ExtentX        =   4445
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo cmbShiftMaintType 
            Height          =   315
            Left            =   660
            TabIndex        =   153
            Top             =   750
            Width           =   2475
            _ExtentX        =   4366
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSComCtl2.DTPicker txtDateRec 
            Height          =   435
            Left            =   60
            TabIndex        =   162
            Top             =   2160
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   767
            _Version        =   393216
            Format          =   145817601
            CurrentDate     =   38784
         End
         Begin MSComCtl2.DTPicker txtTimeRec 
            Height          =   420
            Left            =   60
            TabIndex        =   163
            Top             =   2610
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   741
            _Version        =   393216
            Format          =   145817602
            CurrentDate     =   38784
         End
         Begin MSComCtl2.DTPicker txtDateEnd 
            Height          =   435
            Left            =   11250
            TabIndex        =   170
            Top             =   8430
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   767
            _Version        =   393216
            Format          =   145817601
            CurrentDate     =   38784
         End
         Begin MSComCtl2.DTPicker txtTimeEnd 
            Height          =   420
            Left            =   11250
            TabIndex        =   172
            Top             =   8850
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   741
            _Version        =   393216
            Format          =   145817602
            CurrentDate     =   38784
         End
         Begin ImpulseButton.ISButton CmdAttach 
            Height          =   375
            Left            =   1800
            TabIndex        =   176
            Top             =   8760
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   661
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
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   -2147483637
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "Õ«·… «·„⁄œÂ/«·”Ì«—…"
            Height          =   285
            Index           =   12
            Left            =   2160
            TabIndex        =   175
            Top             =   1680
            Width           =   1275
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "Êﬁ  «·«‰ Â«¡"
            Height          =   285
            Index           =   11
            Left            =   12645
            TabIndex        =   173
            Top             =   8940
            Width           =   1320
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «·«‰ Â«¡"
            Height          =   300
            Index           =   3
            Left            =   12525
            TabIndex        =   171
            Top             =   8430
            Width           =   1260
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·«„—"
            Height          =   300
            Index           =   50
            Left            =   13605
            TabIndex        =   169
            Top             =   780
            Width           =   720
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·Ê—œÌ…"
            Height          =   270
            Index           =   9
            Left            =   2955
            TabIndex        =   167
            Top             =   780
            Width           =   1020
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "»‰«¡« ⁄·Ï «„— «·‘€·"
            Height          =   270
            Index           =   8
            Left            =   5655
            RightToLeft     =   -1  'True
            TabIndex        =   166
            Top             =   810
            Width           =   1575
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «·«” ·«„"
            Height          =   300
            Index           =   13
            Left            =   1320
            TabIndex        =   165
            Top             =   2160
            Width           =   1275
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "Êﬁ  «·«” ·«„"
            Height          =   285
            Index           =   14
            Left            =   1440
            TabIndex        =   164
            Top             =   2640
            Width           =   1335
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "”»» «· √ŒÌ—"
            Height          =   300
            Index           =   10
            Left            =   2280
            TabIndex        =   160
            Top             =   7470
            Width           =   2895
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„·«ÕŸ« "
            Height          =   300
            Index           =   7
            Left            =   11805
            TabIndex        =   158
            Top             =   7515
            Width           =   2880
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«· Ê’Ì«  »«·’Ì«‰… «·„ »ﬁÌ…"
            Height          =   285
            Index           =   4
            Left            =   2535
            TabIndex        =   155
            Top             =   3960
            Width           =   2940
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„«  „ ⁄„·Â ›Ï «·’Ì«‰…"
            Height          =   285
            Index           =   1
            Left            =   11805
            TabIndex        =   154
            Top             =   3870
            Width           =   2580
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "Õ«·… «·’Ì«‰…"
            Height          =   285
            Index           =   63
            Left            =   2160
            TabIndex        =   152
            Top             =   1305
            Width           =   1275
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "Ê’› «·’Ì«‰…"
            Height          =   300
            Index           =   34
            Left            =   6855
            TabIndex        =   151
            Top             =   1785
            Width           =   1215
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„”ƒÊ· «·’Ì«‰…"
            Height          =   285
            Index           =   0
            Left            =   6915
            TabIndex        =   150
            Top             =   1305
            Width           =   1155
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„·«ÕŸ«  «·ﬁ”„"
            Height          =   570
            Index           =   68
            Left            =   17520
            TabIndex        =   149
            Top             =   3330
            Width           =   840
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„·«ÕŸ«  «·ﬂ‘› «·„»œ∆Ì"
            Height          =   420
            Index           =   69
            Left            =   8910
            TabIndex        =   148
            Top             =   3390
            Width           =   1440
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·”‰œ"
            Height          =   270
            Index           =   5
            Left            =   17280
            RightToLeft     =   -1  'True
            TabIndex        =   62
            Top             =   750
            Width           =   1140
         End
         Begin VB.Label LabCount_Rec 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   225
            Index           =   2
            Left            =   600
            RightToLeft     =   -1  'True
            TabIndex        =   61
            Top             =   8400
            Width           =   540
         End
         Begin VB.Label LabCurr_Rec 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   225
            Index           =   2
            Left            =   2655
            RightToLeft     =   -1  'True
            TabIndex        =   60
            Top             =   8400
            Width           =   840
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «·”Ã·« :"
            Height          =   225
            Index           =   4
            Left            =   1320
            RightToLeft     =   -1  'True
            TabIndex        =   59
            Top             =   8385
            Width           =   1335
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”Ã· «·Õ«·Ì:"
            Height          =   225
            Index           =   5
            Left            =   3495
            RightToLeft     =   -1  'True
            TabIndex        =   58
            Top             =   8385
            Width           =   1260
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Õ—— »Ê«”ÿ…  "
            Height          =   285
            Index           =   6
            Left            =   16980
            TabIndex        =   57
            Top             =   8640
            Width           =   1200
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·›—⁄"
            Height          =   270
            Index           =   16
            Left            =   9150
            TabIndex        =   56
            Top             =   810
            Width           =   960
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒÂ"
            ForeColor       =   &H000000FF&
            Height          =   270
            Index           =   19
            Left            =   11190
            RightToLeft     =   -1  'True
            TabIndex        =   55
            Top             =   810
            Width           =   1095
         End
      End
      Begin C1SizerLibCtl.C1Elastic ELe 
         Height          =   9315
         Index           =   1
         Left            =   -19425
         TabIndex        =   63
         TabStop         =   0   'False
         Top             =   45
         Width           =   18780
         _cx             =   33126
         _cy             =   16431
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
         Begin VB.TextBox txtcolor 
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
            Height          =   330
            Left            =   1680
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   86
            TabStop         =   0   'False
            Top             =   5610
            Visible         =   0   'False
            Width           =   1515
         End
         Begin VB.Frame FraHeader 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   570
            Left            =   60
            RightToLeft     =   -1  'True
            TabIndex        =   74
            Top             =   0
            Width           =   6675
            Begin VB.CommandButton Command2 
               Height          =   255
               Left            =   3000
               RightToLeft     =   -1  'True
               TabIndex        =   80
               Top             =   120
               Visible         =   0   'False
               Width           =   615
            End
            Begin VB.TextBox TxtVac_ID 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H000000FF&
               Height          =   240
               Left            =   3030
               RightToLeft     =   -1  'True
               TabIndex        =   79
               Top             =   510
               Visible         =   0   'False
               Width           =   945
            End
            Begin VB.TextBox TxtModFlg 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H0000FF00&
               Enabled         =   0   'False
               Height          =   285
               Left            =   2580
               RightToLeft     =   -1  'True
               TabIndex        =   78
               Text            =   "modflag"
               Top             =   -150
               Visible         =   0   'False
               Width           =   465
            End
            Begin VB.Frame Frmo2 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   375
               Left            =   540
               RightToLeft     =   -1  'True
               TabIndex        =   75
               Top             =   450
               Visible         =   0   'False
               Width           =   3105
               Begin MSDataListLib.DataCombo DCUser 
                  CausesValidation=   0   'False
                  Height          =   315
                  Left            =   -255
                  TabIndex        =   76
                  Tag             =   "„‰ ›÷·ﬂ √œŒ· —ﬁ„ «·ﬁ÷Ì…"
                  Top             =   15
                  Width           =   2340
                  _ExtentX        =   4128
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
                  Left            =   2160
                  RightToLeft     =   -1  'True
                  TabIndex        =   77
                  Top             =   45
                  Visible         =   0   'False
                  Width           =   855
               End
            End
            Begin MSComctlLib.ImageList GrdImageList 
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
                     Picture         =   "project_status.frx":17B50
                     Key             =   "CompanyName"
                  EndProperty
                  BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "project_status.frx":17EEA
                     Key             =   "Ser"
                  EndProperty
                  BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "project_status.frx":18284
                     Key             =   "Vac_Name"
                  EndProperty
                  BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "project_status.frx":1861E
                     Key             =   "ShareCount"
                  EndProperty
                  BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "project_status.frx":189B8
                     Key             =   "Dis_Count"
                  EndProperty
                  BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "project_status.frx":18D52
                     Key             =   "Bouns"
                  EndProperty
                  BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "project_status.frx":190EC
                     Key             =   "SharesValue"
                  EndProperty
                  BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "project_status.frx":19686
                     Key             =   "BuyValue"
                  EndProperty
               EndProperty
            End
            Begin ImpulseButton.ISButton btnLast 
               Height          =   315
               Left            =   90
               TabIndex        =   81
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
               ButtonImage     =   "project_status.frx":19A20
               ColorButton     =   14871017
               AcclimateGrayTones=   -1  'True
               DrawFocusRectangle=   0   'False
               DisabledImageExtraction=   0
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnNext 
               Height          =   315
               Left            =   555
               TabIndex        =   82
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
               ButtonImage     =   "project_status.frx":19DBA
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnPrevious 
               Height          =   315
               Left            =   1155
               TabIndex        =   83
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
               ButtonImage     =   "project_status.frx":1A154
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnFirst 
               Height          =   315
               Left            =   1620
               TabIndex        =   84
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
               ButtonImage     =   "project_status.frx":1A4EE
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin VB.Label LblTitle 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Õ«·«  «·„‘«—Ì⁄"
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
               Left            =   3015
               RightToLeft     =   -1  'True
               TabIndex        =   85
               Top             =   90
               Width           =   2280
            End
         End
         Begin VB.Frame Frm2 
            BackColor       =   &H00E2E9E9&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   2175
            Left            =   60
            RightToLeft     =   -1  'True
            TabIndex        =   64
            Top             =   4230
            Width           =   6255
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
               Left            =   1395
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   69
               Top             =   960
               Width           =   2760
            End
            Begin VB.ComboBox Combo1 
               Height          =   315
               ItemData        =   "project_status.frx":1A888
               Left            =   2760
               List            =   "project_status.frx":1A89B
               RightToLeft     =   -1  'True
               TabIndex        =   68
               Top             =   1320
               Width           =   1335
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
               Left            =   1395
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   67
               Top             =   645
               Width           =   2760
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
               Left            =   3120
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   66
               Top             =   285
               Width           =   1065
            End
            Begin VB.ComboBox CmbType 
               BackColor       =   &H80000018&
               Height          =   315
               ItemData        =   "project_status.frx":1A8BE
               Left            =   2280
               List            =   "project_status.frx":1A8CE
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   65
               Top             =   3150
               Visible         =   0   'False
               Width           =   1005
            End
            Begin VB.Label lblNameE 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·Õ«·… «‰Ã·Ì“Ì"
               Height          =   285
               Left            =   4200
               RightToLeft     =   -1  'True
               TabIndex        =   73
               Top             =   1080
               Width           =   1290
            End
            Begin VB.Label lblColor 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«··Ê‰"
               Height          =   285
               Left            =   3360
               RightToLeft     =   -1  'True
               TabIndex        =   72
               Top             =   1440
               Width           =   1890
            End
            Begin VB.Label lblNameA 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·Õ«·… ⁄—»Ì"
               Height          =   285
               Left            =   4260
               RightToLeft     =   -1  'True
               TabIndex        =   71
               Top             =   720
               Width           =   1170
            End
            Begin VB.Label lblCode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬂÊœ «·Õ«·…"
               Height          =   195
               Left            =   4305
               RightToLeft     =   -1  'True
               TabIndex        =   70
               Top             =   390
               Width           =   990
            End
         End
         Begin C1SizerLibCtl.C1Elastic EltCont 
            Height          =   1080
            Left            =   0
            TabIndex        =   87
            TabStop         =   0   'False
            Top             =   6420
            Width           =   7770
            _cx             =   13705
            _cy             =   1905
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
            Begin ImpulseButton.ISButton btnNew 
               Height          =   330
               Left            =   4575
               TabIndex        =   88
               Top             =   555
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
               ButtonImage     =   "project_status.frx":1A8E7
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnSave 
               Height          =   330
               Left            =   3030
               TabIndex        =   89
               Top             =   555
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
               ButtonImage     =   "project_status.frx":1AC81
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnModify 
               Height          =   330
               Left            =   3795
               TabIndex        =   90
               Top             =   555
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
               ButtonImage     =   "project_status.frx":1B01B
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton BtnUndo 
               Height          =   330
               Left            =   2265
               TabIndex        =   91
               Top             =   555
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
               ButtonImage     =   "project_status.frx":1B3B5
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnDelete 
               Height          =   330
               Left            =   1500
               TabIndex        =   92
               Top             =   555
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
               ButtonImage     =   "project_status.frx":1B74F
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnQuery 
               Height          =   330
               Left            =   5880
               TabIndex        =   93
               TabStop         =   0   'False
               ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
               Top             =   90
               Visible         =   0   'False
               Width           =   840
               _ExtentX        =   1482
               _ExtentY        =   582
               ButtonStyle     =   1
               ButtonPositionImage=   1
               Caption         =   "»ÕÀ"
               BackColor       =   14737632
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
               ButtonImage     =   "project_status.frx":1BCE9
               ColorButton     =   14737632
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton BtnUpdate 
               Height          =   330
               Left            =   6045
               TabIndex        =   94
               TabStop         =   0   'False
               ToolTipText     =   " ÕœÌÀ ﬁ«⁄œ… «·»Ì«‰« "
               Top             =   105
               Visible         =   0   'False
               Width           =   840
               _ExtentX        =   1482
               _ExtentY        =   582
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
               ButtonImage     =   "project_status.frx":1C083
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton BtnPrint 
               Height          =   285
               Left            =   4725
               TabIndex        =   95
               TabStop         =   0   'False
               Top             =   150
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
               ButtonImage     =   "project_status.frx":1C41D
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnCancel 
               Height          =   330
               Left            =   705
               TabIndex        =   96
               Top             =   555
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
               ButtonImage     =   "project_status.frx":1C7B7
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton ISButton1 
               Height          =   330
               Left            =   0
               TabIndex        =   97
               Top             =   555
               Width           =   390
               _ExtentX        =   688
               _ExtentY        =   582
               ButtonStyle     =   1
               ButtonPositionImage=   1
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
               ButtonImage     =   "project_status.frx":1CB51
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin XtremeSuiteControls.CommonDialog cd1 
               Left            =   600
               Top             =   0
               _Version        =   786432
               _ExtentX        =   423
               _ExtentY        =   423
               _StockProps     =   4
            End
            Begin VB.Label lblCurrent 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·”Ã· «·Õ«·Ì:"
               Height          =   210
               Left            =   2505
               RightToLeft     =   -1  'True
               TabIndex        =   101
               Top             =   225
               Width           =   975
            End
            Begin VB.Label lblCounter 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄œœ «·”Ã·« :"
               Height          =   210
               Left            =   810
               RightToLeft     =   -1  'True
               TabIndex        =   100
               Top             =   225
               Width           =   975
            End
            Begin VB.Label LabCurrRec 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Height          =   210
               Left            =   1800
               RightToLeft     =   -1  'True
               TabIndex        =   99
               Top             =   240
               Width           =   675
            End
            Begin VB.Label LabCountRec 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Height          =   210
               Left            =   240
               RightToLeft     =   -1  'True
               TabIndex        =   98
               Top             =   225
               Width           =   540
            End
         End
         Begin VSFlex8Ctl.VSFlexGrid Grid 
            Height          =   3600
            Left            =   60
            TabIndex        =   102
            Top             =   600
            Width           =   6255
            _cx             =   11033
            _cy             =   6350
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
            Cols            =   5
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"project_status.frx":1CEEB
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
      Begin C1SizerLibCtl.C1Elastic ELe 
         Height          =   9315
         Index           =   3
         Left            =   45
         TabIndex        =   177
         TabStop         =   0   'False
         Top             =   45
         Width           =   18780
         _cx             =   33126
         _cy             =   16431
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
            Height          =   570
            Index           =   2
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   186
            Top             =   0
            Width           =   18840
            Begin VB.TextBox Text2 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H000000FF&
               Height          =   240
               Index           =   2
               Left            =   3030
               RightToLeft     =   -1  'True
               TabIndex        =   191
               Top             =   510
               Visible         =   0   'False
               Width           =   945
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
               TabIndex        =   190
               Text            =   "modflag"
               Top             =   90
               Visible         =   0   'False
               Width           =   465
            End
            Begin VB.Frame Frame2 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   375
               Index           =   2
               Left            =   540
               RightToLeft     =   -1  'True
               TabIndex        =   187
               Top             =   450
               Visible         =   0   'False
               Width           =   3105
               Begin MSDataListLib.DataCombo DataCombo4 
                  CausesValidation=   0   'False
                  Height          =   315
                  Index           =   2
                  Left            =   -255
                  TabIndex        =   188
                  Tag             =   "„‰ ›÷·ﬂ √œŒ· —ﬁ„ «·ﬁ÷Ì…"
                  Top             =   -15
                  Width           =   2340
                  _ExtentX        =   4128
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
                  Index           =   0
                  Left            =   2160
                  RightToLeft     =   -1  'True
                  TabIndex        =   189
                  Top             =   45
                  Width           =   855
               End
            End
            Begin MSComctlLib.ImageList GrdImageList2 
               Index           =   3
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
                     Picture         =   "project_status.frx":1CFAC
                     Key             =   "CompanyName"
                  EndProperty
                  BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "project_status.frx":1D346
                     Key             =   "Ser"
                  EndProperty
                  BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "project_status.frx":1D6E0
                     Key             =   "Vac_Name"
                  EndProperty
                  BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "project_status.frx":1DA7A
                     Key             =   "ShareCount"
                  EndProperty
                  BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "project_status.frx":1DE14
                     Key             =   "Dis_Count"
                  EndProperty
                  BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "project_status.frx":1E1AE
                     Key             =   "Bouns"
                  EndProperty
                  BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "project_status.frx":1E548
                     Key             =   "SharesValue"
                  EndProperty
                  BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "project_status.frx":1EAE2
                     Key             =   "BuyValue"
                  EndProperty
               EndProperty
            End
            Begin ImpulseButton.ISButton btn_Last 
               Height          =   315
               Index           =   3
               Left            =   90
               TabIndex        =   192
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
               ButtonImage     =   "project_status.frx":1EE7C
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
               TabIndex        =   193
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
               ButtonImage     =   "project_status.frx":1F216
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btn_Previous 
               Height          =   315
               Index           =   3
               Left            =   1155
               TabIndex        =   194
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
               ButtonImage     =   "project_status.frx":1F5B0
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btn_First 
               Height          =   315
               Index           =   3
               Left            =   1620
               TabIndex        =   195
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
               ButtonImage     =   "project_status.frx":1F94A
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "»Ì«‰«  «·«ÿ»«¡"
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
               Index           =   1
               Left            =   4980
               RightToLeft     =   -1  'True
               TabIndex        =   196
               Top             =   60
               Width           =   2640
            End
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00E2E9E9&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   2190
            Left            =   60
            RightToLeft     =   -1  'True
            TabIndex        =   178
            Top             =   4320
            Width           =   6255
            Begin VB.TextBox txtPercentV 
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
               Left            =   1380
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   214
               Top             =   1320
               Width           =   2775
            End
            Begin VB.TextBox txtNamee3 
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
               Left            =   1395
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   182
               Top             =   960
               Width           =   2760
            End
            Begin VB.TextBox txtName3 
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
               Left            =   1395
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   181
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
               Index           =   3
               Left            =   3120
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   180
               Top             =   285
               Width           =   1065
            End
            Begin VB.ComboBox Combo2 
               BackColor       =   &H80000018&
               Height          =   315
               ItemData        =   "project_status.frx":1FCE4
               Left            =   2280
               List            =   "project_status.frx":1FCF4
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   179
               Top             =   3150
               Visible         =   0   'False
               Width           =   1005
            End
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·‰”»…"
               Height          =   285
               Index           =   6
               Left            =   4950
               RightToLeft     =   -1  'True
               TabIndex        =   215
               Top             =   1365
               Width           =   780
            End
            Begin VB.Label Label8 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·ÿ»Ì» «‰Ã·Ì“Ï"
               Height          =   285
               Left            =   4200
               RightToLeft     =   -1  'True
               TabIndex        =   185
               Top             =   1080
               Width           =   1500
            End
            Begin VB.Label Label7 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·ÿ»Ì» ⁄—»Ì"
               Height          =   285
               Left            =   4350
               RightToLeft     =   -1  'True
               TabIndex        =   184
               Top             =   720
               Width           =   1350
            End
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬂÊœ «·ÿ»Ì»"
               Height          =   195
               Left            =   4695
               RightToLeft     =   -1  'True
               TabIndex        =   183
               Top             =   390
               Width           =   990
            End
         End
         Begin VSFlex8UCtl.VSFlexGrid FgItems 
            Height          =   9195
            Index           =   0
            Left            =   25950
            TabIndex        =   197
            Top             =   810
            Width           =   18585
            _cx             =   32782
            _cy             =   16219
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
            FormatString    =   $"project_status.frx":1FD0D
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
         Begin ImpulseButton.ISButton btn_New 
            Height          =   345
            Index           =   3
            Left            =   6675
            TabIndex        =   198
            Top             =   8610
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   609
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
            ButtonImage     =   "project_status.frx":1FDCD
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Save 
            Height          =   345
            Index           =   3
            Left            =   4875
            TabIndex        =   199
            Top             =   8610
            Width           =   900
            _ExtentX        =   1588
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
            ButtonImage     =   "project_status.frx":20167
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Modify 
            Height          =   345
            Index           =   3
            Left            =   5775
            TabIndex        =   200
            Top             =   8610
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   609
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
            ButtonImage     =   "project_status.frx":20501
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton Btn_Undo 
            Height          =   345
            Index           =   3
            Left            =   3975
            TabIndex        =   201
            Top             =   8610
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   609
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
            ButtonImage     =   "project_status.frx":2089B
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Delete 
            Height          =   345
            Index           =   3
            Left            =   3135
            TabIndex        =   202
            Top             =   8610
            Width           =   840
            _ExtentX        =   1482
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
            ButtonImage     =   "project_status.frx":20C35
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton Btn_Update 
            Height          =   345
            Index           =   3
            Left            =   5475
            TabIndex        =   203
            TabStop         =   0   'False
            ToolTipText     =   " ÕœÌÀ ﬁ«⁄œ… «·»Ì«‰« "
            Top             =   7830
            Visible         =   0   'False
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   609
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
            ButtonImage     =   "project_status.frx":211CF
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Cancel 
            Height          =   345
            Index           =   3
            Left            =   60
            TabIndex        =   204
            Top             =   8610
            Width           =   900
            _ExtentX        =   1588
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
            ButtonImage     =   "project_status.frx":21569
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton Btn_Print 
            Height          =   540
            Index           =   3
            Left            =   2040
            TabIndex        =   205
            TabStop         =   0   'False
            ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
            Top             =   8520
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   953
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
            ButtonImage     =   "project_status.frx":21903
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btn_Query 
            Height          =   600
            Index           =   3
            Left            =   780
            TabIndex        =   206
            TabStop         =   0   'False
            ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
            Top             =   8475
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   1058
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
            ButtonImage     =   "project_status.frx":28165
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton Cmd_DeleteRow 
            Height          =   300
            Index           =   3
            Left            =   2220
            TabIndex        =   207
            Top             =   7785
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   529
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
            ButtonImage     =   "project_status.frx":284FF
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton Cmd_DeleteAll 
            Height          =   300
            Index           =   3
            Left            =   360
            TabIndex        =   208
            Top             =   7785
            Width           =   1740
            _ExtentX        =   3069
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
            ButtonImage     =   "project_status.frx":28A99
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin VSFlex8Ctl.VSFlexGrid Grid3 
            Height          =   3600
            Left            =   60
            TabIndex        =   209
            Top             =   690
            Width           =   7530
            _cx             =   13282
            _cy             =   6350
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
            FormatString    =   $"project_status.frx":29033
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
            Height          =   210
            Index           =   1
            Left            =   5415
            RightToLeft     =   -1  'True
            TabIndex        =   213
            Top             =   8265
            Width           =   1080
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «·”Ã·« :"
            Height          =   210
            Index           =   0
            Left            =   3495
            RightToLeft     =   -1  'True
            TabIndex        =   212
            Top             =   8265
            Width           =   1080
         End
         Begin VB.Label LabCurr_Rec 
            BackColor       =   &H00E2E9E9&
            Height          =   225
            Index           =   3
            Left            =   4575
            RightToLeft     =   -1  'True
            TabIndex        =   211
            Top             =   8280
            Width           =   780
         End
         Begin VB.Label LabCount_Rec 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   225
            Index           =   3
            Left            =   2775
            RightToLeft     =   -1  'True
            TabIndex        =   210
            Top             =   8280
            Width           =   540
         End
      End
   End
End
Attribute VB_Name = "project_status"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsSavRec As ADODB.Recordset
Dim BKGrndPic As ClsBackGroundPic
Dim RecId As String
Dim II As Long




Public mIndex As Integer
Public LngRow As Long
Public FixedAssetsID As Long
Public FixedAssetsName As String
Dim s As String
Dim mGridClicked As Boolean

Dim AccountVATCreit As String


Private Sub CmdAttach_Click()
    On Error Resume Next
          If DoPremis(Do_Attach, Me.Name, True) = False Then
                Exit Sub
            End If
            
ShowAttachments TxtSerial1(2), "21012020"

End Sub

Private Sub DcbType_Change()
'TxtOrder.Visible = False
''lbl(33).Visible = False
'Frame4.Visible = False
'Frame7.Visible = False
'If val(DcbType.ListIndex) = 0 Then
'Frame4.Visible = True
'TxtOrder.Visible = True
''lbl(33).Visible = True
'Else
'Frame7.Visible = True
'End If
End Sub

Private Sub DcbType_Click()
DcbType_Change
End Sub




 

Private Sub ChDrievType_Click(Index As Integer)
If ChDrievType(0).value = True Then
Text6.Enabled = True
DcbDrievID.Enabled = True
TxtDrievName.Enabled = False
TxtDrievName.Text = ""
ElseIf ChDrievType(1).value = True Then
Text6.Enabled = False
DcbDrievID.Enabled = False
TxtDrievName.Enabled = True
DcbDrievID.BoundText = 0
Text6.Text = ""
End If
End Sub

Private Sub ChLeaderType_Click(Index As Integer)
If ChLeaderType(0).value = True Then
txtNamee.Enabled = True
DcbLeaderID.Enabled = True
TxtLeaderName.Enabled = False
TxtLeaderName.Text = ""
ElseIf ChLeaderType(1).value = True Then
txtNamee.Enabled = False
DcbLeaderID.Enabled = False
TxtLeaderName.Enabled = True
DcbLeaderID.BoundText = 0
txtNamee.Text = ""
End If
End Sub

Private Sub DcbDrievID_Change()
DcbDrievID_Click (0)
End Sub

Private Sub DcbDrievID_Click(Area As Integer)
    If val(DcbDrievID.BoundText) = 0 Then Exit Sub
      Dim EmpCode  As String
      GetEmployeeIDFromCode , , DcbDrievID.BoundText, EmpCode
      Text6.Text = EmpCode
End Sub

Private Sub DcbDrievID_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        FrmEmployeeSearch.lbltype = 41
        FrmEmployeeSearch.show
    End If
End Sub

Public Sub DcbEquepment_Change()
    DcbEquepment_Click (0)
End Sub

Private Sub DcbEquepment_Click(Area As Integer)
On Error Resume Next
RetriveCarsInfo val(DcbEquepment.BoundText), , , 0
If Me.TxtModFlg2(mIndex).Text <> "R" Then
Retrive_CarParts
End If
End Sub

Private Sub Retrive_CarParts()
    Dim i As Integer
    Dim rs_CarParts As ADODB.Recordset
    Set rs_CarParts = New ADODB.Recordset
    Dim StrSQL As String
    
'    StrSQL = " SELECT     dbo.TblCarsDataDet.ID AS PID, dbo.TblCarsDataDet.PartID, dbo.FixedAssets.Name, dbo.FixedAssets.code, dbo.FixedAssets.namee"
'    StrSQL = StrSQL & "  FROM         dbo.TblCarsDataDet LEFT OUTER JOIN"
'    StrSQL = StrSQL & "                  dbo.FixedAssets ON dbo.TblCarsDataDet.PartID = dbo.FixedAssets.id"
'    StrSQL = StrSQL & " Where TblCarsDataDet.EqupID = " & val(Me.DcbEquepment.BoundText) & " "
'
'    rs_CarParts.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
'
'    VSFlexGrid13.Rows = 1
'    If rs_CarParts.RecordCount > 0 Then
'        rs_CarParts.MoveFirst
'        With VSFlexGrid13
'            .Rows = rs_CarParts.RecordCount + 1
'            For i = 1 To .Rows - 1
'                .TextMatrix(i, .ColIndex("Serial")) = i
'                .TextMatrix(i, .ColIndex("ID")) = IIf(IsNull(rs_CarParts("PID").value), 0, rs_CarParts("PID").value)
'                .TextMatrix(i, .ColIndex("PartID")) = IIf(IsNull(rs_CarParts("PartID").value), 0, rs_CarParts("PartID").value)
'                .TextMatrix(i, .ColIndex("Code")) = IIf(IsNull(rs_CarParts("code").value), "", rs_CarParts("code").value)
'                If SystemOptions.UserInterface = ArabicInterface Then
'                .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs_CarParts("Name").value), "", rs_CarParts("Name").value)
'                Else
'                .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs_CarParts("namee").value), "", rs_CarParts("namee").value)
'                End If
'                rs_CarParts.MoveNext
'            Next
'         End With
'    End If
End Sub

Sub RetriveCarsInfo(Optional CarID As Double = 0, Optional OperNo As String, Optional BoardNO As String, Optional Typ As Integer = 0)
If Me.TxtModFlg <> "R" Then
Dim sql As String
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
sql = "select * from TblCarsData"
If Typ = 0 Then
sql = sql & "  Where FixedassetId = " & CarID & ""
ElseIf Typ = 1 Then
sql = sql & " where OperatorN='" & OperNo & "'"
ElseIf Typ = 2 Then
sql = sql & " where BoardNO='" & BoardNO & "'"
End If
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
'Me.TxtLastKM.Text = IIf(IsNull(Rs3("LastKMCounter").value), "", Rs3("LastKMCounter").value)
If Typ <> 1 Then
Me.TxtOperatorN.Text = IIf(IsNull(Rs3("OperatorN").value), "", Rs3("OperatorN").value)
End If
If Typ <> 2 Then
txtBoardNo.Text = IIf(IsNull(Rs3("BoardNO").value), "", Rs3("BoardNO").value)
End If
If Typ <> 0 Then
DcbEquepment.BoundText = IIf(IsNull(Rs3("FixedassetId").value), 0, Rs3("FixedassetId").value)
End If
DcbLeaderID.BoundText = IIf(IsNull(Rs3("Emp_id").value), 0, Rs3("Emp_id").value)
Else
If Typ <> 1 Then
TxtOperatorN.Text = ""
End If
If Typ <> 2 Then
txtBoardNo.Text = ""
End If
If Typ <> 0 Then
DcbEquepment.BoundText = 0
End If
End If
End If
End Sub

Private Sub DcbEquepment_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
         Load FrmCasrShearches
          FrmCasrShearches.SendForm = "OrderMaintin"
          FrmCasrShearches.show vbModal
    End If
End Sub

Private Sub DcbLeaderID_Change()
DcbLeaderID_Click (0)
End Sub

Private Sub DcbLeaderID_Click(Area As Integer)
      If val(DcbLeaderID.BoundText) = 0 Then Exit Sub
      Dim EmpCode  As String
      GetEmployeeIDFromCode , , DcbLeaderID.BoundText, EmpCode
      txtNamee.Text = EmpCode
End Sub

Private Sub DcbLeaderID_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        FrmEmployeeSearch.lbltype = 40
        FrmEmployeeSearch.show
    End If
End Sub

Private Sub DcboEmpName_Change()
DcboEmpName_Click (0)
End Sub

Private Sub DcboEmpName_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        FrmEmployeeSearch.lbltype = 42
        FrmEmployeeSearch.show
    End If
End Sub

 

Private Sub EnterTime_Change()
 

End Sub

Private Sub txtNamee_KeyPress(KeyAscii As Integer)
    Dim EmpID As Integer
    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode txtNamee.Text, EmpID
        DcbLeaderID.BoundText = EmpID
    End If
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
    Dim EmpID As Integer
    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode Text6.Text, EmpID
        DcbDrievID.BoundText = EmpID
    End If
End Sub

Private Sub TxtBoardNO_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
RetriveCarsInfo , , txtBoardNo.Text, 2
End If
End Sub



Private Sub TxtOperatorN_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
RetriveCarsInfo , TxtOperatorN.Text, , 1
End If
End Sub

Private Sub txtOrderMaintinNo_Change()
    If Me.TxtModFlg2(mIndex).Text = "N" Or Me.TxtModFlg2(mIndex).Text = "E" Then
        RetriveOrderNo val(txtOrderMaintinNo), False
    End If
End Sub

Private Sub RetriveOrderNo(ByVal mOrder As Long, Optional ByVal mIsDisplay As Boolean = False)

Dim s As String
Dim rs As New ADODB.Recordset
s = "Select * from TblOrderMaint Where Id = " & mOrder
rs.Open s, Cn, adOpenStatic, adLockReadOnly
If rs.EOF Then
    If Not mIsDisplay Then
        MsgBox "Â–« «·«„— €Ì— „”Ã· ›Ï „·› «·’Ì«‰…"
        Exit Sub
    End If
Else


            

     TxtDeptNotes.Text = IIf(IsNull(rs("DeptNotes").value), "", rs("DeptNotes").value)
     TxtInitialNotes.Text = IIf(IsNull(rs("InitialNotes").value), "", rs("InitialNotes").value)
     
     
    
     Me.DcboEmpName.BoundText = IIf(IsNull(rs("SuperVisor").value), "", rs("SuperVisor").value)
     DcbEquepment.BoundText = IIf(IsNull(rs("EquepID").value), "", rs("EquepID").value)
     'txtRemark.text = IIf(IsNull(rs("Remarks").value), "", rs("Remarks").value)
     DcbType.ListIndex = val(IIf(IsNull(rs("TypeMaint").value), -1, rs("TypeMaint").value))
     TxtDes.Text = IIf(IsNull(rs("Des").value), "", rs("Des").value)
  

     
     DcbStutsMaint.ListIndex = IIf(IsNull(rs("StutsMaint").value), -1, rs("StutsMaint").value)
   
    txtBoardNo.Text = IIf(IsNull(rs("BoardNO").value), "", rs("BoardNO").value)
    TxtOperatorN.Text = IIf(IsNull(rs("OperatorN").value), "", rs("OperatorN").value)
    '******************************************

    
    
    
        
       ''///////////////////////
       ''04 05 2016
       Me.DcbBranchFrom.BoundText = IIf(IsNull(rs("DcbBranchFrom").value), "", rs("DcbBranchFrom").value)
       Me.DcbLeaderID.BoundText = IIf(IsNull(rs("LeaderID").value), "", rs("LeaderID").value)
       Me.TxtLeaderName.Text = IIf(IsNull(rs("LeaderName").value), "", rs("LeaderName").value)
       Me.TxtDrievName.Text = IIf(IsNull(rs("DrievName").value), "", rs("DrievName").value)
       
       Me.DcbDrievID.BoundText = IIf(IsNull(rs("DrievID").value), "", rs("DrievID").value)
       If Not IsNull(rs("LeaderType").value) Then
       If rs("LeaderType").value = 1 Then
       ChLeaderType(1).value = True
       Else
       ChLeaderType(0).value = True
       End If
       Else
       ChLeaderType(0).value = True
       End If
       
        If Not IsNull(rs("DrievType").value) Then
       If rs("DrievType").value = 1 Then
       ChDrievType(1).value = True
       Else
       ChDrievType(0).value = True
       End If
       Else
       ChDrievType(0).value = True
       End If
End If
End Sub

'Private Sub ImgFavorites_Click()
'AddTofaforites Me.name, Me.Caption, Me.Caption
'
'End Sub

Private Sub TxtSearchCode_KeyPress(KeyAscii As Integer)
    Dim EmpID As Integer
    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode TxtSearchCode.Text, EmpID
        DcboEmpName.BoundText = EmpID
    End If
End Sub


Private Sub DcboEmpName_Click(Area As Integer)
On Error Resume Next
      If val(DcboEmpName.BoundText) = 0 Then Exit Sub
    Dim EmpCode  As String
    GetEmployeeIDFromCode , , DcboEmpName.BoundText, EmpCode
    TxtSearchCode.Text = EmpCode
    End Sub
Private Sub txtNum4_KeyPress(KeyAscii As Integer)
txtNum4.Text = ""
If Len(txtNum4.Text) > 0 Then
KeyAscii = 0
End If
If Not (CHR(KeyAscii) >= 0 And CHR(KeyAscii) <= 9) Then
KeyAscii = 0
End If
 

End Sub

Private Sub txtNum3_KeyPress(KeyAscii As Integer)
txtNum3.Text = ""
If Len(txtNum3.Text) > 0 Then
KeyAscii = 0
End If
If Not (CHR(KeyAscii) >= 0 And CHR(KeyAscii) <= 9) Then
KeyAscii = 0
Else
        txtNum4.SetFocus
End If
 
End Sub

Private Sub txtNum2_KeyPress(KeyAscii As Integer)
txtNum2.Text = ""
If Len(txtNum2.Text) > 0 Then
KeyAscii = 0
End If
If Not (CHR(KeyAscii) >= 0 And CHR(KeyAscii) <= 9) Then
KeyAscii = 0
Else
        txtNum3.SetFocus
End If

End Sub

Private Sub txtNum1_KeyPress(KeyAscii As Integer)
txtNum1.Text = ""
If Len(txtNum1.Text) > 0 Then
KeyAscii = 0
End If
If Not (CHR(KeyAscii) >= 0 And CHR(KeyAscii) <= 9) Then
KeyAscii = 0
Else
        txtNum2.SetFocus
End If

End Sub

Private Sub txtLetter4_KeyPress(KeyAscii As Integer)
txtLetter4.Text = ""
If Len(txtLetter4.Text) > 0 Then
KeyAscii = 0
End If
Select Case CHR(KeyAscii)
    Case 0 To 9
        KeyAscii = 0
    Case Else
        txtNum1.SetFocus
End Select

End Sub

Private Sub txtLetter3_KeyPress(KeyAscii As Integer)
txtLetter3.Text = ""
If Len(txtLetter3.Text) > 0 Then
KeyAscii = 0
End If
Select Case CHR(KeyAscii)
    Case 0 To 9
        KeyAscii = 0
    Case Else
        txtLetter4.SetFocus
End Select

End Sub

Private Sub txtLetter2_KeyPress(KeyAscii As Integer)
txtLetter2.Text = ""
If Len(txtLetter2.Text) > 0 Then
KeyAscii = 0
End If
Select Case CHR(KeyAscii)
    Case 0 To 9
        KeyAscii = 0
    Case Else
        txtLetter3.SetFocus
End Select

End Sub

Private Sub txtLetter1_KeyPress(KeyAscii As Integer)
txtLetter1.Text = ""
If Len(txtLetter1.Text) > 0 Then
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

    If TxtVac_ID.Text <> "" Then
        If CheckDelCountry(val(Me.TxtVac_ID.Text)) = False Then
            Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã·...!!!"
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Exit Sub
        End If

        MSGType = MsgBox("Â·  —€» ›Ì Õ–› Â–« «·”Ã·", vbQuestion + vbMsgBoxRtlReading + vbYesNo + vbMsgBoxRight, App.title)

        If MSGType = vbYes Then
            RsSavRec.Find "id=" & val(TxtVac_ID.Text), , adSearchForward, 1
            RsSavRec.delete
            MsgBox " „  ⁄„·Ì… «·Õ–› »‰Ã«Õ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
            '------------------------------ Move Next ---------------------------.
            FillGridWithData
            BtnNext_Click
        End If
    End If

    Exit Sub
ErrTrap:
 
    Select Case Err.Number

        Case -2147217873, -2147467259
            StrMSG = "⁄›Ê« ·« ÌÃÊ“ Õ–› «·”Ã· ·«— »«ÿÂ »»Ì«‰«  √Œ—Ì"
            RsSavRec.CancelUpdate
            MsgBox StrMSG, vbOKOnly + vbMsgBoxRight, App.title
            'clear the ConnectiOn Errors
            Cn.Errors.Clear
    End Select

End Sub

Private Sub BtnFirst_Click()
    On Error GoTo ErrTrap

    If Me.TxtModFlg.Text = "N" Then
        FindRec val(TxtVac_ID.Text)
        Me.TxtModFlg.Text = "R"
    End If

    Dim Msg As String
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
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select

End Sub

Private Sub BtnLast_Click()
    On Error GoTo ErrTrap

    Dim Msg As String

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
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
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

    If TxtVac_ID.Text <> "" Then
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
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
        
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
    TxtModFlg.Text = "N"

    My_SQL = "project_status"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable

    If rs.RecordCount > 0 Then
        TxtSerial.Text = rs.RecordCount + 1
    Else
        TxtSerial.Text = 1
    End If

    rs.Close
    CmbType.ListIndex = 0
    TxtVacName.SetFocus
ErrTrap:
End Sub

Private Sub BtnNext_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
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
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select

End Sub

Private Sub BtnPrevious_Click()
    On Error GoTo ErrTrap
    Dim Msg As String

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
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
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
            If CtrlTxt.Text = "" And CtrlTxt.Tag <> "" And CtrlTxt.Enabled = True Then
                MsgBox CtrlTxt.Tag, vbOKOnly + vbMsgBoxRight, App.title
                CtrlTxt.SetFocus
                Exit Sub
            End If
        End If

    Next

    '------------------------------ check if Empcode exist ----------------------

    StrVacName = IsRecExist("project_status", "name", Trim(TxtVacName.Text), "name", "Vac_ID<>'" & Trim(TxtVac_ID.Text) & "'")

    If StrVacName <> "" Then
        Msg = "·ﬁœ ”»ﬁ  ”ÃÌ· Â–« «·‰Ê⁄ „‰ ﬁ»·"
         
        MsgBox Msg, vbOKOnly + vbMsgBoxRight, App.title
        TxtVacName.SetFocus
    
        Exit Sub

    End If

    ' -------------------------------------- txtmodflg type -------------------
    Select Case Me.TxtModFlg.Text

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
    MsgBox "Â‰«ﬂ Œÿ√ „« ›Ì ≈œŒ«· «·»Ì«‰« ", vbOKOnly + vbMsgBoxRight, App.title

End Sub
 
Private Sub BtnUndo_Click()
    FindRec val(TxtVac_ID.Text)
    Me.TxtModFlg.Text = "R"
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

    MsgBox Msg, vbOKOnly + vbMsgBoxRight + vbInformation, App.title
ErrTrap:
End Sub

Private Sub Command1_Click()
GetHeaders Me
End Sub

Private Sub Combo1_Click()
    Dim Color As String

    Color = 16776960

    Select Case Combo1.ListIndex

        Case 0
            Color = "16776960"

        Case 1
            Color = "8421631"

        Case 2
            Color = "8454016"

        Case 3
            Color = "8454143"

        Case 4
            Color = "12632256"

    End Select

    Combo1.backcolor = Color
    txtcolor.Text = Color
End Sub

Private Sub ChangeLang()
    Me.Caption = "Project Status"
    lblTitle.Caption = Me.Caption
    Dim XPic As IPictureDisp
    Set XPic = Me.btnFirst.ButtonImage
    Set Me.btnFirst.ButtonImage = Me.btnLast.ButtonImage
    Set Me.btnLast.ButtonImage = XPic

    Set XPic = Me.btnPrevious.ButtonImage
    Set Me.btnPrevious.ButtonImage = Me.btnNext.ButtonImage
    Set Me.btnNext.ButtonImage = XPic

    btnNew.Caption = "New"
    btnModify.Caption = "Modify"
    btnSave.Caption = "Save"
    BtnUndo.Caption = "Undo"
    btnDelete.Caption = "Delete"
    btnCancel.Caption = "Exit"

    With Grid
        .TextMatrix(0, .ColIndex("ID")) = " ID"
        .TextMatrix(0, .ColIndex("Name")) = " Arabic Name"
        .TextMatrix(0, .ColIndex("Namee")) = " English Name"
        
        .TextMatrix(0, .ColIndex("Color")) = "Status Color"
 
    End With

   lblcode.Caption = "ID"
   Me.LblNameA.Caption = "Status Name"
    LblNameE.Caption = "English Name"
    
    lblColor.Caption = "Color"
    Combo1.Clear
    Combo1.AddItem "Blue"
    Combo1.AddItem "RED"
    Combo1.AddItem "Green"
    Combo1.AddItem "Yellow"
    Combo1.AddItem "Gray"

    lblCurrent.Caption = "Current rec"
lblCounter.Caption = "Records"
End Sub
Public Function IsControlArrayMember(ctl As Object) As Boolean
    IsControlArrayMember = typename(Me.Controls(ctl.Name)) = "Object"
End Function

 
 
  
Private Sub Form_Load()
Dim Dcombos As ClsDataCombos
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
    
    On Error GoTo ErrTrap
    Dim i As Integer
    Dim My_SQL As String
    
            
    
  
    
    TabMain.TabVisible(1) = False
    TabMain.TabVisible(2) = False
    TabMain.TabVisible(0) = False
    TabMain.TabVisible(3) = False
    
    
    If mIndex = 1 Then
        TabMain.TabVisible(0) = True
        TabMain.CurrTab = 0
    ElseIf mIndex = 2 Then
        TabMain.TabVisible(1) = True
        TabMain.CurrTab = 1
    ElseIf mIndex = 3 Then
        TabMain.TabVisible(2) = True
        TabMain.CurrTab = 2
        Me.Width = Grid.Width + 400
    ElseIf mIndex = 4 Then
        TabMain.TabVisible(3) = True
        TabMain.CurrTab = 3
        Me.Width = Grid3.Width + 400

    
    End If
    If mIndex = 3 Then mIndex = 0
    If mIndex = 4 Then mIndex = 3
    
    
    If mIndex = 0 Then
        My_SQL = "project_status"
        Set BKGrndPic = New ClsBackGroundPic
        Set RsSavRec = New ADODB.Recordset
        RsSavRec.CursorLocation = adUseClient
        RsSavRec.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdTableDirect
        Me.TxtModFlg.Text = "R"
        
       My_SQL = "select UserID,UserName From tblUsers "
        fill_combo DCUser, My_SQL
    
        FillGridWithData
        Set ISButton1.ButtonImage = mdifrmmain.ImgLstTree.ListImages("GridOptions").Picture
        
            With Me.Grid
                .Cell(flexcpPicture, 0, .ColIndex("name")) = Me.GrdImageList.ListImages("Vac_Name").ExtractIcon
                .Cell(flexcpPicture, 0, .ColIndex("Ser")) = Me.GrdImageList.ListImages("Ser").ExtractIcon
        
                For i = 0 To .Cols - 1
                    .Cell(flexcpPictureAlignment, 0, i) = flexPicAlignRightCenter
                Next
           
                .ExtendLastCol = True
                .WallPaper = BKGrndPic.Picture
                .RowHeight(-1) = 300
            End With
        
            BtnFirst_Click
            ShowTip
        
            If OPEN_NEW_SCREEN = True Then
                btnNew_Click
            End If
            Dim StrLogFileName  As String
            StrLogFileName = App.path & "\Titles\" & Me.Name & ".txt"
            If Dir(StrLogFileName) <> "" Then
                  ShowFormtitles Me
            End If
  
    ElseIf mIndex = 1 Then
    
        Me.Caption = "«‰Ê«⁄ «·Ê—œÌ« "
        My_SQL = "ShiftMaintType"
        Set BKGrndPic = New ClsBackGroundPic
        Set RsSavRec = New ADODB.Recordset
        RsSavRec.CursorLocation = adUseClient
        RsSavRec.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdTableDirect
        btn_First_Click (mIndex)
        Me.Width = Grid2.Width + 400
        FillGridWithData2
    ElseIf mIndex = 3 Then
    
        Me.Caption = "»Ì«‰«  «·«ÿ»«¡"
        My_SQL = "tblDoctorsType"
        Set BKGrndPic = New ClsBackGroundPic
        Set RsSavRec = New ADODB.Recordset
        RsSavRec.CursorLocation = adUseClient
        RsSavRec.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdTableDirect
        btn_First_Click (mIndex)
        Me.Width = Grid3.Width + 400
        FillGridWithData3
   
    End If
    
    If mIndex = 2 Then

        Me.Caption = " ”·Ì„ Ê—œÌ…"
        If SystemOptions.UserInterface = EnglishInterface Then
            Me.DcbType.AddItem "Internal"
            Me.DcbType.AddItem "External"
            
  
            With DcbStutsMaint
                .Clear
                .AddItem "Current Reform"
                .AddItem "Ready"
                .AddItem "Exit"
                End With
                With cmbCarStatus
                .Clear
                .AddItem "Inside the workshop"
                .AddItem "Damage"
                .AddItem "In the garage"
                End With
                
        Else
            Me.DcbType.AddItem "œ«Œ·Ì"
            Me.DcbType.AddItem "Œ«—ÃÌ"
                   
            With DcbStutsMaint
                .Clear
                .AddItem "Ã«—Ì «·«’·«Õ"
                .AddItem "Ã«Â“"
                .AddItem "Œ—Ã"
            End With
            
            With cmbCarStatus
                .Clear
                .AddItem "›Ï «·Ê—‘…"
                .AddItem "⁄ÿ· ›Ï «·ÿ—Ìﬁ"
                .AddItem "›Ï «·»«—ﬂÌ‰Ã"
            End With
        End If
        
        
        
             Dim str  As String
       If SystemOptions.UserInterface = ArabicInterface Then
      str = " SELECT     dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.FlagDriver, dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmpJobsTypes.JobTypeNamee, "
    str = str & "                   dbo.TblEmployee.Emp_Namee"
   Else
   str = " SELECT     dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Namee, dbo.TblEmployee.FlagDriver, dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmpJobsTypes.JobTypeNamee, "
    str = str & "                   dbo.TblEmployee.Emp_Name"
   End If
    str = str & "    FROM         dbo.TblEmployee LEFT OUTER JOIN"
    str = str & "                    dbo.TblEmpJobsTypes ON dbo.TblEmployee.JobTypeID = dbo.TblEmpJobsTypes.JobTypeID"
     If SystemOptions.ShowDriverOnly = True Then
        str = str & "     where  ( JobTypeName like '%”«∆ﬁ%'  or JobTypeNamee like '%driver%' )or (FlagDriver=1) "
    End If
    fill_combo DcbLeaderID, str
    fill_combo DcbDrievID, str
    
    
    If SystemOptions.UserInterface = ArabicInterface Then
      str = " SELECT     dbo.ShiftMaintType.ID, dbo.ShiftMaintType.Name"
    Else
      str = " SELECT     dbo.ShiftMaintType.ID, dbo.ShiftMaintType.Namee"
    End If
    str = str & " From ShiftMaintType"
    fill_combo cmbShiftMaintType, str
    
    Set Dcombos = New ClsDataCombos
   ' Dcombos.GetBoxes Me.DcboBox
    Dcombos.GetUsers Me.DCboUserName(mIndex)
     
    Dcombos.GetBranches Me.dcBranch(mIndex)
    Dcombos.GetEquipments DcbEquepment
    Dcombos.GetEmployees Me.DcboEmpName
    'Dcombos.GetEmployees Me.reciverid
    Dcombos.GetEquipments DcbEquepment
    
    
    
        
        
        
    
          '  Dim Dcombos As New ClsDataCombos
            'Dcombos.GetCustomersSuppliers 1, Me.DcbSales, True
            



        
            Resize_Form Me
            
            
            
            SetDtpickerDate Me.XPDtbRecordDate
           
            
            'load tblUsers -----------------------------------------------
            My_SQL = "select UserID,UserName From tblUsers "
            fill_combo DCUser, My_SQL
            fill_combo DCboUserName(mIndex), My_SQL
        
              
             If mIndex = 1 Then
                
          
                
                My_SQL = "ShiftRec"
               ' Set BKGrndPic = New ClsBackGroundPic
                Set RsSavRec = New ADODB.Recordset
                RsSavRec.CursorLocation = adUseClient
                RsSavRec.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdTableDirect
                TxtModFlg2(mIndex).Text = "R"
                DCboUserName(mIndex) = user_id
               
       
                
        
                
        
                btn_First_Click (mIndex)
            ElseIf mIndex = 2 Then
   
        '
                My_SQL = "ShiftRec"
                'Set BKGrndPic2 = New ClsBackGroundPic
                Set RsSavRec = New ADODB.Recordset
                RsSavRec.CursorLocation = adUseClient
                RsSavRec.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdTableDirect
                TxtModFlg2(mIndex).Text = "R"
                DCboUserName(mIndex) = user_id
                btn_First_Click (mIndex)
                

             
            

                

             
            End If
        
            If SystemOptions.UserInterface = EnglishInterface Then
            ShowTipE
                SetInterface Me
                ChangeLang
             Else
             ShowTip
            End If
        
            If OPEN_NEW_SCREEN = True Then
           
                    btn_New_Click (mIndex)
          
                
            End If
        

    
    End If

    

    
    Resize_Form Me
    'load tblUsers -----------------------------------------------
 
    
    
ErrTrap:
End Sub

Private Sub Form_Terminate()
    'Set cSearchDCombo = Nothing
    'Set BKGrndPic = Nothing
    Set FrmVacancy = Nothing

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



Public Sub FiLLRec()
    On Error GoTo ErrTrap

    RsSavRec.Fields("name").value = IIf(TxtVacName.Text <> "", Trim(TxtVacName.Text), Null)
    RsSavRec.Fields("namee").value = IIf(TxtVacNamee.Text <> "", Trim(TxtVacNamee.Text), Null)
    
    RsSavRec.Fields("color").value = IIf(txtcolor.Text <> "", Trim(txtcolor.Text), Null)
    RsSavRec.update
    MsgBox " „  ⁄„·Ì… «·Õ›Ÿ »‰Ã«Õ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
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
    Frm2.Enabled = False
    TxtVac_ID.Text = IIf(IsNull(RsSavRec.Fields("id").value), "", RsSavRec.Fields("id").value)
    TxtVacName.Text = IIf(IsNull(RsSavRec.Fields("name").value), "", RsSavRec.Fields("name").value)
    TxtVacNamee.Text = IIf(IsNull(RsSavRec.Fields("namee").value), "", RsSavRec.Fields("namee").value)
    
    txtcolor.Text = IIf(IsNull(RsSavRec.Fields("color").value), "", RsSavRec.Fields("color").value)

    If txtcolor.Text <> "" Then
        Combo1.backcolor = txtcolor.Text
    End If

    LabCurrRec.Caption = RsSavRec.AbsolutePosition
    LabCountRec.Caption = RsSavRec.RecordCount

    With Grid

        For i = 1 To .Rows - 1

            If Trim(TxtVac_ID.Text) = .TextMatrix(i, .ColIndex("id")) Then
                TxtSerial.Text = .TextMatrix(i, .ColIndex("Ser"))
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


Private Sub ISButton1_Click()
Dim X As Integer
Dim StrLogFileName As String
If SystemOptions.UserInterface = ArabicInterface Then
X = MsgBox("Â·  —Ìœ    ⁄œÌ· «·⁄‰«ÊÌ‰ ", vbInformation + vbYesNoCancel)
Else
X = MsgBox("Change Title  yes/no", vbInformation + vbYesNoCancel)
End If



StrLogFileName = App.path & "\Titles\" & Me.Name & ".txt"
    If Dir(StrLogFileName) = "" Then
             Exit Sub
    End If
    
       If X = vbYes Then
             
            ShellExecute 0&, vbNullString, StrLogFileName, vbNullString, vbNullString, vbNormalFocus
        ElseIf X = vbNo Then
                ShowFormtitles Me
                
        End If
        
End Sub



'Private Sub TxtVacCode_KeyPress(KeyAscii As Integer)
'KeyAscii = DataFormat(ChrOnly, KeyAscii)
'End Sub

Private Sub TxtModFlg_Change()

    If TxtModFlg.Text = "N" Then
        Frm2.Enabled = True
        Me.btnNew.Enabled = False
        btnModify.Enabled = False
        btnDelete.Enabled = False
        Me.btnQuery.Enabled = False
        Grid.Enabled = False
        BtnUndo.Enabled = True
        Me.btnSave.Enabled = True
        BtnUpdate.Enabled = False
        ' btnNext.Enabled = False
        ' btnPrevious.Enabled = False
        ' btnFirst.Enabled = False
        ' btnLast.Enabled = False
    
    ElseIf TxtModFlg.Text = "R" Then
        Frm2.Enabled = False
        Grid.Enabled = True
        btnModify.Enabled = False
        btnDelete.Enabled = False

        If TxtVac_ID.Text <> "" Then
            btnModify.Enabled = True
            btnDelete.Enabled = True
        End If

        BtnUpdate.Enabled = True
        Me.btnQuery.Enabled = True
        Me.btnNew.Enabled = True
        BtnUndo.Enabled = False
        Me.btnSave.Enabled = False
    
        btnNext.Enabled = True
        btnPrevious.Enabled = True
        btnFirst.Enabled = True
        btnLast.Enabled = True
    
    ElseIf TxtModFlg.Text = "E" Then
        Frm2.Enabled = True
        Me.btnNew.Enabled = False
        btnModify.Enabled = False
        btnDelete.Enabled = False
        Me.btnQuery.Enabled = False
        BtnUpdate.Enabled = False
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
    My_SQL = "select * From project_status order by id"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText

    With Me.Grid
        .Rows = 2
        .Clear flexClearScrollable

        If rs.RecordCount > 0 Then
            .Rows = rs.RecordCount + 1
            rs.MoveFirst

            For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("Ser")) = i
             
                .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(rs.Fields("name").value), "", rs.Fields("name").value)
                .TextMatrix(i, .ColIndex("namee")) = IIf(IsNull(rs.Fields("namee").value), "", rs.Fields("namee").value)
                
               
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs.Fields("id").value), "", rs.Fields("id").value)
                .Cell(flexcpBackColor, i, 4, i, 4) = IIf(IsNull(rs.Fields("color").value), "", rs.Fields("color").value)
            
                '    .Cell(flexcpBackColor, .Rows - 1, 1, .Rows - 1, .Cols - 1) = vbYellow
                rs.MoveNext
            Next

            rs.Close
        End If

        .RowHeight(-1) = 300
    End With

ErrTrap:
End Sub


Public Sub FillGridWithData3()

    On Error GoTo ErrTrap

    Dim i As Integer
    Dim rs As ADODB.Recordset
    Dim My_SQL As String

    Set rs = New ADODB.Recordset
    My_SQL = "select * From tblDoctorsType order by id"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText

    With Me.Grid3
        .Rows = 2
        .Clear flexClearScrollable

        If rs.RecordCount > 0 Then
            .Rows = rs.RecordCount + 1
            rs.MoveFirst

            For i = 1 To .Rows - 1
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



Public Sub FillGridWithData2()

    On Error GoTo ErrTrap

    Dim i As Integer
    Dim rs As ADODB.Recordset
    Dim My_SQL As String

    Set rs = New ADODB.Recordset
    My_SQL = "select * From ShiftMaintType order by id"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText

    With Me.Grid2
        .Rows = 2
        .Clear flexClearScrollable

        If rs.RecordCount > 0 Then
            .Rows = rs.RecordCount + 1
            rs.MoveFirst

            For i = 1 To .Rows - 1
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

'-------------------------------------------------------------
Private Sub ShowTip()
    On Error GoTo ErrTrap
    Dim TTP As New clstooltip
    Dim Wrap As String
    Dim Msg As String
    Wrap = CHR(13) + CHR(10)

    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "ÃœÌœ" & Wrap & "·› Õ ”Ã· ÃœÌœ " & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F12 √Ê Enter"
            
        .AddControl btnNew, Msg, True
    End With

    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = " ⁄œÌ·" & Wrap & "· ⁄œÌ·  ”Ã· «·Õ«·Ï " & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F11"
        .AddControl btnModify, Msg, True
    End With

    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Õ›Ÿ" & Wrap & "· ”ÃÌ· «·»Ì«‰«  œ«Œ· ﬁ«⁄œ… " & Wrap & "«·»Ì«‰«  ≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F10"
        .AddControl btnSave, Msg, True
    End With

    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = " —«Ã⁄" & Wrap & "·· —«Ã⁄ ⁄‰ «·⁄„·Ì… «·Õ«·Ì…" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F9"
        .AddControl BtnUndo, Msg, True
    End With

    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Õ–› «·”Ã·" & Wrap & "·Õ–› «·”Ã· «·Õ«·Ï" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " F18"
        .AddControl btnDelete, Msg, True
    End With

    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Œ—ÊÃ" & Wrap & "·≈€·«ﬁ Â–Â «·‰«›–…" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " Ctrl+x"
        .AddControl btnCancel, Msg, True
    End With

    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«·«Ê·" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·«Ê·" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " Home √Ê UpArrow"
        .AddControl btnFirst, Msg, True
    End With

    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«·”«»ﬁ" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·”«»ﬁ" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " PageUp √Ê LeftArrow"
        .AddControl btnPrevious, Msg, True
    End With

    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«· «·Ï" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «· «·Ï" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " PageDown √Ê RightArrow"
        .AddControl btnNext, Msg, True
    End With

    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "«·«ŒÌ—" & Wrap & "··«‰ ﬁ«· «·Ï «·”Ã· «·«ŒÌ—" & Wrap & "≈÷€ÿ Â–« «·„› «Õ" & Wrap & "√Ê „› «Õ" & " End √Ê DownArrow"
        .AddControl btnLast, Msg, True
    End With

ErrTrap:
End Sub








 





Private Sub Cmd_DeleteRow_Click(Index As Integer)
If Me.TxtModFlg.Text <> "R" Then
If Index = 3 Then
    RemoveGridRow4
Else
    RemoveGridRow
End If


End If
End Sub
Private Sub RemoveGridRow()

     
End Sub

Private Sub RemoveGridRow4()

    With Me.Grid2
'MsgBox .Row
        If .Row <= 0 Then
                .Rows = 2
        Exit Sub
        Else
        .RemoveItem .Row
        End If
    End With
End Sub

Private Sub btn_Cancel_Click(Index As Integer)

   Unload Me
End Sub

Private Sub btn_Delete_Click(Index As Integer)
   Dim MSGType As Integer
   Dim StrSQL As String
    Dim BlnRecExist As Boolean
    Dim StrMSG  As String
    Dim Msg As String
    On Error GoTo ErrTrap
    'Index = TabMain.CurrTab
    'If DoPremis(Do_Delete, Me.name, True) = False Then
    '    Exit Sub
    'End If
    If TxtSerial1(mIndex).Text <> "" Then
        '    If CheckDelCountry(Val(Me.TxtVac_ID.text)) = False Then
        '        Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã·...!!!"
        '        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        '        Exit Sub
        '    End If
        If SystemOptions.UserInterface = ArabicInterface Then
        MSGType = MsgBox("Â·  —€» ›Ì Õ–› Â–« «·”Ã·", vbQuestion + vbMsgBoxRtlReading + vbYesNo + vbMsgBoxRight, App.title)
        Else
        MSGType = MsgBox("Confirm Delete", vbQuestion + vbMsgBoxRtlReading + vbYesNo + vbMsgBoxRight, App.title)
        End If

        If MSGType = vbYes Then
            RsSavRec.Find "id=" & val(TxtSerial1(mIndex).Text), , adSearchForward, 1
            CuurentLogdata ("D")
            RsSavRec.delete
            Dim s As String

            If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox " „  ⁄„·Ì… «·Õ–› »‰Ã«Õ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
            Else
            MsgBox "Delete Successfully", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
            End If
            '------------------------------ Move Next ---------------------------.
            FillGridWithData2
            FillGridWithData3
            btn_Next_Click mIndex
        End If
    End If

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
            MsgBox StrMSG, vbOKOnly + vbMsgBoxRight, App.title
            'clear the ConnectiOn Errors
            Cn.Errors.Clear
    End Select

End Sub

Private Sub btn_First_Click(Index As Integer)
  On Error GoTo ErrTrap

    Dim Msg As String

  
    If Me.TxtModFlg2(mIndex).Text = "N" Then
        FindRec val(TxtSerial1(mIndex).Text)
        TxtModFlg2(mIndex).Text = "R"
    End If

    TxtModFlg2(mIndex) = "R"

    If RsSavRec.RecordCount = 0 Then
        clear_all Me
        Exit Sub
    End If

BegnieWork:
    RsSavRec.MoveFirst
    If mIndex = 1 Then
        FiLLTXT1
    ElseIf mIndex = 2 Then
        FiLLTXT2
    ElseIf mIndex = 3 Then
        FiLLTXT3

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
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub

Private Sub btn_Modify_Click(Index As Integer)
    Dim Msg As String

    If DoPremis(Do_Edit, Me.Name, True) = False Then
        Exit Sub
    End If

    On Error GoTo ErrTrap

    If TxtSerial1(mIndex).Text <> "" Then
        TxtModFlg2(mIndex) = "E"
    
        Frm2.Enabled = True
        
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
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
        
            If RsSavRec.EditMode <> adEditNone Then
                RsSavRec.CancelUpdate
                'RsSavRec.Requery
            End If

    End Select

End Sub

Public Sub btn_New_Click(Index As Integer)
    Dim My_SQL As String
    Dim rs As ADODB.Recordset

    If DoPremis(Do_New, Me.Name, True) = False Then
        Exit Sub
    End If

    On Error GoTo ErrTrap
    Set rs = New ADODB.Recordset
    Frame1.Enabled = True
    Frame3.Enabled = True
    clear_all Me
    TxtModFlg2(mIndex).Text = "N"
    If mIndex = 1 Then
        My_SQL = "ShiftMaintType"
        'DCboUserName(mIndex) = user_id
     
         clear_all Me
    

   
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable

    If rs.RecordCount > 0 Then
        TxtSerial1(mIndex).Text = rs.RecordCount + 1
    Else
        TxtSerial1(mIndex).Text = 1
    End If

    rs.Close
    'CmbType.ListIndex = 0
    txtName.SetFocus
        
    
    ElseIf mIndex = 2 Then
       
        DCboUserName(mIndex) = user_id
        My_SQL = "ShiftRec"
        'DCboUserName(mIndex) = user_id
     
         clear_all Me
    

   
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable

    If rs.RecordCount > 0 Then
        TxtSerial1(mIndex).Text = rs.RecordCount + 1
    Else
        TxtSerial1(mIndex).Text = 1
    End If

    rs.Close
        
    ElseIf mIndex = 3 Then
   My_SQL = "tblDoctorsType"
        'DCboUserName(mIndex) = user_id
     
         clear_all Me
    

   
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable

    If rs.RecordCount > 0 Then
        TxtSerial1(mIndex).Text = rs.RecordCount + 1
    Else
        TxtSerial1(mIndex).Text = 1
    End If

    rs.Close
    'CmbType.ListIndex = 0
    TxtName3.SetFocus
        
        
    End If
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

    If Me.TxtModFlg2(mIndex).Text = "N" Then
        FindRec val(TxtSerial1(mIndex).Text)
        TxtModFlg2(mIndex).Text = "R"
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

    If mIndex = 1 Then
        FiLLTXT1
    ElseIf mIndex = 2 Then
        FiLLTXT2
    ElseIf mIndex = 3 Then
        FiLLTXT3

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
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub

Private Sub btn_Previous_Click(Index As Integer)
  On Error GoTo ErrTrap
    Dim Msg As String

    If TxtModFlg2(mIndex).Text = "N" Then
        FindRec val(TxtSerial1(mIndex).Text)
        TxtModFlg2(mIndex).Text = "R"
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

    If mIndex = 1 Then
        FiLLTXT1
    ElseIf mIndex = 2 Then
        FiLLTXT2
    ElseIf mIndex = 3 Then
        FiLLTXT3

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
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub

Private Sub Btn_Print_Click(Index As Integer)
 Exit Sub
  If mIndex = 1 Or mIndex = 2 Then
    
    PrintRercord
ElseIf mIndex = 3 Then
    PrintRercord2
End If
End Sub

Private Sub PrintRercord2()
  Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
 
'--------------------------------------------------------------------------------------------
   
   
    Dim mTableName As String, mTableName2 As String
   
  
    
    
     s = "Select TblProcessDEF.ProcessName GroupName,Det.SalPrice Price,Det.InstallPrice widtj,Det.TotalSalPrice TotalDisc,"
    s = s & "    Det.TotalSalPrice TotalDisc,TT.NoteSerial1 NoteSerial11,TT.RecordDate,TT.TradingContractID  TransactionID1, "
    s = s & "     Det.TotalInstallPrice TotalAdd,Det.BDet_Qun Qty1,"
    s = s & "     Det.Total TotalWithVat,"
    s = s & "     Det.Vatyo hight,"
    s = s & "     Det.Vat2,Det.TotalNet Net"
    
    s = s & "  "
    s = s & " From Tbl_TradingContractInv  TT Inner Join Tbl_TradingContractInvDet Det On Det.MasterID = TT.ID"
    s = s & "  Left Outer join TblProcessDEF On TblProcessDEF.TblProcessDEFID =Det.TConID "
    s = s & " Where TT.ID = " & val(TxtSerial1(mIndex))

      
   
        
        
         
            If SystemOptions.UserInterface = ArabicInterface Then
                StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "TradingContractInv.rpt"
            Else
                StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "TradingContractInv.rpt"
            End If
       
    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open s, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        'GetMsgs 138, vbExclamation
        If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        Else
        Msg = "No Data"
        End If
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        RsData.Close
        Set RsData = Nothing
        Screen.MousePointer = vbDefault
        Exit Sub
    End If

    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)
    xReport.Database.SetDataSource RsData

    Dim cCompanyInfo As New ClsCompanyInfo

    If SystemOptions.UserInterface = ArabicInterface Then
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
        ' xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
        StrReportTitle = "" '& StrAccountName

    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
      '  xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        
        StrReportTitle = ""
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name

 
        xReport.ParameterFields(7).AddCurrentValue dcBranch(mIndex).Text
        
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.title
    xReport.ReportAuthor = App.title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault

End Sub


Private Sub PrintRercord()
 
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
            If CtrlTxt.Text = "" And CtrlTxt.Tag <> "" And CtrlTxt.Enabled = True Then
                MsgBox CtrlTxt.Tag, vbOKOnly + vbMsgBoxRight, App.title
                CtrlTxt.SetFocus
                Exit Sub
            End If
        End If

    Next


If mIndex = 2 Then
     If dcBranch(mIndex).Text = "" Then
        MsgBox "Please Enter Branch"
        dcBranch(mIndex).SetFocus
        Exit Sub
    End If
End If
    
    '------------------------------ check if Empcode exist ----------------------

   

    ' -------------------------------------- txtmodflg type -------------------
    Select Case TxtModFlg2(mIndex).Text

            '------------------------------ new record ----------------------------
        Case "N"
      
            '------------------------- save record -----------------------------
            If mIndex = 1 Then
                AddNewRec
            ElseIf mIndex = 2 Then
                FiLLRec2
            ElseIf mIndex = 3 Then
                AddNewRec
                
            End If
            If mIndex = 0 Then
                BtnLast_Click
          
                btn_Last_Click CInt(mIndex)
            End If

        Case "E"

            '----------------------------- save edit -------------------------------
            
            If mIndex = 1 Then
                FiLLRec1
            ElseIf mIndex = 2 Then
                FiLLRec2
            ElseIf mIndex = 3 Then
                FiLLRec3

            End If
    End Select

    Exit Sub
ErrTrap:
If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Â‰«ﬂ Œÿ√ „« ›Ì ≈œŒ«· «·»Ì«‰« ", vbOKOnly + vbMsgBoxRight, App.title
 Else
  MsgBox "Sorry...error douring insert data", vbOKOnly + vbMsgBoxRight, App.title
End If
 
End Sub

Private Sub Btn_Undo_Click(Index As Integer)
    Undo
End Sub

 

 






Private Sub TabMain_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)

'If NewTab = 4 Then
'    mIndex = 3
'Else
'    mIndex = NewTab
'End If
End Sub

Private Sub TabMain_Validate(Cancel As Boolean)
'If TabMain.CurrTab = 4 Then
'    mIndex = 3
'Else
'
'    mIndex = TabMain.CurrTab
'End If
End Sub
Private Sub btn_Last_Click(Index As Integer)
  On Error GoTo ErrTrap

    Dim Msg As String

    If Me.TxtModFlg.Text = "N" Then
        FindRec val(TxtSerial1(mIndex).Text)
        Me.TxtModFlg2(mIndex).Text = "R"
    End If

    Me.TxtModFlg2(mIndex) = "R"

    If RsSavRec.RecordCount = 0 Then
        clear_all Me
        Exit Sub
    End If

BegnieWork:

    RsSavRec.MoveLast
    If mIndex = 1 Then
        FiLLTXT1
    
    ElseIf mIndex = 2 Then
        FiLLTXT2
    ElseIf mIndex = 3 Then
        FiLLTXT3
    
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
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select
End Sub





 
Private Sub Undo()
    On Error GoTo ErrTrap

    Select Case TxtModFlg2(mIndex).Text

        Case "N"
            clear_all Me
            TxtModFlg2(mIndex).Text = "R"
           
            btn_First_Click (mIndex)
        Case "E"
            RsSavRec.Find "ID='" & val(TxtSerial1(mIndex).Text) & "'", , adSearchForward, adBookmarkFirst

            If RsSavRec.EOF Or RsSavRec.BOF Then
                TxtModFlg2(mIndex).Text = "R"
                Exit Sub
            End If

            If mIndex = 1 Then
                FiLLTXT1
            ElseIf mIndex = 2 Then
                FiLLTXT2
            ElseIf mIndex = 3 Then
                FiLLTXT3

            End If
            TxtModFlg2(mIndex).Text = "R"
    End Select

    Exit Sub
ErrTrap:
End Sub
 
 

Function CuurentLogdata(Optional Currentmode As String)
     LogTextA = "    ‘«‘… " & ScreenNameArabic & CHR(13) & "ﬂÊœ " & TxtSerial.Text & CHR(13) & "   «·‰Ê⁄ " & TxtVacName
        LogTexte = "    Screen  " & ScreenNameEnglish & CHR(13) & "Code  " & TxtSerial.Text & CHR(13) & "   Type " & TxtVacName
       If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, Me.TxtModFlg
    Else
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, "D"
    End If
    
End Function


Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)
    Dim IntResult As String
    Dim StrMSG As String
    On Error GoTo ErrTrap
    Dim mSelectModFlg As String
    If mIndex = 0 Then
        mSelectModFlg = Me.TxtModFlg.Text
    Else
        mSelectModFlg = Me.TxtModFlg2(mIndex).Text
    End If
    
    
    If mSelectModFlg <> "R" Then

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

        IntResult = MsgBox(StrMSG, vbMsgBoxRight + vbYesNoCancel + vbMsgBoxRtlReading + vbQuestion, App.title)

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

    End If

    Exit Sub
ErrTrap:

End Sub





Public Sub AddNewRec()
    On Error GoTo ErrTrap
    Dim StrRecID As String
   

    If mIndex = 2 Then
        StrRecID = new_id("ShiftRec", "id", "")
    ElseIf mIndex = 1 Then

        
        StrRecID = new_id("ShiftMaintType", "id", "")
        RsSavRec.AddNew
        RsSavRec.Fields("id").value = IIf(StrRecID <> "", StrRecID, Null)
        FiLLRec1
    ElseIf mIndex = 3 Then

        
        StrRecID = new_id("tblDoctorsType", "id", "")
        RsSavRec.AddNew
        RsSavRec.Fields("id").value = IIf(StrRecID <> "", StrRecID, Null)
        FiLLRec3




        
    End If
    
    
ErrTrap:
   
  
    

End Sub




Public Sub FiLLRec1()
    On Error GoTo ErrTrap

    RsSavRec.Fields("name").value = IIf(txtName.Text <> "", Trim(txtName.Text), Null)
    RsSavRec.Fields("namee").value = IIf(txtNamee.Text <> "", Trim(txtNamee.Text), Null)
    

    RsSavRec.update
    MsgBox " „  ⁄„·Ì… «·Õ›Ÿ »‰Ã«Õ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    FillGridWithData2
    TxtModFlg2(mIndex) = "R"

    Exit Sub
ErrTrap:

    If RsSavRec.EditMode <> adEditNone Then
        RsSavRec.CancelUpdate
    End If

End Sub

Public Sub FiLLRec3()
    On Error GoTo ErrTrap

    RsSavRec.Fields("name").value = IIf(TxtName3.Text <> "", Trim(TxtName3.Text), Null)
    RsSavRec.Fields("namee").value = IIf(txtNamee3.Text <> "", Trim(txtNamee3.Text), Null)
    RsSavRec.Fields("PercentV").value = IIf(txtPercentV.Text <> "", Trim(txtPercentV.Text), Null)
    

    RsSavRec.update
    MsgBox " „  ⁄„·Ì… «·Õ›Ÿ »‰Ã«Õ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    FillGridWithData3
    TxtModFlg2(mIndex) = "R"

    Exit Sub
ErrTrap:

    If RsSavRec.EditMode <> adEditNone Then
        RsSavRec.CancelUpdate
    End If

End Sub





Public Sub FiLLRec2()
    On Error GoTo ErrTrap
    If TxtModFlg2(mIndex).Text = "N" Then
           ' TxtSerial1(mIndex).Text = CStr(new_id("TblAdditionsAssest", "ID", "", True))

       
            RsSavRec.AddNew
            TxtSerial1(mIndex).Text = new_id("ShiftRec", "id", "")
            RsSavRec.Fields("ID").value = val(TxtSerial1(mIndex).Text)
    End If
    RsSavRec.Fields("ShiftMaintTypeID").value = IIf(cmbShiftMaintType.Text <> "", Trim(cmbShiftMaintType.BoundText), Null)
    RsSavRec.Fields("BranchID").value = IIf(dcBranch(mIndex).Text <> "", Trim(dcBranch(mIndex).BoundText), Null)
    RsSavRec("RecordDate").value = XPDtbRecordDate.value
   
     RsSavRec("DateRec").value = txtDateRec.value
    RsSavRec("TimeRec").value = txtTimeRec.value
    
     RsSavRec("DateEnd").value = txtDateEnd.value
    RsSavRec("TimeEnd").value = txtTimeEnd.value
    

               
   RsSavRec("OrderMaintinNo").value = val(txtOrderMaintinNo.Text)
   RsSavRec("typemaint").value = DcbType.ListIndex
   
    RsSavRec("NoteDone").value = txtNoteDone.Text
     
    RsSavRec("NoteStill").value = txtNoteStill.Text
    RsSavRec("Remarks").value = txtRemarks.Text
    RsSavRec("NoteLate").value = txtNoteLate.Text
    RsSavRec("CarStatus").value = cmbCarStatus.ListIndex
    RsSavRec("StutsMaint").value = DcbStutsMaint.ListIndex

    
    DCboUserName(mIndex).BoundText = IIf(DCboUserName(mIndex).Text <> "", Trim(DCboUserName(mIndex).BoundText), user_id)
    RsSavRec.Fields("UserID").value = IIf(DCboUserName(mIndex).Text <> "", Trim(DCboUserName(mIndex).BoundText), user_id)
     
   

    RsSavRec.update
    
  
    
   
If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox " „  ⁄„·Ì… «·Õ›Ÿ »‰Ã«Õ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
 Else
   MsgBox "Save Successfully", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
End If
    CuurentLogdata
    
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
    Frm2.Enabled = False
    TxtSerial1(mIndex).Text = IIf(IsNull(RsSavRec.Fields("id").value), "", RsSavRec.Fields("id").value)
    txtName.Text = IIf(IsNull(RsSavRec.Fields("name").value), "", RsSavRec.Fields("name").value)
    txtNamee.Text = IIf(IsNull(RsSavRec.Fields("nameE").value), "", RsSavRec.Fields("nameE").value)
    LabCurr_Rec(mIndex).Caption = RsSavRec.AbsolutePosition
    LabCount_Rec(mIndex).Caption = RsSavRec.RecordCount

    With Grid2

        For i = 1 To .Rows - 1

            If Trim(TxtSerial1(mIndex).Text) = .TextMatrix(i, .ColIndex("id")) Then
                TxtSerial1(mIndex).Text = .TextMatrix(i, .ColIndex("Ser"))
                .Row = i
                Exit Sub
            End If

        Next

    End With

ErrTrap:

End Sub

Public Sub FiLLTXT3()

    On Error GoTo ErrTrap
    Dim i As Integer
    Frm2.Enabled = False
    TxtSerial1(mIndex).Text = IIf(IsNull(RsSavRec.Fields("id").value), "", RsSavRec.Fields("id").value)
    TxtName3.Text = IIf(IsNull(RsSavRec.Fields("name").value), "", RsSavRec.Fields("name").value)
    txtNamee3.Text = IIf(IsNull(RsSavRec.Fields("nameE").value), "", RsSavRec.Fields("nameE").value)
    txtPercentV.Text = IIf(IsNull(RsSavRec.Fields("PercentV").value), "", RsSavRec.Fields("PercentV").value)
    
    LabCurr_Rec(mIndex).Caption = RsSavRec.AbsolutePosition
    LabCount_Rec(mIndex).Caption = RsSavRec.RecordCount

    With Grid3

        For i = 1 To .Rows - 1

            If Trim(TxtSerial1(mIndex).Text) = .TextMatrix(i, .ColIndex("id")) Then
                TxtSerial1(mIndex).Text = .TextMatrix(i, .ColIndex("Ser"))
                .Row = i
                Exit Sub
            End If

        Next

    End With

ErrTrap:

End Sub

Public Sub FiLLTXT2(Optional Lngid As Long = 0)

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
  
    'Frm2.Enabled = False
     
    

    
    DCboUserName(mIndex).BoundText = IIf(DCboUserName(mIndex).Text <> "", Trim(DCboUserName(mIndex).BoundText), user_id)
    RsSavRec.Fields("UserID").value = IIf(DCboUserName(mIndex).Text <> "", Trim(DCboUserName(mIndex).BoundText), user_id)
      
     
     
    TxtSerial1(mIndex).Text = IIf(IsNull(RsSavRec.Fields("id").value), "", RsSavRec.Fields("id").value)
    XPDtbRecordDate.value = IIf(IsNull(RsSavRec("RecordDate").value), Date, RsSavRec("RecordDate").value)
    txtDateRec.value = IIf(IsNull(RsSavRec("DateRec").value), Date, RsSavRec("DateRec").value)
    
         
txtDateEnd.value = IIf(IsNull(RsSavRec("DateEnd").value), Date, RsSavRec("DateEnd").value)


    cmbCarStatus.ListIndex = IIf(IsNull(RsSavRec("CarStatus").value), -1, RsSavRec("CarStatus").value)
    DcbStutsMaint.ListIndex = IIf(IsNull(RsSavRec("StutsMaint").value), -1, RsSavRec("StutsMaint").value)
        
    
         Dim startmaintenanceTime As Date
    If Not IsNull(RsSavRec("TimeEnd").value) Then
         startmaintenanceTime = FormatDateTime(RsSavRec("TimeEnd").value, vbShortTime)
         Me.txtTimeEnd.value = startmaintenanceTime
    End If
    
    cmbShiftMaintType.BoundText = IIf(IsNull(RsSavRec("ShiftMaintTypeID").value), "", RsSavRec("ShiftMaintTypeID").value)
'    txtTel = IIf(IsNull(RsSavRec("CustTel").value), "", RsSavRec("CustTel").value)
    dcBranch(mIndex).BoundText = IIf(IsNull(RsSavRec("BranchID").value), "", RsSavRec("BranchID").value)
    'TxtRemarks = IIf(IsNull(RsSavRec("Remarks").value), "", RsSavRec("Remarks").value)
    txtOrderMaintinNo = IIf(IsNull(RsSavRec("OrderMaintinNo").value), "", RsSavRec("OrderMaintinNo").value)
    
    DcbType.ListIndex = IIf(IsNull(RsSavRec("typemaint").value), -1, RsSavRec("typemaint").value)
  
    txtNoteLate = IIf(IsNull(RsSavRec("NoteLate").value), "", RsSavRec("NoteLate").value)
    
    txtNoteStill = IIf(IsNull(RsSavRec("NoteStill").value), "", RsSavRec("NoteStill").value)
    txtNoteDone = IIf(IsNull(RsSavRec("NoteDone").value), "", RsSavRec("NoteDone").value)
    txtRemarks.Text = IIf(IsNull(RsSavRec("Remarks").value), "", RsSavRec("Remarks").value)
    Me.DCboUserName(mIndex).BoundText = IIf(IsNull(RsSavRec("UserID").value), "", RsSavRec("UserID").value)
   
    
    
    If Not IsNull(RsSavRec("TimeRec").value) Then
         startmaintenanceTime = FormatDateTime(RsSavRec("TimeRec").value, vbShortTime)
         Me.txtTimeRec.value = startmaintenanceTime
    End If
    RetriveOrderNo val(txtOrderMaintinNo), True
    LabCurr_Rec(mIndex).Caption = RsSavRec.AbsolutePosition
    LabCount_Rec(mIndex).Caption = RsSavRec.RecordCount

   

ErrTrap:

End Sub





Private Sub Grid_EnterCell()
    On Error GoTo ErrTrap
    FindRec val(Me.Grid.TextMatrix(Me.Grid.Row, Me.Grid.ColIndex("id")))
ErrTrap:
End Sub
Private Sub Grid2_EnterCell()
    On Error GoTo ErrTrap
    FindRec val(Me.Grid2.TextMatrix(Me.Grid2.Row, Me.Grid2.ColIndex("id")))
ErrTrap:
End Sub

Private Sub Grid3_EnterCell()
    On Error GoTo ErrTrap
    FindRec val(Me.Grid3.TextMatrix(Me.Grid3.Row, Me.Grid3.ColIndex("id")))
ErrTrap:
End Sub

Private Sub TxtModFlg2_Change(Index As Integer)
 On Error GoTo ErrTrap

    Select Case Me.TxtModFlg2(mIndex).Text

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
      

           ' XPDtbTrans.Enabled = True
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub TxtVac_ID_Change()
    Dim TxtMod As String
    TxtMod = TxtModFlg.Text
    TxtModFlg.Text = ""
    TxtModFlg = TxtMod
End Sub

Public Function FindRec(ByVal RecId As Long, Optional ByVal mIndex2 As Integer = 0)
    On Error GoTo ErrTrap
    RsSavRec.Find "id=" & RecId, , adSearchForward, 1
    If mIndex2 = 0 Then mIndex2 = mIndex
    If Not (RsSavRec.EOF) Then
        If mIndex2 = 0 Then
            FiLLTXT
        ElseIf mIndex2 = 1 Then
            FiLLTXT1
        ElseIf mIndex2 = 2 Then
            FiLLTXT2
        ElseIf mIndex2 = 3 Then
            FiLLTXT3


        End If
    End If

    Exit Function
ErrTrap:

    If RsSavRec.EditMode <> adEditNone Then
        RsSavRec.CancelUpdate
        If mIndex = 0 Then
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

Private Sub ShowTipE()
    On Error GoTo ErrTrap
    Dim TTP As New clstooltip
    Dim Wrap As String
    Dim Msg As String
    Wrap = CHR(13) + CHR(10)

    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "New" & Wrap & "To open a new record" & Wrap & "Press this key" & Wrap & "or Key" & " F12 √Ê Enter"
            
        .AddControl btnNew, Msg, True
    End With

    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Modification" & Wrap & "Modifying the current record " & Wrap & "Press this key" & Wrap & "or Key" & " F11"
        .AddControl btnModify, Msg, True
    End With

    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Save" & Wrap & "To record the data within the database " & Wrap & "Press this key" & Wrap & "or Key" & " F10"
        .AddControl btnSave, Msg, True
    End With

    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Undo" & Wrap & "To undo the current operation" & Wrap & "Press this key" & Wrap & "or Key" & " F9"
        .AddControl BtnUndo, Msg, True
    End With

    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Delete the record" & Wrap & "To delete the current record" & Wrap & "Press this key" & Wrap & "or Key" & " F18"
        .AddControl btnDelete, Msg, True
    End With

    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Exit" & Wrap & "To close this window" & Wrap & "Press this key" & Wrap & "or Key" & " Ctrl+x"
        .AddControl btnCancel, Msg, True
    End With

    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "First" & Wrap & "Move first record" & Wrap & "Press this key" & Wrap & "or Key" & " Home √Ê UpArrow"
        .AddControl btnFirst, Msg, True
    End With

    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Previous" & Wrap & "Move previous record" & Wrap & "Press this key" & Wrap & "or Key" & " PageUp √Ê LeftArrow"
        .AddControl btnPrevious, Msg, True
    End With

    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Next" & Wrap & "Move next record" & Wrap & "Press this key" & Wrap & "or Key" & " PageDown √Ê RightArrow"
        .AddControl btnNext, Msg, True
    End With

    With TTP
        .Create Me.hwnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Last" & Wrap & "Move last record" & Wrap & "Press this key" & Wrap & "or Key" & " End √Ê DownArrow"
        .AddControl btnLast, Msg, True
    End With

ErrTrap:
End Sub
'-------------------------------------------------------------

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap
    If mGridClicked Then Exit Sub
    If KeyCode = vbKeyReturn Then
        If Me.TxtModFlg.Text = "R" Then
            btnNew_Click
        Else
        '    SendKeys "{TAB}"
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
        BtnFirst_Click
    End If

    'Move Previous---------------------------------------------------------
    If KeyCode = vbKeyLeft Or KeyCode = vbKeyPageUp Then
        If btnPrevious.Enabled = False Then Exit Sub
        BtnPrevious_Click
    End If

    'Move Next---------------------------------------------------------
    If KeyCode = vbKeyRight Or KeyCode = vbKeyPageDown Then
        If btnNext.Enabled = False Then Exit Sub
        BtnNext_Click
    End If

    'Move Last---------------------------------------------------------
    If KeyCode = vbKeyDown Or KeyCode = vbKeyEnd Then
        If btnLast.Enabled = False Then Exit Sub
        BtnLast_Click
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


Public Sub GridSerial(vsGrd As Object, _
                      Optional chkRowHidden As Boolean = False, _
                      Optional SerStartRow As Single = 0)
    Dim j As Long
    Dim ss As Long
    If SerStartRow = 0 Then SerStartRow = vsGrd.FixedRows

    For ss = SerStartRow To vsGrd.Rows - 1
        If chkRowHidden Then
            If Not vsGrd.RowHidden(ss) Then j = j + 1
        Else
            j = j + 1
        End If
        vsGrd.TextMatrix(ss, 0) = j
    Next

End Sub

