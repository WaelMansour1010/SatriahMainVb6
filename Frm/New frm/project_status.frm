VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{28D9BF84-BC20-11D2-94CF-004005455FAA}#1.1#0"; "ImpulseAniLabel.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
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
      Caption         =   "«‰Ê«⁄ «·Ê—œÌ« | ”·Ì„ Ê—œÌ…|Õ«·«  «·„‘«—Ì⁄|«‰Ê«⁄ «·«ÿ»«¡|«·”Ì«—« |«·„÷Œ« |«‰Ê«⁄ «·“ÌÊ "
      Align           =   0
      CurrTab         =   5
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
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   9315
         Index           =   0
         Left            =   -20625
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
            Left            =   90
            RightToLeft     =   -1  'True
            TabIndex        =   103
            Top             =   4320
            Width           =   6225
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
            Width           =   18870
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
            Left            =   25920
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
            Width           =   825
            _ExtentX        =   1455
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
            Left            =   3165
            TabIndex        =   18
            Top             =   8610
            Width           =   810
            _ExtentX        =   1429
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
            Left            =   5505
            TabIndex        =   19
            TabStop         =   0   'False
            ToolTipText     =   " ÕœÌÀ ﬁ«⁄œ… «·»Ì«‰« "
            Top             =   7830
            Visible         =   0   'False
            Width           =   990
            _ExtentX        =   1746
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
            Left            =   90
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
            Left            =   2070
            TabIndex        =   21
            TabStop         =   0   'False
            ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
            Top             =   8520
            Width           =   1005
            _ExtentX        =   1773
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
            Left            =   810
            TabIndex        =   22
            TabStop         =   0   'False
            ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
            Top             =   8475
            Width           =   1170
            _ExtentX        =   2064
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
            Left            =   2250
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
            Width           =   1710
            _ExtentX        =   3016
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
            Left            =   90
            TabIndex        =   111
            Top             =   690
            Width           =   7500
            _cx             =   13229
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
            Left            =   2805
            RightToLeft     =   -1  'True
            TabIndex        =   28
            Top             =   8280
            Width           =   540
         End
         Begin VB.Label LabCurr_Rec 
            BackColor       =   &H00E2E9E9&
            Height          =   225
            Index           =   1
            Left            =   4605
            RightToLeft     =   -1  'True
            TabIndex        =   27
            Top             =   8280
            Width           =   720
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «·”Ã·« :"
            Height          =   210
            Index           =   2
            Left            =   3525
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
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   9315
         Index           =   2
         Left            =   -20325
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
            Left            =   12195
            RightToLeft     =   -1  'True
            TabIndex        =   168
            Top             =   780
            Width           =   1350
         End
         Begin VB.TextBox txtNoteLate 
            Alignment       =   1  'Right Justify
            Height          =   645
            Left            =   270
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   161
            Top             =   7770
            Width           =   8670
         End
         Begin VB.TextBox txtRemarks 
            Alignment       =   1  'Right Justify
            Height          =   645
            Left            =   9660
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   159
            Top             =   7770
            Width           =   8310
         End
         Begin VB.TextBox txtNoteStill 
            Alignment       =   1  'Right Justify
            Height          =   3270
            Left            =   270
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   157
            Top             =   4170
            Width           =   8670
         End
         Begin VB.TextBox txtNoteDone 
            Alignment       =   1  'Right Justify
            Height          =   3240
            Left            =   9660
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
            Width           =   7140
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
            Left            =   3525
            RightToLeft     =   -1  'True
            TabIndex        =   134
            Top             =   2280
            Width           =   6765
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
            Left            =   8220
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
            Left            =   3525
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   115
            Top             =   1665
            Width           =   3420
         End
         Begin VB.TextBox TxtSearchCode 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   330
            Left            =   6045
            RightToLeft     =   -1  'True
            TabIndex        =   114
            Top             =   1290
            Width           =   900
         End
         Begin VB.TextBox TxtDeptNotes 
            Alignment       =   1  'Right Justify
            Height          =   480
            Left            =   10470
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   113
            Top             =   3390
            Width           =   6960
         End
         Begin VB.TextBox TxtInitialNotes 
            Alignment       =   1  'Right Justify
            Height          =   480
            Left            =   2625
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   112
            Top             =   3390
            Width           =   6315
         End
         Begin VB.TextBox TxtSerial1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Enabled         =   0   'False
            Height          =   405
            Index           =   2
            Left            =   15975
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
            Width           =   1350
         End
         Begin VB.Frame Fra_Header 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   570
            Index           =   1
            Left            =   90
            RightToLeft     =   -1  'True
            TabIndex        =   30
            Top             =   0
            Width           =   18510
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
            Left            =   9480
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
            Width           =   990
            _ExtentX        =   1746
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
            Left            =   8580
            TabIndex        =   45
            Top             =   8730
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
            ButtonImage     =   "project_status.frx":FB3D
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton Btn_Undo 
            Height          =   345
            Index           =   2
            Left            =   6585
            TabIndex        =   46
            Top             =   8730
            Width           =   1005
            _ExtentX        =   1773
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
            Width           =   990
            _ExtentX        =   1746
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
            Width           =   990
            _ExtentX        =   1746
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
            Left            =   450
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
            Width           =   2625
            _ExtentX        =   4630
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
            Left            =   7140
            TabIndex        =   51
            Top             =   750
            Width           =   2250
            _ExtentX        =   3969
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
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   741
            _Version        =   393216
            Format          =   235405313
            CurrentDate     =   38784
         End
         Begin ImpulseButton.ISButton Btn_Print 
            Height          =   555
            Index           =   2
            Left            =   4245
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
            Left            =   3165
            TabIndex        =   54
            TabStop         =   0   'False
            ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
            Top             =   8640
            Width           =   1080
            _ExtentX        =   1905
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
            Left            =   3525
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
            Left            =   630
            TabIndex        =   153
            Top             =   750
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSComCtl2.DTPicker txtDateRec 
            Height          =   435
            Left            =   90
            TabIndex        =   162
            Top             =   2160
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   767
            _Version        =   393216
            Format          =   235405313
            CurrentDate     =   38784
         End
         Begin MSComCtl2.DTPicker txtTimeRec 
            Height          =   420
            Left            =   90
            TabIndex        =   163
            Top             =   2610
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   741
            _Version        =   393216
            Format          =   235405314
            CurrentDate     =   38784
         End
         Begin MSComCtl2.DTPicker txtDateEnd 
            Height          =   435
            Left            =   11280
            TabIndex        =   170
            Top             =   8430
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   767
            _Version        =   393216
            Format          =   235405313
            CurrentDate     =   38784
         End
         Begin MSComCtl2.DTPicker txtTimeEnd 
            Height          =   420
            Left            =   11280
            TabIndex        =   172
            Top             =   8850
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   741
            _Version        =   393216
            Format          =   235405314
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
            Width           =   1350
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «·«‰ Â«¡"
            Height          =   300
            Index           =   3
            Left            =   12555
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
            Left            =   13635
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
            Left            =   2985
            TabIndex        =   167
            Top             =   780
            Width           =   990
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "»‰«¡« ⁄·Ï «„— «·‘€·"
            Height          =   270
            Index           =   8
            Left            =   5685
            RightToLeft     =   -1  'True
            TabIndex        =   166
            Top             =   810
            Width           =   1545
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «·«” ·«„"
            Height          =   300
            Index           =   13
            Left            =   1350
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
            Width           =   1365
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "”»» «· √ŒÌ—"
            Height          =   300
            Index           =   10
            Left            =   2250
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
            Left            =   11835
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
            Width           =   2970
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„«  „ ⁄„·Â ›Ï «·’Ì«‰…"
            Height          =   285
            Index           =   1
            Left            =   11835
            TabIndex        =   154
            Top             =   3870
            Width           =   2520
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
            Width           =   1185
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„”ƒÊ· «·’Ì«‰…"
            Height          =   285
            Index           =   0
            Left            =   6945
            TabIndex        =   150
            Top             =   1305
            Width           =   1095
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
            Width           =   810
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„·«ÕŸ«  «·ﬂ‘› «·„»œ∆Ì"
            Height          =   420
            Index           =   69
            Left            =   8940
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
            Left            =   17250
            RightToLeft     =   -1  'True
            TabIndex        =   62
            Top             =   750
            Width           =   1170
         End
         Begin VB.Label LabCount_Rec 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   225
            Index           =   2
            Left            =   630
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
            Left            =   2625
            RightToLeft     =   -1  'True
            TabIndex        =   60
            Top             =   8400
            Width           =   900
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «·”Ã·« :"
            Height          =   225
            Index           =   4
            Left            =   1350
            RightToLeft     =   -1  'True
            TabIndex        =   59
            Top             =   8385
            Width           =   1275
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”Ã· «·Õ«·Ì:"
            Height          =   225
            Index           =   5
            Left            =   3525
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
            Width           =   1170
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·›—⁄"
            Height          =   270
            Index           =   16
            Left            =   9120
            TabIndex        =   56
            Top             =   810
            Width           =   990
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
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   9315
         Index           =   1
         Left            =   -20025
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
            Left            =   1710
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   86
            TabStop         =   0   'False
            Top             =   5610
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.Frame FraHeader 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   570
            Left            =   90
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
            Left            =   90
            RightToLeft     =   -1  'True
            TabIndex        =   64
            Top             =   4230
            Width           =   6225
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
            Left            =   90
            TabIndex        =   102
            Top             =   600
            Width           =   6225
            _cx             =   10980
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
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   9315
         Index           =   3
         Left            =   -19725
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
            Width           =   18870
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
            Left            =   90
            RightToLeft     =   -1  'True
            TabIndex        =   178
            Top             =   4320
            Width           =   6225
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
            Left            =   25920
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
            Width           =   825
            _ExtentX        =   1455
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
            Left            =   3165
            TabIndex        =   202
            Top             =   8610
            Width           =   810
            _ExtentX        =   1429
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
            Left            =   5505
            TabIndex        =   203
            TabStop         =   0   'False
            ToolTipText     =   " ÕœÌÀ ﬁ«⁄œ… «·»Ì«‰« "
            Top             =   7830
            Visible         =   0   'False
            Width           =   990
            _ExtentX        =   1746
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
            Left            =   90
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
            Left            =   2070
            TabIndex        =   205
            TabStop         =   0   'False
            ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
            Top             =   8520
            Width           =   1005
            _ExtentX        =   1773
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
            Left            =   810
            TabIndex        =   206
            TabStop         =   0   'False
            ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
            Top             =   8475
            Width           =   1170
            _ExtentX        =   2064
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
            Left            =   2250
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
            Width           =   1710
            _ExtentX        =   3016
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
            Left            =   90
            TabIndex        =   209
            Top             =   690
            Width           =   7500
            _cx             =   13229
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
            Left            =   3525
            RightToLeft     =   -1  'True
            TabIndex        =   212
            Top             =   8265
            Width           =   1080
         End
         Begin VB.Label LabCurr_Rec 
            BackColor       =   &H00E2E9E9&
            Height          =   225
            Index           =   3
            Left            =   4605
            RightToLeft     =   -1  'True
            TabIndex        =   211
            Top             =   8280
            Width           =   720
         End
         Begin VB.Label LabCount_Rec 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   225
            Index           =   3
            Left            =   2805
            RightToLeft     =   -1  'True
            TabIndex        =   210
            Top             =   8280
            Width           =   540
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   9315
         Index           =   4
         Left            =   -19425
         TabIndex        =   216
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
         Begin VB.CommandButton cmdSearch 
            Caption         =   "«” ⁄·«„"
            Height          =   315
            Left            =   630
            RightToLeft     =   -1  'True
            TabIndex        =   257
            Top             =   3780
            Width           =   1260
         End
         Begin VB.TextBox txtBalance 
            Alignment       =   1  'Right Justify
            Height          =   495
            Left            =   2250
            RightToLeft     =   -1  'True
            TabIndex        =   255
            Text            =   "50"
            Top             =   3630
            Width           =   2625
         End
         Begin VB.CheckBox chkOrg 
            Alignment       =   1  'Right Justify
            Caption         =   "√’·Ï"
            Height          =   540
            Left            =   3795
            RightToLeft     =   -1  'True
            TabIndex        =   253
            Top             =   2850
            Width           =   1080
         End
         Begin VB.CheckBox chkcom 
            Alignment       =   1  'Right Justify
            Caption         =   " Ã«—Ï"
            Height          =   540
            Left            =   2625
            RightToLeft     =   -1  'True
            TabIndex        =   252
            Top             =   2835
            Width           =   1080
         End
         Begin VB.CheckBox chkTestd 
            Alignment       =   1  'Right Justify
            Caption         =   "ÃœÌœ"
            Height          =   540
            Left            =   1440
            RightToLeft     =   -1  'True
            TabIndex        =   251
            Top             =   2835
            Width           =   1095
         End
         Begin VB.CheckBox chknormal 
            Alignment       =   1  'Right Justify
            Caption         =   "„” ⁄„·"
            Height          =   540
            Left            =   270
            RightToLeft     =   -1  'True
            TabIndex        =   250
            Top             =   2820
            Width           =   1170
         End
         Begin VB.TextBox TxtSerial1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   330
            Index           =   5
            Left            =   3795
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   239
            Top             =   855
            Width           =   1080
         End
         Begin VB.Frame Fra_Header 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   570
            Index           =   3
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   217
            Top             =   0
            Width           =   18870
            Begin VB.Frame Frame2 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   375
               Index           =   3
               Left            =   540
               RightToLeft     =   -1  'True
               TabIndex        =   220
               Top             =   450
               Visible         =   0   'False
               Width           =   3105
               Begin MSDataListLib.DataCombo DataCombo4 
                  CausesValidation=   0   'False
                  Height          =   315
                  Index           =   3
                  Left            =   -255
                  TabIndex        =   221
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
                  Index           =   2
                  Left            =   2160
                  RightToLeft     =   -1  'True
                  TabIndex        =   222
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
               Index           =   5
               Left            =   4170
               RightToLeft     =   -1  'True
               TabIndex        =   219
               Text            =   "modflag"
               Top             =   240
               Visible         =   0   'False
               Width           =   465
            End
            Begin VB.TextBox Text2 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H000000FF&
               Height          =   240
               Index           =   3
               Left            =   3030
               RightToLeft     =   -1  'True
               TabIndex        =   218
               Top             =   510
               Visible         =   0   'False
               Width           =   945
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
                     Picture         =   "project_status.frx":290C8
                     Key             =   "CompanyName"
                  EndProperty
                  BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "project_status.frx":29462
                     Key             =   "Ser"
                  EndProperty
                  BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "project_status.frx":297FC
                     Key             =   "Vac_Name"
                  EndProperty
                  BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "project_status.frx":29B96
                     Key             =   "ShareCount"
                  EndProperty
                  BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "project_status.frx":29F30
                     Key             =   "Dis_Count"
                  EndProperty
                  BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "project_status.frx":2A2CA
                     Key             =   "Bouns"
                  EndProperty
                  BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "project_status.frx":2A664
                     Key             =   "SharesValue"
                  EndProperty
                  BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "project_status.frx":2ABFE
                     Key             =   "BuyValue"
                  EndProperty
               EndProperty
            End
            Begin ImpulseButton.ISButton btn_Last 
               Height          =   315
               Index           =   5
               Left            =   90
               TabIndex        =   223
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
               ButtonImage     =   "project_status.frx":2AF98
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
               TabIndex        =   224
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
               ButtonImage     =   "project_status.frx":2B332
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btn_Previous 
               Height          =   315
               Index           =   5
               Left            =   1155
               TabIndex        =   225
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
               ButtonImage     =   "project_status.frx":2B6CC
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btn_First 
               Height          =   315
               Index           =   5
               Left            =   1620
               TabIndex        =   226
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
               ButtonImage     =   "project_status.frx":2BA66
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«·»ÕÀ ⁄‰ ﬁÿ⁄ €Ì«—"
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
               Index           =   3
               Left            =   2220
               RightToLeft     =   -1  'True
               TabIndex        =   227
               Top             =   30
               Width           =   2640
            End
         End
         Begin VSFlex8UCtl.VSFlexGrid FgItems 
            Height          =   9195
            Index           =   2
            Left            =   25920
            TabIndex        =   228
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
            FormatString    =   $"project_status.frx":2BE00
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
            Index           =   5
            Left            =   6675
            TabIndex        =   229
            Top             =   8685
            Visible         =   0   'False
            Width           =   825
            _ExtentX        =   1455
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
            ButtonImage     =   "project_status.frx":2BEC0
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Save 
            Height          =   345
            Index           =   5
            Left            =   4875
            TabIndex        =   230
            Top             =   8685
            Visible         =   0   'False
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
            ButtonImage     =   "project_status.frx":2C25A
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Modify 
            Height          =   345
            Index           =   5
            Left            =   5775
            TabIndex        =   231
            Top             =   8670
            Visible         =   0   'False
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
            ButtonImage     =   "project_status.frx":2C5F4
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton Btn_Undo 
            Height          =   345
            Index           =   5
            Left            =   3975
            TabIndex        =   232
            Top             =   8685
            Visible         =   0   'False
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
            ButtonImage     =   "project_status.frx":2C98E
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Delete 
            Height          =   345
            Index           =   5
            Left            =   3165
            TabIndex        =   233
            Top             =   8685
            Visible         =   0   'False
            Width           =   810
            _ExtentX        =   1429
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
            ButtonImage     =   "project_status.frx":2CD28
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Cancel 
            Height          =   345
            Index           =   5
            Left            =   90
            TabIndex        =   234
            Top             =   8685
            Visible         =   0   'False
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
            ButtonImage     =   "project_status.frx":2D2C2
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton Btn_Print 
            Height          =   540
            Index           =   5
            Left            =   2070
            TabIndex        =   235
            TabStop         =   0   'False
            ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
            Top             =   8595
            Visible         =   0   'False
            Width           =   1005
            _ExtentX        =   1773
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
            ButtonImage     =   "project_status.frx":2D65C
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btn_Query 
            Height          =   600
            Index           =   5
            Left            =   810
            TabIndex        =   236
            TabStop         =   0   'False
            ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
            Top             =   8580
            Visible         =   0   'False
            Width           =   1260
            _ExtentX        =   2223
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
            ButtonImage     =   "project_status.frx":33EBE
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin MSComCtl2.DTPicker XPDtbTrans 
            Height          =   315
            Left            =   540
            TabIndex        =   240
            Top             =   825
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   556
            _Version        =   393216
            Format          =   238157825
            CurrentDate     =   38784
         End
         Begin MSDataListLib.DataCombo DcbCarModel 
            Height          =   315
            Left            =   270
            TabIndex        =   241
            Top             =   2100
            Width           =   4605
            _ExtentX        =   8123
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "6"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcbCarGroup 
            Bindings        =   "project_status.frx":34258
            Height          =   315
            Left            =   2250
            TabIndex        =   242
            Top             =   1350
            Width           =   2715
            _ExtentX        =   4789
            _ExtentY        =   556
            _Version        =   393216
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
         Begin MSDataListLib.DataCombo dcItems 
            Height          =   315
            Left            =   270
            TabIndex        =   248
            Top             =   2490
            Width           =   4605
            _ExtentX        =   8123
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "6"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcbCarType 
            Height          =   315
            Left            =   270
            TabIndex        =   254
            Top             =   1740
            Width           =   4605
            _ExtentX        =   8123
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "6"
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "ﬂ„Ì… «·’‰›"
            Height          =   255
            Index           =   18
            Left            =   4965
            RightToLeft     =   -1  'True
            TabIndex        =   256
            Top             =   3720
            Width           =   1260
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«”„ «·ﬁÿ⁄Â"
            Height          =   255
            Index           =   17
            Left            =   4875
            RightToLeft     =   -1  'True
            TabIndex        =   249
            Top             =   2520
            Width           =   1260
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «·”‰œ"
            Height          =   285
            Index           =   15
            Left            =   2625
            TabIndex        =   247
            Top             =   825
            Width           =   1440
         End
         Begin VB.Label lbltycar 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "‰Ê⁄ «·”Ì«—…"
            Height          =   255
            Left            =   4155
            TabIndex        =   246
            Top             =   1380
            Width           =   1980
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·„ÊœÌ·"
            Height          =   255
            Index           =   125
            Left            =   4875
            RightToLeft     =   -1  'True
            TabIndex        =   245
            Top             =   2145
            Width           =   1260
         End
         Begin VB.Label LblYear 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·›∆Â"
            Height          =   270
            Left            =   4695
            TabIndex        =   244
            Top             =   1815
            Width           =   1440
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "—ﬁ„ «·”‰œ"
            Height          =   270
            Index           =   1
            Left            =   5235
            TabIndex        =   243
            Top             =   885
            Width           =   990
         End
         Begin VB.Label LabCount_Rec 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   225
            Index           =   0
            Left            =   2805
            RightToLeft     =   -1  'True
            TabIndex        =   238
            Top             =   8280
            Width           =   540
         End
         Begin VB.Label LabCurr_Rec 
            BackColor       =   &H00E2E9E9&
            Height          =   225
            Index           =   0
            Left            =   4605
            RightToLeft     =   -1  'True
            TabIndex        =   237
            Top             =   8280
            Width           =   720
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   9315
         Index           =   6
         Left            =   45
         TabIndex        =   258
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
            Height          =   555
            Index           =   4
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   259
            Top             =   0
            Width           =   18870
            Begin VB.Frame Frame2 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   375
               Index           =   6
               Left            =   540
               RightToLeft     =   -1  'True
               TabIndex        =   262
               Top             =   450
               Visible         =   0   'False
               Width           =   3105
               Begin MSDataListLib.DataCombo DataCombo4 
                  CausesValidation=   0   'False
                  Height          =   315
                  Index           =   6
                  Left            =   -255
                  TabIndex        =   263
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
                  Index           =   6
                  Left            =   2175
                  RightToLeft     =   -1  'True
                  TabIndex        =   264
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
               Index           =   6
               Left            =   2580
               RightToLeft     =   -1  'True
               TabIndex        =   261
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
               Index           =   6
               Left            =   3030
               RightToLeft     =   -1  'True
               TabIndex        =   260
               Top             =   510
               Visible         =   0   'False
               Width           =   945
            End
            Begin MSComctlLib.ImageList GrdImageList2 
               Index           =   6
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
                     Picture         =   "project_status.frx":3426D
                     Key             =   "CompanyName"
                  EndProperty
                  BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "project_status.frx":34607
                     Key             =   "Ser"
                  EndProperty
                  BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "project_status.frx":349A1
                     Key             =   "Vac_Name"
                  EndProperty
                  BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "project_status.frx":34D3B
                     Key             =   "ShareCount"
                  EndProperty
                  BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "project_status.frx":350D5
                     Key             =   "Dis_Count"
                  EndProperty
                  BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "project_status.frx":3546F
                     Key             =   "Bouns"
                  EndProperty
                  BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "project_status.frx":35809
                     Key             =   "SharesValue"
                  EndProperty
                  BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "project_status.frx":35DA3
                     Key             =   "BuyValue"
                  EndProperty
               EndProperty
            End
            Begin ImpulseButton.ISButton btn_Last 
               Height          =   315
               Index           =   6
               Left            =   90
               TabIndex        =   265
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
               ButtonImage     =   "project_status.frx":3613D
               ColorButton     =   14871017
               AcclimateGrayTones=   -1  'True
               DrawFocusRectangle=   0   'False
               DisabledImageExtraction=   0
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btn_Next 
               Height          =   315
               Index           =   6
               Left            =   555
               TabIndex        =   266
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
               ButtonImage     =   "project_status.frx":364D7
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btn_Previous 
               Height          =   315
               Index           =   6
               Left            =   1155
               TabIndex        =   267
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
               ButtonImage     =   "project_status.frx":36871
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btn_First 
               Height          =   315
               Index           =   6
               Left            =   1620
               TabIndex        =   268
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
               ButtonImage     =   "project_status.frx":36C0B
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "»Ì«‰«  «·„÷Œ« "
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
               Left            =   4980
               RightToLeft     =   -1  'True
               TabIndex        =   269
               Top             =   60
               Width           =   2640
            End
         End
         Begin VSFlex8UCtl.VSFlexGrid FgItems 
            Height          =   9195
            Index           =   3
            Left            =   25920
            TabIndex        =   270
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
            FormatString    =   $"project_status.frx":36FA5
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
            Index           =   6
            Left            =   6675
            TabIndex        =   271
            Top             =   8610
            Width           =   825
            _ExtentX        =   1455
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
            ButtonImage     =   "project_status.frx":37065
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Save 
            Height          =   345
            Index           =   6
            Left            =   4875
            TabIndex        =   272
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
            ButtonImage     =   "project_status.frx":373FF
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Modify 
            Height          =   345
            Index           =   6
            Left            =   5775
            TabIndex        =   273
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
            ButtonImage     =   "project_status.frx":37799
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton Btn_Undo 
            Height          =   345
            Index           =   6
            Left            =   3975
            TabIndex        =   274
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
            ButtonImage     =   "project_status.frx":37B33
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Delete 
            Height          =   345
            Index           =   6
            Left            =   3165
            TabIndex        =   275
            Top             =   8610
            Width           =   810
            _ExtentX        =   1429
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
            ButtonImage     =   "project_status.frx":37ECD
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton Btn_Update 
            Height          =   345
            Index           =   6
            Left            =   5505
            TabIndex        =   276
            TabStop         =   0   'False
            ToolTipText     =   " ÕœÌÀ ﬁ«⁄œ… «·»Ì«‰« "
            Top             =   7830
            Visible         =   0   'False
            Width           =   990
            _ExtentX        =   1746
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
            ButtonImage     =   "project_status.frx":38467
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Cancel 
            Height          =   345
            Index           =   6
            Left            =   90
            TabIndex        =   277
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
            ButtonImage     =   "project_status.frx":38801
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton Btn_Print 
            Height          =   540
            Index           =   6
            Left            =   2070
            TabIndex        =   278
            TabStop         =   0   'False
            ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
            Top             =   8520
            Width           =   1005
            _ExtentX        =   1773
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
            ButtonImage     =   "project_status.frx":38B9B
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btn_Query 
            Height          =   600
            Index           =   6
            Left            =   810
            TabIndex        =   279
            TabStop         =   0   'False
            ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
            Top             =   8475
            Width           =   1170
            _ExtentX        =   2064
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
            ButtonImage     =   "project_status.frx":3F3FD
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton Cmd_DeleteRow 
            Height          =   300
            Index           =   6
            Left            =   2250
            TabIndex        =   280
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
            ButtonImage     =   "project_status.frx":3F797
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton Cmd_DeleteAll 
            Height          =   300
            Index           =   6
            Left            =   360
            TabIndex        =   281
            Top             =   7785
            Width           =   1710
            _ExtentX        =   3016
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
            ButtonImage     =   "project_status.frx":3FD31
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin VSFlex8Ctl.VSFlexGrid Grid6 
            Height          =   3060
            Left            =   90
            TabIndex        =   282
            Top             =   660
            Width           =   7500
            _cx             =   13229
            _cy             =   5397
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
            FormatString    =   $"project_status.frx":402CB
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
         Begin C1SizerLibCtl.C1Tab C1Tab1 
            Height          =   4095
            Left            =   90
            TabIndex        =   287
            Top             =   3810
            Width           =   7410
            _cx             =   13070
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
            Appearance      =   2
            MousePointer    =   0
            Version         =   801
            BackColor       =   14871017
            ForeColor       =   -2147483630
            FrontTabColor   =   -2147483633
            BackTabColor    =   14871017
            TabOutlineColor =   -2147483632
            FrontTabForeColor=   -2147483630
            Caption         =   "«·„÷Œ« |«·”Ì«—« "
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
               Height          =   3720
               Index           =   1
               Left            =   45
               TabIndex        =   288
               TabStop         =   0   'False
               Top             =   45
               Width           =   7320
               _cx             =   12912
               _cy             =   6562
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
               Begin VB.Frame Frame7 
                  BackColor       =   &H00E2E9E9&
                  BorderStyle     =   0  'None
                  Enabled         =   0   'False
                  Height          =   3315
                  Left            =   0
                  RightToLeft     =   -1  'True
                  TabIndex        =   301
                  Top             =   0
                  Width           =   6225
                  Begin VB.TextBox txtPunpPer 
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
                     Left            =   1170
                     MaxLength       =   50
                     RightToLeft     =   -1  'True
                     TabIndex        =   305
                     Top             =   1740
                     Width           =   3165
                  End
                  Begin VB.TextBox txtpumpNameE 
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
                     Left            =   1170
                     MaxLength       =   50
                     RightToLeft     =   -1  'True
                     TabIndex        =   304
                     Top             =   1365
                     Width           =   3165
                  End
                  Begin VB.TextBox txtpumpName 
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
                     Left            =   1170
                     MaxLength       =   50
                     RightToLeft     =   -1  'True
                     TabIndex        =   303
                     Top             =   1065
                     Width           =   3165
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
                     Index           =   6
                     Left            =   1170
                     MaxLength       =   50
                     RightToLeft     =   -1  'True
                     TabIndex        =   302
                     Top             =   705
                     Width           =   3165
                  End
                  Begin MSDataListLib.DataCombo txtItem 
                     Height          =   315
                     Left            =   1155
                     TabIndex        =   306
                     Top             =   2115
                     Width           =   3165
                     _ExtentX        =   5583
                     _ExtentY        =   556
                     _Version        =   393216
                     BackColor       =   16777215
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin MSDataListLib.DataCombo txtBox 
                     Height          =   315
                     Left            =   1170
                     TabIndex        =   307
                     Top             =   2520
                     Width           =   3165
                     _ExtentX        =   5583
                     _ExtentY        =   556
                     _Version        =   393216
                     BackColor       =   16777215
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin MSDataListLib.DataCombo cmbStore 
                     Height          =   315
                     Left            =   1170
                     TabIndex        =   308
                     Top             =   2910
                     Width           =   3165
                     _ExtentX        =   5583
                     _ExtentY        =   556
                     _Version        =   393216
                     BackColor       =   16777215
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin MSDataListLib.DataCombo cmbStore2 
                     Height          =   315
                     Left            =   1230
                     TabIndex        =   376
                     Top             =   0
                     Width           =   3165
                     _ExtentX        =   5583
                     _ExtentY        =   556
                     _Version        =   393216
                     BackColor       =   16777215
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "⁄—÷ „÷Œ«  «·„Œ“‰"
                     Height          =   285
                     Index           =   23
                     Left            =   4755
                     RightToLeft     =   -1  'True
                     TabIndex        =   377
                     Top             =   60
                     Width           =   1425
                  End
                  Begin VB.Label Label5 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·ﬁ—«¡…"
                     Height          =   285
                     Index           =   2
                     Left            =   5130
                     RightToLeft     =   -1  'True
                     TabIndex        =   315
                     Top             =   1815
                     Width           =   780
                  End
                  Begin VB.Label Label9 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«”„ «·„÷ŒÂ «‰Ã·Ì“Ï"
                     Height          =   285
                     Left            =   4410
                     RightToLeft     =   -1  'True
                     TabIndex        =   314
                     Top             =   1500
                     Width           =   1500
                  End
                  Begin VB.Label Label10 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«”„ «·„÷ŒÂ ⁄—»Ì"
                     Height          =   285
                     Left            =   4560
                     RightToLeft     =   -1  'True
                     TabIndex        =   313
                     Top             =   1140
                     Width           =   1350
                  End
                  Begin VB.Label Label11 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "ﬂÊœ «·„÷ŒÂ"
                     Height          =   195
                     Left            =   4920
                     RightToLeft     =   -1  'True
                     TabIndex        =   312
                     Top             =   810
                     Width           =   990
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·’‰›"
                     Height          =   285
                     Index           =   28
                     Left            =   4695
                     RightToLeft     =   -1  'True
                     TabIndex        =   311
                     Top             =   2175
                     Width           =   1215
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·Œ“Ì‰Â"
                     Height          =   285
                     Index           =   20
                     Left            =   4695
                     RightToLeft     =   -1  'True
                     TabIndex        =   310
                     Top             =   2580
                     Width           =   1215
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·„Œ“‰"
                     Height          =   285
                     Index           =   21
                     Left            =   4695
                     RightToLeft     =   -1  'True
                     TabIndex        =   309
                     Top             =   2970
                     Width           =   1215
                  End
               End
               Begin ImpulseAniLabel.ISAniLabel LblLink 
                  Height          =   165
                  Left            =   0
                  TabIndex        =   289
                  Top             =   165
                  Width           =   1260
                  _ExtentX        =   2223
                  _ExtentY        =   291
                  ActiveUnderline =   -1  'True
                  BackStyle       =   1
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  FontBold        =   -1  'True
                  FontName        =   "MS Sans Serif"
                  FontSize        =   8.25
                  ForeColor       =   4210688
                  MousePointer    =   99
                  MouseIcon       =   "project_status.frx":40360
                  BackColor       =   14871017
                  Alignment       =   1
                  Caption         =   ""
                  ColorHover      =   16711680
                  RightToLeft     =   -1  'True
                  ImageCount      =   0
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000080&
                  Height          =   1155
                  Index           =   25
                  Left            =   30
                  RightToLeft     =   -1  'True
                  TabIndex        =   290
                  Top             =   360
                  Width           =   1980
               End
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic3 
               Height          =   3720
               Index           =   1
               Left            =   8055
               TabIndex        =   291
               TabStop         =   0   'False
               Top             =   45
               Width           =   7320
               _cx             =   12912
               _cy             =   6562
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
               Begin C1SizerLibCtl.C1Elastic Ele 
                  Height          =   4860
                  Index           =   11
                  Left            =   -720
                  TabIndex        =   292
                  TabStop         =   0   'False
                  Top             =   0
                  Width           =   7965
                  _cx             =   14049
                  _cy             =   8573
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
                  Begin VB.TextBox txtAmountH 
                     Alignment       =   1  'Right Justify
                     Height          =   345
                     Left            =   3000
                     RightToLeft     =   -1  'True
                     TabIndex        =   318
                     Top             =   900
                     Width           =   1275
                  End
                  Begin VB.CheckBox chkIsOther 
                     Alignment       =   1  'Right Justify
                     Caption         =   "«œ«—… ··€Ì—"
                     Height          =   405
                     Left            =   3030
                     RightToLeft     =   -1  'True
                     TabIndex        =   316
                     Top             =   330
                     Width           =   1245
                  End
                  Begin VB.TextBox XPTxtSerial 
                     Alignment       =   1  'Right Justify
                     Height          =   870
                     Index           =   1
                     Left            =   5295
                     Locked          =   -1  'True
                     TabIndex        =   293
                     Top             =   4830
                     Visible         =   0   'False
                     Width           =   600
                  End
                  Begin MSDataListLib.DataCombo cmbChangeType 
                     Height          =   315
                     Left            =   5895
                     TabIndex        =   294
                     Top             =   4725
                     Width           =   195
                     _ExtentX        =   344
                     _ExtentY        =   556
                     _Version        =   393216
                     ListField       =   "6"
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin MSDataListLib.DataCombo cmbAccount 
                     Height          =   315
                     Left            =   960
                     TabIndex        =   297
                     Top             =   2280
                     Width           =   3585
                     _ExtentX        =   6324
                     _ExtentY        =   556
                     _Version        =   393216
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin MSDataListLib.DataCombo cmbAccountComm 
                     Height          =   315
                     Left            =   960
                     TabIndex        =   298
                     Top             =   2790
                     Width           =   3585
                     _ExtentX        =   6324
                     _ExtentY        =   556
                     _Version        =   393216
                     Locked          =   -1  'True
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "·ﬂ· · —"
                     Height          =   330
                     Index           =   22
                     Left            =   4230
                     RightToLeft     =   -1  'True
                     TabIndex        =   317
                     Top             =   930
                     Width           =   1020
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·Õ”«»"
                     Height          =   150
                     Index           =   26
                     Left            =   4380
                     RightToLeft     =   -1  'True
                     TabIndex        =   300
                     Top             =   2280
                     Width           =   1920
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "Õ”«» «·⁄„Ê·« "
                     Height          =   255
                     Index           =   89
                     Left            =   4320
                     RightToLeft     =   -1  'True
                     TabIndex        =   299
                     Top             =   2880
                     Width           =   1890
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackStyle       =   0  'Transparent
                     Caption         =   "‰Ê⁄ «· €Ì—"
                     Height          =   600
                     Index           =   164
                     Left            =   5895
                     RightToLeft     =   -1  'True
                     TabIndex        =   295
                     Top             =   4995
                     Width           =   510
                  End
               End
               Begin VB.Label Label1100 
                  Alignment       =   2  'Center
                  BackColor       =   &H00FFFFC0&
                  Caption         =   "„ÿ·Ê» «⁄ „«œ… Õ«·Ì«"
                  Height          =   540
                  Left            =   9165
                  RightToLeft     =   -1  'True
                  TabIndex        =   296
                  Top             =   4350
                  Width           =   4095
               End
            End
         End
         Begin VB.Label LabCount_Rec 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   225
            Index           =   4
            Left            =   2805
            RightToLeft     =   -1  'True
            TabIndex        =   286
            Top             =   8280
            Width           =   540
         End
         Begin VB.Label LabCurr_Rec 
            BackColor       =   &H00E2E9E9&
            Height          =   225
            Index           =   4
            Left            =   4605
            RightToLeft     =   -1  'True
            TabIndex        =   285
            Top             =   8280
            Width           =   720
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «·”Ã·« :"
            Height          =   210
            Index           =   7
            Left            =   3525
            RightToLeft     =   -1  'True
            TabIndex        =   284
            Top             =   8265
            Width           =   1080
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”Ã· «·Õ«·Ì:"
            Height          =   210
            Index           =   6
            Left            =   5415
            RightToLeft     =   -1  'True
            TabIndex        =   283
            Top             =   8265
            Width           =   1080
         End
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   9315
         Index           =   5
         Left            =   19515
         TabIndex        =   319
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
            Height          =   555
            Index           =   5
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   320
            Top             =   0
            Width           =   18870
            Begin VB.TextBox Text2 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H000000FF&
               Height          =   240
               Index           =   4
               Left            =   3030
               RightToLeft     =   -1  'True
               TabIndex        =   325
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
               Index           =   7
               Left            =   2580
               RightToLeft     =   -1  'True
               TabIndex        =   324
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
               Index           =   4
               Left            =   540
               RightToLeft     =   -1  'True
               TabIndex        =   321
               Top             =   450
               Visible         =   0   'False
               Width           =   3105
               Begin MSDataListLib.DataCombo DataCombo4 
                  CausesValidation=   0   'False
                  Height          =   315
                  Index           =   4
                  Left            =   -255
                  TabIndex        =   322
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
                  Index           =   10
                  Left            =   2175
                  RightToLeft     =   -1  'True
                  TabIndex        =   323
                  Top             =   45
                  Width           =   855
               End
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
                     Picture         =   "project_status.frx":404C2
                     Key             =   "CompanyName"
                  EndProperty
                  BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "project_status.frx":4085C
                     Key             =   "Ser"
                  EndProperty
                  BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "project_status.frx":40BF6
                     Key             =   "Vac_Name"
                  EndProperty
                  BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "project_status.frx":40F90
                     Key             =   "ShareCount"
                  EndProperty
                  BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "project_status.frx":4132A
                     Key             =   "Dis_Count"
                  EndProperty
                  BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "project_status.frx":416C4
                     Key             =   "Bouns"
                  EndProperty
                  BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "project_status.frx":41A5E
                     Key             =   "SharesValue"
                  EndProperty
                  BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "project_status.frx":41FF8
                     Key             =   "BuyValue"
                  EndProperty
               EndProperty
            End
            Begin ImpulseButton.ISButton btn_Last 
               Height          =   315
               Index           =   7
               Left            =   90
               TabIndex        =   326
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
               ButtonImage     =   "project_status.frx":42392
               ColorButton     =   14871017
               AcclimateGrayTones=   -1  'True
               DrawFocusRectangle=   0   'False
               DisabledImageExtraction=   0
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btn_Next 
               Height          =   315
               Index           =   7
               Left            =   540
               TabIndex        =   327
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
               ButtonImage     =   "project_status.frx":4272C
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btn_Previous 
               Height          =   315
               Index           =   7
               Left            =   1155
               TabIndex        =   328
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
               ButtonImage     =   "project_status.frx":42AC6
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btn_First 
               Height          =   315
               Index           =   7
               Left            =   1620
               TabIndex        =   329
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
               ButtonImage     =   "project_status.frx":42E60
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«‰Ê«⁄ «·“ÌÊ "
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
               Index           =   11
               Left            =   4980
               RightToLeft     =   -1  'True
               TabIndex        =   330
               Top             =   60
               Width           =   2640
            End
         End
         Begin VSFlex8UCtl.VSFlexGrid FgItems 
            Height          =   9195
            Index           =   4
            Left            =   25920
            TabIndex        =   331
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
            FormatString    =   $"project_status.frx":431FA
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
            Index           =   7
            Left            =   6675
            TabIndex        =   332
            Top             =   8610
            Width           =   825
            _ExtentX        =   1455
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
            ButtonImage     =   "project_status.frx":432BA
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Save 
            Height          =   345
            Index           =   7
            Left            =   4875
            TabIndex        =   333
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
            ButtonImage     =   "project_status.frx":43654
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Modify 
            Height          =   345
            Index           =   7
            Left            =   5775
            TabIndex        =   334
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
            ButtonImage     =   "project_status.frx":439EE
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton Btn_Undo 
            Height          =   345
            Index           =   7
            Left            =   3975
            TabIndex        =   335
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
            ButtonImage     =   "project_status.frx":43D88
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Delete 
            Height          =   345
            Index           =   7
            Left            =   3165
            TabIndex        =   336
            Top             =   8610
            Width           =   810
            _ExtentX        =   1429
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
            ButtonImage     =   "project_status.frx":44122
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton Btn_Update 
            Height          =   345
            Index           =   0
            Left            =   5505
            TabIndex        =   337
            TabStop         =   0   'False
            ToolTipText     =   " ÕœÌÀ ﬁ«⁄œ… «·»Ì«‰« "
            Top             =   7830
            Visible         =   0   'False
            Width           =   990
            _ExtentX        =   1746
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
            ButtonImage     =   "project_status.frx":446BC
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Cancel 
            Height          =   345
            Index           =   7
            Left            =   90
            TabIndex        =   338
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
            ButtonImage     =   "project_status.frx":44A56
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton Btn_Print 
            Height          =   540
            Index           =   7
            Left            =   2070
            TabIndex        =   339
            TabStop         =   0   'False
            ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
            Top             =   8520
            Width           =   1005
            _ExtentX        =   1773
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
            ButtonImage     =   "project_status.frx":44DF0
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btn_Query 
            Height          =   600
            Index           =   7
            Left            =   810
            TabIndex        =   340
            TabStop         =   0   'False
            ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
            Top             =   8475
            Width           =   1170
            _ExtentX        =   2064
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
            ButtonImage     =   "project_status.frx":4B652
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton Cmd_DeleteRow 
            Height          =   300
            Index           =   7
            Left            =   2250
            TabIndex        =   341
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
            ButtonImage     =   "project_status.frx":4B9EC
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton Cmd_DeleteAll 
            Height          =   300
            Index           =   7
            Left            =   360
            TabIndex        =   342
            Top             =   7785
            Width           =   1710
            _ExtentX        =   3016
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
            ButtonImage     =   "project_status.frx":4BF86
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin VSFlex8Ctl.VSFlexGrid Grid7 
            Height          =   3060
            Left            =   90
            TabIndex        =   343
            Top             =   660
            Width           =   7500
            _cx             =   13229
            _cy             =   5397
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
            FormatString    =   $"project_status.frx":4C520
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
         Begin C1SizerLibCtl.C1Tab C1Tab2 
            Height          =   4095
            Left            =   90
            TabIndex        =   344
            Top             =   3810
            Width           =   7410
            _cx             =   13070
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
            Appearance      =   2
            MousePointer    =   0
            Version         =   801
            BackColor       =   14871017
            ForeColor       =   -2147483630
            FrontTabColor   =   -2147483633
            BackTabColor    =   14871017
            TabOutlineColor =   -2147483632
            FrontTabForeColor=   -2147483630
            Caption         =   "«·”Ì«—« "
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
               Height          =   3720
               Index           =   0
               Left            =   45
               TabIndex        =   345
               TabStop         =   0   'False
               Top             =   45
               Width           =   7320
               _cx             =   12912
               _cy             =   6562
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
               Begin VB.Frame Frame8 
                  BackColor       =   &H00E2E9E9&
                  BorderStyle     =   0  'None
                  Enabled         =   0   'False
                  Height          =   3315
                  Left            =   0
                  RightToLeft     =   -1  'True
                  TabIndex        =   346
                  Top             =   0
                  Width           =   6225
                  Begin VB.TextBox txtKiloMetr 
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
                     Left            =   1170
                     MaxLength       =   50
                     RightToLeft     =   -1  'True
                     TabIndex        =   374
                     Top             =   1710
                     Width           =   3165
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
                     Index           =   7
                     Left            =   1170
                     MaxLength       =   50
                     RightToLeft     =   -1  'True
                     TabIndex        =   350
                     Top             =   285
                     Width           =   3165
                  End
                  Begin VB.TextBox txtOilName 
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
                     Left            =   1170
                     MaxLength       =   50
                     RightToLeft     =   -1  'True
                     TabIndex        =   349
                     Top             =   645
                     Width           =   3165
                  End
                  Begin VB.TextBox txtOilNameE 
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
                     Left            =   1170
                     MaxLength       =   50
                     RightToLeft     =   -1  'True
                     TabIndex        =   348
                     Top             =   945
                     Width           =   3165
                  End
                  Begin VB.TextBox txtPeriod 
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
                     Left            =   1170
                     MaxLength       =   50
                     RightToLeft     =   -1  'True
                     TabIndex        =   347
                     Top             =   1320
                     Width           =   3165
                  End
                  Begin VB.Label Label5 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "⁄œœ «·ﬂÌ·Ê „ —"
                     Height          =   285
                     Index           =   4
                     Left            =   4950
                     RightToLeft     =   -1  'True
                     TabIndex        =   375
                     Top             =   1785
                     Width           =   960
                  End
                  Begin VB.Label Label14 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "ﬂÊœ «·“Ì "
                     Height          =   195
                     Left            =   4920
                     RightToLeft     =   -1  'True
                     TabIndex        =   354
                     Top             =   390
                     Width           =   990
                  End
                  Begin VB.Label Label13 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«”„ ⁄—»Ì"
                     Height          =   285
                     Left            =   4560
                     RightToLeft     =   -1  'True
                     TabIndex        =   353
                     Top             =   720
                     Width           =   1350
                  End
                  Begin VB.Label Label12 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«”„ «‰Ã·Ì“Ï"
                     Height          =   285
                     Left            =   4410
                     RightToLeft     =   -1  'True
                     TabIndex        =   352
                     Top             =   1080
                     Width           =   1500
                  End
                  Begin VB.Label Label5 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·„œ…"
                     Height          =   285
                     Index           =   3
                     Left            =   5130
                     RightToLeft     =   -1  'True
                     TabIndex        =   351
                     Top             =   1395
                     Width           =   780
                  End
               End
               Begin ImpulseAniLabel.ISAniLabel ISAniLabel1 
                  Height          =   165
                  Left            =   0
                  TabIndex        =   355
                  Top             =   165
                  Width           =   1260
                  _ExtentX        =   2223
                  _ExtentY        =   291
                  ActiveUnderline =   -1  'True
                  BackStyle       =   1
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  FontBold        =   -1  'True
                  FontName        =   "MS Sans Serif"
                  FontSize        =   8.25
                  ForeColor       =   4210688
                  MousePointer    =   99
                  MouseIcon       =   "project_status.frx":4C5B4
                  BackColor       =   14871017
                  Alignment       =   1
                  Caption         =   ""
                  ColorHover      =   16711680
                  RightToLeft     =   -1  'True
                  ImageCount      =   0
               End
               Begin VB.Label lbl 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E2E9E9&
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   178
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000080&
                  Height          =   1155
                  Index           =   30
                  Left            =   30
                  RightToLeft     =   -1  'True
                  TabIndex        =   356
                  Top             =   360
                  Width           =   1980
               End
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic3 
               Height          =   3720
               Index           =   0
               Left            =   8055
               TabIndex        =   357
               TabStop         =   0   'False
               Top             =   45
               Width           =   7320
               _cx             =   12912
               _cy             =   6562
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
               Begin C1SizerLibCtl.C1Elastic Ele 
                  Height          =   4860
                  Index           =   7
                  Left            =   -720
                  TabIndex        =   358
                  TabStop         =   0   'False
                  Top             =   0
                  Width           =   7965
                  _cx             =   14049
                  _cy             =   8573
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
                  Begin VB.TextBox XPTxtSerial 
                     Alignment       =   1  'Right Justify
                     Height          =   870
                     Index           =   0
                     Left            =   5295
                     Locked          =   -1  'True
                     TabIndex        =   361
                     Top             =   4830
                     Visible         =   0   'False
                     Width           =   600
                  End
                  Begin VB.CheckBox Check1 
                     Alignment       =   1  'Right Justify
                     Caption         =   "«œ«—… ··€Ì—"
                     Height          =   405
                     Left            =   3030
                     RightToLeft     =   -1  'True
                     TabIndex        =   360
                     Top             =   330
                     Width           =   1245
                  End
                  Begin VB.TextBox Text7 
                     Alignment       =   1  'Right Justify
                     Height          =   345
                     Left            =   3000
                     RightToLeft     =   -1  'True
                     TabIndex        =   359
                     Top             =   900
                     Width           =   1275
                  End
                  Begin MSDataListLib.DataCombo DataCombo5 
                     Height          =   315
                     Left            =   5895
                     TabIndex        =   362
                     Top             =   4725
                     Width           =   195
                     _ExtentX        =   344
                     _ExtentY        =   556
                     _Version        =   393216
                     ListField       =   "6"
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin MSDataListLib.DataCombo DataCombo6 
                     Height          =   315
                     Left            =   960
                     TabIndex        =   363
                     Top             =   2280
                     Width           =   3585
                     _ExtentX        =   6324
                     _ExtentY        =   556
                     _Version        =   393216
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin MSDataListLib.DataCombo DataCombo7 
                     Height          =   315
                     Left            =   960
                     TabIndex        =   364
                     Top             =   2790
                     Width           =   3585
                     _ExtentX        =   6324
                     _ExtentY        =   556
                     _Version        =   393216
                     Locked          =   -1  'True
                     Text            =   ""
                     RightToLeft     =   -1  'True
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackStyle       =   0  'Transparent
                     Caption         =   "‰Ê⁄ «· €Ì—"
                     Height          =   600
                     Index           =   35
                     Left            =   5895
                     RightToLeft     =   -1  'True
                     TabIndex        =   368
                     Top             =   4995
                     Width           =   510
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "Õ”«» «·⁄„Ê·« "
                     Height          =   255
                     Index           =   33
                     Left            =   4320
                     RightToLeft     =   -1  'True
                     TabIndex        =   367
                     Top             =   2880
                     Width           =   1890
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "«·Õ”«»"
                     Height          =   150
                     Index           =   32
                     Left            =   4380
                     RightToLeft     =   -1  'True
                     TabIndex        =   366
                     Top             =   2280
                     Width           =   1920
                  End
                  Begin VB.Label lbl 
                     Alignment       =   1  'Right Justify
                     BackColor       =   &H00E2E9E9&
                     Caption         =   "·ﬂ· · —"
                     Height          =   330
                     Index           =   31
                     Left            =   4230
                     RightToLeft     =   -1  'True
                     TabIndex        =   365
                     Top             =   930
                     Width           =   1020
                  End
               End
               Begin VB.Label Label15 
                  Alignment       =   2  'Center
                  BackColor       =   &H00FFFFC0&
                  Caption         =   "„ÿ·Ê» «⁄ „«œ… Õ«·Ì«"
                  Height          =   540
                  Left            =   9165
                  RightToLeft     =   -1  'True
                  TabIndex        =   369
                  Top             =   4350
                  Width           =   4095
               End
            End
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”Ã· «·Õ«·Ì:"
            Height          =   210
            Index           =   9
            Left            =   5415
            RightToLeft     =   -1  'True
            TabIndex        =   373
            Top             =   8265
            Width           =   1080
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «·”Ã·« :"
            Height          =   210
            Index           =   8
            Left            =   3525
            RightToLeft     =   -1  'True
            TabIndex        =   372
            Top             =   8265
            Width           =   1080
         End
         Begin VB.Label LabCurr_Rec 
            BackColor       =   &H00E2E9E9&
            Height          =   225
            Index           =   7
            Left            =   4605
            RightToLeft     =   -1  'True
            TabIndex        =   371
            Top             =   8280
            Width           =   720
         End
         Begin VB.Label LabCount_Rec 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   225
            Index           =   7
            Left            =   2805
            RightToLeft     =   -1  'True
            TabIndex        =   370
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
Dim ii As Long




Public mIndex As Integer
Public LngRow As Long
Public FixedAssetsID As Long
Public FixedAssetsName As String
Dim s As String
Dim mGridClicked As Boolean

Dim AccountVATCreit As String


Private Sub cmbAccount_KeyUp(KeyCode As Integer, Shift As Integer)



    If KeyCode = vbKeyF3 Then
        cmbAccount.text = ""
        '   Unload Account_search
        Account_search.show
        Account_search.case_id = 2608180
            
    End If



End Sub



Private Sub cmbAccountComm_KeyUp(KeyCode As Integer, Shift As Integer)



    If KeyCode = vbKeyF3 Then
        cmbAccountComm.text = ""
        '   Unload Account_search
        Account_search.show
        Account_search.case_id = 2608181
            
    End If



End Sub

Private Sub cmbStore2_Click(Area As Integer)
FillGridWithData6 val(cmbStore2.BoundText)
End Sub

Private Sub CmdAttach_Click()
    On Error Resume Next
          If DoPremis(Do_Attach, Me.Name, True) = False Then
                Exit Sub
            End If
            
ShowAttachments TxtSerial1(2), "21012020"

End Sub

Private Sub CmdSearch_Click()
Dim s As String
Dim FirstPeriodDateInthisYear  As Date
Dim rsDummy As New ADODB.Recordset
getFirstPeriodDateInthisYear2 FirstPeriodDateInthisYear
    s = " SELECT"
s = s & "     Transaction_Details.lotNo , Item_ID"
s = s & "    ,dbo.TblUnites.UnitId"
s = s & "    ,SUM(dbo.Transaction_Details.showqty * dbo.TransactionTypes.StockEffect) AS SUMQTY"

s = s & " From dbo.transactions"
s = s & " INNER JOIN dbo.Transaction_Details"
s = s & "     ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID"
s = s & " INNER JOIN dbo.TransactionTypes"
s = s & "     ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type"
s = s & " INNER JOIN dbo.TblItems"
s = s & "     ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID"
s = s & " INNER JOIN dbo.TblStore"
s = s & "     ON dbo.Transactions.StoreID = dbo.TblStore.StoreID"
s = s & " LEFT OUTER JOIN dbo.TblItemsSizes"
s = s & "     ON dbo.Transaction_Details.ItemSize = dbo.TblItemsSizes.SizeId"
s = s & " LEFT OUTER JOIN dbo.TblItemsclasses"
s = s & "     ON dbo.Transaction_Details.ClassId = dbo.TblItemsclasses.SizeId"
s = s & " LEFT OUTER JOIN dbo.TblUnites"
s = s & "     ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID"
s = s & " LEFT OUTER JOIN dbo.TblItemsColors"
s = s & "     ON dbo.Transaction_Details.ColorID = dbo.TblItemsColors.ColorID"


s = s + " where dbo.Transactions.Transaction_Date >=" & SQLDate(FirstPeriodDateInthisYear, True) & ""
s = s + " and dbo.Transactions.Transaction_Date <=" & SQLDate(Date, True) & ""
s = s + " and Item_ID =" & val(dcitems.BoundText)
            


'--AND Transactions.StoreID = 1
s = s & " GROUP BY dbo.TblStore.StoreName"
s = s & "         ,dbo.TblUnites.UnitName"
s = s & "         ,dbo.TblUnites.UnitId"
s = s & "         ,dbo.TblItemsclasses.SizeName"
s = s & "         ,dbo.TblItemsSizes.SizeName"
s = s & "         ,dbo.TblItemsColors.ColorName"
s = s & "         ,Transaction_Details.lotno ,Item_ID"
s = s & " Having (SUM(dbo.Transaction_Details.ShowQty * dbo.TransactionTypes.StockEffect) <> 0)"
Set rsDummy = New ADODB.Recordset
rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly
TxtBalance = 0
If Not rsDummy.EOF Then
    TxtBalance = rsDummy!SumQty & ""
End If

End Sub

Private Sub DcbCarGroup_Click(Area As Integer)
 Dim grpId As Integer
   
    Dim GID   As Integer
    GID = val(DcbCarGroup.BoundText)
    Dim StrSQL As String
    StrSQL = ""
    StrSQL = StrSQL & "SELECT ItemID, "
    StrSQL = StrSQL & "        "
    StrSQL = StrSQL & "       ItemName "
    StrSQL = StrSQL & "FROM TblItems "
    StrSQL = StrSQL & "WHERE GroupID = " & GID & ";"
         
    GetComboData dcitems, StrSQL
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
TxtDrievName.text = ""
ElseIf ChDrievType(1).value = True Then
Text6.Enabled = False
DcbDrievID.Enabled = False
TxtDrievName.Enabled = True
DcbDrievID.BoundText = 0
Text6.text = ""
End If
End Sub

Private Sub ChLeaderType_Click(Index As Integer)
If ChLeaderType(0).value = True Then
TxtNameE.Enabled = True
DcbLeaderID.Enabled = True
TxtLeaderName.Enabled = False
TxtLeaderName.text = ""
ElseIf ChLeaderType(1).value = True Then
TxtNameE.Enabled = False
DcbLeaderID.Enabled = False
TxtLeaderName.Enabled = True
DcbLeaderID.BoundText = 0
TxtNameE.text = ""
End If
End Sub

Private Sub DcbDrievID_Change()
DcbDrievID_Click (0)
End Sub

Private Sub DcbDrievID_Click(Area As Integer)
    If val(DcbDrievID.BoundText) = 0 Then Exit Sub
      Dim EmpCode  As String
      GetEmployeeIDFromCode , , DcbDrievID.BoundText, EmpCode
      Text6.text = EmpCode
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
If Me.TxtModFlg2(mIndex).text <> "R" Then
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
Me.TxtOperatorN.text = IIf(IsNull(Rs3("OperatorN").value), "", Rs3("OperatorN").value)
End If
If Typ <> 2 Then
txtBoardNo.text = IIf(IsNull(Rs3("BoardNO").value), "", Rs3("BoardNO").value)
End If
If Typ <> 0 Then
DcbEquepment.BoundText = IIf(IsNull(Rs3("FixedassetId").value), 0, Rs3("FixedassetId").value)
End If
DcbLeaderID.BoundText = IIf(IsNull(Rs3("Emp_id").value), 0, Rs3("Emp_id").value)
Else
If Typ <> 1 Then
TxtOperatorN.text = ""
End If
If Typ <> 2 Then
txtBoardNo.text = ""
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
      TxtNameE.text = EmpCode
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

Private Sub dcItems_Validate(Cancel As Boolean)
CmdSearch_Click
End Sub

Private Sub Grid6_EnterCell()
 On Error GoTo ErrTrap
    FindRec val(Me.Grid6.TextMatrix(Me.Grid6.Row, Me.Grid6.ColIndex("id")))
ErrTrap:
End Sub

Private Sub txtNamee_KeyPress(KeyAscii As Integer)
    Dim EmpID As Integer
    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode TxtNameE.text, EmpID
        DcbLeaderID.BoundText = EmpID
    End If
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
    Dim EmpID As Integer
    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode Text6.text, EmpID
        DcbDrievID.BoundText = EmpID
    End If
End Sub

Private Sub TxtBoardNO_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
RetriveCarsInfo , , txtBoardNo.text, 2
End If
End Sub



Private Sub TxtOperatorN_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
RetriveCarsInfo , TxtOperatorN.text, , 1
End If
End Sub

Private Sub txtOrderMaintinNo_Change()
    If Me.TxtModFlg2(mIndex).text = "N" Or Me.TxtModFlg2(mIndex).text = "E" Then
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


            

     TxtDeptNotes.text = IIf(IsNull(rs("DeptNotes").value), "", rs("DeptNotes").value)
     TxtInitialNotes.text = IIf(IsNull(rs("InitialNotes").value), "", rs("InitialNotes").value)
     
     
    
     Me.DcboEmpName.BoundText = IIf(IsNull(rs("SuperVisor").value), "", rs("SuperVisor").value)
     DcbEquepment.BoundText = IIf(IsNull(rs("EquepID").value), "", rs("EquepID").value)
     'txtRemark.text = IIf(IsNull(rs("Remarks").value), "", rs("Remarks").value)
     DcbType.ListIndex = val(IIf(IsNull(rs("TypeMaint").value), -1, rs("TypeMaint").value))
     TxtDes.text = IIf(IsNull(rs("Des").value), "", rs("Des").value)
  

     
     DcbStutsMaint.ListIndex = IIf(IsNull(rs("StutsMaint").value), -1, rs("StutsMaint").value)
   
    txtBoardNo.text = IIf(IsNull(rs("BoardNO").value), "", rs("BoardNO").value)
    TxtOperatorN.text = IIf(IsNull(rs("OperatorN").value), "", rs("OperatorN").value)
    '******************************************

    
    
    
        
       ''///////////////////////
       ''04 05 2016
       Me.DcbBranchFrom.BoundText = IIf(IsNull(rs("DcbBranchFrom").value), "", rs("DcbBranchFrom").value)
       Me.DcbLeaderID.BoundText = IIf(IsNull(rs("LeaderID").value), "", rs("LeaderID").value)
       Me.TxtLeaderName.text = IIf(IsNull(rs("LeaderName").value), "", rs("LeaderName").value)
       Me.TxtDrievName.text = IIf(IsNull(rs("DrievName").value), "", rs("DrievName").value)
       
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
        GetEmployeeIDFromCode TxtSearchCode.text, EmpID
        DcboEmpName.BoundText = EmpID
    End If
End Sub


Private Sub DcboEmpName_Click(Area As Integer)
On Error Resume Next
      If val(DcboEmpName.BoundText) = 0 Then Exit Sub
    Dim EmpCode  As String
    GetEmployeeIDFromCode , , DcboEmpName.BoundText, EmpCode
    TxtSearchCode.text = EmpCode
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

Private Sub txtLetter3_KeyPress(KeyAscii As Integer)
txtLetter3.text = ""
If Len(txtLetter3.text) > 0 Then
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

    If TxtVac_ID.text <> "" Then
        If CheckDelCountry(val(Me.TxtVac_ID.text)) = False Then
            Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã·...!!!"
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            Exit Sub
        End If

        MSGType = MsgBox("Â·  —€» ›Ì Õ–› Â–« «·”Ã·", vbQuestion + vbMsgBoxRtlReading + vbYesNo + vbMsgBoxRight, App.Title)

        If MSGType = vbYes Then
            RsSavRec.Find "id=" & val(TxtVac_ID.text), , adSearchForward, 1
            RsSavRec.delete
            MsgBox " „  ⁄„·Ì… «·Õ–› »‰Ã«Õ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
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
            MsgBox StrMSG, vbOKOnly + vbMsgBoxRight, App.Title
            'clear the ConnectiOn Errors
            Cn.Errors.Clear
    End Select

End Sub

Private Sub BtnFirst_Click()
    On Error GoTo ErrTrap

    If Me.TxtModFlg.text = "N" Then
        FindRec val(TxtVac_ID.text)
        Me.TxtModFlg.text = "R"
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
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.Title
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

    My_SQL = "project_status"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable

    If rs.RecordCount > 0 Then
        TxtSerial.text = rs.RecordCount + 1
    Else
        TxtSerial.text = 1
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
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.Title
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

    StrVacName = IsRecExist("project_status", "name", Trim(TxtVacName.text), "name", "Vac_ID<>'" & Trim(TxtVac_ID.text) & "'")

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
    txtColor.text = Color
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
    Dim i      As Integer
    Dim My_SQL As String
    
    TabMain.TabVisible(1) = False
    TabMain.TabVisible(2) = False
    TabMain.TabVisible(0) = False
    TabMain.TabVisible(3) = False
    TabMain.TabVisible(4) = False
    TabMain.TabVisible(5) = False
    TabMain.TabVisible(6) = False
'    TabMain.TabVisible(7) = False
    
    
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
        
    ElseIf mIndex = 5 Then
        TabMain.TabVisible(4) = True
        TabMain.CurrTab = 4
        Me.Width = Grid3.Width + 400
        
    ElseIf mIndex = 6 Then
        TabMain.TabVisible(5) = True
        TabMain.CurrTab = 5
        Me.Width = Grid6.Width + 400
        Frame7.Enabled = True
        Ele(6).Enabled = True
        
    ElseIf mIndex = 7 Then
        TabMain.TabVisible(6) = True
        TabMain.CurrTab = 6
        Me.Width = Grid7.Width + 400
        Frame8.Enabled = True
        Ele(7).Enabled = True
    End If
    If mIndex = 3 Then mIndex = 0
    If mIndex = 4 Then mIndex = 3
    
    If mIndex = 0 Then
        My_SQL = "project_status"
        Set BKGrndPic = New ClsBackGroundPic
        Set RsSavRec = New ADODB.Recordset
        RsSavRec.CursorLocation = adUseClient
        RsSavRec.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdTableDirect
        Me.TxtModFlg.text = "R"
        
        My_SQL = "select UserID,UserName From tblUsers "
        fill_combo DCUser, My_SQL
    
        FillGridWithData
        Set ISButton1.ButtonImage = mdifrmmain.ImgLstTree.ListImages("GridOptions").Picture
        
        With Me.Grid
            .cell(flexcpPicture, 0, .ColIndex("name")) = Me.GrdImageList.ListImages("Vac_Name").ExtractIcon
            .cell(flexcpPicture, 0, .ColIndex("Ser")) = Me.GrdImageList.ListImages("Ser").ExtractIcon
        
            For i = 0 To .Cols - 1
                .cell(flexcpPictureAlignment, 0, i) = flexPicAlignRightCenter
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
        Dim StrLogFileName As String
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
        
    ElseIf mIndex = 7 Then
     
        Me.Caption = "»Ì«‰«  «·“ÌÊ "
        My_SQL = "tblOilsTypes"
        Set BKGrndPic = New ClsBackGroundPic
        Set RsSavRec = New ADODB.Recordset
        RsSavRec.CursorLocation = adUseClient
        RsSavRec.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdTableDirect
        btn_First_Click (mIndex)
        Me.Width = Grid7.Width + 400
        FillGridWithData7
    ElseIf mIndex = 6 Then
       Frame7.Enabled = True
        Me.Caption = "»Ì«‰«  «·„÷Œ« "
        My_SQL = "tblPumpType"
        Set BKGrndPic = New ClsBackGroundPic
        Set RsSavRec = New ADODB.Recordset
        RsSavRec.CursorLocation = adUseClient
        RsSavRec.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdTableDirect
        btn_First_Click (mIndex)
        Me.Width = Grid6.Width + 400
        
        '****************
        Set Dcombos = New ClsDataCombos
        Dcombos.GetBoxes Me.TxtBox
        Dcombos.GetItemsNames Me.TxtItem
        
        Dcombos.GetStores cmbStore
        Dcombos.GetStores cmbStore2
         
        Dcombos.GetAccountingCodes Me.cmbAccount, True
        Dcombos.GetAccountingCodes Me.cmbAccountComm, True
        
        '****************
        FillGridWithData6
     
   
    End If
    
    If mIndex = 1 Then
                
        My_SQL = "ShiftRec"
        ' Set BKGrndPic = New ClsBackGroundPic
        Set RsSavRec = New ADODB.Recordset
        RsSavRec.CursorLocation = adUseClient
        RsSavRec.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdTableDirect
        TxtModFlg2(mIndex).text = "R"
        DCboUserName(mIndex) = user_id
        
        btn_First_Click (mIndex)
    ElseIf mIndex = 2 Then
   
        '
        My_SQL = "ShiftRec"
        'Set BKGrndPic2 = New ClsBackGroundPic
        Set RsSavRec = New ADODB.Recordset
        RsSavRec.CursorLocation = adUseClient
        RsSavRec.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdTableDirect
        TxtModFlg2(mIndex).text = "R"
        DCboUserName(mIndex) = user_id
        btn_First_Click (mIndex)
        
    ElseIf mIndex = 5 Then
    
        Me.Caption = "«‰Ê«⁄ ﬁÿ⁄ «·€Ì«—"
        My_SQL = "tblCarMaint"
        Set BKGrndPic = New ClsBackGroundPic
        Set RsSavRec = New ADODB.Recordset
        RsSavRec.CursorLocation = adUseClient
        RsSavRec.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdTableDirect
        btn_First_Click (mIndex)
        Me.Width = Grid3.Width + 400
        
        Set Dcombos = New ClsDataCombos
        
        Dcombos.GetItemSGroups Me.DcbCarGroup, False
        Dim GID As Integer
        GID = val(DcbCarGroup.BoundText)
        Dim StrSQL As String
        StrSQL = ""
        StrSQL = StrSQL & "SELECT ItemID, "
        StrSQL = StrSQL & "        "
        StrSQL = StrSQL & "       ItemName "
        StrSQL = StrSQL & "FROM TblItems "
        StrSQL = StrSQL & "WHERE GroupID = " & GID & ";"
         
        GetComboData dcitems, StrSQL
        '*********************
        If SystemOptions.UserInterface = EnglishInterface Then
            My_SQL = "SELECT id,ISNULL(ModelE,Model) ModelName from TblCarModels"
        Else
            My_SQL = "SELECT id, Model from TblCarModels"
        End If
        fill_combo DcbCarModel, My_SQL
            
        Dcombos.GetTblCarsDataGroup Me.DcbCarType
        
        TxtModFlg2(mIndex).text = "R"
        DCboUserName(mIndex) = user_id
        btn_First_Click (mIndex)
             
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
        
        Dim str As String
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

Public Sub FiLLRec5()
    On Error GoTo ErrTrap
    
    If TxtModFlg2(mIndex).text = "N" Then
        ' TxtSerial1(mIndex).Text = CStr(new_id("TblAdditionsAssest", "ID", "", True))
       
        RsSavRec.AddNew
        TxtSerial1(mIndex).text = new_id("tblCarMaint", "id", "")
        RsSavRec.Fields("ID").value = val(TxtSerial1(mIndex).text)
    End If
    RsSavRec("NoteSerial1").value = Trim$(Me.TxtSerial1(mIndex).text)
    ' RsSavRec.Fields("BranchID").value = IIf(Dcbranch(mIndex).text <> "", Trim(Dcbranch(mIndex).BoundText), Null)
    RsSavRec("RecordDate").value = XPDtbTrans.value
    RsSavRec.Fields("GroupId").value = val(DcbCarGroup.BoundText)
   
    RsSavRec.Fields("TypeId").value = val(DcbCarType.BoundText)
    RsSavRec.Fields("ModelId").value = val(DcbCarModel.BoundText)
    RsSavRec.Fields("ItemId").value = val(dcitems.BoundText)
    RsSavRec.Fields("ChkOrg").value = chkOrg.value = vbChecked
    RsSavRec.Fields("chkCom").value = chkcom.value = vbChecked
    RsSavRec.Fields("chkTested").value = chkTestd.value = vbChecked
    RsSavRec.Fields("chkNormal").value = chknormal.value = vbChecked
     
    '*********************
    RsSavRec.update
 
    If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox " „  ⁄„·Ì… «·Õ›Ÿ »‰Ã«Õ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    Else
        MsgBox "Save Successfully", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    End If
    'CuurentLogdata
    
    TxtModFlg2(mIndex) = "R"

    Exit Sub
ErrTrap:

    If RsSavRec.EditMode <> adEditNone Then
        RsSavRec.CancelUpdate
    End If

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

    RsSavRec.Fields("name").value = IIf(TxtVacName.text <> "", Trim(TxtVacName.text), Null)
    RsSavRec.Fields("namee").value = IIf(TxtVacNamee.text <> "", Trim(TxtVacNamee.text), Null)
    
    RsSavRec.Fields("color").value = IIf(txtColor.text <> "", Trim(txtColor.text), Null)
    RsSavRec.update
    MsgBox " „  ⁄„·Ì… «·Õ›Ÿ »‰Ã«Õ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
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
    TxtVac_ID.text = IIf(IsNull(RsSavRec.Fields("id").value), "", RsSavRec.Fields("id").value)
    TxtVacName.text = IIf(IsNull(RsSavRec.Fields("name").value), "", RsSavRec.Fields("name").value)
    TxtVacNamee.text = IIf(IsNull(RsSavRec.Fields("namee").value), "", RsSavRec.Fields("namee").value)
    
    txtColor.text = IIf(IsNull(RsSavRec.Fields("color").value), "", RsSavRec.Fields("color").value)

    If txtColor.text <> "" Then
        Combo1.backcolor = txtColor.text
    End If

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

    If TxtModFlg.text = "N" Then
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
    
    ElseIf TxtModFlg.text = "R" Then
        Frm2.Enabled = False
        Grid.Enabled = True
        btnModify.Enabled = False
        btnDelete.Enabled = False

        If TxtVac_ID.text <> "" Then
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
    
    ElseIf TxtModFlg.text = "E" Then
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
                .cell(flexcpBackColor, i, 4, i, 4) = IIf(IsNull(rs.Fields("color").value), "", rs.Fields("color").value)
            
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


Public Sub FillGridWithData7()

    On Error GoTo ErrTrap

    Dim i As Integer
    Dim rs As ADODB.Recordset
    Dim My_SQL As String

    Set rs = New ADODB.Recordset
    My_SQL = "select * From tblOilsTypes order by id"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText

    With Me.Grid7
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



Public Sub FillGridWithData6(Optional ByVal mStoreId As Long = 0)

    On Error GoTo ErrTrap

    Dim i      As Integer
    Dim rs     As ADODB.Recordset
    Dim My_SQL As String

    Set rs = New ADODB.Recordset
    If mStoreId = 0 Then
        My_SQL = "select * From tblpumpType order by id"
    Else
        My_SQL = "select * From tblpumpType where  storeid = " & mStoreId & " order by id"
    End If
    rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText

    With Me.Grid6
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

Public Sub FillGridWithData2()

    On Error GoTo ErrTrap

    Dim i As Integer
    Dim rs As ADODB.Recordset
    Dim My_SQL As String

    Set rs = New ADODB.Recordset
    My_SQL = "select * From ShiftMaintType order by id"
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








 





Private Sub Cmd_DeleteRow_Click(Index As Integer)
If Me.TxtModFlg.text <> "R" Then
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
                .rows = 2
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
            RsSavRec.Find "id=" & val(TxtSerial1(mIndex).text), , adSearchForward, 1
            CuurentLogdata ("D")
            RsSavRec.delete
            Dim s As String

            If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox " „  ⁄„·Ì… «·Õ–› »‰Ã«Õ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
            Else
            MsgBox "Delete Successfully", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
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
            MsgBox StrMSG, vbOKOnly + vbMsgBoxRight, App.Title
            'clear the ConnectiOn Errors
            Cn.Errors.Clear
    End Select

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
    If mIndex = 1 Then
        FiLLTXT1
    ElseIf mIndex = 2 Then
        FiLLTXT2
    ElseIf mIndex = 3 Then
        FiLLTXT3
    ElseIf mIndex = 5 Then
        FiLLTXT5
    ElseIf mIndex = 6 Then
        FiLLTXT6
    ElseIf mIndex = 7 Then
        FiLLTXT7
        
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

Private Sub btn_Modify_Click(Index As Integer)
    Dim Msg As String

    If DoPremis(Do_Edit, Me.Name, True) = False Then
        Exit Sub
    End If

    On Error GoTo ErrTrap

    If TxtSerial1(mIndex).text <> "" Then
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
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.Title
        
            If RsSavRec.EditMode <> adEditNone Then
                RsSavRec.CancelUpdate
                'RsSavRec.Requery
            End If

    End Select

End Sub

Public Sub btn_New_Click(Index As Integer)
    Dim My_SQL As String
    Dim rs     As ADODB.Recordset

    If DoPremis(Do_New, Me.Name, True) = False Then
        Exit Sub
    End If

    On Error GoTo ErrTrap
    Set rs = New ADODB.Recordset
    Frame1.Enabled = True
    Frame3.Enabled = True
    clear_all Me
    TxtModFlg2(mIndex).text = "N"
    If mIndex = 1 Then
        My_SQL = "ShiftMaintType"
        'DCboUserName(mIndex) = user_id
     
        clear_all Me
   
        rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable

        If rs.RecordCount > 0 Then
            TxtSerial1(mIndex).text = rs.RecordCount + 1
        Else
            TxtSerial1(mIndex).text = 1
        End If

        rs.Close
        'CmbType.ListIndex = 0
        TxtName.SetFocus
    
    ElseIf mIndex = 2 Then
       
        DCboUserName(mIndex) = user_id
        My_SQL = "ShiftRec"
        'DCboUserName(mIndex) = user_id
     
        clear_all Me
   
        rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable

        If rs.RecordCount > 0 Then
            TxtSerial1(mIndex).text = rs.RecordCount + 1
        Else
            TxtSerial1(mIndex).text = 1
        End If

        rs.Close
        
    ElseIf mIndex = 3 Then
        My_SQL = "tblDoctorsType"
        'DCboUserName(mIndex) = user_id
     
        clear_all Me
   
        rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable

        If rs.RecordCount > 0 Then
            TxtSerial1(mIndex).text = rs.RecordCount + 1
        Else
            TxtSerial1(mIndex).text = 1
        End If

        rs.Close
        'CmbType.ListIndex = 0
        TxtName3.SetFocus
  
      ElseIf mIndex = 3 Then
        My_SQL = "tblDoctorsType"
        'DCboUserName(mIndex) = user_id
     
        clear_all Me
   
        rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable

        If rs.RecordCount > 0 Then
            TxtSerial1(mIndex).text = rs.RecordCount + 1
        Else
            TxtSerial1(mIndex).text = 1
        End If

        rs.Close
        'CmbType.ListIndex = 0
        TxtName3.SetFocus
      
    ElseIf mIndex = 7 Then
        My_SQL = "tblOilsTypes"
        'DCboUserName(mIndex) = user_id
     
        clear_all Me
   
        rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable

        If rs.RecordCount > 0 Then
            TxtSerial1(mIndex).text = rs.RecordCount + 1
        Else
            TxtSerial1(mIndex).text = 1
        End If

        rs.Close
        'CmbType.ListIndex = 0
        txtOilName.SetFocus
    ElseIf mIndex = 6 Then
    Frame7.Enabled = True
        My_SQL = "tblPumpType"
        'DCboUserName(mIndex) = user_id
     
        clear_all Me
   
        rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable

        If rs.RecordCount > 0 Then
            TxtSerial1(mIndex).text = rs.RecordCount + 1
        Else
            TxtSerial1(mIndex).text = 1
        End If

        rs.Close
        'CmbType.ListIndex = 0
        txtpumpName.SetFocus
        
    ElseIf mIndex = 5 Then
        My_SQL = "tblCarMaint"
        'DCboUserName(mIndex) = user_id
     
        clear_all Me
   
        rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable

        If rs.RecordCount > 0 Then
            TxtSerial1(mIndex).text = rs.RecordCount + 1
        Else
            TxtSerial1(mIndex).text = 1
        End If

        rs.Close
        'CmbType.ListIndex = 0
        DcbCarGroup.SetFocus
        
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

    If mIndex = 1 Then
        FiLLTXT1
    ElseIf mIndex = 2 Then
        FiLLTXT2
    ElseIf mIndex = 3 Then
        FiLLTXT3
    ElseIf mIndex = 7 Then
        FiLLTXT7
    ElseIf mIndex = 5 Then
        FiLLTXT5
    ElseIf mIndex = 6 Then
        FiLLTXT6

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

    If mIndex = 1 Then
        FiLLTXT1
    ElseIf mIndex = 2 Then
        FiLLTXT2
    ElseIf mIndex = 3 Then
        FiLLTXT3
    ElseIf mIndex = 7 Then
        FiLLTXT7
    ElseIf mIndex = 5 Then
        FiLLTXT5
    ElseIf mIndex = 6 Then
        FiLLTXT6

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
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
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

 
        xReport.ParameterFields(7).AddCurrentValue dcBranch(mIndex).text
        
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
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
    Dim Msg        As String
    Dim StrVacCode As String
    Dim StrVacName As String
    Dim CtrlTxt    As Control
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
        If dcBranch(mIndex).text = "" Then
            MsgBox "Please Enter Branch"
            dcBranch(mIndex).SetFocus
            Exit Sub
        End If
    End If
    
    '------------------------------ check if Empcode exist ----------------------

    ' -------------------------------------- txtmodflg type -------------------
    Select Case TxtModFlg2(mIndex).text

            '------------------------------ new record ----------------------------
        Case "N"
      
            '------------------------- save record -----------------------------
            If mIndex = 1 Then
                AddNewRec
            ElseIf mIndex = 2 Then
                FiLLRec2
            ElseIf mIndex = 3 Then
                AddNewRec
            ElseIf mIndex = 5 Then
                FiLLRec5
            ElseIf mIndex = 7 Then
                AddNewRec
                FiLLRec7
                ElseIf mIndex = 6 Then
                FiLLRec6
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
            ElseIf mIndex = 5 Then
                FiLLRec5
            ElseIf mIndex = 6 Then
                FiLLRec6
            ElseIf mIndex = 7 Then
                FiLLRec7

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
    If mIndex = 1 Then
        FiLLTXT1
    
    ElseIf mIndex = 2 Then
        FiLLTXT2
    ElseIf mIndex = 3 Then
        FiLLTXT3
    ElseIf mIndex = 7 Then
        FiLLTXT7
    ElseIf mIndex = 5 Then
        FiLLTXT5
    ElseIf mIndex = 6 Then
        FiLLTXT6
    
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
 
Private Sub Undo()
    On Error GoTo ErrTrap

    Select Case TxtModFlg2(mIndex).text

        Case "N"
            clear_all Me
            TxtModFlg2(mIndex).text = "R"
           
            btn_First_Click (mIndex)
        Case "E"
            RsSavRec.Find "ID='" & val(TxtSerial1(mIndex).text) & "'", , adSearchForward, adBookmarkFirst

            If RsSavRec.EOF Or RsSavRec.BOF Then
                TxtModFlg2(mIndex).text = "R"
                Exit Sub
            End If

            If mIndex = 1 Then
                FiLLTXT1
            ElseIf mIndex = 2 Then
                FiLLTXT2
            ElseIf mIndex = 3 Then
                FiLLTXT3
            ElseIf mIndex = 7 Then
                FiLLTXT7
            ElseIf mIndex = 5 Then
                FiLLTXT5
            ElseIf mIndex = 6 Then
                FiLLTXT6

            End If
            TxtModFlg2(mIndex).text = "R"
    End Select

    Exit Sub
ErrTrap:
End Sub
 
 

Function CuurentLogdata(Optional Currentmode As String)
     LogTextA = "    ‘«‘… " & ScreenNameArabic & CHR(13) & "ﬂÊœ " & TxtSerial.text & CHR(13) & "   «·‰Ê⁄ " & TxtVacName
        LogTexte = "    Screen  " & ScreenNameEnglish & CHR(13) & "Code  " & TxtSerial.text & CHR(13) & "   Type " & TxtVacName
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
        mSelectModFlg = Me.TxtModFlg.text
    Else
        mSelectModFlg = Me.TxtModFlg2(mIndex).text
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
ElseIf mIndex = 7 Then
            
            StrRecID = new_id("tblOilsTypes", "id", "")
            RsSavRec.AddNew
            RsSavRec.Fields("id").value = IIf(StrRecID <> "", StrRecID, Null)
 
            FiLLRec7
 ElseIf mIndex = 5 Then

        
        StrRecID = new_id("tblCarMaint", "id", "")
        RsSavRec.AddNew
        RsSavRec.Fields("id").value = IIf(StrRecID <> "", StrRecID, Null)
        FiLLRec5




        
    End If
    
    
ErrTrap:
   
  
    

End Sub




Public Sub FiLLRec1()
    On Error GoTo ErrTrap

    RsSavRec.Fields("name").value = IIf(TxtName.text <> "", Trim(TxtName.text), Null)
    RsSavRec.Fields("namee").value = IIf(TxtNameE.text <> "", Trim(TxtNameE.text), Null)
    

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


Public Sub FiLLRec3()
    On Error GoTo ErrTrap

    RsSavRec.Fields("name").value = IIf(TxtName3.text <> "", Trim(TxtName3.text), Null)
    RsSavRec.Fields("namee").value = IIf(txtNameE3.text <> "", Trim(txtNameE3.text), Null)
    RsSavRec.Fields("Period").value = IIf(txtPercentV.text <> "", Trim(txtPercentV.text), Null)
    

    RsSavRec.update
    MsgBox " „  ⁄„·Ì… «·Õ›Ÿ »‰Ã«Õ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    FillGridWithData3
    TxtModFlg2(mIndex) = "R"

    Exit Sub
ErrTrap:

    If RsSavRec.EditMode <> adEditNone Then
        RsSavRec.CancelUpdate
    End If

End Sub



Public Sub FiLLRec7()
    On Error GoTo ErrTrap

    RsSavRec.Fields("name").value = IIf(txtOilName.text <> "", Trim(txtOilName.text), Null)
    RsSavRec.Fields("namee").value = IIf(txtOilNameE.text <> "", Trim(txtOilNameE.text), Null)
    RsSavRec.Fields("Period").value = IIf(txtPeriod.text <> "", Trim(txtPeriod.text), Null)
    RsSavRec.Fields("KiloMetr").value = IIf(txtKiloMetr.text <> "", Trim(txtKiloMetr.text), Null)
    

    RsSavRec.update
    MsgBox " „  ⁄„·Ì… «·Õ›Ÿ »‰Ã«Õ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    FillGridWithData7
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

       
            RsSavRec.AddNew
            TxtSerial1(mIndex).text = new_id("tblPumpType", "id", "")
            'RsSavRec.Fields("ID").value = val(TxtSerial1(mIndex).text)
    End If

    RsSavRec.Fields("name").value = IIf(txtpumpName.text <> "", Trim(txtpumpName.text), Null)
    RsSavRec.Fields("namee").value = IIf(txtpumpNameE.text <> "", Trim(txtpumpNameE.text), Null)
    RsSavRec.Fields("PercentV").value = IIf(txtPunpPer.text <> "", Trim(txtPunpPer.text), Null)
    
    RsSavRec.Fields("AmountH").value = IIf(txtAmountH.text <> "", Trim(txtAmountH.text), Null)
    RsSavRec.Fields("IsOther").value = chkIsOther.value = vbChecked
    RsSavRec.Fields("Account_Code").value = Trim(Me.cmbAccount.BoundText)
    RsSavRec.Fields("Account_CodeComm").value = Trim(Me.cmbAccountComm.BoundText)
    
    RsSavRec.Fields("BoxId").value = val(Me.TxtBox.BoundText)
    RsSavRec.Fields("ItemID").value = val(Me.TxtItem.BoundText)
    RsSavRec.Fields("StoreID").value = val(Me.cmbStore.BoundText)
    

    RsSavRec.update
    MsgBox " „  ⁄„·Ì… «·Õ›Ÿ »‰Ã«Õ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    FillGridWithData6
    TxtModFlg2(mIndex) = "R"

    Exit Sub
ErrTrap:

    If RsSavRec.EditMode <> adEditNone Then
        RsSavRec.CancelUpdate
    End If

End Sub





Public Sub FiLLRec2()
    On Error GoTo ErrTrap
    If TxtModFlg2(mIndex).text = "N" Then
           ' TxtSerial1(mIndex).Text = CStr(new_id("TblAdditionsAssest", "ID", "", True))

       
            RsSavRec.AddNew
            TxtSerial1(mIndex).text = new_id("ShiftRec", "id", "")
            RsSavRec.Fields("ID").value = val(TxtSerial1(mIndex).text)
    End If
    RsSavRec.Fields("ShiftMaintTypeID").value = IIf(cmbShiftMaintType.text <> "", Trim(cmbShiftMaintType.BoundText), Null)
    RsSavRec.Fields("BranchID").value = IIf(dcBranch(mIndex).text <> "", Trim(dcBranch(mIndex).BoundText), Null)
    RsSavRec("RecordDate").value = XPDtbRecordDate.value
   
     RsSavRec("DateRec").value = txtDateRec.value
    RsSavRec("TimeRec").value = txtTimeRec.value
    
     RsSavRec("DateEnd").value = txtDateEnd.value
    RsSavRec("TimeEnd").value = txtTimeEnd.value
    

               
   RsSavRec("OrderMaintinNo").value = val(txtOrderMaintinNo.text)
   RsSavRec("typemaint").value = DcbType.ListIndex
   
    RsSavRec("NoteDone").value = txtNoteDone.text
     
    RsSavRec("NoteStill").value = txtNoteStill.text
    RsSavRec("Remarks").value = txtRemarks.text
    RsSavRec("NoteLate").value = txtNoteLate.text
    RsSavRec("CarStatus").value = cmbCarStatus.ListIndex
    RsSavRec("StutsMaint").value = DcbStutsMaint.ListIndex

    
    DCboUserName(mIndex).BoundText = IIf(DCboUserName(mIndex).text <> "", Trim(DCboUserName(mIndex).BoundText), user_id)
    RsSavRec.Fields("UserID").value = IIf(DCboUserName(mIndex).text <> "", Trim(DCboUserName(mIndex).BoundText), user_id)
     
   

    RsSavRec.update
    
  
    
   
If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox " „  ⁄„·Ì… «·Õ›Ÿ »‰Ã«Õ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
 Else
   MsgBox "Save Successfully", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
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
    TxtSerial1(mIndex).text = IIf(IsNull(RsSavRec.Fields("id").value), "", RsSavRec.Fields("id").value)
    TxtName.text = IIf(IsNull(RsSavRec.Fields("name").value), "", RsSavRec.Fields("name").value)
    TxtNameE.text = IIf(IsNull(RsSavRec.Fields("nameE").value), "", RsSavRec.Fields("nameE").value)
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
Public Sub FiLLTXT6()

    On Error GoTo ErrTrap
    Dim i As Integer
    Frm2.Enabled = False
    TxtSerial1(mIndex).text = IIf(IsNull(RsSavRec.Fields("id").value), "", RsSavRec.Fields("id").value)
    txtpumpName.text = IIf(IsNull(RsSavRec.Fields("name").value), "", RsSavRec.Fields("name").value)
    txtpumpNameE.text = IIf(IsNull(RsSavRec.Fields("nameE").value), "", RsSavRec.Fields("nameE").value)
    txtPunpPer.text = IIf(IsNull(RsSavRec.Fields("PercentV").value), "", RsSavRec.Fields("PercentV").value)
    Me.TxtBox.BoundText = IIf(IsNull(RsSavRec("BoxId").value), "", RsSavRec("BoxId").value)
    
    Me.txtAmountH.text = IIf(IsNull(RsSavRec("AmountH").value), "", RsSavRec("AmountH").value)
    cmbAccount.BoundText = IIf(IsNull(RsSavRec("Account_Code").value), "", RsSavRec("Account_Code").value)
    cmbAccountComm.BoundText = IIf(IsNull(RsSavRec("Account_CodeComm").value), "", RsSavRec("Account_CodeComm").value)
    Me.TxtItem.BoundText = IIf(IsNull(RsSavRec("ItemId").value), "", RsSavRec("ItemId").value)
    Me.cmbStore.BoundText = IIf(IsNull(RsSavRec("StoreId").value), "", RsSavRec("StoreId").value)
    chkIsOther.value = IIf(RsSavRec.Fields("IsOther").value, vbChecked, vbUnchecked)
    'LabCurr_Rec(mIndex).Caption = RsSavRec.AbsolutePosition
'    LabCount_Rec(mIndex).Caption = RsSavRec.RecordCount

    With Grid6

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
Public Sub FiLLTXT3()

    On Error GoTo ErrTrap
    Dim i As Integer
    Frm2.Enabled = False
    TxtSerial1(mIndex).text = IIf(IsNull(RsSavRec.Fields("id").value), "", RsSavRec.Fields("id").value)
    TxtName3.text = IIf(IsNull(RsSavRec.Fields("name").value), "", RsSavRec.Fields("name").value)
    txtNameE3.text = IIf(IsNull(RsSavRec.Fields("nameE").value), "", RsSavRec.Fields("nameE").value)
    txtPercentV.text = IIf(IsNull(RsSavRec.Fields("PercentV").value), "", RsSavRec.Fields("PercentV").value)
    
    LabCurr_Rec(mIndex).Caption = RsSavRec.AbsolutePosition
    LabCount_Rec(mIndex).Caption = RsSavRec.RecordCount

    With Grid3

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



Public Sub FiLLTXT7()

    On Error GoTo ErrTrap
    Dim i As Integer
    'Frm2.Enabled = False
    TxtSerial1(mIndex).text = IIf(IsNull(RsSavRec.Fields("id").value), "", RsSavRec.Fields("id").value)
    txtOilName.text = IIf(IsNull(RsSavRec.Fields("name").value), "", RsSavRec.Fields("name").value)
    txtOilNameE.text = IIf(IsNull(RsSavRec.Fields("nameE").value), "", RsSavRec.Fields("nameE").value)
    txtPeriod.text = IIf(IsNull(RsSavRec.Fields("Period").value), "", RsSavRec.Fields("Period").value)
    txtKiloMetr.text = IIf(IsNull(RsSavRec.Fields("KiloMetr").value), "", RsSavRec.Fields("KiloMetr").value)
    
    LabCurr_Rec(mIndex).Caption = RsSavRec.AbsolutePosition
    LabCount_Rec(mIndex).Caption = RsSavRec.RecordCount

    With Grid7

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


Public Sub FiLLTXT5()

    On Error GoTo ErrTrap
    Dim i As Integer
    
    TxtSerial1(mIndex).text = IIf(IsNull(RsSavRec.Fields("id").value), "", RsSavRec.Fields("id").value)
    
    XPDtbTrans.value = RsSavRec("RecordDate").value
    DcbCarGroup.BoundText = RsSavRec.Fields("GroupId").value
    DcbCarGroup_Click 5
    DcbCarType.BoundText = RsSavRec.Fields("TypeId").value
    DcbCarModel.BoundText = RsSavRec.Fields("ModelId").value
    dcitems.BoundText = RsSavRec.Fields("ItemId").value
    chkOrg.value = IIf(RsSavRec.Fields("ChkOrg").value, vbChecked, vbUnchecked)
    chkcom.value = IIf(RsSavRec.Fields("chkCom").value, vbChecked, vbUnchecked)
    chkTestd.value = IIf(RsSavRec.Fields("chkTested").value, vbChecked, vbUnchecked)
    chknormal.value = IIf(RsSavRec.Fields("chkNormal").value, vbChecked, vbUnchecked)

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
     
    

    
    DCboUserName(mIndex).BoundText = IIf(DCboUserName(mIndex).text <> "", Trim(DCboUserName(mIndex).BoundText), user_id)
    RsSavRec.Fields("UserID").value = IIf(DCboUserName(mIndex).text <> "", Trim(DCboUserName(mIndex).BoundText), user_id)
      
     
     
    TxtSerial1(mIndex).text = IIf(IsNull(RsSavRec.Fields("id").value), "", RsSavRec.Fields("id").value)
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
    txtRemarks.text = IIf(IsNull(RsSavRec("Remarks").value), "", RsSavRec("Remarks").value)
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
    TxtMod = TxtModFlg.text
    TxtModFlg.text = ""
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

        ElseIf mIndex2 = 6 Then
            FiLLTXT6
        ElseIf mIndex2 = 7 Then
            FiLLTXT7
            
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
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "New" & Wrap & "To open a new record" & Wrap & "Press this key" & Wrap & "or Key" & " F12 √Ê Enter"
            
        .AddControl btnNew, Msg, True
    End With

    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Modification" & Wrap & "Modifying the current record " & Wrap & "Press this key" & Wrap & "or Key" & " F11"
        .AddControl btnModify, Msg, True
    End With

    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Save" & Wrap & "To record the data within the database " & Wrap & "Press this key" & Wrap & "or Key" & " F10"
        .AddControl btnSave, Msg, True
    End With

    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Undo" & Wrap & "To undo the current operation" & Wrap & "Press this key" & Wrap & "or Key" & " F9"
        .AddControl BtnUndo, Msg, True
    End With

    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Delete the record" & Wrap & "To delete the current record" & Wrap & "Press this key" & Wrap & "or Key" & " F18"
        .AddControl btnDelete, Msg, True
    End With

    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Exit" & Wrap & "To close this window" & Wrap & "Press this key" & Wrap & "or Key" & " Ctrl+x"
        .AddControl btnCancel, Msg, True
    End With

    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "First" & Wrap & "Move first record" & Wrap & "Press this key" & Wrap & "or Key" & " Home √Ê UpArrow"
        .AddControl btnFirst, Msg, True
    End With

    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Previous" & Wrap & "Move previous record" & Wrap & "Press this key" & Wrap & "or Key" & " PageUp √Ê LeftArrow"
        .AddControl btnPrevious, Msg, True
    End With

    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 10000
        .DelayTime = 300
        Msg = "Next" & Wrap & "Move next record" & Wrap & "Press this key" & Wrap & "or Key" & " PageDown √Ê RightArrow"
        .AddControl btnNext, Msg, True
    End With

    With TTP
        .Create Me.hWnd, Me.Caption, 1, 15204351, -2147483630
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
        If Me.TxtModFlg.text = "R" Then
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

    For ss = SerStartRow To vsGrd.rows - 1
        If chkRowHidden Then
            If Not vsGrd.RowHidden(ss) Then j = j + 1
        Else
            j = j + 1
        End If
        vsGrd.TextMatrix(ss, 0) = j
    Next

End Sub


Private Sub GetComboData(My_Combo As DataCombo, _
                         My_SQL As String)
    Dim rs As ADODB.Recordset
    Dim StrTemp As String
    Dim Msg As String
    On Error GoTo ErrorHandler

    If InStr(1, My_SQL, "SELECT", vbTextCompare) = 0 Then
        Exit Sub
    End If

    My_Combo.Tag = My_SQL
    Set rs = New ADODB.Recordset

    If SystemOptions.SysDataBaseType = SQLServerDataBase Then
        rs.CursorLocation = adUseClient
    End If

    rs.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    'Populate the ADO datacombo by setting its properties
    With My_Combo
        StrTemp = .BoundText
        Set .RowSource = rs
        .BoundColumn = rs(0).Name
        .ListField = rs(1).Name

        If Trim(StrTemp) <> "" Then
            .BoundText = StrTemp
        Else
            .BoundText = ""
            .text = ""
        End If

    End With

Exit_Sub:
    Set rs = Nothing
    Exit Sub
ErrorHandler:
    Msg = Now
    Msg = Msg & CHR(13) & "ClsDataCombos:GetComboData"
    Msg = Msg & CHR(13) & Err.Number
    Msg = Msg & CHR(13) & Err.Description
    WriteInLogFile Msg
    'MsgBox "ERROR! Err# " & Err.Number & " Desc: " & Err.Description, vbCritical + vbOKOnly
    Resume Exit_Sub
End Sub

