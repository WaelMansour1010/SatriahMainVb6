VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SUITEC~1.OCX"
Begin VB.Form nationality 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   9870
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14880
   Icon            =   "nationality.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   9870
   ScaleWidth      =   14880
   ShowInTaskbar   =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin C1SizerLibCtl.C1Tab TabMain 
      Height          =   9735
      Left            =   -210
      TabIndex        =   0
      Top             =   0
      Width           =   15030
      _cx             =   26511
      _cy             =   17171
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
      Caption         =   "»Ì«‰«  «·œÊ· Ê «·Ã‰”Ì« | ‰»ÌÂ«  «·ﬁ÷«Ì«| ’›ÌÂ «·⁄Âœ ··”«∆ﬁÌ‰"
      Align           =   0
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
         Height          =   9360
         Index           =   1
         Left            =   -15885
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   45
         Width           =   14940
         _cx             =   26353
         _cy             =   16510
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
            Height          =   675
            Left            =   15
            RightToLeft     =   -1  'True
            TabIndex        =   18
            Top             =   4215
            Width           =   11220
            Begin VB.ComboBox CmbType 
               BackColor       =   &H80000018&
               Height          =   315
               ItemData        =   "nationality.frx":058A
               Left            =   2280
               List            =   "nationality.frx":059A
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   23
               Top             =   870
               Visible         =   0   'False
               Width           =   1005
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
               Left            =   8160
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   22
               Top             =   285
               Width           =   1065
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
               Left            =   5235
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   21
               Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·Ã‰”Ì…"
               Top             =   285
               Width           =   2880
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
               Height          =   315
               Left            =   2280
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   20
               Tag             =   "Enterr English nAme"
               Top             =   285
               Width           =   2880
            End
            Begin VB.ComboBox DcbQuality 
               Height          =   315
               ItemData        =   "nationality.frx":05B3
               Left            =   0
               List            =   "nationality.frx":05B5
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   19
               Top             =   240
               Width           =   885
            End
            Begin MSDataListLib.DataCombo DCPreFix 
               Height          =   315
               Left            =   1080
               TabIndex        =   24
               Top             =   285
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   556
               _Version        =   393216
               BackColor       =   16777215
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬂÊœ  "
               Height          =   195
               Index           =   3
               Left            =   7905
               RightToLeft     =   -1  'True
               TabIndex        =   29
               Top             =   30
               Width           =   990
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·œÊ·…/ «·Ã‰”Ì…/⁄—»Ì"
               Height          =   285
               Index           =   0
               Left            =   5220
               RightToLeft     =   -1  'True
               TabIndex        =   28
               Top             =   0
               Width           =   1890
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·œÊ·…/ «·Ã‰”Ì… «‰Ã·Ì“Ì"
               Height          =   285
               Index           =   1
               Left            =   2505
               RightToLeft     =   -1  'True
               TabIndex        =   27
               Top             =   0
               Width           =   2130
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬂÊœ «·Ã‰”Ì…"
               Height          =   195
               Index           =   4
               Left            =   1200
               RightToLeft     =   -1  'True
               TabIndex        =   26
               Top             =   0
               Width           =   990
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "›—“ «· ’‰Ì⁄"
               Height          =   195
               Index           =   5
               Left            =   0
               RightToLeft     =   -1  'True
               TabIndex        =   25
               Top             =   0
               Width           =   990
            End
         End
         Begin VB.Frame FraHeader 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   570
            Left            =   15
            RightToLeft     =   -1  'True
            TabIndex        =   7
            Top             =   0
            Width           =   10980
            Begin VB.Frame Frmo2 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   375
               Left            =   540
               RightToLeft     =   -1  'True
               TabIndex        =   10
               Top             =   450
               Visible         =   0   'False
               Width           =   3105
               Begin MSDataListLib.DataCombo DCUser 
                  CausesValidation=   0   'False
                  Height          =   315
                  Left            =   -255
                  TabIndex        =   11
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
                  TabIndex        =   12
                  Top             =   45
                  Width           =   855
               End
            End
            Begin VB.TextBox TxtModFlg 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H0000FF00&
               Enabled         =   0   'False
               Height          =   285
               Left            =   2580
               RightToLeft     =   -1  'True
               TabIndex        =   9
               Text            =   "modflag"
               Top             =   90
               Visible         =   0   'False
               Width           =   465
            End
            Begin VB.TextBox TxtVac_ID 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H000000FF&
               Height          =   240
               Left            =   3030
               RightToLeft     =   -1  'True
               TabIndex        =   8
               Top             =   510
               Visible         =   0   'False
               Width           =   945
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
                     Picture         =   "nationality.frx":05B7
                     Key             =   "CompanyName"
                  EndProperty
                  BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "nationality.frx":0951
                     Key             =   "Ser"
                  EndProperty
                  BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "nationality.frx":0CEB
                     Key             =   "Vac_Name"
                  EndProperty
                  BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "nationality.frx":1085
                     Key             =   "ShareCount"
                  EndProperty
                  BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "nationality.frx":141F
                     Key             =   "Dis_Count"
                  EndProperty
                  BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "nationality.frx":17B9
                     Key             =   "Bouns"
                  EndProperty
                  BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "nationality.frx":1B53
                     Key             =   "SharesValue"
                  EndProperty
                  BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "nationality.frx":20ED
                     Key             =   "BuyValue"
                  EndProperty
               EndProperty
            End
            Begin ImpulseButton.ISButton btnLast 
               Height          =   315
               Left            =   90
               TabIndex        =   13
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
               ButtonImage     =   "nationality.frx":2487
               ColorButton     =   14871017
               AcclimateGrayTones=   -1  'True
               DrawFocusRectangle=   0   'False
               DisabledImageExtraction=   0
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnNext 
               Height          =   315
               Left            =   555
               TabIndex        =   14
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
               ButtonImage     =   "nationality.frx":2821
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnPrevious 
               Height          =   315
               Left            =   1155
               TabIndex        =   15
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
               ButtonImage     =   "nationality.frx":2BBB
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnFirst 
               Height          =   315
               Left            =   1620
               TabIndex        =   16
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
               ButtonImage     =   "nationality.frx":2F55
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "»Ì«‰«  «·œÊ· Ê «·Ã‰”Ì« "
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
               Left            =   6855
               RightToLeft     =   -1  'True
               TabIndex        =   17
               Top             =   90
               Width           =   2280
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "›—“ «· ’‰Ì⁄"
            ForeColor       =   &H000000FF&
            Height          =   765
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   5
            Top             =   4890
            Width           =   6510
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               Caption         =   "«·›—“ Ì»œ√ „‰ 1 «·Ì 10 ÊÌ⁄ »— Ê«Õœ ÂÊ «›÷· ÃÊœ…"
               Height          =   375
               Left            =   360
               RightToLeft     =   -1  'True
               TabIndex        =   6
               Top             =   240
               Width           =   4575
            End
         End
         Begin VSFlex8UCtl.VSFlexGrid FgItems 
            Height          =   9240
            Index           =   0
            Left            =   20640
            TabIndex        =   2
            Top             =   795
            Width           =   14715
            _cx             =   25956
            _cy             =   16298
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
            FormatString    =   $"nationality.frx":32EF
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
         Begin C1SizerLibCtl.C1Elastic EltCont 
            Height          =   1065
            Left            =   1260
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   5760
            Width           =   8040
            _cx             =   14182
            _cy             =   1879
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
               TabIndex        =   31
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
               ButtonImage     =   "nationality.frx":33AF
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnSave 
               Height          =   330
               Left            =   3030
               TabIndex        =   32
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
               ButtonImage     =   "nationality.frx":3749
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnModify 
               Height          =   330
               Left            =   3795
               TabIndex        =   33
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
               ButtonImage     =   "nationality.frx":3AE3
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton BtnUndo 
               Height          =   330
               Left            =   2265
               TabIndex        =   34
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
               ButtonImage     =   "nationality.frx":3E7D
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnDelete 
               Height          =   330
               Left            =   1500
               TabIndex        =   35
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
               ButtonImage     =   "nationality.frx":4217
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnQuery 
               Height          =   330
               Left            =   5880
               TabIndex        =   36
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
               ButtonImage     =   "nationality.frx":47B1
               ColorButton     =   14737632
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton BtnUpdate 
               Height          =   330
               Left            =   6045
               TabIndex        =   37
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
               ButtonImage     =   "nationality.frx":4B4B
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton BtnPrint 
               Height          =   285
               Left            =   4725
               TabIndex        =   38
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
               ButtonImage     =   "nationality.frx":4EE5
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnCancel 
               Height          =   330
               Left            =   705
               TabIndex        =   39
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
               ButtonImage     =   "nationality.frx":527F
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin VB.Label LabCountRec 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Height          =   210
               Left            =   240
               RightToLeft     =   -1  'True
               TabIndex        =   43
               Top             =   225
               Width           =   540
            End
            Begin VB.Label LabCurrRec 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Height          =   210
               Left            =   1800
               RightToLeft     =   -1  'True
               TabIndex        =   42
               Top             =   240
               Width           =   675
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄œœ «·”Ã·« :"
               Height          =   210
               Index           =   1
               Left            =   810
               RightToLeft     =   -1  'True
               TabIndex        =   41
               Top             =   225
               Width           =   975
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·”Ã· «·Õ«·Ì:"
               Height          =   210
               Index           =   0
               Left            =   2505
               RightToLeft     =   -1  'True
               TabIndex        =   40
               Top             =   225
               Width           =   975
            End
         End
         Begin VSFlex8Ctl.VSFlexGrid Grid 
            Height          =   3585
            Left            =   0
            TabIndex        =   44
            Top             =   600
            Width           =   10980
            _cx             =   19368
            _cy             =   6324
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
            FormatString    =   $"nationality.frx":5619
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
         Height          =   9360
         Index           =   0
         Left            =   -15585
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   45
         Width           =   14940
         _cx             =   26353
         _cy             =   16510
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
         Begin VB.CommandButton cmdDisplay 
            Caption         =   "⁄—÷ «· ‰»ÌÂ"
            Height          =   300
            Left            =   210
            RightToLeft     =   -1  'True
            TabIndex        =   46
            Top             =   9015
            Width           =   3750
         End
         Begin VSFlex8UCtl.VSFlexGrid FgItems 
            Height          =   9240
            Index           =   1
            Left            =   20640
            TabIndex        =   4
            Top             =   795
            Width           =   14715
            _cx             =   25956
            _cy             =   16298
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
            FormatString    =   $"nationality.frx":56D5
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
         Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid1 
            Height          =   8505
            Left            =   -30
            TabIndex        =   45
            Top             =   90
            Width           =   10965
            _cx             =   19341
            _cy             =   15002
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
            Cols            =   9
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"nationality.frx":5795
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
         Height          =   9360
         Index           =   2
         Left            =   45
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   45
         Width           =   14940
         _cx             =   26353
         _cy             =   16510
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
         Begin VB.CheckBox Check3 
            Alignment       =   1  'Right Justify
            Caption         =   " ÕœÌœ «·ﬂ·"
            Height          =   195
            Left            =   12900
            RightToLeft     =   -1  'True
            TabIndex        =   139
            Top             =   3570
            Width           =   1695
         End
         Begin VB.CommandButton cmdDeleteEntry 
            Caption         =   "Õ–› «·ﬁÌœ"
            Height          =   375
            Left            =   2100
            RightToLeft     =   -1  'True
            TabIndex        =   138
            Top             =   7680
            Width           =   1095
         End
         Begin VB.CommandButton cmdCreateENtry 
            Caption         =   " «‰‘«¡ «·ﬁÌœ"
            Height          =   375
            Left            =   3240
            RightToLeft     =   -1  'True
            TabIndex        =   137
            Top             =   7650
            Width           =   1095
         End
         Begin VB.Frame Frame10 
            Caption         =   "»Ì«‰«  „Õ«”»Ì…"
            Height          =   825
            Left            =   390
            RightToLeft     =   -1  'True
            TabIndex        =   131
            Top             =   6810
            Width           =   3990
            Begin VB.CommandButton Command9 
               Caption         =   "ÿ»«⁄Â «·ﬁÌœ"
               Height          =   375
               Left            =   150
               RightToLeft     =   -1  'True
               TabIndex        =   133
               Top             =   270
               Width           =   1095
            End
            Begin VB.TextBox TxtNoteSerial 
               Alignment       =   1  'Right Justify
               Height          =   375
               Left            =   1320
               RightToLeft     =   -1  'True
               TabIndex        =   132
               Top             =   240
               Width           =   1935
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "—ﬁ„ «·ﬁÌœ"
               Height          =   195
               Index           =   35
               Left            =   2880
               RightToLeft     =   -1  'True
               TabIndex        =   134
               Top             =   360
               Width           =   990
            End
         End
         Begin VB.TextBox txtNoteSerial1 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   8280
            Locked          =   -1  'True
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   127
            Top             =   570
            Visible         =   0   'False
            Width           =   1230
         End
         Begin VB.TextBox TxtNoteID 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   6720
            RightToLeft     =   -1  'True
            TabIndex        =   126
            Text            =   "Text1"
            Top             =   420
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.TextBox txtTotal2 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   405
            Left            =   5385
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   119
            Top             =   8070
            Width           =   1680
         End
         Begin VB.TextBox txtValuePrice 
            Alignment       =   1  'Right Justify
            Height          =   390
            Left            =   6135
            RightToLeft     =   -1  'True
            TabIndex        =   115
            Top             =   6075
            Width           =   1275
         End
         Begin VB.TextBox txtTripNo 
            Alignment       =   1  'Right Justify
            Height          =   405
            Left            =   7410
            RightToLeft     =   -1  'True
            TabIndex        =   114
            Top             =   6060
            Width           =   1320
         End
         Begin VB.TextBox txtBoardNO 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   12255
            RightToLeft     =   -1  'True
            TabIndex        =   113
            Top             =   6030
            Width           =   1305
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Õ–› ”ÿ—"
            Height          =   390
            Left            =   12435
            RightToLeft     =   -1  'True
            TabIndex        =   111
            Top             =   8115
            Width           =   2040
         End
         Begin VB.TextBox TxtQtyDischarge 
            Alignment       =   1  'Right Justify
            Height          =   390
            Left            =   -2070
            RightToLeft     =   -1  'True
            TabIndex        =   107
            Top             =   5430
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.TextBox TxtQtyDownload 
            Alignment       =   1  'Right Justify
            Height          =   390
            Left            =   -1560
            RightToLeft     =   -1  'True
            TabIndex        =   106
            Top             =   5910
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.TextBox TxtPrice 
            Alignment       =   1  'Right Justify
            Height          =   390
            Left            =   -1725
            RightToLeft     =   -1  'True
            TabIndex        =   104
            Top             =   5850
            Visible         =   0   'False
            Width           =   1425
         End
         Begin VB.TextBox TxtTotalValue 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   390
            Left            =   840
            Locked          =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   102
            Top             =   5910
            Width           =   1680
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   330
            Left            =   12150
            RightToLeft     =   -1  'True
            TabIndex        =   100
            Top             =   1965
            Width           =   750
         End
         Begin VB.CheckBox Check2 
            Alignment       =   1  'Right Justify
            Caption         =   "·”«∆ﬁ „Õœœ"
            Height          =   195
            Left            =   9810
            RightToLeft     =   -1  'True
            TabIndex        =   81
            Top             =   3540
            Width           =   1605
         End
         Begin VB.CheckBox Check1 
            Alignment       =   1  'Right Justify
            Caption         =   "·ﬂ· «·”«∆ﬁÌ‰"
            Height          =   195
            Left            =   11055
            RightToLeft     =   -1  'True
            TabIndex        =   80
            Top             =   3570
            Width           =   1755
         End
         Begin VB.Frame Frame12 
            BackColor       =   &H00E2E9E9&
            Height          =   630
            Index           =   2
            Left            =   7815
            RightToLeft     =   -1  'True
            TabIndex        =   76
            Top             =   2385
            Width           =   6285
            Begin VB.TextBox TxtAccount 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   4140
               RightToLeft     =   -1  'True
               TabIndex        =   77
               Top             =   240
               Width           =   705
            End
            Begin MSDataListLib.DataCombo DcbAccount 
               Height          =   315
               Left            =   90
               TabIndex        =   78
               Top             =   240
               Width           =   4035
               _ExtentX        =   7117
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label lbl 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·Õ”«»"
               Height          =   285
               Index           =   91
               Left            =   5250
               TabIndex        =   79
               Top             =   210
               Width           =   585
            End
         End
         Begin VB.TextBox Text6 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   315
            Left            =   12150
            RightToLeft     =   -1  'True
            TabIndex        =   73
            Top             =   2010
            Visible         =   0   'False
            Width           =   750
         End
         Begin VB.ComboBox CboPaymentType1 
            Height          =   315
            Left            =   11400
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   71
            Top             =   1560
            Width           =   1530
         End
         Begin VB.Frame Frame8 
            Caption         =   "Õœœ «· «—ÌŒ"
            Height          =   630
            Left            =   690
            RightToLeft     =   -1  'True
            TabIndex        =   66
            Top             =   2385
            Width           =   6975
            Begin MSComCtl2.DTPicker Fromdate 
               Height          =   330
               Left            =   4560
               TabIndex        =   67
               Top             =   240
               Width           =   1275
               _ExtentX        =   2249
               _ExtentY        =   582
               _Version        =   393216
               Format          =   49414145
               CurrentDate     =   41640
            End
            Begin MSComCtl2.DTPicker todate 
               Height          =   330
               Left            =   1440
               TabIndex        =   68
               Top             =   240
               Width           =   1275
               _ExtentX        =   2249
               _ExtentY        =   582
               _Version        =   393216
               Format          =   49414145
               CurrentDate     =   41640
            End
            Begin Dynamic_Byte.NourHijriCal Fromdate√H 
               Height          =   330
               Left            =   3210
               TabIndex        =   135
               Top             =   240
               Width           =   1275
               _ExtentX        =   2249
               _ExtentY        =   582
            End
            Begin Dynamic_Byte.NourHijriCal todateH 
               Height          =   330
               Left            =   90
               TabIndex        =   136
               Top             =   240
               Width           =   1275
               _ExtentX        =   2249
               _ExtentY        =   582
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "≈«·Ï"
               Height          =   435
               Index           =   14
               Left            =   2580
               RightToLeft     =   -1  'True
               TabIndex        =   70
               Top             =   240
               Width           =   540
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               BackStyle       =   0  'Transparent
               Caption         =   "«·› —… „‰"
               Height          =   315
               Index           =   1
               Left            =   5580
               RightToLeft     =   -1  'True
               TabIndex        =   69
               Top             =   240
               Width           =   945
            End
         End
         Begin VB.TextBox TxtSerial1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Enabled         =   0   'False
            Height          =   405
            Index           =   2
            Left            =   11775
            RightToLeft     =   -1  'True
            TabIndex        =   65
            Top             =   690
            Width           =   1155
         End
         Begin VB.Frame Fra_Header 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   675
            Index           =   4
            Left            =   -30
            RightToLeft     =   -1  'True
            TabIndex        =   48
            Top             =   0
            Width           =   19905
            Begin VB.Frame Frame2 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   375
               Index           =   4
               Left            =   540
               RightToLeft     =   -1  'True
               TabIndex        =   51
               Top             =   450
               Visible         =   0   'False
               Width           =   3105
               Begin MSDataListLib.DataCombo DataCombo4 
                  CausesValidation=   0   'False
                  Height          =   315
                  Index           =   4
                  Left            =   -255
                  TabIndex        =   52
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
                  Index           =   19
                  Left            =   2160
                  RightToLeft     =   -1  'True
                  TabIndex        =   53
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
               Index           =   2
               Left            =   2580
               RightToLeft     =   -1  'True
               TabIndex        =   50
               Text            =   "modflag"
               Top             =   90
               Visible         =   0   'False
               Width           =   465
            End
            Begin VB.TextBox Text5 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H000000FF&
               Height          =   240
               Left            =   3990
               RightToLeft     =   -1  'True
               TabIndex        =   49
               Top             =   390
               Visible         =   0   'False
               Width           =   945
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
                     Picture         =   "nationality.frx":58E7
                     Key             =   "CompanyName"
                  EndProperty
                  BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "nationality.frx":5C81
                     Key             =   "Ser"
                  EndProperty
                  BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "nationality.frx":601B
                     Key             =   "Vac_Name"
                  EndProperty
                  BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "nationality.frx":63B5
                     Key             =   "ShareCount"
                  EndProperty
                  BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "nationality.frx":674F
                     Key             =   "Dis_Count"
                  EndProperty
                  BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "nationality.frx":6AE9
                     Key             =   "Bouns"
                  EndProperty
                  BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "nationality.frx":6E83
                     Key             =   "SharesValue"
                  EndProperty
                  BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "nationality.frx":741D
                     Key             =   "BuyValue"
                  EndProperty
               EndProperty
            End
            Begin ImpulseButton.ISButton btn_Last 
               Height          =   315
               Index           =   5
               Left            =   90
               TabIndex        =   54
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
               ButtonImage     =   "nationality.frx":77B7
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
               TabIndex        =   55
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
               ButtonImage     =   "nationality.frx":7B51
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btn_Previous 
               Height          =   315
               Index           =   5
               Left            =   1155
               TabIndex        =   56
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
               ButtonImage     =   "nationality.frx":7EEB
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btn_First 
               Height          =   315
               Index           =   5
               Left            =   1620
               TabIndex        =   57
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
               ButtonImage     =   "nationality.frx":8285
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin C1SizerLibCtl.C1Elastic C1Elastic1 
               Height          =   495
               Left            =   0
               TabIndex        =   128
               TabStop         =   0   'False
               Top             =   0
               Visible         =   0   'False
               Width           =   3570
               _cx             =   6297
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
               Begin XtremeSuiteControls.RadioButton RdAuto_Manual 
                  Height          =   255
                  Index           =   0
                  Left            =   2040
                  TabIndex        =   129
                  Top             =   120
                  Width           =   1215
                  _Version        =   786432
                  _ExtentX        =   2143
                  _ExtentY        =   450
                  _StockProps     =   79
                  Caption         =   "ÌœÊÌ"
                  BackColor       =   14871017
                  UseVisualStyle  =   -1  'True
                  TextAlignment   =   1
                  RightToLeft     =   -1  'True
               End
               Begin XtremeSuiteControls.RadioButton RdAuto_Manual 
                  Height          =   255
                  Index           =   1
                  Left            =   240
                  TabIndex        =   130
                  Top             =   120
                  Width           =   1215
                  _Version        =   786432
                  _ExtentX        =   2143
                  _ExtentY        =   450
                  _StockProps     =   79
                  Caption         =   "«·Ì"
                  BackColor       =   14871017
                  UseVisualStyle  =   -1  'True
                  TextAlignment   =   1
                  RightToLeft     =   -1  'True
               End
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   " ’›ÌÂ «·⁄Âœ ··”«∆ﬁÌ‰"
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
               Left            =   11340
               RightToLeft     =   -1  'True
               TabIndex        =   58
               Top             =   90
               Width           =   2640
            End
         End
         Begin MSComCtl2.DTPicker XPDtbTrans 
            Height          =   330
            Index           =   2
            Left            =   11010
            TabIndex        =   60
            Top             =   1155
            Width           =   1965
            _ExtentX        =   3466
            _ExtentY        =   582
            _Version        =   393216
            Format          =   49414145
            CurrentDate     =   38784
         End
         Begin MSDataListLib.DataCombo DcBranch 
            Height          =   315
            Left            =   630
            TabIndex        =   63
            Top             =   1185
            Width           =   4080
            _ExtentX        =   7197
            _ExtentY        =   556
            _Version        =   393216
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
         Begin MSDataListLib.DataCombo DcboBankName 
            Height          =   315
            Left            =   8400
            TabIndex        =   74
            Top             =   2280
            Visible         =   0   'False
            Width           =   3810
            _ExtentX        =   6720
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VSFlex8UCtl.VSFlexGrid GridInstallments 
            Height          =   1830
            Left            =   930
            TabIndex        =   82
            Top             =   3840
            Width           =   13770
            _cx             =   24289
            _cy             =   3228
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
            Cols            =   38
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"nationality.frx":861F
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
            ExplorerBar     =   3
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
         Begin VSFlex8Ctl.VSFlexGrid fg 
            Height          =   1635
            Left            =   4395
            TabIndex        =   83
            Top             =   6510
            Width           =   10230
            _cx             =   18045
            _cy             =   2884
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
            Cols            =   11
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"nationality.frx":8BA0
            ScrollTrack     =   0   'False
            ScrollBars      =   2
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
         Begin ImpulseButton.ISButton btn_New 
            Height          =   375
            Index           =   2
            Left            =   7680
            TabIndex        =   84
            Top             =   8880
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   661
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
            ButtonImage     =   "nationality.frx":8D38
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Save 
            Height          =   375
            Index           =   2
            Left            =   5820
            TabIndex        =   85
            Top             =   8880
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   661
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
            ButtonImage     =   "nationality.frx":90D2
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Modify 
            Height          =   375
            Index           =   2
            Left            =   6720
            TabIndex        =   86
            Top             =   8880
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   661
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
            ButtonImage     =   "nationality.frx":946C
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton Btn_Undo 
            Height          =   375
            Index           =   2
            Left            =   4830
            TabIndex        =   87
            Top             =   8880
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   661
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
            ButtonImage     =   "nationality.frx":9806
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Delete 
            Height          =   375
            Index           =   2
            Left            =   3930
            TabIndex        =   88
            Top             =   8880
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   661
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
            ButtonImage     =   "nationality.frx":9BA0
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton Btn_Update 
            Height          =   420
            Index           =   2
            Left            =   630
            TabIndex        =   89
            TabStop         =   0   'False
            ToolTipText     =   " ÕœÌÀ ﬁ«⁄œ… «·»Ì«‰« "
            Top             =   8385
            Visible         =   0   'False
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   741
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
            ButtonImage     =   "nationality.frx":A13A
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Cancel 
            Height          =   375
            Index           =   2
            Left            =   690
            TabIndex        =   90
            Top             =   8880
            Width           =   960
            _ExtentX        =   1693
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
            ButtonImage     =   "nationality.frx":A4D4
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton Btn_Print 
            Height          =   585
            Index           =   2
            Left            =   2835
            TabIndex        =   91
            TabStop         =   0   'False
            ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
            Top             =   8775
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   1032
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
            ButtonImage     =   "nationality.frx":A86E
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btn_Query 
            Height          =   645
            Index           =   2
            Left            =   1455
            TabIndex        =   92
            TabStop         =   0   'False
            ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
            Top             =   8730
            Width           =   1290
            _ExtentX        =   2275
            _ExtentY        =   1138
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
            ButtonImage     =   "nationality.frx":110D0
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton cmdAdd 
            Height          =   285
            Left            =   660
            TabIndex        =   97
            Top             =   3540
            Width           =   1020
            _ExtentX        =   1799
            _ExtentY        =   503
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "≈÷«›…"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "nationality.frx":1146A
            DrawFocusRectangle=   0   'False
         End
         Begin MSDataListLib.DataCombo DcboDebitSide 
            Height          =   315
            Left            =   405
            TabIndex        =   98
            Top             =   1725
            Visible         =   0   'False
            Width           =   4260
            _ExtentX        =   7514
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcboCreditSide 
            Height          =   315
            Left            =   405
            TabIndex        =   99
            Top             =   1980
            Visible         =   0   'False
            Width           =   4260
            _ExtentX        =   7514
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo DcboBox 
            Height          =   315
            Left            =   8325
            TabIndex        =   101
            Top             =   1980
            Width           =   3780
            _ExtentX        =   6668
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton RdQty 
            Height          =   270
            Index           =   0
            Left            =   -1020
            TabIndex        =   108
            Top             =   5820
            Visible         =   0   'False
            Width           =   1290
            _Version        =   786432
            _ExtentX        =   2275
            _ExtentY        =   476
            _StockProps     =   79
            Caption         =   "ﬂ„Ì… «· Õ„Ì·"
            BackColor       =   14871017
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton RdQty 
            Height          =   270
            Index           =   1
            Left            =   -1140
            TabIndex        =   109
            Top             =   6405
            Visible         =   0   'False
            Width           =   1290
            _Version        =   786432
            _ExtentX        =   2275
            _ExtentY        =   476
            _StockProps     =   79
            Caption         =   "ﬂ„Ì… «· ›—Ì€"
            BackColor       =   14871017
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin ImpulseButton.ISButton ISButton1 
            Height          =   285
            Left            =   5250
            TabIndex        =   110
            Top             =   6165
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   503
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "≈÷«›…"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ButtonImage     =   "nationality.frx":11804
            DrawFocusRectangle=   0   'False
         End
         Begin MSDataListLib.DataCombo DcbLeaderID 
            Height          =   315
            Left            =   4935
            TabIndex        =   112
            Top             =   3510
            Width           =   4815
            _ExtentX        =   8493
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
         Begin MSDataListLib.DataCombo DcbEqup 
            Height          =   315
            Left            =   10260
            TabIndex        =   121
            Top             =   6030
            Width           =   1965
            _ExtentX        =   3466
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSComCtl2.DTPicker txtTripDate 
            Height          =   330
            Left            =   8760
            TabIndex        =   122
            Top             =   6030
            Width           =   1485
            _ExtentX        =   2619
            _ExtentY        =   582
            _Version        =   393216
            Format          =   49414145
            CurrentDate     =   38784
         End
         Begin Dynamic_Byte.NourHijriCal XPDtbTransH 
            Height          =   255
            Index           =   2
            Left            =   7470
            TabIndex        =   125
            Top             =   1170
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   450
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„⁄œÂ/«·”Ì«—…"
            Height          =   285
            Index           =   8
            Left            =   10860
            RightToLeft     =   -1  'True
            TabIndex        =   124
            Top             =   5760
            Width           =   735
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «·—Õ·…"
            Height          =   285
            Index           =   7
            Left            =   9000
            RightToLeft     =   -1  'True
            TabIndex        =   123
            Top             =   5790
            Width           =   840
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·«Ã„«·Ì"
            Height          =   330
            Index           =   6
            Left            =   6750
            RightToLeft     =   -1  'True
            TabIndex        =   120
            Top             =   8130
            Width           =   975
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·ﬁÌ„…"
            Height          =   300
            Index           =   5
            Left            =   6330
            RightToLeft     =   -1  'True
            TabIndex        =   118
            Top             =   5790
            Width           =   495
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·—Õ·…"
            Height          =   285
            Index           =   4
            Left            =   7560
            RightToLeft     =   -1  'True
            TabIndex        =   117
            Top             =   5790
            Width           =   840
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·„⁄œÂ/«·”Ì«—…"
            Height          =   285
            Index           =   3
            Left            =   12450
            RightToLeft     =   -1  'True
            TabIndex        =   116
            Top             =   5760
            Width           =   915
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”⁄—"
            Height          =   330
            Index           =   10
            Left            =   -705
            RightToLeft     =   -1  'True
            TabIndex        =   105
            Top             =   5850
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·«Ã„«·Ì"
            Height          =   330
            Index           =   12
            Left            =   2235
            RightToLeft     =   -1  'True
            TabIndex        =   103
            Top             =   5910
            Width           =   975
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”Ã· «·Õ«·Ì:"
            Height          =   225
            Index           =   6
            Left            =   6390
            RightToLeft     =   -1  'True
            TabIndex        =   96
            Top             =   8505
            Width           =   1140
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «·”Ã·« :"
            Height          =   225
            Index           =   7
            Left            =   4305
            RightToLeft     =   -1  'True
            TabIndex        =   95
            Top             =   8505
            Width           =   1185
         End
         Begin VB.Label LabCurr_Rec 
            BackColor       =   &H00E2E9E9&
            Height          =   225
            Index           =   2
            Left            =   5490
            RightToLeft     =   -1  'True
            TabIndex        =   94
            Top             =   8535
            Width           =   810
         End
         Begin VB.Label LabCount_Rec 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   225
            Index           =   2
            Left            =   3540
            RightToLeft     =   -1  'True
            TabIndex        =   93
            Top             =   8535
            Width           =   615
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " «·⁄Âœ… / «·’‰œÊﬁ"
            Height          =   285
            Index           =   17
            Left            =   12975
            RightToLeft     =   -1  'True
            TabIndex        =   75
            Top             =   2040
            Width           =   1290
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ÿ—Ìﬁ… «·œ›⁄"
            Height          =   285
            Index           =   23
            Left            =   12915
            RightToLeft     =   -1  'True
            TabIndex        =   72
            Top             =   1650
            Width           =   1350
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·›—⁄"
            Height          =   225
            Index           =   13
            Left            =   4215
            RightToLeft     =   -1  'True
            TabIndex        =   64
            Top             =   1185
            Width           =   1035
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«· «—ÌŒ "
            Height          =   285
            Index           =   2
            Left            =   12945
            TabIndex        =   62
            Top             =   1185
            Width           =   1830
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ ÂÃ—Ì"
            Height          =   285
            Index           =   0
            Left            =   8835
            TabIndex        =   61
            Top             =   1155
            Width           =   1830
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·”‰œ"
            Height          =   285
            Index           =   16
            Left            =   13320
            TabIndex        =   59
            Top             =   750
            Width           =   1080
         End
      End
   End
End
Attribute VB_Name = "nationality"
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
      
                RsSavRec.find "id=" & val(TxtSerial1(mIndex).Text), , adSearchForward, 1
            
            'CuurentLogdata  ("D")
            RsSavRec.delete
            StrSQL = "Delete From notes Where notes_all=" & val(TxtSerial1(mIndex).Text)
            Cn.Execute StrSQL, , adExecuteNoRecords
            Cn.Execute "delete tblTripTrans2 where TravID=" & val(Me.TxtSerial1(mIndex).Text)
            Cn.Execute "delete tblTripTrans3 where TravID=" & val(Me.TxtSerial1(mIndex).Text)
            
            StrSQL = "Delete From notes Where notes_all=" & val(TxtSerial1(mIndex).Text)
            Cn.Execute StrSQL, , adExecuteNoRecords
            StrSQL = "Delete From notes Where noteId=" & val(TXTNoteID.Text)
            Cn.Execute StrSQL, , adExecuteNoRecords
            
            StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(Me.TXTNoteID.Text)
            Cn.Execute StrSQL, , adExecuteNoRecords
            
            If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox " „  ⁄„·Ì… «·Õ–› »‰Ã«Õ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
            Else
            MsgBox "Delete Successfully", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
            End If
            '------------------------------ Move Next ---------------------------.
            
            If mIndex = 2 Then
                FiLLTXT2
                
                
                
            End If
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

Private Sub btn_Modify_Click(Index As Integer)
    Dim Msg As String
    If mIndex = 2 Then
         If Trim(TxtNoteSerial) <> "" Then
            MsgBox "·« Ì„ﬂ‰ «· ⁄œÌ· ﬁ»· Õ–› «·ﬁÌœ «Ê·«"
            Exit Sub
         End If
    End If
    If DoPremis(Do_Edit, Me.Name, True) = False Then
        Exit Sub
    End If

    On Error GoTo ErrTrap


    If TxtSerial1(mIndex).Text <> "" Then
        TxtModFlg2(mIndex) = "E"
'    Frame1(1).Enabled = True
'    Frame1(2).Enabled = True
        'Frm2.Enabled = True
        
    End If

   ' Exit Sub
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
 Dim currentgroup As Integer
    If DoPremis(Do_New, Me.Name, True) = False Then
        Exit Sub
    End If

    On Error GoTo ErrTrap
    Set rs = New ADODB.Recordset
'    Frame1(1).Enabled = True
'    Frame1(2).Enabled = True
    clear_all Me
    FG.Rows = 1
    TxtModFlg2(mIndex).Text = "N"

    If mIndex = 2 Then
        My_SQL = "tblTripTrans"
        'DCboUserName(mIndex) = user_id
        cmdCreateENtry.Enabled = False
        cmdDeleteEntry.Enabled = False
         clear_all Me
 
        
        dcBranch.BoundText = branch_id
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable

    If rs.RecordCount > 0 Then
        TxtSerial1(mIndex).Text = rs.RecordCount + 1
    Else
        TxtSerial1(mIndex).Text = 1
    End If
    
        
    End If
    
    XPDtbTransH(mIndex).value = ToHijriDate(Date)
    
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

Private Sub Btn_Print_Click(Index As Integer)
    If mIndex = 2 Then
        print_report
    End If
End Sub


Function print_report(Optional NoteSerial As String)
    
    Dim MySQL As String, StrSQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String

  
 StrSQL = " SELECT   tblTripTrans.Id,tblTripTrans.recordDate,tblTripTrans.recordDateH,tblTripTrans.Fromdate,tblTripTrans.todate,"
StrSQL = StrSQL & "                       notesallid , dbo.tblTripTrans2.notesallid, dbo.tblTripTrans2.ID, dbo.tblTripTrans2.TravID, dbo.tblTripTrans2.TripNo, dbo.tblTripTrans2.TripDate, dbo.tblTripTrans2.BranchID, "
StrSQL = StrSQL & "                      TblBoxesData.BoxName , Accounts.account_name,Branches2.branch_name,"
StrSQL = StrSQL & "                      dbo.tblTripTrans2.Price,dbo.tblTripTrans2.TotalValue,tblTripTrans2.RecNo,tblTripTrans2.Weight,"
StrSQL = StrSQL & "                      dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.tblTripTrans2.Typed, dbo.tblTripTrans2.[Value], dbo.tblTripTrans2.Remarks,"
StrSQL = StrSQL & "                      dbo.tblTripTrans2.NoteID, dbo.tblTripTrans2.QtyDownload, dbo.tblTripTrans2.QtyDischarge, dbo.tblTripTrans2.CardNO, dbo.tblTripTrans2.CardNO2,"
StrSQL = StrSQL & "                      dbo.tblTripTrans2.CarType1, dbo.tblTripTrans2.CarID, dbo.TblCarsData.BoardNO, dbo.TblVendorCars.BoardNo AS BoardNo2, dbo.tblTripTrans2.FromID,"
StrSQL = StrSQL & "                      TblCountriesGovernments_2.GovernmentName, dbo.tblTripTrans2.ToID, TblCountriesGovernments_1.GovernmentName AS ToGovernmentName,"
StrSQL = StrSQL & "                      dbo.tblTripTrans2.CarTypeID, dbo.TBLCarTypes.name, dbo.TBLCarTypes.namee, dbo.tblTripTrans2.TypeTrans, dbo.tblTripTrans2.ShipID,"
StrSQL = StrSQL & "                      dbo.TblShipsData.Name AS ShipName, dbo.TblShipsData.NameE AS ShipNameE, dbo.tblTripTrans2.LeaderName,TblCustemers.CusName ,TblCustemers.CusID ,TblCarsData.BoardNO"
StrSQL = StrSQL & " FROM         dbo.tblTripTrans2 LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblShipsData ON dbo.tblTripTrans2.ShipID = dbo.TblShipsData.id LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TBLCarTypes ON dbo.tblTripTrans2.CarTypeID = dbo.TBLCarTypes.id LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblCountriesGovernments TblCountriesGovernments_1 ON dbo.tblTripTrans2.ToID = TblCountriesGovernments_1.GovernmentID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblCountriesGovernments TblCountriesGovernments_2 ON dbo.tblTripTrans2.FromID = TblCountriesGovernments_2.GovernmentID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblVendorCars ON dbo.tblTripTrans2.CarID = dbo.TblVendorCars.ID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblCarsData ON dbo.tblTripTrans2.CarID = dbo.TblCarsData.id LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblBranchesData ON dbo.tblTripTrans2.BranchID = dbo.TblBranchesData.branch_id"
StrSQL = StrSQL & "                      LEFT OUTER JOIN TblCustemers On TblCustemers.CusID = dbo.tblTripTrans2.CusID "
StrSQL = StrSQL & "                      LEFT OUTER JOIN tblTripTrans On tblTripTrans.Id = tblTripTrans2.TravID "
StrSQL = StrSQL & "                      LEFT OUTER JOIN ACCOUNTS On tblTripTrans.AccountPaym = ACCOUNTS.Account_Code "
StrSQL = StrSQL & "                      LEFT OUTER JOIN TblBranchesData Branches2 On tblTripTrans.BranchId = Branches2.branch_id "
StrSQL = StrSQL & "                      LEFT OUTER JOIN TblBoxesData  On tblTripTrans.BoxID = TblBoxesData.BoxID "

StrSQL = StrSQL & "   Where (dbo.tblTripTrans2.TravID = " & val(TxtSerial1(mIndex).Text) & ") and (dbo.tblTripTrans2.TypeTrans is null or dbo.tblTripTrans2.TypeTrans=0)  "


    If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\Reports\" & "Transporter\TripDataR.rpt"
    Else
        StrFileName = App.path & "\Reports\" & "Transporter\TripDataR.rpt"
    End If

    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

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
    
    
    StrSQL = "sELECT *,TblCarsData.EqupName as CarName FROM tblTripTrans3 lEFT oUTER join TblCarsData On TblCarsData.fixedassetid =tblTripTrans3.CarID  where TravID=" & val(Me.TxtSerial1(mIndex).Text)


 If StrSQL <> "" Then
        Dim RsData2  As New ADODB.Recordset
        
         
        RsData2.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        xReport.OpenSubreport("Sub2").Database.SetDataSource RsData2
        
        
    End If
    
    Dim cCompanyInfo As New ClsCompanyInfo

    If SystemOptions.UserInterface = ArabicInterface Then
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
        'xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
        StrReportTitle = "" '& StrAccountName
        'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        'StrReportTitle = StrReportTitle + " »œ«Ì… „‰ " & Format(Me.DTPickerAccFrom.value, "yyyy/M/d") & ""
        'End If
        'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        'StrReportTitle = StrReportTitle + " ≈·Ï " & Format(Me.DTPickerAccTo.value, "yyyy/M/d") & " "
        'End If
    Else
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        'xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""
        'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        'StrReportTitle = StrReportTitle + " From Date " & (Me.DTPickerAccFrom.value) & ""
        'End If
        'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        'StrReportTitle = StrReportTitle + " To Date :  " & (Me.DTPickerAccTo.value) & ""
        'End If
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
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

Private Sub btn_Query_Click(Index As Integer)
 Load FrmSearchCarsPlan
 
   FrmSearchCarsPlan.Caption = "»ÕÀ ⁄‰ „ «»⁄… «·—Õ·« "
    FrmSearchCarsPlan.Indx = 3
    FrmSearchCarsPlan.show vbModal
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


If mIndex = 0 Then
'     If Dcbranch(mIndex).Text = "" Then
'        MsgBox "Please Enter Branch"
'        Dcbranch(mIndex).SetFocus
'        Exit Sub
'    End If
End If
    
    '------------------------------ check if Empcode exist ----------------------

   

    ' -------------------------------------- txtmodflg type -------------------
    Select Case TxtModFlg2(mIndex).Text

            '------------------------------ new record ----------------------------
        Case "N"
      
            '------------------------- save record -----------------------------
            
            If mIndex = 2 Then
                'AddNewRec
                FiLLRec2
                
                
                
            End If
            

        Case "E"

            '----------------------------- save edit -------------------------------
            
            If mIndex = 2 Then
                FiLLRec2
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
Private Sub Undo()
    On Error GoTo ErrTrap
    
    
    Select Case TxtModFlg2(mIndex).Text

        Case "N"
            clear_all Me
            TxtModFlg2(mIndex).Text = "R"
           
            btn_First_Click (mIndex)
        Case "E"
                RsSavRec.find "ID='" & val(TxtSerial1(mIndex).Text) & "'", , adSearchForward, adBookmarkFirst


            If RsSavRec.EOF Or RsSavRec.BOF Then
                TxtModFlg2(mIndex).Text = "R"
                Exit Sub
            End If

            If mIndex = 2 Then
                FiLLTXT2

            End If
            TxtModFlg2(mIndex).Text = "R"
    End Select
    
    Exit Sub
ErrTrap:
End Sub

Private Sub Btn_Undo_Click(Index As Integer)
    Undo
End Sub

Private Sub CboPaymentType1_Change()
 If Me.CboPaymentType1.ListIndex = 0 Then
        
    Me.DcboBox.Enabled = True
    Text1.Enabled = True
    DcbAccount.Text = ""
    TxtAccount.Text = ""
    DcbAccount.Enabled = False
    TxtAccount.Enabled = False
Else
    Me.DcboBox.Enabled = False
    Text1.Enabled = False
    Text1.Text = ""
    TxtAccount.Text = ""
    DcbAccount.Enabled = True
    TxtAccount.Enabled = True
End If
End Sub

Private Sub CboPaymentType1_Click()
 If Me.CboPaymentType1.ListIndex = 0 Then
        
    Me.DcboBox.Enabled = True
    Text1.Enabled = True
    DcbAccount.Text = ""
    TxtAccount.Text = ""
    DcbAccount.Enabled = False
    TxtAccount.Enabled = False
Else
    Me.DcboBox.Enabled = False
    Text1.Enabled = False
    Text1.Text = ""
    TxtAccount.Text = ""
    DcbAccount.Enabled = True
    TxtAccount.Enabled = True
End If
End Sub

Private Sub Check1_Click()
If Check1.value = vbChecked Then
    DcbLeaderID.Enabled = False
    Check2.value = vbUnchecked
End If

End Sub

Private Sub Check2_Click()
If Check2.value = vbUnchecked Then
    DcbLeaderID.Enabled = False
Else
    Check1.value = vbUnchecked
    DcbLeaderID.Enabled = True
End If
End Sub

Private Sub Check3_Click()
    Dim i As Long
    For i = 1 To GridInstallments.Rows - 1
        GridInstallments.TextMatrix(i, GridInstallments.ColIndex("Select")) = IIf(Check3, 1, 0)
    Next
End Sub

Private Sub cmdAdd_Click()

If Me.TxtModFlg2(mIndex) <> "E" And Me.TxtModFlg2(mIndex) <> "N" Then Exit Sub
FillGrid
End Sub

Private Sub cmdDeleteEntry_Click()
Dim StrSQL As String
    StrSQL = "Delete From notes Where notes_all=" & val(TxtSerial1(mIndex).Text)
    Cn.Execute StrSQL, , adExecuteNoRecords
    StrSQL = "Delete From notes Where noteId=" & val(TXTNoteID.Text)
    Cn.Execute StrSQL, , adExecuteNoRecords
    
    StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(Me.TXTNoteID.Text)
    Cn.Execute StrSQL, , adExecuteNoRecords
    RsSavRec.Fields("NoteID").value = Null
    RsSavRec("NoteSerial").value = Null
    RsSavRec.update

    FiLLTXT2
End Sub

Private Sub Command1_Click()
If FG.Rows > 1 Then
    FG.Rows = FG.Rows - 1
End If
End Sub

Private Sub cmdCreateENtry_Click()
    If Trim(TxtNoteSerial) = "" And TxtSerial1(mIndex) <> "" Then
        createVoucher
        FiLLTXT2
    End If
End Sub

Private Sub Command9_Click()
   ShowGL_cc Me.TxtNoteSerial.Text, , 9095
End Sub

Private Sub DcbAccount_Change()
DcbAccount_Click (0)
End Sub
Private Sub DcbAccount_KeyUp(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyF3 Then

        Unload Account_search
        Account_search.show
        Account_search.case_id = 31219
        

        
            
    End If
    
    
End Sub
Private Sub DcbAccount_Click(Area As Integer)
TxtAccount.Text = getAccountSerial_Code("Account_Serial", "Account_Code", DcbAccount.BoundText)
'If Me.TxtModFlg.Text <> "R" Then
        'If CboPaymentType.ListIndex = 4 Then
            Me.DcboCreditSide.BoundText = DcbAccount.BoundText
        'End If
' End If
End Sub

Private Sub DcbEqup_Change()
    If DcbEqup.Text = "" Then
        txtBoardNo = ""
        Dim Dcombos As New ClsDataCombos
        Dcombos.GetEquipments DcbEqup
    End If
End Sub

Private Sub DcbEqup_Click(Area As Integer)
Dim s As String

Dim rsDummy As New ADODB.Recordset
    's = "Select BoardNO FROM TblCarsData Where Id = " & val(DcbEqup.BoundText)
    Dim str As String
    s = " SELECT       TblCarsData.BoardNO                 FROM         dbo.TblCarsData LEFT OUTER JOIN                       dbo.insurance_companies ON dbo.TblCarsData.InsuranceCompanyId = dbo.insurance_companies.id LEFT OUTER JOIN                       dbo.TblEmployee ON dbo.TblCarsData.Emp_id = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN                       dbo.TBLCarTypes ON dbo.TblCarsData.CarsTypeId = dbo.TBLCarTypes.id LEFT OUTER JOIN                       dbo.FixedAssets ON dbo.TblCarsData.fixedAssetid = dbo.FixedAssets.id LEFT OUTER JOIN                       dbo.TblBranchesData ON dbo.TblCarsData.Branch_NO = dbo.TblBranchesData.branch_id  where  (dbo.TblCarsData.branch_no =0 or dbo.TblCarsData.branch_no is null or    dbo.TblCarsData.branch_no  in( SELECT     BranchID From dbo.TblUsersBranches  Where (UserID = " & user_id & " ))) AND  fixedassetid = " & val(DcbEqup.BoundText)
    rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly
    If Not rsDummy.EOF Then
        txtTripDate.Tag = rsDummy!BoardNO & ""
    Else
        txtTripDate.Tag = ""
    End If
    If DcbEqup.Text = "" Then
        txtTripDate.Tag = ""
        txtBoardNo = ""
        Dim Dcombos As New ClsDataCombos
        Dcombos.GetEquipments DcbEqup
    End If
End Sub

Private Sub DcboBox_Change()
Dim acc As String

    
    If DcboBox.BoundText = "" Then Exit Sub
    If Me.TxtModFlg2(mIndex).Text = "N" Or Me.TxtModFlg2(mIndex).Text = "E" Then
        DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblBoxesData", "BoxID", val(Me.DcboBox.BoundText))
    
    acc = ModAccounts.GetMyAccountCode("TblBoxesData", "BoxID", val(Me.DcboBox.BoundText))
    'WriteCustomerBalPublic acc, Balance, balanceString
    'LblLink1.Caption = balanceString
    End If

    
End Sub
Private Sub DcboBox_Click(Area As Integer)
    DcboBox_Change
End Sub

Private Sub Fg_AfterEdit(ByVal Row As Long, ByVal Col As Long)
ReLineGrid2
End Sub

Private Sub fg_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'If Me.TxtModFlg2(mIndex) <> "E" Then Cancel = True Else Cancel = False
If Me.TxtModFlg2(mIndex) <> "E" And Me.TxtModFlg2(mIndex) <> "N" Then Cancel = True: Exit Sub Else Cancel = False

        Select Case FG.ColKey(Col)
         Case "Price"
            FG.EditMaxLength = 10
        Case "Remarks"
                FG.EditMaxLength = 200
        Case "TripDate"
                FG.EditMaxLength = 10
        Case Else
            Cancel = True
        End Select
        
End Sub

Private Sub GridInstallments_AfterEdit(ByVal Row As Long, ByVal Col As Long)

If Me.TxtModFlg2(mIndex) <> "E" And Me.TxtModFlg2(mIndex) <> "N" Then Exit Sub
ReLineGrid
End Sub

Private Sub GridInstallments_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

        If Me.TxtModFlg2(mIndex) <> "E" And Me.TxtModFlg2(mIndex) <> "N" Then Cancel = True: Exit Sub Else Cancel = False
        Select Case GridInstallments.ColKey(Col)
         Case "Select"
            GridInstallments.EditMaxLength = 10
        Case Else
            Cancel = True
        End Select
End Sub

Private Sub ISButton1_Click()

If Me.TxtModFlg2(mIndex) <> "E" And Me.TxtModFlg2(mIndex) <> "N" Then Exit Sub
FG.Rows = FG.Rows + 1
FG.TextMatrix(FG.Rows - 1, FG.ColIndex("TripNo")) = txtTripNo
'GridInstallments.TextMatrix(GridInstallments.Row, GridInstallments.ColIndex("TripNo"))
FG.TextMatrix(FG.Rows - 1, FG.ColIndex("BoardNO")) = txtTripDate.Tag
FG.TextMatrix(FG.Rows - 1, FG.ColIndex("Price")) = txtValuePrice

FG.TextMatrix(FG.Rows - 1, FG.ColIndex("CarName")) = DcbEqup.Text
FG.TextMatrix(FG.Rows - 1, FG.ColIndex("CarID")) = DcbEqup.BoundText

Dim s As String
Dim rsDummy As New ADODB.Recordset
s = "Select ID CarID From TblCarsData Where fixedAssetid = " & val(DcbEqup.BoundText)
rsDummy.Open s, Cn, adOpenKeyset, adLockReadOnly
If Not rsDummy.EOF Then
    'fg.TextMatrix(fg.Rows - 1, fg.ColIndex("fixedAssetid")) = val(rsDummy!FixedassetId & "")
    FG.TextMatrix(FG.Rows - 1, FG.ColIndex("fixedAssetid")) = val(DcbEqup.BoundText)
    FG.TextMatrix(FG.Rows - 1, FG.ColIndex("CarID")) = val(rsDummy!CarID & "")
    
End If
FG.TextMatrix(FG.Rows - 1, FG.ColIndex("TripDate")) = txtTripDate.value
'fg.TextMatrix(fg.Rows - 1, fg.ColIndex("Price")) = txtValuePrice

'GridInstallments.TextMatrix(GridInstallments.Row, GridInstallments.ColIndex("Car"))
ReLineGrid2
End Sub

Private Sub TxtAccount_KeyPress(KeyAscii As Integer)
DcbAccount.BoundText = getAccountSerial_Code("Account_Code", "Account_Serial", TxtAccount.Text)
End Sub



Private Sub DcboBankName_Change()
    On Error Resume Next

    If DcboBankName.BoundText = "" Then Exit Sub
    Dim RsSavRec As ADODB.Recordset
    Dim My_SQL As String

    If Me.TxtModFlg.Text = "N" Or Me.TxtModFlg.Text = "E" Then
        '    Me.DcboCreditSide.BoundText = "a2a3a2"
    
        My_SQL = "  select Account_Code from BanksData WHERE BankID=" & DcboBankName.BoundText

        Set RsSavRec = New ADODB.Recordset
        RsSavRec.CursorLocation = adUseClient
        RsSavRec.Open My_SQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
 
        If SystemOptions.banks_Accounts3 = True Then
            Me.DcboCreditSide.BoundText = get_bank_Account(val(Me.DcboBankName.BoundText), "Account_Code2")
        Else
            Me.DcboCreditSide.BoundText = RsSavRec.Fields("Account_Code").value
        End If
    
        'Me.DcboCreditSide.BoundText = RsSavRec.Fields("Account_Code").value

      

    End If

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
    If mIndex = 2 Then
        FiLLTXT2
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

   
    If mIndex = 2 Then
        FiLLTXT2

        
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

   
    If mIndex = 2 Then
        FiLLTXT2

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
    If mIndex = 2 Then
        FiLLTXT2
    
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

Private Sub TxtBoardNO_Validate(Cancel As Boolean)
On Error Resume Next
   Dim Dcombos As New ClsDataCombos
    Dim str As String
    Dim rsDummy As New ADODB.Recordset
    Dim EmpID As Integer
  
    
    str = " SELECT       fixedassetid                 FROM         dbo.TblCarsData LEFT OUTER JOIN                       dbo.insurance_companies ON dbo.TblCarsData.InsuranceCompanyId = dbo.insurance_companies.id LEFT OUTER JOIN                       dbo.TblEmployee ON dbo.TblCarsData.Emp_id = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN                       dbo.TBLCarTypes ON dbo.TblCarsData.CarsTypeId = dbo.TBLCarTypes.id LEFT OUTER JOIN                       dbo.FixedAssets ON dbo.TblCarsData.fixedAssetid = dbo.FixedAssets.id LEFT OUTER JOIN                       dbo.TblBranchesData ON dbo.TblCarsData.Branch_NO = dbo.TblBranchesData.branch_id  where  (dbo.TblCarsData.branch_no =0 or dbo.TblCarsData.branch_no is null or    dbo.TblCarsData.branch_no  in( SELECT     BranchID From dbo.TblUsersBranches  Where (UserID = " & user_id & " ))) AND  (dbo.TblCarsData.Fullcode like '%" & txtBoardNo.Text & "%' Or dbo.TblCarsData.EqupName like '%" & txtBoardNo.Text & "%' )"

   
   rsDummy.Open str, Cn, adOpenStatic, adLockReadOnly
   Dcombos.GetEquipments DcbEqup, str
   If Not rsDummy.EOF Then
    DcbEqup.BoundText = val(rsDummy!FixedassetId)
   End If
   If txtBoardNo = "" Then
        Dcombos.GetEquipments DcbEqup
   End If

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
    MSGType = MsgBox("Â·  —€» ›Ì Õ–› Â–« «·”Ã·", vbQuestion + vbMsgBoxRtlReading + vbYesNo + vbMsgBoxRight, App.title)

    If MSGType = vbYes Then
        RsSavRec.find "id=" & val(TxtVac_ID.Text), , adSearchForward, 1
        RsSavRec.delete
        MsgBox " „  ⁄„·Ì… «·Õ–› »‰Ã«Õ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
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
            MsgBox StrMSG, vbOKOnly + vbMsgBoxRight, App.title
            'clear the ConnectiOn Errors
            Cn.Errors.Clear
    End Select

End Sub

Private Sub BtnFirst_Click()
    On Error GoTo ErrTrap

    Dim Msg As String

    If Me.TxtModFlg.Text = "N" Then
        FindRec val(TxtVac_ID.Text)
        Me.TxtModFlg.Text = "R"
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
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select

End Sub

Private Sub BtnLast_Click()
    On Error GoTo ErrTrap

    Dim Msg As String

    If Me.TxtModFlg.Text = "N" Then
        FindRec val(TxtVac_ID.Text)
        Me.TxtModFlg.Text = "R"
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

    My_SQL = "Nationality"
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

    If Me.TxtModFlg.Text = "N" Then
        FindRec val(TxtVac_ID.Text)
        Me.TxtModFlg.Text = "R"
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
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.title
            RsSavRec.Requery
            Resume BegnieWork
    End Select

End Sub

Private Sub BtnPrevious_Click()
    On Error GoTo ErrTrap
    Dim Msg As String

    If Me.TxtModFlg.Text = "N" Then
        FindRec val(TxtVac_ID.Text)
        Me.TxtModFlg.Text = "R"
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
If TxtVacNamee.Text = "" Then
TxtVacNamee = TxtVacName
End If

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

    StrVacName = IsRecExist("Nationality", "name", Trim(TxtVacName.Text), "name", "Vac_ID<>'" & Trim(TxtVac_ID.Text) & "'")

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

Private Sub cmdDisplay_Click()
Dim s As String

s = " SELECT tc.CusName CustomerName, Legalcourts.Name                 LegalcourtsName,       LegalIssuesData.RecordDate RecordDate5,* FROM SessionDate"
s = s & " LEFT OUTER JOIN LegalIssuesData ON SessionDate.IssuesNo = LegalIssuesData.IssuesNo"
s = s & " LEFT OUTER JOIN TblCustemers AS tc ON LegalIssuesData.CustID = tc.CusID"
s = s & " LEFT OUTER JOIN Legalcourts  ON LegalIssuesData.LegalcourtsID = Legalcourts.ID"
s = s & " Where SessionDate >= GETDATE() "


loadgrid s, VSFlexGrid1, True, False
End Sub

Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim i As Integer
    Dim My_SQL As String
    Dim k As Integer
    If mIndex = 0 Then
        My_SQL = "Nationality"
        Set BKGrndPic = New ClsBackGroundPic
        Set RsSavRec = New ADODB.Recordset
        RsSavRec.CursorLocation = adUseClient
        RsSavRec.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdTableDirect
        Me.TxtModFlg.Text = "R"
       
    End If
    Resize_Form Me
    
    'load tblUsers -----------------------------------------------
    My_SQL = "select UserID,UserName From tblUsers "
    fill_combo DCUser, My_SQL
  

TabMain.TabVisible(0) = False
TabMain.TabVisible(1) = False
TabMain.TabVisible(2) = False
TabMain.TabVisible(mIndex) = True
TabMain.CurrTab = mIndex


    Dim Dcombos As New ClsDataCombos
    Dcombos.GetCodeing Me.DCPreFix, 6
    
   If mIndex = 2 Then
        Set Dcombos = New ClsDataCombos
        Dcombos.GetBoxes Me.DcboBox
        Dcombos.GetBanks Me.DcboBankName
        'Dcombos.GetUsers Me.DCboUserName
       ' Dcombos.GetExpensesType XPCboExpensesType
        Dcombos.GetBranches Me.dcBranch
        Dcombos.GetAccountingCodes Me.DcbAccount, True, False
        '    Dim Dcombos As ClsDataCombos
        'Set Dcombos = New ClsDataCombos
         
        My_SQL = "tblTripTrans"
        Dcombos.GetAccountingCodes Me.DcboDebitSide
        Dcombos.GetAccountingCodes Me.DcboCreditSide
        Dcombos.GetEquipments DcbEqup
         
        Me.Caption = " ’›ÌÂ «·⁄Âœ ··”«∆ﬁÌ‰  "
        dcBranch.BoundText = branch_id
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
        
          With GridInstallments
            .ColComboList(.ColIndex("CarType1")) = "#1;„„·Êﬂ… |#2;„„·Êﬂ… ··€Ì— "
     If SystemOptions.UserInterface = ArabicInterface Then
            .ColComboList(.ColIndex("Typed")) = "#1;ﬂÃ„  |#2;—œ |#3;Ê“‰ "
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
           .ColComboList(.ColIndex("Typed")) = "#1;Kg |#2;Kg |#3;Weight"
      End If
           If SystemOptions.UserInterface = ArabicInterface Then
         .ColComboList(.ColIndex("Show")) = "⁄—÷"
        Else
        .ColComboList(.ColIndex("Show")) = "View"
        End If
        
              
                
    End With
        With Me.CboPaymentType1
        .Clear
        .AddItem "‰ﬁœÌ"
        .AddItem "„‰ Õ”«»"

        End With
        TabMain.TabVisible(2) = True
        TabMain.CurrTab = 2
   ElseIf mIndex = 0 Then
        Me.Width = Grid.Width + 400
        TabMain.TabVisible(0) = True
        TabMain.CurrTab = 0
    ElseIf mIndex = 1 Then
        Me.Width = VSFlexGrid1.Width + 400
        TabMain.TabVisible(1) = True
        TabMain.CurrTab = 1
    

   End If
    If mIndex = 0 Then
        FillGridWithData
        For k = 1 To 10
        DcbQuality.AddItem k
        Next k
    
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
    
        If SystemOptions.UserInterface = EnglishInterface Then
            SetInterface Me
            ChangeLang
        End If
    
        If OPEN_NEW_SCREEN = True Then
            btnNew_Click
        End If
End If
If mIndex > 1 Then
    Set BKGrndPic = New ClsBackGroundPic
    Set RsSavRec = New ADODB.Recordset
    RsSavRec.CursorLocation = adUseClient
    
    RsSavRec.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdTableDirect

    Me.TxtModFlg2(mIndex).Text = "R"
    btn_First_Click (mIndex)
End If

ErrTrap:
End Sub






Public Sub FillGrid()
    Dim i As Double
    Dim Rs3 As ADODB.Recordset
    Dim My_SQL As String
    Set Rs3 = New ADODB.Recordset
 
'My_SQL = " SELECT     dbo.notes_all.NoteID, dbo.notes_all.NoteDate, dbo.notes_all.NoteType, dbo.notes_all.Note_Value, dbo.notes_all.branch_no, dbo.TblBranchesData.branch_name, "
'My_SQL = My_SQL & "                       dbo.TblBranchesData.branch_namee, dbo.notes_all.CusID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode,"
'My_SQL = My_SQL & "                       dbo.notes_all.TotalQty, dbo.notes_all.Typed, dbo.notes_all.Total, dbo.notes_all.Price, dbo.notes_all.VehicleType, dbo.TBLCarTypes.name, dbo.TBLCarTypes.namee,"
'My_SQL = My_SQL & "                       dbo.notes_all.CarId, dbo.TblCarsData.BoardNO, dbo.TblCarsData.Name AS CarName, dbo.notes_all.general_des, dbo.notes_all.DriverId, dbo.TblEmployee.Emp_ID,"
'My_SQL = My_SQL & "                       dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode AS EmpFullcode, dbo.TblEmployee.Emp_Namee, dbo.notes_all.CityFromId,"
'My_SQL = My_SQL & "                       TblCountriesGovernments_1.GovernmentName, dbo.notes_all.CityToId, TblCountriesGovernments_1.GovernmentName AS ToGovernmentName,"
'My_SQL = My_SQL & "                       dbo.notes_all.allocations ,dbo.notes_all.NoteSerial1"
'My_SQL = My_SQL & "  FROM         dbo.TblCountriesGovernments TblCountriesGovernments_1 RIGHT OUTER JOIN"
'My_SQL = My_SQL & "                       dbo.notes_all ON TblCountriesGovernments_1.GovernmentID = dbo.notes_all.CityToId LEFT OUTER JOIN"
'My_SQL = My_SQL & "                       dbo.TblCountriesGovernments TblCountriesGovernments_2 ON dbo.notes_all.CityFromId = TblCountriesGovernments_2.GovernmentID LEFT OUTER JOIN"
'My_SQL = My_SQL & "                       dbo.TblEmployee ON dbo.notes_all.DriverId = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
'My_SQL = My_SQL & "                       dbo.TblCarsData ON dbo.notes_all.CarId = dbo.TblCarsData.id LEFT OUTER JOIN"
'My_SQL = My_SQL & "                       dbo.TBLCarTypes ON dbo.notes_all.VehicleType = dbo.TBLCarTypes.id LEFT OUTER JOIN"
'My_SQL = My_SQL & "                       dbo.TblCustemers ON dbo.notes_all.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
'My_SQL = My_SQL & "                       dbo.TblBranchesData ON dbo.notes_all.branch_no = dbo.TblBranchesData.branch_id"
'My_SQL = My_SQL & "  Where (dbo.notes_all.notetype = 370) and (dbo.notes_all.allocations=0  or dbo.notes_all.allocations is null)"
'My_SQL = My_SQL + " and (dbo.notes_all.NoteDate >=" & SQLDate(Me.Fromdate, True) & ""
'My_SQL = My_SQL + " and dbo.notes_all.NoteDate <=" & SQLDate(todate, True) & " )"
'My_SQL = My_SQL + "   order by dbo.notes_all.NoteSerial1 "
My_SQL = " SELECT     dbo.TblTripTypesTransport.BillDate,dbo.TblTripTypesTransport.ID, dbo.TblTripTypesTransport.NotesallID, dbo.TblTripTypesTransport.CardNO, dbo.TblTripTypesTransport.QtyDownload, "
My_SQL = My_SQL & "                      dbo.TblTripTypesTransport.CardNO2, dbo.TblTripTypesTransport.QtyDischarge, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee,"
My_SQL = My_SQL & "                      dbo.notes_all.NoteDate,dbo.notes_all.RecNo,dbo.notes_all.Weight, dbo.notes_all.NoteSerial1, dbo.notes_all.general_des, dbo.notes_all.CityFromId, TblCountriesGovernments_2.GovernmentName,"
My_SQL = My_SQL & "                      dbo.notes_all.CityToId, TblCountriesGovernments_1.GovernmentName AS GovernmentNameTO, dbo.notes_all.VehicleType, dbo.TBLCarTypes.name,"
My_SQL = My_SQL & "                      dbo.TBLCarTypes.namee, dbo.notes_all.CarId,dbo.notes_all.Price, dbo.TblCarsData.BoardNO, dbo.notes_all.CarID2, dbo.TblVendorCars.BoardNo AS BoardNo2, dbo.notes_all.CusID,"
My_SQL = My_SQL & "                      dbo.TblTripTypesTransport.ItemID, dbo.notes_all.TypeTransportID, dbo.notes_all.NoteID, dbo.notes_all.NoteType, dbo.notes_all.branch_no, dbo.notes_all.CarType,"
My_SQL = My_SQL & "                      dbo.notes_all.ShipID, dbo.TblShipsData.Name AS ShipName, dbo.TblShipsData.NameE AS ShipNameE, dbo.notes_all.DriverId, dbo.TblEmployee.Emp_Name,"
My_SQL = My_SQL & "                      dbo.TblEmployee.fullcode , dbo.TblEmployee.Emp_Namee, dbo.notes_all.LeaderName,TblTripTypesTransport.HOverVoucher,"
My_SQL = My_SQL & "                      TblCustemers.CusName ,TblCustemers.CusID,TblCarsData.BoardNO,notes_all.DriverValue as TotalValue,TblCarsData.fixedAssetid"
My_SQL = My_SQL & " FROM         dbo.notes_all LEFT OUTER JOIN"
My_SQL = My_SQL & "                      dbo.TblEmployee ON dbo.notes_all.DriverId = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
My_SQL = My_SQL & "                      dbo.TblShipsData ON dbo.notes_all.ShipID = dbo.TblShipsData.id RIGHT OUTER JOIN"
My_SQL = My_SQL & "                      dbo.TblTripTypesTransport ON dbo.notes_all.NoteID = dbo.TblTripTypesTransport.NotesallID LEFT OUTER JOIN"
My_SQL = My_SQL & "                      dbo.TblVendorCars ON dbo.notes_all.CarID2 = dbo.TblVendorCars.ID LEFT OUTER JOIN"
My_SQL = My_SQL & "                      dbo.TblCarsData ON dbo.notes_all.CarId = dbo.TblCarsData.id LEFT OUTER JOIN"
My_SQL = My_SQL & "                      dbo.TBLCarTypes ON dbo.notes_all.VehicleType = dbo.TBLCarTypes.id LEFT OUTER JOIN"
My_SQL = My_SQL & "                      dbo.TblCountriesGovernments TblCountriesGovernments_1 ON dbo.notes_all.CityToId = TblCountriesGovernments_1.GovernmentID LEFT OUTER JOIN"
My_SQL = My_SQL & "                      dbo.TblCountriesGovernments TblCountriesGovernments_2 ON dbo.notes_all.CityFromId = TblCountriesGovernments_2.GovernmentID LEFT OUTER JOIN"
My_SQL = My_SQL & "                      dbo.TblBranchesData ON dbo.notes_all.branch_no = dbo.TblBranchesData.branch_id"
My_SQL = My_SQL & "                      LEFT OUTER JOIN TblCustemers On TblCustemers.CusID = dbo.notes_all.CusID "
My_SQL = My_SQL & "  Where (dbo.notes_all.notetype = 370) AND IsNull(dbo.TblTripTypesTransport.allocations,0) = 0"


My_SQL = My_SQL + "  and (dbo.TblTripTypesTransport.BillDate >=" & SQLDate(Me.FromDate, True) & ""
My_SQL = My_SQL + "  and dbo.TblTripTypesTransport.BillDate <=" & SQLDate(toDate, True) & " )"

If Check2.value = vbChecked Then
    If val(DcbLeaderID.BoundText) <> 0 Then
        My_SQL = My_SQL & "  and (dbo.TblEmployee.Emp_ID = " & val(DcbLeaderID.BoundText) & ") "
    End If

    
End If
'My_SQL = My_SQL & "  and (dbo.notes_all.CusID = " & val(DBCboClientName2.BoundText) & ") "
'If val(DcbTypeTransport.BoundText) <> 0 Then
'My_SQL = My_SQL & "  and (dbo.notes_all.TypeTransportID = " & val(DcbTypeTransport.BoundText) & ") "
'End If
'If val(DCboItemS.BoundText) <> 0 Then
'My_SQL = My_SQL & "  and (dbo.TblTripTypesTransport.ItemID = " & val(DCboItemS.BoundText) & ") "
'End If
'If val(DcCityFromId.BoundText) <> 0 Then
'My_SQL = My_SQL & "  and (dbo.notes_all.CityFromId= " & val(DcCityFromId.BoundText) & ") "
'End If
'If val(DcCityToId.BoundText) <> 0 Then
'My_SQL = My_SQL & "  and (dbo.notes_all.CityToId= " & val(DcCityToId.BoundText) & ") "
'End If
'If val(DcbShip.BoundText) <> 0 Then
'My_SQL = My_SQL & "  and (dbo.notes_all.ShipID= " & val(DcbShip.BoundText) & ") "
'End If


 My_SQL = My_SQL & "  and  NoteSerial1  not in (  "
My_SQL = My_SQL & " SELECT     TripNo FROM         dbo.tblTripTrans2 Where TravID <>  " & val(TxtSerial1(mIndex))
'If val(TxtTransID) <> 0 Then
'    My_SQL = My_SQL & "  WHERE     (TravID <> " & TxtTransID & ")"
'End If
  My_SQL = My_SQL & " )"


'If ChkDate.value = vbChecked Then
'
'My_SQL = My_SQL + "   order by dbo.notes_all.NoteDate "
'Else
'My_SQL = My_SQL + "   order by dbo.notes_all.NoteSerial1 "
'End If

 
    Rs3.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText
    

      With Me.GridInstallments
      Dim Xb As Integer
       .Rows = 1
        .Clear flexClearScrollable
        If Rs3.RecordCount > 0 Then
           .Rows = Rs3.RecordCount + 1
           Rs3.MoveFirst
            For i = 1 To .Rows - 1
            
                       If SystemOptions.UserInterface = ArabicInterface Then
         .ColComboList(.ColIndex("Show")) = "⁄—÷"
        Else
        .ColComboList(.ColIndex("Show")) = "View"
        End If
        
        
            .TextMatrix(i, .ColIndex("Ser")) = i
            .TextMatrix(i, .ColIndex("Select")) = 1
           .TextMatrix(i, .ColIndex("ShipID")) = (IIf(IsNull(Rs3.Fields("ShipID").value), 0, Rs3.Fields("ShipID").value))
            .TextMatrix(i, .ColIndex("NoteID")) = (IIf(IsNull(Rs3.Fields("ID").value), 0, Rs3.Fields("ID").value))
            .TextMatrix(i, .ColIndex("NoteIDA")) = (IIf(IsNull(Rs3.Fields("NoteID").value), 0, Rs3.Fields("NoteID").value))
            
            .TextMatrix(i, .ColIndex("TripNo")) = (IIf(IsNull(Rs3.Fields("NoteSerial1").value), "", Rs3.Fields("NoteSerial1").value))
            .TextMatrix(i, .ColIndex("TripDate")) = (IIf(IsNull(Rs3.Fields("BillDate").value), "", Rs3.Fields("BillDate").value))
            .TextMatrix(i, .ColIndex("BranchID")) = (IIf(IsNull(Rs3.Fields("branch_no").value), 0, Rs3.Fields("branch_no").value))
            
            .TextMatrix(i, .ColIndex("CusID")) = (IIf(IsNull(Rs3.Fields("CusID").value), 0, Rs3.Fields("CusID").value))
            .TextMatrix(i, .ColIndex("CusName")) = (IIf(IsNull(Rs3.Fields("CusName").value), "", Rs3.Fields("CusName").value))
            .TextMatrix(i, .ColIndex("BoardNO")) = (IIf(IsNull(Rs3.Fields("BoardNO").value), "", Rs3.Fields("BoardNO").value))
            .TextMatrix(i, .ColIndex("fixedAssetid")) = (IIf(IsNull(Rs3.Fields("fixedAssetid").value), "", Rs3.Fields("fixedAssetid").value))
            .TextMatrix(i, .ColIndex("CardNO")) = (IIf(IsNull(Rs3.Fields("CardNO").value), "", Rs3.Fields("CardNO").value))
            .TextMatrix(i, .ColIndex("QtyDownload")) = (IIf(IsNull(Rs3.Fields("QtyDownload").value), "", Rs3.Fields("QtyDownload").value))
           ' Xb = (IIf(IsNull(Rs3.Fields("Typed").value), 0, Rs3.Fields("Typed").value))
           ' .TextMatrix(i, .ColIndex("Typed")) = Xb + 1
            .TextMatrix(i, .ColIndex("CarType1")) = (IIf(IsNull(Rs3.Fields("CarType").value), 0, Rs3.Fields("CarType").value)) + 1
            .TextMatrix(i, .ColIndex("CardNO2")) = (IIf(IsNull(Rs3.Fields("CardNO2").value), "", Rs3.Fields("CardNO2").value))
            .TextMatrix(i, .ColIndex("FromID")) = (IIf(IsNull(Rs3.Fields("CityFromId").value), 0, Rs3.Fields("CityFromId").value))
            .TextMatrix(i, .ColIndex("ToID")) = (IIf(IsNull(Rs3.Fields("CityToId").value), 0, Rs3.Fields("CityToId").value))
            .TextMatrix(i, .ColIndex("TotalValue")) = (IIf(IsNull(Rs3.Fields("TotalValue").value), 0, Rs3.Fields("TotalValue").value))
            .TextMatrix(i, .ColIndex("From")) = (IIf(IsNull(Rs3.Fields("GovernmentName").value), "", Rs3.Fields("GovernmentName").value))
            .TextMatrix(i, .ColIndex("To")) = (IIf(IsNull(Rs3.Fields("GovernmentNameTO").value), "", Rs3.Fields("GovernmentNameTO").value))
            .TextMatrix(i, .ColIndex("CarTypeID")) = (IIf(IsNull(Rs3.Fields("VehicleType").value), 0, Rs3.Fields("VehicleType").value))
            If val(.TextMatrix(i, .ColIndex("CarType1"))) = 1 Then
            .TextMatrix(i, .ColIndex("CarType1")) = 1
            End If
            If val(.TextMatrix(i, .ColIndex("CarType1"))) = 2 Then
            .TextMatrix(i, .ColIndex("CarID")) = (IIf(IsNull(Rs3.Fields("CarID2").value), 0, Rs3.Fields("CarID2").value))
            .TextMatrix(i, .ColIndex("Car")) = (IIf(IsNull(Rs3.Fields("BoardNo2").value), "", Rs3.Fields("BoardNo2").value))
            Else
             .TextMatrix(i, .ColIndex("CarID")) = (IIf(IsNull(Rs3.Fields("CarId").value), 0, Rs3.Fields("CarId").value))
            .TextMatrix(i, .ColIndex("Car")) = (IIf(IsNull(Rs3.Fields("BoardNO").value), "", Rs3.Fields("BoardNO").value))
            End If
            .TextMatrix(i, .ColIndex("QtyDischarge")) = (IIf(IsNull(Rs3.Fields("QtyDischarge").value), 0, Rs3.Fields("QtyDischarge").value))
            .TextMatrix(i, .ColIndex("Remarks")) = (IIf(IsNull(Rs3.Fields("general_des").value), "", Rs3.Fields("general_des").value))
            
            .TextMatrix(i, .ColIndex("Value")) = (IIf(IsNull(Rs3.Fields("Price").value), "", Rs3.Fields("Price").value))
            If val(.TextMatrix(i, .ColIndex("TotalValue"))) = 0 Then
           '     .TextMatrix(i, .ColIndex("TotalValue")) = .TextMatrix(i, .ColIndex("Value"))
            End If
            .TextMatrix(i, .ColIndex("RecNo")) = (IIf(IsNull(Rs3.Fields("HOverVoucher").value), IIf(IsNull(Rs3.Fields("RecNo").value), "", Rs3.Fields("RecNo").value), Rs3.Fields("HOverVoucher").value))
            .TextMatrix(i, .ColIndex("Weight")) = (IIf(IsNull(Rs3.Fields("Weight").value), "", Rs3.Fields("Weight").value))
            
            If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(i, .ColIndex("Ship")) = (IIf(IsNull(Rs3.Fields("ShipName").value), "", Rs3.Fields("ShipName").value))
            .TextMatrix(i, .ColIndex("CarType")) = (IIf(IsNull(Rs3.Fields("name").value), "", Rs3.Fields("name").value))
            .TextMatrix(i, .ColIndex("EmpName")) = (IIf(IsNull(Rs3.Fields("Emp_Name").value), (IIf(IsNull(Rs3.Fields("LeaderName").value), "", Rs3.Fields("LeaderName").value)), Rs3.Fields("Emp_Name").value))
            .TextMatrix(i, .ColIndex("Branch")) = (IIf(IsNull(Rs3.Fields("branch_name").value), "", Rs3.Fields("branch_name").value))
            Else
            
            .TextMatrix(i, .ColIndex("Ship")) = (IIf(IsNull(Rs3.Fields("NameE").value), "", Rs3.Fields("NameE").value))
            .TextMatrix(i, .ColIndex("CarType")) = (IIf(IsNull(Rs3.Fields("namee").value), "", Rs3.Fields("namee").value))
            .TextMatrix(i, .ColIndex("Branch")) = (IIf(IsNull(Rs3.Fields("branch_namee").value), "", Rs3.Fields("branch_namee").value))
            .TextMatrix(i, .ColIndex("EmpName")) = (IIf(IsNull(Rs3.Fields("Emp_Namee").value), (IIf(IsNull(Rs3.Fields("LeaderName").value), "", Rs3.Fields("LeaderName").value)), Rs3.Fields("Emp_Namee").value))
            End If
        Rs3.MoveNext

            Next i
 End If
            Rs3.Close
        .RowHeight(-1) = 300
    End With
ReLineGrid
End Sub
Private Sub ReLineGrid()
'ReLineGrid2

Dim SumVal As Double
Dim SumPrice As Double
Dim i As Integer
Dim sumQtyDischarge As Double
sumQtyDischarge = 0
SumVal = 0
If Me.TxtModFlg.Text <> "R" Then
With GridInstallments
For i = 1 To .Rows - 1
'.TextMatrix(i, .ColIndex("Select")) = 1
If .Cell(flexcpChecked, i, .ColIndex("Select")) = flexChecked And SystemOptions.TransBillPriceByGrid = True Then
            If val(.TextMatrix(i, .ColIndex("QtyDownload"))) <> 0 Then
'            .TextMatrix(i, .ColIndex("TotalValue")) = Round(.TextMatrix(i, .ColIndex("QtyDownload")), 3) * Round(.TextMatrix(i, .ColIndex("Value")), 3)
            Else
'            .TextMatrix(i, .ColIndex("TotalValue")) = Round(val(.TextMatrix(i, .ColIndex("QtyDischarge"))), 3) * Round(val(.TextMatrix(i, .ColIndex("Value"))), 3)
            End If
'            .TextMatrix(i, .ColIndex("TotalValue")) = Round(.TextMatrix(i, .ColIndex("TotalValue")), 3)
Else
'.TextMatrix(i, .ColIndex("TotalValue")) = 0
End If
            If .Cell(flexcpChecked, i, .ColIndex("Select")) = flexChecked Then
            SumVal = SumVal + Round(val(.TextMatrix(i, .ColIndex("QtyDownload"))), 3)
            SumPrice = SumPrice + Round(val(.TextMatrix(i, .ColIndex("TotalValue"))), 3)
            sumQtyDischarge = sumQtyDischarge + Round((val(.TextMatrix(i, .ColIndex("QtyDischarge")))), 3)
            End If
Next i
End With
TxtQtyDownload.Text = Round(SumVal, 3)
TxtQtyDischarge.Text = Round(sumQtyDischarge, 3)
If SystemOptions.TransBillPriceByGrid = True Then
TxtTotalValue.Text = SumPrice
End If
Calculte
'FillGrid2
Else

End If
End Sub


Private Sub ReLineGrid2()
'ReLineGrid2

    Dim SumVal As Double
  
    Dim i As Integer
   
    
    SumVal = 0
   ' If Me.TxtModFlg2(mIndex).Text <> "R" Then
        With FG
            For i = 1 To .Rows - 1
                SumVal = SumVal + Round(val(.TextMatrix(i, .ColIndex("Price"))), 3)
                        
            Next i
        End With
        txtTotal2.Text = Round(SumVal, 3)
    
   ' End If
End Sub

Sub Calculte()
If SystemOptions.TransBillPriceByGrid = False Then
If val(Me.TxtPrice.Text) = 0 Then Me.TxtPrice.Text = 0
If RdQty(1).value = True Then
TxtTotalValue.Text = val(Me.TxtPrice.Text) * val(TxtQtyDischarge.Text)
Else
TxtTotalValue.Text = Round(Me.TxtPrice.Text, 3) * Round(val(TxtQtyDownload.Text), 3)
TxtTotalValue.Text = Round(TxtTotalValue.Text, 3)
End If
End If
'TxtNetValue.Text = val(TxtTotalValue.Text) '- val(TxtDiscount.Text)
'TxtVAT.Text = Round((TxtNetValue.Text) * 5 / 100, 3)

'If chkoWithoutVat.value = vbChecked Then
''TxtVAT.Text = 0
'End If

'TxtTotal.Text = val(TxtNetValue.Text) '+ val(TxtVAT.Text)
'TxtTotal.Text = Round(TxtTotal..Text, 3)
End Sub
Function GetOwnerName(Optional ID As Double) As String
Dim sql As String
Dim Rs2 As ADODB.Recordset
Set Rs2 = New ADODB.Recordset
sql = " SELECT     dbo.TblVendorCars.ID, dbo.TblVendorCars.CustomerID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee"
sql = sql & " FROM         dbo.TblVendorCars LEFT OUTER JOIN"
sql = sql & "                      dbo.TblCustemers ON dbo.TblVendorCars.CustomerID = dbo.TblCustemers.CusID"
sql = sql & " Where (dbo.TblVendorCars.ID = " & ID & ")"
Rs2.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs2.RecordCount > 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
GetOwnerName = IIf(IsNull(Rs2("CusName").value), "", Rs2("CusName").value)
Else
GetOwnerName = IIf(IsNull(Rs2("CusNamee").value), "", Rs2("CusNamee").value)
End If
Else
GetOwnerName = ""
End If
End Function

Private Sub ChangeLang()
    Dim XPic As IPictureDisp
    Set XPic = Me.btnFirst.ButtonImage
    Set Me.btnFirst.ButtonImage = Me.btnLast.ButtonImage
    Set Me.btnLast.ButtonImage = XPic
    Set XPic = Me.btnPrevious.ButtonImage
    Set Me.btnPrevious.ButtonImage = Me.btnNext.ButtonImage
    Set Me.btnNext.ButtonImage = XPic

    Me.Caption = "Nationality"
    Label3.Caption = "Quality start from 1 to 10  1 is the best quality"
    Frame1.Caption = "Quality Hint"
    Label1(4).Caption = "N Code"
    Me.Label1(2).Caption = Me.Caption
    Label1(3).Caption = "Code"
    Label1(0).Caption = "Name AR"
    Label1(1).Caption = "Name ENG"
    Label1(5).Caption = "Manf Quality"
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
        .TextMatrix(0, .ColIndex("NCODE")) = "N Code"
    End With

End Sub

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
                btnSave_Click

            Case vbCancel
                Cancel = True
        End Select

    End If

    Exit Sub
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

Public Sub AddNewRec()

    On Error GoTo ErrTrap
    Dim StrRecID As String
    If mIndex = 0 Then
        StrRecID = new_id("Nationality", "id", "")
    
        RsSavRec.AddNew
        RsSavRec.Fields("id").value = IIf(StrRecID <> "", StrRecID, Null)
        FiLLRec
    ElseIf mIndex = 2 Then
        StrRecID = new_id("tblTripTrans", "id", "")
        cmdCreateENtry.Enabled = False
        cmdDeleteEntry.Enabled = False
        dcBranch.BoundText = branch_id
        RsSavRec.AddNew
        RsSavRec.Fields("id").value = IIf(StrRecID <> "", StrRecID, Null)
       'FiLLRec2
    End If
ErrTrap:
End Sub

Public Sub FiLLRec()
    On Error GoTo ErrTrap

    RsSavRec.Fields("name").value = IIf(TxtVacName.Text <> "", Trim(TxtVacName.Text), Null)
    RsSavRec.Fields("namee").value = IIf(TxtVacNamee.Text <> "", Trim(TxtVacNamee.Text), Null)
    RsSavRec.Fields("NCODE").value = IIf(DCPreFix.Text <> "", Trim(DCPreFix.Text), Null)
    RsSavRec.Fields("Quality").value = IIf(val(Me.DcbQuality.ListIndex) <> -1, val(DcbQuality.ListIndex), Null)

    RsSavRec.update
    If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox " „  ⁄„·Ì… «·Õ›Ÿ »‰Ã«Õ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Else
    MsgBox "Saved Success", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
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
    Frm2.Enabled = False
    TxtVac_ID.Text = IIf(IsNull(RsSavRec.Fields("id").value), "", RsSavRec.Fields("id").value)
    TxtVacName.Text = IIf(IsNull(RsSavRec.Fields("name").value), "", RsSavRec.Fields("name").value)
    TxtVacNamee.Text = IIf(IsNull(RsSavRec.Fields("namee").value), "", RsSavRec.Fields("namee").value)
    DCPreFix.Text = IIf(IsNull(RsSavRec.Fields("NCODE").value), "", RsSavRec.Fields("NCODE").value)
     Me.DcbQuality.ListIndex = IIf(IsNull(RsSavRec.Fields("Quality").value), -1, RsSavRec.Fields("Quality").value)
'   RsSavRec.Fields("NCODE").value = IIf(DCPreFix.text <> "", Trim(DCPreFix.text), Null)

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

Private Sub Grid_EnterCell()
    On Error GoTo ErrTrap
    FindRec val(Me.Grid.TextMatrix(Me.Grid.Row, Me.Grid.ColIndex("id")))
ErrTrap:
End Sub

Private Sub TxtVac_ID_Change()
    Dim TxtMod As String
    TxtMod = TxtModFlg.Text
    TxtModFlg.Text = ""
    TxtModFlg = TxtMod
End Sub

Public Function FindRec(ByVal RecId As Long)
    On Error GoTo ErrTrap
    RsSavRec.find "id=" & RecId, , adSearchForward, 1
    
    If Not (RsSavRec.EOF) Then
        If mIndex = 2 Then
            FiLLTXT2
        Else
            FiLLTXT
        End If
    
    End If

    Exit Function
ErrTrap:

    If RsSavRec.EditMode <> adEditNone Then
        RsSavRec.CancelUpdate
        BtnUndo_Click
    End If

    'RsSavRec.Filter = adFilterNone
End Function

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
        '    btnNext.Enabled = False
        '    btnPrevious.Enabled = False
        '    btnFirst.Enabled = False
        '    btnLast.Enabled = False
    
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
    My_SQL = "select * From Nationality order by id"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText

    With Me.Grid
        .Rows = 2
        .Clear flexClearScrollable

        If rs.RecordCount > 0 Then
            .Rows = rs.RecordCount + 1
            rs.MoveFirst

            For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("Ser")) = i
             .TextMatrix(i, .ColIndex("NCODE")) = IIf(IsNull(rs.Fields("NCODE").value), "", rs.Fields("NCODE").value)

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

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        If Me.TxtModFlg.Text = "R" Then
            btnNew_Click
        Else
            SendKeys "{TAB}"
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

Private Sub TxtVacName_GotFocus()
    SwitchKeyboardLang LANG_ARABIC
End Sub

Private Sub TxtVacNamee_GotFocus()
SwitchKeyboardLang LANG_ENGLISH
End Sub



Public Sub FiLLTXT2()

    On Error GoTo ErrTrap
    Dim i As Integer
    'Frame1(2).Enabled = False
    TxtSerial1(mIndex).Text = IIf(IsNull(RsSavRec.Fields("id").value), "", RsSavRec.Fields("id").value)

    TXTNoteID.Text = IIf(IsNull(RsSavRec.Fields("NoteID").value), "", RsSavRec.Fields("NoteID").value)
    TxtNoteSerial1.Text = IIf(IsNull(RsSavRec.Fields("NoteSerial1").value), "", RsSavRec.Fields("NoteSerial1").value)

    Me.TxtNoteSerial.Text = IIf(IsNull(RsSavRec("NoteSerial").value), "", RsSavRec("NoteSerial").value)

    XPDtbTrans(mIndex).value = IIf(IsNull(RsSavRec("RecordDate").value), Date, RsSavRec("RecordDate").value)
    XPDtbTransH(mIndex).value = IIf(IsNull(RsSavRec("RecorddateH").value), ToHijriDate(Date), RsSavRec("RecorddateH").value)


    FromDate.value = IIf(IsNull(RsSavRec("Fromdate").value), Date, RsSavRec("Fromdate").value)
    Me.Fromdate√H.value = IIf(IsNull(RsSavRec("FromDateh").value), ToHijriDate(Date), RsSavRec("FromDateh").value)
    toDate.value = IIf(IsNull(RsSavRec("todate").value), Date, RsSavRec("todate").value)
    toDateH.value = IIf(IsNull(RsSavRec("todateH").value), ToHijriDate(Date), RsSavRec("todateH").value)

    
    DcboBox.BoundText = IIf(IsNull(RsSavRec.Fields("BoxID").value), "", RsSavRec.Fields("BoxID").value)
    dcBranch.BoundText = IIf(IsNull(RsSavRec.Fields("BranchId").value), "", RsSavRec.Fields("BranchId").value)
    DcbAccount.BoundText = IIf(IsNull(RsSavRec.Fields("AccountPaym").value), "", RsSavRec.Fields("AccountPaym").value)
    CboPaymentType1.ListIndex = IIf(IsNull(RsSavRec.Fields("PaymentType").value), 0, RsSavRec.Fields("PaymentType").value)
    
    If RsSavRec("PaymentType").value = 0 Then
        DcboCreditSide.BoundText = ModAccounts.GetMyAccountCode("TblBoxesData", "BoxID", val(Me.DcboBox.BoundText))
    Else
        DcboCreditSide.BoundText = DcbAccount.BoundText
    End If
    
    LabCurr_Rec(mIndex).Caption = RsSavRec.AbsolutePosition
    LabCount_Rec(mIndex).Caption = RsSavRec.RecordCount

   
  Dim StrSQL As String
  Dim RsDev As ADODB.Recordset
  'Me.TxtTransID.Text = IIf(IsNull(rs("ID").value), "", rs("ID").value)
StrSQL = " SELECT   notesallid,  dbo.tblTripTrans2.notesallid,  dbo.tblTripTrans2.ID, dbo.tblTripTrans2.TravID, dbo.tblTripTrans2.TripNo, dbo.tblTripTrans2.TripDate, dbo.tblTripTrans2.BranchID, "
StrSQL = StrSQL & "                      dbo.tblTripTrans2.Price,dbo.tblTripTrans2.TotalValue,tblTripTrans2.RecNo,tblTripTrans2.Weight,"
StrSQL = StrSQL & "                      dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.tblTripTrans2.Typed, dbo.tblTripTrans2.[Value], dbo.tblTripTrans2.Remarks,"
StrSQL = StrSQL & "                      dbo.tblTripTrans2.NoteID, dbo.tblTripTrans2.QtyDownload, dbo.tblTripTrans2.QtyDischarge, dbo.tblTripTrans2.CardNO, dbo.tblTripTrans2.CardNO2,"
StrSQL = StrSQL & "                      dbo.tblTripTrans2.CarType1, dbo.tblTripTrans2.CarID, dbo.TblCarsData.BoardNO, dbo.TblVendorCars.BoardNo AS BoardNo2, dbo.tblTripTrans2.FromID,"
StrSQL = StrSQL & "                      TblCountriesGovernments_2.GovernmentName, dbo.tblTripTrans2.ToID, TblCountriesGovernments_1.GovernmentName AS ToGovernmentName,"
StrSQL = StrSQL & "                      dbo.tblTripTrans2.CarTypeID, dbo.TBLCarTypes.name, dbo.TBLCarTypes.namee, dbo.tblTripTrans2.TypeTrans, dbo.tblTripTrans2.ShipID,"
StrSQL = StrSQL & "                      dbo.TblShipsData.Name AS ShipName, dbo.TblShipsData.NameE AS ShipNameE, dbo.tblTripTrans2.LeaderName,TblCustemers.CusName ,TblCustemers.CusID ,TblCarsData.BoardNO,TblCarsData.fixedAssetid"
StrSQL = StrSQL & " FROM         dbo.tblTripTrans2 LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblShipsData ON dbo.tblTripTrans2.ShipID = dbo.TblShipsData.id LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TBLCarTypes ON dbo.tblTripTrans2.CarTypeID = dbo.TBLCarTypes.id LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblCountriesGovernments TblCountriesGovernments_1 ON dbo.tblTripTrans2.ToID = TblCountriesGovernments_1.GovernmentID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblCountriesGovernments TblCountriesGovernments_2 ON dbo.tblTripTrans2.FromID = TblCountriesGovernments_2.GovernmentID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblVendorCars ON dbo.tblTripTrans2.CarID = dbo.TblVendorCars.ID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblCarsData ON dbo.tblTripTrans2.CarID = dbo.TblCarsData.id LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblBranchesData ON dbo.tblTripTrans2.BranchID = dbo.TblBranchesData.branch_id"
StrSQL = StrSQL & "                      LEFT OUTER JOIN TblCustemers On TblCustemers.CusID = dbo.tblTripTrans2.CusID "
StrSQL = StrSQL & "   Where (dbo.tblTripTrans2.TravID = " & val(TxtSerial1(mIndex).Text) & ") and (dbo.tblTripTrans2.TypeTrans is null or dbo.tblTripTrans2.TypeTrans=0)  "
    Set RsDev = New ADODB.Recordset
    RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    GridInstallments.Rows = 1
    If Not (RsDev.BOF Or RsDev.EOF) Then
        RsDev.MoveFirst
        With Me.GridInstallments
            
            .Rows = .FixedRows + RsDev.RecordCount
            For i = .FixedRows To .Rows - 1
                       If SystemOptions.UserInterface = ArabicInterface Then
         .ColComboList(.ColIndex("Show")) = "⁄—÷"
        Else
        .ColComboList(.ColIndex("Show")) = "View"
        End If
        
        .TextMatrix(i, .ColIndex("Ser")) = i
           .Cell(flexcpChecked, i, .ColIndex("Select")) = flexChecked
           'RsDetails1("notesallid").value = val(.TextMatrix(i, .ColIndex("NoteIDA")))
           
           .TextMatrix(i, .ColIndex("NoteIDA")) = (IIf(IsNull(RsDev.Fields("notesallid").value), 0, RsDev.Fields("notesallid").value))
           
           .TextMatrix(i, .ColIndex("EmpName")) = (IIf(IsNull(RsDev.Fields("LeaderName").value), "", RsDev.Fields("LeaderName").value))
           .TextMatrix(i, .ColIndex("ShipID")) = (IIf(IsNull(RsDev.Fields("ShipID").value), 0, RsDev.Fields("ShipID").value))
           .TextMatrix(i, .ColIndex("TripNo")) = (IIf(IsNull(RsDev.Fields("TripNo").value), "", RsDev.Fields("TripNo").value))
           .TextMatrix(i, .ColIndex("TripDate")) = (IIf(IsNull(RsDev.Fields("TripDate").value), "", RsDev.Fields("TripDate").value))
           .TextMatrix(i, .ColIndex("BranchID")) = (IIf(IsNull(RsDev.Fields("BranchID").value), 0, RsDev.Fields("BranchID").value))
           .TextMatrix(i, .ColIndex("QtyDownload")) = (IIf(IsNull(RsDev.Fields("QtyDownload").value), "", RsDev.Fields("QtyDownload").value))
           .TextMatrix(i, .ColIndex("Value")) = (IIf(IsNull(RsDev.Fields("Price").value), "", RsDev.Fields("Price").value))
           .TextMatrix(i, .ColIndex("TotalValue")) = (IIf(IsNull(RsDev.Fields("TotalValue").value), "", RsDev.Fields("TotalValue").value))
           .TextMatrix(i, .ColIndex("Weight")) = (IIf(IsNull(RsDev.Fields("Weight").value), "", RsDev.Fields("Weight").value))
           .TextMatrix(i, .ColIndex("RecNo")) = (IIf(IsNull(RsDev.Fields("RecNo").value), "", RsDev.Fields("RecNo").value))
           .TextMatrix(i, .ColIndex("QtyDischarge")) = (IIf(IsNull(RsDev.Fields("QtyDischarge").value), "", RsDev.Fields("QtyDischarge").value))
           .TextMatrix(i, .ColIndex("CarType1")) = (IIf(IsNull(RsDev.Fields("CarType1").value), 1, RsDev.Fields("CarType1").value))
           .TextMatrix(i, .ColIndex("CardNO")) = (IIf(IsNull(RsDev.Fields("CardNO").value), "", RsDev.Fields("CardNO").value))
            .TextMatrix(i, .ColIndex("BoardNO")) = (IIf(IsNull(RsDev.Fields("BoardNO").value), "", RsDev.Fields("BoardNO").value))
            .TextMatrix(i, .ColIndex("fixedAssetid")) = (IIf(IsNull(RsDev.Fields("fixedAssetid").value), "", RsDev.Fields("fixedAssetid").value))
            
           .TextMatrix(i, .ColIndex("CardNO2")) = (IIf(IsNull(RsDev.Fields("CardNO2").value), "", RsDev.Fields("CardNO2").value))
           .TextMatrix(i, .ColIndex("Remarks")) = (IIf(IsNull(RsDev.Fields("Remarks").value), "", RsDev.Fields("Remarks").value))
           .TextMatrix(i, .ColIndex("FromID")) = (IIf(IsNull(RsDev.Fields("FromID").value), 0, RsDev.Fields("FromID").value))
          .TextMatrix(i, .ColIndex("From")) = (IIf(IsNull(RsDev.Fields("GovernmentName").value), "", RsDev.Fields("GovernmentName").value))
            
            .TextMatrix(i, .ColIndex("CusID")) = (IIf(IsNull(RsDev.Fields("CusID").value), 0, RsDev.Fields("CusID").value))
            .TextMatrix(i, .ColIndex("CusName")) = (IIf(IsNull(RsDev.Fields("CusName").value), "", RsDev.Fields("CusName").value))
          .TextMatrix(i, .ColIndex("ToID")) = (IIf(IsNull(RsDev.Fields("ToID").value), 0, RsDev.Fields("ToID").value))
          .TextMatrix(i, .ColIndex("To")) = (IIf(IsNull(RsDev.Fields("ToGovernmentName").value), "", RsDev.Fields("ToGovernmentName").value))
          .TextMatrix(i, .ColIndex("CarTypeID")) = (IIf(IsNull(RsDev.Fields("CarTypeID").value), 0, RsDev.Fields("CarTypeID").value))
          .TextMatrix(i, .ColIndex("CarID")) = (IIf(IsNull(RsDev.Fields("CarID").value), 0, RsDev.Fields("CarID").value))
          If val(.TextMatrix(i, .ColIndex("CarType1"))) = 2 Then
          .TextMatrix(i, .ColIndex("Car")) = (IIf(IsNull(RsDev.Fields("BoardNo2").value), "", RsDev.Fields("BoardNo2").value))
          Else
          .TextMatrix(i, .ColIndex("Car")) = (IIf(IsNull(RsDev.Fields("BoardNO").value), "", RsDev.Fields("BoardNO").value))
          End If
            .TextMatrix(i, .ColIndex("NoteID")) = (IIf(IsNull(RsDev.Fields("NoteID").value), 0, RsDev.Fields("NoteID").value))
        If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(i, .ColIndex("Ship")) = (IIf(IsNull(RsDev.Fields("ShipName").value), "", RsDev.Fields("ShipName").value))
            .TextMatrix(i, .ColIndex("CarType")) = (IIf(IsNull(RsDev.Fields("name").value), "", RsDev.Fields("name").value))
            .TextMatrix(i, .ColIndex("Branch")) = (IIf(IsNull(RsDev.Fields("branch_name").value), "", RsDev.Fields("branch_name").value))
         Else
         .TextMatrix(i, .ColIndex("Ship")) = (IIf(IsNull(RsDev.Fields("ShipNameE").value), "", RsDev.Fields("ShipNameE").value))
         .TextMatrix(i, .ColIndex("CarType")) = (IIf(IsNull(RsDev.Fields("namee").value), "", RsDev.Fields("namee").value))
         .TextMatrix(i, .ColIndex("Branch")) = (IIf(IsNull(RsDev.Fields("branch_namee").value), "", RsDev.Fields("branch_namee").value))
        End If
        RsDev.MoveNext
        Next i
        End With
    End If
 RsDev.Close
ReLineGrid

If Trim(TxtNoteSerial) = "" Then
    cmdCreateENtry.Enabled = True
    cmdDeleteEntry.Enabled = False
Else
    cmdCreateENtry.Enabled = False
    cmdDeleteEntry.Enabled = True
End If

StrSQL = "sELECT TblCarsData.Id as CarID,tblTripTrans3.*,TblCarsData.EqupName as CarName,  TblCarsData.fixedAssetid FROM tblTripTrans3 lEFT oUTER join TblCarsData On TblCarsData.id =tblTripTrans3.CarID  where TravID=" & val(Me.TxtSerial1(mIndex).Text)
loadgrid StrSQL, FG, True, False
ReLineGrid2
ErrTrap:

End Sub

Public Sub FiLLRec2()
    On Error GoTo ErrTrap
  Dim StrSQL As String
  
  If CboPaymentType1.Text = "" Then
    MsgBox "«œŒ· ÿ—Ìﬁ… «·œ›⁄ «Ê·«"
    Exit Sub
  End If
  If Me.DcboBox.Enabled And Me.DcboBox.Text = "" Then
        MsgBox "«œŒ· ÿ—Ìﬁ… «·’‰œÊﬁ"
        Exit Sub
  End If
  
If Me.DcbAccount.Enabled And Me.DcbAccount.Text = "" Then
        MsgBox "«œŒ· «·Õ”«»"
        Exit Sub
  End If
  
  If val(TxtTotalValue.Text) = 0 Then
            If SystemOptions.UserInterface = ArabicInterface Then
                MsgBox "Ì—ÃÏ «Œ Ì«— ﬁÌ„… Ê«Õœ… ⁄·Ï «·«ﬁ·"
            Else
                MsgBox "Please Select Value"
            End If
            Exit Sub
         End If
  
   Dim BeginTrans As Boolean
    Cn.BeginTrans
    BeginTrans = True
    If Me.TxtModFlg2(mIndex).Text = "E" Then
        Cn.Execute "delete tblTripTrans2 where TravID=" & val(Me.TxtSerial1(mIndex).Text)
        Cn.Execute "delete tblTripTrans3 where TravID=" & val(Me.TxtSerial1(mIndex).Text)
        

        
        StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & val(Me.TXTNoteID.Text)
        Cn.Execute StrSQL, , adExecuteNoRecords
    ElseIf TxtModFlg2(mIndex).Text = "N" Then
        AddNewRec
    End If
    If TxtNoteSerial1.Text = "" Then
        TxtNoteSerial1.Text = Voucher_coding(val(Me.dcBranch.BoundText), XPDtbTrans(mIndex).value, 76, 76)
    End If
    RsSavRec("NoteSerial1").value = IIf(Me.TxtNoteSerial1 <> "", val(TxtNoteSerial1.Text), Null)
   ' TxtSerial1(mIndex).Text = new_id(mTableName, "id", "")
    TxtSerial1(mIndex).Text = RsSavRec.Fields("ID").value

    'RsSavRec.Fields("CustID").value = IIf(DcCustmer.Text <> "", Trim(DcCustmer.BoundText), Null)
    'RsSavRec.Fields("LegalcourtsID").value = IIf(DcLegalcourts.Text <> "", Trim(DcLegalcourts.BoundText), Null)
    'RsSavRec.Fields("LegalIssuesID").value = IIf(DcLegalIssues.Text <> "", Trim(DcLegalIssues.BoundText), Null)
    
    RsSavRec("RecordDate").value = XPDtbTrans(mIndex).value
    RsSavRec("RecordDateH").value = XPDtbTransH(mIndex).value

    RsSavRec("Fromdate").value = FromDate.value
    RsSavRec("todate").value = toDate.value
    RsSavRec("Fromdateh").value = ToHijriDate(FromDate.value)
    RsSavRec("todateh").value = ToHijriDate(toDate.value)
    
    'RsSavRec("NoteSerial1").value = val(TxtNoteID)
    RsSavRec.Fields("PaymentType").value = val(CboPaymentType1.ListIndex)
    RsSavRec.Fields("BoxID").value = IIf(DcboBox.Text <> "", Trim(DcboBox.BoundText), Null)
    RsSavRec.Fields("AccountPaym").value = IIf(DcbAccount.Text <> "", Trim(DcbAccount.BoundText), Null)
    RsSavRec.Fields("BranchId").value = IIf(dcBranch.Text <> "", Trim(dcBranch.BoundText), Null)
    
    
    

    RsSavRec.update
    
    Dim RsDetails1 As ADODB.Recordset
    
       Set RsDetails1 = New ADODB.Recordset
 'DB_CreateField "TblTravDueKDet", "Price", adCurrency, adColNullable, , , , False, True
   StrSQL = "SELECT  *  from dbo.tblTripTrans2 Where (1 = -1)"
   RsDetails1.Open StrSQL, Cn, adOpenKeyset, adLockOptimistic, adCmdText
      Dim i As Integer
      
    With Me.GridInstallments
'Selected
        For i = 1 To .Rows - 1
        If .Cell(flexcpChecked, i, .ColIndex("Select")) = flexChecked Then
         RsDetails1.AddNew
         RsDetails1("TypeTrans").value = 0
         RsDetails1("TravID").value = val(Me.TxtSerial1(mIndex).Text)
         RsDetails1("ShipID").value = val(.TextMatrix(i, .ColIndex("ShipID")))
        
         RsDetails1("notesallid").value = val(.TextMatrix(i, .ColIndex("NoteIDA")))
         RsDetails1("TripNo").value = (.TextMatrix(i, .ColIndex("TripNo")))
         RsDetails1("TripDate").value = (.TextMatrix(i, .ColIndex("TripDate")))
         RsDetails1("BranchID").value = val(.TextMatrix(i, .ColIndex("BranchID")))
         RsDetails1("CardNO").value = (.TextMatrix(i, .ColIndex("CardNO")))
         RsDetails1("QtyDownload").value = val(.TextMatrix(i, .ColIndex("QtyDownload")))
         RsDetails1("CardNO2").value = (.TextMatrix(i, .ColIndex("CardNO2")))
         RsDetails1("QtyDischarge").value = val(.TextMatrix(i, .ColIndex("QtyDischarge")))
         RsDetails1("FromID").value = val(.TextMatrix(i, .ColIndex("FromID")))
         RsDetails1("ToID").value = val(.TextMatrix(i, .ColIndex("ToID")))
         RsDetails1("CarTypeID").value = val(.TextMatrix(i, .ColIndex("CarTypeID")))
         RsDetails1("CarID").value = val(.TextMatrix(i, .ColIndex("CarID")))
         RsDetails1("CusID").value = val(.TextMatrix(i, .ColIndex("CusID")))
         RsDetails1("CarType1").value = val(.TextMatrix(i, .ColIndex("CarType1")))
         RsDetails1("Price").value = val(.TextMatrix(i, .ColIndex("Value")))
        '  RsDetails1("RecNo").value = val(.TextMatrix(i, .ColIndex("RecNo")))
         RsDetails1("Weight").value = val(.TextMatrix(i, .ColIndex("Weight")))
         RsDetails1("TotalValue").value = val(.TextMatrix(i, .ColIndex("TotalValue")))
         RsDetails1("Remarks").value = (.TextMatrix(i, .ColIndex("Remarks")))
         RsDetails1("NoteID").value = val(.TextMatrix(i, .ColIndex("NoteID")))
         RsDetails1("LeaderName").value = (.TextMatrix(i, .ColIndex("EmpName")))
         RsDetails1.update
        ' Cn.Execute " update  TblTripTypesTransport set  allocations=1 where ID=" & val(.TextMatrix(i, .ColIndex("NoteID")))
         Else
       
       ' Cn.Execute " update  TblTripTypesTransport set  allocations=null where ID=" & val(.TextMatrix(i, .ColIndex("NoteID")))
        End If
           Next i
        RsDetails1.Close
    End With
    StrSQL = "sELECT * FROM tblTripTrans3 "
   
    saveGrid StrSQL, FG, "CarID", "", "TravID", val(Me.TxtSerial1(mIndex).Text)
    
    
    
    
   If TxtNoteSerial.Text = "" Then     'ÃœÌœ ›ﬁÿ
                        If Notes_coding(val(my_branch), Me.XPDtbTrans(mIndex).value) = "error" Then
                            MsgBox " ·« Ì„ﬂ‰ «÷«›… ﬁÌÊœ ·Â–… «·⁄„·Ì… ·«‰ﬂ  ⁄œÌ  «·Õœ «·«ﬁ’Ì ··ﬁÌÊœ ﬂ„« Õœœ  ›Ì  —ﬁÌ„ «·”‰œ«  ": Exit Sub
                        Else
                                       
                                        If Notes_coding(val(my_branch), XPDtbTrans(mIndex).value) = "" Then
                                            MsgBox " ·«Ì„ﬂ‰ «‰‘«¡ «·ﬁÌœ ·Â–« «·„” ‰œ ·«‰ﬂ Õœœ   —ﬁÌ„ ﬁÌÊœ ÌœÊÌ  ": Exit Sub
                                        Else
                                             
                                        End If
                        End If
 End If
 Dim Account_Code_dynamic As String
 '  Account_Code_dynamic = get_account_code_branch(2, my_branch)
 '
 '           If Account_Code_dynamic = "NO branch" Then
 '               If SystemOptions.UserInterface = ArabicInterface Then
 '                   MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
 '               Else
 '                   MsgBox "Branch Not Created", vbCritical
 '               End If
'
'                GoTo ErrTrap
'            Else
'
'                If Account_Code_dynamic = "NO account" Then
'                    If SystemOptions.UserInterface = ArabicInterface Then
'                        MsgBox "·„ Ì „  ÕœÌœ Õ”«»  «·„»Ì⁄«   ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
'                    Else
'                        MsgBox "Sales Account Not Defined in this Branch", vbCritical
'                    End If
'
'                    GoTo ErrTrap
'
'                End If
'            End If
         
         'DcbTypeTransport.BoundText = 1
        
       
         
    Dim TxtNoteSerial1str As String

   

ll:
    Cn.CommitTrans
    BeginTrans = False
    
    MsgBox " „  ⁄„·Ì… «·Õ›Ÿ »‰Ã«Õ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    
    TxtModFlg2(mIndex) = "R"
    FiLLTXT2
    Exit Sub
ErrTrap:

    If RsSavRec.EditMode <> adEditNone Then
        RsSavRec.CancelUpdate
    End If

End Sub



Function createVoucher()
Dim NoteID As Long
Dim NoteDate As Date
Dim NoteSerial As String
Dim Notevalue As Double
Dim des As String
Dim TxtRemarks As TextBox
des = "«À»«  «” Õﬁ«ﬁ «·—Õ·«  ⁄‰ «·› —… „‰  " '& Fromdate√H.value & "  Õ Ï  " & TodateH.value & Chr(13)
des = des & " „‰ " & FromDate.value & "  Õ Ï  " & toDate.value & CHR(13)
'des = des & " ··⁄„Ì· " & DBCboClientName2.Text & CHR(13)
'des = des & " ‰Ê⁄ «·‰ﬁ· " & DcbTypeTransport.Text & CHR(13)
'des = des & " «·’‰› " & DcboItems.Text & CHR(13)
des = des & " ··›—⁄ " & dcBranch.Text ' & "     " & txtRemarks

Dim tablename As String
Dim Filedname As String
Dim ContNo As Long
Dim sql As String
tablename = "tblTripTrans"
Filedname = "ID"
ContNo = TxtSerial1(mIndex)
Notevalue = 0

Notevalue = Format(val(TxtTotalValue.Text), "#.##")

If Me.TxtModFlg2(mIndex) = "N" Or TxtNoteSerial.Text = "" Then

    
    CreateNotes NoteID, (XPDtbTrans(mIndex).value), val(dcBranch.BoundText), 9095, Notevalue, NoteSerial, Me.TxtNoteSerial1, tablename, Filedname, ContNo, des, XPDtbTransH(mIndex).value
    TXTNoteID.Text = NoteID
    TxtNoteSerial.Text = NoteSerial
Else
    sql = "update notes  set Note_Value=" & Notevalue & ",note_value_by_characters='" & WriteNo(val(Notevalue), 0, True) & "'"
    sql = sql & ",NoteSerial1='" & Me.TxtNoteSerial1 & "',remark='" & des & "'"
    sql = sql & " where NoteID=" & val(TXTNoteID.Text)
    Cn.Execute sql
End If

CREATE_VOUCHER_GE val(TXTNoteID.Text), val(dcBranch.BoundText), user_id, XPDtbTrans(mIndex).value

RsSavRec.Resync adAffectCurrent
End Function


Public Function CREATE_VOUCHER_GE(general_noteid As Long, BranchID As Integer, UserID As Long _
, NoteDate As Date)
 Dim Notevalue As Single
    Dim LngDevID As Long
    Dim LngDevNO  As Integer
    Dim StrTempAccountCode As String
    Dim StrTempDes As String
    Dim actiondesdes As String
    Dim SngTemp  As Variant
    Dim Account_Code_dynamic As String
    Dim i As Long
 
 Dim StrSQL As String
 
         StrSQL = "Delete From DOUBLE_ENTREY_VOUCHERS Where Notes_ID=" & general_noteid
        Cn.Execute StrSQL, , adExecuteNoRecords
        
 LngDevNO = 0

    LngDevID = new_id("DOUBLE_ENTREY_VOUCHERS", "Double_Entry_Vouchers_ID", "")
    
    
   
    my_branch = BranchID

                                   '«·ÿ—› «·„Ì‰
    StrTempAccountCode = get_account_code_branch(70, my_branch)
    
    
    Dim mCarId As Double, mfixedAssetid As Integer, mBoardNO As String
    Dim mDesc As String, mEmpName As String
    For i = 1 To GridInstallments.Rows - 1
        mCarId = val(GridInstallments.TextMatrix(i, GridInstallments.ColIndex("CarID")))
        mDesc = Trim(GridInstallments.TextMatrix(i, GridInstallments.ColIndex("Remarks")))
        mEmpName = Trim(GridInstallments.TextMatrix(i, GridInstallments.ColIndex("EmpName")))
        Notevalue = Round(val(GridInstallments.TextMatrix(i, GridInstallments.ColIndex("TotalValue"))), 3)
        mfixedAssetid = val(GridInstallments.TextMatrix(i, GridInstallments.ColIndex("fixedAssetid")))
        mBoardNO = Trim(GridInstallments.TextMatrix(i, GridInstallments.ColIndex("BoardNO")))
                                
        If Notevalue > 0 Then
            LngDevNO = LngDevNO + 1
            actiondesdes = " Õ”«» „ﬂ«›¬  «·”«∆ﬁÌ‰" & CHR(13) & mEmpName & CHR(13) & " ··”Ì«—… " & mBoardNO
            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 0, StrTempDes & actiondesdes, general_noteid, , , , NoteDate, UserID, , , , , , , , mDesc, , , , , , , mfixedAssetid, , , BranchID, mCarId) = False Then
                GoTo ErrTrap
            End If
        End If
    Next i
'
'                                   If val(TxtVAT.Text) > 0 Then
'
'                                           Notevalue = Round(TxtVAT.Text, 3)
'                                           GetValueAddedAccount XPDtbTrans.value, , StrTempAccountCode, 1, 21
'                                           LngDevNO = LngDevNO + 1
'
'                                           actiondesdes = "Õ”«»  «·ﬁÌ„… «·„÷«›… „»Ì⁄«  " & CHR(13) & TxtRemarks.Text
'                                                    If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 1, StrTempDes & actiondesdes, general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
'                                                                 GoTo ErrTrap
'                                                     End If
'                                        End If
                                        
    If Round(txtTotal2.Text, 3) > 0 Then
    
        For i = 1 To FG.Rows - 1
            LngDevNO = LngDevNO + 1
            Notevalue = val(txtTotal2.Text)
            mCarId = val(FG.TextMatrix(i, FG.ColIndex("CarID")))
            mDesc = Trim(FG.TextMatrix(i, FG.ColIndex("Remarks")))
            Notevalue = Round(val(FG.TextMatrix(i, FG.ColIndex("Price"))), 3)
            mfixedAssetid = val(FG.TextMatrix(i, FG.ColIndex("fixedAssetid")))
            mBoardNO = Trim(FG.TextMatrix(i, FG.ColIndex("BoardNO")))
            ' StrTempAccountCode = get_account_code_branch(69, my_branch)
            'GetAccountTypeTrans val(DcbTypeTransport.BoundText), StrTempAccountCode
            actiondesdes = "Õ”«» „’—Ê› «·œÌ“· " & CHR(13) & mDesc & CHR(13) & " ··”Ì«—… " & mBoardNO
            StrTempAccountCode = get_account_code_branch(69, my_branch)
            If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 0, StrTempDes & actiondesdes, general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , mfixedAssetid, , , BranchID, mCarId) = False Then
                GoTo ErrTrap
            End If
        Next
    End If
                                  
    If Round(val(txtTotal2.Text) + val(TxtTotalValue), 3) > 0 Then
        LngDevNO = LngDevNO + 1
        Notevalue = Round(val(txtTotal2.Text) + val(TxtTotalValue), 3)
    
        StrTempAccountCode = DcboCreditSide.BoundText
    
        actiondesdes = "Õ”«» «·⁄Âœ…/«·Œ“‰… " & CHR(13) '& txtRemarks.Text
    If ModAccounts.AddNewDev(LngDevID, LngDevNO, StrTempAccountCode, Notevalue, 1, StrTempDes & actiondesdes, general_noteid, , , , NoteDate, UserID, , , , , , , , , , , , , , , , , , BranchID) = False Then
    GoTo ErrTrap
    End If
    End If
    

    updateNotesValueAndNobytext (general_noteid)
ErrTrap:
End Function

