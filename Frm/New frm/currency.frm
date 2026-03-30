VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FRMcurrency 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "»Ì«‰«  «·⁄„·« "
   ClientHeight    =   9945
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14250
   Icon            =   "currency.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   9945
   ScaleWidth      =   14250
   ShowInTaskbar   =   0   'False
   Begin C1SizerLibCtl.C1Tab TabMain 
      Height          =   9870
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   18870
      _cx             =   33285
      _cy             =   17410
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
      Caption         =   "√‰Ê«⁄ «·„Õ«ﬂ„|«·ﬁ÷«Ì«|»Ì«‰«  «·⁄„·« |»Ì«‰«  «·ﬁ÷Ì…| ”ÃÌ· „Ê«⁄Ìœ «·Ã·”« | ”ÃÌ· ”Ì— «·ﬁ÷Ì…"
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
         Height          =   9495
         Index           =   0
         Left            =   -19725
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   45
         Width           =   18780
         _cx             =   33126
         _cy             =   16748
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
            Height          =   600
            Index           =   1
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   13
            Top             =   0
            Width           =   18870
            Begin VB.Frame Frame2 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   375
               Index           =   1
               Left            =   540
               RightToLeft     =   -1  'True
               TabIndex        =   18
               Top             =   450
               Visible         =   0   'False
               Width           =   3105
               Begin MSDataListLib.DataCombo DataCombo4 
                  CausesValidation=   0   'False
                  Height          =   315
                  Index           =   1
                  Left            =   -255
                  TabIndex        =   19
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
                  TabIndex        =   20
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
               Index           =   0
               Left            =   2580
               RightToLeft     =   -1  'True
               TabIndex        =   17
               Text            =   "modflag"
               Top             =   90
               Visible         =   0   'False
               Width           =   465
            End
            Begin VB.TextBox Text21 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H000000FF&
               Height          =   240
               Index           =   1
               Left            =   3030
               RightToLeft     =   -1  'True
               TabIndex        =   16
               Top             =   510
               Visible         =   0   'False
               Width           =   945
            End
            Begin VB.TextBox TxtMenuState 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   5040
               RightToLeft     =   -1  'True
               TabIndex        =   15
               Text            =   "N"
               Top             =   780
               Visible         =   0   'False
               Width           =   735
            End
            Begin VB.TextBox TxtCutKey 
               Alignment       =   1  'Right Justify
               Height          =   375
               Left            =   5910
               RightToLeft     =   -1  'True
               TabIndex        =   14
               Top             =   750
               Visible         =   0   'False
               Width           =   615
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
                     Picture         =   "currency.frx":058A
                     Key             =   "CompanyName"
                  EndProperty
                  BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "currency.frx":0924
                     Key             =   "Ser"
                  EndProperty
                  BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "currency.frx":0CBE
                     Key             =   "Vac_Name"
                  EndProperty
                  BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "currency.frx":1058
                     Key             =   "ShareCount"
                  EndProperty
                  BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "currency.frx":13F2
                     Key             =   "Dis_Count"
                  EndProperty
                  BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "currency.frx":178C
                     Key             =   "Bouns"
                  EndProperty
                  BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "currency.frx":1B26
                     Key             =   "SharesValue"
                  EndProperty
                  BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "currency.frx":20C0
                     Key             =   "BuyValue"
                  EndProperty
               EndProperty
            End
            Begin ImpulseButton.ISButton btn_Last 
               Height          =   315
               Index           =   0
               Left            =   90
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
               ButtonImage     =   "currency.frx":245A
               ColorButton     =   14871017
               AcclimateGrayTones=   -1  'True
               DrawFocusRectangle=   0   'False
               DisabledImageExtraction=   0
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btn_Next 
               Height          =   315
               Index           =   0
               Left            =   555
               TabIndex        =   22
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
               ButtonImage     =   "currency.frx":27F4
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btn_Previous 
               Height          =   315
               Index           =   0
               Left            =   1155
               TabIndex        =   23
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
               ButtonImage     =   "currency.frx":2B8E
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btn_First 
               Height          =   315
               Index           =   0
               Left            =   1620
               TabIndex        =   24
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
               ButtonImage     =   "currency.frx":2F28
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«‰Ê«⁄ «·„Õ«ﬂ„"
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
               Left            =   4590
               RightToLeft     =   -1  'True
               TabIndex        =   25
               Top             =   90
               Width           =   2640
            End
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00E2E9E9&
            BorderStyle     =   0  'None
            Height          =   5415
            Index           =   1
            Left            =   -225
            RightToLeft     =   -1  'True
            TabIndex        =   5
            Top             =   1065
            Width           =   7605
            Begin VB.ComboBox Combo1 
               BackColor       =   &H80000018&
               Height          =   315
               ItemData        =   "currency.frx":32C2
               Left            =   2550
               List            =   "currency.frx":32D2
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   9
               Top             =   5460
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
               Index           =   0
               Left            =   3120
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   8
               Top             =   4065
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
               Index           =   0
               Left            =   285
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   7
               Top             =   4545
               Width           =   3870
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
               Index           =   0
               Left            =   285
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   6
               Top             =   5010
               Width           =   3870
            End
            Begin VSFlex8Ctl.VSFlexGrid Grid2 
               Height          =   3435
               Left            =   330
               TabIndex        =   122
               Top             =   60
               Width           =   7200
               _cx             =   12700
               _cy             =   6059
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
               FormatString    =   $"currency.frx":32EB
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
            Begin VB.Label Label6 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ﬂÊœ "
               Height          =   195
               Index           =   1
               Left            =   4695
               RightToLeft     =   -1  'True
               TabIndex        =   12
               Top             =   4170
               Width           =   990
            End
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ ⁄—»Ì"
               Height          =   285
               Index           =   1
               Left            =   4350
               RightToLeft     =   -1  'True
               TabIndex        =   11
               Top             =   4620
               Width           =   1350
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «‰Ã·Ì“Ì"
               Height          =   285
               Index           =   1
               Left            =   4200
               RightToLeft     =   -1  'True
               TabIndex        =   10
               Top             =   5130
               Width           =   1500
            End
         End
         Begin VB.CheckBox Chklast 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„Ã„Ê⁄Â ‰Â«∆Ì…"
            Height          =   225
            Left            =   135
            RightToLeft     =   -1  'True
            TabIndex        =   4
            Top             =   3600
            Visible         =   0   'False
            Width           =   1485
         End
         Begin VB.TextBox TxtGroupCode 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   3
            Top             =   0
            Visible         =   0   'False
            Width           =   630
         End
         Begin VSFlex8UCtl.VSFlexGrid FgItems 
            Height          =   9330
            Index           =   1
            Left            =   25935
            TabIndex        =   26
            Top             =   885
            Width           =   18600
            _cx             =   32808
            _cy             =   16457
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
            FormatString    =   $"currency.frx":337A
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
         Begin MSDataListLib.DataCombo DCboUserName 
            Height          =   315
            Index           =   2
            Left            =   315
            TabIndex        =   27
            Top             =   9210
            Visible         =   0   'False
            Width           =   2565
            _ExtentX        =   4524
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin ImpulseButton.ISButton btn_New 
            Height          =   360
            Index           =   0
            Left            =   6435
            TabIndex        =   28
            Top             =   7665
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   635
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
            ButtonImage     =   "currency.frx":343A
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Save 
            Height          =   360
            Index           =   0
            Left            =   4680
            TabIndex        =   29
            Top             =   7665
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   635
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
            ButtonImage     =   "currency.frx":37D4
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Modify 
            Height          =   360
            Index           =   0
            Left            =   5535
            TabIndex        =   30
            Top             =   7665
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   635
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
            ButtonImage     =   "currency.frx":3B6E
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton Btn_Undo 
            Height          =   360
            Index           =   0
            Left            =   3735
            TabIndex        =   31
            Top             =   7665
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   635
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
            ButtonImage     =   "currency.frx":3F08
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Delete 
            Height          =   360
            Index           =   0
            Left            =   2880
            TabIndex        =   32
            Top             =   7665
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   635
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
            ButtonImage     =   "currency.frx":42A2
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton Btn_Update 
            Height          =   405
            Index           =   0
            Left            =   5265
            TabIndex        =   33
            TabStop         =   0   'False
            ToolTipText     =   " ÕœÌÀ ﬁ«⁄œ… «·»Ì«‰« "
            Top             =   6810
            Visible         =   0   'False
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   714
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
            ButtonImage     =   "currency.frx":483C
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Cancel 
            Height          =   360
            Index           =   0
            Left            =   -180
            TabIndex        =   34
            Top             =   7665
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   635
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
            ButtonImage     =   "currency.frx":4BD6
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton Btn_Print 
            Height          =   555
            Index           =   0
            Left            =   1845
            TabIndex        =   35
            TabStop         =   0   'False
            ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
            Top             =   7575
            Width           =   990
            _ExtentX        =   1746
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
            ButtonImage     =   "currency.frx":4F70
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btn_Query 
            Height          =   615
            Index           =   0
            Left            =   540
            TabIndex        =   36
            TabStop         =   0   'False
            ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
            Top             =   7530
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   1085
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
            ButtonImage     =   "currency.frx":B7D2
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”Ã· «·Õ«·Ì:"
            Height          =   225
            Index           =   1
            Left            =   5220
            RightToLeft     =   -1  'True
            TabIndex        =   40
            Top             =   7305
            Width           =   1080
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «·”Ã·« :"
            Height          =   225
            Index           =   0
            Left            =   3240
            RightToLeft     =   -1  'True
            TabIndex        =   39
            Top             =   7305
            Width           =   1125
         End
         Begin VB.Label LabCurr_Rec 
            BackColor       =   &H00E2E9E9&
            Height          =   225
            Index           =   0
            Left            =   4365
            RightToLeft     =   -1  'True
            TabIndex        =   38
            Top             =   7335
            Width           =   765
         End
         Begin VB.Label LabCount_Rec 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   225
            Index           =   0
            Left            =   2520
            RightToLeft     =   -1  'True
            TabIndex        =   37
            Top             =   7335
            Width           =   585
         End
      End
      Begin C1SizerLibCtl.C1Elastic ELe 
         Height          =   9495
         Index           =   1
         Left            =   -19425
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   45
         Width           =   18780
         _cx             =   33126
         _cy             =   16748
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
            Height          =   645
            Index           =   0
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   50
            Top             =   0
            Width           =   18870
            Begin VB.TextBox Text22w 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H000000FF&
               Height          =   240
               Left            =   3030
               RightToLeft     =   -1  'True
               TabIndex        =   55
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
               Index           =   1
               Left            =   2580
               RightToLeft     =   -1  'True
               TabIndex        =   54
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
               Index           =   0
               Left            =   540
               RightToLeft     =   -1  'True
               TabIndex        =   51
               Top             =   450
               Visible         =   0   'False
               Width           =   3105
               Begin MSDataListLib.DataCombo DataCombo4 
                  CausesValidation=   0   'False
                  Height          =   315
                  Index           =   0
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
                  Index           =   4
                  Left            =   2160
                  RightToLeft     =   -1  'True
                  TabIndex        =   53
                  Top             =   45
                  Width           =   855
               End
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
                     Picture         =   "currency.frx":BB6C
                     Key             =   "CompanyName"
                  EndProperty
                  BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "currency.frx":BF06
                     Key             =   "Ser"
                  EndProperty
                  BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "currency.frx":C2A0
                     Key             =   "Vac_Name"
                  EndProperty
                  BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "currency.frx":C63A
                     Key             =   "ShareCount"
                  EndProperty
                  BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "currency.frx":C9D4
                     Key             =   "Dis_Count"
                  EndProperty
                  BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "currency.frx":CD6E
                     Key             =   "Bouns"
                  EndProperty
                  BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "currency.frx":D108
                     Key             =   "SharesValue"
                  EndProperty
                  BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "currency.frx":D6A2
                     Key             =   "BuyValue"
                  EndProperty
               EndProperty
            End
            Begin ImpulseButton.ISButton btn_Last 
               Height          =   315
               Index           =   1
               Left            =   90
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
               ButtonImage     =   "currency.frx":DA3C
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
               ButtonImage     =   "currency.frx":DDD6
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btn_Previous 
               Height          =   315
               Index           =   1
               Left            =   1155
               TabIndex        =   58
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
               ButtonImage     =   "currency.frx":E170
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btn_First 
               Height          =   315
               Index           =   1
               Left            =   1620
               TabIndex        =   59
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
               ButtonImage     =   "currency.frx":E50A
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "«‰Ê«⁄ «·ﬁ÷«Ì«"
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
               Left            =   4560
               RightToLeft     =   -1  'True
               TabIndex        =   60
               Top             =   60
               Width           =   2640
            End
         End
         Begin VB.Frame Frame1 
            BackColor       =   &H00E2E9E9&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   1740
            Index           =   2
            Left            =   270
            RightToLeft     =   -1  'True
            TabIndex        =   42
            Top             =   4170
            Width           =   6255
            Begin VB.ComboBox Combo2 
               BackColor       =   &H80000018&
               Height          =   315
               ItemData        =   "currency.frx":E8A4
               Left            =   2280
               List            =   "currency.frx":E8B4
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   46
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
               Left            =   3030
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   45
               Top             =   330
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
               Index           =   1
               Left            =   1395
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   44
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
               Index           =   1
               Left            =   1395
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   43
               Top             =   1020
               Width           =   2760
            End
            Begin VB.Label Label6 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·ﬂÊœ "
               Height          =   195
               Index           =   2
               Left            =   4695
               RightToLeft     =   -1  'True
               TabIndex        =   49
               Top             =   450
               Width           =   990
            End
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ ⁄—»Ì"
               Height          =   285
               Index           =   2
               Left            =   4350
               RightToLeft     =   -1  'True
               TabIndex        =   48
               Top             =   780
               Width           =   1350
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «‰Ã·Ì“Ì"
               Height          =   285
               Index           =   2
               Left            =   4200
               RightToLeft     =   -1  'True
               TabIndex        =   47
               Top             =   1140
               Width           =   1500
            End
         End
         Begin ImpulseButton.ISButton btn_New 
            Height          =   390
            Index           =   1
            Left            =   6660
            TabIndex        =   61
            Top             =   6720
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   688
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
            ButtonImage     =   "currency.frx":E8CD
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Save 
            Height          =   390
            Index           =   1
            Left            =   4905
            TabIndex        =   62
            Top             =   6720
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   688
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
            ButtonImage     =   "currency.frx":EC67
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Modify 
            Height          =   390
            Index           =   1
            Left            =   5760
            TabIndex        =   63
            Top             =   6720
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   688
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
            ButtonImage     =   "currency.frx":F001
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton Btn_Undo 
            Height          =   390
            Index           =   1
            Left            =   3960
            TabIndex        =   64
            Top             =   6720
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   688
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
            ButtonImage     =   "currency.frx":F39B
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Delete 
            Height          =   390
            Index           =   1
            Left            =   3105
            TabIndex        =   65
            Top             =   6720
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   688
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
            ButtonImage     =   "currency.frx":F735
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton Btn_Update 
            Height          =   240
            Index           =   1
            Left            =   5490
            TabIndex        =   66
            TabStop         =   0   'False
            ToolTipText     =   " ÕœÌÀ ﬁ«⁄œ… «·»Ì«‰« "
            Top             =   5910
            Visible         =   0   'False
            Width           =   1035
            _ExtentX        =   1826
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
            ButtonImage     =   "currency.frx":FCCF
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Cancel 
            Height          =   390
            Index           =   1
            Left            =   45
            TabIndex        =   67
            Top             =   6720
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   688
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
            ButtonImage     =   "currency.frx":10069
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton Btn_Print 
            Height          =   480
            Index           =   1
            Left            =   2070
            TabIndex        =   68
            TabStop         =   0   'False
            ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
            Top             =   6630
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   847
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
            ButtonImage     =   "currency.frx":10403
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btn_Query 
            Height          =   570
            Index           =   1
            Left            =   765
            TabIndex        =   69
            TabStop         =   0   'False
            ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
            Top             =   6570
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
            ButtonImage     =   "currency.frx":16C65
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin VSFlex8Ctl.VSFlexGrid Grid1 
            Height          =   3435
            Left            =   45
            TabIndex        =   70
            Top             =   765
            Width           =   7560
            _cx             =   13335
            _cy             =   6059
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
            FormatString    =   $"currency.frx":16FFF
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
            Left            =   2745
            RightToLeft     =   -1  'True
            TabIndex        =   74
            Top             =   6270
            Width           =   585
         End
         Begin VB.Label LabCurr_Rec 
            BackColor       =   &H00E2E9E9&
            Height          =   240
            Index           =   1
            Left            =   4590
            RightToLeft     =   -1  'True
            TabIndex        =   73
            Top             =   6270
            Width           =   765
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «·”Ã·« :"
            Height          =   225
            Index           =   2
            Left            =   3465
            RightToLeft     =   -1  'True
            TabIndex        =   72
            Top             =   6255
            Width           =   1125
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”Ã· «·Õ«·Ì:"
            Height          =   225
            Index           =   3
            Left            =   5445
            RightToLeft     =   -1  'True
            TabIndex        =   71
            Top             =   6255
            Width           =   1080
         End
      End
      Begin C1SizerLibCtl.C1Elastic ELe 
         Height          =   9495
         Index           =   2
         Left            =   45
         TabIndex        =   75
         TabStop         =   0   'False
         Top             =   45
         Width           =   18780
         _cx             =   33126
         _cy             =   16748
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
         Begin VB.CheckBox chkbasic 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   255
            Left            =   5175
            RightToLeft     =   -1  'True
            TabIndex        =   103
            Top             =   5520
            Width           =   1350
         End
         Begin VB.Frame FraHeader 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   540
            Left            =   45
            RightToLeft     =   -1  'True
            TabIndex        =   92
            Top             =   0
            Width           =   7785
            Begin VB.TextBox TxtVac_ID 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H000000FF&
               Height          =   240
               Left            =   3030
               RightToLeft     =   -1  'True
               TabIndex        =   97
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
               TabIndex        =   96
               Text            =   "modflag"
               Top             =   90
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
               TabIndex        =   93
               Top             =   450
               Visible         =   0   'False
               Width           =   3105
               Begin MSDataListLib.DataCombo DCUser 
                  CausesValidation=   0   'False
                  Height          =   315
                  Left            =   -255
                  TabIndex        =   94
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
                  TabIndex        =   95
                  Top             =   45
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
                     Picture         =   "currency.frx":1708E
                     Key             =   "CompanyName"
                  EndProperty
                  BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "currency.frx":17428
                     Key             =   "Ser"
                  EndProperty
                  BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "currency.frx":177C2
                     Key             =   "Vac_Name"
                  EndProperty
                  BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "currency.frx":17B5C
                     Key             =   "ShareCount"
                  EndProperty
                  BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "currency.frx":17EF6
                     Key             =   "Dis_Count"
                  EndProperty
                  BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "currency.frx":18290
                     Key             =   "Bouns"
                  EndProperty
                  BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "currency.frx":1862A
                     Key             =   "SharesValue"
                  EndProperty
                  BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "currency.frx":18BC4
                     Key             =   "BuyValue"
                  EndProperty
               EndProperty
            End
            Begin ImpulseButton.ISButton btnLast 
               Height          =   315
               Left            =   90
               TabIndex        =   98
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
               ButtonImage     =   "currency.frx":18F5E
               ColorButton     =   14871017
               AcclimateGrayTones=   -1  'True
               DrawFocusRectangle=   0   'False
               DisabledImageExtraction=   0
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnNext 
               Height          =   315
               Left            =   555
               TabIndex        =   99
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
               ButtonImage     =   "currency.frx":192F8
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnPrevious 
               Height          =   315
               Left            =   1155
               TabIndex        =   100
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
               ButtonImage     =   "currency.frx":19692
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnFirst 
               Height          =   315
               Left            =   1620
               TabIndex        =   101
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
               ButtonImage     =   "currency.frx":19A2C
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "»Ì«‰«  «·⁄„·« "
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
               Left            =   5295
               RightToLeft     =   -1  'True
               TabIndex        =   102
               Top             =   90
               Width           =   2280
            End
         End
         Begin VB.Frame Frm2 
            BackColor       =   &H00E2E9E9&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   1365
            Left            =   -1395
            RightToLeft     =   -1  'True
            TabIndex        =   76
            Top             =   4035
            Width           =   9180
            Begin VB.TextBox txtdivnamee 
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
               Left            =   1800
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   84
               Top             =   840
               Width           =   1080
            End
            Begin VB.TextBox txtdivname 
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
               Left            =   1800
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   83
               Top             =   480
               Width           =   1080
            End
            Begin VB.TextBox TxtVacNameE 
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
               Left            =   4320
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   82
               Top             =   840
               Width           =   3480
            End
            Begin VB.TextBox Text2 
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
               Left            =   1800
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   81
               Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· „⁄œ· «·’—›"
               Top             =   120
               Width           =   1080
            End
            Begin VB.TextBox Text1 
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
               Left            =   5520
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   80
               Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· ﬂÊœ «·⁄„·…"
               Top             =   120
               Width           =   2280
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
               Left            =   4320
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   79
               Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·⁄„·Â ⁄—»Ì"
               Top             =   525
               Width           =   3480
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
               Left            =   11160
               MaxLength       =   50
               RightToLeft     =   -1  'True
               TabIndex        =   78
               Top             =   45
               Width           =   1065
            End
            Begin VB.ComboBox CmbType 
               BackColor       =   &H80000018&
               Height          =   315
               ItemData        =   "currency.frx":19DC6
               Left            =   2280
               List            =   "currency.frx":19DD6
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   77
               Top             =   1590
               Visible         =   0   'False
               Width           =   1005
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " »«ﬁÌ «·⁄„·Â «‰Ã·Ì“Ì"
               Height          =   285
               Index           =   8
               Left            =   2880
               RightToLeft     =   -1  'True
               TabIndex        =   91
               Top             =   840
               Width           =   1410
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   " »«ﬁÌ «·⁄„·Â ⁄—»Ì"
               Height          =   285
               Index           =   7
               Left            =   2880
               RightToLeft     =   -1  'True
               TabIndex        =   90
               Top             =   480
               Width           =   1410
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·⁄„·Â «‰Ã·Ì“Ì"
               Height          =   285
               Index           =   6
               Left            =   7320
               RightToLeft     =   -1  'True
               TabIndex        =   89
               Top             =   870
               Width           =   1770
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„⁄œ· «·’—›"
               Height          =   285
               Index           =   3
               Left            =   3360
               RightToLeft     =   -1  'True
               TabIndex        =   88
               Top             =   120
               Width           =   930
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«”„ «·⁄„·Â ⁄—»Ì"
               Height          =   285
               Index           =   9
               Left            =   7920
               RightToLeft     =   -1  'True
               TabIndex        =   87
               Top             =   555
               Width           =   1170
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "ﬂÊœ «·⁄„·…  "
               Height          =   285
               Index           =   10
               Left            =   8220
               RightToLeft     =   -1  'True
               TabIndex        =   86
               Top             =   120
               Width           =   810
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "—ﬁ„ «·⁄„·…  "
               Height          =   195
               Index           =   11
               Left            =   12345
               RightToLeft     =   -1  'True
               TabIndex        =   85
               Top             =   150
               Width           =   990
            End
         End
         Begin C1SizerLibCtl.C1Elastic EltCont 
            Height          =   1020
            Left            =   90
            TabIndex        =   104
            TabStop         =   0   'False
            Top             =   5415
            Width           =   7785
            _cx             =   13732
            _cy             =   1799
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
               Left            =   5055
               TabIndex        =   105
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
               ButtonImage     =   "currency.frx":19DEF
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnSave 
               Height          =   330
               Left            =   3510
               TabIndex        =   106
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
               ButtonImage     =   "currency.frx":1A189
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnModify 
               Height          =   330
               Left            =   4275
               TabIndex        =   107
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
               ButtonImage     =   "currency.frx":1A523
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton BtnUndo 
               Height          =   330
               Left            =   2745
               TabIndex        =   108
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
               ButtonImage     =   "currency.frx":1A8BD
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnDelete 
               Height          =   330
               Left            =   1980
               TabIndex        =   109
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
               ButtonImage     =   "currency.frx":1AC57
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton btnQuery 
               Height          =   330
               Left            =   7320
               TabIndex        =   110
               TabStop         =   0   'False
               ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
               Top             =   690
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
               ButtonImage     =   "currency.frx":1B1F1
               ColorButton     =   14737632
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton BtnUpdate 
               Height          =   330
               Left            =   6045
               TabIndex        =   111
               TabStop         =   0   'False
               ToolTipText     =   " ÕœÌÀ ﬁ«⁄œ… «·»Ì«‰« "
               Top             =   585
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
               ButtonImage     =   "currency.frx":1B58B
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
            End
            Begin ImpulseButton.ISButton BtnPrint 
               Height          =   285
               Left            =   4725
               TabIndex        =   112
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
               ButtonImage     =   "currency.frx":1B925
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btnCancel 
               Height          =   330
               Left            =   1185
               TabIndex        =   113
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
               ButtonImage     =   "currency.frx":1BCBF
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·⁄„·Â «·—∆Ì”Ì…"
               Height          =   285
               Index           =   14
               Left            =   5880
               RightToLeft     =   -1  'True
               TabIndex        =   118
               Top             =   120
               Width           =   1650
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·”Ã· «·Õ«·Ì:"
               Height          =   210
               Index           =   4
               Left            =   2505
               RightToLeft     =   -1  'True
               TabIndex        =   117
               Top             =   225
               Width           =   975
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "⁄œœ «·”Ã·« :"
               Height          =   210
               Index           =   5
               Left            =   810
               RightToLeft     =   -1  'True
               TabIndex        =   116
               Top             =   225
               Width           =   975
            End
            Begin VB.Label LabCurrRec 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Height          =   210
               Left            =   1800
               RightToLeft     =   -1  'True
               TabIndex        =   115
               Top             =   240
               Width           =   675
            End
            Begin VB.Label LabCountRec 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Height          =   210
               Left            =   240
               RightToLeft     =   -1  'True
               TabIndex        =   114
               Top             =   225
               Width           =   540
            End
         End
         Begin VSFlex8Ctl.VSFlexGrid Grid 
            Height          =   3435
            Left            =   0
            TabIndex        =   119
            Top             =   570
            Width           =   7785
            _cx             =   13732
            _cy             =   6059
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
            Cols            =   6
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   320
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"currency.frx":1C059
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
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·”‰œ"
            Height          =   270
            Index           =   21
            Left            =   17250
            RightToLeft     =   -1  'True
            TabIndex        =   120
            Top             =   840
            Width           =   1170
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic12 
         Height          =   9495
         Index           =   0
         Left            =   19515
         TabIndex        =   123
         TabStop         =   0   'False
         Top             =   45
         Width           =   18780
         _cx             =   33126
         _cy             =   16748
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
         Begin VB.TextBox txtIssuesNo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   8070
            RightToLeft     =   -1  'True
            TabIndex        =   211
            Top             =   2340
            Width           =   1845
         End
         Begin VB.TextBox TxtIssuesDesc 
            Alignment       =   2  'Center
            Height          =   3990
            Left            =   60
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   207
            Top             =   3660
            Width           =   5235
         End
         Begin VB.TextBox TxtSerial1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Enabled         =   0   'False
            Height          =   390
            Index           =   3
            Left            =   11970
            RightToLeft     =   -1  'True
            TabIndex        =   200
            Top             =   1500
            Width           =   1095
         End
         Begin VB.TextBox TxtIssuesReason 
            Alignment       =   2  'Center
            Height          =   3720
            Left            =   8760
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   199
            Top             =   3690
            Width           =   5235
         End
         Begin VB.TextBox txtCustCode 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   5340
            TabIndex        =   198
            Top             =   1410
            Width           =   1545
         End
         Begin VB.Frame Fra_Header 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   645
            Index           =   2
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   125
            Top             =   0
            Width           =   18870
            Begin VB.Frame Frame2 
               BackColor       =   &H00E0E0E0&
               BorderStyle     =   0  'None
               Enabled         =   0   'False
               Height          =   375
               Index           =   2
               Left            =   540
               RightToLeft     =   -1  'True
               TabIndex        =   128
               Top             =   450
               Visible         =   0   'False
               Width           =   3105
               Begin MSDataListLib.DataCombo DataCombo4 
                  CausesValidation=   0   'False
                  Height          =   315
                  Index           =   2
                  Left            =   -255
                  TabIndex        =   129
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
                  Index           =   15
                  Left            =   2160
                  RightToLeft     =   -1  'True
                  TabIndex        =   130
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
               Index           =   3
               Left            =   2580
               RightToLeft     =   -1  'True
               TabIndex        =   127
               Text            =   "modflag"
               Top             =   90
               Visible         =   0   'False
               Width           =   465
            End
            Begin VB.TextBox Text3 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H000000FF&
               Height          =   240
               Left            =   3030
               RightToLeft     =   -1  'True
               TabIndex        =   126
               Top             =   510
               Visible         =   0   'False
               Width           =   945
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
                     Picture         =   "currency.frx":1C145
                     Key             =   "CompanyName"
                  EndProperty
                  BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "currency.frx":1C4DF
                     Key             =   "Ser"
                  EndProperty
                  BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "currency.frx":1C879
                     Key             =   "Vac_Name"
                  EndProperty
                  BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "currency.frx":1CC13
                     Key             =   "ShareCount"
                  EndProperty
                  BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "currency.frx":1CFAD
                     Key             =   "Dis_Count"
                  EndProperty
                  BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "currency.frx":1D347
                     Key             =   "Bouns"
                  EndProperty
                  BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "currency.frx":1D6E1
                     Key             =   "SharesValue"
                  EndProperty
                  BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "currency.frx":1DC7B
                     Key             =   "BuyValue"
                  EndProperty
               EndProperty
            End
            Begin ImpulseButton.ISButton btn_Last 
               Height          =   315
               Index           =   3
               Left            =   90
               TabIndex        =   131
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
               ButtonImage     =   "currency.frx":1E015
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
               TabIndex        =   132
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
               ButtonImage     =   "currency.frx":1E3AF
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btn_Previous 
               Height          =   315
               Index           =   3
               Left            =   1155
               TabIndex        =   133
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
               ButtonImage     =   "currency.frx":1E749
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btn_First 
               Height          =   315
               Index           =   3
               Left            =   1620
               TabIndex        =   134
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
               ButtonImage     =   "currency.frx":1EAE3
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "»Ì«‰«  «·ﬁ÷Ì…"
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
               Left            =   11340
               RightToLeft     =   -1  'True
               TabIndex        =   135
               Top             =   90
               Width           =   2640
            End
         End
         Begin ImpulseButton.ISButton btn_New 
            Height          =   360
            Index           =   3
            Left            =   6645
            TabIndex        =   136
            Top             =   9075
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   635
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
            ButtonImage     =   "currency.frx":1EE7D
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Save 
            Height          =   360
            Index           =   3
            Left            =   4890
            TabIndex        =   137
            Top             =   9075
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   635
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
            ButtonImage     =   "currency.frx":1F217
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Modify 
            Height          =   360
            Index           =   3
            Left            =   5745
            TabIndex        =   138
            Top             =   9075
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   635
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
            ButtonImage     =   "currency.frx":1F5B1
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton Btn_Undo 
            Height          =   360
            Index           =   3
            Left            =   3945
            TabIndex        =   139
            Top             =   9075
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   635
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
            ButtonImage     =   "currency.frx":1F94B
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Delete 
            Height          =   360
            Index           =   3
            Left            =   3090
            TabIndex        =   140
            Top             =   9075
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   635
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
            ButtonImage     =   "currency.frx":1FCE5
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton Btn_Update 
            Height          =   405
            Index           =   3
            Left            =   5475
            TabIndex        =   141
            TabStop         =   0   'False
            ToolTipText     =   " ÕœÌÀ ﬁ«⁄œ… «·»Ì«‰« "
            Top             =   8220
            Visible         =   0   'False
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   714
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
            ButtonImage     =   "currency.frx":2027F
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Cancel 
            Height          =   360
            Index           =   3
            Left            =   30
            TabIndex        =   142
            Top             =   9075
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   635
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
            ButtonImage     =   "currency.frx":20619
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton Btn_Print 
            Height          =   555
            Index           =   3
            Left            =   2055
            TabIndex        =   143
            TabStop         =   0   'False
            ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
            Top             =   8985
            Width           =   990
            _ExtentX        =   1746
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
            ButtonImage     =   "currency.frx":209B3
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btn_Query 
            Height          =   615
            Index           =   3
            Left            =   750
            TabIndex        =   144
            TabStop         =   0   'False
            ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
            Top             =   8940
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   1085
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
            ButtonImage     =   "currency.frx":27215
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin MSComCtl2.DTPicker XPDtbTrans 
            Height          =   315
            Index           =   3
            Left            =   8100
            TabIndex        =   201
            Top             =   1530
            Width           =   1860
            _ExtentX        =   3281
            _ExtentY        =   556
            _Version        =   393216
            Format          =   239206401
            CurrentDate     =   38784
         End
         Begin MSDataListLib.DataCombo DcCustmer 
            Height          =   315
            Left            =   645
            TabIndex        =   202
            Top             =   1410
            Width           =   4695
            _ExtentX        =   8281
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
         Begin Dynamic_Byte.NourHijriCal XPDtbTransH 
            Height          =   255
            Index           =   3
            Left            =   8100
            TabIndex        =   209
            Top             =   1890
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   450
         End
         Begin MSDataListLib.DataCombo DcLegalcourts 
            Bindings        =   "currency.frx":275AF
            Height          =   315
            Left            =   660
            TabIndex        =   213
            Top             =   1740
            Width           =   6225
            _ExtentX        =   10980
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
         Begin MSDataListLib.DataCombo DcLegalIssues 
            Bindings        =   "currency.frx":275C4
            Height          =   315
            Left            =   660
            TabIndex        =   215
            Top             =   2070
            Width           =   6225
            _ExtentX        =   10980
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
         Begin ImpulseButton.ISButton Cmd2 
            Height          =   375
            Left            =   7590
            TabIndex        =   252
            Top             =   9060
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
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·ﬁ÷Ì…"
            Height          =   360
            Index           =   5
            Left            =   6870
            TabIndex        =   216
            Top             =   2100
            Width           =   990
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·„Õﬂ„…"
            Height          =   360
            Index           =   3
            Left            =   6870
            TabIndex        =   214
            Top             =   1770
            Width           =   990
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·ﬁ÷Ì…"
            Height          =   270
            Index           =   1
            Left            =   9960
            TabIndex        =   212
            Top             =   2310
            Width           =   1740
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «·ﬁ÷Ì…  ÂÃ—Ì"
            Height          =   270
            Index           =   0
            Left            =   9930
            TabIndex        =   210
            Top             =   1860
            Width           =   1740
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "Ê’› «·ﬁ÷Ì…"
            Height          =   255
            Index           =   22
            Left            =   2250
            RightToLeft     =   -1  'True
            TabIndex        =   208
            Top             =   3330
            Width           =   1080
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «·ﬁ÷Ì…"
            Height          =   270
            Index           =   2
            Left            =   9930
            TabIndex        =   206
            Top             =   1560
            Width           =   1740
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "„"
            Height          =   255
            Index           =   4
            Left            =   13245
            RightToLeft     =   -1  'True
            TabIndex        =   205
            Top             =   1590
            Width           =   900
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "”»» «·ﬁ÷Ì…"
            Height          =   255
            Index           =   21
            Left            =   10950
            RightToLeft     =   -1  'True
            TabIndex        =   204
            Top             =   3360
            Width           =   1080
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·⁄„Ì·"
            Height          =   450
            Index           =   15
            Left            =   6795
            TabIndex        =   203
            Top             =   1440
            Width           =   1170
         End
         Begin VB.Label LabCount_Rec 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   225
            Index           =   3
            Left            =   2730
            RightToLeft     =   -1  'True
            TabIndex        =   148
            Top             =   8745
            Width           =   585
         End
         Begin VB.Label LabCurr_Rec 
            BackColor       =   &H00E2E9E9&
            Height          =   225
            Index           =   3
            Left            =   4575
            RightToLeft     =   -1  'True
            TabIndex        =   147
            Top             =   8745
            Width           =   765
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «·”Ã·« :"
            Height          =   225
            Index           =   7
            Left            =   3450
            RightToLeft     =   -1  'True
            TabIndex        =   146
            Top             =   8715
            Width           =   1125
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”Ã· «·Õ«·Ì:"
            Height          =   225
            Index           =   6
            Left            =   5430
            RightToLeft     =   -1  'True
            TabIndex        =   145
            Top             =   8715
            Width           =   1080
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic12 
         Height          =   9495
         Index           =   1
         Left            =   19815
         TabIndex        =   124
         TabStop         =   0   'False
         Top             =   45
         Width           =   18780
         _cx             =   33126
         _cy             =   16748
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
         Begin VB.TextBox TxtIssuesDesc2 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            Height          =   3990
            Left            =   870
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   253
            Top             =   3570
            Width           =   5235
         End
         Begin VB.TextBox txtIssuesNo2 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   7770
            RightToLeft     =   -1  'True
            TabIndex        =   228
            Top             =   840
            Width           =   1785
         End
         Begin VB.TextBox TxtSerial1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Enabled         =   0   'False
            Height          =   390
            Index           =   4
            Left            =   11535
            RightToLeft     =   -1  'True
            TabIndex        =   218
            Top             =   810
            Width           =   1095
         End
         Begin VB.TextBox txtSessionPlace 
            Alignment       =   2  'Center
            Height          =   3720
            Left            =   8325
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   217
            Top             =   3570
            Width           =   5235
         End
         Begin VB.Frame Fra_Header 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   645
            Index           =   3
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   162
            Top             =   0
            Width           =   18870
            Begin VB.TextBox Text4 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H000000FF&
               Height          =   240
               Left            =   3030
               RightToLeft     =   -1  'True
               TabIndex        =   167
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
               Index           =   4
               Left            =   2580
               RightToLeft     =   -1  'True
               TabIndex        =   166
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
               Index           =   3
               Left            =   540
               RightToLeft     =   -1  'True
               TabIndex        =   163
               Top             =   450
               Visible         =   0   'False
               Width           =   3105
               Begin MSDataListLib.DataCombo DataCombo4 
                  CausesValidation=   0   'False
                  Height          =   315
                  Index           =   3
                  Left            =   -255
                  TabIndex        =   164
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
                  Index           =   17
                  Left            =   2160
                  RightToLeft     =   -1  'True
                  TabIndex        =   165
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
                     Picture         =   "currency.frx":275D9
                     Key             =   "CompanyName"
                  EndProperty
                  BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "currency.frx":27973
                     Key             =   "Ser"
                  EndProperty
                  BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "currency.frx":27D0D
                     Key             =   "Vac_Name"
                  EndProperty
                  BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "currency.frx":280A7
                     Key             =   "ShareCount"
                  EndProperty
                  BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "currency.frx":28441
                     Key             =   "Dis_Count"
                  EndProperty
                  BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "currency.frx":287DB
                     Key             =   "Bouns"
                  EndProperty
                  BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "currency.frx":28B75
                     Key             =   "SharesValue"
                  EndProperty
                  BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "currency.frx":2910F
                     Key             =   "BuyValue"
                  EndProperty
               EndProperty
            End
            Begin ImpulseButton.ISButton btn_Last 
               Height          =   315
               Index           =   4
               Left            =   90
               TabIndex        =   168
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
               ButtonImage     =   "currency.frx":294A9
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
               TabIndex        =   169
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
               ButtonImage     =   "currency.frx":29843
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btn_Previous 
               Height          =   315
               Index           =   4
               Left            =   1155
               TabIndex        =   170
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
               ButtonImage     =   "currency.frx":29BDD
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btn_First 
               Height          =   315
               Index           =   4
               Left            =   1620
               TabIndex        =   171
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
               ButtonImage     =   "currency.frx":29F77
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   " ”ÃÌ· „Ê«⁄Ìœ «·Ã·”« "
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
               Left            =   11340
               RightToLeft     =   -1  'True
               TabIndex        =   172
               Top             =   90
               Width           =   2640
            End
         End
         Begin ImpulseButton.ISButton btn_New 
            Height          =   360
            Index           =   4
            Left            =   6615
            TabIndex        =   149
            Top             =   9045
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   635
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
            ButtonImage     =   "currency.frx":2A311
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Save 
            Height          =   360
            Index           =   4
            Left            =   4860
            TabIndex        =   150
            Top             =   9045
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   635
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
            ButtonImage     =   "currency.frx":2A6AB
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Modify 
            Height          =   360
            Index           =   4
            Left            =   5715
            TabIndex        =   151
            Top             =   9045
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   635
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
            ButtonImage     =   "currency.frx":2AA45
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton Btn_Undo 
            Height          =   360
            Index           =   4
            Left            =   3915
            TabIndex        =   152
            Top             =   9045
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   635
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
            ButtonImage     =   "currency.frx":2ADDF
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Delete 
            Height          =   360
            Index           =   4
            Left            =   3060
            TabIndex        =   153
            Top             =   9045
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   635
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
            ButtonImage     =   "currency.frx":2B179
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton Btn_Update 
            Height          =   405
            Index           =   4
            Left            =   5445
            TabIndex        =   154
            TabStop         =   0   'False
            ToolTipText     =   " ÕœÌÀ ﬁ«⁄œ… «·»Ì«‰« "
            Top             =   8190
            Visible         =   0   'False
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   714
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
            ButtonImage     =   "currency.frx":2B713
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Cancel 
            Height          =   360
            Index           =   4
            Left            =   0
            TabIndex        =   155
            Top             =   9045
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   635
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
            ButtonImage     =   "currency.frx":2BAAD
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton Btn_Print 
            Height          =   555
            Index           =   4
            Left            =   2025
            TabIndex        =   156
            TabStop         =   0   'False
            ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
            Top             =   8955
            Width           =   990
            _ExtentX        =   1746
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
            ButtonImage     =   "currency.frx":2BE47
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btn_Query 
            Height          =   615
            Index           =   4
            Left            =   720
            TabIndex        =   157
            TabStop         =   0   'False
            ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
            Top             =   8910
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   1085
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
            ButtonImage     =   "currency.frx":326A9
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin MSComCtl2.DTPicker XPDtbTrans 
            Height          =   315
            Index           =   4
            Left            =   7665
            TabIndex        =   219
            Top             =   1410
            Width           =   1860
            _ExtentX        =   3281
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   239140865
            CurrentDate     =   38784
         End
         Begin VB.PictureBox Picture1 
            Height          =   0
            Left            =   0
            ScaleHeight     =   0
            ScaleWidth      =   0
            TabIndex        =   223
            Top             =   0
            Width           =   0
         End
         Begin VB.PictureBox Picture2 
            Height          =   0
            Left            =   0
            ScaleHeight     =   0
            ScaleWidth      =   0
            TabIndex        =   224
            Top             =   0
            Width           =   0
         End
         Begin Dynamic_Byte.NourHijriCal XPDtbTransH 
            Height          =   255
            Index           =   4
            Left            =   7680
            TabIndex        =   225
            Top             =   1830
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   450
         End
         Begin MSDataListLib.DataCombo cmbLegalIssuesData 
            Bindings        =   "currency.frx":32A43
            Height          =   315
            Left            =   2190
            TabIndex        =   227
            Top             =   2370
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
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
         Begin MSComCtl2.DTPicker txtSessionDate 
            Height          =   315
            Left            =   4200
            TabIndex        =   244
            Top             =   1470
            Width           =   1860
            _ExtentX        =   3281
            _ExtentY        =   556
            _Version        =   393216
            Format          =   239140865
            CurrentDate     =   38784
         End
         Begin Dynamic_Byte.NourHijriCal txtSessionDateH 
            Height          =   255
            Left            =   4215
            TabIndex        =   246
            Top             =   1890
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   450
         End
         Begin MSComCtl2.DTPicker txtSessionTime 
            Height          =   285
            Left            =   210
            TabIndex        =   248
            Top             =   1470
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   503
            _Version        =   393216
            Format          =   239140866
            CurrentDate     =   38784
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "Ê’› «·ﬁ÷Ì…"
            Height          =   255
            Index           =   25
            Left            =   3060
            RightToLeft     =   -1  'True
            TabIndex        =   254
            Top             =   3270
            Width           =   1080
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·ﬁ÷Ì…"
            Height          =   270
            Index           =   20
            Left            =   6240
            TabIndex        =   251
            Top             =   2370
            Width           =   1020
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·ﬁ÷Ì…"
            Height          =   270
            Index           =   19
            Left            =   0
            TabIndex        =   250
            Top             =   0
            Width           =   1020
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "«· ÊﬁÌ "
            Height          =   270
            Index           =   18
            Left            =   2295
            TabIndex        =   249
            Top             =   1470
            Width           =   1530
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «·Ã·”… ÂÃ—Ì"
            Height          =   270
            Index           =   17
            Left            =   6045
            TabIndex        =   247
            Top             =   1860
            Width           =   1740
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «·Ã·”…"
            Height          =   270
            Index           =   6
            Left            =   6030
            TabIndex        =   245
            Top             =   1500
            Width           =   1740
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·ﬁ÷Ì…"
            Height          =   270
            Index           =   10
            Left            =   9810
            TabIndex        =   229
            Top             =   870
            Width           =   1020
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «·ﬁ÷Ì…  ÂÃ—Ì"
            Height          =   270
            Index           =   9
            Left            =   9510
            TabIndex        =   226
            Top             =   1800
            Width           =   1740
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «·ﬁ÷Ì…"
            Height          =   270
            Index           =   8
            Left            =   9495
            TabIndex        =   222
            Top             =   1440
            Width           =   1740
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "„"
            Height          =   255
            Index           =   7
            Left            =   12810
            RightToLeft     =   -1  'True
            TabIndex        =   221
            Top             =   900
            Width           =   900
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "„ﬂ«‰ «·Ã·”…"
            Height          =   255
            Index           =   23
            Left            =   10515
            RightToLeft     =   -1  'True
            TabIndex        =   220
            Top             =   3240
            Width           =   1080
         End
         Begin VB.Label LabCount_Rec 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   225
            Index           =   4
            Left            =   2700
            RightToLeft     =   -1  'True
            TabIndex        =   161
            Top             =   8715
            Width           =   585
         End
         Begin VB.Label LabCurr_Rec 
            BackColor       =   &H00E2E9E9&
            Height          =   225
            Index           =   4
            Left            =   4545
            RightToLeft     =   -1  'True
            TabIndex        =   160
            Top             =   8715
            Width           =   765
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «·”Ã·« :"
            Height          =   225
            Index           =   9
            Left            =   3420
            RightToLeft     =   -1  'True
            TabIndex        =   159
            Top             =   8685
            Width           =   1125
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”Ã· «·Õ«·Ì:"
            Height          =   225
            Index           =   8
            Left            =   5400
            RightToLeft     =   -1  'True
            TabIndex        =   158
            Top             =   8685
            Width           =   1080
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic12 
         Height          =   9495
         Index           =   2
         Left            =   20115
         TabIndex        =   173
         TabStop         =   0   'False
         Top             =   45
         Width           =   18780
         _cx             =   33126
         _cy             =   16748
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
         Begin VB.TextBox TxtIssuesDesc3 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            Height          =   3990
            Left            =   870
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   255
            Top             =   3270
            Width           =   5235
         End
         Begin VB.TextBox txtSessionDateNo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   4170
            RightToLeft     =   -1  'True
            TabIndex        =   243
            Top             =   1650
            Width           =   1845
         End
         Begin VB.TextBox txtIssuesNo3 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   7500
            RightToLeft     =   -1  'True
            TabIndex        =   240
            Top             =   810
            Width           =   1845
         End
         Begin VB.TextBox TxtSerial1 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Enabled         =   0   'False
            Height          =   390
            Index           =   5
            Left            =   11385
            RightToLeft     =   -1  'True
            TabIndex        =   231
            Top             =   1200
            Width           =   1095
         End
         Begin VB.TextBox txtSessionResult 
            Alignment       =   2  'Center
            Height          =   3720
            Left            =   8175
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   230
            Top             =   3390
            Width           =   5235
         End
         Begin VB.Frame Fra_Header 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   645
            Index           =   4
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   174
            Top             =   0
            Width           =   18870
            Begin VB.TextBox Text5 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H000000FF&
               Height          =   240
               Left            =   3030
               RightToLeft     =   -1  'True
               TabIndex        =   179
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
               Index           =   5
               Left            =   2580
               RightToLeft     =   -1  'True
               TabIndex        =   178
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
               TabIndex        =   175
               Top             =   450
               Visible         =   0   'False
               Width           =   3105
               Begin MSDataListLib.DataCombo DataCombo4 
                  CausesValidation=   0   'False
                  Height          =   315
                  Index           =   4
                  Left            =   -255
                  TabIndex        =   176
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
                  TabIndex        =   177
                  Top             =   45
                  Width           =   855
               End
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
                     Picture         =   "currency.frx":32A58
                     Key             =   "CompanyName"
                  EndProperty
                  BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "currency.frx":32DF2
                     Key             =   "Ser"
                  EndProperty
                  BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "currency.frx":3318C
                     Key             =   "Vac_Name"
                  EndProperty
                  BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "currency.frx":33526
                     Key             =   "ShareCount"
                  EndProperty
                  BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "currency.frx":338C0
                     Key             =   "Dis_Count"
                  EndProperty
                  BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "currency.frx":33C5A
                     Key             =   "Bouns"
                  EndProperty
                  BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "currency.frx":33FF4
                     Key             =   "SharesValue"
                  EndProperty
                  BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "currency.frx":3458E
                     Key             =   "BuyValue"
                  EndProperty
               EndProperty
            End
            Begin ImpulseButton.ISButton btn_Last 
               Height          =   315
               Index           =   5
               Left            =   90
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
               ButtonImage     =   "currency.frx":34928
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
               ButtonImage     =   "currency.frx":34CC2
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btn_Previous 
               Height          =   315
               Index           =   5
               Left            =   1155
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
               ButtonImage     =   "currency.frx":3505C
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin ImpulseButton.ISButton btn_First 
               Height          =   315
               Index           =   5
               Left            =   1620
               TabIndex        =   183
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
               ButtonImage     =   "currency.frx":353F6
               ColorButton     =   14871017
               DrawFocusRectangle=   0   'False
               DisabledImageStyle=   1
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   " ”ÃÌ· ”Ì— «·ﬁ÷Ì…"
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
               TabIndex        =   184
               Top             =   90
               Width           =   2640
            End
         End
         Begin ImpulseButton.ISButton btn_New 
            Height          =   360
            Index           =   5
            Left            =   6645
            TabIndex        =   185
            Top             =   9075
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   635
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
            ButtonImage     =   "currency.frx":35790
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Save 
            Height          =   360
            Index           =   5
            Left            =   4890
            TabIndex        =   186
            Top             =   9075
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   635
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
            ButtonImage     =   "currency.frx":35B2A
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Modify 
            Height          =   360
            Index           =   5
            Left            =   5745
            TabIndex        =   187
            Top             =   9075
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   635
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
            ButtonImage     =   "currency.frx":35EC4
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton Btn_Undo 
            Height          =   360
            Index           =   5
            Left            =   3945
            TabIndex        =   188
            Top             =   9075
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   635
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
            ButtonImage     =   "currency.frx":3625E
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Delete 
            Height          =   360
            Index           =   5
            Left            =   3090
            TabIndex        =   189
            Top             =   9075
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   635
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
            ButtonImage     =   "currency.frx":365F8
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton Btn_Update 
            Height          =   405
            Index           =   5
            Left            =   5475
            TabIndex        =   190
            TabStop         =   0   'False
            ToolTipText     =   " ÕœÌÀ ﬁ«⁄œ… «·»Ì«‰« "
            Top             =   8220
            Visible         =   0   'False
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   714
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
            ButtonImage     =   "currency.frx":36B92
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin ImpulseButton.ISButton btn_Cancel 
            Height          =   360
            Index           =   5
            Left            =   30
            TabIndex        =   191
            Top             =   9075
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   635
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
            ButtonImage     =   "currency.frx":36F2C
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton Btn_Print 
            Height          =   555
            Index           =   5
            Left            =   2055
            TabIndex        =   192
            TabStop         =   0   'False
            ToolTipText     =   "ÿ»«⁄… «·»Ì«‰«  "
            Top             =   8985
            Width           =   990
            _ExtentX        =   1746
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
            ButtonImage     =   "currency.frx":372C6
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
            DisabledImageStyle=   1
         End
         Begin ImpulseButton.ISButton btn_Query 
            Height          =   615
            Index           =   5
            Left            =   750
            TabIndex        =   193
            TabStop         =   0   'False
            ToolTipText     =   "(Ctrl+F)  ··»ÕÀ ≈÷€ÿ Â–« «·„› «Õ √Ê ≈÷€ÿ "
            Top             =   8940
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   1085
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
            ButtonImage     =   "currency.frx":3DB28
            ColorButton     =   14871017
            DrawFocusRectangle=   0   'False
         End
         Begin MSComCtl2.DTPicker XPDtbTrans 
            Height          =   315
            Index           =   5
            Left            =   7515
            TabIndex        =   232
            Top             =   1230
            Width           =   1860
            _ExtentX        =   3281
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   215220225
            CurrentDate     =   38784
         End
         Begin Dynamic_Byte.NourHijriCal XPDtbTransH 
            Height          =   255
            Index           =   5
            Left            =   7530
            TabIndex        =   236
            Top             =   1650
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   450
         End
         Begin MSDataListLib.DataCombo cmbLegalIssuesData3 
            Bindings        =   "currency.frx":3DEC2
            Height          =   315
            Left            =   180
            TabIndex        =   238
            Top             =   1260
            Width           =   3945
            _ExtentX        =   6959
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
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
         Begin MSDataListLib.DataCombo cmbSessionDate 
            Bindings        =   "currency.frx":3DED7
            Height          =   315
            Left            =   420
            TabIndex        =   242
            Top             =   8400
            Visible         =   0   'False
            Width           =   3945
            _ExtentX        =   6959
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
         Begin MSComCtl2.DTPicker txtSessionDate2 
            Height          =   315
            Left            =   180
            TabIndex        =   258
            Top             =   1620
            Width           =   1860
            _ExtentX        =   3281
            _ExtentY        =   556
            _Version        =   393216
            Format          =   215285761
            CurrentDate     =   38784
         End
         Begin Dynamic_Byte.NourHijriCal txtSessionDateH2 
            Height          =   255
            Left            =   195
            TabIndex        =   259
            Top             =   2040
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   450
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «·Ã·”…"
            Height          =   270
            Index           =   24
            Left            =   2010
            TabIndex        =   261
            Top             =   1650
            Width           =   1740
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «·Ã·”… ÂÃ—Ì"
            Height          =   270
            Index           =   23
            Left            =   2025
            TabIndex        =   260
            Top             =   2010
            Width           =   1740
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "‰Ê⁄ «·ﬁ÷Ì…"
            Height          =   270
            Index           =   22
            Left            =   4260
            TabIndex        =   257
            Top             =   1260
            Width           =   1020
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "Ê’› «·ﬁ÷Ì…"
            Height          =   255
            Index           =   26
            Left            =   3060
            RightToLeft     =   -1  'True
            TabIndex        =   256
            Top             =   2970
            Width           =   1080
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·ﬁ÷Ì…"
            Height          =   270
            Index           =   16
            Left            =   9450
            TabIndex        =   241
            Top             =   840
            Width           =   1020
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·Ã·”…"
            Height          =   360
            Index           =   14
            Left            =   6150
            TabIndex        =   239
            Top             =   1680
            Width           =   990
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «·ﬁ÷Ì…  ÂÃ—Ì"
            Height          =   270
            Index           =   13
            Left            =   9360
            TabIndex        =   237
            Top             =   1620
            Width           =   1740
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   " «—ÌŒ «·ﬁ÷Ì…"
            Height          =   270
            Index           =   12
            Left            =   9345
            TabIndex        =   235
            Top             =   1260
            Width           =   1740
         End
         Begin VB.Label lbl 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            Caption         =   "„"
            Height          =   255
            Index           =   11
            Left            =   12660
            RightToLeft     =   -1  'True
            TabIndex        =   234
            Top             =   1290
            Width           =   900
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "‰ ÌÃ… «·Ã·”…"
            Height          =   255
            Index           =   24
            Left            =   10365
            RightToLeft     =   -1  'True
            TabIndex        =   233
            Top             =   3060
            Width           =   1080
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·”Ã· «·Õ«·Ì:"
            Height          =   225
            Index           =   11
            Left            =   5430
            RightToLeft     =   -1  'True
            TabIndex        =   197
            Top             =   8715
            Width           =   1080
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "⁄œœ «·”Ã·« :"
            Height          =   225
            Index           =   10
            Left            =   3450
            RightToLeft     =   -1  'True
            TabIndex        =   196
            Top             =   8715
            Width           =   1125
         End
         Begin VB.Label LabCurr_Rec 
            BackColor       =   &H00E2E9E9&
            Height          =   225
            Index           =   5
            Left            =   4575
            RightToLeft     =   -1  'True
            TabIndex        =   195
            Top             =   8745
            Width           =   765
         End
         Begin VB.Label LabCount_Rec 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   225
            Index           =   5
            Left            =   2730
            RightToLeft     =   -1  'True
            TabIndex        =   194
            Top             =   8745
            Width           =   585
         End
      End
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Label4"
      Height          =   195
      Left            =   4500
      RightToLeft     =   -1  'True
      TabIndex        =   121
      Top             =   9930
      Width           =   435
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Label1"
      Height          =   255
      Index           =   1
      Left            =   1980
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   780
      Width           =   375
   End
End
Attribute VB_Name = "FRMcurrency"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsSavRec As ADODB.Recordset
Dim BKGrndPic As ClsBackGroundPic
Dim RecId As String
Dim ii As Long
Dim cSearchDcbo As clsDCboSearch
Dim mTableName As String
Public mIndex As Integer


Private Sub Dcbranch_Change(Index As Integer)
    If Me.TxtModFlg <> "R" Then
'    TxtNoteSerial1.Text = ""
'   TxtNoteSerial.Text = ""
   End If
End Sub

Private Sub Dcbranch_Click(Index As Integer, Area As Integer)
    If Me.TxtModFlg <> "R" Then
'    TxtNoteSerial1.Text = ""
'   TxtNoteSerial.Text = ""
   End If
   
End Sub

Private Sub cmbLegalIssuesData_Change()
cmbLegalIssuesData_Click 0
End Sub

Private Sub cmbLegalIssuesData_Click(Area As Integer)
 If val(cmbLegalIssuesData.BoundText) = 0 Then Exit Sub
  
  '  txtIssuesNo2 = cmbLegalIssuesData.BoundText

End Sub



Private Sub cmbLegalIssuesData3_Change()
cmbLegalIssuesData3_Click 0
End Sub

Private Sub cmbLegalIssuesData3_Click(Area As Integer)
 If val(cmbLegalIssuesData3.BoundText) = 0 Then Exit Sub
  
    txtIssuesNo3 = cmbLegalIssuesData3.BoundText

End Sub


Private Sub cmbSessionDate_Change()
cmbSessionDate_Click 0
End Sub

Private Sub cmbSessionDate_Click(Area As Integer)
 If val(cmbSessionDate.BoundText) = 0 Then Exit Sub
  
    txtSessionDateNo = cmbSessionDate.BoundText

End Sub





Private Sub Cmd2_Click()
    
            On Error Resume Next
                  If DoPremis(Do_Attach, Me.Name, True) = False Then
                Exit Sub
            End If
            
ShowAttachments TxtSerial1(mIndex).text, "09872201401"

End Sub

Private Sub DcCustmer_Change()
DcCustmer_Click 0
End Sub

Private Sub DcCustmer_Click(Area As Integer)
  If val(DcCustmer.BoundText) = 0 Then Exit Sub
  

    Dim EmpCode  As String
    GetTblCustemersCode , , DcCustmer.BoundText, EmpCode
    'Me.txtTotalValue.Text = EmpCode
    TxtCustCode = EmpCode
If Me.TxtModFlg2(mIndex).text <> "R" Then
If val(DcCustmer.BoundText) <> 0 Then

'GetInformationCustomer (DcCustmer.BoundText)

End If
End If
End Sub

Private Sub DcCustmer_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF3 Then
        Dim Frm As New FrmCustemerSearch
        Frm.SearchType = 604
        Frm.RetrunType = 0
        Frm.mIndex = mIndex
        Frm.show vbModal
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
           

            If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox " „  ⁄„·Ì… «·Õ–› »‰Ã«Õ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
            Else
            MsgBox "Delete Successfully", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
            End If
            '------------------------------ Move Next ---------------------------.
            
            If mIndex = 1 Then
                FiLLTXT1
                FillGridWithData1
            ElseIf mIndex = 0 Then
                FiLLTXT2
                FillGridWithData2
            ElseIf mIndex = 3 Then
                FiLLTXT3
            ElseIf mIndex = 4 Then
                FiLLTXT4
            ElseIf mIndex = 5 Then
                FiLLTXT5
                
                
                
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
    If mIndex = 0 Then
        FiLLTXT2
    ElseIf mIndex = 1 Then
        FiLLTXT1
    ElseIf mIndex = 3 Then
        FiLLTXT3
    ElseIf mIndex = 4 Then
        FiLLTXT4
    ElseIf mIndex = 5 Then
        FiLLTXT5
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
    Frame1(1).Enabled = True
    Frame1(2).Enabled = True
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
    Frame1(1).Enabled = True
    Frame1(2).Enabled = True
    clear_all Me
    TxtModFlg2(mIndex).text = "N"

    If mIndex = 0 Then
        My_SQL = "Legalcourts"
        'DCboUserName(mIndex) = user_id
                 
         clear_all Me
            
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable

    If rs.RecordCount > 0 Then
        TxtSerial1(mIndex).text = rs.RecordCount + 1
    Else
        TxtSerial1(mIndex).text = 1
    End If
    
    ElseIf mIndex = 1 Then
       
      '  DCboUserName(mIndex) = user_id
        My_SQL = "LegalIssues"
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
       
      '  DCboUserName(mIndex) = user_id
        My_SQL = "LegalIssuesData"
        'DCboUserName(mIndex) = user_id
     
         clear_all Me
    

   
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable

    If rs.RecordCount > 0 Then
        TxtSerial1(mIndex).text = rs.RecordCount + 1
    Else
        TxtSerial1(mIndex).text = 1
    End If

    rs.Close
      ElseIf mIndex = 4 Then
       
      '  DCboUserName(mIndex) = user_id
        My_SQL = "SessionDate"
        'DCboUserName(mIndex) = user_id
     
         clear_all Me
    

   
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable

    If rs.RecordCount > 0 Then
        TxtSerial1(mIndex).text = rs.RecordCount + 1
    Else
        TxtSerial1(mIndex).text = 1
    End If

    rs.Close
  ElseIf mIndex = 5 Then
       
      '  DCboUserName(mIndex) = user_id
        My_SQL = "LegalIssuesTrans"
        'DCboUserName(mIndex) = user_id
     
         clear_all Me
    

   
    rs.Open My_SQL, Cn, adOpenKeyset, adLockOptimistic, adCmdTable

    If rs.RecordCount > 0 Then
        TxtSerial1(mIndex).text = rs.RecordCount + 1
    Else
        TxtSerial1(mIndex).text = 1
    End If

    rs.Close
    ElseIf mIndex = 10 Then

        
    End If
    
    XPDtbTransH(mIndex).value = ToHijriDate(Date)
    txtSessionDateH.value = ToHijriDate(Date)
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
    ElseIf mIndex = 0 Then
        FiLLTXT2
 ElseIf mIndex = 3 Then
        FiLLTXT3
 ElseIf mIndex = 4 Then
        FiLLTXT4
 ElseIf mIndex = 5 Then
        FiLLTXT5
        
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
    ElseIf mIndex = 0 Then
        FiLLTXT2
    ElseIf mIndex = 3 Then
        FiLLTXT3
    ElseIf mIndex = 4 Then
        FiLLTXT4
    ElseIf mIndex = 5 Then
        FiLLTXT5

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


If mIndex = 0 Then
'     If Dcbranch(mIndex).Text = "" Then
'        MsgBox "Please Enter Branch"
'        Dcbranch(mIndex).SetFocus
'        Exit Sub
'    End If
End If
    
    '------------------------------ check if Empcode exist ----------------------

   If SystemOptions.ApplyEinvoice = True Then
        If Len(Text1) < 3 Then
        MsgBox "ﬂÊœ «·⁄„·… ÌÃ» «‰ ÌﬂÊ‰ 3 Œ«‰«  ÿ»ﬁ« ·„ ÿ·»«  «·«Ì“Ê "
        Exit Sub
        End If
End If


    ' -------------------------------------- txtmodflg type -------------------
    Select Case TxtModFlg2(mIndex).text

            '------------------------------ new record ----------------------------
        Case "N"
      
            '------------------------- save record -----------------------------
            If mIndex = 1 Then
                AddNewRec
                FiLLRec1
                
                FiLLTXT1
            ElseIf mIndex = 0 Then
                AddNewRec
                FiLLRec2
                FiLLTXT2
            ElseIf mIndex = 3 Then
                AddNewRec
                FiLLRec3
                FiLLTXT3
            ElseIf mIndex = 4 Then
                AddNewRec
                FiLLRec4
                FiLLTXT4
            ElseIf mIndex = 5 Then
                AddNewRec
                FiLLRec5
                FiLLTXT5
                
                
            End If
            

        Case "E"

            '----------------------------- save edit -------------------------------
            
            If mIndex = 0 Then
                FiLLRec1
            ElseIf mIndex = 1 Then
                FiLLRec2
            ElseIf mIndex = 3 Then
                FiLLRec3
            ElseIf mIndex = 4 Then
                FiLLRec4
            ElseIf mIndex = 5 Then
                FiLLRec5
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
            ElseIf mIndex = 0 Then
                FiLLTXT2
    
            ElseIf mIndex = 3 Then
                FiLLTXT3
            ElseIf mIndex = 4 Then
                FiLLTXT4
            ElseIf mIndex = 5 Then
                    FiLLTXT5
            End If
            TxtModFlg2(mIndex).text = "R"
    End Select
    
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

Private Sub Grid2_EnterCell()
    On Error GoTo ErrTrap
    FindRec val(Me.Grid2.TextMatrix(Me.Grid2.row, Me.Grid2.ColIndex("id")))
ErrTrap:
End Sub

Private Sub Grid1_EnterCell()
   '44
    On Error GoTo ErrTrap
    FindRec val(Me.GRID1.TextMatrix(Me.GRID1.row, Me.GRID1.ColIndex("id")))
ErrTrap:
End Sub


Private Sub TreeGroups_MouseUp(Button As Integer, _
                               Shift As Integer, _
                               X As Single, _
                               Y As Single)
    On Error GoTo ErrTrap
    Dim tp            As POINTAPI
    Dim lX            As Single
    Dim lY            As Single
    Dim tr            As RECT
    Dim XNodeSeelcted As MSComctlLib.Node

    'If Me.TreeGroups.SelectedItem Is Nothing Then
    '    Exit Sub
    'End If

    'TxtMenuState_Change
    If Button = vbRightButton Then
'        GetCursorPos tp
'        lX = (tp.X) * Screen.TwipsPerPixelX
'        lY = tp.Y * Screen.TwipsPerPixelY
'        XPPopUp.PopupMenu "mnuDropMenu1", lX, lY
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub TreeGroups_NodeClick(ByVal Node As MSComctlLib.Node)
    

End Sub

Private Sub TxtCustCode_KeyPress(KeyAscii As Integer)
    Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetTblCustemersCode TxtCustCode.text, EmpID
        DcCustmer.BoundText = EmpID
    End If
End Sub

Private Sub TxtCustCode_KeyUp(KeyCode As Integer, Shift As Integer)
'  If KeyCode = vbKeyF3 Then
'        Dim Frm As New FrmCustemerSearch
'        Frm.SearchType = 604
'        Frm.RetrunType = 0
'        Frm.mIndex = Index
'
'        Frm.show vbModal
'
'    End If
End Sub


Private Sub txtIssuesNo2_Change()
    If val(txtIssuesNo2) = 0 Then Exit Sub
    Dim s As String
    s = "Select * from LegalIssuesData Where IssuesNo = " & val(txtIssuesNo2)
    Dim rsDummy As New ADODB.Recordset
    rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly
    If rsDummy.EOF Then
    '    MsgBox "—ﬁ„ «·ﬁ÷Ì… €Ì— „”Ã· ›Ï »Ì«‰«  «·ﬁ÷«Ì«"
    '    txtIssuesNo2 = ""
        Exit Sub
    Else
        cmbLegalIssuesData.BoundText = val(rsDummy!LegalIssuesID & "")
        'val(txtIssuesNo2)
        TxtIssuesDesc2 = Trim(rsDummy!IssuesDesc & "")
        
        XPDtbTrans(4).value = IIf(IsNull(rsDummy("RecordDate").value), Date, rsDummy("RecordDate").value)
        
            
        XPDtbTransH(mIndex).value = ToHijriDate(XPDtbTrans(4).value)
    End If
    
    
End Sub

Private Sub txtIssuesNo2_Validate(Cancel As Boolean)
    If val(txtIssuesNo2) = 0 Then Exit Sub
    Dim s As String
    s = "Select * from LegalIssuesData Where IssuesNo = " & val(txtIssuesNo2)
    Dim rsDummy As New ADODB.Recordset
    rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly
    If rsDummy.EOF Then
        MsgBox "—ﬁ„ «·ﬁ÷Ì… €Ì— „”Ã· ›Ï »Ì«‰«  «·ﬁ÷«Ì«"
        txtIssuesNo2 = ""
        Exit Sub
    Else
        cmbLegalIssuesData.BoundText = val(rsDummy!LegalIssuesID & "")
        TxtIssuesDesc2 = Trim(rsDummy!IssuesDesc & "")
        XPDtbTrans(4).value = IIf(IsNull(rsDummy("RecordDate").value), Date, rsDummy("RecordDate").value)
        XPDtbTransH(mIndex).value = ToHijriDate(XPDtbTrans(4).value)
        'val(txtIssuesNo2)
    End If
End Sub

Private Sub txtIssuesNo3_Change()



cmbLegalIssuesData3.BoundText = val(txtIssuesNo3)


    If val(txtIssuesNo3) = 0 Then Exit Sub
    Dim s As String
    s = "Select * from LegalIssuesData Where IssuesNo = " & val(txtIssuesNo3)
    Dim rsDummy As New ADODB.Recordset
    rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly
    If rsDummy.EOF Then
    '    MsgBox "—ﬁ„ «·ﬁ÷Ì… €Ì— „”Ã· ›Ï »Ì«‰«  «·ﬁ÷«Ì«"
    '    txtIssuesNo2 = ""
        Exit Sub
    Else
        cmbLegalIssuesData3.BoundText = val(rsDummy!LegalIssuesID & "")
        'val(txtIssuesNo2)
        TxtIssuesDesc3 = Trim(rsDummy!IssuesDesc & "")
        
        XPDtbTrans(mIndex).value = IIf(IsNull(rsDummy("RecordDate").value), Date, rsDummy("RecordDate").value)
        
            
        XPDtbTransH(mIndex).value = ToHijriDate(XPDtbTrans(mIndex).value)
    End If

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
        '   btnNext.Enabled = False
        '   btnPrevious.Enabled = False
        '   btnFirst.Enabled = False
        '   btnLast.Enabled = False
    
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

Private Sub BtnCancel_Click()
    Unload Me
End Sub

Private Sub btnDelete_Click()
    Dim MSGType As Integer
    Dim BlnRecExist As Boolean
    Dim StrMSG  As String
    Dim Msg As String
    On Error GoTo ErrTrap

 If val(TxtVac_ID.text) = 1 Then
       Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã·...!!!"
           MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
           Exit Sub
 End If
 
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
        CuurentLogdata ("D")
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

    My_SQL = "currency"
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
    'On Error GoTo ErrTrap
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

    '------------------------------ check if currency exist ----------------------

    StrVacName = IsRecExist("currency", "name", Trim(TxtVacName.text), "name", "Vac_ID<>'" & Trim(TxtVac_ID.text) & "'")

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

Private Sub ChangeLang()
    Dim XPic As IPictureDisp
    Set XPic = Me.btnFirst.ButtonImage
    Set Me.btnFirst.ButtonImage = Me.btnLast.ButtonImage
    Set Me.btnLast.ButtonImage = XPic
    Set XPic = Me.btnPrevious.ButtonImage
    Set Me.btnPrevious.ButtonImage = Me.btnNext.ButtonImage
    Set Me.btnNext.ButtonImage = XPic

    Label1(3).Caption = "ID"
    Label1(1).Caption = "Arabic Name"
    Label1(6).Caption = "English Name"
    Label1(7).Caption = "Div Arabic  "
    Label1(8).Caption = "Div English  "
    Label1(5).Caption = "Default Currency  "
 
    Label1(0).Caption = "Code"
    Label1(4).Caption = "Rate"

    With Grid
        .TextMatrix(0, .ColIndex("ID")) = "  ID"
        .TextMatrix(0, .ColIndex("Code")) = "Code"
        .TextMatrix(0, .ColIndex("name")) = " Arabic Name"
        .TextMatrix(0, .ColIndex("namee")) = "English  Name"

        .TextMatrix(0, .ColIndex("rate")) = "Rate"
        Me.Caption = "Currency  Data"
        Label1(2).Caption = Me.Caption
        btnNew.Caption = "New"
        btnModify.Caption = "Modify"
        btnSave.Caption = "Save"
        BtnUndo.Caption = "Undo"
        btnDelete.Caption = "Delete"
        btnCancel.Caption = "Exit"
        Label2(0).Caption = "Current Record"
        Label2(1).Caption = "NO Of Record"
    End With

End Sub

Private Sub Form_Load()
Dim StrSQL As String
Dim rs As ADODB.Recordset

TabMain.TabVisible(1) = False
TabMain.TabVisible(2) = False
TabMain.TabVisible(0) = False
TabMain.TabVisible(3) = False
TabMain.TabVisible(4) = False
TabMain.TabVisible(5) = False
'TabMain.TabVisible(6) = False
If mIndex <> 2 Then
    TabMain.TabVisible(mIndex) = True
    TabMain.CurrTab = mIndex
End If
If mIndex = 0 Then mIndex = 10
If mIndex = 2 Then mIndex = 0


Select Case mIndex
Case 0
    mTableName = "Legalcourts"
    Me.Width = Grid2.Width + 400
    TabMain.TabVisible(0) = True
    TabMain.CurrTab = 0
    ' Me.Width = Grid.Width + 400
    ScreenNameArabic = " «‰Ê«⁄ «·„Õ«ﬂ„  "
    ScreenNameEnglish = " Items Class "
    FillGridWithData2
Case 1
    mTableName = "LegalIssues"
    TabMain.TabVisible(1) = True
    TabMain.CurrTab = 1
    Me.Width = GRID1.Width + 400
    ScreenNameArabic = " «‰Ê«⁄ «·ﬁ÷«Ì«"
    ScreenNameEnglish = " Items Class "
    
    FillGridWithData1
Case 2
    'mTableName = "Legalcourts"
Case 3
    mTableName = "LegalIssuesData"
Case 4
    mTableName = "SessionDate"
Case 5
    mTableName = "LegalIssuesTrans"
Case 10
    mTableName = "currency"
End Select
Me.Caption = ScreenNameArabic
Dim My_SQL As String
My_SQL = mTableName
Set BKGrndPic = New ClsBackGroundPic
Set RsSavRec = New ADODB.Recordset
RsSavRec.CursorLocation = adUseClient

RsSavRec.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdTableDirect
If mIndex <> 10 Then
    Me.TxtModFlg2(mIndex).text = "R"
    btn_First_Click (mIndex)
End If
       





     

      
    


    Dim Dcombos As New ClsDataCombos
    'Dcombos.GetCustomersSuppliers 1, Me.DcbSales, True
    Dcombos.GetCustomersSuppliers 1, Me.DcCustmer
    
    
      
        'Set cSearchDcbo = New clsDCboSearch
        'Set cSearchDcbo.Client = XPCboGroup
        'Dim Dcombos As New ClsDataCombos

       
        


   If mIndex = 10 Then
        TabMain.TabVisible(0) = False
        TabMain.TabVisible(2) = True
        TabMain.CurrTab = 2
       
       Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500
        Me.Width = Grid.Width + 400
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
    
    On Error GoTo ErrTrap
    LogTextA = "   «·œŒÊ· «·Ì ‘«‘… " & "«·⁄„·«  "
    LogTexte = " Open Window " & " Currency"
    AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, "O", "", ""

    Dim i As Integer
'    Dim My_SQL As String

    Me.TxtModFlg.text = "R"
    Resize_Form Me
    'load tblUsers -----------------------------------------------
    My_SQL = "select UserID,UserName From tblUsers "
    fill_combo DCUser, My_SQL

    FillGridWithData

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


   
    End If
  
    My_SQL = "select ID,Name From Legalcourts "
    fill_combo DcLegalcourts, My_SQL

    My_SQL = "select ID,Name From LegalIssues "
    fill_combo DcLegalIssues, My_SQL


    My_SQL = "select ID,Name From LegalIssues "
    fill_combo cmbLegalIssuesData, My_SQL

    My_SQL = "select ID,Name From LegalIssues "
    fill_combo cmbLegalIssuesData3, My_SQL


    My_SQL = "select ID,SessionPlace From SessionDate "
    fill_combo cmbSessionDate, My_SQL


Resize_Form Me

               'SetInterface Me
     RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish, "1"
'      If mIndex = 0 Then
'        Set rs = New ADODB.Recordset
'        rs.Open "LegalIssues", Cn, adOpenStatic, adLockOptimistic, adCmdTable
'        Me.TxtModFlg.Text = "R"
'        Retrive
'
'        If OPEN_NEW_SCREEN = True Then
'            Cmd_Click (0)
'            End If
'    End If

ErrTrap:
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)
 
'    Dim IntResult As String
'    Dim StrMSG As String
'    On Error GoTo ErrTrap
'
'    If Me.TxtModFlg.Text <> "R" Then
'
'        Select Case Me.TxtModFlg.Text
'
'            Case "N"
'
'                If SystemOptions.UserInterface = EnglishInterface Then
'                    StrMSG = "You will close this screen before save " & CHR(13)
'                    StrMSG = StrMSG & " the new data  " & CHR(13)
'                    StrMSG = StrMSG & " do you want save before exit" & CHR(13)
'                    StrMSG = StrMSG & "yes" & "-" & "save the new data" & CHR(13)
'                    StrMSG = StrMSG & "no" & "-" & "Don't save" & CHR(13)
'                    StrMSG = StrMSG & "cancel" & "-" & "Cancel Exit" & CHR(13)
'
'                Else
'                    StrMSG = "”Ê› Ì „ €·ﬁ «·‘«‘… Ê·„  ‰ Â „‰  ”ÃÌ·" & CHR(13)
'                    StrMSG = StrMSG & " «·»Ì«‰«  «·ÃœÌœ… «·Õ«·Ì… " & CHR(13)
'                    StrMSG = StrMSG & " Â·  —Ìœ «·Õ›Ÿ ﬁ»· «·Œ—ÊÃ" & CHR(13)
'                    StrMSG = StrMSG & "‰⁄„" & "-" & "Ì „ Õ›Ÿ «·»Ì«‰«  «·ÃœÌœ…" & CHR(13)
'                    StrMSG = StrMSG & "·«" & "-" & "·‰ Ì „ «·Õ›Ÿ" & CHR(13)
'                    StrMSG = StrMSG & "≈·€«¡ «·√„—" & "-" & "≈·€«¡ ⁄„·Ì… «·Œ—ÊÃ" & CHR(13)
'
'                End If
'
'            Case "E"
'
'                If SystemOptions.UserInterface = EnglishInterface Then
'                    StrMSG = "You will close this screen before save  " & CHR(13)
'                    StrMSG = StrMSG & " the Modifications  " & CHR(13)
'                    StrMSG = StrMSG & " do you want save before exit" & CHR(13)
'                    StrMSG = StrMSG & "yes" & "-" & "save the new data" & CHR(13)
'                    StrMSG = StrMSG & "no" & "-" & "Don't save" & CHR(13)
'                    StrMSG = StrMSG & "cancel" & "-" & "Cancel Exit" & CHR(13)
'
'                Else
'                    StrMSG = "”Ê› Ì „ €·ﬁ «·‘«‘… Ê·„  ‰ Â „‰  ”ÃÌ·" & CHR(13)
'                    StrMSG = StrMSG & " «· ⁄œÌ·«  «·ÃœÌœ… ⁄·Ï «·”Ã· «·Õ«·Ï " & CHR(13)
'                    StrMSG = StrMSG & " Â·  —Ìœ «·Õ›Ÿ ﬁ»· «·Œ—ÊÃ" & CHR(13)
'                    StrMSG = StrMSG & "‰⁄„" & "-" & "Ì „ Õ›Ÿ «· ⁄œÌ·«   «·ÃœÌœ…" & CHR(13)
'                    StrMSG = StrMSG & "·«" & "-" & "·‰ Ì „ «·Õ›Ÿ" & CHR(13)
'                    StrMSG = StrMSG & "≈·€«¡ «·√„—" & "-" & "≈·€«¡ ⁄„·Ì… «·Œ—ÊÃ" & CHR(13)
'
'                End If
'
'        End Select
'
'        IntResult = MsgBox(StrMSG, vbMsgBoxRight + vbYesNoCancel + vbMsgBoxRtlReading + vbQuestion, App.title)
'
'        Select Case IntResult
'
'            Case vbYes
'                Cancel = True
'                btnSave_Click
'
'            Case vbCancel
'                Cancel = True
'        End Select
'
'    End If
'
'    Exit Sub
'ErrTrap:



   Dim IntResult As String
    Dim StrMSG As String
    On Error GoTo ErrTrap
    Dim mSelectModFlg As String
    If mIndex = 2 Then
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
                If mIndex = 10 Then
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

Private Sub Form_Terminate()
    'Set cSearchDCombo = Nothing
    'Set BKGrndPic = Nothing
    Set FrmVacancy = Nothing

End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap
    LogTextA = "«·Œ—ÊÃ  ‘«‘… " & "«·⁄„·«  "
    LogTexte = " Exit Window " & " Currency"
    AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, "O", "", ""

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
   

    If mIndex = 10 Then
        StrRecID = new_id("currency", "id", "")
        RsSavRec.AddNew
        RsSavRec.Fields("id").value = IIf(StrRecID <> "", StrRecID, Null)
        FiLLRec
    ElseIf mIndex = 1 Then
    On Error GoTo ErrTrap
        
        StrRecID = new_id("LegalIssues", "id", "")
        RsSavRec.AddNew
        RsSavRec.Fields("id").value = IIf(StrRecID <> "", StrRecID, Null)
    ElseIf mIndex = 0 Then
    On Error GoTo ErrTrap
        
        StrRecID = new_id("Legalcourts", "id", "")
        RsSavRec.AddNew
        RsSavRec.Fields("id").value = IIf(StrRecID <> "", StrRecID, Null)
        
    ElseIf mIndex = 3 Or mIndex = 4 Or mIndex = 5 Then
    On Error GoTo ErrTrap
        
        StrRecID = new_id(mTableName, "id", "")
        RsSavRec.AddNew
        RsSavRec.Fields("id").value = IIf(StrRecID <> "", StrRecID, Null)
                
                    ElseIf mIndex = 0 Then
    On Error GoTo ErrTrap
        
        StrRecID = new_id("Legalcourts", "id", "")
        RsSavRec.AddNew
        RsSavRec.Fields("id").value = IIf(StrRecID <> "", StrRecID, Null)
        
       ' FiLLRec1
ErrTrap:

        
    End If
    
    
   
   
End Sub

Public Sub FiLLRec()
'    On Error GoTo ErrTrap
    Dim sql  As String
'    sql = "update   currency  set basic=0"
'    Cn.Execute sql

    If chkbasic.value = vbChecked Then
        RsSavRec.Fields("basic").value = 1
    Else
        RsSavRec.Fields("basic").value = 0
    End If

    RsSavRec.Fields("name").value = IIf(TxtVacName.text <> "", Trim(TxtVacName.text), Null)
    RsSavRec.Fields("namee").value = IIf(TxtVacNamee.text <> "", Trim(TxtVacNamee.text), Null)
    RsSavRec.Fields("code").value = IIf(Text1.text <> "", Trim(Text1.text), Null)
    RsSavRec.Fields("rate").value = IIf(Text2.text <> "", Trim(Text2.text), Null)

    RsSavRec.Fields("divname").value = IIf(txtdivname.text <> "", Trim(txtdivname.text), Null)
    RsSavRec.Fields("divnamee").value = IIf(txtdivnamee.text <> "", Trim(txtdivnamee.text), Null)

    RsSavRec.update
'    RsSavRec.Resync adAffectAllChapters
    MsgBox " „  ⁄„·Ì… «·Õ›Ÿ »‰Ã«Õ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    CuurentLogdata

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

    If RsSavRec.Fields("basic").value = True Then
        chkbasic.value = vbChecked
    Else
        chkbasic.value = Unchecked
    End If

    TxtVac_ID.text = IIf(IsNull(RsSavRec.Fields("id").value), "", RsSavRec.Fields("id").value)
    TxtVacName.text = IIf(IsNull(RsSavRec.Fields("name").value), "", RsSavRec.Fields("name").value)
    TxtVacNamee.text = IIf(IsNull(RsSavRec.Fields("namee").value), "", RsSavRec.Fields("namee").value)

    Text1.text = IIf(IsNull(RsSavRec.Fields("CODE").value), "", RsSavRec.Fields("CODE").value)
    Text2.text = IIf(IsNull(RsSavRec.Fields("RATE").value), "", RsSavRec.Fields("RATE").value)

    txtdivname.text = IIf(IsNull(RsSavRec.Fields("divname").value), "", RsSavRec.Fields("divname").value)
    txtdivnamee.text = IIf(IsNull(RsSavRec.Fields("divnamee").value), "", RsSavRec.Fields("divnamee").value)
    LabCurrRec.Caption = RsSavRec.AbsolutePosition
    LabCountRec.Caption = RsSavRec.RecordCount

    With Grid

        For i = 1 To .rows - 1

            If Trim(TxtVac_ID.text) = .TextMatrix(i, .ColIndex("id")) Then
                TxtSerial.text = .TextMatrix(i, .ColIndex("Ser"))
                .row = i
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
    FindRec val(Me.Grid.TextMatrix(Me.Grid.row, Me.Grid.ColIndex("id")))
ErrTrap:
End Sub

Private Sub txtdivname_GotFocus()
    SwitchKeyboardLang LANG_ARABIC
End Sub

Private Sub txtdivnamee_GotFocus()
    SwitchKeyboardLang LANG_ENGLISH
End Sub

Private Sub txtSessionDate_Change()
    txtSessionDateH.value = ToHijriDate(txtSessionDate.value)
End Sub

Private Sub txtSessionDateH_LostFocus()
              VBA.Calendar = vbCalGreg
            txtSessionDate.value = ToGregorianDate(txtSessionDateH.value)

End Sub

Private Sub txtSessionDateNo_Change()
' cmbSessionDate.BoundText = val(txtSessionDateNo)

 If val(txtIssuesNo3) = 0 Then Exit Sub
    Dim s As String
    s = "Select * from SessionDate Where ID = " & val(txtSessionDateNo)
    Dim rsDummy As New ADODB.Recordset
    rsDummy.Open s, Cn, adOpenStatic, adLockReadOnly
    If rsDummy.EOF Then
    '    MsgBox "—ﬁ„ «·ﬁ÷Ì… €Ì— „”Ã· ›Ï »Ì«‰«  «·ﬁ÷«Ì«"
    '    txtIssuesNo2 = ""
        Exit Sub
    Else
      '  cmbLegalIssuesData3.BoundText = val(rsDummy!LegalIssuesID & "")
        'val(txtIssuesNo2)
       
        
        txtSessionDate2.value = IIf(IsNull(rsDummy("SessionDate").value), Date, rsDummy("SessionDate").value)
        
            
       txtSessionDate2.value = ToHijriDate(txtSessionDate2.value)
    End If
End Sub

Private Sub TxtVac_ID_Change()
    Dim TxtMod As String
    TxtMod = TxtModFlg.text
    TxtModFlg.text = ""
    TxtModFlg = TxtMod
End Sub


Public Sub FiLLTXT3()

    On Error GoTo ErrTrap
    Dim i As Integer
    'Frame1(2).Enabled = False
    TxtSerial1(mIndex).text = IIf(IsNull(RsSavRec.Fields("id").value), "", RsSavRec.Fields("id").value)
    
    XPDtbTrans(mIndex).value = IIf(IsNull(RsSavRec("RecordDate").value), Date, RsSavRec("RecordDate").value)
    XPDtbTransH(mIndex).value = IIf(IsNull(RsSavRec("RecorddateH").value), ToHijriDate(Date), RsSavRec("RecorddateH").value)
    DcCustmer.BoundText = IIf(IsNull(RsSavRec("CustID").value), "", RsSavRec("CustID").value)

    TxtIssuesReason.text = IIf(IsNull(RsSavRec.Fields("IssuesReason").value), "", RsSavRec.Fields("IssuesReason").value)
    TxtIssuesDesc.text = IIf(IsNull(RsSavRec.Fields("IssuesDesc").value), "", RsSavRec.Fields("IssuesDesc").value)
    txtIssuesNo.text = IIf(IsNull(RsSavRec.Fields("IssuesNo").value), "", RsSavRec.Fields("IssuesNo").value)
    
    
    DcLegalcourts.BoundText = IIf(IsNull(RsSavRec("LegalcourtsID").value), "", RsSavRec("LegalcourtsID").value)
    DcLegalIssues.BoundText = IIf(IsNull(RsSavRec("LegalIssuesID").value), "", RsSavRec("LegalIssuesID").value)
    

    
    LabCurr_Rec(mIndex).Caption = RsSavRec.AbsolutePosition
    LabCount_Rec(mIndex).Caption = RsSavRec.RecordCount

   

ErrTrap:

End Sub

Public Sub FiLLTXT4()

    On Error GoTo ErrTrap
    Dim i As Integer
    'Frame1(2).Enabled = False
    TxtSerial1(mIndex).text = IIf(IsNull(RsSavRec.Fields("id").value), "", RsSavRec.Fields("id").value)
    
    XPDtbTrans(mIndex).value = IIf(IsNull(RsSavRec("RecordDate").value), Date, RsSavRec("RecordDate").value)
    XPDtbTransH(mIndex).value = IIf(IsNull(RsSavRec("RecorddateH").value), ToHijriDate(Date), RsSavRec("RecorddateH").value)
   
    txtSessionDate.value = IIf(IsNull(RsSavRec("SessionDate").value), Date, RsSavRec("SessionDate").value)
    XPDtbTransH(mIndex).value = IIf(IsNull(RsSavRec("SessionDateH").value), ToHijriDate(Date), RsSavRec("SessionDateH").value)
   
    
    
    txtSessionPlace.text = IIf(IsNull(RsSavRec.Fields("SessionPlace").value), "", RsSavRec.Fields("SessionPlace").value)
    txtIssuesNo2.text = IIf(IsNull(RsSavRec.Fields("IssuesNo").value), "", RsSavRec.Fields("IssuesNo").value)
    
    
    txtIssuesNo2_Change
    
      Dim startmaintenanceTime As Date
    If Not IsNull(RsSavRec("SessionTime").value) Then
         startmaintenanceTime = FormatDateTime(RsSavRec("SessionTime").value, vbShortTime)
         Me.txtSessionTime.value = startmaintenanceTime
    End If
    
    LabCurr_Rec(mIndex).Caption = RsSavRec.AbsolutePosition
    LabCount_Rec(mIndex).Caption = RsSavRec.RecordCount

   

ErrTrap:

End Sub


Public Sub FiLLTXT5()

    On Error GoTo ErrTrap
    Dim i As Integer
    'Frame1(2).Enabled = False
    TxtSerial1(mIndex).text = IIf(IsNull(RsSavRec.Fields("id").value), "", RsSavRec.Fields("id").value)
    
    XPDtbTrans(mIndex).value = IIf(IsNull(RsSavRec("RecordDate").value), Date, RsSavRec("RecordDate").value)
    XPDtbTransH(mIndex).value = IIf(IsNull(RsSavRec("RecorddateH").value), ToHijriDate(Date), RsSavRec("RecorddateH").value)
   
    'txtSessionDate.value = IIf(IsNull(RsSavRec("SessionDate").value), Date, RsSavRec("SessionDate").value)
    'XPDtbTransH(mIndex).value = IIf(IsNull(RsSavRec("SessionDateH").value), ToHijriDate(Date), RsSavRec("SessionDateH").value)
   
    
    
    txtSessionPlace.text = IIf(IsNull(RsSavRec.Fields("SessionPlace").value), "", RsSavRec.Fields("SessionPlace").value)
    txtIssuesNo3.text = IIf(IsNull(RsSavRec.Fields("IssuesNo").value), "", RsSavRec.Fields("IssuesNo").value)
    
    txtSessionDateNo = IIf(IsNull(RsSavRec.Fields("SessionDateNo").value), "", RsSavRec.Fields("SessionDateNo").value)
  
    
  
    LabCurr_Rec(mIndex).Caption = RsSavRec.AbsolutePosition
    LabCount_Rec(mIndex).Caption = RsSavRec.RecordCount

   

ErrTrap:

End Sub


Public Sub FiLLTXT1()

    On Error GoTo ErrTrap
    Dim i As Integer
    'Frame1(2).Enabled = False
    TxtSerial1(mIndex).text = IIf(IsNull(RsSavRec.Fields("id").value), "", RsSavRec.Fields("id").value)
    TxtName(mIndex).text = IIf(IsNull(RsSavRec.Fields("name").value), "", RsSavRec.Fields("name").value)
    TxtNameE(mIndex).text = IIf(IsNull(RsSavRec.Fields("nameE").value), "", RsSavRec.Fields("nameE").value)
    LabCurr_Rec(mIndex).Caption = RsSavRec.AbsolutePosition
    LabCount_Rec(mIndex).Caption = RsSavRec.RecordCount

    With GRID1

        For i = 1 To .rows - 1

            If Trim(TxtSerial1(mIndex).text) = .TextMatrix(i, .ColIndex("id")) Then
                TxtSerial1(mIndex).text = .TextMatrix(i, .ColIndex("Ser"))
                .row = i
                Exit Sub
            End If

        Next

    End With

ErrTrap:

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

Public Sub FiLLRec1()
    On Error GoTo ErrTrap

    RsSavRec.Fields("name").value = IIf(TxtName(mIndex).text <> "", Trim(TxtName(mIndex).text), Null)
    RsSavRec.Fields("namee").value = IIf(TxtNameE(mIndex).text <> "", Trim(TxtNameE(mIndex).text), Null)
    

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

Public Sub FiLLRec3()
    On Error GoTo ErrTrap

    TxtSerial1(mIndex).text = new_id(mTableName, "id", "")
    RsSavRec.Fields("ID").value = val(TxtSerial1(mIndex).text)

    RsSavRec.Fields("CustID").value = IIf(DcCustmer.text <> "", Trim(DcCustmer.BoundText), Null)
    RsSavRec.Fields("LegalcourtsID").value = IIf(DcLegalcourts.text <> "", Trim(DcLegalcourts.BoundText), Null)
    RsSavRec.Fields("LegalIssuesID").value = IIf(DcLegalIssues.text <> "", Trim(DcLegalIssues.BoundText), Null)
    
    RsSavRec("RecordDate").value = XPDtbTrans(mIndex).value
    RsSavRec("RecordDateH").value = XPDtbTransH(mIndex).value
  
    
    RsSavRec.Fields("IssuesNo").value = IIf(txtIssuesNo.text <> "", Trim(txtIssuesNo.text), Null)
    RsSavRec.Fields("IssuesReason").value = IIf(TxtIssuesReason.text <> "", Trim(TxtIssuesReason.text), Null)
    RsSavRec.Fields("IssuesDesc").value = IIf(TxtIssuesDesc.text <> "", Trim(TxtIssuesDesc.text), Null)
    

    RsSavRec.update
    MsgBox " „  ⁄„·Ì… «·Õ›Ÿ »‰Ã«Õ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    
    TxtModFlg2(mIndex) = "R"

    Exit Sub
ErrTrap:

    If RsSavRec.EditMode <> adEditNone Then
        RsSavRec.CancelUpdate
    End If

End Sub

Public Sub FiLLRec4()
    On Error GoTo ErrTrap

    TxtSerial1(mIndex).text = new_id(mTableName, "id", "")
    RsSavRec.Fields("ID").value = val(TxtSerial1(mIndex).text)

    'RsSavRec.Fields("CustID").value = IIf(DcCustmer.Text <> "", Trim(DcCustmer.BoundText), Null)
    'RsSavRec.Fields("LegalcourtsID").value = IIf(DcLegalcourts.Text <> "", Trim(DcLegalcourts.BoundText), Null)
    'RsSavRec.Fields("LegalIssuesID").value = IIf(DcLegalIssues.Text <> "", Trim(DcLegalIssues.BoundText), Null)
    
    RsSavRec("RecordDate").value = XPDtbTrans(mIndex).value
    RsSavRec("RecordDateH").value = XPDtbTransH(mIndex).value
    RsSavRec("SessionDate").value = txtSessionDate.value
    RsSavRec("SessionDateH").value = txtSessionDateH.value
    
    RsSavRec("SessionTime").value = txtSessionTime.value
    
    RsSavRec.Fields("IssuesNo").value = IIf(txtIssuesNo2.text <> "", Trim(txtIssuesNo2.text), Null)
    RsSavRec.Fields("SessionPlace").value = IIf(txtSessionPlace.text <> "", Trim(txtSessionPlace.text), Null)
    RsSavRec.Fields("IssuesDesc").value = IIf(TxtIssuesDesc.text <> "", Trim(TxtIssuesDesc.text), Null)
    

    RsSavRec.update
    MsgBox " „  ⁄„·Ì… «·Õ›Ÿ »‰Ã«Õ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    
    TxtModFlg2(mIndex) = "R"

    Exit Sub
ErrTrap:

    If RsSavRec.EditMode <> adEditNone Then
        RsSavRec.CancelUpdate
    End If

End Sub

Public Sub FiLLRec5()
    On Error GoTo ErrTrap

    TxtSerial1(mIndex).text = new_id(mTableName, "id", "")
    RsSavRec.Fields("ID").value = val(TxtSerial1(mIndex).text)

    'RsSavRec.Fields("CustID").value = IIf(DcCustmer.Text <> "", Trim(DcCustmer.BoundText), Null)
    'RsSavRec.Fields("LegalcourtsID").value = IIf(DcLegalcourts.Text <> "", Trim(DcLegalcourts.BoundText), Null)
    'RsSavRec.Fields("LegalIssuesID").value = IIf(DcLegalIssues.Text <> "", Trim(DcLegalIssues.BoundText), Null)
    
    RsSavRec("RecordDate").value = XPDtbTrans(mIndex).value
    RsSavRec("RecordDateH").value = XPDtbTransH(mIndex).value


    
    RsSavRec.Fields("IssuesNo").value = IIf(txtIssuesNo3.text <> "", Trim(txtIssuesNo3.text), Null)
     RsSavRec.Fields("SessionDateNo").value = IIf(txtSessionDateNo.text <> "", Trim(txtSessionDateNo.text), Null)
    
    RsSavRec.Fields("SessionResult").value = IIf(txtSessionResult.text <> "", Trim(txtSessionResult.text), Null)
    

    RsSavRec.update
    MsgBox " „  ⁄„·Ì… «·Õ›Ÿ »‰Ã«Õ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    
    TxtModFlg2(mIndex) = "R"

    Exit Sub
ErrTrap:

    If RsSavRec.EditMode <> adEditNone Then
        RsSavRec.CancelUpdate
    End If

End Sub

Public Sub FillGridWithData1()

    On Error GoTo ErrTrap

    Dim i As Integer
    Dim rs As ADODB.Recordset
    Dim My_SQL As String

    Set rs = New ADODB.Recordset
    My_SQL = "select * From LegalIssues order by id"
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
    My_SQL = "select * From Legalcourts order by id"
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


Public Sub FiLLTXT2()

    On Error GoTo ErrTrap
    Dim i As Integer
   ' Frame1(2).Enabled = False
    TxtSerial1(mIndex).text = IIf(IsNull(RsSavRec.Fields("id").value), "", RsSavRec.Fields("id").value)
    TxtName(mIndex).text = IIf(IsNull(RsSavRec.Fields("name").value), "", RsSavRec.Fields("name").value)
    TxtNameE(mIndex).text = IIf(IsNull(RsSavRec.Fields("nameE").value), "", RsSavRec.Fields("nameE").value)
    LabCurr_Rec(mIndex).Caption = RsSavRec.AbsolutePosition
    LabCount_Rec(mIndex).Caption = RsSavRec.RecordCount

    With GRID1

        For i = 1 To .rows - 1

            If Trim(TxtSerial1(mIndex).text) = .TextMatrix(i, .ColIndex("id")) Then
                TxtSerial1(mIndex).text = .TextMatrix(i, .ColIndex("Ser"))
                .row = i
                Exit Sub
            End If

        Next

    End With

ErrTrap:

End Sub

Public Function FindRec(ByVal RecId As Long, Optional ByVal mIndex2 As Integer = 0)
    On Error GoTo ErrTrap
    RsSavRec.Find "id=" & RecId, , adSearchForward, 1
    mIndex2 = mIndex
    If Not (RsSavRec.EOF) Then
        If mIndex2 = 10 Then
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
            


        End If
    End If

    Exit Function
ErrTrap:

    If RsSavRec.EditMode <> adEditNone Then
        RsSavRec.CancelUpdate
        If mIndex = 10 Then
            BtnUndo_Click
        Else
            Btn_Undo_Click (mIndex2)
       
        End If
        
        
    End If

    'RsSavRec.Filter = adFilterNone
End Function


Public Function FindRec6(ByVal RecId As Long)
    On Error GoTo ErrTrap
    RsSavRec.Find "id=" & RecId, , adSearchForward, 1

    If Not (RsSavRec.EOF) Then
        FiLLTXT
    
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

Private Sub TxtModFlg_Change2()

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
        '   btnNext.Enabled = False
        '   btnPrevious.Enabled = False
        '   btnFirst.Enabled = False
        '   btnLast.Enabled = False
    
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
    My_SQL = "select * From currency order by id"
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
           
                .TextMatrix(i, .ColIndex("CODE")) = IIf(IsNull(rs.Fields("CODE").value), "", rs.Fields("CODE").value)
            
                .TextMatrix(i, .ColIndex("RATE")) = IIf(IsNull(rs.Fields("RATE").value), "", rs.Fields("RATE").value)
            
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
            btnNew_Click
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

Function CuurentLogdata(Optional Currentmode As String)
     LogTextA = "  Õ›Ÿ ‘«‘… " & "     »Ì«‰«  «·⁄„·«   " & CHR(13) & " ﬂÊœ   " & Text1.text & CHR(13) & " «·«”„ " & TxtVacName.text & CHR(13) & " „⁄œ· «·’—›  " & Text22w & CHR(13) & "»«ﬁÌ «·⁄„·…  " & txtdivname
        LogTexte = " Save Window " & "    Currency Data " & CHR(13) & " ﬂÊœ   " & Text1.text & CHR(13) & " «·«”„ " & TxtVacNamee.text & CHR(13) & " „⁄œ· «·’—›  " & Text22w & CHR(13) & "»«ﬁÌ «·⁄„·…  " & txtdivnamee
       If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, Me.TxtModFlg, "", ""
    Else
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, "D", "", ""
    End If
                      
End Function

Private Sub TxtVacName_GotFocus()
    SwitchKeyboardLang LANG_ARABIC
End Sub

Private Sub TxtVacNamee_GotFocus()
    SwitchKeyboardLang LANG_ENGLISH
End Sub

Private Sub XPDtbTrans_Change(Index As Integer)
     XPDtbTransH(mIndex).value = ToHijriDate(XPDtbTrans(mIndex).value)
End Sub

Private Sub XPDtbTransH_LostFocus(Index As Integer)
              VBA.Calendar = vbCalGreg
            XPDtbTrans(mIndex).value = ToGregorianDate(XPDtbTransH(mIndex).value)

End Sub


  

