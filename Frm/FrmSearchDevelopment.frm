VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SuiteCtrls.ocx"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form FemSearchDevelopment 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11880
   Icon            =   "FrmSearchDevelopment.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6000
   ScaleWidth      =   11880
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   34
      Top             =   3960
      Visible         =   0   'False
      Width           =   11745
      Begin VB.TextBox XPTxtVal 
         Alignment       =   1  'Right Justify
         Height          =   405
         Left            =   630
         RightToLeft     =   -1  'True
         TabIndex        =   44
         Top             =   480
         Visible         =   0   'False
         Width           =   1965
      End
      Begin VB.ComboBox CboPayMentType 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4275
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   43
         Top             =   600
         Width           =   2190
      End
      Begin VB.TextBox txtPhoneCust 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   3300
         TabIndex        =   39
         Top             =   150
         Width           =   3210
      End
      Begin VB.TextBox txtCustName 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   7470
         TabIndex        =   35
         Top             =   180
         Width           =   3210
      End
      Begin MSDataListLib.DataCombo cmbLocationsName 
         Height          =   315
         Left            =   8115
         TabIndex        =   36
         Top             =   630
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo cmbCarName 
         Height          =   315
         Left            =   0
         TabIndex        =   41
         Top             =   120
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo cmbPaymentClass 
         Height          =   315
         Left            =   60
         TabIndex        =   47
         Top             =   990
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "›∆… «·”œ«œ"
         Height          =   285
         Index           =   2
         Left            =   2655
         TabIndex        =   48
         Top             =   990
         Width           =   945
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÿ—Ìﬁ… «·”œ«œ"
         Height          =   285
         Index           =   1
         Left            =   6495
         TabIndex        =   46
         Top             =   630
         Width           =   945
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·ﬁÌ„…"
         Height          =   285
         Index           =   4
         Left            =   2820
         TabIndex        =   45
         Top             =   540
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·„⁄œÂ/«·”Ì«—…"
         Height          =   285
         Index           =   0
         Left            =   2550
         TabIndex        =   42
         Top             =   150
         Width           =   645
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " ·Ì›Ê‰"
         Height          =   315
         Index           =   46
         Left            =   6480
         TabIndex        =   40
         Top             =   210
         Width           =   645
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·⁄„Ì·"
         Height          =   285
         Index           =   3
         Left            =   11010
         TabIndex        =   38
         Top             =   210
         Width           =   645
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·„Êﬁ⁄"
         Height          =   285
         Index           =   25
         Left            =   11010
         TabIndex        =   37
         Top             =   690
         Width           =   645
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   3720
      TabIndex        =   32
      Top             =   4080
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   3720
      TabIndex        =   31
      Top             =   4440
      Width           =   855
   End
   Begin VB.TextBox txtorder_no 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   8700
      TabIndex        =   25
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Frame Frame2 
      Caption         =   "«·√Ê·ÊÌÂ"
      Height          =   555
      Left            =   6000
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   5340
      Width           =   5775
      Begin XtremeSuiteControls.RadioButton Opt 
         Height          =   375
         Index           =   0
         Left            =   3360
         TabIndex        =   23
         Top             =   120
         Width           =   1815
         _Version        =   786432
         _ExtentX        =   3201
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "⁄«œÌ"
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton Opt 
         Height          =   375
         Index           =   1
         Left            =   1080
         TabIndex        =   24
         Top             =   120
         Width           =   1815
         _Version        =   786432
         _ExtentX        =   3201
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "„Â„"
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
   End
   Begin ALLButtonS.ALLButton ALLButton1 
      Height          =   375
      Left            =   360
      TabIndex        =   9
      Top             =   8400
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "»ÕÀ"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmSearchDevelopment.frx":000C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox txtremark 
      Alignment       =   1  'Right Justify
      Height          =   1020
      Left            =   14160
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   4680
      Width           =   7830
   End
   Begin VB.Frame FraHeader 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   660
      Left            =   60
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   90
      Width           =   11865
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
               Picture         =   "FrmSearchDevelopment.frx":0028
               Key             =   "CompanyName"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmSearchDevelopment.frx":03C2
               Key             =   "Ser"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmSearchDevelopment.frx":075C
               Key             =   "Vac_Name"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmSearchDevelopment.frx":0AF6
               Key             =   "ShareCount"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmSearchDevelopment.frx":0E90
               Key             =   "Dis_Count"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmSearchDevelopment.frx":122A
               Key             =   "Bouns"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmSearchDevelopment.frx":15C4
               Key             =   "SharesValue"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmSearchDevelopment.frx":1B5E
               Key             =   "BuyValue"
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "»ÕÀ „ «»⁄… «· ÿÊÌ—"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   375
         Index           =   2
         Left            =   4320
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   120
         Width           =   7080
      End
   End
   Begin MSComCtl2.DTPicker XPDtbBill 
      Height          =   315
      Left            =   3000
      TabIndex        =   5
      Top             =   3600
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   556
      _Version        =   393216
      CheckBox        =   -1  'True
      Format          =   142475265
      CurrentDate     =   38784
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Height          =   315
      Left            =   10080
      TabIndex        =   6
      Top             =   6120
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "6"
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VSFlex8UCtl.VSFlexGrid Fg 
      Height          =   2745
      Left            =   60
      TabIndex        =   10
      Top             =   750
      Width           =   11835
      _cx             =   20876
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
      Cols            =   12
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmSearchDevelopment.frx":1EF8
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
      Index           =   0
      Left            =   2400
      TabIndex        =   11
      Top             =   5460
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
      Left            =   1380
      TabIndex        =   12
      Top             =   5460
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
      Left            =   480
      TabIndex        =   13
      Top             =   5460
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
   Begin MSDataListLib.DataCombo DcbCustomer 
      Height          =   315
      Left            =   6000
      TabIndex        =   15
      Top             =   6240
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DcboEmpName 
      Height          =   315
      Left            =   120
      TabIndex        =   18
      Top             =   4440
      Width           =   3675
      _ExtentX        =   6482
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DcbTypeVisit1 
      Height          =   315
      Left            =   6000
      TabIndex        =   20
      Top             =   4080
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DcbManager 
      Height          =   315
      Left            =   120
      TabIndex        =   26
      Top             =   4080
      Width           =   3675
      _ExtentX        =   6482
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSComCtl2.DTPicker ToDate 
      Height          =   315
      Left            =   480
      TabIndex        =   27
      Top             =   3600
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   556
      _Version        =   393216
      CheckBox        =   -1  'True
      Format          =   142475265
      CurrentDate     =   38784
   End
   Begin MSDataListLib.DataCombo DcbDes 
      Height          =   315
      Left            =   6000
      TabIndex        =   29
      Top             =   4440
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VSFlex8UCtl.VSFlexGrid FG2 
      Height          =   2745
      Left            =   0
      TabIndex        =   33
      Top             =   660
      Width           =   11835
      _cx             =   20876
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
      Cols            =   14
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmSearchDevelopment.frx":20B7
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
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "«·⁄„·Ì…"
      Height          =   285
      Index           =   1
      Left            =   10530
      TabIndex        =   30
      Top             =   4440
      Width           =   1245
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "«·Ï"
      Height          =   375
      Left            =   2280
      TabIndex        =   28
      Top             =   3600
      Width           =   495
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "«·„Â„Â"
      Height          =   285
      Index           =   2
      Left            =   10530
      TabIndex        =   21
      Top             =   4080
      Width           =   1245
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "„œÌ— «·„Â„Â"
      Height          =   285
      Index           =   3
      Left            =   4770
      TabIndex        =   19
      Top             =   4440
      Width           =   1005
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "«·„ÊŸ›"
      Height          =   285
      Index           =   0
      Left            =   4650
      TabIndex        =   17
      Top             =   4080
      Width           =   1125
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "«”„ «·⁄„Ì·"
      Height          =   285
      Index           =   10
      Left            =   10560
      TabIndex        =   16
      Top             =   6240
      Width           =   1125
   End
   Begin VB.Label lblitemid 
      Alignment       =   1  'Right Justify
      Height          =   495
      Left            =   960
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "„·ÕÊŸ…"
      Height          =   375
      Left            =   14160
      TabIndex        =   7
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "«·»·œ"
      Height          =   375
      Left            =   8640
      TabIndex        =   4
      Top             =   6240
      Width           =   735
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "„‰  «—ÌŒ"
      Height          =   375
      Left            =   4680
      TabIndex        =   3
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "—ﬁ„ «·√„—"
      Height          =   375
      Left            =   10560
      TabIndex        =   2
      Top             =   3600
      Width           =   1215
   End
End
Attribute VB_Name = "FemSearchDevelopment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim rs As ADODB.Recordset
Dim cSearchDcbo As clsDCboSearch

Private m_DcboItems As DataCombo

Private m_RetrunType As Integer
Public WithEvents Fg1 As VSFlex8UCtl.vsFlexGrid
Attribute Fg1.VB_VarHelpID = -1

Public WithEvents NewGrid As VSFlex8UCtl.vsFlexGrid
Attribute NewGrid.VB_VarHelpID = -1
'Public NewGrid As New ClsGrid
 
Public LngRow As Long

Public LngCol As Long
Public mIndex As Long



Private Sub Cmd_Click(Index As Integer)
    On Error GoTo ErrTrap

    Select Case Index

        Case 0

            If rs.State = adStateOpen Then
                rs.Close
            End If
            If mIndex <> 1 Then
                rs.Open Build_Sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
            
                If SystemOptions.UserInterface = ArabicInterface Then
                    '   LblRes.Caption = "‰ ÌÃ… «·»ÕÀ = " & rs.RecordCount
                ElseIf SystemOptions.UserInterface = EnglishInterface Then
                    '   LblRes.Caption = "Search Result=" & rs.RecordCount
                End If
        
                If rs.RecordCount < 1 Then
                    Fg.Clear flexClearScrollable, flexClearEverything
                    Fg.Rows = 2
    
                    If SystemOptions.UserInterface = ArabicInterface Then
                        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
                        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                    Else
                        Msg = "NO Search Results Found...!!!"
                        MsgBox Msg, vbOKOnly + vbExclamation, App.title
                    End If
    
                    Exit Sub
                End If
    
                Retrive
                Fg.SetFocus
            ElseIf mIndex = 1 Then
                Fg.Visible = False
                FG2.Visible = True
                loadgrid Build_Sql2, FG2, True, False
            End If
        Case 1
            clear_all Me
            Fg.Clear flexClearScrollable, flexClearEverything
            XPDtbBill.value = ""
            ToDate.value = ""
 Opt(0).value = False
             Opt(1).value = False
        Case 2
            Unload Me
    End Select

    Exit Sub
ErrTrap:

    If Err.Number = -2147217900 Then
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & CHR(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… „⁄«ÌÌ— «·»ÕÀ Ê√⁄œ «·„Õ«Ê·…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

End Sub






Private Sub DcbManager_Change()
DcbManager_Click (0)
End Sub

Private Sub DcbManager_Click(Area As Integer)
 If val(DcbManager.BoundText) = 0 Then Exit Sub
 Dim EmpCode  As String

    GetEmployeeIDFromCode , , DcbManager.BoundText, EmpCode
    Me.Text1.Text = EmpCode
End Sub

Private Sub DcboEmpName_Change()
DcboEmpName_Click (0)
End Sub

Private Sub DcboEmpName_Click(Area As Integer)
 If val(DcboEmpName.BoundText) = 0 Then Exit Sub
 Dim EmpCode  As String

    GetEmployeeIDFromCode , , DcboEmpName.BoundText, EmpCode
    Me.Text2.Text = EmpCode
End Sub

Private Sub DcbTypeVisit1_Change()
DcbTypeVisit1_Click (0)
End Sub

Private Sub DcbTypeVisit1_Click(Area As Integer)
If val(DcbTypeVisit1.BoundText) <> 0 Then
    Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos
    Dcombos.GetDevelopProcessPand Me.DcbDes, val(DcbTypeVisit1.BoundText)
    End If
End Sub

Private Sub fg_Click()
    On Error GoTo ErrTrap
       
      FrmRegDevelopment.Retrive val(Fg.TextMatrix(Fg.Row, 1))

ErrTrap:
End Sub

Private Sub Retrive()
    Dim Num As Integer
    On Error GoTo ErrTrap
    Fg.Clear flexClearScrollable, flexClearEverything

    If Not (rs.EOF Or rs.BOF) Then
        Fg.Rows = rs.RecordCount + 1

        For Num = 1 To rs.RecordCount

            With Fg
                .TextMatrix(Num, .ColIndex("id")) = IIf(IsNull(rs("id").value), "", rs("id").value)
                .TextMatrix(Num, .ColIndex("RecordDate")) = IIf(IsNull(rs("RecordDate").value), "", rs("RecordDate").value)
                If Not (IsNull(rs("Important").value)) Then
                If SystemOptions.UserInterface = ArabicInterface Then
                If val(rs("Important").value) = 0 Then
                .TextMatrix(Num, .ColIndex("Important")) = "⁄«œÌ"
                ElseIf val(rs("Important").value) = 1 Then
                 .TextMatrix(Num, .ColIndex("Important")) = "„Â„"
                End If
                Else
                   If val(rs("Important").value) = 0 Then
                .TextMatrix(Num, .ColIndex("Important")) = "Normal"
                ElseIf val(rs("Important").value) = 1 Then
                 .TextMatrix(Num, .ColIndex("Important")) = "Important"
                End If
                End If
                End If
                
                

               If SystemOptions.UserInterface = ArabicInterface Then
               .TextMatrix(Num, .ColIndex("Des")) = IIf(IsNull(rs("Des").value), "", rs("Des").value)
                .TextMatrix(Num, .ColIndex("Emp_Name")) = IIf(IsNull(rs("Emp_Name").value), "", Trim(rs("Emp_Name").value))
                .TextMatrix(Num, .ColIndex("MangEmp_Name")) = IIf(IsNull(rs("MangEmp_Name").value), "", Trim(rs("MangEmp_Name").value))
                .TextMatrix(Num, .ColIndex("Name")) = IIf(IsNull(rs("Name").value), "", Trim(rs("Name").value))
                '.TextMatrix(Num, .ColIndex("Emp_Name")) = IIf(IsNull(rs("Emp_Name").value), "", Trim(rs("Emp_Name").value))
               Else
               .TextMatrix(Num, .ColIndex("Des")) = IIf(IsNull(rs("desE").value), "", rs("desE").value)
                .TextMatrix(Num, .ColIndex("Emp_Name")) = IIf(IsNull(rs("Emp_Namee").value), "", Trim(rs("Emp_Namee").value))
                .TextMatrix(Num, .ColIndex("MangEmp_Name")) = IIf(IsNull(rs("MangEmp_NameE").value), "", Trim(rs("MangEmp_NameE").value))
                .TextMatrix(Num, .ColIndex("Name")) = IIf(IsNull(rs("NameE").value), "", Trim(rs("NameE").value))
                    '  .TextMatrix(Num, .ColIndex("Emp_NameM")) = IIf(IsNull(rs("CusNamee").value), "", Trim(rs("CusNamee").value))
                End If
            End With

            rs.MoveNext
        Next Num

        Fg.AutoSize 0, Fg.Cols - 1, False
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Fg_DblClick()
    fg_Click
    Unload Me
End Sub

Private Sub Fg2_Click()
    On Error GoTo ErrTrap
       
      dean.FiLLTXT9 val(FG2.TextMatrix(FG2.Row, 1))

ErrTrap:

End Sub

Private Sub Form_Activate()
    If mIndex = 1 Then
    Label1(2).Caption = "»ÕÀ ⁄‰ œŒÊ· «·„⁄œ« /«·”Ì«—« "
    Me.Caption = "»ÕÀ ⁄‰ œŒÊ· «·„⁄œ« /«·”Ì«—« "
    ReloadCompo
    Frame1.Visible = True
    Fg.Visible = False
    FG2.Visible = True
    
 End If
End Sub

Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim StrSQL As String
    Dim BG As New ClsBackGroundPic
    Dim Dcombos As ClsDataCombos
    If mIndex = 1 Then
    Label1(2).Caption = "»ÕÀ «·—Õ·« "
    ReloadCompo
    Frame1.Visible = True
    Fg.Visible = False
    FG2.Visible = True
    
 End If

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

 XPDtbBill.value = ""
            ToDate.value = ""
            Opt(0).value = False
             Opt(1).value = False
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Clear").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
 
    Dim My_SQL As String
    Set Dcombos = New ClsDataCombos

     Dcombos.GetDevelopProcess Me.DcbTypeVisit1
      Dcombos.GetEmployees Me.DcboEmpName
        Dcombos.GetEmployees Me.DcbManager

   Dcombos.GetBranches cmbLocationsName
       '   Dcombos.GetFileCustomer Me.DcbCustomer
    'RetrunType = -1
 
    CenterForm Me

    FormPostion Me, GetPostion
    Fg.WallPaper = BG.SearchWallpaper
    Set rs = New ADODB.Recordset
    DBCboClientName.BoundText = ""
    If mIndex = 1 Then
    Label1(2).Caption = "»ÕÀ «·—Õ·« "
    ReloadCompo
    Frame1.Visible = True
    Fg.Visible = False
    FG2.Visible = True
    
 End If
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

StrSQL = "SELECT     dbo.TblRegDevelopment.Id, dbo.TblRegDevelopment.RecordDate, dbo.TblRegDevelopment.StrDate, dbo.TblRegDevelopment.EndExptedDate, "
  StrSQL = StrSQL & "                    dbo.TblRegDevelopment.EndActDate, dbo.TblRegDevelopment.UserID, dbo.TblRegDevelopment.Important, dbo.TblRegDevelopment.MoDay,"
  StrSQL = StrSQL & "                    dbo.TblRegDevelopment.CusID, dbo.TblRegDevelopment.DesOp, dbo.TblRegDevelopment.AnlysOp, dbo.TblRegDevelopment.TimeReq,"
  StrSQL = StrSQL & "                    dbo.TblRegDevelopment.StartTime, dbo.TblRegDevelopment.EndExptedTime, dbo.TblRegDevelopment.EndActTIme, dbo.TblRegDevelopment.StatusProcess,"
  StrSQL = StrSQL & "                    dbo.TblRegDevelopment.StatusPand, dbo.TblRegDevelopment.NoDaySatart, dbo.TblRegDevelopment.NoDayEnd, dbo.TblRegDevelopment.FromDate,"
  StrSQL = StrSQL & "                    dbo.TblRegDevelopment.ToDate, dbo.TblRegDevelopment.BranchID, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee,"
  StrSQL = StrSQL & "                    dbo.TblRegDevelopment.EmpID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee, dbo.TblRegDevelopment.MangID,"
  StrSQL = StrSQL & "                    TblEmployee_1.Emp_Name AS MangEmp_Name, TblEmployee_1.Fullcode AS MangFullcode, TblEmployee_1.Emp_Namee AS MangEmp_NameE,"
  StrSQL = StrSQL & "                    dbo.TblRegDevelopment.OpType, dbo.TblProceeDevelper.Name, dbo.TblProceeDevelper.NameE, dbo.TblRegDevelopment.DesID,"
  StrSQL = StrSQL & "                    dbo.TblProceeDevelperDet.des , dbo.TblRegDevelopment.RecordTime ,dbo.TblProceeDevelperDet.desE"
  StrSQL = StrSQL & " FROM         dbo.TblBranchesData RIGHT OUTER JOIN"
  StrSQL = StrSQL & "                    dbo.TblRegDevelopment LEFT OUTER JOIN"
  StrSQL = StrSQL & "                    dbo.TblProceeDevelperDet ON dbo.TblRegDevelopment.DesID = dbo.TblProceeDevelperDet.ID LEFT OUTER JOIN"
  StrSQL = StrSQL & "                    dbo.TblProceeDevelper ON dbo.TblRegDevelopment.OpType = dbo.TblProceeDevelper.ID LEFT OUTER JOIN"
  StrSQL = StrSQL & "                    dbo.TblEmployee TblEmployee_1 ON dbo.TblRegDevelopment.MangID = TblEmployee_1.Emp_ID LEFT OUTER JOIN"
  StrSQL = StrSQL & "                    dbo.TblEmployee ON dbo.TblRegDevelopment.EmpID = dbo.TblEmployee.Emp_ID ON dbo.TblBranchesData.branch_id = dbo.TblRegDevelopment.BranchID"
StrSQL = StrSQL & " Where (1 = 1)"
 
StrWhere = ""
    
    If Me.Opt(0).value = True Then
 
        StrWhere = StrWhere + " and dbo.TblRegDevelopment.Important = 0 "
 
    End If
    If Me.Opt(1).value = True Then
 
        StrWhere = StrWhere + " and dbo.TblRegDevelopment.Important = 1 "
 
    End If
    
    
    If val(Me.DcbManager.BoundText) <> 0 And Me.DcbManager.Text <> "" Then
 
        StrWhere = StrWhere + " and dbo.TblRegDevelopment.MangID =" & Me.DcbManager.BoundText & ""
 
    End If
 
    If val(Me.DcboEmpName.BoundText) <> 0 And Me.DcboEmpName.Text <> "" Then
 
        StrWhere = StrWhere + " and dbo.TblRegDevelopment.EmpID =" & val(Me.DcboEmpName.BoundText)
 
    End If
    
   ' If Me.DcbCustomer.BoundText <> "" And Me.DcbCustomer.text <> "" Then
 
   '     StrWhere = StrWhere + " and dbo.TblRegDevelopment.CusID =" & val(Me.DcbCustomer.BoundText)
 '
 '   End If
    
       If val(Me.DcbTypeVisit1.BoundText) <> 0 And Me.DcbTypeVisit1.Text <> "" Then
 
        StrWhere = StrWhere + " and dbo.TblRegDevelopment.OpType =" & val(Me.DcbTypeVisit1.BoundText)
 
    End If
       If val(Me.DcbDes.BoundText) <> 0 And Me.DcbDes.Text <> "" Then
 
        StrWhere = StrWhere + " and dbo.TblRegDevelopment.DesID =" & val(Me.DcbDes.BoundText)
 
    End If
    
    If Me.txtorder_no.Text <> "" Then
 
        StrWhere = StrWhere + " and dbo.TblRegDevelopment.Id ='" & Me.txtorder_no.Text & "'"
 
    End If
 If Not IsNull(Me.XPDtbBill.value) Then
       
            StrWhere = StrWhere & " AND dbo.TblRegDevelopment.RecordDate >=" & SQLDate(Me.XPDtbBill.value, True) & ""
        End If
        
         If Not IsNull(Me.ToDate.value) Then
       
            StrWhere = StrWhere & " AND dbo.TblRegDevelopment.RecordDate <=" & SQLDate(Me.ToDate.value, True) & ""
        End If
    StrWhere = StrWhere + " order by dbo.TblRegDevelopment.Id"

    Build_Sql = StrSQL + StrWhere
    Exit Function
ErrTrap:
End Function


Private Function Build_Sql2()
    Dim StrSQL As String
    Dim Begin As Boolean
    Dim StrWhere As String
    Dim BolHaveSerial As Boolean
    Dim IntHaveSerial As Integer
 
    On Error GoTo ErrTrap

StrSQL = "SELECT    *,TblBranchesData.branch_name   from TblTripReg  Left  Outer join TblBranchesData On  TblBranchesData.branch_id = TblTripReg.branchid"

StrSQL = StrSQL & " Where (1 = 1)"
 
StrWhere = ""
    

    
    
    If Trim(Me.txtCustName.Text) <> "" Then
 
        StrWhere = StrWhere + " and dbo.TblTripReg.CustName Like '%" & Trim(Me.txtCustName.Text) & "%'"
 
    End If
 
     If Trim(Me.txtPhoneCust.Text) <> "" Then
 
        StrWhere = StrWhere + " and dbo.TblTripReg.PhoneCust Like '%" & Trim(Me.txtPhoneCust.Text) & "%'"
 
    End If
    
     If Trim(Me.XPTxtVal.Text) <> "" Then
 
        StrWhere = StrWhere + " and dbo.TblTripReg.TotalWithVat2 = " & val(Me.XPTxtVal.Text)
 
    End If
    
    
    If val(Me.cmbLocationsName.BoundText) <> 0 And Me.cmbLocationsName.Text <> "" Then
 
        StrWhere = StrWhere + " and dbo.TblTripReg.BranchID =" & val(Me.cmbLocationsName.BoundText)
 
    End If
    
   ' If Me.DcbCustomer.BoundText <> "" And Me.DcbCustomer.text <> "" Then
 
   '     StrWhere = StrWhere + " and dbo.TblRegDevelopment.CusID =" & val(Me.DcbCustomer.BoundText)
 '
 '   End If
    
     If Me.cmbCarName.Text <> "" Then
 
        StrWhere = StrWhere + " and dbo.TblTripReg.nBoardNo Like '%" & Trim(Me.cmbCarName.Text) & "%'"
 
    End If
    
    If val(Me.CboPayMentType.ListIndex) <> -1 Then
 
        StrWhere = StrWhere + " and dbo.TblTripReg.PayType =" & val(Me.CboPayMentType.ListIndex)
 
    End If
    
    
      If val(Me.cmbPaymentClass.BoundText) <> 0 And Me.cmbPaymentClass.Text <> "" Then
 
        StrWhere = StrWhere + " and dbo.TblTripReg.PaymentClassID =" & val(Me.cmbPaymentClass.BoundText)
 
    End If
    
    If Me.txtorder_no.Text <> "" Then
 
        StrWhere = StrWhere + " and dbo.TblTripReg.NoteSerial1 ='" & Me.txtorder_no.Text & "'"
 
    End If
    
 If Not IsNull(Me.XPDtbBill.value) Then
       
            StrWhere = StrWhere & " AND dbo.TblTripReg.RecordDate >=" & SQLDate(Me.XPDtbBill.value, True) & ""
        End If
        
         If Not IsNull(Me.ToDate.value) Then
       
            StrWhere = StrWhere & " AND dbo.TblTripReg.RecordDate <=" & SQLDate(Me.ToDate.value, True) & ""
        End If
    StrWhere = StrWhere + " order by dbo.TblTripReg.Id"

    Build_Sql2 = StrSQL + StrWhere
    Exit Function
ErrTrap:
End Function


Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        If Me.ActiveControl Is Fg Then
            If Not Fg.TextMatrix(Fg.Row, 1) = "" Then
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



Private Sub ChangeLang()
    Me.Caption = "Search For Follow-Up Development"
    Label1(2).Caption = Me.Caption
    Label2.Caption = "Trans No"
    lbl(2).Caption = "Task"
    lbl(1).Caption = "Process"
    lbl(0).Caption = "Employee"
    lbl(3).Caption = "Manger"
' Lbl(0).Caption = "Manager"
    Label3.Caption = "From Date"
 'Lbl(3).Caption = "Based Dev"
    Label4.Caption = "TO"
   ' Lbl(2).Caption = "Type Process"
    Frame2.Caption = "Priority"
Opt(0).RightToLeft = False
Opt(1).RightToLeft = False
Opt(0).Caption = "Normal"
Opt(1).Caption = "Imprtant"
 
    Cmd(0).Caption = "Search"
    Cmd(1).Caption = "Clear"
    Cmd(2).Caption = "Exit"

    'OptType(0).Caption = "Start of the name"
    'OptType(1).Caption = "any part of the name"
    With Me.Fg
        .TextMatrix(0, .ColIndex("NumIndex")) = "Serial"
        .TextMatrix(0, .ColIndex("id")) = "Trans No"
         .TextMatrix(0, .ColIndex("RecordDate")) = "RecordDate  "
        .TextMatrix(0, .ColIndex("Important")) = "Priority"
        .TextMatrix(0, .ColIndex("Name")) = " Task"
        .TextMatrix(0, .ColIndex("Des")) = "Process"
       .TextMatrix(0, .ColIndex("Emp_Name")) = "Task Manager"
       .TextMatrix(0, .ColIndex("MangEmp_Name")) = "Employee"
       .AutoSize 0, .Cols - 1, False
    End With

End Sub

Private Sub lblItemID_Change()
DCboItem.BoundText = val(Me.lblitemid.Caption)
End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode Text1.Text, EmpID
        DcbManager.BoundText = EmpID
    End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode Text2.Text, EmpID
        DcboEmpName.BoundText = EmpID
    End If
End Sub

Sub ReloadCompo()
Dim sql As String
sql = "SELECT DISTINCT LocationsName, LocationsName AS LocationsName"
sql = sql & " From dbo.TblTripReg"
sql = sql & " WHERE     (NOT (LocationsName IS NULL)) "
fill_combo cmbLocationsName, sql
    Dim Dcombos As New ClsDataCombos

sql = "SELECT DISTINCT nBoardNo, nBoardNo AS CarName"
sql = sql & " From dbo.TblTripReg"
sql = sql & " WHERE     (NOT (nBoardNo IS NULL)) "
fill_combo cmbCarName, sql

'sql = "SELECT DISTINCT CustName, CustName AS CustName"
'sql = sql & " From dbo.TblTripReg"
'sql = sql & " WHERE     (NOT (CustName IS NULL)) "
'fill_combo cmbCustName, sql

sql = "SELECT DISTINCT Id, Name ,Namee  from tblPaymentClass "




fill_combo cmbPaymentClass, sql
 If SystemOptions.UserInterface = ArabicInterface Then

        With CboPayMentType
             .Clear
             .AddItem "‰ﬁœ«"
             .AddItem "¬Ã·"
             
         End With
         
    Else
         With CboPayMentType
            .Clear
            'AddItem "Cash"
            
            .AddItem "Cash"
            .AddItem "Cheque"
            .AddItem "Visa"
            .AddItem "Master"
        End With
        
    End If
Dcombos.GetBranches cmbLocationsName
End Sub
