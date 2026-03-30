VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form FrmIqarContractSearch 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6945
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   16095
   Icon            =   "FrmIqarContractSearch.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6945
   ScaleWidth      =   16095
   Begin VB.OptionButton Opt 
      Alignment       =   1  'Right Justify
      Caption         =   "≈ŸÂ«— «·⁄ﬁÊœ «·„ ’›Ì…"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   2
      Left            =   9360
      RightToLeft     =   -1  'True
      TabIndex        =   75
      Top             =   6480
      Width           =   2535
   End
   Begin VB.OptionButton Opt 
      Alignment       =   1  'Right Justify
      Caption         =   "≈ŸÂ«— «·⁄ﬁÊœ «· Ì ·„   ’›Ï"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   1
      Left            =   11880
      RightToLeft     =   -1  'True
      TabIndex        =   74
      Top             =   6480
      Width           =   2535
   End
   Begin VB.OptionButton Opt 
      Alignment       =   1  'Right Justify
      Caption         =   "≈ŸÂ«— ﬂ· «·⁄ﬁÊœ"
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   0
      Left            =   14400
      RightToLeft     =   -1  'True
      TabIndex        =   73
      Top             =   6480
      Width           =   1575
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
      Height          =   315
      Left            =   7800
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   69
      Top             =   7440
      Width           =   825
   End
   Begin VB.ComboBox DcbPeriodsID 
      Height          =   315
      ItemData        =   "FrmIqarContractSearch.frx":000C
      Left            =   6960
      List            =   "FrmIqarContractSearch.frx":0016
      RightToLeft     =   -1  'True
      TabIndex        =   68
      Top             =   7440
      Width           =   735
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
      Height          =   315
      Left            =   3360
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   66
      Top             =   6000
      Width           =   5205
   End
   Begin VB.TextBox TxtIncresYearRate 
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
      Left            =   13020
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   61
      Top             =   6000
      Width           =   1965
   End
   Begin VB.TextBox TxtIncresYearValue 
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
      Left            =   9840
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   60
      Top             =   6000
      Width           =   1725
   End
   Begin VB.ComboBox DcbFurnishing 
      Height          =   315
      ItemData        =   "FrmIqarContractSearch.frx":0024
      Left            =   120
      List            =   "FrmIqarContractSearch.frx":002E
      RightToLeft     =   -1  'True
      TabIndex        =   58
      Top             =   5520
      Width           =   1965
   End
   Begin VB.TextBox TxtWater 
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
      Left            =   13020
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   53
      Top             =   5520
      Width           =   1965
   End
   Begin VB.TextBox TxtElectricity 
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
      Left            =   9840
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   52
      Top             =   5520
      Width           =   1725
   End
   Begin VB.TextBox TxtPhone 
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
      Left            =   6600
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   51
      Top             =   5520
      Width           =   1965
   End
   Begin VB.TextBox TxtEnternet 
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
      Left            =   3360
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   50
      Top             =   5520
      Width           =   1965
   End
   Begin VB.TextBox TxtInsuranceValue 
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
      Left            =   120
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   45
      Top             =   5040
      Width           =   1965
   End
   Begin VB.TextBox TxtPayAmini 
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
      Left            =   6600
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   44
      Top             =   5040
      Width           =   1965
   End
   Begin VB.TextBox TxtCommiValue 
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
      Left            =   3360
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   43
      Top             =   5040
      Width           =   1965
   End
   Begin VB.TextBox TxtTotalContract 
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
      Left            =   120
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   39
      Top             =   4560
      Width           =   1965
   End
   Begin VB.TextBox TxtMeterValue 
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
      Left            =   6600
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   38
      Top             =   4560
      Width           =   1965
   End
   Begin VB.TextBox TxtMeterCount 
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
      Left            =   3360
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   37
      Top             =   4560
      Width           =   1965
   End
   Begin VB.ComboBox DcbRentType 
      Height          =   315
      ItemData        =   "FrmIqarContractSearch.frx":0042
      Left            =   120
      List            =   "FrmIqarContractSearch.frx":004C
      RightToLeft     =   -1  'True
      TabIndex        =   28
      Top             =   4080
      Width           =   1965
   End
   Begin VB.TextBox Text15 
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
      Left            =   13920
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   4560
      Width           =   1065
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
      Left            =   13920
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   5040
      Width           =   1065
   End
   Begin VB.TextBox TxtSearch 
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
      Left            =   13920
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   4080
      Width           =   1035
   End
   Begin VB.ComboBox DcbContType 
      Height          =   315
      ItemData        =   "FrmIqarContractSearch.frx":0060
      Left            =   9840
      List            =   "FrmIqarContractSearch.frx":006A
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   3600
      Width           =   1695
   End
   Begin ALLButtonS.ALLButton ALLButton1 
      Height          =   375
      Left            =   360
      TabIndex        =   10
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
      MICON           =   "FrmIqarContractSearch.frx":0082
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
      Left            =   17040
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   9
      Top             =   4680
      Width           =   7830
   End
   Begin VB.TextBox txtorder_no 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   13260
      TabIndex        =   5
      Top             =   3600
      Width           =   1695
   End
   Begin VB.Frame FraHeader 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   660
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   16065
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
               Picture         =   "FrmIqarContractSearch.frx":009E
               Key             =   "CompanyName"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmIqarContractSearch.frx":0438
               Key             =   "Ser"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmIqarContractSearch.frx":07D2
               Key             =   "Vac_Name"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmIqarContractSearch.frx":0B6C
               Key             =   "ShareCount"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmIqarContractSearch.frx":0F06
               Key             =   "Dis_Count"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmIqarContractSearch.frx":12A0
               Key             =   "Bouns"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmIqarContractSearch.frx":163A
               Key             =   "SharesValue"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmIqarContractSearch.frx":1BD4
               Key             =   "BuyValue"
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "»ÕÀ »Ì«‰«  «·⁄ﬁÊœ"
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
         Left            =   10680
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   0
         Width           =   5280
      End
   End
   Begin MSComCtl2.DTPicker XPDtbBill 
      Height          =   315
      Left            =   6600
      TabIndex        =   6
      Top             =   3600
      Width           =   1965
      _ExtentX        =   3466
      _ExtentY        =   556
      _Version        =   393216
      CheckBox        =   -1  'True
      Format          =   99287041
      CurrentDate     =   38784
   End
   Begin MSDataListLib.DataCombo DataCombo4 
      Height          =   315
      Left            =   10080
      TabIndex        =   7
      Top             =   8160
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
      Left            =   0
      TabIndex        =   11
      Top             =   750
      Width           =   16035
      _cx             =   28284
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
      Cols            =   34
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmIqarContractSearch.frx":1F6E
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
      TabIndex        =   12
      Top             =   6480
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
      TabIndex        =   13
      Top             =   6480
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
      TabIndex        =   14
      Top             =   6480
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
   Begin ImpulseButton.ISButton CmdItemSearch 
      Height          =   345
      Index           =   2
      Left            =   -480
      TabIndex        =   15
      Top             =   4410
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   609
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "..."
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
      ButtonImage     =   "FrmIqarContractSearch.frx":24A1
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin MSDataListLib.DataCombo DcbIqara 
      Height          =   315
      Left            =   9840
      TabIndex        =   20
      Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·⁄ﬁ«—"
      Top             =   4080
      Width           =   4035
      _ExtentX        =   7117
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo dcsupplier 
      Height          =   315
      Left            =   9840
      TabIndex        =   23
      Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·„«·ﬂ"
      Top             =   5040
      Width           =   4035
      _ExtentX        =   7117
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo dcCustomer 
      Height          =   315
      Left            =   9840
      TabIndex        =   26
      Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«— √”„ «·„” «Ã—"
      Top             =   4560
      Width           =   4035
      _ExtentX        =   7117
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DcbUnitNo 
      Height          =   315
      Left            =   3360
      TabIndex        =   29
      Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
      Top             =   4080
      Width           =   1965
      _ExtentX        =   3466
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSComCtl2.DTPicker StrDate 
      Height          =   315
      Left            =   3360
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   3600
      Width           =   1965
      _ExtentX        =   3466
      _ExtentY        =   556
      _Version        =   393216
      CalendarBackColor=   12648447
      CalendarTitleBackColor=   10383715
      CheckBox        =   -1  'True
      CustomFormat    =   "yyyy/M/d"
      Format          =   99287043
      CurrentDate     =   37140
   End
   Begin MSComCtl2.DTPicker EndDate 
      Height          =   315
      Left            =   120
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   3600
      Width           =   1965
      _ExtentX        =   3466
      _ExtentY        =   556
      _Version        =   393216
      CalendarBackColor=   12648447
      CalendarTitleBackColor=   10383715
      CheckBox        =   -1  'True
      CustomFormat    =   "yyyy/M/d"
      Format          =   99287043
      CurrentDate     =   37140
   End
   Begin MSComCtl2.DTPicker FristPaymentDate 
      Height          =   315
      Left            =   120
      TabIndex        =   64
      TabStop         =   0   'False
      Top             =   6000
      Width           =   1965
      _ExtentX        =   3466
      _ExtentY        =   556
      _Version        =   393216
      CalendarBackColor=   12648447
      CalendarTitleBackColor=   10383715
      CheckBox        =   -1  'True
      CustomFormat    =   "yyyy/M/d"
      Format          =   99287043
      CurrentDate     =   37140
   End
   Begin MSDataListLib.DataCombo DcbUnitType 
      Height          =   315
      Left            =   6600
      TabIndex        =   71
      Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
      Top             =   4080
      Width           =   1965
      _ExtentX        =   3466
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DcbUnitNo1 
      Height          =   315
      Left            =   10440
      TabIndex        =   72
      Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
      Top             =   9360
      Width           =   1965
      _ExtentX        =   3466
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   "«·› —Â »Ì‰ «·œ›⁄« "
      Height          =   405
      Index           =   11
      Left            =   8445
      RightToLeft     =   -1  'True
      TabIndex        =   70
      Top             =   7440
      Width           =   1050
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   "⁄œœ «·œ›⁄« "
      Height          =   285
      Index           =   8
      Left            =   8640
      RightToLeft     =   -1  'True
      TabIndex        =   67
      Top             =   6000
      Width           =   930
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   " «—ÌŒ «Ê· œ›⁄Â"
      Height          =   285
      Index           =   9
      Left            =   2040
      RightToLeft     =   -1  'True
      TabIndex        =   65
      Top             =   6000
      Width           =   1170
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   "«·“Ì«œ… «·”‰ÊÌ… %"
      Height          =   195
      Index           =   30
      Left            =   14745
      RightToLeft     =   -1  'True
      TabIndex        =   63
      Top             =   6000
      Width           =   1350
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   "«·“Ì«œ… «·”‰ÊÌ… ﬁÌ„…"
      Height          =   195
      Index           =   31
      Left            =   11625
      RightToLeft     =   -1  'True
      TabIndex        =   62
      Top             =   6000
      Width           =   1350
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   "«· √ÀÌÀ"
      Height          =   285
      Index           =   29
      Left            =   2400
      RightToLeft     =   -1  'True
      TabIndex        =   59
      Top             =   5520
      Width           =   810
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   "„Ì«Â"
      Height          =   195
      Index           =   20
      Left            =   14985
      RightToLeft     =   -1  'True
      TabIndex        =   57
      Top             =   5520
      Width           =   990
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   "ﬂÂ—»«¡"
      Height          =   195
      Index           =   21
      Left            =   11985
      RightToLeft     =   -1  'True
      TabIndex        =   56
      Top             =   5520
      Width           =   990
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   "Â« ›"
      Height          =   195
      Index           =   27
      Left            =   8505
      RightToLeft     =   -1  'True
      TabIndex        =   55
      Top             =   5520
      Width           =   990
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   "Œœ„… «‰ —‰ "
      Height          =   195
      Index           =   28
      Left            =   5460
      RightToLeft     =   -1  'True
      TabIndex        =   54
      Top             =   5520
      Width           =   990
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   " √„Ì‰"
      Height          =   195
      Index           =   19
      Left            =   2220
      RightToLeft     =   -1  'True
      TabIndex        =   49
      Top             =   5040
      Width           =   990
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   "—”Ê„ «œ«—Ì…"
      Height          =   195
      Index           =   24
      Left            =   8505
      RightToLeft     =   -1  'True
      TabIndex        =   48
      Top             =   5040
      Width           =   990
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   "›Ì„… «·⁄„Ê·Â"
      Height          =   405
      Index           =   25
      Left            =   5640
      RightToLeft     =   -1  'True
      TabIndex        =   47
      Top             =   5040
      Width           =   810
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      Height          =   195
      Index           =   26
      Left            =   7440
      RightToLeft     =   -1  'True
      TabIndex        =   46
      Top             =   5040
      Width           =   270
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   "«Ã„«·Ì «·⁄ﬁœ"
      Height          =   195
      Index           =   6
      Left            =   2220
      RightToLeft     =   -1  'True
      TabIndex        =   42
      Top             =   4560
      Width           =   990
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   "›Ì„… «·„ —"
      Height          =   195
      Index           =   17
      Left            =   8610
      RightToLeft     =   -1  'True
      TabIndex        =   41
      Top             =   4560
      Width           =   885
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   "⁄œœ «·«„ «—"
      Height          =   195
      Index           =   18
      Left            =   5460
      RightToLeft     =   -1  'True
      TabIndex        =   40
      Top             =   4560
      Width           =   990
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   "—ﬁ„ «·ÊÕœ…"
      Height          =   195
      Index           =   0
      Left            =   5460
      RightToLeft     =   -1  'True
      TabIndex        =   36
      Top             =   4080
      Width           =   990
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   "‰Â«ÌÂ  «·⁄ﬁœ"
      Height          =   405
      Index           =   23
      Left            =   2400
      RightToLeft     =   -1  'True
      TabIndex        =   35
      Top             =   3600
      Width           =   810
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   "»œ«ÌÂ «·⁄ﬁœ"
      Height          =   285
      Index           =   22
      Left            =   5640
      RightToLeft     =   -1  'True
      TabIndex        =   34
      Top             =   3600
      Width           =   810
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   "‰Ê⁄ «·ÊÕœ…"
      Height          =   195
      Index           =   15
      Left            =   8505
      RightToLeft     =   -1  'True
      TabIndex        =   31
      Top             =   4080
      Width           =   990
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   "‰Ê⁄ «· √ÃÌ—"
      Height          =   195
      Index           =   16
      Left            =   2340
      RightToLeft     =   -1  'True
      TabIndex        =   30
      Top             =   4080
      Width           =   870
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   " «·„” √Ã—"
      Height          =   285
      Index           =   5
      Left            =   15165
      RightToLeft     =   -1  'True
      TabIndex        =   27
      Top             =   4560
      Width           =   810
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   " «·„«·ﬂ"
      Height          =   165
      Index           =   1
      Left            =   15165
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   5040
      Width           =   810
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   "«·⁄ﬁ«—"
      Height          =   195
      Index           =   4
      Left            =   14985
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   4080
      Width           =   990
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      BackStyle       =   0  'Transparent
      Caption         =   "‰Ê⁄ «·⁄ﬁœ"
      Height          =   285
      Index           =   7
      Left            =   11520
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   3600
      Width           =   1410
   End
   Begin VB.Label lblitemid 
      Alignment       =   1  'Right Justify
      Height          =   495
      Left            =   960
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "„·ÕÊŸ…"
      Height          =   375
      Left            =   17040
      TabIndex        =   8
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "«·»·œ"
      Height          =   375
      Left            =   8640
      TabIndex        =   4
      Top             =   8280
      Width           =   735
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   " «—ÌŒ «·⁄ﬁœ"
      Height          =   375
      Left            =   8640
      TabIndex        =   3
      Top             =   3600
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "—ﬁ„ «·⁄ﬁœ"
      Height          =   375
      Left            =   14760
      TabIndex        =   2
      Top             =   3600
      Width           =   1215
   End
End
Attribute VB_Name = "FrmIqarContractSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim rs As ADODB.Recordset
Dim cSearchDcbo As clsDCboSearch

Private m_DcboItems As DataCombo

Public m_RetrunType As Integer
Public WithEvents FG1 As VSFlex8UCtl.VSFlexGrid
Attribute FG1.VB_VarHelpID = -1

Public WithEvents NewGrid As VSFlex8UCtl.VSFlexGrid
Attribute NewGrid.VB_VarHelpID = -1
'Public NewGrid As New ClsGrid
 
Public LngRow As Long

Public LngCol As Long


Private Sub Cmd_Click(Index As Integer)
    On Error GoTo ErrTrap

    Select Case Index

        Case 0

            If rs.State = adStateOpen Then
                rs.Close
            End If

            rs.Open Build_Sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
        
            If SystemOptions.UserInterface = ArabicInterface Then
                '   LblRes.Caption = "‰ ÌÃ… «·»ÕÀ = " & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
                '   LblRes.Caption = "Search Result=" & rs.RecordCount
            End If
    
            If rs.RecordCount < 1 Then
                FG.Clear flexClearScrollable, flexClearEverything
                FG.Rows = 2

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
            FG.SetFocus

        Case 1
            clear_all Me
            
            FG.Clear flexClearScrollable, flexClearEverything
               StrDate.value = ""
                    XPDtbBill.value = ""
                EndDate.value = ""
                FristPaymentDate.value = ""
                 Dcombos.GetIqarUnit -2, 1, DcbUnitNo

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





Private Sub DcbIqara_Change()
DcbIqara_Click (0)
End Sub

Private Sub DcbIqara_Click(Area As Integer)
      If val(DcbIqara.BoundText) = 0 Then Exit Sub

    Dim EmpCode  As String
 
    GetIqarCode , , val(DcbIqara.BoundText), EmpCode
    Me.TxtSearch.Text = EmpCode
End Sub

Private Sub DcbUnitType_Change()
'Dim Dcombos As ClsDataCombos
'Dim idd As Long
'   Set Dcombos = New ClsDataCombos
'
'If val(DcbUnitType.BoundText) > 0 Then
'idd = val(DcbUnitType.BoundText)
'Dcombos.GetIqarUnit idd, Me.DcbUnitNo
'End If
Dim Dcombos As ClsDataCombos
Dim idd As Long
Dim idd1 As Long
   Set Dcombos = New ClsDataCombos

If val(DcbIqara.BoundText) > 0 Then
idd = val(DcbIqara.BoundText)

idd1 = val(DcbUnitType.BoundText)
'If Me.TxtModFlg = "R" Then
Dcombos.GetIqarUnit idd, idd1, Me.DcbUnitNo, "R"
'Else
'Dcombos.GetIqarUnit idd, idd1, Me.DcbUnitNo
'End If
End If

End Sub




Private Sub DcbUnitType_Click(Area As Integer)
DcbUnitType_Change
End Sub

Private Sub dcCustomer_Change()
dcCustomer_Click (0)
End Sub

Private Sub dcCustomer_Click(Area As Integer)
  If val(dcCustomer.BoundText) = 0 Then Exit Sub

    Dim EmpCode  As String
 
    GetTblCustemersCode , , dcCustomer.BoundText, EmpCode, 56
    Me.Text15.Text = EmpCode
End Sub

Private Sub dcCustomer_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        FrmCustemerSearch.SearchType = 1011
        FrmCustemerSearch.show vbModal

    End If
    
End Sub

Private Sub dcsupplier_Change()
dcsupplier_Click (0)
End Sub

Private Sub dcsupplier_Click(Area As Integer)
If val(dcsupplier.BoundText) = 0 Then Exit Sub
    Dim EmpCode  As String
    GetTblCustemersCode , , dcsupplier.BoundText, EmpCode, 57
    Me.Text1.Text = EmpCode
End Sub

Private Sub fg_Click()
    On Error GoTo ErrTrap
    If m_RetrunType = 0 Then
    RSContract.FindRec val(FG.TextMatrix(FG.Row, FG.ColIndex("ContNo"))), True
'RSContract.BtnLast_Click
'RSContract.Retrive val(Fg.TextMatrix(Fg.Row, Fg.ColIndex("ContNo")))
    RSContract.RereivID = val(FG.TextMatrix(FG.Row, FG.ColIndex("ContNo")))
      ' RSContract.FindRec val(Fg.TextMatrix(Fg.Row, Fg.ColIndex("ContNo")))
 ElseIf m_RetrunType = 1 Then
 FrmCashing.TxtContNo = (FG.TextMatrix(FG.Row, FG.ColIndex("ContNo")))
 FrmCashing.TxtContractNo = (FG.TextMatrix(FG.Row, FG.ColIndex("NoteSerial")))
  ElseIf m_RetrunType = 5 Then
 FrmCashing1.TxtContNo = (FG.TextMatrix(FG.Row, FG.ColIndex("ContNo")))
 FrmCashing1.TxtContractNo = (FG.TextMatrix(FG.Row, FG.ColIndex("NoteSerial")))
 FrmCashing1.DcbUnitNo.BoundText = val(FG.TextMatrix(FG.Row, FG.ColIndex("Aqarid")))
 ElseIf m_RetrunType = 2 Then
   
   If val(FG.TextMatrix(FG.Row, FG.ColIndex("NoteSerial"))) = 0 Then
        FrmWaiverSettlement.TxtOrder.Text = val(FG.TextMatrix(FG.Row, FG.ColIndex("ContNo")))
    Else
        FrmWaiverSettlement.TxtOrder.Text = Trim(FG.TextMatrix(FG.Row, FG.ColIndex("NoteSerial")))
       
    End If

                
    
   FrmWaiverSettlement.DcbIqara2.BoundText = val(Trim(FG.TextMatrix(FG.Row, FG.ColIndex("Aqarid"))))
FrmWaiverSettlement.DcbUnitType2.BoundText = val(Trim(FG.TextMatrix(FG.Row, FG.ColIndex("UnitTypeID"))))
FrmWaiverSettlement.DcbUnitNo2.BoundText = val(Trim(FG.TextMatrix(FG.Row, FG.ColIndex("UnitNo"))))
FrmWaiverSettlement.dcCustomer2.BoundText = val(Trim(FG.TextMatrix(FG.Row, FG.ColIndex("CustomerID"))))
      
  FrmWaiverSettlement.TxtContNo.Text = val(FG.TextMatrix(FG.Row, FG.ColIndex("ContNo")))
    FrmWaiverSettlement.RetriveOrder val(FG.TextMatrix(FG.Row, FG.ColIndex("ContNo"))), 1
    
   FrmWaiverSettlement.GetContract (FG.TextMatrix(FG.Row, FG.ColIndex("NoteSerial"))), val(FG.TextMatrix(FG.Row, FG.ColIndex("ContNo")))
    ElseIf m_RetrunType = 3 Then
   FrmRsCustomerAlarm.RetriveOrder val(FG.TextMatrix(FG.Row, FG.ColIndex("ContNo"))), 1
    ElseIf m_RetrunType = 4 Then
    FrmWaiver.TxtContNo.Text = val(FG.TextMatrix(FG.Row, FG.ColIndex("ContNo")))
 End If
ErrTrap:
End Sub

Private Sub Retrive()
    Dim Num As Integer
    On Error GoTo ErrTrap
    FG.Clear flexClearScrollable, flexClearEverything

    If Not (rs.EOF Or rs.BOF) Then
        FG.Rows = rs.RecordCount + 1

        For Num = 1 To rs.RecordCount

            With FG
            
            DcbFurnishing.ListIndex = val(IIf(IsNull(rs("Furnishing").value), -1, rs("Furnishing").value))
                .TextMatrix(Num, .ColIndex("Furnishing")) = DcbFurnishing.Text
                .TextMatrix(Num, .ColIndex("Aqarid")) = IIf(IsNull(rs("Iqar").value), "", rs("Iqar").value)
                .TextMatrix(Num, .ColIndex("UnitTypeID")) = IIf(IsNull(rs("UnitTypeID").value), "", rs("UnitTypeID").value)
                .TextMatrix(Num, .ColIndex("UnitNo")) = IIf(IsNull(rs("UnitNo").value), "", rs("UnitNo").value)
                .TextMatrix(Num, .ColIndex("CustomerID")) = IIf(IsNull(rs("CusID").value), "", rs("CusID").value)
                
     
                
                .TextMatrix(Num, .ColIndex("ContNo")) = IIf(IsNull(rs("ContNo").value), "", rs("ContNo").value)
               
                '
                If val(rs!NoteSerial1 & "") = 0 Then
                
                    .TextMatrix(Num, .ColIndex("NoteSerial")) = IIf(IsNull(rs("ContNo").value), "", rs("ContNo").value)
                Else
                    .TextMatrix(Num, .ColIndex("NoteSerial")) = IIf(IsNull(rs("NoteSerial1").value), "", rs("NoteSerial1").value)
                End If
                DcbContType.ListIndex = val(IIf(IsNull(rs("ContType").value), -1, rs("ContType").value))
                .TextMatrix(Num, .ColIndex("ContType")) = DcbContType.Text
                
                .TextMatrix(Num, .ColIndex("UnitType")) = IIf(IsNull(rs("name").value), "", rs("name").value)
                DcbRentType.ListIndex = val(IIf(IsNull(rs("RentType").value), -1, rs("RentType").value))
                .TextMatrix(Num, .ColIndex("RentType")) = DcbRentType.Text
            .TextMatrix(Num, .ColIndex("ContDate")) = IIf(IsNull(rs("ContDate").value), "", Trim(rs("ContDate").value))
            .TextMatrix(Num, .ColIndex("StrDate")) = IIf(IsNull(rs("StrDate").value), "", Trim(rs("StrDate").value))
                    
        .TextMatrix(Num, .ColIndex("FromdateH")) = IIf(IsNull(rs("FromdateH").value), "", Trim(rs("FromdateH").value))
            .TextMatrix(Num, .ColIndex("TodateH")) = IIf(IsNull(rs("TodateH").value), "", Trim(rs("TodateH").value))
                      
                      
                    
                    .TextMatrix(Num, .ColIndex("EndDate")) = IIf(IsNull(rs("EndDate").value), "", Trim(rs("EndDate").value))
                .TextMatrix(Num, .ColIndex("aqarname")) = IIf(IsNull(rs("aqarname").value), "", Trim(rs("aqarname").value))
              
 .TextMatrix(Num, .ColIndex("unitnoName")) = IIf(IsNull(rs("nameunitno").value), 0, Trim(rs("nameunitno").value))
                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(Num, .ColIndex("CusName")) = IIf(IsNull(rs("CusName").value), "", Trim(rs("CusName").value))
                Else
                    .TextMatrix(Num, .ColIndex("CusName")) = IIf(IsNull(rs("CusNamee").value), "", Trim(rs("CusNamee").value))
                End If
.TextMatrix(Num, .ColIndex("MeterValue")) = IIf(IsNull(rs("MeterValue").value), "", Trim(rs("MeterValue").value))
           
                    .TextMatrix(Num, .ColIndex("MeterCount")) = IIf(IsNull(rs("MeterCount").value), "", Trim(rs("MeterCount").value))
               
                    .TextMatrix(Num, .ColIndex("Water")) = IIf(IsNull(rs("Water").value), "", Trim(rs("Water").value))
           
                    .TextMatrix(Num, .ColIndex("Electricity")) = IIf(IsNull(rs("Electricity").value), "", Trim(rs("Electricity").value))
          .TextMatrix(Num, .ColIndex("Phone")) = IIf(IsNull(rs("Phone").value), "", Trim(rs("Phone").value))
          
           .TextMatrix(Num, .ColIndex("Enternet")) = IIf(IsNull(rs("Enternet").value), "", Trim(rs("Enternet").value))
            .TextMatrix(Num, .ColIndex("IncresYearRate")) = IIf(IsNull(rs("IncresYearRate").value), "", Trim(rs("IncresYearRate").value))
             .TextMatrix(Num, .ColIndex("IncresYearValue")) = IIf(IsNull(rs("IncresYearValue").value), "", Trim(rs("IncresYearValue").value))
              .TextMatrix(Num, .ColIndex("PaymentCount")) = IIf(IsNull(rs("PaymentCount").value), "", Trim(rs("PaymentCount").value))
               .TextMatrix(Num, .ColIndex("FristPaymentDate")) = IIf(IsNull(rs("FristPaymentDate").value), "", Trim(rs("FristPaymentDate").value))
             '   .TextMatrix(Num, .ColIndex("IncresYearValue")) = IIf(IsNull(rs("IncresYearValue").value), "", Trim(rs("IncresYearValue").value))
                

                .TextMatrix(Num, .ColIndex("TotalContract")) = IIf(IsNull(rs("TotalContract").value), "", Trim(rs("TotalContract").value))
                   .TextMatrix(Num, .ColIndex("PayAmini")) = IIf(IsNull(rs("PayAmini").value), "", Trim(rs("PayAmini").value))
           
                  .TextMatrix(Num, .ColIndex("CommiValue")) = IIf(IsNull(rs("CommiValue").value), "", (rs("CommiValue").value))
                    .TextMatrix(Num, .ColIndex("InsuranceValue")) = IIf(IsNull(rs("InsuranceValue").value), "", Trim(rs("InsuranceValue").value))
            
            End With

            rs.MoveNext
        Next Num

        ' Fg.AutoSize 0, Fg.Cols - 1, False
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Fg_DblClick()
    fg_Click
    Unload Me
End Sub

Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim StrSQL As String
    Dim BG As New ClsBackGroundPic
    Dim Dcombos As ClsDataCombos

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
    StrDate.value = ""
    XPDtbBill.value = ""
    Opt(1).value = True
EndDate.value = ""
FristPaymentDate.value = ""
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Clear").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
 
    Dim My_SQL As String
    Set Dcombos = New ClsDataCombos
   ' My_SQL = "select UserID,UserName From tblUsers "
   ' fill_combo DCUser, My_SQL
    Set Dcombos = New ClsDataCombos
    Dcombos.GetCustomersSuppliers 56, Me.dcCustomer
    Dcombos.GetCustomersSuppliers 57, Me.dcsupplier
    Dcombos.GetIqar DcbIqara
    Dcombos.getAkarUnit Me.DcbUnitType
    Dcombos.GetIqarUnit -2, 1, DcbUnitNo
 ' Dcombos.GetIqarUnit 1, DcbUnitNo
'  Dcombos.GetIqarUnit 1, DcbUnitNo1
    Set cSearch = New clsDCboSearch
   ' My_SQL = " select CountryID,CountryName from TblCountriesData"
 
   ' fill_combo Me.DataCombo4, My_SQL
   ' RetrunType = -1
 
    CenterForm Me

    FormPostion Me, GetPostion
    FG.WallPaper = BG.SearchWallpaper
    Set rs = New ADODB.Recordset
    DBCboClientName.BoundText = ""
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

 StrSQL = " SELECT DISTINCT "
  StrSQL = StrSQL & "                     dbo.TblContract.FromdateH, dbo.TblContract.TodateH, dbo.TblContract.NoteSerial1, dbo.TblContract.ContNo, dbo.TblContract.UnitType, dbo.TblContract.ContType,"
 StrSQL = StrSQL & "                      dbo.TblContract.ContDate, dbo.TblContract.Iqar, dbo.TblAqar.aqarname, dbo.TblAqar.aqarNo, dbo.TblAqar.Aqarid, dbo.TblContract.ownerid,"
 StrSQL = StrSQL & "                      TblCustemers_1.CusName AS CusNameOw, TblCustemers_1.CusNamee AS CusNameeOw, TblCustemers_1.Fullcode AS FullcodeOw, dbo.TblContract.RentType,"
 StrSQL = StrSQL & "                      dbo.TblContract.StrDate, dbo.TblContract.EndDate, dbo.TblContract.MeterValue, dbo.TblContract.MeterCount, dbo.TblContract.TotalContract, dbo.TblContract.PayAmini,"
StrSQL = StrSQL & "                       dbo.TblContract.CommiValue, dbo.TblContract.InsuranceValue, dbo.TblContract.Water, dbo.TblContract.Electricity, dbo.TblContract.Phone, dbo.TblContract.Enternet,"
 StrSQL = StrSQL & "                      dbo.TblContract.IncresYearValue, dbo.TblContract.IncresYearRate, dbo.TblContract.PaymentCount, dbo.TblContract.FristPaymentDate, dbo.TblContract.PeriodsID,"
 StrSQL = StrSQL & "                      dbo.TblContract.Periods, dbo.TblContract.Furnishing, TblCustemers_1.CusName, TblCustemers_1.CusNamee, TblCustemers_1.CusID, TblCustemers_1.Fullcode,"
 StrSQL = StrSQL & "                      dbo.TblContract.Remarks, dbo.TblAkarUnit.name, dbo.TblAkarUnit.namee, dbo.TblContract.UnitNo, dbo.TblAqarDetai.unitno AS nameunitno,dbo.TblContract.EndContract,"
  StrSQL = StrSQL & "                      TblContract.Iqar , TblContract.unittype UnitTypeID, TblContract.unitno, TblContract.CusID "
StrSQL = StrSQL & "  FROM         dbo.TblContract LEFT OUTER JOIN"
 StrSQL = StrSQL & "                      dbo.TblAqarDetai ON dbo.TblContract.UnitNo = dbo.TblAqarDetai.Id LEFT OUTER JOIN"
 StrSQL = StrSQL & "                      dbo.TblAkarUnit ON dbo.TblContract.UnitType = dbo.TblAkarUnit.id LEFT OUTER JOIN"
  StrSQL = StrSQL & "                     dbo.TblCustemers TblCustemers_1 ON dbo.TblContract.CusID = TblCustemers_1.CusID LEFT OUTER JOIN"
  StrSQL = StrSQL & "                     dbo.TblCustemers TblCustemers_2 ON dbo.TblContract.ownerid = TblCustemers_2.CusID LEFT OUTER JOIN"
   StrSQL = StrSQL & "                    dbo.TblAqar ON dbo.TblContract.Iqar = dbo.TblAqar.Aqarid"
If SystemOptions.usertype = UserAdminAll Then
    StrSQL = StrSQL & "   where 1=1   "
Else
StrSQL = StrSQL & "   where 1=1   and Branch_NO=" & Current_branch
End If

    If Opt(1).value = True Then
    StrWhere = StrWhere + " and (dbo.TblContract.EndContract IS NULL) "
    ElseIf Opt(2).value = True Then
    StrWhere = StrWhere + " and (dbo.TblContract.EndContract =1 )"
    End If
    
    If m_RetrunType = 2 Then
    StrWhere = StrWhere + " and (dbo.TblContract.EndContract IS NULL) "
    End If
     If Me.DcbRentType.ListIndex <> -1 And Me.DcbRentType.Text <> "" Then
         StrWhere = StrWhere + " and dbo.TblContract.RentType =" & Me.DcbRentType.ListIndex & ""
    End If
    
    If val(DcbUnitType.BoundText) <> 0 And DcbUnitType.Text <> "" Then
 
        StrWhere = StrWhere + " and dbo.TblContract.UnitType =" & DcbUnitType.BoundText & ""
 
    End If
    
    
    If Me.DcbContType.ListIndex <> -1 And Me.DcbContType.Text <> "" Then
 
        StrWhere = StrWhere + " and dbo.TblContract.ContType =" & Me.DcbContType.ListIndex & ""
 
    End If
        If Me.DcbFurnishing.ListIndex <> -1 And Me.DcbFurnishing.Text <> "" Then
 
        StrWhere = StrWhere + " and dbo.TblContract.Furnishing =" & Me.DcbFurnishing.ListIndex & ""
 
    End If
   If val(Me.DcbUnitNo.BoundText) <> 0 And Me.DcbUnitNo.Text <> "" Then
 
        StrWhere = StrWhere + " and dbo.TblAqarDetai.unitno =N'" & (Me.DcbUnitNo.Text) & "'"
 
    End If
 
    If Me.dcsupplier.BoundText <> "" And Me.dcsupplier.Text <> "" Then
 
        StrWhere = StrWhere + " and dbo.TblContract.ownerid =" & val(Me.dcsupplier.BoundText)
 
    End If
    If Me.dcCustomer.BoundText <> "" And Me.dcCustomer.Text <> "" Then
 
        StrWhere = StrWhere + " and dbo.TblContract.CusID =" & val(Me.dcCustomer.BoundText)
 
    End If
       If Me.DcbIqara.BoundText <> "" And Me.DcbIqara.Text <> "" Then
 
        StrWhere = StrWhere + " and dbo.TblContract.Iqar =" & val(Me.DcbIqara.BoundText)
 
    End If
    
    If Me.TXTOrDer_no.Text <> "" Then
 
        StrWhere = StrWhere + " and dbo.TblContract.NoteSerial1='" & Me.TXTOrDer_no.Text & "'"
 
    End If
If Me.TxtMeterValue.Text <> "" Then
 
        StrWhere = StrWhere + " and dbo.TblContract.MeterValue ='" & Me.TxtMeterValue.Text & "'"
 
    End If
      If Me.TxtMeterCount.Text <> "" Then
 
        StrWhere = StrWhere + " and dbo.TblContract.MeterCount ='" & Me.TxtMeterCount.Text & "'"
 
    End If
      If Me.TxtTotalContract.Text <> "" Then
 
        StrWhere = StrWhere + " and dbo.TblContract.TotalContract ='" & Me.TxtTotalContract.Text & "'"
 
    End If
     If Me.TxtPayAmini.Text <> "" Then
 
        StrWhere = StrWhere + " and dbo.TblContract.PayAmini ='" & Me.TxtPayAmini.Text & "'"
 
    End If
       If Me.TxtCommiValue.Text <> "" Then
 
        StrWhere = StrWhere + " and dbo.TblContract.CommiValue ='" & Me.TxtCommiValue.Text & "'"
 
    End If
           If Me.TxtInsuranceValue.Text <> "" Then
 
        StrWhere = StrWhere + " and dbo.TblContract.InsuranceValue ='" & Me.TxtInsuranceValue.Text & "'"
 
    End If
            If Me.TxtWater.Text <> "" Then
 
        StrWhere = StrWhere + " and dbo.TblContract.Water ='" & Me.TxtWater.Text & "'"
 
    End If
                If Me.TxtElectricity.Text <> "" Then
 
        StrWhere = StrWhere + " and dbo.TblContract.Electricity ='" & Me.TxtElectricity.Text & "'"
 
    End If
                    If Me.TxtPhone.Text <> "" Then
         StrWhere = StrWhere + " and dbo.TblContract.Phone ='" & Me.TxtPhone.Text & "'"
     End If
                         If Me.TxtEnternet.Text <> "" Then
         StrWhere = StrWhere + " and dbo.TblContract.Enternet ='" & Me.TxtEnternet.Text & "'"
     End If
                           If Me.TxtIncresYearRate.Text <> "" Then
         StrWhere = StrWhere + " and dbo.TblContract.IncresYearRate ='" & Me.TxtIncresYearRate.Text & "'"
     End If
                             If Me.TxtIncresYearValue.Text <> "" Then
         StrWhere = StrWhere + " and dbo.TblContract.IncresYearValue ='" & Me.TxtIncresYearValue.Text & "'"
     End If
                             If Me.TxtPeriods.Text <> "" Then
         StrWhere = StrWhere + " and dbo.TblContract.Periods ='" & Me.TxtPeriods.Text & "'"
     End If
                                If Me.TxtPaymentCount.Text <> "" Then
         StrWhere = StrWhere + " and dbo.TblContract.PaymentCount ='" & Me.TxtPaymentCount.Text & "'"
     End If
     If Not IsNull(Me.FristPaymentDate.value) Then
    
            StrWhere = StrWhere & " AND dbo.TblContract.FristPaymentDate >=" & SQLDate(Me.FristPaymentDate.value, True) & ""
   End If
         If Not IsNull(Me.EndDate.value) Then
    
            StrWhere = StrWhere & " AND dbo.TblContract.EndDate <=" & SQLDate(Me.EndDate.value, True) & ""
   End If
       If Not IsNull(Me.StrDate.value) Then
    
            StrWhere = StrWhere & " AND dbo.TblContract.StrDate >=" & SQLDate(Me.StrDate.value, True) & ""
   End If
       If Not IsNull(Me.XPDtbBill.value) Then
    
            StrWhere = StrWhere & " AND dbo.TblContract.ContDate >=" & SQLDate(Me.XPDtbBill.value, True) & ""
   End If
    StrWhere = StrWhere + " order by dbo.TblContract.ContNo"

    Build_Sql = StrSQL + StrWhere
    Exit Function
ErrTrap:
End Function

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        If Me.ActiveControl Is FG Then
            If Not FG.TextMatrix(FG.Row, 1) = "" Then
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
    Me.Caption = "Search For Production Orders"
    Label1(2).Caption = Me.Caption
    Label2.Caption = "Order No"
 
    Label3.Caption = "Date"
    Label5.Caption = "Country"
    Label4.Caption = "Vendor"
    Label6.Caption = "Remark"

    Cmd(0).Caption = "Search"
    Cmd(1).Caption = "Clear"
    Cmd(2).Caption = "Exit"

    'OptType(0).Caption = "Start of the name"
    'OptType(1).Caption = "any part of the name"
    With Me.FG
        .TextMatrix(0, .ColIndex("order_no")) = "order no"
        '  .TextMatrix(0, .ColIndex("remark")) = "remark  "
        .TextMatrix(0, .ColIndex("CusName")) = "Customer Name"
        .TextMatrix(0, .ColIndex("Transaction_Date")) = " Date"
        '     .TextMatrix(0, .ColIndex("CountryName")) = "Country Name"
  
        '  .AutoSize 0, .Cols - 1, False
    End With

End Sub



Private Sub Opt_Click(Index As Integer)
Build_Sql
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
 Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetTblCustemersCode Text1.Text, EmpID, , , 57
        dcsupplier.BoundText = EmpID
    End If
End Sub

Private Sub Text15_KeyPress(KeyAscii As Integer)
Dim EmpID As Double

    If KeyAscii = vbKeyReturn Then
        GetTblCustemersCode Text15.Text, EmpID, , , 56
        dcCustomer.BoundText = EmpID
    End If
End Sub

Private Sub TxtSearch_KeyPress(KeyAscii As Integer)
 Dim EmpID As Double

    If KeyAscii = vbKeyReturn Then
        GetIqarCode TxtSearch.Text, EmpID
        DcbIqara.BoundText = EmpID
    End If
End Sub
