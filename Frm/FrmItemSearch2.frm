VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Begin VB.Form FrmItemSearch2 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«·»ÕÀ ⁄‰ ’‰›"
   ClientHeight    =   7080
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10560
   Icon            =   "FrmItemSearch2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   7080
   ScaleWidth      =   10560
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.Frame Frame3 
      Height          =   1935
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   31
      Top             =   4560
      Width           =   10455
      Begin VB.TextBox txtItemDetailedCode 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   5520
         RightToLeft     =   -1  'True
         TabIndex        =   44
         Top             =   720
         Width           =   2865
      End
      Begin VB.TextBox TxtItemName 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   240
         Width           =   3945
      End
      Begin VB.TextBox TxtbarCodeNO 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   5520
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   240
         Width           =   2865
      End
      Begin MSDataListLib.DataCombo DCCOLOR 
         Height          =   315
         Left            =   7440
         TabIndex        =   37
         Top             =   1200
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DCSIZE 
         Height          =   315
         Left            =   5520
         TabIndex        =   39
         Top             =   1200
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DCcLASS 
         Height          =   315
         Left            =   3000
         TabIndex        =   41
         Top             =   1200
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "ﬂÊœ «·’‰›"
         Height          =   345
         Index           =   17
         Left            =   9000
         RightToLeft     =   -1  'True
         TabIndex        =   45
         Top             =   840
         Width           =   1035
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·›—“"
         Height          =   285
         Index           =   16
         Left            =   4080
         RightToLeft     =   -1  'True
         TabIndex        =   42
         Top             =   1320
         Width           =   915
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·„ﬁ«”"
         Height          =   285
         Index           =   15
         Left            =   6360
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   1320
         Width           =   915
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«··Ê‰"
         Height          =   285
         Index           =   14
         Left            =   9120
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   1320
         Width           =   915
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«”„ «·’‰›"
         Height          =   345
         Index           =   1
         Left            =   4200
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   360
         Width           =   915
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·»«—ﬂÊœ"
         Height          =   345
         Index           =   12
         Left            =   9120
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   360
         Width           =   915
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Height          =   195
         Index           =   13
         Left            =   5310
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   360
         Width           =   45
      End
   End
   Begin VB.TextBox TxtPartNo 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   10920
      RightToLeft     =   -1  'True
      TabIndex        =   29
      Top             =   3720
      Width           =   1545
   End
   Begin VB.ComboBox CboItemCodeSearch 
      Height          =   315
      Left            =   12750
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   3630
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
      TabIndex        =   27
      Top             =   8070
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
      TabIndex        =   26
      Top             =   7620
      Width           =   6495
   End
   Begin VB.ComboBox CboArchive 
      Height          =   315
      Left            =   11070
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   4140
      Width           =   1335
   End
   Begin VB.ComboBox CboGuar 
      Height          =   315
      Left            =   13200
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   4170
      Width           =   1305
   End
   Begin VB.ComboBox CboNameSearch 
      Height          =   315
      Left            =   12750
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   4020
      Width           =   1515
   End
   Begin VB.ComboBox CboAttachedItem 
      Height          =   315
      Left            =   11070
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   4500
      Width           =   1335
   End
   Begin VB.ComboBox CboAssbliedItem 
      Height          =   315
      Left            =   13200
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   4500
      Width           =   1305
   End
   Begin VB.ComboBox CboItemType 
      Height          =   315
      Left            =   15420
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   4170
      Width           =   1215
   End
   Begin VB.TextBox TxtItemID 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   13260
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   2670
      Width           =   735
   End
   Begin VB.ComboBox CboSerial 
      Height          =   315
      Left            =   12750
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   4860
      Width           =   1515
   End
   Begin VB.TextBox XPTxtItemCode 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   11010
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   2655
      Width           =   1395
   End
   Begin VSFlex8UCtl.VSFlexGrid FgX 
      Height          =   4185
      Left            =   9600
      TabIndex        =   0
      Top             =   2760
      Visible         =   0   'False
      Width           =   10515
      _cx             =   18547
      _cy             =   7382
      Appearance      =   0
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
      Cols            =   15
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmItemSearch2.frx":030A
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
      Left            =   5520
      TabIndex        =   12
      Top             =   6555
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
      Left            =   4530
      TabIndex        =   13
      Top             =   6555
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
      Left            =   3630
      TabIndex        =   14
      Top             =   6555
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
      Left            =   9000
      TabIndex        =   5
      Top             =   7080
      Width           =   2985
      _ExtentX        =   5265
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VSFlex8UCtl.VSFlexGrid Fg 
      Height          =   4215
      Left            =   0
      TabIndex        =   43
      Top             =   0
      Width           =   10665
      _cx             =   18812
      _cy             =   7435
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
      Rows            =   15
      Cols            =   15
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmItemSearch2.frx":0537
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
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "—ﬁ„ «·ﬁÿ⁄Â/«·„ÊœÌ·"
      Height          =   615
      Left            =   12480
      RightToLeft     =   -1  'True
      TabIndex        =   30
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "„Ã«· «·»ÕÀ"
      Height          =   345
      Index           =   11
      Left            =   14310
      RightToLeft     =   -1  'True
      TabIndex        =   28
      Top             =   3630
      Width           =   975
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·√—‘Ì›"
      Height          =   285
      Index           =   10
      Left            =   12450
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   4170
      Width           =   705
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·÷„«‰"
      Height          =   285
      Index           =   9
      Left            =   14520
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   4170
      Width           =   885
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "„Ã«· «·»ÕÀ"
      Height          =   345
      Index           =   8
      Left            =   14310
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   4020
      Width           =   975
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "„·Õﬁ"
      Height          =   315
      Index           =   7
      Left            =   12450
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   4500
      Width           =   705
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " Ã„Ì⁄"
      Height          =   285
      Index           =   6
      Left            =   14520
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   4500
      Width           =   915
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "‰Ê⁄ «·’‰›"
      Height          =   285
      Index           =   5
      Left            =   16680
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   4200
      Width           =   915
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·’‰›"
      Height          =   345
      Index           =   4
      Left            =   14040
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   2670
      Width           =   915
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "‰Ÿ«„ «·”Ì—Ì«·"
      Height          =   315
      Index           =   2
      Left            =   14310
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   4890
      Width           =   975
   End
   Begin VB.Label LblRes 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      ForeColor       =   &H000000C0&
      Height          =   315
      Left            =   15660
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   4530
      Width           =   1905
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "ﬂÊœ «·’‰›"
      Height          =   345
      Index           =   0
      Left            =   12420
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   2670
      Width           =   795
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·„Ã„Ê⁄…"
      Height          =   285
      Index           =   3
      Left            =   10080
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   7200
      Width           =   915
   End
End
Attribute VB_Name = "FrmItemSearch2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim cSearchDcbo As clsDCboSearch

Private m_DcboItems As DataCombo

Private m_RetrunType As Integer

Private Sub Cmd_Click(Index As Integer)
    On Error GoTo ErrTrap

    Select Case Index

        Case 0

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
                fg.Clear flexClearScrollable, flexClearEverything
                fg.Rows = 2

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
            fg.SetFocus

        Case 1
            clear_all Me
            fg.Clear flexClearScrollable, flexClearEverything

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

Private Sub fg_Click()
    On Error GoTo ErrTrap

    If Not fg.TextMatrix(fg.Row, 3) = "" Then
                           If Me.RetrunType = 0 Then
                        
                                
                               ElseIf Me.RetrunType = 1 Then
                       frmsalebill2.TxtItemCodeB.Text = (fg.TextMatrix(fg.Row, 3))
                       
                         'frmsalebill2.TxtItemCodeB_KeyDown (vbKeyReturn), 0
                         
                      ' frmsalebill2.TxtItemCodeB_KeyDown 13, 0
                               Unload Me
                           ElseIf Me.RetrunType = 2 Then
                    
                       FrmReturnSalling.TxtItemCodeB.Text = (fg.TextMatrix(fg.Row, 3))
                               Unload Me
                           
                           
                                        ElseIf Me.RetrunType = 3 Then
                    
                       FrmItemsDetails.TxtItemCodeB.Text = (fg.TextMatrix(fg.Row, 3))
                       FrmItemsDetails.Cmd_Click 20
                            Unload Me
                            
                            
                                                   ElseIf Me.RetrunType = 4 Then
                   FrmNewGard.TxtItemCodeB.Text = (fg.TextMatrix(fg.Row, 3))
                               Unload Me
                            
                               
                           End If
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Retrive()
    Dim Num As Integer
    On Error GoTo ErrTrap
    fg.Clear flexClearScrollable, flexClearEverything

    If Not (rs.EOF Or rs.BOF) Then
        fg.Rows = rs.RecordCount + 1

        For Num = 1 To rs.RecordCount

            With fg
                .TextMatrix(Num, .ColIndex("NumIndex")) = Num
                 If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(Num, .ColIndex("ItemName")) = IIf(IsNull(rs("ItemName").value), "", Trim(rs("ItemName").value))
        Else
        .TextMatrix(Num, .ColIndex("ItemName")) = IIf(IsNull(rs("ItemNamee").value), "", Trim(rs("ItemNamee").value))
        
        End If
        
        
             .TextMatrix(Num, .ColIndex("ColorName")) = IIf(IsNull(rs("ColorName").value), "", (rs("ColorName").value))
                .TextMatrix(Num, .ColIndex("ItemSize")) = IIf(IsNull(rs("SizeName").value), "", (rs("SizeName").value))
                .TextMatrix(Num, .ColIndex("ClassName")) = IIf(IsNull(rs("cclASS NAME").value), "", (rs("cclASS NAME").value))
                
                 
            .TextMatrix(Num, .ColIndex("ParrtNoCode")) = IIf(IsNull(rs("ParrtNoCode").value), "", (rs("ParrtNoCode").value))
            .TextMatrix(Num, .ColIndex("ItemDetailedCode")) = IIf(IsNull(rs("ItemDetailedCode").value), "", (rs("ItemDetailedCode").value))
          
          
            
            End With

            rs.MoveNext
        Next Num

        fg.AutoSize 0, fg.Cols - 1, False
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

    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Clear").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture

    Set Dcombos = New ClsDataCombos
    Dcombos.GetItemSGroups Me.DCboGroupName
     Dcombos.GetItemsColors Me.DCCOLOR
    Dcombos.GetItemsSizes Me.DCSIZE
    Dcombos.GetItemsClasses Me.dcClass
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
            .AddItem "·Â «’‰«› „·Õﬁ…"
            .AddItem "·Ì” ·Â «’‰«› „·Õﬁ…"
            .AddItem "«·ﬂ·"
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
            .AddItem "ALL"
        End With

    End If

    CenterForm Me
 
    FormPostion Me, GetPostion
    fg.WallPaper = BG.SearchWallpaper
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

    StrSQL = "SELECT   dbo.ItemsDetails.ItemDetailedCode, dbo.ItemsDetails.ParrtNoCode, dbo.TblItemsclasses.SizeName AS [cclASS NAME], dbo.TblItemsColors.ColorName, dbo.TblItemsSizes.SizeName, "
StrSQL = StrSQL + "     dbo.ItemsDetails.ProductionDate , dbo.ItemsDetails.ExpireDate, dbo.TblItems.ItemCode, dbo.TblItems.itemname, dbo.TblItems.ItemNamee"
StrSQL = StrSQL + " FROM         dbo.ItemsDetails INNER JOIN"
StrSQL = StrSQL + "    dbo.Transactions ON dbo.ItemsDetails.Transaction_ID = dbo.Transactions.Transaction_ID INNER JOIN"
StrSQL = StrSQL + "      dbo.TransactionTypes ON dbo.Transactions.Transaction_Type = dbo.TransactionTypes.Transaction_Type INNER JOIN"
StrSQL = StrSQL + "      dbo.TblStore ON dbo.Transactions.StoreID = dbo.TblStore.StoreID LEFT OUTER JOIN"
StrSQL = StrSQL + "         dbo.TblItems ON dbo.ItemsDetails.ItemId = dbo.TblItems.ItemID LEFT OUTER JOIN"
StrSQL = StrSQL + "      dbo.TblUnites ON dbo.ItemsDetails.UnitID = dbo.TblUnites.UnitID LEFT OUTER JOIN"
  StrSQL = StrSQL + "     dbo.TblItemsColors ON dbo.ItemsDetails.ColorID = dbo.TblItemsColors.ColorID LEFT OUTER JOIN"
StrSQL = StrSQL + "       dbo.TblItemsSizes ON dbo.ItemsDetails.SizeID = dbo.TblItemsSizes.SizeId LEFT OUTER JOIN"
StrSQL = StrSQL + "      dbo.TblItemsclasses ON dbo.ItemsDetails.ClassId = dbo.TblItemsclasses.SizeId"
 StrSQL = StrSQL + " Where  1=1"

  If SystemOptions.WorkWithLINKEDiActivity = True Then
    StrSQL = StrSQL & "  and dbo.TblItems.GroupID in(   "
     StrSQL = StrSQL & " select GroupID from fullgroups ()  )"
 End If
   
   
   If TxtbarCodeNO.Text <> "" Then
       
            StrSQL = StrSQL + " and ItemsDetails.ParrtNoCode like '%" & Trim(TxtbarCodeNO.Text) & "%'"
       
    End If
 
   If txtItemName.Text <> "" Then
       
            StrSQL = StrSQL + " and TblItems.ItemName like '%" & Trim(txtItemName.Text) & "%'"
       
    End If
    
    
   If txtItemDetailedCode.Text <> "" Then
       
            StrSQL = StrSQL + " and ItemsDetails.ItemDetailedCode like '%" & Trim(txtItemDetailedCode.Text) & "%'"
       
    End If
    
    

    
    If Me.DCCOLOR.BoundText <> "" And Me.DCCOLOR.Text <> "" Then
        StrSQL = StrSQL + " and ItemsDetails.ColorID =" & val(Me.DCCOLOR.BoundText & "")
    End If
    
    
    If Me.DCSIZE.BoundText <> "" And Me.DCSIZE.Text <> "" Then
        StrSQL = StrSQL + " and ItemsDetails.SizeID =" & val(Me.DCSIZE.BoundText & "")
    End If
    
    If Me.dcClass.BoundText <> "" And Me.dcClass.Text <> "" Then
        StrSQL = StrSQL + " and ItemsDetails.ClassId =" & val(Me.dcClass.BoundText & "")
    End If
    
    
StrSQL = StrSQL + " GROUP BY dbo.ItemsDetails.ParrtNoCode, dbo.ItemsDetails.ColorID, dbo.ItemsDetails.UnitID, dbo.ItemsDetails.SizeID, dbo.ItemsDetails.ClassId,"
StrSQL = StrSQL + "   dbo.TblItemsclasses.SizeName, dbo.TblItemsColors.ColorName, dbo.TblItemsSizes.SizeName, dbo.ItemsDetails.ProductionDate, dbo.ItemsDetails.ExpireDate,"
StrSQL = StrSQL + "      dbo.TblItems.ItemCode , dbo.TblItems.itemname, dbo.TblItems.ItemNamee, dbo.ItemsDetails.ItemDetailedCode"

 
  
  
    Build_Sql = StrSQL
    Debug.Print StrSQL
    Exit Function
ErrTrap:
End Function

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        If Me.ActiveControl Is fg Then
            If Not fg.TextMatrix(fg.Row, 1) = "" Then
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
    Me.Caption = "Search For Item"
 
        lbl(12).Caption = "BarCode"
              lbl(1).Caption = "Name"
 
     lbl(14).Caption = "Color"
 lbl(15).Caption = "Size"
 lbl(16).Caption = "Class"
  lbl(17).Caption = "Item Code"
  
    Cmd(0).Caption = "Search"
    Cmd(1).Caption = "Clear"
    Cmd(2).Caption = "Exit"

    'OptType(0).Caption = "Start of the name"
    'OptType(1).Caption = "any part of the name"
    With Me.fg
  

        .TextMatrix(0, .ColIndex("NumIndex")) = "Serial"
        .TextMatrix(0, .ColIndex("ParrtNoCode")) = "Bar Code"
        .TextMatrix(0, .ColIndex("ItemDetailedCode")) = "Item Detailed Code"
        .TextMatrix(0, .ColIndex("ItemName")) = "Item Name"
        
        .TextMatrix(0, .ColIndex("ItemSize")) = "ItemSize "
        .TextMatrix(0, .ColIndex("ColorName")) = "Color "
         .TextMatrix(0, .ColIndex("ClassName")) = "Class "
         
        .AutoSize 0, .Cols - 1, False
    End With

End Sub

Private Sub TxtItemName_Change()

    If Trim$(Me.txtItemName.Text) = "" Then
        Me.lbl(8).Enabled = False
        Me.CboNameSearch.Enabled = False
    Else
        Me.lbl(8).Enabled = True
        Me.CboNameSearch.Enabled = True
    End If

End Sub

Private Sub XPTxtItemCode_Change()

    If Trim$(Me.XPTxtItemCode.Text) = "" Then
        Me.lbl(11).Enabled = False
        Me.CboItemCodeSearch.Enabled = False
    Else
        Me.lbl(11).Enabled = True
        Me.CboItemCodeSearch.Enabled = True
    End If

End Sub

