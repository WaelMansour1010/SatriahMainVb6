VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmEmpDepartments 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "«·«œ«—«  Ê «·«ﬁ”«„"
   ClientHeight    =   10110
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10215
   Icon            =   "FrmEmpDepartments.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10110
   ScaleWidth      =   10215
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      BackColor       =   &H00E2E9E9&
      BorderStyle     =   0  'None
      Height          =   2985
      Left            =   0
      TabIndex        =   43
      Top             =   5940
      Width           =   10155
      Begin VSFlex8Ctl.VSFlexGrid Grid 
         Height          =   2355
         Left            =   0
         TabIndex        =   44
         Top             =   120
         Width           =   10515
         _cx             =   18547
         _cy             =   4154
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
         FormatString    =   $"FrmEmpDepartments.frx":058A
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
      Begin ImpulseButton.ISButton Cmd 
         Height          =   225
         Index           =   3
         Left            =   9000
         TabIndex        =   49
         Top             =   2640
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   397
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   "Õ–› ”ÿ— "
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
         ButtonImage     =   "FrmEmpDepartments.frx":0674
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E2E9E9&
      BorderStyle     =   0  'None
      Height          =   945
      Left            =   0
      TabIndex        =   37
      Top             =   4950
      Width           =   10155
      Begin VB.TextBox TxtID 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   0
         MaxLength       =   20
         TabIndex        =   50
         Top             =   0
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.TextBox TxtDeptCode 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   7605
         MaxLength       =   20
         TabIndex        =   45
         Top             =   600
         Width           =   945
      End
      Begin VB.ComboBox Combo2 
         BackColor       =   &H80000018&
         Height          =   315
         ItemData        =   "FrmEmpDepartments.frx":0C0E
         Left            =   2280
         List            =   "FrmEmpDepartments.frx":0C1E
         Style           =   2  'Dropdown List
         TabIndex        =   40
         Top             =   3030
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.TextBox TxtNameE 
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
         TabIndex        =   39
         Top             =   165
         Width           =   3375
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
         Left            =   5160
         TabIndex        =   38
         Top             =   120
         Width           =   3375
      End
      Begin MSDataListLib.DataCombo DcbDeptManger 
         Height          =   315
         Left            =   3600
         TabIndex        =   46
         Top             =   600
         Width           =   3900
         _ExtentX        =   6879
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         BackColor       =   16777215
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin ImpulseButton.ISButton ISButton2 
         Height          =   315
         Left            =   120
         TabIndex        =   48
         ToolTipText     =   "«÷«›… «·»Ì«‰«  «·Ï «·œ« «"
         Top             =   600
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   556
         Caption         =   "«÷«›…"
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
         ButtonImage     =   "FrmEmpDepartments.frx":0C37
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         DisabledImageExtraction=   0
         LowerToggledContent=   0   'False
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„œÌ— «·ﬁ”„"
         Height          =   285
         Index           =   0
         Left            =   8520
         TabIndex        =   47
         Top             =   600
         Width           =   1410
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«”„ «·ﬁ”„ «‰Ã·Ì“Ì"
         Height          =   285
         Index           =   3
         Left            =   3525
         TabIndex        =   42
         Top             =   150
         Width           =   1410
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«”„ «·ﬁ”„ ⁄—»Ì"
         Height          =   285
         Index           =   1
         Left            =   8565
         TabIndex        =   41
         Top             =   150
         Width           =   1410
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2E9E9&
      BorderStyle     =   0  'None
      Height          =   960
      Left            =   0
      TabIndex        =   25
      Top             =   600
      Width           =   10155
      Begin VB.TextBox TxtSerial 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
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
         Left            =   7485
         MaxLength       =   50
         TabIndex        =   32
         Top             =   120
         Width           =   1065
      End
      Begin VB.TextBox TxtVacName 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
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
         Left            =   5175
         TabIndex        =   31
         Top             =   525
         Width           =   3375
      End
      Begin VB.TextBox TxtVacNamee 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
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
         TabIndex        =   30
         Top             =   525
         Width           =   3375
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H80000018&
         Height          =   315
         ItemData        =   "FrmEmpDepartments.frx":7499
         Left            =   2280
         List            =   "FrmEmpDepartments.frx":74A9
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   3030
         Visible         =   0   'False
         Width           =   1005
      End
      Begin MSDataListLib.DataCombo DCBranch 
         Height          =   315
         Left            =   4665
         TabIndex        =   33
         Top             =   105
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777152
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcbSectionID 
         Height          =   315
         Left            =   120
         TabIndex        =   35
         Top             =   150
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777152
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "Ì »⁄ ﬁÿ«⁄"
         Height          =   285
         Index           =   0
         Left            =   2070
         TabIndex        =   36
         Top             =   180
         Width           =   1410
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "Ì »⁄ ›—⁄"
         Height          =   285
         Index           =   52
         Left            =   5805
         TabIndex        =   34
         Top             =   120
         Width           =   1470
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ﬂÊœ "
         Height          =   195
         Index           =   25
         Left            =   8985
         TabIndex        =   29
         Top             =   150
         Width           =   990
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«”„ «·«œ«—… ⁄—»Ì"
         Height          =   285
         Index           =   24
         Left            =   8085
         TabIndex        =   28
         Top             =   510
         Width           =   1890
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«”„ «·«œ«—…  «‰Ã·Ì“Ì"
         Height          =   285
         Index           =   23
         Left            =   3165
         TabIndex        =   27
         Top             =   510
         Width           =   1890
      End
   End
   Begin VB.Frame FraHeader 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   660
      Left            =   -15
      TabIndex        =   14
      Top             =   -90
      Width           =   10185
      Begin VB.TextBox TxtVac_ID 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         Height          =   240
         Left            =   3030
         TabIndex        =   19
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
         TabIndex        =   18
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
         TabIndex        =   15
         Top             =   450
         Visible         =   0   'False
         Width           =   3105
         Begin MSDataListLib.DataCombo DCUser 
            CausesValidation=   0   'False
            Height          =   315
            Left            =   -255
            TabIndex        =   16
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
            TabIndex        =   17
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
               Picture         =   "FrmEmpDepartments.frx":74C2
               Key             =   "CompanyName"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmEmpDepartments.frx":785C
               Key             =   "Ser"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmEmpDepartments.frx":7BF6
               Key             =   "Vac_Name"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmEmpDepartments.frx":7F90
               Key             =   "ShareCount"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmEmpDepartments.frx":832A
               Key             =   "Dis_Count"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmEmpDepartments.frx":86C4
               Key             =   "Bouns"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmEmpDepartments.frx":8A5E
               Key             =   "SharesValue"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmEmpDepartments.frx":8FF8
               Key             =   "BuyValue"
            EndProperty
         EndProperty
      End
      Begin ImpulseButton.ISButton btnLast 
         Height          =   315
         Left            =   90
         TabIndex        =   20
         Top             =   150
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
         ButtonImage     =   "FrmEmpDepartments.frx":9392
         ColorButton     =   14871017
         AcclimateGrayTones=   -1  'True
         DrawFocusRectangle=   0   'False
         DisabledImageExtraction=   0
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnNext 
         Height          =   315
         Left            =   555
         TabIndex        =   21
         Top             =   150
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
         ButtonImage     =   "FrmEmpDepartments.frx":972C
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnPrevious 
         Height          =   315
         Left            =   1155
         TabIndex        =   22
         Top             =   150
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
         ButtonImage     =   "FrmEmpDepartments.frx":9AC6
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnFirst 
         Height          =   315
         Left            =   1620
         TabIndex        =   23
         Top             =   150
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
         ButtonImage     =   "FrmEmpDepartments.frx":9E60
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«·«œ«—«  Ê «·«ﬁ”«„"
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
         Left            =   4365
         TabIndex        =   24
         Top             =   210
         Width           =   5280
      End
   End
   Begin C1SizerLibCtl.C1Elastic EltCont 
      Height          =   1020
      Left            =   60
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   8985
      Width           =   10080
      _cx             =   17780
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
         Left            =   5925
         TabIndex        =   1
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
         ButtonImage     =   "FrmEmpDepartments.frx":A1FA
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton btnSave 
         Height          =   330
         Left            =   4380
         TabIndex        =   2
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
         ButtonImage     =   "FrmEmpDepartments.frx":A594
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton btnModify 
         Height          =   330
         Left            =   5145
         TabIndex        =   3
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
         ButtonImage     =   "FrmEmpDepartments.frx":A92E
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton BtnUndo 
         Height          =   330
         Left            =   3615
         TabIndex        =   4
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
         ButtonImage     =   "FrmEmpDepartments.frx":ACC8
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton btnDelete 
         Height          =   330
         Left            =   2850
         TabIndex        =   5
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
         ButtonImage     =   "FrmEmpDepartments.frx":B062
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton btnQuery 
         Height          =   330
         Left            =   5880
         TabIndex        =   6
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
         ButtonImage     =   "FrmEmpDepartments.frx":B5FC
         ColorButton     =   14737632
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton BtnUpdate 
         Height          =   330
         Left            =   6045
         TabIndex        =   7
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
         ButtonImage     =   "FrmEmpDepartments.frx":B996
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton BtnPrint 
         Height          =   285
         Left            =   4725
         TabIndex        =   8
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
         ButtonImage     =   "FrmEmpDepartments.frx":BD30
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnCancel 
         Height          =   330
         Left            =   2055
         TabIndex        =   9
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
         ButtonImage     =   "FrmEmpDepartments.frx":C0CA
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
         Left            =   3465
         TabIndex        =   13
         Top             =   225
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "⁄œœ «·”Ã·« :"
         Height          =   210
         Index           =   1
         Left            =   1170
         TabIndex        =   12
         Top             =   225
         Width           =   975
      End
      Begin VB.Label LabCurrRec 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Height          =   210
         Left            =   2640
         TabIndex        =   11
         Top             =   240
         Width           =   675
      End
      Begin VB.Label LabCountRec 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Height          =   210
         Left            =   360
         TabIndex        =   10
         Top             =   225
         Width           =   540
      End
   End
   Begin C1SizerLibCtl.C1Tab C1Tab1 
      Height          =   4725
      Left            =   30
      TabIndex        =   51
      Top             =   120
      Width           =   9975
      _cx             =   17595
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
      Appearance      =   2
      MousePointer    =   0
      Version         =   801
      BackColor       =   14871017
      ForeColor       =   -2147483630
      FrontTabColor   =   -2147483633
      BackTabColor    =   14871017
      TabOutlineColor =   -2147483632
      FrontTabForeColor=   -2147483630
      Caption         =   "«·»Ì«‰«  «·«”«”Ì…|Õ”«»«  «·—»ÿ"
      Align           =   0
      CurrTab         =   1
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
         Height          =   4350
         Left            =   -10530
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   45
         Width           =   9885
         _cx             =   17436
         _cy             =   7673
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
         Begin VB.Frame Frm2 
            BackColor       =   &H00E2E9E9&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   3330
            Left            =   -150
            TabIndex        =   56
            Top             =   1740
            Width           =   10155
            Begin VB.ComboBox CmbType 
               BackColor       =   &H80000018&
               Height          =   315
               ItemData        =   "FrmEmpDepartments.frx":C464
               Left            =   2280
               List            =   "FrmEmpDepartments.frx":C474
               Style           =   2  'Dropdown List
               TabIndex        =   60
               Top             =   3030
               Visible         =   0   'False
               Width           =   1005
            End
            Begin VB.TextBox TxtCode 
               Alignment       =   1  'Right Justify
               BackColor       =   &H80000018&
               ForeColor       =   &H00000000&
               Height          =   315
               IMEMode         =   3  'DISABLE
               Left            =   8040
               MaxLength       =   20
               TabIndex        =   59
               Top             =   30
               Width           =   945
            End
            Begin VB.TextBox TxtCode1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H80000018&
               ForeColor       =   &H00000000&
               Height          =   315
               IMEMode         =   3  'DISABLE
               Left            =   8040
               MaxLength       =   20
               TabIndex        =   58
               Top             =   390
               Width           =   945
            End
            Begin VB.TextBox TxtCode2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H80000018&
               ForeColor       =   &H00000000&
               Height          =   315
               IMEMode         =   3  'DISABLE
               Left            =   8040
               MaxLength       =   20
               TabIndex        =   57
               Top             =   750
               Width           =   945
            End
            Begin MSDataListLib.DataCombo DcboUsers 
               Height          =   315
               Left            =   4035
               TabIndex        =   61
               Top             =   30
               Width           =   3900
               _ExtentX        =   6879
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               BackColor       =   -2147483624
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcboUsers1 
               Height          =   315
               Left            =   4035
               TabIndex        =   62
               Top             =   390
               Width           =   3900
               _ExtentX        =   6879
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               BackColor       =   -2147483624
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcbDelay 
               Height          =   315
               Left            =   120
               TabIndex        =   63
               Top             =   30
               Width           =   2475
               _ExtentX        =   4366
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               BackColor       =   -2147483624
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcbEarlyexit 
               Height          =   315
               Left            =   120
               TabIndex        =   64
               Top             =   390
               Width           =   2475
               _ExtentX        =   4366
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               BackColor       =   -2147483624
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcbAbscen 
               Height          =   315
               Left            =   120
               TabIndex        =   65
               Top             =   750
               Width           =   2475
               _ExtentX        =   4366
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               BackColor       =   -2147483624
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcbAdd 
               Height          =   315
               Left            =   6990
               TabIndex        =   66
               Top             =   1110
               Width           =   1995
               _ExtentX        =   3519
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               BackColor       =   -2147483624
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcbMokafah 
               Height          =   315
               Left            =   120
               TabIndex        =   67
               Top             =   1110
               Width           =   2475
               _ExtentX        =   4366
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               BackColor       =   -2147483624
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcboUsers2 
               Height          =   315
               Left            =   4035
               TabIndex        =   68
               Top             =   750
               Width           =   3900
               _ExtentX        =   6879
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               BackColor       =   -2147483624
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcbJaza 
               Height          =   315
               Left            =   4035
               TabIndex        =   69
               Top             =   1110
               Width           =   2100
               _ExtentX        =   3704
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               BackColor       =   -2147483624
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo Account_code 
               Height          =   315
               Index           =   0
               Left            =   4020
               TabIndex        =   70
               Top             =   1560
               Width           =   3735
               _ExtentX        =   6588
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo Account_code 
               Height          =   315
               Index           =   1
               Left            =   4020
               TabIndex        =   71
               Top             =   1920
               Width           =   3735
               _ExtentX        =   6588
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo Account_code 
               Height          =   315
               Index           =   2
               Left            =   4020
               TabIndex        =   72
               Top             =   2220
               Width           =   3735
               _ExtentX        =   6588
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin MSDataListLib.DataCombo DcbMokafahVac 
               Height          =   315
               Left            =   120
               TabIndex        =   73
               Top             =   1530
               Width           =   2475
               _ExtentX        =   4366
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               BackColor       =   -2147483624
               Text            =   ""
               RightToLeft     =   -1  'True
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„œÌ— «·«œ«—…"
               Height          =   285
               Index           =   4
               Left            =   8685
               TabIndex        =   86
               Top             =   0
               Width           =   1290
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„œÌ— «·›—⁄ "
               Height          =   285
               Index           =   5
               Left            =   8685
               TabIndex        =   85
               Top             =   360
               Width           =   1290
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„œÌ— «·ﬁÿ«⁄"
               Height          =   285
               Index           =   6
               Left            =   8685
               TabIndex        =   84
               Top             =   720
               Width           =   1290
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "„›—œ  «ŒÌ—"
               Height          =   285
               Index           =   7
               Left            =   2760
               TabIndex        =   83
               Top             =   0
               Width           =   1170
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "„›—œ Œ—ÊÃ „»ﬂ—"
               Height          =   285
               Index           =   8
               Left            =   2760
               TabIndex        =   82
               Top             =   360
               Width           =   1170
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "„›—œ €Ì«»"
               Height          =   285
               Index           =   9
               Left            =   2760
               TabIndex        =   81
               Top             =   720
               Width           =   1170
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„›—œ «÷«›Ì "
               Height          =   285
               Index           =   10
               Left            =   8685
               TabIndex        =   80
               Top             =   1080
               Width           =   1290
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "„›—œ Ã“«¡"
               Height          =   285
               Index           =   11
               Left            =   6000
               TabIndex        =   79
               Top             =   1080
               Width           =   1170
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "„›—œ „ﬂ«›« "
               Height          =   285
               Index           =   12
               Left            =   2760
               TabIndex        =   78
               Top             =   1080
               Width           =   1170
            End
            Begin VB.Label Lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "Õ”«» „’—Ê› «·«Ã«“…"
               Height          =   315
               Index           =   20
               Left            =   8010
               RightToLeft     =   -1  'True
               TabIndex        =   77
               Top             =   1530
               Width           =   1875
            End
            Begin VB.Label Lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "Õ”«» „’—Ê› «· –«ﬂ—"
               Height          =   315
               Index           =   0
               Left            =   8040
               RightToLeft     =   -1  'True
               TabIndex        =   76
               Top             =   1890
               Width           =   1875
            End
            Begin VB.Label Lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "Õ”«» „’—Ê› ‰Â«Ì… «·Œœ„…"
               Height          =   315
               Index           =   1
               Left            =   8010
               RightToLeft     =   -1  'True
               TabIndex        =   75
               Top             =   2220
               Width           =   1875
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00E2E9E9&
               Caption         =   "„›—œ «Ã«“« "
               Height          =   285
               Index           =   14
               Left            =   2760
               TabIndex        =   74
               Top             =   1500
               Width           =   1170
            End
         End
      End
      Begin C1SizerLibCtl.C1Elastic C1Elastic5 
         Height          =   4350
         Left            =   45
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   45
         Width           =   9885
         _cx             =   17436
         _cy             =   7673
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
         Begin C1SizerLibCtl.C1Elastic C1Elastic6 
            Height          =   615
            Left            =   -120
            TabIndex        =   54
            TabStop         =   0   'False
            Top             =   0
            Width           =   10005
            _cx             =   17648
            _cy             =   1085
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
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "«·„Ê«’›« "
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   13.5
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   330
               Index           =   16
               Left            =   3000
               RightToLeft     =   -1  'True
               TabIndex        =   55
               Top             =   45
               Width           =   2430
            End
         End
         Begin MSDataListLib.DataCombo DataCombo1 
            Bindings        =   "FrmEmpDepartments.frx":C48D
            DataSource      =   "Adodc1"
            Height          =   315
            Index           =   7
            Left            =   540
            TabIndex        =   87
            Top             =   1470
            Width           =   7635
            _ExtentX        =   13467
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ListField       =   "Account_Name"
            BoundColumn     =   "Account_Code"
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
         Begin MSDataListLib.DataCombo DataCombo1 
            Bindings        =   "FrmEmpDepartments.frx":C4A2
            DataSource      =   "Adodc1"
            Height          =   315
            Index           =   29
            Left            =   540
            TabIndex        =   88
            Top             =   1800
            Width           =   7635
            _ExtentX        =   13467
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ListField       =   "Account_Name"
            BoundColumn     =   "Account_Code"
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
         Begin MSDataListLib.DataCombo DataCombo1 
            Bindings        =   "FrmEmpDepartments.frx":C4B7
            DataSource      =   "Adodc1"
            Height          =   315
            Index           =   30
            Left            =   540
            TabIndex        =   89
            Top             =   2160
            Width           =   7635
            _ExtentX        =   13467
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ListField       =   "Account_Name"
            BoundColumn     =   "Account_Code"
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
         Begin MSDataListLib.DataCombo DataCombo1 
            Bindings        =   "FrmEmpDepartments.frx":C4CC
            DataSource      =   "Adodc1"
            Height          =   315
            Index           =   65
            Left            =   540
            TabIndex        =   90
            Top             =   3240
            Width           =   7635
            _ExtentX        =   13467
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ListField       =   "Account_Name"
            BoundColumn     =   "Account_Code"
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
         Begin MSDataListLib.DataCombo DataCombo1 
            Bindings        =   "FrmEmpDepartments.frx":C4E1
            DataSource      =   "Adodc1"
            Height          =   315
            Index           =   74
            Left            =   540
            TabIndex        =   91
            Top             =   2520
            Width           =   7635
            _ExtentX        =   13467
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ListField       =   "Account_Name"
            BoundColumn     =   "Account_Code"
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
         Begin MSDataListLib.DataCombo DataCombo1 
            Bindings        =   "FrmEmpDepartments.frx":C4F6
            DataSource      =   "Adodc1"
            Height          =   315
            Index           =   93
            Left            =   540
            TabIndex        =   92
            Top             =   2880
            Width           =   7635
            _ExtentX        =   13467
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            ListField       =   "Account_Name"
            BoundColumn     =   "Account_Code"
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
         Begin VB.Label Labelx 
            Alignment       =   1  'Right Justify
            Caption         =   "„Œ’’ «·«Ã«“« "
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   30
            Left            =   8310
            TabIndex        =   98
            Top             =   2190
            Width           =   1455
         End
         Begin VB.Label Labelx 
            Alignment       =   1  'Right Justify
            Caption         =   "–„„  «·„ÊŸ›Ì‰"
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   7
            Left            =   8400
            TabIndex        =   97
            Top             =   1530
            Width           =   1335
         End
         Begin VB.Label Labelx 
            Alignment       =   1  'Right Justify
            Caption         =   "«·«ÃÊ— «·„” Õﬁ… "
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   29
            Left            =   8310
            TabIndex        =   96
            Top             =   1875
            Width           =   1455
         End
         Begin VB.Label Labelx 
            Alignment       =   1  'Right Justify
            Caption         =   "«·„œ›Ê⁄«  «·„ﬁœ„Â"
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   65
            Left            =   8310
            TabIndex        =   95
            Top             =   3300
            Width           =   1455
         End
         Begin VB.Label Labelx 
            Alignment       =   1  'Right Justify
            Caption         =   "„Œ’’ ‰Â«Ì… «·Œœ„…"
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   72
            Left            =   8310
            TabIndex        =   94
            Top             =   2520
            Width           =   1455
         End
         Begin VB.Label Labelx 
            Alignment       =   1  'Right Justify
            Caption         =   "„Œ’’ «· –«ﬂ—"
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   89
            Left            =   8310
            TabIndex        =   93
            Top             =   2940
            Width           =   1455
         End
      End
   End
End
Attribute VB_Name = "FrmEmpDepartments"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsSavRec As ADODB.Recordset
Dim BKGrndPic As ClsBackGroundPic
Dim RecId As String
Dim ii As Long
Dim Dcombos As New ClsDataCombos

Private Sub DboParentAccount_KeyUp(index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        Account_search.show
        Account_search.case_id = 16112
    End If
End Sub

Private Sub Account_code_KeyUp(index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then
        Account_search.show
        Account_search.mIndex = index
        Account_search.case_id = 654879
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
    Dim i As Integer
    On Error GoTo ErrTrap

    If DoPremis(Do_Delete, Me.Name, True) = False Then
        Exit Sub
    End If

    If TxtVac_ID.text <> "" Then
        If CheckDelDepartment(val(Me.TxtVac_ID.text)) = False Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· „— »ÿ »„·› «·„ÊŸ›Ì‰...!!!"
                  Else
                  Msg = "Can't Delete...!!!"
                  End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            Exit Sub
        End If
   With Grid

For i = 1 To .rows - 1
If CheckDelDepartment2(val(.TextMatrix(i, .ColIndex("ID")))) = True Then
        If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã· „— »ÿ »„·› «·„ÊŸ›Ì‰...!!!"
                  Else
                  Msg = "Can't Delete...!!!"
                  End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
Exit Sub
End If
Next i
End With
If SystemOptions.UserInterface = ArabicInterface Then
        MSGType = MsgBox("Â·  —€» ›Ì Õ–› Â–« «·”Ã·", vbQuestion + vbMsgBoxRtlReading + vbYesNo + vbMsgBoxRight, App.Title)
Else
MSGType = MsgBox("Confirm Deletion", vbQuestion + vbMsgBoxRtlReading + vbYesNo + vbMsgBoxRight, App.Title)
End If
        If MSGType = vbYes Then
        Cn.Execute "Delete from TblEmpDepartmentsDet where DeparmentID=" & val(TxtVac_ID.text) & " "
            RsSavRec.Find "DeparmentID=" & val(TxtVac_ID.text), , adSearchForward, 1
            CuurentLogdata ("D")
            RsSavRec.delete
            If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox " „  ⁄„·Ì… «·Õ–› »‰Ã«Õ", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title
            Else
          MsgBox "Delete Sucess", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.Title

            End If
            '------------------------------ Move Next ---------------------------.
            'FillGridWithData
            BtnNext_Click
        End If
    End If

    Exit Sub
ErrTrap:
 
    Select Case Err.Number

        Case -2147217873, -2147467259
        If SystemOptions.UserInterface = ArabicInterface Then
            StrMSG = "⁄›Ê« ·« ÌÃÊ“ Õ–› «·”Ã· ·«— »«ÿÂ »»Ì«‰«  √Œ—Ì"
         Else
         StrMSG = "Can't Delete Integerty Issue"
         End If
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
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
          Else
          Msg = "Sorry this is record deleted .by another user on network"
          End If
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
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
        Else
          Msg = "Sorry this is record deleted .by another user on network"
        End If
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
If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê«" & CHR(13)
            Msg = Msg & " ·«Ì„ﬂ‰  ⁄œÌ· Â–« «·”Ã· ›Ï «·Êﬁ  «·Õ«·Ï" & CHR(13)
            Msg = Msg & "ÕÌÀ «‰Â ﬁÌœ «· ⁄œÌ· „‰ ﬁ»· „” Œœ„ «Œ— ⁄·Ï «·‘»ﬂ…"
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.Title
 Else
 Msg = "Sorry" & CHR(13)
            Msg = Msg & "Can't Edit This now" & CHR(13)
            Msg = Msg & "Worked in network"
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading + vbOKOnly, App.Title
 End If
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
    Grid.Clear flexClearScrollable, flexClearEverything
    Grid.rows = 1
    My_SQL = "TblEmpDepartments"
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
         If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
        Else
          Msg = "Sorry this is record deleted .by another user on network"
        End If
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
           If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "⁄›Ê« ·ﬁœ  „ Õ–› Â–« «·”Ã· «· «·Ï" & CHR(13)
            Msg = Msg & "„‰ ﬁ»· „” Œœ„ √Œ— ⁄·Ï «·‘»ﬂ… " & CHR(13)
            Msg = Msg & "”Ê› Ì „  ÕœÌÀ «·»Ì«‰« "
        Else
          Msg = "Sorry this is record deleted .by another user on network"
        End If
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
If SystemOptions.UserInterface = ArabicInterface Then
If TxtVacName.text = "" Then
MsgBox "Ì—ÃÏ «œŒ«·  ”„ «·«œ«—…"
TxtVacName.SetFocus
Exit Sub
End If
Else
If TxtVacNamee.text = "" Then
MsgBox "Please Eneter Name"
TxtVacNamee.SetFocus
Exit Sub
End If
End If
    '------------------------------ check if Empcode exist ----------------------

    StrVacName = IsRecExist("TblEmpDepartments", "DepartmentName", Trim(TxtVacName.text), "DepartmentName", "Vac_ID<>'" & Trim(TxtVac_ID.text) & "'")

    If StrVacName <> "" Then
    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "·ﬁœ ”»ﬁ  ”ÃÌ· Â–« «·‰Ê⁄ „‰ ﬁ»·"
      Else
      Msg = "This Type Already exsist"
      End If
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
If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox "Â‰«ﬂ Œÿ√ „« ›Ì ≈œŒ«· «·»Ì«‰« ", vbOKOnly + vbMsgBoxRight, App.Title
 Else
  MsgBox "ÂSorry error douring insert data, vbOKOnly + vbMsgBoxRight, App.title"
 End If
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
    XPLbl(52).Caption = "Branch"
    lbl(20).Caption = "Vacation Acc"
    lbl(0).Caption = "Ticket Acc"
    lbl(1).Caption = "End Service ACC"
     
    
Label1(7).Caption = "Delay"
Label1(8).Caption = "Early Exit"
Label1(9).Caption = "Absence"
Label1(10).Caption = "Additional"
Label1(11).Caption = "Retribution"
Label1(12).Caption = "Reward"
Label1(4).Caption = "Mana.Deptartment"
Label1(5).Caption = "Mana.Branch"
Label1(6).Caption = "Mana.Section"
    Me.Caption = "Departements and Sections"
    Label1(2).Caption = Me.Caption
    Cmd(3).Caption = "Delete"
ISButton2.Caption = "Add"
Label1(25).Caption = "Code"
Label1(24).Caption = "Dept. Name AR"
Label1(23).Caption = "Dept.Name ENG"
XPLbl(0).Caption = "Section"
    With Me.Grid
        .TextMatrix(0, .ColIndex("UserName")) = "Department Manager"
        .TextMatrix(0, .ColIndex("Name")) = "Department Name AR"
        .TextMatrix(0, .ColIndex("NameE")) = "Department Name ENG "

    End With

    Label1(3).Caption = "Section Name ENG"
    Label1(1).Caption = "Section Name AR"
    Label1(0).Caption = "Manager"
 
    Label2(0).Caption = "Curr. Rec."
    Label2(1).Caption = "Rec. Count."

    btnNew.Caption = "New"
    btnModify.Caption = "Modify"
    btnSave.Caption = "Save"
    BtnUndo.Caption = "Undo"
    btnDelete.Caption = "Delete"
    btnCancel.Caption = "Exit"

End Sub

Private Sub DboParentAccount_LostFocus(index As Integer)

End Sub

Private Sub DcbDeptManger_Change()
DcbDeptManger_Click (0)
End Sub

Private Sub DcbDeptManger_Click(Area As Integer)
If val(DcbDeptManger.BoundText) = 0 Then Exit Sub
    Me.TxtDeptCode.text = GeTuserFullCode(val(DcbDeptManger.BoundText))
End Sub

Private Sub DcboUsers_Change()
DcboUsers_Click (0)
End Sub

Private Sub DcboUsers_Click(Area As Integer)
    If val(DcboUsers.BoundText) = 0 Then Exit Sub
   TxtCode.text = GeTuserFullCode(val(DcboUsers.BoundText))
End Sub

Private Sub DcboUsers1_Change()
DcboUsers1_Click (0)
End Sub

Private Sub DcboUsers1_Click(Area As Integer)
    If val(DcboUsers1.BoundText) = 0 Then Exit Sub
     txtCode1.text = GeTuserFullCode(val(DcboUsers1.BoundText))
End Sub

Private Sub DcboUsers2_Change()
DcboUsers2_Click (0)
End Sub

Private Sub DcboUsers2_Click(Area As Integer)
    If val(DcboUsers2.BoundText) = 0 Then Exit Sub
    TxtCode2.text = GeTuserFullCode(val(DcboUsers2.BoundText))
End Sub

Private Sub Form_Load()
'    On Error GoTo ErrTrap
    Dim i As Integer
    Dim My_SQL As String

    My_SQL = "TblEmpDepartments"
    
    Set BKGrndPic = New ClsBackGroundPic
    Set RsSavRec = New ADODB.Recordset
    RsSavRec.CursorLocation = adUseClient
    RsSavRec.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdTableDirect
    If SystemOptions.UserInterface = ArabicInterface Then
        My_SQL = " select id,name from mofrad where ViewComp=1 and  FixedOrChanged=1 and (Unit=1 or Unit=2)"
    Else
        My_SQL = " select id,namee from mofrad where ViewComp=1 and  FixedOrChanged=1 and (Unit=1 or Unit=2)"
    End If
    fill_combo Me.DcbAbscen, My_SQL
    fill_combo Me.DcbDelay, My_SQL
    fill_combo Me.DcbAdd, My_SQL
    fill_combo Me.DcbJaza, My_SQL
    fill_combo Me.DcbEarlyexit, My_SQL
    fill_combo Me.DcbMokafah, My_SQL
    My_SQL = " select id,namee from mofrad where ViewComp=1 "
    fill_combo Me.DcbMokafahVac, My_SQL
    
    Me.TxtModFlg.text = "R"




    ScreenNameArabic = "√ﬁ”«„ «·⁄„· ›Ï «·‘—ﬂ…"
    ScreenNameEnglish = "Departements Data"
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish, "1"
 
    Resize_Form Me

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    'load tblUsers -----------------------------------------------
    My_SQL = "select UserID,UserName From tblUsers "
    fill_combo DCUser, My_SQL
   
    Dcombos.GetUsers Me.DcbDeptManger, False
    Dcombos.GetUsers Me.DcboUsers, False
    Dcombos.GetUsers Me.DcboUsers1, False
    Dcombos.GetUsers Me.DcboUsers2, False
    Dcombos.GetBranches Me.dcBranch
    Dcombos.GetSection Me.DcbSectionID



    Dcombos.GetAccountingCodes Me.Account_code(0), True
    Dcombos.GetAccountingCodes Me.Account_code(1), True
    Dcombos.GetAccountingCodes Me.Account_code(2), True
    
      Dcombos.GetAccountingCodes Me.DataCombo1(7), , True
    Dcombos.GetAccountingCodes Me.DataCombo1(29), , True
    Dcombos.GetAccountingCodes Me.DataCombo1(30), , True
    Dcombos.GetAccountingCodes Me.DataCombo1(65), , True
    Dcombos.GetAccountingCodes Me.DataCombo1(93), , True
    Dcombos.GetAccountingCodes Me.DataCombo1(74), , True
    
   ' FillGridWithData

    With Me.Grid
        '.Cell(flexcpPicture, 0, .ColIndex("DepartmentName")) = Me.GrdImageList.ListImages("Vac_Name").ExtractIcon
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

ErrTrap:
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)
    Dim IntResult As String
    Dim StrMSG As String
    On Error GoTo ErrTrap

    If Me.TxtModFlg.text <> "R" Then

        Select Case Me.TxtModFlg.text

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
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish

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

Function CuurentLogdata(Optional Currentmode As String)
     LogTextA = "    ‘«‘… " & ScreenNameArabic & CHR(13) & "ﬂÊœ " & TxtSerial.text & CHR(13) & "   «”„ «·ﬁ”„ " & TxtVacName
        LogTexte = "    Screen  " & ScreenNameEnglish & CHR(13) & "Code  " & TxtSerial.text & CHR(13) & "   Name " & TxtVacNamee
       If Currentmode <> "D" Then
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, Me.TxtModFlg
    Else
        AddToLogFile CInt(user_id), 0, Date, Time, LogTextA, LogTexte, Me.Name, "D"
    End If
    
End Function

Private Sub Form_Activate()
    Me.ZOrder 0
End Sub

Public Sub AddNewRec()
    On Error GoTo ErrTrap
    Dim StrRecID As String
    StrRecID = new_id("TblEmpDepartments", "DeparmentID", "")
    RsSavRec.AddNew
    TxtVac_ID.text = StrRecID
    RsSavRec.Fields("DeparmentID").value = IIf(StrRecID <> "", StrRecID, Null)
    FiLLRec
ErrTrap:
End Sub

Public Sub FiLLRec()
Dim i As Integer
    On Error GoTo ErrTrap
Dim Rs4 As ADODB.Recordset
Dim sql As String

   
        Dim Account_Code_dynamic1 As String
        Dim Account_Code_dynamic2 As String
        Dim Account_Code_dynamic3 As String
         
        If SystemOptions.ProvisionsByManagement = True Then
            Account_Code_dynamic1 = get_account_code_branch(55, my_branch)
            
            If Account_Code_dynamic1 = "NO branch" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                Else
                    MsgBox "Branch Not Created", vbCritical
                End If
    
                GoTo ErrTrap
            Else
    
                If Account_Code_dynamic1 = "NO account" Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox "·„ Ì „  ÕœÌœ Õ”«» «·«Ã«“… ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                    Else
                        MsgBox "Inventory Account not defined on this branch", vbCritical
                    End If
            
                    GoTo ErrTrap
             
                End If
            End If
            
            
             Account_Code_dynamic1 = get_account_code_branch(56, my_branch)
            
            If Account_Code_dynamic2 = "NO branch" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                Else
                    MsgBox "Branch Not Created", vbCritical
                End If
    
                GoTo ErrTrap
            Else
    
                If Account_Code_dynamic2 = "NO account" Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox "·„ Ì „  ÕœÌœ Õ”«» «· –«ﬂ— ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                    Else
                        MsgBox "Inventory Account not defined on this branch", vbCritical
                    End If
            
                    GoTo ErrTrap
             
                End If
            End If
            
            
             Account_Code_dynamic3 = get_account_code_branch(94, my_branch)
            
            If Account_Code_dynamic1 = "NO branch" Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "·„ Ì „ «‰‘«¡ «·›—⁄", vbCritical
                Else
                    MsgBox "Branch Not Created", vbCritical
                End If
    
                GoTo ErrTrap
            Else
    
                If Account_Code_dynamic3 = "NO account" Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        MsgBox "·„ Ì „  ÕœÌœ Õ”«» ‰Â«Ì… «·Œœ„… ›Ì «·›—⁄ ·Â–… «·⁄„·Ì…", vbCritical
                    Else
                        MsgBox "Inventory Account not defined on this branch", vbCritical
                    End If
            
                    GoTo ErrTrap
             
                End If
            End If
            
        End If

    If Me.TxtModFlg = "N" Then
            Dim last_account As Boolean
            Dim link_account As Boolean
            Dim X As String
        If SystemOptions.ProvisionsByManagement = True Then
            RsSavRec("Account_Code1").value = Account_code(0).BoundText ' ModAccounts.AddNewAccount(Account_Code_dynamic1, " „’—Ê› «Ã«“…  " & TxtVacName.Text, last_account, False, TxtVacNamee.Text)
            RsSavRec("Account_Code2").value = Account_code(1).BoundText  ' ModAccounts.AddNewAccount(Account_Code_dynamic2, " „’—Ê› «· –«ﬂ—  " & TxtVacName.Text, last_account, False, TxtVacNamee.Text)
            RsSavRec("Account_Code3").value = Account_code(2).BoundText ' ModAccounts.AddNewAccount(Account_Code_dynamic3, " „’—Ê› ‰Â«Ì… «·Œœ„…  " & TxtVacName.Text, last_account, False, TxtVacNamee.Text)
        Else
            RsSavRec("Account_Code1").value = Account_Code_dynamic1
            RsSavRec("Account_Code2").value = Account_Code_dynamic2
            RsSavRec("Account_Code3").value = Account_Code_dynamic3
        End If
    End If

If Me.TxtModFlg.text = "E" Then
Cn.Execute "Delete from TblEmpDepartmentsDet  where DeparmentID =" & val(TxtVac_ID.text) & "  "
'With Grid
'   If CheckDelDepartment(val(Me.TxtVac_ID.text)) = True Then
'For i = 1 To .Rows - 1
'If CheckDelDepartment2(val(.TextMatrix(i, .ColIndex("ID")))) = True Then
'Cn.Execute "Delete from TblEmpDepartmentsDet  where id =" & val(.TextMatrix(i, .ColIndex("ID"))) & "  "
'End If
'Next i
'End If
'End With
End If
    RsSavRec.Fields("DepartmentName").value = IIf(TxtVacName.text <> "", Trim(TxtVacName.text), Null)
    RsSavRec.Fields("DepartmentNamee").value = IIf(TxtVacNamee.text <> "", Trim(TxtVacNamee.text), Null)
    RsSavRec.Fields("UserId").value = val(DcboUsers.BoundText)
    RsSavRec.Fields("UserId1").value = val(DcboUsers1.BoundText)
    RsSavRec.Fields("UserId2").value = val(DcboUsers2.BoundText)
    RsSavRec.Fields("BranchId").value = val(dcBranch.BoundText)
    
    RsSavRec.Fields("a7").value = (DataCombo1(7).BoundText)
    RsSavRec.Fields("a29").value = (DataCombo1(29).BoundText)
   RsSavRec.Fields("a30").value = (DataCombo1(30).BoundText)
   RsSavRec.Fields("a65").value = (DataCombo1(65).BoundText)
   RsSavRec.Fields("a74").value = (DataCombo1(74).BoundText)
   RsSavRec.Fields("a93").value = (DataCombo1(93).BoundText)
   
    
    
    
    
    
   
    
    ''//////
    RsSavRec.Fields("DelayID").value = val(Me.DcbDelay.BoundText)
    RsSavRec.Fields("EarlyexitID").value = val(Me.DcbEarlyexit.BoundText)
    RsSavRec.Fields("AbscenID").value = val(Me.DcbAbscen.BoundText)
    RsSavRec.Fields("AddID").value = val(Me.DcbAdd.BoundText)
    RsSavRec.Fields("JazaID").value = val(Me.DcbJaza.BoundText)
    RsSavRec.Fields("MokafahID").value = val(Me.DcbMokafah.BoundText)
    RsSavRec.Fields("MokafahVacID").value = val(Me.DcbMokafahVac.BoundText)
    
    RsSavRec.Fields("MokafahID").value = val(Me.DcbMokafah.BoundText)
    RsSavRec.Fields("SectionID").value = val(Me.DcbSectionID.BoundText)
    RsSavRec.Fields("MangerID").value = val(Me.DcbDeptManger.BoundText)
    RsSavRec.Fields("Name").value = IIf(TxtName.text <> "", Trim(TxtName.text), Null)
    RsSavRec.Fields("NameE").value = IIf(TxtNameE.text <> "", Trim(TxtNameE.text), Null)
    
    RsSavRec.Fields("Account_Code1").value = Trim(Me.Account_code(0).BoundText)
    RsSavRec.Fields("Account_Code2").value = Trim(Me.Account_code(1).BoundText)
    RsSavRec.Fields("Account_Code3").value = Trim(Me.Account_code(2).BoundText)
       
'    If Not IsNull(RsSavRec("Account_Code1").value) Then
'        ModAccounts.EditAccount RsSavRec("Account_Code1").value, Me.TxtVacName.Text & "  „’—Ê› «Ã«“…   ", TxtVacNamee.Text & "   -", , , , , , , , , , , , , , , , , last_account
'    End If
'
'    If Not IsNull(RsSavRec("Account_Code2").value) Then
'        ModAccounts.EditAccount RsSavRec("Account_Code2").value, Me.TxtVacName.Text & "  „’—Ê›  –«ﬂ—   ", TxtVacNamee.Text & "   -Inventory", , , , , , , , , , , , , , , , , last_account
'    End If
'
'    If Not IsNull(RsSavRec("Account_Code3").value) Then
'        ModAccounts.EditAccount RsSavRec("Account_Code3").value, Me.TxtVacName.Text & "  „’—Ê› ‰Â«Ì… Œœ„…   ", TxtVacNamee.Text & "   -Inventory", , , , , , , , , , , , , , , , , last_account
'    End If
    Dcombos.GetAccountingCodes Me.Account_code(0), True
    Dcombos.GetAccountingCodes Me.Account_code(1), True
    Dcombos.GetAccountingCodes Me.Account_code(2), True
    
    
    
  
    
    
        
   
    
    Me.Account_code(0).BoundText = RsSavRec("Account_Code1").value
    Me.Account_code(1).BoundText = RsSavRec("Account_Code2").value
    Me.Account_code(2).BoundText = RsSavRec("Account_Code3").value

    
    
    
    RsSavRec.update
    ''//////////////
    Dim StrRecID As Double
    Set Rs4 = New ADODB.Recordset
    sql = "SELECT  *  from TblEmpDepartmentsDet Where (1 = -1)"
    Rs4.Open sql, Cn, adOpenKeyset, adLockOptimistic, adCmdText
    With Me.Grid
    For i = 1 To .rows - 1
    If .TextMatrix(i, .ColIndex("Name")) <> "" Or .TextMatrix(i, .ColIndex("NameE")) <> "" Then
    If val(.TextMatrix(i, .ColIndex("ID"))) = 0 Then
    StrRecID = new_id("TblEmpDepartmentsDet", "ID", "")
    
    Else
    StrRecID = val(.TextMatrix(i, .ColIndex("ID")))
    End If
    Rs4.AddNew
    Rs4("ID").value = StrRecID
    Rs4("DeparmentID").value = val(TxtVac_ID.text)
    Rs4("Name").value = IIf(.TextMatrix(i, .ColIndex("Name")) = "", Null, .TextMatrix(i, .ColIndex("Name")))
    Rs4("NameE").value = IIf(.TextMatrix(i, .ColIndex("NameE")) = "", Null, .TextMatrix(i, .ColIndex("NameE")))
    Rs4("MangerID").value = IIf(val(.TextMatrix(i, .ColIndex("MangerID"))) = 0, Null, val(.TextMatrix(i, .ColIndex("MangerID"))))
    Rs4.update
    End If
    Next i
    End With
    If SystemOptions.UserInterface = ArabicInterface Then
    MsgBox " „  ⁄„·Ì… «·Õ›Ÿ »‰Ã«Õ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    Else
    MsgBox "Save Success", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
    End If
    CuurentLogdata
    FiLLTXT
   ' FillGridWithData
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
TxtSerial.text = IIf(IsNull(RsSavRec.Fields("DeparmentID").value), "", RsSavRec.Fields("DeparmentID").value)
    TxtVac_ID.text = IIf(IsNull(RsSavRec.Fields("DeparmentID").value), "", RsSavRec.Fields("DeparmentID").value)
    TxtVacName.text = IIf(IsNull(RsSavRec.Fields("DepartmentName").value), "", RsSavRec.Fields("DepartmentName").value)
    TxtVacNamee.text = IIf(IsNull(RsSavRec.Fields("DepartmentNamee").value), "", RsSavRec.Fields("DepartmentNamee").value)
DcboUsers.BoundText = IIf(IsNull(RsSavRec.Fields("UserId").value), "", RsSavRec.Fields("UserId").value)
DcboUsers1.BoundText = IIf(IsNull(RsSavRec.Fields("UserId1").value), "", RsSavRec.Fields("UserId1").value)
DcboUsers2.BoundText = IIf(IsNull(RsSavRec.Fields("UserId2").value), "", RsSavRec.Fields("UserId2").value)
dcBranch.BoundText = IIf(IsNull(RsSavRec.Fields("Branchid").value), "", RsSavRec.Fields("Branchid").value)
Me.DcbDelay.BoundText = IIf(IsNull(RsSavRec.Fields("DelayID").value), "", RsSavRec.Fields("DelayID").value)
Me.DcbEarlyexit.BoundText = IIf(IsNull(RsSavRec.Fields("EarlyexitID").value), "", RsSavRec.Fields("EarlyexitID").value)
Me.DcbAbscen.BoundText = IIf(IsNull(RsSavRec.Fields("AbscenID").value), "", RsSavRec.Fields("AbscenID").value)
Me.DcbAdd.BoundText = IIf(IsNull(RsSavRec.Fields("AddID").value), "", RsSavRec.Fields("AddID").value)
Me.DcbJaza.BoundText = IIf(IsNull(RsSavRec.Fields("JazaID").value), "", RsSavRec.Fields("JazaID").value)
Me.DcbMokafah.BoundText = IIf(IsNull(RsSavRec.Fields("MokafahID").value), "", RsSavRec.Fields("MokafahID").value)

Me.DcbMokafahVac.BoundText = IIf(IsNull(RsSavRec.Fields("MokafahVacID").value), "", RsSavRec.Fields("MokafahVacID").value)

    DataCombo1(7).BoundText = RsSavRec.Fields("a7").value & ""
    DataCombo1(29).BoundText = RsSavRec.Fields("a29").value & ""
   DataCombo1(30).BoundText = RsSavRec.Fields("a30").value & ""
   DataCombo1(65).BoundText = RsSavRec.Fields("a65").value & ""
   DataCombo1(74).BoundText = RsSavRec.Fields("a74").value & ""
   DataCombo1(93).BoundText = RsSavRec.Fields("a93").value & ""
   


Me.DcbSectionID.BoundText = IIf(IsNull(RsSavRec.Fields("SectionID").value), "", RsSavRec.Fields("SectionID").value)
Me.DcbDeptManger.BoundText = IIf(IsNull(RsSavRec.Fields("MangerID").value), "", RsSavRec.Fields("MangerID").value)
Me.TxtName.text = IIf(IsNull(RsSavRec.Fields("Name").value), "", RsSavRec.Fields("Name").value)
Me.TxtNameE.text = IIf(IsNull(RsSavRec.Fields("NameE").value), "", RsSavRec.Fields("NameE").value)

Me.Account_code(0).BoundText = IIf(IsNull(RsSavRec("Account_Code1").value), "", RsSavRec("Account_Code1").value)
Me.Account_code(1).BoundText = IIf(IsNull(RsSavRec("Account_Code2").value), "", RsSavRec("Account_Code2").value)
Me.Account_code(2).BoundText = IIf(IsNull(RsSavRec("Account_Code3").value), "", RsSavRec("Account_Code3").value)

    LabCurrRec.Caption = RsSavRec.AbsolutePosition
    LabCountRec.Caption = RsSavRec.RecordCount
FullGrid
ErrTrap:

End Sub
Sub FullGrid()
Dim i As Integer
Dim sql As String
    Grid.Clear flexClearScrollable, flexClearEverything
    Grid.rows = 1
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
sql = " SELECT     dbo.TblEmpDepartmentsDet.ID, dbo.TblEmpDepartmentsDet.DeparmentID, dbo.TblEmpDepartmentsDet.Name, dbo.TblEmpDepartmentsDet.NameE,"
sql = sql & "                       dbo.TblEmpDepartmentsDet.mangerid , dbo.TblUsers.UserName"
sql = sql & "  FROM         dbo.TblEmpDepartmentsDet LEFT OUTER JOIN"
sql = sql & "                       dbo.TblUsers ON dbo.TblEmpDepartmentsDet.MangerID = dbo.TblUsers.UserID"
sql = sql & "  Where (dbo.TblEmpDepartmentsDet.DeparmentID = " & val(TxtVac_ID.text) & ")"
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
With Grid
.rows = Rs3.RecordCount + 1
Rs3.MoveFirst
For i = 1 To .rows - 1
.TextMatrix(i, .ColIndex("ID")) = IIf(IsNull(Rs3("ID").value), 0, Rs3("ID").value)
'.TextMatrix(i, .ColIndex("DeparmentID")) = IIf(IsNull(RS3("DeparmentID").value), 0, RS3("DeparmentID").value)
.TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(Rs3("Name").value), "", Rs3("Name").value)
.TextMatrix(i, .ColIndex("NameE")) = IIf(IsNull(Rs3("NameE").value), "", Rs3("NameE").value)
.TextMatrix(i, .ColIndex("MangerID")) = IIf(IsNull(Rs3("MangerID").value), 0, Rs3("MangerID").value)
.TextMatrix(i, .ColIndex("UserName")) = IIf(IsNull(Rs3("UserName").value), "", Rs3("UserName").value)
Rs3.MoveNext
Next i

End With
End If
End Sub
Public Sub EditRec(StrTable As String, _
                   RecId As String)
    'My_SQL = "select * From " & StrTable & " where "
    'RsSavRec.Open My_SQL, cn, adOpenStatic, adLockReadOnly, adCmdTableDirect
    FiLLRec

End Sub

Private Sub Grid_EnterCell()
    On Error GoTo ErrTrap
    FillTextFromGrid
ErrTrap:
End Sub
Sub FillTextFromGrid()
Dim i As Integer
With Me.Grid
If .row > 0 Then
txtID.text = .row
TxtName.text = .TextMatrix(.row, .ColIndex("Name"))
TxtNameE.text = .TextMatrix(.row, .ColIndex("NameE"))
Me.DcbDeptManger.BoundText = val(.TextMatrix(.row, .ColIndex("MangerID")))
End If
End With

End Sub


Private Sub ISButton2_Click()
If Me.TxtModFlg.text <> "R" Then
If SystemOptions.UserInterface = ArabicInterface Then
If Me.TxtName.text = "" Then
MsgBox "Ì—ÃÏ «œŒ«· «·«”„"
Me.TxtName.SetFocus
Exit Sub
End If
Else
If Me.TxtNameE.text = "" Then
MsgBox "Please Eneter Name"
Me.TxtNameE.SetFocus
Exit Sub
End If
End If
filgrid
Me.TxtNameE.text = ""
Me.TxtName.text = ""
Me.DcbDeptManger.BoundText = 0
txtID.text = 0
End If
End Sub
Sub filgrid()
Dim i As Integer
Dim k As Integer
If val(txtID.text) = 0 Then
With Grid
k = .rows
.rows = .rows + 1
For i = k To .rows - 1
.TextMatrix(i, .ColIndex("MangerID")) = val(Me.DcbDeptManger.BoundText)
.TextMatrix(i, .ColIndex("UserName")) = Me.DcbDeptManger.text
.TextMatrix(i, .ColIndex("NameE")) = Me.TxtNameE.text
.TextMatrix(i, .ColIndex("Name")) = Me.TxtName.text
.TextMatrix(i, .ColIndex("Ser")) = i
Next i
End With
Else
With Grid
.TextMatrix(val(txtID.text), .ColIndex("MangerID")) = val(Me.DcbDeptManger.BoundText)
.TextMatrix(val(txtID.text), .ColIndex("UserName")) = Me.DcbDeptManger.text
.TextMatrix(val(txtID.text), .ColIndex("NameE")) = Me.TxtNameE.text
.TextMatrix(val(txtID.text), .ColIndex("Name")) = Me.TxtName.text
End With
End If
End Sub
Private Sub TxtCode_Change()
    DcboUsers.BoundText = GeTuserIDByEmpCode(TxtCode.text)
End Sub
Private Sub Cmd_Click(index As Integer)
If Me.TxtModFlg.text <> "R" Then
Select Case index
Case 3
RemoveGridRow

End Select
End If
End Sub

Private Sub RemoveGridRow()
Dim Msg As String
    With Me.Grid
        If .row <= 0 Then Exit Sub
                If CheckDelDepartment(val(Me.TxtVac_ID.text)) = False Then
                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "·«Ì„ﬂ‰ Õ–› Â–« «·”Ã·...!!!"
                  Else
                  Msg = "Can't Delete...!!!"
                  End If
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            Exit Sub
        End If
        Cn.Execute "Delete TblEmpDepartmentsDet  where id =" & val(.TextMatrix(.row, .ColIndex("ID"))) & "  "
        .RemoveItem .row
    End With
   ' ReLineGrid
End Sub
Private Sub TXTCode1_Change()
    DcboUsers1.BoundText = GeTuserIDByEmpCode(txtCode1.text)

End Sub

Private Sub txtCode2_Change()
    DcboUsers.BoundText = GeTuserIDByEmpCode(TxtCode2.text)
End Sub

Private Sub TxtDeptCode_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
Me.DcbDeptManger.BoundText = GeTuserIDByEmpCode(TxtDeptCode.text)
End If
End Sub

Private Sub TxtVac_ID_Change()
    Dim TxtMod As String
    TxtMod = TxtModFlg.text
    TxtModFlg.text = ""
    TxtModFlg = TxtMod
End Sub

Public Function FindRec(ByVal RecId As Long)
    On Error GoTo ErrTrap

    RsSavRec.Find "DeparmentID=" & RecId, , adSearchForward, 1

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

Private Sub TxtModFlg_Change()

    If TxtModFlg.text = "N" Then
        Frm2.Enabled = True
        Me.btnNew.Enabled = False
        btnModify.Enabled = False
        btnDelete.Enabled = False
        Me.btnQuery.Enabled = False
      '  Grid.Enabled = False
        BtnUndo.Enabled = True
        Me.btnSave.Enabled = True
        BtnUpdate.Enabled = False
        '   btnNext.Enabled = False
        '   btnPrevious.Enabled = False
        '   btnFirst.Enabled = False
        '   btnLast.Enabled = False
    
    ElseIf TxtModFlg.text = "R" Then
        Frm2.Enabled = False
     '   Grid.Enabled = True
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
       ' Frm2.Enabled = True
        Me.btnNew.Enabled = False
        btnModify.Enabled = False
        btnDelete.Enabled = False
        Me.btnQuery.Enabled = False
        BtnUpdate.Enabled = False
        BtnUndo.Enabled = True
        Me.btnSave.Enabled = True
       ' Grid.Enabled = False
        btnNext.Enabled = False
        btnPrevious.Enabled = False
        btnFirst.Enabled = False
        btnLast.Enabled = False
    
    End If

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
Private Function CheckDelDepartment2(LngDepartmentID As Long) As Boolean
    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    StrSQL = "Select * From TblEmployee Where DeptID2=" & LngDepartmentID & ""
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.BOF Or rs.EOF) Then
        CheckDelDepartment2 = False
    Else
        CheckDelDepartment2 = True
    End If

    rs.Close
    Set rs = Nothing
End Function

Private Function CheckDelDepartment(LngDepartmentID As Long) As Boolean
    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    StrSQL = "Select * From TblEmployee Where DepartmentID=" & LngDepartmentID & ""
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.BOF Or rs.EOF) Then
        CheckDelDepartment = False
    Else
        CheckDelDepartment = True
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

