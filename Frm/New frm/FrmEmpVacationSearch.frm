VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SuiteCtrls.ocx"
Begin VB.Form FrmEmpVacationSearch 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "«·»ÕÀ ⁄‰ ÿ·» ≈Ã«“…"
   ClientHeight    =   6060
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9450
   Icon            =   "FrmEmpVacationSearch.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6060
   ScaleWidth      =   9450
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   " «—ÌŒ ‰Â«Ì…«·«Ã«“…"
      Height          =   1035
      Index           =   4
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   4560
      Width           =   3255
      Begin MSComCtl2.DTPicker FromEndDate 
         Height          =   330
         Left            =   1290
         TabIndex        =   26
         Top             =   270
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   210436099
         CurrentDate     =   38887
      End
      Begin MSComCtl2.DTPicker ToEndDate 
         Height          =   330
         Left            =   1290
         TabIndex        =   27
         Top             =   630
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   210436099
         CurrentDate     =   38887
      End
      Begin Dynamic_Byte.NourHijriCal FromEndDateH 
         Height          =   315
         Left            =   120
         TabIndex        =   28
         Top             =   270
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
      End
      Begin Dynamic_Byte.NourHijriCal ToEndDateH 
         Height          =   315
         Left            =   120
         TabIndex        =   29
         Top             =   630
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï"
         Height          =   195
         Index           =   11
         Left            =   2895
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   660
         Width           =   255
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   195
         Index           =   7
         Left            =   2820
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   330
         Width           =   300
      End
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Height          =   2235
      Index           =   3
      Left            =   3360
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   3360
      Width           =   6135
      Begin VB.Frame Frame1 
         BackColor       =   &H00E2E9E9&
         Caption         =   "„Ê«›ﬁ… «·„œÌ—"
         Height          =   735
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   600
         Width           =   5892
         Begin XtremeSuiteControls.RadioButton Rd 
            Height          =   495
            Index           =   1
            Left            =   2520
            TabIndex        =   41
            Top             =   120
            Width           =   1095
            _Version        =   786432
            _ExtentX        =   1931
            _ExtentY        =   873
            _StockProps     =   79
            Caption         =   " „  «·„Ê«›ﬁ…"
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton Rd 
            Height          =   495
            Index           =   2
            Left            =   120
            TabIndex        =   42
            Top             =   120
            Width           =   1695
            _Version        =   786432
            _ExtentX        =   2990
            _ExtentY        =   873
            _StockProps     =   79
            Caption         =   "·„   „ «·„Ê«›ﬁ…"
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton Rd 
            Height          =   495
            Index           =   0
            Left            =   3840
            TabIndex        =   43
            Top             =   120
            Width           =   1095
            _Version        =   786432
            _ExtentX        =   1931
            _ExtentY        =   873
            _StockProps     =   79
            Caption         =   "«·ﬂ·"
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
      End
      Begin VB.Frame lbltype 
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰Ê⁄ «·«Ã«“…"
         Height          =   735
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   1320
         Width           =   5892
         Begin XtremeSuiteControls.RadioButton Rdb1 
            Height          =   495
            Index           =   1
            Left            =   2520
            TabIndex        =   36
            Top             =   120
            Width           =   1095
            _Version        =   786432
            _ExtentX        =   1931
            _ExtentY        =   873
            _StockProps     =   79
            Caption         =   "—”„Ì…"
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton Rdb1 
            Height          =   495
            Index           =   2
            Left            =   1440
            TabIndex        =   37
            Top             =   120
            Width           =   1095
            _Version        =   786432
            _ExtentX        =   1931
            _ExtentY        =   873
            _StockProps     =   79
            Caption         =   "«÷ÿ—«—Ì…"
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton Rdb1 
            Height          =   495
            Index           =   3
            Left            =   120
            TabIndex        =   38
            Top             =   120
            Width           =   1095
            _Version        =   786432
            _ExtentX        =   1931
            _ExtentY        =   873
            _StockProps     =   79
            Caption         =   "«Œ—Ï"
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton Rdb1 
            Height          =   495
            Index           =   0
            Left            =   3840
            TabIndex        =   39
            Top             =   120
            Width           =   1095
            _Version        =   786432
            _ExtentX        =   1931
            _ExtentY        =   873
            _StockProps     =   79
            Caption         =   "«·ﬂ·"
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
      End
      Begin VB.TextBox TxtSearchCode 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3720
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   240
         Width           =   1335
      End
      Begin MSDataListLib.DataCombo DcboEmpName 
         Height          =   315
         Left            =   360
         TabIndex        =   33
         Top             =   240
         Width           =   3315
         _ExtentX        =   5847
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«”„ «·„ÊŸ›"
         Height          =   285
         Index           =   0
         Left            =   4950
         TabIndex        =   34
         Top             =   240
         Width           =   1005
      End
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   " «—ÌŒ »œ«Ì… «·«Ã«“…"
      Height          =   1035
      Index           =   0
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   3480
      Width           =   3255
      Begin MSComCtl2.DTPicker FromStartDate 
         Height          =   330
         Left            =   1290
         TabIndex        =   18
         Top             =   270
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   210436099
         CurrentDate     =   38887
      End
      Begin MSComCtl2.DTPicker ToStartDate 
         Height          =   330
         Left            =   1290
         TabIndex        =   19
         Top             =   630
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   210436099
         CurrentDate     =   38887
      End
      Begin Dynamic_Byte.NourHijriCal FromStartDateH 
         Height          =   315
         Left            =   120
         TabIndex        =   22
         Top             =   270
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
      End
      Begin Dynamic_Byte.NourHijriCal ToStartDateH 
         Height          =   315
         Left            =   120
         TabIndex        =   23
         Top             =   630
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   195
         Index           =   9
         Left            =   2700
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   330
         Width           =   420
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï"
         Height          =   195
         Index           =   8
         Left            =   2895
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   660
         Width           =   255
      End
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   " «—ÌŒ «· ”ÃÌ·"
      Height          =   645
      Index           =   1
      Left            =   2310
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   2700
      Width           =   4455
      Begin MSComCtl2.DTPicker DtpDateFrom 
         Height          =   330
         Left            =   1890
         TabIndex        =   6
         Top             =   150
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   210436099
         CurrentDate     =   38887
      End
      Begin MSComCtl2.DTPicker DtpDateTo 
         Height          =   330
         Left            =   90
         TabIndex        =   7
         Top             =   150
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   210436099
         CurrentDate     =   38887
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï"
         Height          =   195
         Index           =   3
         Left            =   1575
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   180
         Width           =   255
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   195
         Index           =   4
         Left            =   3600
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   210
         Width           =   420
      End
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·ÿ·»"
      Height          =   645
      Index           =   2
      Left            =   6720
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   2700
      Width           =   2715
      Begin VB.TextBox TxtIDFrom 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   1320
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   180
         Width           =   675
      End
      Begin VB.TextBox TxtIDTO 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   180
         Width           =   675
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   195
         Index           =   5
         Left            =   2175
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   240
         Width           =   420
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï"
         Height          =   195
         Index           =   6
         Left            =   780
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   240
         Width           =   405
      End
   End
   Begin VSFlex8UCtl.VSFlexGrid Fg 
      Height          =   2625
      Left            =   30
      TabIndex        =   10
      Top             =   0
      Width           =   9435
      _cx             =   16642
      _cy             =   4630
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
      FormatString    =   $"FrmEmpVacationSearch.frx":038A
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
      Left            =   1650
      TabIndex        =   11
      Top             =   5640
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   661
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
      ColorShadow     =   -2147483637
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      DisabledImageExtraction=   0
      ColorToggledHoverText=   16711680
      LowerToggledContent=   0   'False
      ColorTextShadow =   -2147483637
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   1
      Left            =   810
      TabIndex        =   12
      Top             =   5640
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   661
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
      ColorShadow     =   -2147483637
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      LowerToggledContent=   0   'False
      ColorTextShadow =   -2147483637
   End
   Begin ImpulseButton.ISButton Cmd 
      Cancel          =   -1  'True
      Height          =   375
      Index           =   2
      Left            =   30
      TabIndex        =   13
      Top             =   5640
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
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
      ColorShadow     =   -2147483637
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      LowerToggledContent=   0   'False
      ColorTextShadow =   -2147483637
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·≈Ã„«·Ï"
      Height          =   285
      Index           =   2
      Left            =   1290
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   3060
      Width           =   945
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      ForeColor       =   &H00000080&
      Height          =   285
      Index           =   1
      Left            =   60
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   3060
      Width           =   1785
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      ForeColor       =   &H00000080&
      Height          =   315
      Index           =   10
      Left            =   60
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   2700
      Width           =   1815
   End
End
Attribute VB_Name = "FrmEmpVacationSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DCboSearch As clsDCboSearch
Public index As Integer
Private Sub ChangeLang()
    Cmd(1).Caption = "Clear"
    Cmd(0).Caption = "Search"
    Cmd(2).Caption = "Exit"
     Me.Caption = "Search Request Vacation"
    ' labell name
   Fra(2).Caption = "Req No"
   ' Me.lbl(14).Caption = "Trans ID"
    lbl(4).Caption = "From"
    lbl(5).Caption = "From"
    lbl(9).Caption = "From"
    lbl(7).Caption = "From"
    lbl(6).Caption = "To"
    lbl(3).Caption = "To"
    lbl(8).Caption = "To"
    lbl(11).Caption = "To"
    Fra(1).Caption = "Req Date"
    lbl(2).Caption = "Total"
    lbl(0).Caption = "Employee"
    Frame1.Caption = "Accept Manager"
Rd(0).Caption = "All"
Rd(0).RightToLeft = False
Rd(1).Caption = "Accept"
Rd(1).RightToLeft = False
Rd(2).Caption = "Not Accept"
Rd(2).RightToLeft = False
Rdb1(0).Caption = "All"
Rdb1(0).RightToLeft = False
Rdb1(1).Caption = "Official"
Rdb1(1).RightToLeft = False
Rdb1(2).Caption = "Important"
Rdb1(2).RightToLeft = False
Rdb1(3).Caption = "Other"
Rdb1(3).RightToLeft = False
Fra(4).Caption = "End Vacation"
Fra(0).Caption = "Start Vacation"
lbltype.Caption = "Type Vacation"
With Fg
.TextMatrix(0, .ColIndex("Serial")) = "Serial"
.TextMatrix(0, .ColIndex("ID")) = "No"
.TextMatrix(0, .ColIndex("RecordDate")) = "Trans Date"
.TextMatrix(0, .ColIndex("Emp_Name")) = "Employee Name"
.TextMatrix(0, .ColIndex("ManagerApprove")) = "Manager Accept"
.TextMatrix(0, .ColIndex("VocationType")) = "Type Vocation"
.TextMatrix(0, .ColIndex("FromDate")) = "Start Vacation"
.TextMatrix(0, .ColIndex("FromDateH")) = "Start Vacation"
.TextMatrix(0, .ColIndex("ToDate")) = "End Vacation"
.TextMatrix(0, .ColIndex("ToDateH")) = "End Vacation"

End With

  End Sub
Private Sub Cmd_Click(index As Integer)

    Select Case index

        Case 0
            GetData

        Case 1
            clear_all Me
FromStartDate.value = ""
ToStartDate.value = ""
FromEndDate.value = ""
ToEndDate.value = ""
DtpDateFrom.value = ""
DtpDateTo.value = ""
            If SystemOptions.UserInterface = ArabicInterface Then
                Me.lbl(0).Caption = "‰ ÌÃ… «·»ÕÀ"
            Else
                Me.lbl(0).Caption = "Search Results"
            End If

        Case 2
            Unload Me
    End Select

End Sub

Private Sub DcboEmpName_Click(Area As Integer)
   If val(DcboEmpName.BoundText) = 0 Then Exit Sub

    Dim EmpCode  As String
 
    GetEmployeeIDFromCode , , DcboEmpName.BoundText, EmpCode
    TxtSearchCode.text = EmpCode
End Sub

Private Sub fg_Click()

    With Me.Fg

        If .row = -1 Then Exit Sub
        If .Col = -1 Then Exit Sub
        If val(.TextMatrix(.row, .ColIndex("ID"))) = 0 Then
            Exit Sub
        End If
If index = 0 Then
formvocatinl.Retrive val(.TextMatrix(.row, .ColIndex("ID")))
ElseIf index = 1 Then
FrmVocationEntitlements.TxtOrder.text = val(.TextMatrix(.row, .ColIndex("ID")))
FrmVocationEntitlements.TxtOrder_KeyPress (13)
  End If
    End With

End Sub

Private Sub Form_Activate()
    PutFormOnTop Me.hWnd
End Sub

Private Sub Form_Load()
    Dim GrdBack As ClsBackGroundPic
    Dim Dcombos As ClsDataCombos

    Set Dcombos = New ClsDataCombos
    Dcombos.GetEmployees Me.DcboEmpName
    Set DCboSearch = New clsDCboSearch
   ' Set DCboSearch.Client = Me.DCEmp_Name
   ' Dcombos.GetUsers Me.DCUser
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Clear").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture

    CenterForm Me

    FormPostion Me, GetPostion
    Set GrdBack = New ClsBackGroundPic

    With Me.Fg
        Set .WallPaper = GrdBack.Picture
        .AutoSize 0, .Cols - 1, False
    End With
 If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
    SetDtpickerDate Me.FromStartDate
    SetDtpickerDate Me.ToStartDate
    
     SetDtpickerDate Me.FromEndDate
    SetDtpickerDate Me.ToEndDate
     SetDtpickerDate Me.DtpDateFrom
    SetDtpickerDate Me.DtpDateTo

End Sub
Private Sub DcboEmpName_Change()
DcboEmpName_Click (0)
End Sub

Private Sub Form_Unload(Cancel As Integer)

    FormPostion Me, SavePostion
    Set DCboSearch = Nothing
End Sub

Private Sub GetData()
    Dim StrSQL As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer

    StrSQL = "SELECT     dbo.TblVocation.ID, dbo.TblVocation.RecordDate, dbo.TblVocation.EmpID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee, "
    StrSQL = StrSQL + "                  dbo.TblVocation.FromDate, dbo.TblVocation.ToDate, dbo.TblVocation.Reson, dbo.TblVocation.Phone, dbo.TblVocation.Telephone, dbo.TblVocation.OtherAdress,"
    StrSQL = StrSQL + "                   dbo.TblVocation.TypeVocation, dbo.TblVocation.VocationType, dbo.TblVocation.ManagerApprove, dbo.TblVocation.VistCostOnCompany,"
    StrSQL = StrSQL + "                   dbo.TblVocation.VistCostOnEmployee, dbo.TblVocation.VisaCost, dbo.TblVocation.ToDateH, dbo.TblVocation.FromDateH, dbo.TblVocation.Approved,"
    StrSQL = StrSQL + "                   dbo.TblVocation.Adress, dbo.TblVocation.OutOnly, dbo.TblVocation.OutAndBack, dbo.TblVocation.ResumeWork, dbo.TblVocation.ResumeWorkH, dbo.TblVocation.ok,"
    StrSQL = StrSQL + "                   dbo.TblVocation.notok, dbo.TblVocation.ForEmployee, dbo.TblVocation.ForFamily, dbo.TblVocation.WithoutSalary, dbo.TblVocation.WithSalary,"
    StrSQL = StrSQL + "                   dbo.TblVocation.NoVacation ,dbo.TblVocation.TypeVacation "
    StrSQL = StrSQL + "            FROM         dbo.TblVocation LEFT OUTER JOIN"
    StrSQL = StrSQL + "                   dbo.TblEmployee ON dbo.TblVocation.EmpID = dbo.TblEmployee.Emp_ID"
    BolBegine = False
    StrWhere = ""

    If val(Me.TxtIDFrom.text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblVocation.ID >=" & val(Me.TxtIDFrom.text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblVocation.ID >=" & val(Me.TxtIDFrom.text) & ""
        End If
    End If

    If val(Me.TxtIDTO.text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblVocation.ID <=" & val(Me.TxtIDTO.text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblVocation.ID <=" & val(Me.TxtIDTO.text) & ""
        End If
    End If

    If Me.DcboEmpName.BoundText <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblVocation.EmpID=" & Me.DcboEmpName.BoundText & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblVocation.EmpID=" & Me.DcboEmpName.BoundText & ""
        End If
    End If



    If Not IsNull(Me.DtpDateFrom.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblVocation.RecordDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblVocation.RecordDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        End If
    End If

    If Not IsNull(Me.DtpDateTo.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblVocation.RecordDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblVocation.RecordDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        End If
    End If
''''
    If Not IsNull(Me.FromStartDate.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblVocation.FromDate >=" & SQLDate(Me.FromStartDate.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblVocation.FromDate >=" & SQLDate(Me.FromStartDate.value, True) & ""
        End If
    End If
        If Not IsNull(Me.ToStartDate.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblVocation.FromDate <=" & SQLDate(Me.ToStartDate.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblVocation.FromDate <=" & SQLDate(Me.ToStartDate.value, True) & ""
        End If
    End If
''''''
    '-----------------------------------
       If Not IsNull(Me.FromEndDate.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblVocation.ToDate >=" & SQLDate(Me.FromEndDate.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblVocation.ToDate >=" & SQLDate(Me.FromEndDate.value, True) & ""
        End If
    End If
       If Not IsNull(Me.ToEndDate.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblVocation.ToDate <=" & SQLDate(Me.ToEndDate.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblVocation.ToDate <=" & SQLDate(Me.ToEndDate.value, True) & ""
        End If
    End If
    ''''''''''''''''''''''''''
If Rd(1).value = True Then
      If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblVocation.ManagerApprove=1"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblVocation.ManagerApprove = 1"
        End If
End If
If Rd(2).value = True Then
      If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblVocation.ManagerApprove=0"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblVocation.ManagerApprove=0"
        End If
End If
'''''''''///////////////
If Rdb1(1).value = True Then
      If BolBegine = True Then
           StrWhere = StrWhere & " AND dbo.TblVocation.TypeVacation =0"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblVocation.TypeVacation =0"
        End If
End If
If Rdb1(2).value = True Then
      If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblVocation.TypeVacation =1"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblVocation.TypeVacation =1"
        End If
End If
If Rdb1(3).value = True Then
      If BolBegine = True Then
           StrWhere = StrWhere & " AND dbo.TblVocation.TypeVacation =2"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblVocation.TypeVacation =2"
       End If
End If
If index = 1 Then
   If BolBegine = True Then
           StrWhere = StrWhere & " AND 1=1"
        Else
            BolBegine = True
            StrWhere = " Where 1=1"
       End If
 Dim Scren As String
 Scren = "formvocatinl"
 If CheckAprroveScreen("formvocatinl") = True Then
StrWhere = StrWhere & " and   (dbo.ScreenSendAparoved(dbo.TblVocation.ID, '" & Scren & "') > 0)"
StrWhere = StrWhere & " and   (dbo.ScreenIsAparoved(dbo.TblVocation.ID, '" & Scren & "') is null)"
End If
 End If



''''''''''''''''''


    StrSQL = StrSQL & StrWhere
    StrSQL = StrSQL & " Order By dbo.TblVocation.ID"
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            Me.lbl(10).Caption = "Search Results=0"
        End If

        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «·»ÕÀ"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    Else

        With Me.Fg
            .Clear flexClearScrollable, flexClearEverything
            .rows = .FixedRows
            .rows = rs.RecordCount + .FixedRows

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
                Me.lbl(10).Caption = "Search Results=" & rs.RecordCount
            End If

            rs.MoveFirst
        
            For i = .FixedRows To .rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("ID")) = IIf(IsNull(rs("ID").value), "", rs("ID").value)
                        
                If Not (IsNull(rs("RecordDate").value)) Then
                    .TextMatrix(i, .ColIndex("RecordDate")) = Format(rs("RecordDate").value, "yyyy/M/d")
                End If
            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(rs("Emp_Name").value), "", rs("Emp_Name").value)
                Else
                .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(rs("Emp_Namee").value), "", rs("Emp_Namee").value)
                End If
                .TextMatrix(i, .ColIndex("Fullcode")) = IIf(IsNull(rs("Fullcode").value), "", rs("Fullcode").value)
                .TextMatrix(i, .ColIndex("EmpID")) = IIf(IsNull(rs("EmpID").value), "", rs("EmpID").value)
                .TextMatrix(i, .ColIndex("FromDate")) = IIf(IsNull(rs("FromDate").value), "", rs("FromDate").value)
                .TextMatrix(i, .ColIndex("FromDateH")) = IIf(IsNull(rs("FromDateH").value), "", rs("FromDateH").value)
                
                .TextMatrix(i, .ColIndex("ToDate")) = IIf(IsNull(rs("ToDate").value), "", rs("ToDate").value)
                .TextMatrix(i, .ColIndex("ToDateH")) = IIf(IsNull(rs("ToDateH").value), "", rs("ToDateH").value)
                If Not IsNull(rs("ManagerApprove").value) Then
                If rs("ManagerApprove").value = True Then
                If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("ManagerApprove")) = " „  «·„Ê«›ﬁ…"
                Else
                .TextMatrix(i, .ColIndex("ManagerApprove")) = " Accept"
                End If
                Else
                If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("ManagerApprove")) = "·„   „ «·„Ê«›ﬁ…"
                Else
                .TextMatrix(i, .ColIndex("ManagerApprove")) = "Not Accept"
                End If
                End If
                End If
                If Not (IsNull(rs("TypeVacation").value)) Then
                If SystemOptions.UserInterface = EnglishInterface Then
                Select Case val(rs("TypeVacation").value)
                Case 0
                .TextMatrix(i, .ColIndex("VocationType")) = "Official"
                Case 1
                .TextMatrix(i, .ColIndex("VocationType")) = "Important"
                Case 2
                .TextMatrix(i, .ColIndex("VocationType")) = "Other"
                End Select
                Else
                 Select Case val(rs("TypeVacation").value)
                Case 0
                .TextMatrix(i, .ColIndex("VocationType")) = "—”„Ì…"
                Case 1
                .TextMatrix(i, .ColIndex("VocationType")) = "≈Ÿÿ—«—Ì…"
                Case 2
                .TextMatrix(i, .ColIndex("VocationType")) = "≈Œ—Ï"
                End Select
                End If
                
                End If
                               
                rs.MoveNext
            Next i

            .AutoSize 0, .Cols - 1, False
            Me.lbl(1).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("ID"), .rows - 1, .ColIndex("ID"))
        End With

    End If

End Sub

Private Sub FromEndDate_Change()
If Not (IsNull(FromEndDate.value)) Then
FromEndDateH.value = ToHijriDate(FromEndDate.value)
End If
End Sub

Private Sub FromEndDateH_LostFocus()
VBA.Calendar = vbCalGreg
            FromEndDate.value = ToGregorianDate(FromEndDateH.value)
End Sub

Private Sub FromStartDate_Change()
If Not (IsNull(FromStartDate.value)) Then
FromStartDateH.value = ToHijriDate(FromStartDate.value)
End If
End Sub

Private Sub FromStartDateH_LostFocus()
VBA.Calendar = vbCalGreg
            FromStartDate.value = ToGregorianDate(FromStartDateH.value)
End Sub

Private Sub ToEndDate_Change()
If Not (IsNull(ToEndDate.value)) Then
ToEndDateH.value = ToHijriDate(ToEndDate.value)
End If
End Sub

Private Sub ToEndDateH_LostFocus()
VBA.Calendar = vbCalGreg
            ToEndDate.value = ToGregorianDate(ToEndDateH.value)
End Sub

Private Sub ToStartDate_Change()
If Not (IsNull(ToStartDate.value)) Then
ToStartDateH.value = ToHijriDate(ToStartDate.value)
End If
End Sub

Private Sub ToStartDateH_LostFocus()
VBA.Calendar = vbCalGreg
            ToStartDate.value = ToGregorianDate(ToStartDateH.value)
End Sub

Private Sub TxtIDFrom_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtIDFrom.text, 1)
End Sub

Private Sub TxtIDTO_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtIDTO.text, 1)
End Sub

Private Sub TxtSearchCode_KeyPress(KeyAscii As Integer)
    Dim EmpID As Integer
 
    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode TxtSearchCode.text, EmpID
        DcboEmpName.BoundText = EmpID
    End If
End Sub
