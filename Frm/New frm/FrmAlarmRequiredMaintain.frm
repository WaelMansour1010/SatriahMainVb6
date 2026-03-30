VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmAlarmRequiredMaintain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ‰»ÌÂ«  ÿ·»«  «·’Ì«‰…"
   ClientHeight    =   8820
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14535
   Icon            =   "FrmAlarmRequiredMaintain.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8820
   ScaleWidth      =   14535
   Begin VB.Frame lbprocess 
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·ÿ·»"
      ForeColor       =   &H00FF0000&
      Height          =   768
      Left            =   5040
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   840
      Visible         =   0   'False
      Width           =   3435
      Begin VB.TextBox TxtIDTO 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   240
         Width           =   915
      End
      Begin VB.TextBox TxtIDFrom 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   1680
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   240
         Width           =   915
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï"
         Height          =   195
         Index           =   6
         Left            =   1020
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   240
         Width           =   525
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   195
         Index           =   5
         Left            =   2535
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   240
         Width           =   540
      End
   End
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   0
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E2E9E9&
      Height          =   855
      Left            =   0
      TabIndex        =   5
      Top             =   8040
      Width           =   14535
      Begin ImpulseButton.ISButton Cmd 
         Height          =   495
         Index           =   6
         Left            =   480
         TabIndex        =   6
         Top             =   240
         Width           =   3045
         _ExtentX        =   5371
         _ExtentY        =   873
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
         ButtonImage     =   "FrmAlarmRequiredMaintain.frx":6852
         ColorButton     =   14871017
         ColorHighlight  =   16777215
         ColorHoverText  =   16711680
         ColorShadow     =   -2147483637
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         ColorToggledHoverText=   16711680
         ColorTextShadow =   -2147483637
      End
      Begin ImpulseButton.ISButton CmdHelp 
         Height          =   495
         Left            =   5040
         TabIndex        =   7
         Top             =   240
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   873
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
         ButtonImage     =   "FrmAlarmRequiredMaintain.frx":30474
         ColorButton     =   14871017
         ColorHighlight  =   16777215
         ColorHoverText  =   16711680
         ColorShadow     =   -2147483637
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         ColorToggledHoverText=   16711680
         ColorTextShadow =   -2147483637
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   495
         Index           =   9
         Left            =   9240
         TabIndex        =   8
         Top             =   240
         Width           =   3045
         _ExtentX        =   5371
         _ExtentY        =   873
         ButtonPositionImage=   1
         Caption         =   "ÿ»«⁄Â"
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
         ButtonImage     =   "FrmAlarmRequiredMaintain.frx":36CD6
         ColorButton     =   14871017
         ColorHighlight  =   16777215
         ColorHoverText  =   16711680
         ColorShadow     =   -2147483637
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         ColorToggledHoverText=   16711680
         ColorTextShadow =   -2147483637
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E2E9E9&
      Height          =   6495
      Left            =   0
      TabIndex        =   4
      Top             =   1680
      Width           =   14535
      Begin VSFlex8UCtl.VSFlexGrid Fg 
         Height          =   6255
         Left            =   120
         TabIndex        =   17
         Top             =   120
         Width           =   14295
         _cx             =   25215
         _cy             =   11033
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
         BackColorAlternate=   16777152
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
         Cols            =   9
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FrmAlarmRequiredMaintain.frx":3D538
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
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E2E9E9&
      Height          =   1095
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   600
      Width           =   14535
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4320
         TabIndex        =   14
         Text            =   "1"
         Top             =   600
         Width           =   336
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·› —Â"
         Height          =   735
         Left            =   8520
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   240
         Width           =   5895
         Begin MSComCtl2.DTPicker DtpDateFrom 
            Height          =   330
            Left            =   3000
            TabIndex        =   10
            Top             =   240
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   96665601
            CurrentDate     =   41640
         End
         Begin MSComCtl2.DTPicker DtpDateTo 
            Height          =   330
            Left            =   360
            TabIndex        =   11
            Top             =   240
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   96665601
            CurrentDate     =   41640
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "„‰"
            Height          =   315
            Index           =   0
            Left            =   4680
            RightToLeft     =   -1  'True
            TabIndex        =   13
            Top             =   240
            Width           =   585
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "≈«·Ï"
            Height          =   435
            Index           =   14
            Left            =   2100
            RightToLeft     =   -1  'True
            TabIndex        =   12
            Top             =   240
            Width           =   540
         End
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   735
         Index           =   5
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   1296
         ButtonPositionImage=   1
         Caption         =   " ÕœÌÀ"
         BackColor       =   14871017
         FontSize        =   14.25
         FontName        =   "Times New Roman"
         FontBold        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmAlarmRequiredMaintain.frx":3D691
         ColorButton     =   14871017
         ColorHighlight  =   16777215
         ColorHoverText  =   12632064
         ColorShadow     =   -2147483637
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         ColorToggledHoverText=   12632064
         LowerToggledContent=   0   'False
         ColorTextShadow =   -2147483637
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "  œﬁÌﬁ…"
         Height          =   288
         Index           =   2
         Left            =   3600
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   600
         Width           =   648
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   " ÕœÌÀ ﬂ·"
         Height          =   432
         Index           =   4
         Left            =   3720
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   240
         Width           =   780
      End
   End
   Begin C1SizerLibCtl.C1Elastic EleHeader 
      Height          =   585
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   14565
      _cx             =   25691
      _cy             =   1032
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   21.75
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
      BackColor       =   16777215
      ForeColor       =   4210688
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   " ‰»ÌÂ«  ÿ·»«  «·’Ì«‰…  "
      Align           =   0
      AutoSizeChildren=   0
      BorderWidth     =   0
      ChildSpacing    =   0
      Splitter        =   0   'False
      FloodDirection  =   0
      FloodPercent    =   0
      CaptionPos      =   6
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
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H000000FF&
         Height          =   555
         Index           =   27
         Left            =   2520
         TabIndex        =   2
         Top             =   0
         Width           =   2205
      End
   End
End
Attribute VB_Name = "FrmAlarmRequiredMaintain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Cmd_Click(Index As Integer)
Select Case Index
Case 5
'ProgressBar1.Visible = True
': ProgressBar1.value = 10
fillgrid
': ProgressBar1.value = 50
'ProgressBar1.Visible = False
'ProgressBar1.value = 0
Case 6
Me.Hide
Case 9
GetData
End Select

End Sub
Public Sub GetData()

  print_report

End Sub
Function print_report(Optional NoteSerial As String)
     
     
    
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String



MySQL = "  SELECT dbo.TblRequerMainten.ID, dbo.TblRequerMainten.ProblemTimID, dbo.TblRequerMainten.ProblemOther, dbo.TblRequerMainten.StopTime, dbo.TblRequerMainten.StartTime,"
MySQL = MySQL & "           dbo.TblRequerMainten.Des, dbo.TblRequerMainten.Remarks, dbo.TblRequerMainten.RecordDate, dbo.TblRequerMainten.StartDate, dbo.TblRequerMainten.StopDate,"
MySQL = MySQL & "                         dbo.TblRequerMainten.UnitID, dbo.TblEmpDepartments.DepartmentName, dbo.TblEmpDepartments.DepartmentNamee, dbo.TblRequerMainten.EquepID,"
 MySQL = MySQL & "                        dbo.FixedAssets.Name, dbo.FixedAssets.code, dbo.TblRequerMainten.BranchID, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee,"
  MySQL = MySQL & "                       dbo.FixedAssets.NameE , TblOrderMaint.ReqMainID"
 MySQL = MySQL & "      FROM     dbo.TblRequerMainten LEFT OUTER JOIN"
 MySQL = MySQL & "                        dbo.TblBranchesData ON dbo.TblRequerMainten.BranchID = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
 MySQL = MySQL & "                        dbo.FixedAssets ON dbo.TblRequerMainten.EquepID = dbo.FixedAssets.id LEFT OUTER JOIN"
MySQL = MySQL & "                         dbo.TblEmpDepartments ON dbo.TblRequerMainten.UnitID = dbo.TblEmpDepartments.DeparmentID"
 MySQL = MySQL & "                        left join TblOrderMaint on dbo.TblRequerMainten.id = dbo.TblOrderMaint.ReqMainID"
 MySQL = MySQL & "      Where dbo.TblOrderMaint.ReqMainID Is Null"




    If val(Me.TxtIDFrom.text) <> 0 Then
        'If BolBegine = True Then
            MySQL = MySQL & " and  dbo.TblRequerMainten.id >=" & val(Me.TxtIDFrom.text) & ""
       ' Else
       '     BolBegine = True
       '     MySQL = " and dbo.TblRequerMainten.id >=" & val(Me.TxtIDFrom.text) & ""
       ' End If
    End If
   

    If val(Me.TxtIDTO.text) <> 0 Then
       ' If BolBegine = True Then
            MySQL = MySQL & " AND dbo.TblRequerMainten.id <=" & val(Me.TxtIDTO.text) & ""
       ' Else
       '     BolBegine = True
       '     MySQL = "  and  dbo.TblRequerMainten.id <=" & val(Me.TxtIDTO.text) & ""
       ' End If
    End If


   If Not IsNull(Me.DtpDateFrom.value) Then
        'If BolBegine = True Then
            MySQL = MySQL & " AND dbo.TblRequerMainten.RecordDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
       ' Else
       '     BolBegine = True
       '     MySQL = "  and  dbo.TblRequerMainten.RecordDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
       ' End If
    End If

    If Not IsNull(Me.DtpDateTo.value) Then
       ' If BolBegine = True Then
            MySQL = MySQL & " AND  dbo.TblRequerMainten.RecordDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
       ' Else
       '     BolBegine = True
       '    MySQL = "  and  dbo.TblRequerMainten.RecordDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
       ' End If
    End If





 If SystemOptions.UserInterface = ArabicInterface Then
          StrFileName = App.path & "\REPORTS\REPORTS NEW\rpt_ClosedRequirMaintin.rpt"
     Else
        StrFileName = App.path & "\REPORTS\REPORTS NEW\rpt_ClosedRequirMaintinE.rpt"
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
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        RsData.Close
        Set RsData = Nothing
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)
    xReport.Database.SetDataSource RsData

    Dim cCompanyInfo As New ClsCompanyInfo

Dim Sal  As Double
'Sal = GetEmployeeSalaryAccordingToComponent(val(DcboEmpName.BoundText), "", 0)

    If SystemOptions.UserInterface = ArabicInterface Then
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
         xReport.ParameterFields(2).AddCurrentValue ""
       
     
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.EngCompanyName  ' RPTCompany_Name_Eng
        xReport.ParameterFields(2).AddCurrentValue ""
       
   
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

Private Sub CmdHelp_Click()
 
          clear_all Me
            FG.Clear flexClearScrollable, flexClearEverything
            FG.Rows = 1
DtpDateFrom.value = Date
DtpDateTo.value = Date
End Sub

Private Sub DBCboClientName_Change()
    TxtSearchCode.text = ""

    Dim DefaultSalesPersonId As Integer
    Dim fullcode As String

    GetCustomersDetail val(DBCboClientName.BoundText), DefaultSalesPersonId, fullcode

    TxtSearchCode.text = fullcode
 
End Sub
Function cahngelang()
'    EleHeader.Caption = " The Amounts Reserved and the Delivered and the Remaining "
'    Me.Caption = EleHeader.Caption
''    lbl(3).Caption = "Customer"
'    lbl(1).Caption = "Item"
'    lbl(0).Caption = "From"
'    lbl(14).Caption = "To"
'    Frame5.Caption = "Period"
'    lbl(4).Caption = "Update All"
'   CmdHelp.Caption = "Clear"
   
'   Cmd(5).Caption = "Saerch"
'   Cmd(9).Caption = "Print"
'    Cmd(6).Caption = "Exit"
'    With GridInstallments
'    .TextMatrix(0, .ColIndex("Ser")) = "Serial"
'    .TextMatrix(0, .ColIndex("Transaction_Date")) = "Date"
'    .TextMatrix(0, .ColIndex("CusName")) = "Customer"
'    .TextMatrix(0, .ColIndex("ItemName")) = "ItemName"
'    .TextMatrix(0, .ColIndex("mah")) = "Reserved"
'    .TextMatrix(0, .ColIndex("ms")) = "Delivered"
'    .TextMatrix(0, .ColIndex("mt")) = "Remaining "

   
'    End With

' Cmd(1).Caption = "Delete"
 '   Cmd(0).Caption = "Search"
   Cmd(6).Caption = "Exit"
    EleHeader.Caption = "Required Maintenance alarm"
  Me.Caption = "Search Required Maintenance"
lbprocess.Caption = "OrderNo"
lbl(5).Caption = "From"
lbl(6).Caption = "To"
lbl(0).Caption = "To"
Cmd(9).Caption = "Print"
lbl(4).Caption = "Refresh"
lbl(2).Caption = " / Minute "
CmdHelp.Caption = "Clear"
Frame5.Caption = "Period"
lbl(0).Caption = "From"
lbl(14).Caption = "To"
Cmd(5).Caption = "Refresh"



     With Me.FG
        .TextMatrix(0, .ColIndex("Serial")) = "NO"
        .TextMatrix(0, .ColIndex("id")) = "OrderNo "
        .TextMatrix(0, .ColIndex("RecordDate")) = "Date"
         .TextMatrix(0, .ColIndex("branch_name")) = "BranchName"
        .TextMatrix(0, .ColIndex("DepartmentName")) = "DepartmentName"
       .TextMatrix(0, .ColIndex("Name")) = "Machine"
        .TextMatrix(0, .ColIndex("Remarks")) = "Remarks"
       .TextMatrix(0, .ColIndex("des")) = "Failure"
        .TextMatrix(0, .ColIndex("ProblemTimID")) = "TimeProblem"
    End With
  '

End Function
Public Sub RetriveOrder(Optional IDde As Integer = 0, Optional order_no As String = "")
    Dim RsDetails As New ADODB.Recordset
    Dim StrSQL As String
    Dim RsNotes As New ADODB.Recordset
    Dim RsTemp As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Dim Num As Long
   'sa On Error GoTo ErrTrap

    StrSQL = "Select * from transactions  where   ResProductionNo='" & order_no & "'"

    
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If rs.RecordCount < 1 Then
   
 
        Exit Sub
    Else
    Dim IDde2 As Integer
    IDde2 = IDde
      
   RetriveOrder2 IDde2, rs("Transaction_serial").value
    End If

    If rs.EOF Or rs.BOF Then
        Exit Sub
    End If

 
    StrSQL = "SELECT TblItems.HaveSerial, * FROM TblItems INNER JOIN Transaction_Details " & "ON TblItems.ItemID = Transaction_Details.Item_ID INNER JOIN dbo.TblUnites ON dbo.Transaction_Details.UnitID = dbo.TblUnites.UnitID"
    StrSQL = StrSQL + " where Transaction_ID=" & val(rs("Transaction_ID").value)

    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
   
    If Not (RsDetails.EOF Or RsDetails.BOF) Then
    With GridInstallments
    Dim j As Integer
    For j = 1 To RsDetails.RecordCount
    If val(.TextMatrix(IDde, .ColIndex("id"))) = 0 Then
                  .TextMatrix(IDde, .ColIndex("mt")) = IIf(IsNull(RsDetails("showqty")), "", (RsDetails("showqty").value))
                  .TextMatrix(IDde, .ColIndex("id")) = IIf(IsNull(RsDetails("id")), "", (RsDetails("id").value))
                  IDde = IDde + 1
                  
                  RsDetails.MoveNext
                  End If
                  Next j
                  End With
        End If


End Sub
Public Sub RetriveOrder2(Optional IDde As Integer = 0, Optional order_no As String = "")
    Dim RsDetails As New ADODB.Recordset
    Dim StrSQL As String
    Dim RsNotes As New ADODB.Recordset
    Dim RsTemp As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Dim Num As Long
   'sa On Error GoTo ErrTrap

    StrSQL = "Select * from transactions  where   ProductionOrder='" & order_no & "'"

    
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If rs.RecordCount < 1 Then

 
        Exit Sub
    Else
    
    End If

    If rs.EOF Or rs.BOF Then
        Exit Sub
    End If

 
    StrSQL = "SELECT TblItems.HaveSerial, * FROM TblItems INNER JOIN Transaction_Details " & "ON TblItems.ItemID = Transaction_Details.Item_ID INNER JOIN dbo.TblUnites ON dbo.Transaction_Details.UnitID = dbo.TblUnites.UnitID"
    StrSQL = StrSQL + " where Transaction_ID=" & val(rs("Transaction_ID").value)

    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
   
    If Not (RsDetails.EOF Or RsDetails.BOF) Then
    With GridInstallments
    Dim j As Integer
    For j = 1 To RsDetails.RecordCount
    If val(.TextMatrix(IDde, .ColIndex("idd"))) = 0 Then
                  .TextMatrix(IDde, .ColIndex("ms")) = IIf(IsNull(RsDetails("RecivedShippedQty")), "", (RsDetails("RecivedShippedQty").value))
                  .TextMatrix(IDde, .ColIndex("idd")) = IIf(IsNull(RsDetails("id")), "", (RsDetails("id").value))
                  IDde = IDde + 1
                  
                  RsDetails.MoveNext
                  End If
                  Next j
                  End With
        End If


End Sub
Public Sub fillgrid(Optional str As String)

    Dim StrSQL As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
    Dim Problem_Time As String


StrSQL = " SELECT dbo.TblRequerMainten.ID, dbo.TblRequerMainten.ProblemTimID, dbo.TblRequerMainten.ProblemOther, dbo.TblRequerMainten.StopTime, dbo.TblRequerMainten.StartTime,"
StrSQL = StrSQL & "                 dbo.TblRequerMainten.Des, dbo.TblRequerMainten.Remarks, dbo.TblRequerMainten.RecordDate, dbo.TblRequerMainten.StartDate, dbo.TblRequerMainten.StopDate,"
StrSQL = StrSQL & "        dbo.TblRequerMainten.UnitID, dbo.TblEmpDepartments.DepartmentName, dbo.TblEmpDepartments.DepartmentNamee, dbo.TblRequerMainten.EquepID,"
StrSQL = StrSQL & "                       dbo.FixedAssets.Name, dbo.FixedAssets.code, dbo.TblRequerMainten.BranchID, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee,"
StrSQL = StrSQL & "                       dbo.FixedAssets.NameE , TblOrderMaint.ReqMainID"
 StrSQL = StrSQL & "    FROM     dbo.TblRequerMainten LEFT OUTER JOIN"
 StrSQL = StrSQL & "                      dbo.TblBranchesData ON dbo.TblRequerMainten.BranchID = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
  StrSQL = StrSQL & "                     dbo.FixedAssets ON dbo.TblRequerMainten.EquepID = dbo.FixedAssets.id LEFT OUTER JOIN"
    StrSQL = StrSQL & "                   dbo.TblEmpDepartments ON dbo.TblRequerMainten.UnitID = dbo.TblEmpDepartments.DeparmentID"
  StrSQL = StrSQL & "                     left join TblOrderMaint on dbo.TblRequerMainten.id = dbo.TblOrderMaint.ReqMainID"
 StrSQL = StrSQL & "    Where dbo.TblOrderMaint.ReqMainID Is Null"

 BolBegine = False
    StrWhere = ""

    If val(Me.TxtIDFrom.text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " and  dbo.TblRequerMainten.id >=" & val(Me.TxtIDFrom.text) & ""
        Else
            BolBegine = True
            StrWhere = " and dbo.TblRequerMainten.id >=" & val(Me.TxtIDFrom.text) & ""
        End If
    End If
   

    If val(Me.TxtIDTO.text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblRequerMainten.id <=" & val(Me.TxtIDTO.text) & ""
        Else
            BolBegine = True
            StrWhere = "  and  dbo.TblRequerMainten.id <=" & val(Me.TxtIDTO.text) & ""
        End If
    End If

  



   If Not IsNull(Me.DtpDateFrom.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblRequerMainten.RecordDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        Else
            BolBegine = True
            StrWhere = "  and  dbo.TblRequerMainten.RecordDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        End If
    End If

    If Not IsNull(Me.DtpDateTo.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND  dbo.TblRequerMainten.RecordDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        Else
            BolBegine = True
           StrWhere = "  and  dbo.TblRequerMainten.RecordDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        End If
    End If

    '-----------------------------------

    StrSQL = StrSQL & StrWhere
    StrSQL = StrSQL & " Order By dbo.TblRequerMainten.id "
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
           ' Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            'Me.lbl(10).Caption = "Search Results=0"
        End If

    '    Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «·»ÕÀ"
        Cmd_Click (1)
    '    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        FG.Rows = FG.FixedRows
        Exit Sub
    Else

        With Me.FG
            .Clear flexClearScrollable, flexClearEverything
            .Rows = .FixedRows
            .Rows = rs.RecordCount + .FixedRows

            If SystemOptions.UserInterface = ArabicInterface Then
'                Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
'                Me.lbl(10).Caption = "Search Results=" & rs.RecordCount
            End If

            rs.MoveFirst
        
            For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                             
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs("id").value), "", rs("id").value)
                        
                If Not (IsNull(rs("RecordDate").value)) Then
                    .TextMatrix(i, .ColIndex("RecordDate")) = Format(rs("RecordDate").value, "yyyy/M/d")
                End If
                If SystemOptions.UserInterface = EnglishInterface Then
                .TextMatrix(i, .ColIndex("DepartmentName")) = IIf(IsNull(rs("DepartmentNamee").value), "", rs("DepartmentNamee").value)
                .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(rs("branch_namee").value), "", rs("branch_namee").value)
                .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs("namee").value), "", rs("namee").value)
                .TextMatrix(i, .ColIndex("des")) = IIf(IsNull(rs("des").value), "", rs("des").value)
               Problem_Time = ""
               
               If rs("ProblemTimID").value = 0 Then
               Problem_Time = "«À‰«¡ «· ’‰Ì⁄"
               ElseIf rs("ProblemTimID").value = 1 Then
               Problem_Time = "«À‰«¡ »œ¡ «· ‘€Ì·"
               ElseIf rs("ProblemTimID").value = 2 Then
               Problem_Time = "«À‰«¡ «·«’·«Õ"
               ElseIf rs("ProblemTimID").value = 3 Then
               Problem_Time = "«Œ—Ï"
               End If
               .TextMatrix(i, .ColIndex("ProblemTimID")) = Problem_Time
                
                Else
            .TextMatrix(i, .ColIndex("DepartmentName")) = IIf(IsNull(rs("DepartmentName").value), "", rs("DepartmentName").value)
                .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(rs("branch_name").value), "", rs("branch_name").value)
                .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs("Name").value), "", rs("Name").value)
                .TextMatrix(i, .ColIndex("des")) = IIf(IsNull(rs("des").value), "", rs("des").value)
                 Problem_Time = ""
               
               If rs("ProblemTimID").value = 0 Then
               Problem_Time = "During Production"
               ElseIf rs("ProblemTimID").value = 1 Then
               Problem_Time = "During Start up"
               ElseIf rs("ProblemTimID").value = 2 Then
               Problem_Time = "During Repair"
               ElseIf rs("ProblemTimID").value = 3 Then
               Problem_Time = "Others"
               End If
            .TextMatrix(i, .ColIndex("ProblemTimID")) = Problem_Time
               End If
      

               ' ProblemTimID1.ListIndex = val(IIf(IsNull(rs("ProblemTimID").value), -1, rs("ProblemTimID").value))
              '  .TextMatrix(i, .ColIndex("ProblemTimID")) = ProblemTimID1.text
                rs.MoveNext
            Next i

            .AutoSize 0, .Cols - 1, False
         '   Me.lbl(1).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("AdvanceValue"), .Rows - 1, .ColIndex("AdvanceValue"))
        End With

    End If



End Sub

Private Sub DcboItems_Change()
      ' Me.TxtItemCode.text = GetItemCode(val(Me.DcboItems.BoundText))
  End Sub





Private Sub TxtItemCode_KeyDown(KeyCode As Integer, Shift As Integer)
 

End Sub

Private Sub DtpDateFrom_Change()
Cmd_Click (5)
End Sub

Private Sub DtpDateTo_Change()
Cmd_Click (5)
End Sub

Private Sub Fg_DblClick()
On Error GoTo ErrTrap

Dim i As Integer
i = val(Me.FG.TextMatrix(Me.FG.Row, Me.FG.ColIndex("id")))
If i <> 0 Then
Load FrmRequerMainten
    FrmRequerMainten.Retrive (i)
    End If
ErrTrap:
End Sub

Private Sub Form_Load()
   Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500
    Dim Dcombos As New ClsDataCombos
      Set Dcombos = New ClsDataCombos
    'Dcombos.GetItemsNames Me.DcboItems
   'Dcombos.GetCustomersSuppliers 1, Me.DBCboClientName, True  '  2 supplier  1 customer
      

    If SystemOptions.UserInterface = EnglishInterface Then

        SetInterface Me
       cahngelang
    End If

'FromDate.value = Date
'toDate.value = Date
DtpDateFrom.value = Date
DtpDateTo.value = Date

fillgrid


'DtpDateFrom.value = Null
'DtpDateTo.value = Null
    
    
    'set timer interval
    Dim i As Integer
i = val(Text1.text)
Timer1.interval = 60000 * i
 
End Sub



Private Sub TxtSearchCode_KeyPress(KeyAscii As Integer)
    Dim CUSTID As Integer

    If KeyAscii = vbKeyReturn Then
        GetCustomersDetail CUSTID, , TxtSearchCode.text, 1
        DBCboClientName.BoundText = CUSTID
    End If

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)

If Not (KeyAscii >= 48 And KeyAscii <= 57) Then
KeyAscii = 0
End If

End Sub

Private Sub Timer1_Timer()

 

Cmd_Click (5)
End Sub
