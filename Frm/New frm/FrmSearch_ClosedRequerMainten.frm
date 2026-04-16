VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmSearch_ClosedRequerMainten 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "«·»ÕÀ ⁄‰ ÿ·»«  «·’Ì«‰…"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12465
   Icon            =   "FrmSearch_ClosedRequerMainten.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   4500
   ScaleWidth      =   12465
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
   Begin VB.TextBox txtTimer 
      Alignment       =   1  'Right Justify
      Height          =   372
      Left            =   6360
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Text            =   "5"
      Top             =   3000
      Width           =   492
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   "›Ì «·› —…"
      ForeColor       =   &H00FF0000&
      Height          =   645
      Index           =   1
      Left            =   8280
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   2700
      Width           =   4125
      Begin MSComCtl2.DTPicker DtpDateFrom 
         Height          =   345
         Left            =   2100
         TabIndex        =   21
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   609
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/m/yyyy"
         DateIsNull      =   -1  'True
         Format          =   104529921
         CurrentDate     =   38979.743287037
      End
      Begin MSComCtl2.DTPicker DtpDateTo 
         Height          =   375
         Left            =   60
         TabIndex        =   22
         Top             =   270
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   104529921
         CurrentDate     =   38784
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï"
         Height          =   285
         Index           =   0
         Left            =   1620
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   315
         Width           =   345
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   285
         Index           =   11
         Left            =   3720
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   255
         Width           =   285
      End
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Height          =   645
      Index           =   0
      Left            =   2760
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   4440
      Width           =   10635
      Begin VB.ComboBox ProblemTimID 
         Height          =   315
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   240
         Width           =   1815
      End
      Begin MSDataListLib.DataCombo Dcbranch 
         Bindings        =   "FrmSearch_ClosedRequerMainten.frx":038A
         Height          =   315
         Left            =   8160
         TabIndex        =   12
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
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
      Begin MSDataListLib.DataCombo DcbUnit 
         Bindings        =   "FrmSearch_ClosedRequerMainten.frx":039F
         Height          =   315
         Left            =   5520
         TabIndex        =   14
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
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
      Begin MSDataListLib.DataCombo DcbEquepment 
         Height          =   315
         Left            =   3000
         TabIndex        =   16
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Êﬁ  «·„‘ﬂ·…"
         Height          =   255
         Index           =   7
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«·„⁄œÂ"
         Height          =   255
         Index           =   4
         Left            =   4560
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«·ﬁ”„"
         Height          =   255
         Index           =   3
         Left            =   7200
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   300
         Width           =   855
      End
      Begin VB.Label lblbr 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«·›—⁄"
         Height          =   255
         Left            =   9720
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   300
         Width           =   855
      End
   End
   Begin VB.Frame lbprocess 
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·ÿ·»"
      ForeColor       =   &H00FF0000&
      Height          =   645
      Left            =   9360
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   4440
      Width           =   3435
      Begin VB.TextBox TxtIDFrom 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   1680
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   240
         Width           =   915
      End
      Begin VB.TextBox TxtIDTO 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   915
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   195
         Index           =   5
         Left            =   2535
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   240
         Width           =   540
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï"
         Height          =   195
         Index           =   6
         Left            =   1020
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   240
         Width           =   525
      End
   End
   Begin VSFlex8UCtl.VSFlexGrid Fg 
      Height          =   2628
      Left            =   36
      TabIndex        =   5
      Top             =   0
      Width           =   12432
      _cx             =   21929
      _cy             =   4636
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
      Cols            =   9
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmSearch_ClosedRequerMainten.frx":03B4
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
      Height          =   372
      Index           =   0
      Left            =   2496
      TabIndex        =   6
      Top             =   4080
      Width           =   768
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
      TabIndex        =   7
      Top             =   4080
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
      TabIndex        =   8
      Top             =   4080
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
   Begin ImpulseButton.ISButton Cmd 
      Height          =   372
      Index           =   3
      Left            =   1680
      TabIndex        =   26
      Top             =   4080
      Width           =   768
      _ExtentX        =   1349
      _ExtentY        =   661
      ButtonPositionImage=   1
      Caption         =   "ÿ»«⁄…"
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
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«· ÕœÌÀ"
      Height          =   288
      Index           =   8
      Left            =   7080
      RightToLeft     =   -1  'True
      TabIndex        =   28
      Top             =   3000
      Width           =   648
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "\  œﬁÌﬁ…"
      Height          =   288
      Index           =   2
      Left            =   5520
      RightToLeft     =   -1  'True
      TabIndex        =   27
      Top             =   3000
      Width           =   648
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      ForeColor       =   &H00000080&
      Height          =   285
      Index           =   1
      Left            =   60
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   2940
      Width           =   1545
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      ForeColor       =   &H00000080&
      Height          =   315
      Index           =   10
      Left            =   60
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   2700
      Width           =   2775
   End
End
Attribute VB_Name = "FrmSearch_ClosedRequerMainten"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim DCboSearch As clsDCboSearch

Private Sub Cmd_Click(Index As Integer)

    Select Case Index

        Case 0
    
 GetData
           
        Case 1
            clear_all Me
Me.DtpDateFrom.value = ""
Me.DtpDateTo.value = ""
            If SystemOptions.UserInterface = ArabicInterface Then
               ' Me.lbl(0).Caption = "‰ ÌÃ… «·»ÕÀ"
            Else
               ' Me.lbl(0).Caption = "Search Results"
            End If

        Case 2
            Unload Me
        
        Case 3
            print_report
            
    End Select

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
MySQL = MySQL & "                  dbo.TblRequerMainten.Des, dbo.TblRequerMainten.Remarks, dbo.TblRequerMainten.RecordDate, dbo.TblRequerMainten.StartDate, dbo.TblRequerMainten.StopDate,"
MySQL = MySQL & "                      dbo.TblRequerMainten.UnitID, dbo.TblEmpDepartments.DepartmentName, dbo.TblEmpDepartments.DepartmentNamee, dbo.TblRequerMainten.EquepID,"
MySQL = MySQL & "                      dbo.FixedAssets.Name, dbo.FixedAssets.code, dbo.TblRequerMainten.BranchID, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee,"
 MySQL = MySQL & "                     dbo.FixedAssets.NameE"
MySQL = MySQL & "    FROM     dbo.TblRequerMainten LEFT OUTER JOIN"
    MySQL = MySQL & "                  dbo.TblBranchesData ON dbo.TblRequerMainten.BranchID = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
   MySQL = MySQL & "                   dbo.FixedAssets ON dbo.TblRequerMainten.EquepID = dbo.FixedAssets.id LEFT OUTER JOIN"
   MySQL = MySQL & "                   dbo.TblEmpDepartments ON dbo.TblRequerMainten.UnitID = dbo.TblEmpDepartments.DeparmentID CROSS JOIN"
  MySQL = MySQL & "                    dbo.TblOrderMaint"
MySQL = MySQL & "    Where (Not (dbo.TblRequerMainten.id = dbo.TblOrderMaint.ReqMainID))"




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
          StrFileName = App.path & "\REPORTS\REPORTS NEW\Special\" & Report_Folder & "\rpt_ClosedRequirMaintin.rpt"
     Else
        StrFileName = App.path & "\REPORTS\REPORTS NEW\Special\" & Report_Folder & "\rpt_ClosedRequirMaintinE.rpt"
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







Private Sub Fg_DblClick()
On Error GoTo ErrTrap

Dim i As Integer
i = val(Me.Fg.TextMatrix(Me.Fg.Row, Me.Fg.ColIndex("id")))

    FrmRequerMainten.Retrive (i)
ErrTrap:
End Sub

Private Sub Form_Activate()
'   PutFormOnTop Me.hWnd
End Sub

Private Sub Form_Load()
    Dim GrdBack As ClsBackGroundPic
    Dim Dcombos As ClsDataCombos

    Set Dcombos = New ClsDataCombos
  Dcombos.GetEmpDepartments Me.DcbUnit
    Dcombos.GetBranches Me.Dcbranch
    Dcombos.GetEquipments DcbEquepment
    
    '  If SystemOptions.UserInterface = EnglishInterface Then
     
    '  Me.DcbOrderStatus.AddItem "New"
    '    Me.DcbOrderStatus.AddItem "Accept Customer"
    '    Me.DcbOrderStatus.AddItem "Final Maintenance"

       If SystemOptions.UserInterface = EnglishInterface Then
            ProblemTimID.AddItem "During Production"
            ProblemTimID.AddItem "During Start up"
            ProblemTimID.AddItem "During Repair"
            ProblemTimID.AddItem "Others"
         '   ProblemTimID1.AddItem "During Production"
         '   ProblemTimID1.AddItem "During Start up"
         '   ProblemTimID1.AddItem "During Repair"
         '   ProblemTimID1.AddItem "Others"
     '   SetInterface Me
     '   ChangeLang
        Else
        ProblemTimID.AddItem "«À‰«¡ «· ’‰Ì⁄"
        ProblemTimID.AddItem "«À‰«¡ »œ¡ «· ‘€Ì·"
        ProblemTimID.AddItem "«À‰«¡ «·«’·«Õ"
        ProblemTimID.AddItem "«Œ—Ï"
      '      ProblemTimID1.AddItem "«À‰«¡ «· ’‰Ì⁄"
      ''  ProblemTimID1.AddItem "«À‰«¡ »œ¡ «· ‘€Ì·"
      '  ProblemTimID1.AddItem "«À‰«¡ «·«’·«Õ"
      '  ProblemTimID1.AddItem "«Œ—Ï"

    End If
    Set DCboSearch = New clsDCboSearch
   ' Set DCboSearch.Client = Me.DCEmp_Name
    'Dcombos.GetUsers Me.DCUser
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Clear").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture

  '  CenterForm Me
'GetData
'    FormPostion Me, GetPostion
    Set GrdBack = New ClsBackGroundPic

    With Me.Fg
        Set .WallPaper = GrdBack.Picture
        .AutoSize 0, .Cols - 1, False
    End With
 If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
'    SetDtpickerDate Me.DtpDateFrom
'    SetDtpickerDate Me.DtpDateTo

End Sub

Public Sub GetData()
    Dim StrSQL As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer


StrSQL = "  SELECT dbo.TblRequerMainten.ID, dbo.TblRequerMainten.ProblemTimID, dbo.TblRequerMainten.ProblemOther, dbo.TblRequerMainten.StopTime, dbo.TblRequerMainten.StartTime,"
StrSQL = StrSQL & "                  dbo.TblRequerMainten.Des, dbo.TblRequerMainten.Remarks, dbo.TblRequerMainten.RecordDate, dbo.TblRequerMainten.StartDate, dbo.TblRequerMainten.StopDate,"
StrSQL = StrSQL & "                      dbo.TblRequerMainten.UnitID, dbo.TblEmpDepartments.DepartmentName, dbo.TblEmpDepartments.DepartmentNamee, dbo.TblRequerMainten.EquepID,"
StrSQL = StrSQL & "                      dbo.FixedAssets.Name, dbo.FixedAssets.code, dbo.TblRequerMainten.BranchID, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee,"
 StrSQL = StrSQL & "                     dbo.FixedAssets.NameE"
StrSQL = StrSQL & "    FROM     dbo.TblRequerMainten LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.TblBranchesData ON dbo.TblRequerMainten.BranchID = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
   StrSQL = StrSQL & "                   dbo.FixedAssets ON dbo.TblRequerMainten.EquepID = dbo.FixedAssets.id LEFT OUTER JOIN"
   StrSQL = StrSQL & "                   dbo.TblEmpDepartments ON dbo.TblRequerMainten.UnitID = dbo.TblEmpDepartments.DeparmentID CROSS JOIN"
  StrSQL = StrSQL & "                    dbo.TblOrderMaint"
StrSQL = StrSQL & "    Where (Not (dbo.TblRequerMainten.id = dbo.TblOrderMaint.ReqMainID))"


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
            Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            Me.lbl(10).Caption = "Search Results=0"
        End If

        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «·»ÕÀ"
        Cmd_Click (1)
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    Else

        With Me.Fg
            .Clear flexClearScrollable, flexClearEverything
            .Rows = .FixedRows
            .Rows = rs.RecordCount + .FixedRows

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
                Me.lbl(10).Caption = "Search Results=" & rs.RecordCount
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
                Else
            .TextMatrix(i, .ColIndex("DepartmentName")) = IIf(IsNull(rs("DepartmentName").value), "", rs("DepartmentName").value)
                .TextMatrix(i, .ColIndex("branch_name")) = IIf(IsNull(rs("branch_name").value), "", rs("branch_name").value)
                .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs("Name").value), "", rs("Name").value)
                .TextMatrix(i, .ColIndex("des")) = IIf(IsNull(rs("des").value), "", rs("des").value)
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

Private Sub ChangeLang()
 
    Cmd(1).Caption = "Delete"
    Cmd(0).Caption = "Search"
    Cmd(2).Caption = "Exit"
    
  Me.Caption = "Search Requer Maintenance"
lbprocess.Caption = "OrderNo"
lbl(5).Caption = "From"
lbl(6).Caption = "To"
lbl(11).Caption = "From"
lbl(0).Caption = "To"
Fra(1).Caption = "Date"
lblbr.Caption = "Branch"
lbl(4).Caption = "Machine"
lbl(3).Caption = "Dept"
lbl(7).Caption = "Time Problem"
Cmd(3).Caption = "Print"
lbl(8).Caption = "Refresh"
lbl(2).Caption = " / Minute "
'Me.lbreg.Caption = "Date Registration"

     With Me.Fg
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
End Sub

 
Private Sub Timer1_Timer()

End Sub

Private Sub Form_Unload(Cancel As Integer)

End Sub
