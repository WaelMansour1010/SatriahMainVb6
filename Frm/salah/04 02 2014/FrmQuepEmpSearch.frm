VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SUITEC~1.OCX"
Begin VB.Form FrmQuepEmpSearch 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "«·»ÕÀ ⁄‰ «” »Ì«‰ „ÊŸ›"
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6750
   Icon            =   "FrmQuepEmpSearch.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   4890
   ScaleWidth      =   6750
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TxtEmpCode 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   4800
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   3360
      Width           =   915
   End
   Begin VB.TextBox TxtEmpName 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   2880
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   3360
      Width           =   1875
   End
   Begin VB.ComboBox DcbOrderStatus 
      Height          =   315
      Left            =   6840
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   1080
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   6960
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Text            =   " „ „Ê«›ﬁ… «·⁄„"
      Top             =   600
      Width           =   1035
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   "‰Ê⁄ «·ÿ·»"
      Height          =   645
      Index           =   0
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   3840
      Width           =   6675
      Begin VB.CheckBox RegCK 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«÷ÿ—«—ÌÂ"
         Height          =   255
         Index           =   1
         Left            =   360
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   240
         Width           =   1695
      End
      Begin VB.CheckBox RegCK 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—”„Ì…"
         Height          =   255
         Index           =   0
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   240
         Width           =   1335
      End
      Begin XtremeSuiteControls.RadioButton RadioButton1 
         Height          =   255
         Left            =   5400
         TabIndex        =   23
         Top             =   240
         Width           =   1215
         _Version        =   786432
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "«” ⁄·«„"
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton RadioButton2 
         Height          =   255
         Left            =   3960
         TabIndex        =   24
         Top             =   240
         Width           =   1215
         _Version        =   786432
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "«Ã«“…"
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
   End
   Begin VB.Frame lbreg 
      BackColor       =   &H00E2E9E9&
      Caption         =   " «—ÌŒ «·Õ—ﬂÂ"
      Height          =   1035
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   2880
      Width           =   2295
      Begin MSComCtl2.DTPicker DtpDateFrom 
         Height          =   330
         Left            =   90
         TabIndex        =   6
         Top             =   270
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   93913091
         CurrentDate     =   38887
      End
      Begin MSComCtl2.DTPicker DtpDateTo 
         Height          =   330
         Left            =   90
         TabIndex        =   7
         Top             =   630
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   93913091
         CurrentDate     =   38887
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï"
         Height          =   195
         Index           =   3
         Left            =   1695
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   660
         Width           =   480
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   195
         Index           =   4
         Left            =   1680
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   330
         Width           =   540
      End
   End
   Begin VB.Frame lbprocess 
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·⁄„·ÌÂ"
      Height          =   645
      Left            =   2880
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   2700
      Width           =   3795
      Begin VB.TextBox TxtIDFrom 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   1920
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
         Left            =   2775
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
         Left            =   1260
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   240
         Width           =   525
      End
   End
   Begin VSFlex8UCtl.VSFlexGrid Fg 
      Height          =   2625
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   6675
      _cx             =   11774
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
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmQuepEmpSearch.frx":038A
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
      Top             =   4440
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
      Top             =   4440
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
      Top             =   4440
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
      Left            =   1920
      RightToLeft     =   -1  'True
      TabIndex        =   17
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
      TabIndex        =   16
      Top             =   3060
      Width           =   1785
   End
   Begin VB.Label LblClientName 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·„ÊŸ›"
      Height          =   195
      Left            =   5760
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   3390
      Width           =   795
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
      Width           =   2775
   End
End
Attribute VB_Name = "FrmQuepEmpSearch"
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
    End Select

End Sub


Private Sub Fg_Click()
FrmQUesEmp.retrive1 (val(Me.Fg.TextMatrix(Me.Fg.Row, Me.Fg.ColIndex("id"))))


End Sub





Private Sub Form_Activate()
'   PutFormOnTop Me.hWnd
End Sub

Private Sub Form_Load()
    Dim GrdBack As ClsBackGroundPic
 
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
    SetDtpickerDate Me.DtpDateFrom
    SetDtpickerDate Me.DtpDateTo

End Sub

Private Sub Form_Unload(Cancel As Integer)

    FormPostion Me, SavePostion
    Set DCboSearch = Nothing
End Sub

Public Sub GetData()
    Dim StrSQL As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer

StrSQL = " SELECT     dbo.TblQuesEmp.ID, dbo.TblQuesEmp.RecordDate, dbo.TblQuesEmp.EmpID, dbo.TblQuesEmp.ProjectID, dbo.TblQuesEmp.BranchID, dbo.EmpGroupDep.GroupName, "
StrSQL = StrSQL & "                      dbo.TblBranchesData.branch_namee, dbo.TblBranchesData.branch_name, dbo.TblQuesEmp.JobID, dbo.TblEmpJobsTypes.JobTypeName,"
StrSQL = StrSQL & "                      dbo.TblEmpJobsTypes.JobTypeNamee, dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Name1,"
StrSQL = StrSQL & "                      dbo.TblEmployee.Emp_Name2, dbo.TblEmployee.Emp_Name3, dbo.TblEmployee.Emp_Name4, dbo.TblEmployee.Emp_Namee, dbo.TblEmployee.Emp_Namee1,"
StrSQL = StrSQL & "                      dbo.TblEmployee.Emp_Namee2, dbo.TblEmployee.Emp_Namee3, dbo.TblEmployee.Emp_Namee4, dbo.TblQuesEmp.SalaryAppoint, dbo.TblQuesEmp.CurrSalary,"
StrSQL = StrSQL & "                      dbo.TblQuesEmp.FileNo, dbo.TblQuesEmp.NationalID, dbo.TblQuesEmp.Posted, dbo.TblQuesEmp.NutWork, dbo.TblQuesEmp.EndIqama, dbo.TblQuesEmp.StartWork,"
StrSQL = StrSQL & "                      dbo.TblQuesEmp.LastTreatment, dbo.TblQuesEmp.WorkFrom, dbo.TblQuesEmp.WorkTo, dbo.TblQuesEmp.[Day], dbo.TblQuesEmp.[Month], dbo.TblQuesEmp.[Year],"
StrSQL = StrSQL & "                      dbo.TblQuesEmp.Increase, dbo.TblQuesEmp.Additional, dbo.TblQuesEmp.Absent, dbo.TblQuesEmp.other, dbo.TblQuesEmp.LongCont, dbo.TblQuesEmp.PostedDate,"
StrSQL = StrSQL & "                      dbo.TblQuesEmp.SpecificHolidyaType1, dbo.TblQuesEmp.SpecificHolidyaType2, dbo.TblQuesEmp.EndIqamah, dbo.TblQuesEmp.Remarks,"
StrSQL = StrSQL & "                      dbo.TblQuesEmp.HolidayType , dbo.TblEmployee.Nationality, dbo.TblEmployee.fullcode"
StrSQL = StrSQL & " FROM         dbo.TblEmployee RIGHT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblQuesEmp ON dbo.TblEmployee.Emp_ID = dbo.TblQuesEmp.EmpID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblEmpJobsTypes ON dbo.TblQuesEmp.JobID = dbo.TblEmpJobsTypes.JobTypeID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblBranchesData ON dbo.TblQuesEmp.BranchID = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.EmpGroupDep ON dbo.TblQuesEmp.ProjectID = dbo.EmpGroupDep.GroupID"


    BolBegine = False
    StrWhere = ""

    If val(Me.TxtIDFrom.text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " dbo.TblQuesEmp.ID >=" & val(Me.TxtIDFrom.text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblQuesEmp.ID >=" & val(Me.TxtIDFrom.text) & ""
        End If
    End If
  

    If val(Me.TxtIDTO.text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblQuesEmp.ID <=" & val(Me.TxtIDTO.text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblQuesEmp.ID <=" & val(Me.TxtIDTO.text) & ""
        End If
    End If
    '///////////////////
     If Me.TxtEmpCode.text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblEmployee.fullcode ='" & Me.TxtEmpCode.text & "'"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblEmployee.fullcode ='" & Me.TxtEmpCode.text & "'"
        End If
    End If
'////////////////////////
 If Me.TxtEmpName.text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblEmployee.Emp_Name like '%" & Me.TxtEmpName.text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblEmployee.Emp_Name like '%" & Me.TxtEmpName.text & "%'"
        End If
    End If


    If Me.RadioButton1.value = True Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblQuesEmp.HolidayType = 0"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblQuesEmp.HolidayType = 0"
        End If
    End If
  If Me.RadioButton2.value = True Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblQuesEmp.HolidayType = 1"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblQuesEmp.HolidayType = 1 "
        End If
        End If
       If Me.RegCK(0).value = vbChecked Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblQuesEmp.SpecificHolidyaType1 = 1"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblQuesEmp.SpecificHolidyaType1 = 1 "
        End If
        End If
          If Me.RegCK(1).value = vbChecked Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblQuesEmp.SpecificHolidyaType2 = 1"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblQuesEmp.SpecificHolidyaType2 = 1 "
        End If
    End If
    If Not IsNull(Me.DtpDateFrom.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblQuesEmp.RecordDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblQuesEmp.RecordDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        End If
    End If

    If Not IsNull(Me.DtpDateTo.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND  dbo.TblQuesEmp.RecordDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TblQuesEmp.RecordDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        End If
    End If

    '-----------------------------------

    StrSQL = StrSQL & StrWhere
    StrSQL = StrSQL & " Order By dbo.TblQuesEmp.ID"
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
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
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
               
                
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs("ID").value), "", rs("ID").value)
                        
                If Not (IsNull(rs("RecordDate").value)) Then
                    .TextMatrix(i, .ColIndex("RecordDate")) = Format(rs("RecordDate").value, "yyyy/M/d")
                End If
            
                .TextMatrix(i, .ColIndex("empname")) = IIf(IsNull(rs("Emp_Name").value), "", rs("Emp_Name").value)
                .TextMatrix(i, .ColIndex("code")) = IIf(IsNull(rs("fullcode").value), "", rs("fullcode").value)
                If (rs("HolidayType").value) = False Then
                .TextMatrix(i, .ColIndex("type")) = "«” ⁄·«„"
                Else
                .TextMatrix(i, .ColIndex("type")) = "≈Ã«“…"
                End If
              '  .TextMatrix(i, .ColIndex("UserName")) = IIf(IsNull(rs("UserName").value), "", rs("UserName").value)
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
  Me.Caption = "Search Query questionnaire for employees"
  Me.RadioButton1.RightToLeft = True
  Me.RadioButton1.Caption = "Query"
Me.RadioButton2.RightToLeft = True
Me.RadioButton2.Caption = "Vacation"
Me.RegCK(0).Caption = "Official"
Me.RegCK(1).Caption = "COMPELLING"
Me.LblClientName.Caption = "EmpName"
lbl(4).Caption = "From"
lbl(3).Caption = "To"
lbl(5).Caption = "From"
lbl(6).Caption = "To"
Fra(0).Caption = "Type Order"
lbl(2).Caption = "Total"
Me.lbreg.Caption = "Date Registration"
Me.lbprocess.Caption = "Process No"
     With Me.Fg
        .TextMatrix(0, .ColIndex("Serial")) = "NO"
        .TextMatrix(0, .ColIndex("id")) = "No Process"
        .TextMatrix(0, .ColIndex("RecordDate")) = "Date"
         .TextMatrix(0, .ColIndex("code")) = "Code"
        .TextMatrix(0, .ColIndex("empname")) = "Emp Name"
       .TextMatrix(0, .ColIndex("type")) = "Type Order"
    End With
  '
End Sub

Private Sub TxtIDFrom_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtIDFrom.text, 1)
'    FrmCarAuthontication.TxtOrder.text = ""
End Sub

Private Sub TxtIDTO_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtIDTO.text, 1)
'    FrmCarAuthontication.TxtOrder.text = ""
End Sub

