VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmSearchVocationEntitlement 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "»ÕÀ „” Õﬁ«  «·ﬁÌ«„  »«·«Ã«“…"
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8820
   Icon            =   "FrmSearchVocationEntitlement.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   4890
   ScaleWidth      =   8820
   ShowInTaskbar   =   0   'False
   Begin VB.Frame lblLW 
      BackColor       =   &H00E2E9E9&
      Caption         =   "»ÕÀ »Õ”»"
      Height          =   1455
      Left            =   2520
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   3240
      Width           =   6345
      Begin VB.TextBox TxtSearchCode 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3480
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   240
         Width           =   1335
      End
      Begin MSDataListLib.DataCombo DcboEmpName 
         Height          =   315
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   3315
         _ExtentX        =   5847
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcbDept 
         Height          =   315
         Left            =   120
         TabIndex        =   21
         Top             =   600
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo Opretot 
         Bindings        =   "FrmSearchVocationEntitlement.frx":038A
         Height          =   315
         Left            =   120
         TabIndex        =   23
         Top             =   960
         Width           =   4695
         _ExtentX        =   8281
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
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«·ﬁ«∆„"
         Height          =   255
         Index           =   8
         Left            =   4920
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·«œ«—Â/«·ﬁ”„"
         Height          =   285
         Index           =   7
         Left            =   4950
         TabIndex        =   22
         Top             =   600
         Width           =   1245
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«”„ «·„ÊŸ›"
         Height          =   285
         Index           =   0
         Left            =   5070
         TabIndex        =   20
         Top             =   255
         Width           =   1125
      End
   End
   Begin VB.Frame lbreg 
      BackColor       =   &H00E2E9E9&
      Caption         =   " «—ÌŒ «·Õ—ﬂÂ"
      Height          =   1035
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   3360
      Width           =   2295
      Begin MSComCtl2.DTPicker DtpDateFrom 
         Height          =   330
         Left            =   120
         TabIndex        =   6
         Top             =   270
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   113311747
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
         Format          =   113311747
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
      Caption         =   "—ﬁ„ «·Õ—ﬂÂ"
      Height          =   645
      Left            =   5040
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   2580
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
      Width           =   8835
      _cx             =   15584
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
      FormatString    =   $"FrmSearchVocationEntitlement.frx":039F
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
      Width           =   2775
   End
End
Attribute VB_Name = "FrmSearchVocationEntitlement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim DCboSearch As clsDCboSearch
Public index As Integer

Private Sub Cmd_Click(index As Integer)

    Select Case index

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



Private Sub DcboEmpName_Change()
DcboEmpName_Click (0)
End Sub
Private Sub DcboEmpName_Click(Area As Integer)
'    On Error Resume Next
       If val(DcboEmpName.BoundText) = 0 Then Exit Sub


    Dim EmpCode  As String
 
    GetEmployeeIDFromCode , , DcboEmpName.BoundText, EmpCode
    TxtSearchCode.text = EmpCode
    End Sub

Private Sub fg_Click()
If index = 0 Then
FrmVocationEntitlements.Retrive (val(Me.Fg.TextMatrix(Me.Fg.row, Me.Fg.ColIndex("id"))))
ElseIf index = 1 Then
FrmTypeExchange.txtTransaction_ID = (val(Me.Fg.TextMatrix(Me.Fg.row, Me.Fg.ColIndex("id"))))
FrmTypeExchange.TxtOrderNo = (val(Me.Fg.TextMatrix(Me.Fg.row, Me.Fg.ColIndex("id"))))



ElseIf index = 2 Then
FrmPayments.TxtDue = (val(Me.Fg.TextMatrix(Me.Fg.row, Me.Fg.ColIndex("id"))))

End If



End Sub





Private Sub Form_Activate()
'   PutFormOnTop Me.hWnd
End Sub

Private Sub Form_Load()
    Dim GrdBack As ClsBackGroundPic
    Dim Dcombos As ClsDataCombos

    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Clear").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture

  Set Dcombos = New ClsDataCombos
    Dcombos.GetUsers Opretot
    
    Dcombos.GetEmployees Me.DcboEmpName
   
    Dcombos.GetEmpDepartments Me.DcbDept
    
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

StrSQL = " SELECT     dbo.TblVocationEntitlements.RecordDate, dbo.TblVocationEntitlements.DateSta, dbo.TblVocationEntitlements.OpretotID, dbo.TblUsers.UserName, "
StrSQL = StrSQL & "                      dbo.TblVocationEntitlements.EmpID, dbo.TblVocationEntitlements.BranchID, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee,"
StrSQL = StrSQL & "                      dbo.TblVocationEntitlements.JobID, dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmpJobsTypes.JobTypeNamee, dbo.TblVocationEntitlements.DeptID,"
StrSQL = StrSQL & "                      dbo.TblEmpDepartments.DepartmentName, dbo.TblEmpDepartments.DepartmentNamee, dbo.TblVocationEntitlements.BignDate,"
StrSQL = StrSQL & "                      dbo.TblVocationEntitlements.LastVocatinDate, dbo.TblVocationEntitlements.ContDay, dbo.TblVocationEntitlements.LastDayVoc, dbo.TblVocationEntitlements.TotalDay,"
StrSQL = StrSQL & "                      dbo.TblVocationEntitlements.NoDay, dbo.TblVocationEntitlements.NoMonth, dbo.TblVocationEntitlements.NoYear, dbo.TblVocationEntitlements.Remark,"
StrSQL = StrSQL & "                      dbo.TblVocationEntitlements.DaySalary, dbo.TblVocationEntitlements.Salary, dbo.TblVocationEntitlements.DayIncrease, dbo.TblVocationEntitlements.Increase,"
StrSQL = StrSQL & "                      dbo.TblVocationEntitlements.DaySalVocation, dbo.TblVocationEntitlements.SalaryVocation, dbo.TblVocationEntitlements.DayEntitOther,"
StrSQL = StrSQL & "                      dbo.TblVocationEntitlements.SalEntitOther, dbo.TblVocationEntitlements.Other, dbo.TblVocationEntitlements.Advance, dbo.TblVocationEntitlements.ValueTickt,"
StrSQL = StrSQL & "                      dbo.TblVocationEntitlements.Booked, dbo.TblVocationEntitlements.Delivery, dbo.TblVocationEntitlements.ID, dbo.TblEmployee.Emp_Name,"
StrSQL = StrSQL & "                      dbo.TblEmployee.fullcode , dbo.TblEmployee.Emp_Namee"
StrSQL = StrSQL & " FROM         dbo.TblEmpJobsTypes RIGHT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblEmpDepartments RIGHT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblVocationEntitlements LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblEmployee ON dbo.TblVocationEntitlements.EmpID = dbo.TblEmployee.Emp_ID ON"
StrSQL = StrSQL & "                      dbo.TblEmpDepartments.DeparmentID = dbo.TblVocationEntitlements.DeptID ON"
StrSQL = StrSQL & "                      dbo.TblEmpJobsTypes.JobTypeID = dbo.TblVocationEntitlements.JobID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblBranchesData ON dbo.TblVocationEntitlements.BranchID = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblUsers ON dbo.TblVocationEntitlements.OpretotID = dbo.TblUsers.UserID"

    BolBegine = False
    StrWhere = ""
 StrWhere = StrWhere & "  where not ( dbo.TblVocationEntitlements.RecordDate is null)"
 
 If index <> 0 Then
StrWhere = StrWhere & "and   not (NoteSerial is null) "
End If

 If CheckAprroveScreen("FrmVocationEntitlements") = True And index <> 0 Then
     StrWhere = StrWhere & " and approved =1"

End If


    If val(Me.TxtIDFrom.text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " dbo.TblVocationEntitlements.ID >=" & val(Me.TxtIDFrom.text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblVocationEntitlements.ID >=" & val(Me.TxtIDFrom.text) & ""
        End If
    End If
  

    If val(Me.TxtIDTO.text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblVocationEntitlements.ID <=" & val(Me.TxtIDTO.text) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblVocationEntitlements.ID <=" & val(Me.TxtIDTO.text) & ""
        End If
    End If
    '///////////////////
     If Me.TxtSearchCode.text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblEmployee.fullcode ='" & Me.TxtSearchCode.text & "'"
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblEmployee.fullcode ='" & Me.TxtSearchCode.text & "'"
        End If
    End If
    If Me.DcboEmpName.text <> "" And (val(DcboEmpName.BoundText) <> 0) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblVocationEntitlements.EmpID =" & Me.DcboEmpName.BoundText & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblVocationEntitlements.EmpID =" & Me.DcboEmpName.BoundText & ""
        End If
    End If
    
        If Me.DcbDept.text <> "" And (val(DcbDept.BoundText) <> 0) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblVocationEntitlements.DeptID =" & Me.DcbDept.BoundText & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblVocationEntitlements.DeptID =" & Me.DcbDept.BoundText & ""
        End If
    End If
         If Me.Opretot.text <> "" And (val(Opretot.BoundText) <> 0) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblVocationEntitlements.OpretotID =" & Me.Opretot.BoundText & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblVocationEntitlements.OpretotID =" & Me.Opretot.BoundText & ""
        End If
    End If

    If Not IsNull(Me.DtpDateFrom.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblVocationEntitlements.RecordDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblVocationEntitlements.RecordDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        End If
    End If

    If Not IsNull(Me.DtpDateTo.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND  dbo.TblVocationEntitlements.RecordDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TblVocationEntitlements.RecordDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        End If
    End If

    '-----------------------------------

    StrSQL = StrSQL & StrWhere
    StrSQL = StrSQL & " Order By dbo.TblVocationEntitlements.ID"
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
               
                
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs("ID").value), "", rs("ID").value)
                        
                If Not (IsNull(rs("RecordDate").value)) Then
                    .TextMatrix(i, .ColIndex("RecordDate")) = Format(rs("RecordDate").value, "yyyy/M/d")
                End If
            If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(i, .ColIndex("empname")) = IIf(IsNull(rs("Emp_Name").value), "", rs("Emp_Name").value)
            .TextMatrix(i, .ColIndex("DepartmentName")) = IIf(IsNull(rs("DepartmentName").value), "", rs("DepartmentName").value)
          
            Else
            .TextMatrix(i, .ColIndex("empname")) = IIf(IsNull(rs("Emp_Name").value), "", rs("Emp_Name").value)
            .TextMatrix(i, .ColIndex("DepartmentName")) = IIf(IsNull(rs("DepartmentNamee").value), "", rs("DepartmentNamee").value)
            End If
           .TextMatrix(i, .ColIndex("UserName")) = IIf(IsNull(rs("UserName").value), "", rs("UserName").value)
            
                
                
        
            
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
  Me.Caption = "Saerch Due to vacation"
lbprocess.Caption = "No Transection"
lbl(5).Caption = "From"
lbl(6).Caption = "To"
lbl(4).Caption = "From"
lbl(3).Caption = "To"
lbreg.Caption = "Date Transection"
lblLW.Caption = "Saerch By"
lbl(2).Caption = "Total"
lbl(0).Caption = "Employee"
lbl(7).Caption = "Department"
lbl(8).Caption = "Charge d'affaires"
     With Me.Fg
        .TextMatrix(0, .ColIndex("Serial")) = "Serial"
        .TextMatrix(0, .ColIndex("id")) = "No Transection"
        .TextMatrix(0, .ColIndex("RecordDate")) = "Date"
         .TextMatrix(0, .ColIndex("DepartmentName")) = "DepartmentName"
        .TextMatrix(0, .ColIndex("empname")) = "Emp Name"
       .TextMatrix(0, .ColIndex("UserName")) = "Charge d'affaires"
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

Private Sub TxtSearchCode_KeyPress(KeyAscii As Integer)
   Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode TxtSearchCode.text, EmpID
        DcboEmpName.BoundText = EmpID
    End If
    
End Sub
