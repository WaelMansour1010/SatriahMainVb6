VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SuiteCtrls.ocx"
Begin VB.Form FrmInsurance 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " ﬁ«—Ì—  «· √„Ì‰«  ··„ÊŸ›Ì‰"
   ClientHeight    =   5010
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10260
   Icon            =   "frminsurance.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5010
   ScaleWidth      =   10260
   ShowInTaskbar   =   0   'False
   Begin XtremeSuiteControls.RadioButton XPChkSearchTypeClient1 
      Height          =   495
      Left            =   8370
      TabIndex        =   14
      Top             =   5010
      Width           =   1815
      _Version        =   786432
      _ExtentX        =   3201
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "›—“ »Õ”» «·„ÊŸ›"
      UseVisualStyle  =   -1  'True
      TextAlignment   =   1
      RightToLeft     =   -1  'True
   End
   Begin VB.ComboBox DcbOrderStatus 
      Height          =   315
      Left            =   11640
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   1200
      Width           =   1935
   End
   Begin VB.ComboBox ComGranty 
      Height          =   315
      Left            =   11040
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   120
      Width           =   2175
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Height          =   2985
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   480
      Width           =   10155
      Begin VB.CheckBox chkIsPayInsu 
         Alignment       =   1  'Right Justify
         Caption         =   "«·„œ›Ê⁄«  ·· √„Ì‰«  «Ã„«·Ì"
         Height          =   285
         Index           =   1
         Left            =   60
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   2070
         Width           =   2235
      End
      Begin VB.CheckBox chkIsPayInsu 
         Alignment       =   1  'Right Justify
         Caption         =   "«·„œ›Ê⁄«  ·· √„Ì‰« "
         Height          =   285
         Index           =   0
         Left            =   2940
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   2070
         Width           =   1755
      End
      Begin VB.CheckBox chkIsNotInsurance 
         Alignment       =   1  'Right Justify
         Caption         =   "«·€Ì— Œ«÷⁄Ì‰ ·· √„Ì‰"
         Height          =   435
         Left            =   330
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   2370
         Width           =   1935
      End
      Begin VB.CheckBox chkIsInsurance 
         Alignment       =   1  'Right Justify
         Caption         =   "«·Œ«÷⁄Ì‰ ·· √„Ì‰"
         Height          =   435
         Left            =   3210
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   2370
         Width           =   1485
      End
      Begin VB.TextBox TxtRemarks 
         Alignment       =   1  'Right Justify
         Height          =   765
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   3600
         Width           =   3915
      End
      Begin VB.TextBox TxtSearchCode 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   7080
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   255
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker DtpDateFrom 
         Height          =   330
         Left            =   7800
         TabIndex        =   8
         Top             =   2280
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   207290371
         CurrentDate     =   38887
      End
      Begin MSComCtl2.DTPicker DtpDateTo 
         Height          =   330
         Left            =   5280
         TabIndex        =   10
         Top             =   2280
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   207290371
         CurrentDate     =   38887
      End
      Begin MSDataListLib.DataCombo DcboEmpName 
         Height          =   315
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   6915
         _ExtentX        =   12197
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcmbToDepart 
         Height          =   315
         Left            =   4770
         TabIndex        =   22
         Top             =   1920
         Visible         =   0   'False
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo dcmbToProject 
         Height          =   315
         Left            =   120
         TabIndex        =   23
         Top             =   2955
         Width           =   3915
         _ExtentX        =   6906
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcmbToJob 
         Height          =   315
         Left            =   120
         TabIndex        =   24
         Top             =   1260
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo dcstatus 
         Height          =   315
         Left            =   5160
         TabIndex        =   29
         Top             =   840
         Width           =   3795
         _ExtentX        =   6694
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo dcnationality 
         Height          =   315
         Left            =   120
         TabIndex        =   31
         Top             =   840
         Width           =   3795
         _ExtentX        =   6694
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo dcBranch 
         Height          =   315
         Left            =   5160
         TabIndex        =   32
         Top             =   1320
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·›—⁄"
         Height          =   195
         Index           =   8
         Left            =   8895
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   1320
         Width           =   1140
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "Õ«·Â «·⁄„·"
         Height          =   195
         Index           =   6
         Left            =   8880
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   840
         Width           =   1140
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Height          =   195
         Left            =   4770
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   360
         Width           =   45
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·ÊŸÌ›…«·Õ«·Ì…"
         Height          =   195
         Index           =   5
         Left            =   3840
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   1320
         Width           =   1140
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·„·«ÕŸ…"
         Height          =   195
         Index           =   2
         Left            =   3960
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   3840
         Width           =   1065
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·Ã‰”Ì…"
         Height          =   195
         Index           =   0
         Left            =   3990
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   810
         Width           =   1065
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï  «—ÌŒ"
         Height          =   195
         Index           =   3
         Left            =   6930
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   2310
         Width           =   705
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰  «—ÌŒ"
         Height          =   195
         Index           =   4
         Left            =   9270
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   2370
         Width           =   765
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·ﬁ”„ «·Õ«·Ì"
         Height          =   195
         Index           =   7
         Left            =   9180
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   2940
         Width           =   1140
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·„ÊŸ›"
         Height          =   195
         Left            =   9300
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   11520
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Text            =   " „ „Ê«›ﬁ… «·⁄„"
      Top             =   720
      Width           =   915
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   495
      Index           =   0
      Left            =   2850
      TabIndex        =   0
      Top             =   4530
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   873
      ButtonPositionImage=   1
      Caption         =   "⁄—÷ «· ﬁ—Ì—"
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
      Height          =   495
      Index           =   1
      Left            =   1380
      TabIndex        =   1
      Top             =   4530
      Width           =   1365
      _ExtentX        =   2408
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
      Height          =   495
      Index           =   2
      Left            =   0
      TabIndex        =   2
      Top             =   4530
      Width           =   1365
      _ExtentX        =   2408
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
   Begin XtremeSuiteControls.RadioButton XPChkSearchTypeCar 
      Height          =   375
      Left            =   5850
      TabIndex        =   15
      Top             =   5010
      Width           =   1815
      _Version        =   786432
      _ExtentX        =   3201
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "›—“ »Õ”» «·ﬁ”„"
      UseVisualStyle  =   -1  'True
      TextAlignment   =   1
      RightToLeft     =   -1  'True
   End
   Begin XtremeSuiteControls.RadioButton XPChkSearchTypeModel 
      Height          =   375
      Left            =   3450
      TabIndex        =   16
      Top             =   5010
      Width           =   1935
      _Version        =   786432
      _ExtentX        =   3413
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "›—“ »Õ”» «·„Êﬁ⁄"
      UseVisualStyle  =   -1  'True
      TextAlignment   =   1
      RightToLeft     =   -1  'True
   End
   Begin XtremeSuiteControls.RadioButton XPChkSearchTypePlate 
      Height          =   375
      Left            =   930
      TabIndex        =   17
      Top             =   5010
      Width           =   1935
      _Version        =   786432
      _ExtentX        =   3413
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "›—“ »Õ”» «·ÊŸÌ›…"
      UseVisualStyle  =   -1  'True
      TextAlignment   =   1
      RightToLeft     =   -1  'True
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      BackStyle       =   0  'Transparent
      Caption         =   " ﬁ«—Ì—  «· √„Ì‰«  ··„ÊŸ›Ì‰"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6675
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   120
      Width           =   3390
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      ForeColor       =   &H00000080&
      Height          =   285
      Index           =   1
      Left            =   30
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   4350
      Width           =   1785
   End
End
Attribute VB_Name = "FrmInsurance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim DCboSearch As clsDCboSearch

Private Sub chkIsPayInsu_Click(index As Integer)
If index = 0 And chkIsPayInsu(index).value = vbChecked Then
    chkIsPayInsu(1).value = vbUnchecked
ElseIf index = 1 And chkIsPayInsu(index).value = vbChecked Then
    chkIsPayInsu(0).value = vbUnchecked
End If
End Sub

Private Sub Cmd_Click(index As Integer)

    Select Case index

        Case 0
       

 GetData
            
        Case 1
            clear_all Me
DtpDateFrom.value = ""
DtpDateTo.value = ""

            If SystemOptions.UserInterface = ArabicInterface Then
               ' Me.lbl(0).Caption = "‰ ÌÃ… «·»ÕÀ"
            Else
               ' Me.lbl(0).Caption = "Search Results"
            End If

        Case 2
            Unload Me
            Case 3
'print_report
    End Select

End Sub




Private Sub fg_Click()

 

End Sub
'Public Sub FiLLTXT()
'
'    On Error GoTo ErrTrap
'    Dim i As Integer
' '   Frm2.Enabled = False
'    FrmCarAuthontication.XPTxtID.text = IIf(IsNull(RsSavRec.Fields("ID").value), "", RsSavRec.Fields("ID").value)
'    FrmCarAuthontication.TxtCliientName = IIf(IsNull(RsSavRec.Fields("CarID").value), "", RsSavRec.Fields("CarID").value)
'    TxtVacNamee.text = IIf(IsNull(RsSavRec.Fields("model").value), "", RsSavRec.Fields("model").value)
'
'    LabCurrRec.Caption = RsSavRec.AbsolutePosition
'    LabCountRec.Caption = RsSavRec.RecordCount

'    With Grid
'
'        For i = 1 To .Rows - 1
'
'            If Trim(TxtVac_ID.text) = .TextMatrix(i, .ColIndex("id")) Then
'                TxtSerial.text = .TextMatrix(i, .ColIndex("Ser"))
'                .Row = i
'                Exit Sub
'            End If

'        Next
'
'    End With
'
'ErrTrap:
'
'End Sub


'Private Sub Fg_EnterCell()
'   On Error GoTo ErrTrap
  '  FindRec val(Me.Fg.TextMatrix(Me.Grid.Row, Me.Fg.ColIndex("id")))
 'If FrmBillCarMaintExtra.ch = True Then
 'FrmBillCarMaintExtra.Retrive1 (val(Me.Fg.TextMatrix(Me.Fg.Row, Me.Fg.ColIndex("id"))))
 'Else
 ' FrmCarAuthontication.Retrive2 (val(Me.Fg.TextMatrix(Me.Fg.Row, Me.Fg.ColIndex("id"))))
 ' FrmCarAuthontication.TxtAmoutAccept.text = 0
 '   FrmCarAuthontication.TxtFirstPrice.text = 0
 '   FrmCarAuthontication.TXtCarMeter.text = ""
 '   FrmCarAuthontication.DcbOrderStatus.ListIndex = 0
'FrmCarAuthontication.ComGranty.ListIndex = 2
'  End If
'ErrTrap:
'End Sub
Public Function FindRec(ByVal RecId As Long)
 
End Function
Private Sub Form_Activate()
   PutFormOnTop Me.hWnd
End Sub
Private Sub ChangeLang()

 

 
 Cmd(1).Caption = "Delete"
    Cmd(0).Caption = "View Report"
   Cmd(2).Caption = "Exit"
  Me.Caption = "Reports of Movement Of Employee "
Label5.Caption = Me.Caption
Label1.Caption = "Emp"


lbl(7).Caption = "Current Dept"
lbl(0).Caption = "Current Location"
lbl(2).Caption = "Remraks"
lbl(5).Caption = "Current Job"
XPChkSearchTypeCar.RightToLeft = False
Me.XPChkSearchTypeCar.Caption = "By Dept"
Me.XPChkSearchTypeClient1.RightToLeft = False
Me.XPChkSearchTypeClient1.Caption = "By Emp"
Me.XPChkSearchTypePlate.RightToLeft = False
Me.XPChkSearchTypePlate.Caption = "By Job"
Me.XPChkSearchTypeModel.RightToLeft = False
Me.XPChkSearchTypeModel.Caption = "By Location"

lbl(3).Caption = "To Date"
lbl(4).Caption = "From Date"
End Sub

Private Sub Form_Load()
    'Dim GrdBack As ClsBackGroundPic
    Dim Dcombos As ClsDataCombos
    
        Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500




    Set Dcombos = New ClsDataCombos
   ' Dcombos.GetEmployees Me.DCEmp_Name
  Set Dcombos = New ClsDataCombos
     'Dcombos.GetUsers Me.DCboUserName
     Dcombos.GetEmployees Me.DcboEmpName
    
     Dcombos.GetEmpDepartments Me.DcmbToDepart
    
   Dcombos.GetBranches Me.Dcbranch
   
   Dcombos.GetEmpJobsTypes Me.DcmbToJob
   
   Dcombos.GetEmpLocations Me.dcmbToProject ' locatione
    Set DCboSearch = New clsDCboSearch
  '  Set DCboSearch.Client = Me.DcbClientname
    'Dcombos.GetUsers Me.DCUser
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Clear").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
Dim My_SQL As String

    My_SQL = "  select id,name from jopstatus   "
    fill_combo dcstatus, My_SQL


 
    
  My_SQL = "  select id,name from Nationality   "
    fill_combo DCNationality, My_SQL
 
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
    Dim StrSQL As String, sql As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
    
If chkIsPayInsu(0).value = vbChecked Or chkIsPayInsu(1).value = vbChecked Then
    sql = "SELECT     dbo.TBLInsurancesJoin.IDINS AS IDINSJOIN,TBLInsurancesJoin.BignDateWork,TBLInsurancesJoin.WorkDays, dbo.TBLInsurancesJoin.EmpCode, dbo.TBLInsurancesJoin.EmpInsurances, dbo.TBLInsurancesJoin.InsValue, "
    sql = sql + "                     dbo.TBLInsurancesJoin.InsTotal, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee, dbo.TblEmployee.Emp_ID,"
    sql = sql + "                       dbo.TBLInsurancesJoin.BranchId, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblEmployee.Nationality,"
    sql = sql + "                     dbo.TblEmployee.NationalityE, dbo.TBLInsurancesJoin.payed, dbo.TBLInsurancesJoin.Citirent, dbo.TBLInsurancesJoin.InsTotal2,"
    sql = sql + "                       dbo.TBLInsurancesJoin.CompRate,TBLInsurances.Monthe +1 as Monthe,TBLInsurances.SubYear"
    sql = sql + "  FROM         dbo.TblBranchesData RIGHT OUTER JOIN"
    sql = sql + "                      dbo.TBLInsurancesJoin ON dbo.TblBranchesData.branch_id = dbo.TBLInsurancesJoin.BranchId LEFT OUTER JOIN"
    sql = sql + "                     dbo.TblEmployee ON dbo.TBLInsurancesJoin.EmpCode = dbo.TblEmployee.Emp_ID"
    sql = sql + "                     Inner join TBLInsurances On TBLInsurances.IDINS =TBLInsurancesJoin.IDINS "
    
    sql = sql & "  Where 1 = 1 "
      If Not IsNull(DtpDateFrom.value) Then
        sql = sql & " and TBLInsurances.DateM >=" & SQLDate(DtpDateFrom.value, True) & ""
    End If
    If Not IsNull(DtpDateTo.value) Then
        sql = sql & "and TBLInsurances.DateM <=" & SQLDate(DtpDateTo.value, True) & ""
    End If
    StrSQL = sql
Else
    StrSQL = "SELECT     dbo.jopstatus.name, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_Code, dbo.TblBranchesData.branch_namee, dbo.TblEmployee.*, "
    StrSQL = StrSQL & "                        dbo.TblEmpJobsTypes.JobTypeName , dbo.TblEmpJobsTypes.JobTypeNamee"
    StrSQL = StrSQL & "  FROM         dbo.TblEmployee INNER JOIN"
     StrSQL = StrSQL & "   dbo.TblEmpJobsTypes ON dbo.TblEmployee.JobTypeID = dbo.TblEmpJobsTypes.JobTypeID INNER JOIN"
     StrSQL = StrSQL & "   dbo.jopstatus ON dbo.TblEmployee.jopstatusid = dbo.jopstatus.id INNER JOIN"
    StrSQL = StrSQL & "   dbo.TblBranchesData ON dbo.TblEmployee.BranchId = dbo.TblBranchesData.branch_id"
                          
    StrSQL = StrSQL & "  Where 1 = 1 "
       If Not IsNull(DtpDateFrom.value) Then
        StrSQL = StrSQL & " and TblEmployee.InstanceDateM >=" & SQLDate(DtpDateFrom.value, True) & ""
    End If
    If Not IsNull(DtpDateTo.value) Then
        StrSQL = StrSQL & "and TblEmployee.InstanceDateM <=" & SQLDate(DtpDateTo.value, True) & ""
    End If
   
End If
 
    If chkIsInsurance.value = vbChecked And chkIsNotInsurance.value = vbUnchecked Then
        StrSQL = StrSQL & " and  IsNull(dbo.TblEmployee.InsuranceState,0) = 1"
    ElseIf chkIsInsurance.value = vbUnchecked And chkIsNotInsurance.value = vbChecked Then
        StrSQL = StrSQL & " and  IsNull(dbo.TblEmployee.InsuranceState,0) = 0"
    End If
    
    BolBegine = False
    StrWhere = ""



 
 If (Me.TxtSearchCode.text) <> "" Then
        
            StrWhere = StrWhere & " AND dbo.TblEmployee.Fullcode like '%" & Me.TxtSearchCode.text & "%'"
        
    End If
   If Me.DcboEmpName.BoundText <> "" Then
     
            StrWhere = StrWhere & " AND dbo.TblEmployee.Emp_ID=" & val(Me.DcboEmpName.BoundText)
      
    End If
    If Me.DcmbToDepart.BoundText <> "" Then
     
            StrWhere = StrWhere & " AND dbo.TblEmployee.DepartmentID=" & val(Me.DcmbToDepart.BoundText)
      
    End If
   If dcstatus.BoundText <> "" Then
     
            StrWhere = StrWhere + " and TblEmployee.jopstatusid =" & val(dcstatus.BoundText)
 
   
    End If

    If DCNationality.BoundText <> "" Then
  
            StrWhere = StrWhere + " and TblEmployee.nationality ='" & Trim(DCNationality.text) & "'"
 
   
    End If

   If Me.Dcbranch.BoundText <> "" Then
     
            StrWhere = StrWhere + " and TblEmployee.BranchId =" & val(Dcbranch.BoundText)
 
   
    End If
 

    If Me.DcmbToJob.BoundText <> "" Then
     
            StrWhere = StrWhere + " and TblEmployee.JobTypeID =" & val(DcmbToJob.BoundText)
 
   
    End If
    

    

    '-----------------------------------
StrSQL = StrSQL & StrWhere
 
   StrSQL = StrSQL & " Order By dbo.TblEmployee.Emp_ID"
  
  
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
           ' Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
          '  Me.lbl(10).Caption = "Search Results=0"
        End If

        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «· ﬁ—Ì—"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    Else
    ' Me.DTPicker1.value = Format(rs("DateAcutExite").value, "yyyy/M/d")
 rs.MoveFirst
' MsgBox rs("Granty").value
 print_report StrSQL
'print_report StrSQL
       ' With Me.Fg
       '     .Clear flexClearScrollable, flexClearEverything
       '     .Rows = .FixedRows
       '     .Rows = rs.RecordCount + .FixedRows
'
            If SystemOptions.UserInterface = ArabicInterface Then
             '   Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
               ' Me.lbl(10).Caption = "Search Results=" & rs.RecordCount
            End If

           
 

    End If

End Sub
Function print_report(Optional NoteSerial As String)
     
    Set rs = New ADODB.Recordset
    rs.Open NoteSerial, Cn, adOpenStatic, adLockReadOnly, adCmdText
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
    If chkIsPayInsu(0).value = vbUnchecked And chkIsPayInsu(1).value = vbUnchecked And chkIsInsurance.value = vbUnchecked And chkIsNotInsurance.value = vbUnchecked Then
        MsgBox "ÌÃ» «Œ Ì«— ‰Ê⁄ «· ﬁ—Ì—"
        Exit Function
    End If
     If chkIsInsurance.value = vbChecked And chkIsNotInsurance.value = vbUnchecked Then
        StrReportTitle = "«·„ÊŸ›Ì‰ " & chkIsInsurance.Caption
    ElseIf chkIsInsurance.value = vbUnchecked And chkIsNotInsurance.value = vbChecked Then
        StrReportTitle = "«·„ÊŸ›Ì‰ " & chkIsNotInsurance.Caption
    ElseIf chkIsInsurance.value = vbChecked And chkIsNotInsurance.value = vbChecked Then
    StrReportTitle = " ﬁ—Ì— «· √„Ì‰«  ·ﬂ· «·„ÊŸ›Ì‰ "
    End If
    
        If chkIsPayInsu(0) Then
        
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "employeeInsurnace2.rpt"
        ElseIf chkIsPayInsu(1) Then
        
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "employeeInsurnaceTotal.rpt"
        Else
            If SystemOptions.UserInterface = ArabicInterface Then
            
                StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "employeeInsurnace.rpt"
              
            Else
                  
                StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "employeeInsurnace.rpt"
               
               
            End If
        End If


    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open NoteSerial, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        'GetMsgs 138, vbExclamation
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        RsData.Close
        Set RsData = Nothing
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)
    xReport.Database.SetDataSource RsData

    Dim cCompanyInfo As New ClsCompanyInfo

    If SystemOptions.UserInterface = ArabicInterface Then
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
        ' xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
       ' StrReportTitle = "" '& StrAccountName
        'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        '    StrReportTitle = StrReportTitle + " »œ«Ì… „‰ " & Format(Me.DTPickerAccFrom.value, "yyyy/M/d") & ""
        'End If
        'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        '    StrReportTitle = StrReportTitle + " ≈·Ï " & Format(Me.DTPickerAccTo.value, "yyyy/M/d") & " "
        'End If
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        'xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
    '    StrReportTitle = ""
        'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        '    StrReportTitle = StrReportTitle + " From Date " & (Me.DTPickerAccFrom.value) & ""
        'End If
        'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        '    StrReportTitle = StrReportTitle + " To Date :  " & (Me.DTPickerAccTo.value) & ""
        'End If
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
      '  xReport.ParameterFields(4).AddCurrentValue WriteNo(Format(val(TxtAdvanceValue.text), "0.00"), 0, True, ".")
       ' xReport.ParameterFields(6).AddCurrentValue val(lbl(23).Caption)
       '  xReport.ParameterFields(7).AddCurrentValue DBIssueDate.value
'    xReport.ParameterFields(8).AddCurrentValue IIf(IsNumeric(fg.TextMatrix(Me.fg.FixedRows, fg.ColIndex("PartValue"))), val(fg.TextMatrix(Me.fg.FixedRows, fg.ColIndex("PartValue"))), 0)
' xReport.ParameterFields(9).AddCurrentValue val(lbl(22).Caption)
 ' xReport.ParameterFields(10).AddCurrentValue val(TxtDiscount.text)
  ' xReport.ParameterFields(11).AddCurrentValue txtDiscountDES.text
  Dim total As String
  Dim totl As Double
 ' totl = val(LbToTalExtra.Caption) + val(Me.lbTotalMente.Caption)
 ' total = totl
 '  xReport.ParameterFields(12).AddCurrentValue Me.lbTotalMente.Caption
 '     xReport.ParameterFields(13).AddCurrentValue LbToTalExtra.Caption
 '       xReport.ParameterFields(14).AddCurrentValue total
   ' xReport.ParameterFields(5).AddCurrentValue ToHijriDate(RsData("notedate").value)
   
   
       Dim i As Integer
      For i = 1 To xReport.FormulaFields.count
        Select Case xReport.FormulaFields.Item(i).Name
        Case "{@Title}"
            xReport.FormulaFields.Item(i).text = "'" & Trim(StrReportTitle) & "'"
        Case "{@FromDate}"
            xReport.FormulaFields.Item(i).text = "'" & Trim(DtpDateFrom.value) & "'"
        Case "{@ToDate}"
            xReport.FormulaFields.Item(i).text = "'" & Trim(DtpDateTo.value) & "'"
        End Select
    Next i
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault


 
  
 
End Function


 
Private Sub TxtSearchCode_KeyPress(KeyAscii As Integer)
Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode TxtSearchCode.text, EmpID
        DcboEmpName.BoundText = EmpID
    End If
End Sub
