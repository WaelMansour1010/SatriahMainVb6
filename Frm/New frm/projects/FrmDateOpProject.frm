VERSION 5.00
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmDateOpProject 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " ”ÃÌ· «· «—ÌŒ "
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4875
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   4875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin ImpulseButton.ISButton CmdOk 
      Height          =   405
      Left            =   1020
      TabIndex        =   4
      Top             =   1170
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   714
      Caption         =   "Õ›Ÿ"
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
   Begin VB.TextBox TxtComment 
      Alignment       =   1  'Right Justify
      Height          =   975
      Left            =   30
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   4830
      Width           =   4425
   End
   Begin ImpulseButton.ISButton CmdCancel 
      Height          =   405
      Left            =   60
      TabIndex        =   5
      Top             =   1170
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   714
      Caption         =   "«·€«¡"
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
   Begin MSComCtl2.DTPicker XPDtbBill 
      Height          =   330
      Left            =   2400
      TabIndex        =   6
      Top             =   120
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   582
      _Version        =   393216
      Format          =   238878721
      CurrentDate     =   38784
   End
   Begin MSComCtl2.DTPicker DcTime 
      Height          =   330
      Left            =   2400
      TabIndex        =   10
      Top             =   600
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   582
      _Version        =   393216
      CheckBox        =   -1  'True
      Format          =   238878722
      CurrentDate     =   38784
   End
   Begin Dynamic_Byte.NourHijriCal Txt_DateHigri 
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   120
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
   End
   Begin MSComCtl2.DTPicker ToTime 
      Height          =   330
      Left            =   240
      TabIndex        =   14
      Top             =   600
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   582
      _Version        =   393216
      CheckBox        =   -1  'True
      Format          =   238878722
      CurrentDate     =   38784
   End
   Begin Dynamic_Byte.NourHijriCal NourHijriCal1 
      Height          =   255
      Left            =   2280
      TabIndex        =   15
      Top             =   1200
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   450
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   330
      Left            =   240
      TabIndex        =   16
      Top             =   120
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   582
      _Version        =   393216
      CheckBox        =   -1  'True
      Format          =   238878722
      CurrentDate     =   38784
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·”«⁄…"
      Height          =   255
      Index           =   2
      Left            =   1440
      TabIndex        =   17
      Top             =   120
      Width           =   795
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·Ï «·”«⁄…"
      Height          =   255
      Index           =   1
      Left            =   1200
      TabIndex        =   12
      Top             =   600
      Width           =   1155
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "„‰ «·”«⁄…"
      Height          =   255
      Index           =   0
      Left            =   3600
      TabIndex        =   11
      Top             =   600
      Width           =   1155
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   255
      Index           =   8
      Left            =   3000
      TabIndex        =   9
      Top             =   4920
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   255
      Index           =   7
      Left            =   0
      TabIndex        =   8
      Top             =   3120
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«· «—ÌŒ"
      Height          =   255
      Index           =   6
      Left            =   3600
      TabIndex        =   7
      Top             =   120
      Width           =   1155
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   4425
      X2              =   0
      Y1              =   1080
      Y2              =   1095
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   255
      Index           =   5
      Left            =   0
      TabIndex        =   2
      Top             =   4380
      Width           =   3645
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   255
      Index           =   4
      Left            =   2100
      TabIndex        =   1
      Top             =   3480
      Width           =   1545
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   255
      Index           =   3
      Left            =   2100
      TabIndex        =   0
      Top             =   1020
      Width           =   1545
   End
End
Attribute VB_Name = "FrmDateOpProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Public fg As VSFlex8UCtl.vsFlexGrid

'Public LngRow As Long

Public index As Integer

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdOk_Click()
Dim Period As Double
    Dim Msg As String
    Dim dateenter As Date
    Dim timEnter As Date
    Dim Askinterval As String
    Period = 0
On Error Resume Next

If index = 0 Then

    If Not Projects.VSFlexGrid2 Is Nothing Then
If Projects.LngCol = 21 Then
 Projects.VSFlexGrid2.TextMatrix(Projects.LngRow, Projects.VSFlexGrid2.ColIndex("StartDate")) = XPDtbBill.value
 ElseIf Projects.LngCol = 22 Then
 Projects.VSFlexGrid2.TextMatrix(Projects.LngRow, Projects.VSFlexGrid2.ColIndex("EndDate")) = XPDtbBill.value
 End If
   Unload Me
    End If
          ElseIf index = 540 Then
        FrmAccEditJournal.Fg_Journal.TextMatrix(FrmAccEditJournal.Fg_Journal.LngRow, FrmAccEditJournal.Fg_Journal.ColIndex("DueDate")) = XPDtbBill.value
        Unload Me

      ElseIf index = 541 Then
        FrmAccEditJournal1.Fg_Journal.TextMatrix(FrmAccEditJournal1.LngRow, FrmAccEditJournal1.Fg_Journal.ColIndex("DueDate")) = XPDtbBill.value
        Unload Me
      ElseIf index = 542 Then
      

      ElseIf index = 543 Then
        FrmAccEditJournal3.Fg_Journal.TextMatrix(FrmAccEditJournal3.LngRow, FrmAccEditJournal3.Fg_Journal.ColIndex("DueDate")) = XPDtbBill.value
        Unload Me

      ElseIf index = 544 Then
        FrmAccEditJournal4.Fg_Journal.TextMatrix(FrmAccEditJournal4.LngRow, FrmAccEditJournal4.Fg_Journal.ColIndex("DueDate")) = XPDtbBill.value
        Unload Me
  
    ElseIf index = 1 Then


    If Not FrmInstalVacation.Grid Is Nothing Then
    If FrmInstalVacation.LngCol = 4 Or FrmInstalVacation.LngCol = 5 Then
 FrmInstalVacation.Grid.TextMatrix(FrmInstalVacation.LngRow, FrmInstalVacation.Grid.ColIndex("BeginDate")) = XPDtbBill.value
  FrmInstalVacation.Grid.TextMatrix(FrmInstalVacation.LngRow, FrmInstalVacation.Grid.ColIndex("BeginDateH")) = Txt_DateHigri.value
 ElseIf FrmInstalVacation.LngCol = 6 Or FrmInstalVacation.LngCol = 7 Then
 FrmInstalVacation.Grid.TextMatrix(FrmInstalVacation.LngRow, FrmInstalVacation.Grid.ColIndex("LastDate")) = XPDtbBill.value
  FrmInstalVacation.Grid.TextMatrix(FrmInstalVacation.LngRow, FrmInstalVacation.Grid.ColIndex("LastDateH")) = Txt_DateHigri.value
 End If
   Unload Me

    End If
    
'      ElseIf Index = 2 Then
'    FrmEmpIncreaseSalaries.VSFlexGrid1.TextMatrix(FrmEmpIncreaseSalaries.LngRow, FrmEmpIncreaseSalaries.VSFlexGrid1.ColIndex("RecoedDate")) = XPDtbBill.value
'    Unload Me
'

'
'
'
     ElseIf index = 17 Then
    FrmExpensesInvestment.GridInstallments.TextMatrix(FrmExpensesInvestment.LngRow, FrmExpensesInvestment.GridInstallments.ColIndex("StartDate")) = XPDtbBill.value
    Unload Me
     ElseIf index = 18 Then
    FrmExpensesInvestment.GridInstallments.TextMatrix(FrmExpensesInvestment.LngRow, FrmExpensesInvestment.GridInstallments.ColIndex("EndDate")) = XPDtbBill.value
    Unload Me
        ElseIf index = 19 Then
    FrmImportShifts.FG.TextMatrix(FrmImportShifts.LngRow, FrmImportShifts.FG.ColIndex("MachinDate")) = XPDtbBill.value
    If Not IsNull(DcTime.value) Then
    FrmImportShifts.FG.TextMatrix(FrmImportShifts.LngRow, FrmImportShifts.FG.ColIndex("RecTime")) = FormatDateTime(DcTime.value, vbShortTime)
    Else
    FrmImportShifts.FG.TextMatrix(FrmImportShifts.LngRow, FrmImportShifts.FG.ColIndex("RecTime")) = ""
    End If
    If Not (IsNull(ToTime.value)) Then
    FrmImportShifts.FG.TextMatrix(FrmImportShifts.LngRow, FrmImportShifts.FG.ColIndex("ToTime")) = FormatDateTime(ToTime.value, vbShortTime)
    Else
    FrmImportShifts.FG.TextMatrix(FrmImportShifts.LngRow, FrmImportShifts.FG.ColIndex("ToTime")) = ""
    End If
    Unload Me
        ElseIf index = 20 Then
    FrmImportShifts.FG.TextMatrix(FrmImportShifts.LngRow, FrmImportShifts.FG.ColIndex("MachinDate")) = XPDtbBill.value
    If Not IsNull(DcTime.value) Then
    FrmImportShifts.FG.TextMatrix(FrmImportShifts.LngRow, FrmImportShifts.FG.ColIndex("RecTime")) = FormatDateTime(DcTime.value, vbShortTime)
    Else
    FrmImportShifts.FG.TextMatrix(FrmImportShifts.LngRow, FrmImportShifts.FG.ColIndex("RecTime")) = ""
    End If
    If Not (IsNull(ToTime.value)) Then
    FrmImportShifts.FG.TextMatrix(FrmImportShifts.LngRow, FrmImportShifts.FG.ColIndex("ToTime")) = FormatDateTime(ToTime.value, vbShortTime)
    Else
    FrmImportShifts.FG.TextMatrix(FrmImportShifts.LngRow, FrmImportShifts.FG.ColIndex("ToTime")) = ""
    End If
   Unload Me
      ElseIf index = 21 Then
    FrmImportShifts.FG.TextMatrix(FrmImportShifts.LngRow, FrmImportShifts.FG.ColIndex("MachinDate")) = XPDtbBill.value
    If Not IsNull(DcTime.value) Then
    FrmImportShifts.FG.TextMatrix(FrmImportShifts.LngRow, FrmImportShifts.FG.ColIndex("RecTime")) = FormatDateTime(DcTime.value, vbShortTime)
    Else
    FrmImportShifts.FG.TextMatrix(FrmImportShifts.LngRow, FrmImportShifts.FG.ColIndex("RecTime")) = ""
    End If
    If Not (IsNull(ToTime.value)) Then
    FrmImportShifts.FG.TextMatrix(FrmImportShifts.LngRow, FrmImportShifts.FG.ColIndex("ToTime")) = FormatDateTime(ToTime.value, vbShortTime)
    Else
    FrmImportShifts.FG.TextMatrix(FrmImportShifts.LngRow, FrmImportShifts.FG.ColIndex("ToTime")) = ""
     End If
    Unload Me
    ElseIf index = 22 Or index = 24 Then
    FrmComponentYear.GridInstallments.TextMatrix(FrmComponentYear.LngRow, FrmComponentYear.GridInstallments.ColIndex("RecDate1")) = XPDtbBill.value
    FrmComponentYear.GridInstallments.TextMatrix(FrmComponentYear.LngRow, FrmComponentYear.GridInstallments.ColIndex("RecDate1H")) = Txt_DateHigri.value
    Unload Me
     ElseIf index = 23 Or index = 25 Then
    FrmComponentYear.GridInstallments.TextMatrix(FrmComponentYear.LngRow, FrmComponentYear.GridInstallments.ColIndex("RecDate2")) = XPDtbBill.value
    FrmComponentYear.GridInstallments.TextMatrix(FrmComponentYear.LngRow, FrmComponentYear.GridInstallments.ColIndex("RecDate2H")) = Txt_DateHigri.value
    Unload Me
    ElseIf index = 26 Then
    FrmExitvisasReturn.GridInstallments.TextMatrix(FrmExitvisasReturn.LngRow, FrmExitvisasReturn.GridInstallments.ColIndex("IssuDate")) = XPDtbBill.value
    FrmExitvisasReturn.GridInstallments.TextMatrix(FrmExitvisasReturn.LngRow, FrmExitvisasReturn.GridInstallments.ColIndex("IssuDateH")) = Txt_DateHigri.value
    FrmExitvisasReturn.GridInstallments.TextMatrix(FrmExitvisasReturn.LngRow, FrmExitvisasReturn.GridInstallments.ColIndex("BeforDate")) = DateAdd("d", 90, XPDtbBill.value)
    NourHijriCal1.value = ToHijriDate(DateAdd("d", 90, XPDtbBill.value))
    FrmExitvisasReturn.GridInstallments.TextMatrix(FrmExitvisasReturn.LngRow, FrmExitvisasReturn.GridInstallments.ColIndex("BeforDateH")) = NourHijriCal1.value
    Unload Me
    ElseIf index = 27 Then
    Period = FrmExitvisasReturn.Period
    FrmExitvisasReturn.GridInstallments.TextMatrix(FrmExitvisasReturn.LngRow, FrmExitvisasReturn.GridInstallments.ColIndex("TravlDate")) = XPDtbBill.value
    FrmExitvisasReturn.GridInstallments.TextMatrix(FrmExitvisasReturn.LngRow, FrmExitvisasReturn.GridInstallments.ColIndex("TravlDateH")) = Txt_DateHigri.value
        FrmExitvisasReturn.GridInstallments.TextMatrix(FrmExitvisasReturn.LngRow, FrmExitvisasReturn.GridInstallments.ColIndex("ReturnDate")) = DateAdd("d", Period, XPDtbBill.value)
        NourHijriCal1.value = ToHijriDate(DateAdd("d", Period, XPDtbBill.value))
    FrmExitvisasReturn.GridInstallments.TextMatrix(FrmExitvisasReturn.LngRow, FrmExitvisasReturn.GridInstallments.ColIndex("ReturnDateH")) = NourHijriCal1.value
    Unload Me
    ElseIf index = 28 Then
    FrmExitvisasReturn.GridInstallments.TextMatrix(FrmExitvisasReturn.LngRow, FrmExitvisasReturn.GridInstallments.ColIndex("ActRetDate")) = XPDtbBill.value
    FrmExitvisasReturn.GridInstallments.TextMatrix(FrmExitvisasReturn.LngRow, FrmExitvisasReturn.GridInstallments.ColIndex("ActRetDateH")) = Txt_DateHigri.value
    Unload Me
    ElseIf index = 29 Then
    FrmApproveRequset.FG1.TextMatrix(FrmApproveRequset.LngRow, FrmApproveRequset.FG1.ColIndex("ApproveDate")) = XPDtbBill.value
    FrmApproveRequset.FG1.TextMatrix(FrmApproveRequset.LngRow, FrmApproveRequset.FG1.ColIndex("ApproveTime")) = DTPicker1.value
        Unload Me

    ElseIf index = 30 Then
    FrmEmpSalary3.Grid.TextMatrix(FrmEmpSalary3.LngRow, FrmEmpSalary3.Grid.ColIndex("FromDate")) = XPDtbBill.value
    If FrmEmpSalary3.Grid.TextMatrix(FrmEmpSalary3.LngRow, FrmEmpSalary3.Grid.ColIndex("FromDate")) <> "" And FrmEmpSalary3.Grid.TextMatrix(FrmEmpSalary3.LngRow, FrmEmpSalary3.Grid.ColIndex("ToDate")) <> "" Then
    FrmEmpSalary3.Grid.TextMatrix(FrmEmpSalary3.LngRow, FrmEmpSalary3.Grid.ColIndex("interval")) = DateDiff("d", FrmEmpSalary3.Grid.TextMatrix(FrmEmpSalary3.LngRow, FrmEmpSalary3.Grid.ColIndex("FromDate")), FrmEmpSalary3.Grid.TextMatrix(FrmEmpSalary3.LngRow, FrmEmpSalary3.Grid.ColIndex("ToDate"))) + 1
    End If
    Unload Me
    ElseIf index = 31 Then
    FrmEmpSalary3.Grid.TextMatrix(FrmEmpSalary3.LngRow, FrmEmpSalary3.Grid.ColIndex("ToDate")) = XPDtbBill.value & ""
     If FrmEmpSalary3.Grid.TextMatrix(FrmEmpSalary3.LngRow, FrmEmpSalary3.Grid.ColIndex("FromDate")) <> "" And FrmEmpSalary3.Grid.TextMatrix(FrmEmpSalary3.LngRow, FrmEmpSalary3.Grid.ColIndex("ToDate")) <> "" Then
    FrmEmpSalary3.Grid.TextMatrix(FrmEmpSalary3.LngRow, FrmEmpSalary3.Grid.ColIndex("interval")) = DateDiff("d", FrmEmpSalary3.Grid.TextMatrix(FrmEmpSalary3.LngRow, FrmEmpSalary3.Grid.ColIndex("FromDate")), FrmEmpSalary3.Grid.TextMatrix(FrmEmpSalary3.LngRow, FrmEmpSalary3.Grid.ColIndex("ToDate"))) + 1
    Else
        FrmEmpSalary3.Grid.TextMatrix(FrmEmpSalary3.LngRow, FrmEmpSalary3.Grid.ColIndex("interval")) = ""
    End If
    Unload Me
    ElseIf index = 610 Then

        If Not IsNull(XPDtbBill.value) Then
            FrmLC.GrdMargin3.TextMatrix(FrmLC.LngRow, FrmLC.GrdMargin3.ColIndex("GuaranteeDate")) = XPDtbBill.value
        End If
         Unload Me
    ElseIf index = 611 Then

        If Not IsNull(XPDtbBill.value) Then
            FrmLC.GrdMargin.TextMatrix(FrmLC.LngRow, FrmLC.GrdMargin.ColIndex("GuaranteeDate")) = XPDtbBill.value
        End If
         Unload Me
         
    ElseIf index = 612 Then

        If Not IsNull(XPDtbBill.value) Then
            FrmLC.GrdMargin2.TextMatrix(FrmLC.LngRow, FrmLC.GrdMargin2.ColIndex("OrderDate")) = XPDtbBill.value
        End If
         Unload Me
         
    ElseIf index = 613 Then

        If Not IsNull(XPDtbBill.value) Then
            FrmLC.GrdMargin2.TextMatrix(FrmLC.LngRow, FrmLC.GrdMargin2.ColIndex("GuaranteeDate")) = XPDtbBill.value
        End If
         Unload Me
         
    ElseIf index = 614 Then

        If Not IsNull(XPDtbBill.value) Then
            FrmLC.GrdMargin4.TextMatrix(FrmLC.LngRow, FrmLC.GrdMargin4.ColIndex("GuaranteeDate")) = XPDtbBill.value
        End If
         Unload Me
         
    ElseIf index = 615 Then

        If Not IsNull(XPDtbBill.value) Then
            FrmLC.GrdMargin4.TextMatrix(FrmLC.LngRow, FrmLC.GrdMargin4.ColIndex("OrderDate")) = XPDtbBill.value
        End If
         Unload Me
         
         
    ElseIf index = 616 Then

        If Not IsNull(XPDtbBill.value) Then
            FrmLC.GrdMargin4.TextMatrix(FrmLC.LngRow, FrmLC.GrdMargin4.ColIndex("payDate")) = XPDtbBill.value
        End If
         Unload Me
         
    ElseIf index = 617 Then

        If Not IsNull(XPDtbBill.value) Then
            FrmLC.GrdMargin2.TextMatrix(FrmLC.LngRow, FrmLC.GrdMargin2.ColIndex("payDate")) = XPDtbBill.value
        End If
         Unload Me
    ElseIf index = 33 Then

        If Not IsNull(XPDtbBill.value) Then
            FrmInstallmentVendorAlarm.grd.TextMatrix(FrmInstallmentVendorAlarm.LngRow, FrmInstallmentVendorAlarm.grd.ColIndex("DateStart")) = XPDtbBill.value
        End If
        If Not IsNull(DcTime.value) Then
            FrmInstallmentVendorAlarm.grd.TextMatrix(FrmInstallmentVendorAlarm.LngRow, FrmInstallmentVendorAlarm.grd.ColIndex("TimeStart")) = DcTime.value
        End If
        FrmInstallmentVendorAlarm.grd.TextMatrix(FrmInstallmentVendorAlarm.LngRow, FrmInstallmentVendorAlarm.grd.ColIndex("UserId")) = user_id
        FrmInstallmentVendorAlarm.grd.TextMatrix(FrmInstallmentVendorAlarm.LngRow, FrmInstallmentVendorAlarm.grd.ColIndex("UserName")) = user_name
        Unload Me
    ElseIf index = 34 Then
        If Not IsNull(XPDtbBill.value) Then
            FrmInstallmentVendorAlarm.grd.TextMatrix(FrmInstallmentVendorAlarm.LngRow, FrmInstallmentVendorAlarm.grd.ColIndex("DateEnd")) = XPDtbBill.value
        End If
        If Not IsNull(DcTime.value) Then
            FrmInstallmentVendorAlarm.grd.TextMatrix(FrmInstallmentVendorAlarm.LngRow, FrmInstallmentVendorAlarm.grd.ColIndex("TimeEnd")) = DcTime.value
        End If
        FrmInstallmentVendorAlarm.grd.TextMatrix(FrmInstallmentVendorAlarm.LngRow, FrmInstallmentVendorAlarm.grd.ColIndex("UserId")) = user_id
        FrmInstallmentVendorAlarm.grd.TextMatrix(FrmInstallmentVendorAlarm.LngRow, FrmInstallmentVendorAlarm.grd.ColIndex("UserName")) = user_name

    'FrmInstallmentVendorAlarm.grd.TextMatrix(FrmInstallmentVendorAlarm.LngRow, FrmInstallmentVendorAlarm.grd.ColIndex("DateEnd")) = XPDtbBill.value
    Unload Me
    ElseIf index = 35 Then
        If Not IsNull(XPDtbBill.value) Then
            Frm_TradingContract.GridInstallments.TextMatrix(Frm_TradingContract.LngRow, Frm_TradingContract.GridInstallments.ColIndex("FinishDate")) = XPDtbBill.value
        End If
        Unload Me
    ElseIf index = 36 Then
        If Not IsNull(XPDtbBill.value) Then
            emp_CONTRACT_TYPE.grd2(1).TextMatrix(emp_CONTRACT_TYPE.LngRow, emp_CONTRACT_TYPE.grd2(1).ColIndex("FromDate")) = XPDtbBill.value
        End If
        Unload Me
    ElseIf index = 37 Then
        If Not IsNull(XPDtbBill.value) Then
            emp_CONTRACT_TYPE.grd2(1).TextMatrix(emp_CONTRACT_TYPE.LngRow, emp_CONTRACT_TYPE.grd2(1).ColIndex("ToDate")) = XPDtbBill.value
        End If
        Unload Me
    ElseIf index = 38 Then
        If Not IsNull(XPDtbBill.value) Then
            emp_CONTRACT_TYPE.grd2(2).TextMatrix(emp_CONTRACT_TYPE.LngRow, emp_CONTRACT_TYPE.grd2(2).ColIndex("FromDate")) = XPDtbBill.value
        End If
        Unload Me
    ElseIf index = 39 Then
        If Not IsNull(XPDtbBill.value) Then
            emp_CONTRACT_TYPE.grd2(2).TextMatrix(emp_CONTRACT_TYPE.LngRow, emp_CONTRACT_TYPE.grd2(2).ColIndex("ToDate")) = XPDtbBill.value
        End If
        Unload Me
        
        
End If

End Sub
Private Sub ChangeLang()
    cmdCancel.Caption = "Cancel"
cmdOK.Caption = "Save"
lbl(6).Caption = "DateEnter"
' 'bl(9).Caption = "TimeEnter"
Me.Caption = "Register Date "


End Sub

Private Sub DatePicker1_SelectionChanged()
'XPDtbBill.value = DatePicker1.AttachToCalendar
End Sub

Private Sub Form_Load()
    CenterForm Me
XPDtbBill.value = Date
DcTime = Time
ToTime.value = Time
lbl(0).Visible = False
lbl(1).Visible = False
ToTime.Visible = False
DTPicker1.Visible = False
lbl(2).Visible = False
DcTime.value = ""
Me.ToTime.value = ""
XPDtbBill.value = Date
    FormPostion Me, GetPostion
XPDtbBill.Visible = True
Txt_DateHigri.Visible = True


If index = 30 Then
DTPicker1.Visible = False
lbl(2).Visible = True
DTPicker1.Visible = True
Txt_DateHigri.Visible = False

 If FrmEmpSalary3.Grid.TextMatrix(FrmEmpSalary3.LngRow, FrmEmpSalary3.Grid.ColIndex("FromDate")) <> "" Then
   XPDtbBill.value = FrmEmpSalary3.Grid.TextMatrix(FrmEmpSalary3.LngRow, FrmEmpSalary3.Grid.ColIndex("FromDate"))
   End If

End If
If index = 31 Then
DTPicker1.Visible = False
lbl(2).Visible = True
DTPicker1.Visible = True
Txt_DateHigri.Visible = False

 If FrmEmpSalary3.Grid.TextMatrix(FrmEmpSalary3.LngRow, FrmEmpSalary3.Grid.ColIndex("ToDate")) <> "" Then
   XPDtbBill.value = FrmEmpSalary3.Grid.TextMatrix(FrmEmpSalary3.LngRow, FrmEmpSalary3.Grid.ColIndex("ToDate"))
   End If

End If

If index = 29 Then
lbl(2).Visible = True
DTPicker1.Visible = True
Txt_DateHigri.Visible = False
On Error Resume Next

         If FrmApproveRequset.FG1.TextMatrix(FrmApproveRequset.LngRow, FrmApproveRequset.FG1.ColIndex("ApproveDate")) <> "" Then
           XPDtbBill.value = FrmApproveRequset.FG1.TextMatrix(FrmApproveRequset.LngRow, FrmApproveRequset.FG1.ColIndex("ApproveDate"))
           End If
           If FrmApproveRequset.FG1.TextMatrix(FrmApproveRequset.LngRow, FrmApproveRequset.FG1.ColIndex("ApproveTime")) <> "" Then
           DTPicker1.value = FrmApproveRequset.FG1.TextMatrix(FrmApproveRequset.LngRow, FrmApproveRequset.FG1.ColIndex("ApproveTime"))
           End If

        End If

        If index = 26 Then
         If FrmExitvisasReturn.GridInstallments.TextMatrix(FrmExitvisasReturn.LngRow, FrmExitvisasReturn.GridInstallments.ColIndex("IssuDate")) <> "" Then
           XPDtbBill.value = FrmExitvisasReturn.GridInstallments.TextMatrix(FrmExitvisasReturn.LngRow, FrmExitvisasReturn.GridInstallments.ColIndex("IssuDate"))
           End If
           If FrmExitvisasReturn.GridInstallments.TextMatrix(FrmExitvisasReturn.LngRow, FrmExitvisasReturn.GridInstallments.ColIndex("IssuDateH")) <> "" Then
           Txt_DateHigri.value = FrmExitvisasReturn.GridInstallments.TextMatrix(FrmExitvisasReturn.LngRow, FrmExitvisasReturn.GridInstallments.ColIndex("IssuDateH"))
           End If
        End If
        If index = 27 Then
         If FrmExitvisasReturn.GridInstallments.TextMatrix(FrmExitvisasReturn.LngRow, FrmExitvisasReturn.GridInstallments.ColIndex("TravlDate")) <> "" Then
           XPDtbBill.value = FrmExitvisasReturn.GridInstallments.TextMatrix(FrmExitvisasReturn.LngRow, FrmExitvisasReturn.GridInstallments.ColIndex("TravlDate"))
           End If
           If FrmExitvisasReturn.GridInstallments.TextMatrix(FrmExitvisasReturn.LngRow, FrmExitvisasReturn.GridInstallments.ColIndex("TravlDateH")) <> "" Then
           Txt_DateHigri.value = FrmExitvisasReturn.GridInstallments.TextMatrix(FrmExitvisasReturn.LngRow, FrmExitvisasReturn.GridInstallments.ColIndex("TravlDateH"))
           End If
        End If
        If index = 28 Then
         If FrmExitvisasReturn.GridInstallments.TextMatrix(FrmExitvisasReturn.LngRow, FrmExitvisasReturn.GridInstallments.ColIndex("ActRetDate")) <> "" Then
           XPDtbBill.value = FrmExitvisasReturn.GridInstallments.TextMatrix(FrmExitvisasReturn.LngRow, FrmExitvisasReturn.GridInstallments.ColIndex("ActRetDate"))
           End If
           If FrmExitvisasReturn.GridInstallments.TextMatrix(FrmExitvisasReturn.LngRow, FrmExitvisasReturn.GridInstallments.ColIndex("ActRetDateH")) <> "" Then
           Txt_DateHigri.value = FrmExitvisasReturn.GridInstallments.TextMatrix(FrmExitvisasReturn.LngRow, FrmExitvisasReturn.GridInstallments.ColIndex("ActRetDateH"))
           End If
        End If

        If index = 22 Or index = 24 Then
         If FrmComponentYear.GridInstallments.TextMatrix(FrmComponentYear.LngRow, FrmComponentYear.GridInstallments.ColIndex("RecDate1")) <> "" Then
           XPDtbBill.value = FrmComponentYear.GridInstallments.TextMatrix(FrmComponentYear.LngRow, FrmComponentYear.GridInstallments.ColIndex("RecDate1"))
           End If
           If FrmComponentYear.GridInstallments.TextMatrix(FrmComponentYear.LngRow, FrmComponentYear.GridInstallments.ColIndex("RecDate1H")) <> "" Then
           Txt_DateHigri.value = FrmComponentYear.GridInstallments.TextMatrix(FrmComponentYear.LngRow, FrmComponentYear.GridInstallments.ColIndex("RecDate1H"))
           End If
        End If
        If index = 23 Or index = 25 Then
         If FrmComponentYear.GridInstallments.TextMatrix(FrmComponentYear.LngRow, FrmComponentYear.GridInstallments.ColIndex("RecDate2")) <> "" Then
           XPDtbBill.value = FrmComponentYear.GridInstallments.TextMatrix(FrmComponentYear.LngRow, FrmComponentYear.GridInstallments.ColIndex("RecDate2"))
           End If
          If FrmComponentYear.GridInstallments.TextMatrix(FrmComponentYear.LngRow, FrmComponentYear.GridInstallments.ColIndex("RecDate2H")) <> "" Then
           Txt_DateHigri.value = FrmComponentYear.GridInstallments.TextMatrix(FrmComponentYear.LngRow, FrmComponentYear.GridInstallments.ColIndex("RecDate2H"))
           End If
        End If

lbl(6).Caption = "«· «—ÌŒ"
DcTime.Visible = False
If index = 19 Or index = 20 Or index = 21 Then
ToTime.Visible = True
lbl(0).Visible = True
lbl(1).Visible = True
Txt_DateHigri.Visible = False
DcTime.Visible = True


       If FrmImportShifts.FG.TextMatrix(FrmImportShifts.LngRow, FrmImportShifts.FG.ColIndex("MachinDate")) <> "" Then
       XPDtbBill.value = FrmImportShifts.FG.TextMatrix(FrmImportShifts.LngRow, FrmImportShifts.FG.ColIndex("MachinDate"))
       End If
      If FrmImportShifts.FG.TextMatrix(FrmImportShifts.LngRow, FrmImportShifts.FG.ColIndex("RecTime")) <> "" Then
      DcTime.value = FrmImportShifts.FG.TextMatrix(FrmImportShifts.LngRow, FrmImportShifts.FG.ColIndex("RecTime"))
      End If
        If FrmImportShifts.FG.TextMatrix(FrmImportShifts.LngRow, FrmImportShifts.FG.ColIndex("ToTime")) <> "" Then
        Me.ToTime.value = FrmImportShifts.FG.TextMatrix(FrmImportShifts.LngRow, FrmImportShifts.FG.ColIndex("ToTime"))
        End If
    
      End If
 

If index = 33 Or index = 34 Then
    lbl(0).Visible = True
    lbl(0).Caption = "«·Êﬁ "
    DcTime.Visible = True
    DcTime.CheckBox = True
    
End If
If index = 35 Then
    lbl(0).Visible = False
    
    DcTime.Visible = False
    DcTime.CheckBox = False
    
End If

If index = 36 Or index = 37 Or index = 38 Or index = 39 Then
    lbl(0).Visible = False
    
    DcTime.Visible = False
    DcTime.CheckBox = False
    
End If



    Me.cmdOK.ButtonStyle = impActive
    Set cmdOK.ButtonImage = mdifrmmain.ImgLstMenuIcons.ListImages("Save").Picture
    cmdOK.ButtonPositionImage = impRightOfText

    Me.cmdCancel.ButtonStyle = impActive
    Set cmdCancel.ButtonImage = mdifrmmain.ImgLstMenuIcons.ListImages("Hide").Picture
    cmdCancel.ButtonPositionImage = impRightOfText
    
'Me.timeEnter.value = Time
If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)

    FormPostion Me, SavePostion
End Sub



Private Sub Txt_DateHigri_LostFocus()
VBA.Calendar = vbCalGreg
            XPDtbBill.value = ToGregorianDate(Txt_DateHigri.value)
End Sub

Private Sub XPDtbBill_Change()
If IsDate(XPDtbBill.value) Then
    Txt_DateHigri.value = ToHijriDate(XPDtbBill.value)
End If
End Sub
