VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SUITEC~1.OCX"
Begin VB.Form FrmConttractTotalService 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   5835
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10050
   Icon            =   "FrmContracTotalService.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5835
   ScaleWidth      =   10050
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      BackColor       =   &H00E2E9E9&
      Height          =   4605
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   480
      Width           =   10035
      Begin VB.TextBox TxtSearch 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4440
         TabIndex        =   28
         Top             =   720
         Width           =   855
      End
      Begin XtremeSuiteControls.CheckBox RDRentValue 
         Height          =   255
         Left            =   4560
         TabIndex        =   21
         Top             =   1320
         Width           =   1455
         _Version        =   786432
         _ExtentX        =   2566
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "«·«ÌÃ«— Œ·«· › —Â"
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
      Begin VB.Frame Frame3 
         Height          =   4455
         Left            =   6600
         TabIndex        =   19
         Top             =   120
         Width           =   3375
         Begin VB.Label lblCompanyname 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "«·”« —Ì…"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   27.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   5295
            Left            =   120
            TabIndex        =   20
            Top             =   2520
            Width           =   2895
         End
         Begin VB.Image Image1 
            Height          =   2310
            Left            =   120
            Picture         =   "FrmContracTotalService.frx":038A
            Stretch         =   -1  'True
            Top             =   120
            Width           =   3225
         End
      End
      Begin MSComCtl2.DTPicker DtpDateFrom 
         Height          =   330
         Left            =   2040
         TabIndex        =   8
         Top             =   3720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   94830595
         CurrentDate     =   38887
      End
      Begin MSComCtl2.DTPicker DtpDateTo 
         Height          =   330
         Left            =   2040
         TabIndex        =   9
         Top             =   4080
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   94830595
         CurrentDate     =   38887
      End
      Begin MSDataListLib.DataCombo Dcbranch 
         Bindings        =   "FrmContracTotalService.frx":10A48
         Height          =   315
         Left            =   1320
         TabIndex        =   14
         Top             =   360
         Width           =   3975
         _ExtentX        =   7011
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
      Begin Dynamic_Byte.NourHijriCal DtpDateFromH 
         Height          =   315
         Left            =   360
         TabIndex        =   17
         Top             =   3720
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   556
      End
      Begin Dynamic_Byte.NourHijriCal DtpDateToH 
         Height          =   315
         Left            =   360
         TabIndex        =   18
         Top             =   4080
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   556
      End
      Begin XtremeSuiteControls.CheckBox RdWater 
         Height          =   255
         Left            =   4560
         TabIndex        =   22
         Top             =   1680
         Width           =   1455
         _Version        =   786432
         _ExtentX        =   2566
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "«·„Ì«Â Œ·«· › —Â"
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox RdElctricity 
         Height          =   255
         Left            =   4440
         TabIndex        =   23
         Top             =   2040
         Width           =   1575
         _Version        =   786432
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "«·ﬂÂ—»«¡ Œ·«· › —Â"
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox RdCommiion 
         Height          =   255
         Left            =   4440
         TabIndex        =   24
         Top             =   2400
         Width           =   1575
         _Version        =   786432
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "«·”⁄Ì Œ·«· › —Â"
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox RdInsurance 
         Height          =   255
         Left            =   4440
         TabIndex        =   25
         Top             =   2760
         Width           =   1575
         _Version        =   786432
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "«· √„Ì‰ Œ·«· › —Â"
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox RdTelandNet 
         Height          =   255
         Left            =   4440
         TabIndex        =   26
         Top             =   3120
         Width           =   1575
         _Version        =   786432
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "«·Œœ„«  Œ·«· › —Â"
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox RdAll 
         Height          =   255
         Left            =   4080
         TabIndex        =   27
         Top             =   3480
         Width           =   1935
         _Version        =   786432
         _ExtentX        =   3413
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "ﬂ· «· Õ’Ì·«  Œ·«· › —Â"
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo dcbAqarType 
         Height          =   315
         Left            =   1320
         TabIndex        =   29
         Top             =   720
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   "·⁄ﬁ«— „⁄Ì‰"
         Height          =   195
         Index           =   1
         Left            =   5310
         TabIndex        =   30
         Top             =   720
         Width           =   720
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Caption         =   "Ì—ÃÏ «Œ Ì«— «·›—⁄ «Ê «· «—ÌŒ «Ê ”Ê› ÌﬂÊ‰ «· ﬁ—Ì— «Ã„«·Ì ·ﬂ· «·›—Ê⁄  Ê«·„œ…"
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
         Height          =   1890
         Index           =   2
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Shape Shape2 
         BorderWidth     =   2
         Height          =   2175
         Left            =   120
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label lblBr 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "·›—⁄ „⁄Ì‰"
         Height          =   255
         Left            =   5160
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   195
         Index           =   4
         Left            =   3450
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   3690
         Width           =   540
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï"
         Height          =   195
         Index           =   3
         Left            =   3510
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   4050
         Width           =   480
      End
   End
   Begin VB.ComboBox DcbOrderStatus 
      Height          =   315
      Left            =   12240
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   1200
      Width           =   1935
   End
   Begin VB.ComboBox ComGranty 
      Height          =   315
      Left            =   12360
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   120
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   12360
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Text            =   " „ „Ê«›ﬁ… «·⁄„"
      Top             =   720
      Width           =   915
   End
   Begin ImpulseButton.ISButton Cmd 
      Cancel          =   -1  'True
      Height          =   495
      Index           =   2
      Left            =   30
      TabIndex        =   0
      Top             =   5160
      Width           =   1125
      _ExtentX        =   1984
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
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   330
      Left            =   12480
      TabIndex        =   7
      Top             =   2400
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   582
      _Version        =   393216
      CheckBox        =   -1  'True
      Format          =   94830595
      CurrentDate     =   38887
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   495
      Index           =   0
      Left            =   2400
      TabIndex        =   12
      Top             =   5160
      Width           =   1125
      _ExtentX        =   1984
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
      Left            =   1170
      TabIndex        =   13
      Top             =   5160
      Width           =   1125
      _ExtentX        =   1984
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
   Begin VB.Image ImgFavorites 
      Height          =   390
      Left            =   120
      Picture         =   "FrmContracTotalService.frx":10A5D
      Stretch         =   -1  'True
      Top             =   0
      Width           =   525
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "«· ﬁ«—Ì— «·⁄«„…  ·· Õ’Ì·« "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   -30
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   0
      Width           =   10080
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      ForeColor       =   &H00000080&
      Height          =   285
      Index           =   1
      Left            =   60
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   3060
      Width           =   1785
   End
End
Attribute VB_Name = "FrmConttractTotalService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
'Dim DCboSearch As clsDCboSearch

Dim TTP As clstooltip

Dim TTD As clstooltipdemand

'Public Order As String
'Public gr As String


Private Sub Cmd_Click(Index As Integer)

    Select Case Index

        Case 0
       

 GetData
            
        Case 1
            clear_all Me
DtpDateFrom.value = ""
DtpDateTo.value = ""
RDRentValue.value = vbUnchecked
RdWater.value = vbUnchecked
RdElctricity.value = vbUnchecked
RdCommiion.value = vbUnchecked
RdInsurance.value = vbUnchecked
RdTelandNet.value = vbUnchecked
RdAll.value = vbUnchecked

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




Private Sub DtpDateFrom_Change()
If DtpDateFrom.value <> "" Then
   DtpDateFromH.value = ToHijriDate(DtpDateFrom.value)
   End If
End Sub

Private Sub DtpDateFromH_LostFocus()
 VBA.Calendar = vbCalGreg
            DtpDateFrom.value = ToGregorianDate(DtpDateFromH.value)
End Sub



Private Sub DtpDateTo_Change()
If DtpDateTo.value <> "" Then
   DtpDateToH.value = ToHijriDate(DtpDateTo.value)
   End If
End Sub



Private Sub DtpDateToH_LostFocus()
 VBA.Calendar = vbCalGreg
            DtpDateTo.value = ToGregorianDate(DtpDateToH.value)
End Sub

Private Sub Form_Activate()
   PutFormOnTop Me.hWnd
End Sub
Private Sub dcbAqarType_Click(Area As Integer)
      If val(dcbAqarType.BoundText) = 0 Then: Exit Sub
Dim str As String
    Dim EmpCode  As String
 Dim ownerid As Double
    GetIqarCode , , dcbAqarType.BoundText, EmpCode, ownerid
    
    Me.TxtSearch.Text = EmpCode
End Sub
Private Sub dcbAqarType_Change()
dcbAqarType_Click (0)
End Sub
Private Sub TxtSearch_KeyPress(KeyAscii As Integer)
  Dim EmpID As Double
'GetTblCustemersCode
    If KeyAscii = vbKeyReturn Then
        GetIqarCode TxtSearch.Text, EmpID
        dcbAqarType.BoundText = EmpID
        dcbAqarType_Click (0)
    End If
End Sub

Private Sub Form_Load()
    Dim GrdBack As ClsBackGroundPic
    Dim Dcombos As ClsDataCombos
    'AddTip
        Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500

    Set Dcombos = New ClsDataCombos
 Dcombos.GetBranches Me.Dcbranch
    'Dcombos.GetUsers Me.DCUser
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Clear").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
Dcombos.GetIqar dcbAqarType
DtpDateFrom.value = ""
DtpDateTo.value = ""
RDRentValue.value = vbUnchecked
RdWater.value = vbUnchecked
RdElctricity.value = vbUnchecked
RdCommiion.value = vbUnchecked
RdInsurance.value = vbUnchecked
RdTelandNet.value = vbUnchecked
RdAll.value = vbUnchecked
    Set GrdBack = New ClsBackGroundPic

 
 If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
       ' ChangeLang
    End If
    SetDtpickerDate Me.DtpDateFrom
    SetDtpickerDate Me.DtpDateTo

End Sub

Private Sub Form_Unload(Cancel As Integer)

    FormPostion Me, SavePostion
   ' Set DCboSearch = Nothing
End Sub

Public Sub GetData()
   
    Dim StrSQL As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer


StrSQL = " SELECT     dbo.ContracttBillInstallmentsDone.RentValuePayed, dbo.ContracttBillInstallmentsDone.CommissionsPayed, dbo.ContracttBillInstallmentsDone.InsurancePayed, "
StrSQL = StrSQL & "                      dbo.ContracttBillInstallmentsDone.WaterPayed, dbo.ContracttBillInstallmentsDone.ElectricPayed, dbo.ContracttBillInstallmentsDone.TelandNetPayed,"
StrSQL = StrSQL & "                      dbo.ContracttBillInstallmentsDone.NoteID, dbo.Notes.EmpId, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Name1,"
StrSQL = StrSQL & "                      dbo.TblEmployee.Emp_Name2, dbo.TblEmployee.Emp_Name3, dbo.TblEmployee.Emp_Name4, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee4,"
StrSQL = StrSQL & "                      dbo.TblEmployee.Emp_Namee3, dbo.TblEmployee.Emp_Namee2, dbo.TblEmployee.Emp_Namee1, dbo.TblEmployee.Emp_Namee,"
StrSQL = StrSQL & "                      dbo.ContracttBillInstallmentsDone.istallid, dbo.TblContractInstallments.ContNo, dbo.TblContract.ContDate, dbo.TblContract.NoteSerial1, dbo.TblContract.Iqar,"
StrSQL = StrSQL & "                      dbo.TblAqar.Aqarid, dbo.TblAqar.aqarNo, dbo.TblAqar.aqarname, dbo.TblContract.UnitType, dbo.TblAkarUnit.name, dbo.TblAkarUnit.namee, dbo.TblContract.UnitNo,"
StrSQL = StrSQL & "                      dbo.TblAqarDetai.unitno AS unitnoName, dbo.ContracttBillInstallmentsDone.RecordDate, dbo.ContracttBillInstallmentsDone.RecordDateH,"
StrSQL = StrSQL & "                      dbo.ContracttBillInstallmentsDone.total, dbo.ContracttBillInstallmentsDone.id, dbo.TblContract.CusID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee,"
StrSQL = StrSQL & "                      dbo.TblCustemers.Fullcode AS FullcodeCus, dbo.TblContract.Branch_NO, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee"
StrSQL = StrSQL & " FROM         dbo.TblBranchesData RIGHT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblContract ON dbo.TblBranchesData.branch_id = dbo.TblContract.Branch_NO LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblCustemers ON dbo.TblContract.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblAqarDetai ON dbo.TblContract.UnitNo = dbo.TblAqarDetai.Id LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblAkarUnit ON dbo.TblContract.UnitType = dbo.TblAkarUnit.id LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblAqar ON dbo.TblContract.Iqar = dbo.TblAqar.Aqarid RIGHT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblContractInstallments ON dbo.TblContract.ContNo = dbo.TblContractInstallments.ContNo RIGHT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.ContracttBillInstallmentsDone ON dbo.TblContractInstallments.id = dbo.ContracttBillInstallmentsDone.istallid LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.Notes ON dbo.ContracttBillInstallmentsDone.NoteID = dbo.Notes.NoteID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblEmployee ON dbo.Notes.EmpId = dbo.TblEmployee.Emp_ID"
StrSQL = StrSQL & " Where  (1 = 1)"
    BolBegine = False
    StrWhere = ""



    If Dcbranch.Text <> "" And val(Dcbranch.BoundText) <> 0 Then
        
            StrWhere = StrWhere + " and dbo.TblContract.Branch_NO =" & val((Trim(Dcbranch.BoundText)))
         
    End If
    
       If Me.dcbAqarType.Text <> "" And val(dcbAqarType.BoundText) <> 0 Then
        
            StrWhere = StrWhere + " and dbo.TblAqar.Aqarid=" & val((Trim(dcbAqarType.BoundText)))
         
    End If
    
    If Not IsNull(Me.DtpDateFrom.value) Then
                   StrWhere = StrWhere & " AND dbo.ContracttBillInstallmentsDone.RecordDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
      End If

    If Not IsNull(Me.DtpDateTo.value) Then
            StrWhere = StrWhere & " AND  dbo.ContracttBillInstallmentsDone.RecordDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
     
    End If


 


    '-----------------------------------

    StrSQL = StrSQL & StrWhere
 
   StrSQL = StrSQL & " Order By dbo.ContracttBillInstallmentsDone.ID"
   
   

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
           ' Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
          '  Me.lbl(10).Caption = "Search Results=0"
        End If

        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «· ﬁ—Ì—"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    Else
     'Me.DTPicker1.value = Format(rs("DateAcutExite").value, "yyyy/M/d")
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

Dim v1, v2, v3, v4, v5, v6 As Integer
If Me.RdAll.value = vbChecked Then
v1 = 1
v2 = 1
v3 = 1
v4 = 1
v5 = 1
v6 = 1
Else

If Me.RDRentValue.value = vbChecked Then
v1 = 1
Else
v1 = 0
End If
If Me.RdWater.value = vbChecked Then
v2 = 1
Else
v2 = 0
End If
If Me.RdElctricity.value = vbChecked Then
v3 = 1
Else
v3 = 0
End If
If Me.RdCommiion.value = vbChecked Then
v4 = 1
Else
v4 = 0
End If
If Me.RdInsurance.value = vbChecked Then
v5 = 1
Else
v5 = 0
End If
If Me.RdInsurance.value = vbChecked Then
v6 = 1
Else
v6 = 0
End If
End If
If v1 = 0 And v2 = 0 And v3 = 0 And v4 = 0 And v5 = 0 And v6 = 0 Then
MsgBox "ÌÃ» «Œ Ì«— «Õœ «· Õ’Ì·«  «Ê «ﬂÀ—"
Exit Function
End If


        If SystemOptions.UserInterface = ArabicInterface Then
        
       ' If Me.RDRentValue.value = vbChecked Then
       '     StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepCotctSerceByRentValue.rpt"
       '     Else
       'If Me.RdWater.value = vbChecked Then
       '     StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepCotctSerceByWater.rpt"
       '     Else
      'If Me.RdElctricity.value = vbChecked Then
      '      StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepCotctSerceByElectr.rpt"
      '      Else
     ' If Me.RdCommiion.value = vbChecked Then
     '       StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepCotctSerceByCommission.rpt"
     '       Else
     ' If Me.RdInsurance.value = vbChecked Then
     '       StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepCotctSerceByInsurance.rpt"
     '       Else
     ' If Me.RdTelandNet.value = vbChecked Then
     '       StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepCotctSerceByTelandNet.rpt"
     '        Else
     '       If Me.RdAll.value = vbChecked Then
     '       StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepCotctServices.rpt"
     '        Else
    
             StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepCotctServicesDynamic.rpt"
     '
     '     End If
     '       End If
     '  End If
     '        End If
     '
     '       End If
     '       End If
     '       End If
            
          
            
           
             
        Else
     ' If Me.RDRentValue.value = vbChecked Then
     '      StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepCotctSerceByRentValue.rpt"
     '       Else
  ' If Me.RdWater.value = vbChecked Then
     '       StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepCotctSerceByWater.rpt"
     '       Else
     '  If Me.RdElctricity.value = vbChecked Then
     '       StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepCotctSerceByElectr.rpt"
     '       Else
     '  If Me.RdCommiion.value = vbChecked Then
    ''        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepCotctSerceByCommission.rpt"
    '        Else
    '   If Me.RdInsurance.value = vbChecked Then
    '        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepCotctSerceByInsurance.rpt"
    '        Else
    '     If Me.RdTelandNet.value = vbChecked Then
    '        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepCotctSerceByTelandNet.rpt"
    '         Else
    '        If Me.RdAll.value = vbChecked Then
    '        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepCotctServices.rpt"
    '         Else
    '
          '   StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepCotctServices.rpt"
             StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepCotctServicesDynamic.rpt"
          '  End If
          ' End If
          '  End If
          '  End If
          '   End If
          '  End If
          '  End If
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

    If SystemOptions.UserInterface = ArabicInterface Then
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
        ' xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
        StrReportTitle = "" '& StrAccountName
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
        StrReportTitle = ""
        'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        '    StrReportTitle = StrReportTitle + " From Date " & (Me.DTPickerAccFrom.value) & ""
        'End If
        'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        '    StrReportTitle = StrReportTitle + " To Date :  " & (Me.DTPickerAccTo.value) & ""
        'End If
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
    If DtpDateFrom.value <> "" And DtpDateTo.value <> "" Then
    xReport.ParameterFields(4).AddCurrentValue Me.DtpDateFrom.value
    xReport.ParameterFields(5).AddCurrentValue DtpDateFromH.value
     xReport.ParameterFields(6).AddCurrentValue DtpDateTo.value
    xReport.ParameterFields(7).AddCurrentValue DtpDateToH.value
    End If
    xReport.ParameterFields(8).AddCurrentValue v1
    xReport.ParameterFields(9).AddCurrentValue v2
     xReport.ParameterFields(10).AddCurrentValue v3
    xReport.ParameterFields(11).AddCurrentValue v4
     xReport.ParameterFields(12).AddCurrentValue v5
    xReport.ParameterFields(13).AddCurrentValue v6
       ' xReport.ParameterFields(6).AddCurrentValue val(lbl(23).Caption)
        ' xReport.ParameterFields(13).AddCurrentValue Me.DTPicker1.value
'    xReport.ParameterFields(8).AddCurrentValue IIf(IsNumeric(fg.TextMatrix(Me.fg.FixedRows, fg.ColIndex("PartValue"))), val(fg.TextMatrix(Me.fg.FixedRows, fg.ColIndex("PartValue"))), 0)
'Dim gr, order As Integer
' xReport.ParameterFields(14).AddCurrentValue Order
' xReport.ParameterFields(15).AddCurrentValue gr
 ' xReport.ParameterFields(15).AddCurrentValue gr
 ' xReport.ParameterFields(10).AddCurrentValue val(TxtDiscount.text)
  ' xReport.ParameterFields(11).AddCurrentValue txtDiscountDES.text
'  Dim total As String
'  Dim totl As Double
 ' totl = val(LbToTalExtra.Caption) + val(Me.lbTotalMente.Caption)
 ' total = totl
 '  xReport.ParameterFields(12).AddCurrentValue Me.lbTotalMente.Caption
 '     xReport.ParameterFields(13).AddCurrentValue LbToTalExtra.Caption
 '       xReport.ParameterFields(14).AddCurrentValue total
   ' xReport.ParameterFields(5).AddCurrentValue ToHijriDate(RsData("notedate").value)
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

 
Private Sub ImgFavorites_Click()
AddTofaforites Me.Name, Me.Caption, Me.Caption
End Sub



