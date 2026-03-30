VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SUITEC~1.OCX"
Begin VB.Form FrmOwnerAqarReport 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   5940
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10425
   Icon            =   "FrmOwnerAqarReport.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5940
   ScaleWidth      =   10425
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton btnClear 
      BackColor       =   &H00E2E9E9&
      Caption         =   "„”Õ"
      Height          =   495
      Left            =   2760
      TabIndex        =   21
      Top             =   5400
      Width           =   1335
   End
   Begin VB.Frame XPPnlTime 
      BackColor       =   &H00E2E9E9&
      Caption         =   "›Ï «·› —…"
      Height          =   1185
      Left            =   4320
      TabIndex        =   11
      Top             =   6720
      Visible         =   0   'False
      Width           =   2415
      Begin MSComCtl2.DTPicker XPDtbFrom 
         Height          =   345
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   609
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   98893825
         CurrentDate     =   38784
      End
      Begin MSComCtl2.DTPicker XPDtpTo 
         Height          =   345
         Left            =   120
         TabIndex        =   13
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   609
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   98893825
         CurrentDate     =   38784
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   285
         Index           =   2
         Left            =   1560
         TabIndex        =   15
         Top             =   360
         Width           =   465
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï"
         Height          =   285
         Index           =   0
         Left            =   1560
         TabIndex        =   14
         Top             =   720
         Width           =   465
      End
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   "»Ì«‰«  "
      Height          =   4605
      Left            =   0
      TabIndex        =   3
      Top             =   720
      Width           =   10395
      Begin VB.Frame Frame1 
         BackColor       =   &H00E2E9E9&
         Caption         =   "Õœœ ‰Ê⁄ «· ﬁ—Ì—"
         Height          =   960
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   1320
         Width           =   6555
         Begin XtremeSuiteControls.RadioButton Rd1 
            Height          =   375
            Left            =   3360
            TabIndex        =   32
            Top             =   120
            Width           =   1935
            _Version        =   786432
            _ExtentX        =   3413
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   " ﬁ—Ì— «·„·«ﬂ"
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton Rd2 
            Height          =   375
            Left            =   0
            TabIndex        =   33
            Top             =   120
            Width           =   3135
            _Version        =   786432
            _ExtentX        =   5530
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "ﬂ‘›  «·«ÌÃ«—«  «·„” Õﬁ… ·„·«ﬂ «·⁄„«∆—"
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton Rd3 
            Height          =   375
            Left            =   2280
            TabIndex        =   35
            Top             =   480
            Width           =   3015
            _Version        =   786432
            _ExtentX        =   5318
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "ﬂ‘› «·„»«·€ «·„”œœ ··„·«ﬂ "
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
      End
      Begin VB.TextBox TxtSearch 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4320
         TabIndex        =   30
         Top             =   600
         Width           =   855
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00E2E9E9&
         Caption         =   "Õœœ «·› —… "
         Height          =   960
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   2160
         Width           =   6555
         Begin MSComCtl2.DTPicker Fromdate 
            Height          =   330
            Left            =   3495
            TabIndex        =   23
            Top             =   240
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   98893825
            CurrentDate     =   41640
         End
         Begin Dynamic_Byte.NourHijriCal Fromdateh 
            Height          =   330
            Left            =   3480
            TabIndex        =   24
            Top             =   600
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   582
         End
         Begin Dynamic_Byte.NourHijriCal todateH 
            Height          =   330
            Left            =   240
            TabIndex        =   25
            Top             =   600
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   582
         End
         Begin MSComCtl2.DTPicker toDate 
            Height          =   330
            Left            =   240
            TabIndex        =   26
            Top             =   240
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   98893825
            CurrentDate     =   41640
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "„‰"
            Height          =   315
            Index           =   3
            Left            =   4980
            RightToLeft     =   -1  'True
            TabIndex        =   28
            Top             =   480
            Width           =   945
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "≈«·Ï"
            Height          =   435
            Index           =   14
            Left            =   2580
            RightToLeft     =   -1  'True
            TabIndex        =   27
            Top             =   480
            Width           =   540
         End
      End
      Begin VB.Frame Frame3 
         Height          =   4455
         Left            =   6960
         TabIndex        =   19
         Top             =   120
         Width           =   3375
         Begin VB.Image Image1 
            Height          =   2310
            Left            =   120
            Picture         =   "FrmOwnerAqarReport.frx":038A
            Stretch         =   -1  'True
            Top             =   120
            Width           =   3300
         End
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
      End
      Begin VB.TextBox txtCodeBranch 
         Height          =   285
         Left            =   6360
         TabIndex        =   18
         Top             =   4680
         Width           =   855
      End
      Begin VB.TextBox txtCodeOwner 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4320
         TabIndex        =   17
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   5400
         TabIndex        =   16
         Top             =   4680
         Width           =   855
      End
      Begin MSDataListLib.DataCombo DcbBranch 
         Height          =   315
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo dcsupplier 
         Height          =   315
         Left            =   240
         TabIndex        =   7
         Top             =   960
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo dcbAqarType 
         Height          =   315
         Left            =   240
         TabIndex        =   8
         Top             =   600
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox Chck 
         Height          =   375
         Left            =   4800
         TabIndex        =   34
         Top             =   3240
         Width           =   1815
         _Version        =   786432
         _ExtentX        =   3201
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "ﬂ‘› Õ”«» «·„«·ﬂ"
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
      Begin VB.Shape Shape2 
         BorderWidth     =   2
         Height          =   855
         Left            =   0
         Top             =   3720
         Width           =   6975
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
         Height          =   810
         Index           =   4
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   3720
         Width           =   6975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÿ»ﬁ« ·„«·ﬂ „Õœœ"
         Height          =   195
         Index           =   2
         Left            =   5400
         TabIndex        =   10
         Top             =   960
         Width           =   1065
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÿ»ﬁ« ·⁄ﬁ«— „⁄Ì‰"
         Height          =   195
         Index           =   1
         Left            =   5355
         TabIndex        =   9
         Top             =   600
         Width           =   1065
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÿ»ﬁ« ·›—⁄ „⁄Ì‰"
         Height          =   195
         Index           =   0
         Left            =   5400
         TabIndex        =   5
         Top             =   240
         Width           =   1020
      End
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   495
      Index           =   0
      Left            =   1440
      TabIndex        =   0
      Top             =   5400
      Width           =   1245
      _ExtentX        =   2196
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
      Index           =   2
      Left            =   240
      TabIndex        =   1
      Top             =   5400
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
   Begin VB.Image ImgFavorites 
      Height          =   390
      Left            =   600
      Picture         =   "FrmOwnerAqarReport.frx":10A48
      Stretch         =   -1  'True
      Top             =   0
      Width           =   525
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "‘«‘…  ﬁ«—Ì— „·«ﬂ «·⁄ﬁ«—« "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   15
      TabIndex        =   6
      Top             =   0
      Width           =   10335
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      ForeColor       =   &H00000080&
      Height          =   285
      Index           =   1
      Left            =   60
      TabIndex        =   2
      Top             =   3060
      Width           =   1785
   End
End
Attribute VB_Name = "FrmOwnerAqarReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RsSavRec As ADODB.Recordset
Dim BKGrndPic As ClsBackGroundPic
Dim RecId As String
Dim II As Long
Dim cSearch  As clsDCboSearch
Dim rs As ADODB.Recordset
Dim DCboSearch As clsDCboSearch
'

Private Sub btnClear_Click()
Cmd_Click (1)
End Sub
Sub Acount()
If val(dcsupplier.BoundText) = 0 Then
MsgBox "Ì—ÃÏ «Œ Ì«— «·„«·ﬂ"
dcsupplier.SetFocus
Exit Sub
End If
If IsNull(ToDate.value) Or IsNull(FromDate.value) Then
MsgBox "Ì—ÃÏ  ÕœÌœ «·› —…"
Exit Sub
End If
        Dim Account_code As String
        Account_code = GetMyAccountCode("TblCustemers", "CusID", val(Me.dcsupplier.BoundText))
updateopeningbalanceNewFromsql FromDate.value, ToDate.value, False, 0, 0, Account_code, 3
        ShowReport Account_code, dcsupplier.Text, FromDate.value, ToDate.value
End Sub
Private Sub Cmd_Click(Index As Integer)

    Select Case Index

        Case 0
If Chck.value = vbChecked Then
Acount
Else
 GetData
End If
            
        Case 1
   
            clear_all Me
         FromDate.value = ""
    ToDate.value = ""
      Rd1.value = False
    Rd2.value = False
        Rd3.value = False
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






Private Sub dcbAqarType_Change()
dcbAqarType_Click (0)
End Sub

Private Sub dcbAqarType_Click(Area As Integer)
      If val(dcbAqarType.BoundText) = 0 Then: Exit Sub
Dim str As String
    Dim EmpCode  As String
 Dim ownerid As Double
    GetIqarCode , , dcbAqarType.BoundText, EmpCode, ownerid
    
    Me.TxtSearch.Text = EmpCode
End Sub

Private Sub dcsupplier_Change()
    dcsupplier_Click (0)
End Sub

Private Sub dcsupplier_Click(Area As Integer)
  If val(dcsupplier.BoundText) = 0 Then Exit Sub

    Dim EmpCode  As String
 
    GetTblCustemersCode , , dcsupplier.BoundText, EmpCode
    Me.txtCodeOwner.Text = EmpCode
End Sub

Private Sub Form_Activate()
   PutFormOnTop Me.hwnd
End Sub




Private Sub Form_Load()
   
 'On Error GoTo ErrTrap
    Dim i As Integer
    Dim My_SQL As String
    Dim Dcombos As ClsDataCombos
    
    Set Dcombos = New ClsDataCombos
    
    Dcombos.GetIqar dcbAqarType
    
   ' Dcombos.GetCountriesGovernCities dcmCityID
    
   ' Dcombos.getCountriesGovernments dcbCityId2
    
    Dcombos.GetCustomersSuppliers 57, Me.dcsupplier
    
    'Dcombos.getAkarUnit Me.DCAkarUnit
    
   ' Dcombos.GetSalesRepData Me.dcbSalesSpec
    
   ' Dcombos.GetCustomersSuppliers 1, Me.dbcClient
    
    Dcombos.GetBranches DcbBranch
    
 ' Dcombos.GetRentStatus dbcAqarStatus
    
    FromDate.value = ""
    ToDate.value = ""
    Rd1.value = False
    Rd2.value = False
    
    Set cSearch = New clsDCboSearch
  '  My_SQL = "TblContract"
    
    Set BKGrndPic = New ClsBackGroundPic
    Set RsSavRec = New ADODB.Recordset
    
   ' RsSavRec.CursorLocation = adUseClient
   ' RsSavRec.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdTableDirect
    
 
    
    Resize_Form Me
    
   
    
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
    'gr = 9
    'Order = 9
 
If Rd2.value = True Or Rd3.value = True Then
 StrSQL = "SELECT    suckno, ContValue ,isnull(NOOFYears,1) NOOFYears,    dbo.TblAqar.Priod, dbo.TblAqar.PriodDMY, dbo.TblAqrOwin.ID, dbo.TblAqrOwin.AqrID, dbo.TblAqar.aqarNo, dbo.TblAqar.aqarname, dbo.TblAqrOwin.RecDateH, dbo.TblAqrOwin.RecDate, dbo.TblAqrOwin.[value], "
 StrSQL = StrSQL & "                     dbo.TblAqrOwin.DMY, dbo.TblAqrOwin.Cont, dbo.TblAqrOwin.AllowDateH, dbo.TblAqrOwin.AllowDate, dbo.TblAqrOwin.PaymentNo, dbo.TblAqar.ownerid,"
 StrSQL = StrSQL & "                     dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.Cus_Phone, dbo.TblCustemers.Cus_mobile, dbo.TblAqar.BranchId,"
 StrSQL = StrSQL & "                     dbo.TblBranchesData.branch_name , dbo.TblBranchesData.branch_nameE , dbo.GetOwnerPayment(dbo.TblAqrOwin.ID) AS PayedValue"
 StrSQL = StrSQL & " FROM         dbo.TblBranchesData RIGHT OUTER JOIN"
 StrSQL = StrSQL & "                      dbo.TblAqar ON dbo.TblBranchesData.branch_id = dbo.TblAqar.BranchId LEFT OUTER JOIN"
 StrSQL = StrSQL & "                      dbo.TblCustemers ON dbo.TblAqar.ownerid = dbo.TblCustemers.CusID RIGHT OUTER JOIN"
 StrSQL = StrSQL & "                     dbo.TblAqrOwin ON dbo.TblAqar.Aqarid = dbo.TblAqrOwin.AqrID"
 StrSQL = StrSQL & "   Where  (1 = 1)"
    BolBegine = False
    StrWhere = ""
  If Not IsNull(Me.FromDate.value) Then
                   StrWhere = StrWhere & " AND dbo.TblAqrOwin.RecDate >=" & SQLDate(Me.FromDate.value, True) & ""
      End If

    If Not IsNull(Me.ToDate.value) Then
            StrWhere = StrWhere & " AND  dbo.TblAqrOwin.RecDate <=" & SQLDate(Me.ToDate.value, True) & ""
     
    End If
End If
If Rd1.value = True Then
StrSQL = "SELECT     dbo.TblAqar.aqarNo, dbo.TblAqar.aqarname, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.Cus_Phone, "
StrSQL = StrSQL & "                      dbo.TblCustemers.Cus_mobile, dbo.TblAqar.BranchId, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblAqar.aqartypeid,"
StrSQL = StrSQL & "                      dbo.TblAqar.Remarks, dbo.TblAqar.Provide, dbo.TblAqar.BanckName, dbo.TblAqar.SalesEmp, dbo.TblAqar.FristPaymentDate, dbo.TblAqar.FirstInstallDateH,"
StrSQL = StrSQL & "                      dbo.TblAqar.ToCotDate, dbo.TblAqar.ToCotDateH, dbo.TblAqar.FromCotDate, dbo.TblAqar.FromCotDateH, dbo.TblAqar.DateCont, dbo.TblAqar.DateHCont,"
StrSQL = StrSQL & "                      dbo.TblAqar.PriodAlowDMY, dbo.TblAqar.PriodAlow, dbo.TblAqar.PriodDMY, dbo.TblAqar.Priod, dbo.TblAqar.PaymentNo, dbo.TblAqar.ContValue,"
StrSQL = StrSQL & "                      dbo.TblAqar.AccountBank, dbo.TblAqar.Fax, dbo.TblAqar.Email, dbo.TblAqar.Mobile, dbo.TblAqar.Telephone, dbo.TblAqar.AgemcyNo, dbo.TblAqar.westWriiten,"
StrSQL = StrSQL & "                      dbo.TblAqar.eastWriiten, dbo.TblAqar.PriceHad, dbo.TblAqar.PriceSom, dbo.TblAqar.Price, dbo.TblAqar.PriceHadW, dbo.TblAqar.PriceSomW, dbo.TblAqar.StreetNo,"
StrSQL = StrSQL & "                      dbo.TblAqar.Part, dbo.TblAqar.Block, dbo.TblAqar.UnitNo, dbo.TblAqar.Rate, dbo.TblAqar.authorizationname, dbo.TblAqar.suckno, dbo.TblAqar.suckdateH,"
StrSQL = StrSQL & "                      dbo.TblAqar.suckdate, dbo.TblAqar.statusdate, dbo.TblAqar.GoogleMap, dbo.TblAqar.metersalevalue, dbo.TblAqar.meterRentvalue, dbo.TblAqar.totallength,"
StrSQL = StrSQL & "                      dbo.TblAqar.Westlength, dbo.TblAqar.eastlength, dbo.TblAqar.Southlength, dbo.TblAqar.northlength, dbo.TblAqar.lastrentvalue, dbo.TblAqar.currentPrice,"
StrSQL = StrSQL & "                      dbo.TblAqar.aqarage, dbo.TblAqar.Location, dbo.TblAqar.streetname, dbo.TblAqar.ownerid, dbo.TblAqar.maintenancetypeid, dbo.TblAqar.noofparking,"
StrSQL = StrSQL & "                      dbo.TblAqar.noofoffices, dbo.TblAqar.noofapartement, dbo.TblAqar.EntryCount, dbo.TblAqar.StatusId, dbo.TblAqar.interfaceid, dbo.TblAqar.schemeid,"
StrSQL = StrSQL & "                      dbo.TblAqar.floorcount , dbo.TblAqar.heyid, dbo.TblAqar.cityid, dbo.TblAqar.CountryID, dbo.TblAqar.Aqarid, dbo.TblCountriesGovernmentsCities.CityName"
StrSQL = StrSQL & " FROM         dbo.TblCountriesGovernmentsCities RIGHT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblAqar ON dbo.TblCountriesGovernmentsCities.CityID = dbo.TblAqar.heyid LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblBranchesData ON dbo.TblAqar.BranchId = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblCustemers ON dbo.TblAqar.ownerid = dbo.TblCustemers.CusID"
StrSQL = StrSQL & "   Where  (1 = 1)"
    BolBegine = False
    StrWhere = ""
      If Not IsNull(Me.FromDate.value) Then
                   StrWhere = StrWhere & " AND dbo.TblAqar.FromCotDate >=" & SQLDate(Me.FromDate.value, True) & ""
      End If

    If Not IsNull(Me.ToDate.value) Then
            StrWhere = StrWhere & " AND  dbo.TblAqar.FromCotDate <=" & SQLDate(Me.ToDate.value, True) & ""
     
    End If
End If



    
    
If val(Me.DcbBranch.BoundText) <> 0 Or Me.DcbBranch.Text <> "" Then
StrWhere = StrWhere & " AND TblAqar.BranchId = " & val(Me.DcbBranch.BoundText)
'gr = 0
End If


If val(Me.dcbAqarType.BoundText) <> 0 Or Me.dcbAqarType.Text <> "" Then
'gr = 1
StrWhere = StrWhere & " AND dbo.TblAqar.Aqarid = " & val(Me.dcbAqarType.BoundText)
'gr = 1
End If


If val(Me.dcsupplier.BoundText) <> 0 Or Me.dcsupplier.Text <> "" Then
StrWhere = StrWhere & " AND dbo.TblAqar.ownerid  = " & val(dcsupplier.BoundText)
'gr = 2
End If

  




    '-----------------------------------
If Rd1.value = True Or Rd2.value = True Or Rd3.value = True Then
    StrSQL = StrSQL & StrWhere
 
  StrSQL = StrSQL & " order by  dbo.TblAqar.Aqarid "
  
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
   '  Me.DTPicker1.value = Format(rs("DateAcutExite").value, "yyyy/M/d")
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
If Me.Rd2.value = True Then


        If SystemOptions.UserInterface = ArabicInterface Then
        
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepOwnerAqarReport.rpt"
            Else
             StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepOwnerAqarReport.rpt"
            
       End If
    End If
    
   
   
   
 If Me.Rd3.value = True Then


        If SystemOptions.UserInterface = ArabicInterface Then
        
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepOwnerAqarReport3.rpt"
            Else
             StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepOwnerAqarReport3.rpt"
            
       End If
    End If
    
    
        If Me.Rd1.value = True Then


        If SystemOptions.UserInterface = ArabicInterface Then
        
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepOwnerAqarReportsh.rpt"
            Else
             StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepOwnerAqarReportsh.rpt"
            
       End If
    End If
            
    ' If Me.RdDept.value = True Then
           ' StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rep1byDept.rpt"
     '       Else
      '      If Me.RdSuper.value = True Then
       '     StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rep1bySuper.rpt"
        '    Else
         '   If Me.RdFitter.value = True Then
           ' StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rep1byFitter.rpt"
          ' Else
             
            '        If Me.RdAll2.value = True Then
         '   StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rep1all.rpt"
          '  Else
           '  StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rep1.rpt"
            
     '      End If
      '      End If
       '     End If
        '     End If
         '   End If
          '  End If
        '    End If
           ' End If
          '  End If
       '      End If
           '
      '  End If
If Rd1.value = True Or Rd2.value = True Or Rd3.value = True Then


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
        If Me.FromDate.value <> Empty Or Me.FromDate.value <> Null Then
            StrReportTitle = StrReportTitle + " »œ«Ì… „‰ " & Format(Me.FromDate.value, "yyyy/M/d") & ""
        End If
        If Me.ToDate.value <> Empty Or Me.ToDate.value <> Null Then
           StrReportTitle = StrReportTitle + " ≈·Ï " & Format(Me.ToDate.value, "yyyy/M/d") & " "
        End If
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

    xReport.ParameterFields(4).AddCurrentValue user_name
   ' xReport.ParameterFields(13).AddCurrentValue Me.DTPicker1.value
      '  xReport.ParameterFields(4).AddCurrentValue WriteNo(Format(val(TxtAdvanceValue.text), "0.00"), 0, True, ".")
       ' xReport.ParameterFields(6).AddCurrentValue val(lbl(23).Caption)
        'xReport.ParameterFields(13).AddCurrentValue Me.DTPicker1.value
'    xReport.ParameterFields(8).AddCurrentValue IIf(IsNumeric(fg.TextMatrix(Me.fg.FixedRows, fg.ColIndex("PartValue"))), val(fg.TextMatrix(Me.fg.FixedRows, fg.ColIndex("PartValue"))), 0)
'Dim gr, order As Integer
' xReport.ParameterFields(14).AddCurrentValue Order
 'xReport.ParameterFields(15).AddCurrentValue gr
 ' xReport.ParameterFields(15).AddCurrentValue gr
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
 xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.title
    xReport.ReportAuthor = App.title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName, , NoteSerial
    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault

End If
 
  
 
End Function




Private Sub FromDate_Change()
If FromDate.value <> "" Then
   FromDateH.value = ToHijriDate(FromDate.value)
   End If
End Sub



Private Sub Fromdateh_LostFocus()

 VBA.Calendar = vbCalGreg
            FromDate.value = ToGregorianDate(FromDateH.value)

End Sub

Private Sub ImgFavorites_Click()
AddTofaforites Me.Name, Me.Caption, Me.Caption
End Sub

Private Sub ToDate_Change()
If ToDate.value <> "" Then
   todateH.value = ToHijriDate(ToDate.value)
   End If
End Sub

Private Sub ToDateH_LostFocus()

 VBA.Calendar = vbCalGreg
            ToDate.value = ToGregorianDate(todateH.value)

End Sub


Private Sub txtCodeOwner_KeyPress(KeyAscii As Integer)
 Dim EmpID As Integer
  If KeyAscii = vbKeyReturn Then
        GetTblCustemersCode txtCodeOwner.Text, EmpID, , , 57
        dcsupplier.BoundText = EmpID
   End If
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
