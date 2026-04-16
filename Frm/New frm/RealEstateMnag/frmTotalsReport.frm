VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SUITEC~1.OCX"
Begin VB.Form FrmTotalsReport 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "  ﬁ«—Ì— «Ã„«·ÌÂ"
   ClientHeight    =   5460
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10365
   Icon            =   "frmTotalsReport.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   10365
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
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   495
      Left            =   5880
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   5760
      Width           =   1095
      _cx             =   1931
      _cy             =   873
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
   End
   Begin VB.CommandButton btnClear 
      Caption         =   "„”Õ"
      Height          =   495
      Left            =   2760
      TabIndex        =   9
      Top             =   4920
      Width           =   1335
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   "«· ﬁ«—Ì— «·«Ã„«·Ì…"
      Height          =   4005
      Left            =   0
      TabIndex        =   3
      Top             =   840
      Width           =   10395
      Begin VB.Frame Frame1 
         BackColor       =   &H00E2E9E9&
         Caption         =   "»Ì‰«  Œ«’Â »«·⁄ﬁÊœ"
         Height          =   1095
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   2760
         Width           =   5775
         Begin XtremeSuiteControls.CheckBox RdEntry 
            Height          =   255
            Left            =   3960
            TabIndex        =   25
            Top             =   720
            Width           =   1095
            _Version        =   786432
            _ExtentX        =   1931
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "œ«Œ·Ì"
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin VB.TextBox Text1 
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
            Left            =   3960
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   22
            Top             =   240
            Width           =   1065
         End
         Begin MSDataListLib.DataCombo dcsupplier 
            Height          =   315
            Left            =   120
            TabIndex        =   23
            Tag             =   "⁄›Ê« Ì—ÃÏ «Œ Ì«—√”„ «·„«·ﬂ"
            Top             =   240
            Width           =   3795
            _ExtentX        =   6694
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox RdyExtranl 
            Height          =   255
            Left            =   2520
            TabIndex        =   26
            Top             =   720
            Width           =   1095
            _Version        =   786432
            _ExtentX        =   1931
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Œ«—ÃÌ"
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox RdAll 
            Height          =   255
            Left            =   -120
            TabIndex        =   27
            Top             =   720
            Width           =   975
            _Version        =   786432
            _ExtentX        =   1720
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "«·ﬂ·"
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox LegalIssue 
            Height          =   255
            Left            =   960
            TabIndex        =   28
            Top             =   720
            Width           =   1215
            _Version        =   786432
            _ExtentX        =   2143
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "‘∆Ê‰ ﬁ«‰Ê‰Ì…"
            ForeColor       =   255
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   " «·„«·ﬂ"
            Height          =   165
            Index           =   1
            Left            =   4920
            RightToLeft     =   -1  'True
            TabIndex        =   24
            Top             =   240
            Width           =   810
         End
      End
      Begin VB.OptionButton opt_4 
         BackColor       =   &H00E2E9E9&
         Height          =   615
         Left            =   5160
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   2160
         Width           =   255
      End
      Begin VB.OptionButton opt_3 
         BackColor       =   &H00E2E9E9&
         Height          =   615
         Left            =   5160
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   1800
         Width           =   255
      End
      Begin VB.OptionButton opt_2 
         BackColor       =   &H00E2E9E9&
         Height          =   615
         Left            =   5160
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   1320
         Width           =   255
      End
      Begin VB.OptionButton opt_1 
         BackColor       =   &H00E2E9E9&
         Height          =   615
         Left            =   5160
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   840
         Width           =   255
      End
      Begin VB.Frame Frame3 
         Height          =   3855
         Left            =   6600
         TabIndex        =   7
         Top             =   120
         Width           =   3735
         Begin VB.Image Image1 
            Height          =   2190
            Left            =   240
            Picture         =   "frmTotalsReport.frx":038A
            Stretch         =   -1  'True
            Top             =   120
            Width           =   2940
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
            Height          =   1095
            Left            =   240
            TabIndex        =   8
            Top             =   2400
            Width           =   2895
         End
      End
      Begin MSDataListLib.DataCombo DcbBranch 
         Height          =   315
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "√Ã„«·Ï «⁄ﬁÊœ"
         Height          =   285
         Index           =   0
         Left            =   2400
         TabIndex        =   20
         Top             =   2400
         Width           =   2505
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "√Ã„«·Ï ⁄œœ «·ÊÕœ«  ›Ï «·‘—ﬂ…"
         Height          =   285
         Index           =   5
         Left            =   2400
         TabIndex        =   17
         Top             =   2040
         Width           =   2505
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "√Ã„«·Ï ⁄œœ «·„”ÊﬁÌ‰"
         Height          =   285
         Index           =   4
         Left            =   3000
         TabIndex        =   15
         Top             =   1560
         Width           =   1905
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "√Ã„«·Ï ⁄œœ «·⁄ﬁ«—« "
         Height          =   285
         Index           =   3
         Left            =   3000
         TabIndex        =   13
         Top             =   1080
         Width           =   1905
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
         Top             =   480
         Width           =   1020
      End
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   495
      Index           =   0
      Left            =   1440
      TabIndex        =   0
      Top             =   4920
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
      Top             =   4920
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
   Begin MSDataListLib.DataCombo DataCombo2 
      Height          =   315
      Left            =   240
      TabIndex        =   10
      Top             =   2040
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VB.Image ImgFavorites 
      Height          =   390
      Left            =   3000
      Picture         =   "frmTotalsReport.frx":10A48
      Stretch         =   -1  'True
      Top             =   0
      Width           =   525
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00E2E9E9&
      Caption         =   "ÿ»ﬁ« ·„” √Ã— „Õœœ"
      Height          =   195
      Index           =   5
      Left            =   5400
      TabIndex        =   11
      Top             =   2040
      Width           =   1290
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "‘«‘…«· ﬁ«—Ì— «·«Ã„«·Ì… "
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
      Left            =   -30
      TabIndex        =   6
      Top             =   0
      Width           =   10365
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
Attribute VB_Name = "FrmTotalsReport"
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


Private Sub btnClear_Click()
clear_all Me
Me.opt_4.value = False
RdEntry.value = vbUnchecked
Me.RdyExtranl.value = vbUnchecked
Me.RdAll.value = vbUnchecked
End Sub

Private Sub Cmd_Click(Index As Integer)

    Select Case Index

        Case 0
 If Me.opt_4.value = True Then
GetData
End If
If opt_1.value = True Then
GetData
End If

If opt_2.value = True Then
GetData
End If

If opt_3.value = True Then
GetData
End If


' GetData
            
        Case 1
            clear_all Me
'DtpDateFrom.value = ""
'DtpDateTo.value = ""
'Me.DtStart.value = ""
'Me.DtEnd.value = ""
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
Private Sub dcsupplier_Change()
dcsupplier_Click (0)
End Sub

Private Sub Form_Activate()
   PutFormOnTop Me.hwnd
End Sub
Private Sub Form_Load()
   
 'On Error GoTo ErrTrap
    Dim i As Integer
    Dim My_SQL As String
    Dim Dcombos As ClsDataCombos
    Me.opt_4.value = False
    Set Dcombos = New ClsDataCombos
    Frame1.Enabled = False
    
    Dcombos.GetCustomersSuppliers 57, Me.dcsupplier
    Dcombos.GetBranches DcbBranch
    Set cSearch = New clsDCboSearch
    My_SQL = "TblContract"
    
    Set BKGrndPic = New ClsBackGroundPic
    Set RsSavRec = New ADODB.Recordset
    
    RsSavRec.CursorLocation = adUseClient
    RsSavRec.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdTableDirect
    
    
    
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
If Me.opt_4.value = True Then
StrSQL = " SELECT     dbo.TblContract.ContNo, dbo.TblContract.ContType, dbo.TblContract.ContDate, dbo.TblContract.Iqar, dbo.TblAqar.aqarname, dbo.TblContract.UnitType, "
             StrSQL = StrSQL & "           dbo.TblAkarUnit.name, dbo.TblAkarUnit.namee, dbo.TblContract.UnitNo, dbo.TblAqarDetai.unitno AS unitnoname, dbo.TblContract.Branch_NO,"
           StrSQL = StrSQL & "             dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblContract.Water, dbo.TblContract.Electricity, dbo.TblContract.Phone,"
             StrSQL = StrSQL & "           dbo.TblContract.Enternet, dbo.TblContract.IncresYearValue, dbo.TblContract.IncresYearRate, dbo.TblContract.TotalContract, dbo.TblContract.CommiValue,"
           StrSQL = StrSQL & "             dbo.TblContract.PayAmini, dbo.TblContract.InsuranceValue, dbo.TblContract.PaymentCount, dbo.TblContract.FristPaymentDate, dbo.TblContract.PeriodsID,"
         StrSQL = StrSQL & "               dbo.TblContract.Periods, dbo.TblContract.Furnishing, dbo.TblContract.Remarks, dbo.TblContract.RecorddateH, dbo.TblContract.FromdateH, dbo.TblContract.TodateH,"
           StrSQL = StrSQL & "             dbo.TblContract.StrDate, dbo.TblContract.RentType, dbo.TblContract.EndDate, dbo.TblContract.MeterValue, dbo.TblContract.MeterCount,"
         StrSQL = StrSQL & "               dbo.TblContract.FirstInstallDateH, dbo.TblContract.NoteID, dbo.TblContract.NoteSerial, dbo.TblContract.NoteSerial1, dbo.TblContract.NewOrOpeneing,"
      StrSQL = StrSQL & "                  dbo.TblContract.OthersRules, dbo.TblContract.OutContract, dbo.TblContract.OldRent, dbo.TblContract.OldWater, dbo.TblContract.OldElectric, dbo.TblContract.oldCommi,"
     StrSQL = StrSQL & "                   dbo.TblContract.DivWater, dbo.TblContract.DivElectric, dbo.TblContract.OldInsurance, dbo.TblContract.balanceDate, dbo.TblContract.balanceDateH,"
     StrSQL = StrSQL & "                   dbo.TblContract.balanceDes, dbo.TblContract.Renew, dbo.TblContract.ContNoOld, dbo.TblContract.FromdateHO, dbo.TblContract.FromdateO,"
    StrSQL = StrSQL & "                    dbo.TblContract.EndContract, dbo.TblContract.Employeecontract, dbo.TblContract.Emp_IDContract, dbo.TblContract.OutOffice, dbo.TblContract.LegalIssue,"
    StrSQL = StrSQL & "                    dbo.TblContract.NotID, dbo.TblContract.NoteSrial1, dbo.TblContract.NotValue, dbo.TblBranchesData.branch_id, dbo.TblContract.CusID, dbo.TblCustemers.CusName,"
   StrSQL = StrSQL & "                     dbo.TblCustemers.CusNamee, dbo.TblContract.ownerid, TblCustemers_1.CusName AS CusNameOwe, TblCustemers_1.CusNamee AS CusNameOweE"
  StrSQL = StrSQL & " FROM         dbo.TblContract LEFT OUTER JOIN"
     StrSQL = StrSQL & "                   dbo.TblCustemers TblCustemers_1 ON dbo.TblContract.ownerid = TblCustemers_1.CusID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                    dbo.TblCustemers ON dbo.TblContract.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
    StrSQL = StrSQL & "                    dbo.TblBranchesData ON dbo.TblContract.Branch_NO = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
    StrSQL = StrSQL & "                    dbo.TblAqarDetai ON dbo.TblContract.UnitNo = dbo.TblAqarDetai.Id LEFT OUTER JOIN"
    StrSQL = StrSQL & "                    dbo.TblAkarUnit ON dbo.TblContract.UnitType = dbo.TblAkarUnit.id LEFT OUTER JOIN"
   StrSQL = StrSQL & "                     dbo.TblAqar ON dbo.TblContract.Iqar = dbo.TblAqar.Aqarid"
  StrSQL = StrSQL & " Where  (1 = 1)"
If val(Me.DcbBranch.BoundText) <> 0 Then
    StrSQL = StrSQL & " AND   dbo.TblBranchesData.branch_id = " & val(Me.DcbBranch.BoundText)

End If
If val(Me.dcsupplier.BoundText) <> 0 Then
    StrSQL = StrSQL & " AND   dbo.TblContract.ownerid= " & val(Me.dcsupplier.BoundText)

End If
If Me.RdEntry.value = vbChecked Then
 StrSQL = StrSQL & " AND   dbo.TblContract.OutContract is null "
End If

If Me.LegalIssue.value = vbChecked Then
 StrSQL = StrSQL & " AND   dbo.TblContract.LegalIssue =1"
End If


If Me.RdyExtranl.value = vbChecked Then
 StrSQL = StrSQL & " AND   dbo.TblContract.OutContract= 1"
End If
If Me.RdAll.value = vbChecked Then
 StrSQL = StrSQL & " AND  ( dbo.TblContract.OutContract =1 or dbo.TblContract.OutContract is null )"
End If
End If

If opt_1.value = True Then
StrSQL = "SELECT     MAX(dbo.TblAqar.aqarname) AS aqarname, dbo.TblBranchesData.branch_name, COUNT(dbo.TblAqar.aqarname) AS aqar_count ,  dbo.TblBranchesData.branch_id, "
   StrSQL = StrSQL + "                   dbo.TblBranchesData.branch_nameE"
StrSQL = StrSQL + " FROM         dbo.TblAqar RIGHT OUTER JOIN"
StrSQL = StrSQL + "                      dbo.TblBranchesData ON dbo.TblAqar.BranchId = dbo.TblBranchesData.branch_id"

StrSQL = StrSQL & " Where  (1 = 1)"
If Me.DcbBranch.BoundText <> "" Then
    StrSQL = StrSQL & " AND   dbo.TblBranchesData.branch_id = " & val(Me.DcbBranch.BoundText)
'gr = 0
End If
StrSQL = StrSQL + " GROUP BY dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee ,dbo.TblBranchesData.branch_id"

End If






If opt_2.value = True Then

StrSQL = " SELECT     dbo.TBLSalesRepData.BranchId, dbo.TBLSalesRepData.EmpID, dbo.TBLSalesRepData.GroupID, dbo.TBLSalesRepGroups.name, dbo.TBLSalesRepGroups.namee, "
StrSQL = StrSQL & "                      dbo.TblBranchesData.branch_name , dbo.TblBranchesData.branch_nameE, dbo.TblBranchesData.branch_id"
StrSQL = StrSQL & " FROM         dbo.TBLSalesRepData LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblBranchesData ON dbo.TBLSalesRepData.BranchId = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TBLSalesRepGroups ON dbo.TBLSalesRepData.GroupID = dbo.TBLSalesRepGroups.id"


StrSQL = StrSQL & " Where  (1 = 1)"
If Me.DcbBranch.BoundText <> "" Then
    StrSQL = StrSQL & " AND   dbo.TblBranchesData.branch_id = " & val(Me.DcbBranch.BoundText)
'gr = 0
End If
'StrSQL = StrSQL + " dbo.TblBranchesData.branch_id, dbo.TblBranchesData.branch_name"
End If







If opt_3.value = True Then

StrSQL = " SELECT     dbo.TblBranchesData.branch_id, dbo.TblBranchesData.branch_name, MAX(dbo.TblAqar.aqarname) AS aqarname, dbo.TblAkarUnit.name,"
StrSQL = StrSQL + "                       COUNT(dbo.TblAqar.aqarname) AS Unit_Count, dbo.TblBranchesData.branch_name + '  ' + dbo.TblAkarUnit.name AS grp"
StrSQL = StrSQL + " FROM         dbo.TblAqar INNER JOIN"
                   StrSQL = StrSQL + "    dbo.TblAqarDetai ON dbo.TblAqar.Aqarid = dbo.TblAqarDetai.Aqarid INNER JOIN"
 StrSQL = StrSQL + "  dbo.TblAkarUnit ON dbo.TblAqarDetai.unittype = dbo.TblAkarUnit.id RIGHT OUTER JOIN"
 StrSQL = StrSQL + "    dbo.TblBranchesData ON dbo.TblAqar.BranchId = dbo.TblBranchesData.branch_id"




'StrSQL = "SELECT     dbo.TblBranchesData.branch_id, dbo.TblBranchesData.branch_name, MAX(dbo.TblAqar.aqarname) AS aqarname, dbo.TblAkarUnit.name, "
'StrSQL = StrSQL + " COUNT(dbo.TblAqar.aqarname) As Unit_Count"
'StrSQL = StrSQL + " FROM         dbo.TblAqar INNER JOIN"
'StrSQL = StrSQL + " dbo.TblAqarDetai ON dbo.TblAqar.Aqarid = dbo.TblAqarDetai.Aqarid INNER JOIN "
'StrSQL = StrSQL + " dbo.TblAkarUnit ON dbo.TblAqarDetai.unittype = dbo.TblAkarUnit.id RIGHT OUTER JOIN "
'StrSQL = StrSQL + "dbo.TblBranchesData ON dbo.TblAqar.BranchId = dbo.TblBranchesData.branch_id "

StrSQL = StrSQL & " Where  (1 = 1)"
If Me.DcbBranch.BoundText <> "" Then
    StrSQL = StrSQL & " AND   dbo.TblBranchesData.branch_id = " & val(Me.DcbBranch.BoundText)
'gr = 0
End If
'StrSQL = StrSQL + " GROUP BY dbo.TblBranchesData.branch_id, dbo.TblBranchesData.branch_name, dbo.TblAkarUnit.name"
StrSQL = StrSQL + " GROUP BY dbo.TblBranchesData.branch_id, dbo.TblBranchesData.branch_name, dbo.TblAkarUnit.name"
End If



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
    
If Me.opt_4.value = True Then
        If SystemOptions.UserInterface = ArabicInterface Then
        
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepToltalCotctService.rpt"
            
       End If
End If

If opt_1.value = True Then
        If SystemOptions.UserInterface = ArabicInterface Then
        
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rpt_aqarCountReport.rpt"
            
       End If
    End If
    
     
If opt_2.value = True Then
        If SystemOptions.UserInterface = ArabicInterface Then
        
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rpt_RepCountReport.rpt"
            
       End If
End If
    
If opt_3.value = True Then
        If SystemOptions.UserInterface = ArabicInterface Then
        
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rpt_AqarUnitCountReport.rpt"
            
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
       'If XPDtbFrom.value <> Null Or XPDtbFrom.value <> "" Then xReport.ParameterFields(3).AddCurrentValue Format(Me.XPDtbFrom.value, "yyyy/M/d")
       ' If XPDtpTo.value <> Null Or XPDtpTo.value <> "" Then xReport.ParameterFields(4).AddCurrentValue Format(Me.XPDtpTo.value, "yyyy/M/d")
       ' If XPDtbFrom.value <> Null Or XPDtbFrom.value <> "" Then xReport.ParameterFields(5).AddCurrentValue Me.Txt_DateHigriFrom.value
       ' If XPDtpTo.value <> Null Or XPDtpTo.value <> "" Then xReport.ParameterFields(6).AddCurrentValue Me.Txt_DateHigriTO.value
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
        'xReport.ParameterFields(3).AddCurrentValue Format(Me.XPDtbFrom.value, "yyyy/M/d")
        'xReport.ParameterFields(4).AddCurrentValue Format(Me.XPDtpTo.value, "yyyy/M/d")
       'xReport.ParameterFields(5).AddCurrentValue Me.Txt_DateHigriFrom.value
       ' xReport.ParameterFields(6).AddCurrentValue Me.Txt_DateHigriTO.value
    End If

   ' xReport.ParameterFields(3).AddCurrentValue user_name
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
  Dim Total As String
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
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault


 
  
 
End Function

Private Sub dcsupplier_Click(Area As Integer)
   If val(dcsupplier.BoundText) = 0 Then Exit Sub

    Dim EmpCode  As String
 
    GetTblCustemersCode , , dcsupplier.BoundText, EmpCode
    Me.Text1.Text = EmpCode
End Sub
'Public Function GetBranchIDFromCode(Optional brancHcode As String, _
'Optional ByRef Emp_id As Integer) ' As Integer
'
'    Dim sql As String
'    Dim rs As New ADODB.Recordset
'    Dim Balance As Double
'    Dim id As Integer
'
'
'
'    sql = "select * from TblBranchesData where branch_code= '" & brancHcode & "'"
'
'
'    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    If rs.RecordCount > 0 Then
'        id = IIf(IsNull(rs("branch_Id").value), 0, rs("branch_Id").value)
'    Else
'        id = 0
'    End If
'
'    rs.Close
'    Emp_id = id
    'GetBranchIDFromCode = id

Private Sub ImgFavorites_Click()
AddTofaforites Me.Name, Me.Caption, Me.Caption
End Sub

Private Sub opt_1_Click()
If Me.opt_4.value = True Then

Me.Frame1.Enabled = True
Else

Me.Frame1.Enabled = False
End If
End Sub

Private Sub opt_2_Click()
If Me.opt_4.value = True Then

Me.Frame1.Enabled = True
Else

Me.Frame1.Enabled = False
End If
End Sub

Private Sub opt_3_Click()
If Me.opt_4.value = True Then

Frame1.Enabled = True
Else
Frame1.Enabled = False
End If
End Sub

Private Sub opt_4_Click()
If Me.opt_4.value = True Then
Frame1.Enabled = True

Else

Me.Frame1.Enabled = False
End If
End Sub

Private Sub RdAll_Click()
If Me.RdAll.value = vbChecked Then
Me.RdyExtranl.value = vbUnchecked
Me.RdEntry.value = vbUnchecked
End If
End Sub

Private Sub RdEntry_Click()
If Me.RdEntry.value = vbChecked Then
Me.RdAll.value = vbUnchecked
Me.RdyExtranl.value = vbUnchecked
End If
End Sub

Private Sub RdyExtranl_Click()
If Me.RdyExtranl.value = vbChecked Then
Me.RdAll.value = vbUnchecked
Me.RdEntry.value = vbUnchecked
End If
End Sub

'End Function

Private Sub Text1_KeyPress(KeyAscii As Integer)
 Dim EmpID As Integer

  If KeyAscii = vbKeyReturn Then
        GetTblCustemersCode Text1.Text, EmpID, , , 57
        dcsupplier.BoundText = EmpID
   End If
   
End Sub
 
