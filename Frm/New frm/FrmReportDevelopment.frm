VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SuiteCtrls.ocx"
Begin VB.Form FrmReportDevelopment 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " ﬁ«—Ì— «·„Â«„ Ê «·⁄„·Ì«   "
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6975
   Icon            =   "FrmReportDevelopment.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton btnClear 
      BackColor       =   &H00E2E9E9&
      Caption         =   "„”Õ"
      Height          =   495
      Left            =   2640
      TabIndex        =   22
      Top             =   6360
      Width           =   1335
   End
   Begin VB.Frame XPPnlTime 
      BackColor       =   &H00E2E9E9&
      Caption         =   "›Ï «·› —…"
      Height          =   1185
      Left            =   4320
      TabIndex        =   6
      Top             =   7440
      Visible         =   0   'False
      Width           =   2415
      Begin MSComCtl2.DTPicker XPDtbFrom 
         Height          =   345
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   609
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   159645697
         CurrentDate     =   38784
      End
      Begin MSComCtl2.DTPicker XPDtpTo 
         Height          =   345
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   609
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   159645697
         CurrentDate     =   38784
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   285
         Index           =   2
         Left            =   1560
         TabIndex        =   10
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
         TabIndex        =   9
         Top             =   720
         Width           =   465
      End
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Height          =   5565
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   720
      Width           =   10395
      Begin VB.Frame Frame2 
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·›—“"
         Height          =   795
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   2760
         Width           =   6735
         Begin XtremeSuiteControls.RadioButton Opt 
            Height          =   375
            Index           =   0
            Left            =   4320
            TabIndex        =   34
            Top             =   240
            Width           =   1815
            _Version        =   786432
            _ExtentX        =   3201
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "»«·⁄„·Ì…"
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton Opt 
            Height          =   375
            Index           =   1
            Left            =   2280
            TabIndex        =   35
            Top             =   240
            Width           =   1815
            _Version        =   786432
            _ExtentX        =   3201
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "»„œÌ— «·⁄„·Ì…"
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton Opt 
            Height          =   375
            Index           =   2
            Left            =   120
            TabIndex        =   36
            Top             =   240
            Width           =   1815
            _Version        =   786432
            _ExtentX        =   3201
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "»«· «—ÌŒ"
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
      End
      Begin VB.TextBox Text3 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4320
         TabIndex        =   24
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4320
         TabIndex        =   23
         Top             =   960
         Width           =   855
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00E2E9E9&
         Caption         =   "Õœœ «·› —…"
         Height          =   720
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   2040
         Width           =   6795
         Begin MSComCtl2.DTPicker Fromdate 
            Height          =   330
            Left            =   3495
            TabIndex        =   16
            Top             =   240
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   159645697
            CurrentDate     =   41640
         End
         Begin MSComCtl2.DTPicker toDate 
            Height          =   330
            Left            =   240
            TabIndex        =   17
            Top             =   240
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   159711233
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
            TabIndex        =   19
            Top             =   240
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
            TabIndex        =   18
            Top             =   240
            Width           =   540
         End
      End
      Begin VB.Frame Frame3 
         Height          =   5415
         Left            =   6960
         TabIndex        =   13
         Top             =   120
         Width           =   3375
         Begin VB.Image Image1 
            Height          =   2310
            Left            =   120
            Picture         =   "FrmReportDevelopment.frx":038A
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
            TabIndex        =   14
            Top             =   2520
            Width           =   2895
         End
      End
      Begin VB.TextBox txtCodeBranch 
         Height          =   285
         Left            =   6360
         TabIndex        =   12
         Top             =   5760
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   5400
         TabIndex        =   11
         Top             =   6000
         Width           =   855
      End
      Begin MSDataListLib.DataCombo DcbBranch 
         Height          =   315
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcbManager 
         Height          =   315
         Left            =   240
         TabIndex        =   25
         Top             =   600
         Width           =   4035
         _ExtentX        =   7117
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcboEmpName 
         Height          =   315
         Left            =   240
         TabIndex        =   26
         Top             =   960
         Width           =   4035
         _ExtentX        =   7117
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcbTypeVisit1 
         Height          =   315
         Left            =   240
         TabIndex        =   29
         Top             =   1320
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcbDes 
         Height          =   315
         Left            =   240
         TabIndex        =   30
         Top             =   1680
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Shape Shape1 
         BorderWidth     =   2
         Height          =   1815
         Left            =   120
         Top             =   3720
         Width           =   6735
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
         Height          =   1530
         Index           =   8
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   3840
         Width           =   6615
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰Ê⁄ «·⁄„·Ì…"
         Height          =   285
         Index           =   7
         Left            =   5640
         TabIndex        =   32
         Top             =   1320
         Width           =   1125
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·»‰œ"
         Height          =   285
         Index           =   19
         Left            =   5760
         TabIndex        =   31
         Top             =   1680
         Width           =   1005
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„œÌ— «·⁄„·Ì…"
         Height          =   285
         Index           =   6
         Left            =   5670
         TabIndex        =   28
         Top             =   960
         Width           =   1125
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·„”ƒÊ·"
         Height          =   285
         Index           =   5
         Left            =   5790
         TabIndex        =   27
         Top             =   600
         Width           =   1005
      End
      Begin VB.Shape Shape2 
         BorderWidth     =   2
         Height          =   495
         Left            =   0
         Top             =   6000
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
         Height          =   450
         Index           =   4
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   6240
         Width           =   6975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÿ»ﬁ« ·›—⁄ „Õœ"
         Height          =   195
         Index           =   0
         Left            =   5910
         TabIndex        =   4
         Top             =   240
         Width           =   945
      End
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   495
      Index           =   2
      Left            =   120
      TabIndex        =   0
      Top             =   6360
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
   Begin ImpulseButton.ISButton Cmd 
      Height          =   495
      Index           =   1
      Left            =   1320
      TabIndex        =   21
      Top             =   6360
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
   Begin VB.Image ImgFavorites 
      Height          =   390
      Left            =   1920
      Picture         =   "FrmReportDevelopment.frx":10A48
      Stretch         =   -1  'True
      Top             =   120
      Width           =   525
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   " ﬁ«—Ì— «·„Â«„ Ê «·⁄„·Ì«   "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   6975
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      ForeColor       =   &H00000080&
      Height          =   285
      Index           =   1
      Left            =   60
      TabIndex        =   1
      Top             =   3060
      Width           =   1785
   End
End
Attribute VB_Name = "FrmReportDevelopment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim amoutId As Integer
Dim RsSavRec As ADODB.Recordset
Dim BKGrndPic As ClsBackGroundPic
Dim RecID As String
Dim II As Long
Dim cSearch  As clsDCboSearch
Dim rs As ADODB.Recordset
Dim DCboSearch As clsDCboSearch
Public indexx As Integer

Private Sub btnClear_Click()
Cmd_Click (7)
End Sub

Private Sub ChangeLang()
    Dim XPic As IPictureDisp
    'Set XPic = Me.XPBtnMove(1).ButtonImage
    'Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    'Set Me.XPBtnMove(2).ButtonImage = XPic
    'Set XPic = Me.XPBtnMove(0).ButtonImage
    'Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    'Set Me.XPBtnMove(3).ButtonImage = XPic
  '  Label1.Visible = False

Label5.Caption = "Report of Development"
Label1(0).Caption = "Brand"
lbl(5).Caption = "Salesman"
   lbl(6).Caption = "Manger"
  lbl(7).Caption = "Process"
  lbl(19).Caption = "Des"
 Frame8.Caption = "Priod"
 lbl(3).Caption = "From"
 lbl(14).Caption = "To"
 Frame2.Caption = "Sort By"
 Opt(0).Caption = "Process"
 Opt(1).Caption = "Manger Pr."
 Opt(2).Caption = "Date "
 lbl(8).Caption = ""
 btnClear.Caption = "Clear"
 Cmd(1).Caption = "Show"
 Cmd(2).Caption = "Exit"
End Sub

Private Sub Cmd_Click(Index As Integer)

    Select Case Index

        Case 0
       indexx = 1
 GetData
            
        Case 1
   
       indexx = 0
 GetData
  Case 7
  clear_all Me
  Fromdate.value = ""
    todate.value = ""

        Case 2
        
            Unload Me
            Case 3
'print_report
    End Select

End Sub




Private Sub DcbManager_Change()
DcbManager_Click (0)
End Sub

Private Sub DcbManager_Click(Area As Integer)
 If val(DcbManager.BoundText) = 0 Then Exit Sub
 Dim EmpCode  As String

    GetEmployeeIDFromCode , , DcbManager.BoundText, EmpCode
    Me.Text3.Text = EmpCode
End Sub

Private Sub DcboEmpName_Change()
DcboEmpName_Click (0)
End Sub

Private Sub DcboEmpName_Click(Area As Integer)
 If val(DcboEmpName.BoundText) = 0 Then Exit Sub
 Dim EmpCode  As String

    GetEmployeeIDFromCode , , DcboEmpName.BoundText, EmpCode
    Me.Text1.Text = EmpCode
End Sub

Private Sub DcbTypeVisit1_Click(Area As Integer)
If val(DcbTypeVisit1.BoundText) <> 0 Then
    Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos
    Dcombos.GetDevelopProcessPand Me.DcbDes, val(DcbTypeVisit1.BoundText)
    End If
End Sub

Private Sub Form_Activate()
   PutFormOnTop Me.hWnd
End Sub




Private Sub Form_Load()
   
 'On Error GoTo ErrTrap
    Dim I As Integer
    Dim My_SQL As String
    Dim Dcombos As ClsDataCombos
    
    Set Dcombos = New ClsDataCombos
    Dcombos.GetBranches DcbBranch
    Dcombos.GetEmployees DcbManager
    Dcombos.GetEmployees DcboEmpName
    Dcombos.GetDevelopProcess Me.DcbTypeVisit1
    
    Fromdate.value = ""
    todate.value = ""
               If SystemOptions.UserInterface = EnglishInterface Then
         
        SetInterface Me
        ChangeLang
        Else
      
    End If
    
    Set cSearch = New clsDCboSearch
    Set BKGrndPic = New ClsBackGroundPic
    Set RsSavRec = New ADODB.Recordset
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
    Dim I As Integer
    'gr = 9
    'Order = 9

 StrSQL = "SELECT     dbo.TblRegDevelopment.Id, dbo.TblRegDevelopment.RecordDate, dbo.TblRegDevelopment.StrDate, dbo.TblRegDevelopment.EndExptedDate, "
 StrSQL = StrSQL & "                   dbo.TblRegDevelopment.EndActDate, dbo.TblRegDevelopment.Important, dbo.TblRegDevelopment.MoDay, dbo.TblRegDevelopment.DesOp,"
 StrSQL = StrSQL & "                     dbo.TblRegDevelopment.AnlysOp, dbo.TblRegDevelopment.TimeReq, dbo.TblRegDevelopment.StartTime, dbo.TblRegDevelopment.StatusProcess,"
 StrSQL = StrSQL & "                     dbo.TblRegDevelopment.StatusPand, dbo.TblRegDevelopment.NoDaySatart, dbo.TblRegDevelopment.NoDayEnd, dbo.TblRegDevelopment.FromDate,"
 StrSQL = StrSQL & "                     dbo.TblRegDevelopment.ToDate, dbo.TblRegDevelopment.RecordTime, dbo.TblRegDevelopment.BranchID, dbo.TblBranchesData.branch_name,"
 StrSQL = StrSQL & "                     dbo.TblBranchesData.branch_namee, dbo.TblRegDevelopment.EmpID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee,"
 StrSQL = StrSQL & "                     dbo.TblRegDevelopment.MangID, TblEmployee_1.Emp_Name AS MangEmp_Name, TblEmployee_1.Fullcode AS MangFullcode,"
 StrSQL = StrSQL & "                     TblEmployee_1.Emp_Namee AS MangEmp_NameE, dbo.TblRegDevelopment.OpType, dbo.TblProceeDevelper.Name, dbo.TblProceeDevelper.NameE,"
 StrSQL = StrSQL & "                     dbo.TblRegDevelopment.DesID , dbo.TblProceeDevelperDet.des"
 StrSQL = StrSQL & " FROM         dbo.TblRegDevelopment LEFT OUTER JOIN"
 StrSQL = StrSQL & "                     dbo.TblProceeDevelperDet ON dbo.TblRegDevelopment.DesID = dbo.TblProceeDevelperDet.ID LEFT OUTER JOIN"
 StrSQL = StrSQL & "                     dbo.TblProceeDevelper ON dbo.TblRegDevelopment.OpType = dbo.TblProceeDevelper.ID LEFT OUTER JOIN"
 StrSQL = StrSQL & "                     dbo.TblEmployee TblEmployee_1 ON dbo.TblRegDevelopment.MangID = TblEmployee_1.Emp_ID LEFT OUTER JOIN"
 StrSQL = StrSQL & "                     dbo.TblEmployee ON dbo.TblRegDevelopment.EmpID = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
 StrSQL = StrSQL & "                     dbo.TblBranchesData ON dbo.TblRegDevelopment.BranchID = dbo.TblBranchesData.branch_id"
 StrSQL = StrSQL & " WHERE  (1=1)  "

    BolBegine = False
    StrWhere = ""
If val(Me.DcbBranch.BoundText) <> 0 And Me.DcbBranch.Text <> "" Then
StrWhere = StrWhere & " AND dbo.TblRegDevelopment.BranchID = " & val(Me.DcbBranch.BoundText)

End If


If val(Me.DcbManager.BoundText) <> 0 And Me.DcbManager.Text <> "" Then

StrWhere = StrWhere & " AND dbo.TblRegDevelopment.MangID = " & val(Me.DcbManager.BoundText)

End If

If val(Me.DcboEmpName.BoundText) <> 0 And Me.DcboEmpName.Text <> "" Then
StrWhere = StrWhere & " AND dbo.TblRegDevelopment.EmpID    = " & val(DcboEmpName.BoundText)

End If

If val(Me.DcbTypeVisit1.BoundText) <> 0 And Me.DcbTypeVisit1.Text <> "" Then
StrWhere = StrWhere & " AND dbo.TblRegDevelopment.OpType  = " & val(DcbTypeVisit1.BoundText)

End If
If val(Me.DcbDes.BoundText) <> 0 And Me.DcbDes.Text <> "" Then
StrWhere = StrWhere & " AND dbo.TblRegDevelopment.DesID  = " & val(DcbDes.BoundText)

End If


   If Not IsNull(Me.Fromdate.value) Then
                   StrWhere = StrWhere & " AND dbo.TblRegDevelopment.RecordDate >=" & SQLDate(Me.Fromdate.value, True) & ""
      End If

    If Not IsNull(Me.todate.value) Then
            StrWhere = StrWhere & " AND  dbo.TblRegDevelopment.RecordDate <=" & SQLDate(Me.todate.value, True) & ""
     
    End If




    '-----------------------------------

    StrSQL = StrSQL & StrWhere
 If Opt(1).value = True Then

 StrSQL = StrSQL & " order by   dbo.TblRegDevelopment.EmpID  "
 ElseIf Opt(0).value = True Then
 StrSQL = StrSQL & " order by  dbo.TblRegDevelopment.OpType "
 Else
  StrSQL = StrSQL & " order by  dbo.TblRegDevelopment.RecordDate "
  End If
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
             Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «· ﬁ—Ì—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            Msg = "There's no data to show that matches the specified conditions"
        End If
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    Else
  
 rs.MoveFirst

 print_report StrSQL

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



        If SystemOptions.UserInterface = ArabicInterface Then
        
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "ReportDevelopmnet.rpt"
            Else
             StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "ReportDevelopmnetE.rpt"
            
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
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        Else
            Msg = "There's no data to show"
        End If
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
    Dim MSGType As Integer
   
    If SystemOptions.UserInterface = ArabicInterface Then
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
        ' xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
        StrReportTitle = "" '& StrAccountName
   
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        'xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""
     
    End If

   ' xReport.ParameterFields(3).AddCurrentValue user_name
   If Fromdate.value <> "" And todate.value <> "" Then
    xReport.ParameterFields(14).AddCurrentValue Fromdate.value
       
       xReport.ParameterFields(16).AddCurrentValue todate.value
      
       End If
       
       'val(lbl(23).Caption)
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

Private Sub ImgFavorites_Click()
AddTofaforites Me.Name, Me.Caption, Me.Caption
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode Text1.Text, EmpID
        DcboEmpName.BoundText = EmpID
    End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode Text3.Text, EmpID
        DcbManager.BoundText = EmpID
    End If
End Sub
