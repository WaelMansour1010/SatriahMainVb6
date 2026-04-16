VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SUITEC~1.OCX"
Begin VB.Form FixedAssetReportsEmp 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " ﬁ«—Ì— «·⁄Âœ ⁄‰œ «·„ÊŸ›Ì‰"
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10260
   Icon            =   "FrmAsestReport.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   3345
   ScaleWidth      =   10260
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
   Begin XtremeSuiteControls.RadioButton RdByName 
      Height          =   375
      Left            =   8280
      TabIndex        =   14
      Top             =   2760
      Width           =   1815
      _Version        =   786432
      _ExtentX        =   3201
      _ExtentY        =   661
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
      Height          =   2085
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   480
      Width           =   10155
      Begin VB.Frame Frame1 
         BackColor       =   &H00E2E9E9&
         Caption         =   "Frame1"
         Height          =   735
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   0
         Width           =   3015
         Begin XtremeSuiteControls.RadioButton Opt 
            Height          =   375
            Index           =   0
            Left            =   1800
            TabIndex        =   30
            Top             =   240
            Width           =   855
            _Version        =   786432
            _ExtentX        =   1508
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "⁄ÂœÂ"
            ForeColor       =   16711680
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton Opt 
            Height          =   375
            Index           =   1
            Left            =   480
            TabIndex        =   31
            Top             =   240
            Width           =   855
            _Version        =   786432
            _ExtentX        =   1508
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "„” Êœ⁄"
            ForeColor       =   16711680
            UseVisualStyle  =   -1  'True
            TextAlignment   =   1
            RightToLeft     =   -1  'True
         End
      End
      Begin VB.CheckBox ChkStatus 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈ŸÂ«— ﬂ· «·„ÊŸ›Ì‰ „⁄ «·„‰ ÂÌ… Œœ„« Â„"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   720
         Width           =   3255
      End
      Begin VB.TextBox TxtQuntity 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   1080
         Width           =   3855
      End
      Begin VB.TextBox TxtSearchCode 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   7320
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   240
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker DtpDateFrom 
         Height          =   330
         Left            =   7560
         TabIndex        =   8
         Top             =   1680
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   183697411
         CurrentDate     =   38887
      End
      Begin MSComCtl2.DTPicker DtpDateTo 
         Height          =   330
         Left            =   5040
         TabIndex        =   10
         Top             =   1680
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   183697411
         CurrentDate     =   38887
      End
      Begin MSDataListLib.DataCombo DcboEmpName 
         Height          =   315
         Left            =   3240
         TabIndex        =   19
         Top             =   240
         Width           =   4035
         _ExtentX        =   7117
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo dcmboassest 
         Height          =   315
         Left            =   5040
         TabIndex        =   20
         Top             =   720
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton RdAll 
         Height          =   375
         Left            =   3120
         TabIndex        =   23
         Top             =   1560
         Width           =   855
         _Version        =   786432
         _ExtentX        =   1508
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "ﬂ· «·⁄Âœ"
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton RdMove 
         Height          =   375
         Left            =   -120
         TabIndex        =   24
         Top             =   1560
         Width           =   1335
         _Version        =   786432
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "«·⁄Âœ «·„‰ﬁÊ·…"
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton RdFound 
         Height          =   375
         Left            =   1560
         TabIndex        =   25
         Top             =   1560
         Width           =   1335
         _Version        =   786432
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "«·⁄Âœ «·„ÊÃÊœ…"
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DCboStoreName 
         Height          =   315
         Left            =   5040
         TabIndex        =   26
         Top             =   1080
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "7"
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·„Œ“‰"
         Height          =   195
         Index           =   24
         Left            =   8775
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   1095
         Width           =   1140
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Height          =   195
         Left            =   4770
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   360
         Width           =   45
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·ﬂ„ÌÂ"
         Height          =   195
         Index           =   0
         Left            =   3960
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   1170
         Width           =   585
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï  «—ÌŒ"
         Height          =   195
         Index           =   3
         Left            =   6750
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   1650
         Width           =   705
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰  «—ÌŒ"
         Height          =   195
         Index           =   4
         Left            =   9090
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   1650
         Width           =   765
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·⁄Âœ…"
         Height          =   195
         Index           =   7
         Left            =   8775
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   720
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
      Left            =   2880
      TabIndex        =   0
      Top             =   2760
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
      Left            =   1410
      TabIndex        =   1
      Top             =   2760
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
      Left            =   30
      TabIndex        =   2
      Top             =   2760
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
   Begin XtremeSuiteControls.RadioButton RdByAssest 
      Height          =   375
      Left            =   5760
      TabIndex        =   15
      Top             =   2760
      Width           =   1815
      _Version        =   786432
      _ExtentX        =   3201
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "›—“ »Õ”» «·⁄Âœ…"
      UseVisualStyle  =   -1  'True
      TextAlignment   =   1
      RightToLeft     =   -1  'True
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      BackStyle       =   0  'Transparent
      Caption         =   " ﬁ«—Ì— «·⁄Âœ ⁄‰œ «·„ÊŸ›Ì‰"
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
      Left            =   4260
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   120
      Width           =   3405
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      ForeColor       =   &H00000080&
      Height          =   285
      Index           =   1
      Left            =   60
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   3060
      Width           =   1785
   End
End
Attribute VB_Name = "FixedAssetReportsEmp"
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
            If Opt(0).value = True Then
                GetData
            ElseIf Opt(1).value = True Then
                GetData1
            Else
                If SystemOptions.UserInterface = ArabicInterface Then
                    MsgBox "Ì—ÃÏ ≈Œ Ì«— ‰Ê⁄ «· ﬁ—Ì—"
                Else
                    MsgBox "Please Select Type of Reports"
                End If
            End If
        Case 1
            clear_all Me
DtpDateFrom.value = ""
DtpDateTo.value = ""
Me.RdAll.value = False
Me.RdByAssest.value = False
Me.RdByName.value = False
Me.RdFound.value = False
Me.RdMove.value = False
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





Private Sub DcboEmpName_Change()
DcboEmpName_Click (0)
End Sub

Private Sub DcboEmpName_Click(Area As Integer)
    If val(DcboEmpName.BoundText) = 0 Then Exit Sub

    Dim EmpCode  As String
 
    GetEmployeeIDFromCode , , DcboEmpName.BoundText, EmpCode
    TxtSearchCode.Text = EmpCode
End Sub

Private Sub Form_Activate()
   PutFormOnTop Me.hWnd
End Sub
Private Sub ChangeLang()

 
ChkStatus.Caption = "All Employees With End Service"
 
 Cmd(1).Caption = "Clear"
    Cmd(0).Caption = "View Report"
   Cmd(2).Caption = "Exit"
  Me.Caption = "Reports of Assest Of Employee "
Label5.Caption = Me.Caption
Label1.Caption = "Emp"

Opt(0).RightToLeft = False
Opt(1).RightToLeft = False
Opt(0).Caption = "Assest"
Opt(1).Caption = "Store"
lbl(24).Caption = "Store"

lbl(7).Caption = "Assest "
lbl(0).Caption = "Qty"


Me.RdByAssest.RightToLeft = False
RdByAssest.Caption = "By Assest"
Me.RdAll.RightToLeft = False
Me.RdAll.Caption = "All"
Me.RdByName.RightToLeft = False
Me.RdByName.Caption = "By EmpName"
Me.RdFound.RightToLeft = False
Me.RdFound.Caption = "Existing"
Me.RdMove.RightToLeft = False
Me.RdMove.Caption = "Pushers"

lbl(3).Caption = "To Date"
lbl(4).Caption = "From Date"
End Sub

Private Sub Form_Load()
    'Dim GrdBack As ClsBackGroundPic
    Dim Dcombos As ClsDataCombos
    
        Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500

dcmboassest.Enabled = False
DCboStoreName.Enabled = True

    Set Dcombos = New ClsDataCombos
   ' Dcombos.GetEmployees Me.DCEmp_Name
  Set Dcombos = New ClsDataCombos
     'Dcombos.GetUsers Me.DCboUserName
     Dcombos.GetEmployees Me.DcboEmpName
    
     Dcombos.GetStores Me.DCboStoreName
    
   Dcombos.GetAssests Me.dcmboassest
   'Dcombos.GetEmpJobsTypes Me.DcmbToJob
   
   'Dcombos.GetEmpLocations Me.dcmbToProject ' locatione
    Set DCboSearch = New clsDCboSearch
  '  Set DCboSearch.Client = Me.DcbClientname
    'Dcombos.GetUsers Me.DCUser
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Clear").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture


 
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
Sub GetData()
    Dim StrSQL As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim I As Integer

StrSQL = " SELECT     dbo.TblAssestes.AsName, dbo.TblAssestes.AsID, dbo.TblAssestes.AsDes, dbo.TblEmpAsestDetails.Remarks, dbo.TblEmpAsestDetails.Qunt,"
 StrSQL = StrSQL & "                      dbo.TblEmpAsestDetails.FlagAs, dbo.TblEmpAsest.EmpAsID, dbo.TblEmpAsestDetails.IDAseset, dbo.TblEmpAsest.EmpAsestID, dbo.TblEmpAsest.PostedDate,"
StrSQL = StrSQL & "                       dbo.TblEmpAsest.RecordDate, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name1, dbo.TblEmployee.Emp_Name2,"
 StrSQL = StrSQL & "                      dbo.TblEmployee.Emp_Name3, dbo.TblEmployee.Emp_Name4, dbo.TblEmployee.Nationality, dbo.TblEmployee.dean, dbo.TblEmployee.Emp_Namee,"
StrSQL = StrSQL & "                       dbo.TblEmployee.Emp_Namee1, dbo.TblEmployee.Emp_Namee2, dbo.TblEmployee.Emp_Namee3, dbo.TblEmployee.Emp_Namee4, dbo.TblEmployee.Fullcode,"
StrSQL = StrSQL & "                       dbo.TblEmpAsestDetails.Remark2 , dbo.TblEmpAsestDetails.EmpID , dbo.TblEmployee.jopstatusid "
StrSQL = StrSQL & "  FROM         dbo.TblAssestes INNER JOIN"
StrSQL = StrSQL & "                       dbo.TblEmpAsestDetails ON dbo.TblAssestes.AsID = dbo.TblEmpAsestDetails.AsID INNER JOIN"
StrSQL = StrSQL & "                       dbo.TblEmpAsest ON dbo.TblEmpAsestDetails.IDAseset = dbo.TblEmpAsest.EmpAsID LEFT OUTER JOIN"
 StrSQL = StrSQL & "                      dbo.TblEmployee ON dbo.TblEmpAsestDetails.EmpID = dbo.TblEmployee.Emp_ID"
StrSQL = StrSQL & " WHERE     (1 = 1) "
    BolBegine = False
    StrWhere = ""
    
   If ChkStatus.value = vbUnchecked Then
 StrSQL = StrSQL & " and dbo.TblEmployee.jopstatusid <> 2"
 StrSQL = StrSQL & " and dbo.TblEmployee.jopstatusid <> 5"
 StrSQL = StrSQL & " and dbo.TblEmployee.jopstatusid <> 6"
 End If
If Me.RdAll.value = False And Me.RdMove.value = False Then
StrWhere = StrWhere & " AND  (dbo.TblEmpAsestDetails.FlagAs IS NULL)  "
End If
If Me.RdFound.value = True Then
StrWhere = StrWhere & " AND  (dbo.TblEmpAsestDetails.FlagAs IS NULL)  "
End If
If Me.RdMove.value = True Then
StrWhere = StrWhere & " AND dbo.TblEmpAsestDetails.FlagAs =1 "
End If

 If (Me.TxtSearchCode.Text) <> "" Then
        
            StrWhere = StrWhere & " AND dbo.TblEmployee.Fullcode like '%" & Me.TxtSearchCode.Text & "%'"
        
    End If
    If (Me.TxtQuntity.Text) <> "" Then
        
            StrWhere = StrWhere & " AND dbo.TblEmpAsestDetails.Qunt like '%" & Me.TxtQuntity.Text & "%'"
        
    End If
   If Me.DcboEmpName.BoundText <> "" Then
     
            StrWhere = StrWhere & " AND dbo.TblEmpAsestDetails.EmpID=" & Me.DcboEmpName.BoundText & ""
      
    End If
    If Me.dcmboassest.BoundText <> "" Then
    
            StrWhere = StrWhere & " AND dbo.TblEmpAsestDetails.AsID=" & Me.dcmboassest.BoundText & ""
    
    End If
   

    If Not IsNull(Me.DtpDateFrom.value) Then
                   StrWhere = StrWhere & " AND dbo.TblEmpAsest.RecordDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
      End If

    If Not IsNull(Me.DtpDateTo.value) Then
            StrWhere = StrWhere & " AND  dbo.TblEmpAsest.RecordDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
     
    End If


    

    '-----------------------------------
StrSQL = StrSQL & StrWhere
 
   StrSQL = StrSQL & " Order By dbo.TblAssestes.AsID"
  
  
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
           ' Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
          '  Me.lbl(10).Caption = "Search Results=0"
        End If
If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «· ﬁ—Ì—"
        Else
        Msg = "Not Found Data TO Show"
        
        End If
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    Else

 rs.MoveFirst

 print_report StrSQL


            If SystemOptions.UserInterface = ArabicInterface Then
             '   Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
               ' Me.lbl(10).Caption = "Search Results=" & rs.RecordCount
            End If

           
 

    End If

End Sub
Sub GetData1()
    Dim StrSQL As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim I As Integer

StrSQL = "SELECT     dbo.Transactions.Emp_ID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee, dbo.TblItems.ItemID, "
StrSQL = StrSQL & "                      SUM(dbo.TransactionTypes.StockEffect * dbo.Transaction_Details.ShowQty) AS Qty, dbo.TblItems.ItemName, dbo.TblItems.ItemCode, dbo.TblItems.ItemNamee,"
StrSQL = StrSQL & "                      dbo.TblStore.STORENAME , dbo.TblStore.StoreNamee , dbo.Transactions.StoreID"
StrSQL = StrSQL & " FROM         dbo.TransactionTypes RIGHT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.Transactions LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblStore ON dbo.Transactions.StoreID = dbo.TblStore.StoreID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblEmployee ON dbo.Transactions.Emp_ID = dbo.TblEmployee.Emp_ID ON"
StrSQL = StrSQL & "                      dbo.TransactionTypes.Transaction_Type = dbo.Transactions.Transaction_Type LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblItems RIGHT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.Transaction_Details ON dbo.TblItems.ItemID = dbo.Transaction_Details.Item_ID ON"
StrSQL = StrSQL & "                      dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID"
StrSQL = StrSQL & " WHERE     ((dbo.Transactions.Transaction_Type = 19) OR"
StrSQL = StrSQL & "                      (dbo.Transactions.Transaction_Type = 20))"
StrSQL = StrSQL & " AND (NOT (dbo.Transactions.Emp_ID IS NULL)) "
    BolBegine = False
    StrWhere = ""
    



 If (Me.TxtSearchCode.Text) <> "" Then
        
            StrWhere = StrWhere & " AND dbo.TblEmployee.Fullcode like '%" & Me.TxtSearchCode.Text & "%'"
        
    End If
    If (Me.TxtQuntity.Text) <> "" Then
        
            StrWhere = StrWhere & " AND Qty like '%" & Me.TxtQuntity.Text & "%'"
        
    End If
   If Me.DcboEmpName.Text <> "" And val(Me.DcboEmpName.BoundText) <> 0 Then
     
            StrWhere = StrWhere & " AND dbo.Transactions.Emp_ID=" & Me.DcboEmpName.BoundText & ""
      
    End If
       If Me.DCboStoreName.Text <> "" And val(Me.DCboStoreName.BoundText) <> 0 Then
     
            StrWhere = StrWhere & " AND dbo.Transactions.StoreID=" & Me.DCboStoreName.BoundText & ""
      
    End If
    
    If Me.dcmboassest.BoundText <> "" Then
    
            StrWhere = StrWhere & " AND dbo.Transactions.Emp_ID=" & Me.dcmboassest.BoundText & ""
    
    End If
   

    If Not IsNull(Me.DtpDateFrom.value) Then
                   StrWhere = StrWhere & " AND dbo.Transactions.Transaction_Date >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
      End If

    If Not IsNull(Me.DtpDateTo.value) Then
            StrWhere = StrWhere & " AND  dbo.Transactions.Transaction_Date <=" & SQLDate(Me.DtpDateTo.value, True) & ""
     
    End If


    

    '-----------------------------------
StrSQL = StrSQL & StrWhere
 
   StrSQL = StrSQL & " GROUP BY dbo.Transactions.Emp_ID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee, dbo.TblItems.ItemID, dbo.TblItems.ItemName, "
   StrSQL = StrSQL & "                   dbo.TblItems.ItemCode, dbo.TblItems.ItemNamee, dbo.TblStore.StoreName, dbo.TblStore.StoreAdress, dbo.TblStore.StoreNamee , dbo.Transactions.StoreID"
  
  
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
           ' Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
          '  Me.lbl(10).Caption = "Search Results=0"
        End If
If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «· ﬁ—Ì—"
        Else
        Msg = "Not Found Data TO Show"
        
        End If
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    Else

 rs.MoveFirst

 print_report1 StrSQL


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
        If Me.RdByName.value = True Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAssestByEmp.rpt"
            Else
            If Me.RdByAssest.value = True Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAssestByAsest.rpt"
            Else
           
            
             StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "repAsestAll.rpt"
            
         
            
            
            End If
             End If
        Else
               If Me.RdByName.value = True Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAssestByEmp.rpt"
            Else
            If Me.RdByAssest.value = True Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAssestByAsest.rpt"
            Else
          
            
             StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "repAsestAll.rpt"
        
         
            
            
            End If
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

  Dim Total As String
  Dim totl As Double

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
Function print_report1(Optional NoteSerial As String)
     
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
           StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAssestStore.rpt"
        Else
           StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAssestStoreE.rpt"
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
      '  xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""
        'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        '    StrReportTitle = StrReportTitle + " From Date " & (Me.DTPickerAccFrom.value) & ""
        'End If
        'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        '    StrReportTitle = StrReportTitle + " To Date :  " & (Me.DTPickerAccTo.value) & ""
        'End If
    End If
    If Not (IsNull(DtpDateFrom.value)) Then
    xReport.ParameterFields(4).AddCurrentValue DtpDateFrom.value
    End If
    If Not (IsNull(DtpDateTo.value)) Then
    xReport.ParameterFields(5).AddCurrentValue DtpDateTo.value
    End If
  Dim Total As String
  Dim totl As Double

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


 
Private Sub Opt_Click(Index As Integer)
If Index = 0 Then
dcmboassest.Enabled = True
DCboStoreName.Enabled = True
ElseIf Index = 1 Then
DCboStoreName.Enabled = True
dcmboassest.Enabled = False
End If
End Sub

Private Sub TxtSearchCode_KeyPress(KeyAscii As Integer)
Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode TxtSearchCode.Text, EmpID
        DcboEmpName.BoundText = EmpID
    End If
End Sub
