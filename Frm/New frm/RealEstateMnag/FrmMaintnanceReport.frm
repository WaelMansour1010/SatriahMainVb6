VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmMaintnanceReport 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10365
   Icon            =   "FrmMaintnanceReport.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5895
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
   Begin VB.CommandButton btnClear 
      Caption         =   "„”Õ"
      Height          =   495
      Left            =   2760
      TabIndex        =   14
      Top             =   5040
      Width           =   1335
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   " ﬁ«—Ì— «·’Ì«‰…"
      Height          =   4125
      Left            =   0
      TabIndex        =   3
      Top             =   840
      Width           =   10395
      Begin VB.CheckBox chkMaintStatus 
         Alignment       =   1  'Right Justify
         Caption         =   "„Êﬁ› ÿ·»«  «·’Ì«‰…"
         Height          =   195
         Left            =   3720
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   2460
         Width           =   2835
      End
      Begin VB.TextBox txtUnitNo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   240
         TabIndex        =   29
         Top             =   2100
         Width           =   4935
      End
      Begin VB.Frame XPPnlTime 
         BackColor       =   &H00E2E9E9&
         Caption         =   "›Ï «·› —…"
         Height          =   1185
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   2760
         Width           =   2415
         Begin MSComCtl2.DTPicker XPDtbFrom 
            Height          =   345
            Left            =   120
            TabIndex        =   21
            Top             =   360
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   609
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   251330561
            CurrentDate     =   41640
         End
         Begin MSComCtl2.DTPicker XPDtpTo 
            Height          =   345
            Left            =   120
            TabIndex        =   22
            Top             =   720
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   609
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   251330561
            CurrentDate     =   41640
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„‰"
            Height          =   285
            Index           =   2
            Left            =   1560
            TabIndex        =   24
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
            TabIndex        =   23
            Top             =   720
            Width           =   465
         End
      End
      Begin VB.Frame FrameDateH 
         Caption         =   " ÕœÌœ «· «—ÌŒ «·ÂÃ—Ì"
         Height          =   1185
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   2760
         Width           =   2220
         Begin Dynamic_Byte.NourHijriCal Txt_DateHigriFrom 
            Height          =   315
            Left            =   120
            TabIndex        =   16
            Top             =   360
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
         End
         Begin Dynamic_Byte.NourHijriCal Txt_DateHigriTO 
            Height          =   315
            Left            =   120
            TabIndex        =   17
            Top             =   720
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            Caption         =   "«·Ï"
            Height          =   255
            Left            =   1560
            RightToLeft     =   -1  'True
            TabIndex        =   19
            Top             =   720
            Width           =   495
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            Caption         =   "„‰"
            Height          =   255
            Left            =   1560
            RightToLeft     =   -1  'True
            TabIndex        =   18
            Top             =   360
            Width           =   495
         End
      End
      Begin VB.Frame Frame3 
         Height          =   3855
         Left            =   6960
         TabIndex        =   12
         Top             =   120
         Width           =   3375
         Begin VB.Image Image1 
            Height          =   2190
            Left            =   240
            Picture         =   "FrmMaintnanceReport.frx":038A
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
            TabIndex        =   13
            Top             =   2400
            Visible         =   0   'False
            Width           =   2895
         End
      End
      Begin MSDataListLib.DataCombo DcbBranch 
         Height          =   315
         Left            =   240
         TabIndex        =   4
         Top             =   300
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcboEmpName 
         Height          =   315
         Left            =   240
         TabIndex        =   7
         Top             =   660
         Width           =   4935
         _ExtentX        =   8705
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
         Top             =   1380
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DCAkarUnit 
         Height          =   315
         Left            =   240
         TabIndex        =   11
         Top             =   1740
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcboEmpNameSuper 
         Height          =   315
         Left            =   240
         TabIndex        =   31
         Top             =   1020
         Width           =   4935
         _ExtentX        =   8705
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
         Caption         =   "ÿ»ﬁ« ·—ﬁ„ «·ÊÕœ…"
         Height          =   195
         Index           =   7
         Left            =   5400
         TabIndex        =   30
         Top             =   2100
         Width           =   1125
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÿ»ﬁ« ·„ÊŸ› «·’Ì«‰…"
         Height          =   195
         Index           =   6
         Left            =   5400
         TabIndex        =   28
         Top             =   1020
         Width           =   1320
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÿ»ﬁ« ··„ÊŸ› «·„”ƒÊ· "
         Height          =   195
         Index           =   3
         Left            =   5400
         TabIndex        =   25
         Top             =   660
         Width           =   1515
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÿ»ﬁ« ·‰Ê⁄ «·ÊÕœ…"
         Height          =   195
         Index           =   9
         Left            =   5400
         TabIndex        =   10
         Top             =   1740
         Width           =   1140
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÿ»ﬁ«·⁄ﬁ«— „⁄Ì‰"
         Height          =   195
         Index           =   1
         Left            =   5400
         TabIndex        =   9
         Top             =   1380
         Width           =   1020
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
         Top             =   300
         Width           =   1020
      End
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   495
      Index           =   0
      Left            =   1440
      TabIndex        =   0
      Top             =   5040
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
      Top             =   5040
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
      TabIndex        =   26
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
      Left            =   1680
      Picture         =   "FrmMaintnanceReport.frx":10A48
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
      TabIndex        =   27
      Top             =   2040
      Width           =   1290
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "‘«‘…  ﬁ«—Ì— «·’Ì«‰…"
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
      Left            =   -90
      TabIndex        =   6
      Top             =   0
      Width           =   10425
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
Attribute VB_Name = "FrmMaintnanceReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RsSavRec As ADODB.Recordset
Dim BKGrndPic As ClsBackGroundPic
Dim RecId As String
Dim ii As Long
Dim cSearch  As clsDCboSearch
Dim rs As ADODB.Recordset
Dim DCboSearch As clsDCboSearch


Private Sub btnClear_Click()
clear_all Me
xpdtbfrom.value = ""
   XPDtpTo.value = ""
End Sub

Private Sub Cmd_Click(index As Integer)

    Select Case index

        Case 0
       
If chkMaintStatus.value = vbChecked Then
    GetData2
Else
    GetData
End If
            
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






Private Sub Form_Activate()
   PutFormOnTop Me.hWnd
End Sub


Private Sub ChangeLang()
 
  '
End Sub

Private Sub Form_Load()
   
 'On Error GoTo ErrTrap
    Dim i As Integer
    Dim My_SQL As String
    Dim Dcombos As ClsDataCombos
    
   xpdtbfrom.value = ""
   XPDtpTo.value = ""
    
    Set Dcombos = New ClsDataCombos
    
    Dcombos.GetIqar dcbAqarType
    
    
    'Dcombos.GetCustomersSuppliers 1, Me.dcsupplier
    
    Dcombos.getAkarUnit Me.DCAkarUnit
    
    'Dcombos.GetSalesRepData Me.dcbSalesSpec
    
    
    Dcombos.GetBranches DcbBranch
    
    'Dcombos.GetRentStatus dbcAqarStatus
    
   Dcombos.GetEmployees Me.DcboEmpName
    Dcombos.GetEmployees Me.DcboEmpNameSuper
    
    
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

StrSQL = " SELECT     TOP 100 PERCENT dbo.TblOrderMaintenance.BranchID, dbo.TblOrderMaintenance.RecDate, dbo.TblOrderMaintenance.RecDateH, dbo.TblAqar.Aqarid, "
StrSQL = StrSQL & "                      dbo.TblOrderMaintenance.SuperVM, dbo.TblBranchesData.branch_name, TblEmployee_1.Emp_Name AS SuperVM_Name, dbo.TblAqar.aqarname,"
StrSQL = StrSQL & "                      dbo.TblAkarUnit.name, dbo.TblOrderMaintenanceDet.TypeUnit, TblEmployee_2.Emp_Name, dbo.TblOrderMaintenance.EmpID, dbo.TblOrderMaintenance.Lock,"
StrSQL = StrSQL & "                      dbo.TblCustemers.CusID, dbo.TblCustemers.CusName, dbo.TblOrderMaintenance.ID, dbo.TblOrderMaintenance.LockDateH, dbo.TblOrderMaintenance.LockDate,"
StrSQL = StrSQL & "                      dbo.TblOrderMaintenance.EndFate, DATEDIFF(Day, dbo.TblOrderMaintenance.EndFate, dbo.TblOrderMaintenance.LockDate) AS DayLate,"
StrSQL = StrSQL & "                      dbo.TblOrderMaintenance.TimOrder, dbo.TblOrderMaintenance.LocationIqar, dbo.TblOrderMaintenance.Des, dbo.TblOrderMaintenance.DMY,"
StrSQL = StrSQL & "                      dbo.TblOrderMaintenance.Cont, dbo.TblOrderMaintenance.EndFateH, dbo.TblOrderMaintenanceDet.Mobile, dbo.TblOrderMaintenanceDet.Ms,"
StrSQL = StrSQL & "                      dbo.TblOrderMaintenanceDet.UnitStatus, dbo.TblOrderMaintenanceDet.RenterID, dbo.TblOrderMaintenanceDet.UnitNo, dbo.TblAqarDetai.unitno AS unitnoName"
StrSQL = StrSQL & " FROM         dbo.TblCustemers RIGHT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblAqarDetai RIGHT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblOrderMaintenanceDet ON dbo.TblAqarDetai.Id = dbo.TblOrderMaintenanceDet.UnitNo ON"
StrSQL = StrSQL & "                      dbo.TblCustemers.CusID = dbo.TblOrderMaintenanceDet.RenterID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblAkarUnit ON dbo.TblOrderMaintenanceDet.TypeUnit = dbo.TblAkarUnit.id RIGHT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblBranchesData RIGHT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblOrderMaintenance ON dbo.TblBranchesData.branch_id = dbo.TblOrderMaintenance.BranchID ON"
StrSQL = StrSQL & "                      dbo.TblOrderMaintenanceDet.ORderID = dbo.TblOrderMaintenance.ID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblEmployee TblEmployee_1 ON dbo.TblOrderMaintenance.SuperVM = TblEmployee_1.Emp_ID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblEmployee TblEmployee_2 ON dbo.TblOrderMaintenance.EmpID = TblEmployee_2.Emp_ID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblAqar ON dbo.TblOrderMaintenance.AqrID = dbo.TblAqar.Aqarid"


StrSQL = StrSQL & " Where  (1 = 1)"
    BolBegine = False
    StrWhere = ""
    
    
If Me.DcbBranch.BoundText <> "" Then
StrWhere = StrWhere & " AND TblOrderMaintenance.BranchID = " & val(Me.DcbBranch.BoundText)
'gr = 0
End If


If Me.dcbAqarType.BoundText <> "" Then
'gr = 1
StrWhere = StrWhere & " AND dbo.TblAqar.Aqarid = " & val(Me.dcbAqarType.BoundText)
'gr = 1
End If


If Me.DCAkarUnit.BoundText <> "" Then
StrWhere = StrWhere & " AND  TblOrderMaintenanceDet.TypeUnit = " & val(DCAkarUnit.BoundText)
'gr = 2
End If

If Me.txtUnitNo.text <> "" Then
StrWhere = StrWhere & " AND  TblOrderMaintenanceDet.unitno = " & val(txtUnitNo.text)
'gr = 2
End If

If Me.DcboEmpNameSuper.BoundText <> "" Then
StrWhere = StrWhere & " AND   dbo.TblOrderMaintenance.SuperVM = " & val(DcboEmpNameSuper.BoundText)
'gr = 2
End If

If Me.DcboEmpName.BoundText <> "" Then
StrWhere = StrWhere & " AND  dbo.TblOrderMaintenance.EmpID = " & val(DcboEmpName.BoundText)
'gr = 2
End If




 If Me.xpdtbfrom <> Empty Or Me.xpdtbfrom <> Null Then
        StrWhere = StrWhere + " and (RecDate >=" & SQLDate(Me.xpdtbfrom.value, True) & ")"
    End If

    If Me.XPDtpTo <> Empty Or Me.XPDtpTo <> Null Then
        StrWhere = StrWhere + " and (RecDate <=" & SQLDate(XPDtpTo.value, True) & ")"
    End If


    '-----------------------------------



    StrSQL = StrSQL & StrWhere
 
  StrSQL = StrSQL + " ORDER BY dbo.TblOrderMaintenance.Lock"
  
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

Public Sub GetData2()
    Dim StrSQL As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
    'gr = 9
    'Order = 9

StrSQL = " SELECT     TOP 100 PERCENT dbo.TblOrderMaintenance.BranchID, dbo.TblOrderMaintenance.RecDate, dbo.TblOrderMaintenance.RecDateH, dbo.TblAqar.Aqarid, "
StrSQL = StrSQL & "                      dbo.TblOrderMaintenance.SuperVM, dbo.TblBranchesData.branch_name, TblEmployee_1.Emp_Name AS SuperVM_Name, dbo.TblAqar.aqarname,"
StrSQL = StrSQL & "                      dbo.TblAkarUnit.name, dbo.TblOrderMaintenanceDet.TypeUnit, TblEmployee_2.Emp_Name, dbo.TblOrderMaintenance.EmpID, dbo.TblOrderMaintenance.Lock,"
StrSQL = StrSQL & "                      dbo.TblCustemers.CusID, dbo.TblCustemers.CusName, dbo.TblOrderMaintenance.ID, dbo.TblOrderMaintenance.LockDateH, dbo.TblOrderMaintenance.LockDate,"
StrSQL = StrSQL & "                      dbo.TblOrderMaintenance.EndFate, DATEDIFF(Day, dbo.TblOrderMaintenance.EndFate, dbo.TblOrderMaintenance.LockDate) AS DayLate,"
StrSQL = StrSQL & "                      dbo.TblOrderMaintenance.TimOrder, dbo.TblOrderMaintenance.LocationIqar, dbo.TblOrderMaintenance.Des, dbo.TblOrderMaintenance.DMY,"
StrSQL = StrSQL & "                      dbo.TblOrderMaintenance.Cont, dbo.TblOrderMaintenance.EndFateH, dbo.TblOrderMaintenanceDet.Mobile, dbo.TblOrderMaintenanceDet.Ms,"
StrSQL = StrSQL & "                      dbo.TblOrderMaintenanceDet.UnitStatus, dbo.TblOrderMaintenanceDet.RenterID, dbo.TblOrderMaintenanceDet.UnitNo, dbo.TblAqarDetai.unitno AS unitnoName"
StrSQL = StrSQL & " FROM         dbo.TblCustemers RIGHT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblAqarDetai RIGHT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblOrderMaintenanceDet ON dbo.TblAqarDetai.Id = dbo.TblOrderMaintenanceDet.UnitNo ON"
StrSQL = StrSQL & "                      dbo.TblCustemers.CusID = dbo.TblOrderMaintenanceDet.RenterID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblAkarUnit ON dbo.TblOrderMaintenanceDet.TypeUnit = dbo.TblAkarUnit.id RIGHT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblBranchesData RIGHT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblOrderMaintenance ON dbo.TblBranchesData.branch_id = dbo.TblOrderMaintenance.BranchID ON"
StrSQL = StrSQL & "                      dbo.TblOrderMaintenanceDet.ORderID = dbo.TblOrderMaintenance.ID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblEmployee TblEmployee_1 ON dbo.TblOrderMaintenance.SuperVM = TblEmployee_1.Emp_ID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblEmployee TblEmployee_2 ON dbo.TblOrderMaintenance.EmpID = TblEmployee_2.Emp_ID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblAqar ON dbo.TblOrderMaintenance.AqrID = dbo.TblAqar.Aqarid"




StrSQL = "SELECT"

StrSQL = StrSQL & "             TOP 100 PERCENT            TblEmployee.Emp_ID"
StrSQL = StrSQL & "                         ,TblEmployee.Emp_Name"
StrSQL = StrSQL & "                         ,TblEmployee.Emp_Namee"
StrSQL = StrSQL & "                         ,TblOrderMaintenance.Mobile"
StrSQL = StrSQL & "                         ,ACCOUNTS.Account_Code"
StrSQL = StrSQL & "                         ,ACCOUNTS.Account_Name"
StrSQL = StrSQL & "                         ,ACCOUNTS.Account_NameEng"
StrSQL = StrSQL & "                         ,TblAqar.aqarname"
StrSQL = StrSQL & "                         ,TblCustemers.CusID"
StrSQL = StrSQL & "                         ,TblCustemers.CusName"
StrSQL = StrSQL & "                         ,TblCustemers.CusNamee"
StrSQL = StrSQL & "                         ,notes_all.OrderMaintenanceId"
StrSQL = StrSQL & "                         ,TblOrderMaintenance.RecDate"
StrSQL = StrSQL & "                         ,TblExpensesDet.value"
StrSQL = StrSQL & "                         ,TblExpensesDet.PriceTotal"
StrSQL = StrSQL & "                         ,TblExpensesDet.Vat"
StrSQL = StrSQL & "                         ,TblExpensesDet.Vatyo"
StrSQL = StrSQL & "                         ,TblExpensesDet.vaTotalPayedlue"
StrSQL = StrSQL & "                         ,TblExpensesDet.TotalPayed"
StrSQL = StrSQL & "                      From notes_all"
StrSQL = StrSQL & "                      INNER JOIN TblExpensesDet"
StrSQL = StrSQL & "                          ON TblExpensesDet.ExpID = notes_all.NoteID"

StrSQL = StrSQL & "                      INNER JOIN TblOrderMaintenance"
StrSQL = StrSQL & "                          ON TblOrderMaintenance.ID = notes_all.OrderMaintenanceId"
StrSQL = StrSQL & "                      LEFT OUTER JOIN TblAqar"
StrSQL = StrSQL & "                          ON TblOrderMaintenance.AqrID = TblAqar.Aqarid"
StrSQL = StrSQL & "                      LEFT OUTER JOIN TblCustemers"
StrSQL = StrSQL & "                          ON TblAqar.ownerid = TblCustemers.CusID"
StrSQL = StrSQL & "                      LEFT OUTER JOIN ACCOUNTS"
StrSQL = StrSQL & "                          ON ACCOUNTS.Account_Code = TblExpensesDet.AccountCode"
StrSQL = StrSQL & "                      LEFT OUTER JOIN TblEmployee"
StrSQL = StrSQL & "                          ON TblEmployee.Emp_ID = TblOrderMaintenance.SuperVM"
    
StrSQL = StrSQL & " Where  (1 = 1)"
    BolBegine = False
    StrWhere = ""
    
    
If Me.DcbBranch.BoundText <> "" Then
StrWhere = StrWhere & " AND TblOrderMaintenance.BranchID = " & val(Me.DcbBranch.BoundText)
'gr = 0
End If


If Me.dcbAqarType.BoundText <> "" Then
'gr = 1
StrWhere = StrWhere & " AND dbo.TblAqar.Aqarid = " & val(Me.dcbAqarType.BoundText)
'gr = 1
End If


If Me.DCAkarUnit.BoundText <> "" Then
StrWhere = StrWhere & " AND  TblOrderMaintenanceDet.TypeUnit = " & val(DCAkarUnit.BoundText)
'gr = 2
End If

If Me.txtUnitNo.text <> "" Then
StrWhere = StrWhere & " AND  TblOrderMaintenanceDet.unitno = " & val(txtUnitNo.text)
'gr = 2
End If

If Me.DcboEmpNameSuper.BoundText <> "" Then
StrWhere = StrWhere & " AND   dbo.TblOrderMaintenance.SuperVM = " & val(DcboEmpNameSuper.BoundText)
'gr = 2
End If

If Me.DcboEmpName.BoundText <> "" Then
StrWhere = StrWhere & " AND  dbo.TblOrderMaintenance.EmpID = " & val(DcboEmpName.BoundText)
'gr = 2
End If




 If Me.xpdtbfrom <> Empty Or Me.xpdtbfrom <> Null Then
        StrWhere = StrWhere + " and (TblOrderMaintenance.RecDate >=" & SQLDate(Me.xpdtbfrom.value, True) & ")"
    End If

    If Me.XPDtpTo <> Empty Or Me.XPDtpTo <> Null Then
        StrWhere = StrWhere + " and (TblOrderMaintenance.RecDate <=" & SQLDate(XPDtpTo.value, True) & ")"
    End If


    '-----------------------------------



    StrSQL = StrSQL & StrWhere
 
  StrSQL = StrSQL + " ORDER BY dbo.TblOrderMaintenance.Lock"
  
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

    If chkMaintStatus.value = vbChecked Then
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rpt_MaintnanceReportStatus.rpt"
    Else
        If SystemOptions.UserInterface = ArabicInterface Then
        
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rpt_MaintnanceReport.rpt"
            
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
        StrReportTitle = "" '& StrAccountName
        'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        '    StrReportTitle = StrReportTitle + " »œ«Ì… „‰ " & Format(Me.DTPickerAccFrom.value, "yyyy/M/d") & ""
        'End If
        'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        '    StrReportTitle = StrReportTitle + " ≈·Ï " & Format(Me.DTPickerAccTo.value, "yyyy/M/d") & " "
        'End If
        If xpdtbfrom.value <> Null Or xpdtbfrom.value <> "" Then xReport.ParameterFields(3).AddCurrentValue Format(Me.xpdtbfrom.value, "yyyy/M/d")
        If XPDtpTo.value <> Null Or XPDtpTo.value <> "" Then xReport.ParameterFields(4).AddCurrentValue Format(Me.XPDtpTo.value, "yyyy/M/d")
        If xpdtbfrom.value <> Null Or xpdtbfrom.value <> "" Then xReport.ParameterFields(5).AddCurrentValue Me.Txt_DateHigriFrom.value
        If XPDtpTo.value <> Null Or XPDtpTo.value <> "" Then xReport.ParameterFields(6).AddCurrentValue Me.Txt_DateHigriTO.value
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
        xReport.ParameterFields(3).AddCurrentValue Format(Me.xpdtbfrom.value, "yyyy/M/d")
        xReport.ParameterFields(4).AddCurrentValue Format(Me.XPDtpTo.value, "yyyy/M/d")
       xReport.ParameterFields(5).AddCurrentValue Me.Txt_DateHigriFrom.value
        xReport.ParameterFields(6).AddCurrentValue Me.Txt_DateHigriTO.value
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
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault


 
  
 
End Function


Private Sub ImgFavorites_Click()
AddTofaforites Me.Name, Me.Caption, Me.Caption
End Sub

Private Sub Txt_DateHigriFrom_LostFocus()
 VBA.Calendar = vbCalGreg
            xpdtbfrom.value = ToGregorianDate(Txt_DateHigriFrom.value)
End Sub

Private Sub Txt_DateHigriTO_LostFocus()
 VBA.Calendar = vbCalGreg
            XPDtpTo.value = ToGregorianDate(Txt_DateHigriTO.value)
End Sub



Public Function GetBranchIDFromCode(Optional brancHcode As String, _
Optional ByRef Emp_id As Integer) ' As Integer
            
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim Balance As Double
    Dim ID As Integer
    

    
    sql = "select * from TblBranchesData where branch_code= '" & brancHcode & "'"
   
 
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        ID = IIf(IsNull(rs("branch_Id").value), 0, rs("branch_Id").value)
    Else
        ID = 0
    End If

    rs.Close
    Emp_id = ID
    'GetBranchIDFromCode = id

End Function



Private Sub xpdtbfrom_Change()
If xpdtbfrom.value <> Null Or xpdtbfrom.value <> "" Then
 Txt_DateHigriFrom.value = ToHijriDate(xpdtbfrom.value)
 End If
End Sub

Private Sub XPDtpTo_Change()
If XPDtpTo.value <> Null Or XPDtpTo.value <> "" Then
 Txt_DateHigriTO.value = ToHijriDate(XPDtpTo.value)
 End If
End Sub
