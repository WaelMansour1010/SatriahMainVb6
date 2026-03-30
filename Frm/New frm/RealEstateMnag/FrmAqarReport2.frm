VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmExpiredContract 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10425
   Icon            =   "FrmAqarReport2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   10425
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton btnClear 
      Caption         =   "„”Õ"
      Height          =   495
      Left            =   2760
      TabIndex        =   26
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Frame XPPnlTime 
      BackColor       =   &H00E2E9E9&
      Caption         =   "›Ï «·› —…"
      Height          =   1185
      Left            =   4320
      TabIndex        =   15
      Top             =   6720
      Visible         =   0   'False
      Width           =   2415
      Begin MSComCtl2.DTPicker XPDtbFrom 
         Height          =   345
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   609
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   62390273
         CurrentDate     =   38784
      End
      Begin MSComCtl2.DTPicker XPDtpTo 
         Height          =   345
         Left            =   120
         TabIndex        =   17
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   609
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   62390273
         CurrentDate     =   38784
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   285
         Index           =   2
         Left            =   1560
         TabIndex        =   19
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
         TabIndex        =   18
         Top             =   720
         Width           =   465
      End
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   " ﬁ«—Ì— «·⁄ﬁÊœ «·„‰ ÂÌ…"
      Height          =   2805
      Left            =   0
      TabIndex        =   3
      Top             =   720
      Width           =   10395
      Begin VB.Frame Frame3 
         Height          =   2655
         Left            =   6960
         TabIndex        =   24
         Top             =   120
         Width           =   3375
         Begin VB.Image Image1 
            Height          =   1710
            Left            =   480
            Picture         =   "FrmAqarReport2.frx":038A
            Stretch         =   -1  'True
            Top             =   120
            Width           =   2460
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
            Left            =   240
            TabIndex        =   25
            Top             =   1920
            Width           =   2895
         End
      End
      Begin VB.TextBox txtCodeBranch 
         Height          =   285
         Left            =   6360
         TabIndex        =   23
         Top             =   4680
         Width           =   855
      End
      Begin VB.TextBox txtCodeOwner 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4320
         TabIndex        =   22
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox txtCodeSalesRep 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4320
         TabIndex        =   21
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   5400
         TabIndex        =   20
         Top             =   4680
         Width           =   855
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
      Begin MSDataListLib.DataCombo dcsupplier 
         Height          =   315
         Left            =   240
         TabIndex        =   7
         Top             =   1200
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
         Top             =   840
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo dcbSalesSpec 
         Height          =   315
         Left            =   240
         TabIndex        =   11
         Top             =   1560
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DCAkarUnit 
         Height          =   315
         Left            =   240
         TabIndex        =   14
         Top             =   1920
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
         Caption         =   "ÿ»ﬁ« ·‰Ê⁄ «·ÊÕœ…"
         Height          =   195
         Index           =   9
         Left            =   5400
         TabIndex        =   13
         Top             =   1920
         Width           =   1140
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÿ»ﬁ« ·„‰œÊ» „Õœœ"
         Height          =   195
         Index           =   4
         Left            =   5400
         TabIndex        =   12
         Top             =   1560
         Width           =   1185
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
         Top             =   1200
         Width           =   1065
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
         Top             =   840
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
         Top             =   480
         Width           =   1020
      End
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   495
      Index           =   0
      Left            =   1440
      TabIndex        =   0
      Top             =   3600
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
      Top             =   3600
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
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "‘«‘…  ﬁ«—Ì— «·⁄ﬁÊœ «·„‰ ÂÌ…"
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
      Width           =   10485
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
Attribute VB_Name = "FrmExpiredContract"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RsSavRec As ADODB.Recordset
Dim BKGrndPic As ClsBackGroundPic
Dim RecID As String
Dim II As Long
Dim cSearch  As clsDCboSearch
Dim rs As ADODB.Recordset
Dim DCboSearch As clsDCboSearch


Private Sub btnClear_Click()
clear_all Me
End Sub

Private Sub Cmd_Click(Index As Integer)

    Select Case Index

        Case 0
       

 GetData
            
        Case 1
            clear_all Me

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

Private Sub dcsupplier_Click(Area As Integer)
   If val(dcsupplier.BoundText) = 0 Then Exit Sub
    Dim EmpCode  As String
    GetTblCustemersCode , , dcsupplier.BoundText, EmpCode
    Me.txtCodeOwner.Text = EmpCode
End Sub
Private Sub Form_Activate()
   PutFormOnTop Me.hWnd
End Sub


Private Sub Form_Load()
   
 'On Error GoTo ErrTrap
    Dim i As Integer
    Dim My_SQL As String
    Dim Dcombos As ClsDataCombos
    
    Set Dcombos = New ClsDataCombos
    
    Dcombos.GetIqar dcbAqarType
    
    
    Dcombos.GetCustomersSuppliers 57, Me.dcsupplier
    
    Dcombos.getAkarUnit Me.DCAkarUnit
    
    Dcombos.GetSalesRepData Me.dcbSalesSpec
    
    
    Dcombos.GetBranches DcbBranch
    
    'Dcombos.GetRentStatus dbcAqarStatus
    
    
    
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

StrSQL = "SELECT   dbo.TblContract.ContNo ,  dbo.TblContract.NoteSerial1, dbo.TblContract.TodateH, dbo.TblContract.EndDate, dbo.TblContract.OldRent, dbo.TblContract.CusID,  dbo.TblCustemers.CusName AS CustName, dbo.TblCustemers.CusNamee AS CustNamee, dbo.TblContract.ownerid, dbo.TblContract.UnitType, dbo.TblAqar.BranchId, dbo.TblContract.Iqar , dbo.TblContract.Emp_id , dbo.TblContract.PayAmini + dbo.TblContract.TotalContract AS Total_Value  FROM         dbo.TblContract INNER JOIN  dbo.TblAqar ON dbo.TblContract.Iqar = dbo.TblAqar.Aqarid LEFT OUTER JOIN  dbo.TblCustemers ON dbo.TblContract.CusID = dbo.TblCustemers.CusID "





StrSQL = StrSQL & " Where  (1 = 1)"
    BolBegine = False
    StrWhere = ""
    
    
If Me.DcbBranch.BoundText <> "" Then
StrWhere = StrWhere & " AND TblAqar.BranchId = " & val(Me.DcbBranch.BoundText)
'gr = 0
End If


If Me.dcbAqarType.BoundText <> "" Then
'gr = 1
StrWhere = StrWhere & " AND Iqar = " & val(Me.dcbAqarType.BoundText)
'gr = 1
End If


If Me.dcsupplier.BoundText <> "" Then
StrWhere = StrWhere & " AND dbo.TblContract.ownerid = " & val(dcsupplier.BoundText)
'gr = 2
End If

If Me.dcbSalesSpec.BoundText <> "" Then
StrWhere = StrWhere & " AND Emp_ID = " & val(dcbSalesSpec.BoundText)
'gr = 2
End If



If Me.DCAkarUnit.BoundText <> "" Then
StrWhere = StrWhere & " AND  UnitType = " & val(DCAkarUnit.BoundText)
'gr = 2
End If


StrWhere = StrWhere & " and EndDate <=" & SQLDate(Date, True)



    '-----------------------------------

    StrSQL = StrSQL & StrWhere
 
  
  
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



        If SystemOptions.UserInterface = ArabicInterface Then
        
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rpt_ExpiredAqar.rpt"
            
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
    xReport.ApplicationName = App.title
    xReport.ReportAuthor = App.title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault


 
  
 
End Function


Private Sub Text1_Change()

End Sub

Private Sub txtCodeBranch_KeyPress(KeyAscii As Integer)
 Dim EmpID As Integer
    If KeyAscii = vbKeyReturn Then
       GetBranchIDFromCode txtCodeBranch.Text, EmpID
       DcbBranch.BoundText = EmpID
    End If
End Sub

 

Private Sub txtCodeOwner_KeyPress(KeyAscii As Integer)
 Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetTblCustemersCode txtCodeOwner.Text, EmpID, , , 57
        dcsupplier.BoundText = EmpID
    End If
End Sub

Private Sub txtCodeSalesRep_KeyPress(KeyAscii As Integer)
Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode txtCodeSalesRep.Text, EmpID
        dcbSalesSpec.BoundText = EmpID
    End If

End Sub


 

Public Function GetBranchIDFromCode(Optional brancHcode As String, _
Optional ByRef Emp_id As Integer) ' As Integer
            
    Dim sql As String
    Dim rs As New ADODB.Recordset
    Dim Balance As Double
    Dim id As Integer
    

    
    sql = "select * from TblBranchesData where branch_code= '" & brancHcode & "'"
   
 
    rs.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        id = IIf(IsNull(rs("branch_Id").value), 0, rs("branch_Id").value)
    Else
        id = 0
    End If

    rs.Close
    Emp_id = id
    'GetBranchIDFromCode = id

End Function


















