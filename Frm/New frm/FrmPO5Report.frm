VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SUITEC~1.OCX"
Begin VB.Form FrmPO5Report 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10425
   Icon            =   "FrmPO5Report.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   10425
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
      BackColor       =   &H00E2E9E9&
      Caption         =   "„”Õ"
      Height          =   495
      Left            =   2760
      TabIndex        =   14
      Top             =   5880
      Width           =   1335
   End
   Begin VB.Frame XPPnlTime 
      BackColor       =   &H00E2E9E9&
      Caption         =   "›Ï «·› —…"
      Height          =   1185
      Left            =   4320
      TabIndex        =   5
      Top             =   6720
      Visible         =   0   'False
      Width           =   2415
      Begin MSComCtl2.DTPicker XPDtbFrom 
         Height          =   345
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   609
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   69926913
         CurrentDate     =   38784
      End
      Begin MSComCtl2.DTPicker XPDtpTo 
         Height          =   345
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   609
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   69926913
         CurrentDate     =   38784
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   285
         Index           =   2
         Left            =   1560
         TabIndex        =   9
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
         TabIndex        =   8
         Top             =   720
         Width           =   465
      End
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   "»Ì«‰«  "
      Height          =   5205
      Left            =   0
      TabIndex        =   3
      Top             =   720
      Width           =   10395
      Begin VB.TextBox TxtCodeAother 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4320
         TabIndex        =   31
         Top             =   960
         Width           =   1065
      End
      Begin VB.TextBox TxtEmployeeID 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4320
         TabIndex        =   27
         Top             =   615
         Width           =   1065
      End
      Begin VB.TextBox TxtSearchCode 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4320
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   240
         Width           =   1065
      End
      Begin XtremeSuiteControls.CheckBox ChApproved 
         Height          =   255
         Left            =   5160
         TabIndex        =   23
         Top             =   1320
         Width           =   1455
         _Version        =   786432
         _ExtentX        =   2566
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "„⁄ „œ"
         ForeColor       =   -2147483635
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00E2E9E9&
         Caption         =   "Õœœ «·› —…"
         Height          =   1080
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   1800
         Width           =   6555
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
            Format          =   69926913
            CurrentDate     =   41640
         End
         Begin Dynamic_Byte.NourHijriCal Fromdateh 
            Height          =   330
            Left            =   3480
            TabIndex        =   17
            Top             =   600
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   582
         End
         Begin Dynamic_Byte.NourHijriCal todateH 
            Height          =   330
            Left            =   240
            TabIndex        =   18
            Top             =   600
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   582
         End
         Begin MSComCtl2.DTPicker toDate 
            Height          =   330
            Left            =   240
            TabIndex        =   19
            Top             =   240
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   69926913
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
            TabIndex        =   21
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
            TabIndex        =   20
            Top             =   480
            Width           =   540
         End
      End
      Begin VB.Frame Frame3 
         Height          =   5055
         Left            =   6960
         TabIndex        =   12
         Top             =   120
         Width           =   3375
         Begin VB.Image Image1 
            Height          =   2310
            Left            =   120
            Picture         =   "FrmPO5Report.frx":038A
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
            TabIndex        =   13
            Top             =   2520
            Width           =   2895
         End
      End
      Begin VB.TextBox txtCodeBranch 
         Height          =   285
         Left            =   6360
         TabIndex        =   11
         Top             =   5280
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   5400
         TabIndex        =   10
         Top             =   5520
         Width           =   855
      End
      Begin MSDataListLib.DataCombo DBCboClientName 
         Height          =   315
         Left            =   240
         TabIndex        =   25
         Top             =   240
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "6"
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcboEmp 
         Height          =   315
         Left            =   240
         TabIndex        =   28
         Top             =   600
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "7"
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo Dcbiteem 
         Height          =   315
         Left            =   240
         TabIndex        =   32
         Top             =   960
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox ChNotApproved 
         Height          =   255
         Left            =   3480
         TabIndex        =   33
         Top             =   1320
         Width           =   1455
         _Version        =   786432
         _ExtentX        =   2566
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "€Ì— „⁄ „œ"
         ForeColor       =   -2147483635
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "‘«‘…  ﬁ«—Ì— ⁄—Ê÷ «·«”⁄«—   "
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
         Height          =   2055
         Index           =   17
         Left            =   120
         TabIndex        =   34
         Top             =   3000
         Width           =   6645
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«”„ «·’‰›"
         Height          =   345
         Index           =   5
         Left            =   5655
         TabIndex        =   30
         Top             =   990
         Width           =   930
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«”„ «·„‰œÊ»"
         Height          =   345
         Index           =   32
         Left            =   5655
         TabIndex        =   29
         Top             =   630
         Width           =   930
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "   «·„Ê—œ"
         Height          =   285
         Index           =   7
         Left            =   5565
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   240
         Width           =   1095
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
         TabIndex        =   22
         Top             =   6240
         Width           =   6975
      End
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   495
      Index           =   0
      Left            =   1440
      TabIndex        =   0
      Top             =   5880
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
      Top             =   5880
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
      Caption         =   "‘«‘…  ﬁ«—Ì— ⁄—Ê÷ «·«”⁄«—"
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
      Left            =   -45
      TabIndex        =   4
      Top             =   0
      Width           =   10500
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
Attribute VB_Name = "FrmPO5Report"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim amoutId As Integer
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



Private Sub ChApproved_Click()
If ChApproved = xtpChecked Then
Me.ChNotApproved.value = xtpUnchecked
End If

End Sub

Private Sub ChNotApproved_Click()
If ChNotApproved = xtpChecked Then
Me.ChApproved.value = xtpUnchecked
End If
End Sub

Private Sub Cmd_Click(Index As Integer)

    Select Case Index

        Case 0
       

 GetData
            
        Case 1
   
            clear_all Me
            Me.ChApproved.value = vbUnchecked
             Me.ChNotApproved.value = vbUnchecked
         FromDate.value = ""
    todate.value = ""
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



Private Sub DBCboClientName_Change()
  If val(DBCboClientName.BoundText) = 0 Then Exit Sub

    Dim EmpCode  As String
 
    GetTblCustemersCode , , DBCboClientName.BoundText, EmpCode
    Me.TxtSearchCode.Text = EmpCode
End Sub

Private Sub DBCboClientName_Click(Area As Integer)
DBCboClientName_Change
End Sub

Private Sub DcboEmp_Change()
 'If Me.TxtModFlg = "N" Or Me.TxtModFlg = "E" Then
         If val(Me.DcboEmp.BoundText) = 0 Then Exit Sub
           Me.TxtEmployeeID.Text = get_EMPLOYEE_Data(val(Me.DcboEmp.BoundText), "Fullcode")
        'DCEmP.text = DCEmP.text
'End If
End Sub

Private Sub DcboEmp_Click(Area As Integer)
DcboEmp_Change
End Sub

Private Sub Form_Activate()
   PutFormOnTop Me.hwnd
End Sub

Private Sub DcboEmp_KeyUp(KeyCode As Integer, _
                          Shift As Integer)

    If KeyCode = vbKeyF5 Then
        Dim Dcombos As ClsDataCombos
        
        Set Dcombos = New ClsDataCombos
        Dcombos.GetSalesRepData Me.DcboEmp

    End If

End Sub

Private Sub TxtCodeAother_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Then
        If TxtCodeAother.Text = "" Then
            Me.Dcbiteem.BoundText = ""
        Else
            Me.Dcbiteem.BoundText = GetItemID(Trim$(Me.TxtCodeAother.Text))
        End If
    End If
End Sub


Private Sub Dcbiteem_Change()
     Me.TxtCodeAother.Text = GetItemCode(val(Me.Dcbiteem.BoundText))
End Sub

Private Sub Dcbiteem_Click(Area As Integer)
 Dcbiteem_Change
End Sub

Private Sub Form_Load()
   
 'On Error GoTo ErrTrap
    Dim i As Integer
    Dim My_SQL As String
    Dim Dcombos As ClsDataCombos
    
    Set Dcombos = New ClsDataCombos

     Dcombos.GetSalesRepDatapurchase Me.DcboEmp
     Dcombos.GetCustomersSuppliers 2, Me.DBCboClientName, True
 
    Dcombos.GetItemsNames Me.Dcbiteem
    FromDate.value = ""
    todate.value = ""
    Cmd_Click (1)
    
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


StrSQL = " SELECT     dbo.TblCustemers.CusName AS CuCusName, dbo.TblCustemers.CusNamee AS CuCusNameE, dbo.TblCustemers.Fullcode AS CusFullcode, "
      StrSQL = StrSQL & "                 dbo.Transaction_Details.Item_ID AS Item_IDD, dbo.TblItems.Fullcode AS ItemFullcode, dbo.TblItems.ItemCode AS IItemCode, dbo.TblItems.ItemName AS IItemName,"
     StrSQL = StrSQL & "                  dbo.TblItems.ItemNamee AS IItemNameE, dbo.Transactions.*, dbo.ApprovalData.ApprovDate AS ApprovDateD,"
     StrSQL = StrSQL & "                  dbo.Transactions.Transaction_Date AS Transaction_Dated, dbo.Transaction_Details.ShowQty AS ShowQty, dbo.Transaction_Details.showPrice AS showPrice,"
     StrSQL = StrSQL & "                  dbo.Transaction_Details.Quantity AS Quantity, dbo.Transaction_Details.ItemCase AS ItemCase, dbo.Transaction_Details.ItemSerial AS ItemSerial,"
    StrSQL = StrSQL & "                   dbo.Transaction_Details.ItemDiscountType AS ItemDiscountType, dbo.Transaction_Details.ItemDiscount AS ItemDiscount, dbo.Transaction_Details.ClassId AS ClassId,"
    StrSQL = StrSQL & "                    dbo.Transaction_Details.ParrtNoCode AS ParrtNoCode, dbo.TblEmployee.Emp_Name AS Emp_Name, dbo.TblEmployee.Emp_Name1 AS Emp_Name1,"
    StrSQL = StrSQL & "                   dbo.TblEmployee.Emp_Name2 AS Emp_Name2, dbo.TblEmployee.Emp_Name3 AS Emp_Name3, dbo.TblEmployee.Emp_Name4 AS Emp_Name4,"
    StrSQL = StrSQL & "                   dbo.TblEmployee.Fullcode AS EmpFullcode, dbo.TblEmployee.Emp_Namee4 AS Emp_Namee4, dbo.TblEmployee.Emp_Namee3 AS Emp_Namee3,"
    StrSQL = StrSQL & "                   dbo.TblEmployee.Emp_Namee2 AS Emp_Namee2, dbo.TblEmployee.Emp_Namee1 AS Emp_Namee1, dbo.TblEmployee.Emp_Namee AS Emp_Namee"
   StrSQL = StrSQL & " FROM         dbo.ApprovalData RIGHT OUTER JOIN"
  StrSQL = StrSQL & "                     dbo.Transactions ON dbo.ApprovalData.Transaction_ID = dbo.Transactions.Transaction_ID LEFT OUTER JOIN"
   StrSQL = StrSQL & "                    dbo.TblItems RIGHT OUTER JOIN"
  StrSQL = StrSQL & "                     dbo.Transaction_Details ON dbo.TblItems.ItemID = dbo.Transaction_Details.Item_ID ON"
   StrSQL = StrSQL & "                    dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID LEFT OUTER JOIN"
  StrSQL = StrSQL & "                     dbo.TblEmployee ON dbo.Transactions.Emp_ID = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
  StrSQL = StrSQL & "                     dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID"
 StrSQL = StrSQL & " WHERE  Transaction_Type=46  "

 
 
  
    BolBegine = False
    StrWhere = ""
    
    
If val(Me.DBCboClientName.BoundText) <> 0 Or Me.DBCboClientName.Text <> "" Then
StrWhere = StrWhere & " AND dbo.Transactions.CusID  = " & val(Me.DBCboClientName.BoundText)

End If


If val(Me.DcboEmp.BoundText) <> 0 Or Me.DcboEmp.Text <> "" Then

StrWhere = StrWhere & " AND dbo.Transactions.Emp_ID  = " & val(Me.DcboEmp.BoundText)

End If

If val(Me.Dcbiteem.BoundText) <> 0 Or Me.Dcbiteem.Text <> "" Then
StrWhere = StrWhere & " AND dbo.Transaction_Details.Item_ID  = " & val(Dcbiteem.BoundText)

End If


If Me.ChNotApproved.value = vbChecked Then
StrWhere = StrWhere & " AND dbo.ApprovalData.ApprovDate  IS NULL"
End If
If Me.ChApproved.value = vbChecked Then
StrWhere = StrWhere & " AND (NOT (dbo.ApprovalData.ApprovDate IS NULL))"
End If

   If Not IsNull(Me.FromDate.value) Then
                   StrWhere = StrWhere & " AND dbo.Transactions.Transaction_Date >=" & SQLDate(Me.FromDate.value, True) & ""
      End If

    If Not IsNull(Me.todate.value) Then
            StrWhere = StrWhere & " AND  dbo.Transactions.Transaction_Date <=" & SQLDate(Me.todate.value, True) & ""
     
    End If




    '-----------------------------------

    StrSQL = StrSQL & StrWhere
 
  StrSQL = StrSQL & " order by  dbo.Transactions.Transaction_ID "
  
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
           ' Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
          '  Me.lbl(10).Caption = "Search Results=0"
        End If

        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «· ﬁ—Ì—"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.TITLE
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
        
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepPO5Report.rpt"
            Else
             StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepPO5Report.rpt"
            
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
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.TITLE
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
   If FromDate.value <> "" And todate.value <> "" Then
    xReport.ParameterFields(12).AddCurrentValue CStr(FromDate.value)
   '    xReport.ParameterFields(15).AddCurrentValue Fromdateh.value
       xReport.ParameterFields(13).AddCurrentValue CStr(todate.value)
   '    xReport.ParameterFields(17).AddCurrentValue todateH.value
       End If
       
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
    xReport.ApplicationName = App.TITLE
    xReport.ReportAuthor = App.TITLE
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault


 
  
 
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



Private Sub ToDate_Change()
If todate.value <> "" Then
   todateH.value = ToHijriDate(todate.value)
   End If
End Sub

Private Sub ToDateH_LostFocus()

 VBA.Calendar = vbCalGreg
            todate.value = ToGregorianDate(todateH.value)

End Sub






Private Sub TxtEmployeeID_Change()

    DcboEmp.BoundText = GeTEmpIDByEmpCode(TxtEmployeeID.Text, True)

End Sub

Private Sub TxtSearchCode_Change()
 Dim EmpID As Integer

  
        GetTblCustemersCode TxtSearchCode.Text, EmpID
        DBCboClientName.BoundText = EmpID
   
End Sub
