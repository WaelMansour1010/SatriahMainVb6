VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmStoreExchangReport 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   6195
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9945
   Icon            =   "FrmStoreExchangReport.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   9945
   ShowInTaskbar   =   0   'False
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   495
      Left            =   5880
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   6480
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
      TabIndex        =   7
      Top             =   5640
      Width           =   1335
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Height          =   4725
      Index           =   1
      Left            =   0
      TabIndex        =   3
      Top             =   840
      Width           =   9915
      Begin VB.TextBox txtEmpCode 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3750
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   2055
         Width           =   705
      End
      Begin VB.TextBox TxtSearchCode 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   3750
         TabIndex        =   28
         Top             =   1680
         Width           =   705
      End
      Begin VB.TextBox TxtStoreID 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3750
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   1320
         Width           =   705
      End
      Begin VB.OptionButton ChDetails 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " ﬁ—Ì—  ›’Ì·Ì"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   360
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   24
         ToolTipText     =   "«’€— „‰"
         Top             =   2400
         Width           =   2355
      End
      Begin VB.OptionButton ChAll 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " ﬁ—Ì— «Ã„«·Ì"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   360
         Left            =   2400
         RightToLeft     =   -1  'True
         TabIndex        =   23
         ToolTipText     =   "«’€— „‰"
         Top             =   2400
         Width           =   2835
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰ «·› —Â"
         Height          =   735
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   2760
         Width           =   4455
         Begin MSComCtl2.DTPicker DtpDateFrom 
            Height          =   330
            Left            =   2280
            TabIndex        =   12
            Top             =   270
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   104202243
            CurrentDate     =   41640
         End
         Begin MSComCtl2.DTPicker DtpDateTo 
            Height          =   330
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   104202243
            CurrentDate     =   41640
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈·Ï"
            Height          =   195
            Index           =   3
            Left            =   1710
            RightToLeft     =   -1  'True
            TabIndex        =   15
            Top             =   240
            Width           =   480
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„‰"
            Height          =   195
            Index           =   4
            Left            =   3690
            RightToLeft     =   -1  'True
            TabIndex        =   14
            Top             =   240
            Width           =   540
         End
      End
      Begin VB.Frame Frame3 
         Height          =   4575
         Left            =   5520
         TabIndex        =   5
         Top             =   120
         Width           =   4335
         Begin VB.Image Image1 
            Height          =   3675
            Left            =   120
            Picture         =   "FrmStoreExchangReport.frx":038A
            Stretch         =   -1  'True
            Top             =   120
            Width           =   4395
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
            Left            =   480
            TabIndex        =   6
            Top             =   3840
            Width           =   2895
         End
      End
      Begin MSDataListLib.DataCombo dcBranch 
         Height          =   315
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcboEmpDepartments 
         Height          =   315
         Left            =   120
         TabIndex        =   19
         Top             =   600
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DCEquipments 
         Height          =   315
         Left            =   120
         TabIndex        =   21
         Top             =   960
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "6"
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DCboStoreName 
         Height          =   315
         Left            =   120
         TabIndex        =   26
         Top             =   1320
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "7"
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DBCboClientName 
         Height          =   315
         Left            =   120
         TabIndex        =   29
         Top             =   1680
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "6"
         BoundColumn     =   ""
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcboEmpName 
         Height          =   315
         Left            =   120
         TabIndex        =   32
         Top             =   2040
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«·„ÊŸ›"
         Height          =   210
         Index           =   64
         Left            =   4320
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·⁄„Ì·"
         Height          =   210
         Index           =   7
         Left            =   4320
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   1695
         Width           =   1095
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«·„Œ“‰"
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   2
         Left            =   4320
         TabIndex        =   27
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«·„⁄œ…"
         Height          =   210
         Index           =   62
         Left            =   4320
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«·«œ«—… "
         Height          =   210
         Index           =   61
         Left            =   4320
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«·›—⁄"
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   0
         Left            =   4320
         TabIndex        =   18
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Caption         =   "‘«‘…  ﬁ«—Ì— ”‰œ«  «·’—› «·„Œ“‰Ì "
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
         Height          =   900
         Index           =   25
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   3600
         Width           =   5175
      End
      Begin VB.Shape Shape1 
         BorderWidth     =   2
         Height          =   975
         Left            =   120
         Top             =   3600
         Width           =   5295
      End
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   495
      Index           =   0
      Left            =   1440
      TabIndex        =   0
      Top             =   5640
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
      Top             =   5640
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
      TabIndex        =   8
      Top             =   2040
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
      Caption         =   "ÿ»ﬁ« ·„” √Ã— „Õœœ"
      Height          =   195
      Index           =   5
      Left            =   5400
      TabIndex        =   9
      Top             =   2040
      Width           =   1290
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "‘«‘…  ﬁ—Ì— ”‰œ«  «·’—› «·„Œ“‰Ì"
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
      Left            =   15
      TabIndex        =   4
      Top             =   0
      Width           =   9960
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
Attribute VB_Name = "FrmStoreExchangReport"
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

Private Sub DBCboClientName_Click(Area As Integer)
DBCboClientName_Change
End Sub
Private Sub DBCboClientName_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF5 Then
       LoadCombosData
    End If
End Sub

Private Sub DcboEmpDepartments_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF5 Then
       LoadCombosData
    End If
End Sub

Private Sub DcboEmpName_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF5 Then
       LoadCombosData
    End If
End Sub

Private Sub DCboStoreName_Click(Area As Integer)
DCboStoreName_Change
End Sub



Private Sub DCboStoreName_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF5 Then
       LoadCombosData
    End If
End Sub

Private Sub dcBranch_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF5 Then
       LoadCombosData
    End If
End Sub



Private Sub DCEquipments_KeyUp(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyF5 Then
       LoadCombosData
    End If
End Sub

Private Sub TxtSearchCode_KeyPress(KeyAscii As Integer)
  Dim CUSTID As Integer

    If KeyAscii = vbKeyReturn Then
        GetCustomersDetail CUSTID, , TxtSearchCode.Text, 2
        DBCboClientName.BoundText = CUSTID
    End If
End Sub
Private Sub DcboEmpName_Change()
DcboEmpName_Click (0)
End Sub

Private Sub DcboEmpName_Click(Area As Integer)
       If val(DcboEmpName.BoundText) = 0 Then Exit Sub

    Dim EmpCode  As String
 
    GetEmployeeIDFromCode , , DcboEmpName.BoundText, EmpCode
    txtEmpCode.Text = EmpCode
    
End Sub
Private Sub ChangeLang()
    Dim XPic As IPictureDisp
   ' Set XPic = Me.btnFirst.ButtonImage
   ' Set Me.btnFirst.ButtonImage = Me.btnLast.ButtonImage
   ' Set Me.btnLast.ButtonImage = XPic
   ' Set XPic = Me.btnPrevious.ButtonImage
   ' Set Me.btnPrevious.ButtonImage = Me.btnNext.ButtonImage
   ' Set Me.btnNext.ButtonImage = XPic
   lblCompanyname.Caption = "Al SATTATYAH Group"
    Label5.Caption = "Stock Exchange Bonds Reports"
    lbl(25).Caption = Label5.Caption
    lbl(0).Caption = "Branch"
    lbl(61).Caption = "Management"
    lbl(62).Caption = "Equipment"
    lbl(2).Caption = "Store"
    lbl(7).Caption = "Customer"
    lbl(64).Caption = "Employee"
    ChAll.RightToLeft = False
    ChAll.Caption = "Total Report"
    Frame1.Caption = "Period"
    lbl(3).Caption = "To"
    lbl(4).Caption = "From"
ChDetails.RightToLeft = False
ChDetails.Caption = "Analytical Report"
btnClear.Caption = "Clear"
Cmd(0).Caption = "Show Report"
Cmd(2).Caption = "Exit"

   

End Sub
Private Sub btnClear_Click()
clear_all Me

DtpDateFrom.value = ""
DtpDateTo.value = ""
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
Private Sub Form_Activate()
   PutFormOnTop Me.hwnd
End Sub
Private Sub DCboStoreName_Change()
 TxtStoreID.Text = getStoreCoding(val(DCboStoreName.BoundText))
End Sub
Private Sub TxtStoreID_KeyDown(KeyCode As Integer, Shift As Integer)
 Dim StoreId As Integer

    If KeyCode = vbKeyReturn Then
    StoreId = getStoreInformatin(TxtStoreID)
        DCboStoreName.BoundText = StoreId
    End If
End Sub
Private Sub DBCboClientName_Change()
     Dim Fullcode As String
     GetCustomersDetail val(DBCboClientName.BoundText), , Fullcode, 1
    TxtSearchCode.Text = Fullcode
End Sub
Private Sub Form_Load()
   
 'On Error GoTo ErrTrap
LoadCombosData
DtpDateFrom.value = Date
DtpDateTo.value = Date


DtpDateFrom.value = ""
DtpDateTo.value = ""
    Resize_Form Me
    If SystemOptions.UserInterface = EnglishInterface Then
    SetInterface Me
    ChangeLang
    End If

End Sub
Private Sub txtEmpCode_KeyPress(KeyAscii As Integer)
   Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode txtEmpCode.Text, EmpID
        DcboEmpName.BoundText = EmpID
    End If
    
    
End Sub


Private Sub LoadCombosData()
  Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos
     Dcombos.GetEmpDepartments Me.DcboEmpDepartments
     Dcombos.GetEquipments DCEquipments
    Dcombos.GetEmployees Me.DcboEmpName
    Dcombos.GetCustomersSuppliers 0, Me.DBCboClientName
    Dcombos.GetStores Me.DCboStoreName
    Dcombos.GetBranches Me.dcBranch
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
    If SystemOptions.UserInterface = ArabicInterface Then
     If Me.ChAll.value = False And Me.ChDetails.value = False Then
    MsgBox "Ì—ÃÏ «Œ Ì«— ‰Ê⁄ «· ﬁ—Ì— «Ã„«·Ì «Ê  ›’Ì·Ì"
    Exit Sub
    End If
    Else
       If Me.ChAll.value = False And Me.ChDetails.value = False Then
    MsgBox "Please Select Type of Report"
    Exit Sub
    End If
    End If
    
    If ChAll.value = True Or Me.ChDetails.value = True Then
If ChAll.value = True Then
StrSQL = " SELECT     dbo.Transactions.Transaction_ID, dbo.Transactions.Transaction_Serial, dbo.Transactions.Transaction_Date, dbo.Transactions.Transaction_Type, "
StrSQL = StrSQL & "                      dbo.Transactions.PaymentType, dbo.Transactions.Transaction_HijriDate, dbo.Transactions.Trans_Discount, dbo.Transactions.Trans_DiscountType,"
StrSQL = StrSQL & "                      dbo.Transactions.CusID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode, dbo.Transactions.StoreID, dbo.TblStore.StoreName,"
StrSQL = StrSQL & "                      dbo.TblStore.StoreAdress, dbo.TblStore.StorePhone, dbo.TblStore.StoreNamee, dbo.Transactions.TaxFound, dbo.Transactions.TaxValue, dbo.Transactions.BranchId,"
StrSQL = StrSQL & "                      dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.Transactions.FixesAssetsID, dbo.FixedAssets.Name, dbo.FixedAssets.namee,"
StrSQL = StrSQL & "                      dbo.Transactions.DepartementID, dbo.TblEmpDepartments.DepartmentName, dbo.TblEmpDepartments.DepartmentNamee, dbo.Transactions.CashCustomerName,"
StrSQL = StrSQL & "                      dbo.Transactions.Emp_ID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Name1, dbo.TblEmployee.Emp_Name2, dbo.TblEmployee.Emp_Name3,"
StrSQL = StrSQL & "                      dbo.TblEmployee.Emp_Name4, dbo.TblEmployee.Fullcode AS EmpFullcode, dbo.TblEmployee.Emp_Namee4, dbo.TblEmployee.Emp_Namee3,"
StrSQL = StrSQL & "                      dbo.TblEmployee.Emp_Namee2, dbo.TblEmployee.Emp_Namee1, dbo.TblEmployee.Emp_Namee, dbo.Transactions.SaleType, dbo.Transactions.CashCustomerPhone,"
StrSQL = StrSQL & "                       dbo.Transactions.CashCustomerMobile, dbo.Transactions.CashCustomerAddress, dbo.Transactions.CashCustomerComment, dbo.Transactions.TransactionComment,"
StrSQL = StrSQL & "                      dbo.Transactions.TaxAddValue , dbo.Transactions.TaxStampValue, dbo.Transactions.NoteSerial, dbo.Transactions.NoteSerial1, dbo.Transactions.netvalue"
StrSQL = StrSQL & " FROM         dbo.Transactions LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblEmployee ON dbo.Transactions.Emp_ID = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblEmpDepartments ON dbo.Transactions.DepartementID = dbo.TblEmpDepartments.DeparmentID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.FixedAssets ON dbo.Transactions.FixesAssetsID = dbo.FixedAssets.id LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblBranchesData ON dbo.Transactions.BranchId = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblStore ON dbo.Transactions.StoreID = dbo.TblStore.StoreID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID"
StrSQL = StrSQL & " Where (dbo.Transactions.Transaction_Type = 19)"
End If
If Me.ChDetails.value = True Then
StrSQL = "SELECT     dbo.Transactions.Transaction_ID, dbo.Transactions.Transaction_Serial, dbo.Transactions.Transaction_Date, dbo.Transactions.Transaction_Type, "
StrSQL = StrSQL & "                      dbo.Transactions.PaymentType, dbo.Transactions.Transaction_HijriDate, dbo.Transactions.Trans_Discount, dbo.Transactions.Trans_DiscountType,"
StrSQL = StrSQL & "                      dbo.Transactions.CusID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode, dbo.Transactions.StoreID, dbo.TblStore.StoreName,"
StrSQL = StrSQL & "                      dbo.TblStore.StoreAdress, dbo.TblStore.StorePhone, dbo.TblStore.StoreNamee, dbo.Transactions.TaxFound, dbo.Transactions.TaxValue, dbo.Transactions.BranchId,"
StrSQL = StrSQL & "                      dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.Transactions.FixesAssetsID, dbo.FixedAssets.Name, dbo.FixedAssets.namee,"
StrSQL = StrSQL & "                      dbo.Transactions.DepartementID, dbo.TblEmpDepartments.DepartmentName, dbo.TblEmpDepartments.DepartmentNamee, dbo.Transactions.CashCustomerName,"
StrSQL = StrSQL & "                      dbo.Transactions.Emp_ID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Name1, dbo.TblEmployee.Emp_Name2, dbo.TblEmployee.Emp_Name3,"
StrSQL = StrSQL & "                      dbo.TblEmployee.Emp_Name4, dbo.TblEmployee.Fullcode AS EmpFullcode, dbo.TblEmployee.Emp_Namee4, dbo.TblEmployee.Emp_Namee3,"
StrSQL = StrSQL & "                      dbo.TblEmployee.Emp_Namee2, dbo.TblEmployee.Emp_Namee1, dbo.TblEmployee.Emp_Namee, dbo.Transactions.SaleType, dbo.Transactions.CashCustomerPhone,"
StrSQL = StrSQL & "                       dbo.Transactions.CashCustomerMobile, dbo.Transactions.CashCustomerAddress, dbo.Transactions.CashCustomerComment, dbo.Transactions.TransactionComment,"
StrSQL = StrSQL & "                      dbo.Transactions.TaxAddValue, dbo.Transactions.TaxStampValue, dbo.Transactions.NoteSerial, dbo.Transactions.NoteSerial1, dbo.Transactions.NetValue,"
StrSQL = StrSQL & "                      dbo.Transaction_Details.Item_ID, dbo.TblItems.ItemName, dbo.TblItems.ItemNamee, dbo.TblItems.Fullcode AS ItemFullcode, dbo.Transaction_Details.UnitId,"
StrSQL = StrSQL & "                      dbo.TblUnites.UnitName, dbo.TblUnites.UnitNamee, dbo.Transaction_Details.Quantity, dbo.Transaction_Details.Price, dbo.Transaction_Details.ShowQty,"
If SystemOptions.HideCost = True Then
StrSQL = StrSQL & "                      dbo.Transaction_Details.ItemCase,0 as showPrice, dbo.Transaction_Details.ItemSerial, dbo.Transaction_Details.ItemDiscountType,"
Else
StrSQL = StrSQL & "                      dbo.Transaction_Details.ItemCase, dbo.Transaction_Details.showPrice, dbo.Transaction_Details.ItemSerial, dbo.Transaction_Details.ItemDiscountType,"
End If

StrSQL = StrSQL & "                      dbo.Transaction_Details.ItemDiscount, dbo.Transaction_Details.guaranteeTime, dbo.Transaction_Details.CostPrice, dbo.Transaction_Details.CostTransID,"
StrSQL = StrSQL & "                      dbo.Transaction_Details.ItemProfit, dbo.Transaction_Details.Remarks, dbo.Transaction_Details.ItemSize, dbo.Transaction_Details.QtyBySmalltUnit,"
StrSQL = StrSQL & "                      dbo.Transaction_Details.Remarks1"
StrSQL = StrSQL & " FROM         dbo.TblUnites RIGHT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.Transaction_Details ON dbo.TblUnites.UnitID = dbo.Transaction_Details.UnitId RIGHT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.Transactions ON dbo.Transaction_Details.Transaction_ID = dbo.Transactions.Transaction_ID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblEmployee ON dbo.Transactions.Emp_ID = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblEmpDepartments ON dbo.Transactions.DepartementID = dbo.TblEmpDepartments.DeparmentID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.FixedAssets ON dbo.Transactions.FixesAssetsID = dbo.FixedAssets.id LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblBranchesData ON dbo.Transactions.BranchId = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblStore ON dbo.Transactions.StoreID = dbo.TblStore.StoreID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID RIGHT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblItems ON dbo.Transaction_Details.Item_ID = dbo.TblItems.ItemID"
StrSQL = StrSQL & " Where (dbo.Transactions.Transaction_Type = 19)"
End If
If Me.dcBranch.Text <> "" And val(dcBranch.BoundText) <> 0 Then
    StrSQL = StrSQL & " AND   dbo.Transactions.BranchId = " & val(Me.dcBranch.BoundText)
End If
If Me.DcboEmpDepartments.Text <> "" And val(DcboEmpDepartments.BoundText) <> 0 Then
    StrSQL = StrSQL & " AND   dbo.Transactions.DepartementID = " & val(Me.DcboEmpDepartments.BoundText)
End If
If Me.DCEquipments.Text <> "" And val(DCEquipments.BoundText) <> 0 Then
    StrSQL = StrSQL & " AND   dbo.Transactions.FixesAssetsID = " & val(Me.DCEquipments.BoundText)
End If
If Me.DCboStoreName.Text <> "" And val(DCboStoreName.BoundText) <> 0 Then
    StrSQL = StrSQL & " AND   dbo.Transactions.StoreID = " & val(Me.DCboStoreName.BoundText)
End If
If Me.DBCboClientName.Text <> "" And val(DBCboClientName.BoundText) <> 0 Then
    StrSQL = StrSQL & " AND   dbo.Transactions.CusID = " & val(Me.DBCboClientName.BoundText)
End If
If Me.DcboEmpName.Text <> "" And val(DcboEmpName.BoundText) <> 0 Then
    StrSQL = StrSQL & " AND   dbo.Transactions.Emp_ID = " & val(Me.DcboEmpName.BoundText)
End If

 If Not IsNull(Me.DtpDateFrom.value) Then
                   StrSQL = StrSQL & " AND dbo.Transactions.Transaction_Date >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
      End If
       If Not IsNull(Me.DtpDateTo.value) Then
                   StrSQL = StrSQL & " AND dbo.Transactions.Transaction_Date<=" & SQLDate(Me.DtpDateTo.value, True) & ""
      End If
 

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
   Msg = "No Data"
   End If
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
   
   If Me.ChAll.value = True Then
        If SystemOptions.UserInterface = ArabicInterface Then
        
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAllStoreExchangReport.rpt"
            Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAllStoreExchangReportE.rpt"
            
       End If
       End If
       If Me.ChDetails.value = True Then
             If SystemOptions.UserInterface = ArabicInterface Then
        
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepDetailsStoreExchangReport.rpt"
            Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepDetailsStoreExchangReportE.rpt"
            
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
      If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
       Else
       Msg = "No Data"
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

    If SystemOptions.UserInterface = ArabicInterface Then
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
        ' xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
        StrReportTitle = "" '& StrAccountName
    Else
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""
  
    End If

   
  If DtpDateFrom.value <> "" And DtpDateTo.value <> "" Then
   xReport.ParameterFields(8).AddCurrentValue DtpDateFrom.value

    xReport.ParameterFields(10).AddCurrentValue DtpDateTo.value
  '  xReport.ParameterFields(11).AddCurrentValue DtpDateToH.value
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




