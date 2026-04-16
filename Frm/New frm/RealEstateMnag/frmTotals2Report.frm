VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SuiteCtrls.ocx"
Begin VB.Form FrmTotals2Report 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   6060
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13575
   Icon            =   "frmTotals2Report.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   13575
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
      TabIndex        =   12
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
      Height          =   4005
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   840
      Width           =   13635
      Begin XtremeSuiteControls.RadioButton RdOmala 
         Height          =   375
         Left            =   3840
         TabIndex        =   15
         Top             =   1320
         Width           =   1095
         _Version        =   786432
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "«·⁄„·«¡"
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
      Begin VB.TextBox TxtSearchCode 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6180
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   480
         Width           =   915
      End
      Begin VB.Frame Frame3 
         Height          =   3855
         Left            =   8520
         TabIndex        =   7
         Top             =   120
         Width           =   5175
         Begin VB.Image Image1 
            Height          =   2535
            Left            =   0
            Picture         =   "frmTotals2Report.frx":038A
            Stretch         =   -1  'True
            Top             =   120
            Width           =   5100
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
            Height          =   855
            Left            =   240
            TabIndex        =   8
            Top             =   2640
            Width           =   4815
         End
      End
      Begin MSDataListLib.DataCombo DcbBranch 
         Height          =   315
         Left            =   240
         TabIndex        =   4
         Top             =   120
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcboEmpName 
         Height          =   315
         Left            =   240
         TabIndex        =   14
         Top             =   480
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton RdRenter 
         Height          =   375
         Left            =   600
         TabIndex        =   17
         Top             =   2760
         Width           =   1095
         _Version        =   786432
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "«·„” √Ã—Ì‰"
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton RdCustomer 
         Height          =   375
         Left            =   3840
         TabIndex        =   18
         Top             =   1800
         Width           =   1095
         _Version        =   786432
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "«·„Ê—œÌ‰"
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton RdOwner 
         Height          =   375
         Left            =   2520
         TabIndex        =   19
         Top             =   3240
         Width           =   1095
         _Version        =   786432
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "«·„·«ﬂ"
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton RdBanck 
         Height          =   375
         Left            =   2160
         TabIndex        =   20
         Top             =   1800
         Width           =   1455
         _Version        =   786432
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "«·»‰Êﬂ"
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton RdStore 
         Height          =   375
         Left            =   3840
         TabIndex        =   21
         Top             =   2280
         Width           =   1095
         _Version        =   786432
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "«·„Œ«“‰"
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton RdGroupItem 
         Height          =   375
         Left            =   2160
         TabIndex        =   22
         Top             =   2280
         Width           =   1455
         _Version        =   786432
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "„Ã„Ê⁄… «·«’‰«›"
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton RdSales 
         Height          =   375
         Left            =   3840
         TabIndex        =   23
         Top             =   2760
         Width           =   1095
         _Version        =   786432
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "«·„‰«œÌ»"
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton RdCushCus 
         Height          =   375
         Left            =   2160
         TabIndex        =   24
         Top             =   1320
         Width           =   1455
         _Version        =   786432
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "«·⁄„·«¡ «·‰ﬁœÌ"
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton RdIqar 
         Height          =   375
         Left            =   2160
         TabIndex        =   25
         Top             =   2760
         Width           =   1455
         _Version        =   786432
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "«·⁄ﬁ«—« "
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton RdSuppller 
         Height          =   375
         Left            =   600
         TabIndex        =   26
         Top             =   1320
         Width           =   1095
         _Version        =   786432
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "«·„ ⁄ÂœÌ‰"
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton RdBox 
         Height          =   375
         Left            =   600
         TabIndex        =   27
         Top             =   1800
         Width           =   1095
         _Version        =   786432
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "«·Œ“‰ Ê«·⁄Âœ"
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton RdGroupSales 
         Height          =   375
         Left            =   240
         TabIndex        =   28
         Top             =   2280
         Width           =   1455
         _Version        =   786432
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "„Ã„Ê⁄… «·„‰«œÌ»"
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DboParentAccount 
         Height          =   315
         Left            =   240
         TabIndex        =   29
         Top             =   840
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton rdACC 
         Height          =   375
         Left            =   600
         TabIndex        =   31
         Top             =   3240
         Width           =   1095
         _Version        =   786432
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "«·Õ”«»« "
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
      Begin VB.Label lblParentAcc 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   " «·Õ”«» «·—∆Ì”Ì   "
         Height          =   315
         Left            =   7110
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   840
         Width           =   1230
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÿ»ﬁ« ·„ÊŸ› „⁄Ì‰"
         Height          =   195
         Index           =   1
         Left            =   7170
         TabIndex        =   16
         Top             =   480
         Width           =   1170
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÿ»ﬁ« ·›—⁄ „⁄Ì‰"
         Height          =   195
         Index           =   0
         Left            =   7320
         TabIndex        =   5
         Top             =   120
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
      Left            =   1800
      Picture         =   "frmTotals2Report.frx":28E2
      Stretch         =   -1  'True
      Top             =   120
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
      Height          =   825
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   13485
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
Attribute VB_Name = "FrmTotals2Report"
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
Private Sub ChangeLang()
 Cmd(0).Caption = "Show"
 Cmd(2).Caption = "Exit"
btnClear.Caption = "Clear"
Label5.Caption = "Total Reports"
Label1(0).Caption = "Branch"
Label1(1).Caption = "Employee"
lblCompanyname.Caption = "ALL SATTARYAH "
RdOmala.Caption = "Customers"
RdOmala.RightToLeft = False
RdCushCus.RightToLeft = False
RdCushCus.Caption = "Cash Customers "
RdCustomer.RightToLeft = False
RdCustomer.Caption = "Vendors"
RdBanck.Caption = "Bancks"
RdBanck.RightToLeft = False
RdStore.Caption = "Stores"
RdStore.RightToLeft = False
RdSales.RightToLeft = False
RdSales.Caption = "SuperVisor"
RdGroupItem.RightToLeft = False
RdGroupItem.Caption = "Group Items"
RdIqar.RightToLeft = False
RdIqar.Caption = "Real Estate"
RdOwner.RightToLeft = False
RdOwner.Caption = "Owners"
RdRenter.RightToLeft = False
RdRenter.Caption = "Renters"
RdGroupSales.RightToLeft = False
RdGroupSales.Caption = "Group SuperVisor"
RdSuppller.RightToLeft = False
RdSuppller.Caption = "Contractors"
RdBox.RightToLeft = False
RdBox.Caption = "Cash On Hand"
End Sub
Private Sub LoadPremis()
    Dim i As Integer
    Dim BolTemp As Boolean
    Dim Msg As String

   RdOmala.Visible = DoPremis(Do_Open, "ReportSales", False)
     RdCustomer.Visible = DoPremis(Do_Open, "ReportPurchase", False)
    RdBox.Visible = DoPremis(Do_Open, "ReportBoxes", False)
  RdStore.Visible = DoPremis(Do_Open, "ReportStock", False)
  RdSales.Visible = DoPremis(Do_Open, "ReportSales", False)
    RdIqar.Visible = DoPremis(Do_Open, "ReportSales", False)
   
 
   RdOwner.Visible = DoPremis(Do_Open, "ReportPurchase", False)
   
  RdCushCus.Visible = DoPremis(Do_Open, "ReportSales", False)
   
   RdBanck.Visible = DoPremis(Do_Open, "ReporBanks", False)
 RdBanck.Visible = DoPremis(Do_Open, "ReporBanks", False)
  RdGroupItem.Visible = DoPremis(Do_Open, "ReportItems", False)
 
      RdGroupSales.Visible = DoPremis(Do_Open, "ReportSales", False)
      RdRenter.Visible = DoPremis(Do_Open, "ReportSales", False)
Me.RdSuppller.Visible = DoPremis(Do_Open, "ReportSales", False)
 

  
 
 If mdifrmmain.AssetsMngBase.Visible = False Then
 RdRenter.Visible = False
 RdOwner.Visible = False
 
 End If
 
End Sub


Private Sub btnClear_Click()
clear_all Me
Me.RdBanck.value = False
Me.RdBox.value = False
Me.RdCustomer.value = False
Me.RdGroupItem.value = False
Me.RdGroupSales.value = False
Me.RdOmala.value = False
Me.RdOwner.value = False
Me.RdRenter.value = False
Me.RdSales.value = False
Me.RdStore.value = False
Me.RdSuppller.value = False

End Sub

Private Sub Cmd_Click(Index As Integer)

    Select Case Index

        Case 0
        If Me.rdACC.value = True Or Me.RdSuppller.value = True Or Me.RdCushCus.value = True Or Me.RdBanck.value = True Or Me.RdBox.value = True Or Me.RdCustomer.value = True Or Me.RdGroupItem.value = True Or Me.RdGroupSales.value = True Or Me.RdOmala.value = True Or Me.RdOwner.value = True Or Me.RdRenter.value = True Or Me.RdSales.value = True Or Me.RdStore.value = True Or Me.RdIqar.value = True Then
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






Private Sub DboParentAccount_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF3 Then
  Account_search.show
            Account_search.case_id = 180620
            End If
            
End Sub

Private Sub Form_Activate()
   PutFormOnTop Me.hWnd
End Sub
Private Sub DcboEmpName_Change()
DcboEmpName_Click (0)
End Sub

Private Sub DcboEmpName_Click(Area As Integer)
'    On Error Resume Next
       If val(DcboEmpName.BoundText) = 0 Then Exit Sub

    Dim EmpCode  As String
 
    GetEmployeeIDFromCode , , DcboEmpName.BoundText, EmpCode
    TxtSearchCode.text = EmpCode
    End Sub

Private Sub ImgFavorites_Click()
AddTofaforites Me.Name, Me.Caption, Me.Caption
End Sub

Private Sub TxtSearchCode_KeyPress(KeyAscii As Integer)
Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode TxtSearchCode.text, EmpID
        DcboEmpName.BoundText = EmpID
    End If

End Sub

Private Sub Form_Load()
   
 'On Error GoTo ErrTrap
    Dim i As Integer
    Dim My_SQL As String
    Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos
     Dcombos.GetEmployees Me.DcboEmpName
     Dcombos.GetBranches DcbBranch
Dcombos.GetAccountingCodes Me.DboParentAccount, False, True


    Set cSearch = New clsDCboSearch
    My_SQL = "TblContract"
    
    Set BKGrndPic = New ClsBackGroundPic
    Set RsSavRec = New ADODB.Recordset
    
    RsSavRec.CursorLocation = adUseClient
    RsSavRec.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdTableDirect
     If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
    Resize_Form Me
    LoadPremis
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
    If Me.RdSuppller.value = True Or Me.RdCustomer.value = True Or Me.RdOmala.value = True Or Me.RdOwner.value = True Or Me.RdRenter.value = True Or RdIqar.value = True Then
StrSQL = "SELECT     dbo.TblCustemers.CusID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.ResponsibleContact, dbo.TblCustemers.Cus_Phone, "
StrSQL = StrSQL & "                      dbo.TblCustemers.Cus_mobile, dbo.TblCustemers.Remark, dbo.TblCustemers.OpenBalance, dbo.TblCustemers.OpenBalanceType,"
StrSQL = StrSQL & "                      dbo.TblCustemers.OpenBalanceDate, dbo.TblCustemers.CreditLimit, dbo.TblCustemers.Account_Code_As_Client, dbo.TblCustemers.Account_Code_As_Supplier,"
StrSQL = StrSQL & "                      dbo.TblCustemers.CreditlimitCredit, dbo.TblCustemers.FaxNumber, dbo.TblCustemers.E_mail, dbo.TblCustemers.SaleType, dbo.TblCustemers.Account_Code,"
StrSQL = StrSQL & "                      dbo.TblCustemers.Trans_Discount, dbo.TblCustemers.Trans_DiscountType, dbo.TblCustemers.CountryID, dbo.TblCountriesData.CountryName,"
StrSQL = StrSQL & "                      dbo.TblCustemers.CityID, dbo.TblCountriesGovernmentsCities.CityName, dbo.TblCustemers.GovernmentID, dbo.TblCountriesGovernments.GovernmentName,"
StrSQL = StrSQL & "                      dbo.TblCustemers.Address, dbo.TblCustemers.Trans_DiscountPur, dbo.TblCustemers.Trans_DiscountTypePur, dbo.TblCustemers.CountEmp, dbo.TblCustemers.ToTal,"
StrSQL = StrSQL & "                       dbo.TblCustemers.c1, dbo.TblCustemers.c2, dbo.TblCustemers.Remark2, dbo.TblCustemers.locked, dbo.TblCustemers.parent_account,"
StrSQL = StrSQL & "                      dbo.TblCustemers.opening_balance_voucher_id, dbo.TblCustemers.DepitInterval, dbo.TblCustemers.CreditInterval, dbo.TblCustemers.DepitIntervalID,"
StrSQL = StrSQL & "                      dbo.TblCustemers.CreditIntervalID, dbo.TblCustemers.prifix, dbo.TblCustemers.code, dbo.TblCustemers.Fullcode, dbo.TblCustemers.CustomerandVendor,"
StrSQL = StrSQL & "                      dbo.TblCustemers.CustomerTypeID, dbo.TblCustemers.BranchId, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee,"
StrSQL = StrSQL & "                      dbo.TblCustemers.CustGID, dbo.TblCustemers.ExpireDateH, dbo.TblCustemers.Company, dbo.TblCustemers.JobTitle, dbo.TblCustemers.Salary,"
StrSQL = StrSQL & "                      dbo.TblCustemers.JobAddress, dbo.TblCustemers.JobTel, dbo.TblCustemers.JobTelConvert, dbo.TblCustemers.HomeTel, dbo.TblCustemers.Mobile1,"
StrSQL = StrSQL & "                      dbo.TblCustemers.Mobile2, dbo.TblCustemers.CountryID2, dbo.TblCustemers.Sex, dbo.TblCustemers.Account_Code1, dbo.TblCustemers.Account_Code2,"
StrSQL = StrSQL & "                      dbo.TblCustemers.ParentAccount, dbo.TblCustemers.OpenBalanceType1, dbo.TblCustemers.OpenBalance1, dbo.TblCustemers.OpenBalanceType2,"
StrSQL = StrSQL & "                      dbo.TblCustemers.OpenBalance2, dbo.TblCustemers.ShowQty1, dbo.TblCustemers.showPrice1, dbo.TblCustemers.showPrice2, dbo.TblCustemers.Salaries1,"
StrSQL = StrSQL & "                      dbo.TblCustemers.Salaries2, dbo.TblCustemers.ShowQty1c, dbo.TblCustemers.showPrice1c, dbo.TblCustemers.showPrice2c, dbo.TblCustemers.Salaries1c,"
StrSQL = StrSQL & "                      dbo.TblCustemers.Salaries2c, dbo.TblCustemers.Totald, dbo.TblCustemers.Totalc, dbo.TblCustemers.RecordDate, dbo.TblCustemers.balanced,"
StrSQL = StrSQL & "                      dbo.TblCustemers.balancec, dbo.TblCustemers.TypeCustomer, dbo.TblCustemers.BoxMil, dbo.TblCustemers.ZipCode, dbo.ACCOUNTS.Account_Serial,"
StrSQL = StrSQL & "                      dbo.TblCustemers.Type, dbo.TblCustemers.EmpId, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Name1,"
StrSQL = StrSQL & "                      dbo.TblEmployee.Emp_Name2, dbo.TblEmployee.Emp_Name3, dbo.TblEmployee.Emp_Name4, dbo.TblEmployee.Fullcode AS EmpFullcode,"
StrSQL = StrSQL & "                      dbo.TblEmployee.Emp_Namee4, dbo.TblEmployee.Emp_Namee3, dbo.TblEmployee.Emp_Namee2, dbo.TblEmployee.Emp_Namee1, dbo.TblEmployee.Emp_Namee,"
StrSQL = StrSQL & "                      dbo.TblBranchesData.branch_id, dbo.TblCustemers.RsID, dbo.ACCOUNTS.Account_Name, dbo.ACCOUNTS.Account_NameEng, dbo.TblCustemers.BankAccount,"
StrSQL = StrSQL & "                      dbo.TblCustemers.BankName, dbo.TblCustemers.RsIDDateH, dbo.TblCustemers.RsDegree, dbo.TblCustemers.RsIDDate, dbo.TblCustemers.IBAN,"
StrSQL = StrSQL & "                      dbo.TblCustemers.GroupInvestor, dbo.TblCustemers.TypeInvestor, dbo.TblCustemers.Flg, dbo.TblCustemers.BankID, dbo.TblCustemers.Category,"
StrSQL = StrSQL & "                      dbo.TblCustemers.RecorddateH , dbo.TblCustemers.RecordNo,"
StrSQL = StrSQL & "                      TblCustemers.CustGID , TblCustemers.VATNO, TblCustemers.PlotIdentification, TblCustemers.PostalZone,"
StrSQL = StrSQL & "                      GroupsCustomers.GroupName,GroupsCustomers.GroupNamee,"
StrSQL = StrSQL & "                      ClassCustomers.Name ClassCustomersName,ClassCustomers.NameE ClassCustomersNamee,TblCustemers.StreetName,TblCustemers.BuildingNumber,TblCustemers.BankIBAN ,TblCustemers.BankCode"

StrSQL = StrSQL & " FROM         dbo.TblCustemers LEFT OUTER JOIN "
StrSQL = StrSQL & "                      dbo.TblEmployee ON dbo.TblCustemers.EmpId = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.ACCOUNTS ON dbo.TblCustemers.Account_Code = dbo.ACCOUNTS.Account_Code LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblBranchesData ON dbo.TblCustemers.BranchId = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblCountriesGovernments ON dbo.TblCustemers.GovernmentID = dbo.TblCountriesGovernments.GovernmentID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblCountriesGovernmentsCities ON dbo.TblCustemers.CityID = dbo.TblCountriesGovernmentsCities.CityID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblCountriesData ON dbo.TblCustemers.CountryID = dbo.TblCountriesData.CountryID"

StrSQL = StrSQL & "                      Left outer join ClassCustomers On ClassCustomers.ID = TblCustemers.ClassCustomersId"
StrSQL = StrSQL & "                      Left outer join GroupsCustomers On GroupsCustomers.GroupID = TblCustemers.GroupsCustomersId"

StrSQL = StrSQL & " Where  (1 = 1)"
If Me.RdCustomer.value = True Or Me.RdSuppller.value = True Then
    StrSQL = StrSQL & " AND   dbo.TblCustemers.Type = 2"
End If
If Me.RdOwner.value = True Then
    StrSQL = StrSQL & " AND   dbo.TblCustemers.Type = 57"
End If

If Me.RdOmala.value = True Then
    StrSQL = StrSQL & " AND   dbo.TblCustemers.Type = 1"
End If
If Me.RdRenter.value = True Then
    StrSQL = StrSQL & " AND   dbo.TblCustemers.Type = 56"
End If

End If
If Me.RdBanck.value = True Then

StrSQL = " SELECT     dbo.BanksData.BankID, dbo.BanksData.BankName, dbo.BanksData.Remarks, dbo.BanksData.Account_Code, dbo.BanksData.Branch, "
StrSQL = StrSQL & "                       dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.BanksData.Account_Code1, dbo.BanksData.Account_Code2, dbo.BanksData.report_no,"
StrSQL = StrSQL & "                       dbo.BanksData.BranchId, dbo.BanksData.Account_code3, dbo.BanksData.Commision, dbo.BanksData.ParetnAccount, dbo.BanksData.BankNamee,"
StrSQL = StrSQL & "                       dbo.BanksData.opening_balance_voucher_id, dbo.BanksData.OpenBalanceDate, dbo.BanksData.OpenBalanceType, dbo.BanksData.OpenBalance,"
StrSQL = StrSQL & "                       dbo.BanksData.account_no, dbo.BanksData.IBan, dbo.BanksData.Branch_NO, dbo.BanksData.Tel, dbo.BanksData.Address, dbo.BanksData.Email,"
StrSQL = StrSQL & "                       dbo.BanksData.Currency_ID , dbo.BanksData.chkapprov, dbo.BanksData.chkLoan, dbo.TblBranchesData.branch_id"
StrSQL = StrSQL & "  FROM         dbo.BanksData LEFT OUTER JOIN"
StrSQL = StrSQL & "                       dbo.TblBranchesData ON dbo.BanksData.BranchId = dbo.TblBranchesData.branch_id"
StrSQL = StrSQL & " Where  (1 = 1)"
End If
If Me.RdBox.value = True Then
StrSQL = " SELECT     dbo.TblBoxesData.BoxID, dbo.TblBoxesData.BoxName, dbo.TblBoxesData.Comments, dbo.TblBoxesData.Type, dbo.TblBoxesData.Account_Code,"
StrSQL = StrSQL & "                      dbo.TblBoxesData.empid, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Name1, dbo.TblEmployee.Emp_Name2,"
StrSQL = StrSQL & "                      dbo.TblEmployee.Emp_Name3, dbo.TblEmployee.Emp_Name4, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee4, dbo.TblEmployee.Emp_Namee3,"
StrSQL = StrSQL & "                      dbo.TblEmployee.Emp_Namee2, dbo.TblEmployee.Emp_Namee1, dbo.TblEmployee.Emp_Namee, dbo.TblBoxesData.BranchId, dbo.TblBranchesData.branch_name,"
StrSQL = StrSQL & "                      dbo.TblBranchesData.branch_namee, dbo.TblBoxesData.BoxNameE, dbo.TblBoxesData.Account_Code1, dbo.TblBoxesData.OpenBalanceDate,"
StrSQL = StrSQL & "                      dbo.TblBoxesData.OpenBalanceType, dbo.TblBoxesData.OpenBalance, dbo.TblBoxesData.boxValue, dbo.TblBoxesData.Account_Code2, dbo.TblBoxesData.BTtype,"
StrSQL = StrSQL & "                      dbo.TblBoxesData.DriverId, dbo.TblBoxesData.opening_balance_voucher_id, dbo.TblBoxesData.ChequeBox, dbo.TblBoxesData.ParentAccount,"
StrSQL = StrSQL & "                      dbo.TblBranchesData.branch_id"
StrSQL = StrSQL & " FROM         dbo.TblBoxesData LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblBranchesData ON dbo.TblBoxesData.BranchId = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblEmployee ON dbo.TblBoxesData.empid = dbo.TblEmployee.Emp_ID"
StrSQL = StrSQL & " Where  (1 = 1)"
End If
If Me.RdGroupItem.value = True Then
StrSQL = " SELECT     dbo.Groups.GroupID, dbo.Groups.GroupName, dbo.Groups.GroupCode, dbo.Groups.Fullcode, dbo.Groups.EXpirType, dbo.Groups.prifix, dbo.Groups.EXpireValue,"
 StrSQL = StrSQL & "                     dbo.Groups.GroupNamee, dbo.Groups.OverHead, dbo.Groups.LastGroup, dbo.Groups.code, dbo.Groups.Branch_NO, dbo.Groups.ParentID,"
 StrSQL = StrSQL & "                     Groups_1.GroupName AS ParGroupName, Groups_1.GroupNamee AS ParGroupNameE, Groups_1.code AS Parcode, Groups_1.Fullcode AS ParFullcode,"
StrSQL = StrSQL & "                      dbo.TblBranchesData.branch_id , dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_nameE"
StrSQL = StrSQL & " FROM         dbo.Groups LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.TblBranchesData ON dbo.Groups.Branch_NO = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
    StrSQL = StrSQL & "                  dbo.Groups Groups_1 ON dbo.Groups.ParentID = Groups_1.GroupID"
    StrSQL = StrSQL & " Where  (1 = 1)"
End If
If Me.RdGroupSales.value = True Then
StrSQL = " SELECT     dbo.TBLSalesRepGroups.*"
StrSQL = StrSQL & " From dbo.TBLSalesRepGroups"
 StrSQL = StrSQL & " Where  (1 = 1)"
End If
If Me.RdSales.value = True Then
StrSQL = " SELECT     dbo.TBLSalesRepData.id, dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Fullcode,"
StrSQL = StrSQL & "                      dbo.TblEmployee.Emp_Namee, dbo.TblBranchesData.branch_id, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee,"
StrSQL = StrSQL & "                      dbo.TBLSalesRepGroups.id AS Expr1, dbo.TBLSalesRepGroups.name, dbo.TBLSalesRepGroups.namee, dbo.TblEmpJobsTypes.JobTypeID,"
StrSQL = StrSQL & "                      dbo.TblEmpJobsTypes.JobTypeName , dbo.TblEmpJobsTypes.JobTypeNamee, dbo.TBLSalesRepData.discountvalue"
StrSQL = StrSQL & " FROM         dbo.TBLSalesRepData INNER JOIN"
StrSQL = StrSQL & "                      dbo.TblEmployee ON dbo.TBLSalesRepData.EmpID = dbo.TblEmployee.Emp_ID INNER JOIN"
StrSQL = StrSQL & "                      dbo.TblBranchesData ON dbo.TBLSalesRepData.BranchId = dbo.TblBranchesData.branch_id INNER JOIN"
StrSQL = StrSQL & "                      dbo.TBLSalesRepGroups ON dbo.TBLSalesRepData.GroupID = dbo.TBLSalesRepGroups.id LEFT OUTER JOIN"
StrSQL = StrSQL & "                     dbo.TblEmpJobsTypes ON dbo.TBLSalesRepData.JobID = dbo.TblEmpJobsTypes.JobTypeID"
StrSQL = StrSQL & " Where  (1 = 1)"

End If
If Me.RdStore.value = True Then

   StrSQL = " SELECT     dbo.TblStore.StoreID, dbo.TblStore.StoreName, dbo.TblStore.StoreAdress, dbo.TblStore.StorePhone, dbo.TblStore.Remarks, dbo.TblStore.Account_Code,"
   StrSQL = StrSQL & "                   dbo.TblStore.Account_Code1, dbo.TblStore.Account_Code2, dbo.TblStore.Account_Code3, dbo.TblStore.linked, dbo.TblStore.Code, dbo.TblStore.StoreNamee,"
   StrSQL = StrSQL & "                   dbo.TblStore.ParetnAccount, dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Namee,"
   StrSQL = StrSQL & "                   dbo.TblBranchesData.branch_id , dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_nameE"
   StrSQL = StrSQL & " FROM         dbo.TblStore INNER JOIN"
   StrSQL = StrSQL & "                   dbo.TblBranchesData ON dbo.TblStore.BranchId = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
   StrSQL = StrSQL & "                   dbo.TblEmployee ON dbo.TblStore.Emp_ID = dbo.TblEmployee.Emp_ID"
   StrSQL = StrSQL & " Where  (1 = 1)"

End If
If Me.RdIqar.value = True Then

StrSQL = "SELECT     dbo.TblAqar.Aqarid, dbo.TblAqar.aqarNo, dbo.TblAqar.aqartypeid, dbo.tblAkarType.name, dbo.tblAkarType.namee, dbo.TblAqar.BranchId, "
StrSQL = StrSQL & "                      dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblAqar.ownerid, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee,"
StrSQL = StrSQL & "                      dbo.TblAqar.SalesEmp, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Name1, dbo.TblEmployee.Emp_Name2, dbo.TblEmployee.Emp_Name3,"
StrSQL = StrSQL & "                      dbo.TblEmployee.Emp_Name4, dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee4, dbo.TblEmployee.Emp_Namee3,"
StrSQL = StrSQL & "                      dbo.TblEmployee.Emp_Namee2, dbo.TblEmployee.Emp_Namee1, dbo.TblAqar.DateHCont, dbo.TblAqar.DateCont, dbo.TblAqar.FromCotDateH,"
StrSQL = StrSQL & "                      dbo.TblAqar.FromCotDate, dbo.TblAqar.ToCotDateH, dbo.TblAqar.aqarname, dbo.TblAqarDetai.unittype, dbo.TblAqarDetai.roomscount, dbo.TblAqarDetai.meterPrice,"
StrSQL = StrSQL & "                      dbo.TblAqarDetai.rentType, dbo.TblAqarDetai.WCcount, dbo.TblAqarDetai.kithchencount, dbo.TblAqarDetai.length, dbo.TblAqarDetai.unitdesc, dbo.TblAqarDetai.unitno,"
StrSQL = StrSQL & "                      dbo.TblAqarDetai.RentValue, dbo.TblAqarDetai.haveFurniture, dbo.TblAqarDetai.namerentType, dbo.TblAqarDetai.Floor, dbo.TblAqarDetai.LoungeCount,"
StrSQL = StrSQL & "                      dbo.TblAqarDetai.ACCount, dbo.TblAqarDetai.Status, dbo.TblAqarDetai.Services, dbo.TblAqarDetai.Water, dbo.TblAqarDetai.electric, dbo.TblAqarDetai.ACCountspleat,"
StrSQL = StrSQL & "                      dbo.TblAqarDetai.UnitElectric , dbo.TblAqarDetai.MiniRentValue"
StrSQL = StrSQL & " FROM         dbo.TblAqar LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblAqarDetai ON dbo.TblAqar.Aqarid = dbo.TblAqarDetai.Aqarid LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblEmployee ON dbo.TblAqar.SalesEmp = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblCustemers ON dbo.TblAqar.ownerid = dbo.TblCustemers.CusID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblBranchesData ON dbo.TblAqar.BranchId = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.tblAkarType ON dbo.TblAqar.aqartypeid = dbo.tblAkarType.id"
StrSQL = StrSQL & " Where  (1 = 1)"
End If
If Me.RdCushCus.value = True Then
StrSQL = " SELECT     dbo.TblCusCsh.*"
StrSQL = StrSQL & " From dbo.TblCusCsh"
StrSQL = StrSQL & " Where  (1 = 1)"

End If
If Me.RdCushCus.value = False Then
If Me.RdGroupSales.value = False Then
If Me.DcbBranch.BoundText <> "" Then
    StrSQL = StrSQL & " AND   dbo.TblBranchesData.branch_id = " & val(Me.DcbBranch.BoundText)

End If
End If
End If
If Me.RdSuppller.value = True Or Me.RdOmala.value = True Or Me.RdRenter.value = True Or Me.RdBox.value = True Or Me.RdSales.value = True Or Me.RdStore.value = True Or RdIqar.value = True Then
If Me.DcboEmpName.BoundText <> "" Then
    StrSQL = StrSQL & " AND   dbo.TblEmployee.Emp_ID = " & val(Me.DcboEmpName.BoundText)

End If
End If


If rdACC.value = True And Me.DboParentAccount.BoundText <> "" Then
 
StrSQL = " select * from  Accounts"
StrSQL = StrSQL & "         Where Account_code"

StrSQL = StrSQL & "  IN (SELECT Code"

StrSQL = StrSQL & "      FROM   [FN_MAIN_ACCOUNT_SUB_CODES]('" & Me.DboParentAccount.BoundText & "', '" & Me.DboParentAccount.BoundText & "', 1)"

StrSQL = StrSQL & "  )"
StrSQL = StrSQL & "      OR (Account_Code = '" & Me.DboParentAccount.BoundText & "')"
 
 
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
    
      If Me.RdIqar.value = True Then
        If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAllIqar.rpt"
        Else
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAllIqar.rpt"
         
       End If
End If

    If Me.RdCushCus.value = True Then
        If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAllCashCustomers.rpt"
        Else
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAllCashCustomers.rpt"
         
       End If
End If

If Me.RdOwner.value = True Then
        If SystemOptions.UserInterface = ArabicInterface Then
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAllOwner.rpt"
        Else
        StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAllOwner.rpt"
         
       End If
End If

If Me.RdCustomer.value = True Then
        If SystemOptions.UserInterface = ArabicInterface Then
         StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAllCustomers.rpt"
         Else
          StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAllCustomers.rpt"
            
       End If
End If
If Me.RdOmala.value = True Then
        If SystemOptions.UserInterface = ArabicInterface Then
         StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAllOmala.rpt"
         Else
          StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAllOmalaE.rpt"
            
       End If
End If
If Me.RdSuppller.value = True Then
        If SystemOptions.UserInterface = ArabicInterface Then
         StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAllSuppler.rpt"
         Else
          StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAllSuppler.rpt"
            
       End If
End If

If Me.RdRenter.value = True Then
        If SystemOptions.UserInterface = ArabicInterface Then
         StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAllRenter.rpt"
         Else
          StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAllRenter.rpt"
            
       End If
End If
If Me.RdBanck.value = True Then
        If SystemOptions.UserInterface = ArabicInterface Then
         StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAllBanck.rpt"
         Else
          StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAllBanck.rpt"
            
       End If
End If

 If Me.RdBox.value = True Then
        If SystemOptions.UserInterface = ArabicInterface Then
         StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAllBoxes.rpt"
         Else
          StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAllBoxes.rpt"
            
       End If
End If

If Me.RdGroupItem.value = True Then
    If SystemOptions.UserInterface = ArabicInterface Then
         StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAllGroupItem.rpt"
         Else
          StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAllGroupItem.rpt"
            
       End If
End If
If Me.RdGroupSales.value = True Then
    If SystemOptions.UserInterface = ArabicInterface Then
         StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAllGroupSales.rpt"
         Else
          StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAllGroupSales.rpt"
            
       End If
End If
If Me.RdSales.value = True Then
    If SystemOptions.UserInterface = ArabicInterface Then
         StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAllSales.rpt"
         Else
          StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAllSales.rpt"
            
       End If
End If
If Me.RdStore.value = True Then
    If SystemOptions.UserInterface = ArabicInterface Then
         StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAllSroress.rpt"
         Else
          StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAllSroress.rpt"
            
       End If
End If

If rdACC.value = True Then
    If SystemOptions.UserInterface = ArabicInterface Then
         StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "newAccReports.rpt"
         Else
          StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "newAccReportsE.rpt"
            
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
       'If XPDtbFrom.value <> Null Or XPDtbFrom.value <> "" Then xReport.ParameterFields(3).AddCurrentValue Format(Me.XPDtbFrom.value, "yyyy/M/d")
       ' If XPDtpTo.value <> Null Or XPDtpTo.value <> "" Then xReport.ParameterFields(4).AddCurrentValue Format(Me.XPDtpTo.value, "yyyy/M/d")
       ' If XPDtbFrom.value <> Null Or XPDtbFrom.value <> "" Then xReport.ParameterFields(5).AddCurrentValue Me.Txt_DateHigriFrom.value
       ' If XPDtpTo.value <> Null Or XPDtpTo.value <> "" Then xReport.ParameterFields(6).AddCurrentValue Me.Txt_DateHigriTO.value
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        'xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
'        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
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
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName, , NoteSerial

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault


 
  
 
End Function










 
