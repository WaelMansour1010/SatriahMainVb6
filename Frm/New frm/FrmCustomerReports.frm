VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SUITEC~1.OCX"
Begin VB.Form FrmCustomerReports 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   5460
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10365
   Icon            =   "FrmCustomerReports.frx":0000
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
      Index           =   1
      Left            =   0
      TabIndex        =   3
      Top             =   840
      Width           =   10395
      Begin VB.Frame Frame2 
         Caption         =   "„Õœœ« "
         Height          =   615
         Left            =   120
         TabIndex        =   34
         Top             =   2160
         Width           =   2895
         Begin VB.CheckBox Chkpayment 
            Caption         =   "«·ﬂ·"
            Height          =   255
            Index           =   0
            Left            =   2040
            TabIndex        =   35
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.Frame Fra 
         BackColor       =   &H00E2E9E9&
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   0
         Left            =   1200
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   1560
         Width           =   3225
         Begin VB.OptionButton Opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "<"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   2
            Left            =   2640
            RightToLeft     =   -1  'True
            TabIndex        =   33
            ToolTipText     =   "«’€— „‰"
            Top             =   0
            Width           =   555
         End
         Begin VB.OptionButton Opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "="
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   1
            Left            =   2160
            RightToLeft     =   -1  'True
            TabIndex        =   32
            ToolTipText     =   "Ì”«ÊÏ"
            Top             =   0
            Value           =   -1  'True
            Width           =   495
         End
         Begin VB.OptionButton Opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   ">"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   0
            Left            =   1560
            RightToLeft     =   -1  'True
            TabIndex        =   31
            ToolTipText     =   "«ﬂ»— „‰"
            Top             =   0
            Width           =   465
         End
         Begin VB.OptionButton Opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "=>"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   3
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   30
            ToolTipText     =   "«ﬂ»— „‰"
            Top             =   0
            Width           =   705
         End
         Begin VB.OptionButton Opt 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "=<"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   4
            Left            =   720
            RightToLeft     =   -1  'True
            TabIndex        =   29
            ToolTipText     =   "«’€— „‰"
            Top             =   0
            Width           =   795
         End
      End
      Begin VB.TextBox TxtMobile 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   1200
         Width           =   4215
      End
      Begin VB.Frame Frame1 
         Caption         =   "„‰ «·› —Â"
         Height          =   975
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   3000
         Width           =   4455
         Begin MSComCtl2.DTPicker DtpDateFrom 
            Height          =   330
            Left            =   1800
            TabIndex        =   20
            Top             =   150
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   104792067
            CurrentDate     =   41640
         End
         Begin MSComCtl2.DTPicker DtpDateTo 
            Height          =   330
            Left            =   1800
            TabIndex        =   21
            Top             =   480
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   104792067
            CurrentDate     =   41640
         End
         Begin Dynamic_Byte.NourHijriCal DtpDateFromH 
            Height          =   315
            Left            =   120
            TabIndex        =   22
            Top             =   150
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
         End
         Begin Dynamic_Byte.NourHijriCal DtpDateToH 
            Height          =   315
            Left            =   120
            TabIndex        =   23
            Top             =   510
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   556
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈·Ï"
            Height          =   195
            Index           =   3
            Left            =   3270
            RightToLeft     =   -1  'True
            TabIndex        =   25
            Top             =   600
            Width           =   480
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„‰"
            Height          =   195
            Index           =   4
            Left            =   3210
            RightToLeft     =   -1  'True
            TabIndex        =   24
            Top             =   120
            Width           =   540
         End
      End
      Begin XtremeSuiteControls.RadioButton RdTotal 
         Height          =   495
         Left            =   2760
         TabIndex        =   17
         Top             =   2040
         Width           =   2895
         _Version        =   786432
         _ExtentX        =   5106
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "«Ã„«·Ì „»Ì⁄«  «·⁄„·«¡ «·‰ﬁœÌ"
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
         Value           =   -1  'True
      End
      Begin VB.TextBox TxtValue 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   210
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   1530
         Width           =   885
      End
      Begin VB.TextBox TxtName 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   840
         Width           =   4215
      End
      Begin VB.Frame Frame3 
         Height          =   3855
         Left            =   5880
         TabIndex        =   7
         Top             =   120
         Width           =   4455
         Begin VB.Image Image1 
            Height          =   2835
            Left            =   0
            Picture         =   "FrmCustomerReports.frx":038A
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
            Left            =   240
            TabIndex        =   8
            Top             =   3000
            Width           =   2895
         End
      End
      Begin MSDataListLib.DataCombo DcbBranch 
         Height          =   315
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton RdAnalis 
         Height          =   495
         Left            =   2520
         TabIndex        =   18
         Top             =   2400
         Width           =   3135
         _Version        =   786432
         _ExtentX        =   5530
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   " Õ·Ì·Ì „»Ì⁄«  «·⁄„·«¡ «·‰ﬁœÌ"
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÿ»ﬁ« ·ÃÊ«· «·⁄„Ì·"
         Height          =   195
         Index           =   2
         Left            =   4455
         TabIndex        =   27
         Top             =   1200
         Width           =   1200
      End
      Begin VB.Label XPLbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÿ»ﬁ« ·„»·€ „Õœœ"
         Height          =   195
         Index           =   1
         Left            =   4485
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÿ»ﬁ« ·⁄„Ì· „Õœœ"
         Height          =   195
         Index           =   1
         Left            =   4485
         TabIndex        =   14
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÿ»ﬁ« ·›—⁄ „Õœœ"
         Height          =   195
         Index           =   0
         Left            =   4560
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
      Caption         =   "‘«‘…«· ﬁ«—Ì— «·⁄„·«¡ «·‰ﬁœÌ"
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
      Left            =   -15
      TabIndex        =   6
      Top             =   0
      Width           =   10350
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
Attribute VB_Name = "FrmCustomerReports"
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
Me.RdAnalis.value = False
Me.RdTotal.value = False
DtpDateFrom.value = ""
DtpDateTo.value = ""
End Sub

Private Sub Cmd_Click(Index As Integer)

    Select Case Index

        Case 0
If Me.RdAnalis.value = True Or Me.RdTotal.value = True Then
GetData
End If



            
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




Private Sub Form_Load()
   
 'On Error GoTo ErrTrap
    Dim i As Integer
    Dim My_SQL As String
    Dim Dcombos As ClsDataCombos
  
  
    Set Dcombos = New ClsDataCombos
 
   
    Dcombos.GetBranches DcbBranch
    
    'Dcombos.GetRentStatus dbcAqarStatus
    Me.RdAnalis.value = False
Me.RdTotal.value = False
DtpDateFrom.value = ""
DtpDateTo.value = ""
    Opt(0).value = False
    Opt(1).value = False
    Opt(2).value = False
    Opt(3).value = False
    Opt(4).value = False
    
  '  Set cSearch = New clsDCboSearch
  '  My_SQL = "TblContract"
  '
  '  Set BKGrndPic = New ClsBackGroundPic
  '  Set RsSavRec = New ADODB.Recordset
  '
  '  RsSavRec.CursorLocation = adUseClient
  '  RsSavRec.Open My_SQL, Cn, adOpenKeyset, adLockPessimistic, adCmdTableDirect
    
    
    
    Resize_Form Me
    
   
    
End Sub




Private Sub Form_Unload(Cancel As Integer)

    FormPostion Me, SavePostion
    Set DCboSearch = Nothing
End Sub

Public Sub GetData()
    Dim StrSQL As String
    Dim StrSQL1 As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer

If Me.RdTotal.value = True Then

StrSQL = "SELECT     dbo.Transactions.CashCustomerName, dbo.Transactions.CashCustomerPhone, SUM(QryTransactionsTotal.TransNet) AS totals, dbo.Transactions.BranchId, "
StrSQL = StrSQL & "                      dbo.TblBranchesData.branch_name , dbo.TblBranchesData.branch_nameE"
StrSQL = StrSQL & " FROM         dbo.Transactions INNER JOIN"
StrSQL = StrSQL & "                      dbo.QryTransactionsTotal() QryTransactionsTotal ON dbo.Transactions.Transaction_ID = QryTransactionsTotal.Transaction_ID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblBranchesData ON dbo.Transactions.BranchId = dbo.TblBranchesData.branch_id"
StrSQL = StrSQL & " WHERE 1=1"
If Chkpayment(0).value = vbUnchecked Then
StrSQL = StrSQL & "  and  ( dbo.Transactions.Transaction_Type =21 and    dbo.transactions.PaymentType = 0)"
Else
StrSQL = StrSQL & "  and  ( dbo.Transactions.Transaction_Type =21 )"
End If

 


End If
If Me.RdAnalis.value = True Then
StrSQL = " SELECT     dbo.Transactions.CashCustomerName, dbo.Transactions.CashCustomerPhone, QryTransactionsTotal.TransNet AS totals, dbo.Transactions.BranchId,"
StrSQL = StrSQL & "                      dbo.Transactions.Transaction_Date , dbo.Transactions.NoteSerial1, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_nameE"
StrSQL = StrSQL & " FROM         dbo.Transactions INNER JOIN"
StrSQL = StrSQL & "                      dbo.QryTransactionsTotal() QryTransactionsTotal ON dbo.Transactions.Transaction_ID = QryTransactionsTotal.Transaction_ID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblBranchesData ON dbo.Transactions.BranchId = dbo.TblBranchesData.branch_id"
'StrSQL = StrSQL & " Where    (dbo.Transactions.Transaction_Type = 21) and (dbo.Transactions.PaymentType = 0) And (Not (dbo.Transactions.CashCustomerName Is Null))"

 
 
 StrSQL = StrSQL & " WHERE   (Not (dbo.Transactions.CashCustomerName Is Null)) "
If Chkpayment(0).value = vbUnchecked Then
StrSQL = StrSQL & "  and  ( dbo.Transactions.Transaction_Type =21 and    dbo.transactions.PaymentType = 0)"
Else
StrSQL = StrSQL & "  and  ( dbo.Transactions.Transaction_Type =21 )"
End If





End If
 If Not IsNull(Me.DtpDateFrom.value) Then
                   StrSQL = StrSQL & " AND dbo.Transactions.Transaction_Date >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
      End If
       If Not IsNull(Me.DtpDateTo.value) Then
                   StrSQL = StrSQL & " AND dbo.Transactions.Transaction_Date <=" & SQLDate(Me.DtpDateTo.value, True) & ""
      End If
      
If Me.DcbBranch.BoundText <> "" Then
    StrSQL = StrSQL & " AND   dbo.TblBranchesData.branch_id = " & val(Me.DcbBranch.BoundText)

End If
If Me.RdTotal.value = True Then
If Me.TxtValue.Text <> "" Then
If Opt(0).value = True Then
    StrSQL1 = StrSQL1 & " AND  SUM(QryTransactionsTotal.TransNet) > " & val(Me.TxtValue.Text)
ElseIf Opt(1).value = True Then
   StrSQL1 = StrSQL1 & " AND  SUM(QryTransactionsTotal.TransNet) = " & val(Me.TxtValue.Text)
ElseIf Opt(2).value = True Then
    StrSQL1 = StrSQL1 & " AND  SUM(QryTransactionsTotal.TransNet) < " & val(Me.TxtValue.Text)
ElseIf Opt(4).value = True Then
    StrSQL1 = StrSQL1 & " AND  SUM(QryTransactionsTotal.TransNet) <= " & val(Me.TxtValue.Text)
ElseIf Opt(3).value = True Then
    StrSQL1 = StrSQL1 & " AND  SUM(QryTransactionsTotal.TransNet) >= " & val(Me.TxtValue.Text)
End If
End If
End If
''//
If Me.RdAnalis.value = True Then
If Me.TxtValue.Text <> "" Then
If Opt(0).value = True Then
    StrSQL = StrSQL & " AND  QryTransactionsTotal.TransNet > " & val(Me.TxtValue.Text)
ElseIf Opt(1).value = True Then
   StrSQL = StrSQL & " AND  QryTransactionsTotal.TransNet = " & val(Me.TxtValue.Text)
ElseIf Opt(2).value = True Then
    StrSQL = StrSQL & " AND  QryTransactionsTotal.TransNet < " & val(Me.TxtValue.Text)
ElseIf Opt(4).value = True Then
    StrSQL = StrSQL & " AND  QryTransactionsTotal.TransNet <= " & val(Me.TxtValue.Text)
ElseIf Opt(3).value = True Then
    StrSQL = StrSQL & " AND  QryTransactionsTotal.TransNet >= " & val(Me.TxtValue.Text)
End If
End If
End If
If Me.TxtName.Text <> "" Then
    StrSQL = StrSQL & " AND  ( dbo.Transactions.CashCustomerName Like '%" & Trim(Me.TxtName.Text) & "%')"

End If
If Me.TxtMobile.Text <> "" Then
    StrSQL = StrSQL & " AND  ( dbo.Transactions.CashCustomerPhone Like '%" & Trim(Me.TxtMobile.Text) & "%')"

End If


If Me.RdTotal.value = True Then

StrSQL = StrSQL & " GROUP BY dbo.Transactions.CashCustomerName, dbo.Transactions.CashCustomerPhone, dbo.Transactions.BranchId, dbo.TblBranchesData.branch_name,"
StrSQL = StrSQL & "                     dbo.TblBranchesData.branch_nameE"
StrSQL = StrSQL & " Having (Not (dbo.Transactions.CashCustomerName Is Null))"
StrSQL = StrSQL & StrSQL1

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
   
If Me.RdTotal.value = True Then
        If SystemOptions.UserInterface = ArabicInterface Then
        
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAllCustomerCashRep.rpt"
            Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAllCustomerCashRep.rpt"
            
       End If
End If
If Me.RdAnalis.value = True Then
        If SystemOptions.UserInterface = ArabicInterface Then
        
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAnlysisCustomerCashRep.rpt"
            Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAnlysisCustomerCashRep.rpt"
            
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
       'If XPDtbFrom.value <> Null Or XPDtbFrom.value <> "" Then xReport.ParameterFields(3).AddCurrentValue Format(Me.XPDtbFrom.value, "yyyy/M/d")
       ' If XPDtpTo.value <> Null Or XPDtpTo.value <> "" Then xReport.ParameterFields(4).AddCurrentValue Format(Me.XPDtpTo.value, "yyyy/M/d")
       ' If XPDtbFrom.value <> Null Or XPDtbFrom.value <> "" Then xReport.ParameterFields(5).AddCurrentValue Me.Txt_DateHigriFrom.value
       ' If XPDtpTo.value <> Null Or XPDtpTo.value <> "" Then xReport.ParameterFields(6).AddCurrentValue Me.Txt_DateHigriTO.value
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        'xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""
  
    End If

   ' xReport.ParameterFields(3).AddCurrentValue user_name
  If DtpDateFrom.value <> "" And DtpDateTo.value <> "" Then
    xReport.ParameterFields(8).AddCurrentValue DtpDateFrom.value
    xReport.ParameterFields(9).AddCurrentValue DtpDateFromH.value
    xReport.ParameterFields(10).AddCurrentValue DtpDateTo.value
    xReport.ParameterFields(11).AddCurrentValue DtpDateToH.value
    End If
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



 
