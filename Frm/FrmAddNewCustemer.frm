VERSION 5.00
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Begin VB.Form FrmAddNewCustemer 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "≈÷«›… „Ê—œ √Ê ⁄„Ì· ÃœÌœ"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4995
   Icon            =   "FrmAddNewCustemer.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   3690
   ScaleWidth      =   4995
   Begin VB.TextBox XPTxtCusName 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   780
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   465
      Width           =   3045
   End
   Begin VB.TextBox XPMTxtRemarks 
      Alignment       =   1  'Right Justify
      Height          =   1335
      Left            =   780
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   1590
      Width           =   3045
   End
   Begin VB.TextBox XPTxtMobile 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   780
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   1200
      Width           =   3045
   End
   Begin VB.TextBox XPTxtCusID 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   780
      Locked          =   -1  'True
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   75
      Width           =   3045
   End
   Begin VB.TextBox XPTxtPhone 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   780
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   825
      Width           =   3045
   End
   Begin ImpulseButton.ISButton XPButton301 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   60
      TabIndex        =   7
      Top             =   3210
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   661
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "≈·€«¡"
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
      ButtonImage     =   "FrmAddNewCustemer.frx":1CCA
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
   Begin ImpulseButton.ISButton XPBtnsave 
      Height          =   375
      Left            =   930
      TabIndex        =   5
      Top             =   3210
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   661
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "Õ›Ÿ"
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
      ButtonImage     =   "FrmAddNewCustemer.frx":2064
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
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   0
      X2              =   4995
      Y1              =   3090
      Y2              =   3105
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·„Ê—œ"
      Height          =   345
      Index           =   0
      Left            =   3870
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   465
      Width           =   1035
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·Â« ›"
      Height          =   345
      Index           =   3
      Left            =   3870
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   825
      Width           =   1035
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·ÃÊ«·"
      Height          =   345
      Index           =   2
      Left            =   3870
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   1200
      Width           =   1035
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "„·«ÕŸ« "
      Height          =   315
      Index           =   4
      Left            =   3870
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   1590
      Width           =   1035
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "ﬂÊœ «·„Ê—œ"
      Height          =   345
      Index           =   1
      Left            =   3870
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   75
      Width           =   1035
   End
End
Attribute VB_Name = "FrmAddNewCustemer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_DealingForm As GridTransType

Private m_DcboCustomers As DataCombo

Dim m_AddType As Integer

Private Sub Form_Activate()
    SetMeForAdd
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If Shift = 2 Then
        If KeyCode = vbKeyX Then
            XPButton301_Click
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Load()
    On Error GoTo ErrTrap
    CenterForm Me

    FormPostion Me, GetPostion
    clear_all Me
    Exit Sub
ErrTrap:
End Sub

Private Sub PassData()
    Dim StrSQL As String
    On Error GoTo ErrTrap
    StrSQL = "SELECT * From TblCustemers"

    Select Case Me.DealingForm

        Case PurchaseTransaction
            fill_combo Me.DcboCustomers, StrSQL
            Me.DcboCustomers.BoundText = val(XPTxtCusID.text)

        Case InvoiceTransaction
            fill_combo Me.DcboCustomers, StrSQL
            Me.DcboCustomers.BoundText = val(XPTxtCusID.text)

        Case Maintenance
            fill_combo FrmMaintenence.DBCboClientName, StrSQL
            FrmMaintenence.DBCboClientName.BoundText = val(XPTxtCusID.text)

        Case PriceList
            StrSQL = "SELECT * From TblCustemers where Type=2"
            fill_combo FrmMainPriceList.DBCboSupplierName, StrSQL
            FrmMainPriceList.DBCboSupplierName.BoundText = val(XPTxtCusID.text)

            '⁄—÷ «·√”⁄«—
        Case ShowPrice
            fill_combo FrmShowPrice.DBCboClientName, StrSQL
            FrmShowPrice.DBCboClientName.BoundText = val(XPTxtCusID.text)
    End Select

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap

    FormPostion Me, SavePostion

    Exit Sub
ErrTrap:
End Sub

Private Sub XPBtnsave_Click()
    Dim RsTemp As ADODB.Recordset
    Dim StrSQL As String
    Dim Msg As String
    Dim BeginTrans  As Boolean
    Dim rs As ADODB.Recordset

    On Error GoTo ErrTrap

    If XPTxtCusName.text = "" Then
        Msg = "ÌÃ» «œŒ«· «”„ «·⁄„Ì·...!!!"
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        XPTxtCusName.SetFocus
        Exit Sub
    End If

    StrSQL = "select * From TblCustemers where CusName='" & Trim(XPTxtCusName.text) & "'"
    Set RsTemp = New ADODB.Recordset
    RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If RsTemp.RecordCount > 0 Then
        Msg = "ÌÊÃœ ⁄„Ì· „”Ã· „”»ﬁ« »Â–« «·«”„" & Chr(13)
        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·»Ì«‰«  «·„œŒ·… " & Chr(13)
        Msg = Msg + "√Ê  €ÌÌ— √Ê  „ÌÌ“ «·»Ì«‰«  «·„œŒ·…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        XPTxtCusName.SetFocus
        Exit Sub
    End If

    If Me.AddType = 1 Or Me.AddType = 2 Then
        Cn.BeginTrans
        BeginTrans = True
        Set rs = New ADODB.Recordset
        rs.Open "[TblCustemers]", Cn, adOpenStatic, adLockOptimistic, adCmdTable
        rs.AddNew
        XPTxtCusID.text = CStr(new_id("TblCustemers", "CusID", "", True))
        rs("CusID").value = val(XPTxtCusID.text)
        rs("CusName").value = Trim(XPTxtCusName.text)
        rs("Cus_Phone").value = IIf(XPTxtPhone.text = "", "", Trim(XPTxtPhone.text))
        rs("Cus_mobile").value = IIf(XPTxtmobile.text = "", "", Trim(XPTxtmobile.text))
        rs("Remark").value = IIf(XPMTxtRemarks.text = "", "", Trim(XPMTxtRemarks.text))

        If Me.AddType = 1 Then
            '⁄„Ì·
            rs("Type").value = 1
            rs("Account_Code").value = ModAccounts.AddNewAccount("a1a2a3", Trim$(Me.XPTxtCusName.text), True, False)
        ElseIf Me.AddType = 2 Then
            '„Ê—œ
            rs("Type").value = 2
            rs("Account_Code").value = ModAccounts.AddNewAccount("a2a3a1", Trim$(Me.XPTxtCusName.text), True, False)
        End If

        rs.update
        Cn.CommitTrans
        BeginTrans = False
    ElseIf Me.AddType = 0 Then
    End If

    PassData
    Unload Me
    Exit Sub
ErrTrap:

    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If

    If Err.Number = -2147217900 Then
        Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & Chr(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
End Sub

Private Sub XPButton301_Click()
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If

End Sub

Public Property Get DealingForm() As GridTransType
    DealingForm = m_DealingForm
End Property

Public Property Let DealingForm(ByVal vNewValue As GridTransType)
    'If vNewValue = OpeningBalance Or vNewValue = PurchaseTransaction Or vNewValue = InvoiceTransaction Then
    m_DealingForm = vNewValue
    'End If
End Property

Public Property Get DcboCustomers() As DataCombo
    Set DcboCustomers = m_DcboCustomers
End Property

Public Property Set DcboCustomers(ByVal vNewValue As DataCombo)
    Set m_DcboCustomers = vNewValue
End Property

Public Property Get AddType() As Integer
    AddType = m_AddType
End Property

Public Property Let AddType(ByVal vNewValue As Integer)
    m_AddType = vNewValue
End Property

Private Sub SetMeForAdd()

    If Me.AddType = 0 Then
        'Add a Cash Customer
        Me.Caption = "≈÷«›… »Ì«‰«  ⁄„Ì· ‰ﬁœÌ"
        Me.lbl(1).Visible = False
        Me.XPTxtCusID.Visible = False
    
    ElseIf Me.AddType = 1 Then
        '≈÷«›… ⁄„Ì·
    
    ElseIf Me.AddType = 2 Then
        '≈÷«›… „Ê—œ
    
    End If

End Sub
