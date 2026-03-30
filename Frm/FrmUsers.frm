VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Begin VB.Form FrmAddUser 
   Appearance      =   0  'Flat
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "‘«‘Â «÷«›Â „” Œœ„"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7290
   HelpContextID   =   260
   Icon            =   "FrmUsers.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   7290
   Begin VB.ComboBox CboPriv 
      BackColor       =   &H00C0FFFF&
      Height          =   315
      ItemData        =   "FrmUsers.frx":038A
      Left            =   120
      List            =   "FrmUsers.frx":0394
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   720
      Width           =   2415
   End
   Begin VB.TextBox XPTxtPassConfirm 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   3480
      MaxLength       =   50
      PasswordChar    =   "*"
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   1485
      Width           =   2415
   End
   Begin VB.TextBox XPTxtPass 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   3480
      MaxLength       =   50
      PasswordChar    =   "*"
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   1065
      Width           =   2415
   End
   Begin VB.TextBox XPTxtUserName 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      Height          =   330
      Left            =   3480
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   690
      Width           =   2415
   End
   Begin ImpulseButton.ISButton XPBtnOK 
      Height          =   375
      Left            =   930
      TabIndex        =   3
      Top             =   2835
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   661
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "„Ê«›ﬁ"
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
      ButtonImage     =   "FrmUsers.frx":03B0
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
   Begin ImpulseButton.ISButton XPBtnCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   90
      TabIndex        =   4
      Top             =   2835
      Width           =   795
      _ExtentX        =   1402
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
      ButtonImage     =   "FrmUsers.frx":074A
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
   Begin ImpulseButton.ISButton CmdHelp 
      Height          =   405
      Left            =   2940
      TabIndex        =   5
      Top             =   2835
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   714
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "„”«⁄œ…"
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
      ButtonImage     =   "FrmUsers.frx":0AE4
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   -2147483637
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   -2147483637
   End
   Begin MSDataListLib.DataCombo dcBranch 
      Height          =   315
      Left            =   120
      TabIndex        =   10
      Top             =   1560
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   12648447
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DCboStoreName 
      Height          =   315
      Left            =   3480
      TabIndex        =   13
      Top             =   2265
      Width           =   2325
      _ExtentX        =   4101
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   12648447
      ListField       =   "7"
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DcboBox 
      Height          =   315
      Left            =   3480
      TabIndex        =   14
      Top             =   1920
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   12648447
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin C1SizerLibCtl.C1Elastic Ele 
      Height          =   615
      Index           =   9
      Left            =   0
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   0
      Width           =   7275
      _cx             =   12832
      _cy             =   1085
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial (Arabic)"
         Size            =   24
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
      BackColor       =   16777215
      ForeColor       =   4210688
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   "‘«‘Â «÷«›Â „” Œœ„ "
      Align           =   0
      AutoSizeChildren=   7
      BorderWidth     =   0
      ChildSpacing    =   0
      Splitter        =   0   'False
      FloodDirection  =   0
      FloodPercent    =   0
      CaptionPos      =   7
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
      Begin VB.TextBox TxtModFlg 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000000FF&
         Height          =   0
         Left            =   3180
         TabIndex        =   18
         Top             =   0
         Visible         =   0   'False
         Width           =   0
      End
      Begin ImpulseButton.ISButton CmdInfo 
         Height          =   615
         Left            =   4995
         TabIndex        =   19
         Top             =   -120
         Visible         =   0   'False
         Width           =   30
         _ExtentX        =   53
         _ExtentY        =   1085
         ButtonStyle     =   1
         ButtonPositionImage=   4
         Caption         =   ""
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmUsers.frx":0E7E
         ButtonImageHover=   "FrmUsers.frx":1B58
         ColorButton     =   16777215
         DrawFocusRectangle=   0   'False
      End
   End
   Begin MSDataListLib.DataCombo Dcbobank 
      Height          =   315
      Left            =   120
      TabIndex        =   21
      Top             =   1920
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   12648447
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DCEmployee 
      Height          =   315
      Left            =   120
      TabIndex        =   23
      Top             =   1080
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   12648447
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·„ÊŸ›"
      Height          =   330
      Index           =   7
      Left            =   2640
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·»‰ﬂ «·«› —«÷Ì"
      Height          =   465
      Index           =   6
      Left            =   2520
      TabIndex        =   22
      Top             =   1905
      Width           =   900
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·„Œ“‰ «·«› —«÷Ì"
      Height          =   225
      Index           =   5
      Left            =   6000
      TabIndex        =   20
      Top             =   2280
      Width           =   1260
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·„Œ“‰"
      Height          =   240
      Index           =   24
      Left            =   10215
      TabIndex        =   16
      Top             =   4710
      Width           =   1710
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·Œ“Ì‰Â «·«› —«÷ÌÂ"
      Height          =   225
      Index           =   11
      Left            =   6000
      TabIndex        =   15
      Top             =   2025
      Width           =   1260
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "’·«ÕÌ« "
      Height          =   330
      Index           =   4
      Left            =   2640
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   720
      Width           =   735
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·›—⁄ «·«› —«÷Ì"
      Height          =   450
      Index           =   0
      Left            =   2640
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·„” Œœ„"
      Height          =   330
      Index           =   1
      Left            =   6300
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   690
      Width           =   975
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " √ﬂÌœ ﬂ·„… «·„—Ê—"
      Height          =   330
      Index           =   2
      Left            =   5940
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   1485
      Width           =   1215
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "ﬂ·„… «·„—Ê—"
      Height          =   330
      Index           =   3
      Left            =   6300
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   1065
      Width           =   855
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   7320
      X2              =   -60
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   1
      Left            =   5970
      Picture         =   "FrmUsers.frx":2832
      Top             =   1095
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   0
      Left            =   5970
      Picture         =   "FrmUsers.frx":2BBC
      Top             =   735
      Width           =   240
   End
End
Attribute VB_Name = "FrmAddUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdHelp_Click()
    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
    SystemOptions.SysHelp.HHDisplayTopicID Me.hWnd
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If Shift = 2 Then
        If KeyCode = vbKeyX Then
            XPBtnCancel_Click
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub ChangeLang()
    Me.Caption = "Add New User"
    Ele(9).Caption = Me.Caption
    lbl(1).Caption = "User Name"
    lbl(3).Caption = "PassWord"
    lbl(2).Caption = "Confirm Pw"
    lbl(11).Caption = "Default Box"
    lbl(5).Caption = "Default Store"
    lbl(4).Caption = "Type"
    lbl(7).Caption = "Employee"
    lbl(0).Caption = " Branch"
    lbl(6).Caption = "Default Bank"
    XPBtnOK.Caption = "Create"
    XPBtnCancel.Caption = "Cancel"

End Sub

Private Sub Form_Load()
    Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos
    Dcombos.GetBoxes Me.DcboBox
    Dcombos.GetStores Me.DCboStoreName
    Dcombos.GetBanks Me.Dcbobank
    Dcombos.GetBranches Me.Dcbranch
    Dcombos.GetEmployees DCEmployee

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
    
    Me.CboPriv.ListIndex = 0
    CenterForm Me

    FormPostion Me, GetPostion
 
End Sub

Private Sub Form_Unload(Cancel As Integer)

    FormPostion Me, SavePostion
End Sub

Private Sub XPBtnCancel_Click()
    Unload Me
End Sub

Private Sub XPBtnOK_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
    Dim RsTemp As ADODB.Recordset
    Dim StrSQL As String

    If CboPriv.ListIndex = -1 Then
        Msg = "Õœœ «·’·«ÕÌ« "
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        CboPriv.SetFocus
        Exit Sub
    End If
 
    If Trim(Dcbranch.BoundText) = "" Then
        If SystemOptions.UserInterface = EnglishInterface Then
            Msg = "Specify Branch"
        Else
            Msg = "Õœœ «·›—⁄ «·«› —«÷Ì  "
        End If

        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Dcbranch.SetFocus
        SendKeys "{F4}"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
 
    If Trim(Me.DCEmployee.BoundText) = "" Then
        If SystemOptions.UserInterface = EnglishInterface Then
            Msg = "Specify Employee"
        Else
            Msg = "Õœœ «·„ÊŸ›    "
        End If

        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        DCEmployee.SetFocus
        SendKeys "{F4}"
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
 
    If XPTxtUserName.text = "" Then
        Msg = "√œŒ· «”„ «·„” Œœ„"
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        XPTxtUserName.SetFocus
        Exit Sub
    End If

    If XPTxtPass.text = "" Then
        Msg = "√œŒ· ﬂ·„… «·„—Ê—"
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        XPTxtPass.SetFocus
        Exit Sub
    End If
 
    If XPTxtPassConfirm.text = "" Then
        Msg = "√œŒ·  √ﬂÌœ ﬂ·„… «·„—Ê—"
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        XPTxtPassConfirm.SetFocus
        Exit Sub
    End If

    If StrComp(XPTxtPass.text, XPTxtPassConfirm.text, vbTextCompare) <> 0 Then
        Msg = "ﬂ·„… «·„—Ê— Ê √ﬂÌœ ﬂ·„… «·„—Ê— " & Chr(13)
        Msg = Msg + "€Ì— „ ÿ«»ﬁ Ì‰"
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        XPTxtPassConfirm.SetFocus
        Exit Sub
    End If

    StrSQL = "select * From TblUsers where UserName='" & Trim(XPTxtUserName.text) & "'"
    Set RsTemp = New ADODB.Recordset
    RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Not (RsTemp.EOF Or RsTemp.BOF) Then
        Msg = "ÌÊÃœ „” Œœ„ „”Ã· „”»ﬁ« »Â–« «·«”„" & Chr(13)
        Msg = Msg + "»—Ã«¡ «· √ﬂœ „‰ «·«”„ «·’ÕÌÕ" & Chr(13)
        Msg = Msg + "√Ê  €ÌÌ— √Ê  „ÌÌ“ «”„ «·„” Œœ„"
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        XPTxtUserName.SetFocus
        RsTemp.Close
        Exit Sub
    End If

    RsTemp.Close
    RsTemp.Open "TblUsers", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    RsTemp.AddNew
    RsTemp("UserID").value = CStr(new_id("TblUsers", "UserID", "", True))
    RsTemp("UserName").value = Trim(XPTxtUserName.text)

    If Me.CboPriv.ListIndex = 0 Then
        RsTemp("UserType").value = 2
    Else
        RsTemp("InvPrices").value = 1
        RsTemp("ShowInvProfit").value = 1
        RsTemp("AllowOverMax").value = 1

        RsTemp("FullPremis").value = 1
        RsTemp("UserType").value = 0
    End If

    RsTemp("EmpID").value = val(Me.DCEmployee.BoundText)
    RsTemp("BranchId").value = val(Dcbranch.BoundText)
    RsTemp("StoreID").value = val(Me.DCboStoreName.BoundText)

    RsTemp("BoxID").value = val(Me.DcboBox.BoundText)
    RsTemp("BankID").value = val(Me.Dcbobank.BoundText)

    RsTemp("PassWord").value = Trim(XPTxtPass.text)
    RsTemp("IsActive").value = 1
    RsTemp.update
    Msg = " „ Õ›Ÿ »Ì«‰«  Â–« «·„” Œœ„ »‰Ã«Õ" & Chr(13)
    Msg = Msg + "Ê⁄·Ìﬂ «·¬‰  ”ÃÌ· ’·«ÕÌ«  Â–« «·„” Œœ„"
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title

    If SystemOptions.SysDataBaseType = AccessDataBase Then
        FrmUserAbility.Show
        FrmUserAbility.DCboUserName.BoundText = RsTemp("UserID").value
    End If

    Unload Me
    Exit Sub
ErrTrap:

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

Private Sub XPTxtPass_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyEscape Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Then
        Exit Sub
    End If

    If (InStr(1, "0123456789", Chr(KeyAscii), vbBinaryCompare) <> 0) Or (KeyAscii >= Asc("a") And KeyAscii <= Asc("z")) Or (KeyAscii >= Asc("A") And KeyAscii <= Asc("Z")) Or (KeyAscii >= Asc("√") And KeyAscii <= Asc("Ì")) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
        Beep
    End If

End Sub

Private Sub XPTxtPassConfirm_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyEscape Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Then
        Exit Sub
    End If

    If (InStr(1, "0123456789", Chr(KeyAscii), vbBinaryCompare) <> 0) Or (KeyAscii >= Asc("a") And KeyAscii <= Asc("z")) Or (KeyAscii >= Asc("A") And KeyAscii <= Asc("Z")) Or (KeyAscii >= Asc("√") And KeyAscii <= Asc("Ì")) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
        Beep
    End If

End Sub
