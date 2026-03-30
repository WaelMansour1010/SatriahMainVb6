VERSION 5.00
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Begin VB.Form FrmEditPW 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " €ÌÌ— «·—ﬁ„ «·”—Ì"
   ClientHeight    =   2925
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4530
   HelpContextID   =   260
   Icon            =   "FrmEditPW.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   2925
   ScaleWidth      =   4530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox XPTxtOldPW 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   330
      MaxLength       =   50
      PasswordChar    =   "*"
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   870
      Width           =   2415
   End
   Begin VB.TextBox XPTXTNewPw 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   330
      MaxLength       =   50
      PasswordChar    =   "*"
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   1290
      Width           =   2415
   End
   Begin VB.TextBox XPTxtConfirm 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   330
      MaxLength       =   50
      PasswordChar    =   "*"
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   1710
      Width           =   2415
   End
   Begin ImpulseButton.ISButton XPBtnOK 
      Default         =   -1  'True
      Height          =   375
      Left            =   900
      TabIndex        =   3
      Top             =   2430
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
      ButtonImage     =   "FrmEditPW.frx":038A
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
      Left            =   60
      TabIndex        =   4
      Top             =   2430
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
      ButtonImage     =   "FrmEditPW.frx":0724
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
      Height          =   375
      Left            =   3480
      TabIndex        =   5
      Top             =   2430
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
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
      ButtonImage     =   "FrmEditPW.frx":0ABE
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   -2147483637
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      ColorTextShadow =   -2147483637
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "ﬂ·„… «·„—Ê— «·”«»ﬁ…"
      Height          =   315
      Index           =   4
      Left            =   2850
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   878
      Width           =   1365
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "ﬂ·„… «·„—Ê— «·ÃœÌœ…"
      Height          =   315
      Index           =   3
      Left            =   2850
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   1298
      Width           =   1365
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " √ﬂÌœ ﬂ·„… «·„—Ê—"
      Height          =   315
      Index           =   2
      Left            =   2850
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   1718
      Width           =   1365
   End
   Begin VB.Label Lbl 
      Alignment       =   2  'Center
      BackColor       =   &H00E2E9E9&
      Caption         =   " €ÌÌ— ﬂ·„… «·„—Ê— ··„” Œœ„"
      Height          =   315
      Index           =   0
      Left            =   2220
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   270
      Width           =   1875
   End
   Begin VB.Label XPLblUserName 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   270
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   0
      Left            =   4230
      Picture         =   "FrmEditPW.frx":0E58
      Top             =   300
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   1
      Left            =   4230
      Picture         =   "FrmEditPW.frx":11E2
      Top             =   870
      Width           =   240
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   6300
      X2              =   -60
      Y1              =   2310
      Y2              =   2310
   End
End
Attribute VB_Name = "FrmEditPW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdHelp_Click()
    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
    SystemOptions.SysHelp.HHDisplayTopicID Me.hwnd
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

Function ChangeLang()
    Me.Caption = "Change PW"
    lbl(0).Caption = Me.Caption
    lbl(4).Caption = "Old PW"

    lbl(3).Caption = "New PW"
    lbl(2).Caption = "Confirm PassWord"

    XPBtnOK.Caption = "OK"
    XPBtnCancel.Caption = "Cancel"

End Function

Private Sub Form_Load()
    On Error GoTo ErrTrap
    CenterForm Me

    FormPostion Me, GetPostion
    XPLblUserName.Caption = user_name

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Unload(Cancel As Integer)

    FormPostion Me, SavePostion
End Sub

Private Sub XPBtnCancel_Click()
    Unload Me
End Sub

Private Sub XPBtnOK_Click()
    On Error GoTo ErrTrap
    Dim StrSQL As String
    Dim Msg As String
    Dim BeginTrans As Boolean

    Dim rs As ADODB.Recordset

    If XPTxtOldPW.text = "" Then
        '    Msg = "√œŒ· ﬂ·„… «·„—Ê— «·”«»ﬁ…"
        '    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        '    XPTxtOldPW.SetFocus
        '    Exit Sub
    End If

    If StrComp(XPTxtOldPW.text, User_Password, vbTextCompare) <> 0 Then
        Msg = "ﬂ·„… «·„—Ê— «·ﬁœÌ„… €Ì— ’ÕÌÕ…"
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        XPTxtOldPW.SetFocus
        Exit Sub
    End If

    If XPTXTNewPw.text = "" Then
        Msg = "√œŒ· ﬂ·„… «·„—Ê— «·ÃœÌœ…"
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        XPTXTNewPw.SetFocus
        Exit Sub
    End If

    If XPTxtConfirm.text = "" Then
        Msg = "√œŒ·  √ﬂÌœ ﬂ·„… «·„—Ê— "
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        XPTxtConfirm.SetFocus
        Exit Sub
    End If

    If StrComp(XPTXTNewPw.text, XPTxtConfirm.text, vbTextCompare) <> 0 Then
        Msg = "ﬂ·„… «·„—Ê— Ê √ﬂÌœ ﬂ·„… «·„—Ê— " & Chr(13)
        Msg = Msg + "€Ì— „ ”«ÊÌ Ì‰"
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        XPTxtConfirm.SetFocus
        Exit Sub
    End If

    StrSQL = "SELECT * FROM TblUsers WHERE UserID=" & user_id
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
    Cn.BeginTrans
    BeginTrans = True
    rs("PassWord").value = Trim(XPTXTNewPw.text)
    rs("PassConfirm").value = Trim(XPTXTNewPw.text)
    
    rs.update
    Cn.CommitTrans
    BeginTrans = False
    User_Password = rs("PassWord").value
    rs.Close

    If SystemOptions.UserInterface = ArabicInterface Then
        Msg = " „  ⁄œÌ· ﬂ·„… «·„—Ê— »‰Ã«Õ"
    Else
        Msg = "Password Changed"
    End If

    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    Unload Me
    Exit Sub
ErrTrap:

    If BeginTrans = True Then
        BeginTrans = False
        Cn.RollbackTrans
    End If

    If rs.State = adStateOpen Then
        If rs.EditMode <> adEditNone Then
            rs.CancelUpdate
        End If
    End If

    If Err.Number = -2147217900 Then
        Msg = "·« Ì„ﬂ‰ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & Chr(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… «·»Ì«‰«  Ê√⁄œ «·„Õ«Ê·…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

    Msg = "⁄›Ê«...ÕœÀ Œÿ√ „« √À‰«¡ Õ›Ÿ Â–Â «·»Ì«‰«  " & Chr(13)
    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
End Sub

