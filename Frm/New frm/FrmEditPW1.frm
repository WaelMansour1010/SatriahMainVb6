VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form FrmEditPW1 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " €ÌÌ— «·—ﬁ„ «·”—Ì"
   ClientHeight    =   2310
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5955
   HelpContextID   =   260
   Icon            =   "FrmEditPW1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   2310
   ScaleWidth      =   5955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FrameAdmin 
      Caption         =   "»Ì«‰«  «·„‘—›"
      Height          =   2295
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   6015
      Begin VB.TextBox XPTxtPass1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   120
         MaxLength       =   20
         PasswordChar    =   "*"
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   600
         Width           =   3465
      End
      Begin MSDataListLib.DataCombo DCboUserName1 
         Height          =   360
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   3465
         _ExtentX        =   6112
         _ExtentY        =   635
         _Version        =   393216
         BackColor       =   16761024
         ForeColor       =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ALLButtonS.ALLButton CmdLogin 
         Default         =   -1  'True
         Height          =   375
         Left            =   2160
         TabIndex        =   14
         Top             =   1080
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   " ”ÃÌ· «·œŒÊ·"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   16711680
         BCOLO           =   16711680
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmEditPW1.frx":038A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Image Image3 
         Height          =   495
         Left            =   3840
         Picture         =   "FrmEditPW1.frx":03A6
         Stretch         =   -1  'True
         Top             =   480
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "ﬂ·„… «·„—Ê—"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   9
         Left            =   4200
         TabIndex        =   16
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Labelx 
         Alignment       =   1  'Right Justify
         Caption         =   "«”„ «·„‘—›"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   10
         Left            =   4200
         TabIndex        =   15
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.TextBox XPTxtOldPW 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   6090
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
      Top             =   810
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
      Top             =   1230
      Width           =   2415
   End
   Begin ImpulseButton.ISButton XPBtnOK 
      Height          =   375
      Left            =   900
      TabIndex        =   3
      Top             =   1830
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
      ButtonImage     =   "FrmEditPW1.frx":0A83
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
      Height          =   375
      Left            =   60
      TabIndex        =   4
      Top             =   1830
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
      ButtonImage     =   "FrmEditPW1.frx":0E1D
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
      Top             =   1830
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
      ButtonImage     =   "FrmEditPW1.frx":11B7
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
      Left            =   5850
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   885
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
      Top             =   825
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
      Top             =   1245
      Width           =   1365
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " €ÌÌ— ﬂ·„… «·„—Ê— ··„” Œœ„"
      Height          =   315
      Index           =   0
      Left            =   1380
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   30
      Width           =   1875
   End
   Begin VB.Label XPLblUserName 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   315
      Left            =   360
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   390
      Width           =   3735
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   0
      Left            =   3990
      Picture         =   "FrmEditPW1.frx":1551
      Top             =   60
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   1
      Left            =   6150
      Picture         =   "FrmEditPW1.frx":18DB
      Top             =   870
      Width           =   240
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   6300
      X2              =   -60
      Y1              =   1710
      Y2              =   1710
   End
End
Attribute VB_Name = "FrmEditPW1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

 

Private Sub CmdHelp_Click()
    SystemOptions.SysHelp.HHTopicID = Me.HelpContextID
    SystemOptions.SysHelp.HHDisplayTopicID Me.hWnd
End Sub

Private Sub CMDLogin_Click()
On Error Resume Next
    'On Error GoTo ErrTrap
    If DCboUserName1.Text = "" Then
        
If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "ÌÃ» «œŒ«· «”„ «·„‘—›"
 Else
 Msg = "Select dmin User Name"
 
 
 End If
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        DCboUserName.SetFocus
        Exit Sub
    End If
 
 

    StrSQL = "Select * From cachierData Where id=" & Me.DCboUserName1.BoundText & " AND password='" & Trim(Me.XPTxtPass1.Text) & "'"
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.EOF Or rs.BOF Then
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = " √ﬂœ „‰ ’Õ… «”„ «·„‘—› " & Chr(13)
                Msg = Msg + "Êﬂ·„… «·„—Ê— Ê√⁄œ «·„Õ«Ê·…"
            Else
            
            Msg = "User Name Or Password Incorrect " & Chr(13)
            End If
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        DCboUserName1.SetFocus
        Exit Sub
    End If
SystemOptions.usertype = UserAdminAll
If SystemOptions.UserInterface = ArabicInterface Then
      Msg = "‰„  ”ÃÌ·  œŒÊ·  «·„‘—› " & Chr(13)
 Else
 Msg = "Admin was Login Success" & Chr(13)
 End If
 
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        
     AddToLogFile CInt(user_id), 0, Date, Time, "  ”ÃÌ· «·œŒÊ· ·‰›ÿ… «·»Ì⁄ »√”„ «·„‘—›  " & DCboUserName1.Text, " System Login", Me.Name, "L", "", ""
   Unload Me
    Exit Sub
ErrTrap:



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
    Me.Caption = "Change PW "
    lbl(0).Caption = Me.Caption & " For User : "
    lbl(4).Caption = "Old PW"
FrameAdmin.Caption = "Admin Login"
Labelx(10).Caption = "User Name"
Labelx(9).Caption = "PassWord"
CMDLogin.Caption = "Login"

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
Public Function LoadAdmins(PointID As Double)
Dim My_SQL As String

  If SystemOptions.UserInterface = ArabicInterface Then
        My_SQL = "SELECT     dbo.cachierData.id, dbo.cachierData.name"
    My_SQL = My_SQL & " FROM         dbo.cachierData  "
   My_SQL = My_SQL & " WHERE     (dbo.cachierData.Ctype = 1)"
   
    My_SQL = My_SQL & " and  PointId =" & PointID
    Else
    
            My_SQL = "SELECT     dbo.cachierData.id, dbo.cachierData.namee"
    My_SQL = My_SQL & " FROM         dbo.cachierData  "
   My_SQL = My_SQL & " WHERE     (dbo.cachierData.Ctype = 1)"
   My_SQL = My_SQL & " and  PointId =" & PointID

    End If
    
        fill_combo DCboUserName1, My_SQL
        
End Function
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

    If XPTxtOldPW.Text = "" Then
        '    Msg = "√œŒ· ﬂ·„… «·„—Ê— «·”«»ﬁ…"
        '    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        '    XPTxtOldPW.SetFocus
        '    Exit Sub
    End If

'    If StrComp(XPTxtOldPW.text, User_Password, vbTextCompare) <> 0 Then
'        Msg = "ﬂ·„… «·„—Ê— «·ﬁœÌ„… €Ì— ’ÕÌÕ…"
'        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'        XPTxtOldPW.SetFocus
'        Exit Sub
'    End If

    If XPTXTNewPw.Text = "" Then
        Msg = "√œŒ· ﬂ·„… «·„—Ê— «·ÃœÌœ…"
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        XPTXTNewPw.SetFocus
        Exit Sub
    End If

    If XPTxtConfirm.Text = "" Then
        Msg = "√œŒ·  √ﬂÌœ ﬂ·„… «·„—Ê— "
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        XPTxtConfirm.SetFocus
        Exit Sub
    End If

    If StrComp(XPTXTNewPw.Text, XPTxtConfirm.Text, vbTextCompare) <> 0 Then
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
    rs("PassWord").value = Trim(XPTXTNewPw.Text)
    rs("PassConfirm").value = Trim(XPTxtConfirm.Text)
    
    rs("ChangePW").value = 0
    
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

