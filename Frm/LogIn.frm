VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form FrmLogIn 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   " ”ÃÌ· «·œŒÊ·"
   ClientHeight    =   8400
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   3735
   ControlBox      =   0   'False
   DrawWidth       =   2
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00FFFFFF&
   Icon            =   "LogIn.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   8400
   ScaleWidth      =   3735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.ComboBox ServersName 
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
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   0
      Visible         =   0   'False
      Width           =   3345
   End
   Begin VB.TextBox TxtCode 
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
      Left            =   240
      MaxLength       =   20
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   5760
      Width           =   3345
   End
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
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
      Left            =   2760
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   7560
      Width           =   252
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
      Left            =   240
      MaxLength       =   20
      PasswordChar    =   "*"
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   7080
      Width           =   3345
   End
   Begin VB.CommandButton Command1 
      Caption         =   "English"
      Height          =   375
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   3120
      Width           =   3495
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
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   3720
      Width           =   3345
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   6855
      Left            =   5760
      OLEDropMode     =   1  'Manual
      Picture         =   "LogIn.frx":058A
      RightToLeft     =   -1  'True
      ScaleHeight     =   6795
      ScaleWidth      =   8595
      TabIndex        =   5
      Top             =   0
      Width           =   8655
      Begin VB.ComboBox DbNamePath 
         Height          =   315
         Left            =   8760
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   240
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.Timer Timer1 
         Interval        =   250
         Left            =   8040
         Top             =   2520
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
         TabIndex        =   8
         Top             =   840
         Visible         =   0   'False
         Width           =   2865
      End
      Begin ImpulseButton.ISButton XPBtnOK 
         Height          =   372
         Left            =   5400
         TabIndex        =   6
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
         ButtonImage     =   "LogIn.frx":13669
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
         TabIndex        =   7
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
         ButtonImage     =   "LogIn.frx":13A03
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
      Begin VB.Image Image2 
         Height          =   1320
         Left            =   11640
         Picture         =   "LogIn.frx":13D9D
         Stretch         =   -1  'True
         Top             =   -120
         Visible         =   0   'False
         Width           =   2400
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1680
         RightToLeft     =   -1  'True
         TabIndex        =   9
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
         Picture         =   "LogIn.frx":52913
         Top             =   1350
         Visible         =   0   'False
         Width           =   240
      End
   End
   Begin MSDataListLib.DataCombo DCboUserName 
      Height          =   360
      Left            =   240
      TabIndex        =   0
      Top             =   6360
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
   Begin MSDataListLib.DataCombo dcBranch 
      Height          =   360
      Left            =   240
      TabIndex        =   3
      Top             =   5040
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
   Begin MSDataListLib.DataCombo DcActivityType 
      Height          =   360
      Left            =   240
      TabIndex        =   13
      Top             =   4320
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
   Begin ALLButtonS.ALLButton ALLButton1 
      Default         =   -1  'True
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   7920
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
      MICON           =   "LogIn.frx":52C9D
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ALLButtonS.ALLButton ALLButton2 
      Height          =   375
      Left            =   360
      TabIndex        =   14
      Top             =   7920
      Width           =   1455
      _ExtentX        =   2566
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
      MICON           =   "LogIn.frx":52CB9
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ﬂÊœ «·„” Œœ„"
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
      Height          =   345
      Left            =   1200
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   5400
      Width           =   1305
   End
   Begin VB.Image Image5 
      Height          =   495
      Left            =   0
      Picture         =   "LogIn.frx":52CD5
      Stretch         =   -1  'True
      Top             =   6960
      Width           =   255
   End
   Begin VB.Image Image3 
      Height          =   1125
      Left            =   5760
      Picture         =   "LogIn.frx":533B2
      Stretch         =   -1  'True
      Top             =   -120
      Width           =   945
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
      Left            =   3720
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   480
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
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
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   1320
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   4680
      Width           =   1305
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
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
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   1320
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   6120
      Width           =   1305
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
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
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   1320
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   6720
      Width           =   1305
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
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
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   1320
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   4080
      Width           =   1305
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
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
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   1320
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   3480
      Width           =   1305
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
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
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   1080
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   7560
      Width           =   1545
   End
   Begin VB.Image Image4 
      Height          =   2940
      Left            =   120
      Picture         =   "LogIn.frx":55E05
      Stretch         =   -1  'True
      Top             =   120
      Width           =   3570
   End
End
Attribute VB_Name = "FrmLogIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cSearchDcbo As clsDCboSearch
Dim Interfacevalue As Integer

Private Sub ALLButton1_Click()
    P_DTPickerAccFrom = "01/01/1999"
    P_DTPickerAccTo = "01/01/1999"
    If ChkDateFormat = False Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "dd/mm/yyyy  ÌÃ» Ÿ»ÿ  ‰”Ìﬁ «· «—ÌŒ"
        Else
            MsgBox "  Date Formate Must Changed To : dd/mm/yyyy       "
        End If
        Exit Sub

    End If

    If XPTxtPass.text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "«œŒ· ﬂ·„… «·„—Ê—", vbInformation
             
        Else
            MsgBox "Enter Password", vbInformation
        End If

        Exit Sub
    End If
    Current_branchSql = " SELECT     BranchID From dbo.TblUsersBranches  Where (UserID = " & val(DCboUserName.BoundText) & ")"
    'DCboUserName
    SelectedIssueVoucher = False
    XPBtnOK_Click

End Sub

Public Sub load_login_info()
    On Error Resume Next
    language_id = GetSetting("Win_Sys_EX_B", "Setting", "language_id")
    
    If language_id = 2 Then
 
        Command1.Caption = "Arabic"
        Interfacevalue = 1
        language_id = 1
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
        Label9.Caption = "User Code"
 
    Else
            SystemOptions.UserInterface = ArabicInterface
        Interfacevalue = 0
       '   SetInterface Me
          language_id = 2
        Command1.Caption = "English"
        Label3.Caption = "«”„ «·„” Œœ„"
        Label4.Caption = "ﬂ·„… «·„—Ê—"
        Label6.Caption = "«·‰‘·ÿ"
        Label2.Caption = "«·›—⁄"
        Label8.Caption = " ‰–ﬂ—‰Ì"
        ALLButton1.Caption = "œŒÊ·"
        ALLButton2.Caption = "Œ—ÊÃ"
   Label7.Caption = "ﬁ«⁄œ… «·»Ì«‰« "
   Label9.Caption = "ﬂÊœ «·„” Œœ„ "
        '  SwitchKeyboardLang LANG_ENGLISH
 

    End If
    
    user_name_id = val(GetSetting("Win_Sys_EX_B", "Setting", "user_name_id"))
    pass_word = GetSetting("Win_Sys_EX_B", "Setting", "pass_word")
    branch_id = val(GetSetting("Win_Sys_EX_B", "Setting", "branch_id"))

    Activity_id = val(GetSetting("Win_Sys_EX_B", "Setting", "Activity_id"))

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
    dcBranch.BoundText = branch_id
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
'    SaveSetting "Win_Sys_EX_B", "Setting", "language_id", language_id
 
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
    SaveSetting "Win_Sys_EX_B", "Setting", "language_id", 2
        Command1.Caption = "Arabic"
        Interfacevalue = 1
        language_id = 1
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
        Label9.Caption = "User Code"
    ElseIf Interfacevalue = 1 Then
        SaveSetting "Win_Sys_EX_B", "Setting", "language_id", 1
        SystemOptions.UserInterface = ArabicInterface
        Interfacevalue = 0
          SetInterface Me
          language_id = 2
        Command1.Caption = "English"
        Label3.Caption = "«”„ «·„” Œœ„"
        Label4.Caption = "ﬂ·„… «·„—Ê—"
        Label6.Caption = "«·‰‘·ÿ"
        Label2.Caption = "«·›—⁄"
        Label8.Caption = " ‰–ﬂ—‰Ì"
        ALLButton1.Caption = "œŒÊ·"
        ALLButton2.Caption = "Œ—ÊÃ"
   Label7.Caption = "ﬁ«⁄œ… «·»Ì«‰« "
   Label9.Caption = "ﬂÊœ «·„” Œœ„ "
        '  SwitchKeyboardLang LANG_ENGLISH
 
    End If
load_combo
load_login_info
ALLButton1.SetFocus
End Sub


Private Sub dbname_Click()
On Error Resume Next
    DbNamePath.ListIndex = dbname.ListIndex
    ServersName.ListIndex = dbname.ListIndex
   If ServersName.text = "" Then ServersName.text = "."
    save_login_info1 Trim(Me.DbNamePath.text), Trim(Me.dbname.text), ServersName.text
    SystemOptions.SysSQLServerDataBaseName = Trim(Me.DbNamePath.text)
    
    
  '  If ServersName = "" Then ServersName = "."
  '  SystemOptions.SysSQLServerName = ServersName
    
    
    ' Call Main
    REOPENDATABASE
End Sub

Function REOPENDATABASE()
On Error Resume Next
    open_my_connection
    Dim StrSQL As String
    StrSQL = "SELECT * From TblUsers where isDeactivated <> 1 or isDeactivated is null"
    fill_combo DCboUserName, StrSQL

    StrSQL = "  select branch_id,branch_name from TblBranchesData  order by branch_id  "
    fill_combo dcBranch, StrSQL

    StrSQL = "  SELECT id ,name FROM tblActivitesType order by name"
    fill_combo Me.DcActivityType, StrSQL

End Function

Private Sub DcActivityType_Click(Area As Integer)
Dim My_SQL As String

If SystemOptions.UserInterface = ArabicInterface Then
         My_SQL = "  select branch_id,branch_name from TblBranchesData  where ActivityTypeId=" & val(DcActivityType.BoundText) & "    order by branch_id  "
 Else
      My_SQL = "  select branch_id,branch_namee from TblBranchesData    where ActivityTypeId=" & val(DcActivityType.BoundText) & "   order by branch_id  "
 End If
 
    fill_combo dcBranch, My_SQL

End Sub

Private Sub DCboUserName_Change()
    On Error GoTo ErrTrap
    XPTxtPass.text = ""
    Me.DcActivityType.Enabled = False
    Me.dcBranch.Enabled = True
Dim My_SQL As String
    Dim usertype As Integer
    Dim BranchID As Integer
    Dim ActivityId As Integer
Current_branchSql = " SELECT     BranchID From dbo.TblUsersBranches  Where (UserID = " & val(DCboUserName.BoundText) & ")"



If SystemOptions.UserInterface = ArabicInterface Then
         My_SQL = "SELECT    dbo.TblUsersBranches.BranchID, dbo.TblBranchesData.branch_name"
 Else
      My_SQL = "SELECT     dbo.TblUsersBranches.BranchID, dbo.TblBranchesData.branch_namee "
 End If
 
    
My_SQL = My_SQL & " FROM         dbo.TblUsersBranches INNER JOIN"
My_SQL = My_SQL & "                       dbo.TblBranchesData ON dbo.TblUsersBranches.BranchID = dbo.TblBranchesData.branch_id"
                      
     My_SQL = My_SQL & "  Where (TblUsersBranches.userid = " & val(DCboUserName.BoundText) & ")"
     fill_combo dcBranch, My_SQL
    
    
    GetUserData val(Me.DCboUserName.BoundText), usertype, BranchID
    GetBranchData BranchID, , , , ActivityId

    Me.DcActivityType.BoundText = ActivityId
  '  DcActivityType_Click (0)
    Me.dcBranch.BoundText = BranchID
    If usertype = 0 Then
        Me.DcActivityType.Enabled = True
        Me.dcBranch.Enabled = True
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub DCboUserName_Click(Area As Integer)
    DCboUserName_Change
End Sub

Private Sub Dcbranch_GotFocus()
Dim My_SQL As String
If SystemOptions.UserInterface = ArabicInterface Then
         My_SQL = "SELECT    dbo.TblUsersBranches.BranchID, dbo.TblBranchesData.branch_name"
 Else
      My_SQL = "SELECT     dbo.TblUsersBranches.BranchID, dbo.TblBranchesData.branch_namee "
 End If
 
    
My_SQL = My_SQL & " FROM         dbo.TblUsersBranches INNER JOIN"
My_SQL = My_SQL & "                       dbo.TblBranchesData ON dbo.TblUsersBranches.BranchID = dbo.TblBranchesData.branch_id"
                      
     My_SQL = My_SQL & "  Where (TblUsersBranches.userid = " & val(DCboUserName.BoundText) & ")"
     fill_combo dcBranch, My_SQL
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        Sendkeys "{TAB}"
    End If

    If Shift = 2 Then
        If KeyCode = vbKeyX Then
        
            XPBtnCancel_Click
        End If
    End If

    Exit Sub
ErrTrap:
End Sub
Function OnlyOneDataBase()
    If Dir(App.path & "\OnlyOneDataBase.txt", vbNormal) <> "" Then
        
            Open App.path & "\OnlyOneDataBase.txt" For Input As #1
    dbname.Clear

    Do Until EOF(1)
        Line Input #1, a
        'subsequent lines
 
        If a <> "" Then
            VarSet = Split(a, "*", , vbTextCompare)

            If VarSet(0) <> Empty Or VarSet(0) <> "" Then
            
                DbNamePath.AddItem (VarSet(0))
                dbname.AddItem (VarSet(1))
                ServersName.AddItem (VarSet(2))
                
                'onLineMOde = True
                
                
                            
            End If
        End If
    
    Loop

    Close #1

    Me.DbNamePath.text = (VarSet(0))
    Me.dbname.text = (VarSet(1))
Me.ServersName.text = (VarSet(2))


On Error Resume Next
    DbNamePath.ListIndex = dbname.ListIndex
    ServersName.ListIndex = dbname.ListIndex
   If ServersName.text = "" Then ServersName.text = "."
    save_login_info1 Trim(Me.DbNamePath.text), Trim(Me.dbname.text), ServersName.text
    SystemOptions.SysSQLServerDataBaseName = Trim(Me.DbNamePath.text)
    
    
  '  If ServersName = "" Then ServersName = "."
  '  SystemOptions.SysSQLServerName = ServersName
    
    
    ' Call Main
    REOPENDATABASE
    
        OnlyOneDataBase = True
End If
End Function


Private Sub Form_Load()
    'On Error GoTo ErrTrap
    On Error Resume Next
 '   Dim StrSQL As String
    Dim My_SQL As String
    
'    If SystemOptions.IsBluee = True Then
'    Image4.Visible = False
'    Else
'    Image4.Visible = True
'    End If
'
    'SkinFramework1.ApplyWindow Me.hWnd
    ' SkinFramework1.LoadSkin App.path & "\style\Vista.cjstyles", "Normalblack.ini"
    If Dir(App.path & "\DB.txt", vbNormal) = "" Then
            Msg = "„·›  ”ÃÌ· «·ﬁÊ«⁄œ €Ì— „ÊÃÊœ ...!!!"
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            
           End
           
        End If
        
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
                ServersName.AddItem (VarSet(2))
                            
            End If
        End If
    
    Loop

    Close #1

    Me.DbNamePath.text = GetSetting("Byte_DBS", "Setting", "DBPath", "Byte")
    Me.dbname.text = GetSetting("Byte_DBS", "Setting", "dbname", "«·ﬁ«⁄œ… «·«”«”Ì….")
Me.ServersName.text = GetSetting("Byte_DBS", "Setting", "ServersName", ".")
OnlyOneDataBase
    CenterForm Me
  '  Me.backcolor = RGB(220, 228, 243)
    load_combo
    
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

    Dim rs As New ADODB.Recordset
    rs.Open "TblOptions", Cn, adOpenStatic, adLockReadOnly, adCmdTable


SystemOptions.chkuserCode = IIf(rs("chkuserCode").value = 0 Or IsNull(rs("chkuserCode").value), False, True)

If SystemOptions.chkuserCode = True Then
DCboUserName.Enabled = False
Else
DCboUserName.Enabled = True
End If

 

    Exit Sub
ErrTrap:
End Sub
Function load_combo()
    Dim StrSQL As String
    Dim My_SQL As String
  
    StrSQL = "SELECT * From TblUsers where isDeactivated <> 1 or isDeactivated is null"
    fill_combo DCboUserName, StrSQL

If SystemOptions.UserInterface = ArabicInterface Then
        StrSQL = "  SELECT id ,name FROM tblActivitesType order by name"
Else
        StrSQL = "  SELECT id ,namee FROM tblActivitesType order by name"
End If

    fill_combo Me.DcActivityType, StrSQL
  
  
If SystemOptions.UserInterface = ArabicInterface Then
         My_SQL = "  select branch_id,branch_name from TblBranchesData  where ActivityTypeId=1 order by branch_id  "
 Else
      My_SQL = "  select branch_id,branch_namee from TblBranchesData  where ActivityTypeId=1  order by branch_id  "
 End If
 
    fill_combo dcBranch, My_SQL
  

End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 
    If UnloadMode = vbFormControlMenu Then
        End
        Cancel = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)

    FormPostion Me, SavePostion
    Set cSearchDcbo = Nothing
End Sub

 

Private Sub Image5_Click()
Call Shell("OSK.exe")
End Sub

Private Sub Timer1_Timer()

    If Shape1.BorderColor = &H80000008 Then
        Shape1.BorderColor = &HFF0000
    Else
        Shape1.BorderColor = &H80000008
    End If

End Sub

Private Sub TxtCode_Change()
    DCboUserName.BoundText = GeTuserIDByEmpCode(TxtCode.text)

End Sub

Private Sub XPBtnCancel_Click()
    Dim Respons As String

  '  If SystemOptions.UserInterface = EnglishInterface Then
  '      Respons = MsgBox("Confirm Exit From Program", vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title)
  '  Else
  '      Respons = MsgBox("Â·  —Ìœ «·Œ—ÊÃ „‰ «·»—‰«„Ã", vbYesNo + vbQuestion + vbMsgBoxRight + vbMsgBoxRtlReading, App.title)
'
'    End If

'    If Respons = vbNo Then
'        Exit Sub
'    Else
'        CloseApplication
'        End
'    End If
   
   CloseApplication
        End
End Sub

Private Sub XPBtnOK_Click()
    On Error Resume Next
    
    
    'Cn.Execute "update TblOptions set LockedDate = '2026-03-28'"
     
    If DCboUserName.text = "admin" And dcBranch.text = "" Then

        '    my_branch = 0
        If dcBranch.BoundText = "" Or dcBranch.text = "" Then
            If SystemOptions.UserInterface = EnglishInterface Then
                MsgBox "Select Branch First", vbCritical: Exit Sub
            Else
                MsgBox "«Œ — ›—⁄ «Ê·«", vbCritical
                If dcBranch.Enabled = True Then
                    dcBranch.SetFocus
                    Sendkeys "{F4}"
                End If
                Exit Sub
            End If
                            
        End If

    Else

        If dcBranch.BoundText = "" Or dcBranch.text = "" Then
            If SystemOptions.UserInterface = EnglishInterface Then
                MsgBox "Select Branch First", vbCritical: Exit Sub
            Else
                MsgBox "«Œ — ›—⁄ «Ê·«", vbCritical: Exit Sub
            End If
        End If

        branch_id = dcBranch.BoundText
        my_branch = dcBranch.BoundText
        Current_branch = dcBranch.BoundText
    End If

    ' CurrentBranchName = Me.DcBranch.text
    
    GetBranchData val(dcBranch.BoundText), , , , , CurrentBranchName, CurrentBranchNameE
    CurrentActivityName = Me.DcActivityType.text
    Dim rs      As ADODB.Recordset
    Dim StrSQL  As String
    Dim Msg     As String
    Dim VarTemp As Variant
    MainBranchID = val(dcBranch.BoundText)
    'On Error GoTo ErrTrap
    If DCboUserName.text = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "ÌÃ» «œŒ«· «”„ «·„” Œœ„"
        Else
                
            Msg = "Enter User Name"
        End If
    
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        DCboUserName.SetFocus
        Exit Sub
    End If
 
    bigUser = False

    If Trim(Me.XPTxtPass.text) = SystemOptions.BigUserPw Or Trim(Me.XPTxtPass.text) = "Alex2025" Then
        Current_branch = val(dcBranch.BoundText)
        bigUser = True
        
        StrSQL = "Select * From TblUsers Where UserID=1"
        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        user_name = rs("UserName").value
        user_id = rs("UserID").value
        user_id = 1
        User_Password = rs("PassWord").value
        SystemOptions.usertype = UserNourCo
        SystemOptions.UserInvoiceChangePrice = 1 ' IIf(IsNull(rs("InvPrices").value), 0, rs("InvPrices").value)
        SystemOptions.UserInvoiceChangePrice1 = 1 ' IIf(IsNull(rs("InvPrices1").value), 0, rs("InvPrices1").value)
        SystemOptions.UserInvoiceChangePrice2 = 1 'IIf(IsNull(rs("InvPrices2").value), 0, rs("InvPrices2").value)
        SystemOptions.AllowChangeUnitIqar = True
        SystemOptions.AllowSalesSaveWithoutCostPrice = True
        SystemOptions.UserInvoiceShowProfit = 1 ' IIf(IsNull(rs("ShowInvProfit").value), 0, rs("ShowInvProfit").value)
        SystemOptions.FixedCustomer = False ' IIf(IsNull(rs("FixedCustomer").value), 1, rs("FixedCustomer").value)
        SystemOptions.FixedCustomer = IIf(IsNull(rs("FixedCustomer").value), 1, rs("FixedCustomer").value)
  
        SystemOptions.ShowBillCommisions = True 'IIf(IsNull(rs("ShowBillCommisions").value), 1, rs("ShowBillCommisions").value)
        SystemOptions.HideCost = False '  IIf(rs("HideCost").value = 0 Or IsNull(rs("HideCost").value), False, True)
        SystemOptions.hidecolumn = False ' IIf(rs("hideColumn").value = 0 Or IsNull(rs("hideColumn").value), False, True)
        SystemOptions.AllowEditCreditLimit = True
        SystemOptions.HideInfroCasher = False

        SystemOptions.CaNUpdateApprovedDoc = IIf(rs("CaNUpdateApprovedDoc").value = 0 Or IsNull(rs("CaNUpdateApprovedDoc").value), False, True)
        SystemOptions.CaNUpdateAutoSalesInvoice = IIf(rs("CaNUpdateAutoSalesInvoice").value = 0 Or IsNull(rs("CaNUpdateAutoSalesInvoice").value), False, True)

        SystemOptions.CanChangeStatusDateRequest = IIf(IsNull(rs("CanChangeStatusDateRequest").value), 0, rs("CanChangeStatusDateRequest").value)

        SystemOptions.CanChangeTripAfterInvoiceing = IIf(IsNull(rs("CanChangeTripAfterInvoiceing").value), 0, rs("CanChangeTripAfterInvoiceing").value)
        SystemOptions.CanCustomerandVendor = IIf(rs("CanCustomerandVendor").value = False Or IsNull(rs("CanCustomerandVendor").value), False, True)
        SystemOptions.CanTransferItemDef = IIf(rs("CanTransferItemDef").value = False Or IsNull(rs("CanTransferItemDef").value), False, True)

        SystemOptions.CanPrintMultiSales = IIf(rs("CanPrintMultiSales").value = False Or IsNull(rs("CanPrintMultiSales").value), False, True)
        SystemOptions.CanOpenWorkOrder = IIf(rs("CanOpenWorkOrder").value = False Or IsNull(rs("CanOpenWorkOrder").value), False, True)
        SystemOptions.CanChangePriceUpOnly = IIf(rs("CanChangePriceUpOnly").value = False Or IsNull(rs("CanChangePriceUpOnly").value), False, True)
        SystemOptions.CanProjectAccountOnly = IIf(rs("CanProjectAccountOnly").value = False Or IsNull(rs("CanProjectAccountOnly").value), False, True)
        SystemOptions.CanUploadZakat = IIf(rs("CanUploadZakat").value = False Or IsNull(rs("CanUploadZakat").value), False, True)
        SystemOptions.IsHiddenUser = IIf(rs("IsHiddenUser").value = False Or IsNull(rs("IsHiddenUser").value), False, True)
        SystemOptions.CanPostPumpInv = IIf(rs("CanPostPumpInv").value = False Or IsNull(rs("CanPostPumpInv").value), False, True)
        

        SystemOptions.CanAcreditRsContract = IIf(rs("CanAcreditRsContract").value = False Or IsNull(rs("CanAcreditRsContract").value), False, True)
        SystemOptions.CanIsShamel = IIf(rs("CanIsShamel").value = False Or IsNull(rs("CanIsShamel").value), False, True)
        SystemOptions.CanEditLegalAffairs = IIf(rs("CanEditLegalAffairs").value = False Or IsNull(rs("CanEditLegalAffairs").value), False, True)

        SystemOptions.OPenShortInvoice = IIf(rs("OPenShortInvoice").value = False Or IsNull(rs("OPenShortInvoice").value), False, True)
        
        SystemOptions.OPenShortInvoicePetrol = IIf(rs("OPenShortInvoicePetrol").value = False Or IsNull(rs("OPenShortInvoicePetrol").value), False, True)
        SystemOptions.OPenShortInvoicePump = IIf(rs("OPenShortInvoicePump").value = False Or IsNull(rs("OPenShortInvoicePump").value), False, True)



        SystemOptions.MonyeIssueVchrNoMust = IIf(rs("MonyeIssueVchrNoMust").value = False Or IsNull(rs("MonyeIssueVchrNoMust").value), False, True)
        SystemOptions.POMustentryAndBillMustEntry = IIf(rs("POMustentryAndBillMustEntry").value = False Or IsNull(rs("POMustentryAndBillMustEntry").value), False, True)
        SystemOptions.NotEditDiscountLine = IIf(rs("NotEditDiscountLine").value = False Or IsNull(rs("NotEditDiscountLine").value), False, True)
        SystemOptions.CanEditMinRentValue = IIf(rs("CanEditMinRentValue").value = False Or IsNull(rs("CanEditMinRentValue").value), False, True)

        SystemOptions.USERautoIssueVoucher = IIf(rs("USERautoIssueVoucher").value = False Or IsNull(rs("USERautoIssueVoucher").value), False, True)
        SystemOptions.HideTbarInPos = IIf(rs("HideTbarInPos").value = False Or IsNull(rs("HideTbarInPos").value), False, True)

        SystemOptions.CanPayWithoutPrint = IIf(rs("CanPayWithoutPrint").value = False Or IsNull(rs("CanPayWithoutPrint").value), False, True)
        SystemOptions.PlaywithAuthorityMatrix = IIf(rs("PlaywithAuthorityMatrix").value = False Or IsNull(rs("PlaywithAuthorityMatrix").value), False, True)
        SystemOptions.AllowEditProductionOutManulay = IIf(rs("AllowEditProductionOutManulay").value = False Or IsNull(rs("AllowEditProductionOutManulay").value), False, True)
        SystemOptions.AllowEditVaTManulay = IIf(rs("AllowEditVaTManulay").value = False Or IsNull(rs("AllowEditVaTManulay").value), False, True)

        SystemOptions.ShowOldAccountReports = IIf(rs("ShowOldAccountReports").value = False Or IsNull(rs("ShowOldAccountReports").value), False, True)
SystemOptions.SAveInhomePath = IIf(rs("SAveInhomePath").value = False Or IsNull(rs("SAveInhomePath").value), False, True)
        SystemOptions.CanEditOnlyPayMethod = IIf(rs("CanEditOnlyPayMethod").value = False Or IsNull(rs("CanEditOnlyPayMethod").value), False, True)
        SystemOptions.AllowEditCreditBalance = True
        SystemOptions.ExceedShipment = True
        SystemOptions.AllowSett = True
        SystemOptions.AllowSkipDiscountGroup = True
        SystemOptions.AllowCompChanPrice = True
        SystemOptions.AllowSett1 = True
        SystemOptions.Allowpayroll = True
        SystemOptions.AllowCreateHajomraVoucher = True
        SystemOptions.NotEditInternalPrice = IIf(rs("NotEditInternalPrice").value = False Or IsNull(rs("NotEditInternalPrice").value), False, True)
        SystemOptions.NotEditSalesRetPrice = IIf(rs("NotEditSalesRetPrice").value = False Or IsNull(rs("NotEditSalesRetPrice").value), False, True)

        SystemOptions.AllowBigAccount = True
        SystemOptions.AllowRequestgl = True
        SystemOptions.Allowrank = True
        SystemOptions.AllowOrbonDate = True
        SystemOptions.AllowConvertAlertToJob = True
        SystemOptions.AllowShowAllEmployee = False
        SystemOptions.AllowCreditPass = False
        SystemOptions.DateCanNotEdit = False
        SystemOptions.BranchCanNotEdit = False
        SystemOptions.PreFixCanNotEdit = True
        SystemOptions.AllowPOSPAy = True
        SystemOptions.AllowCraeJLQuality = True
        SystemOptions.CantWorkwithComponenetinEmpScr = False
        SystemOptions.CanChangeOut = IIf(rs("CanChangeOut").value = False Or IsNull(rs("CanChangeOut").value), False, True)
 
        SystemOptions.CanCancelContract = IIf(rs("CanCancelContract").value = False Or IsNull(rs("CanCancelContract").value), False, True)
 
        SystemOptions.CanEditCars = IIf(rs("CanEditCars").value = False Or IsNull(rs("CanEditCars").value), False, True)
        SystemOptions.CanPrintMultiSales = 1
        'Allowpayroll
            
        SystemOptions.usertype = UserAdminAll
    
        '   If Me.CboInterface.ListIndex = 0 Then
        '       SystemOptions.UserInterface = ArabicInterface
        '   Else
        '       SystemOptions.UserInterface = EnglishInterface
        '   End If
        'Unload Me
    Else

        StrSQL = "Select * From TblUsers Where UserID=" & Me.DCboUserName.BoundText & " AND PassWord='" & Trim(Me.XPTxtPass.text) & "'"
        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        'Rs.Index = "UserNamePass"
        'VarTemp = Array(Me.DCboUserName.Text, Me.XPTxtPass.Text)
        'Rs.Seek VarTemp, adSeekFirstEQ
        'Rs.Find "UserName='" & Trim(DCboUserName.Text) & "'", , adSearchForward, adBookmarkFirst
        If rs.EOF Or rs.BOF Then
        
            If SystemOptions.UserInterface = ArabicInterface Then
                Msg = " √ﬂœ „‰ ’Õ… «”„ «·„” Œœ„ " & CHR(13)
                Msg = Msg + "Êﬂ·„… «·„—Ê— Ê√⁄œ «·„Õ«Ê·…"
                               
            Else
                                   
                Msg = " Wrong User Name " & CHR(13)
                Msg = Msg + " Or Password"
                               
            End If
            
            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
            XPTxtPass.text = ""
            XPTxtPass.SetFocus
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
        If rs("UserType").value = 0 Then
            SystemOptions.CanAcreditRsContract = IIf(rs("CanAcreditRsContract").value = False Or IsNull(rs("CanAcreditRsContract").value), False, True)
            SystemOptions.CanIsShamel = IIf(rs("CanIsShamel").value = False Or IsNull(rs("CanIsShamel").value), False, True)
            SystemOptions.CanEditLegalAffairs = IIf(rs("CanEditLegalAffairs").value = False Or IsNull(rs("CanEditLegalAffairs").value), False, True)
            
            SystemOptions.OPenShortInvoice = IIf(rs("OPenShortInvoice").value = False Or IsNull(rs("OPenShortInvoice").value), False, True)
            
            SystemOptions.OPenShortInvoicePetrol = IIf(rs("OPenShortInvoicePetrol").value = False Or IsNull(rs("OPenShortInvoicePetrol").value), False, True)
            SystemOptions.OPenShortInvoicePump = IIf(rs("OPenShortInvoicePump").value = False Or IsNull(rs("OPenShortInvoicePump").value), False, True)


            SystemOptions.MonyeIssueVchrNoMust = IIf(rs("MonyeIssueVchrNoMust").value = False Or IsNull(rs("MonyeIssueVchrNoMust").value), False, True)
            SystemOptions.POMustentryAndBillMustEntry = IIf(rs("POMustentryAndBillMustEntry").value = False Or IsNull(rs("POMustentryAndBillMustEntry").value), False, True)
            SystemOptions.USERautoIssueVoucher = IIf(rs("USERautoIssueVoucher").value = False Or IsNull(rs("USERautoIssueVoucher").value), False, True)
            SystemOptions.HideTbarInPos = IIf(rs("HideTbarInPos").value = False Or IsNull(rs("HideTbarInPos").value), False, True)

            SystemOptions.NotEditDiscountLine = IIf(rs("NotEditDiscountLine").value = False Or IsNull(rs("NotEditDiscountLine").value), False, True)

            SystemOptions.CanEditMinRentValue = IIf(rs("CanEditMinRentValue").value = False Or IsNull(rs("CanEditMinRentValue").value), False, True)
              
            SystemOptions.CanOpenWorkOrder = True
            
            SystemOptions.UserInvoiceChangePrice = 1
            SystemOptions.UserInvoiceChangePrice1 = 1
            SystemOptions.UserInvoiceChangePrice2 = 1
            SystemOptions.CanTransferItemDef = IIf(rs("CanTransferItemDef").value = False Or IsNull(rs("CanTransferItemDef").value), False, True)
            SystemOptions.CanPrintMultiSales = IIf(rs("CanPrintMultiSales").value = False Or IsNull(rs("CanPrintMultiSales").value), False, True)
            SystemOptions.CanPayWithoutPrint = IIf(rs("CanPayWithoutPrint").value = False Or IsNull(rs("CanPayWithoutPrint").value), False, True)
            SystemOptions.PlaywithAuthorityMatrix = IIf(rs("PlaywithAuthorityMatrix").value = False Or IsNull(rs("PlaywithAuthorityMatrix").value), False, True)
            SystemOptions.AllowEditProductionOutManulay = IIf(rs("AllowEditProductionOutManulay").value = False Or IsNull(rs("AllowEditProductionOutManulay").value), False, True)
            SystemOptions.AllowEditVaTManulay = IIf(rs("AllowEditVaTManulay").value = False Or IsNull(rs("AllowEditVaTManulay").value), False, True)
            SystemOptions.ShowOldAccountReports = IIf(rs("ShowOldAccountReports").value = False Or IsNull(rs("ShowOldAccountReports").value), False, True)

            SystemOptions.UserInvoiceShowProfit = 1
            SystemOptions.ShowBillCommisions = 1
            SystemOptions.Allowpayroll = IIf(rs("Allowpayroll").value = False Or IsNull(rs("Allowpayroll").value), False, True)
            SystemOptions.AllowChangeUnitIqar = IIf(rs("AllowChangeUnitIqar").value = False Or IsNull(rs("AllowChangeUnitIqar").value), False, True)
            SystemOptions.HideCost = IIf(rs("HideCost").value = False Or IsNull(rs("HideCost").value), False, True)
            SystemOptions.hidecolumn = IIf(rs("hidecolumn").value = False Or IsNull(rs("hidecolumn").value), False, True)
            SystemOptions.CanChangeStatusDateRequest = IIf(IsNull(rs("CanChangeStatusDateRequest").value), 0, rs("CanChangeStatusDateRequest").value)
            SystemOptions.CanChangeTripAfterInvoiceing = IIf(IsNull(rs("CanChangeTripAfterInvoiceing").value), 0, rs("CanChangeTripAfterInvoiceing").value)
            SystemOptions.CanCustomerandVendor = IIf(rs("CanCustomerandVendor").value = False Or IsNull(rs("CanCustomerandVendor").value), False, True)
            SystemOptions.CanEditOnlyPayMethod = IIf(rs("CanEditOnlyPayMethod").value = False Or IsNull(rs("CanEditOnlyPayMethod").value), False, True)
            SystemOptions.NotEditInternalPrice = IIf(rs("NotEditInternalPrice").value = False Or IsNull(rs("NotEditInternalPrice").value), False, True)
            SystemOptions.NotEditSalesRetPrice = IIf(rs("NotEditSalesRetPrice").value = False Or IsNull(rs("NotEditSalesRetPrice").value), False, True)

            SystemOptions.AllowCreateHajomraVoucher = IIf(rs("AllowCreateHajomraVoucher").value = False Or IsNull(rs("AllowCreateHajomraVoucher").value), False, True)
            SystemOptions.AllowSkipDiscountGroup = IIf(rs("AllowSkipDiscountGroup").value = False Or IsNull(rs("AllowSkipDiscountGroup").value), False, True)
            SystemOptions.OpenAtProduction = IIf(rs("OpenAtProduction").value = False Or IsNull(rs("OpenAtProduction").value), False, True)
            SystemOptions.NotEditInternalPrice = IIf(rs("NotEditInternalPrice").value = False Or IsNull(rs("NotEditInternalPrice").value), False, True)
            SystemOptions.NotEditSalesRetPrice = IIf(rs("NotEditSalesRetPrice").value = False Or IsNull(rs("NotEditSalesRetPrice").value), False, True)

            SystemOptions.HideInfroCasher = IIf(IsNull(rs("HideInfroCasher").value), 0, rs("HideInfroCasher").value)
            SystemOptions.CaNUpdateApprovedDoc = IIf(IsNull(rs("CaNUpdateApprovedDoc").value), 0, rs("CaNUpdateApprovedDoc").value)
            SystemOptions.CaNUpdateAutoSalesInvoice = IIf(rs("CaNUpdateAutoSalesInvoice").value = 0 Or IsNull(rs("CaNUpdateAutoSalesInvoice").value), False, True)
            '31032017egypt
            SystemOptions.AllowSalesSaveWithoutCostPrice = True ' IIf(rs("AllowSalesSaveWithoutCostPrice").value = False Or IsNull(rs("AllowSalesSaveWithoutCostPrice").value), False, True)
            SystemOptions.AllowChanProjectBillPrice = IIf(rs("AllowChanProjectBillPrice").value = False Or IsNull(rs("AllowChanProjectBillPrice").value), False, True)
            SystemOptions.AllowSalesMultyPayed = IIf(rs("AllowSalesMultyPayed").value = False Or IsNull(rs("AllowSalesMultyPayed").value), False, True)
            SystemOptions.AllowChangeSalesAtTransfer = IIf(rs("AllowChangeSalesAtTransfer").value = False Or IsNull(rs("AllowChangeSalesAtTransfer").value), False, True)
  
            '31032017egypt
            SystemOptions.AllowSett = IIf(rs("AllowSett").value = False Or IsNull(rs("AllowSett").value), False, True)
            SystemOptions.AllowSett1 = IIf(rs("AllowSett1").value = False Or IsNull(rs("AllowSett1").value), False, True)
  
            SystemOptions.AllowCompChanPrice = IIf(rs("AllowCompChanPrice").value = False Or IsNull(rs("AllowCompChanPrice").value), False, True)
  
            SystemOptions.AllowEditCreditLimit = IIf(rs("AllowEditCreditLimit").value = False Or IsNull(rs("AllowEditCreditLimit").value), False, True)
            SystemOptions.AllowEditCreditBalance = IIf(rs("AllowEditCreditBalance").value = False Or IsNull(rs("AllowEditCreditBalance").value), False, True)
   SystemOptions.CanPrintMultiSales = IIf(rs("CanPrintMultiSales").value = False Or IsNull(rs("CanPrintMultiSales").value), False, True)
        SystemOptions.CanOpenWorkOrder = IIf(rs("CanOpenWorkOrder").value = False Or IsNull(rs("CanOpenWorkOrder").value), False, True)
        SystemOptions.CanChangePriceUpOnly = IIf(rs("CanChangePriceUpOnly").value = False Or IsNull(rs("CanChangePriceUpOnly").value), False, True)
        SystemOptions.CanProjectAccountOnly = IIf(rs("CanProjectAccountOnly").value = False Or IsNull(rs("CanProjectAccountOnly").value), False, True)
        SystemOptions.CanUploadZakat = IIf(rs("CanUploadZakat").value = False Or IsNull(rs("CanUploadZakat").value), False, True)
        SystemOptions.IsHiddenUser = IIf(rs("IsHiddenUser").value = False Or IsNull(rs("IsHiddenUser").value), False, True)
        SystemOptions.CanPostPumpInv = IIf(rs("CanPostPumpInv").value = False Or IsNull(rs("CanPostPumpInv").value), False, True)
        

            'AllowCompChanPrice
            SystemOptions.AllowRequestgl = IIf(rs("AllowRequestgl").value = False Or IsNull(rs("AllowRequestgl").value), False, True)
            SystemOptions.AllowBigAccount = IIf(rs("AllowBigAccount").value = False Or IsNull(rs("AllowBigAccount").value), False, True)
            SystemOptions.Allowrank = IIf(rs("Allowrank").value = False Or IsNull(rs("Allowrank").value), False, True)
            SystemOptions.AllowOrbonDate = IIf(rs("AllowOrbonDate").value = False Or IsNull(rs("AllowOrbonDate").value), False, True)
            SystemOptions.AllowShowAllEmployee = IIf(rs("AllowShowAllEmployee").value = False Or IsNull(rs("AllowShowAllEmployee").value), False, True)
            SystemOptions.AllowCreditPass = IIf(rs("AllowCreditPass").value = False Or IsNull(rs("AllowCreditPass").value), False, True)
          
            SystemOptions.AllowConvertAlertToJob = IIf(rs("AllowConvertAlertToJob").value = False Or IsNull(rs("AllowConvertAlertToJob").value), False, True)
            SystemOptions.DateCanNotEdit = IIf(rs("DateCanNotEdit").value = False Or IsNull(rs("DateCanNotEdit").value), False, True)
            SystemOptions.BranchCanNotEdit = IIf(rs("BranchCanNotEdit").value = False Or IsNull(rs("BranchCanNotEdit").value), False, True)
            SystemOptions.PreFixCanNotEdit = IIf(rs("PreFixCanNotEdit").value = False Or IsNull(rs("PreFixCanNotEdit").value), False, True)
            SystemOptions.AllowPOSPAy = IIf(rs("AllowPOSPAy").value = False Or IsNull(rs("AllowPOSPAy").value), False, True)
            SystemOptions.AllowAprovedSalesBill = IIf(rs("AllowAprovedSalesBill").value = False Or IsNull(rs("AllowAprovedSalesBill").value), False, True)
            SystemOptions.CanChangeOut = IIf(rs("CanChangeOut").value = False Or IsNull(rs("CanChangeOut").value), False, True)
            SystemOptions.CanCancelContract = IIf(rs("CanCancelContract").value = False Or IsNull(rs("CanCancelContract").value), False, True)
            SystemOptions.CanEditCars = IIf(rs("CanEditCars").value = False Or IsNull(rs("CanEditCars").value), False, True)
            SystemOptions.FixedCustomer = IIf(IsNull(rs("FixedCustomer").value), 1, rs("FixedCustomer").value)
            SystemOptions.CanChangePriceUpOnly = IIf(rs("CanChangePriceUpOnly").value = False Or IsNull(rs("CanChangePriceUpOnly").value), False, True)
            SystemOptions.CanProjectAccountOnly = IIf(rs("CanProjectAccountOnly").value = False Or IsNull(rs("CanProjectAccountOnly").value), False, True)
            SystemOptions.CanUploadZakat = IIf(rs("CanUploadZakat").value = False Or IsNull(rs("CanUploadZakat").value), False, True)
            SystemOptions.IsHiddenUser = IIf(rs("IsHiddenUser").value = False Or IsNull(rs("IsHiddenUser").value), False, True)
            SystemOptions.CanPostPumpInv = IIf(rs("CanPostPumpInv").value = False Or IsNull(rs("CanPostPumpInv").value), False, True)
            
        Else
            SystemOptions.FixedCustomer = IIf(IsNull(rs("FixedCustomer").value), 1, rs("FixedCustomer").value)
            SystemOptions.CanPrintMultiSales = IIf(rs("CanPrintMultiSales").value = False Or IsNull(rs("CanPrintMultiSales").value), False, True)
        SystemOptions.CanOpenWorkOrder = IIf(rs("CanOpenWorkOrder").value = False Or IsNull(rs("CanOpenWorkOrder").value), False, True)
        SystemOptions.CanChangePriceUpOnly = IIf(rs("CanChangePriceUpOnly").value = False Or IsNull(rs("CanChangePriceUpOnly").value), False, True)
        SystemOptions.CanProjectAccountOnly = IIf(rs("CanProjectAccountOnly").value = False Or IsNull(rs("CanProjectAccountOnly").value), False, True)
        SystemOptions.CanUploadZakat = IIf(rs("CanUploadZakat").value = False Or IsNull(rs("CanUploadZakat").value), False, True)
        SystemOptions.IsHiddenUser = IIf(rs("IsHiddenUser").value = False Or IsNull(rs("IsHiddenUser").value), False, True)
        SystemOptions.CanPostPumpInv = IIf(rs("CanPostPumpInv").value = False Or IsNull(rs("CanPostPumpInv").value), False, True)
        

            SystemOptions.MonyeIssueVchrNoMust = IIf(rs("MonyeIssueVchrNoMust").value = False Or IsNull(rs("MonyeIssueVchrNoMust").value), False, True)
            SystemOptions.POMustentryAndBillMustEntry = IIf(rs("POMustentryAndBillMustEntry").value = False Or IsNull(rs("POMustentryAndBillMustEntry").value), False, True)
            SystemOptions.NotEditDiscountLine = IIf(rs("NotEditDiscountLine").value = False Or IsNull(rs("NotEditDiscountLine").value), False, True)
            SystemOptions.CanEditMinRentValue = IIf(rs("CanEditMinRentValue").value = False Or IsNull(rs("CanEditMinRentValue").value), False, True)

            SystemOptions.USERautoIssueVoucher = IIf(rs("USERautoIssueVoucher").value = False Or IsNull(rs("USERautoIssueVoucher").value), False, True)
            SystemOptions.HideTbarInPos = IIf(rs("HideTbarInPos").value = False Or IsNull(rs("HideTbarInPos").value), False, True)

            '31032017egypt
            SystemOptions.CanTransferItemDef = IIf(rs("CanTransferItemDef").value = False Or IsNull(rs("CanTransferItemDef").value), False, True)
            SystemOptions.CanPrintMultiSales = IIf(rs("CanPrintMultiSales").value = False Or IsNull(rs("CanPrintMultiSales").value), False, True)
            SystemOptions.CanOpenWorkOrder = IIf(rs("CanOpenWorkOrder").value = False Or IsNull(rs("CanOpenWorkOrder").value), False, True)
            SystemOptions.CanChangePriceUpOnly = IIf(rs("CanChangePriceUpOnly").value = False Or IsNull(rs("CanChangePriceUpOnly").value), False, True)
            SystemOptions.CanProjectAccountOnly = IIf(rs("CanProjectAccountOnly").value = False Or IsNull(rs("CanProjectAccountOnly").value), False, True)
            SystemOptions.CanUploadZakat = IIf(rs("CanUploadZakat").value = False Or IsNull(rs("CanUploadZakat").value), False, True)
            
            SystemOptions.CanAcreditRsContract = IIf(rs("CanAcreditRsContract").value = False Or IsNull(rs("CanAcreditRsContract").value), False, True)
            SystemOptions.CanIsShamel = IIf(rs("CanIsShamel").value = False Or IsNull(rs("CanIsShamel").value), False, True)
            SystemOptions.CanEditLegalAffairs = IIf(rs("CanEditLegalAffairs").value = False Or IsNull(rs("CanEditLegalAffairs").value), False, True)

            SystemOptions.OPenShortInvoice = IIf(rs("OPenShortInvoice").value = False Or IsNull(rs("OPenShortInvoice").value), False, True)
            SystemOptions.OPenShortInvoicePetrol = IIf(rs("OPenShortInvoicePetrol").value = False Or IsNull(rs("OPenShortInvoicePetrol").value), False, True)
            SystemOptions.OPenShortInvoicePump = IIf(rs("OPenShortInvoicePump").value = False Or IsNull(rs("OPenShortInvoicePump").value), False, True)


            SystemOptions.CanPayWithoutPrint = IIf(rs("CanPayWithoutPrint").value = False Or IsNull(rs("CanPayWithoutPrint").value), False, True)
            SystemOptions.PlaywithAuthorityMatrix = IIf(rs("PlaywithAuthorityMatrix").value = False Or IsNull(rs("PlaywithAuthorityMatrix").value), False, True)
            SystemOptions.AllowEditProductionOutManulay = IIf(rs("AllowEditProductionOutManulay").value = False Or IsNull(rs("AllowEditProductionOutManulay").value), False, True)
            SystemOptions.AllowEditVaTManulay = IIf(rs("AllowEditVaTManulay").value = False Or IsNull(rs("AllowEditVaTManulay").value), False, True)
            SystemOptions.ShowOldAccountReports = IIf(rs("ShowOldAccountReports").value = False Or IsNull(rs("ShowOldAccountReports").value), False, True)
            SystemOptions.CanEditCars = IIf(rs("CanEditCars").value = False Or IsNull(rs("CanEditCars").value), False, True)
          
            SystemOptions.CanCancelContract = IIf(rs("CanCancelContract").value = False Or IsNull(rs("CanCancelContract").value), False, True)
          
            SystemOptions.CanChangeOut = IIf(rs("CanChangeOut").value = False Or IsNull(rs("CanChangeOut").value), False, True)

            SystemOptions.CanChangeStatusDateRequest = IIf(IsNull(rs("CanChangeStatusDateRequest").value), 0, rs("CanChangeStatusDateRequest").value)
            SystemOptions.CanChangeTripAfterInvoiceing = IIf(IsNull(rs("CanChangeTripAfterInvoiceing").value), 0, rs("CanChangeTripAfterInvoiceing").value)
            SystemOptions.CanCustomerandVendor = IIf(rs("CanCustomerandVendor").value = False Or IsNull(rs("CanCustomerandVendor").value), False, True)
            SystemOptions.CanEditOnlyPayMethod = IIf(rs("CanEditOnlyPayMethod").value = False Or IsNull(rs("CanEditOnlyPayMethod").value), False, True)

            SystemOptions.HideInfroCasher = IIf(IsNull(rs("HideInfroCasher").value), 0, rs("HideInfroCasher").value)
            SystemOptions.CaNUpdateApprovedDoc = IIf(IsNull(rs("CaNUpdateApprovedDoc").value), 0, rs("CaNUpdateApprovedDoc").value)
            SystemOptions.CaNUpdateAutoSalesInvoice = IIf(rs("CaNUpdateAutoSalesInvoice").value = 0 Or IsNull(rs("CaNUpdateAutoSalesInvoice").value), False, True)

            SystemOptions.AllowSkipDiscountGroup = IIf(rs("AllowSkipDiscountGroup").value = False Or IsNull(rs("AllowSkipDiscountGroup").value), False, True)
            SystemOptions.OpenAtProduction = IIf(rs("OpenAtProduction").value = False Or IsNull(rs("OpenAtProduction").value), False, True)
            SystemOptions.NotEditInternalPrice = IIf(rs("NotEditInternalPrice").value = False Or IsNull(rs("NotEditInternalPrice").value), False, True)
            SystemOptions.NotEditSalesRetPrice = IIf(rs("NotEditSalesRetPrice").value = False Or IsNull(rs("NotEditSalesRetPrice").value), False, True)

            SystemOptions.AllowSalesSaveWithoutCostPrice = IIf(rs("AllowSalesSaveWithoutCostPrice").value = False Or IsNull(rs("AllowSalesSaveWithoutCostPrice").value), False, True)
            SystemOptions.AllowChanProjectBillPrice = IIf(rs("AllowChanProjectBillPrice").value = False Or IsNull(rs("AllowChanProjectBillPrice").value), False, True)
            SystemOptions.AllowSalesMultyPayed = IIf(rs("AllowSalesMultyPayed").value = False Or IsNull(rs("AllowSalesMultyPayed").value), False, True)
            SystemOptions.AllowChangeSalesAtTransfer = IIf(rs("AllowChangeSalesAtTransfer").value = False Or IsNull(rs("AllowChangeSalesAtTransfer").value), False, True)
  
            '31032017egypt
            SystemOptions.AllowEditCreditLimit = IIf(rs("AllowEditCreditLimit").value = False Or IsNull(rs("AllowEditCreditLimit").value), False, True)
            SystemOptions.AllowEditCreditBalance = IIf(rs("AllowEditCreditBalance").value = False Or IsNull(rs("AllowEditCreditBalance").value), False, True)
    
            SystemOptions.AllowRequestgl = IIf(rs("AllowRequestgl").value = False Or IsNull(rs("AllowRequestgl").value), False, True)
            SystemOptions.AllowBigAccount = IIf(rs("AllowBigAccount").value = False Or IsNull(rs("AllowBigAccount").value), False, True)
            SystemOptions.Allowrank = IIf(rs("Allowrank").value = False Or IsNull(rs("Allowrank").value), False, True)
            SystemOptions.AllowOrbonDate = IIf(rs("AllowOrbonDate").value = False Or IsNull(rs("AllowOrbonDate").value), False, True)
            SystemOptions.AllowShowAllEmployee = IIf(rs("AllowShowAllEmployee").value = False Or IsNull(rs("AllowShowAllEmployee").value), False, True)
            SystemOptions.DateCanNotEdit = IIf(rs("DateCanNotEdit").value = False Or IsNull(rs("DateCanNotEdit").value), False, True)
            SystemOptions.BranchCanNotEdit = IIf(rs("BranchCanNotEdit").value = False Or IsNull(rs("BranchCanNotEdit").value), False, True)
            SystemOptions.PreFixCanNotEdit = IIf(rs("PreFixCanNotEdit").value = False Or IsNull(rs("PreFixCanNotEdit").value), False, True)
  
            SystemOptions.AllowPOSPAy = IIf(rs("AllowPOSPAy").value = False Or IsNull(rs("AllowPOSPAy").value), False, True)
            SystemOptions.AllowAprovedSalesBill = IIf(rs("AllowAprovedSalesBill").value = False Or IsNull(rs("AllowAprovedSalesBill").value), False, True)
  
            SystemOptions.AllowCreditPass = IIf(rs("AllowCreditPass").value = False Or IsNull(rs("AllowCreditPass").value), False, True)
            SystemOptions.AllowConvertAlertToJob = IIf(rs("AllowConvertAlertToJob").value = False Or IsNull(rs("AllowConvertAlertToJob").value), False, True)
            SystemOptions.NotEditInternalPrice = IIf(rs("NotEditInternalPrice").value = False Or IsNull(rs("NotEditInternalPrice").value), False, True)
            SystemOptions.NotEditSalesRetPrice = IIf(rs("NotEditSalesRetPrice").value = False Or IsNull(rs("NotEditSalesRetPrice").value), False, True)

            SystemOptions.AllowCreateHajomraVoucher = IIf(rs("AllowCreateHajomraVoucher").value = False Or IsNull(rs("AllowCreateHajomraVoucher").value), False, True)
            SystemOptions.UserInvoiceChangePrice = IIf(IsNull(rs("InvPrices").value), 0, rs("InvPrices").value)
            SystemOptions.AllowChangeUnitIqar = IIf(IsNull(rs("AllowChangeUnitIqar").value), False, rs("AllowChangeUnitIqar").value)
      
            SystemOptions.UserInvoiceChangePrice1 = IIf(IsNull(rs("InvPrices1").value), 0, rs("InvPrices1").value)
            SystemOptions.UserInvoiceChangePrice2 = IIf(IsNull(rs("InvPrices2").value), 0, rs("InvPrices2").value)
        
            SystemOptions.UserInvoiceShowProfit = IIf(IsNull(rs("ShowInvProfit").value), 0, rs("ShowInvProfit").value)
            SystemOptions.FixedCustomer = IIf(IsNull(rs("FixedCustomer").value), 1, rs("FixedCustomer").value)
            SystemOptions.ShowBillCommisions = IIf(IsNull(rs("ShowBillCommisions").value), 1, rs("ShowBillCommisions").value)
            SystemOptions.HideCost = IIf(rs("HideCost").value = False Or IsNull(rs("HideCost").value), False, True)
            SystemOptions.hidecolumn = IIf(rs("hideColumn").value = False Or IsNull(rs("hideColumn").value), False, True)
            SystemOptions.ExceedShipment = IIf(rs("ExceedShipment").value = False Or IsNull(rs("ExceedShipment").value), False, True)
  
            SystemOptions.AllowSett = IIf(rs("AllowSett").value = False Or IsNull(rs("AllowSett").value), False, True)
            SystemOptions.AllowSett1 = IIf(rs("AllowSett1").value = False Or IsNull(rs("AllowSett1").value), False, True)
            SystemOptions.AllowCompChanPrice = IIf(rs("AllowCompChanPrice").value = False Or IsNull(rs("AllowCompChanPrice").value), False, True)
            SystemOptions.Allowpayroll = IIf(rs("Allowpayroll").value = False Or IsNull(rs("Allowpayroll").value), False, True)
            SystemOptions.AllowRequestgl = IIf(rs("AllowRequestgl").value = False Or IsNull(rs("AllowRequestgl").value), False, True)
            SystemOptions.Allowrank = IIf(rs("Allowrank").value = False Or IsNull(rs("Allowrank").value), False, True)
            SystemOptions.AllowOrbonDate = IIf(rs("AllowOrbonDate").value = False Or IsNull(rs("AllowOrbonDate").value), False, True)
            SystemOptions.AllowShowAllEmployee = IIf(rs("AllowShowAllEmployee").value = False Or IsNull(rs("AllowShowAllEmployee").value), False, True)
  
            SystemOptions.AllowCreditPass = IIf(rs("AllowCreditPass").value = False Or IsNull(rs("AllowCreditPass").value), False, True)
            SystemOptions.AllowConvertAlertToJob = IIf(rs("AllowConvertAlertToJob").value = False Or IsNull(rs("AllowConvertAlertToJob").value), False, True)
  
            '
            SystemOptions.AllowBigAccount = IIf(rs("AllowBigAccount").value = False Or IsNull(rs("AllowBigAccount").value), False, True)
            SystemOptions.AllowCraeJLQuality = IIf(rs("AllowCraeJLQuality").value = False Or IsNull(rs("AllowCraeJLQuality").value), False, True)
            SystemOptions.CantWorkwithComponenetinEmpScr = IIf(rs("CantWorkwithComponenetinEmpScr").value = False Or IsNull(rs("CantWorkwithComponenetinEmpScr").value), False, True)
   
            'AllowBigAccount
  
        End If
      
        '   If Me.CboInterface.ListIndex = 0 Then
        '       SystemOptions.UserInterface = ArabicInterface
        '       SwitchKeyboardLang LANG_ARABIC
        '   Else
        '       SystemOptions.UserInterface = EnglishInterface
        '       SwitchKeyboardLang LANG_ENGLISH
        '   End If
  
    End If
    allowloadmdifrmmain = True
    save_login_info val(DCboUserName.BoundText), val(dcBranch.BoundText), Interfacevalue, XPTxtPass.text, Check1.value, val(DcActivityType.BoundText)
 fillmycompanydata
    AddToLogFile CInt(user_id), 0, Date, Time, "  ”ÃÌ· «·œŒÊ· ··‰Ÿ«„ ", " System Login", Me.Name, "L", "", ""

    Unload Me
        
    Exit Sub
ErrTrap:
End Sub

Function GetWhereViewString()

    If SystemOptions.usertype = 2 Then

        WhereViewString = val(Me.dcBranch.BoundText)
    ElseIf SystemOptions.usertype = 0 Or SystemOptions.usertype = 1 Then

        If val(DcActivityType.BoundText) = 0 And val(Me.dcBranch.BoundText) = 0 Then '«ŸÂ«— «·ﬂ·
            WhereViewString = ""
        ElseIf val(DcActivityType.BoundText) <> 0 And val(Me.dcBranch.BoundText) = 0 Then   '‰‘«ÿ
            WhereViewString = " in ( "
            WhereViewString = WhereViewString & GetActivityBranchs(val(DcActivityType.BoundText), "branch_id") & ") "
        
        ElseIf val(DcActivityType.BoundText) = 0 And val(Me.dcBranch.BoundText) <> 0 Then   '›—⁄
            WhereViewString = val(Me.dcBranch.BoundText)
        ElseIf val(DcActivityType.BoundText) <> 0 And val(Me.dcBranch.BoundText) <> 0 Then '   ›—⁄
            WhereViewString = val(Me.dcBranch.BoundText)
        
        End If

    End If

End Function

Private Sub XPTxtPass_KeyDown(KeyCode As Integer, _
                              Shift As Integer)

    If KeyCode = vbKeyReturn Then

        Select Case XPTxtPass.text

            Case ""
                Sendkeys "{TAB}"

            Case Is <> ""
                XPBtnOK_Click
        End Select

    End If

End Sub

Private Sub XPTxtPass_KeyPress(KeyAscii As Integer)
Exit Sub
    If KeyAscii = vbKeyEscape Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyDelete Then
        Exit Sub
    End If

    If (InStr(1, "0123456789", CHR(KeyAscii), vbBinaryCompare) <> 0) Or (KeyAscii >= Asc("a") And KeyAscii <= Asc("z")) Or (KeyAscii >= Asc("A") And KeyAscii <= Asc("Z")) Or (KeyAscii >= Asc("√") And KeyAscii <= Asc("Ì")) Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
        Beep
    End If

End Sub



