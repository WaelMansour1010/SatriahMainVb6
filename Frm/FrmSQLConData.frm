VERSION 5.00
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Begin VB.Form FrmSQLConData 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "»Ì«‰«  «·√ ’«· »ﬁ«⁄œ… «·»Ì«‰« "
   ClientHeight    =   4755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4755
   Icon            =   "FrmSQLConData.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   4755
   ScaleWidth      =   4755
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ ﬁ«⁄œ… «·»Ì«‰« "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1965
      Index           =   1
      Left            =   570
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   1980
      Width           =   4065
      Begin VB.TextBox TxtDisplayname 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Text            =   "«·ﬁ«⁄œ… «·«”«”Ì…"
         Top             =   240
         Width           =   2535
      End
      Begin VB.TextBox TxtPassword 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   240
         PasswordChar    =   "*"
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Text            =   "Admin@123"
         Top             =   1320
         Width           =   2535
      End
      Begin VB.TextBox TxtUserID 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Text            =   "sa"
         Top             =   960
         Width           =   2535
      End
      Begin VB.TextBox TxtDbName 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Text            =   "byte"
         Top             =   600
         Width           =   2535
      End
      Begin VB.ComboBox CboDataBaseName 
         Height          =   315
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   2160
         Visible         =   0   'False
         Width           =   2835
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·«”„ «·Ÿ«Â—"
         Height          =   375
         Index           =   4
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   360
         Width           =   915
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "PassWord"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   1440
         Width           =   915
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "User Id"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   1080
         Width           =   915
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«”„ «·ﬁ«⁄œÂ"
         Height          =   375
         Index           =   1
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   720
         Width           =   915
      End
   End
   Begin ImpulseButton.ISButton CmdTest 
      Height          =   375
      Left            =   2820
      TabIndex        =   4
      Top             =   4260
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   661
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "≈Œ »«— «·« ’«·"
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
      ButtonImage     =   "FrmSQLConData.frx":058A
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·ÃÂ«“ «·”Ì—›—"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1905
      Index           =   0
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   30
      Width           =   4545
      Begin VB.Frame Frame1 
         Height          =   495
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   1320
         Width           =   3495
         Begin VB.OptionButton optserver 
            Alignment       =   1  'Right Justify
            Caption         =   "SQL Auth."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   1
            Left            =   360
            RightToLeft     =   -1  'True
            TabIndex        =   13
            Top             =   240
            Width           =   1215
         End
         Begin VB.OptionButton optserver 
            Alignment       =   1  'Right Justify
            Caption         =   "Windows Auth."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   0
            Left            =   1680
            RightToLeft     =   -1  'True
            TabIndex        =   12
            Top             =   240
            Value           =   -1  'True
            Width           =   1695
         End
      End
      Begin VB.TextBox TxtServerName 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H00FF0000&
         Height          =   345
         Left            =   210
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Text            =   "."
         Top             =   870
         Width           =   2865
      End
      Begin VB.OptionButton opt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÃÂ«“ √Œ— „ÊÃÊœ ⁄·Ï «·‘»ﬂ…"
         ForeColor       =   &H00FF0000&
         Height          =   405
         Index           =   1
         Left            =   1350
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   570
         Width           =   3135
      End
      Begin VB.OptionButton opt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·ÃÂ«“ «·„Õ·Ï ( «·ÃÂ«“ «·„ÊÃÊœ ⁄·ÌÂ «·»—‰«„Ã)"
         Height          =   375
         Index           =   0
         Left            =   870
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   3615
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈ﬂ » «”„ «·”Ì—›—"
         ForeColor       =   &H00FF0000&
         Height          =   405
         Index           =   0
         Left            =   2850
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   960
         Width           =   1515
      End
   End
   Begin ImpulseButton.ISButton ISBXPBtnOK 
      Height          =   375
      Left            =   960
      TabIndex        =   5
      Top             =   4230
      Width           =   915
      _ExtentX        =   1614
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
      ButtonImage     =   "FrmSQLConData.frx":0B24
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      LowerToggledContent=   0   'False
      ColorTextShadow =   4210752
   End
   Begin ImpulseButton.ISButton ISBXPBtnCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   60
      TabIndex        =   6
      Top             =   4230
      Width           =   885
      _ExtentX        =   1561
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
      ButtonImage     =   "FrmSQLConData.frx":0EBE
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      LowerToggledContent=   0   'False
      ColorTextShadow =   4210752
   End
   Begin VB.Line LinLine1 
      BorderWidth     =   2
      X1              =   4725
      X2              =   0
      Y1              =   3990
      Y2              =   4005
   End
End
Attribute VB_Name = "FrmSQLConData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim m_UserCanceled As Boolean

Private Sub CmdTest_Click()
    TestSetting
End Sub

Private Sub Form_Load()
    CenterForm Me
    Opt_Click 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)

    If UnloadMode <> QueryUnloadConstants.vbFormCode Then
        Me.UserCanceled = True
    End If

End Sub

Private Sub ISBXPBtnCancel_Click()
    Me.UserCanceled = True
user_id = -1
    Me.Hide
    End
End Sub

Private Sub ISBXPBtnOK_Click()

    If TestSetting = True Then
        Me.UserCanceled = False
          save_login_info1 Trim(Me.TxtDbName.Text), Trim(Me.TxtDisplayname.Text), TxtServerName
        Me.Hide
        user_id = -1
        Unload Me
        
'End
    End If

End Sub

Private Sub Opt_Click(Index As Integer)
    Me.TxtServerName.Enabled = opt(1).value
    Me.lbl(0).Enabled = opt(1).value
 If Index = 1 Then
    optserver(1).value = True
    
 End If
End Sub

Public Property Get UserCanceled() As Boolean
    UserCanceled = m_UserCanceled
End Property

Public Property Let UserCanceled(ByVal vNewValue As Boolean)
    m_UserCanceled = vNewValue
End Property

Private Function TestSetting() As Boolean
    Dim TestCon As ADODB.Connection
    Dim Msg As String
    Dim StrConn As String
    On Error GoTo ErrTrap

    If Trim(Me.TxtDbName.Text) = "" Then
        Msg = "ÌÃ» ﬂ «»… «”„ «·”Ì—›— ﬁ«⁄œ… «·»Ì«‰«  "
 
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Function
    End If

    SystemOptions.SysSQLServerDataBaseName = (Me.TxtDbName.Text)
         SaveSetting "Byte_DBS", "Setting", "DBname", TxtDisplayname
     SystemOptions.SysSQLServerUserId = (Me.TxtUserID.Text)
     SystemOptions.SysSQLServerUserpassword = (Me.TxtPassword.Text)
     
    If Me.opt(1).value = True Then
        If Trim(Me.TxtServerName.Text) = "" Then
            Msg = "ÌÃ» ﬂ «»… «”„ «·”Ì—›— «·„ÊÃÊœ " & Chr(13) & "⁄·ÌÂ ﬁ«⁄œ… «·»Ì«‰« "
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Exit Function
        End If
  
        '  MsgBox "salim" & SystemOptions.SysSQLServerDataBaseName
        Set TestCon = New ADODB.Connection

        With TestCon
            .CursorLocation = adUseClient
            .ConnectionTimeout = 30

            If Me.optserver(0).value = True Then
                .ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI; " & "Persist Security Info=False;Initial Catalog=" & SystemOptions.SysSQLServerDataBaseName & ";Data Source=" & (Me.TxtServerName.Text) & ""
            Else
            'TxtPassword
            'TxtUserID
            
                .ConnectionString = "Provider=SQLOLEDB.1;Password=" & SystemOptions.SysSQLServerUserpassword & ";Persist Security Info=True;User ID=" & SystemOptions.SysSQLServerUserId & ";Initial Catalog=" & SystemOptions.SysSQLServerDataBaseName & ";Data Source=" & (Me.TxtServerName.Text)
            End If

            .Open
        End With

    Else
        Set TestCon = New ADODB.Connection

        With TestCon
            .CursorLocation = adUseClient
            .ConnectionTimeout = 30
            .ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI; " & "Persist Security Info=False;Initial Catalog=" & SystemOptions.SysSQLServerDataBaseName & ";Data Source=(LOCAL)"
            .Open
        End With

    End If

    If Me.opt(0).value = True Then
        SaveSetting SystemOptions.SysRegsAppPath, "ServerCon", "ServerType", 1
        SaveSetting SystemOptions.SysRegsAppPath, "ServerCon", "ServerName", "(LOCAL)"
    ElseIf Me.opt(1).value = True Then
        SaveSetting SystemOptions.SysRegsAppPath, "ServerCon", "ServerType", 2
        SaveSetting SystemOptions.SysRegsAppPath, "ServerCon", "ServerName", Trim(Me.TxtServerName.Text)

        If Me.optserver(0).value = True Then 'win
            SaveSetting SystemOptions.SysRegsAppPath, "ServerCon", "SysSQLServerTypeTechnical", 0
        Else
            SaveSetting SystemOptions.SysRegsAppPath, "ServerCon", "SysSQLServerTypeTechnical", 1
            
              
        End If
    
    End If

    SystemOptions.SysSQLServerType = val(GetSetting(SystemOptions.SysRegsAppPath, "ServerCon", "ServerType", 0))
    SystemOptions.SysSQLServerName = GetSetting(SystemOptions.SysRegsAppPath, "ServerCon", "ServerName", "")

    SystemOptions.SysSQLServerTypeTechnical = GetSetting(SystemOptions.SysRegsAppPath, "ServerCon", "SysSQLServerTypeTechnical", "0")


      Msg = " „  ⁄„·Ì… «·√ ’«· »‰Ã«Õ"
    MsgBox Msg, vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    TestCon.Close
    Set TestCon = Nothing
    TestSetting = True
    Exit Function
ErrTrap:
    Msg = "›‘·  ⁄„·Ì… «·√ ’«·.."
    Msg = Msg & Chr(13) & "Description:" & Err.description
    Msg = Msg & Chr(13) & "Number:" & Err.Number
    Msg = Msg & Chr(13) & "Source:" & Err.Source
    MsgBox Msg, vbCritical + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
    TestSetting = False
End Function

Private Sub TxtPassword_GotFocus()
TxtPassword.Text = ""
End Sub

Private Sub TxtServerName_GotFocus()
TxtServerName.Text = ""
End Sub
