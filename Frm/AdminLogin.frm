VERSION 5.00
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form AdminLogin 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ÕœÌÀ «·‰Ÿ«„"
   ClientHeight    =   1260
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   5820
   ControlBox      =   0   'False
   Icon            =   "AdminLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   744.449
   ScaleMode       =   0  'User
   ScaleWidth      =   5464.666
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.TextBox txtfuncid 
      Height          =   375
      Left            =   480
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   1320
      Width           =   615
   End
   Begin VB.TextBox txtUserName 
      Alignment       =   2  'Center
      Height          =   345
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   1605
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   6720
      PasswordChar    =   "*"
      TabIndex        =   2
      Text            =   "stars123"
      Top             =   4680
      Width           =   1725
   End
   Begin ALLButtonS.ALLButton cmdOK 
      Height          =   615
      Left            =   1920
      TabIndex        =   4
      Top             =   600
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   " ÕœÌÀ"
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   16711680
      BCOLO           =   12582912
      FCOL            =   16777215
      FCOLO           =   0
      MCOL            =   192
      MPTR            =   1
      MICON           =   "AdminLogin.frx":6852
      PICN            =   "AdminLogin.frx":686E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   1
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ALLButtonS.ALLButton cmdCancel 
      Height          =   615
      Left            =   240
      TabIndex        =   5
      Top             =   600
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   "Œ—ÊÃ"
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   16711680
      BCOLO           =   12582912
      FCOL            =   16777215
      FCOLO           =   0
      MCOL            =   192
      MPTR            =   1
      MICON           =   "AdminLogin.frx":D0D0
      PICN            =   "AdminLogin.frx":D0EC
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   1
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin ALLButtonS.ALLButton ALLButton1 
      Cancel          =   -1  'True
      Height          =   615
      Left            =   2400
      TabIndex        =   6
      Top             =   1560
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   "Full update"
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
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   16711680
      BCOLO           =   12582912
      FCOL            =   16777215
      FCOLO           =   0
      MCOL            =   192
      MPTR            =   1
      MICON           =   "AdminLogin.frx":1394E
      PICN            =   "AdminLogin.frx":1396A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   1
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Image Image2 
      Height          =   1080
      Left            =   4200
      Picture         =   "AdminLogin.frx":1A1CC
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1560
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "«Œ—  ÕœÌÀ ··»—‰«„Ã » «—ÌŒ"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   0
      Left            =   1680
      TabIndex        =   3
      Top             =   120
      Width           =   2505
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&Password:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   1
      Left            =   8640
      TabIndex        =   1
      Top             =   4800
      Width           =   1320
   End
   Begin VB.Image Image1 
      Height          =   2085
      Left            =   -3120
      Picture         =   "AdminLogin.frx":1C90C
      Top             =   -2880
      Width           =   3885
   End
End
Attribute VB_Name = "AdminLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LoginSucceeded As Boolean
Dim funcid As Integer

Private Sub ALLButton1_Click()

    'check for correct password
    If TxtPassWord = "stars123" Then
        DB_CreateField "TblCarBillMentains", "AuthoOrder", adInteger, adColNullable, 255, , "C?C??   ", False, True, , True
        UpdateDataBase
        Me.txtUserName = getLastDataBaseUpdateDate
        
        '  Unload Me
    Else
        MsgBox "Invalid Password, try again!", , "Login"
        '        txtpassword.SetFocus
        '        SendKeys "{Home}+{End}"
    End If


End Sub

Private Sub CmdCancel_Click()
'            If CurrentVersion <> getLastDataBaseUpdateDate Then
'            If SystemOptions.UserInterface = ArabicInterface Then
'        MsgBox "·«»œ „‰  ÕœÌÀ «·‰”Œ…", vbCritical
'        Else
'        MsgBox "Must Update Data base", vbCritical
'        End If
'        End
   
        
         
      '  End If
        
    'set the global var to false
    'to denote a failed login
    LoginSucceeded = False
    Me.Hide
    mdifrmmain.Enabled = True
     Exit Sub
   
End Sub
Private Sub CmdOk_Click()
'Dim funcid As Integer

    'check for correct password
    If TxtPassWord = "stars123" Then
    'chek
    
   If funcid = 0 Then
   UpdateDataBase
   ElseIf funcid = 11 Then
   UpdateDataBasePart11
   
      ElseIf funcid = 12 Then
   UpdateDataBasePart12
      ElseIf funcid = 13 Then
   UpdateDataBasePart13
         ElseIf funcid = 14 Then
   UpdateDataBasePart14
         ElseIf funcid = 15 Then
   UpdateDataBasePart15
         ElseIf funcid = 16 Then
         UpdateDataBasePart15
'   UpdateDataBasePart16
         ElseIf funcid = 17 Then
         UpdateDataBasePart17
         ElseIf funcid = 18 Then
         UpdateDataBasePart18
         ElseIf funcid = 19 Then
         UpdateDataBasePart19
         

   
   End If
   
      
        
        Me.txtUserName = getLastDataBaseUpdateDate(funcid)
  txtfuncid.Text = funcid
  
    mdifrmmain.Enabled = True
  
        '  Unload Me
    Else
        MsgBox "Invalid Password, try again!", , "Login"
        '        txtpassword.SetFocus
        '        SendKeys "{Home}+{End}"
    End If


End Sub

Private Sub Form_Activate()
PutFormOnTop Me.hWnd, True
mdifrmmain.Enabled = False
End Sub

Private Sub Form_Load()
    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500

 '   Me.txtUserName = getLastDataBaseUpdateDate
  Me.txtUserName = getLastDataBaseUpdateDate(funcid)
  txtfuncid.Text = funcid
      If SystemOptions.UserInterface = EnglishInterface Then
    SetInterface Me
    ChangeLang
    End If
    
End Sub
Private Sub ChangeLang()

   Me.Caption = "Update Database"
    lblLabels(0).Caption = "Last Update "
    CmdOk.Caption = "Update"
     CmdCancel.Caption = "Exit"
End Sub
Private Sub Image2_Click()
UpdateDataBase
End Sub
