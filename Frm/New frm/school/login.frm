VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form login 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ãÏÑÓÉ ÇáãäÇåÌ"
   ClientHeight    =   1785
   ClientLeft      =   2370
   ClientTop       =   4245
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "login.frx":0000
   ScaleHeight     =   1785
   ScaleWidth      =   6000
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
   Begin VB.TextBox txtpassword 
      Height          =   405
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   840
      Width           =   1935
   End
   Begin VB.TextBox txtusername 
      Height          =   405
      Left            =   1920
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
   Begin VB.TextBox txtrecived 
      Height          =   855
      Left            =   360
      TabIndex        =   6
      Top             =   4200
      Width           =   2295
   End
   Begin VB.CommandButton cmdsend 
      Caption         =   "Send"
      Height          =   255
      Left            =   3000
      TabIndex        =   5
      Top             =   3480
      Width           =   1095
   End
   Begin VB.TextBox txtsend 
      Height          =   855
      Left            =   360
      TabIndex        =   4
      Top             =   3120
      Width           =   2295
   End
   Begin VB.TextBox serverip 
      Height          =   375
      Left            =   4800
      TabIndex        =   3
      Text            =   "127.0.0.1"
      Top             =   3120
      Width           =   2295
   End
   Begin VB.CommandButton cmdconnect 
      Caption         =   "ÏÎæá"
      Default         =   -1  'True
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   1320
      Width           =   1335
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   360
      Top             =   3480
      Width           =   7320
      _ExtentX        =   12912
      _ExtentY        =   1296
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=MAY"
      OLEDBString     =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=MAY"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "users"
      Caption         =   " "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "ÇÓã ÇáãÓÊÎÏã"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   375
      Left            =   2280
      TabIndex        =   7
      Top             =   0
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "ßáãÉ ÇáãÑæÑ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   375
      Left            =   2280
      TabIndex        =   8
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label4"
      ForeColor       =   &H00FFFFFF&
      Height          =   1815
      Left            =   1800
      TabIndex        =   9
      Top             =   0
      Width           =   2175
   End
End
Attribute VB_Name = "login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdconnect_Click()
On Error GoTo ll
Adodc1.CommandType = adCmdText
Adodc1.RecordSource = "select * from users where username='" & TXTUSERNAME.Text & "' and password='" & txtpassword.Text & "'"
  Adodc1.Refresh
  
    If Adodc1.Recordset.RecordCount > 0 Then
    
        main.Show
        main.TXTUSERNAME = Me.TXTUSERNAME
        main.user_id.Caption = Adodc1.Recordset.Fields!user_id
      alarm_frm.Show
        
        Unload Me
    Else
        MsgBox "Invalid Password, try again!", , "Login"
        
    End If
ll:
End Sub

Private Sub Label5_Click()
IMAGE_PATH_FRM.Show
End Sub
