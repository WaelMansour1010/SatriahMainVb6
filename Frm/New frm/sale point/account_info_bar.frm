VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Begin VB.Form b 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   0  'None
   Caption         =   "—"
   ClientHeight    =   1980
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4950
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   1980
   ScaleWidth      =   4950
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
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1095
      Left            =   120
      TabIndex        =   17
      Top             =   600
      Visible         =   0   'False
      Width           =   1455
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Account Bance"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   0
         TabIndex        =   21
         Top             =   0
         Width           =   1215
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Credit"
         Height          =   375
         Left            =   240
         TabIndex        =   20
         Top             =   600
         Width           =   2535
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Depit"
         Height          =   375
         Left            =   240
         TabIndex        =   19
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Balance"
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   840
         Width           =   2535
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1215
      Left            =   3120
      TabIndex        =   16
      Top             =   480
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   4920
      Top             =   240
   End
   Begin MSAdodcLib.Adodc Adodc6 
      Height          =   495
      Left            =   -480
      Top             =   3960
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
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
   Begin VB.Shape Shape1 
      Height          =   1815
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   4695
   End
   Begin VB.Label inventory_name 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Left            =   4440
      TabIndex        =   15
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label item_name 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "xxxxxxxxxxxxxx"
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
      Height          =   495
      Left            =   3960
      TabIndex        =   14
      Top             =   3000
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   0
      Width           =   255
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "—Ì«·"
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   360
      TabIndex        =   12
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label d3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "??"
      Height          =   255
      Left            =   1440
      TabIndex        =   11
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "—Ì«·"
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   1200
      TabIndex        =   10
      Top             =   2760
      Width           =   495
   End
   Begin VB.Label d2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   255
      Left            =   1440
      TabIndex        =   9
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "—Ì«·"
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   2280
      TabIndex        =   8
      Top             =   2760
      Width           =   495
   End
   Begin VB.Label d1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   255
      Left            =   1440
      TabIndex        =   7
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "«·—’Ìœ «·Õ«·Ì"
      Height          =   255
      Left            =   3240
      TabIndex        =   6
      Top             =   1440
      Width           =   2535
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "„œÌ‰"
      Height          =   375
      Left            =   3360
      TabIndex        =   5
      Top             =   840
      Width           =   2535
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "œ«∆‰"
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Top             =   1200
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "—’Ìœ «·Õ”«»"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   3120
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label inventory_id 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      Height          =   255
      Left            =   3360
      TabIndex        =   2
      Top             =   2760
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label item_code 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "xxxxxxxxxxxxx"
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
      Height          =   255
      Left            =   1200
      TabIndex        =   1
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "„⁄·Ê„«  ⁄‰ Õ”«»"
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   0
      Width           =   2055
   End
End
Attribute VB_Name = "b"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim first_run  As Boolean

Private Sub Form_Activate()

    If first_run = True Then
        d3 = get_balance(item_code.Caption)
        first_run = False
    End If

End Sub

Private Sub Form_Load()
    connection_string = Cn.ConnectionString
    Adodc6.ConnectionString = connection_string
    Adodc6.CommandType = adCmdText
    first_run = True
    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500

End Sub

Function get_balance(account_serial As String) As Double
    Dim total_credit As Double
    Dim total_depit As Double
    Dim total As Double
    total_credit = 0: total_depit = 0: total = 0
    Adodc6.RecordSource = "select sum(DEV_Value) As total_credit from RptLedger_Sub where  Credit_Or_Debit=0 and  Account_Serial='" & account_serial & "'"
    Adodc6.Refresh

    If Not IsNull(Adodc6.Recordset.Fields!total_credit) Then
        total_credit = Adodc6.Recordset.Fields!total_credit
    Else
        total_credit = 0
    End If

    d2 = total_credit

    Adodc6.RecordSource = "select sum(DEV_Value) As total_depit from RptLedger_Sub where    Credit_Or_Debit=1 and  Account_Serial='" & account_serial & "'"
    Adodc6.Refresh

    If Not IsNull(Adodc6.Recordset.Fields!total_depit) Then
        total_depit = Adodc6.Recordset.Fields!total_depit
    Else
        total_depit = 0
    End If

    d1 = total_depit
    'Total = total_credit - total_depit
    get_balance = total_credit - total_depit
    d3 = get_balance

    '1 œ∆«∆‰
    '2„œÌ‰
End Function

Private Sub item_code_Change()
    d3 = get_balance(item_code.Caption)
End Sub

Private Sub Timer1_Timer()
    Unload Me
End Sub
