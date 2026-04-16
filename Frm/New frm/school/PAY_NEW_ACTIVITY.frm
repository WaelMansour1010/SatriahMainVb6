VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form PAY_NEW_ACTIVITY 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "     ‘«‘… ”œ« œ «·«‰‘ÿ…"
   ClientHeight    =   3645
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   9660
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   3645
   ScaleWidth      =   9660
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
   Begin VB.TextBox Text19 
      Alignment       =   2  'Center
      DataField       =   "USER_NAME"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   4440
      TabIndex        =   17
      Top             =   3120
      Width           =   2175
   End
   Begin VB.TextBox Text9 
      Alignment       =   2  'Center
      DataField       =   "bill_no"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   4440
      TabIndex        =   15
      Top             =   2640
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF0000&
      Caption         =   "”œ«œ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3120
      Width           =   3135
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      DataField       =   "member_type"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   1920
      Width           =   2295
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      DataField       =   "OPERATION_TYPE"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   720
      Width           =   2295
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      DataField       =   "OPERATION_NO"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   240
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      DataField       =   "OPERATION_DATE"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   240
      Width           =   2295
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      DataField       =   "MEMBER_NAME"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   2040
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      DataField       =   "MEMBER_ID"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   1440
      Width           =   2172
   End
   Begin VB.TextBox Text18 
      Alignment       =   2  'Center
      DataField       =   "ACTIVITY_VALUE"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   960
      Width           =   2295
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   435
      Left            =   0
      Top             =   2640
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   767
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      RecordSource    =   "select  *  from operations where PAYED=0  and (operation_type='‰‘«ÿ ÃœÌœ'  or operation_type=' ÃœÌœ ‰‘«ÿ')"
      Caption         =   "«·«‰ ﬁ«· »Ì‰ «·⁄„·Ì« "
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
   Begin MSAdodcLib.Adodc Adodc20 
      Height          =   375
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   661
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
      RecordSource    =   "OPERATIONS"
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
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      Caption         =   "»Ê«”ÿ…"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   7680
      TabIndex        =   18
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "—ﬁ„ «·«Ì’«·"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   612
      Left            =   6720
      TabIndex        =   16
      Top             =   2520
      Width           =   3972
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "‰Ê⁄ «·⁄÷ÊÌ…"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   612
      Left            =   2400
      TabIndex        =   13
      Top             =   1920
      Width           =   1812
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "‰Ê⁄ «·⁄„·Ì…"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   612
      Left            =   7560
      TabIndex        =   11
      Top             =   720
      Width           =   1452
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   " —ﬁ„ «·⁄·„Ì…"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   372
      Left            =   6720
      TabIndex        =   10
      Top             =   240
      Width           =   3012
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   " «—ÌŒ «·⁄„·Ì…"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   612
      Left            =   2640
      TabIndex        =   9
      Top             =   240
      Width           =   1452
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "«”„ «·ÿ«·»"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   612
      Left            =   6480
      TabIndex        =   8
      Top             =   1920
      Width           =   3972
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   " —ﬁ„ «·ÿ«·»"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   372
      Left            =   6840
      TabIndex        =   7
      Top             =   1440
      Width           =   3012
   End
   Begin VB.Label Label25 
      Alignment       =   2  'Center
      Caption         =   "ﬁÌ„… «·‰‘«ÿ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   492
      Left            =   2280
      TabIndex        =   6
      Top             =   960
      Width           =   1932
   End
End
Attribute VB_Name = "PAY_NEW_ACTIVITY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

If Text9.Text = "" Then
MsgBox "  ·«»œ „‰ ﬂ «»… —ﬁ„ «·«Ì’«· «Ê·«  ", vbCritical
Text9.BackColor = &HFF&
Text9.SetFocus
Exit Sub
End If

Adodc20.CommandType = adCmdText
Adodc20.RecordSource = "select * from operations where bill_no=" & Text9.Text
Adodc20.Refresh

If Adodc20.Recordset.RecordCount > 0 Then
MsgBox "—ﬁ„ «·«Ì’«·  „” Œœ„ „‰ ﬁ»·", vbCritical
Exit Sub
End If



If Adodc1.Recordset.RecordCount = 0 Then Exit Sub
Adodc1.Recordset.Fields!payed = 1
Adodc1.Recordset.Fields!ACTUAL_VALUE = Text18.Text
Adodc1.Recordset.Update
Adodc1.Refresh
End Sub

Private Sub Text9_Change()
If Text9.Text = "" Then Exit Sub
If Not IsNumeric(Text9.Text) Then
MsgBox "  —ﬁ„ «·«Ì’«· „ﬂÊ‰ „‰ «—ﬁ«„ ›ﬁÿ    ", vbCritical
Text9.Text = ""
Exit Sub
End If

End Sub
