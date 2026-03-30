VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form ADD_MEMBER_FINES 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "     «÷«›… €—«„… ⁄·Ï ⁄÷Ê „⁄Ì‰"
   ClientHeight    =   5580
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   7245
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5580
   ScaleWidth      =   7245
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "  ÿ»Ìﬁ €—«„… ⁄·Ï  «·“ÊÃ…"
      Enabled         =   0   'False
      Height          =   855
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   4680
      Width           =   2295
   End
   Begin VB.TextBox Text9 
      DataField       =   "MEMBER_ID"
      DataSource      =   "Adodc5"
      Height          =   615
      Left            =   8280
      TabIndex        =   22
      Text            =   "Text9"
      Top             =   7080
      Width           =   495
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  'Center
      DataField       =   "VALUE_OF_MEMBER"
      DataSource      =   "Adodc3"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3120
      TabIndex        =   19
      Text            =   "0"
      Top             =   1320
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000FF00&
      Caption         =   "ÃœÌœ"
      Height          =   612
      Left            =   360
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   3240
      Width           =   2172
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      DataField       =   "Fine_DATE"
      DataSource      =   "Adodc3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   3120
      TabIndex        =   16
      Top             =   4320
      Width           =   2295
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      DataField       =   "Fine_nO_OF_YEAR"
      DataSource      =   "Adodc3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   3120
      TabIndex        =   14
      Top             =   3960
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      DataField       =   "FINES_TOTAL"
      DataSource      =   "Adodc3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   3480
      Visible         =   0   'False
      Width           =   2295
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "FINES.frx":0000
      DataField       =   "FINES_TYPE"
      DataSource      =   "Adodc3"
      Height          =   315
      Left            =   3000
      TabIndex        =   11
      Top             =   1800
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "Fines_NAME"
      Text            =   "DataCombo1"
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      DataField       =   "Fines_PERCENT"
      DataSource      =   "Adodc3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   9
      Text            =   "0"
      Top             =   2880
      Width           =   2295
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      DataField       =   "Fines_VALUE"
      DataSource      =   "Adodc3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "0"
      Top             =   2280
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      DataField       =   "MEMBER_ID"
      DataSource      =   "Adodc3"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3120
      TabIndex        =   3
      Top             =   240
      Width           =   2295
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      DataField       =   "MEMBER_NAME"
      DataSource      =   "Adodc3"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   840
      Width           =   2295
   End
   Begin VB.CommandButton Command3 
      Caption         =   "»ÕÀ"
      Enabled         =   0   'False
      Height          =   855
      Left            =   1800
      Picture         =   "FINES.frx":0015
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "  ÿ»Ìﬁ €—«„… ⁄·Ï «·⁄÷Ê «·«”«”Ì"
      Enabled         =   0   'False
      Height          =   855
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   3840
      Width           =   2295
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   0
      Top             =   6240
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   873
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
      RecordSource    =   "select *  FROM MEMBERS where MEMBER_ID =0"
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
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   7200
      Top             =   1800
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
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
      RecordSource    =   "SELECT  * FROM Fine_TYPES WHERE Fines_ID=0"
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
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   495
      Left            =   1080
      Top             =   6480
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   873
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
      RecordSource    =   "SELECT * FROM FINES WHERE MEMBER_ID=0"
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
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   330
      Left            =   480
      Top             =   2280
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
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
      RecordSource    =   "Fine_TYPES"
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
   Begin MSAdodcLib.Adodc Adodc5 
      Height          =   495
      Left            =   1080
      Top             =   7200
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   873
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
      RecordSource    =   "FINES_DETAILS"
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
   Begin MSAdodcLib.Adodc Adodc6 
      Height          =   330
      Left            =   720
      Top             =   2760
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
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
      RecordSource    =   "Fine_TYPES"
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
   Begin MSAdodcLib.Adodc Adodc7 
      Height          =   492
      Left            =   3240
      Top             =   6120
      Width           =   1332
      _ExtentX        =   2355
      _ExtentY        =   873
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
      RecordSource    =   ""
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
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2280
      TabIndex        =   21
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "«·«‘ —«ﬂ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   20
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   " «—ÌŒ «·€—«„…"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5520
      TabIndex        =   17
      Top             =   4320
      Width           =   1815
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "⁄œœ «·œ›⁄« "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5400
      TabIndex        =   15
      Top             =   3960
      Width           =   1815
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "«·ﬁÌ„… «·›⁄·Ì…"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5400
      TabIndex        =   13
      Top             =   3480
      Width           =   1815
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "«·‰”»…"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5400
      TabIndex        =   10
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "‰Ê⁄ «·€—«„…"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5400
      TabIndex        =   8
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "«·ﬁÌ„…"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5400
      TabIndex        =   7
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   " —ﬁ„ «·⁄÷Ê"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   6
      Top             =   240
      Width           =   3015
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "«”„ «·⁄÷Ê"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4320
      TabIndex        =   5
      Top             =   840
      Width           =   3975
   End
End
Attribute VB_Name = "ADD_MEMBER_FINES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CHECKFRM As Integer
Dim VALUE_OF_MEMBER As Single
Private Sub Command1_Click()


If Text4.Text <> 0 And Text8.Text <> 0 Then
Text2.Text = (Text8 * Text4) / 100
Else
Text2 = Text3
End If
Adodc3.Recordset.Fields!MEMBER_NAME = Text6.Text
Adodc3.Recordset.Fields!FINES_VALUE = Text3.Text
Adodc3.Recordset.Fields!Fines_PERCENT = Text4.Text
           
    
Adodc3.Recordset.Update
If Text5.Text <> "" Then
 Dim i As Integer
        For i = 1 To Text5.Text
         Adodc5.Recordset.AddNew
          Adodc5.Recordset.Fields!fines_no = i
         Adodc5.Recordset.Fields!member_id = Text1.Text
         Adodc5.Recordset.Fields!MEMBER_NAME = Text6.Text
         Adodc5.Recordset.Fields!FINES_VALUE = Val(Text2.Text) / Val(Text5.Text)

                Adodc5.Recordset.Fields!FINES_TYPE = DataCombo1.Text
           Adodc5.Recordset.Fields!FINES_TOTAL = Text2.Text
           Adodc5.Recordset.Fields!FINES_DATE = DateValue(Now)
         Adodc5.Recordset.Update
        Next i

Else
MsgBox "·«»œ „‰ ﬂ «»… ⁄œœ «·œ›⁄« ", vbCritical
Exit Sub
End If


Text2.Visible = True
MsgBox " „ «· ‰›ÌÌ–", vbInformation
Command1.Enabled = False
Command3.Enabled = False
 
Command5.Enabled = False
End Sub

Private Sub Command2_Click()
Adodc3.Recordset.AddNew
Text7.Text = DateValue(Now)
Command1.Enabled = True
Command3.Enabled = True
 
Command5.Enabled = True
End Sub

Private Sub Command3_Click()
member_search.Show
member_search.from = 0
 
End Sub

 

Private Sub Command5_Click()
Dim i As Integer
Adodc7.CommandType = adCmdText
Adodc7.RecordSource = "select * from member_child where member_id= " & Text1.Text & " and MEMBER_TITLE='«·“ÊÃ…'"
Adodc7.Refresh

If Adodc7.Recordset.RecordCount > 0 Then

If Text4.Text <> 0 And Text8.Text <> 0 Then
Text2.Text = (Text8.Text * Text4) / 100
Else
Text2 = Text3
End If

Adodc3.Recordset.Fields!FINES_VALUE = Text3.Text
Adodc3.Recordset.Fields!Fines_PERCENT = Text4.Text
Adodc3.Recordset.Fields!CHILD_ID = Adodc7.Recordset.Fields!MEMBER_CHILD_ID
Adodc3.Recordset.Fields!MEMBER_NAME = Adodc7.Recordset.Fields!MEMBER_CHILD_NAME
Adodc3.Recordset.Update
If Text5.Text <> "" Then
 
        For i = 1 To Text5.Text
         Adodc5.Recordset.AddNew
          Adodc5.Recordset.Fields!fines_no = i
         Adodc5.Recordset.Fields!member_id = Text1.Text
         Adodc5.Recordset.Fields!MEMBER_NAME = Adodc7.Recordset.Fields!MEMBER_CHILD_NAME
         Adodc5.Recordset.Fields!FINES_VALUE = Val(Text2.Text) / Val(Text5.Text)
         Adodc5.Recordset.Fields!CHILD_ID = Adodc7.Recordset.Fields!MEMBER_CHILD_ID
                  Adodc5.Recordset.Fields!FINES_TYPE = DataCombo1.Text
                   Adodc5.Recordset.Fields!wife = 1
           Adodc5.Recordset.Fields!FINES_TOTAL = Text2.Text
           Adodc5.Recordset.Fields!FINES_DATE = DateValue(Now)
           
         Adodc5.Recordset.Update
        Next i

Else
MsgBox "·«»œ „‰ ﬂ «»… ⁄œœ «·œ›⁄« ", vbCritical
Exit Sub
End If


Text2.Visible = True
MsgBox " „ «· ‰›ÌÌ–", vbInformation
Command1.Enabled = False
Command3.Enabled = False
 
Command5.Enabled = False
Else
MsgBox "·« ÌÊÃœ “ÊÃ… ·Â–« «·⁄÷Ê", vbCritical


End If
End Sub


Private Sub DataCombo1_Click(Area As Integer)
Adodc2.CommandType = adCmdText
Adodc2.RecordSource = "SELECT  * FROM Fine_TYPES WHERE Fines_NAME LIKE'%" & DataCombo1.Text & "%'"
Adodc2.Refresh

Text3.Text = Adodc2.Recordset.Fields!FINES_VALUE
Text4.Text = Adodc2.Recordset.Fields!Fines_PERCENT
End Sub

Private Sub Form_Load()
VALUE_OF_MEMBER = 0

End Sub

