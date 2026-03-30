VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form REPORTSFRM 
   BackColor       =   &H80000007&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ﬁ«—Ì—  "
   ClientHeight    =   7095
   ClientLeft      =   1455
   ClientTop       =   525
   ClientWidth     =   8595
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   7095
   ScaleWidth      =   8595
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text5 
      Alignment       =   1  'Right Justify
      DataField       =   "MEMBER_ID"
      DataSource      =   "Adodc3"
      Height          =   285
      Left            =   2160
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Text            =   "Text5"
      Top             =   5880
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
      DataField       =   "MEMBER_ID"
      DataSource      =   "Adodc2"
      Height          =   375
      Left            =   2040
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Text            =   "Text4"
      Top             =   5160
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      DataField       =   "MEMBER_ID"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   2040
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Text            =   "Text3"
      Top             =   4800
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      Height          =   4575
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   120
      Width           =   8295
      Begin MSACAL.Calendar Calendar2 
         Height          =   2175
         Left            =   120
         TabIndex        =   12
         Top             =   1680
         Width           =   3975
         _Version        =   524288
         _ExtentX        =   7011
         _ExtentY        =   3836
         _StockProps     =   1
         BackColor       =   0
         Year            =   2009
         Month           =   10
         Day             =   9
         DayLength       =   1
         MonthLength     =   1
         DayFontColor    =   16777215
         FirstDay        =   7
         GridCellEffect  =   2
         GridFontColor   =   65535
         GridLinesColor  =   -2147483640
         ShowDateSelectors=   -1  'True
         ShowDays        =   -1  'True
         ShowHorizontalGrid=   -1  'True
         ShowTitle       =   -1  'True
         ShowVerticalGrid=   -1  'True
         TitleFontColor  =   65535
         ValueIsNull     =   0   'False
         BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSACAL.Calendar Calendar1 
         Height          =   2175
         Left            =   4080
         TabIndex        =   13
         Top             =   1680
         Width           =   3975
         _Version        =   524288
         _ExtentX        =   7011
         _ExtentY        =   3836
         _StockProps     =   1
         BackColor       =   0
         Year            =   2009
         Month           =   10
         Day             =   9
         DayLength       =   1
         MonthLength     =   1
         DayFontColor    =   16777215
         FirstDay        =   7
         GridCellEffect  =   2
         GridFontColor   =   65535
         GridLinesColor  =   -2147483640
         ShowDateSelectors=   -1  'True
         ShowDays        =   -1  'True
         ShowHorizontalGrid=   -1  'True
         ShowTitle       =   -1  'True
         ShowVerticalGrid=   -1  'True
         TitleFontColor  =   65535
         ValueIsNull     =   0   'False
         BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "„‰ ›÷·ﬂ «Œ — «·› —…"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   615
         Left            =   3000
         TabIndex        =   18
         Top             =   360
         Width           =   3375
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "≈·Ï"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   2760
         TabIndex        =   17
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "„‰"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   6600
         TabIndex        =   16
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label from_date 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3073
            SubFormatType   =   3
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   495
         Left            =   4560
         TabIndex        =   15
         Top             =   1080
         Width           =   2775
      End
      Begin VB.Label to_date 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3073
            SubFormatType   =   3
         EndProperty
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   495
         Left            =   360
         TabIndex        =   14
         Top             =   1080
         Width           =   2775
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Height          =   1455
      Left            =   2760
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   240
      Visible         =   0   'False
      Width           =   5535
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   360
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "«Œ «— «·”‰… «·„«·Ì… „‰ ›÷·ﬂ"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   2760
         TabIndex        =   9
         Top             =   240
         Width           =   3015
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Height          =   1455
      Left            =   2640
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   4800
      Visible         =   0   'False
      Width           =   5535
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   1200
         TabIndex        =   6
         Top             =   720
         Width           =   2895
      End
      Begin VB.CommandButton Command3 
         Height          =   735
         Left            =   240
         Picture         =   "emp_reports.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   2280
         TabIndex        =   3
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "«”„ «·ÿ«·»"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   4320
         TabIndex        =   7
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "—ﬁ„ «·ÿ«·»"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   4200
         TabIndex        =   5
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.CommandButton COMMAND1 
      Caption         =   "⁄—÷  ﬁ—Ì— «· ÃœÌœ  "
      Height          =   615
      Left            =   360
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   6480
      Width           =   7935
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   240
      Top             =   4800
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
      RecordSource    =   "MEMBERS"
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
      Height          =   375
      Left            =   240
      Top             =   5280
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
      RecordSource    =   "MEMBER_CHILD"
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
      Height          =   375
      Left            =   240
      Top             =   5760
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
      RecordSource    =   "gam3ea"
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
   Begin VB.Label case_id 
      Alignment       =   1  'Right Justify
      Caption         =   "case+id"
      Height          =   375
      Left            =   1920
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   2655
   End
End
Attribute VB_Name = "REPORTSFRM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Calendar1_Click()
from_date.Caption = Calendar1.year & "-" & Calendar1.Month & "-" & Calendar1.Day
End Sub

Private Sub Calendar2_Click()
to_date.Caption = Calendar2.year & "-" & Calendar2.Month & "-" & Calendar2.Day
End Sub

Private Sub display_logins_Click()
'If DataCombo1.Text = "" Then
'MsgBox "·«»œ „‰ «Œ Ì«— «·„ÊŸ›", vbCritical
'DataCombo1.SetFocus
'SendKeys ("{F4}")
'Exit Sub
'End If


'If from_date.Caption = "" Or to_date.Caption = "" Then
'MsgBox "·«»œ „«Œ Ì«— «·› —…", vbCritical
'
'Exit Sub
'End If


'Form3.criteria_user_id = DataCombo1.BoundText
'Form3.case_id = 5
'Form3.Show
End Sub

Private Sub Command1_Click()
If case_id = 8 And Text1.Text = "" Then
MsgBox "·«»œ „‰  ÕœÌœ «·⁄÷Ê Ê–·ﬂ »«·»ÃÀ ⁄‰… «Ê ﬂ «»… —ﬁ„… „»«‘—… ›Ì «·Œ«‰… «·„Ÿ··… »«··Ê‰ «·«Õ„—", vbCritical
Text1.BackColor = &HFF&
Exit Sub
End If


If case_id = 7 Then
 Adodc1.CommandType = adCmdText
 Adodc1.RecordSource = "select * from MEMBERS where last_update_year='" & Combo1.Text & "'"
 Adodc1.Refresh
 
 If Adodc1.Recordset.RecordCount > 0 Then
 Adodc1.Recordset.MoveFirst
 End If
  Adodc2.CommandType = adCmdText
 Adodc2.RecordSource = "select * from MEMBER_CHILD where MEMBER_TITLE='«·“ÊÃ…' and last_update_year='" & Combo1.Text & "'"
 Adodc2.Refresh
 
  If Adodc2.Recordset.RecordCount > 0 Then
 Adodc2.Recordset.MoveFirst
 End If
 
   If Adodc3.Recordset.RecordCount > 0 Then
 Adodc3.Recordset.MoveFirst
 End If
 
 
 For i = 1 To Adodc3.Recordset.RecordCount
 Adodc3.Recordset.Delete
 Adodc3.Recordset.MoveNext
 Next i
 
  For i = 1 To Adodc1.Recordset.RecordCount
 Adodc3.Recordset.AddNew
 Adodc3.Recordset.Fields!member_id = Adodc1.Recordset.Fields!member_id
 Adodc3.Recordset.Fields!MEMBER_NAME = Adodc1.Recordset.Fields!MEMBER_NAME
 Adodc3.Recordset.Fields!member_type = Adodc1.Recordset.Fields!member_type
  Adodc3.Recordset.Fields![year] = Combo1.Text
 Adodc3.Recordset.Update
 Adodc1.Recordset.MoveNext
 Next i
 
 
  For i = 1 To Adodc2.Recordset.RecordCount
 Adodc3.Recordset.AddNew
 Adodc3.Recordset.Fields!member_id = Adodc2.Recordset.Fields!member_id
  Adodc3.Recordset.Fields!MEMBER_CHILD_ID = Adodc2.Recordset.Fields!MEMBER_CHILD_ID
  
 Adodc3.Recordset.Fields!MEMBER_NAME = Adodc2.Recordset.Fields!MEMBER_CHILD_NAME
 Adodc3.Recordset.Fields!member_type = Adodc2.Recordset.Fields!member_type
 Adodc3.Recordset.Fields![year] = Combo1.Text
 Adodc3.Recordset.Update
 Adodc2.Recordset.MoveNext
 Next i
 
 
 

 
End If





If from_date.Caption = "" Or to_date.Caption = "" Then
MsgBox "·«»œ „«Œ Ì«— «·› —…", vbCritical
'
Exit Sub
End If

'Form3.Label5.Caption = Combo1.Text
Form3.from_date = from_date
Form3.to_date = to_date
Form3.case_id = case_id.Caption
Form3.Show

End Sub

Private Sub Command3_Click()
member_search.Show
member_search.from = 11
End Sub

Private Sub Form_Activate()
If case_id = 8 Then
Frame1.Visible = True
Else
Frame1.Visible = False
End If
End Sub

Private Sub Form_Load()
Calendar1.Value = Date
Calendar2.Value = Date
from_date.Caption = Calendar1.year & "-" & Calendar1.Month & "-" & Calendar1.Day
to_date.Caption = Calendar2.year & "-" & Calendar2.Month & "-" & Calendar2.Day

For i = 2007 To 2999
Combo1.AddItem i & "/" & i + 1
Next i
End Sub

