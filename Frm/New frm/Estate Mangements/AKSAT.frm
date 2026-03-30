VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form AKSAT 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Õ”«» «·«ﬁ”«ÿ"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6720
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   6720
   Begin MSACAL.Calendar Calendar1 
      Height          =   2415
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Visible         =   0   'False
      Width           =   3495
      _Version        =   524288
      _ExtentX        =   6165
      _ExtentY        =   4260
      _StockProps     =   1
      BackColor       =   -2147483633
      Year            =   2010
      Month           =   11
      Day             =   23
      DayLength       =   1
      MonthLength     =   1
      DayFontColor    =   0
      FirstDay        =   7
      GridCellEffect  =   1
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   -1  'True
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame7 
      Height          =   1215
      Left            =   2400
      TabIndex        =   11
      Top             =   6120
      Visible         =   0   'False
      Width           =   3975
      Begin ALLButtonS.ALLButton Command1 
         Height          =   495
         Index           =   1
         Left            =   1440
         TabIndex        =   12
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "Õ›Ÿ"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   16711680
         BCOLO           =   12582912
         FCOL            =   16777215
         FCOLO           =   0
         MCOL            =   192
         MPTR            =   1
         MICON           =   "AKSAT.frx":0000
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ALLButtonS.ALLButton Command1 
         Height          =   495
         Index           =   2
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "Õ–›"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   255
         BCOLO           =   192
         FCOL            =   16777215
         FCOLO           =   0
         MCOL            =   192
         MPTR            =   1
         MICON           =   "AKSAT.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ALLButtonS.ALLButton Command1 
         Height          =   495
         Index           =   3
         Left            =   120
         TabIndex        =   14
         Top             =   2040
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "ÿ»«⁄…"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   16711680
         BCOLO           =   12582912
         FCOL            =   16777215
         FCOLO           =   0
         MCOL            =   192
         MPTR            =   1
         MICON           =   "AKSAT.frx":0038
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ALLButtonS.ALLButton Command1 
         Height          =   495
         Index           =   0
         Left            =   2640
         TabIndex        =   15
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "ÃœÌœ"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   16711680
         BCOLO           =   12582912
         FCOL            =   16777215
         FCOLO           =   0
         MCOL            =   192
         MPTR            =   1
         MICON           =   "AKSAT.frx":0054
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   960
         Top             =   840
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   582
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
         Caption         =   "  Õ—Ìﬂ"
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
      Begin VB.Label Label5 
         Caption         =   "Label2"
         Height          =   15
         Left            =   -120
         TabIndex        =   16
         Top             =   1440
         Width           =   855
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      TabIndex        =   10
      Top             =   1200
      Width           =   495
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      ItemData        =   "AKSAT.frx":0070
      Left            =   3480
      List            =   "AKSAT.frx":007D
      TabIndex        =   9
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox Text14 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      TabIndex        =   3
      Top             =   0
      Width           =   1575
   End
   Begin VB.TextBox Text16 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      TabIndex        =   2
      Top             =   600
      Width           =   1575
   End
   Begin VB.TextBox Text17 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   0
      Width           =   1575
   End
   Begin VB.CommandButton Command40 
      Height          =   492
      Index           =   2
      Left            =   0
      Picture         =   "AKSAT.frx":0090
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   492
   End
   Begin ALLButtonS.ALLButton Command100 
      Height          =   255
      Left            =   480
      TabIndex        =   7
      Top             =   4680
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   450
      BTYPE           =   3
      TX              =   "Õ”«» «·«ﬁ”«ÿ"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16711680
      BCOLO           =   12582912
      FCOL            =   16777215
      FCOLO           =   0
      MCOL            =   192
      MPTR            =   1
      MICON           =   "AKSAT.frx":08F2
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSDataGridLib.DataGrid ÌŒ 
      Bindings        =   "AKSAT.frx":090E
      Height          =   2535
      Left            =   480
      TabIndex        =   17
      Top             =   1800
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   4471
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   -2147483648
      ColumnHeaders   =   -1  'True
      HeadLines       =   1
      RowHeight       =   19
      FormatLocked    =   -1  'True
      RightToLeft     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "id"
         Caption         =   "id"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3073
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "index"
         Caption         =   "—ﬁ„ «·ﬁ”ÿ"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3073
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "contarct_id"
         Caption         =   "contarct_id"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3073
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "date"
         Caption         =   " «—ÌŒ «·«Ì Õﬁ«ﬁ"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3073
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "value"
         Caption         =   "ﬁÌ„… «·ﬁ”ÿ"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3073
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            Object.Visible         =   0   'False
            ColumnWidth     =   1275.024
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1275.024
         EndProperty
         BeginProperty Column02 
            Object.Visible         =   0   'False
            ColumnWidth     =   1275.024
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2429.858
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1590.236
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "«·› —… »Ì‰ «·«ﬁ”«ÿ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   5160
      TabIndex        =   8
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label Label24 
      Caption         =   "⁄œœ «·«ﬁ”«ÿ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   5160
      TabIndex        =   6
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label Label26 
      Caption         =   "«·„»·€"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   5520
      TabIndex        =   5
      Top             =   720
      Width           =   615
   End
   Begin VB.Label Label27 
      Caption         =   " «—ÌŒ «Ê· ﬁ”ÿ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Top             =   0
      Width           =   1455
   End
End
Attribute VB_Name = "AKSAT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Calendar1_Click()
Text17.Text = Calendar1.value
Calendar1.Visible = False
End Sub

Private Sub Command100_Click()
On Error Resume Next
If Not IsNumeric(Text1.Text) Then MsgBox "·«»œ „‰ «œŒ«· «·› —… »Ì‰ «·«ﬁ”«ÿ —ﬁ„", vbCritical: Exit Sub


If Not IsNumeric(Text14.Text) Then MsgBox "·«»œ „‰ «œŒ«· ⁄œœ «·«ﬁ”«ÿ —ﬁ„", vbCritical: Exit Sub
If Not IsNumeric(Text16.Text) Then MsgBox "·«»œ „‰ «œŒ«· «·„»·€ —ﬁ„", vbCritical: Exit Sub

If Combo1.ListIndex = -1 Then MsgBox "·«»œ „‰ «Œ Ì«— «·› —… »«·ÌÊ„ «Ê »«·‘Â— «Ê »«·”‰…", vbCritical:  Combo1.SetFocus: SendKeys "{f4}": Exit Sub

For i = 1 To Adodc1.Recordset.RecordCount
Adodc1.Recordset.Delete
Adodc1.Recordset.MoveNext
Next i
Dim current_date As Date
Dim new_date As Date
Dim flag As String

If Combo1.ListIndex = 0 Then
flag = "d"
Else
If Combo1.ListIndex = 1 Then
flag = "m"
Else
If Combo1.ListIndex = 2 Then
flag = "yyyy"
End If
End If
End If


'Print DateAdd("yyyy", 3, Now) 'Year
'Print DateAdd("m", 3, Now) 'Month
'Print DateAdd("ww", 3, Now) 'week
'Print DateAdd("d", 3, Now) 'day
'Print DateAdd("h", 3, Now) 'Hour
' Print DateAdd("n", 3, Now) 'minute
current_date = Text17.Text
 Adodc1.Recordset.AddNew
 Adodc1.Recordset.Fields![Date] = current_date
  Adodc1.Recordset.Fields!contarct_id = 0
  Adodc1.Recordset.Fields![Index] = 1
  Adodc1.Recordset.Fields![value] = Round(Text16.Text / Text14.Text, 2)
 Adodc1.Recordset.Update

For i = 2 To Text14.Text
new_date = DateAdd(flag, Text1.Text, current_date)  'Year

Adodc1.Recordset.AddNew
Adodc1.Recordset.Fields!contarct_id = 12345
 Adodc1.Recordset.Fields![Date] = new_date
  Adodc1.Recordset.Fields![Index] = i
    Adodc1.Recordset.Fields![value] = Round(Text16.Text / Text14.Text, 2)
    
 Adodc1.Recordset.Update
 current_date = new_date

Next i
Adodc1.Refresh



End Sub

Private Sub Command40_Click(Index As Integer)
Calendar1.Visible = True
Calendar1.value = Date
End Sub

Private Sub Form_Click()
Calendar1.Visible = False

End Sub

Private Sub Form_Load()
      On Error Resume Next
          login.SkinFramework.ApplyWindow Me.hWnd
Me.Left = (MDIForm1.Width - Me.Width) / 2
   Me.Top = (MDIForm1.Height - Me.Height) / 2 - 500
   
Adodc1.ConnectionString = connection_string
 Adodc1.CommandType = adCmdText
Adodc1.RecordSource = "select * from aksat where contarct_id=12345" '  where branch_no =" & branch_no ' where departement_no=0"
Adodc1.Refresh

For i = 1 To Adodc1.Recordset.RecordCount
Adodc1.Recordset.Delete
Adodc1.Recordset.MoveNext
Next i


End Sub

