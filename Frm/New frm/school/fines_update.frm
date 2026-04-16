VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form fines_update 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "                  «Œ «— «·€—«„«  «·„ÿ·Ê» ”œ«œÂ«"
   ClientHeight    =   7200
   ClientLeft      =   -15
   ClientTop       =   345
   ClientWidth     =   4725
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7200
   ScaleWidth      =   4725
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FF00&
      Caption         =   "Õ›Ÿ"
      Height          =   495
      Left            =   480
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6600
      Width           =   1335
   End
   Begin VB.CheckBox Check1 
      Caption         =   " ›€Ì· «·€—«„…"
      DataField       =   "ACTIVATED"
      DataSource      =   "Adodc2"
      Height          =   615
      Left            =   2880
      TabIndex        =   4
      Top             =   2040
      Width           =   1575
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      DataField       =   "MEMBER_NAME"
      DataSource      =   "Adodc2"
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Top             =   600
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      DataField       =   "MEMBER_ID"
      DataSource      =   "Adodc2"
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   0
      Width           =   2295
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      DataField       =   "FINES_TOTAL"
      DataSource      =   "Adodc2"
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
      Left            =   1080
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1560
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      DataField       =   "FINES_TYPE"
      DataSource      =   "Adodc2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1080
      TabIndex        =   0
      Text            =   " "
      Top             =   1080
      Width           =   2295
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "fines_update.frx":0000
      Height          =   3495
      Left            =   480
      TabIndex        =   6
      Top             =   3000
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   6165
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      ColumnHeaders   =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      RightToLeft     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   495
      Left            =   -720
      Top             =   7320
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
   Begin VB.Label wife 
      Caption         =   "0"
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   840
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "—ﬁ„"
      DataField       =   "FINES_NO"
      DataSource      =   "Adodc2"
      Height          =   255
      Left            =   1800
      TabIndex        =   15
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "—ﬁ„"
      Height          =   255
      Left            =   2280
      TabIndex        =   14
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   " «—ÌŒ «·œ›⁄"
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
      Left            =   120
      TabIndex        =   13
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label Label4 
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
      Height          =   375
      Left            =   1800
      TabIndex        =   12
      Top             =   2520
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "—ﬁ„  «·€—«„…"
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
      Left            =   3120
      TabIndex        =   11
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "«”„ «·ÿ«·»  "
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
      Left            =   2040
      TabIndex        =   10
      Top             =   480
      Width           =   3975
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   " —ﬁ„ «·ÿ«·»  "
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
      Height          =   375
      Left            =   2520
      TabIndex        =   9
      Top             =   0
      Width           =   3015
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
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   3120
      TabIndex        =   8
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "‰Ê⁄ «·€—«„… "
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
      Left            =   3120
      TabIndex        =   7
      Top             =   1080
      Width           =   1815
   End
End
Attribute VB_Name = "fines_update"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CHECKFRM As Integer

Private Sub Command1_Click()
    Dim SUM As Single
    Dim x, i As Integer
    Dim dtmTest As Date
    'If Text2.Text = "€—«„… Ã„⁄Ì… ⁄„Ê„Ì…" And wife.Caption = "0" Then
 
    SUM = 0
    x = MsgBox("Â· «‰  „ √ﬂœ „‰ Â–… «·⁄„·Ì… ÕÌÀ «‰… ·«Ì„ﬂ‰ «·—ÃÊ⁄ ›ÌÂ«", vbCritical + vbYesNo)

    If x = vbNo Then Exit Sub
    If Adodc2.Recordset.RecordCount = 0 Then Exit Sub

    Adodc2.Recordset.MoveFirst

    For i = 1 To Adodc2.Recordset.RecordCount

        If Adodc2.Recordset.Fields!ACTIVATED = True Then
            Adodc2.Recordset.Fields!payed = True
            Adodc2.Recordset.Fields!PAYED_DATE = DateValue(Now)
            SUM = SUM + Adodc2.Recordset.Fields!fines_value
            Adodc2.Recordset.update
        End If
        
        Adodc2.Recordset.MoveNext
    Next i

    operatiomn_update_frm.Text8.text = operatiomn_update_frm.Text8.text + SUM
    Call operatiomn_update_frm.update_date
    Unload Me
    'End If

    If Text2.text = "€—«„…  √ŒÌ—" And wife.Caption = "0" Then
 
        SUM = 0
        x = MsgBox("Â· «‰  „ √ﬂœ „‰ Â–… «·⁄„·Ì… ÕÌÀ «‰… ·«Ì„ﬂ‰ «·—ÃÊ⁄ ›ÌÂ«", vbCritical + vbYesNo)

        If x = vbNo Then Exit Sub

        Adodc2.Recordset.MoveFirst

        For i = 1 To Adodc2.Recordset.RecordCount

            If Adodc2.Recordset.Fields!ACTIVATED = True Then
                Adodc2.Recordset.Fields!payed = True
                Adodc2.Recordset.Fields!PAYED_DATE = DateValue(Now)
                SUM = SUM + Adodc2.Recordset.Fields!fines_value
                Adodc2.Recordset.update
            End If
        
            Adodc2.Recordset.MoveNext
        Next i

        operatiomn_update_frm.Text8.text = operatiomn_update_frm.Text8.text + SUM
        Call operatiomn_update_frm.update_date
        Unload Me

    End If

    If Text2.text = "€—«„… Ã„⁄Ì… ⁄„Ê„Ì…" And wife.Caption = "1" Then
 
        SUM = 0
        x = MsgBox("Â· «‰  „ √ﬂœ „‰ Â–… «·⁄„·Ì… ÕÌÀ «‰… ·«Ì„ﬂ‰ «·—ÃÊ⁄ ›ÌÂ«", vbCritical + vbYesNo)

        If x = vbNo Then Exit Sub

        Adodc2.Recordset.MoveFirst

        For i = 1 To Adodc2.Recordset.RecordCount

            If Adodc2.Recordset.Fields!ACTIVATED = True Then
                Adodc2.Recordset.Fields!payed = True
                Adodc2.Recordset.Fields!PAYED_DATE = DateValue(Now)
                SUM = SUM + Adodc2.Recordset.Fields!fines_value
                Adodc2.Recordset.update
            End If
        
            Adodc2.Recordset.MoveNext
        Next i

        operatiomn_update_frm.Text26.text = operatiomn_update_frm.Text26.text + SUM
        Call operatiomn_update_frm.update_date
        Unload Me
    End If

    If Text2.text = "€—«„…  √ŒÌ—" And wife.Caption = "1" Then
 
        SUM = 0
        x = MsgBox("Â· «‰  „ √ﬂœ „‰ Â–… «·⁄„·Ì… ÕÌÀ «‰… ·«Ì„ﬂ‰ «·—ÃÊ⁄ ›ÌÂ«", vbCritical + vbYesNo)

        If x = vbNo Then Exit Sub

        Adodc2.Recordset.MoveFirst

        For i = 1 To Adodc2.Recordset.RecordCount

            If Adodc2.Recordset.Fields!ACTIVATED = True Then
                Adodc2.Recordset.Fields!payed = True
                Adodc2.Recordset.Fields!PAYED_DATE = DateValue(Now)
                SUM = SUM + Adodc2.Recordset.Fields!fines_value
                Adodc2.Recordset.update
            End If
        
            Adodc2.Recordset.MoveNext
        Next i

        operatiomn_update_frm.Text25.text = operatiomn_update_frm.Text25.text + SUM
        Call operatiomn_update_frm.update_date
        Unload Me

    End If

End Sub
 
Private Sub Form_Load()
    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500

    connection_string = Cn.ConnectionString
    Adodc2.ConnectionString = connection_string
    Adodc2.CommandType = adCmdText
    Adodc2.RecordSource = "  select FINES_NO,FINES_VALUE,PAYED_DATE, ACTIVATED,FINES_TYPE,FINES_TOTAL,MEMBER_ID,MEMBER_name  FROM FINES_DETAILS where MEMBER_ID=0 "
    Adodc2.Refresh

End Sub
