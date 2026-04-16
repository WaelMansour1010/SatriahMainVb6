VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frm_templates 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4995
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6885
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   6885
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.Frame Frame6 
      Height          =   2775
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   1335
      Begin ALLButtonS.ALLButton Command1 
         Height          =   495
         Index           =   4
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "ÈÇáÑÞã"
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
         MICON           =   "frm_templates.frx":0000
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
         Index           =   5
         Left            =   120
         TabIndex        =   8
         Top             =   1440
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "ÈÇáÇÓã"
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
         MICON           =   "frm_templates.frx":001C
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
         Index           =   6
         Left            =   120
         TabIndex        =   10
         Top             =   2040
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   873
         BTYPE           =   3
         TX              =   "ÈÇáÞÓã"
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
         MICON           =   "frm_templates.frx":0038
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label1 
         Caption         =   "ÈÍË"
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
         Height          =   495
         Left            =   480
         TabIndex        =   9
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.TextBox Text2 
      DataField       =   "opr_id"
      DataSource      =   "templates_details"
      Height          =   285
      Left            =   7920
      TabIndex        =   5
      Text            =   "Text2"
      Top             =   5400
      Width           =   255
   End
   Begin VB.TextBox Text1 
      DataField       =   "templates_id"
      DataSource      =   "Adodc2"
      Height          =   285
      Left            =   7680
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   3120
      Width           =   375
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   5760
      Top             =   120
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frm_templates.frx":0054
      Height          =   4095
      Left            =   1560
      TabIndex        =   0
      Top             =   600
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   7223
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   -2147483648
      ColumnHeaders   =   -1  'True
      HeadLines       =   2
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
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "templates_id"
         Caption         =   "ÑÞã ÇáäãæÐÌ"
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
         DataField       =   "templates_name"
         Caption         =   "ÇÓã ÇáäãæÐÌ"
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
      BeginProperty Column02 
         DataField       =   "image_name"
         Caption         =   "image_name"
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
      BeginProperty Column03 
         DataField       =   "departement_name"
         Caption         =   "ÇáÞÓã ÇáÊÇÈÚ áÉ"
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
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column02 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1739.906
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   5880
      Top             =   120
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
   Begin MSAdodcLib.Adodc templates_details 
      Height          =   330
      Left            =   7080
      Top             =   4920
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
   Begin ALLButtonS.ALLButton Command2 
      Height          =   615
      Left            =   240
      TabIndex        =   11
      Top             =   3960
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   "ÇÎÊíÇÑ"
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
      BCOL            =   32896
      BCOLO           =   32896
      FCOL            =   16777215
      FCOLO           =   0
      MCOL            =   192
      MPTR            =   1
      MICON           =   "frm_templates.frx":0069
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label case_id 
      Caption         =   "1000"
      Height          =   615
      Left            =   -120
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H000080FF&
      Caption         =   "ÇÎÊíÇÑ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "ÇÎÊÇÑ ÇáäãæÐÌ ÇáãÑÇÏ ÇáÊÚÇãá ãÚÉ"
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
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Top             =   0
      Width           =   4215
   End
End
Attribute VB_Name = "frm_templates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x As String

Private Sub Command1_Click(Index As Integer)

    Select Case Index

        Case 0
            Adodc1.Recordset.AddNew

        Case 1
            Adodc1.Recordset.update

        Case 2
            x = MsgBox("åá ÇäÊ ãÊÃßÏ ãä ÚãáíÉ ÇáÍÐÝ", vbCritical + vbYesNo)

            If x = vbNo Then
                Exit Sub
            End If

            If Adodc1.Recordset.RecordCount > 0 Then
                Adodc1.Recordset.delete
                Adodc1.Refresh
                DataGrid1.Refresh
            End If

        Case 3

            If Adodc1.Recordset.RecordCount > 0 Then
    
                'Form3.case_id = Me.Name
   
                'Form3.show
            End If

        Case 4
            On Error Resume Next
            x = InputBox("ÇÏÎá ÇáÑÞã ÇáãØáæÈ ÇáÈÍË ÚäÉ")

            If IsNumeric(x) Then
                Adodc1.CommandType = adCmdText
                Adodc1.RecordSource = "select * from  templates where   subject_no=0 and templates_id=" & x
                Adodc1.Refresh
            Else
                MsgBox "áÇÈÏ ãä ÇÏÎÇá ÑÞã ÝÞØ", vbCritical
            End If

        Case 5
            x = InputBox("ÇÏÎá ßáãÉ ÇáÈÍË")
            Adodc1.CommandType = adCmdText
            Adodc1.RecordSource = "select * from templates where subject_no=0 and  templates_name like '%" & x & "%'"
            Adodc1.Refresh

        Case 5
            x = InputBox("ÇÏÎá ßáãÉ ÇáÈÍË")
            Adodc1.CommandType = adCmdText
            Adodc1.RecordSource = "select * from templates where subject_no=0 and  departement_name like '%" & x & "%'"
            Adodc1.Refresh

    End Select

End Sub

Private Sub Command2_Click()
    Label4_Click
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500
    
    connection_string = Cn.ConnectionString
 
    Adodc1.ConnectionString = connection_string
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "select * from templatesnew where subject_no is null"
    Adodc1.Refresh

    Adodc2.ConnectionString = connection_string
    Adodc2.CommandType = adCmdText
    Adodc2.RecordSource = "select * from  templates_detailsnew"
    Adodc2.Refresh

    templates_details.ConnectionString = connection_string
    templates_details.CommandType = adCmdText
 
End Sub

Private Sub Label4_Click()
    Dim template_id As Integer
    Dim template_name, departement_name As String
    Dim SUBJECT_NO, no_of_images As Integer
   On Error Resume Next

    If case_id = 0 Then
        template_id = Adodc1.Recordset.Fields!templates_id
        template_name = Adodc1.Recordset.Fields!templates_name
        SUBJECT_NO = imaged.SUBJECT_NO.Caption
        no_of_images = Adodc1.Recordset.Fields!no_of_images

        If IsNull(Adodc1.Recordset.Fields!departement_name) Then
            departement_name = ""
        Else
            departement_name = Adodc1.Recordset.Fields!departement_name
        End If

        templates_details.CommandType = adCmdText
        templates_details.RecordSource = "select * from templates_detailsnew where templates_id=" & template_id ' & "and subject_no=0"
        templates_details.Refresh

        If templates_details.Recordset.RecordCount = 0 Then
            MsgBox "åÐÇ ÇáäãæÐÌ ÈÉ ÎØÃ íÑÌí ÇÎÊíÇÑ äãæÐÌ ÇÎÑ", vbCritical
            Exit Sub
        End If

        Adodc1.Recordset.AddNew 'templates table
        'Adodc1.Recordset.Fields!templates_id = templates_id
        Adodc1.Recordset.Fields!templates_name = template_name
        Adodc1.Recordset.Fields!SUBJECT_NO = SUBJECT_NO
        Adodc1.Recordset.Fields!no_of_images = no_of_images
        Adodc1.Recordset.Fields!departement_name = departement_name
        Adodc1.Recordset.update
        Adodc1.Recordset.MoveLast

        imaged.Adodc3.Recordset.AddNew 'subject_templates_table
        imaged.Adodc3.Recordset.Fields!SUBJECT_NO = SUBJECT_NO
        imaged.Adodc3.Recordset.Fields!template_id = Adodc1.Recordset.Fields!templates_id
        imaged.Adodc3.Recordset.Fields!template_name = template_name
        imaged.Adodc3.Recordset.Fields!no_of_images = no_of_images
        imaged.Adodc3.Recordset.Fields!date_added = DateValue(Now)

        imaged.Adodc3.Recordset.update
        imaged.Adodc3.Refresh
        imaged.DataGrid1.Refresh

        templates_details.CommandType = adCmdText
        templates_details.RecordSource = "select * from templates_detailsnew where templates_id=" & template_id ' & "and subject_no=0"
        templates_details.Refresh

        If templates_details.Recordset.RecordCount > 0 Then
            templates_details.Recordset.MoveFirst
        End If

        For i = 1 To templates_details.Recordset.RecordCount
            Adodc2.Recordset.AddNew ' templates_details table
            Adodc2.Recordset.Fields!templates_id = Adodc1.Recordset.Fields!templates_id 'templates_details.Recordset.Fields!templates_id
            Adodc2.Recordset.Fields!X1 = templates_details.Recordset.Fields!X1
            Adodc2.Recordset.Fields!X2 = templates_details.Recordset.Fields!X2
            Adodc2.Recordset.Fields!Y1 = templates_details.Recordset.Fields!Y1
            Adodc2.Recordset.Fields!Y2 = templates_details.Recordset.Fields!Y2
            Adodc2.Recordset.Fields!Text = templates_details.Recordset.Fields!Text
            Adodc2.Recordset.Fields!image_id = templates_details.Recordset.Fields!image_id
            Adodc2.Recordset.Fields!IMAGE_NAME = templates_details.Recordset.Fields!IMAGE_NAME
            Adodc2.Recordset.Fields!color = templates_details.Recordset.Fields!color
            Adodc2.Recordset.Fields!backcolor = templates_details.Recordset.Fields!backcolor
            Adodc2.Recordset.Fields!FontName = templates_details.Recordset.Fields!FontName
            Adodc2.Recordset.Fields!fontsize = templates_details.Recordset.Fields!fontsize
            Adodc2.Recordset.Fields!FontBold = templates_details.Recordset.Fields!FontBold
            Adodc2.Recordset.Fields!FontItalic = templates_details.Recordset.Fields!FontItalic
            Adodc2.Recordset.Fields!FontUnderline = templates_details.Recordset.Fields!FontUnderline
            Adodc2.Recordset.Fields!Strikethrough = templates_details.Recordset.Fields!Strikethrough
            Adodc2.Recordset.Fields!image_direction = templates_details.Recordset.Fields!image_direction
            Adodc2.Recordset.Fields!SUBJECT_NO = SUBJECT_NO
            Adodc2.Recordset.update
            templates_details.Recordset.MoveNext
        Next i

        Me.Hide

    Else

        If Adodc1.Recordset.RecordCount <> 0 Then
            loading_temolates.show
            loading_temolates.Label6.Caption = Adodc1.Recordset.Fields!templates_id
            'loading_temolates.Label7.Caption = Adodc1.Recordset.Fields!IMAGE_NAME
            loading_temolates.Label9.Caption = Adodc1.Recordset.Fields!no_of_images
            Me.Hide
        End If

    End If

End Sub
