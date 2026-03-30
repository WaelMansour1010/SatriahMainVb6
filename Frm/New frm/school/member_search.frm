VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form f 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "    ÔÇÔÉ ÇáÈÍË  Úä ÇáØÇáÈ"
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   5265
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   6105
   ScaleWidth      =   5265
   Begin VB.CommandButton Command1 
      Caption         =   "ãæÇÝÞ"
      Default         =   -1  'True
      Height          =   315
      Left            =   0
      TabIndex        =   5
      Top             =   6120
      Width           =   3135
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "member_search.frx":0000
      Height          =   4455
      Left            =   0
      TabIndex        =   4
      Top             =   1560
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   7858
      _Version        =   393216
      AllowUpdate     =   -1  'True
      ColumnHeaders   =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
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
      ColumnCount     =   23
      BeginProperty Column00 
         DataField       =   "MEMBER_ID"
         Caption         =   "MEMBER_ID"
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
         DataField       =   "waly_name"
         Caption         =   "waly_name"
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
         DataField       =   "waly_tel"
         Caption         =   "waly_tel"
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
         DataField       =   "MEMBER_NAME"
         Caption         =   "MEMBER_NAME"
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
         DataField       =   "MEMBER_TYPE"
         Caption         =   "MEMBER_TYPE"
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
      BeginProperty Column05 
         DataField       =   "MEMBER_DOB"
         Caption         =   "MEMBER_DOB"
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
      BeginProperty Column06 
         DataField       =   "MEMBER_born_place"
         Caption         =   "MEMBER_born_place"
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
      BeginProperty Column07 
         DataField       =   "MEMBER_address"
         Caption         =   "MEMBER_address"
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
      BeginProperty Column08 
         DataField       =   "MEMBER_certificate"
         Caption         =   "MEMBER_certificate"
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
      BeginProperty Column09 
         DataField       =   "MEMBER_telephone"
         Caption         =   "MEMBER_telephone"
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
      BeginProperty Column10 
         DataField       =   "MEMBER_job"
         Caption         =   "MEMBER_job"
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
      BeginProperty Column11 
         DataField       =   "MEMBER_job_address"
         Caption         =   "MEMBER_job_address"
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
      BeginProperty Column12 
         DataField       =   "MEMBER_NATIONAL_id"
         Caption         =   "MEMBER_NATIONAL_id"
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
      BeginProperty Column13 
         DataField       =   "MEMBER_date_of_issue"
         Caption         =   "MEMBER_date_of_issue"
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
      BeginProperty Column14 
         DataField       =   "SEX"
         Caption         =   "SEX"
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
      BeginProperty Column15 
         DataField       =   "VALUE_OF_MEMBER"
         Caption         =   "VALUE_OF_MEMBER"
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
      BeginProperty Column16 
         DataField       =   "IMAGE_PATH"
         Caption         =   "IMAGE_PATH"
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
      BeginProperty Column17 
         DataField       =   "expire_date"
         Caption         =   "expire_date"
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
      BeginProperty Column18 
         DataField       =   "MEMBER_TYPE1"
         Caption         =   "MEMBER_TYPE1"
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
      BeginProperty Column19 
         DataField       =   "last_update_year"
         Caption         =   "last_update_year"
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
      BeginProperty Column20 
         DataField       =   "MEMBER_TITLE"
         Caption         =   "MEMBER_TITLE"
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
      BeginProperty Column21 
         DataField       =   "image_location"
         Caption         =   "image_location"
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
      BeginProperty Column22 
         DataField       =   "X"
         Caption         =   "X"
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
            ColumnWidth     =   2294.929
         EndProperty
         BeginProperty Column01 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column02 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2505.26
         EndProperty
         BeginProperty Column04 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column05 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column06 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column07 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column08 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column09 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column10 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column11 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column12 
            Object.Visible         =   0   'False
            ColumnWidth     =   1830.047
         EndProperty
         BeginProperty Column13 
            Object.Visible         =   0   'False
            ColumnWidth     =   1800
         EndProperty
         BeginProperty Column14 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column15 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column16 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column17 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column18 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column19 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column20 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column21 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column22 
            Object.Visible         =   0   'False
            ColumnWidth     =   915.024
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   1320
      TabIndex        =   3
      Top             =   480
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   1320
      TabIndex        =   0
      Top             =   0
      Width           =   2055
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   480
      Top             =   7200
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   1085
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
   Begin VB.Label Label4 
      Caption         =   "ÑÞã ÇáØÇáÈ"
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   3240
      TabIndex        =   8
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "ÇÓã ÇáØÇáÈ"
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   840
      TabIndex        =   7
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label from 
      Caption         =   "Label3"
      Height          =   375
      Left            =   480
      TabIndex        =   6
      Top             =   6720
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "ÇÓã ÇáØÇáÈ"
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   3840
      TabIndex        =   2
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "ÑÞã ÇáØÇáÈ"
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   3840
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "f"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

    If From = 0 And Not IsNull(Adodc1.Recordset.Fields!member_id) Then
 
        ADD_MEMBER_FINES.Text6.text = Adodc1.Recordset.Fields!member_name

        ADD_MEMBER_FINES.Text1.text = Adodc1.Recordset.Fields!member_id
 
        ADD_MEMBER_FINES.Text8.text = Adodc1.Recordset.Fields!VALUE_OF_MEMBER

        'Else
        'ADD_MEMBER_FINES.Text6.text = ""
        '
        'ADD_MEMBER_FINES.Text1.text = ""
 
    End If

    If From = 1 And Not IsNull(Adodc1.Recordset.Fields!member_id) Then
        member_activity.Text3 = Adodc1.Recordset.Fields!member_name
        member_activity.Text2 = Adodc1.Recordset.Fields!member_id
        member_activity.Label60 = Adodc1.Recordset.Fields!member_id
        member_activity.Label70 = "0"
        member_activity.Text4.text = Adodc1.Recordset.Fields!MEMBER_TYPE
    End If

    If From = 2 And Not IsNull(Adodc1.Recordset.Fields!member_id) And Adodc1.Recordset.RecordCount > 0 Then
        ADD_MEMBER_INSTALLMENTS.Text6.text = Adodc1.Recordset.Fields!member_name

        ADD_MEMBER_INSTALLMENTS.Text1.text = Adodc1.Recordset.Fields!member_id

    End If

    If From = 3 And Not IsNull(Adodc1.Recordset.Fields!member_id) And Adodc1.Recordset.RecordCount > 0 Then
        delete_member_activity.Text1.text = Adodc1.Recordset.Fields!member_id

        delete_member_activity.Text2.text = Adodc1.Recordset.Fields!member_name
        delete_member_activity.Text3.text = Adodc1.Recordset.Fields!MEMBER_TYPE
    End If

    If From = 4 And Not IsNull(Adodc1.Recordset.Fields!member_id) And Adodc1.Recordset.RecordCount > 0 Then

        MEMBERS.Adodc1.CommandType = adCmdText
        MEMBERS.Adodc1.RecordSource = "select *  FROM MEMBERS where MEMBER_ID LIKE'%" & Adodc1.Recordset.Fields!member_id & "%'"
        MEMBERS.Adodc1.Refresh

    End If

    'X = InputBox("ÇÏÎá ÇáÑÞã Çæ ÌÒÁ ãä ÇáÑÞã", "ÔÇÔÉ ÇáÈÍË ÈÇáÑÞã")
    If From = 5 And Not IsNull(Adodc1.Recordset.Fields!member_id) And Adodc1.Recordset.RecordCount > 0 Then
        'select * from operations where PAYED=0 and
        operatiomn_update_frm.Adodc1.CommandType = adCmdText
        operatiomn_update_frm.Adodc1.RecordSource = "select *  FROM OPERATIONS where  operation_type='ÊÌÏíÏ ÚÖæíÉ' and  PAYED=0 and MEMBER_ID ='" & Adodc1.Recordset.Fields!member_id & "'"
        operatiomn_update_frm.Adodc1.Refresh

        If operatiomn_update_frm.Text1.text = "" Then operatiomn_update_frm.Text1.text = 0
        operatiomn_update_frm.Adodc2.CommandType = adCmdText
        '   operatiomn_update_frm.Adodc2.RecordSource = "SELECT op_id,CHILD_ID,CHILD_NAME, MEMBER_VALUE,member_card_value,Fines_NAME,Fines_value ,PAYED FROM OPERATION_DETAILS  WHERE PAYED =0 AND MEMBER_ID=" & Adodc1.Recordset.Fields!member_id
        operatiomn_update_frm.Adodc2.RecordSource = "SELECT op_id,CHILD_ID,CHILD_NAME, MEMBER_VALUE,member_card_value,fines_value,fines_value1 ,PAYED FROM OPERATION_DETAILS  WHERE PAYED =0 AND  MEMBER_ID=" & Adodc1.Recordset.Fields!member_id
        operatiomn_update_frm.Adodc2.Refresh

    End If

    If From = 6 And Not IsNull(Adodc1.Recordset.Fields!member_id) And Adodc1.Recordset.RecordCount > 0 Then
        operation_from.Adodc1.CommandType = adCmdText
        operation_from.Adodc1.RecordSource = "select *  FROM OPERATIONS where   PAYED=0 and MEMBER_ID ='" & Adodc1.Recordset.Fields!member_id & "'" '& " and operation_type LIKE'%" & operation_from.Label25.Caption & "%'"
        operation_from.Adodc1.Refresh

        If operation_from.Text1.text = "" Then operation_from.Text1.text = 0
        operation_from.Adodc2.CommandType = adCmdText
    
        ' operation_from.Adodc2.RecordSource = "SELECT op_id,CHILD_ID,CHILD_NAME, MEMBER_VALUE,member_card_value,Fines_NAME,Fines_value ,activity_value,PAYED FROM OPERATION_DETAILS  WHERE PAYED =0 AND MEMBER_ID=" & Adodc1.Recordset.Fields!member_id
        operation_from.Adodc2.RecordSource = "SELECT op_id,CHILD_ID,CHILD_NAME, MEMBER_VALUE,member_card_value,fines_value,fines_value1 ,PAYED FROM OPERATION_DETAILS  WHERE PAYED =0 AND  MEMBER_ID=" & Adodc1.Recordset.Fields!member_id
        operation_from.Adodc2.Refresh

    End If

    If From = 7 And Not IsNull(Adodc1.Recordset.Fields!member_id) And Adodc1.Recordset.RecordCount > 0 Then

        'X = InputBox("ÇÏÎá ÇáÑÞã Çæ ÌÒÁ ãä ÇáÑÞã", "ÔÇÔÉ ÇáÈÍË ÈÇáÑÞã")
        'If X = "" Then Exit Sub
        update_member.Adodc1.CommandType = adCmdText
        update_member.Adodc1.RecordSource = "select *  FROM members WHERE MEMBER_ID=" & Adodc1.Recordset.Fields!member_id
        update_member.Adodc1.Refresh

        If update_member.Adodc1.Recordset.RecordCount > 0 Then
            update_member.Text1.text = Adodc1.Recordset.Fields!member_id
            update_member.Text4.text = Adodc1.Recordset.Fields!member_name
        End If
    End If

    If From.Caption = 8 And Not IsNull(Adodc1.Recordset.Fields!member_id) And Adodc1.Recordset.RecordCount > 0 Then
        losed_card.Text2.text = Adodc1.Recordset.Fields!member_id
        losed_card.Text5.text = 0
        losed_card.Text5.Visible = False
        losed_card.Text3.text = Adodc1.Recordset.Fields!member_name
        losed_card.Text4.text = Adodc1.Recordset.Fields!MEMBER_TYPE
    End If

    If From.Caption = 9 And Not IsNull(Adodc1.Recordset.Fields!member_id) And Adodc1.Recordset.RecordCount > 0 Then
        services.Text1.text = Adodc1.Recordset.Fields!member_id
        services.Text6.text = Adodc1.Recordset.Fields!member_name
    End If

    If From = 10 And Not IsNull(Adodc1.Recordset.Fields!member_id) Then
        renew_member_activity.Text3 = Adodc1.Recordset.Fields!member_name
        renew_member_activity.Text2 = Adodc1.Recordset.Fields!member_id
        renew_member_activity.Label60 = Adodc1.Recordset.Fields!member_id
        renew_member_activity.Label70 = "0"
        renew_member_activity.Text4.text = Adodc1.Recordset.Fields!MEMBER_TYPE
    End If

    If From = 11 And Not IsNull(Adodc1.Recordset.Fields!member_id) And Adodc1.Recordset.RecordCount > 0 Then
        REPORTSFRM.Text1 = Adodc1.Recordset.Fields!member_id
        REPORTSFRM.Text2 = Adodc1.Recordset.Fields!member_name

    End If

    If From = 40 And Not IsNull(Adodc1.Recordset.Fields!member_id) And Adodc1.Recordset.RecordCount > 0 Then
        taslem.Combo3 = Adodc1.Recordset.Fields!member_name
        taslem.Combo5 = Adodc1.Recordset.Fields!MEMBER_TYPE

    End If

    If From = 50 And Not IsNull(Adodc1.Recordset.Fields!member_id) And Adodc1.Recordset.RecordCount > 0 Then
        ATTENDANCE.Text1.text = Adodc1.Recordset.Fields!member_name
        ATTENDANCE.Text2.text = Adodc1.Recordset.Fields!MEMBER_TYPE

    End If

    Unload Me

End Sub

Private Sub DataGrid1_Click()
    Command1_Click
End Sub

Private Sub Form_Load()
    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500

    connection_string = Cn.ConnectionString
    Adodc1.ConnectionString = connection_string
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = " select *  FROM members  "
    Adodc1.Refresh

End Sub

Private Sub Text1_Change()
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "select * FROM members where MEMBER_ID LIKE'%" & Text1.text & "%'"
    Adodc1.Refresh
 
End Sub

Private Sub Text2_Change()
    Adodc1.CommandType = adCmdText
    Adodc1.RecordSource = "select * FROM members where MEMBER_name LIKE'%" & Text2.text & "%'"
    Adodc1.Refresh
End Sub
