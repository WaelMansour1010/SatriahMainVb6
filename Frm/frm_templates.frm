VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frm_templates 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12840
   Icon            =   "frm_templates.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   3840
   ScaleWidth      =   12840
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frm_templates.frx":000C
      Height          =   2535
      Left            =   3480
      TabIndex        =   12
      Top             =   1080
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   4471
      _Version        =   393216
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
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "templates_id"
         Caption         =   "ÑÞã ÇáãÓÊäÏ"
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
         Caption         =   "ÇÓã ÇáãÓÊäÏ"
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
         Caption         =   "ÇáÞÓã"
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
      BeginProperty Column04 
         DataField       =   "no_of_images"
         Caption         =   "no_of_images"
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
      BeginProperty Column05 
         DataField       =   "subject_no"
         Caption         =   "subject_no"
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
            ColumnWidth     =   3495.118
         EndProperty
         BeginProperty Column02 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   3495.118
         EndProperty
         BeginProperty Column04 
            Object.Visible         =   0   'False
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column05 
            Object.Visible         =   0   'False
            ColumnWidth     =   915.024
         EndProperty
      EndProperty
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Bindings        =   "frm_templates.frx":0021
      DataSource      =   "Adodc3"
      Height          =   315
      Left            =   6240
      TabIndex        =   11
      Top             =   720
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "iso_departement_name"
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ÊÍÏíÏ"
      Height          =   255
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Frame Frame6 
      Height          =   1575
      Left            =   0
      TabIndex        =   5
      Top             =   1080
      Width           =   1335
      Begin VB.CommandButton Command1 
         Caption         =   "ÈÍË ÈÇáÇÓã"
         Height          =   255
         Index           =   5
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "ÈÍË ÈÇáÑÞã"
         Height          =   255
         Index           =   4
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "ÈÍË"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.TextBox Text2 
      DataField       =   "opr_id"
      DataSource      =   "templates_details"
      Height          =   285
      Left            =   7920
      TabIndex        =   4
      Text            =   "Text2"
      Top             =   5400
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text1 
      DataField       =   "templates_id"
      DataSource      =   "Adodc2"
      Height          =   285
      Left            =   7680
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   3120
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   465
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   820
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
      Caption         =   "ÊÍÑíß"
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
      Height          =   465
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   820
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
      Caption         =   "ÊÍÑíß"
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
      Height          =   465
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   820
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
      Caption         =   "ÊÍÑíß"
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
      Height          =   465
      Left            =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   820
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
      Caption         =   "ÊÍÑíß"
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
      Caption         =   "ÊÍÏíÏ ÇáÞÓã"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   11160
      TabIndex        =   7
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label case_id 
      Caption         =   "1000"
      Height          =   615
      Left            =   -120
      TabIndex        =   2
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
      TabIndex        =   1
      Top             =   6960
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "ÇÎÊÇÑ ÇáäãæÐÌ ÇáãÑÇÏ ÇáÊÚÇãá ãÚÉ"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12855
   End
End
Attribute VB_Name = "frm_templates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim X As String

Private Sub Command1_Click(Index As Integer)
Select Case Index

Case 0
    Adodc1.Recordset.AddNew
Case 1
    Adodc1.Recordset.Update

Case 2

         If my_language = "E" Then
         X = MsgBox("Confirm Deletion", vbCritical + vbYesNo)
        Else
        X = MsgBox("åá ÇäÊ ãÊÃßÏ ãä ÇáÍÐÝ", vbCritical + vbYesNo)
        
        End If
If X = vbNo Then
Exit Sub
End If

    If Adodc1.Recordset.RecordCount > 0 Then
    Adodc1.Recordset.Delete
    Adodc1.Refresh
    DataGrid1.Refresh
    End If

Case 3
    If Adodc1.Recordset.RecordCount > 0 Then
    
    Form3.case_id = Me.name
   
    Form3.Show
    End If

Case 4
On Error Resume Next
If my_language = "E" Then
X = InputBox("Enter Form No.")
Else

X = InputBox("ÇÏÎá   ÑÞã  ÇáäãæÐÌ ÇáãØáæÈ ÇáÈÍË ÚäÉ")
End If
        If IsNumeric(X) Then
        Adodc1.CommandType = adCmdText
        Adodc1.RecordSource = "select * from  templates where   subject_no=0 and templates_id=" & X
        Adodc1.Refresh
        Else
        If my_language = "E" Then
        MsgBox " Must be  Numeric", vbCritical
        Else
        MsgBox "áÇÈÏ ãä ÇÏÎÇá ÑÞã ÝÞØ", vbCritical
        End If
        End If

Case 5
If my_language = "E" Then
X = InputBox("Enter Form Name")
Else

    X = InputBox("ÇÏÎá ßáãÉ ÇáÈÍË")
End If
            Adodc1.CommandType = adCmdText
        Adodc1.RecordSource = "select * from templates where subject_no=0 and  templates_name like '%" & X & "%'"
        Adodc1.Refresh

Case 5
    X = InputBox("ÇÏÎá ßáãÉ ÇáÈÍË")
            Adodc1.CommandType = adCmdText
        Adodc1.RecordSource = "select * from templates where subject_no=0 and  departement_name like '%" & X & "%'"
        Adodc1.Refresh
        

End Select

End Sub

Private Sub Command2_Click()
Label4_Click
End Sub

Private Sub DataCombo1_Click(Area As Integer)

Adodc1.RecordSource = "select * from templates where departement_name='" & DataCombo1.text & "'"
Adodc1.Refresh
End Sub

Private Sub ChangeLang()
 

Me.Caption = "Select Form"
Label3.Caption = Me.Caption
 Label2.Caption = "DEPT."
 Command1(4).Caption = "By ID"
 Command1(5).Caption = "By Name"
 
 Label1.Caption = "Search"
Command2.Caption = "Select"
DataGrid1.RightToLeft = False
DataGrid1.Columns(0).Caption = "ID"
DataGrid1.Columns(1).Caption = "Form Name"

DataGrid1.Columns(3).Caption = "Departement"


End Sub


Private Sub Form_Load()
 Me.left = (MDIFrmMain.Width - Me.Width) / 2
    Me.top = (MDIFrmMain.Height - Me.Height) / 2 - 500

On Error Resume Next
LoadSettings
If my_language = "E" Then
  SetInterface Me
    ChangeLang
End If


Adodc1.ConnectionString = connection_string
Adodc1.CommandType = adCmdText
Adodc1.RecordSource = "select * from templates where subject_no=0"
Adodc1.Refresh


Adodc2.ConnectionString = connection_string
Adodc2.CommandType = adCmdText
Adodc2.RecordSource = "select * from  templates_details"
Adodc2.Refresh


Adodc3.ConnectionString = connection_string
Adodc3.CommandType = adCmdText
Adodc3.RecordSource = "select * from  iso_department"
Adodc3.Refresh


templates_details.ConnectionString = connection_string
templates_details.CommandType = adCmdText
 



 

 
End Sub

Private Sub Label4_Click()
Dim template_id As Integer
Dim template_name, departement_name As String
Dim subject_no, no_of_images As Integer
'On Error Resume Next



If case_id = 0 Then
template_id = Adodc1.Recordset.Fields!templates_id
template_name = Adodc1.Recordset.Fields!templates_name
subject_no = imaged.subject_no.Caption
no_of_images = Adodc1.Recordset.Fields!no_of_images
    If IsNull(Adodc1.Recordset.Fields!departement_name) Then
    departement_name = ""
    Else
    departement_name = Adodc1.Recordset.Fields!departement_name
    End If

templates_details.CommandType = adCmdText
templates_details.RecordSource = "select * from templates_details where templates_id=" & template_id ' & "and subject_no=0"
templates_details.Refresh
    If templates_details.Recordset.RecordCount = 0 Then
    If my_language = "E" Then
  MsgBox "This For  Have An Error", vbCritical
   Else

    MsgBox "åÐÇ ÇáäãæÐÌ ÈÉ ÎØÃ íÑÌí ÇÎÊíÇÑ äãæÐÌ ÇÎÑ", vbCritical
   End If
    Exit Sub
    End If



Adodc1.Recordset.AddNew 'templates table
'Adodc1.Recordset.Fields!templates_id = templates_id
Adodc1.Recordset.Fields!templates_name = template_name
Adodc1.Recordset.Fields!subject_no = subject_no
Adodc1.Recordset.Fields!no_of_images = no_of_images
Adodc1.Recordset.Fields!departement_name = departement_name
Adodc1.Recordset.Update
Adodc1.Recordset.MoveLast

imaged.Adodc3.Recordset.AddNew 'subject_templates_table
imaged.Adodc3.Recordset.Fields!subject_no = subject_no
imaged.Adodc3.Recordset.Fields!template_id = Adodc1.Recordset.Fields!templates_id
imaged.Adodc3.Recordset.Fields!template_name = template_name
imaged.Adodc3.Recordset.Fields!no_of_images = no_of_images
imaged.Adodc3.Recordset.Fields!date_added = DateValue(Now)

imaged.Adodc3.Recordset.Update
imaged.Adodc3.Refresh
imaged.DataGrid1.Refresh



templates_details.CommandType = adCmdText
templates_details.RecordSource = "select * from templates_details where templates_id=" & template_id ' & "and subject_no=0"
templates_details.Refresh
    If templates_details.Recordset.RecordCount > 0 Then
    templates_details.Recordset.MoveFirst
    End If

For i = 1 To templates_details.Recordset.RecordCount
 Adodc2.Recordset.AddNew ' templates_details table
 Adodc2.Recordset.Fields!templates_id = Adodc1.Recordset.Fields!templates_id 'templates_details.Recordset.Fields!templates_id
 Adodc2.Recordset.Fields!X1 = templates_details.Recordset.Fields!X1
 Adodc2.Recordset.Fields!x2 = templates_details.Recordset.Fields!x2
 Adodc2.Recordset.Fields!Y1 = templates_details.Recordset.Fields!Y1
 Adodc2.Recordset.Fields!y2 = templates_details.Recordset.Fields!y2
 Adodc2.Recordset.Fields!text = templates_details.Recordset.Fields!text
 Adodc2.Recordset.Fields!image_id = templates_details.Recordset.Fields!image_id
 Adodc2.Recordset.Fields!IMAGE_NAME = templates_details.Recordset.Fields!IMAGE_NAME
 Adodc2.Recordset.Fields!color = templates_details.Recordset.Fields!color
 Adodc2.Recordset.Fields!BackColor = templates_details.Recordset.Fields!BackColor
 Adodc2.Recordset.Fields!FontName = templates_details.Recordset.Fields!FontName
 Adodc2.Recordset.Fields!FontSize = templates_details.Recordset.Fields!FontSize
 Adodc2.Recordset.Fields!FontBold = templates_details.Recordset.Fields!FontBold
 Adodc2.Recordset.Fields!FontItalic = templates_details.Recordset.Fields!FontItalic
 Adodc2.Recordset.Fields!FontUnderline = templates_details.Recordset.Fields!FontUnderline
 Adodc2.Recordset.Fields!Strikethrough = templates_details.Recordset.Fields!Strikethrough
 Adodc2.Recordset.Fields!image_direction = templates_details.Recordset.Fields!image_direction
 Adodc2.Recordset.Fields!subject_no = subject_no
Adodc2.Recordset.Update
templates_details.Recordset.MoveNext
Next i



Me.Hide

Else
If Adodc1.Recordset.RecordCount <> 0 Then
loading_temolates.Show
loading_temolates.Label6.Caption = Adodc1.Recordset.Fields!templates_id
'loading_temolates.Label7.Caption = Adodc1.Recordset.Fields!IMAGE_NAME
loading_temolates.Label9.Caption = Adodc1.Recordset.Fields!no_of_images
Me.Hide
End If


End If



End Sub
