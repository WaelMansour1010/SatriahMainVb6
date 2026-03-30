VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form bus 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9360
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11460
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9360
   ScaleWidth      =   11460
   Begin VB.CommandButton Command5 
      BackColor       =   &H000080FF&
      Caption         =   "ÇáãÑÝÞÇÊ"
      Height          =   495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   6240
      Width           =   1935
   End
   Begin VB.CommandButton Command4 
      Caption         =   "ÇÖÇÝÉ"
      Height          =   495
      Left            =   120
      TabIndex        =   26
      Top             =   6960
      Width           =   2055
   End
   Begin VB.ComboBox Combo4 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      ItemData        =   "Hafela.frx":0000
      Left            =   3000
      List            =   "Hafela.frx":000D
      TabIndex        =   25
      Top             =   6240
      Width           =   2295
   End
   Begin VB.ComboBox Combo3 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      ItemData        =   "Hafela.frx":0031
      Left            =   6720
      List            =   "Hafela.frx":003B
      TabIndex        =   24
      Top             =   6240
      Width           =   2295
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ÇÖÇÝÉ"
      Height          =   495
      Left            =   480
      TabIndex        =   19
      Top             =   2760
      Width           =   2055
   End
   Begin VB.ComboBox Combo2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      ItemData        =   "Hafela.frx":0054
      Left            =   2880
      List            =   "Hafela.frx":005E
      TabIndex        =   18
      Top             =   2880
      Width           =   2295
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      ItemData        =   "Hafela.frx":0077
      Left            =   7200
      List            =   "Hafela.frx":0081
      TabIndex        =   17
      Top             =   2880
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000000FF&
      Caption         =   "ÌÏíÏ"
      Height          =   375
      Left            =   360
      TabIndex        =   13
      Top             =   8520
      Width           =   1575
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      DataField       =   "member_name"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   960
      TabIndex        =   12
      Top             =   1200
      Width           =   2295
   End
   Begin VB.OptionButton Option2 
      Caption         =   "ááÛíÑ äÙíÑ äÓÉ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7200
      TabIndex        =   10
      Top             =   1320
      Width           =   1935
   End
   Begin VB.OptionButton Option1 
      Caption         =   "ÊÇÈÚå ááãÏÑÓÉ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9600
      TabIndex        =   9
      Top             =   1320
      Width           =   1935
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      DataField       =   "member_name"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   960
      TabIndex        =   8
      Top             =   600
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      DataField       =   "member_name"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   492
      Left            =   7200
      TabIndex        =   6
      Text            =   "1"
      Top             =   600
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000000FF&
      Caption         =   "ÍÝÙ"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   8880
      Width           =   1575
   End
   Begin VB.TextBox Text6 
      DataField       =   "INSTALLMENTS_TOTAL"
      DataSource      =   "Adodc5"
      Height          =   288
      Left            =   4440
      TabIndex        =   1
      Text            =   "Text6"
      Top             =   5280
      Width           =   492
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      DataField       =   "member_type"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   7200
      TabIndex        =   0
      Top             =   1920
      Width           =   2295
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Hafela.frx":0094
      Height          =   2415
      Left            =   4200
      TabIndex        =   20
      Top             =   3480
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   4260
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      ColumnHeaders   =   -1  'True
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
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "FINES_NO"
         Caption         =   "ã"
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
         DataField       =   "FINES_VALUE"
         Caption         =   "ÇáÍí"
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
         DataField       =   "PAYED_DATE"
         Caption         =   "ÇáÔÇÑÚ"
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
         DataField       =   "ACTIVATED"
         Caption         =   "ACTIVATED"
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
            ColumnWidth     =   2505.26
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2505.26
         EndProperty
         BeginProperty Column03 
            Object.Visible         =   0   'False
            ColumnWidth     =   915.024
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "Hafela.frx":00A9
      Height          =   2175
      Left            =   3720
      TabIndex        =   27
      Top             =   6840
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   3836
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      ColumnHeaders   =   -1  'True
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
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "FINES_NO"
         Caption         =   "ã"
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
         DataField       =   "FINES_VALUE"
         Caption         =   "ÇÓã ÇáØÇáÈ"
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
         DataField       =   "PAYED_DATE"
         Caption         =   "äæÚ ÇáäÞá"
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
         DataField       =   "ACTIVATED"
         Caption         =   "ACTIVATED"
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
            ColumnWidth     =   2505.26
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2505.26
         EndProperty
         BeginProperty Column03 
            Object.Visible         =   0   'False
            ColumnWidth     =   915.024
         EndProperty
      EndProperty
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      Caption         =   "äæÚ ÇáäÞá"
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
      Left            =   5040
      TabIndex        =   23
      Top             =   6240
      Width           =   1815
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Caption         =   "ÇÓã ÇáØÇáÈ"
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
      Left            =   9120
      TabIndex        =   22
      Top             =   6240
      Width           =   1815
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "ÇáØáÇÈ ÇáãÔÊÑßíä"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9120
      TabIndex        =   21
      Top             =   5880
      Width           =   2535
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "ãä ÔÇÑÚ"
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
      TabIndex        =   16
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "ãä Íí"
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
      Left            =   9720
      TabIndex        =   15
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "ÇáÎØ"
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
      Left            =   9840
      TabIndex        =   14
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "ÇáäÓÈÉ"
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
      Left            =   3480
      TabIndex        =   11
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "ÇáÓÇÆÞ"
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
      Left            =   3480
      TabIndex        =   7
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "ÑÞã ÇáãÚÏå/ÇáÓíÇÑÉ"
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
      Left            =   9720
      TabIndex        =   5
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "ÊÚÑíÝ ÇáÍÇÝáÉ"
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
      Left            =   3600
      TabIndex        =   4
      Top             =   0
      Width           =   3975
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "ÚÏÏ ÇáãÞÇÚÏ"
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
      Left            =   9840
      TabIndex        =   3
      Top             =   1920
      Width           =   1815
   End
End
Attribute VB_Name = "bus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command5_Click()
    On Error Resume Next
    'If my_language = "E" Then
    'If Text2.Text = "" Then MsgBox "Select Voucher First": Exit Sub

    'Else
    'If Text2.Text = "" Then MsgBox "áÇÈÏ ãä ÇÍÊíÇÑ ØÇáÈ  ÇæáÇ": Exit Sub
    'End If

    imaged.Show
    imaged.txtopeation_type = "ßÑæßí ÇáØÇáÈ"
    imaged.SUBJECT_NO = Text2.text

    If my_language = "E" Then
        imaged.Label6.Caption = "studnet #"
        imaged.Caption = "student Attachments"
    Else
        imaged.Label6.Caption = "ÑÞã ÇáØÇáÈ"
        imaged.Caption = "ßÑæßí ÇáØÇáÈ"
    End If

    imaged.Adodc1.CommandType = adCmdText
    imaged.Adodc1.RecordSource = "SELECT * FROM subjects_images WHERE operation_type ='ßÑæßí ÇáØÇáÈ' and subject_no='" & Text2.text & "'"
    imaged.Adodc1.Refresh

    If imaged.Adodc1.Recordset.RecordCount > 0 Then

        imaged.DBPix201.Visible = True
    Else
        imaged.DBPix201.Visible = False
    End If

End Sub

Private Sub Form_Load()
    system_path = App.path
    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500

    'connection_string = Cn.ConnectionString
    'Adodc1.ConnectionString = connection_string
    'Adodc1.CommandType = adCmdText
    'Adodc1.RecordSource = "select * from  "
    'Adodc1.Refresh

End Sub
