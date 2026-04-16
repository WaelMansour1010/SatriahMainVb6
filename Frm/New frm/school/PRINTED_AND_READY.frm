VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{D155F1AE-D9A4-458C-8CEE-498CB717DB7B}#1.0#0"; "DBPix20.ocx"
Begin VB.Form PRINTED_AND_READY 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "                                                         ‘«·‘… «·ﬂ«—‰ÌÂ«  «·Ã«Â“… ·· ”·Ì„"
   ClientHeight    =   8475
   ClientLeft      =   90
   ClientTop       =   390
   ClientWidth     =   8265
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8475
   ScaleWidth      =   8265
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
   Begin VB.CommandButton Command1 
      Caption         =   " "
      Height          =   855
      Left            =   7080
      Picture         =   "PRINTED_AND_READY.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   4200
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   4200
      TabIndex        =   26
      Top             =   7560
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   240
      TabIndex        =   24
      Top             =   7080
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   4200
      TabIndex        =   22
      Top             =   7200
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ÿ»«⁄… «·ﬂ·"
      Height          =   192
      Left            =   4800
      TabIndex        =   2
      Top             =   10680
      Width           =   1935
   End
   Begin VB.CommandButton cmdprint 
      BackColor       =   &H0080FF80&
      Caption         =   " „  ”·Ì„ «·ﬂ«—‰Ì… «·„Õœœ ›ﬁÿ"
      Height          =   735
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6240
      Width           =   3255
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0000FF00&
      Caption         =   " „  ”·Ì„ ﬂ· «·ﬂ«—‰ÌÂ« "
      Height          =   732
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6240
      Width           =   3135
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "PRINTED_AND_READY.frx":1992
      Height          =   2055
      Left            =   480
      TabIndex        =   3
      Top             =   4080
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   3625
      _Version        =   393216
      AllowUpdate     =   -1  'True
      ColumnHeaders   =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
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
      ColumnCount     =   13
      BeginProperty Column00 
         DataField       =   "inedx"
         Caption         =   "inedx"
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
      BeginProperty Column02 
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
      BeginProperty Column03 
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
      BeginProperty Column04 
         DataField       =   "SELECTED"
         Caption         =   "SELECTED"
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
         DataField       =   "update_year"
         Caption         =   "update_year"
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
         DataField       =   "sex"
         Caption         =   "sex"
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
         DataField       =   "OPR_TYPE"
         Caption         =   "OPR_TYPE"
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
      BeginProperty Column09 
         DataField       =   "PRINTED"
         Caption         =   "PRINTED"
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
         DataField       =   "BILL_NO"
         Caption         =   "BILL_NO"
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
         DataField       =   "CENTER_MANAGER"
         Caption         =   "CENTER_MANAGER"
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
         DataField       =   "RECIVED"
         Caption         =   "RECIVED"
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
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   854.929
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1140.095
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   764.787
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   764.787
         EndProperty
      EndProperty
   End
   Begin DBPIXLib.DBPix20 DBPIX1 
      DataField       =   "IMAGE_PATH"
      DataSource      =   "Adodc1"
      Height          =   1335
      Left            =   960
      TabIndex        =   4
      Top             =   720
      Width           =   1455
      _Version        =   131072
      _ExtentX        =   2566
      _ExtentY        =   2355
      _StockProps     =   1
      _Image          =   "PRINTED_AND_READY.frx":19A7
      ImageResampleWidth=   100
      ImageResampleHeight=   100
      ImageResampleMode=   1
      ImageSaveFormat =   0
      JPEGQuality     =   75
      JPEGEncoding    =   0
      JPEGColorMode   =   0
      JPEGNoRecompress=   -1  'True
      JPEGRotateWarning=   0
      PNGColorDepth   =   0
      PNGCompression  =   0
      PNGFilter       =   0
      PNGInterlace    =   1
      ImageDitherMethod=   3
      ImagePaletteMethod=   4
      ImagePreviewMode=   0   'False
      ImageKeepMetaData=   -1  'True
      UseAmbientBackcolor=   -1  'True
      ViewAsyncDecoding=   -1  'True
      ViewEnableMouseZoom=   -1  'True
      ViewInitialZoom =   0
      ViewHAlign      =   1
      ViewVAlign      =   1
      ViewMenuMode    =   0
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   492
      Left            =   6120
      Top             =   360
      Visible         =   0   'False
      Width           =   2160
      _ExtentX        =   3810
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
      RecordSource    =   "SELECT * FROM ready_to_print WHERE PRINTED=1 AND RECIVED=0"
      Caption         =   ""
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
   Begin VB.Label Label23 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   960
      TabIndex        =   31
      Top             =   8040
      Width           =   1935
   End
   Begin VB.Label Label22 
      Caption         =   "«· «—ÌŒ"
      Height          =   255
      Left            =   3120
      TabIndex        =   30
      Top             =   8040
      Width           =   855
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   4200
      TabIndex        =   29
      Top             =   8040
      Width           =   1935
   End
   Begin VB.Label Label20 
      Caption         =   "»Ê«”ÿ…"
      Height          =   255
      Left            =   6480
      TabIndex        =   28
      Top             =   8040
      Width           =   855
   End
   Begin VB.Label Label18 
      Caption         =   " ·Ì›Ê‰ «·„” ·„"
      Height          =   495
      Left            =   6240
      TabIndex        =   25
      Top             =   7560
      Width           =   1575
   End
   Begin VB.Label Label13 
      Caption         =   "—ﬁ„ ÂÊÌ… «·„” ·„"
      Height          =   495
      Left            =   2280
      TabIndex        =   23
      Top             =   7200
      Width           =   2175
   End
   Begin VB.Label Label12 
      Caption         =   "«”„  «·„” ·„"
      Height          =   495
      Left            =   6360
      TabIndex        =   21
      Top             =   7200
      Width           =   1575
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      Caption         =   "«·”‰… «·œ—«”Ì…"
      Height          =   330
      Left            =   5760
      TabIndex        =   20
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      Caption         =   "—ﬁ„ «·ÿ«·»"
      Height          =   330
      Left            =   2280
      TabIndex        =   19
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      Caption         =   "«”„ «·ÿ«·»"
      Height          =   330
      Left            =   3840
      TabIndex        =   18
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      Caption         =   "„”·”·"
      Height          =   330
      Left            =   960
      TabIndex        =   17
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label LabelX 
      Alignment       =   2  'Center
      Caption         =   "«·ﬂ«—‰ÌÂ«  ·Ã«Â“… ·· ”·Ì„"
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
      Height          =   330
      Left            =   360
      TabIndex        =   16
      Top             =   3000
      Width           =   6375
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      DataField       =   "CENTER_MANAGER"
      DataSource      =   "Adodc1"
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   960
      TabIndex        =   15
      Top             =   2400
      Width           =   1212
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "„œÌ— ⁄«„ «·„—ﬂ“"
      DataSource      =   "Adodc1"
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   1080
      TabIndex        =   14
      Top             =   2160
      Width           =   1212
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      DataField       =   "update_year"
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
      ForeColor       =   &H00FF0000&
      Height          =   372
      Left            =   2280
      TabIndex        =   13
      Top             =   2520
      Width           =   2532
   End
   Begin VB.Shape Shape1 
      Height          =   2535
      Left            =   720
      Top             =   600
      Width           =   5295
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "—ﬁ„ «·«Ì’«·"
      ForeColor       =   &H00000000&
      Height          =   372
      Left            =   4800
      TabIndex        =   12
      Top             =   2280
      Width           =   1212
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      DataField       =   "BILL_NO"
      DataSource      =   "Adodc1"
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   2520
      TabIndex        =   11
      Top             =   2280
      Width           =   2052
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "—ﬁ„ «·„”·”·"
      ForeColor       =   &H00000000&
      Height          =   372
      Left            =   4800
      TabIndex        =   10
      Top             =   1920
      Width           =   1212
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      DataField       =   "inedx"
      DataSource      =   "Adodc1"
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   2520
      TabIndex        =   9
      Top             =   1920
      Width           =   2052
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "—ﬁ„ «·ÿ«·»"
      ForeColor       =   &H00000000&
      Height          =   372
      Left            =   4800
      TabIndex        =   8
      Top             =   1560
      Width           =   1212
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      DataField       =   "MEMBER_ID"
      DataSource      =   "Adodc1"
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   2520
      TabIndex        =   7
      Top             =   1560
      Width           =   2052
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      DataField       =   "MEMBER_NAME"
      DataSource      =   "Adodc1"
      ForeColor       =   &H00000000&
      Height          =   252
      Left            =   2640
      TabIndex        =   6
      Top             =   1080
      Width           =   2052
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      DataField       =   "MEMBER_TYPE"
      DataSource      =   "Adodc1"
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
      Height          =   372
      Left            =   2160
      TabIndex        =   5
      Top             =   600
      Width           =   2052
   End
End
Attribute VB_Name = "PRINTED_AND_READY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdprint_Click()
If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Then
MsgBox "·«»œ „‰ ﬂ «»… »Ì«‰«  «·„” ·„ ﬂ«„·…", vbCritical
Exit Sub
End If

Dim X As Integer
X = MsgBox("Â· «‰  „ √„œ „‰ Â–… «·⁄„·Ì…", vbExclamation + vbYesNo)
If X = vbNo Then
Exit Sub
End If

If Adodc1.Recordset.RecordCount > 0 Then
If Adodc1.Recordset.EOF = False And Adodc1.Recordset.BOF = False Then

Adodc1.Recordset.Fields!recived_name = Text1.Text
Adodc1.Recordset.Fields!SSN = Text2.Text
Adodc1.Recordset.Fields!recived_tel = Text3.Text

Adodc1.Recordset.Fields!recived_user_name = main.txtusername


Adodc1.Recordset.Fields!received_date = DateValue(Now)
Adodc1.Recordset.Fields!recived = True
Adodc1.Recordset.Update
Adodc1.Refresh
DataGrid1.Refresh
End If
End If

'Adodc1.CommandType = adCmdText
'Adodc1.RecordSource = "SELECT * FROM ready_to_print WHERE PRINTED=1 AND RECIVED=1"
'Adodc1.Refresh

'For i = 1 To Adodc1.Recordset.RecordCount
'Adodc1.Recordset.Fields!recived = True
'Adodc1.Recordset.Update

'Adodc1.Recordset.MoveNext
'Next i

'Adodc1.CommandType = adCmdText
'Adodc1.RecordSource = "SELECT * FROM ready_to_print WHERE PRINTED=1 AND RECIVED=0"
'Adodc1.Refresh


End Sub

Private Sub Command1_Click()
X = InputBox("„‰ ›÷·ﬂ «œŒ· —ﬁ„ «·⁄÷Ê ··»ÕÀ ⁄‰…")
Adodc1.RecordSource = "select * from ready_to_print where member_id like'" & X & "%' and RECIVED=0"
Adodc1.Refresh

If X = "" Then

Adodc1.RecordSource = "select * from ready_to_print where  RECIVED=0"
Adodc1.Refresh
End If
End Sub

Private Sub Command3_Click()
Dim i As Integer
Dim X As Integer
If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Then
MsgBox "·«»œ „‰ ﬂ «»… »Ì«‰«  «·„” ·„ ﬂ«„·…", vbCritical
Exit Sub
End If

X = MsgBox("Â· «‰  „ √„œ „‰ Â–… «·⁄„·Ì…", vbExclamation + vbYesNo)
If X = vbNo Then
Exit Sub
End If

If Adodc1.Recordset.RecordCount = 0 Then Exit Sub
Adodc1.Recordset.MoveFirst

For i = 1 To Adodc1.Recordset.RecordCount

 Adodc1.Recordset.Fields!recived_name = Text1.Text
Adodc1.Recordset.Fields!SSN = Text2.Text
Adodc1.Recordset.Fields!recived_tel = Text3.Text
Adodc1.Recordset.Fields!recived_user_name = main.txtusername
Adodc1.Recordset.Fields!received_date = DateValue(Now)
Adodc1.Recordset.Fields!recived = True
 
 

Adodc1.Recordset.Fields!recived = True
Adodc1.Recordset.Update

Adodc1.Recordset.MoveNext
Next i



Adodc1.CommandType = adCmdText
Adodc1.RecordSource = "SELECT * FROM ready_to_print WHERE PRINTED=1 AND RECIVED=0"
Adodc1.Refresh

 




End Sub
 
Private Sub Command6_Click()

End Sub

Private Sub Form_Load()
Label19.Caption = main.txtusername
Label23.Caption = DateValue(Now)
End Sub
