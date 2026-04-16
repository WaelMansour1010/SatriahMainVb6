VERSION 5.00
Object = "{3C62B3DD-12BE-4941-A787-EA25415DCD27}#10.0#0"; "crviewer.dll"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÔÇÔÉ ÇáØÈÇÚå"
   ClientHeight    =   11490
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   15270
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11490
   ScaleWidth      =   15270
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
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame6 
      Caption         =   "ØÈÇÚÉ ÇáßÇÑäíåÇÊ"
      Height          =   1935
      Left            =   13200
      RightToLeft     =   -1  'True
      TabIndex        =   45
      Top             =   6360
      Visible         =   0   'False
      Width           =   1935
      Begin VB.CommandButton Command22 
         Caption         =   "ØÈÇÚÉ  ÚãæÏíä"
         Height          =   375
         Left            =   360
         TabIndex        =   47
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton Command21 
         Caption         =   "ØÈÇÚÉ ÚãæÏ æÇÍÏ"
         Height          =   375
         Left            =   360
         TabIndex        =   46
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.Frame Frame5 
      Height          =   975
      Left            =   360
      TabIndex        =   37
      Top             =   120
      Width           =   12735
      Begin VB.Label case_id 
         Caption         =   "7"
         Height          =   615
         Left            =   240
         TabIndex        =   43
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "ÑÞã ÇáÊÞÑíÑ"
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
         Left            =   2160
         TabIndex        =   42
         Top             =   360
         Width           =   1335
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
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   4200
         TabIndex        =   41
         Top             =   360
         Width           =   2775
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
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   8400
         TabIndex        =   40
         Top             =   360
         Width           =   2775
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ãä"
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
         Left            =   10080
         TabIndex        =   39
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Åáì"
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
         TabIndex        =   38
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "ÇÎÊÑ ÇäæÇÚ ÇáÚÖæíÉ ÇáãÑÇÏ ÚÑÖåÇ"
      Height          =   4335
      Left            =   12840
      RightToLeft     =   -1  'True
      TabIndex        =   30
      Top             =   1680
      Visible         =   0   'False
      Width           =   2415
      Begin VB.CommandButton Command19 
         BackColor       =   &H0000FF00&
         Caption         =   "ÚÑÖ ÇáÊÞÑíÑ"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   3240
         Width           =   2055
      End
      Begin VB.CommandButton Command11 
         Caption         =   "ÇÖÇÝÉ Çáì ÇáÊÞÑíÑ"
         Height          =   375
         Left            =   240
         TabIndex        =   33
         Top             =   960
         Width           =   1815
      End
      Begin VB.CommandButton Command18 
         Caption         =   "ÍÐÝ ãä ÇáÊÞÑíÑ"
         Height          =   375
         Left            =   360
         TabIndex        =   32
         Top             =   2760
         Width           =   1815
      End
      Begin VB.ListBox List1 
         Height          =   1035
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   1680
         Width           =   2055
      End
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "Form3.frx":0000
         Height          =   315
         Left            =   120
         TabIndex        =   34
         Top             =   480
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "MEMBER_NAME"
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSAdodcLib.Adodc Adodc2 
         Height          =   375
         Left            =   0
         Top             =   3840
         Visible         =   0   'False
         Width           =   2535
         _ExtentX        =   4471
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
         RecordSource    =   "MEMBER_TYPES"
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
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "ÇäæÇÚ ÇáÚÖæíÉ ÇáÊí ÓÊÙåÑ Ýí ÇáÊÞÑíÑ"
         Height          =   495
         Left            =   0
         TabIndex        =   35
         Top             =   1440
         Width           =   2415
      End
   End
   Begin VB.Frame Frame2 
      Height          =   3615
      Left            =   13080
      TabIndex        =   19
      Top             =   1680
      Visible         =   0   'False
      Width           =   1695
      Begin VB.CommandButton Command20 
         Caption         =   "ÇáãÕÑæÝÇÊ ÝÞØ"
         Height          =   495
         Left            =   120
         TabIndex        =   44
         Top             =   2760
         Width           =   1455
      End
      Begin VB.CommandButton Command16 
         Caption         =   "ßá ÇáÚãáíÇÊ ÇáÊí ÊãÊ ÈÇáÎÒíäÉ"
         Height          =   495
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton Command15 
         Caption         =   "ÊÌÏíÏ æÚÖæíÉ ÌÏíÏÉ"
         Height          =   495
         Left            =   120
         TabIndex        =   22
         Top             =   2040
         Width           =   1455
      End
      Begin VB.CommandButton Command14 
         Caption         =   "ÊÌÏíÏ ÇáÚÖæíÉ ÝÞØ"
         Height          =   495
         Left            =   120
         TabIndex        =   21
         Top             =   1440
         Width           =   1455
      End
      Begin VB.CommandButton Command13 
         Caption         =   "ÇáÚÖæíÉ ÇáÌÏíÏÉ ÝÞØ"
         CausesValidation=   0   'False
         Height          =   495
         Left            =   120
         TabIndex        =   20
         Top             =   840
         Width           =   1455
      End
   End
   Begin VB.Frame Frame3 
      Height          =   2415
      Left            =   12960
      TabIndex        =   23
      Top             =   1560
      Visible         =   0   'False
      Width           =   2055
      Begin VB.TextBox Text3 
         DataField       =   "VALUE"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   360
         TabIndex        =   28
         Text            =   "Text3"
         Top             =   2040
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.ComboBox Combo1 
         CausesValidation=   0   'False
         Height          =   315
         Left            =   240
         TabIndex        =   26
         Top             =   1200
         Width           =   1575
      End
      Begin VB.CommandButton Command17 
         Caption         =   "ÊÞÑíÑ ;ßá áÇäÔØÉ"
         Height          =   495
         Left            =   240
         TabIndex        =   25
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton Command12 
         Caption         =   "ÊÞÑíÑ ÇáäÔÇØ ÇáãÍÏÏ"
         Height          =   495
         Left            =   360
         TabIndex        =   24
         Top             =   1560
         Width           =   1455
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   615
         Left            =   -1080
         Top             =   2400
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   1085
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
         RecordSource    =   "ACTIVITIES_type"
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
         Caption         =   "ÇÎÊÇÑ äÔÇØ ãÍÏÏ"
         Height          =   255
         Left            =   600
         TabIndex        =   27
         Top             =   840
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4455
      Left            =   13080
      TabIndex        =   12
      Top             =   1560
      Visible         =   0   'False
      Width           =   1935
      Begin VB.CommandButton Command10 
         Caption         =   "ÇáßÇÑäíåÇÊ ÇáÚÇÏíÉ"
         Height          =   495
         Left            =   120
         TabIndex        =   18
         Top             =   3840
         Width           =   1815
      End
      Begin VB.CommandButton Command9 
         Caption         =   "ÇáßÇÑäíåÇÊ   ÇáÊí Êã ÇÚÇÏÉ ØÈÇÚÊåÇ"
         Height          =   495
         Left            =   120
         TabIndex        =   17
         Top             =   3240
         Width           =   1815
      End
      Begin VB.CommandButton Command8 
         Caption         =   "ÇáßÇÑäíåÇÊ ÇáÈÏá ÝÇÆÏ"
         Height          =   495
         Left            =   120
         TabIndex        =   16
         Top             =   2640
         Width           =   1815
      End
      Begin VB.CommandButton Command7 
         Caption         =   "ÇáßÇÑäíåÇÊ ÇáãØÈæÚÉ  ÇáÊí Êã ÊÓáíãåÇ"
         Height          =   495
         Left            =   120
         TabIndex        =   15
         Top             =   1080
         Width           =   1815
      End
      Begin VB.CommandButton Command6 
         Caption         =   "ßá ÇáßÇÑäíåÇÊ ÇáãØÈæÚÉ"
         Height          =   495
         Left            =   120
         TabIndex        =   14
         Top             =   480
         Width           =   1815
      End
      Begin VB.CommandButton Command5 
         Caption         =   "ÇáßÇÑäíåÇÊ ÇáãØÈæÚÉ  ÇáÊí áã ÊÓáã ÈÚÏ"
         Height          =   495
         Left            =   120
         TabIndex        =   13
         Top             =   1680
         Width           =   1815
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Çáßá"
      Height          =   375
      Left            =   13680
      TabIndex        =   11
      Top             =   2880
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ÇäËì"
      Height          =   375
      Left            =   13680
      TabIndex        =   10
      Top             =   2400
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ÐßÑ"
      Height          =   375
      Left            =   13680
      TabIndex        =   9
      Top             =   1920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin CrystalActiveXReportViewerLib10Ctl.CrystalActiveXReportViewer CRV1 
      Height          =   10095
      Left            =   240
      TabIndex        =   8
      Top             =   1200
      Width           =   12735
      lastProp        =   600
      _cx             =   22463
      _cy             =   17806
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
      EnableLogonPrompts=   -1  'True
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ØÈÇÚÉ "
      Height          =   735
      Left            =   7920
      TabIndex        =   0
      Top             =   9720
      Width           =   2055
   End
   Begin VB.Label table 
      Caption         =   "0120"
      Height          =   375
      Left            =   5280
      TabIndex        =   7
      Top             =   10080
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label caller 
      Caption         =   "1"
      Height          =   615
      Left            =   3240
      TabIndex        =   6
      Top             =   6840
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label criteria_user_id 
      Caption         =   "criteria_user_id"
      Height          =   495
      Left            =   1080
      TabIndex        =   5
      Top             =   9000
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Label shift_no 
      Caption         =   "Label1"
      Height          =   255
      Left            =   6120
      TabIndex        =   4
      Top             =   8280
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label user_name 
      Caption         =   "Label1"
      Height          =   495
      Left            =   960
      TabIndex        =   3
      Top             =   6840
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label user_id 
      Caption         =   "Label1"
      Height          =   735
      Left            =   7680
      TabIndex        =   2
      Top             =   2520
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label bill_id 
      Caption         =   "197"
      Height          =   495
      Left            =   5640
      TabIndex        =   1
      Top             =   6120
      Visible         =   0   'False
      Width           =   1455
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
   Dim rep As CRAXDRT.Report
    Dim cry As New CRAXDRT.Application
    Dim found As Boolean

'Option Explicit

    Dim sql As String
    Dim strConnect As String

Dim m_Connection As ADODB.Connection
Dim adoRS As ADODB.Recordset

Private Sub Command10_Click()
  Set rep = cry.OpenReport(IMAGE_PATH_FRM.IMAGE_PATH & "\reports\report14.rpt")
      sql = "SELECT  * from  printed_card WHERE     (operation_date >= CONVERT(DATETIME, '" & REPORTSFRM.from_date.Caption & " 00:00:00', 102)) AND (operation_date <= CONVERT(DATETIME, '" & REPORTSFRM.to_date.Caption & " 00:00:00', 102) and OPR_TYPE='ÚÇÏí' ) ORDER BY BILL_NO "
      rep.SQLQueryString = sql
      rep.DiscardSavedData
      CRV1.ReportSource = rep
      CRV1.RefreshEx True
      CRV1.ViewReport
      CRV1.Zoom 100
End Sub

Function search(X As String)
found = False
For i = 0 To List1.ListCount
If List1.List(i) = X Then
found = True
End If

Next i


End Function

Private Sub Command11_Click()
search (DataCombo1.Text)
If found = False Then
List1.AddItem DataCombo1.Text
Else
MsgBox "Êã ÇÖÇÝÉ åÐÇ ÇáäæÚ ãä ÞÈá", vbInformation
End If
End Sub

Private Sub Command12_Click()
   Set rep = cry.OpenReport(IMAGE_PATH_FRM.IMAGE_PATH & "\reports\report9.rpt")
      sql = "SELECT  * from  OPERATIONS WHERE     (operation_date >= CONVERT(DATETIME, '" & REPORTSFRM.from_date.Caption & " 00:00:00', 102)) AND (operation_date <= CONVERT(DATETIME, '" & REPORTSFRM.to_date.Caption & " 00:00:00', 102) and (operation_type='äÔÇØ ÌÏíÏ'  or operation_type='ÊÌÏíÏ äÔÇØ')and activity_name='" & Combo1.Text & "' ) ORDER BY BILL_NO"
      rep.SQLQueryString = sql
      rep.DiscardSavedData
      CRV1.ReportSource = rep
      CRV1.RefreshEx True
      CRV1.ViewReport
      CRV1.Zoom 100
End Sub

Private Sub Command13_Click()
     Set rep = cry.OpenReport(IMAGE_PATH_FRM.IMAGE_PATH & "\reports\report15.rpt")
      sql = "SELECT  * from  OPERATIONS WHERE     (operation_date >= CONVERT(DATETIME, '" & REPORTSFRM.from_date.Caption & " 00:00:00', 102)) AND (operation_date <= CONVERT(DATETIME, '" & REPORTSFRM.to_date.Caption & " 00:00:00', 102) AND OPERATION_TYPE='ÚÖæíÉ ÌÏíÏÉ') ORDER BY BILL_NO"
      rep.SQLQueryString = sql
      rep.DiscardSavedData
      CRV1.ReportSource = rep
      CRV1.RefreshEx True
      CRV1.ViewReport
      CRV1.Zoom 100
End Sub

Private Sub Command14_Click()
     Set rep = cry.OpenReport(IMAGE_PATH_FRM.IMAGE_PATH & "\reports\report15.rpt")
      sql = "SELECT  * from  OPERATIONS WHERE     (operation_date >= CONVERT(DATETIME, '" & REPORTSFRM.from_date.Caption & " 00:00:00', 102)) AND (operation_date <= CONVERT(DATETIME, '" & REPORTSFRM.to_date.Caption & " 00:00:00', 102) AND OPERATION_TYPE='ÊÌÏíÏ ÚÖæíÉ') ORDER BY BILL_NO"
      rep.SQLQueryString = sql
      rep.DiscardSavedData
      CRV1.ReportSource = rep
      CRV1.RefreshEx True
      CRV1.ViewReport
      CRV1.Zoom 100
End Sub

Private Sub Command15_Click()
     Set rep = cry.OpenReport(IMAGE_PATH_FRM.IMAGE_PATH & "\reports\report15.rpt")
      sql = "SELECT  * from  OPERATIONS WHERE     (operation_date >= CONVERT(DATETIME, '" & REPORTSFRM.from_date.Caption & " 00:00:00', 102)) AND (operation_date <= CONVERT(DATETIME, '" & REPORTSFRM.to_date.Caption & " 00:00:00', 102) AND (OPERATION_TYPE='ÚÖæíÉ ÌÏíÏÉ'OR OPERATION_TYPE='ÊÌÏíÏ ÚÖæíÉ')) ORDER BY BILL_NO"
      rep.SQLQueryString = sql
      rep.DiscardSavedData
      CRV1.ReportSource = rep
      CRV1.RefreshEx True
      CRV1.ViewReport
      CRV1.Zoom 100
End Sub

Private Sub Command16_Click()
     Set rep = cry.OpenReport(IMAGE_PATH_FRM.IMAGE_PATH & "\reports\report15.rpt")
      sql = "SELECT  * from  OPERATIONS WHERE     (operation_date >= CONVERT(DATETIME, '" & REPORTSFRM.from_date.Caption & " 00:00:00', 102)) AND (operation_date <= CONVERT(DATETIME, '" & REPORTSFRM.to_date.Caption & " 00:00:00', 102)) ORDER BY BILL_NO"
      rep.SQLQueryString = sql
      rep.DiscardSavedData
      CRV1.ReportSource = rep
      CRV1.RefreshEx True
      CRV1.ViewReport
      CRV1.Zoom 100
End Sub

Private Sub Command18_Click()
If List1.ListIndex = -1 Then Exit Sub
List1.RemoveItem (List1.ListIndex)
End Sub

Private Sub Command19_Click()
If List1.ListCount = 0 Then
MsgBox "áÇíæÌÏ ÇäæÇÚ ÚÖæíÉ ãÍÏÏÉ áØÈÇÚÊåÇ"
Exit Sub
End If

types = "'" + List1.List(0) + "'"
For i = 1 To List1.ListCount - 1
types = types + " or MEMBER_TYPE ='" + List1.List(i) + "'"
Next i

  sql = "SELECT  * from  gam3ea WHERE  MEMBER_TYPE = " & types
'  MsgBox sql
     Set rep = cry.OpenReport(IMAGE_PATH_FRM.IMAGE_PATH & "\reports\report7.rpt")
    
      rep.SQLQueryString = sql
      rep.DiscardSavedData
      CRV1.ReportSource = rep
      CRV1.RefreshEx True
      CRV1.ViewReport
      CRV1.Zoom 100
End Sub

Private Sub Command2_Click()
     Set rep = cry.OpenReport(IMAGE_PATH_FRM.IMAGE_PATH & "\reports\report10.rpt")
      sql = "SELECT  * from  report10 WHERE sex='ÐßÑ' and     (operation_date >= CONVERT(DATETIME, '" & REPORTSFRM.from_date.Caption & " 00:00:00', 102)) AND (operation_date <= CONVERT(DATETIME, '" & REPORTSFRM.to_date.Caption & " 00:00:00', 102))"
      rep.SQLQueryString = sql
      rep.DiscardSavedData
      CRV1.ReportSource = rep
      CRV1.RefreshEx True
      CRV1.ViewReport
      CRV1.Zoom 100
End Sub

Private Sub Command20_Click()
     Set rep = cry.OpenReport(IMAGE_PATH_FRM.IMAGE_PATH & "\reports\report15.rpt")
      sql = "SELECT  * from  OPERATIONS WHERE     (operation_date >= CONVERT(DATETIME, '" & REPORTSFRM.from_date.Caption & " 00:00:00', 102)) AND (operation_date <= CONVERT(DATETIME, '" & REPORTSFRM.to_date.Caption & " 00:00:00', 102) AND (member_name='ãÕÑæÝÇÊ')) ORDER BY BILL_NO"
      rep.SQLQueryString = sql
      rep.DiscardSavedData
      CRV1.ReportSource = rep
      CRV1.RefreshEx True
      CRV1.ViewReport
      CRV1.Zoom 100

End Sub

Private Sub Command21_Click()
Set rep = cry.OpenReport(IMAGE_PATH_FRM.IMAGE_PATH & "\reports\cardid _on_column.rpt")
 sql = "Select * from ready_to_print  where SELECTED=1 AND PRINTED=0"
  rep.SQLQueryString = sql
    rep.DiscardSavedData
   
    CRV1.ReportSource = rep
    CRV1.RefreshEx True
    CRV1.ViewReport
   CRV1.Zoom 100

End Sub

Private Sub Command22_Click()

Set rep = cry.OpenReport(IMAGE_PATH_FRM.IMAGE_PATH & "\reports\CARDID.rpt")
 sql = "Select * from ready_to_print  where SELECTED=1 AND PRINTED=0"
  rep.SQLQueryString = sql
    rep.DiscardSavedData
   
    CRV1.ReportSource = rep
    CRV1.RefreshEx True
    CRV1.ViewReport
   CRV1.Zoom 100
End Sub

Private Sub Command3_Click()
    Set rep = cry.OpenReport(IMAGE_PATH_FRM.IMAGE_PATH & "\reports\report10.rpt")
      sql = "SELECT  * from  report10 WHERE sex='ÇäËì' and     (operation_date >= CONVERT(DATETIME, '" & REPORTSFRM.from_date.Caption & " 00:00:00', 102)) AND (operation_date <= CONVERT(DATETIME, '" & REPORTSFRM.to_date.Caption & " 00:00:00', 102))"
      rep.SQLQueryString = sql
      rep.DiscardSavedData
      CRV1.ReportSource = rep
      CRV1.RefreshEx True
      CRV1.ViewReport
      CRV1.Zoom 100
End Sub

Private Sub Command4_Click()
   Set rep = cry.OpenReport(IMAGE_PATH_FRM.IMAGE_PATH & "\reports\report10.rpt")
      sql = "SELECT  * from  report10 WHERE     (operation_date >= CONVERT(DATETIME, '" & REPORTSFRM.from_date.Caption & " 00:00:00', 102)) AND (operation_date <= CONVERT(DATETIME, '" & REPORTSFRM.to_date.Caption & " 00:00:00', 102))"
      rep.SQLQueryString = sql
      rep.DiscardSavedData
      CRV1.ReportSource = rep
      CRV1.RefreshEx True
      CRV1.ViewReport
      CRV1.Zoom 100
End Sub

Private Sub Command5_Click()
  Set rep = cry.OpenReport(IMAGE_PATH_FRM.IMAGE_PATH & "\reports\report14.rpt")
      sql = "SELECT  * from  printed_card WHERE     (operation_date >= CONVERT(DATETIME, '" & REPORTSFRM.from_date.Caption & " 00:00:00', 102)) AND (operation_date <= CONVERT(DATETIME, '" & REPORTSFRM.to_date.Caption & " 00:00:00', 102) and RECIVED=0 ) ORDER BY BILL_NO "
      rep.SQLQueryString = sql
      rep.DiscardSavedData
      CRV1.ReportSource = rep
      CRV1.RefreshEx True
      CRV1.ViewReport
      CRV1.Zoom 100
End Sub

Private Sub Command6_Click()
  Set rep = cry.OpenReport(IMAGE_PATH_FRM.IMAGE_PATH & "\reports\report14.rpt")
      sql = "SELECT  * from  printed_card WHERE     (operation_date >= CONVERT(DATETIME, '" & REPORTSFRM.from_date.Caption & " 00:00:00', 102)) AND (operation_date <= CONVERT(DATETIME, '" & REPORTSFRM.to_date.Caption & " 00:00:00', 102)) ORDER BY BILL_NO"
      rep.SQLQueryString = sql
      rep.DiscardSavedData
      CRV1.ReportSource = rep
      CRV1.RefreshEx True
      CRV1.ViewReport
      CRV1.Zoom 100
End Sub

Private Sub Command7_Click()
  Set rep = cry.OpenReport(IMAGE_PATH_FRM.IMAGE_PATH & "\reports\report14.rpt")
      sql = "SELECT  * from  printed_card WHERE     (operation_date >= CONVERT(DATETIME, '" & REPORTSFRM.from_date.Caption & " 00:00:00', 102)) AND (operation_date <= CONVERT(DATETIME, '" & REPORTSFRM.to_date.Caption & " 00:00:00', 102) and RECIVED=1) ORDER BY BILL_NO   "
      rep.SQLQueryString = sql
      rep.DiscardSavedData
      CRV1.ReportSource = rep
      CRV1.RefreshEx True
      CRV1.ViewReport
      CRV1.Zoom 100
End Sub

Private Sub Command8_Click()
  Set rep = cry.OpenReport(IMAGE_PATH_FRM.IMAGE_PATH & "\reports\report14.rpt")
      sql = "SELECT  * from  printed_card WHERE     (operation_date >= CONVERT(DATETIME, '" & REPORTSFRM.from_date.Caption & " 00:00:00', 102)) AND (operation_date <= CONVERT(DATETIME, '" & REPORTSFRM.to_date.Caption & " 00:00:00', 102) and OPR_TYPE='ÈÏá ÝÇÆÏ') ORDER BY BILL_NO  "
      rep.SQLQueryString = sql
      rep.DiscardSavedData
      CRV1.ReportSource = rep
      CRV1.RefreshEx True
      CRV1.ViewReport
      CRV1.Zoom 100
End Sub

Private Sub Command9_Click()
  Set rep = cry.OpenReport(IMAGE_PATH_FRM.IMAGE_PATH & "\reports\report14.rpt")
      sql = "SELECT  * from  printed_card WHERE     (operation_date >= CONVERT(DATETIME, '" & REPORTSFRM.from_date.Caption & " 00:00:00', 102)) AND (operation_date <= CONVERT(DATETIME, '" & REPORTSFRM.to_date.Caption & " 00:00:00', 102)  and opr_type=  'ÇÚÇÏÉ ØÈÇÚÉ ' ) ORDER BY BILL_NO "
      rep.SQLQueryString = sql
      rep.DiscardSavedData
      CRV1.ReportSource = rep
      CRV1.RefreshEx True
      CRV1.ViewReport
      CRV1.Zoom 100
End Sub

Private Sub Form_Activate()


found = False

    Select Case case_id
 

   Case 1
   
   
   
        Set rep = cry.OpenReport(IMAGE_PATH_FRM.IMAGE_PATH + "\reports\greaterthan18.rpt")
      rep.DiscardSavedData
      CRV1.ReportSource = rep
      CRV1.RefreshEx True
      CRV1.ViewReport
      CRV1.Zoom 100
   

  Case 2
  

Set rep = cry.OpenReport(IMAGE_PATH_FRM.IMAGE_PATH & "\reports\CARDID.rpt")
 sql = "Select * from ready_to_print  where SELECTED=1 AND PRINTED=0"
  rep.SQLQueryString = sql
    rep.DiscardSavedData
   
    CRV1.ReportSource = rep
    CRV1.RefreshEx True
    CRV1.ViewReport
   CRV1.Zoom 100
  Frame6.Visible = True
  Case 4
 

Set rep = cry.OpenReport(IMAGE_PATH_FRM.IMAGE_PATH & "\reports\REPORT4.rpt")

 sql = "SELECT * from report4 WHERE   (OPERATION_DATE >= CONVERT(DATETIME, '" & REPORTSFRM.from_date.Caption & " 00:00:00', 102) AND OPERATION_DATE <= CONVERT(DATETIME,'" & REPORTSFRM.to_date.Caption & " 00:00:00', 102))"
       
      rep.SQLQueryString = sql
      rep.DiscardSavedData
      CRV1.ReportSource = rep
         
    ' rep.FormulaFields(1).Text = Str(emp_reports.from_date.Caption)
     ' rep.FormulaFields(2).Text = Str(emp_reports.to_date.Caption)
      
   '   rep.FormulaFields(1).Text = 1 ' Mid(REPORTSFRM.from_date.Caption, 1, 2)
   '   rep.FormulaFields(2).Text = 2 ' Mid(REPORTSFRM.to_date.Caption, 1, 2)
   '   rep.FormulaFields(3).Text = 3 'Mid(REPORTSFRM.from_date.Caption, 3, 2)
   '   rep.FormulaFields(4).Text = 4 'Mid(REPORTSFRM.to_date.Caption, 3, 2)
   '   rep.FormulaFields(5).Text = 5 'Mid(REPORTSFRM.from_date.Caption, 7, 4)
   '   rep.FormulaFields(6).Text = 6 'Mid(REPORTSFRM.to_date.Caption, 7, 4)
'rep.ParameterFields.GetItemByName("P1", "REPORT4").Value = "AAA"
      CRV1.RefreshEx True
      CRV1.ViewReport
      CRV1.Zoom 100
   
   
  Case 5
  
Set rep = cry.OpenReport(IMAGE_PATH_FRM.IMAGE_PATH & "\reports\REPORT5.rpt")

 sql = "SELECT * from report5 WHERE   (OPERATION_DATE >= CONVERT(DATETIME, '" & REPORTSFRM.from_date.Caption & " 00:00:00', 102) AND OPERATION_DATE <= CONVERT(DATETIME,'" & REPORTSFRM.to_date.Caption & " 00:00:00', 102))"
       
      rep.SQLQueryString = sql
      rep.DiscardSavedData
      CRV1.ReportSource = rep
         
    ' rep.FormulaFields(1).Text = Str(emp_reports.from_date.Caption)
     ' rep.FormulaFields(2).Text = Str(emp_reports.to_date.Caption)
      
   '   rep.FormulaFields(1).Text = 1 ' Mid(REPORTSFRM.from_date.Caption, 1, 2)
   '   rep.FormulaFields(2).Text = 2 ' Mid(REPORTSFRM.to_date.Caption, 1, 2)
   '   rep.FormulaFields(3).Text = 3 'Mid(REPORTSFRM.from_date.Caption, 3, 2)
   '   rep.FormulaFields(4).Text = 4 'Mid(REPORTSFRM.to_date.Caption, 3, 2)
   '   rep.FormulaFields(5).Text = 5 'Mid(REPORTSFRM.from_date.Caption, 7, 4)
   '   rep.FormulaFields(6).Text = 6 'Mid(REPORTSFRM.to_date.Caption, 7, 4)
'rep.ParameterFields.GetItemByName("P1", "REPORT4").Value = "AAA"
      CRV1.RefreshEx True
      CRV1.ViewReport
      CRV1.Zoom 100
      
  Case 6
  Set rep = cry.OpenReport(IMAGE_PATH_FRM.IMAGE_PATH & "\reports\REPORT6.rpt")
      sql = "SELECT  * from  report6 WHERE     (OPERATION_DATE >= CONVERT(DATETIME, '" & REPORTSFRM.from_date.Caption & " 00:00:00', 102)) AND (OPERATION_DATE <= CONVERT(DATETIME, '" & REPORTSFRM.to_date.Caption & " 00:00:00', 102))"
      rep.SQLQueryString = sql
      rep.DiscardSavedData
      CRV1.ReportSource = rep
      CRV1.RefreshEx True
      CRV1.ViewReport
      CRV1.Zoom 100
  
    Case 7
  Frame4.Visible = True
     Set rep = cry.OpenReport(IMAGE_PATH_FRM.IMAGE_PATH & "\reports\report7.rpt")
     sql = "SELECT  * from  gam3ea" ' WHERE last_update_year='" & Label5.Caption & "'"
      rep.SQLQueryString = sql
     rep.DiscardSavedData
      CRV1.ReportSource = rep
      CRV1.RefreshEx True
      CRV1.ViewReport
      CRV1.Zoom 100
   
   Frame5.Visible = False
   
      
      Case 8
     Set rep = cry.OpenReport(IMAGE_PATH_FRM.IMAGE_PATH & "\reports\report8.rpt")
      sql = "SELECT  * from  OPERATIONS WHERE     (operation_date >= CONVERT(DATETIME, '" & REPORTSFRM.from_date.Caption & " 00:00:00', 102)) AND (operation_date <= CONVERT(DATETIME, '" & REPORTSFRM.to_date.Caption & " 00:00:00', 102) AND MEMBER_ID='" & REPORTSFRM.Text1.Text & "' ) ORDER BY BILL_NO"
      rep.SQLQueryString = sql
      rep.DiscardSavedData
      CRV1.ReportSource = rep
      CRV1.RefreshEx True
      CRV1.ViewReport
      CRV1.Zoom 100
   Case 9
     Set rep = cry.OpenReport(IMAGE_PATH_FRM.IMAGE_PATH & "\reports\report9.rpt")
      sql = "SELECT  * from  OPERATIONS WHERE     (operation_date >= CONVERT(DATETIME, '" & REPORTSFRM.from_date.Caption & " 00:00:00', 102)) AND (operation_date <= CONVERT(DATETIME, '" & REPORTSFRM.to_date.Caption & " 00:00:00', 102) and (operation_type='äÔÇØ ÌÏíÏ'  or operation_type='ÊÌÏíÏ äÔÇØ') ) ORDER BY BILL_NO"
      rep.SQLQueryString = sql
      rep.DiscardSavedData
      CRV1.ReportSource = rep
      CRV1.RefreshEx True
      CRV1.ViewReport
      CRV1.Zoom 100
      Frame3.Visible = True
   
   For i = 1 To Adodc1.Recordset.RecordCount
   Combo1.AddItem Adodc1.Recordset.Fields!Activities_NAME
   Adodc1.Recordset.MoveNext
   Next i
   
   
   Case 10
     Set rep = cry.OpenReport(IMAGE_PATH_FRM.IMAGE_PATH & "\reports\report10.rpt")
      sql = "SELECT  * from  report10 WHERE     (operation_date >= CONVERT(DATETIME, '" & REPORTSFRM.from_date.Caption & " 00:00:00', 102)) AND (operation_date <= CONVERT(DATETIME, '" & REPORTSFRM.to_date.Caption & " 00:00:00', 102))"
      rep.SQLQueryString = sql
      rep.DiscardSavedData
      CRV1.ReportSource = rep
      CRV1.RefreshEx True
      CRV1.ViewReport
      CRV1.Zoom 100
   Command2.Visible = True
   Command3.Visible = True
   Command4.Visible = True
       Case 14
       
     Set rep = cry.OpenReport(IMAGE_PATH_FRM.IMAGE_PATH & "\reports\report14.rpt")
      sql = "SELECT  * from  printed_card WHERE     (operation_date >= CONVERT(DATETIME, '" & REPORTSFRM.from_date.Caption & " 00:00:00', 102)) AND (operation_date <= CONVERT(DATETIME, '" & REPORTSFRM.to_date.Caption & " 00:00:00', 102)) ORDER BY BILL_NO,member_id"
      rep.SQLQueryString = sql
      rep.DiscardSavedData
      CRV1.ReportSource = rep
      CRV1.RefreshEx True
      CRV1.ViewReport
      CRV1.Zoom 100
      Frame1.Visible = True
   
      Case 15
     Set rep = cry.OpenReport(IMAGE_PATH_FRM.IMAGE_PATH & "\reports\report15.rpt")
      sql = "SELECT  * from  OPERATIONS WHERE     (operation_date >= CONVERT(DATETIME, '" & REPORTSFRM.from_date.Caption & " 00:00:00', 102)) AND (operation_date <= CONVERT(DATETIME, '" & REPORTSFRM.to_date.Caption & " 00:00:00', 102)) ORDER BY BILL_NO"
      rep.SQLQueryString = sql
      rep.DiscardSavedData
      CRV1.ReportSource = rep
      CRV1.RefreshEx True
      CRV1.ViewReport
      CRV1.Zoom 100
      
   Frame2.Visible = True
   
   End Select
  
'Set rep = cry.OpenReport(IMAGE_PATH_FRM.IMAGE_PATH  & "\reports\delivery_bill.rpt")
' sql = "Select * from [deliveryOrder]  where bill_id=" & Me.bill_id.Caption
'    adoRS.Open sql, m_Connection, adOpenDynamic, adLockBatchOptimistic
'    rep.Database.SetDataSource adoRS
   
   ' Report.
 '   report.re
'CRV1.ReportSource = rep
  '  CRV1.ViewReport

    'CRV1.ViewReport
   ' CRV1.Zoom 100



 

    'cRV1.ViewReport
 ' DoEvents
 '   CRV1.Refresh
 ' ' Set crv1.Zoom = 150
 '   Screen.MousePointer = vbDefault
'   crv1.Refresh
End Sub

' *************************************************************
' Load the Report in the viewer
'
Private Sub Form_Load()
'For i = 1 To Adodc2.Recordset.RecordCount
'List1.AddItem Adodc2.Recordset.Fields!MEMBER_NAME
'Adodc2.Recordset.MoveNext
'Next i
   
    ' Create and bind the ADO Recordset object
    Set m_Connection = New ADODB.Connection
    Set adoRS = New ADODB.Recordset

    ' Open the connection
    strConnect = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=MAY"
     m_Connection.Open strConnect
    
End Sub

' *************************************************************
Private Sub cmdAbout_Click()
   ' frmAbout.Show vbModal
End Sub

' *************************************************************
Private Sub cmdExit_Click()
   ' Unload Me
End Sub

Private Sub lblTotal_Click()
End Sub

Private Sub Form_Unload(Cancel As Integer)
'bill_id.Caption = 0
End Sub


Private Sub Text2_Change()

End Sub

