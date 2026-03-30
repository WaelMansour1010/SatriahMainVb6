VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmAveProdPriceMatrialReports 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   6195
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9945
   Icon            =   "FrmAveProdPriceMatrialReports.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   9945
   ShowInTaskbar   =   0   'False
   Begin C1SizerLibCtl.C1Elastic C1Elastic1 
      Height          =   495
      Left            =   5880
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   6480
      Width           =   1095
      _cx             =   1931
      _cy             =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   4
      MousePointer    =   0
      Version         =   801
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   0
      AutoSizeChildren=   0
      BorderWidth     =   6
      ChildSpacing    =   4
      Splitter        =   0   'False
      FloodDirection  =   0
      FloodPercent    =   0
      CaptionPos      =   1
      WordWrap        =   -1  'True
      MaxChildSize    =   0
      MinChildSize    =   0
      TagWidth        =   0
      TagPosition     =   0
      Style           =   0
      TagSplit        =   2
      PicturePos      =   4
      CaptionStyle    =   0
      ResizeFonts     =   0   'False
      GridRows        =   0
      GridCols        =   0
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   ""
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
   End
   Begin VB.CommandButton btnClear 
      Caption         =   "ãÓÍ"
      Height          =   495
      Left            =   2760
      TabIndex        =   7
      Top             =   5640
      Width           =   1335
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Height          =   4725
      Index           =   1
      Left            =   0
      TabIndex        =   3
      Top             =   600
      Width           =   9915
      Begin VB.TextBox TxtAttachedItemCode 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   3975
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÍÏÏ ÝÊÑÉ ÇáÔÑÇÁ"
         Height          =   735
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   1200
         Width           =   4455
         Begin MSComCtl2.DTPicker DtpDateFrom 
            Height          =   330
            Left            =   2280
            TabIndex        =   12
            Top             =   270
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   68747267
            CurrentDate     =   41640
         End
         Begin MSComCtl2.DTPicker DtpDateTo 
            Height          =   330
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   68747267
            CurrentDate     =   41640
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "Åáì"
            Height          =   195
            Index           =   3
            Left            =   1710
            RightToLeft     =   -1  'True
            TabIndex        =   15
            Top             =   240
            Width           =   480
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "ãä"
            Height          =   195
            Index           =   4
            Left            =   3690
            RightToLeft     =   -1  'True
            TabIndex        =   14
            Top             =   240
            Width           =   540
         End
      End
      Begin VB.Frame Frame3 
         Height          =   4575
         Left            =   5520
         TabIndex        =   5
         Top             =   120
         Width           =   4335
         Begin VB.Image Image1 
            Height          =   3675
            Left            =   120
            Picture         =   "FrmAveProdPriceMatrialReports.frx":038A
            Stretch         =   -1  'True
            Top             =   120
            Width           =   4395
         End
         Begin VB.Label lblCompanyname 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "ÇáÓÇÊÑíÉ"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   27.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   1095
            Left            =   480
            TabIndex        =   6
            Top             =   3840
            Width           =   2895
         End
      End
      Begin MSDataListLib.DataCombo DcboItemID1 
         Height          =   315
         Left            =   120
         TabIndex        =   19
         Top             =   720
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "ßæÏ ÇáÕäÝ"
         Height          =   315
         Index           =   0
         Left            =   4320
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   240
         Width           =   1125
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "ÇÓã ÇáÕäÝ"
         Height          =   315
         Index           =   2
         Left            =   4320
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   720
         Width           =   1125
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Caption         =   "ÚäÏ ÚÏã ÅÎÊíÇÑ ÇÍÏ ÇáÍÞæá ÓæÝ íßæä ÇáÊÞÑíÑ ÅÌãÇáí"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   2460
         Index           =   25
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   2040
         Width           =   5295
      End
      Begin VB.Shape Shape1 
         BorderWidth     =   2
         Height          =   2535
         Left            =   120
         Top             =   2040
         Width           =   5295
      End
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   495
      Index           =   0
      Left            =   1440
      TabIndex        =   0
      Top             =   5640
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   873
      ButtonPositionImage=   1
      Caption         =   "ÚÑÖ ÇáÊÞÑíÑ"
      BackColor       =   14871017
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackStyle       =   0
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   -2147483637
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      DisabledImageExtraction=   0
      ColorToggledHoverText=   16711680
      LowerToggledContent=   0   'False
      ColorTextShadow =   -2147483637
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   495
      Index           =   2
      Left            =   240
      TabIndex        =   1
      Top             =   5640
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   873
      ButtonPositionImage=   1
      Caption         =   "ÎÑæÌ"
      BackColor       =   14871017
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackStyle       =   0
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   -2147483637
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      LowerToggledContent=   0   'False
      ColorTextShadow =   -2147483637
   End
   Begin MSDataListLib.DataCombo DataCombo2 
      Height          =   315
      Left            =   240
      TabIndex        =   8
      Top             =   2040
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00E2E9E9&
      Caption         =   "ØÈÞÇ áãÓÊÃÌÑ ãÍÏÏ"
      Height          =   195
      Index           =   5
      Left            =   5400
      TabIndex        =   9
      Top             =   2040
      Width           =   1290
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "ÊÞÇÑíÑ ãÊæÓØ ÇáÊßÇáíÝ ááÇÕäÇÝ ÇáãäÊÌÉ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   45
      TabIndex        =   4
      Top             =   0
      Width           =   9945
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      ForeColor       =   &H00000080&
      Height          =   285
      Index           =   1
      Left            =   60
      TabIndex        =   2
      Top             =   3060
      Width           =   1785
   End
End
Attribute VB_Name = "FrmAveProdPriceMatrialReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsSavRec As ADODB.Recordset
Dim BKGrndPic As ClsBackGroundPic
Dim RecID As String
Dim II As Long
Dim cSearch  As clsDCboSearch
Dim rs As ADODB.Recordset
Dim DCboSearch As clsDCboSearch

Private Sub ChangeLang()
    Dim XPic As IPictureDisp
      Label5.Caption = "Report of  Average price for the purchase of raw materials"
    lbl(25).Caption = Label5.Caption
    lbl(0).Caption = "Item Code"
    Me.lbl(2).Caption = "Item Name."
   lblCompanyname.Caption = "El-Sattaryh"
   
    Frame1.Caption = "Period"
    lbl(3).Caption = "To"
    lbl(4).Caption = "From"
btnClear.Caption = "Clear"
Cmd(0).Caption = "Show Report"
Cmd(2).Caption = "Exit"

   

End Sub
Private Sub btnClear_Click()
clear_all Me

DtpDateFrom.value = ""
DtpDateTo.value = ""
End Sub




Private Sub Cmd_Click(Index As Integer)

    Select Case Index

        Case 0

GetData
          
        Case 1
            clear_all Me

            If SystemOptions.UserInterface = ArabicInterface Then
               ' Me.lbl(0).Caption = "äÊíÌÉ ÇáÈÍË"
            Else
               ' Me.lbl(0).Caption = "Search Results"
            End If

        Case 2
            Unload Me
            Case 3
'print_report
    End Select

End Sub








Private Sub DcboItemID1_KeyUp(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyF5 Then
       LoadCombosData
    End If
End Sub

Private Sub Form_Activate()
   PutFormOnTop Me.hWnd
End Sub


Private Sub Form_Load()
   
 'On Error GoTo ErrTrap
LoadCombosData
DtpDateFrom.value = ""
DtpDateTo.value = ""
    Resize_Form Me
    If SystemOptions.UserInterface = EnglishInterface Then
    SetInterface Me
    ChangeLang
    End If

End Sub



Private Sub LoadCombosData()
  Dim Dcombos As ClsDataCombos
   
    Set Dcombos = New ClsDataCombos
Dcombos.GetItemsNames Me.DcboItemID1
 
End Sub

Private Sub Form_Unload(Cancel As Integer)

    FormPostion Me, SavePostion
    Set DCboSearch = Nothing
End Sub

Public Sub GetData()
    Dim StrSQL As String
      Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
  

StrSQL = "SELECT     dbo.TblItems.ItemCode, dbo.TblItems.ItemName, dbo.TblItems.ItemNamee, dbo.Transactions.Transaction_ID, dbo.Transactions.Transaction_Serial, "
StrSQL = StrSQL & "                      dbo.Transactions.Transaction_Date, dbo.Transactions.Transaction_Type, dbo.Transaction_Details.ShowQty,"
StrSQL = StrSQL & "                       dbo.GetTotalProudctWork(dbo.Transactions.Transaction_Serial) AS Total"
StrSQL = StrSQL & "  FROM         dbo.TblItems RIGHT OUTER JOIN"
StrSQL = StrSQL & "                       dbo.Transactions LEFT OUTER JOIN"
StrSQL = StrSQL & "                       dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID ON"
StrSQL = StrSQL & "                       dbo.TblItems.ItemID = dbo.Transaction_Details.Item_ID"
StrSQL = StrSQL & "  Where (dbo.Transactions.Transaction_Type = 26) And (dbo.GetTotalProudctWork(dbo.Transactions.Transaction_serial) >= 0)"



  StrWhere = ""
   '///////////////////
   
          If (Me.DcboItemID1.text <> "") And (val(DcboItemID1.BoundText) <> 0) Then
             StrWhere = StrWhere & " AND dbo.Transaction_Details.Item_ID =" & Me.DcboItemID1.BoundText & ""
          End If

     If Not IsNull(Me.DtpDateFrom.value) Then
        StrWhere = StrWhere & " AND dbo.Transactions.Transaction_Date >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
      End If
        If Not IsNull(Me.DtpDateTo.value) Then
         StrWhere = StrWhere & " AND dbo.Transactions.Transaction_Date <=" & SQLDate(Me.DtpDateTo.value, True) & ""
                     
      End If
    StrSQL = StrSQL & StrWhere
  StrSQL = StrSQL & "  ORDER BY dbo.Transactions.Transaction_ID"
 

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
           ' Me.lbl(10).Caption = "äÊíÌÉ ÇáÈÍË=ÕÝÑ"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
          '  Me.lbl(10).Caption = "Search Results=0"
        End If

        Msg = "áÇÊæÌÏ ÈíÇäÇÊ ááÚÑÖ ÊæÇÝÞ ÔÑæØ ÇáÊÞÑíÑ"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    Else
   '  Me.DTPicker1.value = Format(rs("DateAcutExite").value, "yyyy/M/d")
 rs.MoveFirst
' MsgBox rs("Granty").value
 print_report StrSQL
'
            If SystemOptions.UserInterface = ArabicInterface Then
             '   Me.lbl(10).Caption = "äÊíÌÉ ÇáÈÍË=" & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
               ' Me.lbl(10).Caption = "Search Results=" & rs.RecordCount
            End If

    End If

End Sub
Function print_report(Optional NoteSerial As String)
     
    Set rs = New ADODB.Recordset
    rs.Open NoteSerial, Cn, adOpenStatic, adLockReadOnly, adCmdText
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
   

        If SystemOptions.UserInterface = ArabicInterface Then
        
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RePAveProdPriceMatrialsReports.rpt"
            Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RePAveProdPriceMatrialsReportsE.rpt"
            
       End If
     
  


    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open NoteSerial, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        'GetMsgs 138, vbExclamation
        If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "áÇÊæÌÏ ÈíÇäÇÊ ááÚÑÖ"
        Else
        Msg = "Not Found Data to Show"
        End If
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        RsData.Close
        Set RsData = Nothing
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)
    xReport.Database.SetDataSource RsData

    Dim cCompanyInfo As New ClsCompanyInfo

    If SystemOptions.UserInterface = ArabicInterface Then
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
        ' xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
        StrReportTitle = "" '& StrAccountName

    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
       
        StrReportTitle = ""
  
    End If

   
  If DtpDateFrom.value <> "" And DtpDateTo.value <> "" Then
   xReport.ParameterFields(8).AddCurrentValue DtpDateFrom.value

    xReport.ParameterFields(10).AddCurrentValue DtpDateTo.value
  '  xReport.ParameterFields(11).AddCurrentValue DtpDateToH.value
    End If
    xReport.ParameterFields(12).AddCurrentValue val(SystemOptions.IndirectCostPercentage)

  Dim total As String
  Dim totl As Double


    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault

 
End Function
Private Sub DcboItemID1_Change()
    Me.TxtAttachedItemCode.text = GetItemCode(val(Me.DcboItemID1.BoundText))
End Sub

Private Sub DcboItemID1_Click(Area As Integer)
    DcboItemID1_Change
End Sub

Private Sub TxtAttachedItemCode_KeyDown(KeyCode As Integer, _
                                        Shift As Integer)

    If KeyCode = vbKeyReturn Then
        If TxtAttachedItemCode.text = "" Then
            Me.DcboItemID1.BoundText = ""
        Else
            Me.DcboItemID1.BoundText = GetItemID(Trim$(Me.TxtAttachedItemCode.text))
        End If
    End If

End Sub

