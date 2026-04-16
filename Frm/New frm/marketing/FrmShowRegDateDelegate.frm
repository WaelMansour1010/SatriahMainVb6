VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmShowRegDateDelegate 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÔÇÔÇÉ ãÊÇÈÚÉ ãæÇÚíÏ ÇáãäÇÏíÈ"
   ClientHeight    =   8535
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   17040
   ClipControls    =   0   'False
   Icon            =   "FrmShowRegDateDelegate.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   8535
   ScaleWidth      =   17040
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   0
      TabIndex        =   6
      Top             =   7800
      Width           =   1815
      Begin VB.CommandButton menue 
         Height          =   555
         Index           =   1
         Left            =   3240
         Picture         =   "FrmShowRegDateDelegate.frx":038A
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   840
         Width           =   735
      End
      Begin VB.CommandButton menue 
         BackColor       =   &H8000000D&
         Caption         =   "ÊÍÏíË"
         DownPicture     =   "FrmShowRegDateDelegate.frx":082B
         Height          =   555
         Index           =   16
         Left            =   840
         Picture         =   "FrmShowRegDateDelegate.frx":7B5D
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   120
         Width           =   735
      End
      Begin ImpulseButton.ISButton Cmd 
         Cancel          =   -1  'True
         Height          =   555
         Index           =   2
         Left            =   0
         TabIndex        =   9
         Top             =   120
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   979
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
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   5880
      Top             =   9360
   End
   Begin VB.ComboBox DcbOrderStatus 
      Height          =   315
      Left            =   20400
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   3480
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.ComboBox ComGranty 
      Height          =   315
      Left            =   20520
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   2520
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   20760
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Text            =   "Êã ãæÇÝÞÉ ÇáÚã"
      Top             =   3000
      Visible         =   0   'False
      Width           =   915
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   615
      Index           =   0
      Left            =   3360
      TabIndex        =   0
      Top             =   9000
      Visible         =   0   'False
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   1085
      ButtonPositionImage=   1
      Caption         =   "ÚÑÖ ÇáÊÞÑíÑ"
      BackColor       =   14871017
      Enabled         =   0   'False
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
      Height          =   615
      Index           =   1
      Left            =   1920
      TabIndex        =   1
      Top             =   9120
      Visible         =   0   'False
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   1085
      ButtonPositionImage=   1
      Caption         =   "ãÓÍ"
      BackColor       =   14871017
      Enabled         =   0   'False
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
   Begin VSFlex8UCtl.VSFlexGrid Fg 
      Height          =   7785
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   16995
      _cx             =   29977
      _cy             =   13732
      Appearance      =   1
      BorderStyle     =   0
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483639
      BackColorFixed  =   14871017
      ForeColorFixed  =   -2147483639
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483628
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483624
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   31
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmShowRegDateDelegate.frx":80F7
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   7
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   -1  'True
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   2
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin MSDataListLib.DataCombo DcbTypeVisit1 
      Height          =   315
      Left            =   13560
      TabIndex        =   10
      Top             =   3840
      Visible         =   0   'False
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSComCtl2.DTPicker DataeDaye 
      Height          =   330
      Left            =   15480
      TabIndex        =   11
      Top             =   5520
      Visible         =   0   'False
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   582
      _Version        =   393216
      CheckBox        =   -1  'True
      Format          =   94437379
      CurrentDate     =   38887
   End
   Begin MSComCtl2.DTPicker LastDate 
      Height          =   330
      Left            =   15240
      TabIndex        =   14
      Top             =   5160
      Visible         =   0   'False
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   582
      _Version        =   393216
      CheckBox        =   -1  'True
      CustomFormat    =   "dd/mm/yyyy"
      Format          =   94437379
      CurrentDate     =   38887
   End
   Begin MSComCtl2.DTPicker FirstDate 
      Height          =   330
      Left            =   15120
      TabIndex        =   15
      Top             =   4680
      Visible         =   0   'False
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   582
      _Version        =   393216
      CheckBox        =   -1  'True
      CustomFormat    =   "dd/mm/yyyy"
      Format          =   94437379
      CurrentDate     =   38887
   End
   Begin MSComCtl2.DTPicker ChekDate 
      Height          =   330
      Left            =   14880
      TabIndex        =   16
      Top             =   4200
      Visible         =   0   'False
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   582
      _Version        =   393216
      CheckBox        =   -1  'True
      Format          =   94437379
      CurrentDate     =   38887
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   705
      Index           =   0
      Left            =   1800
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   7800
      Width           =   6735
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
      Caption         =   "    ÔÇÔÉ ãæÇÚíÏ ÇáãäÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÇÏíÜÜÜÜÜÈ áåÐÇ ÇáÇÓÈæÚ  ãä ÊÇÑíÎ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   705
      Index           =   10
      Left            =   8520
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   7800
      Width           =   8535
   End
End
Attribute VB_Name = "FrmShowRegDateDelegate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim DCboSearch As clsDCboSearch


Dim TTP As clstooltip
Dim cSearchDcbo  As clsDCboSearch
Dim TTD As clstooltipdemand
Dim Employee_account As String



Private Sub Cmd_Click(Index As Integer)

    Select Case Index

        Case 0
       

 GetData
            
        Case 1
            clear_all Me
'DtpDateFrom.value = ""
'DtpDateTo.value = ""
'Me.DtStart.value = ""
'Me.DtEnd.value = ""
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
Private Sub Fg_Click()
FrmRegDateDelgate.Retrive (val(Me.Fg.TextMatrix(Me.Fg.Row, Me.Fg.ColIndex("id"))))
End Sub



Private Sub Coloring()
    Dim i As Integer
    Dim IntCounter As Integer
Dim line_no1 As Integer
    With Me.Fg

        For i = .FixedRows To .Rows - 1
        
            If i Mod 2 = 0 Then
                .Cell(flexcpBackColor, i, 1, i, 21) = &HFFFFC0
            Else
                .Cell(flexcpBackColor, i, 1, i, 21) = vbWhite
            End If

        Next i

    End With

    line_no1 = IntCounter

End Sub

Private Sub Form_Activate()
   PutFormOnTop Me.hWnd
End Sub
Sub GetNameofDayes(Optional dat As Date, Optional ByRef sweekday As String)


    Dim daynum As Integer
    Dim str As String
    str = dat
    'Figure out the day
    daynum = DatePart("w", str)
    If SystemOptions.UserInterface = ArabicInterface Then
    Select Case (daynum)
        Case 1
            sweekday = "ÇáÇÍÏ"
        Case 2
            sweekday = "ÇáÇËäíä"
        Case 3
            sweekday = "ÇáËáÇËÇÁ"
        Case 4
            sweekday = "ÇáÇÑÈÚÇÁ"
        Case 5
            sweekday = "ÇáÎãíÓ"
        Case 6
            sweekday = "ÇáÌãÚå"
        Case 7
            sweekday = "ÇáÓÈÊ"
        Case Else
            sweekday = "ÛíÑ ãÚÑæÝ"
    End Select
    Else
    Select Case (daynum)
        Case 1
            sweekday = "Sunday"
        Case 2
            sweekday = "Monday"
        Case 3
            sweekday = "Tuesday"
        Case 4
            sweekday = "Wednesday"
        Case 5
            sweekday = "Thursday"
        Case 6
            sweekday = "Friday"
        Case 7
            sweekday = "Saturday"
        Case Else
            sweekday = "unknown"
    End Select
    End If
  '  Text2.text = sweekday
    End Sub
'Private Function Retrive(id As Integer, Optional ByRef str As String, Optional ByRef x As Integer, Optional ByRef strb As String)
' Dim Rs3 As ADODB.Recordset
'    Set Rs3 = New ADODB.Recordset
'    Dim sql As String
'    Dim Rs4 As ADODB.Recordset
'    Dim Index As Integer
'    Set Rs4 = New ADODB.Recordset
'    Dim Sql1 As String
'sql = "SELECT     dbo.TblCardAuthorizationReform.ID, dbo.TblCardAuthorizationReformDetails.Type, dbo.TblCardAuthorizationReformDetails.ID2, dbo.TblCardAuthorizationReformDetails.ID AS idd,"
'sql = sql & "                      dbo.TblMaintenanceWork.name AS NameM, dbo.TblMaintenanceWork.namee AS Nameem, dbo.TblCardAuthorizationReformDetails.Mainte,"
' sql = sql & "                      dbo.TblCardAuthorizationReformDetails.finish, dbo.TblCardAuthorizationReform.OrderStatus, dbo.TblMaintenanceWork.Type AS typemw,"
'sql = sql & "                       dbo.TblCardAuthorizationReformDetails.ID2"
'sql = sql & "  FROM         dbo.TblCardAuthorizationReform FULL OUTER JOIN"
'sql = sql & "                       dbo.TblMaintenanceWork RIGHT OUTER JOIN"
'sql = sql & "                       dbo.TblCardAuthorizationReformDetails ON dbo.TblMaintenanceWork.Id = dbo.TblCardAuthorizationReformDetails.Mainte ON"
'sql = sql & "                       dbo.TblCardAuthorizationReform.id = dbo.TblCardAuthorizationReformDetails.id"
''sql = sql & "  Where (dbo.TblCardAuthorizationReform.id =" & id & ") And (dbo.TblCardAuthorizationReformDetails.Type = 0) And (dbo.TblCardAuthorizationReformDetails.finish = 0)"
''sql = sql & " WHERE     (dbo.TblCardAuthorizationReform.ID = " & id & ")"
 '  Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
 '  If Rs3.RecordCount > 0 Then
 '   str = IIf(Not IsNull(Rs3("NameM").value), Rs3("NameM").value, "")
 '   Index = Rs3("ID2").value
 '   Index = Index - 1
 'Sql1 = " SELECT     dbo.TblCardAuthorizationReformDetails.ID2, dbo.TblCardAuthorizationReformDetails.ID, dbo.TblMaintenanceWork.Id AS idm,"
 ' Sql1 = Sql1 & "                     dbo.TblCardAuthorizationReformDetails.Type , dbo.TblMaintenanceWork.name, dbo.TblMaintenanceWork.namee"
'Sql1 = Sql1 & "  FROM         dbo.TblCardAuthorizationReformDetails INNER JOIN"
'Sql1 = Sql1 & "                       dbo.TblMaintenanceWork ON dbo.TblCardAuthorizationReformDetails.Mainte = dbo.TblMaintenanceWork.Id"
'Sql1 = Sql1 & "  Where (dbo.TblCardAuthorizationReformDetails.Type = 0) And (dbo.TblCardAuthorizationReformDetails.ID2 =" & Index & ")And (dbo.TblCardAuthorizationReformDetails.finish = 1)"
'     Rs4.Open Sql1, Cn, adOpenStatic, adLockOptimistic, adCmdText
'    If Rs4.RecordCount > 0 Then
'      strb = IIf(Not IsNull(Rs4("name").value), Rs4("name").value, "")
'      End If
'  If Rs3("typemw").value = True Then
' x = 1
'' Else
 'x = 0
 'End If
 'End If
 'Exit Function
 '
'End Function

Private Sub ChangeLang()

    Cmd(1).Caption = "Delete"
 '  Cmd(0).Caption = "View Report"
   Cmd(2).Caption = "Exit"
 
  lbl(10).Caption = "  Sales Person  to this Week on dates "

Me.Caption = "  Screen  dates for Sales Person"


     With Me.Fg
     .TextMatrix(0, .ColIndex("Emp_NameD")) = "Sales Person"
        .TextMatrix(0, .ColIndex("Serial")) = "NO"
        .TextMatrix(0, .ColIndex("CusName")) = "Customer Name"
        .TextMatrix(0, .ColIndex("PersonConc")) = "Person Responsible"
        .TextMatrix(0, .ColIndex("Tel")) = "Telephone"
        .TextMatrix(0, .ColIndex("Mobile")) = "Mobile"
       .TextMatrix(0, .ColIndex("Email")) = "Email"
        .TextMatrix(0, .ColIndex("Adress")) = "Address"
        .TextMatrix(0, .ColIndex("Tem")) = "Team"
        .TextMatrix(0, .ColIndex("VisitID")) = "Visit Type"
        .TextMatrix(0, .ColIndex("VisitDate1")) = "VisitDate"
        .TextMatrix(0, .ColIndex("FromTime11")) = "FromTime"
       .TextMatrix(0, .ColIndex("ToTime11")) = "ToTime"
        .TextMatrix(0, .ColIndex("Remark")) = "Remark"
       .TextMatrix(0, .ColIndex("a1")) = "Requirements"
        .TextMatrix(0, .ColIndex("wait")) = " SMS"
       ' .TextMatrix(0, .ColIndex("dateday")) = "DateNow"
         .TextMatrix(0, .ColIndex("sendsms")) = "Send SMS"
    End With
  '


  '
End Sub
Private Sub Form_Load()

    Dim GrdBack As ClsBackGroundPic
    Dim NameDay As String
    Dim Dcombos As ClsDataCombos
     Set Dcombos = New ClsDataCombos
     Dim FirstDayInWeek As Date
Dim LatDayInWeek As Date
Me.DataeDaye.value = Date
       SetDtpickerDate Me.FirstDate
     FirstDate.value = Date
   SetDtpickerDate Me.LastDate
   LastDate = Date
FirstDayInWeek = Me.FirstDate - _
 Weekday(Me.FirstDate, vbUseSystemDayOfWeek) + 1
 LatDayInWeek = Me.FirstDate - _
 Weekday(Me.FirstDate, vbUseSystemDayOfWeek) + 7
    Dcombos.GetTypeVisit Me.DcbTypeVisit1
    
 '   GetNameofDayes DataeDaye, NameDay
   ' lbl(0).Caption = NameDay & "         " & DataeDaye.value
   lbl(0).Caption = Format(FirstDayInWeek, "yyyy/mm/dd") & "   Çáì  " & Format(LatDayInWeek, "yyyy/mm/dd")
     SetDtpickerDate Me.FirstDate
     FirstDate.value = Date
   SetDtpickerDate Me.LastDate
   LastDate = Date
    GetData
    
        Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500
    
    
'Me.DtStart.value = ""
'Me.DtEnd.value = ""

'Me.RDALL.value = True
'Me.RdAll2.value = True
'    Set Dcombos = New ClsDataCombos
   ' Dcombos.GetEmployees Me.DCEmp_Name
'     Dcombos.GetClientName DcbClientname
'     Dcombos.GetTblCarModels DcbCarModel
'      Dcombos.GetTblMaintenanceWork Me.DCBMinten
'     Dcombos.GetTblCarsDataGroup DcbCarType
'    Set DCboSearch = New clsDCboSearch
'    Set DCboSearch.Client = Me.DcbClientname
    'Dcombos.GetUsers Me.DCUser
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Clear").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture

    Set GrdBack = New ClsBackGroundPic
 
 
 If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
   

End Sub

Private Sub Form_Unload(Cancel As Integer)

    FormPostion Me, SavePostion
    Set DCboSearch = Nothing
End Sub
Public Sub RetriveTems(Optional id As Integer, Optional ByRef StrTem As String)
   Dim StrSQL As String
    Dim rs As ADODB.Recordset
    Dim i As Integer
    Dim str As String
    Set rs = New ADODB.Recordset
  StrSQL = "  SELECT     TOP 100 PERCENT TblEmployee_1.Emp_ID, TblEmployee_1.Emp_Name, TblEmployee_1.Emp_Namee, TblEmployee_1.Fullcode,"
  StrSQL = StrSQL & "                    dbo.TblRegDateDelgateDails.DelgID"
StrSQL = StrSQL & " FROM         dbo.TblEmployee TblEmployee_1 RIGHT OUTER JOIN"
  StrSQL = StrSQL & "                    dbo.TblRegDateDelgateDails ON TblEmployee_1.Emp_ID = dbo.TblRegDateDelgateDails.EmpID"
StrSQL = StrSQL & " Where (dbo.TblRegDateDelgateDails.DelgID = " & val(id) & ") And (dbo.TblRegDateDelgateDails.Type = 0)"
rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
StrTem = ""
If rs.RecordCount > 0 Then
For i = 1 To rs.RecordCount
  If SystemOptions.UserInterface = ArabicInterface Then
StrTem = StrTem & rs("Emp_Name").value
StrTem = StrTem & ", " 'vbNewLine
Else
StrTem = StrTem & rs("Emp_Namee").value
StrTem = StrTem & " ," 'vbNewLine
End If
rs.MoveNext
Next i
Else
StrTem = "'"
End If
End Sub
Public Sub RetriveCompo(Optional id As Integer, Optional ByRef StrTem As String)
   Dim StrSQL As String
    Dim rs As ADODB.Recordset
    Dim i As Integer
    Dim str As String
    Set rs = New ADODB.Recordset
  StrSQL = " SELECT     dbo.TblRegDateDelgateDails.Id, dbo.TblRegDateDelgateDails.DelgID, dbo.TblRegDateDelgateDails.EmpID, dbo.TblRegDateDelgateDails.remark,"
   StrSQL = StrSQL & "                   dbo.TblRegDateDelgateDails.Type , dbo.TblCompo.name, dbo.TblCompo.namee, dbo.TblRegDateDelgateDails.Quantity"
StrSQL = StrSQL & " FROM         dbo.TblRegDateDelgateDails LEFT OUTER JOIN"
 StrSQL = StrSQL & "                      dbo.TblCompo ON dbo.TblRegDateDelgateDails.EmpID = dbo.TblCompo.Id"
 
StrSQL = StrSQL & " Where (dbo.TblRegDateDelgateDails.DelgID = " & val(id) & ") And (dbo.TblRegDateDelgateDails.Type = 1)"
rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
StrTem = ""
If rs.RecordCount > 0 Then
For i = 1 To rs.RecordCount
  If SystemOptions.UserInterface = ArabicInterface Then
StrTem = StrTem & rs("name").value & "  Çáßãíå  " & rs("Quantity").value
StrTem = StrTem & ", " 'vbNewLine
Else
StrTem = StrTem & rs("namee").value & "  Quantity  " & rs("Quantity").value
StrTem = StrTem & " ," 'vbNewLine
End If
rs.MoveNext
Next i
Else
StrTem = "'"
End If
End Sub
Public Sub GetData()
    Dim StrSQL As String
    Dim StrSQL1 As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
     Dim rs1 As ADODB.Recordset
    Dim id As Integer
    Dim StrTem As String
    Dim Msg As String
    Dim i As Integer
Dim ID1 As Integer
Dim cod As Integer
Dim strname As String
Dim StrTem1 As String
Dim strnameb As String
Dim FirstDayInWeek As Date
Dim LatDayInWeek As Date
FirstDayInWeek = Me.FirstDate - _
 Weekday(Me.FirstDate, vbUseSystemDayOfWeek) + 1
 LatDayInWeek = Me.FirstDate - _
 Weekday(Me.FirstDate, vbUseSystemDayOfWeek) + 7
 'MsgBox FirstDayInWeek
 'MsgBox LatDayInWeek

StrSQL = " SELECT     TOP 100 PERCENT dbo.TblRegDateDelgate.Id, TblEmployee_1.Emp_ID, dbo.TblRegDateDelgate.RecordDate, dbo.TblRegDateDelgate.BranchID, "
StrSQL = StrSQL & "                      dbo.TblRegDateDelgate.DelgID, TblEmployee_1.Emp_Code AS Emp_CodeD, TblEmployee_1.Emp_Name AS Emp_NameD,"
StrSQL = StrSQL & "                       TblEmployee_1.Emp_Name1 AS Emp_Name1D, TblEmployee_1.Emp_Name2 AS Emp_Name2D, TblEmployee_1.Emp_Name3 AS Emp_Name3D,"
 StrSQL = StrSQL & "                      TblEmployee_1.Emp_Name4 AS Emp_Name4D, TblEmployee_1.Nationality AS NationalityD, TblEmployee_1.Emp_Namee AS Emp_NameeD,"
 StrSQL = StrSQL & "                      TblEmployee_1.Emp_Namee1 AS Emp_Namee1D, TblEmployee_1.Emp_Namee2 AS Emp_Namee2D, TblEmployee_1.Emp_Namee3 AS Emp_Namee3D,"
StrSQL = StrSQL & "                       TblEmployee_1.Emp_Namee4 AS Emp_Namee4D, TblEmployee_1.Fullcode AS FullcodeD, dbo.TblRegDateDelgate.CustomerName, dbo.TblRegDateDelgate.Remark,"
StrSQL = StrSQL & "                       dbo.TblRegDateDelgate.VisitID, TblTypeVisit_1.name, TblTypeVisit_1.namee, dbo.TblRegDateDelgate.VisitID2, TblTypeVisit_1.name AS name2,"
StrSQL = StrSQL & "                       TblTypeVisit_1.namee AS namee2, dbo.TblRegDateDelgate.SpAsID, dbo.TblSpeciaAsement.name AS nameSp, dbo.TblSpeciaAsement.namee AS nameeSp,"
StrSQL = StrSQL & "                       dbo.TblRegDateDelgate.VisitDate, dbo.TblRegDateDelgate.Remark2, dbo.TblRegDateDelgate.TimeFrom1, dbo.TblRegDateDelgate.TimeFrom2,"
StrSQL = StrSQL & "                       dbo.TblRegDateDelgate.TimeTo1, dbo.TblRegDateDelgate.TimeTo2, dbo.TblRegDateDelgate.PersonConc, dbo.TblRegDateDelgate.Tel, dbo.TblRegDateDelgate.Mobile,"
StrSQL = StrSQL & "                       dbo.TblRegDateDelgate.Email, dbo.TblRegDateDelgate.JobID, dbo.TblRegDateDelgate.LongTime, dbo.TblRegDateDelgate.VisitDate1, dbo.TblRegDateDelgate.Entry,"
StrSQL = StrSQL & "                       dbo.TblRegDateDelgate.Map, dbo.TblRegDateDelgate.Adress, dbo.TblRegDateDelgate.NotAcept, dbo.TblRegDateDelgate.BillNo, dbo.TblRegDateDelgate.CustomerID,"
StrSQL = StrSQL & "                       dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblRegDateDelgate.FromTime1, TblRegTimeDelgate_1.name AS FromTime11,"
StrSQL = StrSQL & "                       dbo.TblRegDateDelgate.FromTime2, dbo.TblRegTimeDelgate.name AS FromTime22, dbo.TblRegDateDelgate.ToTime1, TblRegTimeDelgate_3.name AS ToTime11,"
StrSQL = StrSQL & "                       dbo.TblRegDateDelgate.ToTime2, TblRegTimeDelgate_2.name AS ToTime22"
StrSQL = StrSQL & "  FROM         dbo.TblSpeciaAsement RIGHT OUTER JOIN"
StrSQL = StrSQL & "                       dbo.TblCustemers RIGHT OUTER JOIN"
StrSQL = StrSQL & "                       dbo.TblRegTimeDelgate RIGHT OUTER JOIN"
 StrSQL = StrSQL & "                      dbo.TblRegTimeDelgate TblRegTimeDelgate_2 RIGHT OUTER JOIN"
StrSQL = StrSQL & "                       dbo.TblRegDateDelgate ON TblRegTimeDelgate_2.Id = dbo.TblRegDateDelgate.ToTime2 LEFT OUTER JOIN"
 StrSQL = StrSQL & "                      dbo.TblRegTimeDelgate TblRegTimeDelgate_3 ON dbo.TblRegDateDelgate.ToTime1 = TblRegTimeDelgate_3.Id ON"
StrSQL = StrSQL & "                       dbo.TblRegTimeDelgate.Id = dbo.TblRegDateDelgate.FromTime2 LEFT OUTER JOIN"
  StrSQL = StrSQL & "                     dbo.TblRegTimeDelgate TblRegTimeDelgate_1 ON dbo.TblRegDateDelgate.FromTime1 = TblRegTimeDelgate_1.Id ON"
 StrSQL = StrSQL & "                      dbo.TblCustemers.CusID = dbo.TblRegDateDelgate.CustomerID ON dbo.TblSpeciaAsement.Id = dbo.TblRegDateDelgate.SpAsID LEFT OUTER JOIN"
 StrSQL = StrSQL & "                      dbo.TblTypeVisit TblTypeVisit_1 ON dbo.TblRegDateDelgate.VisitID = TblTypeVisit_1.Id LEFT OUTER JOIN"
  StrSQL = StrSQL & "                     dbo.TblEmployee TblEmployee_1 ON dbo.TblRegDateDelgate.DelgID = TblEmployee_1.Emp_ID"
  StrSQL = StrSQL & " WHERE     (dbo.TblRegDateDelgate.VisitDate1 >='" & SQLDate(FirstDayInWeek) & "' ) and (dbo.TblRegDateDelgate.VisitDate1 <='" & SQLDate(LatDayInWeek) & "' )"
    BolBegine = False
    StrWhere = ""

 
  StrSQL = StrSQL & "  ORDER BY dbo.TblRegDateDelgate.VisitDate1, dbo.TblRegDateDelgate.TimeFrom1"
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
 'Set rs1 = New ADODB.Recordset
  '  rs1.Open StrSQL1, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
           ' Me.lbl(10).Caption = "äÊíÌÉ ÇáÈÍË=ÕÝÑ"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
          '  Me.lbl(10).Caption = "Search Results=0"
        End If

    ' Msg = "áÇÊæÌÏ ÈíÇäÇÊ ááÚÑÖ ÊæÇÝÞ ÔÑæØ ÇáÊÞÑíÑ"
    '    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    Else
'print_report StrSQL

        With Me.Fg
           .Clear flexClearScrollable, flexClearEverything
            .Rows = .FixedRows
            .Rows = rs.RecordCount + .FixedRows
'
            If SystemOptions.UserInterface = ArabicInterface Then
             '   Me.lbl(10).Caption = "äÊíÌÉ ÇáÈÍË=" & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
               ' Me.lbl(10).Caption = "Search Results=" & rs.RecordCount
            End If

           rs.MoveFirst
        
           For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs("ID").value), "", rs("ID").value)
                ID1 = val(IIf(IsNull(rs("ID").value), "", rs("ID").value))
                RetriveCompo ID1, StrTem1
                RetriveTems ID1, StrTem
             '   MsgBox StrTem
             .TextMatrix(i, .ColIndex("a1")) = StrTem1
                .TextMatrix(i, .ColIndex("Remark")) = StrTem
                     If Not (IsNull(rs("VisitDate1").value)) Then
                     If (rs("VisitDate1").value) = ChekDate.value Then
                   .TextMatrix(i, .ColIndex("VisitDate1")) = "" 'Format(rs("VisitDate1").value, "yyyy/M/d")
                   Else
                   .TextMatrix(i, .ColIndex("VisitDate1")) = Format(rs("VisitDate1").value, "yyyy/M/d")
                   ChekDate.value = Format(rs("VisitDate1").value, "yyyy/M/d")
                   End If
                End If
                   If Not (IsNull(rs("VisitID").value)) Then
                   Me.DcbTypeVisit1.BoundText = val(IIf(IsNull(rs("VisitID").value), 0, rs("VisitID").value))
            .TextMatrix(i, .ColIndex("VisitID")) = Me.DcbTypeVisit1.text
            End If
                  
                 If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("Emp_NameD")) = IIf(IsNull(rs("Emp_NameD").value), "", rs("Emp_NameD").value)
                       .TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(rs("CusName").value), "", rs("CusName").value)
                    Else
                        .TextMatrix(i, .ColIndex("Emp_NameD")) = IIf(IsNull(rs("Emp_NameeD").value), "", rs("Emp_NameeD").value)
                        .TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(rs("CusNamee").value), "", rs("CusNamee").value)
                       End If
                         .TextMatrix(i, .ColIndex("PersonConc")) = IIf(IsNull(rs("PersonConc").value), "", rs("PersonConc").value)
                .TextMatrix(i, .ColIndex("Tel")) = IIf(IsNull(rs("Tel").value), "", rs("Tel").value)
              
               .TextMatrix(i, .ColIndex("Mobile")) = IIf(IsNull(rs("Mobile").value), "", rs("Mobile").value)
                .TextMatrix(i, .ColIndex("Email")) = IIf(IsNull(rs("Email").value), "", rs("Email").value)
                 .TextMatrix(i, .ColIndex("Adress")) = IIf(IsNull(rs("Adress").value), "", rs("Adress").value)
         
                ' .Cell(flexcpBackColor, i, 12, i, 12) = &HFFFF&
               
                '  .TextMatrix(i, .ColIndex("Remark")) = IIf(IsNull(rs("Remark").value), "", rs("Remark").value)
                   .TextMatrix(i, .ColIndex("FromTime11")) = val(IIf(IsNull(rs("FromTime11").value), 0, rs("FromTime11").value))
                    .TextMatrix(i, .ColIndex("ToTime11")) = val(IIf(IsNull(rs("ToTime11").value), 0, rs("ToTime11").value))
                rs.MoveNext
               ' rs1.MoveNext
               Coloring
            Next i

            .AutoSize 0, .Cols - 1, False
          '  Me.lbl(1).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("AdvanceValue"), .Rows - 1, .ColIndex("AdvanceValue"))
        End With

    End If
'Retrivcoulme1
End Sub


Private Sub menue_Click(Index As Integer)
showsforms Index
Select Case Index

Case 16
GetData
Case 15

End Select
End Sub

Private Sub Timer1_Timer()
'retrivgride
GetData
End Sub
 

