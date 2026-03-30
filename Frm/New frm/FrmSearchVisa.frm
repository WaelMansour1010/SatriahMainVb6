VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmSearchVisa 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "«·»ÕÀ ⁄‰ »Ì«‰«  «· √‘Ì—« "
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10080
   Icon            =   "FrmSearchVisa.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5625
   ScaleWidth      =   10080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   " «—ÌŒ «· ”ÃÌ·"
      Height          =   1035
      Index           =   0
      Left            =   5640
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   3840
      Width           =   4335
      Begin MSComCtl2.DTPicker StarDateFrom 
         Height          =   330
         Left            =   2370
         TabIndex        =   22
         Top             =   270
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   97714179
         CurrentDate     =   38887
      End
      Begin MSComCtl2.DTPicker StarDateTo 
         Height          =   330
         Left            =   2370
         TabIndex        =   23
         Top             =   630
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   97714179
         CurrentDate     =   38887
      End
      Begin Dynamic_Byte.NourHijriCal StarDateFromH 
         Height          =   315
         Left            =   720
         TabIndex        =   24
         Top             =   270
         Width           =   1575
         _extentx        =   2778
         _extenty        =   556
      End
      Begin Dynamic_Byte.NourHijriCal StarDateToH 
         Height          =   315
         Left            =   720
         TabIndex        =   25
         Top             =   630
         Width           =   1575
         _extentx        =   2778
         _extenty        =   556
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   195
         Index           =   8
         Left            =   3540
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   330
         Width           =   660
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï"
         Height          =   195
         Index           =   7
         Left            =   3735
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   660
         Width           =   495
      End
   End
   Begin VB.TextBox TxtOrder 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   4560
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   2760
      Width           =   4455
   End
   Begin VB.TextBox TxtVisa 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   4560
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   3120
      Width           =   4455
   End
   Begin VB.TextBox TxtHodono 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   4560
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   3510
      Width           =   4455
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   " «—ÌŒ «· ”ÃÌ·"
      Height          =   1035
      Index           =   1
      Left            =   1320
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   3870
      Width           =   4335
      Begin MSComCtl2.DTPicker EndDateFrom 
         Height          =   330
         Left            =   2370
         TabIndex        =   1
         Top             =   270
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   97714179
         CurrentDate     =   38887
      End
      Begin MSComCtl2.DTPicker EndDateTo 
         Height          =   330
         Left            =   2370
         TabIndex        =   2
         Top             =   630
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   97714179
         CurrentDate     =   38887
      End
      Begin Dynamic_Byte.NourHijriCal EndDateFromH 
         Height          =   315
         Left            =   720
         TabIndex        =   15
         Top             =   270
         Width           =   1575
         _extentx        =   2778
         _extenty        =   556
      End
      Begin Dynamic_Byte.NourHijriCal EndDateToH 
         Height          =   315
         Left            =   720
         TabIndex        =   16
         Top             =   630
         Width           =   1575
         _extentx        =   2778
         _extenty        =   556
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï"
         Height          =   195
         Index           =   3
         Left            =   3735
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   660
         Width           =   495
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   195
         Index           =   4
         Left            =   3540
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   330
         Width           =   660
      End
   End
   Begin VSFlex8UCtl.VSFlexGrid Fg 
      Height          =   2625
      Left            =   30
      TabIndex        =   5
      Top             =   0
      Width           =   10035
      _cx             =   17701
      _cy             =   4630
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   14871017
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
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
      Cols            =   15
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmSearchVisa.frx":038A
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
      Editable        =   0
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
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   0
      Left            =   1650
      TabIndex        =   6
      Top             =   5040
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   661
      ButtonPositionImage=   1
      Caption         =   "»ÕÀ"
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
      Height          =   375
      Index           =   1
      Left            =   810
      TabIndex        =   7
      Top             =   5040
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   661
      ButtonPositionImage=   1
      Caption         =   "„”Õ"
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
   Begin ImpulseButton.ISButton Cmd 
      Cancel          =   -1  'True
      Height          =   375
      Index           =   2
      Left            =   30
      TabIndex        =   8
      Top             =   5040
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      ButtonPositionImage=   1
      Caption         =   "Œ—ÊÃ"
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
   Begin MSDataListLib.DataCombo DcbNtionality 
      Height          =   315
      Left            =   0
      TabIndex        =   28
      Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
      Top             =   3120
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DcbCity 
      Height          =   315
      Left            =   0
      TabIndex        =   29
      Tag             =   "⁄›Ê« Ì—ÃÏ «œŒ«· √”„ «·ÕÌ"
      Top             =   3480
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·„œÌ‰…"
      Height          =   195
      Index           =   12
      Left            =   4050
      RightToLeft     =   -1  'True
      TabIndex        =   31
      Top             =   3480
      Width           =   435
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·Ã‰”Ì…"
      Height          =   195
      Index           =   11
      Left            =   3960
      RightToLeft     =   -1  'True
      TabIndex        =   30
      Top             =   3120
      Width           =   525
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·ÕœÊœ"
      Height          =   195
      Index           =   6
      Left            =   9285
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   3480
      Width           =   720
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «· √‘Ì—…"
      Height          =   195
      Index           =   5
      Left            =   9135
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   3120
      Width           =   870
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·ÿ·»"
      Height          =   195
      Index           =   9
      Left            =   9315
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   2760
      Width           =   645
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·≈Ã„«·Ï"
      Height          =   285
      Index           =   2
      Left            =   2970
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   2820
      Width           =   945
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      ForeColor       =   &H00000080&
      Height          =   285
      Index           =   1
      Left            =   60
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   2820
      Width           =   1185
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·’‰›"
      Height          =   195
      Index           =   0
      Left            =   11115
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   3630
      Width           =   720
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      ForeColor       =   &H00000080&
      Height          =   435
      Index           =   10
      Left            =   780
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   2700
      Width           =   2055
   End
End
Attribute VB_Name = "FrmSearchVisa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DCboSearch As clsDCboSearch

Private Sub Cmd_Click(Index As Integer)

    Select Case Index

        Case 0
            GetData

        Case 1
            clear_all Me
Me.StarDateFrom.value = ""
Me.StarDateTo.value = ""
Me.EndDateFrom.value = ""
Me.EndDateTO.value = ""
            If SystemOptions.UserInterface = ArabicInterface Then
                Me.lbl(0).Caption = "‰ ÌÃ… «·»ÕÀ"
            Else
                Me.lbl(0).Caption = "Search Results"
            End If

        Case 2
            Unload Me
    End Select

End Sub










Private Sub EndDateFrom_Change()
If EndDateFrom.value <> "" Then
 EndDateFromH.value = ToHijriDate(EndDateFrom.value)
 End If
End Sub

Private Sub EndDateFromH_GotFocus()
VBA.Calendar = vbCalGreg
           EndDateFrom.value = ToGregorianDate(EndDateFromH.value)
End Sub

Private Sub EndDateTo_Change()
If EndDateTO.value <> "" Then
 EndDateTOH.value = ToHijriDate(EndDateTO.value)
 End If
End Sub

Private Sub EndDateToH_GotFocus()
VBA.Calendar = vbCalGreg
           EndDateTO.value = ToGregorianDate(EndDateTOH.value)
End Sub

Private Sub Fg_Click()

    With Me.Fg

        If .Row = -1 Then Exit Sub
        If .Col = -1 Then Exit Sub
        If val(.TextMatrix(.Row, .ColIndex("id"))) = 0 Then
            Exit Sub
        End If

                FrmVisa.Retrive val(.TextMatrix(.Row, .ColIndex("id")))
       
    

    End With

End Sub

Private Sub Form_Activate()
    PutFormOnTop Me.hWnd
End Sub

Private Sub ChangeLang()
    'Dim XPic As IPictureDisp
    'Set XPic = Me.XPBtnMove(1).ButtonImage
    'Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    'Set Me.XPBtnMove(2).ButtonImage = XPic
    'Set XPic = Me.XPBtnMove(0).ButtonImage
    'Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    'Set Me.XPBtnMove(3).ButtonImage = XPic
'    Label1.Visible = False

    'Cmd(0).Caption = "New"
    'Cmd(1).Caption = "Edit"
    'Cmd(2).Caption = "Save"
    'Cmd(3).Caption = "Undo"
    Cmd(1).Caption = "Delete"
    Cmd(0).Caption = "Search"
 'Cmd(9).Caption = "Print"
    Cmd(2).Caption = "Exit"
 '   CmdHelp.Caption = "Help"

    Me.Caption = " Search For Distribution Expenses on Items "
    'EleHeader.Caption = Me.Caption
    lbl(5).Caption = "From"
    lbl(6).Caption = "To"
    lbl(4).Caption = "From"
    lbl(3).Caption = "To "
   ' Frame10.Caption = "Select Store"
    Fra(2).Caption = "Registration No  "
    Fra(1).Caption = "Registration Date  "
lbl(7).Caption = "AccountName"
lbl(8).Caption = "GroupName"
lbl(0).Caption = "ItemName"
lbl(2).Caption = "Total"
   'lbl(8).Caption = "By"
   ' lbl(7).Caption = "Curr rec."
   ' lbl(6).Caption = "rec. count"

   With Me.Fg
        .TextMatrix(0, .ColIndex("Serial")) = "Serial"
        .TextMatrix(0, .ColIndex("ID")) = "NO"
        .TextMatrix(0, .ColIndex("Account_Name")) = "AccountName"
         .TextMatrix(0, .ColIndex("RecordDate")) = "RecordDate"
        .TextMatrix(0, .ColIndex("GroupName")) = "GroupName"
         .TextMatrix(0, .ColIndex("ItemName")) = "ItemName"
        .TextMatrix(0, .ColIndex("TypeValue")) = "TypeValue"
        .TextMatrix(0, .ColIndex("Vlue")) = "Value"
        .TextMatrix(0, .ColIndex("RemarkD3")) = "Remarks"

    End With

End Sub

Private Sub Form_Load()
    Dim GrdBack As ClsBackGroundPic
    Dim Dcombos As ClsDataCombos

    Set Dcombos = New ClsDataCombos
  
    Set DCboSearch = New clsDCboSearch
   
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
Dcombos.GetCountrieshay Me.DcbCity
Dcombos.GETNationality Me.DcbNtionality
  
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Clear").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture

    CenterForm Me

    FormPostion Me, GetPostion
    Set GrdBack = New ClsBackGroundPic

    With Me.Fg
        Set .WallPaper = GrdBack.Picture
        .AutoSize 0, .Cols - 1, False
    End With

  '  SetDtpickerDate Me.DtpDateFrom
  '  SetDtpickerDate Me.DtpDateTo

End Sub

Private Sub Form_Unload(Cancel As Integer)

    FormPostion Me, SavePostion
    Set DCboSearch = Nothing
End Sub

Private Sub GetData()
    Dim StrSQL As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
StrSQL = "SELECT     dbo.TbVisaDeti.ID, dbo.TbVisaDeti.VisaID, dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode, "
StrSQL = StrSQL & "                      dbo.TblEmployee.Emp_Namee, dbo.TbVisaDeti.HododNo, dbo.TbVisaDeti.JobID, dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmpJobsTypes.JobTypeNamee,"
StrSQL = StrSQL & "                      dbo.TbVisaDeti.NotionalID, dbo.Nationality.name, dbo.Nationality.namee, dbo.TbVisaDeti.CityID, dbo.TblCountriesGovernments.GovernmentName,"
StrSQL = StrSQL & "                      dbo.TbVisa.OrderNo, dbo.TbVisa.VisaNo, dbo.TbVisa.Priod, dbo.TbVisa.DMYPriod, dbo.TbVisa.StarDate, dbo.TbVisa.StarDateH, dbo.TbVisa.EndDate,"
StrSQL = StrSQL & "                      dbo.TbVisa.EndDateH , dbo.TbVisa.ID AS IDM "
StrSQL = StrSQL & " FROM         dbo.Nationality RIGHT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TbVisa LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TbVisaDeti ON dbo.TbVisa.ID = dbo.TbVisaDeti.VisaID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblCountriesGovernments ON dbo.TbVisaDeti.CityID = dbo.TblCountriesGovernments.GovernmentID ON"
StrSQL = StrSQL & "                      dbo.Nationality.id = dbo.TbVisaDeti.NotionalID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblEmpJobsTypes ON dbo.TbVisaDeti.JobID = dbo.TblEmpJobsTypes.JobTypeID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblEmployee ON dbo.TbVisaDeti.EmpID = dbo.TblEmployee.Emp_ID"
StrSQL = StrSQL & ""
 
    BolBegine = False
    StrWhere = ""



    If val(Me.DcbCity.BoundText) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TbVisaDeti.CityID=" & Me.DcbCity.BoundText & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TbVisaDeti.CityID =" & Me.DcbCity.BoundText & ""
        End If
    End If
    

    If val(Me.DcbNtionality.BoundText) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TbVisaDeti.NotionalID=" & Me.DcbNtionality.BoundText & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TbVisaDeti.NotionalID =" & Me.DcbNtionality.BoundText & ""
        End If
    End If

    If Me.TxtOrder.text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TbVisa.OrderNo like '%" & Me.TxtOrder.text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TbVisa.OrderNo like '%" & Me.TxtOrder.text & "%'"
        End If
    End If
   If Me.TxtVisa.text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TbVisa.VisaNo like '%" & Me.TxtVisa.text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TbVisa.VisaNo like '%" & Me.TxtVisa.text & "%'"
        End If
    End If
       If Me.TxtHodono.text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TbVisaDeti.HododNo like '%" & Me.TxtHodono.text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TbVisaDeti.HododNo like '%" & Me.TxtHodono.text & "%'"
        End If
    End If

    If Not IsNull(Me.StarDateFrom.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND  dbo.TbVisa.StarDate >=" & SQLDate(Me.StarDateFrom.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TbVisa.StarDate >=" & SQLDate(Me.StarDateFrom.value, True) & ""
        End If
    End If

    If Not IsNull(Me.StarDateTo.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND  dbo.TbVisa.StarDate <=" & SQLDate(Me.StarDateTo.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TbVisa.StarDate <=" & SQLDate(Me.StarDateTo.value, True) & ""
        End If
    End If
    
    If Not IsNull(Me.EndDateFrom.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND  dbo.TbVisa.EndDate >=" & SQLDate(Me.EndDateFrom.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TbVisa.EndDate >=" & SQLDate(Me.EndDateFrom.value, True) & ""
        End If
    End If

    If Not IsNull(Me.EndDateTO.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND  dbo.TbVisa.EndDate <=" & SQLDate(Me.EndDateTO.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TbVisa.EndDate <=" & SQLDate(Me.EndDateTO.value, True) & ""
        End If
    End If
    '-----------------------------------

    StrSQL = StrSQL & StrWhere
    StrSQL = StrSQL & " Order By dbo.TbVisa.ID "
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            Me.lbl(10).Caption = "Search Results=0"
        End If

        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «·»ÕÀ"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    Else

        With Me.Fg
            .Clear flexClearScrollable, flexClearEverything
            .Rows = .FixedRows
            .Rows = rs.RecordCount + .FixedRows

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
                Me.lbl(10).Caption = "Search Results=" & rs.RecordCount
            End If

            rs.MoveFirst
        
            For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs("IDM").value), "", rs("IDM").value)
                        
                If Not (IsNull(rs("EndDate").value)) Then
                    .TextMatrix(i, .ColIndex("EndDate")) = Format(rs("EndDate").value, "yyyy/M/d")
                End If
                    If Not (IsNull(rs("StarDate").value)) Then
                    .TextMatrix(i, .ColIndex("StarDate")) = Format(rs("StarDate").value, "yyyy/M/d")
                End If
             
                .TextMatrix(i, .ColIndex("StarDateH")) = IIf(IsNull(rs("StarDateH").value), "", rs("StarDateH").value)
                .TextMatrix(i, .ColIndex("EndDateH")) = IIf(IsNull(rs("EndDateH").value), "", rs("EndDateH").value)
                If SystemOptions.UserInterface = ArabicInterface Then
                 .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(rs("name").value), "", rs("name").value)
               
                Else
                 .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(rs("namee").value), "", rs("namee").value)
                End If
                 .TextMatrix(i, .ColIndex("HododNo")) = IIf(IsNull(rs("HododNo").value), "", rs("HododNo").value)
                 .TextMatrix(i, .ColIndex("GovernmentName")) = IIf(IsNull(rs("GovernmentName").value), "", rs("GovernmentName").value)
                  .TextMatrix(i, .ColIndex("VisaNo")) = IIf(IsNull(rs("VisaNo").value), "", rs("VisaNo").value)
                 .TextMatrix(i, .ColIndex("OrderNo")) = IIf(IsNull(rs("OrderNo").value), "", rs("OrderNo").value)
              
                rs.MoveNext
            Next i

            .AutoSize 0, .Cols - 1, False
          '  Me.lbl(1).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("AdvanceValue"), .Rows - 1, .ColIndex("AdvanceValue"))
        End With

    End If

End Sub




Private Sub StarDateFrom_Change()
If StarDateFrom.value <> "" Then
 StarDateFromH.value = ToHijriDate(StarDateFrom.value)
 End If
End Sub

Private Sub StarDateFromH_LostFocus()
 VBA.Calendar = vbCalGreg
           StarDateFrom.value = ToGregorianDate(StarDateFromH.value)

End Sub



Private Sub StarDateTo_Change()
If StarDateTo.value <> "" Then
 StarDateToH.value = ToHijriDate(StarDateTo.value)
 End If
End Sub

Private Sub StarDateToH_GotFocus()
VBA.Calendar = vbCalGreg
           StarDateTo.value = ToGregorianDate(StarDateToH.value)
End Sub
