VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmSearchShowTech 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "«·»ÕÀ ⁄‰ ⁄—Ê÷ «”⁄«— «·„’«⁄œ"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10080
   Icon            =   "FrmSearchShowTech.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   4815
   ScaleWidth      =   10080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.TextBox TxtProject 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   4560
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   3720
      Width           =   4455
   End
   Begin VB.TextBox TxtCustomer 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   4560
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   3390
      Width           =   4455
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   " «—ÌŒ «· ”ÃÌ·"
      Height          =   1035
      Index           =   1
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   3270
      Width           =   4335
      Begin MSComCtl2.DTPicker DtpDateFrom 
         Height          =   330
         Left            =   2370
         TabIndex        =   6
         Top             =   270
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   96010243
         CurrentDate     =   38887
      End
      Begin MSComCtl2.DTPicker DtpDateTo 
         Height          =   330
         Left            =   2370
         TabIndex        =   7
         Top             =   630
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   96010243
         CurrentDate     =   38887
      End
      Begin Dynamic_Byte.NourHijriCal DtpDateFromH 
         Height          =   315
         Left            =   720
         TabIndex        =   22
         Top             =   270
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
      End
      Begin Dynamic_Byte.NourHijriCal DtpDateToH 
         Height          =   315
         Left            =   720
         TabIndex        =   23
         Top             =   630
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï"
         Height          =   195
         Index           =   3
         Left            =   3735
         RightToLeft     =   -1  'True
         TabIndex        =   9
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
         TabIndex        =   8
         Top             =   330
         Width           =   660
      End
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·⁄—÷"
      Height          =   645
      Index           =   2
      Left            =   4560
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   2700
      Width           =   5475
      Begin VB.TextBox TxtIDFrom 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   180
         Width           =   1635
      End
      Begin VB.TextBox TxtIDTO 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   180
         Width           =   1515
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   195
         Index           =   5
         Left            =   4335
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   240
         Width           =   660
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï"
         Height          =   195
         Index           =   6
         Left            =   1860
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   240
         Width           =   645
      End
   End
   Begin VSFlex8UCtl.VSFlexGrid Fg 
      Height          =   2625
      Left            =   30
      TabIndex        =   10
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
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmSearchShowTech.frx":038A
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
      TabIndex        =   11
      Top             =   4440
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
      TabIndex        =   12
      Top             =   4440
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
      TabIndex        =   13
      Top             =   4440
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
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·⁄„Ì·"
      Height          =   195
      Index           =   9
      Left            =   9240
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   3480
      Width           =   720
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·„‘—Ê⁄"
      Height          =   285
      Index           =   8
      Left            =   9090
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   3870
      Width           =   945
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·≈Ã„«·Ï"
      Height          =   285
      Index           =   2
      Left            =   2970
      RightToLeft     =   -1  'True
      TabIndex        =   17
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
      TabIndex        =   16
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
      TabIndex        =   15
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
      TabIndex        =   14
      Top             =   2700
      Width           =   2055
   End
End
Attribute VB_Name = "FrmSearchShowTech"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DCboSearch As clsDCboSearch
Public searchtype As Integer

Private Sub Cmd_Click(Index As Integer)

    Select Case Index

        Case 0
            GetData

        Case 1
            clear_all Me
DtpDateFrom.value = ""
Me.DtpDateTo.value = ""
            If SystemOptions.UserInterface = ArabicInterface Then
                Me.lbl(0).Caption = "‰ ÌÃ… «·»ÕÀ"
            Else
                Me.lbl(0).Caption = "Search Results"
            End If

        Case 2
            Unload Me
    End Select

End Sub



Private Sub DtpDateFrom_Change()
If DtpDateFrom.value <> "" Then
 DtpDateFromH.value = ToHijriDate(DtpDateFrom.value)
 End If
End Sub



Private Sub DtpDateFromH_LostFocus()
    VBA.Calendar = vbCalGreg
           DtpDateFrom.value = ToGregorianDate(DtpDateFromH.value)
End Sub


Private Sub DtpDateTo_Change()
If DtpDateTo.value <> "" Then
 DtpDateToH.value = ToHijriDate(DtpDateTo.value)
 End If
End Sub



Private Sub DtpDateToH_LostFocus()
 VBA.Calendar = vbCalGreg
           DtpDateTo.value = ToGregorianDate(DtpDateToH.value)
End Sub

Private Sub fg_Click()

    With Me.Fg

        If .Row = -1 Then Exit Sub
        If .Col = -1 Then Exit Sub
        If val(.TextMatrix(.Row, .ColIndex("ID"))) = 0 Then
            Exit Sub
        End If

                
                        Select Case searchtype
                        Case 0
                                FrmQotation.Retrive val(.TextMatrix(.Row, .ColIndex("ID")))
                       
                     Case 1
                             FrmPO1.TxtPONo.text = val(.TextMatrix(.Row, .ColIndex("ID")))
                       End Select
                       
    End With

End Sub

Private Sub Form_Activate()
    PutFormOnTop Me.hwnd
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

    SetDtpickerDate Me.DtpDateFrom
    SetDtpickerDate Me.DtpDateTo

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
StrSQL = " SELECT     ID, BranchID, UserID, CityID, ElevatorCount, Supply, Granty, Instal, Show, InstalDMY, SupplyDMY, GrantyDMY, ShowDMY, RecordDate, RecordDateH, CusName, "
StrSQL = StrSQL & "                      projectname , Xy, ZibCode, CityIDH"
StrSQL = StrSQL & " From dbo.TblQutations"

 
    BolBegine = False
    StrWhere = ""

    If val(Me.TxtIDFrom.text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND  ID >=" & val(Me.TxtIDFrom.text) & ""
        Else
            BolBegine = True
            StrWhere = " Where  ID >=" & val(Me.TxtIDFrom.text) & ""
        End If
    End If

    If val(Me.TxtIDTO.text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND  ID <=" & val(Me.TxtIDTO.text) & ""
        Else
            BolBegine = True
            StrWhere = " Where  ID <=" & val(Me.TxtIDTO.text) & ""
        End If
    End If

   ' If Me.DCItem.BoundText <> "" Then
   '     If BolBegine = True Then
   '         StrWhere = StrWhere & " AND dbo.TblDistriExpensItemDet2.ItemID=" & Me.DCItem.BoundText & ""
   '     Else
   '         BolBegine = True
   '         StrWhere = " Where dbo.TblDistriExpensItemDet2.ItemID=" & Me.DCItem.BoundText & ""
   ''     End If
    'End If

    If Me.TxtCustomer.text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND CusName like '%" & Me.TxtCustomer.text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where  CusName  like '%" & Me.TxtCustomer.text & "%'"
        End If
    End If
     If Me.TxtProject.text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND ProjectName like '%" & Me.TxtProject.text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where  ProjectName like '%" & Me.TxtProject.text & "%'"
        End If
    End If
       ' If Me.TxtItemCode.text <> "" Then
       ' If BolBegine = True Then
       '     StrWhere = StrWhere & " AND  dbo.TblLink_Item_To_Store_Details2.ItemID=" & Me.TxtItemCode.text & ""
       ' Else
       '     BolBegine = True
       '     StrWhere = " Where  dbo.TblLink_Item_To_Store_Details2.ItemID=" & Me.TxtItemCode.text & ""
       ' End If
  '  End If

    If Not IsNull(Me.DtpDateFrom.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND  RecordDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where  RecordDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        End If
    End If

    If Not IsNull(Me.DtpDateTo.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND  RecordDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where  RecordDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        End If
    End If

    '-----------------------------------

    StrSQL = StrSQL & StrWhere
    StrSQL = StrSQL & " Order By id "
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
                .TextMatrix(i, .ColIndex("ID")) = IIf(IsNull(rs("ID").value), "", rs("ID").value)
                        
                If Not (IsNull(rs("RecordDate").value)) Then
                    .TextMatrix(i, .ColIndex("RecordDate")) = Format(rs("RecordDate").value, "yyyy/M/d")
                End If
             
                .TextMatrix(i, .ColIndex("RecordDateH")) = IIf(IsNull(rs("RecordDateH").value), "", rs("RecordDateH").value)
                .TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(rs("CusName").value), "", rs("CusName").value)
                .TextMatrix(i, .ColIndex("ProjectName")) = IIf(IsNull(rs("ProjectName").value), "", rs("ProjectName").value)
            
            
                
                
              
                rs.MoveNext
            Next i

            .AutoSize 0, .Cols - 1, False
          '  Me.lbl(1).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("AdvanceValue"), .Rows - 1, .ColIndex("AdvanceValue"))
        End With

    End If

End Sub





