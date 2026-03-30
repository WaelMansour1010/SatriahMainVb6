VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmFilCustomerSearch 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "««·»ÕÀ ⁄‰ »ÿ«ﬁ… «–‰ «’·«Õ"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   16485
   Icon            =   "FrmFilCustomerSearch.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5145
   ScaleWidth      =   16485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox DcbyearFactor 
      Height          =   315
      Left            =   7080
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   6000
      Width           =   2295
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   "»Ì«‰«  „·› «·⁄„Ì·"
      Height          =   2205
      Index           =   1
      Left            =   3960
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   2640
      Width           =   12435
      Begin VB.TextBox Txtzipcode 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   120
         TabIndex        =   31
         Top             =   1320
         Width           =   5115
      End
      Begin VB.TextBox TxtCode 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   9720
         TabIndex        =   30
         Top             =   180
         Width           =   1635
      End
      Begin VB.TextBox txtBox 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   6240
         TabIndex        =   7
         Top             =   1320
         Width           =   5115
      End
      Begin VB.TextBox TxtAdress 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   120
         TabIndex        =   4
         Top             =   1680
         Width           =   11235
      End
      Begin VB.TextBox TxtFax 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   6240
         TabIndex        =   6
         Top             =   960
         Width           =   5115
      End
      Begin VB.TextBox TxtEmail 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   5115
      End
      Begin VB.TextBox TxtClientName 
         Alignment       =   2  'Center
         Height          =   345
         Left            =   120
         TabIndex        =   1
         Top             =   180
         Width           =   9555
      End
      Begin VB.TextBox txtmobile 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   6240
         TabIndex        =   2
         Top             =   600
         Width           =   5115
      End
      Begin VB.TextBox TxtPhone 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   5115
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«·—„“ «·»—ÌœÌ"
         Height          =   255
         Index           =   0
         Left            =   5160
         TabIndex        =   32
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "’‰œÊﬁ »—Ìœ"
         Height          =   195
         Index           =   18
         Left            =   11400
         TabIndex        =   29
         Top             =   1380
         Width           =   885
      End
      Begin VB.Label lbladdress 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«·⁄‰Ê«‰"
         Height          =   255
         Left            =   11400
         TabIndex        =   28
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·›«ﬂ”"
         Height          =   195
         Index           =   17
         Left            =   11400
         TabIndex        =   27
         Top             =   1020
         Width           =   885
      End
      Begin VB.Label lblemail 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«·«Ì„Ì·"
         Height          =   255
         Left            =   5280
         TabIndex        =   26
         Top             =   960
         Width           =   855
      End
      Begin VB.Label LblClientName 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·⁄„Ì·"
         Height          =   195
         Left            =   11940
         TabIndex        =   25
         Top             =   240
         Width           =   420
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÃÊ«· «·⁄„Ì·"
         Height          =   195
         Index           =   11
         Left            =   11280
         TabIndex        =   24
         Top             =   600
         Width           =   1020
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "Â« › «·⁄„Ì·"
         Height          =   195
         Index           =   9
         Left            =   5280
         TabIndex        =   23
         Top             =   660
         Width           =   885
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   16920
      TabIndex        =   21
      Text            =   " „ „Ê«›ﬁ… «·⁄„"
      Top             =   120
      Width           =   1035
   End
   Begin VB.ComboBox DcbOrderStatus 
      Height          =   315
      Left            =   16800
      TabIndex        =   20
      Top             =   600
      Width           =   2775
   End
   Begin VB.Frame lbreg 
      BackColor       =   &H00E2E9E9&
      Caption         =   " «—ÌŒ «· ”ÃÌ·"
      Height          =   1275
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   2640
      Width           =   3915
      Begin MSComCtl2.DTPicker DtpDateFrom 
         Height          =   330
         Left            =   90
         TabIndex        =   9
         Top             =   270
         Width           =   3030
         _ExtentX        =   5345
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   102694915
         CurrentDate     =   38887
      End
      Begin MSComCtl2.DTPicker DtpDateTo 
         Height          =   330
         Left            =   90
         TabIndex        =   10
         Top             =   630
         Width           =   3030
         _ExtentX        =   5345
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   102694915
         CurrentDate     =   38887
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   195
         Index           =   4
         Left            =   3240
         TabIndex        =   17
         Top             =   330
         Width           =   540
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï"
         Height          =   195
         Index           =   3
         Left            =   3255
         TabIndex        =   16
         Top             =   660
         Width           =   480
      End
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   0
      Left            =   2730
      TabIndex        =   0
      Top             =   4800
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
      Left            =   1890
      TabIndex        =   11
      Top             =   4800
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
      Left            =   1110
      TabIndex        =   12
      Top             =   4800
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
   Begin VSFlex8UCtl.VSFlexGrid Fg 
      Height          =   2625
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   16395
      _cx             =   28919
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
      Cols            =   22
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmFilCustomerSearch.frx":038A
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
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      ForeColor       =   &H00000080&
      Height          =   315
      Index           =   10
      Left            =   60
      TabIndex        =   19
      Top             =   2640
      Width           =   2775
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·≈Ã„«·Ï"
      Height          =   285
      Index           =   2
      Left            =   1920
      TabIndex        =   18
      Top             =   4200
      Width           =   945
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      ForeColor       =   &H00000080&
      Height          =   285
      Index           =   1
      Left            =   60
      TabIndex        =   13
      Top             =   4260
      Width           =   1785
   End
End
Attribute VB_Name = "FrmFilCustomerSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim DCboSearch As clsDCboSearch

Private Sub Cmd_Click(Index As Integer)

    Select Case Index

        Case 0
 
 GetData
      
        Case 1
            clear_all Me
Me.DtpDateFrom.value = ""
Me.DtpDateTo.value = ""
            If SystemOptions.UserInterface = ArabicInterface Then
               ' Me.lbl(0).Caption = "‰ ÌÃ… «·»ÕÀ"
            Else
               ' Me.lbl(0).Caption = "Search Results"
            End If

        Case 2
            Unload Me
    End Select

End Sub





Private Sub Fg_Click()


  FrmCarAuthontication.TxtClientCode.Text = ((Me.fg.TextMatrix(Me.fg.Row, Me.fg.ColIndex("id"))))
  FrmCarAuthontication.TxtCliientName.Text = Me.fg.TextMatrix(Me.fg.Row, Me.fg.ColIndex("ClientName"))
  FrmCarAuthontication.txtmobile.Text = Me.fg.TextMatrix(Me.fg.Row, Me.fg.ColIndex("mobile"))
  FrmCarAuthontication.TxtClientPhone.Text = Me.fg.TextMatrix(Me.fg.Row, Me.fg.ColIndex("Telephone"))
  FrmCarAuthontication.txtFax.Text = Me.fg.TextMatrix(Me.fg.Row, Me.fg.ColIndex("fax"))
FrmCarAuthontication.TxtBox.Text = Me.fg.TextMatrix(Me.fg.Row, Me.fg.ColIndex("box"))
FrmCarAuthontication.txtboxzip.Text = Me.fg.TextMatrix(Me.fg.Row, Me.fg.ColIndex("zipcode"))
FrmCarAuthontication.txtEmail.Text = Me.fg.TextMatrix(Me.fg.Row, Me.fg.ColIndex("email"))
FrmCarAuthontication.txtAddres.Text = Me.fg.TextMatrix(Me.fg.Row, Me.fg.ColIndex("address"))
FrmCarAuthontication.retInfoCustomer Me.fg.TextMatrix(Me.fg.Row, Me.fg.ColIndex("id"))



End Sub





Private Sub Form_Activate()
'   PutFormOnTop Me.hWnd
End Sub

Private Sub Form_Load()
    Dim GrdBack As ClsBackGroundPic
    Dim Dcombos As ClsDataCombos

    Set Dcombos = New ClsDataCombos
   ' Dcombos.GetEmployees Me.DCEmp_Name
   '  Dcombos.GetClientName Me.DCEmp_Name

    Set DCboSearch = New clsDCboSearch
    


 
    
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Clear").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture

  '  CenterForm Me
'GetData
'    FormPostion Me, GetPostion
    Set GrdBack = New ClsBackGroundPic

    With Me.fg
        Set .WallPaper = GrdBack.Picture
        .AutoSize 0, .Cols - 1, False
    End With
 If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
    SetDtpickerDate Me.DtpDateFrom
    SetDtpickerDate Me.DtpDateTo

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
StrSQL = "select * from TblCustemers "

    BolBegine = False
    StrWhere = ""

    'If val(Me.TxtIDFrom.text) <> 0 Then
    '    If BolBegine = True Then
    '        StrWhere = StrWhere & " dbo.TblCardAuthorizationReform.ID >=" & val(Me.TxtIDFrom.text) & ""
    '    Else
    '        BolBegine = True
    '        StrWhere = " Where dbo.TblCardAuthorizationReform.ID>=" & val(Me.TxtIDFrom.text) & ""
    '    End If
    'End If
 

    'If val(Me.TxtIDTO.text) <> 0 Then
    ''    If BolBegine = True Then
    '        StrWhere = StrWhere & " AND dbo.TblCardAuthorizationReform.ID <=" & val(Me.TxtIDTO.text) & ""
    '    Else
    '        BolBegine = True
    '        StrWhere = " Where dbo.TblCardAuthorizationReform.ID <=" & val(Me.TxtIDTO.text) & ""
    '    End If
    'End If
    '///////////////////
         If Me.TxtCode.Text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND Fullcode like '%" & Me.TxtCode.Text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where Fullcode like '%" & Me.TxtCode.Text & "%'"
        End If
    End If
         If TxtClientName.Text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND CusName like '%" & Me.TxtClientName.Text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where CusName like '%" & Me.TxtClientName.Text & "%'"
        End If
    End If
    '''////////////////
     If TxtPhone.Text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND Cus_Phone like '%" & Me.TxtPhone.Text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where Cus_Phone like '%" & Me.TxtPhone.Text & "%'"
        End If
    End If
    ''''''''''''''/
    
         If txtmobile.Text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND Cus_mobile like '%" & Me.txtmobile.Text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where Cus_mobile like '%" & Me.txtmobile.Text & "%'"
        End If
    End If
    '////////////////////
    
          If txtEmail.Text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND E_mail like '%" & Me.txtEmail.Text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where E_mail like '%" & Me.txtEmail.Text & "%'"
        End If
    End If
    '''''''''''''''''/
          If txtFax.Text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND FaxNumber like '%" & Me.txtFax.Text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where FaxNumber like '%" & Me.txtFax.Text & "%'"
        End If
    End If
    ''''''''/////////////////
          If Me.TxtAdress.Text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND Remark like '%" & Me.TxtAdress.Text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where Remark like '%" & Me.TxtAdress.Text & "%'"
        End If
    End If
    ''''''''''''///////
          If TxtBox.Text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND BoxMil like '%" & Me.TxtBox.Text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where BoxMil like '%" & Me.TxtBox.Text & "%'"
        End If
    End If
'////////////////////////Data of cars

 If Me.Txtzipcode.Text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND ZipCode like '%" & Me.Txtzipcode.Text & "%'"
        Else
            BolBegine = True
            StrWhere = " Where ZipCode like '%" & Me.Txtzipcode.Text & "%'"
        End If
    End If
  




    If Not IsNull(Me.DtpDateFrom.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND RecordDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where RecordDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
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
    StrSQL = StrSQL & " Order By dbo.TblCustemers.CusID"
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            Me.lbl(10).Caption = "Search Results=0"
        End If

        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «·»ÕÀ"
        Cmd_Click (1)
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    Else

        With Me.fg
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
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs("Fullcode").value), "", rs("Fullcode").value)
                        
                If Not (IsNull(rs("RecordDate").value)) Then
                    .TextMatrix(i, .ColIndex("RecordDate")) = Format(rs("RecordDate").value, "yyyy/M/d")
                End If
            
                .TextMatrix(i, .ColIndex("ClientName")) = IIf(IsNull(rs("CusName").value), "", rs("CusName").value)
                .TextMatrix(i, .ColIndex("Telephone")) = IIf(IsNull(rs("Cus_Phone").value), "", rs("Cus_Phone").value)
                .TextMatrix(i, .ColIndex("mobile")) = IIf(IsNull(rs("Cus_mobile").value), "", rs("Cus_mobile").value)
                .TextMatrix(i, .ColIndex("box")) = IIf(IsNull(rs("BoxMil").value), "", rs("BoxMil").value)
                .TextMatrix(i, .ColIndex("fax")) = IIf(IsNull(rs("FaxNumber").value), "", rs("FaxNumber").value)
                .TextMatrix(i, .ColIndex("email")) = IIf(IsNull(rs("E_mail").value), "", rs("E_mail").value)
                .TextMatrix(i, .ColIndex("address")) = IIf(IsNull(rs("Remark").value), "", rs("Remark").value)
                
                .TextMatrix(i, .ColIndex("zipcode")) = IIf(IsNull(rs("ZipCode").value), "", rs("ZipCode").value)
              
                rs.MoveNext
            Next i

            .AutoSize 0, .Cols - 1, False
         '   Me.lbl(1).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("AdvanceValue"), .Rows - 1, .ColIndex("AdvanceValue"))
        End With

    End If

End Sub

Private Sub ChangeLang()
 
    Cmd(1).Caption = "Delete"
    Cmd(0).Caption = "Search"
    Cmd(2).Caption = "Exit"
  Me.Caption = "Search Customer"

Me.LblClientName.Caption = "ClientName"
lbl(4).Caption = "From"
lbl(3).Caption = "To"
lbreg.Caption = "Data Regstration"
Fra(1).Caption = "Data Of Customer"
lbl(9).Caption = "Telephone"
lbl(11).Caption = "Mobile"
Me.lblemail.Caption = "Email"
lbl(17).Caption = "Fax"
Me.lbladdress.Caption = "Address"
lbl(0).Caption = "Zip Code"
lbl(18).Caption = "BoxMail"
lbl(2).Caption = "Total"
     With Me.fg
        .TextMatrix(0, .ColIndex("Serial")) = "NO"
        .TextMatrix(0, .ColIndex("id")) = "Code"
        .TextMatrix(0, .ColIndex("RecordDate")) = "Date"
         .TextMatrix(0, .ColIndex("ClientName")) = "Customer Name"
        .TextMatrix(0, .ColIndex("Telephone")) = "Telephone"
       
       '''''''/////////////////////
          
                  .TextMatrix(0, .ColIndex("mobile")) = "Mobile"
                .TextMatrix(0, .ColIndex("box")) = "Mailbox"
                .TextMatrix(0, .ColIndex("fax")) = "Fax"
                .TextMatrix(0, .ColIndex("email")) = "Email"
                .TextMatrix(0, .ColIndex("address")) = "Address"
                
                .TextMatrix(0, .ColIndex("zipcode")) = " Zip Code."
               
    End With
  '
End Sub

 

