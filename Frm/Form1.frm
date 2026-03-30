VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmPresentTimeSearch 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "«·»ÕÀ ⁄‰ «·Õ÷Ê— Ê«·√‰’—«›"
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6660
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5295
   ScaleWidth      =   6660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   "ÌÊ„ «·√‰’—«›"
      Height          =   1035
      Index           =   3
      Left            =   330
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   3720
      Width           =   2055
      Begin MSComCtl2.DTPicker DtpDepFrom 
         Height          =   330
         Left            =   90
         TabIndex        =   24
         Top             =   270
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   100073475
         CurrentDate     =   38887
      End
      Begin MSComCtl2.DTPicker DtpDepTo 
         Height          =   330
         Left            =   90
         TabIndex        =   25
         Top             =   630
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   100073475
         CurrentDate     =   38887
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   195
         Index           =   9
         Left            =   1740
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   330
         Width           =   180
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï"
         Height          =   195
         Index           =   8
         Left            =   1695
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   660
         Width           =   255
      End
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·⁄„·Ì…"
      Height          =   645
      Index           =   2
      Left            =   2850
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   2700
      Width           =   3795
      Begin VB.TextBox TxtIDTO 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   180
         Width           =   915
      End
      Begin VB.TextBox TxtIDFrom 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   1560
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   180
         Width           =   915
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï"
         Height          =   195
         Index           =   6
         Left            =   1020
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   240
         Width           =   405
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   195
         Index           =   5
         Left            =   2535
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   240
         Width           =   300
      End
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   " «—ÌŒ «· ”ÃÌ·"
      Height          =   1035
      Index           =   1
      Left            =   4560
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   3750
      Width           =   2055
      Begin MSComCtl2.DTPicker DtpDateFrom 
         Height          =   330
         Left            =   90
         TabIndex        =   12
         Top             =   270
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   100073475
         CurrentDate     =   38887
      End
      Begin MSComCtl2.DTPicker DtpDateTo 
         Height          =   330
         Left            =   90
         TabIndex        =   13
         Top             =   630
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   100073475
         CurrentDate     =   38887
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   195
         Index           =   4
         Left            =   1740
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   330
         Width           =   180
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï"
         Height          =   195
         Index           =   3
         Left            =   1695
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   660
         Width           =   255
      End
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   "ÌÊ„ «·Õ÷Ê—"
      Height          =   1035
      Index           =   0
      Left            =   2460
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   3750
      Width           =   2055
      Begin MSComCtl2.DTPicker DtpPresentFrom 
         Height          =   330
         Left            =   30
         TabIndex        =   7
         Top             =   270
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   100073475
         CurrentDate     =   38887
      End
      Begin MSComCtl2.DTPicker DtpPresentTO 
         Height          =   330
         Left            =   30
         TabIndex        =   8
         Top             =   630
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   100073475
         CurrentDate     =   38887
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï"
         Height          =   195
         Index           =   2
         Left            =   1605
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   660
         Width           =   375
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   195
         Index           =   1
         Left            =   1650
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   330
         Width           =   300
      End
   End
   Begin VSFlex8UCtl.VSFlexGrid Fg 
      Height          =   2625
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   6615
      _cx             =   11668
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
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"Form1.frx":038A
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
   Begin MSDataListLib.DataCombo DCEmp_Name 
      Height          =   315
      Left            =   2850
      TabIndex        =   1
      Top             =   3390
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      Text            =   "DCEmp_Name"
      RightToLeft     =   -1  'True
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   0
      Left            =   1680
      TabIndex        =   3
      Top             =   4830
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
      Left            =   840
      TabIndex        =   4
      Top             =   4830
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
      Left            =   60
      TabIndex        =   5
      Top             =   4830
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
   Begin MSDataListLib.DataCombo DCUser 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   2910
      TabIndex        =   22
      Tag             =   "„‰ ›÷·ﬂ √œŒ· —ﬁ„ «·ﬁ÷Ì…"
      Top             =   4830
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   -2147483624
      Text            =   ""
      RightToLeft     =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      ForeColor       =   &H00000080&
      Height          =   315
      Index           =   10
      Left            =   60
      RightToLeft     =   -1  'True
      TabIndex        =   28
      Top             =   2700
      Width           =   2745
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·„” Œœ„"
      Height          =   285
      Index           =   7
      Left            =   5400
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   4860
      Width           =   1155
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·„ÊŸ›"
      Height          =   315
      Index           =   0
      Left            =   5670
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   3390
      Width           =   975
   End
End
Attribute VB_Name = "FrmPresentTimeSearch"
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

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.lbl(0).Caption = "‰ ÌÃ… «·»ÕÀ"
            Else
                Me.lbl(0).Caption = "Search Results"
            End If

        Case 2
            Unload Me
    End Select

End Sub

Private Sub Fg_Click()

    With Me.FG

        If .Row = -1 Then Exit Sub
        If .Col = -1 Then Exit Sub
        If val(.TextMatrix(.Row, .ColIndex("Present_ID"))) = 0 Then
            Exit Sub
        End If

        If Not mdifrmmain.ActiveForm Is Nothing Then
            If mdifrmmain.ActiveForm.name = "FrmPresentTime" Then
                mdifrmmain.ActiveForm.Retrive val(.TextMatrix(.Row, .ColIndex("Present_ID")))
            End If
        End If

    End With

End Sub

Private Sub Form_Activate()
    PutFormOnTop Me.hWnd
End Sub

Private Sub Form_Load()
    Dim GrdBack As ClsBackGroundPic
    Dim Dcombos As ClsDataCombos

    Set Dcombos = New ClsDataCombos
    Dcombos.GetEmployees Me.DCEmp_Name
    Set DCboSearch = New clsDCboSearch
    Set DCboSearch.Client = Me.DCEmp_Name
    Dcombos.GetUsers Me.DCUser
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Clear").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture

    CenterForm Me

    FormPostion Me, GetPostion
    Set GrdBack = New ClsBackGroundPic

    With Me.FG
        Set .WallPaper = GrdBack.Picture
        .AutoSize 0, .Cols - 1, False
    End With

    SetDtpickerDate Me.DtpPresentFrom
    SetDtpickerDate Me.DtpPresentTO
    SetDtpickerDate Me.DtpDateFrom
    SetDtpickerDate Me.DtpDateTo
    SetDtpickerDate Me.DtpDepFrom
    SetDtpickerDate Me.DtpDepTo

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

    StrSQL = "SELECT dbo.TblEmployee.Emp_ID , dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name," & "dbo.tblPresentTime.Present_ID, dbo.tblPresentTime.PresentDate, dbo.tblPresentTime.Present_Type, " & "dbo.tblPresentTime.Present_Code, dbo.tblPresentTime.IntervalNO,dbo.tblPresentTime.GenPresentTime ," & "dbo.tblPresentTime.GenDepartureTime,dbo.TblUsers.UserName"
    StrSQL = StrSQL + " FROM dbo.TblEmployee INNER JOIN dbo.tblPresentTime ON dbo.TblEmployee.Emp_ID =" & "dbo.tblPresentTime.Emp_ID  INNER JOIN dbo.TblUsers ON dbo.tblPresentTime.UserID = dbo.TblUsers.UserID"

    BolBegine = False
    StrWhere = ""

    If val(Me.TxtIDFrom.text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND tblPresentTime.Present_ID >=" & val(Me.TxtIDFrom.text) & ""
        Else
            BolBegine = True
            StrWhere = " Where tblPresentTime.Present_ID >=" & val(Me.TxtIDFrom.text) & ""
        End If
    End If

    If val(Me.TxtIDTO.text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND tblPresentTime.Present_ID <=" & val(Me.TxtIDTO.text) & ""
        Else
            BolBegine = True
            StrWhere = " Where tblPresentTime.Present_ID <=" & val(Me.TxtIDTO.text) & ""
        End If
    End If

    If Me.DCEmp_Name.BoundText <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblEmployee.Emp_ID=" & Me.DCEmp_Name.BoundText & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblEmployee.Emp_ID=" & Me.DCEmp_Name.BoundText & ""
        End If
    End If

    If Me.DCUser.BoundText <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND tblPresentTime.UserID=" & Me.DCUser.BoundText & ""
        Else
            BolBegine = True
            StrWhere = " Where tblPresentTime.UserID=" & Me.DCUser.BoundText & ""
        End If
    End If

    If Not IsNull(Me.DtpDateFrom.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND tblPresentTime.PresentDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where tblPresentTime.PresentDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        End If
    End If

    If Not IsNull(Me.DtpDateTo.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND tblPresentTime.PresentDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where tblPresentTime.PresentDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        End If
    End If

    '-----------------------------------
    If Not IsNull(Me.DtpPresentFrom.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND CONVERT (nvarchar(50),GenPresentTime ,101) >=" & SQLDate(Me.DtpPresentFrom.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where CONVERT (nvarchar(50),GenPresentTime ,101) >=" & SQLDate(Me.DtpPresentFrom.value, True) & ""
        End If
    End If

    If Not IsNull(Me.DtpPresentTO.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND CONVERT (nvarchar(50),GenPresentTime ,101) <=" & SQLDate(Me.DtpPresentTO.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where CONVERT (nvarchar(50),GenPresentTime ,101) <=" & SQLDate(Me.DtpPresentTO.value, True) & ""
        End If
    End If

    '------------------------------
    If Not IsNull(Me.DtpDepFrom.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND CONVERT (nvarchar(50),GenDepartureTime ,101) >=" & SQLDate(Me.DtpDepFrom.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where CONVERT (nvarchar(50),GenDepartureTime ,101) >=" & SQLDate(Me.DtpDepFrom.value, True) & ""
        End If
    End If

    If Not IsNull(Me.DtpDepTo.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND CONVERT (nvarchar(50),GenDepartureTime ,101) <=" & SQLDate(Me.DtpDepTo.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where CONVERT (nvarchar(50),GenDepartureTime ,101) <=" & SQLDate(Me.DtpDepTo.value, True) & ""
        End If
    End If

    '----------------------------
    StrSQL = StrSQL & StrWhere
    StrSQL = StrSQL & " Order By tblPresentTime.Present_ID"
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            Me.lbl(10).Caption = "Search Results=0"
        End If

        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «·»ÕÀ"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    Else

        With Me.FG
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
                .TextMatrix(i, .ColIndex("Present_ID")) = IIf(IsNull(rs("Present_ID").value), "", rs("Present_ID").value)
                .TextMatrix(i, .ColIndex("Present_Code")) = IIf(IsNull(rs("Present_Code").value), "", rs("Present_Code").value)
                        
                If Not (IsNull(rs("PresentDate").value)) Then
                    .TextMatrix(i, .ColIndex("PresentDate")) = Format(rs("PresentDate").value, "yyyy/M/d")
                End If
            
                .TextMatrix(i, .ColIndex("Emp_Name")) = IIf(IsNull(rs("Emp_Name").value), "", rs("Emp_Name").value)

                If Not (IsNull(rs("GenPresentTime").value)) Then
                    .TextMatrix(i, .ColIndex("GenPresentTime")) = Format(rs("GenPresentTime").value, "hh:mm AMPM  yyyy/M/d")
                End If

                If Not (IsNull(rs("GenDepartureTime").value)) Then
                    .TextMatrix(i, .ColIndex("GenDepartureTime")) = Format(rs("GenDepartureTime").value, "hh:mm AMPM  yyyy/M/d")
                End If

                .TextMatrix(i, .ColIndex("UserName")) = IIf(IsNull(rs("UserName").value), "", rs("UserName").value)
                rs.MoveNext
            Next i

            .AutoSize 0, .Cols - 1, False
        End With

    End If

End Sub

Private Sub TxtIDFrom_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtIDFrom.text, 1)
End Sub

Private Sub TxtIDTO_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtIDTO.text, 1)
End Sub

