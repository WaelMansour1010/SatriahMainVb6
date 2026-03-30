VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{49003D3A-66CD-11D7-A449-E937BE2D9041}#1.0#0"; "ALLBUTTONS.ocx"
Begin VB.Form FrmEmployeeSearch1 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«·»ÕÀ ⁄‰ ⁄ﬁœ „ÊŸ›"
   ClientHeight    =   5100
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14295
   Icon            =   "FrmProfSearch1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5100
   ScaleWidth      =   14295
   Begin VB.TextBox Contract_ID 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   8520
      RightToLeft     =   -1  'True
      TabIndex        =   30
      Top             =   2400
      Width           =   2775
   End
   Begin VB.TextBox DCboEmployeesName 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   4200
      RightToLeft     =   -1  'True
      TabIndex        =   28
      Top             =   2760
      Width           =   2655
   End
   Begin ALLButtonS.ALLButton ALLButton1 
      Height          =   375
      Left            =   120
      TabIndex        =   21
      Top             =   4560
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "œ·«·«  «·«·Ê«‰"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmProfSearch1.frx":030A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.CheckBox XPChkSearchType 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·„ÊŸ› »«·ﬂ«„· ›ﬁÿ"
      Height          =   345
      Left            =   10050
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   6510
      Visible         =   0   'False
      Width           =   2265
   End
   Begin VSFlex8UCtl.VSFlexGrid Fg 
      Height          =   2085
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   14325
      _cx             =   25268
      _cy             =   3678
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
      FormatString    =   $"FrmProfSearch1.frx":0326
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
      Left            =   6360
      TabIndex        =   1
      Top             =   4440
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   661
      ButtonStyle     =   1
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
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      DisabledImageExtraction=   0
      ColorToggledHoverText=   16711680
      LowerToggledContent=   0   'False
      ColorTextShadow =   4210752
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   1
      Left            =   5040
      TabIndex        =   2
      Top             =   4440
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      ButtonStyle     =   1
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
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      LowerToggledContent=   0   'False
      ColorTextShadow =   4210752
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   2
      Left            =   3840
      TabIndex        =   3
      Top             =   4440
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      ButtonStyle     =   1
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
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      LowerToggledContent=   0   'False
      ColorTextShadow =   4210752
   End
   Begin MSDataListLib.DataCombo dcdep 
      Height          =   315
      Left            =   8520
      TabIndex        =   7
      Top             =   3120
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo dcnationality 
      Height          =   315
      Left            =   4200
      TabIndex        =   9
      Top             =   3120
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo dcjop 
      Height          =   315
      Left            =   360
      TabIndex        =   11
      Top             =   3120
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo dcdean 
      Height          =   315
      Left            =   360
      TabIndex        =   13
      Top             =   3480
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo dcstatus 
      Height          =   315
      Left            =   8520
      TabIndex        =   15
      Top             =   3480
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin ImpulseButton.ISButton Cmd1 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   2640
      TabIndex        =   17
      Top             =   4440
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "ÿ»«⁄…"
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
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      LowerToggledContent=   0   'False
      ColorTextShadow =   4210752
   End
   Begin MSDataListLib.DataCombo dckafel 
      Height          =   315
      Left            =   4200
      TabIndex        =   18
      Top             =   3480
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo dcEmpcode 
      Height          =   315
      Left            =   8520
      TabIndex        =   20
      Top             =   2760
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo dcekama 
      Height          =   315
      Left            =   360
      TabIndex        =   22
      Top             =   2760
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   315
      Left            =   8520
      TabIndex        =   26
      Top             =   3840
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·⁄ﬁœ"
      Height          =   285
      Index           =   11
      Left            =   11280
      RightToLeft     =   -1  'True
      TabIndex        =   29
      Top             =   2400
      Width           =   1035
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·„Êﬁ⁄"
      Height          =   405
      Index           =   10
      Left            =   11280
      RightToLeft     =   -1  'True
      TabIndex        =   27
      Top             =   3960
      Width           =   1035
   End
   Begin VB.Label lblType 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   135
      Left            =   9240
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   6840
      Width           =   855
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·«”„"
      Height          =   285
      Index           =   9
      Left            =   6840
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   2760
      Width           =   1275
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·ÊŸÌ›…"
      Height          =   285
      Index           =   8
      Left            =   2490
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   30
      Width           =   1275
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·ﬂ›Ì·"
      Height          =   285
      Index           =   7
      Left            =   6930
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   3510
      Width           =   1275
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "Õ«·… «·⁄„·"
      Height          =   405
      Index           =   6
      Left            =   11250
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   3510
      Width           =   1035
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·œÌ«‰…"
      Height          =   285
      Index           =   5
      Left            =   3090
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   3510
      Width           =   915
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·ÊŸÌ›…"
      Height          =   285
      Index           =   4
      Left            =   3210
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   3150
      Width           =   915
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·Ã‰”Ì…"
      Height          =   285
      Index           =   3
      Left            =   6930
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   3150
      Width           =   1275
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·ﬁ”„"
      Height          =   285
      Index           =   2
      Left            =   11280
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   3150
      Width           =   1035
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "ﬂÊœ «·„ÊŸ›"
      Height          =   285
      Index           =   1
      Left            =   11280
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   2760
      Width           =   1035
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·«ﬁ«„…"
      Height          =   285
      Index           =   0
      Left            =   3120
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   2760
      Width           =   1035
   End
End
Attribute VB_Name = "FrmEmployeeSearch1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim cSearchDcbo As clsDCboSearch

Public RetrunFrm As Form

Private Sub ALLButton1_Click()
    jobstatus.show
End Sub

Private Sub Cmd_Click(Index As Integer)
    On Error GoTo ErrTrap
    Set rs = New ADODB.Recordset

    Select Case Index

        Case 0

            If rs.State = adStateOpen Then
                rs.Close
            End If

            rs.Open Build_Sql, Cn, adOpenStatic, adLockOptimistic, adCmdText

            If rs.RecordCount < 1 Then
                FG.Clear flexClearScrollable, flexClearEverything
                FG.rows = 2
                ' Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
                ' MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Exit Sub
            End If

            Retrive
            FG.SetFocus

        Case 1
            clear_all Me
            FG.Clear flexClearScrollable, flexClearEverything

        Case 2
            Unload Me

    End Select

    Exit Sub
ErrTrap:

    If Err.Number = -2147217900 Then
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & CHR(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… „⁄«ÌÌ— «·»ÕÀ Ê√⁄œ «·„Õ«Ê·…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

End Sub

Private Sub Cmd1_Click()
    Dim sql As String
    
    'Dim Rs As ADODB.Recordset
    Dim xReport As New CRAXDRT.Report
    Dim xApp As New CRAXDRT.Application
    Dim rs As New ADODB.Recordset
    Dim reportpatath As String
    Dim CViewer As ClsReportViewer
    sql = Build_Sql
    rs.Open sql, Cn, adOpenStatic, adLockPessimistic, adCmdText
   
    If SystemOptions.UserInterface = ArabicInterface Then
        reportpatath = system_path & "\reports\emp\REPORT9.rpt"
    Else
        reportpatath = system_path & "\reports\emp\REPORT9e.rpt"
    End If

    Set xReport = xApp.OpenReport(reportpatath)
    xReport.Database.SetDataSource rs
'
'    Set FrmReport = New FrmReportViewer
'    FrmReport.CRViewer.ReportSource = xReport
'
'    'FrmReport.CRViewer.viewReport
'    'FrmReport.FireReport xReport, WindowTarget, "", , , , StrFileName
'
'    FrmReport.txtPath = reportpatath
'    FrmReport.show
'
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , reportpatath, , sql
 

    Screen.MousePointer = vbDefault
    '      xReport.ReportTitle = X
    Sendkeys "{RIGHT}"
 
End Sub

Private Sub DCboEmployeesName_KeyUp(KeyCode As Integer, _
                                    Shift As Integer)

    If KeyCode = 13 Then
        Cmd_Click (0)
    End If

End Sub

Private Sub dcdean_KeyUp(KeyCode As Integer, _
                         Shift As Integer)

    If KeyCode = 13 Then
        Cmd_Click (0)
    End If

End Sub

Private Sub dcdep_KeyUp(KeyCode As Integer, _
                        Shift As Integer)

    If KeyCode = 13 Then
        Cmd_Click (0)
    End If

End Sub

Private Sub dcekama_Change()
    'If dcekama.text = "" Then Exit Sub

End Sub

Private Sub dcekama_Click(Area As Integer)
    'dcekama
End Sub

Private Sub dcekama_KeyUp(KeyCode As Integer, _
                          Shift As Integer)

    If KeyCode = 13 Then
        Cmd_Click (0)
    End If

End Sub

Private Sub dcEmpcode_KeyUp(KeyCode As Integer, _
                            Shift As Integer)

    If KeyCode = 13 Then
        Cmd_Click (0)
    End If

End Sub

Private Sub dcjop_KeyUp(KeyCode As Integer, _
                        Shift As Integer)

    If KeyCode = 13 Then
        Cmd_Click (0)
    End If

End Sub

Private Sub dckafel_KeyUp(KeyCode As Integer, _
                          Shift As Integer)

    If KeyCode = 13 Then
        Cmd_Click (0)
    End If

End Sub

Private Sub dcnationality_KeyUp(KeyCode As Integer, _
                                Shift As Integer)

    If KeyCode = 13 Then
        Cmd_Click (0)
    End If

End Sub

Private Sub dcstatus_KeyUp(KeyCode As Integer, _
                           Shift As Integer)

    If KeyCode = 13 Then
        Cmd_Click (0)
    End If

End Sub

Private Sub fg_Click()
    On Error GoTo ErrTrap

    If Not FG.TextMatrix(FG.row, 1) = "" Then
                    If lbltype.Caption = 0 Then
                     'mdifrmmain.ActiveForm.Retrive val(FG.TextMatrix(FG.Row, 1))
                        RetrunFrm.Retrive val(FG.TextMatrix(FG.row, 1))
                    ElseIf lbltype.Caption = 1 Then
RetrunFrm.DCEmployee.BoundText = val(FG.TextMatrix(FG.row, 1))
'                        FrmChangedComponentData.DCEmployee.BoundText = val(Fg.TextMatrix(Fg.Row, 1))
                    ElseIf lbltype.Caption = 2 Then
RetrunFrm.DCEmployee.BoundText = val(FG.TextMatrix(FG.row, 1))
'                        FrmAccountingReport.DCEmployee.BoundText = val(Fg.TextMatrix(Fg.Row, 1))
               ElseIf lbltype.Caption = 3 Then
                         RetrunFrm.DcboEmpName.BoundText = val(FG.TextMatrix(FG.row, 1))
                         'FrmEmpsAdvanceRequest.DcboEmpName.BoundText = val(Fg.TextMatrix(Fg.Row, 1))
               ElseIf lbltype.Caption = 4 Then
                         FrmPassover.DcboEmpName.BoundText = val(FG.TextMatrix(FG.row, 1))
             ElseIf lbltype.Caption = 5 Then
'                         FrmBusinessJob.DcboEmpName.BoundText = val(Fg.TextMatrix(Fg.Row, 1))
RetrunFrm.DcboEmpName.BoundText = val(FG.TextMatrix(FG.row, 1))

              ElseIf lbltype.Caption = 6 Then
            'FrmEmbarkation.DcboEmpName.BoundText = val(Fg.TextMatrix(Fg.Row, 1))
            RetrunFrm.DcboEmpName.BoundText = val(FG.TextMatrix(FG.row, 1))
              ElseIf lbltype.Caption = 7 Then
       '     FrmHolidayData.DcboEmpName.BoundText = val(Fg.TextMatrix(Fg.Row, 1))
            RetrunFrm.DcboEmpName.BoundText = val(FG.TextMatrix(FG.row, 1))
            
            
             ElseIf lbltype.Caption = 8 Then
            'FrmQUesEmp.DcboEmpName.BoundText = val(Fg.TextMatrix(Fg.Row, 1))
             RetrunFrm.DcboEmpName.BoundText = val(FG.TextMatrix(FG.row, 1))
                    ElseIf lbltype.Caption = 9 Then
            'FrmTreament.DcboEmpName.BoundText = val(Fg.TextMatrix(Fg.Row, 1))
            RetrunFrm.DcboEmpName.BoundText = val(FG.TextMatrix(FG.row, 1))
            
                              ElseIf lbltype.Caption = 10 Then
'            FrmRepInjuy.DcboEmpName.BoundText = val(Fg.TextMatrix(Fg.Row, 1))
RetrunFrm.DcboEmpName.BoundText = val(FG.TextMatrix(FG.row, 1))

                             ElseIf lbltype.Caption = 11 Then
            'FrmMovingEmp.DcboEmpName.BoundText = val(Fg.TextMatrix(Fg.Row, 1))
            RetrunFrm.DcboEmpName.BoundText = val(FG.TextMatrix(FG.row, 1))
            
                             ElseIf lbltype.Caption = 12 Then
            'FrmAdvancedHousingpayments.DcboEmpName.BoundText = val(Fg.TextMatrix(Fg.Row, 1))
            RetrunFrm.DcboEmpName.BoundText = val(FG.TextMatrix(FG.row, 1))
                                     ElseIf lbltype.Caption = 13 Then
'            FormEmpMoveDepartment.DcboEmpName.BoundText = val(Fg.TextMatrix(Fg.Row, 1))
RetrunFrm.DcboEmpName.BoundText = val(FG.TextMatrix(FG.row, 1))

   ElseIf lbltype.Caption = 14 Then
'   formvocatinl.DcboEmpName.BoundText = val(Fg.TextMatrix(Fg.Row, 1))
RetrunFrm.DcboEmpName.BoundText = val(FG.TextMatrix(FG.row, 1))

   ElseIf lbltype.Caption = 15 Then
'   frmdriveassest.DcboEmpName.BoundText = val(Fg.TextMatrix(Fg.Row, 1))
RetrunFrm.DcboEmpName.BoundText = val(FG.TextMatrix(FG.row, 1))

   ElseIf lbltype.Caption = 16 Then
 '  FrmPassports.DcboEmpName.BoundText = val(Fg.TextMatrix(Fg.Row, 1))
  RetrunFrm.DcboEmpName.BoundText = val(FG.TextMatrix(FG.row, 1))
                          
   ElseIf lbltype.Caption = 17 Then
   'FRmEmployeeWarning.DcboEmpName.BoundText = val(Fg.TextMatrix(Fg.Row, 1))
   RetrunFrm.DcboEmpName.BoundText = val(FG.TextMatrix(FG.row, 1))
   
           ElseIf lbltype.Caption = 18 Then
   'FrmCars.DCEmployee.BoundText = val(Fg.TextMatrix(Fg.Row, 1))
      RetrunFrm.DCEmployee.BoundText = val(FG.TextMatrix(FG.row, 1))
    
    
               ElseIf lbltype.Caption = 19 Then
   'FrmCars.DCEmployee.BoundText = val(Fg.TextMatrix(Fg.Row, 1))
      FrmOut.DcboEmpName.BoundText = val(FG.TextMatrix(FG.row, 1))
    
    
    
 
                    End If

    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Retrive()
    Dim Num As Integer
    On Error GoTo ErrTrap
    FG.Clear flexClearScrollable, flexClearEverything

    If Not (rs.EOF Or rs.BOF) Then
        FG.rows = rs.RecordCount + 1

        For Num = 1 To rs.RecordCount

            With FG
                .TextMatrix(Num, .ColIndex("NumIndex")) = Num
                
               .TextMatrix(Num, .ColIndex("Contract_ID")) = IIf(IsNull(rs("Contract_ID").value), "", (rs("Contract_ID").value))
               
                .TextMatrix(Num, .ColIndex("ProfID")) = IIf(IsNull(rs("Emp_ID").value), 0, (rs("Emp_ID").value))
                .TextMatrix(Num, .ColIndex("ProfCode")) = IIf(IsNull(rs("Emp_Code").value), "", (rs("Emp_Code").value))
                .TextMatrix(Num, .ColIndex("ProfNme")) = IIf(IsNull(rs("Emp_Name").value), "", Trim(rs("Emp_Name").value))
                .TextMatrix(Num, .ColIndex("ProfPhone")) = IIf(IsNull(rs("Emp_Phone").value), "", Trim(rs("Emp_Phone").value))
                .TextMatrix(Num, .ColIndex("JobTypeName")) = IIf(IsNull(rs("JobTypeName").value), "", Trim(rs("JobTypeName").value))
                .TextMatrix(Num, .ColIndex("DepartmentName")) = IIf(IsNull(rs("DepartmentName").value), "", Trim(rs("DepartmentName").value))
                .TextMatrix(Num, .ColIndex("dean")) = IIf(IsNull(rs("dean").value), "", Trim(rs("dean").value))
                .TextMatrix(Num, .ColIndex("nationality")) = IIf(IsNull(rs("nationality").value), "", Trim(rs("nationality").value))
                .TextMatrix(Num, .ColIndex("name")) = IIf(IsNull(rs("name").value), "", Trim(rs("name").value))
            .TextMatrix(Num, .ColIndex("NumEkama")) = IIf(IsNull(rs("NumEkama").value), "", Trim(rs("NumEkama").value))
            .TextMatrix(Num, .ColIndex("Fullcode")) = IIf(IsNull(rs("Fullcode").value), "", Trim(rs("Fullcode").value))
            
                .TextMatrix(Num, .ColIndex("kafelid")) = IIf(IsNull(rs("kafelid").value), "", Trim(rs("kafelid").value))
            
                .TextMatrix(Num, .ColIndex("kafelname")) = IIf(IsNull(rs("kafelname").value), "", Trim(rs("kafelname").value))
            
                .Cell(flexcpBackColor, Num, 1, Num, 1) = IIf(IsNull(rs.Fields("color").value), "", rs.Fields("color").value)
            
                .Cell(flexcpBackColor, Num, 2, Num, 2) = IIf(IsNull(rs.Fields("color").value), "", rs.Fields("color").value)
            
                .Cell(flexcpBackColor, Num, 3, Num, 3) = IIf(IsNull(rs.Fields("color").value), "", rs.Fields("color").value)
                .Cell(flexcpBackColor, Num, 4, Num, 4) = IIf(IsNull(rs.Fields("color").value), "", rs.Fields("color").value)
                .Cell(flexcpBackColor, Num, 5, Num, 5) = IIf(IsNull(rs.Fields("color").value), "", rs.Fields("color").value)
                .Cell(flexcpBackColor, Num, 6, Num, 6) = IIf(IsNull(rs.Fields("color").value), "", rs.Fields("color").value)
                .Cell(flexcpBackColor, Num, 7, Num, 7) = IIf(IsNull(rs.Fields("color").value), "", rs.Fields("color").value)
                .Cell(flexcpBackColor, Num, 8, Num, 8) = IIf(IsNull(rs.Fields("color").value), "", rs.Fields("color").value)
                .Cell(flexcpBackColor, Num, 9, Num, 9) = IIf(IsNull(rs.Fields("color").value), "", rs.Fields("color").value)
            
                .Cell(flexcpBackColor, Num, 10, Num, 10) = IIf(IsNull(rs.Fields("color").value), "", rs.Fields("color").value)
            
                .Cell(flexcpBackColor, Num, 11, Num, 11) = IIf(IsNull(rs.Fields("color").value), "", rs.Fields("color").value)
            
            End With

            rs.MoveNext
        Next Num

        FG.AutoSize 0, FG.Cols - 1, False
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Fg_DblClick()
    fg_Click
    Unload Me
End Sub

Private Sub Fg_KeyDown(KeyCode As Integer, _
                       Shift As Integer)

    If KeyCode = vbKeyReturn Then
        fg_Click
    End If

End Sub

Private Sub Form_Load()
    Dim StrSQL As String
    Dim BG As New ClsBackGroundPic
    Dim Dcombos As New ClsDataCombos

    On Error GoTo ErrTrap
    Set rs = New ADODB.Recordset

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Clear").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
  '  Dcombos.GetEmployees Me.DCboEmployeesName
    'Set cSearchDcbo = New clsDCboSearch
    'Set cSearchDcbo.Client = Me.DCboEmployeesName

    CenterForm Me

    FormPostion Me, GetPostion
    FG.WallPaper = BG.SearchWallpaper

    Dim My_SQL As String

    My_SQL = "  select JobTypeID,JobTypeName from TblEmpJobsTypes   "
    fill_combo dcjop, My_SQL

    My_SQL = "  select DeparmentID,DepartmentName from TblEmpDepartments   "
    fill_combo Dcdep, My_SQL

    My_SQL = "  select id,name from Nationality   "
    fill_combo DCNationality, My_SQL

    My_SQL = "  select id,name from dean   "
    fill_combo Dcdean, My_SQL

    My_SQL = "  select id,name from jopstatus   "
    fill_combo dcstatus, My_SQL

    My_SQL = "  select distinct kafelid,kafelname from    emp_all_details "
    fill_combo dckafel, My_SQL

    My_SQL = "  select     emp_name,fullcode from    TblEmployee order by fullcode "
    fill_combo dcEmpcode, My_SQL

    My_SQL = "  select  distinct  emp_name,NumEkama from    emp_all_details "
    fill_combo dcekama, My_SQL

    Exit Sub
ErrTrap:

End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap

    If rs.State = adStateOpen Then
        rs.Close
        Set rs = Nothing
    End If

    Set cSearchDcbo = Nothing

    FormPostion Me, SavePostion
    Exit Sub
ErrTrap:
End Sub

Private Function Build_Sql()
    Dim StrSQL As String
    Dim Begin As Boolean
    Dim StrWhere As String
    'On Error GoTo ErrTrap
   ' StrSQL = "select * From emp_all_details"
StrSQL = "SELECT     dbo.TblEmployee.Fullcode, dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmpDepartments.DepartmentName, dbo.jopstatus.color, dbo.jopstatus.name, "
StrSQL = StrSQL & "   dbo.TblEmployee.Emp_ID, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Name1, dbo.TblEmployee.Emp_Name2,"
StrSQL = StrSQL & "   dbo.TblEmployee.Emp_Name3, dbo.TblEmployee.Emp_Name4, dbo.TblEmployee.Emp_Mail, dbo.TblEmployee.Emp_Phone, dbo.TblEmployee.Emp_mobile,"
StrSQL = StrSQL & "   dbo.TblEmployee.Emp_Remark, dbo.TblEmployee.Emp_Salary, dbo.TblEmployee.Emp_Comm, dbo.TblEmployee.EmpProfitCom, dbo.TblEmployee.workstate,"
StrSQL = StrSQL & "   dbo.TblEmployee.DepartmentID, dbo.TblEmployee.JobTypeID, dbo.TblEmployee.SpecificationID, dbo.TblEmployee.Region, dbo.TblEmployee.InsuranceState,"
StrSQL = StrSQL & "   dbo.TblEmployee.InsuranceValue, dbo.TblEmployee.OtherDiscounts, dbo.TblEmployee.placeEkama, dbo.TblEmployee.NumEkama, dbo.TblEmployee.DateExpoekama,"
StrSQL = StrSQL & "   dbo.TblEmployee.DateEndekama, dbo.TblEmployee.DateExpoekamaH, dbo.TblEmployee.DateEndekamah, dbo.TblEmployee.NumLicn, dbo.TblEmployee.DateExpLinc,"
StrSQL = StrSQL & "   dbo.TblEmployee.DateEndLinc, dbo.TblEmployee.DateExpLincH, dbo.TblEmployee.DateEndLincH, dbo.TblEmployee.NumPoket, dbo.TblEmployee.Dateexppoket,"
StrSQL = StrSQL & "   dbo.TblEmployee.dateendpoket, dbo.TblEmployee.NumPasp, dbo.TblEmployee.DateEndPasp, dbo.TblEmployee.DateExpPasp, dbo.TblEmployee.EmpNum,"
StrSQL = StrSQL & "   dbo.TblEmployee.CustNum, dbo.TblEmployee.ChekEndWork, dbo.TblEmployee.ChekStkala, dbo.TblEmployee.BignDateWork, dbo.TblEmployee.EndWork,"
StrSQL = StrSQL & "   dbo.TblEmployee.Notsstkala, dbo.TblEmployee.checkbox1, dbo.TblEmployee.DOB, dbo.TblEmployee.KafelID, dbo.TblEmployee.KafelName,"
StrSQL = StrSQL & "   dbo.TblEmployee.pasplace, dbo.TblEmployee.Nationality, dbo.TblEmployee.dean, dbo.TblEmployee.hdodno, dbo.TblEmployee.hdoddate, dbo.TblEmployee.hdomnfaz,"
StrSQL = StrSQL & "   dbo.TblEmployee.kafeltel, dbo.TblEmployee.jopstatusid, dbo.TblEmployee.kafeladd, dbo.TblEmployee.Emp_Salary_sakn, dbo.TblEmployee.Emp_Salary_bus,"
StrSQL = StrSQL & "   dbo.TblEmployee.Emp_Salary_food, dbo.TblEmployee.Emp_Salary_mob, dbo.TblEmployee.Emp_Salary_mang, dbo.TblEmployee.Emp_Salary_others,"
StrSQL = StrSQL & "   dbo.TblEmployee.Emp_Salary_sakn1, dbo.TblEmployee.Emp_Salary_bus1, dbo.TblEmployee.Emp_Salary_food1, dbo.TblEmployee.Emp_Salary_others1,"
StrSQL = StrSQL & "   dbo.TblEmployee.Emp_Salary_mob1, dbo.TblEmployee.Emp_Salary_mang1, dbo.TblEmployee.Account_code, dbo.TblEmployee.Account_code1,"
StrSQL = StrSQL & "   dbo.TblEmployee.Emp_Salary_saknc, dbo.TblEmployee.Emp_Salary_busc, dbo.TblEmployee.Emp_Salary_foodc, dbo.TblEmployee.Emp_Salary_othersc,"
StrSQL = StrSQL & "   dbo.TblEmployee.Emp_Salary_mobc, dbo.TblEmployee.Emp_Salary_mangc, dbo.TblEmployee.Emp_Salary_saknc1, dbo.TblEmployee.Emp_Salary_busc1,"
StrSQL = StrSQL & "   dbo.TblEmployee.Emp_Salary_foodc1, dbo.TblEmployee.Emp_Salary_othersc1, dbo.TblEmployee.Emp_Salary_mobc1, dbo.TblEmployee.Emp_Salary_mangc1,"
StrSQL = StrSQL & "   dbo.TblEmployee.ItemPhoto, dbo.TblEmployee.placeWORK, dbo.TblEmployee.project_id, dbo.TblEmployee.Account_Code2, dbo.TblEmployee.Dateexppoketh,"
StrSQL = StrSQL & "   dbo.TblEmployee.dateendpoketh, dbo.TblEmployee.opr_fullcode, dbo.TblEmployee.term_id, dbo.TblEmployee.opr_id, dbo.TblEmployee.term_fullcode,"
StrSQL = StrSQL & "   dbo.TblEmployee.Fullcode AS Expr1, dbo.TblEmployee.prifix, dbo.TblEmployee.BranchId, dbo.TblEmployee.cost_center_id, dbo.TblEmployee.Emp_Namee1,"
StrSQL = StrSQL & "   dbo.TblEmployee.Emp_Namee, dbo.TblEmployee.Emp_Namee3, dbo.TblEmployee.Emp_Namee4, dbo.TblEmployee.Emp_Namee2, dbo.TblEmployee.VisaNo,"
StrSQL = StrSQL & "   dbo.TblEmployee.GroupID, dbo.TblEmployee.mangerid, dbo.TblEmployee.swapedempid, dbo.EmpGroupDep.GroupName,"
StrSQL = StrSQL & "   TblEmployee_1.Emp_Name AS Mangername, TblEmployee_1.Emp_Code AS mangerCode, TblEmployee_1.Emp_Namee AS MangerNamee,"
StrSQL = StrSQL & "   dbo.Contract.Contract_ID,"

StrSQL = StrSQL & "                      dbo.GetEmployeeSalary(dbo.TblEmployee.Emp_ID, GETDATE()) AS Salary"

StrSQL = StrSQL & "   FROM         dbo.jopstatus RIGHT OUTER JOIN"
StrSQL = StrSQL & "   dbo.TblEmployee RIGHT OUTER JOIN"
StrSQL = StrSQL & "   dbo.Contract ON dbo.TblEmployee.Emp_ID = dbo.Contract.Emp_id LEFT OUTER JOIN"
StrSQL = StrSQL & "    dbo.TblEmployee TblEmployee_1 ON dbo.TblEmployee.mangerid = TblEmployee_1.Emp_ID LEFT OUTER JOIN"
StrSQL = StrSQL & "    dbo.EmpGroupDep ON dbo.TblEmployee.GroupID = dbo.EmpGroupDep.GroupID ON dbo.jopstatus.id = dbo.TblEmployee.jopstatusid LEFT OUTER JOIN"
StrSQL = StrSQL & "    dbo.TblEmpDepartments ON dbo.TblEmployee.DepartmentID = dbo.TblEmpDepartments.DeparmentID LEFT OUTER JOIN"
StrSQL = StrSQL & "      dbo.TblEmpJobsTypes ON dbo.TblEmployee.JobTypeID = dbo.TblEmpJobsTypes.JobTypeID"
StrSQL = StrSQL & "   Where (1 = 1)"
 



    If Contract_ID.text <> "" Then
  
              StrWhere = StrWhere + " and Contract.Contract_ID like'%" & (Contract_ID.text) & "%'"
         
    End If
    
    
    If dcEmpcode.text <> "" Then
  
              StrWhere = StrWhere + " and TblEmployee.Fullcode like'%" & (dcEmpcode.text) & "%'"
         
    End If

    If dcekama.text <> "" Then
 
            StrWhere = StrWhere + " and  TblEmployee.NumEkama='" & val(dcekama.text) & "'"
         
    End If

    'DCboEmployeesName   XPTxtEmpName
    If DCboEmployeesName.text <> "" Then
        
                StrWhere = StrWhere + " and  TblEmployee.Emp_Name like'%" & Trim(DCboEmployeesName.text) & "%'"
       
    End If

    If dcjop.BoundText <> "" Then
 
            StrWhere = StrWhere + " and TblEmployee.JobTypeID =" & Trim(dcjop.BoundText)
       
    
   
    End If

    If Dcdep.BoundText <> "" Then
 
            StrWhere = StrWhere + " and TblEmployee.DepartmentID =" & Trim(Dcdep.BoundText)
   
   
    End If

    If DCNationality.BoundText <> "" Then
  
            StrWhere = StrWhere + " and TblEmployee.nationality ='" & Trim(DCNationality.text) & "'"
 
   
    End If

    If Dcdean.BoundText <> "" Then
    
    
            StrWhere = StrWhere + " and TblEmployee.dean ='" & Trim(Dcdean.text) & "'"
 
   
    End If

    If dcstatus.BoundText <> "" Then
     
            StrWhere = StrWhere + " and TblEmployee.jopstatusid =" & dcstatus.BoundText
 
   
    End If

    If dckafel.text <> "" Then
 
            StrWhere = StrWhere + " and TblEmployee.kafelname  like'%" & dckafel.text & "%'"
    
   
    End If

    StrSQL = StrSQL + StrWhere
    'StrSQL = StrSQL & " ORDER BY CAST(Emp_Code AS integer) ASC "
   StrSQL = StrSQL & " ORDER BY TblEmployee.fullcode ASC "
    
    Build_Sql = StrSQL

    Exit Function
ErrTrap:
End Function

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        If Not FG.TextMatrix(FG.row, FG.ColIndex("ProfCode")) = "" Then
            fg_Click
        Else
            Cmd_Click (0)
        End If
    End If

    If Shift = 2 Then
        If KeyCode = vbKeyX Then
            Cmd_Click (2)
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub ChangeLang()
    Me.Caption = "Search for Employee"

    XPLbl(1).Caption = "Employee Code"
    XPLbl(0).Caption = "Ekama No"
    XPChkSearchType.Caption = "Math Complete Name"
    Cmd(0).Caption = "Search"
    Cmd(1).Caption = "Clear"
    Cmd(2).Caption = "Exit"
    XPLbl(2).Caption = "Departement"
    XPLbl(6).Caption = "Status"
    XPLbl(9).Caption = "Name"
    XPLbl(3).Caption = "Nationality"
    XPLbl(7).Caption = "sponsor"
    XPLbl(9).Caption = "Name"
    XPLbl(4).Caption = "Job"
    XPLbl(5).Caption = "Religon"
    Cmd1.Caption = "Print"
    ALLButton1.Caption = "Color Map"

    With Me.FG
        .TextMatrix(0, .ColIndex("NumIndex")) = "Serial"
        .TextMatrix(0, .ColIndex("ProfCode")) = "Employee Code"
        .TextMatrix(0, .ColIndex("ProfNme")) = "Employee Name"
        .TextMatrix(0, .ColIndex("ProfPhone")) = "Employee Phone"

        .TextMatrix(0, .ColIndex("ProfID")) = "Employee N0"
        .TextMatrix(0, .ColIndex("JobTypeName")) = " JobTypeName "
        .TextMatrix(0, .ColIndex("DepartmentName")) = " DepartmentName "
        .TextMatrix(0, .ColIndex("nationality")) = " Nationality"
        .TextMatrix(0, .ColIndex("dean")) = "  Religon"
        .TextMatrix(0, .ColIndex("name")) = " name "
        .TextMatrix(0, .ColIndex("kafelid")) = " Sponsor No "
        .TextMatrix(0, .ColIndex("kafelname")) = " Sponsor Name  "
     
    End With

End Sub

Private Sub XPTxtEmpID_Change()

End Sub

