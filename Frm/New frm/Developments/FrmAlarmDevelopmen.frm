VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmAlarmDevelopmen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "  ‰»ÌÂ«  «·„Â«„  Ê «·⁄„·Ì«              "
   ClientHeight    =   7080
   ClientLeft      =   4125
   ClientTop       =   3525
   ClientWidth     =   14550
   Icon            =   "FrmAlarmDevelopmen.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   14550
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.Frame Frame3 
      BackColor       =   &H00E2E9E9&
      Height          =   855
      Left            =   0
      TabIndex        =   7
      Top             =   6240
      Width           =   14535
      Begin ImpulseButton.ISButton Cmd 
         Height          =   495
         Index           =   6
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   3045
         _ExtentX        =   5371
         _ExtentY        =   873
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
         ButtonImage     =   "FrmAlarmDevelopmen.frx":6852
         ColorButton     =   14871017
         ColorHighlight  =   16777215
         ColorHoverText  =   16711680
         ColorShadow     =   -2147483637
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         ColorToggledHoverText=   16711680
         ColorTextShadow =   -2147483637
      End
      Begin ImpulseButton.ISButton CmdHelp 
         Height          =   495
         Left            =   3360
         TabIndex        =   9
         Top             =   240
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   873
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
         ButtonImage     =   "FrmAlarmDevelopmen.frx":30474
         ColorButton     =   14871017
         ColorHighlight  =   16777215
         ColorHoverText  =   16711680
         ColorShadow     =   -2147483637
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         ColorToggledHoverText=   16711680
         ColorTextShadow =   -2147483637
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   495
         Index           =   9
         Left            =   6480
         TabIndex        =   10
         Top             =   240
         Width           =   3045
         _ExtentX        =   5371
         _ExtentY        =   873
         ButtonPositionImage=   1
         Caption         =   "ÿ»«⁄Â"
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
         ButtonImage     =   "FrmAlarmDevelopmen.frx":36CD6
         ColorButton     =   14871017
         ColorHighlight  =   16777215
         ColorHoverText  =   16711680
         ColorShadow     =   -2147483637
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         ColorToggledHoverText=   16711680
         ColorTextShadow =   -2147483637
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   495
         Index           =   0
         Left            =   10080
         TabIndex        =   22
         Top             =   240
         Visible         =   0   'False
         Width           =   3045
         _ExtentX        =   5371
         _ExtentY        =   873
         ButtonPositionImage=   1
         Caption         =   "Õ›Ÿ"
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
         ButtonImage     =   "FrmAlarmDevelopmen.frx":3D538
         ColorButton     =   14871017
         ColorHighlight  =   16777215
         ColorHoverText  =   16711680
         ColorShadow     =   -2147483637
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         ColorToggledHoverText=   16711680
         ColorTextShadow =   -2147483637
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E2E9E9&
      Height          =   4575
      Left            =   0
      TabIndex        =   3
      Top             =   1680
      Width           =   14535
      Begin VSFlex8UCtl.VSFlexGrid GridInstallments 
         Height          =   4215
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   14355
         _cx             =   25321
         _cy             =   7435
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
         BackColorAlternate=   16777152
         GridColor       =   0
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
         FormatString    =   $"FrmAlarmDevelopmen.frx":3D88A
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
         Begin MSComctlLib.ProgressBar ProgressBar1 
            Height          =   615
            Left            =   1920
            TabIndex        =   5
            Top             =   1440
            Visible         =   0   'False
            Width           =   8415
            _ExtentX        =   14843
            _ExtentY        =   1085
            _Version        =   393216
            Appearance      =   0
         End
         Begin VB.Label Label2 
            Caption         =   "%"
            Height          =   375
            Index           =   0
            Left            =   10440
            TabIndex        =   6
            Top             =   -600
            Width           =   375
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2E9E9&
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   14535
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2640
         TabIndex        =   20
         Top             =   480
         Width           =   810
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·› —Â"
         Height          =   735
         Left            =   3600
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   120
         Width           =   5415
         Begin MSComCtl2.DTPicker Fromdate 
            Height          =   330
            Left            =   3000
            TabIndex        =   15
            Top             =   240
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   157483009
            CurrentDate     =   41640
         End
         Begin MSComCtl2.DTPicker todate 
            Height          =   330
            Left            =   360
            TabIndex        =   16
            Top             =   240
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   157483009
            CurrentDate     =   41640
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "≈«·Ï"
            Height          =   435
            Index           =   14
            Left            =   2100
            RightToLeft     =   -1  'True
            TabIndex        =   18
            Top             =   240
            Width           =   540
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "„‰"
            Height          =   315
            Index           =   0
            Left            =   4680
            RightToLeft     =   -1  'True
            TabIndex        =   17
            Top             =   240
            Width           =   585
         End
      End
      Begin VB.TextBox TxtSearchCode 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   12870
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   360
         Width           =   825
      End
      Begin MSDataListLib.DataCombo DcboEmpName 
         Bindings        =   "FrmAlarmDevelopmen.frx":3DA11
         Height          =   315
         Left            =   9120
         TabIndex        =   12
         Top             =   360
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         BackColor       =   16777215
         ListField       =   "account_name"
         BoundColumn     =   "code"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   735
         Index           =   5
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   1296
         ButtonPositionImage=   1
         Caption         =   "»ÕÀ"
         BackColor       =   14871017
         FontSize        =   14.25
         FontName        =   "Times New Roman"
         FontBold        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmAlarmDevelopmen.frx":3DA26
         ColorButton     =   14871017
         ColorHighlight  =   16777215
         ColorHoverText  =   12632064
         ColorShadow     =   -2147483637
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         ColorToggledHoverText=   12632064
         LowerToggledContent=   0   'False
         ColorTextShadow =   -2147483637
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   " ÕœÌÀ ﬂ·"
         Height          =   435
         Index           =   4
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   240
         Width           =   780
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·„ÊŸ›"
         Height          =   285
         Index           =   10
         Left            =   13440
         TabIndex        =   13
         Top             =   360
         Width           =   1245
      End
   End
   Begin C1SizerLibCtl.C1Elastic EleHeader 
      Height          =   585
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   14565
      _cx             =   25691
      _cy             =   1032
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   4
      MousePointer    =   0
      Version         =   801
      BackColor       =   16777215
      ForeColor       =   4210688
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   "  ‰»ÌÂ«  «·„Â«„  Ê «·⁄„·Ì«              "
      Align           =   0
      AutoSizeChildren=   0
      BorderWidth     =   0
      ChildSpacing    =   0
      Splitter        =   0   'False
      FloodDirection  =   0
      FloodPercent    =   0
      CaptionPos      =   6
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
      Begin VB.Image ImgFavorites 
         Height          =   390
         Left            =   5400
         Picture         =   "FrmAlarmDevelopmen.frx":44288
         Stretch         =   -1  'True
         Top             =   120
         Width           =   525
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H000000FF&
         Height          =   555
         Index           =   27
         Left            =   2520
         TabIndex        =   2
         Top             =   0
         Width           =   2205
      End
   End
End
Attribute VB_Name = "FrmAlarmDevelopmen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim My_SQL As String
Private Sub Cmd_Click(Index As Integer)
Select Case Index
Case 5
If val(DcboEmpName.BoundText) = 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "Ì—ÃÏ «Œ Ì«— «·„ÊŸ›"
Else
MsgBox "Please Select Employee"
End If
Exit Sub
End If
ProgressBar1.Visible = True
: ProgressBar1.value = 10
fillgrid
: ProgressBar1.value = 50
ProgressBar1.Visible = False
ProgressBar1.value = 0
Case 6
Me.Hide
Case 9
print_report My_SQL
End Select

End Sub

Function print_report(Optional NoteSerial As String)
     
  If NoteSerial = "" Then
  Exit Function
  End If
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String

 If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAlertDevelopment.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAlertDevelopmentE.rpt"
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
            Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        Else
            Msg = "There's no data to show"
        End If
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
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
        'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        '    StrReportTitle = StrReportTitle + " »œ«Ì… „‰ " & Format(Me.DTPickerAccFrom.value, "yyyy/M/d") & ""
        'End If
        'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        '    StrReportTitle = StrReportTitle + " ≈·Ï " & Format(Me.DTPickerAccTo.value, "yyyy/M/d") & " "
        'End If
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        'xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""
        'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        '    StrReportTitle = StrReportTitle + " From Date " & (Me.DTPickerAccFrom.value) & ""
        'End If
        'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        '    StrReportTitle = StrReportTitle + " To Date :  " & (Me.DTPickerAccTo.value) & ""
        'End If
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
   ' xReport.ParameterFields(13).AddCurrentValue Me.DTPicker1.value
      '  xReport.ParameterFields(4).AddCurrentValue WriteNo(Format(val(TxtAdvanceValue.text), "0.00"), 0, True, ".")
       ' xReport.ParameterFields(6).AddCurrentValue val(lbl(23).Caption)
        ' xReport.ParameterFields(13).AddCurrentValue Me.DTPicker1.value
'    xReport.ParameterFields(8).AddCurrentValue IIf(IsNumeric(fg.TextMatrix(Me.fg.FixedRows, fg.ColIndex("PartValue"))), val(fg.TextMatrix(Me.fg.FixedRows, fg.ColIndex("PartValue"))), 0)
' xReport.ParameterFields(9).AddCurrentValue val(lbl(22).Caption)
 ' xReport.ParameterFields(10).AddCurrentValue val(TxtDiscount.text)
  ' xReport.ParameterFields(11).AddCurrentValue txtDiscountDES.text
  If FromDate.value <> "" And ToDate.value <> "" Then
   xReport.ParameterFields(8).AddCurrentValue FromDate.value

    xReport.ParameterFields(9).AddCurrentValue ToDate.value
  
    End If
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.title
    xReport.ReportAuthor = App.title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault


 
  
 
End Function

Private Sub CmdHelp_Click()
          clear_all Me
            GridInstallments.Clear flexClearScrollable, flexClearEverything
            GridInstallments.Rows = 1
ToDate.value = Date
FromDate.value = Date

End Sub


Function cahngelang()
    EleHeader.Caption = " Task Follow-up Alerts "
    Me.Caption = EleHeader.Caption
    lbl(10).Caption = "Employee"
  
    Frame5.Caption = "Period"
    lbl(0).Caption = "From"
    lbl(14).Caption = "To"
   Cmd(0).Caption = "Save"
    lbl(4).Caption = "Update All"
   CmdHelp.Caption = "Clear"
   
   Cmd(5).Caption = "Saerch"
   Cmd(9).Caption = "Print"
    Cmd(6).Caption = "Exit"
    With GridInstallments
    .TextMatrix(0, .ColIndex("Ser")) = "Serial"
    .TextMatrix(0, .ColIndex("name")) = "Task"
    .TextMatrix(0, .ColIndex("FromDate")) = "Start Date"
    .TextMatrix(0, .ColIndex("Uppdatee")) = "Last Edit"
    .TextMatrix(0, .ColIndex("DesOp")) = "Employee Comments"
    .TextMatrix(0, .ColIndex("AnlysOp")) = "Management Comments"
    .TextMatrix(0, .ColIndex("Endd")) = "End Tasks"
    End With
End Function

Public Sub fillgrid(Optional str As String)
Dim cont1 As Double
Dim cont As Double
Dim typ As Integer
My_SQL = ""
  '  On Error GoTo ErrTrap
On Error Resume Next
    Dim i As Integer
    Dim rs As ADODB.Recordset

    Set rs = New ADODB.Recordset
 'My_SQL = "SELECT     dbo.TblRegDevelopment.Id, dbo.TblRegDevelopment.RecordDate, dbo.TblRegDevelopment.StrDate, dbo.TblRegDevelopment.EndExptedDate, "
 'My_SQL = My_SQL & "                      dbo.TblRegDevelopment.EndActDate, dbo.TblRegDevelopment.BranchID, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee,"
 'My_SQL = My_SQL & "                     dbo.TblRegDevelopment.EmpID, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee, dbo.TblRegDevelopment.Important,"
 'My_SQL = My_SQL & "                      dbo.TblRegDevelopment.MoDay, dbo.TblRegDevelopment.MangID, TblEmployee_1.Emp_Name AS MangEmp_Name, TblEmployee_1.Fullcode AS MangFullcode,"
 'My_SQL = My_SQL & "                     TblEmployee_1.Emp_Namee AS MangEmp_NameE, dbo.TblRegDevelopment.DesOp, dbo.TblRegDevelopment.AnlysOp, dbo.TblRegDevelopment.TimeReq,"
 'My_SQL = My_SQL & "                      dbo.TblRegDevelopment.StartTime, dbo.TblRegDevelopment.EndExptedTime, dbo.TblRegDevelopment.EndActTIme, dbo.TblRegDevelopment.NoDaySatart,"
 'My_SQL = My_SQL & "                     dbo.TblRegDevelopment.NoDayEnd, dbo.TblRegDevelopment.FromDate, dbo.TblRegDevelopment.ToDate, dbo.TblRegDevelopment.RecordTime,"
 'My_SQL = My_SQL & "                     dbo.TblRegDevelopment.OpType, dbo.TblProceeDevelper.Name, dbo.TblProceeDevelper.NameE, dbo.TblRegDevelopment.StatusProcess,"
 'My_SQL = My_SQL & "                     dbo.TblRegDevelopment.DesID , dbo.TblProceeDevelperDet.des, dbo.TblRegDevelopment.StatusPand,dbo.GetLastUpdateDevelopment(dbo.TblRegDevelopment.OpType,dbo.TblRegDevelopment.MangID)as LastUpdate"
 'My_SQL = My_SQL & "  FROM         dbo.TblRegDevelopment LEFT OUTER JOIN"
 'My_SQL = My_SQL & "                     dbo.TblProceeDevelperDet ON dbo.TblRegDevelopment.DesID = dbo.TblProceeDevelperDet.ID LEFT OUTER JOIN"
 'My_SQL = My_SQL & "                     dbo.TblProceeDevelper ON dbo.TblRegDevelopment.OpType = dbo.TblProceeDevelper.ID LEFT OUTER JOIN"
 'My_SQL = My_SQL & "                     dbo.TblEmployee TblEmployee_1 ON dbo.TblRegDevelopment.MangID = TblEmployee_1.Emp_ID LEFT OUTER JOIN"
' My_SQL = My_SQL & "                     dbo.TblEmployee ON dbo.TblRegDevelopment.EmpID = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
'My_SQL = My_SQL & "                    dbo.TblBranchesData ON dbo.TblRegDevelopment.BranchID = dbo.TblBranchesData.branch_id"
My_SQL = "SELECT     TOP 100 PERCENT dbo.TblRegDevelopment.Id, dbo.TblRegDevelopment.RecordDate, dbo.TblRegDevelopment.StrDate, dbo.TblRegDevelopment.EndExptedDate, "
My_SQL = My_SQL & "                            dbo.TblRegDevelopment.EndActDate, dbo.TblRegDevelopment.BranchID, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee,"
My_SQL = My_SQL & "                            dbo.TblRegDevelopment.EmpID, TblEmployee_2.Emp_Name, TblEmployee_2.Fullcode, TblEmployee_2.Emp_Namee, dbo.TblRegDevelopment.Important,"
My_SQL = My_SQL & "                            dbo.TblRegDevelopment.MoDay, dbo.TblRegDevelopment.MangID, TblEmployee_1.Emp_Name AS MangEmp_Name, TblEmployee_1.Fullcode AS MangFullcode,"
My_SQL = My_SQL & "                            TblEmployee_1.Emp_Namee AS MangEmp_NameE, dbo.TblRegDevelopment.DesOp, dbo.TblRegDevelopment.AnlysOp, dbo.TblRegDevelopment.TimeReq,"
My_SQL = My_SQL & "                            dbo.TblRegDevelopment.StartTime, dbo.TblRegDevelopment.EndExptedTime, dbo.TblRegDevelopment.EndActTIme, dbo.TblRegDevelopment.NoDaySatart,"
My_SQL = My_SQL & "                            dbo.TblRegDevelopment.NoDayEnd, dbo.TblRegDevelopment.FromDate, dbo.TblRegDevelopment.ToDate, dbo.TblRegDevelopment.RecordTime,"
My_SQL = My_SQL & "                            dbo.TblRegDevelopment.OpType, dbo.TblProceeDevelper.Name, dbo.TblProceeDevelper.NameE, dbo.TblRegDevelopment.StatusProcess,"
My_SQL = My_SQL & "                            dbo.TblRegDevelopment.DesID, dbo.TblProceeDevelperDet.Des, dbo.TblRegDevelopment.StatusPand,"
My_SQL = My_SQL & "                            dbo.GetLastUpdateDevelopment(dbo.TblRegDevelopment.OpType, dbo.TblRegDevelopment.MangID) AS LastUpdate"
My_SQL = My_SQL & "      FROM         dbo.TblRegDevelopment LEFT OUTER JOIN"
My_SQL = My_SQL & "                            dbo.TblProceeDevelperDet ON dbo.TblRegDevelopment.DesID = dbo.TblProceeDevelperDet.ID LEFT OUTER JOIN"
My_SQL = My_SQL & "                            dbo.TblProceeDevelper ON dbo.TblRegDevelopment.OpType = dbo.TblProceeDevelper.ID LEFT OUTER JOIN"
My_SQL = My_SQL & "                            dbo.TblEmployee TblEmployee_1 ON dbo.TblRegDevelopment.MangID = TblEmployee_1.Emp_ID LEFT OUTER JOIN"
My_SQL = My_SQL & "                            dbo.TblEmployee TblEmployee_2 ON dbo.TblRegDevelopment.EmpID = TblEmployee_2.Emp_ID LEFT OUTER JOIN"
My_SQL = My_SQL & "                            dbo.TblBranchesData ON dbo.TblRegDevelopment.BranchID = dbo.TblBranchesData.branch_id"

My_SQL = My_SQL & " where 1=1"

  If Not (IsNull(Me.FromDate.value)) Then
 My_SQL = My_SQL + " and (dbo.TblRegDevelopment.RecordDate >='" & SQLDate(FromDate.value) & "')"
 End If
 If Not (IsNull(Me.ToDate.value)) Then
 My_SQL = My_SQL + " and (dbo.TblRegDevelopment.RecordDate <='" & SQLDate(ToDate.value) & "')"
 End If
'
If Me.DcboEmpName.Text <> "" And val(Me.DcboEmpName.BoundText) <> 0 Then
My_SQL = My_SQL + "and TblRegDevelopment.empid =" & val(Me.DcboEmpName.BoundText) & ""
End If


My_SQL = My_SQL + "   order by  dbo.TblRegDevelopment.ID "
  
Dim ActualTotal As Double
'rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText
    rs.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
      With Me.GridInstallments
       .Rows = 1
        .Clear flexClearScrollable

        If rs.RecordCount > 0 Then
           .Rows = rs.RecordCount + 1
           rs.MoveFirst

            For i = 1 To .Rows - 1
              .TextMatrix(i, .ColIndex("Ser")) = i
              .TextMatrix(i, .ColIndex("OpType")) = (IIf(IsNull(rs.Fields("OpType").value), 0, rs.Fields("OpType").value))
              .TextMatrix(i, .ColIndex("MangID")) = (IIf(IsNull(rs.Fields("MangID").value), 0, rs.Fields("MangID").value))
              .TextMatrix(i, .ColIndex("ID")) = (IIf(IsNull(rs.Fields("ID").value), "", rs.Fields("ID").value))
              .TextMatrix(i, .ColIndex("FromDate")) = (IIf(IsNull(rs.Fields("FromDate").value), "", rs.Fields("FromDate").value))
              .TextMatrix(i, .ColIndex("AnlysOp")) = (IIf(IsNull(rs.Fields("AnlysOp").value), "", rs.Fields("AnlysOp").value))
              .TextMatrix(i, .ColIndex("DesOp")) = IIf(IsNull(rs("DesOp")), "", (rs("DesOp").value))
              If Not (IsNull(rs.Fields("StatusProcess").value)) Then
              If rs.Fields("StatusProcess").value = 2 Then
              .TextMatrix(i, .ColIndex("Endd")) = True
              Else
              .TextMatrix(i, .ColIndex("Endd")) = False
              End If
              End If
              .TextMatrix(i, .ColIndex("Uppdatee")) = IIf(IsNull(rs("LastUpdate")), "", (CDate(rs("LastUpdate").value)))
              If SystemOptions.UserInterface = ArabicInterface Then
              .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(rs("Name")), "", (rs("Name").value))
              Else
              .TextMatrix(i, .ColIndex("name")) = IIf(IsNull(rs("NameE")), "", (rs("NameE").value))
              End If
        rs.MoveNext
            Next i

            rs.Close
        End If
  .AutoSize 1, .Cols - 1, False

        .RowHeight(-1) = 300
    End With

    

End Sub








Private Sub Form_Load()
   Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500
    Dim Dcombos As New ClsDataCombos
      Set Dcombos = New ClsDataCombos
    Dcombos.GetEmployees Me.DcboEmpName

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
       cahngelang
    End If
FromDate.value = Date
ToDate.value = Date

End Sub




Private Sub ImgFavorites_Click()
AddTofaforites Me.Name, Me.Caption, Me.Caption
End Sub

Private Sub TxtSearchCode_KeyPress(KeyAscii As Integer)
Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode TxtSearchCode.Text, EmpID
        DcboEmpName.BoundText = EmpID
    End If

End Sub
Private Sub DcboEmpName_Change()
DcboEmpName_Click (0)
End Sub

Private Sub DcboEmpName_Click(Area As Integer)
 If val(DcboEmpName.BoundText) = 0 Then Exit Sub
 Dim EmpCode  As String
    GetEmployeeIDFromCode , , DcboEmpName.BoundText, EmpCode
    TxtSearchCode.Text = EmpCode
   ' GetEmployee val(DcboEmpName.BoundText)
End Sub
