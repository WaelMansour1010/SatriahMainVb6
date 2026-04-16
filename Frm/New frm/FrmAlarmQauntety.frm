VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmAlarmQauntety 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ‰»ÌÂ«  «·ﬂ„Ì« "
   ClientHeight    =   7080
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14550
   Icon            =   "FrmAlarmQauntety.frx":0000
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
      TabIndex        =   9
      Top             =   6240
      Width           =   14535
      Begin ImpulseButton.ISButton Cmd 
         Height          =   495
         Index           =   6
         Left            =   480
         TabIndex        =   10
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
         ButtonImage     =   "FrmAlarmQauntety.frx":6852
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
         Left            =   5040
         TabIndex        =   11
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
         ButtonImage     =   "FrmAlarmQauntety.frx":30474
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
         Left            =   9240
         TabIndex        =   12
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
         ButtonImage     =   "FrmAlarmQauntety.frx":36CD6
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
      Height          =   3855
      Left            =   0
      TabIndex        =   5
      Top             =   2400
      Width           =   14535
      Begin VSFlex8UCtl.VSFlexGrid GridInstallments 
         Height          =   3615
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   14355
         _cx             =   25321
         _cy             =   6376
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
         Cols            =   12
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FrmAlarmQauntety.frx":3D538
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
            TabIndex        =   7
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
            TabIndex        =   8
            Top             =   -600
            Width           =   375
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2E9E9&
      Height          =   735
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   14535
      Begin VB.TextBox TxtItemCode 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4485
         TabIndex        =   21
         Top             =   240
         Width           =   1305
      End
      Begin VB.TextBox TxtSearchCode 
         Height          =   315
         Left            =   12135
         TabIndex        =   19
         Top             =   240
         Width           =   1305
      End
      Begin MSDataListLib.DataCombo DBCboClientName 
         Height          =   315
         Left            =   8040
         TabIndex        =   20
         Top             =   240
         Width           =   4110
         _ExtentX        =   7250
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "6"
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcboItems 
         Height          =   315
         Left            =   360
         TabIndex        =   22
         Top             =   240
         Width           =   4110
         _ExtentX        =   7250
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·’‰›"
         Height          =   285
         Index           =   1
         Left            =   5760
         TabIndex        =   23
         Top             =   240
         Width           =   1005
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·⁄„Ì·"
         Height          =   285
         Index           =   3
         Left            =   13200
         TabIndex        =   18
         Top             =   240
         Width           =   1005
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E2E9E9&
      Height          =   1095
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   1320
      Width           =   14535
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6960
         TabIndex        =   24
         Top             =   600
         Width           =   810
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·› —Â"
         Height          =   735
         Left            =   8520
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   240
         Width           =   5895
         Begin MSComCtl2.DTPicker Fromdate 
            Height          =   330
            Left            =   3000
            TabIndex        =   14
            Top             =   240
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   92733441
            CurrentDate     =   41640
         End
         Begin MSComCtl2.DTPicker todate 
            Height          =   330
            Left            =   360
            TabIndex        =   15
            Top             =   240
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   92733441
            CurrentDate     =   41640
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
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "≈«·Ï"
            Height          =   435
            Index           =   14
            Left            =   2100
            RightToLeft     =   -1  'True
            TabIndex        =   16
            Top             =   240
            Width           =   540
         End
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   735
         Index           =   5
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   6285
         _ExtentX        =   11086
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
         ButtonImage     =   "FrmAlarmQauntety.frx":3D6F7
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
         Left            =   6960
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   240
         Width           =   780
      End
   End
   Begin C1SizerLibCtl.C1Elastic EleHeader 
      Height          =   585
      Left            =   0
      TabIndex        =   2
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
      Caption         =   " ‰»ÌÂ«  «·ﬂ„Ì«  /„ÕÃÊ“/„”·„ /„ »ﬁÌ                                                  "
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
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H000000FF&
         Height          =   555
         Index           =   27
         Left            =   2520
         TabIndex        =   3
         Top             =   0
         Width           =   2205
      End
      Begin VB.Image Image1 
         Height          =   555
         Index           =   0
         Left            =   10440
         Picture         =   "FrmAlarmQauntety.frx":43F59
         Stretch         =   -1  'True
         Top             =   0
         Width           =   795
      End
   End
End
Attribute VB_Name = "FrmAlarmQauntety"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Cmd_Click(Index As Integer)
Select Case Index
Case 5
ProgressBar1.Visible = True
: ProgressBar1.value = 10
FillGrid
: ProgressBar1.value = 50
ProgressBar1.Visible = False
ProgressBar1.value = 0
Case 6
Me.Hide
Case 9
GetData
End Select

End Sub
Public Sub GetData()
    Dim StrSQL As String
    Dim StrSQL1 As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
     Dim Rs1 As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
Dim id As Integer
Dim cod As Integer
Dim strname As String
 
                StrSQL = "Delete From TblAlarmQuntity Where 1<>-1"
                Cn.Execute StrSQL, , adExecuteNoRecords

 Set rs = New ADODB.Recordset
       rs.Open "TblAlarmQuntity", Cn, adOpenKeyset, adLockOptimistic, adCmdTable
       With GridInstallments
If .Rows > 1 Then
          
       For i = .FixedRows To .Rows - 1
         If val(.TextMatrix(i, .ColIndex("Ser"))) <> 0 Then
           rs.AddNew
          rs("ID").value = val(.TextMatrix(i, .ColIndex("Ser")))
        rs("Transaction_Date").value = IIf(IsDate(.TextMatrix(i, .ColIndex("Transaction_Date"))), .TextMatrix(i, .ColIndex("Transaction_Date")), Null)
        rs("mah").value = val(.TextMatrix(i, .ColIndex("mah")))
        rs("ms").value = val(.TextMatrix(i, .ColIndex("ms")))
        rs("mt").value = val(.TextMatrix(i, .ColIndex("mt")))
        rs("CusName").value = .TextMatrix(i, .ColIndex("CusName"))
        rs("ItemName").value = .TextMatrix(i, .ColIndex("ItemName"))
                rs.update
        
        End If
        Next i
        End If
End With
StrSQL = "SELECT *  from TblAlarmQuntity "

   StrSQL = StrSQL & " Order By dbo.TblAlarmQuntity.ID"
   
  print_report StrSQL

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
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepShippingQuntity.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepShippingQuntityE.rpt"
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
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
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
  If Fromdate.value <> "" And todate.value <> "" Then
   xReport.ParameterFields(8).AddCurrentValue Fromdate.value

    xReport.ParameterFields(9).AddCurrentValue todate.value
  
    End If
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

Private Sub CmdHelp_Click()
          clear_all Me
            GridInstallments.Clear flexClearScrollable, flexClearEverything
            GridInstallments.Rows = 1
todate.value = Date
Fromdate.value = Date
End Sub

Private Sub DBCboClientName_Change()
    TxtSearchCode.text = ""

    Dim DefaultSalesPersonId As Integer
    Dim fullcode As String

    GetCustomersDetail val(DBCboClientName.BoundText), DefaultSalesPersonId, fullcode

    TxtSearchCode.text = fullcode
 
End Sub
Function cahngelang()
    EleHeader.Caption = " The Amounts Reserved and the Delivered and the Remaining "
    Me.Caption = EleHeader.Caption
    lbl(3).Caption = "Customer"
    lbl(1).Caption = "Item"
    lbl(0).Caption = "From"
    lbl(14).Caption = "To"
    Frame5.Caption = "Period"
    lbl(4).Caption = "Update All"
   CmdHelp.Caption = "Clear"
   
   Cmd(5).Caption = "Saerch"
   Cmd(9).Caption = "Print"
    Cmd(6).Caption = "Exit"
    With GridInstallments
    .TextMatrix(0, .ColIndex("Ser")) = "Serial"
    .TextMatrix(0, .ColIndex("Transaction_Date")) = "Date"
    .TextMatrix(0, .ColIndex("CusName")) = "Customer"
    .TextMatrix(0, .ColIndex("ItemName")) = "ItemName"
    .TextMatrix(0, .ColIndex("mah")) = "Reserved"
    .TextMatrix(0, .ColIndex("ms")) = "Delivered"
    .TextMatrix(0, .ColIndex("mt")) = "Remaining "

   
    End With
End Function
Public Sub RetriveOrder(Optional IDde As Integer = 0, Optional order_no As String = "")
    Dim RsDetails As New ADODB.Recordset
    Dim StrSQL As String
    Dim RsNotes As New ADODB.Recordset
    Dim RsTemp As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Dim Num As Long
   'sa On Error GoTo ErrTrap

    StrSQL = "Select * from transactions  where   ResProductionNo='" & order_no & "'"

    
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If rs.RecordCount < 1 Then
   
 
        Exit Sub
    Else
    Dim IDde2 As Integer
    IDde2 = IDde
      
   RetriveOrder2 IDde2, rs("ResProductionNo").value
    End If

    If rs.EOF Or rs.BOF Then
        Exit Sub
    End If

 
    StrSQL = "SELECT TblItems.HaveSerial, * FROM TblItems INNER JOIN Transaction_Details " & "ON TblItems.ItemID = Transaction_Details.Item_ID INNER JOIN dbo.TblUnites ON dbo.Transaction_Details.UnitID = dbo.TblUnites.UnitID"
    StrSQL = StrSQL + " where Transaction_ID=" & val(rs("Transaction_ID").value)

    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
   
    If Not (RsDetails.EOF Or RsDetails.BOF) Then
    With GridInstallments
    Dim J As Integer
    For J = 1 To RsDetails.RecordCount
    If val(.TextMatrix(IDde, .ColIndex("id"))) = 0 Then
                  .TextMatrix(IDde, .ColIndex("mt")) = IIf(IsNull(RsDetails("showqty")), "", (RsDetails("showqty").value))
                  .TextMatrix(IDde, .ColIndex("id")) = IIf(IsNull(RsDetails("id")), "", (RsDetails("id").value))
                  IDde = IDde + 1
                  
                  RsDetails.MoveNext
                  End If
                  Next J
                  End With
        End If


End Sub
Public Sub RetriveOrder2(Optional IDde As Integer = 0, Optional order_no As String = "")
    Dim RsDetails As New ADODB.Recordset
    Dim StrSQL As String
    Dim RsNotes As New ADODB.Recordset
    Dim RsTemp As ADODB.Recordset
    Dim rs As ADODB.Recordset
    Dim Num As Long
   'sa On Error GoTo ErrTrap

    StrSQL = "SELECT  dbo.gettotalShippedQty1(ProkerId) AS totaRQty, ResProductionNo"
    StrSQL = StrSQL & " FROM         dbo.Transactions  where   ResProductionNo='" & order_no & "'"

    
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If rs.RecordCount < 1 Then

 
        Exit Sub
    Else
        With GridInstallments
    Dim J As Integer
    For J = 1 To 1
    If val(.TextMatrix(IDde, .ColIndex("idd"))) = 0 Then
                  .TextMatrix(IDde, .ColIndex("ms")) = IIf(IsNull(rs("totaRQty")), "", (rs("totaRQty").value))
                  .TextMatrix(IDde, .ColIndex("idd")) = IIf(IsNull(rs("id")), "", (rs("id").value))
                  IDde = IDde + 1
                  
                  rs.MoveNext
                  End If
                  Next J
                  End With
    End If

  
  '  StrSQL = "SELECT TblItems.HaveSerial, * FROM TblItems INNER JOIN Transaction_Details " & "ON TblItems.ItemID = Transaction_Details.Item_ID INNER JOIN dbo.TblUnites ON dbo.Transaction_Details.UnitID = dbo.TblUnites.UnitID"
  '  StrSQL = StrSQL + " where Transaction_ID=" & val(rs("Transaction_ID").value)
'
'    RsDetails.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText
'
'    If Not (RsDetails.EOF Or RsDetails.BOF) Then

'        End If


End Sub
Public Sub FillGrid(Optional str As String)
Dim cont1 As Double
Dim cont As Double
Dim typ As Integer
  '  On Error GoTo ErrTrap
On Error Resume Next
    Dim i As Integer
    Dim rs As ADODB.Recordset

    Set rs = New ADODB.Recordset
 My_SQL = "SELECT     dbo.TblItems.HaveSerial, dbo.Transactions.Transaction_Type, dbo.Transactions.CusID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, "
 My_SQL = My_SQL & "                     dbo.TblCustemers.Fullcode, dbo.TblItems.ItemName, dbo.TblItems.ItemNamee, dbo.TblUnites.UnitName, dbo.TblUnites.UnitNamee,"
 My_SQL = My_SQL & "                     dbo.Transactions.Transaction_ID AS Transaction_IDH, dbo.Transactions.Transaction_Date AS Transaction_DateH, dbo.Transactions.Transaction_Serial,"
 My_SQL = My_SQL & "                     dbo.Transactions.Transaction_HijriDate, dbo.Transactions.PaymentType, dbo.Transactions.Trans_Discount, dbo.Transactions.Trans_DiscountType,"
 My_SQL = My_SQL & "                     dbo.Transactions.StoreID, dbo.Transactions.NoteSerial, dbo.Transactions.NoteSerial1, dbo.Transaction_Details.*"
 My_SQL = My_SQL & " FROM         dbo.TblCustemers RIGHT OUTER JOIN"
 My_SQL = My_SQL & "                     dbo.Transactions ON dbo.TblCustemers.CusID = dbo.Transactions.CusID LEFT OUTER JOIN"
 My_SQL = My_SQL & "                     dbo.TblItems INNER JOIN"
 My_SQL = My_SQL & "                     dbo.Transaction_Details ON dbo.TblItems.ItemID = dbo.Transaction_Details.Item_ID INNER JOIN"
 My_SQL = My_SQL & "                     dbo.TblUnites ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID"
 My_SQL = My_SQL & " Where (dbo.Transactions.Transaction_Type = 61)"
  If Not (IsNull(Me.Fromdate.value)) Then
 My_SQL = My_SQL + " and (dbo.Transactions.Transaction_Date >='" & SQLDate(Fromdate.value) & "')"
 End If
 If Not (IsNull(Me.todate.value)) Then
 My_SQL = My_SQL + " and (dbo.Transactions.Transaction_Date <='" & SQLDate(todate.value) & "')"
 End If

If Me.DBCboClientName.text <> "" And val(Me.DBCboClientName.BoundText) <> 0 Then
My_SQL = My_SQL + "and Transactions.CusID =" & val(Me.DBCboClientName.BoundText) & ""
End If
If Me.DcboItems.text <> "" And val(Me.DcboItems.BoundText) <> 0 Then
My_SQL = My_SQL + "and Transaction_Details.Item_ID =" & val(Me.DcboItems.BoundText) & ""
End If

My_SQL = My_SQL + "   order by  dbo.Transactions.Transaction_Serial "
  
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
              .TextMatrix(i, .ColIndex("Transaction_Serial")) = (IIf(IsNull(rs.Fields("NoteSerial1").value), "", rs.Fields("NoteSerial1").value))
              .TextMatrix(i, .ColIndex("Transaction_ID")) = (IIf(IsNull(rs.Fields("Transaction_ID").value), "", rs.Fields("Transaction_ID").value))
              .TextMatrix(i, .ColIndex("Transaction_Date")) = (IIf(IsNull(rs.Fields("Transaction_DateH").value), "", rs.Fields("Transaction_DateH").value))
             
            If SystemOptions.UserInterface = ArabicInterface Then
           
              .TextMatrix(i, .ColIndex("CusName")) = (IIf(IsNull(rs.Fields("CusName").value), "", rs.Fields("CusName").value))
              .TextMatrix(i, .ColIndex("ItemName")) = (IIf(IsNull(rs.Fields("ItemName").value), "", rs.Fields("ItemName").value))
               
            Else
           
              .TextMatrix(i, .ColIndex("CusName")) = (IIf(IsNull(rs.Fields("CusNamee").value), "", rs.Fields("CusNamee").value))
              .TextMatrix(i, .ColIndex("ItemName")) = (IIf(IsNull(rs.Fields("ItemNamee").value), "", rs.Fields("ItemNamee").value))
              
            End If
            Dim IDde As Integer
            IDde = i
            RetriveOrder IDde, (IIf(IsNull(rs.Fields("NoteSerial1").value), "", rs.Fields("NoteSerial1").value))
           .TextMatrix(i, .ColIndex("mah")) = IIf(IsNull(rs("showqty")), "", (rs("showqty").value))
           
    .TextMatrix(i, .ColIndex("mt")) = val(.TextMatrix(i, .ColIndex("mah"))) - val(.TextMatrix(i, .ColIndex("ms")))

        rs.MoveNext
            Next i
 
            rs.Close
        End If
 'sa .AutoSize 1, .Cols - 1, False

        .RowHeight(-1) = 300
    End With

End Sub

Private Sub DcboItems_Change()
       Me.TxtItemCode.text = GetItemCode(val(Me.DcboItems.BoundText))
  End Sub





Private Sub TxtItemCode_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Then
        If TxtItemCode.text = "" Then
            Me.DcboItems.BoundText = ""
        Else
            Me.DcboItems.BoundText = GetItemID(Trim$(Me.TxtItemCode.text))
        End If
    End If

End Sub
Private Sub Form_Load()
   Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500
    Dim Dcombos As New ClsDataCombos
      Set Dcombos = New ClsDataCombos
    Dcombos.GetItemsNames Me.DcboItems
   Dcombos.GetCustomersSuppliers 1, Me.DBCboClientName, True  '  2 supplier  1 customer
      

    If SystemOptions.UserInterface = EnglishInterface Then

        SetInterface Me
       cahngelang
    End If

Fromdate.value = Date
todate.value = Date
FillGrid
    
 
End Sub



Private Sub TxtSearchCode_KeyPress(KeyAscii As Integer)
    Dim CUSTID As Integer

    If KeyAscii = vbKeyReturn Then
        GetCustomersDetail CUSTID, , TxtSearchCode.text, 1
        DBCboClientName.BoundText = CUSTID
    End If

End Sub
