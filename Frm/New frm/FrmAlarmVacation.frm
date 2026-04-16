VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SUITEC~1.OCX"
Begin VB.Form FrmAlarmVacation 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÊäÈíåÇÊ ÇáÇÌÇÒÇÊ"
   ClientHeight    =   8115
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14550
   Icon            =   "FrmAlarmVacation.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8115
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
      Top             =   7320
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
         ButtonImage     =   "FrmAlarmVacation.frx":6852
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
         Caption         =   "ãÓÍ"
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
         ButtonImage     =   "FrmAlarmVacation.frx":30474
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
         Caption         =   "ØÈÇÚå"
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
         ButtonImage     =   "FrmAlarmVacation.frx":36CD6
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
      Height          =   4935
      Left            =   0
      TabIndex        =   5
      Top             =   2400
      Width           =   14535
      Begin VSFlex8UCtl.VSFlexGrid GridInstallments 
         Height          =   4695
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   14355
         _cx             =   25321
         _cy             =   8281
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
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FrmAlarmVacation.frx":3D538
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
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   600
      Width           =   14535
      Begin XtremeSuiteControls.RadioButton Opt 
         Height          =   255
         Index           =   0
         Left            =   8160
         TabIndex        =   28
         Top             =   240
         Width           =   1575
         _Version        =   786432
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "ãæÙÝ áã íÑÌÚ"
         ForeColor       =   16711680
         BackColor       =   14871017
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
      Begin VB.TextBox txtEmpCode 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   12765
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   255
         Width           =   885
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÈÏÇíÉ ÇáÇÌÇÒÉ"
         Height          =   735
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   0
         Width           =   5895
         Begin MSComCtl2.DTPicker Fromdate 
            Height          =   330
            Left            =   3240
            TabIndex        =   24
            Top             =   240
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   66977793
            CurrentDate     =   41640
         End
         Begin MSComCtl2.DTPicker todate 
            Height          =   330
            Left            =   360
            TabIndex        =   25
            Top             =   240
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   66977793
            CurrentDate     =   41640
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "ÅÇáì"
            Height          =   435
            Index           =   2
            Left            =   2100
            RightToLeft     =   -1  'True
            TabIndex        =   20
            Top             =   240
            Width           =   540
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "ãä"
            Height          =   315
            Index           =   1
            Left            =   4680
            RightToLeft     =   -1  'True
            TabIndex        =   19
            Top             =   240
            Width           =   585
         End
      End
      Begin MSDataListLib.DataCombo DcboEmpName 
         Height          =   315
         Left            =   9840
         TabIndex        =   22
         Top             =   255
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton Opt 
         Height          =   255
         Index           =   1
         Left            =   6120
         TabIndex        =   29
         Top             =   240
         Width           =   1575
         _Version        =   786432
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "ãÓÊÍÞÇÊ ÇáÇÌÇÒÉ"
         ForeColor       =   16711680
         BackColor       =   14871017
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ÇáãæÙÝ"
         Height          =   240
         Index           =   64
         Left            =   13590
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   240
         Width           =   825
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
         TabIndex        =   16
         Top             =   600
         Width           =   810
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00E2E9E9&
         Caption         =   "äåÇíÉ ÇáÇÌÇÒÉ"
         Height          =   735
         Left            =   8520
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   240
         Width           =   5895
         Begin MSComCtl2.DTPicker FromEnd 
            Height          =   330
            Left            =   2880
            TabIndex        =   26
            Top             =   240
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   66977793
            CurrentDate     =   41640
         End
         Begin MSComCtl2.DTPicker ToEnd 
            Height          =   330
            Left            =   360
            TabIndex        =   27
            Top             =   240
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   582
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   66977793
            CurrentDate     =   41640
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "ãä"
            Height          =   315
            Index           =   0
            Left            =   4680
            RightToLeft     =   -1  'True
            TabIndex        =   15
            Top             =   240
            Width           =   585
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            BackStyle       =   0  'Transparent
            Caption         =   "ÅÇáì"
            Height          =   435
            Index           =   14
            Left            =   2100
            RightToLeft     =   -1  'True
            TabIndex        =   14
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
         Caption         =   "ÈÍË"
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
         ButtonImage     =   "FrmAlarmVacation.frx":3D64E
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
         Caption         =   "ÊÍÏíË ßá"
         Height          =   435
         Index           =   4
         Left            =   6960
         RightToLeft     =   -1  'True
         TabIndex        =   17
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
      Caption         =   "ÊäÈíåÇÊ ÇáÇÌÇÒÇÊ   "
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
         Left            =   8640
         Picture         =   "FrmAlarmVacation.frx":43EB0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   795
      End
   End
End
Attribute VB_Name = "FrmAlarmVacation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Cmd_Click(Index As Integer)
Select Case Index
Case 5
ProgressBar1.Visible = True
: ProgressBar1.value = 10
If Opt(0).value = False And Opt(1).value = False Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "íÑÌì ÅÎÊíÇÑ äæÚ ÇáÊäÈíå"
Else
MsgBox "Please Select Type Alert"
End If
Exit Sub
End If
If Opt(0).value = True Then
Opt_Click (0)
Else
Opt_Click (1)
End If
: ProgressBar1.value = 50
ProgressBar1.Visible = False
ProgressBar1.value = 0
Case 6
Me.Hide
Case 9
If Opt(0).value = False And Opt(1).value = False Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "íÑÌì ÅÎÊíÇÑ äæÚ ÇáÊäÈíå"
Else
MsgBox "Please Select Type Alert"
End If
Exit Sub
End If
If Opt(0).value = True Then
FillGrid , 1
Else
fillgrid1 , 1
End If
End Select

End Sub

Function print_report(Optional NoteSerial As String)
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
MySQL = NoteSerial
If Opt(0).value = True Then
  If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAlarmVacation.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAlarmVacationE.rpt"
 End If
 Else
   If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAlarmVacation2.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAlarmVacation2E.rpt"
 End If
 End If


    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        'GetMsgs 138, vbExclamation
        If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "áÇÊæÌÏ ÈíÇäÇÊ ááÚÑÖ"
        Else
        Msg = "Not Found Data to Show"
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
        StrReportTitle = "" '& StrAccountName
    Else
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.EngCompanyName  ' RPTCompany_Name_Eng
        StrReportTitle = ""

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

Function GetHobStatus() As Integer
Dim Sql As String
GetHobStatus = 0
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
Sql = " SELECT     id, Vacation"
Sql = Sql & " From dbo.jopstatus"
Sql = Sql & " Where (Vacation is null or Vacation=0 )and (resignationInt is null or resignationInt=0 ) "
Rs3.Open Sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
GetHobStatus = IIf(IsNull(Rs3("id").value), 0, Rs3("id").value)
Else
GetHobStatus = 0
End If
End Function
Private Sub Opt_Click(Index As Integer)
If Opt(1).value = True Then

If SystemOptions.UserInterface = ArabicInterface Then
With GridInstallments
.ColHidden(.ColIndex("late")) = True
.TextMatrix(0, .ColIndex("Rec")) = "ÊÇÑíÎ ÇáÇÓÊÍÞÇÞ"
.TextMatrix(0, .ColIndex("RecH")) = "ÊÇÑíÎ ÇáÇÓÊÍÞÇÞ åÜ"
End With
Frame5.Caption = "ÊÇÑíÎ ÇáÇÓÊÍÞÇÞ"
Else
With GridInstallments
.ColHidden(.ColIndex("late")) = True
.TextMatrix(0, .ColIndex("Rec")) = " Date"
.TextMatrix(0, .ColIndex("RecH")) = "Date of Hegira"
End With
Frame5.Caption = " Date"
End If
Frame6.Visible = False
fillgrid1
Else
With GridInstallments
.ColHidden(.ColIndex("late")) = False
If SystemOptions.UserInterface = ArabicInterface Then
.TextMatrix(0, .ColIndex("Rec")) = "ÈÏÇíÉ ÇáÇÌÇÒÉ"
.TextMatrix(0, .ColIndex("RecH")) = "äåÇíÉ ÇáÇÌÇÒÉ"

Frame5.Caption = "äåÇíÉ ÇáÇÌÇÒÉ"
Else
.TextMatrix(0, .ColIndex("Rec")) = "Start Vacation"
.TextMatrix(0, .ColIndex("RecH")) = "End Vacation"
Frame5.Caption = " End Vaction"
End If
End With
Frame6.Visible = True
FillGrid
End If

End Sub

Private Sub txtEmpCode_KeyPress(KeyAscii As Integer)
   Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode txtEmpCode.Text, EmpID
        DcboEmpName.BoundText = EmpID
    End If
    
    
End Sub



Private Sub CmdHelp_Click()
          clear_all Me
            GridInstallments.Clear flexClearScrollable, flexClearEverything
            GridInstallments.Rows = 1
ToDate.value = ""
FromDate.value = ""
FromEnd.value = ""
ToEnd.value = ""
End Sub


Function cahngelang()
    EleHeader.Caption = "Staff Vacations Alerts "
    Me.Caption = EleHeader.Caption
    lbl(64).Caption = "Employee"
   Frame6.Caption = "Start Date"
   Opt(0).RightToLeft = False
   Opt(1).RightToLeft = False
   Opt(0).Caption = "Vaca Entitlements"
   Opt(1).Caption = "Not Return"
   
    lbl(0).Caption = "From"
    lbl(1).Caption = "From"
    lbl(14).Caption = "To"
    lbl(2).Caption = "To"
    Frame5.Caption = "End Date"
    lbl(4).Caption = "Update All"
   CmdHelp.Caption = "Clear"
   
   Cmd(5).Caption = "Saerch"
   Cmd(9).Caption = "Print"
    Cmd(6).Caption = "Exit"
    With GridInstallments
    .TextMatrix(0, .ColIndex("Ser")) = "Serial"
    .TextMatrix(0, .ColIndex("Fullcode")) = "Code"
    .TextMatrix(0, .ColIndex("Emp_Name")) = "Employee Name"
  '  .TextMatrix(0, .ColIndex("ItemName")) = "ItemName"
    .TextMatrix(0, .ColIndex("late")) = "Day delay"
    .TextMatrix(0, .ColIndex("Rec")) = "Start Date"
    .TextMatrix(0, .ColIndex("RecH")) = "End Date "
    .TextMatrix(0, .ColIndex("Val")) = "No Vacation "
   
    End With
End Function
Public Sub fillgrid1(Optional str As String, Optional Trans As Integer = 0)
Dim cont1 As Double
Dim cont As Double
Dim typ As Integer
Dim Sql As String
  '  On Error GoTo ErrTrap
On Error Resume Next
    Dim i As Integer
    Dim rs As ADODB.Recordset
Sql = ""
    Set rs = New ADODB.Recordset
 Sql = "SELECT     TOP 100 PERCENT dbo.tblVacationData.ID, dbo.tblVacationData.ExpectedacationDate, dbo.tblVacationData.ExpectedacationDateH, dbo.tblVacationData.Remark, "
 Sql = Sql + "                      dbo.tblVacationData.[value] , dbo.TblEmployee.emp_Name, dbo.TblEmployee.fullcode, dbo.TblEmployee.Emp_Namee, dbo.TblEmployee.Emp_id , dbo.tblVacationData.EmpID"
 Sql = Sql + "  FROM         dbo.tblVacationData LEFT OUTER JOIN"
 Sql = Sql + "                      dbo.TblEmployee ON dbo.tblVacationData.EmpID = dbo.TblEmployee.Emp_ID"
 Sql = Sql + " where 1=1 and dbo.TblEmployee.jopstatusid =" & GetHobStatus() & " "
  If Not (IsNull(Me.FromEnd.value)) Then
 Sql = Sql + " and (dbo.tblVacationData.ExpectedacationDate >='" & SQLDate(FromEnd.value) & "')"
 End If
 If Not (IsNull(Me.ToEnd.value)) Then
 Sql = Sql + " and (dbo.tblVacationData.ExpectedacationDate <='" & SQLDate(ToEnd.value) & "')"
 End If

If Me.DcboEmpName.Text <> "" And val(Me.DcboEmpName.BoundText) <> 0 Then
Sql = Sql + "and tblVacationData.EmpID =" & val(Me.DcboEmpName.BoundText) & ""
End If


Sql = Sql + "   order by  dbo.tblVacationData.ID "
  
Dim ActualTotal As Double
If Trans = 1 Then
print_report Sql
Exit Sub
End If
'rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText
    rs.Open Sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
      With Me.GridInstallments
       .Rows = 1
        .Clear flexClearScrollable

        If rs.RecordCount > 0 Then
           .Rows = rs.RecordCount + 1
           rs.MoveFirst

            For i = 1 To .Rows - 1
              .TextMatrix(i, .ColIndex("Ser")) = i
              .TextMatrix(i, .ColIndex("Fullcode")) = (IIf(IsNull(rs.Fields("Fullcode").value), "", rs.Fields("Fullcode").value))
              .TextMatrix(i, .ColIndex("Rec")) = (IIf(IsNull(rs.Fields("ExpectedacationDate").value), "", rs.Fields("ExpectedacationDate").value))
              .TextMatrix(i, .ColIndex("RecH")) = (IIf(IsNull(rs.Fields("ExpectedacationDateH").value), "", rs.Fields("ExpectedacationDateH").value))
             .TextMatrix(i, .ColIndex("Val")) = (IIf(IsNull(rs.Fields("Value").value), "", rs.Fields("Value").value))
            If SystemOptions.UserInterface = ArabicInterface Then
           
              .TextMatrix(i, .ColIndex("Emp_Name")) = (IIf(IsNull(rs.Fields("Emp_Name").value), "", rs.Fields("Emp_Name").value))
               
            Else
           
              .TextMatrix(i, .ColIndex("Emp_Name")) = (IIf(IsNull(rs.Fields("Emp_Namee").value), "", rs.Fields("Emp_Namee").value))
              
            End If

        rs.MoveNext
            Next i
 
            rs.Close
        End If
  .AutoSize 1, .Cols - 1, False

        .RowHeight(-1) = 300
    End With

End Sub



Public Sub FillGrid(Optional str As String, Optional Trans As Integer = 0)
Dim cont1 As Double
Dim cont As Double
Dim typ As Integer
Dim Sql As String
  '  On Error GoTo ErrTrap
On Error Resume Next
    Dim i As Integer
    Dim rs As ADODB.Recordset
Sql = ""
    Set rs = New ADODB.Recordset
 Sql = "SELECT     dbo.TblVocationEntitlements.ID, dbo.TblVocationEntitlements.RecordDate, dbo.TblVocationEntitlements.EmpID, dbo.TblEmployee.Emp_Name, "
 Sql = Sql & "                     dbo.TblEmployee.Fullcode, dbo.TblEmployee.Emp_Namee, dbo.TblVocationEntitlements.stratDate, dbo.TblVocationEntitlements.EndDate,"
 Sql = Sql & "                     dbo.TblVocationEntitlements.stratDateH, dbo.TblVocationEntitlements.EndDateH, dbo.TblVocationEntitlements.NoDayAct, dbo.TblVocationEntitlements.NoDayDelay,"
 Sql = Sql & "                     dbo.TblVocationEntitlements.AcuDateH , dbo.TblVocationEntitlements.AcuDate"
 Sql = Sql & " FROM         dbo.TblVocationEntitlements LEFT OUTER JOIN"
 Sql = Sql & "                     dbo.TblEmployee ON dbo.TblVocationEntitlements.EmpID = dbo.TblEmployee.Emp_ID"
  Sql = Sql & " Where (dbo.TblVocationEntitlements.AcuDate Is Null) and dbo.TblEmployee.jopstatusid =" & GetHobStatus()
  If Not (IsNull(Me.FromEnd.value)) Then
 Sql = Sql + " and (dbo.TblVocationEntitlements.EndDate >=" & SQLDate(FromEnd.value, True) & ")"
 End If
 If Not (IsNull(Me.ToEnd.value)) Then
 Sql = Sql + " and (dbo.TblVocationEntitlements.EndDate <=" & SQLDate(ToEnd.value, True) & ")"
 End If
''''''''''''''
  If Not (IsNull(Me.FromDate.value)) Then
 Sql = Sql + " and (dbo.TblVocationEntitlements.stratDate >=" & SQLDate(FromDate.value, True) & ")"
 End If
 If Not (IsNull(Me.ToDate.value)) Then
 Sql = Sql + " and (dbo.TblVocationEntitlements.stratDate <=" & SQLDate(ToDate.value, True) & ")"
 End If
''/////
If Me.DcboEmpName.Text <> "" And val(Me.DcboEmpName.BoundText) <> 0 Then
Sql = Sql + "and TblVocationEntitlements.EmpID =" & val(Me.DcboEmpName.BoundText) & ""
End If
Sql = Sql + "   order by  dbo.TblVocationEntitlements.ID "
  
Dim ActualTotal As Double
If Trans = 1 Then
print_report Sql
Exit Sub
End If
'rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText
    rs.Open Sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
      With Me.GridInstallments
       .Rows = 1
        .Clear flexClearScrollable

        If rs.RecordCount > 0 Then
           .Rows = rs.RecordCount + 1
           rs.MoveFirst

            For i = 1 To .Rows - 1
              .TextMatrix(i, .ColIndex("Ser")) = i
              .TextMatrix(i, .ColIndex("Fullcode")) = (IIf(IsNull(rs.Fields("Fullcode").value), "", rs.Fields("Fullcode").value))
              .TextMatrix(i, .ColIndex("Rec")) = (IIf(IsNull(rs.Fields("stratDate").value), "", rs.Fields("stratDate").value))
              .TextMatrix(i, .ColIndex("RecH")) = (IIf(IsNull(rs.Fields("EndDate").value), "", rs.Fields("EndDate").value))
              If Not (IsNull(rs.Fields("stratDate").value)) And Not (IsNull(rs.Fields("EndDate").value)) Then
              .TextMatrix(i, .ColIndex("Val")) = DateDiff("d", rs.Fields("stratDate").value, rs.Fields("EndDate").value)
              End If
              If Not (IsNull(rs.Fields("EndDate").value)) Then
              .TextMatrix(i, .ColIndex("late")) = DateDiff("d", rs.Fields("EndDate").value, Date)
              End If
            If SystemOptions.UserInterface = ArabicInterface Then
           
              .TextMatrix(i, .ColIndex("Emp_Name")) = (IIf(IsNull(rs.Fields("Emp_Name").value), "", rs.Fields("Emp_Name").value))
               
            Else
           
              .TextMatrix(i, .ColIndex("Emp_Name")) = (IIf(IsNull(rs.Fields("Emp_Namee").value), "", rs.Fields("Emp_Namee").value))
              
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
   'Dcombos.GetCustomersSuppliers 1, Me.DBCboClientName, True  '  2 supplier  1 customer
      

    If SystemOptions.UserInterface = EnglishInterface Then

        SetInterface Me
       cahngelang
    End If

FromDate.value = Date
ToDate.value = Date
FromEnd.value = Date
ToEnd.value = Date
'fillgrid
    
 
End Sub

Private Sub DcboEmpName_Change()
DcboEmpName_Click (0)
End Sub

Private Sub DcboEmpName_Click(Area As Integer)
       If val(DcboEmpName.BoundText) = 0 Then Exit Sub

    Dim EmpCode  As String
 
    GetEmployeeIDFromCode , , DcboEmpName.BoundText, EmpCode
    txtEmpCode.Text = EmpCode
    
End Sub



