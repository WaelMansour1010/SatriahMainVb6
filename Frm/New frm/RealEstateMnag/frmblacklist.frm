VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Begin VB.Form frmblacklist 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«·ﬁ«∆„… «·”Êœ«¡"
   ClientHeight    =   9240
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   8625
   HelpContextID   =   440
   Icon            =   "frmblacklist.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   9240
   ScaleWidth      =   8625
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin C1SizerLibCtl.C1Elastic EleMain 
      Height          =   9255
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   8655
      _cx             =   15266
      _cy             =   16325
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
      AutoSizeChildren=   7
      BorderWidth     =   2
      ChildSpacing    =   1
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
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   2265
         Left            =   30
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   7065
         Width           =   8565
         _cx             =   15108
         _cy             =   3995
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
         BackColor       =   14871017
         ForeColor       =   -2147483630
         FloodColor      =   6553600
         ForeColorDisabled=   -2147483631
         Caption         =   ""
         Align           =   0
         AutoSizeChildren=   7
         BorderWidth     =   2
         ChildSpacing    =   1
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
         Begin VB.TextBox txtId 
            Alignment       =   1  'Right Justify
            Height          =   420
            Left            =   2400
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   12
            Top             =   1605
            Width           =   5205
         End
         Begin VB.TextBox txtMessage 
            Alignment       =   1  'Right Justify
            Height          =   420
            Left            =   2415
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   10
            Top             =   1035
            Width           =   5205
         End
         Begin VB.CheckBox ChkShow 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "·«  ŸÂ— Â–Â «·‰«›–… ⁄‰œ  ‘€Ì· «·»—‰«„Ã"
            ForeColor       =   &H000000FF&
            Height          =   1095
            Left            =   3465
            RightToLeft     =   -1  'True
            TabIndex        =   4
            Top             =   2955
            Width           =   4875
         End
         Begin ImpulseButton.ISButton CmdExit 
            Cancel          =   -1  'True
            Height          =   840
            Left            =   120
            TabIndex        =   5
            Top             =   630
            Width           =   705
            _ExtentX        =   1244
            _ExtentY        =   1482
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
            ButtonImage     =   "frmblacklist.frx":038A
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
         Begin ImpulseButton.ISButton CmdPrint 
            Height          =   855
            Left            =   1950
            TabIndex        =   6
            Top             =   2190
            Visible         =   0   'False
            Width           =   450
            _ExtentX        =   794
            _ExtentY        =   1508
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
            ButtonImage     =   "frmblacklist.frx":0724
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   -2147483637
         End
         Begin ImpulseButton.ISButton SendMessage 
            Height          =   615
            Left            =   960
            TabIndex        =   8
            Top             =   2730
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   1085
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "«—”«·"
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
            Height          =   390
            Index           =   9
            Left            =   960
            TabIndex        =   14
            Top             =   870
            Width           =   765
            _ExtentX        =   1349
            _ExtentY        =   688
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
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   -2147483637
         End
         Begin MSDataListLib.DataCombo DcbBranch 
            Height          =   315
            Left            =   2400
            TabIndex        =   15
            Top             =   600
            Width           =   5205
            _ExtentX        =   9181
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "»ÕÀ »«·›—⁄"
            Height          =   540
            Left            =   7680
            RightToLeft     =   -1  'True
            TabIndex        =   16
            Top             =   600
            Width           =   765
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "»ÕÀ »«·ÂÊÌ…"
            Height          =   525
            Left            =   7485
            RightToLeft     =   -1  'True
            TabIndex        =   13
            Top             =   1605
            Width           =   1005
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "»ÕÀ »«·«”„"
            Height          =   540
            Left            =   7740
            RightToLeft     =   -1  'True
            TabIndex        =   9
            Top             =   1035
            Width           =   765
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "Ì „  ÕœÌœ Â–Â «·»Ì«‰«  »‰«¡« ⁄·Ï «· «—ÌŒ «·Õ«·Ì ›Ì «·ÃÂ«“"
            ForeColor       =   &H000000FF&
            Height          =   705
            Left            =   4035
            RightToLeft     =   -1  'True
            TabIndex        =   7
            Top             =   3540
            Visible         =   0   'False
            Width           =   4425
         End
      End
      Begin VSFlex8UCtl.VSFlexGrid Fg 
         Height          =   6735
         Left            =   30
         TabIndex        =   2
         Top             =   1005
         Width           =   8520
         _cx             =   15028
         _cy             =   11880
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
         BackColorBkg    =   16777215
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
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   12
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmblacklist.frx":0ABE
         ScrollTrack     =   0   'False
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
      Begin VB.CheckBox Check17 
         Alignment       =   1  'Right Justify
         Caption         =   " ÕœÌœ «·ﬂ·"
         Height          =   270
         Left            =   9000
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   795
         Width           =   1425
      End
      Begin VB.Image ImgFavorites 
         Height          =   390
         Left            =   2160
         Picture         =   "frmblacklist.frx":0C8A
         Stretch         =   -1  'True
         Top             =   120
         Width           =   525
      End
      Begin VB.Image Image2 
         Height          =   645
         Left            =   120
         Picture         =   "frmblacklist.frx":48F2
         Stretch         =   -1  'True
         Top             =   120
         Width           =   855
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   30
         Picture         =   "frmblacklist.frx":576E
         Top             =   30
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label LblCaption 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         Caption         =   "«·ﬁ«∆„… «·”Êœ«¡"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1035
         Left            =   30
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   30
         Width           =   8580
      End
   End
End
Attribute VB_Name = "frmblacklist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Askinterval As String
Dim Askcount As Integer
Dim sql As String
Function print_report(Optional NoteSerial As String)
     Dim rs As ADODB.Recordset
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
        
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAqarBLockCustomer.rpt"
            Else
             StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepAqarBLockCustomer.rpt"
            
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

   ' xReport.ParameterFields(3).AddCurrentValue user_name
   ' xReport.ParameterFields(13).AddCurrentValue Me.DTPicker1.value
      '  xReport.ParameterFields(4).AddCurrentValue WriteNo(Format(val(TxtAdvanceValue.text), "0.00"), 0, True, ".")
       ' xReport.ParameterFields(6).AddCurrentValue val(lbl(23).Caption)
        'xReport.ParameterFields(13).AddCurrentValue Me.DTPicker1.value
'    xReport.ParameterFields(8).AddCurrentValue IIf(IsNumeric(fg.TextMatrix(Me.fg.FixedRows, fg.ColIndex("PartValue"))), val(fg.TextMatrix(Me.fg.FixedRows, fg.ColIndex("PartValue"))), 0)
'Dim gr, order As Integer
' xReport.ParameterFields(14).AddCurrentValue Order
 'xReport.ParameterFields(15).AddCurrentValue gr
 ' xReport.ParameterFields(15).AddCurrentValue gr
 ' xReport.ParameterFields(10).AddCurrentValue val(TxtDiscount.text)
  ' xReport.ParameterFields(11).AddCurrentValue txtDiscountDES.text
  Dim total As String
  Dim totl As Double
 ' totl = val(LbToTalExtra.Caption) + val(Me.lbTotalMente.Caption)
 ' total = totl
 '  xReport.ParameterFields(12).AddCurrentValue Me.lbTotalMente.Caption
 '     xReport.ParameterFields(13).AddCurrentValue LbToTalExtra.Caption
 '       xReport.ParameterFields(14).AddCurrentValue total
   ' xReport.ParameterFields(5).AddCurrentValue ToHijriDate(RsData("notedate").value)
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
Private Sub Check17_Click()
    Dim i As Integer

    If Check17.value = vbChecked Then

        With Me.Fg
 
            For i = 1 To .Rows - 2
        
                .TextMatrix(i, .ColIndex("Send")) = True
            Next i

        End With

    Else

        With Me.Fg

            For i = 1 To .Rows - 2
        
                .TextMatrix(i, .ColIndex("Send")) = False
            Next i

        End With

    End If

End Sub


Private Sub Cmd_Click(Index As Integer)
print_report sql
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CmdPrint_Click()

    If DoPremis(Do_Print, Me.Name, True) = False Then
        Exit Sub
    End If
        
    On Error GoTo ErrTrap
    Dim Reports As ClsRepoerts
    Dim StrSQL As String

    Askinterval = GetSetting(StrAppRegPath, "Setting", "INTERVAL_InstallmentMustPayed", True)
    Askcount = GetSetting(StrAppRegPath, "Setting", "count_InstallmentMustPayed", True)
    
    'StrSQL = "select * From QestNotReceipted where  DueDate<='" & SQLDate(DateAdd(Askinterval, Askcount, Date)) & "'"
    ' StrSQL = StrSQL + " order by CusName,Transaction_ID,QeqtNum"

    StrSQL = "SELECT     TOP 100 PERCENT dbo.QryCust_Qest.QestID, dbo.QryCust_Qest.NoteID, dbo.QryCust_Qest.QeqtNum, dbo.QryCust_Qest.PartID, dbo.QryCust_Qest.[Value], "
    StrSQL = StrSQL + " dbo.QryCust_Qest.DueDate, dbo.QryCust_Qest.Receipt, dbo.QryCust_Qest.Summition, dbo.QryCust_Qest.CustID, dbo.QryCust_Qest.CusName,"
    StrSQL = StrSQL + "  dbo.QryCust_Qest.Transaction_ID , dbo.QryCust_Qest.Transaction_Date, dbo.Transactions.NoteSerial1"
    StrSQL = StrSQL + " FROM         dbo.QryCust_Qest LEFT OUTER JOIN"
    StrSQL = StrSQL + "  dbo.Transactions ON dbo.QryCust_Qest.Transaction_ID = dbo.Transactions.Transaction_ID"
    StrSQL = StrSQL + " WHERE     (dbo.QryCust_Qest.QestID NOT IN"
    StrSQL = StrSQL + " (SELECT     QestID"
    StrSQL = StrSQL + "  from InstallmentDet_Junc_Receipt"
    StrSQL = StrSQL + " WHERE     Status <> 1))"
    StrSQL = StrSQL + "  and DueDate <" & SQLDate(Date, True) & "'"
    StrSQL = StrSQL + "  order by CusName,QryCust_Qest.Transaction_ID,QeqtNum"
 
    Set Reports = New ClsRepoerts
    Reports.QestMustPayed StrSQL, , LblCaption.Caption
    Exit Sub
ErrTrap:
End Sub





Private Sub DcbBranch_Click(Area As Integer)
loadgrid
End Sub

Private Sub Fg_BeforeEdit(ByVal Row As Long, _
                          ByVal Col As Long, _
                          Cancel As Boolean)

    If Col <> Fg.ColIndex("Send") Then
        Cancel = True
    End If

End Sub

Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim My_SQL As String
    Dim RowNum As Integer
    Dim ReCount As Integer
    Dim Dcombos As ClsDataCombos
    
    Set Dcombos = New ClsDataCombos
 
    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500
    
  Dcombos.GetBranches DcbBranch
    LoadIcons
loadgrid
 
    
 
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    Exit Sub
ErrTrap:
End Sub
Function loadgrid()
   Dim BGround As New ClsBackGroundPic
    Dim BolShowRequest As Boolean

    Dim RsTemp As New ADODB.Recordset
        Dim StrSQL As String
  StrSQL = "SELECT     *, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee"
  StrSQL = StrSQL & " FROM         dbo.TblCustemers LEFT OUTER JOIN"
  StrSQL = StrSQL & "                    dbo.TblBranchesData ON dbo.TblCustemers.BranchId = dbo.TblBranchesData.branch_id"
  StrSQL = StrSQL & " Where (dbo.TblCustemers.CusID > 2) And (dbo.TblCustemers.Type = 1 or dbo.TblCustemers.Type = 20 or dbo.TblCustemers.Type = 55 or dbo.TblCustemers.Type = 56 or dbo.TblCustemers.Type = 57) And (dbo.TblCustemers.locked = 1)"
If txtMessage.Text <> "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
                StrSQL = StrSQL & " and  CusName like'%" & txtMessage.Text & "%'"
        Else
            StrSQL = StrSQL & " and  CusNamee like'%" & txtMessage.Text & "%'"
        End If
End If

If txtid.Text <> "" Then
 
                StrSQL = StrSQL & " and  CustGID =" & txtid.Text
      
End If
If val(Me.DcbBranch.BoundText) <> 0 Or Me.DcbBranch.Text <> "" Then
StrSQL = StrSQL & " AND TblCustemers.BranchId = " & val(Me.DcbBranch.BoundText)

End If


        If SystemOptions.UserInterface = ArabicInterface Then
            StrSQL = StrSQL + " Order by CusName "
        Else
            StrSQL = StrSQL + " Order by CusNamee "
        End If

     With Fg
            .Rows = .FixedRows
 
                
        End With
        
Dim ReCount As Integer
Dim RowNum As Integer
sql = StrSQL
    RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Not (RsTemp.EOF Or RsTemp.BOF) Then

        With Fg
            .Rows = .FixedRows

            For ReCount = 1 To RsTemp.RecordCount
                .Rows = .Rows + 1
                RowNum = .Rows - 1
                   
                ', dbo.QryCust_Qest.CustID
                If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(RowNum, .ColIndex("Name")) = IIf(IsNull(RsTemp("CusNamee").value), "", RsTemp("CusNamee").value)

                    .TextMatrix(RowNum, .ColIndex("Branch")) = IIf(IsNull(RsTemp("branch_name").value), "", RsTemp("branch_name").value)
                Else
                .TextMatrix(RowNum, .ColIndex("Name")) = IIf(IsNull(RsTemp("CusNamee").value), "", RsTemp("CusNamee").value)

                    .TextMatrix(RowNum, .ColIndex("Branch")) = IIf(IsNull(RsTemp("branch_namee").value), "", RsTemp("branch_namee").value)

                End If
 .TextMatrix(RowNum, .ColIndex("remark2")) = IIf(IsNull(RsTemp("remark2").value), "", RsTemp("remark2").value)
                .TextMatrix(RowNum, .ColIndex("Numbers")) = IIf(IsNull(RsTemp("Cus_mobile").value), "", RsTemp("Cus_mobile").value)
            
                RsTemp.MoveNext
            Next ReCount

            .AutoSize 0, .Cols - 1, False
        End With

    End If

    Fg.WallPaper = BGround.Picture
    BolShowRequest = GetSetting(StrAppRegPath, "View_Type", "InstallmentMustPayed", True)

    If BolShowRequest = True Then
        ChkShow.value = Unchecked
    Else
        ChkShow.value = Checked
    End If

End Function
Private Sub ChangeLang()
    Me.Caption = "Installment Must Pay"
    LblCaption.Caption = Me.Caption
    ChkShow.Caption = "Dont Show at Start"
    Label1.Caption = "Data Based in your System Date"
    Me.CmdExit.Caption = "Exit"
    Me.CmdPrint.Caption = "Print"

    With Me.Fg
        .TextMatrix(0, .ColIndex("Name")) = "Customer Name"
        .TextMatrix(0, .ColIndex("BillIID")) = "BillI ID"
        .TextMatrix(0, .ColIndex("TransDate")) = "Trans Date"
        .TextMatrix(0, .ColIndex("QestNum")) = "installm. #"
        .TextMatrix(0, .ColIndex("DueDate")) = "DueDate"
        .TextMatrix(0, .ColIndex("value")) = "value"

    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap

    If ChkShow.value = Checked Then
        SaveSetting StrAppRegPath, "View_Type", "InstallmentMustPayed", False
    Else
        SaveSetting StrAppRegPath, "View_Type", "InstallmentMustPayed", True
    End If

    FormPostion Me, SavePostion
    Exit Sub
ErrTrap:
End Sub

Private Sub LoadIcons()
    On Error GoTo ErrTrap

    With Fg
        .Cell(flexcpPicture, 0, .ColIndex("Name")) = mdifrmmain.ImgLstTree.ListImages("User").Picture
        .Cell(flexcpPicture, 0, .ColIndex("BillIID")) = mdifrmmain.ImgLstTree.ListImages("number").Picture
        .Cell(flexcpPicture, 0, .ColIndex("TransDate")) = mdifrmmain.ImgLstTree.ListImages("qty").Picture
        .Cell(flexcpPicture, 0, .ColIndex("QestNum")) = mdifrmmain.ImgLstTree.ListImages("number").Picture
        .Cell(flexcpPicture, 0, .ColIndex("DueDate")) = mdifrmmain.ImgLstTree.ListImages("Date").Picture
        .Cell(flexcpPicture, 0, .ColIndex("Value")) = mdifrmmain.ImgLstTree.ListImages("Price").Picture
        .Cell(flexcpPictureAlignment, 0, 0, 0, .Cols - 1) = flexPicAlignRightCenter
    End With

    Exit Sub
ErrTrap:
End Sub

Private Sub ImgFavorites_Click()
AddTofaforites Me.Name, Me.Caption, Me.Caption
End Sub

Private Sub LblCaption_Click()
    On Error GoTo ErrTrap

    If Me.WindowState = vbNormal Then
        Me.WindowState = vbMaximized
    Else
        Me.WindowState = vbNormal
    End If

    Exit Sub
ErrTrap:
End Sub

Function GetNumbers()

End Function

Private Sub SendMessage_Click()
    Dim Numbers As String
    Dim RowNum As Integer
    Dim Opt As Integer
    Dim CurrentMessage As String
    Numbers = ""

    With Fg

        For RowNum = .FixedRows To .Rows - 1
    
            If .Cell(flexcpChecked, RowNum, .ColIndex("Send")) = flexChecked Then

                '  MsgBox (.TextMatrix(RowNum, .ColIndex("Numbers")))
                If (.TextMatrix(RowNum, .ColIndex("Numbers"))) <> "" Then
                    If Numbers = "" Then
                        Numbers = (.TextMatrix(RowNum, .ColIndex("Numbers")))
                    Else
                        Numbers = Numbers & "," & (.TextMatrix(RowNum, .ColIndex("Numbers")))
                    End If
             
                End If
            End If
          
        Next RowNum
      
        CurrentMessage = txtMessage.Text

        If Numbers = "" Then Exit Sub
        SMSSeTTings.SendMessage CurrentMessage, Numbers
        SMSSeTTings.Hide
                                    
    End With

End Sub

Private Sub TXTid_Change()
loadgrid
End Sub

Private Sub txtMessage_Change()
loadgrid
End Sub
