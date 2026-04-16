VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmBoxesAccounts 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "—’Ìœ «·Œ“‰… «·√‰"
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8610
   HelpContextID   =   1000
   Icon            =   "FrmBoxesAccounts.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5610
   ScaleWidth      =   8610
   ShowInTaskbar   =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.CheckBox chkAll 
      Alignment       =   1  'Right Justify
      Caption         =   "«Œ Ì«— «·ﬂ·"
      Height          =   225
      Left            =   6360
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   4620
      Width           =   1365
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   " ﬁ—Ì— »ﬂ‘› Õ”«» «·Œ“‰…"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1575
      Index           =   1
      Left            =   60
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   3420
      Width           =   2235
      Begin MSComCtl2.DTPicker DtpBoxFrom 
         Height          =   330
         Left            =   90
         TabIndex        =   6
         Top             =   300
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   582
         _Version        =   393216
         CalendarTrailingForeColor=   0
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   143458305
         CurrentDate     =   38845
      End
      Begin MSComCtl2.DTPicker DtpBoxTo 
         Height          =   360
         Left            =   90
         TabIndex        =   7
         Top             =   690
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   635
         _Version        =   393216
         CalendarTitleBackColor=   14737632
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   143458305
         CurrentDate     =   38845
      End
      Begin ImpulseButton.ISButton CmdShowReport 
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   1110
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   661
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   "⁄—÷ «· ﬁ—Ì—"
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
         ButtonImage     =   "FrmBoxesAccounts.frx":038A
         ColorButton     =   14871017
         ColorShadow     =   4210752
         DrawFocusRectangle=   0   'False
         ColorTextShadow =   4210752
      End
      Begin VB.Label Lab 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   1710
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   330
         Width           =   345
      End
      Begin VB.Label Lab 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·Ï"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   1740
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   720
         Width           =   315
      End
   End
   Begin ImpulseButton.ISButton CmdExit 
      Height          =   435
      Left            =   60
      TabIndex        =   2
      Top             =   5160
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   767
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
      ButtonImage     =   "FrmBoxesAccounts.frx":0724
      ColorButton     =   14871017
      ColorShadow     =   4210752
      DrawFocusRectangle=   0   'False
      ColorTextShadow =   4210752
   End
   Begin VSFlex8UCtl.VSFlexGrid FgBoxes 
      Height          =   2775
      Left            =   -120
      TabIndex        =   0
      Top             =   600
      Width           =   8655
      _cx             =   15266
      _cy             =   4895
      Appearance      =   2
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
      BackColorFixed  =   -2147483633
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
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   9
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   320
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmBoxesAccounts.frx":0ABE
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
   Begin C1SizerLibCtl.C1Elastic EleHeader 
      Height          =   585
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   8535
      _cx             =   15055
      _cy             =   1032
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial (Arabic)"
         Size            =   24
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   0
      MousePointer    =   0
      Version         =   801
      BackColor       =   16777215
      ForeColor       =   4210688
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Picture         =   "FrmBoxesAccounts.frx":0C1E
      Caption         =   "—’Ìœ «·Œ“‰… «·√‰"
      Align           =   0
      AutoSizeChildren=   0
      BorderWidth     =   2
      ChildSpacing    =   1
      Splitter        =   0   'False
      FloodDirection  =   0
      FloodPercent    =   0
      CaptionPos      =   7
      WordWrap        =   -1  'True
      MaxChildSize    =   0
      MinChildSize    =   0
      TagWidth        =   0
      TagPosition     =   0
      Style           =   0
      TagSplit        =   2
      PicturePos      =   1
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
   End
   Begin ImpulseButton.ISButton CmdRefresh 
      Height          =   405
      Left            =   2220
      TabIndex        =   4
      Top             =   4500
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   714
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   " ÕœÌÀ «·»Ì«‰« "
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
      ButtonImage     =   "FrmBoxesAccounts.frx":18F8
      ColorButton     =   14871017
      ColorShadow     =   4210752
      DrawFocusRectangle=   0   'False
      ColorTextShadow =   4210752
   End
   Begin ImpulseButton.ISButton ISButton1 
      Height          =   375
      Left            =   4410
      TabIndex        =   12
      Top             =   4530
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   661
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "ÿ»«⁄… «·«—’œ…"
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
      ButtonImage     =   "FrmBoxesAccounts.frx":1C92
      ColorButton     =   14871017
      ColorShadow     =   4210752
      DrawFocusRectangle=   0   'False
      ColorTextShadow =   4210752
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   120
      X2              =   8520
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Label Lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "Ì„ﬂ‰ﬂ «‰  ⁄—÷ ﬂ‘› Õ”«» ··Œ“‰… «·ÌÊ„ » ÕœÌœ «”„ «·Œ“‰… À„ ≈÷€ÿ “— ﬂ‘› Õ”«» ··Œ“‰…"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   645
      Left            =   2460
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   3480
      Visible         =   0   'False
      Width           =   6045
   End
End
Attribute VB_Name = "FrmBoxesAccounts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkAll_Click()
Dim i As Long
For i = 1 To FgBoxes.Rows - 1
    If chkAll.value = vbChecked Then
        FgBoxes.TextMatrix(i, FgBoxes.ColIndex("Select")) = 1
    Else
        FgBoxes.TextMatrix(i, FgBoxes.ColIndex("Select")) = 0
    End If
Next

End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CmdRefresh_Click()
    ShowBoxesAccouns
End Sub

Private Sub ShowBoxesAccouns()
  Dim rs As ADODB.Recordset
    Dim StrSQL As String
    Dim Msg As String
    Dim i As Integer
    Dim FirstPeriod As Date
    Dim LastPeriod As Date
    Dim Balance As Double
    'On Error GoTo ErrTrap
    'StrSQL = "SELECT * from TblBoxesData where type=0 "
   StrSQL = "SELECT * from TblBoxesData  "
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.BOF Or rs.EOF) Then
        Load FrmBoxesAccounts

        With FrmBoxesAccounts.FgBoxes
            .Rows = .FixedRows + rs.RecordCount
            rs.MoveFirst

            For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("BoxID")) = IIf(IsNull(rs("BoxID").value), "", rs("BoxID").value)
            If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("BoxName")) = IIf(IsNull(rs("BoxName").value), "", rs("BoxName").value)
             Else
             .TextMatrix(i, .ColIndex("BoxName")) = IIf(IsNull(rs("BoxNamee").value), "", rs("BoxNamee").value)
             End If
                .TextMatrix(i, .ColIndex("AccountCode")) = IIf(IsNull(rs("Account_Code").value), "", rs("Account_Code").value)
      
                getFirstPeriodDateInthisYear FirstPeriod
                If Not IsNull(Me.DtpBoxFrom.value) And Not IsNull(Me.DtpBoxFrom.value) Then
                    FirstPeriod = DtpBoxFrom.value
                End If
                If Not IsNull(Me.DtpBoxTo.value) And Not IsNull(Me.DtpBoxTo.value) Then
                    LastPeriod = DtpBoxTo.value
                Else
                    LastPeriod = Date
                End If
                
                Balance = GetActualAccountBalance(rs("Account_Code").value, 0, FirstPeriod, LastPeriod)
            
                '        .TextMatrix(i, .ColIndex("BoxCredit")) = (get_balanceFromGl(rs("Account_Code").value))
                .TextMatrix(i, .ColIndex("BoxCredit")) = Abs(Balance) 'GetActualAccountBalance(rs("Account_Code").value, branch_id, FirstPeriod, Date)

                If SystemOptions.UserInterface = ArabicInterface Then
                    If Balance > 0 Then
                        .TextMatrix(i, .ColIndex("Type")) = "„œÌ‰"
                        .TextMatrix(i, .ColIndex("DebitValue")) = Balance
                    ElseIf Balance < 0 Then
                        .TextMatrix(i, .ColIndex("Type")) = "œ«∆‰"
                        .TextMatrix(i, .ColIndex("CreditValue")) = Balance
                    Else
            
                        .TextMatrix(i, .ColIndex("Type")) = " "
                    End If

                Else

                    If Balance > 0 Then
                        .TextMatrix(i, .ColIndex("Type")) = "Debit"
                        .TextMatrix(i, .ColIndex("DebitValue")) = Balance
                    ElseIf Balance < 0 Then
                        .TextMatrix(i, .ColIndex("Type")) = "Credit"
                        .TextMatrix(i, .ColIndex("CreditValue")) = Balance
                    Else
            
                        .TextMatrix(i, .ColIndex("Type")) = " "
                    End If

                End If
            
                rs.MoveNext
            Next i

            .AutoSize 0, .Cols - 1, False
        End With

    End If

    Exit Sub

    If SystemOptions.SysDataBaseType = AccessDataBase Then
        StrSQL = "SELECT TblBoxesData.BoxID, TblBoxesData.BoxName, QryBoxesCredit.BoxCredit " & " FROM TblBoxesData LEFT JOIN QryBoxesCredit ON TblBoxesData.BoxID =" & "QryBoxesCredit.BoxID "

        If SystemOptions.usertype = UserNormal Then
            StrSQL = StrSQL + " Where TblBoxesData.BoxID <>1"
        End If

    ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
        StrSQL = "SELECT dbo.TblBoxesData.BoxID,dbo.TblBoxesData.BoxName, QryBoxesCredit.BoxCredit" & " FROM dbo.TblBoxesData INNER JOIN " & "dbo.QryBoxesCredit() QryBoxesCredit ON dbo.TblBoxesData.BoxID = QryBoxesCredit.BoxID"

        If SystemOptions.usertype = UserNormal Then
            StrSQL = StrSQL + " Where dbo.TblBoxesData.BoxID <>1"
        End If
    End If

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.BOF Or rs.EOF) Then
        Load FrmBoxesAccounts

        With FrmBoxesAccounts.FgBoxes
            .Rows = .FixedRows + rs.RecordCount
            rs.MoveFirst

            For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("BoxID")) = IIf(IsNull(rs("BoxID").value), "", rs("BoxID").value)
                .TextMatrix(i, .ColIndex("BoxName")) = IIf(IsNull(rs("BoxName").value), "", rs("BoxName").value)

                If Not IsNull(rs("BoxCredit").value) Then
                    .TextMatrix(i, .ColIndex("BoxCredit")) = Format(rs("BoxCredit").value, SystemOptions.SysDefCurrencyForamt)
                Else
                    .TextMatrix(i, .ColIndex("BoxCredit")) = 0
                End If
            
                rs.MoveNext
            Next i

            .AutoSize 0, .Cols - 1, False
        End With

        FrmBoxesAccounts.show
        FrmBoxesAccounts.ZOrder 0
    Else
        Msg = "·«ÌÊÃœ «Ï Œ“‰ „”Ã·… ›Ï «·»—‰«„Ã"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Screen.MousePointer = vbDefault
    End If

    rs.Close
    Set rs = Nothing
    Exit Sub
ErrTrap:
    Msg = "⁄›Ê«·«Ì„ﬂ‰ ⁄—÷ «·Œ“‰ «·Õ«·Ì… ›Ï «·»—‰«„Ã...!!!"
    Msg = Msg & CHR(13) & "»—Ã«¡ «·√ ’«· »«·œ⁄„ «·›‰Ï."
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
End Sub

Private Sub CmdShowReport_Click()

    If DoPremis(Do_Print, Me.Name, True) = False Then
        Exit Sub
    End If
        
    With FgBoxes

        If Not .TextMatrix(.Row, .ColIndex("AccountCode")) = "" Then
            If Not IsNull(Me.DtpBoxFrom.value) And Not IsNull(Me.DtpBoxTo.value) Then
                ShowReport .TextMatrix(.Row, .ColIndex("AccountCode")), .TextMatrix(.Row, .ColIndex("BoxName")), Me.DtpBoxFrom.value, Me.DtpBoxTo.value
            ElseIf Not IsNull(Me.DtpBoxFrom.value) And IsNull(Me.DtpBoxTo.value) Then
                ShowReport .TextMatrix(.Row, .ColIndex("AccountCode")), .TextMatrix(.Row, .ColIndex("BoxName")), Me.DtpBoxFrom.value
            ElseIf IsNull(Me.DtpBoxFrom.value) And Not IsNull(Me.DtpBoxTo.value) Then
                ShowReport .TextMatrix(.Row, .ColIndex("AccountCode")), .TextMatrix(.Row, .ColIndex("BoxName")), , Me.DtpBoxTo.value
            Else
                ShowReport .TextMatrix(.Row, .ColIndex("AccountCode")), .TextMatrix(.Row, .ColIndex("BoxName"))
            End If
        End If

    End With

End Sub

Private Sub FgBoxes_DblClick()
    'ShowReport
    CmdShowReport_Click
End Sub

Private Sub Form_Activate()
    ShowDynamicHelp Me.HelpContextID
End Sub

Private Sub ChangeLang()
    'CmdConvert.Caption = "Convert to bill"
 
    Me.Caption = "Boxs Balancs"
    EleHeader.Caption = Me.Caption

    Fra(1).Caption = "Box Report"
    Lab(4).Caption = "From"
    Lab(3).Caption = "To"
 
    CmdShowReport.Caption = "View"

    CmdRefresh.Caption = "Refredh"

    With Me.FgBoxes
        .TextMatrix(0, .ColIndex("serial")) = "Index"
        .TextMatrix(0, .ColIndex("BoxID")) = "Box Code"
        .TextMatrix(0, .ColIndex("BoxName")) = "Box Name"
        .TextMatrix(0, .ColIndex("BoxCredit")) = "Balance"
        .TextMatrix(0, .ColIndex("Type")) = "Type"
        
    End With

    CmdExit.Caption = "Exit"
    
End Sub

Private Sub Form_Load()
    Dim GrdBack As New ClsBackGroundPic
    Resize_Form Me

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    Dim FirstPeriod As Date
    getFirstPeriodDateInthisYear FirstPeriod
 
    DtpBoxFrom.value = FirstPeriod
    DtpBoxTo.value = Date
    Set Me.FgBoxes.WallPaper = GrdBack.Picture
    ShowBoxesAccouns
    'SetDtpickerDate Me.DtpBoxFrom
    'SetDtpickerDate Me.DtpBoxTo

End Sub

'Public Sub ShowReport(Optional StrAccountCode As String, Optional StrAccountName As String, Optional fromdate As Date, Optional todate As Date)
'Dim cAccountReport As ClsAccReports
'                Set cAccountReport = New ClsAccReports
'                cAccountReport.BegineDate = fromdate '
'                cAccountReport.EndDate = todate
'                cAccountReport.ShowLedger StrAccountCode, _
'                StrAccountName
'                Set cAccountReport = Nothing
'

Private Sub ISButton1_Click()
Dim s As String



s = "Delete tblBoxesTemp "
Cn.Execute s

s = "Select * from tblBoxesTemp "
saveGrid s, FgBoxes, "Select", ""
'saveGrid StrSQL, fg3, "AccountCode", "", "UserId", val(Me.DcboUsers.BoundText)
print_reportBoxes
End Sub
Function print_reportBoxes(Optional NoteSerial As String)
    
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String
MySQL = "SELECT    * from tblBoxesTemp"


 
          StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "ReptblBoxesTemp.rpt"
     
    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open MySQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

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
        StrReportTitle = "" '& StrAccountName
    Else
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
      '  xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""
    End If
   '
    xReport.ParameterFields(3).AddCurrentValue user_name
     If Not DtpBoxFrom.value Then
     xReport.ParameterFields(8).AddCurrentValue DtpBoxFrom.value
     End If
     If Not DtpBoxTo.value Then
      xReport.ParameterFields(10).AddCurrentValue DtpBoxTo.value
    End If
    
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.title
    xReport.ReportAuthor = App.title
              Dim xLogo As CRAXDRT.OLEObject
         
    ''///////
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
End Function
'Exit Sub
'Dim Msg As String
'Dim cBoxReport As ClsBoxesReports
'Dim LngBoxID As Long
'With Me.FgBoxes
'    If .Row = -1 Then
'        Exit Sub
'    End If
'    If .Col = -1 Then
'        Exit Sub
'    End If
'    LngBoxID = Val(.TextMatrix(.Row, .ColIndex("BoxID")))
'    If LngBoxID > 0 Then
'        Set cBoxReport = New ClsBoxesReports
'        cBoxReport.BoxBalance LngBoxID, Me.DtpBoxFrom.value, Me.DtpBoxTo.value
'        Set cBoxReport = Nothing
'    End If
'End With
'End Sub
Private Sub lbl_Click()

End Sub
