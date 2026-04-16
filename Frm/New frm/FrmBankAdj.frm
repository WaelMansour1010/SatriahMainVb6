VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmBankAdj 
   Caption         =   " ﬁ—Ì— „—«Ã⁄Â »‰ﬂ „⁄Ì‰"
   ClientHeight    =   8160
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9855
   Icon            =   "FrmBankAdj.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   8160
   ScaleWidth      =   9855
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
      Height          =   8160
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   9855
      _cx             =   17383
      _cy             =   14393
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
      Align           =   5
      AutoSizeChildren=   8
      BorderWidth     =   1
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
      GridRows        =   4
      GridCols        =   4
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"FrmBankAdj.frx":000C
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin VSFlex8UCtl.VSFlexGrid Fg 
         Height          =   7185
         Left            =   15
         TabIndex        =   2
         Top             =   960
         Width           =   9825
         _cx             =   17330
         _cy             =   12674
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
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
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
         ExplorerBar     =   0
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
      Begin C1SizerLibCtl.C1Elastic EleHeader 
         Height          =   930
         Left            =   15
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   15
         Width           =   9825
         _cx             =   17330
         _cy             =   1640
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
         BorderWidth     =   6
         ChildSpacing    =   4
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
         Begin VB.CommandButton Command1 
            Caption         =   " ’œÌ—«·Ï «·«ﬂ”Ì·"
            Height          =   375
            Left            =   1320
            RightToLeft     =   -1  'True
            TabIndex        =   10
            Top             =   90
            Width           =   1545
         End
         Begin ImpulseButton.ISButton ISButton2 
            Height          =   375
            Left            =   210
            TabIndex        =   9
            Top             =   540
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   661
            Caption         =   "ÿ»«⁄…"
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
         Begin C1SizerLibCtl.C1Elastic Ele 
            Height          =   915
            Left            =   7650
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   -30
            Width           =   2145
            _cx             =   3784
            _cy             =   1614
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
            Caption         =   " ÕœÌœ «·› —…"
            Align           =   0
            AutoSizeChildren=   7
            BorderWidth     =   6
            ChildSpacing    =   4
            Splitter        =   0   'False
            FloodDirection  =   0
            FloodPercent    =   0
            CaptionPos      =   6
            WordWrap        =   -1  'True
            MaxChildSize    =   0
            MinChildSize    =   0
            TagWidth        =   0
            TagPosition     =   0
            Style           =   1
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
            Begin MSComCtl2.DTPicker DtpFrom 
               Height          =   345
               Left            =   60
               TabIndex        =   5
               Top             =   180
               Width           =   1515
               _ExtentX        =   2672
               _ExtentY        =   609
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   94437377
               CurrentDate     =   39561
            End
            Begin MSComCtl2.DTPicker DtpTo 
               Height          =   345
               Left            =   60
               TabIndex        =   6
               Top             =   540
               Width           =   1515
               _ExtentX        =   2672
               _ExtentY        =   609
               _Version        =   393216
               CheckBox        =   -1  'True
               Format          =   94437377
               CurrentDate     =   39561
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "≈·Ï"
               Height          =   315
               Index           =   1
               Left            =   1590
               RightToLeft     =   -1  'True
               TabIndex        =   8
               Top             =   510
               Width           =   405
            End
            Begin VB.Label lbl 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E2E9E9&
               Caption         =   "„‰"
               Height          =   315
               Index           =   0
               Left            =   1590
               RightToLeft     =   -1  'True
               TabIndex        =   7
               Top             =   210
               Width           =   405
            End
         End
         Begin ImpulseButton.ISButton ISButton1 
            Height          =   420
            Left            =   210
            TabIndex        =   3
            Top             =   90
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   741
            Caption         =   "⁄—÷"
            BackColor       =   12632256
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColorButton     =   12632256
         End
         Begin MSDataListLib.DataCombo dcBranch 
            Height          =   315
            Left            =   3480
            TabIndex        =   11
            Top             =   120
            Width           =   2985
            _ExtentX        =   5265
            _ExtentY        =   556
            _Version        =   393216
            BackColor       =   16777215
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin MSDataListLib.DataCombo Dcbank 
            Height          =   315
            Left            =   3480
            TabIndex        =   13
            Top             =   480
            Width           =   2985
            _ExtentX        =   5265
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
            RightToLeft     =   -1  'True
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·»‰ﬂ"
            Height          =   285
            Index           =   15
            Left            =   6840
            RightToLeft     =   -1  'True
            TabIndex        =   14
            Top             =   390
            Width           =   495
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Õœœ «·›—⁄"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   6645
            TabIndex        =   12
            Top             =   120
            Width           =   780
         End
      End
   End
End
Attribute VB_Name = "FrmBankAdj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type TransactionsValues
    TotalCash As Double
    TotalDue As Double
    TotalNet As Double
End Type

Public Sub ShowBoxesAccouns2()
    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    Dim Msg As String
    Dim i As Integer
    Dim FirstPeriod As Date
    Dim Balance As Double
    'On Error GoTo ErrTrap
    StrSQL = "SELECT * from TblBoxesData "
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
                .TextMatrix(i, .ColIndex("AccountCode")) = IIf(IsNull(rs("Account_Code").value), "", rs("Account_Code").value)
      
                getFirstPeriodDateInthisYear FirstPeriod
 
                Balance = GetActualAccountBalance(rs("Account_Code").value, branch_id, FirstPeriod, Date)
            
                '        .TextMatrix(i, .ColIndex("BoxCredit")) = (get_balanceFromGl(rs("Account_Code").value))
                .TextMatrix(i, .ColIndex("BoxCredit")) = Abs(Balance) 'GetActualAccountBalance(rs("Account_Code").value, branch_id, FirstPeriod, Date)

                If SystemOptions.UserInterface = ArabicInterface Then
                    If Balance > 0 Then
                        .TextMatrix(i, .ColIndex("Type")) = "„œÌ‰"
                    ElseIf Balance < 0 Then
                        .TextMatrix(i, .ColIndex("Type")) = "œ«∆‰"
                    Else
            
                        .TextMatrix(i, .ColIndex("Type")) = " "
                    End If

                Else

                    If Balance > 0 Then
                        .TextMatrix(i, .ColIndex("Type")) = "Debit"
                    ElseIf Balance < 0 Then
                        .TextMatrix(i, .ColIndex("Type")) = "Credit"
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
 
    rs.Close
    Set rs = Nothing
    Exit Sub
ErrTrap:
    Msg = "⁄›Ê«·«Ì„ﬂ‰ ⁄—÷ «·Œ“‰ «·Õ«·Ì… ›Ï «·»—‰«„Ã...!!!"
    Msg = Msg & Chr(13) & "»—Ã«¡ «·√ ’«· »«·œ⁄„ «·›‰Ï."
    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title

End Sub

Private Sub Command1_Click()
    On Error Resume Next
    Dim StrFileName As String
    StrFileName = App.path & "\Temp1.xls"

    If Dir(StrFileName) <> "" Then
        Kill StrFileName
    End If

    Me.Fg.SaveGrid StrFileName, flexFileExcel, True
    OpenFile StrFileName
End Sub

Private Sub Dcbank_Click(Area As Integer)
If Dcbank.text = "" Then Exit Sub
ISButton1_Click
End Sub

Private Sub Fg_Click()
    Dim LngMouseRow As Long
    Dim LngMouseCol As Long

    Dim FgNode As VSFlex8UCtl.VSFlexNode
    LngMouseRow = Me.Fg.MouseRow

    If (LngMouseRow < 0 Or Fg.Rows = 0) Then
        Exit Sub
    End If

    If (Fg.IsSubtotal(LngMouseRow) = True And LngMouseCol = 0) Then
        Set FgNode = Fg.GetNode(LngMouseRow)
        FgNode.Expanded = Not FgNode.Expanded
    
    End If

    'With FG
    '   If Not .TextMatrix(.Row, .ColIndex("AccountCode")) = "" Then
    '       If Not IsNull(Me.DTPFrom.value) And Not IsNull(Me.DTPTo.value) Then
    '      ShowReport .TextMatrix(.Row, .ColIndex("AccountCode")), .TextMatrix(.Row, .ColIndex("BoxName")), Me.DTPFrom.value, Me.DTPTo.value
    '      ElseIf Not IsNull(Me.DTPFrom.value) And IsNull(Me.DTPTo.value) Then
    '      ShowReport .TextMatrix(.Row, .ColIndex("AccountCode")), .TextMatrix(.Row, .ColIndex("BoxName")), Me.DTPFrom.value
    '      ElseIf IsNull(Me.DTPFrom.value) And Not IsNull(Me.DTPTo.value) Then
    '      ShowReport .TextMatrix(.Row, .ColIndex("AccountCode")), .TextMatrix(.Row, .ColIndex("BoxName")), , Me.DTPFrom.value
    '      Else
    '     ShowReport .TextMatrix(.Row, .ColIndex("AccountCode")), .TextMatrix(.Row, .ColIndex("BoxName"))
    '      End If
    '  End If
    'End With

End Sub

Private Sub Fg_MouseMove(Button As Integer, _
                         Shift As Integer, _
                         x As Single, _
                         Y As Single)
    Dim LngMouseRow As Long
    Dim LngMouseCol As Long

    LngMouseRow = Me.Fg.MouseRow
    LngMouseCol = Me.Fg.MouseCol

    If (LngMouseRow < 0) Or (LngMouseCol < 0) Or Fg.Rows = 0 Then
        Fg.MousePointer = flexDefault
        Exit Sub
    End If

    If Fg.IsSubtotal(LngMouseRow) = True Then
        Fg.MousePointer = flexHand
    Else
        Fg.MousePointer = flexDefault
    End If

    WriteToolTip LngMouseRow, LngMouseCol

End Sub

Private Sub Form_Load()
    Dim GrdBack As ClsBackGroundPic
  Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500
       
    With Me.Fg
        .Clear flexClearScrollable, flexClearEverything
        .Cols = 10
        .Rows = 0
        .FixedCols = 0
        .FixedRows = 0
        .MergeCells = flexMergeOutline
        .ColAlignment(0) = flexAlignRightCenter
        .RowHeightMin = 320
        .ExplorerBar = flexExNone
        ' appearance
        .GridLines = flexGridNone
        .BackColorBkg = .backcolor
        .SheetBorder = .backcolor
        .ExtendLastCol = True
        .Redraw = flexRDBuffered ' << new setting
        .OutlineCol = 0
        .OutlineBar = flexOutlineBarCompleteLeaf
        '.NodeClosedPicture = MDIFrmMain.ImgLstTree.ListImages("Close").Picture
        '.NodeOpenPicture = MDIFrmMain.ImgLstTree.ListImages("OpenFolder").Picture
        .Ellipsis = flexEllipsisEnd
        'Set Grdback = New ClsBackGroundPic
        'Set .WallPaper = Grdback.Picture
        ' behavior
        .AllowSelection = False
        .Highlight = flexHighlightWithFocus
        .ScrollTrack = True
        .AutoSearch = flexSearchFromCursor
    End With

    'SetDtpickerDate Me.DtpFrom
    'SetDtpickerDate Me.DtpTo

    Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos
    Dcombos.GetBranches dcBranch
 
    Dcombos.GetBanks Me.Dcbank

    dcBranch.BoundText = Current_branch

    Dim FirstPeriodDateInthisYear  As Date
    getFirstPeriodDateInthisYear FirstPeriodDateInthisYear
    Me.DtpFrom = FirstPeriodDateInthisYear
    Me.DtpTo = Date

    Me.Width = 11000
    Me.Height = 9000
    Resize_Form Me
    Cn.CommandTimeout = 180
    
       If SystemOptions.UserInterface = EnglishInterface Then
     
        SetInterface Me
        ChangeLang
    End If
    
    
End Sub
Private Sub ChangeLang()
Me.Caption = "Bank Adj"
Ele.Caption = "Period"
lbl(0).Caption = "From"
lbl(1).Caption = "To"
Label3.Caption = "Branch"
lbl(15).Caption = "Bank"
Command1.Caption = "Excel Export"
ISButton1.Caption = "View"
ISButton2.Caption = "Print"
End Sub
Private Sub ISButton1_Click()
    Dim Msg As String

    If Trim(Me.Dcbank.BoundText) = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Msg = "Õœœ »‰ﬂ «Ê·« "
        Else
            Msg = "Specify Bank.!!"
        End If

        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Dcbank.SetFocus
        SendKeys "{F4}"
        Exit Sub
    End If

    LoadData
End Sub

Private Sub LoadData()
    Dim XFont As IFontDisp
    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    Dim i As Integer, J As Integer
    Dim IntStartSelect As Integer, IntEndSelect As Integer
    Dim SngTempValue As Double
    Dim StrOneRowData As String
    Dim SngHeaderBackColor As Single
    Dim SngDataBackColor As Single
    Dim StrStartDate As String
    Dim SngTemp1 As Double, SngTemp2 As Double, SngTemp3 As Double
    Dim Boxname As String, BoxBalance As Double, balancetype As String
    Dim TransValues As TransactionsValues
    Dim LngCustomersCount As Double, LngTempRow As Double

    SngHeaderBackColor = &HC0C0C0
    SngDataBackColor = &HE2E9E9
    Dim Account_Code As String
    Dim Account_code1 As String
    Dim Account_code2 As String

    Dim Account_codeBalance As Double
    Dim Account_code1Balance As Double
    Dim Account_code2Balance As Double
    Dim FirstPeriod As Date
    Account_Code = ModAccounts.GetMyAccountCodeRefined("BanksData", "BankId", val(Me.Dcbank.BoundText), "Account_code")
    Account_code1 = ModAccounts.GetMyAccountCodeRefined("BanksData", "BankId", val(Me.Dcbank.BoundText), "Account_code1")
    Account_code2 = ModAccounts.GetMyAccountCodeRefined("BanksData", "BankId", val(Me.Dcbank.BoundText), "Account_code2")
    'getFirstPeriodDateInthisYear FirstPeriod
    Account_codeBalance = GetActualAccountBalance(Account_Code, Current_branch, DtpFrom.value, DtpTo.value)
    Account_code1Balance = GetActualAccountBalance(Account_code1, Current_branch, DtpFrom.value, DtpTo.value)
    Account_code2Balance = GetActualAccountBalance(Account_code2, Current_branch, DtpFrom.value, DtpTo.value)
    
    Dim CurrentText As String

    With Me.Fg
        .Clear flexClearScrollable, flexClearEverything
        .Rows = 0
        .OutlineBar = flexOutlineBarComplete
        .MergeCells = flexMergeOutline
        If SystemOptions.UserInterface = ArabicInterface Then
        StrOneRowData = "«· ﬁ—Ì— «·„Ã„⁄ ··»‰ﬂ  ⁄‰ «·› —… "
       Else
               StrOneRowData = "Bank Report  Period "
       End If
        If Not IsNull(Me.DtpFrom.value) Then
                If SystemOptions.UserInterface = ArabicInterface Then

            StrOneRowData = StrOneRowData & "„‰ " & DisplayDate(Me.DtpFrom.value)
            Else
            StrOneRowData = StrOneRowData & "From " & DisplayDate(Me.DtpFrom.value)
            End If
        End If

        If Not IsNull(Me.DtpTo.value) Then
        If SystemOptions.UserInterface = ArabicInterface Then
            StrOneRowData = StrOneRowData & " ≈·Ï " & DisplayDate(Me.DtpTo.value)
       Else
       StrOneRowData = StrOneRowData & " To " & DisplayDate(Me.DtpTo.value)
       End If
        End If

        .AddItem StrOneRowData
        .RowOutlineLevel(0) = 1
        .IsSubtotal(.Rows - 1) = True
        .RowHeight(.Rows - 1) = 450
        .Cell(flexcpFontBold, .Rows - 1, 0) = True
        Set XFont = Me.Font
        XFont.name = "Tahoma"
        XFont.Size = 12
        XFont.Charset = 178
        .Cell(flexcpFont, .Rows - 1, 0) = XFont
        '-------------------------------------------------
        GoTo ll
                If SystemOptions.UserInterface = ArabicInterface Then

        .AddItem "«·«—’œ…"
        Else
        .AddItem "Balances"
        End If
        .RowOutlineLevel(.Rows - 1) = 2
        .IsSubtotal(.Rows - 1) = True
        '     TransValues = GetTransactionsValues("2 Or transaction_type = 21")
                If SystemOptions.UserInterface = ArabicInterface Then
        
        CurrentText = " —’Ìœ «·»‰ﬂ ›Ï «·› —…:" & vbTab & FormatNumber(Abs(Account_codeBalance), 2, vbUseDefault, , vbTrue)
Else
        CurrentText = "Bank Balance" & vbTab & FormatNumber(Abs(Account_codeBalance), 2, vbUseDefault, , vbTrue)
End If
        If Account_codeBalance >= 0 Then
           If SystemOptions.UserInterface = ArabicInterface Then
            CurrentText = CurrentText & "  „œÌ‰  "
            Else
            CurrentText = CurrentText & "  Depit  "
            End If
        Else
                   If SystemOptions.UserInterface = ArabicInterface Then

            CurrentText = CurrentText & "   œ«∆‰  "
            Else
            CurrentText = CurrentText & "   Credit  "
            End If
        End If
            
        .AddItem CurrentText
            
        .RowOutlineLevel(.Rows - 1) = 3
        .IsSubtotal(.Rows - 1) = False
        .Rowdata(.Rows - 1) = "Items1"
                            If SystemOptions.UserInterface = ArabicInterface Then
   
        CurrentText = "—’Ìœ «·‘Ìﬂ«   Õ  «· Õ’Ì·:" & vbTab & FormatNumber(Abs(Account_code1Balance), 2, vbUseDefault, , vbTrue)
Else
        CurrentText = "Under Collection Balance:" & vbTab & FormatNumber(Abs(Account_code1Balance), 2, vbUseDefault, , vbTrue)

End If
        If Account_code1Balance >= 0 Then
                                    If SystemOptions.UserInterface = ArabicInterface Then

            CurrentText = CurrentText & "   „œÌ‰  "
            Else
            CurrentText = CurrentText & "   Depit  "
            End If
        Else
        If SystemOptions.UserInterface = ArabicInterface Then
            CurrentText = CurrentText & "   œ «∆‰  "
        Else
        CurrentText = CurrentText & "  Credit  "
        End If
        End If
                      
        .RowOutlineLevel(.Rows - 1) = 3
        .IsSubtotal(.Rows - 1) = False
                
        .AddItem CurrentText
            
        .RowOutlineLevel(.Rows - 1) = 3
        .IsSubtotal(.Rows - 1) = False
                If SystemOptions.UserInterface = ArabicInterface Then

        CurrentText = "  —’Ìœ «·‘Ìﬂ«    «·„ƒÃ·…:" & vbTab & FormatNumber(Abs(Account_code2Balance), 2, vbUseDefault, , vbTrue)
Else
        CurrentText = "  Post-dated checks Balance:" & vbTab & FormatNumber(Abs(Account_code2Balance), 2, vbUseDefault, , vbTrue)
End If
        If Account_code2Balance >= 0 Then
                                   If SystemOptions.UserInterface = ArabicInterface Then

            CurrentText = CurrentText & "   „œÌ‰  "
            Else
            CurrentText = CurrentText & "   Depit  "
            End If
        Else
                                             If SystemOptions.UserInterface = ArabicInterface Then

            CurrentText = CurrentText & "   œ«∆‰  "
            Else
            CurrentText = CurrentText & "   Credit  "
            End If
        End If
                      
        .AddItem CurrentText
        '  .RowOutlineLevel(.Rows - 1) = 3
        '  .IsSubtotal(.Rows - 1) = False
            
      If SystemOptions.UserInterface = ArabicInterface Then
        .AddItem " «·«Ìœ«⁄«  «·‰ﬁœÌ… "
       Else
       .AddItem "Cash Deposits "
       End If
        .RowOutlineLevel(.Rows - 1) = 3
        .IsSubtotal(.Rows - 1) = True
        LngTempRow = .Rows - 1
        'LngCustomersCount = 0 ' LoadTransCustomers(2, TransValues.TotalNet)
        LngCustomersCount = BankDepositeData(val(Me.Dcbank.BoundText), 0, val(Me.dcBranch.BoundText), DtpFrom.value, DtpTo.value)
        .TextMatrix(LngTempRow, 0) = .TextMatrix(LngTempRow, 0) & "  :  " & LngCustomersCount
            
        LoadTransItems 2
            
        TransValues = LoadSalTypeTrans
        .AddItem "«·„»Ì⁄«  «·ﬁÿ«⁄Ï:" & vbTab & TransValues.TotalCash
        .RowOutlineLevel(.Rows - 1) = 3
        .IsSubtotal(.Rows - 1) = True
        .AddItem "»Ì«‰«  «·⁄„·«¡ «·ﬁÿ«⁄Ï Ê„”ÕÊ»« Â„" & ""
        .RowOutlineLevel(.Rows - 1) = 4
        .IsSubtotal(.Rows - 1) = True
            
        .AddItem "«·„»Ì⁄«  «· Ã«—Ï:" & vbTab & TransValues.TotalDue
        .RowOutlineLevel(.Rows - 1) = 3
        .IsSubtotal(.Rows - 1) = True
        .AddItem "»Ì«‰«  «·⁄„·«¡ «· Ã«—Ï Ê„”ÕÊ»« Â„" & ""
        .RowOutlineLevel(.Rows - 1) = 4
        .IsSubtotal(.Rows - 1) = True
        'Exit Sub
        .AddItem "«·„‘ —Ì« "
        .RowOutlineLevel(.Rows - 1) = 2
        .IsSubtotal(.Rows - 1) = True
        TransValues = GetTransactionsValues("1 Or transaction_type = 22")
        .AddItem "≈Ã„«·Ï «·„‘ —Ì«  ›Ï «·› —…:" & vbTab & TransValues.TotalNet
        .RowOutlineLevel(.Rows - 1) = 3
        .IsSubtotal(.Rows - 1) = False
        .Rowdata(.Rows - 1) = "Items1"
    
        .AddItem "«·„‘ —Ì«  «·‰ﬁœÌ…:" & vbTab & TransValues.TotalCash
        .RowOutlineLevel(.Rows - 1) = 3
        .IsSubtotal(.Rows - 1) = False
            
        .AddItem "«·„‘ —Ì«  «·√Ã·…:" & vbTab & TransValues.TotalDue
        .RowOutlineLevel(.Rows - 1) = 3
        .IsSubtotal(.Rows - 1) = False
            
        .AddItem "«·„Ê—œÌ‰"
        .RowOutlineLevel(.Rows - 1) = 3
        .IsSubtotal(.Rows - 1) = True
        LngTempRow = .Rows - 1
        LngCustomersCount = LoadTransCustomers(1, TransValues.TotalNet)
        .TextMatrix(LngTempRow, 0) = .TextMatrix(LngTempRow, 0) & ":" & LngCustomersCount
            
        '       LoadTransItems 1
            
        .AddItem "„— Ã⁄ «·„»Ì⁄« "
        .RowOutlineLevel(.Rows - 1) = 2
        .IsSubtotal(.Rows - 1) = True
        TransValues = GetTransactionsValues(9)
        .AddItem "≈Ã„«·Ï „— Ã⁄ «·„»Ì⁄«  ›Ï «·› —…:" & vbTab & TransValues.TotalNet
        .RowOutlineLevel(.Rows - 1) = 3
        .IsSubtotal(.Rows - 1) = False
        .Rowdata(.Rows - 1) = "Items1"
    
        .AddItem "„— Ã⁄ «·„»Ì⁄«  «·‰ﬁœÌ…:" & vbTab & TransValues.TotalCash
        .RowOutlineLevel(.Rows - 1) = 3
        .IsSubtotal(.Rows - 1) = False
            
        .AddItem "„— Ã⁄ «·„»Ì⁄«  «·√Ã·…:" & vbTab & TransValues.TotalDue
        .RowOutlineLevel(.Rows - 1) = 3
        .IsSubtotal(.Rows - 1) = False
            
        .AddItem "«·⁄„·«¡"
        .RowOutlineLevel(.Rows - 1) = 3
        .IsSubtotal(.Rows - 1) = True
        LngTempRow = .Rows - 1
        LngCustomersCount = LoadTransCustomers(9, TransValues.TotalNet)
        .TextMatrix(LngTempRow, 0) = .TextMatrix(LngTempRow, 0) & ":" & LngCustomersCount
        '  LoadTransItems 9
        '-----------------------------------------------------------------------------------------
        .AddItem "„— Ã⁄ «·„‘ —Ì« "
        .RowOutlineLevel(.Rows - 1) = 2
        .IsSubtotal(.Rows - 1) = True
        TransValues = GetTransactionsValues(5)
        .AddItem "≈Ã„«·Ï „— Ã⁄ «·„‘ —Ì«  ›Ï «·› —…:" & vbTab & TransValues.TotalNet
        .RowOutlineLevel(.Rows - 1) = 3
        .IsSubtotal(.Rows - 1) = False
        .Rowdata(.Rows - 1) = "Items1"
    
        .AddItem "„— Ã⁄ «·„‘ —Ì«  «·‰ﬁœÌ…:" & vbTab & TransValues.TotalCash
        .RowOutlineLevel(.Rows - 1) = 3
        .IsSubtotal(.Rows - 1) = False
            
        .AddItem "„— Ã⁄ «·„‘ —Ì«  «·√Ã·…:" & vbTab & TransValues.TotalDue
        .RowOutlineLevel(.Rows - 1) = 3
        .IsSubtotal(.Rows - 1) = False
        .AddItem "«·„Ê—œÌ‰"
        .RowOutlineLevel(.Rows - 1) = 3
        .IsSubtotal(.Rows - 1) = True
        LngTempRow = .Rows - 1
        LngCustomersCount = LoadTransCustomers(5, TransValues.TotalNet)
        .TextMatrix(LngTempRow, 0) = .TextMatrix(LngTempRow, 0) & ":" & LngCustomersCount
        '     LoadTransItems 5
        '------------------------------------------------------------------------------------------
        '        .AddItem "„Œ“Ê‰ «·»÷«⁄…"
        '            .RowOutlineLevel(.Rows - 1) = 2
        '            .IsSubtotal(.Rows - 1) = True
        '
        '            .AddItem "ﬁÌ„… «·„Œ“Ê‰ «Ê· «·› —…"
        '            .RowOutlineLevel(.Rows - 1) = 3
        '            .IsSubtotal(.Rows - 1) = False
        '
        '            .AddItem "ﬁÌ„… «·„Œ“Ê‰ ‰Â«Ì… «·› —…"
        '            .RowOutlineLevel(.Rows - 1) = 3
        '            .IsSubtotal(.Rows - 1) = False
        '
        '            .AddItem "«·√’‰«› «·ÃœÌœ… «· Ï ≈÷Ì›  Œ·«· «·› —…"
        '            .RowOutlineLevel(.Rows - 1) = 3
        '            .IsSubtotal(.Rows - 1) = False
            
        '-------------------------------------------------------------------------------

        .AddItem "«·Œ“‰"
        .RowOutlineLevel(.Rows - 1) = 2
        .IsSubtotal(.Rows - 1) = True
        Set rs = New ADODB.Recordset

        If Not (IsNull(Me.DtpFrom.value)) Then
            StrSQL = " SELECT SUM(Note_Value * TransDir) AS BoxAccount "
            StrSQL = StrSQL & " FROM dbo.QryBoxBalance() QryBoxBalance "
            StrSQL = StrSQL + " Where QryBoxBalance.NoteDate <" & SQLDate(Me.DtpFrom.value, True)
            rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

            If Not (rs.BOF Or rs.EOF) Then
                SngTempValue = IIf(IsNull(rs("BoxAccount").value), 0, rs("BoxAccount").value)
            End If

        Else
            SngTempValue = 0
        End If

        '    .AddItem "≈Ã„«·Ï —’Ìœ «·Œ“‰ «Ê· «·› —… :" & SngTempValue
        .AddItem " "
             
        .RowOutlineLevel(.Rows - 1) = 3
        .IsSubtotal(.Rows - 1) = False
        StrSQL = " SELECT SUM(Note_Value * TransDir) AS BoxAccount "
        StrSQL = StrSQL & " FROM dbo.QryBoxBalance() QryBoxBalance "

        If Not (IsNull(Me.DtpTo.value)) Then
            StrSQL = StrSQL + " Where QryBoxBalance.NoteDate <" & SQLDate(Me.DtpTo.value, True)
        End If

        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        SngTempValue = 0

        If Not (rs.BOF Or rs.EOF) Then
            SngTempValue = IIf(IsNull(rs("BoxAccount").value), 0, rs("BoxAccount").value)
        End If

        '    .AddItem "≈Ã„«·Ï —’Ìœ «·Œ“‰ ‰Â«Ì… «·› —… :" & vbTab & SngTempValue
        .AddItem " " & vbTab
        .RowOutlineLevel(.Rows - 1) = 3
        .IsSubtotal(.Rows - 1) = False
            
        '  If IsNull(Me.DtpFrom.value) Then
        '      StrStartDate = SQLDate(CDate("01/01/1900"), True)
        '  Else
        '      StrStartDate = SQLDate(Me.DtpFrom.value, True)
        '  End If
        '
        '  StrSQL = "SELECT dbo.TblBoxesData.BoxID, dbo.TblBoxesData.BoxName, "
        '  StrSQL = StrSQL + " dbo.QryBoxCreditUptoDate(dbo.TblBoxesData.BoxID," & StrStartDate & ") AS StartBal,"
        '  StrSQL = StrSQL + " Convert(Decimal(38,2),SUM(CASE TransDir WHEN 1 THEN  Note_Value ELSE 0 END)) AS SumIn "
        '  StrSQL = StrSQL + ",Convert(Decimal(38,2),SUM(CASE TransDir WHEN -1 THEN  Note_Value ELSE 0 END)) AS SumOut"
        '  StrSQL = StrSQL + " FROM         dbo.TblBoxesData INNER JOIN dbo.QryBoxBalance() QryBoxBalance ON " & _
        '  "dbo.TblBoxesData.BoxID = QryBoxBalance.BoxID "
        '  StrSQL = StrSQL + " Where dbo.TblBoxesData.BoxID <> 0"
        '  If Not IsNull(Me.DtpFrom.value) Then
        '      StrSQL = StrSQL + " AND  QryBoxBalance.NoteDate >=" & SQLDate(Me.DtpFrom, True) & ""
        '  End If
        '  If Not IsNull(Me.DtpTo.value) Then
        '      StrSQL = StrSQL + " AND  QryBoxBalance.NoteDate <=" & SQLDate(Me.DtpTo.value, True) & ""
        '  End If
        '  StrSQL = StrSQL + " Group BY dbo.TblBoxesData.BoxID, dbo.TblBoxesData.BoxName "
            
        StrSQL = "SELECT * from TblBoxesData"
            
        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        .AddItem " ›«’Ì· √—’œ… «·Œ“‰"
        .RowOutlineLevel(.Rows - 1) = 3
        .IsSubtotal(.Rows - 1) = False

        If Not (rs.BOF Or rs.EOF) Then
            StrOneRowData = "«”„ «·Œ“‰…" & vbTab & "«·—’Ìœ " & vbTab & "  ÿ»Ì⁄Â «·—’Ìœ" & vbTab & "   "
            StrOneRowData = StrOneRowData & vbTab & " „·«ÕŸ« "
            .AddItem StrOneRowData
            .Cell(flexcpBackColor, .Rows - 1, 0, .Rows - 1, 4) = SngHeaderBackColor
            IntStartSelect = 0: IntEndSelect = 0
            IntStartSelect = .Rows
                
            rs.MoveFirst
                      
            Dim Balance As Double

            For i = 0 To rs.RecordCount - 1
                StrOneRowData = ""
                '     SngTemp1 = IIf(IsNull(rs("BoxName").value), 0, rs("StartBal").value)
                '     SngTemp2 = IIf(IsNull(rs("SumIn").value), 0, rs("SumIn").value)
                '     SngTemp3 = IIf(IsNull(rs("SumOut").value), 0, rs("SumOut").value)
                    
                getFirstPeriodDateInthisYear FirstPeriod
 
                'Balance = GetActualAccountBalance(rs("Account_Code").value, branch_id, DTPFrom.value, DTPTo.value)
                Balance = GetActualAccountBalance(rs("Account_Code").value, , DtpFrom.value, DtpTo.value)
                BoxBalance = Abs(Balance)

                If SystemOptions.UserInterface = ArabicInterface Then
                    If Balance > 0 Then
                        balancetype = "„œÌ‰"
                    ElseIf Balance < 0 Then
                        balancetype = "œ«∆‰"
                    Else
            
                        balancetype = ""
                    End If

                Else

                    If Balance > 0 Then
                        balancetype = "Debit"
                    ElseIf Balance < 0 Then
                        .balancetype = "Credit"
                    Else
            
                        balancetype = " "
                    End If

                End If
            
                '
                '             SngTemp3 = ""
                    
                StrOneRowData = rs("BoxName").value & vbTab & BoxBalance & vbTab & balancetype & vbTab  '& SngTemp3 & vbTab & ((SngTemp1 + SngTemp2) - SngTemp3)
                .AddItem StrOneRowData
                rs.MoveNext
            Next i

            IntEndSelect = .Rows - 1
            .Cell(flexcpBackColor, IntStartSelect, 0, .Rows - 1, 4) = SngDataBackColor
        End If

        '-------------------------------------------------------------------------------
ll:
        CurrentText = " —’Ìœ «·»‰ﬂ ›Ï «·› —…:" & vbTab & FormatNumber(Abs(Account_codeBalance), 2, vbUseDefault, , vbTrue)

        If Account_codeBalance >= 0 Then
            CurrentText = CurrentText & "  „œÌ‰  "
        Else
            CurrentText = CurrentText & "   œ«∆‰  "
        End If
            
        .AddItem CurrentText
             
        CurrentText = "—’Ìœ «·‘Ìﬂ«   Õ  «· Õ’Ì·:" & vbTab & FormatNumber(Abs(Account_code1Balance), 2, vbUseDefault, , vbTrue)
           
        If Account_code1Balance >= 0 Then
            CurrentText = CurrentText & "   „œÌ‰  "
        Else
            CurrentText = CurrentText & "   œ «∆‰  "
        End If
                 
        .AddItem CurrentText
            
        CurrentText = "  —’Ìœ «·‘Ìﬂ«    «·„ƒÃ·…:" & vbTab & FormatNumber(Abs(Account_code2Balance), 2, vbUseDefault, , vbTrue)

        If Account_code2Balance >= 0 Then
            CurrentText = CurrentText & "     „œÌ‰"
        Else
            CurrentText = CurrentText & "       œ«∆‰"
        End If
                      
        .AddItem CurrentText

        .AddItem "«·«Ìœ«⁄«  «·‰ﬁœÌ…"
        .RowOutlineLevel(.Rows - 1) = 2
        .IsSubtotal(.Rows - 1) = True
        SngTempValue = BankDepositeData(val(Me.Dcbank.BoundText), 0, val(Me.dcBranch.BoundText), DtpFrom.value, DtpTo.value)
        .AddItem "≈Ã„«·Ï «Ìœ«⁄«  «·‰ﬁœÌ…  ›Ì «·ﬁ —… : " & vbTab & SngTempValue
             
        .RowOutlineLevel(.Rows - 1) = 3
        .IsSubtotal(.Rows - 1) = False
        .MergeRow(.Rows - 1) = True
            
        StrSQL = "SELECT     TOP 100 PERCENT dbo.TblBanksDeposite.RecordDate, dbo.TblBanksDepositeDetails.[value], dbo.TblBoxesData.BoxName, dbo.TblBanksDeposite.Remarks"
        StrSQL = StrSQL & " FROM         dbo.TblBanksDeposite INNER JOIN"
        StrSQL = StrSQL & " dbo.TblBanksDepositeDetails ON dbo.TblBanksDeposite.id = dbo.TblBanksDepositeDetails.TblBanksDepositeId LEFT OUTER JOIN"
        StrSQL = StrSQL & " dbo.TblBoxesData ON dbo.TblBanksDepositeDetails.BoxID = dbo.TblBoxesData.BoxID"
        StrSQL = StrSQL & " Where (dbo.TblBanksDeposite.BankID = " & val(Dcbank.BoundText) & ") And (dbo.TblBanksDepositeDetails.box_or_bank = 0)"
        StrSQL = StrSQL & "  AND (dbo.TblBanksDeposite.branch_no = " & val(Me.dcBranch.BoundText) & ")"
    
        If Not (IsNull(Me.DtpFrom.value)) Then
            StrSQL = StrSQL + " AND dbo.TblBanksDeposite.RecordDate >=" & SQLDate(Me.DtpFrom.value, True)
        End If

        If Not (IsNull(Me.DtpTo.value)) Then
            StrSQL = StrSQL + " AND dbo.TblBanksDeposite.RecordDate<=" & SQLDate(Me.DtpTo.value, True)
        End If
    
        StrSQL = StrSQL & " ORDER BY dbo.TblBanksDeposite.RecordDate, dbo.TblBanksDeposite.NoteID"
            
        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        .AddItem " ›«’Ì· «·«Ìœ«⁄«  «·‰ﬁœÌ… ›Ï «·› —…"
        .RowOutlineLevel(.Rows - 1) = 3
        .IsSubtotal(.Rows - 1) = False

        If Not (rs.BOF Or rs.EOF) Then
            StrOneRowData = "  «· «—ÌŒ" & vbTab & "«·Œ“Ì‰…" & vbTab & "«·ﬁÌ„…" & vbTab & vbTab & vbTab & "„·«ÕŸ« "  ' & vbTab & "«·‰”»… „‰ ≈Ã„«·Ï «·«Ìœ«⁄«  «·‰ﬁœÌ… «·› —…"
            .AddItem StrOneRowData
            .Cell(flexcpBackColor, .Rows - 1, 0, .Rows - 1, 5) = SngHeaderBackColor
            IntStartSelect = 0: IntEndSelect = 0
            IntStartSelect = .Rows
                
            rs.MoveFirst

            For i = 0 To rs.RecordCount - 1
                StrOneRowData = ""
                StrOneRowData = rs("RecordDate").value & vbTab & rs("BoxName").value & vbTab & rs("value").value & vbTab & vbTab & vbTab & rs("Remarks").value
                .AddItem StrOneRowData
                '        .Cell(flexcpFloodPercent, .Rows - 1, 4) = 100 * Val(.TextMatrix(.Rows - 1, 2)) / SngTempValue
                rs.MoveNext
            Next i

            IntEndSelect = .Rows - 1
            .Cell(flexcpBackColor, IntStartSelect, 0, .Rows - 1, 5) = SngDataBackColor
        End If
        
        .AddItem "  «Ìœ«⁄ «·‘Ìﬂ« "
        .RowOutlineLevel(.Rows - 1) = 2
        .IsSubtotal(.Rows - 1) = True
        SngTempValue = BankDepositeData(val(Me.Dcbank.BoundText), 1, val(Me.dcBranch.BoundText), DtpFrom.value, DtpTo.value)
        .AddItem "≈Ã„«·Ï «Ìœ«⁄ «·‘Ìﬂ«   ›Ì «·ﬁ —… : " & vbTab & SngTempValue
             
        .RowOutlineLevel(.Rows - 1) = 3
        .IsSubtotal(.Rows - 1) = False
        .MergeRow(.Rows - 1) = True
            
        StrSQL = "SELECT     TOP 100 PERCENT dbo.TblBanksDeposite.RecordDate, dbo.TblBanksDepositeDetails.[value], dbo.TblBoxesData.BoxName, dbo.TblBanksDeposite.Remarks , dbo.TblBanksDepositeDetails.ChequeNo, dbo.TblBanksDepositeDetails.BankName, dbo.TblBanksDepositeDetails.DueDate "
        StrSQL = StrSQL & "  FROM         dbo.TblBanksDeposite INNER JOIN"
        StrSQL = StrSQL & " dbo.TblBanksDepositeDetails ON dbo.TblBanksDeposite.id = dbo.TblBanksDepositeDetails.TblBanksDepositeId LEFT OUTER JOIN"
        StrSQL = StrSQL & " dbo.TblBoxesData ON dbo.TblBanksDepositeDetails.BoxID = dbo.TblBoxesData.BoxID"
        StrSQL = StrSQL & " Where (dbo.TblBanksDeposite.BankID = " & val(Dcbank.BoundText) & ") And (dbo.TblBanksDepositeDetails.box_or_bank = 1)"
        StrSQL = StrSQL & "  AND (dbo.TblBanksDeposite.branch_no = " & val(Me.dcBranch.BoundText) & ")"
    
        If Not (IsNull(Me.DtpFrom.value)) Then
            StrSQL = StrSQL + " AND dbo.TblBanksDeposite.RecordDate >=" & SQLDate(Me.DtpFrom.value, True)
        End If

        If Not (IsNull(Me.DtpTo.value)) Then
            StrSQL = StrSQL + " AND dbo.TblBanksDeposite.RecordDate<=" & SQLDate(Me.DtpTo.value, True)
        End If
    
        StrSQL = StrSQL & " ORDER BY dbo.TblBanksDeposite.RecordDate, dbo.TblBanksDeposite.NoteID"
            
        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        .AddItem " ›«’Ì·  «·‘Ìﬂ«  «·„Êœ⁄Â ›Ï «·› —…"
        .RowOutlineLevel(.Rows - 1) = 3
        .IsSubtotal(.Rows - 1) = False

        If Not (rs.BOF Or rs.EOF) Then
            StrOneRowData = "   «—ÌŒ «·«Ìœ«⁄" & vbTab & "»‰ﬂ «·‘Ìﬂ" & vbTab & "«·ﬁÌ„…" & vbTab & "—ﬁ„ «·‘Ìﬂ" & vbTab & " «—ÌŒ «·«” Õ›«ﬁ  " & vbTab & "„·«ÕŸ«   " ' & vbTab & "«·‰”»… „‰ ≈Ã„«·Ï «·‘Ìﬂ«  «·„Êœ⁄Â "
            .AddItem StrOneRowData
            .Cell(flexcpBackColor, .Rows - 1, 0, .Rows - 1, 5) = SngHeaderBackColor
            IntStartSelect = 0: IntEndSelect = 0
            IntStartSelect = .Rows
                
            rs.MoveFirst

            For i = 0 To rs.RecordCount - 1
                StrOneRowData = ""
                StrOneRowData = rs("RecordDate").value & vbTab & rs("BankName").value & vbTab & rs("value").value & vbTab & rs("ChequeNo").value & vbTab & rs("DueDate").value & vbTab & rs("Remarks").value
                .AddItem StrOneRowData
                '.Cell(flexcpFloodPercent, .Rows - 1, 5) = 100 * Val(.TextMatrix(.Rows - 1, 2)) / SngTempValue
                rs.MoveNext
            Next i

            IntEndSelect = .Rows - 1
            .Cell(flexcpBackColor, IntStartSelect, 0, .Rows - 1, 5) = SngDataBackColor
        End If
        
        .AddItem "  ‘Ìﬂ«   „  Õ’Ì·Â«  "
        .RowOutlineLevel(.Rows - 1) = 2
        .IsSubtotal(.Rows - 1) = True
        SngTempValue = BankCollectData(val(Me.Dcbank.BoundText), 0, val(Me.dcBranch.BoundText), DtpFrom.value, DtpTo.value)
        .AddItem "≈Ã„«·Ï ‘Ìﬂ«   „  Õ’Ì·Â«  ›Ì «·ﬁ —… : " & vbTab & SngTempValue
             
        .RowOutlineLevel(.Rows - 1) = 3
        .IsSubtotal(.Rows - 1) = False
        .MergeRow(.Rows - 1) = True
            
        StrSQL = "SELECT     dbo.TblBanksCollect.RecordDate, dbo.TblBanksCollectDetails.[value], dbo.TblBanksCollectDetails.Remarks, dbo.TblBanksCollectDetails.ChequeNo, "
        StrSQL = StrSQL & "  dbo.TblBanksCollectDetails.BankName , dbo.TblBanksCollectDetails.DueDate, dbo.TblBanksCollect.BankDate"
        StrSQL = StrSQL & " FROM         dbo.TblBanksCollect INNER JOIN"
        StrSQL = StrSQL & " dbo.TblBanksCollectDetails ON dbo.TblBanksCollect.id = dbo.TblBanksCollectDetails.TblBanksCollectId"

        StrSQL = StrSQL & " Where (dbo.TblBanksCollect.BankID = " & val(Dcbank.BoundText) & ") And (dbo.TblBanksCollect.OperationType = 0)"
        StrSQL = StrSQL & "  AND (dbo.TblBanksCollect.branch_no = " & val(Me.dcBranch.BoundText) & ")"
    
        If Not (IsNull(Me.DtpFrom.value)) Then
            StrSQL = StrSQL + " AND dbo.TblBanksCollect.RecordDate >=" & SQLDate(Me.DtpFrom.value, True)
        End If

        If Not (IsNull(Me.DtpTo.value)) Then
            StrSQL = StrSQL + " AND dbo.TblBanksCollect.RecordDate<=" & SQLDate(Me.DtpTo.value, True)
        End If
    
        StrSQL = StrSQL & " ORDER BY dbo.TblBanksCollect.RecordDate "
            
        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        .AddItem " ›«’Ì·  «·‘Ìﬂ«  «·„Õ’·…  ›Ï «·› —…"
        .RowOutlineLevel(.Rows - 1) = 3
        .IsSubtotal(.Rows - 1) = False

        If Not (rs.BOF Or rs.EOF) Then
            StrOneRowData = "   «—ÌŒ «·«Ìœ«⁄" & vbTab & "»‰ﬂ «·‘Ìﬂ" & vbTab & "«·ﬁÌ„…" & vbTab & "—ﬁ„ «·‘Ìﬂ" & vbTab & " «—ÌŒ «· Õ’Ì·  " & vbTab & "„·«ÕŸ« " ' & vbTab & "«·‰”»…  "
            .AddItem StrOneRowData
            .Cell(flexcpBackColor, .Rows - 1, 0, .Rows - 1, 5) = SngHeaderBackColor
            IntStartSelect = 0: IntEndSelect = 0
            IntStartSelect = .Rows
                
            rs.MoveFirst

            For i = 0 To rs.RecordCount - 1
                StrOneRowData = ""
                StrOneRowData = rs("RecordDate").value & vbTab & rs("BankName").value & vbTab & rs("value").value & vbTab & rs("ChequeNo").value & vbTab & rs("BankDate").value & vbTab & rs("Remarks").value
                .AddItem StrOneRowData
                '.Cell(flexcpFloodPercent, .Rows - 1, 6) = 100 * Val(.TextMatrix(.Rows - 1, 2)) / SngTempValue
                rs.MoveNext
            Next i

            IntEndSelect = .Rows - 1
            .Cell(flexcpBackColor, IntStartSelect, 0, .Rows - 1, 5) = SngDataBackColor
        End If
        
        .AddItem "  ‘Ìﬂ«  „ƒÃ·…  ·„  ”œœ  "
        .RowOutlineLevel(.Rows - 1) = 2
        .IsSubtotal(.Rows - 1) = True
        SngTempValue = BankPendingCheques(val(Me.Dcbank.BoundText), 1, val(Me.dcBranch.BoundText), DtpFrom.value, DtpTo.value)
        .AddItem "≈Ã„«·Ï   ‘Ìﬂ«  „ƒÃ·…  ·„  ”œœ ›Ì  «·ﬁ —… : " & vbTab & SngTempValue
             
        .RowOutlineLevel(.Rows - 1) = 3
        .IsSubtotal(.Rows - 1) = False
        .MergeRow(.Rows - 1) = True
            
        'StrSQL = "SELECT     dbo.TblBanksCollect.RecordDate, dbo.TblBanksCollectDetails.[value], dbo.TblBanksCollectDetails.Remarks, dbo.TblBanksCollectDetails.ChequeNo, "
        'StrSQL = StrSQL & "  dbo.TblBanksCollectDetails.BankName , dbo.TblBanksCollectDetails.DueDate, dbo.TblBanksCollect.BankDate"
        'StrSQL = StrSQL & " FROM         dbo.TblBanksCollect INNER JOIN"
        'StrSQL = StrSQL & " dbo.TblBanksCollectDetails ON dbo.TblBanksCollect.id = dbo.TblBanksCollectDetails.TblBanksCollectId"

        ' StrSQL = StrSQL & " Where (dbo.TblBanksCollect.BankID = " & val(Dcbank.BoundText) & ") And (dbo.TblBanksCollect.OperationType = 1)"
        ' StrSQL = StrSQL & "  AND (dbo.TblBanksCollect.branch_no = " & val(Me.dcBranch.BoundText) & ")"
        '
        StrSQL = "SELECT  DueDate, RecordDate, BankID, BankName, ChequeNo, Remarks, ChequeValue, NoteID, Payed, DepitAccount, notes_all"
        StrSQL = StrSQL & "  From dbo.TblChecqueBoxContent1"
        StrSQL = StrSQL & "  Where (payed = 0 Or payed Is Null)"
 
        StrSQL = StrSQL & " AND   BankID = " & val(Dcbank.BoundText)
  
        If Not (IsNull(Me.DtpFrom.value)) Then
            StrSQL = StrSQL + " AND DueDate >=" & SQLDate(Me.DtpFrom.value, True)
        End If

        If Not (IsNull(Me.DtpTo.value)) Then
            StrSQL = StrSQL + " AND DueDate<=" & SQLDate(Me.DtpTo.value, True)
        End If
    
        StrSQL = StrSQL & " ORDER BY DueDate "
            
        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        .AddItem " ›«’Ì·  «·‘Ìﬂ«  «·„ƒÃ·… «· Ì ·„  ”œœ  ›Ï «·› —…"
        .RowOutlineLevel(.Rows - 1) = 3
        .IsSubtotal(.Rows - 1) = False

        If Not (rs.BOF Or rs.EOF) Then
            StrOneRowData = "   «—ÌŒ «·Õ—ﬂ…" & vbTab & "»‰ﬂ «·‘Ìﬂ" & vbTab & "«·ﬁÌ„…" & vbTab & "—ﬁ„ «·‘Ìﬂ" & vbTab & " «—ÌŒ «·«” Õﬁ«ﬁ  " & vbTab & "„·«ÕŸ« " ' & vbTab & "«·‰”»…  "
            .AddItem StrOneRowData
            .Cell(flexcpBackColor, .Rows - 1, 0, .Rows - 1, 5) = SngHeaderBackColor
            IntStartSelect = 0: IntEndSelect = 0
            IntStartSelect = .Rows
                
            rs.MoveFirst

            For i = 0 To rs.RecordCount - 1
                StrOneRowData = ""
                StrOneRowData = rs("RecordDate").value & vbTab & rs("BankName").value & vbTab & rs("ChequeValue").value & vbTab & rs("ChequeNo").value & vbTab & rs("DUEDate").value & vbTab & rs("Remarks").value
                .AddItem StrOneRowData
                '.Cell(flexcpFloodPercent, .Rows - 1, 6) = 100 * Val(.TextMatrix(.Rows - 1, 2)) / SngTempValue
                rs.MoveNext
            Next i

            IntEndSelect = .Rows - 1
            .Cell(flexcpBackColor, IntStartSelect, 0, .Rows - 1, 5) = SngDataBackColor
        End If
        
        '888888888888888888888888888888888888888888888888
        .AddItem "  ‘Ìﬂ«   „ ”œ«œÂ«  "
        .RowOutlineLevel(.Rows - 1) = 2
        .IsSubtotal(.Rows - 1) = True
        SngTempValue = BankCollectData(val(Me.Dcbank.BoundText), 1, val(Me.dcBranch.BoundText), DtpFrom.value, DtpTo.value)
        .AddItem "≈Ã„«·Ï ‘Ìﬂ«   „ ”œ«œÂ«  ›Ì «·ﬁ —… : " & vbTab & SngTempValue
             
        .RowOutlineLevel(.Rows - 1) = 3
        .IsSubtotal(.Rows - 1) = False
        .MergeRow(.Rows - 1) = True
            
        StrSQL = "SELECT     dbo.TblBanksCollect.RecordDate, dbo.TblBanksCollectDetails.[value], dbo.TblBanksCollectDetails.Remarks, dbo.TblBanksCollectDetails.ChequeNo, "
        StrSQL = StrSQL & "  dbo.TblBanksCollectDetails.BankName , dbo.TblBanksCollectDetails.DueDate, dbo.TblBanksCollect.BankDate"
        StrSQL = StrSQL & " FROM         dbo.TblBanksCollect INNER JOIN"
        StrSQL = StrSQL & " dbo.TblBanksCollectDetails ON dbo.TblBanksCollect.id = dbo.TblBanksCollectDetails.TblBanksCollectId"

        StrSQL = StrSQL & " Where (dbo.TblBanksCollect.BankID = " & val(Dcbank.BoundText) & ") And (dbo.TblBanksCollect.OperationType = 1)"
        StrSQL = StrSQL & "  AND (dbo.TblBanksCollect.branch_no = " & val(Me.dcBranch.BoundText) & ")"
    
        If Not (IsNull(Me.DtpFrom.value)) Then
            StrSQL = StrSQL + " AND dbo.TblBanksCollect.RecordDate >=" & SQLDate(Me.DtpFrom.value, True)
        End If

        If Not (IsNull(Me.DtpTo.value)) Then
            StrSQL = StrSQL + " AND dbo.TblBanksCollect.RecordDate<=" & SQLDate(Me.DtpTo.value, True)
        End If
    
        StrSQL = StrSQL & " ORDER BY dbo.TblBanksCollect.RecordDate "
            
        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        .AddItem " ›«’Ì·  «·‘Ìﬂ«  «·„”œœÂ  ›Ï «·› —…"
        .RowOutlineLevel(.Rows - 1) = 3
        .IsSubtotal(.Rows - 1) = False

        If Not (rs.BOF Or rs.EOF) Then
            StrOneRowData = "   «—ÌŒ «·«Ìœ«⁄" & vbTab & "»‰ﬂ «·‘Ìﬂ" & vbTab & "«·ﬁÌ„…" & vbTab & "—ﬁ„ «·‘Ìﬂ" & vbTab & " «—ÌŒ «·”œ«œ  " & vbTab & "„·«ÕŸ« " ' & vbTab & "«·‰”»…  "
            .AddItem StrOneRowData
            .Cell(flexcpBackColor, .Rows - 1, 0, .Rows - 1, 5) = SngHeaderBackColor
            IntStartSelect = 0: IntEndSelect = 0
            IntStartSelect = .Rows
                
            rs.MoveFirst

            For i = 0 To rs.RecordCount - 1
                StrOneRowData = ""
                StrOneRowData = rs("RecordDate").value & vbTab & rs("BankName").value & vbTab & rs("value").value & vbTab & rs("ChequeNo").value & vbTab & rs("BankDate").value & vbTab & rs("Remarks").value
                .AddItem StrOneRowData
                '.Cell(flexcpFloodPercent, .Rows - 1, 6) = 100 * Val(.TextMatrix(.Rows - 1, 2)) / SngTempValue
                rs.MoveNext
            Next i

            IntEndSelect = .Rows - 1
            .Cell(flexcpBackColor, IntStartSelect, 0, .Rows - 1, 5) = SngDataBackColor
        End If

        '888888888888888888888888888888888888888888888888
                  
        .AddItem "  ‘Ìﬂ«  „— œ… ⁄·Ï «·‘—ﬂ…    "
        .RowOutlineLevel(.Rows - 1) = 2
        .IsSubtotal(.Rows - 1) = True
        SngTempValue = BankCollectData(val(Me.Dcbank.BoundText), 2, val(Me.dcBranch.BoundText), DtpFrom.value, DtpTo.value)
        .AddItem "≈Ã„«·Ï ‘Ìﬂ«   „— œ… ⁄·Ï «·‘—ﬂ…    ›Ì «·ﬁ —… : " & vbTab & SngTempValue
             
        .RowOutlineLevel(.Rows - 1) = 3
        .IsSubtotal(.Rows - 1) = False
        .MergeRow(.Rows - 1) = True
            
        StrSQL = "SELECT     dbo.TblBanksCollect.RecordDate, dbo.TblBanksCollectDetails.[value], dbo.TblBanksCollectDetails.Remarks, dbo.TblBanksCollectDetails.ChequeNo, "
        StrSQL = StrSQL & "  dbo.TblBanksCollectDetails.BankName , dbo.TblBanksCollectDetails.DueDate, dbo.TblBanksCollect.BankDate"
        StrSQL = StrSQL & " FROM         dbo.TblBanksCollect INNER JOIN"
        StrSQL = StrSQL & " dbo.TblBanksCollectDetails ON dbo.TblBanksCollect.id = dbo.TblBanksCollectDetails.TblBanksCollectId"

        StrSQL = StrSQL & " Where (dbo.TblBanksCollect.BankID = " & val(Dcbank.BoundText) & ") And (dbo.TblBanksCollect.OperationType = 2)"
        StrSQL = StrSQL & "  AND (dbo.TblBanksCollect.branch_no = " & val(Me.dcBranch.BoundText) & ")"
    
        If Not (IsNull(Me.DtpFrom.value)) Then
            StrSQL = StrSQL + " AND dbo.TblBanksCollect.RecordDate >=" & SQLDate(Me.DtpFrom.value, True)
        End If

        If Not (IsNull(Me.DtpTo.value)) Then
            StrSQL = StrSQL + " AND dbo.TblBanksCollect.RecordDate<=" & SQLDate(Me.DtpTo.value, True)
        End If
    
        StrSQL = StrSQL & " ORDER BY dbo.TblBanksCollect.RecordDate "
            
        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        .AddItem " ›«’Ì·  «·‘Ìﬂ«    «·„— œ… ⁄·Ï «·‘—ﬂ…  ›Ï «·› —…"
        .RowOutlineLevel(.Rows - 1) = 3
        .IsSubtotal(.Rows - 1) = False

        If Not (rs.BOF Or rs.EOF) Then
            StrOneRowData = "   «—ÌŒ «·«Ìœ«⁄" & vbTab & "»‰ﬂ «·‘Ìﬂ" & vbTab & "«·ﬁÌ„…" & vbTab & "—ﬁ„ «·‘Ìﬂ" & vbTab & " «—ÌŒ «·«— œ«œ  " & vbTab & "„·«ÕŸ« " ' & vbTab & "«·‰”»…  "
            .AddItem StrOneRowData
            .Cell(flexcpBackColor, .Rows - 1, 0, .Rows - 1, 5) = SngHeaderBackColor
            IntStartSelect = 0: IntEndSelect = 0
            IntStartSelect = .Rows
                
            rs.MoveFirst

            For i = 0 To rs.RecordCount - 1
                StrOneRowData = ""
                StrOneRowData = rs("RecordDate").value & vbTab & rs("BankName").value & vbTab & rs("value").value & vbTab & rs("ChequeNo").value & vbTab & rs("BankDate").value & vbTab & rs("Remarks").value
                .AddItem StrOneRowData
                '.Cell(flexcpFloodPercent, .Rows - 1, 6) = 100 * Val(.TextMatrix(.Rows - 1, 2)) / SngTempValue
                rs.MoveNext
            Next i

            IntEndSelect = .Rows - 1
            .Cell(flexcpBackColor, IntStartSelect, 0, .Rows - 1, 5) = SngDataBackColor
        End If
          
        .AddItem "  ‘Ìﬂ«  „— œ…   ··‘—ﬂÂ    "
        .RowOutlineLevel(.Rows - 1) = 2
        .IsSubtotal(.Rows - 1) = True
        SngTempValue = BankCollectData(val(Me.Dcbank.BoundText), 3, val(Me.dcBranch.BoundText), DtpFrom.value, DtpTo.value)
        .AddItem "≈Ã„«·Ï ‘Ìﬂ«   „— œ…   ··‘—ﬂÂ    ›Ì «·ﬁ —… : " & vbTab & SngTempValue
             
        .RowOutlineLevel(.Rows - 1) = 3
        .IsSubtotal(.Rows - 1) = False
        .MergeRow(.Rows - 1) = True
            
        StrSQL = "SELECT     dbo.TblBanksCollect.RecordDate, dbo.TblBanksCollectDetails.[value], dbo.TblBanksCollectDetails.Remarks, dbo.TblBanksCollectDetails.ChequeNo, "
        StrSQL = StrSQL & "  dbo.TblBanksCollectDetails.BankName , dbo.TblBanksCollectDetails.DueDate, dbo.TblBanksCollect.BankDate"
        StrSQL = StrSQL & " FROM         dbo.TblBanksCollect INNER JOIN"
        StrSQL = StrSQL & " dbo.TblBanksCollectDetails ON dbo.TblBanksCollect.id = dbo.TblBanksCollectDetails.TblBanksCollectId"

        StrSQL = StrSQL & " Where (dbo.TblBanksCollect.BankID = " & val(Dcbank.BoundText) & ") And (dbo.TblBanksCollect.OperationType = 3)"
        StrSQL = StrSQL & "  AND (dbo.TblBanksCollect.branch_no = " & val(Me.dcBranch.BoundText) & ")"
    
        If Not (IsNull(Me.DtpFrom.value)) Then
            StrSQL = StrSQL + " AND dbo.TblBanksCollect.RecordDate >=" & SQLDate(Me.DtpFrom.value, True)
        End If

        If Not (IsNull(Me.DtpTo.value)) Then
            StrSQL = StrSQL + " AND dbo.TblBanksCollect.RecordDate<=" & SQLDate(Me.DtpTo.value, True)
        End If
    
        StrSQL = StrSQL & " ORDER BY dbo.TblBanksCollect.RecordDate "
            
        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        .AddItem " ›«’Ì·  «·‘Ìﬂ«    «·„— œ…   ··‘—ﬂÂ  ›Ï «·› —…"
        .RowOutlineLevel(.Rows - 1) = 3
        .IsSubtotal(.Rows - 1) = False

        If Not (rs.BOF Or rs.EOF) Then
            StrOneRowData = "   «—ÌŒ «·«Ìœ«⁄" & vbTab & "»‰ﬂ «·‘Ìﬂ" & vbTab & "«·ﬁÌ„…" & vbTab & "—ﬁ„ «·‘Ìﬂ" & vbTab & " «—ÌŒ «·«— œ«œ  " & vbTab & "„·«ÕŸ« " ' & vbTab & "«·‰”»…  "
            .AddItem StrOneRowData
            .Cell(flexcpBackColor, .Rows - 1, 0, .Rows - 1, 5) = SngHeaderBackColor
            IntStartSelect = 0: IntEndSelect = 0
            IntStartSelect = .Rows
                
            rs.MoveFirst

            For i = 0 To rs.RecordCount - 1
                StrOneRowData = ""
                StrOneRowData = rs("RecordDate").value & vbTab & rs("BankName").value & vbTab & rs("value").value & vbTab & rs("ChequeNo").value & vbTab & rs("BankDate").value & vbTab & rs("Remarks").value
                .AddItem StrOneRowData
                '.Cell(flexcpFloodPercent, .Rows - 1, 6) = 100 * Val(.TextMatrix(.Rows - 1, 2)) / SngTempValue
                rs.MoveNext
            Next i

            IntEndSelect = .Rows - 1
            .Cell(flexcpBackColor, IntStartSelect, 0, .Rows - 1, 5) = SngDataBackColor
        End If
                  
        '-------------------------------------------------------------------------------
        'LoadCustomersAccounts
        '     CustomersBalances
        '     SupplierBalances
        '--------------------------------------------------------------
        If SystemOptions.UserInterface = ArabicInterface Then

            For i = 0 To .Cols - 1
                .ColAlignment(i) = flexAlignRightCenter
                .FixedAlignment(i) = flexAlignRightCenter
            Next i

        ElseIf SystemOptions.UserInterface = EnglishInterface Then

            For i = 0 To .Cols - 1
                .ColAlignment(i) = flexAlignLeftCenter
                .FixedAlignment(i) = flexAlignLeftCenter
            Next i

        End If

        FormatGrid
        '--------------------------------------------------------------
        .AutoSize 0, .Cols - 1, False
        .Outline 2
    End With

End Sub

Private Function GetTransactionsValues(IntTransType As String) As TransactionsValues
    Dim rs As ADODB.Recordset
    Dim StrSQL As String

    If SystemOptions.SysDataBaseType = SQLServerDataBase Then
        StrSQL = "Select Isnull(TotalCash,0)as TotalCash, Isnull(TotalDue,0)as TotalDue,"
        StrSQL = StrSQL + " Isnull(TotalCash,0)+Isnull(TotalDue,0)as  NET"
        StrSQL = StrSQL + " From"
        StrSQL = StrSQL + "("
        StrSQL = StrSQL + " SELECT   Convert(Decimal(38,2), SUM(CASE WHEN PaymentType = 0 THEN TotalAfterTax ELSE 0 END)) AS TotalCash,"
        StrSQL = StrSQL + " Convert(Decimal(38,2), SUM(CASE WHEN PaymentType =1 THEN TotalAfterTax ELSE 0 END)) AS TotalDue"
        StrSQL = StrSQL + " FROM         dbo.QryTransactionsTotal() QryTransactionsTotal "
        StrSQL = StrSQL + " Where (Transaction_Type = " & IntTransType & ")"
    
        If Not (IsNull(Me.DtpFrom.value)) Then
            StrSQL = StrSQL + " AND Transaction_Date >=" & SQLDate(Me.DtpFrom.value, True)
        End If

        If Not (IsNull(Me.DtpTo.value)) Then
            StrSQL = StrSQL + " AND Transaction_Date <=" & SQLDate(Me.DtpTo.value, True)
        End If

        '       If Val(dcBranch.BoundText) <> 0 Then
        '     StrSQL = StrSQL + "  AND  (BranchId = " & Val(dcBranch.text) & ")"
        '    End If
     
        StrSQL = StrSQL + ") as xTable"
    End If

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.BOF Or rs.EOF) Then
        GetTransactionsValues.TotalCash = IIf(IsNull(rs("TotalCash").value), 0, rs("TotalCash").value)
        GetTransactionsValues.TotalDue = IIf(IsNull(rs("TotalDue").value), 0, rs("TotalDue").value)
        GetTransactionsValues.TotalNet = IIf(IsNull(rs("NET").value), 0, rs("NET").value)
    End If

    rs.Close
    Set rs = Nothing
End Function

Private Function LoadTransCustomers(IntTransType As Integer, _
                                    SngTransTotals As Double) As Long
    Exit Function
    Dim StrSQL As String
    Dim rs As ADODB.Recordset
    Dim i As Integer
    Dim StrOneRowData As String
    Dim IntStartSelect As Integer, IntEndSelect As Integer
    Dim SngTempValue As Single
    Dim SngHeaderBackColor As Single
    Dim SngDataBackColor As Single
    Dim StrStartDate As String
    Dim SngTemp1 As Double, SngTemp2 As Double, SngTemp3 As Double, SngTemp4 As Double

    Dim SngCashCount As Double
    Dim SngDueCount As Double
    Dim SngCashTotal As Double
    Dim SngDueTotal As Double
    Dim LngCustomersCount As Long

    SngHeaderBackColor = &HC0C0C0
    SngDataBackColor = &HE2E9E9

    If SystemOptions.SysDataBaseType = SQLServerDataBase Then
        StrSQL = "SELECT Convert(Decimal(38,2),SUM(Case WHEN PaymentType=0 THEN QryTransactionsTotal.TotalAfterTax ELSE 0 END)) AS TotalCash,"
        StrSQL = StrSQL + " Count(Case WHEN PaymentType=0 THEN QryTransactionsTotal.TotalAfterTax END) AS CountCash,"
        StrSQL = StrSQL + " Convert(Decimal(38,2),SUM(Case WHEN PaymentType=1 THEN QryTransactionsTotal.TotalAfterTax ELSE 0 END)) AS TotalDue,"
        StrSQL = StrSQL + " Count(Case WHEN PaymentType=1 THEN QryTransactionsTotal.TotalAfterTax  END) AS CountDue,"
        StrSQL = StrSQL + " Convert(Decimal(38,2),SUM( QryTransactionsTotal.TotalAfterTax))as Total,"
        StrSQL = StrSQL + " dbo.TblCustemers.CusName"
        StrSQL = StrSQL + " FROM dbo.QryTransactionsTotal() QryTransactionsTotal INNER JOIN "
        StrSQL = StrSQL + " dbo.TblCustemers ON QryTransactionsTotal.CusID = dbo.TblCustemers.CusID"
    
        If IntTransType = 2 Then
            StrSQL = StrSQL + " Where (QryTransactionsTotal.Transaction_Type = " & " 2 OR QryTransactionsTotal.Transaction_Type = 21" & ")"
        ElseIf IntTransType = 1 Then
            StrSQL = StrSQL + " Where (QryTransactionsTotal.Transaction_Type = " & " 1 OR  QryTransactionsTotal.Transaction_Type =22" & ")"
        Else
            StrSQL = StrSQL + " Where (QryTransactionsTotal.Transaction_Type = " & IntTransType & ")"
        End If
    
        If Not (IsNull(Me.DtpFrom.value)) Then
            StrSQL = StrSQL + " AND Transaction_Date >=" & SQLDate(Me.DtpFrom.value, True)
        End If

        If Not (IsNull(Me.DtpTo.value)) Then
            StrSQL = StrSQL + " AND Transaction_Date <=" & SQLDate(Me.DtpTo.value, True)
        End If

        StrSQL = StrSQL + " GROUP BY  dbo.TblCustemers.CusName"
        StrSQL = StrSQL + " Order By SUM( QryTransactionsTotal.TotalAfterTax) DESC"
    End If

    With Me.Fg
        StrOneRowData = "«”„ «·⁄„Ì·" & vbTab & "‰ﬁœÏ(⁄œœ ---≈Ã„«·Ï)" & vbTab & "√Ã·(⁄œœ ---≈Ã„«·Ï)" & vbTab & "≈Ã„«·Ï " & vbTab & "«·‰”»…"
        .AddItem StrOneRowData
        .Cell(flexcpBackColor, .Rows - 1, 0, .Rows - 1, 4) = SngHeaderBackColor
    End With

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
            
    If Not (rs.BOF Or rs.EOF) Then

        With Me.Fg
            IntStartSelect = 0: IntEndSelect = 0
            IntStartSelect = .Rows
            rs.MoveFirst

            For i = 0 To rs.RecordCount - 1
                LngCustomersCount = LngCustomersCount + 1
                SngTemp1 = IIf(IsNull(rs("TotalCash").value), 0, rs("TotalCash").value)
                SngTemp2 = IIf(IsNull(rs("CountCash").value), 0, rs("CountCash").value)
                SngTemp3 = IIf(IsNull(rs("TotalDue").value), 0, rs("TotalDue").value)
                SngTemp4 = IIf(IsNull(rs("CountDue").value), 0, rs("CountDue").value)
            
                SngCashCount = SngCashCount + SngTemp2
                SngDueCount = SngDueCount + SngTemp4
                SngCashTotal = SngCashTotal + SngTemp1
                SngDueTotal = SngDueTotal + SngTemp3
            
                StrOneRowData = rs("CusName").value & vbTab
                StrOneRowData = StrOneRowData & "" & SngTemp2 & " --- " & SngTemp1 & vbTab
            
                StrOneRowData = StrOneRowData & " " & SngTemp4 & " --- " & SngTemp3 & vbTab
                StrOneRowData = StrOneRowData & " " & (SngTemp2 + SngTemp4) & " --- " & (SngTemp1 + SngTemp3)
                .AddItem StrOneRowData
                .RowOutlineLevel(.Rows - 1) = 4
                .IsSubtotal(.Rows - 1) = False

                If SngTransTotals <> 0 Then
                    .TextMatrix(.Rows - 1, 4) = Format((100 * (SngTemp1 + SngTemp3)) / SngTransTotals, SystemOptions.SysDefCurrencyForamt)
                    .Cell(flexcpFloodPercent, .Rows - 1, 4) = (100 * (SngTemp1 + SngTemp3)) / SngTransTotals
                End If

                rs.MoveNext
            Next i

            IntEndSelect = .Rows - 1
            .Cell(flexcpBackColor, IntStartSelect, 0, .Rows - 1, 4) = SngDataBackColor
            StrOneRowData = ""
            StrOneRowData = "⁄œœ «·⁄„·«¡ «Ê «·„Ê—œÌ‰ : " & LngCustomersCount
            StrOneRowData = StrOneRowData & vbTab & SngCashCount & "---" & SngCashTotal
            StrOneRowData = StrOneRowData & vbTab & SngDueCount & "---" & SngDueTotal
            StrOneRowData = StrOneRowData & vbTab & (SngCashCount + SngDueCount) & "---" & (SngCashTotal + SngDueTotal)
            .AddItem StrOneRowData
            .RowOutlineLevel(.Rows - 1) = 4
            .IsSubtotal(.Rows - 1) = False
            .Cell(flexcpForeColor, .Rows - 1, 1, .Rows - 1, 4) = vbRed
            '.Cell(flexcpBackColor, .Rows - 1, 1, .Rows - 1, 5) = SngHeaderBackColor
            LoadTransCustomers = LngCustomersCount
        End With

    Else
        Exit Function
    End If

End Function

Private Sub FormatGrid()
    Dim XFont As IFontDisp
    Dim i As Long

    With Me.Fg
        Set XFont = Me.Font
        XFont.name = "Tahoma"
        XFont.Charset = 178
        XFont.Bold = True
        XFont.Underline = True
        XFont.Size = 10

        For i = 1 To .Rows - 1

            If .IsSubtotal(i) = True Then
                .RowHeight(i) = 450
                XFont.Size = (14 - (.RowOutlineLevel(i) + 1))
                Set .Cell(flexcpFont, i, 0, i, 0) = XFont

                If .RowOutlineLevel(i) = 2 Then
                    .Cell(flexcpForeColor, i, 0, i, 0) = vbBlue
                ElseIf .RowOutlineLevel(i) = 3 Then
                    .Cell(flexcpForeColor, i, 0, i, 0) = vbRed
                ElseIf .RowOutlineLevel(i) = 4 Then
                    .Cell(flexcpForeColor, i, 0, i, 0) = vbGreen
                Else
                    .Cell(flexcpForeColor, i, 0, i, 0) = vbBlack
                End If
            End If

        Next i

    End With

End Sub

Private Sub ISButton2_Click()

    If DoPremis(Do_Print, Me.name, True) = False Then
        Exit Sub
    End If
        
    PrintData
End Sub

Private Sub PrintData()
    On Error Resume Next
    Dim Frm As FrmViewListPrint

    Set Frm = New FrmViewListPrint
    Frm.VSPrinter1.Zoom = 100
    Frm.VSPrinter1.Orientation = orLandscape
    Frm.VSPrinter1.StartDoc
    Frm.VSPrinter1.MarginLeft = 100
    Frm.VSPrinter1.MarginRight = 100
    Frm.VSPrinter1.CurrentX = 100
    Frm.VSPrinter1.CurrentY = 100
    Frm.VSPrinter1.text = "‰Ÿ«„ œÌ‰«„Ìﬂ »«Ì  «·„ ﬂ«„·  "
    'Frm.VSPrinter1.CurrentX = 100
    'Frm.VSPrinter1.CurrentY = 500
    Frm.Caption = " ﬁ—Ì— „Ã„⁄"
    Frm.VSPrinter1.RenderControl = Fg.hwnd
    Frm.VSPrinter1.EndDoc
    Set Frm.Fg = Me.Fg
    Frm.show
End Sub

Private Function LoadTransItems(IntTransType As Integer) As Long
    Dim StrSQL As String
    Dim rs As ADODB.Recordset
    Dim i As Integer
    Dim StrOneRowData As String
    Dim IntStartSelect As Integer, IntEndSelect As Integer
    Dim SngTempValue As Single
    Dim SngHeaderBackColor As Single
    Dim SngDataBackColor As Single
    Dim StrStartDate As String
    Dim SngTemp1 As Single, SngTemp2 As Single, SngTemp3 As Single, SngTemp4 As Single
    Dim SngItemsTotal As Single
    Dim LngItemsCount As Long
    Dim Remarks As String
    SngHeaderBackColor = &HC0C0C0
    SngDataBackColor = &HE2E9E9

    ' «Ìœ«⁄«  «·‰›œÌ…
   
    With Me.Fg
        .AddItem StrOneRowData
        .RowOutlineLevel(.Rows - 1) = 3
        .IsSubtotal(.Rows - 1) = True
        StrOneRowData = "«· «—ÌŒ  " & vbTab & "«·ﬁ”„…" & vbTab & "«·Œ“Ì‰…  " & vbTab & "„·«ÕŸ«  " & vbTab & "«·‰”»…"
        .AddItem StrOneRowData
        .Cell(flexcpBackColor, .Rows - 1, 0, .Rows - 1, 4) = SngHeaderBackColor
    
        StrSQL = "SELECT     TOP 100 PERCENT dbo.TblBanksDeposite.RecordDate, dbo.TblBanksDepositeDetails.[value], dbo.TblBoxesData.BoxName, dbo.TblBanksDeposite.Remarks"
        StrSQL = StrSQL & " FROM         dbo.TblBanksDeposite INNER JOIN"
        StrSQL = StrSQL & " dbo.TblBanksDepositeDetails ON dbo.TblBanksDeposite.id = dbo.TblBanksDepositeDetails.TblBanksDepositeId LEFT OUTER JOIN"
        StrSQL = StrSQL & " dbo.TblBoxesData ON dbo.TblBanksDepositeDetails.BoxID = dbo.TblBoxesData.BoxID"
        StrSQL = StrSQL & " Where (dbo.TblBanksDeposite.BankID = " & val(Dcbank.BoundText) & ") And (dbo.TblBanksDepositeDetails.box_or_bank = 0)"
        StrSQL = StrSQL & "  AND (dbo.TblBanksDeposite.branch_no = 1)"
    
        If Not (IsNull(Me.DtpFrom.value)) Then
            StrSQL = StrSQL + " AND dbo.TblBanksDeposite.RecordDate >=" & SQLDate(Me.DtpFrom.value, True)
        End If

        If Not (IsNull(Me.DtpTo.value)) Then
            StrSQL = StrSQL + " AND dbo.TblBanksDeposite.RecordDate<=" & SQLDate(Me.DtpTo.value, True)
        End If
    
        StrSQL = StrSQL & " ORDER BY dbo.TblBanksDeposite.RecordDate, dbo.TblBanksDeposite.NoteID"

        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
            
        If Not (rs.BOF Or rs.EOF) Then

            IntStartSelect = 0: IntEndSelect = 0
            IntStartSelect = .Rows
            rs.MoveFirst

            For i = 0 To rs.RecordCount - 1
                LngItemsCount = LngItemsCount + 1
            
                StrOneRowData = rs("RecordDate").value & vbTab
                StrOneRowData = StrOneRowData & rs("value").value & vbTab
            
                StrOneRowData = StrOneRowData & IIf(IsNull(rs("BoxName").value), "", rs("BoxName").value) & vbTab
                Remarks = IIf(IsNull(rs("Remarks").value), "", rs("Remarks").value)
                StrOneRowData = StrOneRowData & Remarks
                SngTemp1 = rs("value").value
                SngItemsTotal = SngItemsTotal + SngTemp1
            
                .AddItem StrOneRowData
                .RowOutlineLevel(.Rows - 1) = 4
                .IsSubtotal(.Rows - 1) = False
                '  If SngItemsTotal <> 0 Then
                '      .TextMatrix(.Rows - 1, 4) = Format((100 * (SngTemp1)) / SngItemsTotal, SystemOptions.SysDefCurrencyForamt)
                '      .Cell(flexcpFloodPercent, .Rows - 1, 4) = (100 * (SngTemp1)) / SngItemsTotal
                '  End If
                rs.MoveNext
            Next i

            IntEndSelect = .Rows - 1
            .Cell(flexcpBackColor, IntStartSelect, 0, .Rows - 1, 4) = SngDataBackColor

            For i = IntStartSelect To IntEndSelect
                .Cell(flexcpFloodPercent, i, 4, i, 4) = 100 * val(.TextMatrix(i, 3)) / SngItemsTotal
                .TextMatrix(i, 4) = 100 * val(.TextMatrix(i, 1)) / SngItemsTotal
                .Cell(flexcpFontBold, i, 4, i, 4) = True
            Next i

            StrOneRowData = ""
            StrOneRowData = "⁄œœ «·Õ—ﬂ« : " & LngItemsCount
            .AddItem StrOneRowData
            .RowOutlineLevel(.Rows - 1) = 4
            .IsSubtotal(.Rows - 1) = False
            .Cell(flexcpForeColor, .Rows - 1, 1, .Rows - 1, 4) = vbRed
            '.Cell(flexcpBackColor, .Rows - 1, 1, .Rows - 1, 5) = SngHeaderBackColor
            LoadTransItems = LngItemsCount
    
        Else
            Exit Function
        End If

    End With

End Function

Private Sub WriteToolTip(LngMouseRow As Long, _
                         LngMouseCol As Long)
    Dim StrTemp As String
    Dim StrToolTip As String
    Dim VarTemp  As Variant
    On Error GoTo hErr

    With Fg
        StrTemp = Trim$(.TextMatrix(LngMouseRow, LngMouseCol))

        If StrTemp = "" Then
            Fg.ToolTipText = ""
            Exit Sub
        ElseIf InStr(1, StrTemp, "---", vbTextCompare) <> 0 Then
            VarTemp = Split(StrTemp, "---", , vbTextCompare)

            If val(VarTemp(0)) <> 0 Then
                StrToolTip = WriteNo(CStr(VarTemp(0)), 0, False)
            End If

            If val(VarTemp(1)) <> 0 Then
                StrToolTip = StrToolTip & "     " & WriteNo(CStr(VarTemp(1)), 0, False)
            End If

        ElseIf val(StrTemp) <> 0 Then 'Â–Â «·ﬁÌ„…  Õ ÊÏ ⁄·Ï «—ﬁ«„

            If IsDblValue(StrTemp) Then
                StrToolTip = WriteNo(IsDblValue(StrTemp), 0, False)
            Else
                StrToolTip = WriteNo(val(StrTemp), 0, False)
            End If
        End If

        Fg.ToolTipText = StrToolTip
    End With

    Exit Sub
hErr:
    Fg.ToolTipText = ""
End Sub

Private Function IsDblValue(strValue As String) As Double
    On Error GoTo hErr
    'IsDblValue = CDbl(strValue)
    Exit Function
hErr:
    IsDblValue = 0
End Function

Private Sub CustomersBalances()
    Dim StrSQL As String
    Dim rs As ADODB.Recordset
    Dim StrOneRowData As String
    Dim SngHeaderBackColor As Single
    Dim IntStartSelect As Integer
    Dim IntEndSelect As Integer
    Dim Boxname As String, BoxBalance As Double, balancetype As String
    Dim SngDataBackColor As Single
    SngHeaderBackColor = &HC0C0C0
    SngDataBackColor = &HE2E9E9

    With Me.Fg

        .AddItem "«—’œ…  «·⁄„·«¡ Œ·«· «·› —…"
        .RowOutlineLevel(.Rows - 1) = 2
        .IsSubtotal(.Rows - 1) = True
        Set rs = New ADODB.Recordset
         
        If SystemOptions.UserInterface = ArabicInterface Then
            StrSQL = "SELECT * from TblCustemers where type=1 and CusID>2 order by CusName"
        Else
            StrSQL = "SELECT * from TblCustemers where type=1 and CusID>2 order by CusNamee"
        End If
            
        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        .AddItem " ›«’Ì· √—’œ… «·⁄„·«¡"
        .RowOutlineLevel(.Rows - 1) = 3
        .IsSubtotal(.Rows - 1) = False

        If Not (rs.BOF Or rs.EOF) Then
            StrOneRowData = "«”„ «·⁄„Ì·" & vbTab & "«·—’Ìœ " & vbTab & "  ÿ»Ì⁄Â «·—’Ìœ" & vbTab & "   "
            StrOneRowData = StrOneRowData & vbTab & " „·«ÕŸ« "
            .AddItem StrOneRowData
            .Cell(flexcpBackColor, .Rows - 1, 0, .Rows - 1, 4) = SngHeaderBackColor
            IntStartSelect = 0: IntEndSelect = 0
            IntStartSelect = .Rows
                
            rs.MoveFirst
            Dim FirstPeriod As Date
            Dim i As Integer
            Dim Balance As Double

            For i = 0 To rs.RecordCount - 1
                StrOneRowData = ""
                    
                getFirstPeriodDateInthisYear FirstPeriod
 
                Balance = GetActualAccountBalance(rs("Account_Code").value, , DtpFrom.value, DtpTo.value)
                BoxBalance = Abs(Balance)

                If SystemOptions.UserInterface = ArabicInterface Then
                    If Balance > 0 Then
                        balancetype = "„œÌ‰"
                    ElseIf Balance < 0 Then
                        balancetype = "œ«∆‰"
                    Else
            
                        balancetype = ""
                    End If

                Else

                    If Balance > 0 Then
                        balancetype = "Debit"
                    ElseIf Balance < 0 Then
                        .balancetype = "Credit"
                    Else
            
                        balancetype = " "
                    End If

                End If
            
                '
                '             SngTemp3 = ""
                    
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrOneRowData = rs("CusName").value & vbTab & BoxBalance & vbTab & balancetype & vbTab  '& SngTemp3 & vbTab & ((SngTemp1 + SngTemp2) - SngTemp3)
                Else
                    StrOneRowData = rs("CusNamee").value & vbTab & BoxBalance & vbTab & balancetype & vbTab  '& SngTemp3 & vbTab & ((SngTemp1 + SngTemp2) - SngTemp3)
                End If

                .AddItem StrOneRowData
                rs.MoveNext
            Next i

            IntEndSelect = .Rows - 1
            .Cell(flexcpBackColor, IntStartSelect, 0, .Rows - 1, 4) = SngDataBackColor
        End If

    End With

End Sub

Private Sub SupplierBalances()
    Dim StrSQL As String
    Dim rs As ADODB.Recordset
    Dim StrOneRowData As String
    Dim SngHeaderBackColor As Single
    Dim IntStartSelect As Integer
    Dim IntEndSelect As Integer
    Dim Boxname As String, BoxBalance As Double, balancetype As String
    Dim SngDataBackColor As Single
    SngHeaderBackColor = &HC0C0C0
    SngDataBackColor = &HE2E9E9

    With Me.Fg

        .AddItem "«—’œ…  «·„Ê—œÌ‰ Œ·«· «·› —…"
        .RowOutlineLevel(.Rows - 1) = 2
        .IsSubtotal(.Rows - 1) = True
        Set rs = New ADODB.Recordset
         
        If SystemOptions.UserInterface = ArabicInterface Then
            StrSQL = "SELECT * from TblCustemers where type=2 and CusID>2 order by CusName"
        Else
            StrSQL = "SELECT * from TblCustemers where type=2 and CusID>2 order by CusNamee"
        End If
            
        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        .AddItem " ›«’Ì· √—’œ… «·„Ê—œÌ‰"
        .RowOutlineLevel(.Rows - 1) = 3
        .IsSubtotal(.Rows - 1) = False

        If Not (rs.BOF Or rs.EOF) Then
            StrOneRowData = "«”„ «·„Ê—œ" & vbTab & "«·—’Ìœ " & vbTab & "  ÿ»Ì⁄Â «·—’Ìœ" & vbTab & "   "
            StrOneRowData = StrOneRowData & vbTab & " „·«ÕŸ« "
            .AddItem StrOneRowData
            .Cell(flexcpBackColor, .Rows - 1, 0, .Rows - 1, 4) = SngHeaderBackColor
            IntStartSelect = 0: IntEndSelect = 0
            IntStartSelect = .Rows
                
            rs.MoveFirst
            Dim FirstPeriod As Date
            Dim i As Integer
            Dim Balance As Double

            For i = 0 To rs.RecordCount - 1
                StrOneRowData = ""
                    
                getFirstPeriodDateInthisYear FirstPeriod
 
                Balance = GetActualAccountBalance(rs("Account_Code").value, , DtpFrom.value, DtpTo.value)
                BoxBalance = Abs(Balance)

                If SystemOptions.UserInterface = ArabicInterface Then
                    If Balance > 0 Then
                        balancetype = "„œÌ‰"
                    ElseIf Balance < 0 Then
                        balancetype = "œ«∆‰"
                    Else
            
                        balancetype = ""
                    End If

                Else

                    If Balance > 0 Then
                        balancetype = "Debit"
                    ElseIf Balance < 0 Then
                        .balancetype = "Credit"
                    Else
            
                        balancetype = " "
                    End If

                End If
            
                '
                '             SngTemp3 = ""
                    
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrOneRowData = rs("CusName").value & vbTab & BoxBalance & vbTab & balancetype & vbTab  '& SngTemp3 & vbTab & ((SngTemp1 + SngTemp2) - SngTemp3)
                Else
                    StrOneRowData = rs("CusNamee").value & vbTab & BoxBalance & vbTab & balancetype & vbTab  '& SngTemp3 & vbTab & ((SngTemp1 + SngTemp2) - SngTemp3)
                End If

                .AddItem StrOneRowData
                rs.MoveNext
            Next i

            IntEndSelect = .Rows - 1
            .Cell(flexcpBackColor, IntStartSelect, 0, .Rows - 1, 4) = SngDataBackColor
        End If

    End With

End Sub

Private Sub LoadCustomersAccounts()
    Dim StrSQL As String
    Dim rs As ADODB.Recordset
    Dim StrOneRowData As String

    With Me.Fg
        .AddItem "«·œ«∆‰Ì‰"
        .RowOutlineLevel(.Rows - 1) = 2
        .IsSubtotal(.Rows - 1) = True

        If IsNull(Me.DtpFrom.value) Then
            StrOneRowData = "—’Ìœ «·œ«∆‰Ì‰ «Ê· «·› —…:" & vbTab & "0"
        Else
            StrSQL = CustomersAccountsSQL(Me.DtpFrom.value, 1)
            Set rs = New ADODB.Recordset
            rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

            If Not (rs.BOF Or rs.EOF) Then
                StrOneRowData = "—’Ìœ «·œ«∆‰Ì‰ «Ê· «·› —…:" & vbTab & rs("SumX").value
            Else
                StrOneRowData = "—’Ìœ «·œ«∆‰Ì‰ «Ê· «·› —…:" & vbTab & "0"
            End If
        End If

        .AddItem StrOneRowData
        .RowOutlineLevel(.Rows - 1) = 3
        .IsSubtotal(.Rows - 1) = False
     
        StrSQL = CustomersAccountsSQL(Me.DtpTo.value, 1)
        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (rs.BOF Or rs.EOF) Then
            StrOneRowData = "—’Ìœ «·œ«∆‰Ì‰ ‰Â«Ì… «·› —…:" & vbTab & rs("SumX").value
        Else
            StrOneRowData = "—’Ìœ «·œ«∆‰Ì‰ ‰Â«Ì… «·› —…:" & vbTab & "0"
        End If

        .AddItem StrOneRowData
        .RowOutlineLevel(.Rows - 1) = 3
        .IsSubtotal(.Rows - 1) = False
        '-------------------------------------------------------------------------------
        .AddItem "«·„œ‰Ì‰"
        .RowOutlineLevel(.Rows - 1) = 2
        .IsSubtotal(.Rows - 1) = True

        If IsNull(Me.DtpFrom.value) Then
            StrOneRowData = "—’Ìœ «·„œ‰Ì‰ «Ê· «·› —…:" & vbTab & "0"
        Else
            StrSQL = CustomersAccountsSQL(Me.DtpFrom.value, 0)
            Set rs = New ADODB.Recordset
            rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

            If Not (rs.BOF Or rs.EOF) Then
                StrOneRowData = "—’Ìœ «·„œ‰Ì‰ «Ê· «·› —…:" & vbTab & Abs(rs("SumX").value)
            Else
                StrOneRowData = "—’Ìœ «·„œ‰Ì‰ «Ê· «·› —…:" & vbTab & "0"
            End If
        End If
        
        .AddItem StrOneRowData
        .RowOutlineLevel(.Rows - 1) = 3
        .IsSubtotal(.Rows - 1) = False
        
        StrSQL = CustomersAccountsSQL(Me.DtpTo.value, 0)
        Set rs = New ADODB.Recordset
        rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (rs.BOF Or rs.EOF) Then
            StrOneRowData = "—’Ìœ «·„œ‰Ì‰ ‰Â«Ì… «·› —…:" & vbTab & Abs(rs("SumX").value)
        Else
            StrOneRowData = "—’Ìœ «·„œ‰Ì‰ ‰Â«Ì… «·› —…:" & vbTab & "0"
        End If
        
        .AddItem StrOneRowData
        .RowOutlineLevel(.Rows - 1) = 3
        .IsSubtotal(.Rows - 1) = False

    End With

End Sub

Private Function CustomersAccountsSQL(ToDate As Variant, _
                                      IntAccountType As Integer) As String

    Dim StrSQL As String
    StrSQL = "Select Cast(Sum(CurrentAccount) as decimal(38,2) )as SumX "
    StrSQL = StrSQL + " From"
    StrSQL = StrSQL + "( "
    StrSQL = StrSQL + "Select "
    StrSQL = StrSQL + "IsNull"
    StrSQL = StrSQL + "("
    StrSQL = StrSQL + "Case OpenBalanceType WHEN 0 THEN (CustomerAccount )+ (-1*OpenBalance)"
    StrSQL = StrSQL + " WHEN 1 THEN (CustomerAccount)+ (OpenBalance)"
    StrSQL = StrSQL + " Else"
    StrSQL = StrSQL + "    CustomerAccount"
    StrSQL = StrSQL + " End"
    StrSQL = StrSQL + ",0)as CurrentAccount"
    StrSQL = StrSQL + " From"
    StrSQL = StrSQL + "("
    StrSQL = StrSQL + "SELECT TOP 100 PERCENT dbo.TblCustemers.CusID, dbo.TblCustemers.CusName," & "dbo.TblCustemers.Type,"
    StrSQL = StrSQL + "dbo.TblCustemers.OpenBalance, dbo.TblCustemers.OpenBalanceType," & "dbo.TblCustemers.OpenBalanceDate,"
    StrSQL = StrSQL + "SUM(dbo.QryCustomerBalance.Note_Value * dbo.QryCustomerBalance.CreditOrDebit)" & "As CustomerAccount"
    StrSQL = StrSQL + " FROM dbo.QryCustomerBalance(0)  RIGHT JOIN dbo.TblCustemers ON " & "dbo.QryCustomerBalance.CusID = dbo.TblCustemers.CusID"

    If Not (IsNull(ToDate)) Then
        StrSQL = StrSQL + " Where dbo.QryCustomerBalance.NoteDate < " & SQLDate(CDate(ToDate), True) & ""
    End If

    StrSQL = StrSQL + " GROUP BY dbo.TblCustemers.CusID, dbo.TblCustemers.CusName,dbo.TblCustemers.OpenBalance ,"
    StrSQL = StrSQL + " dbo.TblCustemers.OpenBalanceType , dbo.TblCustemers.OpenBalanceDate, dbo.TblCustemers.Type"
    StrSQL = StrSQL + " ORDER BY dbo.TblCustemers.CusID"
    StrSQL = StrSQL + ")"
    StrSQL = StrSQL + "XTable"
    StrSQL = StrSQL + ")XXTable"

    If IntAccountType = 0 Then
        '«·√—’œ… «·„œÌ‰…
        StrSQL = StrSQL + " Where XXTable.CurrentAccount < 0"
    ElseIf IntAccountType = 1 Then
        '«·√—’œ… «·œ«∆‰…
        StrSQL = StrSQL + " Where XXTable.CurrentAccount > 0"
    End If

    CustomersAccountsSQL = StrSQL
End Function

Private Function LoadSalTypeTrans() As TransactionsValues
    'Â‰« ‰ﬁÊ„ »«·√” ⁄·«„ ⁄‰
    '⁄‰ «·„»Ì⁄«  «· Ã«—Ï Ê«·ﬁÿ«⁄Ï
    Dim rs As ADODB.Recordset
    Dim StrSQL As String

    If SystemOptions.SysDataBaseType = SQLServerDataBase Then
        StrSQL = "SELECT SUM(CASE WHEN SAleType=0 THEN QryTransactionsTotal.TotalAfterTax ELSE 0 END) AS SumSaleType0 "
        StrSQL = StrSQL + ",SUM(CASE WHEN SAleType=1 THEN QryTransactionsTotal.TotalAfterTax ELSE 0 END) AS SumSaleType1"
        StrSQL = StrSQL + " FROM         dbo.Transactions INNER JOIN "
        StrSQL = StrSQL + "dbo.QryTransactionsTotal() QryTransactionsTotal ON dbo.Transactions.Transaction_ID = QryTransactionsTotal.Transaction_ID"
        StrSQL = StrSQL + " WHERE     (QryTransactionsTotal.Transaction_Type=2) or (QryTransactionsTotal.Transaction_Type=21) AND ((dbo.Transactions.SaleType = 0) OR"
        StrSQL = StrSQL + " (dbo.Transactions.SaleType = 1))"

        If Not (IsNull(Me.DtpFrom.value)) Then
            StrSQL = StrSQL + " AND QryTransactionsTotal.Transaction_Date >=" & SQLDate(Me.DtpFrom.value, True)
        End If

        If Not (IsNull(Me.DtpTo.value)) Then
            StrSQL = StrSQL + " AND QryTransactionsTotal.Transaction_Date <=" & SQLDate(Me.DtpTo.value, True)
        End If
    
    End If

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.BOF Or rs.EOF) Then
        LoadSalTypeTrans.TotalCash = Format(IIf(IsNull(rs("SumSaleType0").value), 0, rs("SumSaleType0").value), SystemOptions.SysDefCurrencyForamt)
        LoadSalTypeTrans.TotalDue = Format(IIf(IsNull(rs("SumSaleType1").value), 0, rs("SumSaleType1").value), SystemOptions.SysDefCurrencyForamt)
        LoadSalTypeTrans.TotalNet = LoadSalTypeTrans.TotalCash + LoadSalTypeTrans.TotalDue
    End If

    rs.Close
    Set rs = Nothing
End Function
