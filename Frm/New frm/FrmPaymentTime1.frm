VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Begin VB.Form FrmPaymentTime1 
   Caption         =   "»Ì«‰«  «·›Ê« Ì— «· Ì ·„  ”œœ"
   ClientHeight    =   6855
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13410
   HelpContextID   =   420
   Icon            =   "FrmPaymentTime1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   6855
   ScaleWidth      =   13410
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
      Height          =   6855
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   13410
      _cx             =   23654
      _cy             =   12091
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
      Align           =   5
      AutoSizeChildren=   8
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
      GridRows        =   7
      GridCols        =   5
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"FrmPaymentTime1.frx":038A
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin VB.Frame Frame3 
         BackColor       =   &H00E2E9E9&
         Height          =   360
         Left            =   810
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   435
         Visible         =   0   'False
         Width           =   2655
         Begin VB.Label lblvalue 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   240
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   16
            Top             =   120
            Width           =   1455
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "«·ﬁÌ„…"
            Height          =   255
            Left            =   1560
            RightToLeft     =   -1  'True
            TabIndex        =   15
            Top             =   120
            Width           =   975
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E2E9E9&
         Height          =   360
         Left            =   3480
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   435
         Visible         =   0   'False
         Width           =   4635
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "—ﬁ„ «·⁄„Ì·"
            Height          =   255
            Left            =   2400
            RightToLeft     =   -1  'True
            TabIndex        =   13
            Top             =   120
            Width           =   1695
         End
         Begin VB.Label lblcusid 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Height          =   240
            Left            =   960
            RightToLeft     =   -1  'True
            TabIndex        =   12
            Top             =   120
            Width           =   1215
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00E2E9E9&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   765
         Left            =   30
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   30
         Width           =   765
         Begin VB.CommandButton Cmd_Pic 
            Appearance      =   0  'Flat
            BackColor       =   &H00E2E9E9&
            Height          =   675
            Left            =   0
            MaskColor       =   &H00FF0000&
            RightToLeft     =   -1  'True
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   30
            UseMaskColor    =   -1  'True
            Visible         =   0   'False
            Width           =   735
         End
      End
      Begin VB.Timer Timer1 
         Interval        =   25
         Left            =   6660
         Top             =   0
      End
      Begin ImpulseButton.ISButton CmdRef 
         Height          =   390
         Left            =   3480
         TabIndex        =   6
         Top             =   30
         Visible         =   0   'False
         Width           =   4635
         _ExtentX        =   8176
         _ExtentY        =   688
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
         ButtonImage     =   "FrmPaymentTime1.frx":0439
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton CmdPrint 
         Height          =   390
         Index           =   0
         Left            =   810
         TabIndex        =   5
         Top             =   30
         Visible         =   0   'False
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   688
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
         ButtonImage     =   "FrmPaymentTime1.frx":07D3
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   450
         Index           =   0
         Left            =   30
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   6375
         Width           =   13350
         _cx             =   23548
         _cy             =   794
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
         Begin ImpulseButton.ISButton CmdExit 
            Height          =   390
            Left            =   195
            TabIndex        =   2
            Top             =   15
            Width           =   1680
            _ExtentX        =   2963
            _ExtentY        =   688
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
            ButtonImage     =   "FrmPaymentTime1.frx":0B6D
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
         Begin ImpulseButton.ISButton ISButton1 
            Cancel          =   -1  'True
            Height          =   390
            Left            =   2325
            TabIndex        =   10
            Top             =   0
            Width           =   1665
            _ExtentX        =   2937
            _ExtentY        =   688
            ButtonStyle     =   1
            ButtonPositionImage=   1
            Caption         =   "„Ê«›ﬁ"
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
      End
      Begin VSFlex8UCtl.VSFlexGrid Fg1 
         Height          =   5550
         Left            =   30
         TabIndex        =   3
         Top             =   810
         Width           =   13350
         _cx             =   23548
         _cy             =   9790
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
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   17
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FrmPaymentTime1.frx":0F07
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
      Begin ImpulseButton.ISButton CmdOptions 
         Height          =   390
         Left            =   11985
         TabIndex        =   9
         Top             =   30
         Visible         =   0   'False
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   688
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   "ŒÌ«—« ..."
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
         ButtonImage     =   "FrmPaymentTime1.frx":1207
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         RightToLeft     =   -1  'True
      End
      Begin VB.Label LblCaption 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«Œ «— «·›Ê« Ì— · ”œÌœÂ«"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   360
         Index           =   0
         Left            =   8130
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   435
         Width           =   5250
      End
   End
End
Attribute VB_Name = "FrmPaymentTime1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim StrAlramSoundPath As String
Dim BolPlaySound As Boolean
Dim SngTimer As Single
Dim first_run As Boolean

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CmdOptions_Click()
    'FrmAlramOptions.show vbModal
End Sub

Private Sub CmdPrint_Click(Index As Integer)
    Dim cNoteReport As ClsNotesReports

    If Index = 0 Then
        Set cNoteReport = New ClsNotesReports
        cNoteReport.ShowCompanyDebitValues 1, Null, Date, False, WindowTarget
        Set cNoteReport = Nothing
    ElseIf Index = 1 Then
        Set cNoteReport = New ClsNotesReports
        cNoteReport.ShowCompanyDebitValues 0, Null, Date, False, WindowTarget
        Set cNoteReport = Nothing
    End If

End Sub

Private Sub CmdRef_Click()
    LoadData (val(lblcusid.Caption))
End Sub

Private Sub Form_Activate()
    On Error GoTo ErrTrap

    If first_run = True Then
        CmdRef_Click
        first_run = False
    End If

    If BolPlaySound = True Then
        '    If Dir(StrAlramSoundPath) <> "" Then
        '        PlaySoundEx StrAlramSoundPath, False, True
        '    End If
    End If

    ShowDynamicHelp Me.HelpContextID
    Exit Sub
ErrTrap:
    'App.Path & "\Sound\ALARM3.WAV"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If Shift = 2 Then
        If KeyCode = vbKeyX Then
            If CmdExit.Enabled = False Then Exit Sub
            CmdExit_Click
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Load()
    Dim BGround As New ClsBackGroundPic
    Dim rs As ADODB.Recordset

    CenterForm Me

    FormPostion Me, GetPostion
    Set FG1.WallPaper = BGround.Picture

    FG1.Rows = FG1.FixedRows
 
    'Me.ChartPay.PointLabels = False
    'Me.ChartRecv.PointLabels = False
    'Resize_Form Me, ReportSize
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    first_run = True
End Sub

Private Sub ChangeLang()
    CmdOptions.Visible = False

    Me.Caption = "Securities outstanding"
    CmdOptions.Caption = "Option"
    CmdRef.Caption = "Reresh"
    LblCaption(0).Caption = "Amounts To receipt"
    'lbl(0).Caption = "No. of securities"
    CmdPrint(0).Caption = "Print"
    'LblCaption(1).Caption = "Amounts To Payed"
    'lbl(1).Caption = "No. of securities"
    'CmdPrint(1).Caption = "Print"

    CmdExit.Caption = "Exit"
    ISButton1.Caption = "OK"
    Label1.Caption = "Customer"
    Label2.Caption = "Value"

    With Me.FG1
        .TextMatrix(0, .ColIndex("pay")) = "pay"
        .TextMatrix(0, .ColIndex("RequiredValue")) = "Required Value"
        .TextMatrix(0, .ColIndex("Note_Value")) = "Note Value"
        .TextMatrix(0, .ColIndex("LateInterval")) = "Late Interval"
        .TextMatrix(0, .ColIndex("PreRelease")) = "PreRelease"
        .TextMatrix(0, .ColIndex("CusName")) = "Cust. Name"
        .TextMatrix(0, .ColIndex("DueDate")) = "DueDate"
        .TextMatrix(0, .ColIndex("TransactionTypeName")) = "Transaction Type  "
        .TextMatrix(0, .ColIndex("Transaction_Serial")) = "Transaction Serial"
        .TextMatrix(0, .ColIndex("Transaction_Date")) = "Transaction Date"
        .TextMatrix(0, .ColIndex("NotesTypeName")) = "Notes Type"
    End With

End Sub

Private Sub Form_Resize()
    'If Me.WindowState = vbMaximized Then
    '    Me.ChartRecv.LegendBox = True
    '    Me.ChartPay.LegendBox = True
    'Else
    '    Me.ChartRecv.LegendBox = False
    '    Me.ChartPay.LegendBox = False
    'End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap

    FormPostion Me, SavePostion
 
    FormPostion Me, SavePostion
    Exit Sub
ErrTrap:
End Sub

Private Sub ISButton1_Click()
    Dim SQLString As String
    Dim lblinvoices As String
    lblinvoices = "«·›Ê« Ì— «·„Õœœ… " & CHR(13)
    Dim i As Integer

    With FG1

        For i = 1 To .Rows - 1

            If .Cell(flexcpChecked, i, .ColIndex("pay")) = flexChecked Then
              
                SQLString = SQLString & " or TransactionsID=" & val(.TextMatrix(i, .ColIndex("TransactionsID")))
                lblinvoices = lblinvoices & "," & (.TextMatrix(i, .ColIndex("Transaction_Serial")))
                
                
            End If
               
            ' .TextMatrix(I, .ColIndex("TransactionsID")) = IIf(IsNull(RsCreditValues("TransactionsID").value), "", RsCreditValues("TransactionsID").value)
        Next i

        If Len(SQLString) > 4 Then
            FrmCashing.lblsqlstring.Caption = "(" & mId(SQLString, 4, Len(SQLString) - 3) & ")"
           FrmCashing.XPMTxtRemarks = FrmCashing.XPMTxtRemarks & lblinvoices
           
           
            
        End If

    End With

    Unload Me

End Sub

Private Sub LblCaption_DblClick(Index As Integer)
    On Error GoTo ErrTrap

    If Index = 0 Then
        If Me.WindowState = vbNormal Then
            Me.WindowState = vbMaximized
        Else
            Me.WindowState = vbNormal
        End If

        Exit Sub
    End If

ErrTrap:
End Sub

Private Sub LoadData(CusID As Double)
    Dim StrSQL As String
    Dim i As Integer

   ' StrSQL = "SELECT CompanyCreditValues.*  FROM dbo.CompanyCreditValues() CompanyCreditValues   where   (cusid=" & CusID & " and requiredvalue>0  and (transaction_type=21 or transaction_type=1) ) order by duedate"
    
  StrSQL = "  SELECT     TOP 100 PERCENT CompanyCreditValues.*, dbo.Transactions.NoteSerial1 "
StrSQL = StrSQL & "  FROM         dbo.CompanyCreditValues() CompanyCreditValues LEFT OUTER JOIN"
StrSQL = StrSQL & "   dbo.Transactions ON CompanyCreditValues.TransactionsID = dbo.Transactions.Transaction_ID"
StrSQL = StrSQL & "  WHERE     (CompanyCreditValues.CusID = " & CusID & ") AND (CompanyCreditValues.RequiredValue > 0) AND (CompanyCreditValues.Transaction_Type = 21 OR"
StrSQL = StrSQL & "     CompanyCreditValues.Transaction_Type = 1)"
StrSQL = StrSQL & "   ORDER BY CompanyCreditValues.DueDate"

    
    Dim RsCreditValues As New ADODB.Recordset
    ' RsCreditValues.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText + adAsyncExecute + adAsyncFetch
    RsCreditValues.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    'Rs3.Open Sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If Not (RsCreditValues.BOF Or RsCreditValues.EOF) Then

        With FrmPaymentTime1.FG1
            .Rows = .FixedRows + RsCreditValues.RecordCount
            RsCreditValues.MoveFirst

            For i = 1 To RsCreditValues.RecordCount
                .TextMatrix(i, .ColIndex("TransactionsID")) = IIf(IsNull(RsCreditValues("TransactionsID").value), "", RsCreditValues("TransactionsID").value)
                .TextMatrix(i, .ColIndex("Transaction_Type")) = IIf(IsNull(RsCreditValues("Transaction_Type").value), "", RsCreditValues("Transaction_Type").value)
                .TextMatrix(i, .ColIndex("TransactionTypeName")) = IIf(IsNull(RsCreditValues("TransactionTypeName").value), "", RsCreditValues("TransactionTypeName").value)
                .TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(RsCreditValues("CusName").value), "", RsCreditValues("CusName").value)
                .TextMatrix(i, .ColIndex("NoteID")) = IIf(IsNull(RsCreditValues("NoteID").value), "", RsCreditValues("NoteID").value)
                .TextMatrix(i, .ColIndex("NoteType")) = IIf(IsNull(RsCreditValues("NoteType").value), "", RsCreditValues("NoteType").value)
                .TextMatrix(i, .ColIndex("Note_Value")) = IIf(IsNull(RsCreditValues("Note_Value").value), "", RsCreditValues("Note_Value").value)

                If Not IsNull(RsCreditValues("DueDate").value) Then
                    .TextMatrix(i, .ColIndex("DueDate")) = DisplayDate(RsCreditValues("DueDate").value)
                    .TextMatrix(i, .ColIndex("LateInterval")) = DateDiff("d", RsCreditValues("DueDate").value, Date)
                End If

                .TextMatrix(i, .ColIndex("Transaction_Serial")) = IIf(IsNull(RsCreditValues("NoteSerial1").value), "", RsCreditValues("NoteSerial1").value)

                If Not IsNull(RsCreditValues("Transaction_Date").value) Then
                    .TextMatrix(i, .ColIndex("Transaction_Date")) = DisplayDate(RsCreditValues("Transaction_Date").value)
                End If

                .TextMatrix(i, .ColIndex("CusID")) = IIf(IsNull(RsCreditValues("CusID").value), "", RsCreditValues("CusID").value)

                If Not IsNull(RsCreditValues("NoteDate").value) Then
                    .TextMatrix(i, .ColIndex("NoteDate")) = DisplayDate(RsCreditValues("NoteDate").value)
                End If

                .TextMatrix(i, .ColIndex("NotesTypeName")) = IIf(IsNull(RsCreditValues("NotesTypeName").value), "", RsCreditValues("NotesTypeName").value)
                .TextMatrix(i, .ColIndex("PreRelease")) = IIf(IsNull(RsCreditValues("PreRelease").value), "", RsCreditValues("PreRelease").value)
                .TextMatrix(i, .ColIndex("RequiredValue")) = IIf(IsNull(RsCreditValues("RequiredValue").value), "", RsCreditValues("RequiredValue").value)
                RsCreditValues.MoveNext
            Next i

            '   DrawFloodProgress FrmPaymentTime.Fg1, .ColIndex("LateInterval")
            .AutoSize 0, .Cols - 1, False
        End With

    End If

End Sub

Public Sub ApplySetting()
    Dim IntDateDiff As Integer
    Dim rs As ADODB.Recordset
    Dim RowNum As Integer
    Dim SngColor1 As Single, SngColor2 As Single, SngColor3 As Single

    Set rs = New ADODB.Recordset
    rs.Open "TblOptions", Cn, adOpenStatic, adLockReadOnly, adCmdTable

    If rs("PlayNotesAlramSound").value = 1 Then
        BolPlaySound = True
    Else
        BolPlaySound = False
    End If

    StrAlramSoundPath = IIf(IsNull(rs("AlramSoundFilePath").value), App.path & "\Sound\ALARM3.WAV", rs("AlramSoundFilePath").value)

    If rs("EnableNotesAlramColors").value = 1 Then
        SngColor1 = IIf(IsNull(rs("Color1").value), vbWhite, rs("Color1").value)
        SngColor2 = IIf(IsNull(rs("Color2").value), vbWhite, rs("Color2").value)
        SngColor3 = IIf(IsNull(rs("Color3").value), vbWhite, rs("Color3").value)
    Else
        Me.FG1.Clear flexClearScrollable, flexClearFormatting
   
        Exit Sub
    End If

    With Me.FG1

        For RowNum = .FixedRows To .Rows - 1
            IntDateDiff = DateDiff("d", .TextMatrix(RowNum, .ColIndex("DueDate")), Date)

            If IntDateDiff > 0 Then
                .Cell(flexcpBackColor, RowNum, 0, RowNum, .Cols - 1) = SngColor1
            ElseIf IntDateDiff = 0 Then
                .Cell(flexcpBackColor, RowNum, 0, RowNum, .Cols - 1) = SngColor2
            Else
                .Cell(flexcpBackColor, RowNum, 0, RowNum, .Cols - 1) = SngColor3
            End If

        Next RowNum

    End With

End Sub
