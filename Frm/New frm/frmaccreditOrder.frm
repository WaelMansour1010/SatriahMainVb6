VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Begin VB.Form frmaccreditOrder 
   Caption         =   "ÿ·»«  «·⁄„·«¡ «· Ì  „ «⁄ „«œÂ«"
   ClientHeight    =   7680
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10920
   HelpContextID   =   440
   Icon            =   "frmaccreditOrder.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   7680
   ScaleWidth      =   10920
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
      Height          =   7680
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   10920
      _cx             =   19262
      _cy             =   13547
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
      GridRows        =   3
      GridCols        =   2
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   $"frmaccreditOrder.frx":038A
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   990
         Left            =   30
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   6660
         Width           =   10860
         _cx             =   19156
         _cy             =   1746
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
         Begin VB.CheckBox ChkShow 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "·«  ŸÂ— Â–Â «·‰«›–… ⁄‰œ  ‘€Ì· «·»—‰«„Ã"
            ForeColor       =   &H000000FF&
            Height          =   480
            Left            =   4650
            RightToLeft     =   -1  'True
            TabIndex        =   4
            Top             =   450
            Visible         =   0   'False
            Width           =   5940
         End
         Begin ImpulseButton.ISButton CmdExit 
            Cancel          =   -1  'True
            Height          =   375
            Left            =   105
            TabIndex        =   5
            Top             =   495
            Width           =   1065
            _ExtentX        =   1879
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
            ButtonImage     =   "frmaccreditOrder.frx":03E4
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
            Height          =   375
            Left            =   1335
            TabIndex        =   6
            Top             =   495
            Width           =   960
            _ExtentX        =   1693
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
            ButtonImage     =   "frmaccreditOrder.frx":077E
            ColorButton     =   14871017
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            ColorTextShadow =   -2147483637
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "Ì „  ÕœÌœ Â–Â «·»Ì«‰«  »‰«¡« ⁄·Ï «· «—ÌŒ «·Õ«·Ì ›Ì «·ÃÂ«“"
            ForeColor       =   &H000000FF&
            Height          =   315
            Left            =   5250
            RightToLeft     =   -1  'True
            TabIndex        =   7
            Top             =   180
            Width           =   5505
         End
      End
      Begin VSFlex8UCtl.VSFlexGrid Fg 
         Height          =   6015
         Left            =   30
         TabIndex        =   2
         Top             =   630
         Width           =   10860
         _cx             =   19156
         _cy             =   10610
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
         Cols            =   9
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmaccreditOrder.frx":0B18
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
         ExplorerBar     =   1
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
      Begin VB.Image Image1 
         Height          =   480
         Left            =   30
         Picture         =   "frmaccreditOrder.frx":0C9E
         Top             =   30
         Width           =   480
      End
      Begin VB.Label LblCaption 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "ÿ·»«  «·⁄„·«¡ «· Ì  „ «⁄ „«œÂ«"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   585
         Left            =   30
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   30
         Width           =   10860
      End
   End
End
Attribute VB_Name = "frmaccreditOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Askinterval As String
Dim Askcount As Integer

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CmdPrint_Click()
    On Error GoTo ErrTrap
    Dim Reports As ClsRepoerts
    Dim StrSQL As String

    Askinterval = GetSetting(StrAppRegPath, "Setting", "INTERVAL_InstallmentMustPayed", True)
    Askcount = GetSetting(StrAppRegPath, "Setting", "count_InstallmentMustPayed", True)
    
    'StrSQL = "select * From QestNotReceipted where  DueDate<='" & SQLDate(DateAdd(Askinterval, Askcount, Date)) & "'"
    ' StrSQL = StrSQL + " order by CusName,Transaction_ID,QeqtNum"

    StrSQL = "SELECT     dbo.Transactions.Transaction_ID, dbo.Transactions.Transaction_Date, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.Transactions.Posted, "
    StrSQL = StrSQL + "  dbo.TblUsers.UserName , dbo.Transactions.order_no,  dbo.Transactions.PostedDate"
    StrSQL = StrSQL + " FROM         dbo.Transactions INNER JOIN"
    StrSQL = StrSQL + " dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
    StrSQL = StrSQL + " dbo.TblUsers ON dbo.Transactions.Posted = dbo.TblUsers.UserID"
    StrSQL = StrSQL + " WHERE     (NOT (dbo.Transactions.Posted IS NULL)) AND (dbo.Transactions.order_no NOT IN"
    StrSQL = StrSQL + " (SELECT     order_no"
    StrSQL = StrSQL + " From Transactions"
    StrSQL = StrSQL + " WHERE     Transaction_Type = 21 AND NOT (order_no IS NULL))) AND (dbo.Transactions.Transaction_Type = 17)"
    StrSQL = StrSQL + " ORDER BY dbo.Transactions.PostedDate"
   
    Set Reports = New ClsRepoerts
    Reports.AccreditOrders StrSQL, , LblCaption.Caption
    Exit Sub
ErrTrap:
End Sub

Private Sub FG_CellButtonClick(ByVal row As Long, _
                               ByVal Col As Long)

    With Me.FG

        Select Case .ColKey(Col)

            Case "Convert"
                frmsalebill.show
                frmsalebill.NewBillFromOrder .TextMatrix(row, .ColIndex("order_no"))
        End Select

    End With

End Sub

Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim My_SQL As String
    Dim RowNum As Integer
    Dim ReCount As Integer
    Dim RsTemp As New ADODB.Recordset
    Dim BGround As New ClsBackGroundPic
    Dim BolShowRequest As Boolean

    FormPostion Me, GetPostion
    LoadIcons

    If SystemOptions.SysDataBaseType = AccessDataBase Then
        My_SQL = "Select * From QestNotReceipted where  DueDate <=#" & SQLDate(Date) & "#"
        My_SQL = My_SQL + " order by CusName,Transaction_ID,QeqtNum"
    ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
     
        '    My_SQL = "Select * From QestNotReceipted where  DueDate <='" & SQLDate(DateAdd(Askinterval, Askcount, Date)) & "'"
        '    My_SQL = My_SQL + " order by CusName,Transaction_ID,QeqtNum"
        Dim StrSQL As String

        StrSQL = "SELECT     dbo.Transactions.Transaction_ID, dbo.Transactions.Transaction_Date, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.Transactions.Posted, "
        StrSQL = StrSQL + "  dbo.TblUsers.UserName , dbo.Transactions.order_no, "
        StrSQL = StrSQL + "  PostedDate = (SELECT TOP 1"
        StrSQL = StrSQL + "              ApprovDate"
        StrSQL = StrSQL + "  From ApprovalData"
        StrSQL = StrSQL + "  WHERE ScreenName = (CASE Transactions.Transaction_Type"
        StrSQL = StrSQL + "              WHEN 29 THEN 'FrmPO10'"
        StrSQL = StrSQL + "      Else 'FrmPO3'"
        StrSQL = StrSQL + "  END))"
        
        StrSQL = StrSQL + " FROM         dbo.Transactions INNER JOIN"
        StrSQL = StrSQL + " dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
        StrSQL = StrSQL + " dbo.TblUsers ON dbo.Transactions.Posted = dbo.TblUsers.UserID"
        StrSQL = StrSQL + " WHERE     (NOT (dbo.Transactions.Posted IS NULL)) AND (dbo.Transactions.order_no NOT IN"
        StrSQL = StrSQL + " (SELECT     order_no"
        StrSQL = StrSQL + " From Transactions"
        StrSQL = StrSQL + " WHERE     Transaction_Type = 21 AND NOT (order_no IS NULL))) AND (dbo.Transactions.Transaction_Type = 6 or dbo.Transactions.Transaction_Type = 29)"
        StrSQL = StrSQL + " and Transactions.Approved = 1"
        StrSQL = StrSQL + " ORDER BY "
        StrSQL = StrSQL + "   (SELECT TOP 1"
        StrSQL = StrSQL + "              ApprovDate"
        StrSQL = StrSQL + "  From ApprovalData"
        StrSQL = StrSQL + "  WHERE ScreenName = (CASE Transactions.Transaction_Type"
        StrSQL = StrSQL + "              WHEN 29 THEN 'FrmPO10'"
        StrSQL = StrSQL + "      Else 'FrmPO3'"
        StrSQL = StrSQL + "  END))"
        
       

    End If

    RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Not (RsTemp.EOF Or RsTemp.BOF) Then

        With FG
            .rows = .FixedRows

            For ReCount = 1 To RsTemp.RecordCount
                .rows = .rows + 1
                RowNum = .rows - 1

                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(RowNum, .ColIndex("CusName")) = IIf(IsNull(RsTemp("CusName").value), "", RsTemp("CusName").value)
                Else
                    .TextMatrix(RowNum, .ColIndex("CusName")) = IIf(IsNull(RsTemp("CusNamee").value), "", RsTemp("CusNamee").value)
                End If
        
                .TextMatrix(RowNum, .ColIndex("order_no")) = IIf(IsNull(RsTemp("order_no").value), "", RsTemp("order_no").value)
            
                .TextMatrix(RowNum, .ColIndex("Transaction_Date")) = IIf(IsNull(RsTemp("Transaction_Date").value), "", Format(RsTemp("Transaction_Date").value, "yyyy/mm/dd"))
            
                '          .TextMatrix(RowNum, .ColIndex("Transaction_Date")) = _
                           IIf(IsNull(RsTemp("Transaction_Date").value), "", Format(RsTemp("Transaction_Date").value, "yyyy/mm/dd"))
            
                .TextMatrix(RowNum, .ColIndex("PostedDate")) = IIf(IsNull(RsTemp("PostedDate").value), "", Format(RsTemp("PostedDate").value, "yyyy/mm/dd"))
                   
                .TextMatrix(RowNum, .ColIndex("UserName")) = IIf(IsNull(RsTemp("UserName").value), "", Format(RsTemp("UserName").value, "yyyy/mm/dd"))
            
                .ColComboList(.ColIndex("CC")) = "..."
             
                RsTemp.MoveNext
            Next ReCount

            .AutoSize 0, .Cols - 1, False
        End With

    End If

    FG.WallPaper = BGround.Picture
    BolShowRequest = GetSetting(StrAppRegPath, "View_Type", "InstallmentMustPayed", True)

    If BolShowRequest = True Then
        ChkShow.value = Unchecked
    Else
        ChkShow.value = Checked
    End If

    Resize_Form Me, ReportSize

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub ChangeLang()
    Me.Caption = "Accredit Sales Order"
    LblCaption.Caption = Me.Caption
    ChkShow.Caption = "Dont Show at Start"
    Label1.Caption = "Data Based in your System Date"
    Me.CmdExit.Caption = "Exit"
    Me.CmdPrint.Caption = "Print"

    With Me.FG
        .TextMatrix(0, .ColIndex("order_no")) = "Order No"
        .TextMatrix(0, .ColIndex("Transaction_Date")) = "Trans Date"
        .TextMatrix(0, .ColIndex("CusName")) = "Cus. Name"
        .TextMatrix(0, .ColIndex("PostedDate")) = "PostedDate"
        .TextMatrix(0, .ColIndex("UserName")) = "By User"
        .TextMatrix(0, .ColIndex("Convert")) = "Convert To Bill"

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

    With FG
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

