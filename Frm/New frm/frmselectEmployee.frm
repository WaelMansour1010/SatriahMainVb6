VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Begin VB.Form FrmSelectEmployee 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "  ÕœÌœ «·„‰«œÌ»"
   ClientHeight    =   7905
   ClientLeft      =   -15
   ClientTop       =   375
   ClientWidth     =   8625
   HelpContextID   =   440
   Icon            =   "frmselectEmployee.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   7905
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
      Height          =   8895
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   8655
      _cx             =   15266
      _cy             =   15690
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
         Height          =   2175
         Left            =   30
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   6690
         Width           =   8565
         _cx             =   15108
         _cy             =   3836
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
         Begin VB.TextBox txtMessage 
            Alignment       =   1  'Right Justify
            Height          =   1590
            Left            =   10320
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   10
            Top             =   0
            Width           =   5205
         End
         Begin VB.CheckBox ChkShow 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "·«  ŸÂ— Â–Â «·‰«›–… ⁄‰œ  ‘€Ì· «·»—‰«„Ã"
            ForeColor       =   &H000000FF&
            Height          =   1050
            Left            =   3465
            RightToLeft     =   -1  'True
            TabIndex        =   4
            Top             =   2835
            Width           =   4875
         End
         Begin ImpulseButton.ISButton CmdExit 
            Cancel          =   -1  'True
            Height          =   810
            Left            =   480
            TabIndex        =   5
            Top             =   480
            Width           =   705
            _ExtentX        =   1244
            _ExtentY        =   1429
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
            ButtonImage     =   "frmselectEmployee.frx":038A
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
            Height          =   825
            Left            =   1950
            TabIndex        =   6
            Top             =   2100
            Visible         =   0   'False
            Width           =   450
            _ExtentX        =   794
            _ExtentY        =   1455
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
            ButtonImage     =   "frmselectEmployee.frx":0724
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
            Height          =   585
            Left            =   1560
            TabIndex        =   8
            Top             =   600
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   1032
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
            DisabledImageExtraction=   0
            ColorToggledHoverText=   16711680
            LowerToggledContent=   0   'False
            ColorTextShadow =   4210752
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«·—”«·…"
            Height          =   510
            Left            =   7620
            RightToLeft     =   -1  'True
            TabIndex        =   9
            Top             =   420
            Width           =   3045
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "Ì „  ÕœÌœ Â–Â «·»Ì«‰«  »‰«¡« ⁄·Ï «· «—ÌŒ «·Õ«·Ì ›Ì «·ÃÂ«“"
            ForeColor       =   &H000000FF&
            Height          =   675
            Left            =   4035
            RightToLeft     =   -1  'True
            TabIndex        =   7
            Top             =   1845
            Visible         =   0   'False
            Width           =   4425
         End
      End
      Begin VSFlex8UCtl.VSFlexGrid Fg 
         Height          =   5760
         Left            =   30
         TabIndex        =   2
         Top             =   915
         Width           =   8520
         _cx             =   15028
         _cy             =   10160
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
         Cols            =   11
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmselectEmployee.frx":0ABE
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
         Height          =   375
         Left            =   7080
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   645
         Width           =   1425
      End
      Begin VB.Label lblFlag 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         Height          =   135
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   240
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   30
         Picture         =   "frmselectEmployee.frx":0C70
         Top             =   30
         Width           =   480
      End
      Begin VB.Label LblCaption 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "  ÕœÌœ «·„‰«œÌ»"
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
         Height          =   630
         Left            =   30
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   0
         Width           =   8580
      End
   End
End
Attribute VB_Name = "FrmSelectEmployee"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Askinterval As String
Dim Askcount As Integer
Public supplierVendor As Integer

Private Sub Check17_Click()
    Dim i As Integer

    If Check17.value = vbChecked Then

        With Me.FG
 
            For i = 1 To .Rows - 1
        
                .TextMatrix(i, .ColIndex("Send")) = True
            Next i

        End With

    Else

        With Me.FG

            For i = 1 To .Rows - 1
        
                .TextMatrix(i, .ColIndex("Send")) = False
            Next i

        End With

    End If

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

Private Sub FG_BeforeEdit(ByVal Row As Long, _
                          ByVal Col As Long, _
                          Cancel As Boolean)

    If Col <> FG.ColIndex("Send") Then
        Cancel = True
    End If

End Sub

Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim My_SQL As String
    Dim RowNum As Integer
    Dim ReCount As Integer
    Dim RsTemp As New ADODB.Recordset
    Dim BGround As New ClsBackGroundPic
    Dim BolShowRequest As Boolean

    Me.Left = (mdifrmmain.Width - Me.Width) / 2
    Me.Top = (mdifrmmain.Height - Me.Height) / 2 - 500
    
    'FormPostion Me, GetPostion
    LoadIcons

    If SystemOptions.SysDataBaseType = AccessDataBase Then
        My_SQL = "Select * From QestNotReceipted where  DueDate <=#" & SQLDate(Date) & "#"
        My_SQL = My_SQL + " order by CusName,Transaction_ID,QeqtNum"
    ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then

 
        '    My_SQL = "Select * From QestNotReceipted where  DueDate <='" & SQLDate(DateAdd(Askinterval, Askcount, Date)) & "'"
        '    My_SQL = My_SQL + " order by CusName,Transaction_ID,QeqtNum"
        Dim StrSQL As String
        
        If supplierVendor = 0 Then
          If SystemOptions.IsMashghal Then
        StrSQL = " SELECT     dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Namee, dbo.TblEmployee.Emp_ID"
         StrSQL = StrSQL & " FROM         "
         StrSQL = StrSQL & "  dbo.TblEmployee "
         StrSQL = StrSQL & " where ( dbo.TblEmployee.BranchId is null or dbo.TblEmployee.BranchId in(" & Current_branchSql & "))  and IsNull(TblEmployee.chkShowTasks,0) = 1"
    Else
   StrSQL = " SELECT     dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Namee, dbo.TBLSalesRepData.EmpID as Emp_ID"
   StrSQL = StrSQL & " FROM         dbo.TBLSalesRepData INNER JOIN"
   StrSQL = StrSQL & "  dbo.TblEmployee ON dbo.TBLSalesRepData.EmpID = dbo.TblEmployee.Emp_ID"
   StrSQL = StrSQL & " where ( dbo.TBLSalesRepData.BranchID is null or dbo.TBLSalesRepData.BranchID in(" & Current_branchSql & "))  "
   End If
Else
  
  If SystemOptions.IsMashghal Then
        StrSQL = " SELECT     dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Namee, dbo.TblEmployee.Emp_ID"
         StrSQL = StrSQL & " FROM         "
         StrSQL = StrSQL & "  dbo.TblEmployee "
         StrSQL = StrSQL & " where ( dbo.TblEmployee.BranchId is null or dbo.TblEmployee.BranchId in(" & Current_branchSql & "))  and IsNull(TblEmployee.chkShowTasks,0) = 1"
    Else
            StrSQL = " SELECT     dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Namee, dbo.TBLSalesRepData2.EmpID"
         StrSQL = StrSQL & " FROM         dbo.TBLSalesRepData2 INNER JOIN"
         StrSQL = StrSQL & "  dbo.TblEmployee ON dbo.TBLSalesRepData2.EmpID = dbo.TblEmployee.Emp_ID"
         StrSQL = StrSQL & " where ( dbo.TBLSalesRepData2.BranchID is null or dbo.TBLSalesRepData2.BranchID in(" & Current_branchSql & "))  "

    End If
End If

    End If

    RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Not (RsTemp.EOF Or RsTemp.BOF) Then

        With FG
            .Rows = .FixedRows

            For ReCount = 1 To RsTemp.RecordCount
                .Rows = .Rows + 1
                RowNum = .Rows - 1
                   
                ', dbo.QryCust_Qest.CustID
                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(RowNum, .ColIndex("Name")) = IIf(IsNull(RsTemp("Emp_Name").value), "", RsTemp("Emp_Name").value)
                Else
                    .TextMatrix(RowNum, .ColIndex("Name")) = IIf(IsNull(RsTemp("Emp_Namee").value), "", RsTemp("Emp_Name").value)

                End If
 
                .TextMatrix(RowNum, .ColIndex("Emp_Code")) = IIf(IsNull(RsTemp("Emp_Code").value), "", RsTemp("Emp_Code").value)
                .TextMatrix(RowNum, .ColIndex("EmpID")) = IIf(IsNull(RsTemp("Emp_ID").value), "", RsTemp("Emp_ID").value)
            
                RsTemp.MoveNext
            Next ReCount

            .AutoSize 0, .Cols - 1, False
        End With

    End If



RsTemp.Close
Set RsTemp = Nothing

 


    FG.WallPaper = BGround.Picture
    BolShowRequest = GetSetting(StrAppRegPath, "View_Type", "InstallmentMustPayed", True)

    If BolShowRequest = True Then
        ChkShow.value = Unchecked
    Else
        ChkShow.value = Checked
    End If

    'Resize_Form Me, ReportSize
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub ChangeLang()
    Me.Caption = "Installment Must Pay"
    LblCaption.Caption = Me.Caption
    ChkShow.Caption = "Dont Show at Start"
    Label1.Caption = "Data Based in your System Date"
    Me.CmdExit.Caption = "Exit"
    Me.CmdPrint.Caption = "Print"
   SendMessage.Caption = "Ok"
   Check17.RightToLeft = False
   Check17.Caption = "Select All"
   LblCaption.Caption = "Employee"
   FrmSelectEmployee.Caption = LblCaption.Caption
    With Me.FG
        .TextMatrix(0, .ColIndex("Name")) = "Name"
        .TextMatrix(0, .ColIndex("BillIID")) = "BillI ID"
        .TextMatrix(0, .ColIndex("TransDate")) = "Trans Date"
        .TextMatrix(0, .ColIndex("QestNum")) = "Installm. #"
        .TextMatrix(0, .ColIndex("DueDate")) = "DueDate"
        .TextMatrix(0, .ColIndex("value")) = "Value"
        .TextMatrix(0, .ColIndex("Emp_Code")) = "Code"
        .TextMatrix(0, .ColIndex("Send")) = "Select"
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

Function GetNumbers()

End Function

Private Sub SendMessage_Click()
    Dim Numbers As String
    Dim Names As String
    Dim RowNum As Integer
    Dim Opt As Integer
    Dim CurrentMessage As String
    Numbers = "0"
If Me.lblflag.Caption = 0 Then
   FrmReports.CurrenrEmployeeIDs = ""
   FrmReports.CurrenrEmployeeNames = ""
End If
    With FG

        For RowNum = .FixedRows To .Rows - 1
    
            If .Cell(flexcpChecked, RowNum, .ColIndex("Send")) = flexChecked Then

                '  MsgBox (.TextMatrix(RowNum, .ColIndex("Numbers")))
                If (.TextMatrix(RowNum, .ColIndex("EmpID"))) <> "" Then
                    If Numbers = "" Then
                        Numbers = (.TextMatrix(RowNum, .ColIndex("EmpID")))
                    Else
                        Numbers = Numbers & "," & (.TextMatrix(RowNum, .ColIndex("EmpID")))
                    End If
             
             
             
                               If Names = "" Then
                        Names = (.TextMatrix(RowNum, .ColIndex("Name")))
                    Else
                        Names = Names & "," & (.TextMatrix(RowNum, .ColIndex("Name")))
                    End If
                    
                    
                End If
            End If
          
        Next RowNum
      
 
        If Numbers = "0" Then Exit Sub
                                     
    End With

If Me.lblflag.Caption = 0 Then
FrmReports.CurrenrEmployeeIDs = Numbers
FrmReports.CurrenrEmployeeNames = Names
ElseIf Me.lblflag.Caption = 1 Then
Ageng_all.CurrenrEmployeeIDs = Numbers
ElseIf Me.lblflag.Caption = 2 Then
FrmVizitScreen.CurrenrEmployeeIDs = Numbers
End If
Unload Me
End Sub

