VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Begin VB.Form FrmInstallmentMustPay 
   Caption         =   "«·√ﬁ”«ÿ «·„ÿ·Ê»…"
   ClientHeight    =   8430
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11790
   HelpContextID   =   440
   Icon            =   "FrmInstallmentMustPay.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   8430
   ScaleWidth      =   11790
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
      Height          =   8430
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   11790
      _cx             =   20796
      _cy             =   14870
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
      _GridInfo       =   $"FrmInstallmentMustPay.frx":038A
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin C1SizerLibCtl.C1Elastic Ele 
         Height          =   990
         Left            =   30
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   7410
         Width           =   11730
         _cx             =   20690
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
         Begin VB.CheckBox Check17 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   " ÕœÌœ «·ﬂ·"
            Height          =   195
            Left            =   1440
            RightToLeft     =   -1  'True
            TabIndex        =   11
            Top             =   240
            Width           =   1095
         End
         Begin VB.Frame Frame1 
            Caption         =   "œ·«·«  «·«·Ê«‰"
            Height          =   735
            Left            =   3750
            RightToLeft     =   -1  'True
            TabIndex        =   8
            Top             =   120
            Width           =   3225
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               Caption         =   "«·ﬁ”ÿ „”œœ Ã“¡ „‰…"
               Height          =   255
               Left            =   360
               RightToLeft     =   -1  'True
               TabIndex        =   9
               Top             =   240
               Width           =   1815
            End
            Begin VB.Shape Shape1 
               FillColor       =   &H0000C000&
               FillStyle       =   0  'Solid
               Height          =   255
               Left            =   2280
               Top             =   240
               Width           =   255
            End
         End
         Begin VB.CheckBox ChkShow 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "·«  ŸÂ— Â–Â «·‰«›–… ⁄‰œ  ‘€Ì· «·»—‰«„Ã"
            ForeColor       =   &H000000FF&
            Height          =   480
            Left            =   4995
            RightToLeft     =   -1  'True
            TabIndex        =   4
            Top             =   450
            Width           =   6465
         End
         Begin ImpulseButton.ISButton CmdExit 
            Cancel          =   -1  'True
            Height          =   375
            Left            =   105
            TabIndex        =   5
            Top             =   495
            Width           =   1155
            _ExtentX        =   2037
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
            ButtonImage     =   "FrmInstallmentMustPay.frx":03E4
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
            Left            =   2520
            TabIndex        =   6
            Top             =   495
            Width           =   1005
            _ExtentX        =   1773
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
            ButtonImage     =   "FrmInstallmentMustPay.frx":077E
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
            Height          =   375
            Left            =   1440
            TabIndex        =   10
            Top             =   480
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   661
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
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "Ì „  ÕœÌœ Â–Â «·»Ì«‰«  »‰«¡« ⁄·Ï «· «—ÌŒ «·Õ«·Ì ›Ì «·ÃÂ«“"
            ForeColor       =   &H000000FF&
            Height          =   315
            Left            =   5580
            RightToLeft     =   -1  'True
            TabIndex        =   7
            Top             =   180
            Width           =   6045
         End
      End
      Begin VSFlex8UCtl.VSFlexGrid Fg 
         Height          =   6765
         Left            =   30
         TabIndex        =   2
         Top             =   630
         Width           =   11730
         _cx             =   20690
         _cy             =   11933
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
         Cols            =   21
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FrmInstallmentMustPay.frx":0B18
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
      Begin VB.Image Image1 
         Height          =   585
         Left            =   30
         Picture         =   "FrmInstallmentMustPay.frx":0E24
         Top             =   30
         Width           =   1125
      End
      Begin VB.Label LblCaption 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "«·√ﬁ”«ÿ «·„” Õﬁ… Œ·«· › —…"
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
         Width           =   11730
      End
   End
End
Attribute VB_Name = "FrmInstallmentMustPay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Askinterval As String
Dim Askcount As Integer

Private Sub Check17_Click()
    Dim i As Integer

    If Check17.value = vbChecked Then

        With Me.FG
 
            For i = 1 To .Rows - 2
        
                .TextMatrix(i, .ColIndex("Send")) = True
            Next i

        End With

    Else

        With Me.FG

            For i = 1 To .Rows - 2
        
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
    StrSQL = StrSQL + "  and DueDate ='" & SQLDate(DateAdd(Askinterval, Askcount, Date)) & "'"
    StrSQL = StrSQL + "  order by CusName,QryCust_Qest.Transaction_ID,QeqtNum"
 
    Set Reports = New ClsRepoerts
    Reports.QestMustPayed StrSQL, , LblCaption.Caption
    Exit Sub
ErrTrap:
End Sub

Private Sub FG_BeforeEdit(ByVal Row As Long, _
                          ByVal Col As Long, _
                          Cancel As Boolean)

    If Col <> FG.ColIndex("Send") And Col <> FG.ColIndex("Show") Then
        Cancel = True
    End If

End Sub
Private Sub FG_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
With FG
Select Case .ColKey(Col)
Case "Show"
Load FrmCustemers
FrmCustemers.Retrive val(.TextMatrix(Row, .ColIndex("CustID")))
FrmCustemers.show
End Select
End With
End Sub

Private Sub fg_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With FG
Select Case .ColKey(Col)
Case "Show"
.ColComboList(.ColIndex("Show")) = "..."
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
        Askinterval = GetSetting(StrAppRegPath, "Setting", "INTERVAL_InstallmentMustPayed", True)
        Askcount = GetSetting(StrAppRegPath, "Setting", "count_InstallmentMustPayed", True)

        If Askinterval = "D" Then
            LblCaption.Caption = LblCaption.Caption & Askcount & "  ÌÊ„  "
        ElseIf Askinterval = "M" Then
            LblCaption.Caption = LblCaption.Caption & Askcount & "  ‘Â—  "
        ElseIf Askinterval = "Y" Then
            LblCaption.Caption = LblCaption.Caption & Askcount & "  ”‰…  "
        End If
    
        '    My_SQL = "Select * From QestNotReceipted where  DueDate <='" & SQLDate(DateAdd(Askinterval, Askcount, Date)) & "'"
        '    My_SQL = My_SQL + " order by CusName,Transaction_ID,QeqtNum"
        Dim StrSQL As String
        'StrSQL = "SELECT     TOP 100 PERCENT dbo.QryCust_Qest.QestID, dbo.QryCust_Qest.NoteID, dbo.QryCust_Qest.QeqtNum, dbo.QryCust_Qest.PartID, dbo.QryCust_Qest.[Value], "
        'StrSQL = StrSQL + " dbo.QryCust_Qest.DueDate, dbo.QryCust_Qest.Receipt, dbo.QryCust_Qest.Summition, dbo.QryCust_Qest.CustID, dbo.QryCust_Qest.CusName,"
        'StrSQL = StrSQL + "  dbo.QryCust_Qest.Transaction_ID , dbo.QryCust_Qest.Transaction_Date, dbo.Transactions.NoteSerial1"
        'StrSQL = StrSQL + " FROM         dbo.QryCust_Qest LEFT OUTER JOIN"
        'StrSQL = StrSQL + "  dbo.Transactions ON dbo.QryCust_Qest.Transaction_ID = dbo.Transactions.Transaction_ID"
        'StrSQL = StrSQL + " WHERE     (dbo.QryCust_Qest.QestID NOT IN"
        'StrSQL = StrSQL + " (SELECT     QestID"
        'StrSQL = StrSQL + "  from InstallmentDet_Junc_Receipt"
        'StrSQL = StrSQL + " WHERE     Status <> 1))"
        'StrSQL = StrSQL + "  and DueDate  ='" & SQLDate(DateAdd(Askinterval, Askcount, Date)) & "'"
        'StrSQL = StrSQL + "  order by CusName,QryCust_Qest.Transaction_ID,QeqtNum"
StrSQL = " SELECT     TOP 100 PERCENT dbo.QryCust_Qest.QestID, dbo.QryCust_Qest.NoteID, dbo.QryCust_Qest.QeqtNum, dbo.QryCust_Qest.PartID, dbo.QryCust_Qest.[Value],"
StrSQL = StrSQL + "                      dbo.QryCust_Qest.DueDate, dbo.QryCust_Qest.Receipt, dbo.QryCust_Qest.Summition, dbo.QryCust_Qest.CustID, dbo.QryCust_Qest.Transaction_ID,"
StrSQL = StrSQL + "                      dbo.QryCust_Qest.Transaction_Date, dbo.Transactions.NoteSerial1, dbo.TblCustemers.CusName, dbo.TblCustemers.Fullcode, dbo.TblCustemers.Cus_Phone,"
StrSQL = StrSQL + "                      dbo.TblCustemers.Cus_mobile, dbo.TblCustemers.JobTel, dbo.TblCustemers.JobTelConvert, dbo.TblCustemers.HomeTel, dbo.TblCustemers.Mobile1,"
StrSQL = StrSQL + "                      dbo.TblCustemers.Mobile2 , dbo.TblCustemers.Entry, dbo.TblCustemers.CusNamee"
StrSQL = StrSQL + " FROM         dbo.QryCust_Qest LEFT OUTER JOIN"
StrSQL = StrSQL + "                      dbo.TblCustemers ON dbo.QryCust_Qest.CustID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
StrSQL = StrSQL + "                      dbo.Transactions ON dbo.QryCust_Qest.Transaction_ID = dbo.Transactions.Transaction_ID"
StrSQL = StrSQL + " WHERE     (dbo.QryCust_Qest.QestID NOT IN"
StrSQL = StrSQL + "                          (SELECT     QestID"
StrSQL = StrSQL + "                             From InstallmentDet_Junc_Receipt"
StrSQL = StrSQL + "                             WHERE     Status <> 1)) AND (dbo.QryCust_Qest.DueDate = '" & SQLDate(DateAdd(Askinterval, Askcount, Date)) & "')"
StrSQL = StrSQL + " ORDER BY dbo.QryCust_Qest.CusName, dbo.QryCust_Qest.Transaction_ID, dbo.QryCust_Qest.QeqtNum"

    End If

    RsTemp.Open StrSQL, Cn, adOpenStatic, adLockOptimistic, adCmdText

    If Not (RsTemp.EOF Or RsTemp.BOF) Then

        With FG
            .Rows = .FixedRows

            For ReCount = 1 To RsTemp.RecordCount
                .Rows = .Rows + 1
                RowNum = .Rows - 1
                .TextMatrix(RowNum, .ColIndex("Send")) = "...."
                   
                ', dbo.QryCust_Qest.CustID
              
            '    .TextMatrix(RowNum, .ColIndex("Name")) = IIf(IsNull(RsTemp("CusName").value), "", RsTemp("CusName").value)
          .TextMatrix(RowNum, .ColIndex("CustID")) = (IIf(IsNull(RsTemp("CustID").value), 0, RsTemp("CustID").value))
               ' .TextMatrix(RowNum, .ColIndex("Numbers")) = GetCustomerNumber(IIf(IsNull(RsTemp("CustID").value), 0, RsTemp("CustID").value))
            
                .TextMatrix(RowNum, .ColIndex("BillIID")) = get_transaction_NoteSerial1ByiD(IIf(IsNull(RsTemp("Transaction_ID").value), 0, RsTemp("Transaction_ID").value), "Transaction_Type=21")
                .TextMatrix(RowNum, .ColIndex("TransDate")) = IIf(IsNull(RsTemp("Transaction_Date").value), "", Format(RsTemp("Transaction_Date").value, "yyyy/mm/dd"))
                .TextMatrix(RowNum, .ColIndex("QestNum")) = IIf(IsNull(RsTemp("QeqtNum").value), "0", RsTemp("QeqtNum").value)
                .TextMatrix(RowNum, .ColIndex("DueDate")) = IIf(IsNull(RsTemp("DueDate").value), "0", Format(RsTemp("DueDate").value, "yyyy/mm/dd"))
                .TextMatrix(RowNum, .ColIndex("Value")) = IIf(IsNull(RsTemp("Value").value), "0", Format(RsTemp("Value").value, "##.00"))
                .Cell(flexcpPictureAlignment, RowNum, 0) = flexPicAlignRightCenter
            
                .TextMatrix(RowNum, .ColIndex("Payed")) = IIf(IsNull(RsTemp("Summition").value), "0", Format(RsTemp("Summition").value, "##.00"))
                .Cell(flexcpPictureAlignment, RowNum, 0) = flexPicAlignRightCenter
            
                .TextMatrix(RowNum, .ColIndex("Results")) = val(.TextMatrix(RowNum, .ColIndex("Value"))) - val(.TextMatrix(RowNum, .ColIndex("Payed")))
            
                If val(.TextMatrix(RowNum, .ColIndex("Payed"))) <> 0 Then
                    .Cell(flexcpBackColor, ReCount, 0, ReCount, 7) = vbGreen
                End If
                         .TextMatrix(RowNum, .ColIndex("Fullcode")) = IIf(IsNull(RsTemp("fullcode").value), "", RsTemp("fullcode").value)
            .TextMatrix(RowNum, .ColIndex("Cus_Phone")) = IIf(IsNull(RsTemp("Cus_Phone").value), "", RsTemp("Cus_Phone").value)
            .TextMatrix(RowNum, .ColIndex("Cus_mobile")) = IIf(IsNull(RsTemp("Cus_mobile").value), "", RsTemp("Cus_mobile").value)
            .TextMatrix(RowNum, .ColIndex("JobTel")) = IIf(IsNull(RsTemp("JobTel").value), "", RsTemp("JobTel").value)
            .TextMatrix(RowNum, .ColIndex("JobTelConvert")) = IIf(IsNull(RsTemp("JobTelConvert").value), "", RsTemp("JobTelConvert").value)
            .TextMatrix(RowNum, .ColIndex("HomeTel")) = IIf(IsNull(RsTemp("HomeTel").value), "", RsTemp("HomeTel").value)
            .TextMatrix(RowNum, .ColIndex("Mobile1")) = IIf(IsNull(RsTemp("Mobile1").value), "", RsTemp("Mobile1").value)
            .TextMatrix(RowNum, .ColIndex("Mobile2")) = IIf(IsNull(RsTemp("Mobile2").value), "", RsTemp("Mobile2").value)
            .TextMatrix(RowNum, .ColIndex("Entry")) = IIf(IsNull(RsTemp("Entry").value), "", RsTemp("Entry").value)
               If SystemOptions.UserInterface = ArabicInterface Then
            .TextMatrix(RowNum, .ColIndex("Name")) = IIf(IsNull(RsTemp("CusName").value), "", RsTemp("CusName").value)
            Else
            .TextMatrix(RowNum, .ColIndex("Name")) = IIf(IsNull(RsTemp("CusNamee").value), "", RsTemp("CusNamee").value)
            End If
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
    Me.cmdPrint.Caption = "Print"
SendMessage.Caption = "Send"
Check17.Caption = "Select All"
Frame1.Caption = "Colors"
Label2.Caption = "Payed part of Installment"
    With Me.FG
        .TextMatrix(0, .ColIndex("Name")) = "Customer Name"
        .TextMatrix(0, .ColIndex("BillIID")) = "BillI ID"
        .TextMatrix(0, .ColIndex("TransDate")) = "Trans Date"
        .TextMatrix(0, .ColIndex("QestNum")) = "installm. #"
        .TextMatrix(0, .ColIndex("DueDate")) = "DueDate"
        .TextMatrix(0, .ColIndex("value")) = "Value"
        .TextMatrix(0, .ColIndex("Fullcode")) = "Code"
        .TextMatrix(0, .ColIndex("JobTel")) = "Work Phone"
        .TextMatrix(0, .ColIndex("JobTelConvert")) = "Convert"
        .TextMatrix(0, .ColIndex("HomeTel")) = "Home Phone"
        .TextMatrix(0, .ColIndex("Mobile1")) = "Mobile"
        .TextMatrix(0, .ColIndex("Mobile2")) = "Mobile"
        .TextMatrix(0, .ColIndex("Show")) = "Show"
        .TextMatrix(0, .ColIndex("Send")) = "Send"
        .TextMatrix(0, .ColIndex("Results")) = "Remaining"
        .TextMatrix(0, .ColIndex("Payed")) = "Payed"
        .TextMatrix(0, .ColIndex("Cus_mobile")) = "Mobile"
        .TextMatrix(0, .ColIndex("Cus_Phone")) = "Phone"

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
    Dim RowNum As Integer
    Dim Opt As Integer
    Dim CurrentMessage As String
    Numbers = ""

    With FG

        For RowNum = .FixedRows To .Rows - 1
    
            If .Cell(flexcpChecked, RowNum, .ColIndex("Send")) = flexChecked Then

                '  MsgBox (.TextMatrix(RowNum, .ColIndex("Numbers")))
                If (.TextMatrix(RowNum, .ColIndex("Numbers"))) <> "" Then
                    'Numbers = Numbers & "," & (.TextMatrix(RowNum, .ColIndex("Numbers")))
                    CurrentMessage = ComposMessage(Me.Name, 0, "", MonthName(Month((.TextMatrix(RowNum, .ColIndex("DueDate"))))), Opt)
            
                    SMSSeTTings.SendMessage CurrentMessage, (.TextMatrix(RowNum, .ColIndex("Numbers")))
                    SMSSeTTings.Hide
             
                End If
            End If
          
        Next RowNum

    End With

End Sub

Private Sub SendMessage_MouseMove(Button As Integer, _
                                  Shift As Integer, _
                                  X As Single, _
                                  Y As Single)
 
    Dim Opt As Integer
    Dim CurrentMessage As String
    SendMessage.ToolTipText = ComposMessage(Me.Name, 0, "", "", Opt)
 
End Sub
