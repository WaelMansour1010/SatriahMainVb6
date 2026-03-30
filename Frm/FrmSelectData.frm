VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmSelectData 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "⁄—÷  ﬁ«—Ì— «·√’‰«›"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5430
   Icon            =   "FrmSelectData.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   4935
   ScaleWidth      =   5430
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox Chk 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "€·ﬁ Â–Â «·‘«‘… »⁄œ ⁄—÷ «· ﬁ—Ì—"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   600
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   4050
      Width           =   4725
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   "ŒÌ«—«  Œ«’… »«· ﬁ—Ì—"
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
      Height          =   2025
      Index           =   3
      Left            =   30
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   1950
      Width           =   5355
      Begin VB.TextBox TxtCusID 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3480
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   570
         Width           =   525
      End
      Begin VB.TextBox TxtStoreID 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3480
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   240
         Width           =   525
      End
      Begin MSDataListLib.DataCombo DcboCusSupp 
         Height          =   315
         Left            =   150
         TabIndex        =   5
         Top             =   570
         Width           =   3315
         _ExtentX        =   5847
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Frame Fra 
         BackColor       =   &H00E2E9E9&
         Caption         =   "›Ì «·› —…"
         ForeColor       =   &H00FF0000&
         Height          =   1035
         Index           =   1
         Left            =   60
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   930
         Width           =   2415
         Begin MSComCtl2.DTPicker DTPFrom 
            Height          =   345
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   1665
            _ExtentX        =   2937
            _ExtentY        =   609
            _Version        =   393216
            CheckBox        =   -1  'True
            CustomFormat    =   "dd/m/yyyy"
            DateIsNull      =   -1  'True
            Format          =   140967937
            CurrentDate     =   36494
         End
         Begin MSComCtl2.DTPicker DTPTo 
            Height          =   345
            Left            =   120
            TabIndex        =   8
            Top             =   600
            Width           =   1665
            _ExtentX        =   2937
            _ExtentY        =   609
            _Version        =   393216
            CheckBox        =   -1  'True
            DateIsNull      =   -1  'True
            Format          =   140967937
            CurrentDate     =   38784
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "„‰"
            Height          =   285
            Index           =   1
            Left            =   1830
            RightToLeft     =   -1  'True
            TabIndex        =   20
            Top             =   255
            Width           =   495
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E2E9E9&
            Caption         =   "≈·Ï"
            Height          =   285
            Index           =   0
            Left            =   1800
            RightToLeft     =   -1  'True
            TabIndex        =   19
            Top             =   645
            Width           =   555
         End
      End
      Begin VB.CheckBox ChkShowTable 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "⁄—÷ ÃœÊ·Ï"
         Height          =   255
         Left            =   3060
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   1050
         Width           =   2205
      End
      Begin MSDataListLib.DataCombo DcboStores 
         Height          =   315
         Left            =   150
         TabIndex        =   4
         Top             =   240
         Width           =   3315
         _ExtentX        =   5847
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«”„ «·⁄„Ì· «Ê «·„Ê—œ"
         Height          =   405
         Index           =   7
         Left            =   4110
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   540
         Width           =   1185
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«”„ «·„Œ“‰"
         Height          =   315
         Index           =   2
         Left            =   4110
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   270
         Width           =   1185
      End
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   "≈Œ — ‰Ê⁄ «· ﬁ—Ì— «·„ÿ·Ê»"
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
      Height          =   765
      Index           =   2
      Left            =   30
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   1170
      Width           =   5355
      Begin VB.ComboBox CboReportType 
         Height          =   315
         Left            =   510
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   270
         Width           =   3795
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«”„ «· ﬁ—Ì—"
         Height          =   315
         Index           =   6
         Left            =   4290
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   300
         Width           =   1005
      End
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   "≈Œ Ì«— «·’‰›"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1125
      Index           =   0
      Left            =   30
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   30
      Width           =   5355
      Begin ImpulseButton.ISButton Cmd 
         Height          =   315
         Left            =   30
         TabIndex        =   25
         Top             =   690
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   556
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   "..."
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
         DrawFocusRectangle=   0   'False
      End
      Begin VB.TextBox TxtItemCode 
         Height          =   345
         Left            =   2640
         TabIndex        =   1
         Top             =   270
         Width           =   1785
      End
      Begin MSDataListLib.DataCombo DCboItemName 
         Height          =   315
         Left            =   690
         TabIndex        =   2
         Top             =   690
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«”„ «·’‰›"
         Height          =   315
         Index           =   5
         Left            =   4320
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   720
         Width           =   945
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ﬂÊœ «·’‰›"
         Height          =   315
         Index           =   3
         Left            =   4320
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   330
         Width           =   945
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«ﬂ » ﬂÊœ «·’‰› À„ ≈÷€ÿ ≈‰ —"
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
         Height          =   285
         Index           =   4
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   300
         Width           =   2505
      End
   End
   Begin ImpulseButton.ISButton CmdPreview 
      Height          =   375
      Left            =   930
      TabIndex        =   10
      Top             =   4470
      Width           =   1275
      _ExtentX        =   2249
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
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      LowerToggledContent=   0   'False
      ColorTextShadow =   -2147483637
   End
   Begin ImpulseButton.ISButton CmdCancel 
      Cancel          =   -1  'True
      CausesValidation=   0   'False
      Height          =   375
      Left            =   90
      TabIndex        =   11
      Top             =   4470
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   661
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "≈·€«¡"
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
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   5400
      X2              =   -60
      Y1              =   4020
      Y2              =   4020
   End
End
Attribute VB_Name = "FrmSelectData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TTP As clstooltipdemand
Dim cDcboSearch(2) As clsDCboSearch

Private Sub CboReportType_Change()

    DisableAll

    If Me.CboReportType.ListIndex = 0 Then
        'ﬂ«—  «·’‰›
        Me.DcboStores.Enabled = True
        ChkShowTable.Enabled = True
        DateFramStatus True
    ElseIf (Me.CboReportType.ListIndex = 1 Or Me.CboReportType.ListIndex = 4 Or Me.CboReportType.ListIndex = 7 Or Me.CboReportType.ListIndex = 10) Then
        DisableAll
    ElseIf (Me.CboReportType.ListIndex = 2 Or Me.CboReportType.ListIndex = 3 Or Me.CboReportType.ListIndex = 5 Or Me.CboReportType.ListIndex = 6) Then
        '„»Ì⁄«  «·’‰›
        ' ﬁ—Ì— „— Ã⁄ „»Ì⁄«  «·’‰›
        ' ﬁ—Ì—  „‘ —Ì«  «·’‰›
        ' ﬁ—Ì— „— Ã⁄ „‘ —Ì«  «·’‰›
        Me.DcboStores.Enabled = True
        Me.DcboCusSupp.Enabled = True
        DateFramStatus True
    Else
        DisableAll
    End If

End Sub

Private Sub CboReportType_Click()
    CboReportType_Change
End Sub

Private Sub CboReportType_Validate(Cancel As Boolean)
    Dim StrMSG As String

    If CboReportType.text Like "----*" Then
        Set TTP = New clstooltipdemand
        Set TTP.m_From = Me
        TTP.Style = TTBalloon
        TTP.Icon = TTIconError
        TTP.Centered = True
        TTP.RightToLeft = True
        TTP.CreateToolTip CboReportType.hWnd
        TTP.DelayTime = 250
        TTP.VisibleTime = 5000
        StrMSG = "Œÿ« ›Ï ≈Œ Ì«— «· ﬁ—Ì—...!!!"
        TTP.title = StrMSG
        StrMSG = "ÌÃ» «‰  ﬁÊ„ »≈Œ ÌÌ«— «· ﬁ—Ì— «·„—«œ ⁄—÷Â"
        TTP.TipText = StrMSG
        TTP.PopupOnDemand = True
        TTP.show (CboReportType.Width / Screen.TwipsPerPixelY), (CboReportType.Height / Screen.TwipsPerPixelX - 1)     '//In Pixel only
        Cancel = True
    Else

        If Not TTP Is Nothing Then
            TTP.Destroy
        End If
    End If

End Sub

Private Sub Cmd_Click()
    Load FrmItemSearch
    FrmItemSearch.RetrunType = 1
    Set FrmItemSearch.DcboItems = Me.DCboItemName
    FrmItemSearch.show vbModal
End Sub

Private Sub DcboItemName_Change()

    If val(Me.DCboItemName.BoundText) <> 0 Then
        Me.TxtItemCode.text = GetItemCode(Me.DCboItemName.BoundText)
    End If

End Sub

Private Sub Form_Activate()
    'PutFormOnTop Me.hwnd
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Sendkeys "{TAB}"
    End If

End Sub

Private Sub Form_Load()
    Dim Dcombos As ClsDataCombos

    Set Me.cmdCancel.ButtonImage = mdifrmmain.ImgLstMenuIcons.ListImages("Hide").Picture
    Set Me.CmdPreview.ButtonImage = mdifrmmain.ImgLstMenuIcons.ListImages("Preview").Picture
    Set Me.Cmd.ButtonImage = mdifrmmain.ImgLstMenuIcons.ListImages("BrowseFile").Picture

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    Set Dcombos = New ClsDataCombos
    Dcombos.GetItemsNames Me.DCboItemName
    Set cDcboSearch(0) = New clsDCboSearch
    Set cDcboSearch(0).Client = Me.DCboItemName
    Dcombos.GetStores Me.DcboStores
    Set cDcboSearch(1) = New clsDCboSearch
    Set cDcboSearch(1).Client = Me.DcboStores
    cDcboSearch(1).SetBuddyText Me.TxtStoreID
    Dcombos.GetCustomersSuppliers 0, Me.DcboCusSupp, True
    Set cDcboSearch(2) = New clsDCboSearch
    Set cDcboSearch(2).Client = Me.DcboCusSupp
    cDcboSearch(2).SetBuddyText Me.TxtCusID

    With Me.CboReportType
        .Clear

        If SystemOptions.UserInterface = ArabicInterface Then
            .AddItem " ﬁ—Ì— ﬂ«—  «·’‰›"
            .ItemData(0) = 0
            .AddItem "-----------------------"
            .AddItem " ﬁ—Ì— „»Ì⁄«  «·’‰›"
            .ItemData(2) = 2
            .AddItem " ﬁ—Ì— „— Ã⁄ „»Ì⁄«  «·’‰›"
            .ItemData(3) = 3
            .AddItem "-----------------------"
            .AddItem " ﬁ—Ì— „‘ —Ì«  «·’‰›"
            .ItemData(5) = 5
            .AddItem " ﬁ—Ì— „— Ã⁄ „‘ —Ì«  «·’‰›"
            .ItemData(6) = 6
            .AddItem "-----------------------"
            .AddItem " ﬁ—Ì— »ﬂ„Ì«  «·—’Ìœ «·√›  «ÕÏ"
            .ItemData(8) = 8
            .AddItem " ﬁ—Ì— »«·ﬂ„Ì«  «· «·›… „‰ «·’‰›"
            .ItemData(9) = 9
            .AddItem "-----------------------"
            .AddItem "ÿ»«⁄… ﬁ«∆„… √”⁄«— «·’‰›"
            .ItemData(11) = 11
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            .AddItem "Item Cart Report"
            .ItemData(0) = 0
            .AddItem "-----------------------"
            .AddItem "Item Sales Report"
            .ItemData(2) = 2
            .AddItem "Item Return Sales Report"
            .ItemData(3) = 3
            .AddItem "-----------------------"
            .AddItem "Item Purchases Report"
            .ItemData(5) = 5
            .AddItem "Item Return Purchases Report"
            .ItemData(6) = 6
            .AddItem "-----------------------"
            .AddItem "Item Openning Balance Quantity Report"
            .ItemData(8) = 8
            .AddItem "Item Waste Quantity Report"
            .ItemData(9) = 9
            .AddItem "-----------------------"
            .AddItem "Item Price List Report"
            .ItemData(11) = 11
        End If

    End With

    CenterForm Me

    FormPostion Me, GetPostion
    SetDtpickerDate Me.DTPFrom
    SetDtpickerDate Me.DTPTo
    Me.ZOrder 0
    'PutFormOnTop Me.hwnd
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer

    FormPostion Me, SavePostion
    Set TTP = Nothing

    For i = LBound(cDcboSearch) To UBound(cDcboSearch)
        Set cDcboSearch(i) = Nothing
    Next i

End Sub

Private Sub CmdPreview_Click()
    Dim Msg As String
    Dim cItemReport As ClsItemsReport
    Dim BolShowTable As Boolean
    Dim BolRetrun As Boolean

    Dim LngItemID As Long
    Dim LngCusID As Long
    Dim LngStoreID As Long

    If Me.DCboItemName.BoundText = "" Then
        Msg = "ÌÃ» «Œ Ì«— «”„ «·’‰› «· Ì  —€»" & CHR(13)
        Msg = Msg + "›Ì ⁄—÷ «· ﬁ—Ì— ·Â ...!!"
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        DCboItemName.SetFocus
        Sendkeys "{F4}"
        Exit Sub
    End If

'    If Me.CboReportType.ListIndex = 0 Then
'        If Me.DcboStores.BoundText = "" Then
'            Msg = "ÌÃ» «Œ Ì«— «”„ «·„Œ“‰...!!" & CHR(13)
'            MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
'            DcboStores.SetFocus
'            Sendkeys "{F4}"
'            Exit Sub
'        End If
'    End If

    LngItemID = val(Me.DCboItemName.BoundText)
    LngCusID = val(Me.DcboCusSupp.BoundText)
    LngStoreID = val(Me.DcboStores.BoundText)

    If Me.CboReportType.ListIndex = 0 Then
        Set cItemReport = New ClsItemsReport
        BolShowTable = IIf(Me.ChkShowTable.value = vbChecked, True, False)
        If val(Me.DcboStores.BoundText) = 0 Then
            BolRetrun = cItemReport.ItemCart(Me.DCboItemName.BoundText, 0, DTPFrom.value, DTPTo.value, BolShowTable)
        Else
            BolRetrun = cItemReport.ItemCart(Me.DCboItemName.BoundText, val(Me.DcboStores.BoundText), DTPFrom.value, DTPTo.value, BolShowTable)
        End If

        If BolRetrun = True Then
            If Me.Chk.value = vbChecked Then
                Me.Hide
                Unload Me
            End If
        End If

    ElseIf Me.CboReportType.ListIndex = 1 Then
        '------
        'Sep
    ElseIf Me.CboReportType.ListIndex = 2 Then
        '„»Ì⁄«  ’‰›
        Set cItemReport = New ClsItemsReport
        BolRetrun = cItemReport.ShowItemTransCustomer(LngItemID, LngCusID, 2, LngStoreID, Me.DTPFrom.value, Me.DTPTo.value, WindowTarget)
    ElseIf Me.CboReportType.ListIndex = 3 Then
        '„— Ã⁄ „»Ì⁄«  «·’‰›
        Set cItemReport = New ClsItemsReport
        BolRetrun = cItemReport.ShowItemTransCustomer(LngItemID, LngCusID, 9, LngStoreID, Me.DTPFrom.value, Me.DTPTo.value, WindowTarget)
    ElseIf Me.CboReportType.ListIndex = 5 Then
        '„‘ —Ì«  «·’‰›
        Set cItemReport = New ClsItemsReport
        BolRetrun = cItemReport.ShowItemTransCustomer(LngItemID, LngCusID, 1, LngStoreID, Me.DTPFrom.value, Me.DTPTo.value, WindowTarget)
    ElseIf Me.CboReportType.ListIndex = 6 Then
        ' „— Ã⁄ „‘ —Ì«  «·’‰›
        Set cItemReport = New ClsItemsReport
        BolRetrun = cItemReport.ShowItemTransCustomer(LngItemID, LngCusID, 5, LngStoreID, Me.DTPFrom.value, Me.DTPTo.value, WindowTarget)
    End If
    
    If BolRetrun = True Then
        If Me.Chk.value = vbChecked Then
            Unload Me
        End If
    End If

End Sub

Private Sub TxtItemCode_KeyDown(KeyCode As Integer, _
                                Shift As Integer)
    Dim LngTempID As Long

    If KeyCode = vbKeyReturn Then
        If Trim(Me.TxtItemCode.text) = "" Then Exit Sub
        LngTempID = GetItemID(Trim(Me.TxtItemCode.text))

        If LngTempID = 0 Then
            Me.DCboItemName.BoundText = ""
            Msg = "·«ÌÊÃœ ’‰› „”Ã· »Â–« «·ﬂÊœ..!"
            MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        ElseIf val(Me.DCboItemName.BoundText) <> LngTempID Then
            DCboItemName.BoundText = LngTempID
        End If
    End If

End Sub

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub ChangeLang()
    Me.Caption = "Show Item Card Report"
    Me.Fra(0).Caption = "Choose an Item"
    Fra(1).Caption = "Choose Date Interval"
    Me.lbl(0).Caption = "To"
    Me.lbl(1).Caption = "From"
    Me.lbl(2).Caption = "Store Name"
    Me.lbl(3).Caption = "Item Code"
    Me.lbl(5).Caption = "Item Name"
    Me.lbl(4).Caption = "Enter Item Code then press Enter"
    Fra(2).Caption = "Choose Report"
    Me.lbl(6).Caption = "Report Name"
    Fra(3).Caption = "More Report Options"
    Me.lbl(7).Caption = "Customer OR Supplier Name"
    ChkShowTable.Caption = "Table Show"
    Me.CmdPreview.Caption = "Show Report"
    Me.Chk.Caption = "Close this window after Report show"
    Me.cmdCancel.Caption = "Cancel"
End Sub

Private Sub DisableAll()
    Me.DcboCusSupp.BoundText = ""
    Me.DcboStores.BoundText = ""
    '---------------------------
    Me.DcboCusSupp.Enabled = False
    Me.DcboStores.Enabled = False
    ChkShowTable.Enabled = False
    '----------------------------
    'Fra(3).Enabled = False
    DateFramStatus False
End Sub

Private Sub DateFramStatus(BolStatus As Boolean)
    Fra(1).Enabled = BolStatus
    lbl(0).Enabled = BolStatus
    lbl(1).Enabled = BolStatus
    DTPFrom.Enabled = BolStatus
    DTPTo.Enabled = BolStatus
End Sub
