VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Begin VB.Form FrmFillBalance 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "√’‰«› «·—’Ìœ «·«›  «ÕÌ"
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4290
   HelpContextID   =   90
   Icon            =   "FrmFillBalance.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   3870
   ScaleWidth      =   4290
   Begin VB.CheckBox XPChkQuantity 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "  ”ÃÌ· ﬂ„Ì…"
      Enabled         =   0   'False
      Height          =   285
      Left            =   2700
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   2400
      Width           =   1515
   End
   Begin VB.ComboBox XPCboItemCase 
      Height          =   315
      Left            =   1860
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   885
      Width           =   1425
   End
   Begin VB.TextBox XPTxtSerial 
      Alignment       =   1  'Right Justify
      Height          =   330
      Left            =   870
      MaxLength       =   20
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   1950
      Width           =   2415
   End
   Begin VB.TextBox XPTxtPrice 
      Alignment       =   1  'Right Justify
      Height          =   330
      Left            =   870
      MaxLength       =   10
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   1575
      Width           =   2415
   End
   Begin VB.TextBox XPTxtQuantity 
      Alignment       =   1  'Right Justify
      Height          =   330
      Left            =   870
      MaxLength       =   10
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   1215
      Width           =   2415
   End
   Begin VB.TextBox TxtJeneralPart 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   330
      Left            =   60
      MaxLength       =   30
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   2790
      Width           =   2385
   End
   Begin VB.TextBox XPTxtCode 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   930
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   180
      Width           =   2355
   End
   Begin MSDataListLib.DataCombo DCboItemsName 
      Height          =   315
      Left            =   450
      TabIndex        =   1
      Top             =   540
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin ImpulseButton.ISButton CmdAdd 
      Height          =   375
      Left            =   915
      TabIndex        =   8
      Top             =   3435
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   661
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "≈÷«›…"
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
      ButtonImage     =   "FrmFillBalance.frx":038A
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
   Begin ImpulseButton.ISButton CmdExit 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   60
      TabIndex        =   9
      Top             =   3435
      Width           =   795
      _ExtentX        =   1402
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
      ButtonImage     =   "FrmFillBalance.frx":0724
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
   Begin ImpulseButton.ISButton CmdItemSearch 
      Height          =   345
      Left            =   30
      TabIndex        =   19
      Top             =   510
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   609
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "..."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonImage     =   "FrmFillBalance.frx":0ABE
      DrawFocusRectangle=   0   'False
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "Ì„ﬂ‰ﬂ ≈ŸÂ«— ‘«‘… «·»ÕÀ ⁄‰ «·√’‰«› »«·÷€ÿ ⁄·Ï F7"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   405
      Index           =   8
      Left            =   2130
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   3420
      Width           =   2085
   End
   Begin VB.Label LblSerial 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   225
      Left            =   30
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   930
      Width           =   1725
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·’‰›"
      Height          =   315
      Index           =   6
      Left            =   3330
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   540
      Width           =   915
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "ﬂÊœ «·’‰›"
      Height          =   315
      Index           =   5
      Left            =   3330
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   180
      Width           =   915
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "Õ«·… «·’‰›"
      Height          =   315
      Index           =   4
      Left            =   3330
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   885
      Width           =   915
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·ﬂ„Ì…"
      Height          =   315
      Index           =   3
      Left            =   3330
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   1215
      Width           =   915
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·”⁄—"
      Height          =   315
      Index           =   1
      Left            =   3330
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   1605
      Width           =   915
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·”Ì—Ì«·"
      Height          =   315
      Index           =   4
      Left            =   3330
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   1950
      Width           =   825
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·Ã“¡ «·À«»  „‰ «·”Ì—Ì«·"
      Enabled         =   0   'False
      Height          =   315
      Index           =   1
      Left            =   2490
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   2790
      Width           =   1725
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   5145
      X2              =   -90
      Y1              =   3285
      Y2              =   3300
   End
   Begin VB.Label LblRequireSerial 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   4140
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   1935
      Visible         =   0   'False
      Width           =   135
   End
End
Attribute VB_Name = "FrmFillBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cSearchDcbo  As clsDCboSearch

Private Sub cmdAdd_Click()
    On Error GoTo ErrTrap
    Dim Msg As String
    Dim ItemCount As Integer
    Dim StrSerial As String
    Dim VarNum As Integer

    If DCboItemsName.text = "" Then
        Msg = "ÌÃ»  ÕœÌœ «”„ «·’‰›"
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        DCboItemsName.SetFocus
        SendKeys "{F4}"
        Exit Sub
    End If

    If XPTxtQuantity.text = "" Then
        Msg = "ÌÃ»  ÕœÌœ «·ﬂ„Ì…"
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        XPTxtQuantity.SetFocus
        Exit Sub
    End If

    If Not IsNumeric(XPTxtQuantity.text) Then
        Msg = "«·ﬂ„Ì… ÌÃ» √‰  ﬂÊ‰ ﬁÌ„… —ﬁ„Ì…"
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        XPTxtQuantity.SetFocus
        Exit Sub
    End If

    If XPTxtPrice.text = "" Then
        Msg = "ÌÃ»  ÕœÌœ «·”⁄—"
        MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        XPTxtPrice.SetFocus
        Exit Sub
    End If

    If Not IsNumeric(XPTxtPrice.text) Then
        Msg = "«·”⁄— ÌÃ» √‰ ÌﬂÊ‰ ﬁÌ„… —ﬁ„Ì…"
        XPTxtPrice.SetFocus
        Exit Sub
    End If

    With FrmOpeningBalance.FG

        If XPChkQuantity.value = Checked Then
            If TxtJeneralPart.text = "" Then
                Msg = "ÌÃ»  ÕœÌœ «·Ã“¡ «·À«»  „‰ «·”Ì—Ì«· "
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                TxtJeneralPart.SetFocus
                Exit Sub
            End If

            If XPTxtSerial.text = "" Then
                Msg = "ÌÃ»  ÕœÌœ »œ«Ì… «· —ﬁÌ„  "
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                XPTxtSerial.SetFocus
                Exit Sub
            End If

            If Not IsNumeric(XPTxtSerial.text) Then
                Msg = "»œ«Ì… «· —ﬁÌ„ ÌÃ» √‰  ﬂÊ‰ ﬁÌ„… —ﬁ„Ì…  "
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                XPTxtSerial.SetFocus
                Exit Sub
            End If

            Screen.MousePointer = vbArrowHourglass

            For ItemCount = 1 To XPTxtQuantity

                If .TextMatrix(.Rows - 1, .ColIndex("Code")) <> "" Then
                    .Rows = .Rows + 1
                End If

                .TextMatrix(.Rows - 1, .ColIndex("Code")) = DCboItemsName.BoundText
                .TextMatrix(.Rows - 1, .ColIndex("Name")) = DCboItemsName.BoundText

                If XPCboItemCase.ListIndex <> -1 Then
                    .TextMatrix(.Rows - 1, .ColIndex("ItemCase")) = XPCboItemCase.ListIndex + 1
                End If

                VarNum = XPTxtSerial + ItemCount - 1
                StrSerial = TxtJeneralPart & VarNum
                .TextMatrix(.Rows - 1, .ColIndex("HaveSerial")) = True
                .TextMatrix(.Rows - 1, .ColIndex("Serial")) = StrSerial
                .TextMatrix(.Rows - 1, .ColIndex("Count")) = 1
                .TextMatrix(.Rows - 1, .ColIndex("Price")) = XPTxtPrice.text
            Next ItemCount

        Else

            If .TextMatrix(.Rows - 1, .ColIndex("Code")) <> "" Then
                .Rows = .Rows + 1
            End If

            .TextMatrix(.Rows - 1, .ColIndex("Name")) = DCboItemsName.BoundText
            .TextMatrix(.Rows - 1, .ColIndex("Code")) = DCboItemsName.BoundText

            If XPCboItemCase.ListIndex <> -1 Then
                .TextMatrix(.Rows - 1, .ColIndex("ItemCase")) = XPCboItemCase.ListIndex + 1
            End If

            .TextMatrix(.Rows - 1, .ColIndex("Serial")) = XPTxtSerial.text
            .TextMatrix(.Rows - 1, .ColIndex("Count")) = XPTxtQuantity.text
            .TextMatrix(.Rows - 1, .ColIndex("Price")) = XPTxtPrice.text

            If LblSerial.Tag = "T" Then
                .Cell(flexcpChecked, .Rows - 1, .ColIndex("HaveSerial")) = flexChecked
            ElseIf LblSerial.Tag = "F" Then
                .Cell(flexcpChecked, .Rows - 1, .ColIndex("HaveSerial")) = flexUnchecked
            End If
        End If

    End With

    Screen.MousePointer = vbDefault
    clear_all Me
    XPTxtCode.SetFocus
    XPCboItemCase.ListIndex = 0
    Exit Sub
ErrTrap:
    Screen.MousePointer = vbDefault
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CmdItemSearch_Click()
    Load FrmItemSearch
    FrmItemSearch.RetrunType = 1
    Set FrmItemSearch.DcboItems = Me.DCboItemsName
    FrmItemSearch.Show vbModal
    XPTxtCode.SetFocus
End Sub

Private Sub DCboItemsName_Change()
    Dim StrSQL As String
    Dim RsTemp As New ADODB.Recordset

    On Error GoTo ErrTrap

    If DCboItemsName.BoundText <> "" Then
        StrSQL = "select * From TblItems where ItemID=" & DCboItemsName.BoundText
        RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
        XPTxtSerial.Enabled = RsTemp("HaveSerial").value
        LblRequireSerial.Visible = RsTemp("HaveSerial").value
        XPTxtCode.text = RsTemp("ItemCode").value

        If XPTxtSerial.Enabled = False Then
            XPTxtSerial.text = ""
        End If

        XPLbl(4).Caption = "«·”Ì—Ì«·"

        If RsTemp("HaveSerial").value = True Then
            XPTxtQuantity.Enabled = False
            XPTxtQuantity.text = "1"
            XPChkQuantity.Enabled = True
            XPChkQuantity.value = False

            LblSerial.ForeColor = vbRed
            LblSerial.Caption = "·Â ”Ì—Ì«·"
            LblSerial.Tag = "T"
        Else
            XPTxtQuantity.Enabled = True
            XPTxtQuantity.text = ""
            XPChkQuantity.Enabled = False
            XPChkQuantity.value = Unchecked
            LblSerial.ForeColor = vbBlue
            LblSerial.Caption = "·Ì” ·Â ”Ì—Ì«·"
            LblSerial.Tag = "F"
        End If

        Me.XPTxtPrice.text = IIf(IsNull(RsTemp("PurchasePrice").value), 0, RsTemp("PurchasePrice").value)
        RsTemp.Close
        XPChkQuantity_Click
    Else
        XPChkQuantity.Enabled = False
        XPChkQuantity.value = Unchecked
        XPLbl(1).Enabled = False
        TxtJeneralPart.Enabled = False
        XPChkQuantity_Click
        LblSerial.Caption = ""
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If Shift = 2 Then
        If KeyCode = vbKeyX Then
            CmdExit_Click
        End If
    End If

    If KeyCode = vbKeyReturn Then
        SendKeys "{TAB}"
    ElseIf KeyCode = vbKeyF3 Then
        CmdItemSearch_Click
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim StrSQL As String
    Dim rs As ADODB.Recordset
    Dim Dcombos As ClsDataCombos

    CenterForm Me

    FormPostion Me, GetPostion
    Set Dcombos = New ClsDataCombos
    Dcombos.GetItemsNames Me.DCboItemsName
    Set cSearchDcbo = New clsDCboSearch
    Set cSearchDcbo.Client = Me.DCboItemsName

    With XPCboItemCase
        .AddItem "ÃœÌœ"
        .AddItem "„” ⁄„·"
    End With

    XPCboItemCase.ListIndex = 0
    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Unload(Cancel As Integer)

    FormPostion Me, SavePostion
    Set cSearchDcbo = Nothing
End Sub

Private Sub XPChkQuantity_Click()
    On Error GoTo ErrTrap

    If XPChkQuantity.value = Checked Then
        XPLbl(4).Caption = "»œ«Ì… «· —ﬁÌ„"
        XPTxtQuantity.Enabled = True
        XPTxtQuantity.text = ""
        XPLbl(1).Enabled = True
        TxtJeneralPart.Enabled = True
        TxtJeneralPart.Enabled = True
        XPLbl(1).Enabled = True
        TxtJeneralPart.text = ""
    Else
        XPLbl(4).Caption = "«·”Ì—Ì«·"
        '    XPTxtQuantity.Enabled = False
        '    XPTxtQuantity.Text = ""
        XPLbl(1).Enabled = False
        TxtJeneralPart.Enabled = False
        TxtJeneralPart.Enabled = False
        XPLbl(1).Enabled = False
        TxtJeneralPart.text = ""
    End If

    Exit Sub
ErrTrap:

End Sub

Private Sub XPTxtCode_Change()
    On Error GoTo ErrTrap
    Dim StrSQL As String
    Dim RsTemp As New ADODB.Recordset

    If XPTxtCode.text <> "" Then
        StrSQL = "select * From TblItems where ItemCode='" & XPTxtCode.text & "'"
        RsTemp.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

        If Not (RsTemp.EOF Or RsTemp.BOF) Then
            DCboItemsName.BoundText = RsTemp("ItemID").value
            'Me.XPTxtPrice.Text = RsTemp("ItemID").Value
        End If

        RsTemp.Close
        Set RsTemp = Nothing
    End If

    Exit Sub
ErrTrap:
End Sub
