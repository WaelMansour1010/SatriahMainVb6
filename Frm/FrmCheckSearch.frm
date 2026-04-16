VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmCheckSearch 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "«·»ÕÀ ⁄‰ «·‘Ìﬂ«  «·€»— „Õ’·… «Ê «·€Ì— „”œœ…"
   ClientHeight    =   4875
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5355
   DrawWidth       =   10
   Icon            =   "FrmCheckSearch.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   4875
   ScaleWidth      =   5355
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.ComboBox CboCheckType 
      Height          =   315
      Left            =   2040
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   24
      Top             =   2850
      Width           =   2445
   End
   Begin VB.TextBox TxtValue 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   3600
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   3210
      Width           =   885
   End
   Begin VB.TextBox TxtSerial 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   3300
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   2490
      Width           =   1185
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E2E9E9&
      Caption         =   "›Ì «·› —…"
      ForeColor       =   &H00FF0000&
      Height          =   1065
      Index           =   0
      Left            =   60
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   2880
      Width           =   1935
      Begin MSComCtl2.DTPicker DTPFrom 
         Height          =   345
         Left            =   60
         TabIndex        =   5
         Top             =   240
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   609
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/m/yyyy"
         DateIsNull      =   -1  'True
         Format          =   100073473
         CurrentDate     =   38979.743287037
      End
      Begin MSComCtl2.DTPicker DTPTo 
         Height          =   345
         Left            =   60
         TabIndex        =   6
         Top             =   600
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   609
         _Version        =   393216
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   100073473
         CurrentDate     =   38784
      End
      Begin VB.Label Lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï"
         Height          =   285
         Index           =   10
         Left            =   1500
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   615
         Width           =   345
      End
      Begin VB.Label Lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   285
         Index           =   11
         Left            =   1560
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   255
         Width           =   285
      End
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   2040
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   3210
      Width           =   1545
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   3
         ToolTipText     =   "«ﬂ»— „‰"
         Top             =   0
         Width           =   465
      End
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "="
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   480
         RightToLeft     =   -1  'True
         TabIndex        =   2
         ToolTipText     =   "Ì”«ÊÏ"
         Top             =   0
         Value           =   -1  'True
         Width           =   495
      End
      Begin VB.OptionButton Opt 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   960
         RightToLeft     =   -1  'True
         TabIndex        =   1
         ToolTipText     =   "«’€— „‰"
         Top             =   0
         Width           =   555
      End
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   0
      Left            =   1740
      TabIndex        =   11
      Top             =   4440
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   661
      ButtonPositionImage=   1
      Caption         =   "»ÕÀ"
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
      ColorShadow     =   -2147483637
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      DisabledImageExtraction=   0
      ColorToggledHoverText=   16711680
      LowerToggledContent=   0   'False
      ColorTextShadow =   -2147483637
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   1
      Left            =   855
      TabIndex        =   12
      Top             =   4440
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      ButtonPositionImage=   1
      Caption         =   "„”Õ"
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
      ColorShadow     =   -2147483637
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      LowerToggledContent=   0   'False
      ColorTextShadow =   -2147483637
   End
   Begin ImpulseButton.ISButton Cmd 
      Cancel          =   -1  'True
      Height          =   375
      Index           =   2
      Left            =   30
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   4440
      Width           =   795
      _ExtentX        =   1402
      _ExtentY        =   661
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
   Begin VSFlex8UCtl.VSFlexGrid Fg 
      Height          =   2445
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   5355
      _cx             =   9446
      _cy             =   4313
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
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmCheckSearch.frx":000C
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
   End
   Begin MSDataListLib.DataCombo DcboUsers 
      Height          =   315
      Left            =   2040
      TabIndex        =   15
      Top             =   3990
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DcboBanks 
      Height          =   315
      Left            =   30
      TabIndex        =   16
      Top             =   2490
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo DcboCustomers 
      Height          =   315
      Left            =   2040
      TabIndex        =   17
      Top             =   3630
      Width           =   2445
      _ExtentX        =   4313
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "‰Ê⁄ «·‘Ìﬂ"
      Height          =   315
      Index           =   0
      Left            =   4530
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   2850
      Width           =   765
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·„” Œœ„"
      Height          =   315
      Index           =   2
      Left            =   4320
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   3990
      Width           =   975
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·„»·€"
      Height          =   345
      Index           =   1
      Left            =   4530
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   3240
      Width           =   765
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·‘Ìﬂ"
      Height          =   315
      Index           =   3
      Left            =   4530
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   2490
      Width           =   765
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·»‰ﬂ"
      Height          =   315
      Index           =   4
      Left            =   2310
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   2490
      Width           =   945
   End
   Begin VB.Label XPLbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·„ ⁄«„·"
      Height          =   315
      Index           =   5
      Left            =   4530
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   3630
      Width           =   765
   End
End
Attribute VB_Name = "FrmCheckSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim m_SearchType As Integer
Dim cSearchDcbo(4) As clsDCboSearch

Private Sub Cmd_Click(Index As Integer)
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer

    On Error GoTo ErrTrap

    Select Case Index

        Case 0
            Set rs = New ADODB.Recordset
            rs.Open Build_Sql, Cn, adOpenStatic, adLockOptimistic, adCmdText

            If rs.RecordCount < 1 Then
                FG.Clear flexClearScrollable, flexClearEverything
                FG.Rows = 2
                Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Exit Sub
            Else

                With Me.FG
                    .Clear flexClearScrollable, flexClearEverything
                    .Rows = .FixedRows
                    .Rows = .FixedRows + rs.RecordCount
                    rs.MoveFirst

                    For i = .FixedRows To rs.RecordCount
                        .TextMatrix(i, .ColIndex("Serial")) = i
                        .TextMatrix(i, .ColIndex("NoteID")) = IIf(IsNull(rs("NoteID").value), "", rs("NoteID").value)
                        .TextMatrix(i, .ColIndex("ChqueNum")) = IIf(IsNull(rs("ChqueNum").value), "", rs("ChqueNum").value)
                        .TextMatrix(i, .ColIndex("BankName")) = IIf(IsNull(rs("BankName").value), "", rs("BankName").value)
                        .TextMatrix(i, .ColIndex("NoteValue")) = IIf(IsNull(rs("Note_Value").value), "", rs("Note_Value").value)
                        .TextMatrix(i, .ColIndex("NoteDate")) = IIf(IsNull(rs("NoteDate").value), "", rs("NoteDate").value)
                        .TextMatrix(i, .ColIndex("DueDate")) = IIf(IsNull(rs("DueDate").value), "", rs("DueDate").value)
                        .TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(rs("CusName").value), "", rs("CusName").value)

                        If IsNull(rs("TransactionTypeName").value) Then
                            .TextMatrix(i, .ColIndex("Notes")) = "Õ—ﬂ… ’Ì«‰… —ﬁ„ " & IIf(IsNull(rs("MaintananceID").value), "", rs("MaintananceID").value)
                        Else
                            .TextMatrix(i, .ColIndex("Notes")) = IIf(IsNull(rs("TransactionTypeName").value), "", rs("TransactionTypeName").value) & " —ﬁ„ " & IIf(IsNull(rs("Transaction_Serial").value), "", rs("Transaction_Serial").value)
                        End If

                        '.TextMatrix(I, .ColIndex("UserName")) = IIf(IsNull(Rs("UserName").Value), "", Rs("UserName").Value)
                    
                        rs.MoveNext
                    Next i

                    .AutoSize 0, .Cols - 1, False
                End With

            End If

        Case 1
            clear_all Me
            FG.Clear flexClearScrollable, flexClearEverything
            FG.Rows = 1

        Case 2
            Unload Me
    End Select

    Exit Sub
ErrTrap:

    If Err.Number = -2147217900 Then
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & Chr(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… „⁄«ÌÌ— «·»ÕÀ Ê√⁄œ «·„Õ«Ê·…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

End Sub

Private Sub Fg_Click()

    With Me.FG

        If .Row = -1 Then Exit Sub
        If .Col = -1 Then Exit Sub
        If val(.TextMatrix(.Row, .ColIndex("NoteID"))) = 0 Then
            Exit Sub
        End If

        mdifrmmain.ActiveForm.TXTNoteID.text = val(.TextMatrix(.Row, .ColIndex("NoteID")))
    End With

End Sub

Private Sub Form_Load()
    Dim Dcombos As ClsDataCombos
    Dim GrdBack As New ClsBackGroundPic

    Set Dcombos = New ClsDataCombos
    CenterForm Me

    FormPostion Me, GetPostion
    Dcombos.GetCustomersSuppliers 0, Me.DcboCustomers, False
    Dcombos.GetBanks Dcbobanks
    Dcombos.GetUsers Me.DcboUsers

    Set cSearchDcbo(0) = New clsDCboSearch
    Set cSearchDcbo(0).Client = Me.DcboCustomers
    Set cSearchDcbo(1) = New clsDCboSearch
    Set cSearchDcbo(1).Client = Me.Dcbobanks

    Set cSearchDcbo(2) = New clsDCboSearch
    Set cSearchDcbo(2).Client = Me.DcboUsers

    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Clear").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture

    With Me.CboCheckType
        .Clear
        .AddItem "‘Ìﬂ ··‘—ﬂ…"
        .AddItem "‘Ìﬂ ⁄·Ï «·‘—ﬂ…"
        .AddItem "«·ﬂ·"
    End With

    With Me.FG
        Set .WallPaper = GrdBack.SearchWallpaper
        .AutoSize 0, .Cols - 1, False
    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)

    FormPostion Me, SavePostion
End Sub

Private Sub TxtValue_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, TxtValue.text, 0)
End Sub

Public Property Get SearchType() As Integer
    SearchType = m_SearchType
End Property

Public Property Let SearchType(ByVal vNewValue As Integer)
    m_SearchType = vNewValue

    With Me.FG

        If m_SearchType = 3 Then
            ' 3 «·»ÕÀ ⁄‰ «·„’—Ê›« 
            .ColHidden(.ColIndex("PaymentType")) = False
            .ColHidden(.ColIndex("CustName")) = True
            Me.Caption = "«·»ÕÀ ⁄‰ «·„’—Ê›« "
            Me.XPLbl(4).Visible = True

            Me.XPLbl(5).Visible = False
            Me.DcboCustomers.Visible = False
        ElseIf m_SearchType = 4 Then
            ' 4 «·»ÕÀ ⁄‰ «·„ﬁ»Ê÷« 
            .ColHidden(.ColIndex("PaymentType")) = True
            .ColHidden(.ColIndex("CustName")) = False
            Me.Caption = "«·»ÕÀ ⁄‰ «·„ﬁ»Ê÷« "
            Me.XPLbl(4).Visible = False

            Me.XPLbl(5).Visible = True
            Me.DcboCustomers.Visible = True

        ElseIf m_SearchType = 5 Then
            '5     «·»ÕÀ «·„œ›Ê⁄« 
            Me.Caption = "«·»ÕÀ ⁄‰ «·„œ›Ê⁄« "
            .ColHidden(.ColIndex("PaymentType")) = True
            .ColHidden(.ColIndex("CustName")) = False
            Me.XPLbl(4).Visible = False

            Me.XPLbl(5).Visible = True
            Me.DcboCustomers.Visible = True
        End If

    End With

End Property

Private Function Build_Sql() As String
    Dim StrSQL As String
    Dim StrWhere As String

    StrSQL = "SELECT Notes.NoteID, Notes.NoteDate, Notes.NoteType, Notes.NoteSerial," & "Notes.Note_Value, BanksData.BankName, Notes.ChqueNum, Notes.DueDate," & "Transactions.Transaction_Serial, Transactions.Transaction_Date," & "TblNotesTypes.NotesTypeName, TransactionTypes.TransactionTypeName," & "TblCustemers.CusName, TblMaintenece.MaintananceID, Notes.BankID "
    StrSQL = StrSQL + " FROM TransactionTypes RIGHT JOIN (Transactions RIGHT JOIN " & "(TblNotesTypes INNER JOIN (TblMaintenece RIGHT JOIN (TblCustemers RIGHT JOIN " & "(BanksData RIGHT JOIN Notes ON BanksData.BankID = Notes.BankID) " & "ON TblCustemers.CusID = Notes.CusID) ON TblMaintenece.MaintananceID = " & "Notes.MaintananceID) ON TblNotesTypes.NotesType = Notes.NoteType) ON " & "Transactions.Transaction_ID = Notes.Transaction_ID) " & "ON TransactionTypes.Transaction_Type = Transactions.Transaction_Type"

    If Me.CboCheckType.ListIndex = -1 Or Me.CboCheckType.ListIndex = 2 Then
        StrSQL = StrSQL + " Where (Notes.NoteType = 2 Or Notes.NoteType = 13)"
    ElseIf Me.CboCheckType.ListIndex = 0 Then
        StrSQL = StrSQL + " Where (Notes.NoteType = 2)"
    ElseIf Me.CboCheckType.ListIndex = 1 Then
        StrSQL = StrSQL + " Where (Notes.NoteType = 13)"
    End If

    StrSQL = StrSQL + " And NoteID NOT IN(Select NoteID From TblCheckRelease) "

    If Trim(Me.TxtSerial.text) <> "" Then
        StrWhere = StrWhere + " AND NoteSerial Like '" & Trim(Me.TxtSerial.text) & "'"
    End If

    If val(Me.TxtValue.text) > 0 Then
        If Me.Opt(1).value = True Then
            StrWhere = StrWhere + " AND Notes.Note_Value =" & val(Me.TxtValue.text) & ""
        ElseIf Me.Opt(0).value = True Then
            StrWhere = StrWhere + " AND Notes.Note_Value >" & val(Me.TxtValue.text) & ""
        Else
            StrWhere = StrWhere + " AND Notes.Note_Value <" & val(Me.TxtValue.text) & ""
        End If
    End If

    If Me.DcboUsers.BoundText <> "" Then
        StrWhere = StrWhere + " AND Notes.UserID=" & Me.DcboUsers.BoundText & ""
    End If

    If Me.DcboCustomers.BoundText <> "" Then
        StrWhere = StrWhere + " AND Notes.CusID=" & Me.DcboCustomers.BoundText & ""
    End If

    If Me.Dcbobanks.BoundText <> "" Then
        StrWhere = StrWhere + " AND  Notes.BankID=" & Me.Dcbobanks.BoundText & ""
    End If

    If Not IsNull(Me.DTPFrom.value) Then
        StrWhere = StrWhere + " AND  Notes.NoteDate >=#" & SQLDate(Me.DTPFrom.value) & "#"
    End If

    If Not IsNull(Me.DTPTo.value) Then
        StrWhere = StrWhere + " AND  Notes.NoteDate <=#" & SQLDate(Me.DTPTo.value) & "#"
    End If

    StrSQL = StrSQL + StrWhere + " Order By Notes.NoteID"
    Build_Sql = StrSQL
End Function

