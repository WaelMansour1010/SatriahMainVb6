VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmMaintanenceSearch 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«·»ÕÀ ⁄‰ ⁄„·Ì«  «·’Ì«‰…"
   ClientHeight    =   5940
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5730
   Icon            =   "FrmMaintanenceSearch.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5940
   ScaleWidth      =   5730
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin MSDataListLib.DataCombo DcboClients 
      Height          =   315
      Left            =   2160
      TabIndex        =   28
      Top             =   3330
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VB.TextBox TxtCusCashName 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   480
      RightToLeft     =   -1  'True
      TabIndex        =   26
      Top             =   3690
      Width           =   3885
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   " «—ÌŒ «·√” ·«„"
      Height          =   1035
      Index           =   0
      Left            =   30
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   2610
      Width           =   1995
      Begin MSComCtl2.DTPicker XPDtbGoInDtae 
         Height          =   345
         Left            =   60
         TabIndex        =   22
         Top             =   240
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   609
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   100073473
         CurrentDate     =   38784
      End
      Begin MSComCtl2.DTPicker XPDtbGoOutDtae 
         Height          =   345
         Left            =   60
         TabIndex        =   23
         Top             =   615
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   609
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   100073473
         CurrentDate     =   38784
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï"
         Height          =   255
         Index           =   3
         Left            =   1620
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   705
         Width           =   285
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   315
         Index           =   1
         Left            =   1620
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   300
         Width           =   285
      End
   End
   Begin VB.ComboBox CboMaintenanceType 
      Height          =   315
      Left            =   2520
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   2580
      Width           =   2295
   End
   Begin VB.ComboBox CboPaymentType 
      Height          =   315
      Left            =   2520
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   2970
      Width           =   2295
   End
   Begin VB.TextBox TxtTransID 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   3660
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   2220
      Width           =   1155
   End
   Begin VB.TextBox XPTxtName 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   2490
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   2250
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Height          =   1335
      Index           =   1
      Left            =   420
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   4530
      Width           =   5235
      Begin VB.TextBox TxtItemCode 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   3030
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   150
         Width           =   1275
      End
      Begin MSDataListLib.DataCombo DCboItemsName 
         Height          =   315
         Left            =   420
         TabIndex        =   8
         Top             =   510
         Width           =   3885
         _ExtentX        =   6853
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.TextBox XPTxtSerial 
         Alignment       =   1  'Right Justify
         Height          =   360
         Left            =   2040
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   840
         Width           =   2235
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ﬂÊœ «·’‰›"
         Height          =   285
         Index           =   9
         Left            =   4350
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   180
         Width           =   795
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«”„ «·’‰›"
         Height          =   315
         Index           =   6
         Left            =   4380
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   540
         Width           =   795
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·”Ì—Ì«·"
         Height          =   315
         Index           =   2
         Left            =   4320
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   900
         Width           =   795
      End
   End
   Begin VSFlex8UCtl.VSFlexGrid FG 
      Height          =   2175
      Left            =   0
      TabIndex        =   0
      Top             =   30
      Width           =   5715
      _cx             =   10081
      _cy             =   3836
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
      Rows            =   15
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmMaintanenceSearch.frx":030A
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
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   0
      Left            =   2250
      TabIndex        =   5
      Top             =   4080
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      ButtonStyle     =   1
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
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      DisabledImageExtraction=   0
      ColorToggledHoverText=   16711680
      LowerToggledContent=   0   'False
      ColorTextShadow =   4210752
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   1
      Left            =   1350
      TabIndex        =   6
      Top             =   4080
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      ButtonStyle     =   1
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
      ColorShadow     =   4210752
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      ColorToggledHoverText=   16711680
      LowerToggledContent=   0   'False
      ColorTextShadow =   4210752
   End
   Begin ImpulseButton.ISButton Cmd 
      Cancel          =   -1  'True
      Height          =   375
      Index           =   2
      Left            =   450
      TabIndex        =   7
      Top             =   4080
      Width           =   855
      _ExtentX        =   1508
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
   Begin ImpulseButton.ISButton CmdShowMoreOptions 
      Height          =   375
      Left            =   4290
      TabIndex        =   18
      Top             =   4080
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   661
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "»ÕÀ „ ﬁœ„..."
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
      ButtonImage     =   "FrmMaintanenceSearch.frx":03DF
      ColorButton     =   14871017
      ColorHoverText  =   12582912
      ButtonToggles   =   1
      DrawFocusRectangle=   0   'False
      RightToLeft     =   -1  'True
      ButtonImageToggled=   "FrmMaintanenceSearch.frx":0779
      ColorToggledHoverText=   192
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·⁄„Ì· «·‰ﬁœÏ"
      Height          =   315
      Index           =   10
      Left            =   4410
      RightToLeft     =   -1  'True
      TabIndex        =   27
      Top             =   3720
      Width           =   1275
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "‰ ÌÃ… «·»ÕÀ: "
      Height          =   225
      Index           =   8
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   2250
      Width           =   2205
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "ÿ—Ìﬁ…  «·œ›⁄"
      Height          =   315
      Index           =   7
      Left            =   4860
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   2970
      Width           =   855
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "‰Ê⁄ «·’Ì«‰…"
      Height          =   345
      Index           =   4
      Left            =   4860
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   2580
      Width           =   855
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·⁄„·Ì…"
      Height          =   285
      Index           =   0
      Left            =   4860
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   2250
      Width           =   855
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·⁄„Ì·"
      Height          =   345
      Index           =   5
      Left            =   4860
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   3315
      Width           =   855
   End
End
Attribute VB_Name = "FrmMaintanenceSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim cSearchDcbo(1)  As clsDCboSearch
Dim M_ExtraRetrunObject As Object
Dim m_SearchType As Integer

Private Sub Cmd_Click(Index As Integer)
    Dim Msg As String

    On Error GoTo ErrTrap

    Select Case Index

        Case 0

            If rs.State = adStateOpen Then
                rs.Close
            End If

            rs.Open Build_Sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

            If rs.RecordCount < 1 Then
                FG.Clear flexClearScrollable, flexClearEverything
                Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Exit Sub
            End If

            Me.lbl(8).Caption = "‰ ÌÃ… «·»ÕÀ : " & rs.RecordCount
            Retrive

        Case 1
            clear_all Me
            FG.Clear flexClearScrollable, flexClearEverything

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

Private Sub CmdShowMoreOptions_Click()

    If CmdShowMoreOptions.value = True Then
        Me.Fra(1).Visible = True
        'Me.Height = Me.Fra(1).top + Fra(1).Height + 400
        Me.Height = Me.Fra(1).top + Fra(1).Height + 500 ' GetMyTitleBarHight(Me.hwnd)
        'Me.Height = Me.ScaleHeight
    Else
        Me.Fra(1).Visible = False
        Me.Height = Me.Fra(1).top + 500
    End If

End Sub

Private Sub DcboClients_Change()

    If val(DcboClients.BoundText) = 1 Or val(Me.DcboClients.BoundText) = 2 Then
        lbl(10).Enabled = True
        Me.TxtCusCashName.Enabled = True
    Else
        lbl(10).Enabled = False
        Me.TxtCusCashName.Enabled = False
    End If

End Sub

Private Sub Fg_Click()
    On Error GoTo ErrTrap

    If Not FG.TextMatrix(FG.Row, 1) = "" Then
        If Me.ExtraRetrunObject Is Nothing Then
            mdifrmmain.ActiveForm.Retrive val(FG.TextMatrix(FG.Row, 1))
        Else
            Me.ExtraRetrunObject = val(FG.TextMatrix(FG.Row, 1))
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Public Property Get ExtraRetrunObject() As Object
    Set ExtraRetrunObject = M_ExtraRetrunObject
End Property

Public Property Set ExtraRetrunObject(ByVal vNewValue As Object)
    'ﬁ„  »⁄„· Â–Â «·Œ«’Ì… „Œ’Ê’ Õ Ï Ì„ﬂ‰‰Ï
    '«‰ «” Œœ„ ‘«‘… «·»ÕÀ ⁄‰ «·Õ—ﬂ«  «·’Ì«‰…
    '„‰ Œ·«· ‘«‘… «·„ﬁ»Ê÷«  ÕÌÀ Ì„ﬂ‰‰Ï
    '«‰ «” —Ã⁄ ﬂÊœ «·Õ—ﬂ… «· Ã«—Ì…
    '›Ï ‘«‘… „À· ‘«‘… «·„ﬁ»Ê÷« 
    Set M_ExtraRetrunObject = vNewValue
End Property

Private Sub Retrive()
    Dim Num As Integer
    Dim StrCusName As String
    On Error GoTo ErrTrap
    FG.Clear flexClearScrollable, flexClearEverything

    If Not (rs.EOF Or rs.BOF) Then
        FG.Rows = rs.RecordCount + 1

        For Num = 1 To rs.RecordCount

            With FG
                .TextMatrix(Num, .ColIndex("NumIndex")) = Num
                .TextMatrix(Num, .ColIndex("Num")) = IIf(IsNull(rs("MaintananceID").value), "", (rs("MaintananceID").value))
                .TextMatrix(Num, .ColIndex("Name")) = IIf(IsNull(rs("CusName").value), "", Trim(rs("CusName").value))

                If Not IsNull(rs("CashCustomerName").value) Then
                    StrCusName = IIf(IsNull(rs("CusName").value), "", Trim(rs("CusName").value))
                    StrCusName = StrCusName + "(" + rs("CashCustomerName").value + ")"
                Else
                    StrCusName = IIf(IsNull(rs("CusName").value), "", Trim(rs("CusName").value))
                End If

                .TextMatrix(Num, .ColIndex("Name")) = StrCusName
                .TextMatrix(Num, .ColIndex("DateGoIn")) = IIf(IsNull(rs("DateGoIN").value), "", Format((rs("DateGoIN").value), "yyyy/m/d"))
                .TextMatrix(Num, .ColIndex("DateGoOut")) = IIf(IsNull(rs("DateGoOUT").value), "", Format((rs("DateGoOUT").value), "yyyy/m/d"))
            End With

            rs.MoveNext
        Next Num

    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Fg_DblClick()
    Fg_Click
    Unload Me
End Sub

Private Sub Form_Load()
    Dim BG As ClsBackGroundPic
    Dim Num As Integer
    Dim Dcombos As ClsDataCombos
    Set rs = New ADODB.Recordset
    Dim RsTemp As ADODB.Recordset
    Dim StrSQL As String
    On Error GoTo ErrTrap
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Clear").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture

    CenterForm Me

    FormPostion Me, GetPostion
    SetDtpickerDate XPDtbGoInDtae
    SetDtpickerDate XPDtbGoOutDtae

    Set BG = New ClsBackGroundPic
    FG.WallPaper = BG.SearchWallpaper
    Set Dcombos = New ClsDataCombos
    Dcombos.GetCustomersSuppliers 0, Me.DcboClients, True
    Dcombos.GetItemsNames Me.DCboItemsName

    With Me.CboPaymentType
        .Clear
        .AddItem "‰ﬁœÌ"
        .AddItem "«Ã·"
        .AddItem "«·ﬂ·"
    End With

    With CboMaintenanceType
        .AddItem "»«· ﬂ·›…"
        .AddItem " »⁄ «·÷„«‰"
        .AddItem "«·ﬂ·"
    End With

    Set cSearchDcbo(0) = New clsDCboSearch
    Set cSearchDcbo(0).Client = Me.DcboClients

    Set cSearchDcbo(1) = New clsDCboSearch
    Set cSearchDcbo(1).Client = Me.DCboItemsName

    CmdShowMoreOptions.value = False
    CmdShowMoreOptions_Click

    lbl(10).Enabled = False
    Me.TxtCusCashName.Enabled = False
    Exit Sub
ErrTrap:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer
    On Error GoTo ErrTrap

    If rs.State = adStateOpen Then
        rs.Close
        Set rs = Nothing
    End If

    FormPostion Me, SavePostion
    For i = LBound(cSearchDcbo) To UBound(cSearchDcbo)
        Set cSearchDcbo(i) = Nothing
    Next i

    Exit Sub
ErrTrap:
End Sub

Private Function Build_Sql()
    Dim StrSQL As String
    Dim BolBegin As Boolean
    Dim StrWhere As String
    Dim StrWhereItem As String

    On Error GoTo ErrTrap

    'StrSQL = "select * From QryMaintanenceSearch"
    StrSQL = "SELECT Distinct dbo.TblMaintenece.MaintananceID, dbo.TblMaintenece.CusID, dbo.TblMaintenece.DateGoIN," & "dbo.TblMaintenece.DateGoOUT,dbo.TblMaintenece.GoOut," & "dbo.TblCustemers.CusName, dbo.TblMaintenece.PaymentType, dbo.TblMaintenece.MType," & "dbo.TblMaintenece.CashCustomerName FROM  " & "dbo.TblMainteneceDetails INNER JOIN dbo.TblMaintenece ON dbo.TblMainteneceDetails." & "MaintananceID = dbo.TblMaintenece.MaintananceID LEFT OUTER JOIN dbo.TblCustemers ON " & "dbo.TblMaintenece.CusID= dbo.TblCustemers.CusID"

    BolBegin = True
    StrWhere = " Where dbo.TblMaintenece.ManOperationTypeID=" & Me.SearchType

    If Me.TxtTransID.text <> "" Then

        'MaintananceID
        If BolBegin = True Then
            StrWhere = StrWhere + " and MaintananceID =" & Trim(TxtTransID.text) & ""
        Else
            StrWhere = StrWhere + " Where MaintananceID =" & Trim(TxtTransID.text) & ""
            BolBegin = True
        End If
    End If

    If Me.CboPaymentType.ListIndex <> -1 Then
        If Me.CboPaymentType.ListIndex = 0 Then
    
            If BolBegin = True Then
                StrWhere = StrWhere + " and PaymentType=0 "
            Else
                StrWhere = StrWhere + " Where PaymentType=0 "
                BolBegin = True
            End If
        
        ElseIf Me.CboPaymentType.ListIndex = 1 Then

            If BolBegin = True Then
                StrWhere = StrWhere + " and PaymentType=1 "
            Else
                StrWhere = StrWhere + " Where PaymentType=1 "
                BolBegin = True
            End If
        End If
    End If

    If Me.CboMaintenanceType.ListIndex <> -1 Then
        If Me.CboMaintenanceType.ListIndex = 0 Then
            If BolBegin = True Then
                StrWhere = StrWhere + " and MType=0 "
            Else
                StrWhere = StrWhere + " Where MType=0 "
                BolBegin = True
            End If

        ElseIf Me.CboMaintenanceType.ListIndex = 1 Then

            If BolBegin = True Then
                StrWhere = StrWhere + " and MType=1 "
            Else
                StrWhere = StrWhere + " Where MType=1 "
                BolBegin = True
            End If
        End If
    End If

    If Me.DcboClients.BoundText <> "" Then
        If BolBegin = True Then
            StrWhere = StrWhere + " and dbo.TblMaintenece.CusID =" & Me.DcboClients.BoundText & ""
        Else
            StrWhere = StrWhere + " where dbo.TblMaintenece.CusID =" & Me.DcboClients.BoundText & ""
            BolBegin = True
        End If
    End If

    If Me.TxtCusCashName.Enabled = True Then
        If Trim$(Me.TxtCusCashName.text) <> "" Then
            If BolBegin = True Then
                StrWhere = StrWhere + " and CashCustomerName Like '%" & Trim$(Me.TxtCusCashName.text) & "%'"
            Else
                StrWhere = StrWhere + " where CashCustomerName '%" & Trim$(Me.TxtCusCashName.text) & "%'"
                BolBegin = True
            End If
        End If
    End If

    If XPDtbGoInDtae.value <> "" Then
        If BolBegin = True Then
            StrWhere = StrWhere + " and DateGoIN >=" & SQLDate(XPDtbGoInDtae.value, True) & ""
        Else
            StrWhere = StrWhere + " where DateGoIN >=" & SQLDate(XPDtbGoInDtae.value, True) & ""
            BolBegin = True
        End If
    End If

    If XPDtbGoOutDtae.value <> "" Then
        If BolBegin = True Then
            StrWhere = StrWhere + " and DateGoOUT=" & SQLDate(XPDtbGoOutDtae.value, True) & ""
        Else
            StrWhere = StrWhere + " where DateGoOUT=" & SQLDate(XPDtbGoOutDtae.value, True) & ""
            BolBegin = True
        End If
    End If

    If Me.DCboItemsName.BoundText <> "" Then
        If BolBegin = True Then
            StrWhere = StrWhere + " and dbo.TblMainteneceDetails.ItemID=" & DCboItemsName.BoundText
        Else
            StrWhere = StrWhere + " where dbo.TblMainteneceDetails.ItemID=" & DCboItemsName.BoundText
            BolBegin = True
        End If
    End If

    If Me.XPTxtSerial.text <> "" Then
        If BolBegin = True Then
            StrWhere = StrWhere + " and dbo.TblMainteneceDetails.ItemSerial='" & XPTxtSerial.text & "'"
        Else
            StrWhere = StrWhere + " where dbo.TblMainteneceDetails.ItemSerial='" & XPTxtSerial.text & "'"
            BolBegin = True
        End If
    End If

    Build_Sql = StrSQL + StrWhere

    Exit Function
ErrTrap:
End Function

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        If Not FG.TextMatrix(FG.Row, 1) = "" Then
            Fg_Click
        Else
            Cmd_Click (0)
        End If
    End If

    If Shift = vbShiftMask Then
        If KeyCode = vbKeyEscape Then
            Cmd_Click (2)
        End If
    End If

    Exit Sub
    Exit Sub
ErrTrap:
End Sub

Private Sub TxtItemCode_KeyDown(KeyCode As Integer, _
                                Shift As Integer)

    If KeyCode = vbKeyReturn Then
        If Trim$(Me.TxtItemCode.text) <> "" Then
            Me.DCboItemsName.BoundText = GetItemID(Trim$(Me.TxtItemCode.text))
        End If
    End If

End Sub

Public Property Get SearchType() As Integer
    SearchType = m_SearchType
End Property

Public Property Let SearchType(ByVal vNewValue As Integer)
    m_SearchType = vNewValue

    If m_SearchType = 1 Then
        Me.Caption = "«·’Ì«‰… ( œŒÊ· ··’Ì«‰…)"
    ElseIf m_SearchType = 2 Then
        Me.Caption = "«·’Ì«‰… ( Œ—ÊÃ ÷„«‰ ··„Ê—œ)"
    ElseIf m_SearchType = 3 Then
        Me.Caption = "«·’Ì«‰… ( —ÃÊ⁄ „‰ «·„Ê—œ )"
    ElseIf m_SearchType = 4 Then
        Me.Caption = "«·’Ì«‰… (  ”·Ì„ ··⁄„Ì· )"
    End If

End Property
