VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmMaintanenceSearch1 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "m"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8640
   Icon            =   "FrmMaintanenceSearch1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   3495
   ScaleWidth      =   8640
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox XPTxtSerial 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   1320
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   27
      Top             =   2640
      Width           =   2235
   End
   Begin MSDataListLib.DataCombo DcboClients 
      Height          =   315
      Left            =   4920
      TabIndex        =   26
      Top             =   2610
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
      Left            =   -4080
      RightToLeft     =   -1  'True
      TabIndex        =   24
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
      TabIndex        =   19
      Top             =   5610
      Visible         =   0   'False
      Width           =   1995
      Begin MSComCtl2.DTPicker XPDtbGoInDtae 
         Height          =   345
         Left            =   60
         TabIndex        =   20
         Top             =   240
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   609
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   126615553
         CurrentDate     =   38784
      End
      Begin MSComCtl2.DTPicker XPDtbGoOutDtae 
         Height          =   345
         Left            =   60
         TabIndex        =   21
         Top             =   615
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   609
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   126615553
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
         TabIndex        =   23
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
         TabIndex        =   22
         Top             =   300
         Width           =   285
      End
   End
   Begin VB.ComboBox CboMaintenanceType 
      Height          =   315
      Left            =   -2520
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   2580
      Width           =   2295
   End
   Begin VB.ComboBox CboPaymentType 
      Height          =   315
      Left            =   -1800
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   5370
      Width           =   2295
   End
   Begin VB.TextBox TxtTransID 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   6420
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   2220
      Width           =   1155
   End
   Begin VB.TextBox XPTxtName 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   2490
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   2250
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Height          =   975
      Index           =   1
      Left            =   3300
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   3570
      Width           =   5235
      Begin VB.TextBox TxtItemCode 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   3030
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   150
         Width           =   1275
      End
      Begin MSDataListLib.DataCombo DCboItemsName 
         Height          =   315
         Left            =   420
         TabIndex        =   7
         Top             =   510
         Width           =   3885
         _ExtentX        =   6853
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ﬂÊœ «·’‰›"
         Height          =   285
         Index           =   9
         Left            =   4350
         RightToLeft     =   -1  'True
         TabIndex        =   18
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
         TabIndex        =   3
         Top             =   540
         Width           =   795
      End
   End
   Begin VSFlex8UCtl.VSFlexGrid FG 
      Height          =   2175
      Left            =   0
      TabIndex        =   0
      Top             =   30
      Width           =   8595
      _cx             =   15161
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
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmMaintanenceSearch1.frx":030A
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
      Left            =   2970
      TabIndex        =   4
      Top             =   3120
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
      Left            =   2070
      TabIndex        =   5
      Top             =   3120
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
      Left            =   1170
      TabIndex        =   6
      Top             =   3120
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
      Left            =   6930
      TabIndex        =   16
      Top             =   3120
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
      ButtonImage     =   "FrmMaintanenceSearch1.frx":045A
      ColorButton     =   14871017
      ColorHoverText  =   12582912
      ButtonToggles   =   1
      DrawFocusRectangle=   0   'False
      RightToLeft     =   -1  'True
      ButtonImageToggled=   "FrmMaintanenceSearch1.frx":07F4
      ColorToggledHoverText=   192
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·”Ì—Ì«·"
      Height          =   315
      Index           =   2
      Left            =   3720
      RightToLeft     =   -1  'True
      TabIndex        =   28
      Top             =   2700
      Width           =   795
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·⁄„Ì· «·‰ﬁœÏ"
      Height          =   315
      Index           =   10
      Left            =   -2670
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   3720
      Width           =   1275
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   225
      Index           =   8
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   2250
      Width           =   2205
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "ÿ—Ìﬁ…  «·œ›⁄"
      Height          =   315
      Index           =   7
      Left            =   540
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   5370
      Width           =   855
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "‰Ê⁄ «·’Ì«‰…"
      Height          =   345
      Index           =   4
      Left            =   -180
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   6300
      Width           =   855
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·«Ì’«·"
      Height          =   285
      Index           =   0
      Left            =   7380
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   2250
      Width           =   1095
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·⁄„Ì·"
      Height          =   345
      Index           =   5
      Left            =   7620
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   2595
      Width           =   855
   End
End
Attribute VB_Name = "FrmMaintanenceSearch1"
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

                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
                Else
                    Msg = "No Data"
                End If

                MsgBox Msg, vbOKOnly + vbInformation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Exit Sub
            End If

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.lbl(8).Caption = "‰ ÌÃ… «·»ÕÀ : " & rs.RecordCount
            Else
                Me.lbl(8).Caption = "Search Result : " & rs.RecordCount
            End If
       
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
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & CHR(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… „⁄«ÌÌ— «·»ÕÀ Ê√⁄œ «·„Õ«Ê·…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
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

Private Sub fg_Click()
    On Error GoTo ErrTrap

    If Not FG.TextMatrix(FG.Row, 2) = "" Then
        If Me.ExtraRetrunObject Is Nothing Then
            mdifrmmain.ActiveForm.Retrive val(FG.TextMatrix(FG.Row, 2))
        Else
            Me.ExtraRetrunObject = val(FG.TextMatrix(FG.Row, 2))
        
        End If
    
        ' FrmManCusRecive.TxtOrgManSerial.text = val(FG.TextMatrix(FG.Row, 2))
    
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
                .TextMatrix(Num, .ColIndex("ReciptNumber")) = IIf(IsNull(rs("ReciptNumber").value), "", (rs("ReciptNumber").value))

                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(Num, .ColIndex("Name")) = IIf(IsNull(rs("CusName").value), "", Trim(rs("CusName").value))
                    .TextMatrix(Num, .ColIndex("itemname")) = IIf(IsNull(rs("ItemName").value), "", Trim(rs("ItemName").value))
                Else
                    .TextMatrix(Num, .ColIndex("Name")) = IIf(IsNull(rs("CusNamee").value), "", Trim(rs("CusNamee").value))
                    .TextMatrix(Num, .ColIndex("itemname")) = IIf(IsNull(rs("ItemNamee").value), "", Trim(rs("ItemNamee").value))
                End If
            
                If Not IsNull(rs("CashCustomerName").value) Then
                    StrCusName = IIf(IsNull(rs("CusName").value), "", Trim(rs("CusName").value))
                    StrCusName = StrCusName + "(" + rs("CashCustomerName").value + ")"
                Else
                    StrCusName = IIf(IsNull(rs("CusName").value), "", Trim(rs("CusName").value))
                End If
             
                .TextMatrix(Num, .ColIndex("Item_Serial")) = IIf(IsNull(rs("ItemSerial").value), "", Trim(rs("ItemSerial").value))

                '.TextMatrix(Num, .ColIndex("Name")) = StrCusName
                '.TextMatrix(Num, .ColIndex("DateGoIn")) = IIf(IsNull(Rs("DateGoIN").value), "", Format((Rs("DateGoIN").value), "yyyy/m/d"))
                '.TextMatrix(Num, .ColIndex("DateGoOut")) = IIf(IsNull(Rs("DateGoOUT").value), "", Format((Rs("DateGoOUT").value), "yyyy/m/d"))
            End With

            rs.MoveNext
        Next Num

    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Fg_DblClick()
    fg_Click
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

    With Me.CboPayMentType
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

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    Exit Sub
ErrTrap:
End Sub

Function ChangeLang()
 
    Me.Caption = "Maintenance Search"
 
    With Me.FG
        .TextMatrix(0, .ColIndex("NumIndex")) = "Index"
        .TextMatrix(0, .ColIndex("ReciptNumber")) = "Recipt Number"
        .TextMatrix(0, .ColIndex("Name")) = "Customer Number"
        .TextMatrix(0, .ColIndex("ItemName")) = "Item Number"
        .TextMatrix(0, .ColIndex("Item_Serial")) = "Item Serial"

    End With

    lbl(0).Caption = "Recipt Number"
    lbl(5).Caption = "Customer"
    lbl(2).Caption = "Serial No"

    CmdShowMoreOptions.Caption = "Adv Search"
    Cmd(0).Caption = "Search"
    Cmd(1).Caption = "Clear"
    Cmd(2).Caption = "Exit"
    lbl(9).Caption = "Item Code"
    lbl(6).Caption = "Item Name"

End Function

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
    StrSQL = " SELECT     dbo.TblMainteneceNew.ReciptNumber, dbo.TblMainteneceNew.Transaction_ID, dbo.TblMainteneceNew.CusID, dbo.TblMainteneceNew.CashCustomerName, "
    StrSQL = StrSQL & "  dbo.TblMainteneceNew.CashCustomerPhone, dbo.TblMainteneceNew.CashCustomerMobile, dbo.TblMainteneceNew.CashCustomerEmail,"
    StrSQL = StrSQL & " dbo.TblMainteneceNew.CashCustomerAddress, dbo.TblMainteneceNew.OperationDate, dbo.TblMainteneceNew.DateGoIN, dbo.TblMainteneceNew.DateGoOUT,"
    StrSQL = StrSQL & " dbo.TblMainteneceNew.EmpID, dbo.TblMainteneceNew.StoreID, dbo.TblMainteneceNew.Remarks, dbo.TblMainteneceNew.GoOut, dbo.TblMainteneceNew.UserID,"
    StrSQL = StrSQL & " dbo.TblMainteneceNew.PaymentType, dbo.TblMainteneceNew.MType, dbo.TblMainteneceNew.ManOperationTypeID, dbo.TblMainteneceNew.ItemID,"
    StrSQL = StrSQL & " dbo.TblMainteneceNew.ItemSerial, dbo.TblMainteneceNew.Quantity, dbo.TblMainteneceNew.TicketNO, dbo.TblMainteneceNew.CustomerNotes,"
    StrSQL = StrSQL & " dbo.TblMainteneceNew.EmpNotes, dbo.TblMainteneceNew.Cost, dbo.TblMainteneceNew.SupDeci, dbo.TblMainteneceNew.MainOperationID,"
    StrSQL = StrSQL & " dbo.TblMainteneceNew.RetrunOrgID, dbo.TblMainteneceNew.FastReplace, dbo.TblMainteneceNew.FastReplaceType, dbo.TblMainteneceNew.ReItemID,"
    StrSQL = StrSQL & " dbo.TblMainteneceNew.ReItemSerial, dbo.TblMainteneceNew.ReItemQuantity, dbo.TblMainteneceNew.ReItemPrice, dbo.TblMainteneceNew.ReItemStore,"
    StrSQL = StrSQL & " dbo.TblItems.ItemCode, dbo.TblItems.ItemName, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblItems.ItemNamee, dbo.TblStore.StoreName,"
    StrSQL = StrSQL & " dbo.TblStore.StoreNamee"
    StrSQL = StrSQL & " FROM         dbo.TblMainteneceNew INNER JOIN"
    StrSQL = StrSQL & " dbo.TblItems ON dbo.TblMainteneceNew.ItemID = dbo.TblItems.ItemID LEFT OUTER JOIN"
    StrSQL = StrSQL & " dbo.TblCustemers ON dbo.TblMainteneceNew.CusID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
    StrSQL = StrSQL & " dbo.TblStore ON dbo.TblMainteneceNew.StoreID = dbo.TblStore.StoreID"

    If Me.SearchType = 2 Then
        StrWhere = " Where (dbo.TblMainteneceNew.ManOperationTypeID = 1) And (dbo.TblMainteneceNew.GoOut Is Null)"
    Else
        StrWhere = " Where (dbo.TblMainteneceNew.ManOperationTypeID = 4) And (dbo.TblMainteneceNew.GoOut Is Null)"
    End If

    BolBegin = True

    If Me.TxtTransID.Text <> "" Then
        'MaintananceID
 
        StrWhere = StrWhere + "   AND (dbo.TblMainteneceNew.ReciptNumber = N'" & TxtTransID.Text & "') "
 
    End If
 
    If Me.DcboClients.BoundText <> "" Then
 
        StrWhere = StrWhere + " and dbo.TblMainteneceNew.CusID =" & Me.DcboClients.BoundText & ""
 
    End If
 
    If Me.DCboItemsName.BoundText <> "" Then
        If BolBegin = True Then
            StrWhere = StrWhere + " and dbo.TblMainteneceNew.ItemID=" & DCboItemsName.BoundText
        Else
            StrWhere = StrWhere + " where dbo.TblMainteneceNew.ItemID=" & DCboItemsName.BoundText
            BolBegin = True
        End If
    End If

    If Me.XPTxtSerial.Text <> "" Then
        If BolBegin = True Then
            StrWhere = StrWhere + " and dbo.TblMainteneceNew.ItemSerial='" & XPTxtSerial.Text & "'"
        Else
            StrWhere = StrWhere + " where dbo.TblMainteneceNew.ItemSerial='" & XPTxtSerial.Text & "'"
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
            fg_Click
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
        If Trim$(Me.txtItemCode.Text) <> "" Then
            Me.DCboItemsName.BoundText = GetItemID(Trim$(Me.txtItemCode.Text))
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
