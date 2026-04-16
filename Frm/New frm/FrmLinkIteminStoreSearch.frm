VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmLinkIteminStoreSearch 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "«·»ÕÀ ⁄‰ «·«’‰«› «·„— »ÿ… »«·„Œ«“‰"
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8355
   Icon            =   "FrmLinkIteminStoreSearch.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   4890
   ScaleWidth      =   8355
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
   Begin VB.TextBox TxtItemCode 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   6600
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   3390
      Width           =   675
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   " «—ÌŒ «· ”ÃÌ·"
      Height          =   1035
      Index           =   1
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   3270
      Width           =   4215
      Begin MSComCtl2.DTPicker DtpDateFrom 
         Height          =   330
         Left            =   90
         TabIndex        =   6
         Top             =   270
         Width           =   2910
         _ExtentX        =   5133
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   93323267
         CurrentDate     =   38887
      End
      Begin MSComCtl2.DTPicker DtpDateTo 
         Height          =   330
         Left            =   90
         TabIndex        =   7
         Top             =   630
         Width           =   2910
         _ExtentX        =   5133
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   93323267
         CurrentDate     =   38887
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï"
         Height          =   195
         Index           =   3
         Left            =   3375
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   660
         Width           =   495
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   195
         Index           =   4
         Left            =   3180
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   330
         Width           =   660
      End
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·⁄„·Ì…"
      Height          =   645
      Index           =   2
      Left            =   2880
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   2700
      Width           =   5475
      Begin VB.TextBox TxtIDFrom 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   180
         Width           =   1635
      End
      Begin VB.TextBox TxtIDTO 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   180
         Width           =   1515
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   195
         Index           =   5
         Left            =   4335
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   240
         Width           =   660
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï"
         Height          =   195
         Index           =   6
         Left            =   1860
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   240
         Width           =   645
      End
   End
   Begin VSFlex8UCtl.VSFlexGrid Fg 
      Height          =   2625
      Left            =   30
      TabIndex        =   10
      Top             =   0
      Width           =   8355
      _cx             =   14737
      _cy             =   4630
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
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmLinkIteminStoreSearch.frx":038A
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
   Begin MSDataListLib.DataCombo DCItem 
      Height          =   315
      Left            =   4560
      TabIndex        =   11
      Top             =   3390
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   16777215
      Text            =   "DCEmp_Name"
      RightToLeft     =   -1  'True
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   0
      Left            =   1650
      TabIndex        =   12
      Top             =   4440
      Width           =   765
      _ExtentX        =   1349
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
      Left            =   810
      TabIndex        =   13
      Top             =   4440
      Width           =   795
      _ExtentX        =   1402
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
      TabIndex        =   14
      Top             =   4440
      Width           =   735
      _ExtentX        =   1296
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
   Begin MSDataListLib.DataCombo DCStore 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   4560
      TabIndex        =   15
      Tag             =   "„‰ ›÷·ﬂ √œŒ· —ﬁ„ «·ﬁ÷Ì…"
      Top             =   3750
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   -2147483624
      Text            =   ""
      RightToLeft     =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo DCGroup 
      CausesValidation=   0   'False
      Height          =   315
      Left            =   4560
      TabIndex        =   21
      Tag             =   "„‰ ›÷·ﬂ √œŒ· —ﬁ„ «·ﬁ÷Ì…"
      Top             =   4080
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   556
      _Version        =   393216
      BackColor       =   -2147483624
      Text            =   ""
      RightToLeft     =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·„Ã„Ê⁄…"
      Height          =   285
      Index           =   8
      Left            =   7380
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   4110
      Width           =   945
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·≈Ã„«·Ï"
      Height          =   285
      Index           =   2
      Left            =   1890
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   3060
      Width           =   945
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      ForeColor       =   &H00000080&
      Height          =   285
      Index           =   1
      Left            =   60
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   3060
      Width           =   1785
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·’‰›"
      Height          =   195
      Index           =   0
      Left            =   7635
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   3390
      Width           =   720
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·„Œ“‰"
      Height          =   285
      Index           =   7
      Left            =   7380
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   3780
      Width           =   945
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      ForeColor       =   &H00000080&
      Height          =   315
      Index           =   10
      Left            =   60
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   2700
      Width           =   2775
   End
End
Attribute VB_Name = "FrmLinkIteminStoreSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DCboSearch As clsDCboSearch

Private Sub Cmd_Click(Index As Integer)

    Select Case Index

        Case 0
            GetData

        Case 1
            clear_all Me
DtpDateFrom.value = ""
Me.DtpDateTo.value = ""
            If SystemOptions.UserInterface = ArabicInterface Then
                Me.lbl(0).Caption = "‰ ÌÃ… «·»ÕÀ"
            Else
                Me.lbl(0).Caption = "Search Results"
            End If

        Case 2
            Unload Me
    End Select

End Sub

Private Sub Fg_Click()

    With Me.Fg

        If .Row = -1 Then Exit Sub
        If .Col = -1 Then Exit Sub
        If val(.TextMatrix(.Row, .ColIndex("ID"))) = 0 Then
            Exit Sub
        End If

                FrmLinkItemToStore.Retrive val(.TextMatrix(.Row, .ColIndex("ID")))
       
    

    End With

End Sub

Private Sub Form_Activate()
    PutFormOnTop Me.hwnd
End Sub

Private Sub ChangeLang()
    'Dim XPic As IPictureDisp
    'Set XPic = Me.XPBtnMove(1).ButtonImage
    'Set Me.XPBtnMove(1).ButtonImage = Me.XPBtnMove(2).ButtonImage
    'Set Me.XPBtnMove(2).ButtonImage = XPic
    'Set XPic = Me.XPBtnMove(0).ButtonImage
    'Set Me.XPBtnMove(0).ButtonImage = Me.XPBtnMove(3).ButtonImage
    'Set Me.XPBtnMove(3).ButtonImage = XPic
'    Label1.Visible = False

    'Cmd(0).Caption = "New"
    'Cmd(1).Caption = "Edit"
    'Cmd(2).Caption = "Save"
    'Cmd(3).Caption = "Undo"
    Cmd(1).Caption = "Delete"
    Cmd(0).Caption = "Search"
 'Cmd(9).Caption = "Print"
    Cmd(2).Caption = "Exit"
 '   CmdHelp.Caption = "Help"

    Me.Caption = " Search For Linking Items in Storage "
    'EleHeader.Caption = Me.Caption
    lbl(5).Caption = "From"
    lbl(6).Caption = "To"
    lbl(4).Caption = "From"
    lbl(3).Caption = "To "
   ' Frame10.Caption = "Select Store"
    Fra(2).Caption = "Registration No  "
    Fra(1).Caption = "Registration Date  "
lbl(7).Caption = "StoreName"
lbl(8).Caption = "GroupName"
lbl(0).Caption = "ItemName"
lbl(2).Caption = "Total"
   'lbl(8).Caption = "By"
   ' lbl(7).Caption = "Curr rec."
   ' lbl(6).Caption = "rec. count"

   With Me.Fg
        .TextMatrix(0, .ColIndex("Serial")) = "Serial"
        .TextMatrix(0, .ColIndex("ID")) = "NO"
        .TextMatrix(0, .ColIndex("StoreName")) = "StoreName"
         .TextMatrix(0, .ColIndex("RecordDate")) = "RecordDate"
        .TextMatrix(0, .ColIndex("GroupName")) = "GroupName"
         .TextMatrix(0, .ColIndex("ItemName")) = "ItemName"
        .TextMatrix(0, .ColIndex("LinkType")) = "LinkType"

    End With

End Sub

Private Sub Form_Load()
    Dim GrdBack As ClsBackGroundPic
    Dim Dcombos As ClsDataCombos

    Set Dcombos = New ClsDataCombos
    Dcombos.GetItemsNames DCItem, , , , True
    Set DCboSearch = New clsDCboSearch
    Set DCboSearch.Client = Me.DCItem
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    Dcombos.GetItemSGroups Me.DCGroup, False

Dcombos.GetStores DCStore
  
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Clear").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture

    CenterForm Me

    FormPostion Me, GetPostion
    Set GrdBack = New ClsBackGroundPic

    With Me.Fg
        Set .WallPaper = GrdBack.Picture
        .AutoSize 0, .Cols - 1, False
    End With

    SetDtpickerDate Me.DtpDateFrom
    SetDtpickerDate Me.DtpDateTo

End Sub

Private Sub Form_Unload(Cancel As Integer)

    FormPostion Me, SavePostion
    Set DCboSearch = Nothing
End Sub

Private Sub GetData()
    Dim StrSQL As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
StrSQL = " SELECT     dbo.TblLink_Item_To_StoreH.Ind, dbo.TblLink_Item_To_StoreH.LinkType, dbo.TblLink_Item_To_StoreH.UserID, dbo.TblUsers.UserName,"
 StrSQL = StrSQL & "                     dbo.TblLink_Item_To_StoreH.BranchID, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblLink_Item_To_StoreH.RecordeDate,"
StrSQL = StrSQL & "                       dbo.TblLink_Item_To_StoreH.Remarks, dbo.TblLink_Item_To_StoreH.Selected, dbo.TblLink_Item_To_StoreH.Posted, dbo.TblLink_Item_To_Store_Details2.Ind AS Ind2,"
StrSQL = StrSQL & "                       dbo.TblLink_Item_To_Store_Details2.StoreID, dbo.TblStore.StoreName, dbo.TblStore.StoreNamee, dbo.TblLink_Item_To_Store_Details2.ItemID,"
StrSQL = StrSQL & "                       dbo.TblItems.ItemName, dbo.TblItems.ItemNamee, dbo.TblLink_Item_To_Store_Details2.LinkType AS linktype2, dbo.TblLink_Item_To_Store_Details2.GroupID,"
 StrSQL = StrSQL & "                      dbo.Groups.GroupName , dbo.Groups.GroupNamee"
StrSQL = StrSQL & "  FROM         dbo.TblLink_Item_To_StoreH LEFT OUTER JOIN"
 StrSQL = StrSQL & "                      dbo.TblBranchesData ON dbo.TblLink_Item_To_StoreH.BranchID = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
  StrSQL = StrSQL & "                     dbo.TblUsers ON dbo.TblLink_Item_To_StoreH.UserID = dbo.TblUsers.UserID RIGHT OUTER JOIN"
  StrSQL = StrSQL & "                     dbo.Groups RIGHT OUTER JOIN"
  StrSQL = StrSQL & "                     dbo.TblLink_Item_To_Store_Details2 ON dbo.Groups.GroupID = dbo.TblLink_Item_To_Store_Details2.GroupID LEFT OUTER JOIN"
   StrSQL = StrSQL & "                    dbo.TblItems ON dbo.TblLink_Item_To_Store_Details2.ItemID = dbo.TblItems.ItemID LEFT OUTER JOIN"
  StrSQL = StrSQL & "                     dbo.TblStore ON dbo.TblLink_Item_To_Store_Details2.StoreID = dbo.TblStore.StoreID ON"
     StrSQL = StrSQL & "                  dbo.TblLink_Item_To_StoreH.Ind = dbo.TblLink_Item_To_Store_Details2.Ind"
'StrSQL = StrSQL & " Where (dbo.TblLink_Item_To_StoreH.Ind = 1)"
 
    BolBegine = False
    StrWhere = ""

    If val(Me.TxtIDFrom.text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND  dbo.TblLink_Item_To_StoreH.Ind >=" & val(Me.TxtIDFrom.text) & ""
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TblLink_Item_To_StoreH.Ind >=" & val(Me.TxtIDFrom.text) & ""
        End If
    End If

    If val(Me.TxtIDTO.text) <> 0 Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND  dbo.TblLink_Item_To_StoreH.Ind <=" & val(Me.TxtIDTO.text) & ""
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TblLink_Item_To_StoreH.Ind <=" & val(Me.TxtIDTO.text) & ""
        End If
    End If

    If Me.DCItem.BoundText <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND dbo.TblLink_Item_To_Store_Details2.ItemID=" & Me.DCItem.BoundText & ""
        Else
            BolBegine = True
            StrWhere = " Where dbo.TblLink_Item_To_Store_Details2.ItemID=" & Me.DCItem.BoundText & ""
        End If
    End If

    If Me.DCStore.BoundText <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND  dbo.TblLink_Item_To_Store_Details2.StoreID=" & Me.DCStore.BoundText & ""
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TblLink_Item_To_Store_Details2.StoreID=" & Me.DCStore.BoundText & ""
        End If
    End If
        If Me.TxtItemCode.text <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND  dbo.TblLink_Item_To_Store_Details2.ItemID=" & Me.TxtItemCode.text & ""
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TblLink_Item_To_Store_Details2.ItemID=" & Me.TxtItemCode.text & ""
        End If
    End If
 If Me.DCGroup.BoundText <> "" Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND  dbo.TblLink_Item_To_Store_Details2.GroupID=" & Me.DCGroup.BoundText & ""
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TblLink_Item_To_Store_Details2.GroupID=" & Me.DCGroup.BoundText & ""
        End If
    End If
    If Not IsNull(Me.DtpDateFrom.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND  dbo.TblLink_Item_To_StoreH.RecordeDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TblLink_Item_To_StoreH.RecordeDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
        End If
    End If

    If Not IsNull(Me.DtpDateTo.value) Then
        If BolBegine = True Then
            StrWhere = StrWhere & " AND  dbo.TblLink_Item_To_StoreH.RecordeDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        Else
            BolBegine = True
            StrWhere = " Where  dbo.TblLink_Item_To_StoreH.RecordeDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
        End If
    End If

    '-----------------------------------

    StrSQL = StrSQL & StrWhere
    StrSQL = StrSQL & " Order By dbo.TblLink_Item_To_StoreH.Ind "
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
            Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
            Me.lbl(10).Caption = "Search Results=0"
        End If

        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «·»ÕÀ"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    Else

        With Me.Fg
            .Clear flexClearScrollable, flexClearEverything
            .Rows = .FixedRows
            .Rows = rs.RecordCount + .FixedRows

            If SystemOptions.UserInterface = ArabicInterface Then
                Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
                Me.lbl(10).Caption = "Search Results=" & rs.RecordCount
            End If

            rs.MoveFirst
        
            For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                .TextMatrix(i, .ColIndex("ID")) = IIf(IsNull(rs("Ind").value), "", rs("Ind").value)
                        
                If Not (IsNull(rs("RecordeDate").value)) Then
                    .TextMatrix(i, .ColIndex("RecordDate")) = Format(rs("RecordeDate").value, "yyyy/M/d")
                End If
            
                .TextMatrix(i, .ColIndex("StoreName")) = IIf(IsNull(rs("StoreName").value), "", rs("StoreName").value)
                .TextMatrix(i, .ColIndex("GroupName")) = IIf(IsNull(rs("GroupName").value), "", rs("GroupName").value)
                .TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(rs("ItemName").value), "", rs("ItemName").value)
                If rs("linktype2").value = 0 Then
                .TextMatrix(i, .ColIndex("LinkType")) = "€Ì— „— »ÿ"
                Else
                 .TextMatrix(i, .ColIndex("LinkType")) = "„— »ÿ"
                 End If
                rs.MoveNext
            Next i

            .AutoSize 0, .Cols - 1, False
          '  Me.lbl(1).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("AdvanceValue"), .Rows - 1, .ColIndex("AdvanceValue"))
        End With

    End If

End Sub

Private Sub TxtIDFrom_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtIDFrom.text, 1)
End Sub

Private Sub TxtIDTO_KeyPress(KeyAscii As Integer)
    KeyAscii = KeyAscii_Num(KeyAscii, Me.TxtIDTO.text, 1)
End Sub

