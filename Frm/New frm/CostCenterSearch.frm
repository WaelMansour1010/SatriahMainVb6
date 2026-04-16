VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Begin VB.Form CostCenterSearch 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«·»ÕÀ ⁄‰ „—ﬂ“  ﬂ·›…"
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8340
   Icon            =   "CostCenterSearch.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   4485
   ScaleWidth      =   8340
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.ComboBox CboItemCodeSearch 
      Height          =   315
      Left            =   30
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2790
      Width           =   1515
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·’‰› «·„—«œ «·»ÕÀ ⁄‰Â ÌÕ ÊÏ ⁄·Ï Â–« «·’‰› ﬂ«Õœ „·Õﬁ« Â"
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
      Height          =   885
      Left            =   30
      RightToLeft     =   -1  'True
      TabIndex        =   29
      Top             =   6030
      Width           =   6495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·’‰› «·„—«œ «·»ÕÀ ⁄‰Â ÌÕ ÊÏ ⁄·Ï Â–« «·’‰› ﬂ«Õœ „ﬂÊ‰« Â"
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
      Height          =   885
      Left            =   30
      RightToLeft     =   -1  'True
      TabIndex        =   28
      Top             =   5100
      Width           =   6495
   End
   Begin VB.ComboBox CboArchive 
      Height          =   315
      Left            =   30
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   5700
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ComboBox CboGuar 
      Height          =   315
      Left            =   2160
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   5850
      Width           =   1305
   End
   Begin VB.ComboBox CboNameSearch 
      Height          =   315
      Left            =   30
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   3180
      Width           =   1515
   End
   Begin VB.ComboBox CboAttachedItem 
      Height          =   315
      Left            =   6150
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   3900
      Width           =   1095
   End
   Begin VB.ComboBox CboAssbliedItem 
      Height          =   315
      Left            =   2160
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   5700
      Width           =   1305
   End
   Begin VB.ComboBox CboItemType 
      Height          =   315
      Left            =   4380
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   5970
      Width           =   1215
   End
   Begin VB.TextBox TxtItemID 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   6540
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   2790
      Width           =   735
   End
   Begin VB.ComboBox CboSerial 
      Height          =   315
      ItemData        =   "CostCenterSearch.frx":030A
      Left            =   30
      List            =   "CostCenterSearch.frx":030C
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   3540
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.TextBox TxtItemName 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   4290
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   3180
      Width           =   2985
   End
   Begin VB.TextBox XPTxtItemCode 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   4290
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   2775
      Width           =   1395
   End
   Begin VSFlex8UCtl.VSFlexGrid Fg 
      Height          =   2745
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8355
      _cx             =   14737
      _cy             =   4842
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
      Cols            =   11
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"CostCenterSearch.frx":030E
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
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   0
      Left            =   1950
      TabIndex        =   13
      Top             =   4035
      Width           =   945
      _ExtentX        =   1667
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
      Left            =   930
      TabIndex        =   14
      Top             =   4035
      Width           =   975
      _ExtentX        =   1720
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
      Left            =   30
      TabIndex        =   15
      Top             =   4035
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
   Begin MSDataListLib.DataCombo DCboGroupName 
      Height          =   315
      Left            =   4290
      TabIndex        =   6
      Top             =   3570
      Width           =   2985
      _ExtentX        =   5265
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "„Ã«· «·»ÕÀ"
      Height          =   345
      Index           =   11
      Left            =   1590
      RightToLeft     =   -1  'True
      TabIndex        =   30
      Top             =   2790
      Width           =   975
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·√—‘Ì›"
      Height          =   285
      Index           =   10
      Left            =   1410
      RightToLeft     =   -1  'True
      TabIndex        =   27
      Top             =   5610
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·÷„«‰"
      Height          =   285
      Index           =   9
      Left            =   3480
      RightToLeft     =   -1  'True
      TabIndex        =   26
      Top             =   5850
      Width           =   885
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "„Ã«· «·»ÕÀ"
      Height          =   345
      Index           =   8
      Left            =   1590
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   3180
      Width           =   975
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«Ìﬁ«› «· ⁄«„·"
      Height          =   315
      Index           =   7
      Left            =   7170
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   3900
      Width           =   1065
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " Ã„Ì⁄"
      Height          =   285
      Index           =   6
      Left            =   3480
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   5700
      Width           =   915
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "‰Ê⁄ «·’‰›"
      Height          =   285
      Index           =   5
      Left            =   5640
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   5760
      Width           =   915
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·„—ﬂ“"
      Height          =   345
      Index           =   4
      Left            =   7320
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   2790
      Width           =   915
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«Ìﬁ«› «· ⁄«„·"
      Height          =   315
      Index           =   2
      Left            =   1470
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   3570
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label LblRes 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      ForeColor       =   &H000000C0&
      Height          =   315
      Left            =   6300
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   4290
      Width           =   1905
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·„—ﬂ“"
      Height          =   345
      Index           =   1
      Left            =   7320
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   3180
      Width           =   915
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "ﬂÊœ «·„—ﬂ“"
      Height          =   345
      Index           =   0
      Left            =   5700
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   2790
      Width           =   795
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·›∆…"
      Height          =   285
      Index           =   3
      Left            =   7320
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   3570
      Width           =   915
   End
End
Attribute VB_Name = "CostCenterSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim cSearchDcbo As clsDCboSearch

Private m_DcboItems As DataCombo

Private m_RetrunType As Integer

Private Sub Cmd_Click(Index As Integer)
    On Error GoTo ErrTrap

    Select Case Index

        Case 0

            If rs.State = adStateOpen Then
                rs.Close
            End If

            rs.Open Build_Sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

            If SystemOptions.UserInterface = ArabicInterface Then
                LblRes.Caption = "‰ ÌÃ… «·»ÕÀ = " & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
                LblRes.Caption = "Search Result=" & rs.RecordCount
            End If
    
            If rs.RecordCount < 1 Then
                Fg.Clear flexClearScrollable, flexClearEverything
                Fg.Rows = 2

                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
                Else
                    Msg = "NO Search Results Found...!!!"
                    MsgBox Msg, vbOKOnly + vbExclamation, App.title
                End If

                Exit Sub
            End If

            Retrive
            Fg.SetFocus

        Case 1
            clear_all Me
            Fg.Clear flexClearScrollable, flexClearEverything

        Case 2
            Unload Me
    End Select

    Exit Sub
ErrTrap:

    If Err.Number = -2147217900 Then
        Msg = Msg + "·ﬁœ  „ «œŒ«· ﬁÌ„ €Ì— ’«·Õ… " & Chr(13)
        Msg = Msg + " √ﬂœ „‰ œﬁ… „⁄«ÌÌ— «·»ÕÀ Ê√⁄œ «·„Õ«Ê·…"
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If

End Sub

Private Sub Fg_Click()
    On Error GoTo ErrTrap

    If Not Fg.TextMatrix(Fg.Row, 1) = "" Then
        If Me.RetrunType = 0 Then
            CostCenter.FiLLTXT val(Fg.TextMatrix(Fg.Row, 1))
        ElseIf Me.RetrunType = 1 Then
        
            FrmAccEditJournal.DcCostCenter.BoundText = Fg.TextMatrix(Fg.Row, 2)
        
        ElseIf Me.RetrunType = 2 Then
    
            marakes_taklefa_tawze3.DataCombo1.BoundText = Fg.TextMatrix(Fg.Row, 2)
        ElseIf Me.RetrunType = 3 Then
    
            FrmExpenses3.DcCostCenter.BoundText = Fg.TextMatrix(Fg.Row, 2)
    
        ElseIf Me.RetrunType = 4 Then
    
            FrmExpenses2.DcCostCenter.BoundText = Fg.TextMatrix(Fg.Row, 2)
    
        ElseIf Me.RetrunType = 5 Then
            FrmPayments.DcCostCenter.BoundText = Fg.TextMatrix(Fg.Row, 2)
    
        ElseIf Me.RetrunType = 6 Then
    
            FrmCashing.DcCostCenter.BoundText = Fg.TextMatrix(Fg.Row, 2)
    
        ElseIf Me.RetrunType = 7 Then
    
            FrmEmployee.DcCostCenter.BoundText = Fg.TextMatrix(Fg.Row, 2)
    
        ElseIf Me.RetrunType = 8 Then
    
            FrmOpeningBalance.DCboItemsCode.BoundText = val(Fg.TextMatrix(Fg.Row, 1))
    
        ElseIf Me.RetrunType = 178 Then
    
            FrmAccountDestribution.DCAccountDist.BoundText = Fg.TextMatrix(Fg.Row, 2)
    
        ElseIf Me.RetrunType = 9 Then
    
            FrmOut.DcCostCenter.BoundText = Fg.TextMatrix(Fg.Row, 2)
    
    ElseIf Me.RetrunType = 10 Then
    
            FrmPO6.DcCostCenter.BoundText = Fg.TextMatrix(Fg.Row, 2)
    
   ElseIf Me.RetrunType = 11 Then
    
            FrmInpout.DcCostCenter.BoundText = Fg.TextMatrix(Fg.Row, 2)
     
    
       ElseIf Me.RetrunType = 12 Then
    
            FixedAssets.DcCostCenter.BoundText = Fg.TextMatrix(Fg.Row, 2)
     
    
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Retrive()
    Dim Num As Integer
    On Error GoTo ErrTrap
    Fg.Clear flexClearScrollable, flexClearEverything

    If Not (rs.EOF Or rs.BOF) Then
        Fg.Rows = rs.RecordCount + 1

        For Num = 1 To rs.RecordCount

            With Fg
                .TextMatrix(Num, .ColIndex("NumIndex")) = Num
                .TextMatrix(Num, .ColIndex("ItemNum")) = IIf(IsNull(rs("id").value), "", val(rs("id").value))
                .TextMatrix(Num, .ColIndex("KindCode")) = IIf(IsNull(rs("CODE").value), "", Trim(rs("CODE").value))
                .TextMatrix(Num, .ColIndex("KindNme")) = IIf(IsNull(rs("account_name").value), "", Trim(rs("account_name").value))
                .TextMatrix(Num, .ColIndex("category")) = IIf(IsNull(rs("name").value), "", Trim(rs("name").value))
                        
                If rs("Block").value Eqv True Then
                    .TextMatrix(Num, .ColIndex("RelatedItem")) = 1
                Else
                    .TextMatrix(Num, .ColIndex("RelatedItem")) = 0
                End If
            
            End With

            rs.MoveNext
        Next Num

        Fg.AutoSize 0, Fg.Cols - 1, False
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Fg_DblClick()
    Fg_Click
    Unload Me
End Sub

Private Sub Form_Load()
    On Error GoTo ErrTrap
    Dim StrSQL As String
    Dim BG As New ClsBackGroundPic
    Dim Dcombos As ClsDataCombos

    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If

    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Clear").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture

    Set Dcombos = New ClsDataCombos
    Dcombos.GetCCTypes Me.DCboGroupName
    Set cSearchDcbo = New clsDCboSearch
    'cSearchDcbo.AllowWriting = False
    Set cSearchDcbo.Client = Me.DCboGroupName

    If SystemOptions.UserInterface = ArabicInterface Then

        With Me.CboItemCodeSearch
            .Clear
            .AddItem "»ÕÀ „ÿ«»ﬁ"
            .AddItem "»ÕÀ „‰ «·»œ«Ì…"
            .AddItem "»ÕÀ „‰ «·‰Â«Ì…"
            .AddItem "»ÕÀ ›Ï «Ï „ﬂ«‰"
        End With

        With Me.CboSerial
            .Clear
            .AddItem "«·ﬂ·"
            .ItemData(0) = 0
            .AddItem "·Â ”Ì—Ì«·"
            .ItemData(1) = 1
            .AddItem "·Ì” ·Â ”Ì—Ì«·"
            .ItemData(2) = 2
        End With

        With Me.CboNameSearch
            .Clear
            .AddItem "„‰ «Ê· «·√”„"
            .AddItem "›Ï «Ï Ã“¡ „‰ «·√”„"
        End With

        With Me.CboItemType
            .Clear
            .AddItem "”·⁄…"
            .AddItem "Œœ„…"
            .AddItem "«·ﬂ·"
        End With

        With Me.CboGuar
            .Clear
            .AddItem "·Â ÷„«‰"
            .AddItem "·Ì” ·Â ÷„«‰"
            .AddItem "«·ﬂ·"
        End With

        With Me.CboArchive
            .Clear
            .AddItem "›Ï «·√—‘Ì›"
            .AddItem "·Ì” ›Ï «·√—‘Ì›"
            .AddItem "«·ﬂ·"
        End With

        With Me.CboAssbliedItem
            .Clear
            .AddItem "’‰› „Ã„⁄"
            .AddItem "’‰› ⁄«œÏ"
            .AddItem "«·ﬂ·"
        End With

        With Me.CboAttachedItem
            .Clear
            .AddItem "‰⁄„"
            .AddItem "·«"
         
        End With

    ElseIf SystemOptions.UserInterface = EnglishInterface Then

        With Me.CboItemCodeSearch
            .Clear
            .AddItem "Typical Search"
            .AddItem "From The Start"
            .AddItem "From The End"
            .AddItem "Any Where"
        End With

        With Me.CboSerial
            .Clear
            .AddItem "All"
            .ItemData(0) = 0
            .AddItem "Has Serial"
            .ItemData(1) = 1
            .AddItem "NO Serial"
            .ItemData(2) = 2
        End With

        With Me.CboNameSearch
            .Clear
            .AddItem "Start Name"
            .AddItem "Any Part of Name"
        End With

        With Me.CboItemType
            .Clear
            .AddItem "Goods"
            .AddItem "Services"
            .AddItem "All"
        End With

        With Me.CboGuar
            .Clear
            .AddItem "YES"
            .AddItem "NO"
            .AddItem "ALL"
        End With

        With Me.CboArchive
            .Clear
            .AddItem "YES"
            .AddItem "NO"
            .AddItem "ALL"
        End With

        With Me.CboAssbliedItem
            .Clear
            .AddItem "YES"
            .AddItem "NO"
            .AddItem "ALL"
        End With

        With Me.CboAttachedItem
            .Clear
            .AddItem "YES"
            .AddItem "NO"
 
        End With

    End If

    CenterForm Me

    FormPostion Me, GetPostion
    Fg.WallPaper = BG.SearchWallpaper
    Set rs = New ADODB.Recordset

    Exit Sub
ErrTrap:

End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap

    If rs.State = adStateOpen Then
        rs.Close
        Set rs = Nothing
    End If

    Set cSearchDcbo = Nothing

    FormPostion Me, SavePostion
    Set m_DcboItems = Nothing
    Exit Sub
ErrTrap:
End Sub

Private Function Build_Sql()
    Dim StrSQL As String
    Dim Begin As Boolean
    Dim StrWhere As String
    Dim BolHaveSerial As Boolean
    Dim IntHaveSerial As Integer

    On Error GoTo ErrTrap

    'StrSQL = "Select * From markaas_taklefa "
    'StrSQL = StrSQL + " Where id <> 0 "

    StrSQL = "SELECT      dbo.markaas_taklefa.id,dbo.markaas_taklefa.account_no, dbo.markaas_taklefa.account_name, dbo.markaas_taklefa.Code, dbo.Marakes_taklefa_type.name , dbo.markaas_taklefa.Block" & " FROM         dbo.markaas_taklefa LEFT OUTER JOIN " & "                     dbo.Marakes_taklefa_type ON dbo.markaas_taklefa.Type = dbo.Marakes_taklefa_type.id  Where dbo.markaas_taklefa.id <> 0"

    If val(Me.TxtItemID.text) <> 0 Then
        StrSQL = StrSQL + " AND dbo.markaas_taklefa.id =" & val(Me.TxtItemID.text)
    End If

    If XPTxtItemCode.text <> "" Then
        If Me.CboItemCodeSearch.ListIndex = 0 Then
            StrWhere = StrWhere + " and CODE ='" & Trim(XPTxtItemCode.text) & "'"
        ElseIf Me.CboItemCodeSearch.ListIndex = 1 Then
            StrWhere = StrWhere + " and CODE like '" & Trim(XPTxtItemCode.text) & "%'"
        ElseIf Me.CboItemCodeSearch.ListIndex = 2 Then
            StrWhere = StrWhere + " and CODE like '%" & Trim(XPTxtItemCode.text) & "'"
        ElseIf Me.CboItemCodeSearch.ListIndex = 3 Then
            StrWhere = StrWhere + " and CODE like '%" & Trim(XPTxtItemCode.text) & "%'"
        ElseIf Me.CboItemCodeSearch.ListIndex = -1 Then
            StrWhere = StrWhere + " and CODE like '%" & Trim(XPTxtItemCode.text) & "%'"
        End If
    End If
 
    If Trim(Me.TxtItemName.text) <> "" Then
        If Me.CboNameSearch.ListIndex = 0 Then
            StrWhere = StrWhere + " and account_name Like '" & Trim(Me.TxtItemName.text) & "%'"
        ElseIf (Me.CboNameSearch.ListIndex = 1 Or Me.CboNameSearch.ListIndex = -1) Then
            StrWhere = StrWhere + " and account_name like '%" & Trim(Me.TxtItemName.text) & "%'"
        End If
    End If

    If Me.DCboGroupName.BoundText <> "" Then
        StrWhere = StrWhere + " and type =" & Me.DCboGroupName.BoundText & ""
    End If

    If Me.CboSerial.ListIndex <> -1 Then
        If Me.CboSerial.ListIndex = 0 Then
            StrSQL = StrSQL + " AND ,  Block =0"
        ElseIf Me.CboSerial.ListIndex = 1 Then
            StrSQL = StrSQL + " AND ,  Block =1"
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
        If Me.ActiveControl Is Fg Then
            If Not Fg.TextMatrix(Fg.Row, 1) = "" Then
                Fg_Click
                Unload Me
            End If

        Else
            Cmd_Click (0)
        End If
    End If

    On Error GoTo ErrTrap

    If Shift = 2 Then
        If KeyCode = vbKeyX Then
            Cmd_Click (2)
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Public Property Get DcboItems() As DataCombo
    Set DcboItems = m_DcboItems
End Property

Public Property Set DcboItems(ByVal vNewValue As DataCombo)
    Set m_DcboItems = vNewValue
End Property

Public Property Get RetrunType() As Integer
    RetrunType = m_RetrunType
End Property

Public Property Let RetrunType(ByVal vNewValue As Integer)
    m_RetrunType = vNewValue
    ' 0 = Retrun in the Items Screen
    ' 1 = Retrun in the Data Combo
End Property

Private Sub ChangeLang()
    Me.Caption = "Search For Item"
    lbl(0).Caption = "Item Code"
    lbl(1).Caption = "Item Name"
    lbl(2).Caption = "Serial Type"
    lbl(3).Caption = "Group Name"
    lbl(4).Caption = "Item ID"
    lbl(5).Caption = "Item Type"
    lbl(6).Caption = "Assembled"
    lbl(7).Caption = "Attached"
    lbl(8).Caption = "Match Type"
    lbl(9).Caption = "Guarantee"
    lbl(10).Caption = "Archives"
    lbl(11).Caption = "Match Type"

    Cmd(0).Caption = "Search"
    Cmd(1).Caption = "Clear"
    Cmd(2).Caption = "Exit"

    'OptType(0).Caption = "Start of the name"
    'OptType(1).Caption = "any part of the name"
    With Me.Fg
        .TextMatrix(0, .ColIndex("NumIndex")) = "Serial"
        .TextMatrix(0, .ColIndex("ItemNum")) = "Item ID"
        .TextMatrix(0, .ColIndex("KindCode")) = "Item Code"
        .TextMatrix(0, .ColIndex("KindNme")) = "Item Name"
        .TextMatrix(0, .ColIndex("ItemType")) = "Item Type"
        .TextMatrix(0, .ColIndex("HaveSerial")) = "Have Serial"
        .TextMatrix(0, .ColIndex("IsArchive")) = "Archive"
        .TextMatrix(0, .ColIndex("HaveGuarantee")) = "Guarantee"
        .TextMatrix(0, .ColIndex("AssbliedItem")) = "Assblied"
        .TextMatrix(0, .ColIndex("RelatedItem")) = "Attached Items"
        .AutoSize 0, .Cols - 1, False
    End With

End Sub

Private Sub TxtItemName_Change()

    If Trim$(Me.TxtItemName.text) = "" Then
        Me.lbl(8).Enabled = False
        Me.CboNameSearch.Enabled = False
    Else
        Me.lbl(8).Enabled = True
        Me.CboNameSearch.Enabled = True
    End If

End Sub

Private Sub XPTxtItemCode_Change()

    If Trim$(Me.XPTxtItemCode.text) = "" Then
        Me.lbl(11).Enabled = False
        Me.CboItemCodeSearch.Enabled = False
    Else
        Me.lbl(11).Enabled = True
        Me.CboItemCodeSearch.Enabled = True
    End If

End Sub

