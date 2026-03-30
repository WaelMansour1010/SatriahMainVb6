VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Begin VB.Form FrmItemSearch1 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«·»ÕÀ ⁄‰ ’‰›"
   ClientHeight    =   5685
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12975
   Icon            =   "FrmItemSearch1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5685
   ScaleWidth      =   12975
   Begin VB.TextBox TxtPartNo 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   2640
      RightToLeft     =   -1  'True
      TabIndex        =   36
      Top             =   3600
      Width           =   1545
   End
   Begin VB.TextBox TxtbarCodeNO 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   5640
      RightToLeft     =   -1  'True
      TabIndex        =   34
      Top             =   3600
      Width           =   2985
   End
   Begin VB.TextBox TxtFillData 
      Alignment       =   1  'Right Justify
      BackColor       =   &H000000FF&
      Height          =   285
      Left            =   0
      TabIndex        =   33
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ComboBox CboItemCodeSearch 
      Height          =   315
      Left            =   2670
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   2910
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
      Top             =   7110
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
      Top             =   7500
      Width           =   6495
   End
   Begin VB.ComboBox CboArchive 
      Height          =   315
      Left            =   2670
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   4380
      Width           =   1335
   End
   Begin VB.ComboBox CboGuar 
      Height          =   315
      Left            =   5160
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   4410
      Width           =   1305
   End
   Begin VB.ComboBox CboNameSearch 
      Height          =   315
      Left            =   2670
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   3300
      Width           =   1515
   End
   Begin VB.ComboBox CboAttachedItem 
      Height          =   315
      Left            =   2670
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   4740
      Width           =   1335
   End
   Begin VB.ComboBox CboAssbliedItem 
      Height          =   315
      Left            =   5160
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   4740
      Width           =   1305
   End
   Begin VB.ComboBox CboItemType 
      Height          =   315
      Left            =   7380
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   4410
      Width           =   1215
   End
   Begin VB.TextBox TxtItemID 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   7860
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   2910
      Width           =   735
   End
   Begin VB.ComboBox CboSerial 
      Height          =   315
      Left            =   2670
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   4020
      Width           =   1515
   End
   Begin VB.TextBox TxtItemName 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   5610
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   3300
      Width           =   2985
   End
   Begin VB.TextBox XPTxtItemCode 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   5610
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   2895
      Width           =   1395
   End
   Begin VSFlex8UCtl.VSFlexGrid Fg 
      Height          =   2745
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   12915
      _cx             =   22781
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
      FormatString    =   $"FrmItemSearch1.frx":030A
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
      Left            =   5430
      TabIndex        =   13
      Top             =   5235
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
      Left            =   4410
      TabIndex        =   14
      Top             =   5235
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
      Left            =   3510
      TabIndex        =   15
      Top             =   5235
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
      Left            =   5610
      TabIndex        =   6
      Top             =   4050
      Width           =   2985
      _ExtentX        =   5265
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·ﬁÿ⁄Â/«·„ÊœÌ·"
      Height          =   375
      Left            =   4200
      RightToLeft     =   -1  'True
      TabIndex        =   37
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·»«—ﬂÊœ"
      Height          =   345
      Index           =   14
      Left            =   8640
      RightToLeft     =   -1  'True
      TabIndex        =   35
      Top             =   3720
      Width           =   915
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   255
      Index           =   13
      Left            =   8040
      RightToLeft     =   -1  'True
      TabIndex        =   32
      Top             =   5040
      Width           =   1335
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   255
      Index           =   12
      Left            =   6840
      RightToLeft     =   -1  'True
      TabIndex        =   31
      Top             =   5040
      Width           =   1335
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "„Ã«· «·»ÕÀ"
      Height          =   345
      Index           =   11
      Left            =   4590
      RightToLeft     =   -1  'True
      TabIndex        =   30
      Top             =   2910
      Width           =   975
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·√—‘Ì›"
      Height          =   285
      Index           =   10
      Left            =   4410
      RightToLeft     =   -1  'True
      TabIndex        =   27
      Top             =   4410
      Width           =   705
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«·÷„«‰"
      Height          =   285
      Index           =   9
      Left            =   6480
      RightToLeft     =   -1  'True
      TabIndex        =   26
      Top             =   4410
      Width           =   885
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "„Ã«· «·»ÕÀ"
      Height          =   345
      Index           =   8
      Left            =   4590
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   3300
      Width           =   975
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "„·Õﬁ"
      Height          =   315
      Index           =   7
      Left            =   4410
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   4740
      Width           =   705
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   " Ã„Ì⁄"
      Height          =   285
      Index           =   6
      Left            =   6480
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   4740
      Width           =   915
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "‰Ê⁄ «·’‰›"
      Height          =   285
      Index           =   5
      Left            =   8640
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   4440
      Width           =   915
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «·’‰›"
      Height          =   345
      Index           =   4
      Left            =   8640
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   2910
      Width           =   915
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "‰Ÿ«„ «·”Ì—Ì«·"
      Height          =   315
      Index           =   2
      Left            =   4590
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   4050
      Width           =   975
   End
   Begin VB.Label LblRes 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      ForeColor       =   &H000000C0&
      Height          =   315
      Left            =   7620
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   4770
      Width           =   1905
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·’‰›"
      Height          =   345
      Index           =   1
      Left            =   8640
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   3300
      Width           =   915
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "ﬂÊœ «·’‰›"
      Height          =   345
      Index           =   0
      Left            =   7020
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   2910
      Width           =   795
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "«”„ «·„Ã„Ê⁄…"
      Height          =   285
      Index           =   3
      Left            =   8640
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   4050
      Width           =   915
   End
End
Attribute VB_Name = "FrmItemSearch1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim cSearchDcbo As clsDCboSearch

Private m_DcboItems As DataCombo

Private m_RetrunType As Integer

Public WithEvents FG1 As VSFlex8UCtl.VSFlexGrid
Attribute FG1.VB_VarHelpID = -1

 Public WithEvents NewGrid As VSFlex8UCtl.VSFlexGrid
Attribute NewGrid.VB_VarHelpID = -1
'Public WithEvents NewGrid As ClsGrid
 
Public LngRow As Long

Public LngCol As Long

Private Sub Cmd_Click(index As Integer)

    On Error GoTo ErrTrap
    Dim Msg As String

    Select Case index

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
                FG.Clear flexClearScrollable, flexClearEverything
                FG.rows = 2

                If SystemOptions.UserInterface = ArabicInterface Then
                    Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
                    MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
                Else
                    Msg = "NO Search Results Found...!!!"
                    MsgBox Msg, vbOKOnly + vbExclamation, App.Title
                End If

                Exit Sub
            End If

            Retrive
            FG.SetFocus

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
        MsgBox Msg, vbOKOnly + vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    End If

End Sub

Private Sub fg_Click()

    'On Error GoTo ErrTrap
    If Not FG.TextMatrix(FG.row, 1) = "" Then
    
        Dim Msg As String
        Dim ExpiryDate As Date
        Dim Askinterval As String

        If Not NewGrid Is Nothing Then
            '
            'FG1.TextMatrix(LngRow, LngCol) = val(Fg.TextMatrix(Fg.Row, 1))  'Trim$(Me.TxtComment.text)
 
            If Me.NewGrid.ColIndex("Code") <> -1 Then
                NewGrid.TextMatrix(LngRow, NewGrid.ColIndex("Code")) = val(FG.TextMatrix(FG.row, 1))
                NewGrid.TextMatrix(LngRow, NewGrid.ColIndex("Name")) = val(FG.TextMatrix(FG.row, 1))
                'Set NewGrid.TxtFillData = TxtFillData
                '    NewGrid.FillGrid
              '  NewGrid.Grid_AfterEdit LngRow, 1
                
                Dim StrSQL As String
                Dim RsUnitData As New ADODB.Recordset
                StrSQL = " SELECT TblItemsUnits.ItemID, TblItemsUnits.UnitID, TblUnites.UnitName," & "TblItemsUnits.UnitFactor, TblItemsUnits.SecOrder, TblItemsUnits.DefaultUnit," & "TblItemsUnits.UnitSalesPrice, TblItemsUnits.UnitPurPrice, TblItemsUnits.FactorByDefaultUnit," & "TblItemsUnits.FactorBySmallUnit "
                StrSQL = StrSQL + " FROM TblItemsUnits INNER JOIN TblUnites ON TblItemsUnits.UnitID =" & "TblUnites.UnitID"
                StrSQL = StrSQL + " Where TblItemsUnits.ItemID=" & val(FG.TextMatrix(FG.row, 1))
                StrSQL = StrSQL + " AND DefaultUnit=1"
                RsUnitData.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                                    If Not (RsUnitData.BOF Or RsUnitData.EOF) Then
                    
                                                        With NewGrid
                                                    
                                                                            If .ColIndex("Code") <> -1 Then
                                                                                .cell(flexcpData, LngRow, .ColIndex("UnitID")) = RsUnitData("UnitID").value
                                                                                .TextMatrix(LngRow, .ColIndex("UnitID")) = RsUnitData("UnitName").value
                                                                            End If
                                                                   
                                                        End With
                    
                                    End If

                RsUnitData.Close
                Set RsUnitData = Nothing
                Dim LngItemID As Long
                Dim LngUnitID As Long
 
            End If

            If Me.FG.ColIndex("Name") <> -1 Then
                '   FG1.TextMatrix(LngRow, FG1.ColIndex("Name")) = val(Fg.TextMatrix(Fg.Row, 2))

            End If
   'NewGrid.Grid_AfterEdit
  
            Unload Me
        End If
    
        If Me.RetrunType = 0 Then
        '    FrmItems.Retrive val(FG.TextMatrix(FG.row, 1))
        ElseIf Me.RetrunType = 1 Then
         
        ElseIf Me.RetrunType = 2 Then
    
            FrmShowPrice.DCboItemsCode.BoundText = val(FG.TextMatrix(FG.row, 1))
        ElseIf Me.RetrunType = 3 Then
    
            FrmBillBuy.DCboItemsCode.BoundText = val(FG.TextMatrix(FG.row, 1))
    
        ElseIf Me.RetrunType = 4 Then
    
            FrmInpout.DCboItemsCode.BoundText = val(FG.TextMatrix(FG.row, 1))
    
        ElseIf Me.RetrunType = 5 Then
    
            frmsalebill.DCboItemsCode.BoundText = val(FG.TextMatrix(FG.row, 1))
    
        ElseIf Me.RetrunType = 6 Then
    
            FrmOut.DCboItemsCode.BoundText = val(FG.TextMatrix(FG.row, 1))
    
        ElseIf Me.RetrunType = 7 Then
    
            FrmDestruction.DCboItemsCode.BoundText = val(FG.TextMatrix(FG.row, 1))
    
        ElseIf Me.RetrunType = 8 Then
    
            FrmOpeningBalance.DCboItemsCode.BoundText = val(FG.TextMatrix(FG.row, 1))
    
        ElseIf Me.RetrunType = 9 Then
    
            FrmReturnSalling.DCboItemsCode.BoundText = val(FG.TextMatrix(FG.row, 1))
    
        ElseIf Me.RetrunType = 10 Then
    
            FrmMoving.DCboItemsCode.BoundText = val(FG.TextMatrix(FG.row, 1))
        ElseIf Me.RetrunType = 11 Then
 
            FrmSallingPlan.dcitems.BoundText = val(FG.TextMatrix(FG.row, 1))
        ElseIf Me.RetrunType = 12 Then
 
            FrmNewGard.DCboItemsCode = val(FG.TextMatrix(FG.row, 1))
            
         ElseIf Me.RetrunType = 66 Then
 
         '   frmsalebill6.FG.TextMatrix(LngRow, frmsalebill6.FG.ColIndex("Code")) = val(FG.TextMatrix(FG.row, 0))
         '   frmsalebill6.FG.TextMatrix(LngRow, frmsalebill6.FG.ColIndex("Name")) = (FG.TextMatrix(FG.row, 0))
        
        ElseIf Me.RetrunType = 13 Then
 
            FrmOutProductionOrder.DCboItemsCode = val(FG.TextMatrix(FG.row, 1))
        
        ElseIf Me.RetrunType = 14 Then
 
            FrmInpoutWorkOrder.DCboItemsCode = val(FG.TextMatrix(FG.row, 1))
       
        ElseIf Me.RetrunType = 15 Then
 
            FrmProductionOrder1.DCboItemsCode = val(FG.TextMatrix(FG.row, 1))
        
        End If
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub AddItemPrice(LngItemID As Long, _
                         Col As Long, _
                         row As Long, _
                         Optional LngUnitID As Long)

    With NewGrid.Grid
     
        If Col = .ColIndex("unitId") Or Col = .ColIndex("Code") Or Col = .ColIndex("Name") Then
            .TextMatrix(row, .ColIndex("Price")) = GetItemPrice(LngItemID, val(.TextMatrix(row, .ColIndex("Count"))), LngUnitID)
        End If
  
    End With

End Sub

Private Sub Retrive()
    Dim Num As Integer
    On Error GoTo ErrTrap
    FG.Clear flexClearScrollable, flexClearEverything

    If Not (rs.EOF Or rs.BOF) Then
        FG.rows = rs.RecordCount + 1

        For Num = 1 To rs.RecordCount

            With FG
                .TextMatrix(Num, .ColIndex("NumIndex")) = Num
                .TextMatrix(Num, .ColIndex("ItemNum")) = IIf(IsNull(rs("ItemID").value), "", val(rs("ItemID").value))
                .TextMatrix(Num, .ColIndex("KindCode")) = IIf(IsNull(rs("ItemCode").value), "", Trim(rs("ItemCode").value))
'                .TextMatrix(Num, .ColIndex("KindNme")) = IIf(IsNull(rs("ItemName").value), "", Trim(rs("ItemName").value))
        .TextMatrix(Num, .ColIndex("barCodeNO")) = IIf(IsNull(rs("barCodeNO").value), "", Trim(rs("barCodeNO").value))
        
         If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(Num, .ColIndex("KindNme")) = IIf(IsNull(rs("ItemName").value), "", Trim(rs("ItemName").value))
        Else
        .TextMatrix(Num, .ColIndex("KindNme")) = IIf(IsNull(rs("ItemNamee").value), "", Trim(rs("ItemNamee").value))
        
        End If
        
                If rs("ItemType").value = 0 Then
                    If SystemOptions.UserInterface = ArabicInterface Then
                        .TextMatrix(Num, .ColIndex("ItemType")) = "”·⁄…"
                    Else
                        .TextMatrix(Num, .ColIndex("ItemType")) = "Goods"
                    End If

                Else

                    If SystemOptions.UserInterface = ArabicInterface Then
                        .TextMatrix(Num, .ColIndex("ItemType")) = "Œœ„…"
                    Else
                        .TextMatrix(Num, .ColIndex("ItemType")) = "Service"
                    End If
                End If

                If rs("HaveSerial").value Eqv True Then
                    .TextMatrix(Num, .ColIndex("HaveSerial")) = 1
                Else
                    .TextMatrix(Num, .ColIndex("HaveSerial")) = 0
                End If

                If rs("HaveGuarantee").value Eqv True Then
                    .TextMatrix(Num, .ColIndex("HaveGuarantee")) = 1
                Else
                    .TextMatrix(Num, .ColIndex("HaveGuarantee")) = 0
                End If

                If rs("IsArchive").value Eqv True Then
                    .TextMatrix(Num, .ColIndex("IsArchive")) = 1
                Else
                    .TextMatrix(Num, .ColIndex("IsArchive")) = 0
                End If
            
                If rs("AssbliedItem").value Eqv True Then
                    .TextMatrix(Num, .ColIndex("AssbliedItem")) = 1
                Else
                    .TextMatrix(Num, .ColIndex("AssbliedItem")) = 0
                End If
                        
                If rs("RelatedItem").value Eqv True Then
                    .TextMatrix(Num, .ColIndex("RelatedItem")) = 1
                Else
                    .TextMatrix(Num, .ColIndex("RelatedItem")) = 0
                End If
            
            End With

            rs.MoveNext
        Next Num

        FG.AutoSize 0, FG.Cols - 1, False
    End If

    Exit Sub
ErrTrap:
End Sub

Private Sub Fg_DblClick()
    fg_Click
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
    Dcombos.GetItemSGroups Me.DCboGroupName
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
            .AddItem "·Â «’‰«› „·Õﬁ…"
            .AddItem "·Ì” ·Â «’‰«› „·Õﬁ…"
            .AddItem "«·ﬂ·"
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
            .AddItem "ALL"
        End With

    End If

    CenterForm Me

    FormPostion Me, GetPostion
    FG.WallPaper = BG.SearchWallpaper
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

    StrSQL = "Select * From TblItems "
    StrSQL = StrSQL + " Where ItemID <> 0 "


  If SystemOptions.WorkWithLINKEDiActivity = True Then
    StrSQL = StrSQL & "  and dbo.TblItems.GroupID in(   "
     StrSQL = StrSQL & " select GroupID from fullgroups ()  )"
 End If


    If val(Me.TxtItemID.text) <> 0 Then
        StrSQL = StrSQL + " AND ItemID =" & val(Me.TxtItemID.text)
    End If

    If XPTxtItemCode.text <> "" Then
        If Me.CboItemCodeSearch.ListIndex = 0 Then
            StrWhere = StrWhere + " and ItemCode ='" & Trim(XPTxtItemCode.text) & "'"
        ElseIf Me.CboItemCodeSearch.ListIndex = 1 Then
            StrWhere = StrWhere + " and ItemCode like '" & Trim(XPTxtItemCode.text) & "%'"
        ElseIf Me.CboItemCodeSearch.ListIndex = 2 Then
            StrWhere = StrWhere + " and ItemCode like '%" & Trim(XPTxtItemCode.text) & "'"
        ElseIf Me.CboItemCodeSearch.ListIndex = 3 Then
            StrWhere = StrWhere + " and ItemCode like '%" & Trim(XPTxtItemCode.text) & "%'"
        ElseIf Me.CboItemCodeSearch.ListIndex = -1 Then
            StrWhere = StrWhere + " and ItemCode like '%" & Trim(XPTxtItemCode.text) & "%'"
        End If
    End If



 If TxtbarCodeNO.text <> "" Then
        If Me.CboItemCodeSearch.ListIndex = 0 Then
            StrWhere = StrWhere + " and barCodeNO ='" & Trim(TxtbarCodeNO.text) & "'"
        ElseIf Me.CboItemCodeSearch.ListIndex = 1 Then
            StrWhere = StrWhere + " and barCodeNO like '" & Trim(TxtbarCodeNO.text) & "%'"
        ElseIf Me.CboItemCodeSearch.ListIndex = 2 Then
            StrWhere = StrWhere + " and barCodeNO like '%" & Trim(TxtbarCodeNO.text) & "'"
        ElseIf Me.CboItemCodeSearch.ListIndex = 3 Then
            StrWhere = StrWhere + " and barCodeNO like '%" & Trim(TxtbarCodeNO.text) & "%'"
        ElseIf Me.CboItemCodeSearch.ListIndex = -1 Then
            StrWhere = StrWhere + " and barCodeNO like '%" & Trim(TxtbarCodeNO.text) & "%'"
        End If
    End If
    


    If Me.CboSerial.ListIndex > 0 Then
        If Me.CboSerial.ItemData(CboSerial.ListIndex) = 1 Then
            BolHaveSerial = True
        ElseIf Me.CboSerial.ItemData(CboSerial.ListIndex) = 2 Then
            BolHaveSerial = False
        End If

        If SystemOptions.SysDataBaseType = AccessDataBase Then
            StrWhere = StrWhere + " and HaveSerial =" & BolHaveSerial & ""
        ElseIf SystemOptions.SysDataBaseType = SQLServerDataBase Then
            IntHaveSerial = IIf(BolHaveSerial = True, 1, 0)
            StrWhere = StrWhere + " and HaveSerial =" & IntHaveSerial & ""
        End If
    End If

'If SystemOptions.UserInterface = ArabicInterface Then
    If Trim(Me.txtItemName.text) <> "" Then
        If Me.CboNameSearch.ListIndex = 0 Then
            StrWhere = StrWhere + " and ItemName Like '" & Trim(Me.txtItemName.text) & "%'"
        ElseIf (Me.CboNameSearch.ListIndex = 1 Or Me.CboNameSearch.ListIndex = -1) Then
            StrWhere = StrWhere + " and ItemName like '%" & Trim(Me.txtItemName.text) & "%'"
        End If
    End If


'Else



If Trim(Me.txtItemName.text) <> "" Then
        If Me.CboNameSearch.ListIndex = 0 Then
            StrWhere = StrWhere + " or ItemNamee Like '" & Trim(Me.txtItemName.text) & "%'"
        ElseIf (Me.CboNameSearch.ListIndex = 1 Or Me.CboNameSearch.ListIndex = -1) Then
            StrWhere = StrWhere + " or ItemNamee like '%" & Trim(Me.txtItemName.text) & "%'"
        End If
    End If
 


'End If


    If Me.DCboGroupName.BoundText <> "" Then
        StrWhere = StrWhere + " and GroupID =" & Me.DCboGroupName.BoundText & ""
    End If

    If Me.CboItemType.ListIndex <> -1 Then
        If Me.CboItemType.ListIndex = 0 Then
            StrSQL = StrSQL + " AND ItemType =0"
        ElseIf Me.CboItemType.ListIndex = 1 Then
            StrSQL = StrSQL + " AND ItemType =1"
        End If
    End If

    If Me.CboGuar.ListIndex <> -1 Then
        If Me.CboGuar.ListIndex = 0 Then
            If SystemOptions.SysDataBaseType = SQLServerDataBase Then
                StrWhere = StrWhere + " and HaveGuarantee =1"
            Else
                StrWhere = StrWhere + " and HaveGuarantee =True"
            End If

        ElseIf Me.CboGuar.ListIndex = 1 Then

            If SystemOptions.SysDataBaseType = SQLServerDataBase Then
                StrWhere = StrWhere + " and HaveGuarantee =0"
            Else
                StrWhere = StrWhere + " and HaveGuarantee =False"
            End If
        End If
    End If

    If Me.CboArchive.ListIndex <> -1 Then
        If Me.CboArchive.ListIndex = 0 Then
            If SystemOptions.SysDataBaseType = SQLServerDataBase Then
                StrWhere = StrWhere + " and IsArchive =1"
            Else
                StrWhere = StrWhere + " and IsArchive =True"
            End If

        ElseIf Me.CboArchive.ListIndex = 1 Then

            If SystemOptions.SysDataBaseType = SQLServerDataBase Then
                StrWhere = StrWhere + " and IsArchive =0"
            Else
                StrWhere = StrWhere + " and IsArchive =False"
            End If
        End If
    End If

    If Me.CboAssbliedItem.ListIndex <> -1 Then
        If Me.CboAssbliedItem.ListIndex = 0 Then
            If SystemOptions.SysDataBaseType = SQLServerDataBase Then
                StrWhere = StrWhere + " and AssbliedItem =1"
            Else
                StrWhere = StrWhere + " and AssbliedItem =True"
            End If

        ElseIf Me.CboAssbliedItem.ListIndex = 1 Then

            If SystemOptions.SysDataBaseType = SQLServerDataBase Then
                StrWhere = StrWhere + " and AssbliedItem =0"
            Else
                StrWhere = StrWhere + " and AssbliedItem =False"
            End If
        End If
    End If

    If Me.CboAttachedItem.ListIndex <> -1 Then
        If Me.CboAttachedItem.ListIndex = 0 Then
            If SystemOptions.SysDataBaseType = SQLServerDataBase Then
                StrWhere = StrWhere + " and RelatedItem =1"
            Else
                StrWhere = StrWhere + " and RelatedItem =True"
            End If

        ElseIf Me.CboAttachedItem.ListIndex = 1 Then

            If SystemOptions.SysDataBaseType = SQLServerDataBase Then
                StrWhere = StrWhere + " and RelatedItem =0"
            Else
                StrWhere = StrWhere + " and RelatedItem =False"
            End If
        End If
    End If

    If TxtPartNo.text <> "" Then
        StrWhere = StrWhere + " and PartNo like '%" & Trim(TxtPartNo.text) & "%'"
    End If
    
    Build_Sql = StrSQL + StrWhere
    Exit Function
ErrTrap:
End Function

Private Sub Form_KeyDown(KeyCode As Integer, _
                         Shift As Integer)
    On Error GoTo ErrTrap

    If KeyCode = vbKeyReturn Then
        If Me.ActiveControl Is FG Then
            If Not FG.TextMatrix(FG.row, 1) = "" Then
                fg_Click
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
    lbl(14).Caption = "BarCode"
    Cmd(0).Caption = "Search"
    Cmd(1).Caption = "Clear"
    Cmd(2).Caption = "Exit"

    'OptType(0).Caption = "Start of the name"
    'OptType(1).Caption = "any part of the name"
    With Me.FG
        .TextMatrix(0, .ColIndex("barCodeNO")) = "BarCode"
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

    If Trim$(Me.txtItemName.text) = "" Then
        Me.lbl(8).Enabled = False
        Me.CboNameSearch.Enabled = False
    Else
        Me.lbl(8).Enabled = True
        Me.CboNameSearch.Enabled = True
    End If

End Sub

Private Sub txtItemName_GotFocus()
If SystemOptions.UserInterface = EnglishInterface Then
    SwitchKeyboardLang LANG_ENGLISH
Else
    SwitchKeyboardLang LANG_ARABIC
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

