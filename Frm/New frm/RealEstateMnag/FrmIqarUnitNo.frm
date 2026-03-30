VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Begin VB.Form FrmIqarUnitNo 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ÊÕœ«  «·⁄ﬁ«—"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4635
   Icon            =   "FrmIqarUnitNo.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   4425
   ScaleWidth      =   4635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame12 
      Caption         =   "«·„’—Ê›« "
      Height          =   3615
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   360
      Width           =   4575
      Begin VB.TextBox TxtTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   3240
         Width           =   1425
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid3 
         Height          =   2820
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   4305
         _cx             =   7594
         _cy             =   4974
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
         BackColorFixed  =   -2147483633
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
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmIqarUnitNo.frx":038A
         ScrollTrack     =   0   'False
         ScrollBars      =   2
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
         ExplorerBar     =   0
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
         Height          =   270
         Index           =   8
         Left            =   3600
         TabIndex        =   6
         Top             =   3240
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   476
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   "Õ–›"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmIqarUnitNo.frx":0417
         DrawFocusRectangle=   0   'False
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·«Ã„«·Ì"
         Height          =   210
         Index           =   56
         Left            =   1680
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   3240
         Width           =   630
      End
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   0
      Left            =   1650
      TabIndex        =   0
      Top             =   3960
      Width           =   765
      _ExtentX        =   1349
      _ExtentY        =   661
      ButtonPositionImage=   1
      Caption         =   "Õ›Ÿ"
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
      TabIndex        =   1
      Top             =   3960
      Visible         =   0   'False
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
      TabIndex        =   2
      Top             =   3960
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      ButtonPositionImage=   1
      Caption         =   ""
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
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "‘«‘… ÊÕœ«  «·⁄ﬁ«—"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   -90
      TabIndex        =   3
      Top             =   0
      Width           =   4755
   End
End
Attribute VB_Name = "FrmIqarUnitNo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim DCboSearch As clsDCboSearch
Public TypIndex As Integer
Dim isExit As Boolean
Private Sub Cmd_Click(Index As Integer)

    Select Case Index

        Case 0
    save
    If isExit = True Then Exit Sub
    
    Unload Me
    isExit = False
       Case 24
       Case 8
            DeleteFgRowAther
    End Select

End Sub
Private Sub Retrive()
    Dim i As Integer
    Dim StrSQL As String
    Dim astrSplit2tems2() As String
    Dim astrSplitItems() As String
    Dim ItemName As String
    Dim j As Integer
    Dim st As String
    Dim nElements As Integer
        With Me.VSFlexGrid3
        st = ""
        If TypIndex = 2 Then
         st = FrmExpenses301.VSFlexGrid1.TextMatrix(FrmExpenses301.LngRow, FrmExpenses301.VSFlexGrid1.ColIndex("StrUnit"))
        ElseIf TypIndex = 3 Then
          st = FrmAccEditJournal4.Fg_Journal.TextMatrix(FrmAccEditJournal4.LngRow, FrmAccEditJournal4.Fg_Journal.ColIndex("StrUnit"))
        Else
          st = RsExpenses.Fg_Journal.TextMatrix(RsExpenses.LngRow, RsExpenses.Fg_Journal.ColIndex("StrUnit"))
        End If
   If st <> "" Then
          st = Trim(st)
          astrSplitItems = Split(st, "@")
          nElements = UBound(astrSplitItems) - LBound(astrSplitItems)
            .rows = .FixedRows + nElements

            For j = 0 To nElements - 1
            astrSplit2tems2 = Split(astrSplitItems(j), "#")
            i = j + 1
            StrSQL = Replace(Replace(astrSplit2tems2(0), CHR(10), ""), CHR(13), "")
            StrSQL = Trim(StrSQL)
            .TextMatrix(i, .ColIndex("Name")) = StrSQL
            .TextMatrix(i, .ColIndex("id")) = val(astrSplit2tems2(1))
            .TextMatrix(i, .ColIndex("Valu")) = val(astrSplit2tems2(2))
          Next j
      End If
        End With
ReLineGrid
End Sub

Sub save()
Dim str As String
Dim i As Integer
Dim str2 As String
str = ""
str2 = ""
With Me.VSFlexGrid3
For i = 1 To .rows - 1
 If .TextMatrix(i, .ColIndex("Name")) <> "" Then
  str = str & Trim(.TextMatrix(i, .ColIndex("Name"))) & "#"
  str = str & Trim(.TextMatrix(i, .ColIndex("id"))) & "#"
  str = str & Trim(.TextMatrix(i, .ColIndex("Valu"))) & "#"
  str = str & Trim("@")
  str = str & CHR(13)
  str = Trim(str)
 End If
Next
str2 = ""
For i = 1 To .rows - 1
 If .TextMatrix(i, .ColIndex("Name")) <> "" Then
 str2 = str2 & Trim(.TextMatrix(i, .ColIndex("Name")))
 If i <> .rows - 1 Then
  str2 = str2 & " " & ","
 End If
  str2 = Trim(str2)
 End If
Next
If TypIndex = 2 Then


If val(Me.txtTotal.text) <> val(FrmExpenses301.VSFlexGrid1.TextMatrix(FrmExpenses301.LngRow, FrmExpenses301.VSFlexGrid1.ColIndex("value"))) Then
   isExit = True
    MsgBox "ÌÃ» «‰  ﬂÊ‰ ﬁÌ„… «·ÊÕœ«  „”«ÊÌ… ·‰›” «·ﬁÌ„… ›Ï «·”ÿ— «Ï «‰ «·ﬁÌ„… ÌÃ» «‰  ﬂ‰ = " & val(FrmExpenses301.VSFlexGrid1.TextMatrix(FrmExpenses301.LngRow, FrmExpenses301.VSFlexGrid1.ColIndex("value")))
    Exit Sub
End If
isExit = False
FrmExpenses301.VSFlexGrid1.TextMatrix(FrmExpenses301.LngRow, FrmExpenses301.VSFlexGrid1.ColIndex("StrUnit")) = str
FrmExpenses301.VSFlexGrid1.TextMatrix(FrmExpenses301.LngRow, FrmExpenses301.VSFlexGrid1.ColIndex("Unitss")) = str2
FrmExpenses301.VSFlexGrid1.TextMatrix(FrmExpenses301.LngRow, FrmExpenses301.VSFlexGrid1.ColIndex("Value")) = val(Me.txtTotal.text)

ElseIf TypIndex = 3 Then


    Dim mValue As Double
    If val(FrmAccEditJournal4.Fg_Journal.TextMatrix(FrmAccEditJournal4.LngRow, FrmAccEditJournal4.Fg_Journal.ColIndex("DebitValue"))) <> 0 Then
        mValue = val(FrmAccEditJournal4.Fg_Journal.TextMatrix(FrmAccEditJournal4.LngRow, FrmAccEditJournal4.Fg_Journal.ColIndex("DebitValue")))
    ElseIf val(FrmAccEditJournal4.Fg_Journal.TextMatrix(FrmAccEditJournal4.LngRow, FrmAccEditJournal4.Fg_Journal.ColIndex("CreditValue"))) <> 0 Then
        mValue = val(FrmAccEditJournal4.Fg_Journal.TextMatrix(FrmAccEditJournal4.LngRow, FrmAccEditJournal4.Fg_Journal.ColIndex("CreditValue")))
    End If
    
If val(Me.txtTotal.text) <> mValue Then
   isExit = True
    MsgBox "ÌÃ» «‰  ﬂÊ‰ ﬁÌ„… «·ÊÕœ«  „”«ÊÌ… ·‰›” «·ﬁÌ„… ›Ï «·”ÿ— «Ï «‰ «·ﬁÌ„… ÌÃ» «‰  ﬂ‰ = " & mValue
    Exit Sub
End If
isExit = False
FrmAccEditJournal4.Fg_Journal.TextMatrix(FrmAccEditJournal4.LngRow, FrmAccEditJournal4.Fg_Journal.ColIndex("StrUnit")) = str
FrmAccEditJournal4.Fg_Journal.TextMatrix(FrmAccEditJournal4.LngRow, FrmAccEditJournal4.Fg_Journal.ColIndex("Unitss")) = str2
'FrmAccEditJournal4.Fg_Journal.TextMatrix(FrmAccEditJournal4.LngRow, FrmAccEditJournal4.Fg_Journal.ColIndex("Value")) = val(Me.txtTotal.text)
Else
RsExpenses.Fg_Journal.TextMatrix(RsExpenses.LngRow, RsExpenses.Fg_Journal.ColIndex("StrUnit")) = str
RsExpenses.Fg_Journal.TextMatrix(RsExpenses.LngRow, RsExpenses.Fg_Journal.ColIndex("Unitss")) = str2
RsExpenses.Fg_Journal.TextMatrix(RsExpenses.LngRow, RsExpenses.Fg_Journal.ColIndex("value")) = val(Me.txtTotal.text)
End If
End With
End Sub

Private Sub DeleteFgRowAther()
    With Me.VSFlexGrid3
        If .Row = -1 Then Exit Sub
        If .Row = 0 Then Exit Sub
        .RemoveItem .Row
     ReLineGrid
    End With
End Sub

Private Sub Form_Activate()
  PutFormOnTop Me.hWnd
End Sub

Private Sub Form_Load()
    Dim GrdBack As ClsBackGroundPic
    Dim Dcombos As ClsDataCombos
Dim Xpid As Integer
Dim rwOp As Integer
Dim rwpand As Integer
If Me.TypIndex = 2 Then
    If FrmExpenses301.TxtModFlg.text <> "R" Then
        Cmd(0).Enabled = True
        VSFlexGrid3.Enabled = True
    Else
        VSFlexGrid3.Enabled = False
        Cmd(0).Enabled = False
    End If
ElseIf Me.TypIndex = 3 Then
    If FrmAccEditJournal4.TxtModFlg.text <> "R" Then
        Cmd(0).Enabled = True
        VSFlexGrid3.Enabled = True
    Else
        VSFlexGrid3.Enabled = False
        Cmd(0).Enabled = False
    End If


Else
    If RsExpenses.TxtModFlg.text <> "R" Then
        Cmd(0).Enabled = True
        VSFlexGrid3.Enabled = True
    Else
        VSFlexGrid3.Enabled = False
        Cmd(0).Enabled = False
    End If
End If
    Set Dcombos = New ClsDataCombos
    Set DCboSearch = New clsDCboSearch
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Clear").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
    Set GrdBack = New ClsBackGroundPic
Retrive
 If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)

    FormPostion Me, SavePostion
    Set DCboSearch = Nothing
End Sub
'

Private Sub ChangeLang()
 
    Cmd(1).Caption = "Delete"
    Cmd(0).Caption = "Save"
    Cmd(2).Caption = "Exit"
    
  Me.Caption = "Distribution Expenses on Items"
  
Label5.Caption = Me.Caption
Frame12.Caption = Me.Caption

Cmd(8).Caption = "Delete"
'lbl(6).Caption = "Totals"

    With Me.VSFlexGrid3
    .TextMatrix(0, .ColIndex("LineNo")) = "LineNo"
    .TextMatrix(0, .ColIndex("Name")) = "Name"
    .TextMatrix(0, .ColIndex("Valu")) = "Value"
    End With
End Sub

Private Sub VSFlexGrid3_AfterEdit(ByVal Row As Long, _
                                  ByVal Col As Long)

    Dim StrAccountCode As String
    Dim Msg As String
    Dim rs As New ADODB.Recordset
    Dim StrSQL As String
    Dim ClsAcc As New ClsAccounts
    Dim LngRow As Long

    With VSFlexGrid3

        Select Case .ColKey(Col)
            Case "Name"
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("id"), False, True)
                .TextMatrix(Row, .ColIndex("id")) = StrAccountCode
        End Select
        If Row = .rows - 1 Then
            .rows = .rows + 1
        End If
    End With

    ReLineGrid
End Sub
Private Sub ReLineGrid(Optional current_terms As String = "")
    Dim i As Integer
    Dim IntCounter As Integer
    txtTotal.text = 0
    IntCounter = 0
    With VSFlexGrid3
        For i = .FixedRows To .rows - 1
            If .TextMatrix(i, .ColIndex("Name")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("LineNo")) = IntCounter
                txtTotal = val(txtTotal.text) + val(.TextMatrix(i, .ColIndex("Valu")))
           End If
        Next i
    End With
End Sub

Private Sub VSFlexGrid3_StartEdit(ByVal Row As Long, _
                                  ByVal Col As Long, _
                                  Cancel As Boolean)
    Dim rs As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrAccountType As String
    Dim StrComboList As String
    Dim Msg As String
    With VSFlexGrid3
        Select Case .ColKey(Col)
            Case "Name"
         If Me.TypIndex = 2 Then
             StrSQL = "select * from dbo.TblAqarDetai  where  Aqarid=" & val(FrmExpenses301.VSFlexGrid1.TextMatrix(FrmExpenses301.LngRow, FrmExpenses301.VSFlexGrid1.ColIndex("Aqarid"))) & " and unittype=" & val(FrmExpenses301.VSFlexGrid1.TextMatrix(FrmExpenses301.LngRow, FrmExpenses301.VSFlexGrid1.ColIndex("UnitType")))
            ElseIf Me.TypIndex = 3 Then
             StrSQL = "select * from dbo.TblAqarDetai  where  Aqarid=" & val(FrmAccEditJournal4.Fg_Journal.TextMatrix(FrmAccEditJournal4.LngRow, FrmAccEditJournal4.Fg_Journal.ColIndex("iqarid"))) & " and unittype=" & val(FrmAccEditJournal4.Fg_Journal.TextMatrix(FrmAccEditJournal4.LngRow, FrmAccEditJournal4.Fg_Journal.ColIndex("type")))
           Else
           StrSQL = "select * from dbo.TblAqarDetai  where  Aqarid=" & val(RsExpenses.Fg_Journal.TextMatrix(RsExpenses.LngRow, RsExpenses.Fg_Journal.ColIndex("iqarid"))) & " and unittype=" & val(RsExpenses.Fg_Journal.TextMatrix(RsExpenses.LngRow, RsExpenses.Fg_Journal.ColIndex("type")))
           
        End If
             rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                StrComboList = VSFlexGrid3.BuildComboList(rs, "unitno", "ID")
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
        End Select

    End With

End Sub
Private Sub VSFlexGrid3_BeforeEdit(ByVal Row As Long, _
                                   ByVal Col As Long, _
                                   Cancel As Boolean)

    With VSFlexGrid3

        Select Case .ColKey(Col)
        Case "LineNo"
                .ComboList = ""
            Case "Valu"
                .ComboList = ""
        End Select

    End With

End Sub

