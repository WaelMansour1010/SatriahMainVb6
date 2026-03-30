VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Begin VB.Form FrmExchangeOper 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "‘«‘… „’«—Ì› «·⁄„·Ì« "
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13935
   Icon            =   "FrmExchangeOper.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   4485
   ScaleWidth      =   13935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame12 
      Caption         =   "«·„’—Ê›« "
      Height          =   3615
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   360
      Width           =   13935
      Begin VB.TextBox txt_expenses_total 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5400
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   3240
         Width           =   1530
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid3 
         Height          =   2820
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   13665
         _cx             =   24104
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
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmExchangeOper.frx":038A
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
         Left            =   13080
         TabIndex        =   8
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
         ButtonImage     =   "FrmExchangeOper.frx":04FD
         DrawFocusRectangle=   0   'False
      End
      Begin VB.Label Lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«Ã„«·Ì ﬁÌ„… «·„’—Ê›« "
         Height          =   255
         Index           =   6
         Left            =   7320
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   3240
         Width           =   2535
      End
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   0
      Left            =   1650
      TabIndex        =   0
      Top             =   4080
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
      Top             =   4080
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
      Top             =   4080
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
      Caption         =   "‘«‘… „’«—Ì› «·⁄„·Ì« "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   75
      TabIndex        =   3
      Top             =   0
      Width           =   13830
   End
End
Attribute VB_Name = "FrmExchangeOper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim DCboSearch As clsDCboSearch
Dim currentterms As String


Private Sub Cmd_Click(Index As Integer)

    Select Case Index

        Case 0
    save
    Unload Me

       Case 24
     '  AddNewFgRowother
       Case 8
            DeleteFgRowAther
    End Select

End Sub
Private Sub Retrive(Optional project_id As Integer = 0, Optional Pand As Integer = 0, Optional Oper As Integer = 0)
 
    Dim StrSQL As String
    Dim AccountName As String
    Dim i As Integer
    Dim astrSplit2tems2() As String
    Dim astrSplitItems() As String
    Dim itemname As String
    Dim j As Integer
    Dim st As String
    Dim nElements As Integer
    'On Error GoTo ErrTrap
   ' VSFlexGrid3.Clear flexClearScrollable, flexClearEverything
   ' VSFlexGrid3.Rows = 2
   ' VSFlexGrid3.Enabled = True
    'txt_opr_total.text = 0
          
 '  StrSQL = " SELECT     dbo.TblExpensiveOper.ID, dbo.TblExpensiveOper.ProjectID, dbo.TblExpensiveOper.Pand, dbo.TblExpensiveOper.Opr, dbo.TblExpensiveOper.EsToal, "
 '  StrSQL = StrSQL & "                  dbo.TblExpensiveOper.Des, dbo.TblExpensiveOper.[value], REPLACE(REPLACE(dbo.TblExpensiveOper.AccountCode, CHAR(10), ''), CHAR(13), '') AS Account_Code1,"
 '  StrSQL = StrSQL & "                   dbo.ACCOUNTS.account_name , dbo.ACCOUNTS.Account_NameEng, dbo.ACCOUNTS.account_serial"
 '   StrSQL = StrSQL & "   FROM         dbo.TblExpensiveOper LEFT OUTER JOIN"
 '  StrSQL = StrSQL & "                    dbo.ACCOUNTS ON REPLACE(REPLACE(dbo.TblExpensiveOper.AccountCode, CHAR(10), ''), CHAR(13), '') = dbo.ACCOUNTS.Account_Code"
'StrSQL = StrSQL & "  Where (dbo.TblExpensiveOper.Projectid =" & project_id & ") And (dbo.TblExpensiveOper.Pand =" & Pand & ") And (dbo.TblExpensiveOper.OPR =" & Oper & ")"
'    Set RsDev = New ADODB.Recordset
'    RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

'    If Not (RsDev.BOF Or RsDev.EOF) Then
'        RsDev.MoveFirst
     
 
        With Me.VSFlexGrid3
        If Projects.VSFlexGrid2.TextMatrix(Projects.LngRow, Projects.VSFlexGrid2.ColIndex("expen")) <> "" Then
          st = Projects.VSFlexGrid2.TextMatrix(Projects.LngRow, Projects.VSFlexGrid2.ColIndex("expen"))
          st = Trim(st)
          astrSplitItems = Split(st, "@")
          nElements = UBound(astrSplitItems) - LBound(astrSplitItems)
            .Rows = .FixedRows + nElements

            For j = 0 To nElements - 1
            astrSplit2tems2 = Split(astrSplitItems(j), "#")
            i = j + 1
            StrSQL = Replace(Replace(astrSplit2tems2(0), Chr(10), ""), Chr(13), "")
            StrSQL = Trim(StrSQL)
          
                 .TextMatrix(i, .ColIndex("AccountCode")) = StrSQL
                 .TextMatrix(i, .ColIndex("EsToal")) = val(astrSplit2tems2(1))
                 .TextMatrix(i, .ColIndex("value")) = val(astrSplit2tems2(2))
                 .TextMatrix(i, .ColIndex("des")) = astrSplit2tems2(3)
                 RetriveAccuntExp StrSQL, AccountName
                 .TextMatrix(i, .ColIndex("AccountName")) = AccountName
           
Next j
           ' Me.txt_opr_total.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("total"), .Rows - 1, .ColIndex("total"))
        End If
          
        End With
ReLineGrid
    
          
  

End Sub

Sub save()
Dim str As String
Dim i As Integer
str = ""

With Me.VSFlexGrid3
For i = 1 To .Rows - 1
 If .TextMatrix(i, .ColIndex("AccountCode")) <> "" Then
 str = str & Trim(.TextMatrix(i, .ColIndex("AccountCode"))) & "#"
 str = str & Trim(.TextMatrix(i, .ColIndex("EsToal"))) & "#"
 str = str & Trim(.TextMatrix(i, .ColIndex("value"))) & "#"
 str = str & Trim(.TextMatrix(i, .ColIndex("des"))) & "#"

 str = str & Trim("@")
  str = str & Chr(13)
  str = Trim(str)
 End If
Next
Projects.VSFlexGrid2.TextMatrix(Projects.LngRow, Projects.VSFlexGrid2.ColIndex("expen")) = str
Projects.VSFlexGrid2.TextMatrix(Projects.LngRow, Projects.VSFlexGrid2.ColIndex("total_expenses")) = val(txt_expenses_total.text)

End With
End Sub

Private Sub DeleteFgRowAther()

    With Me.VSFlexGrid3

        If .Row = -1 Then Exit Sub
        If .Row = 0 Then Exit Sub
        .RemoveItem .Row
        '.AutoSize 0, .Cols - 1, False
     ReLineGrid
    End With

End Sub

Private Sub Form_Activate()
  PutFormOnTop Me.hwnd
End Sub


Sub RetriveAccuntExp(Optional AccountCode As String, Optional ByRef AccountName As String)
Dim rs1 As ADODB.Recordset
Dim sql As String
'(rs, "Account_Name", "Account_Code")
Set rs1 = New ADODB.Recordset
If SystemOptions.UserInterface = ArabicInterface Then
sql = " select * from Expenses_accounts where Account_Code ='" & AccountCode & "'"
Else
sql = " select * from Expenses_accounts_eng where Account_Code ='" & AccountCode & "'"
End If
rs1.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
If rs1.RecordCount > 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
AccountName = IIf(IsNull(rs1("Account_Name").value), "", rs1("Account_Name").value)
Else
AccountName = IIf(IsNull(rs1("Account_NameEng").value), "", rs1("Account_NameEng").value)
End If
End If
End Sub

Private Sub Form_Load()
    Dim GrdBack As ClsBackGroundPic
    Dim Dcombos As ClsDataCombos
Dim Xpid As Integer
Dim rwOp As Integer
Dim rwpand As Integer
If Projects.TxtModFlg.text <> "R" Then
Cmd(0).Enabled = True
Else
Cmd(0).Enabled = False

End If
    Set Dcombos = New ClsDataCombos
 
   'Dcombos.GetAccountingCodes Me.DcbAccount


    Set DCboSearch = New clsDCboSearch
   
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Clear").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture

 'Frame6.Visible = True
    Set GrdBack = New ClsBackGroundPic
    currentterms = Projects.VSFlexGrid2.TextMatrix(Projects.LngRow, Projects.VSFlexGrid2.ColIndex("fullcode"))
    If SystemOptions.UserInterface = ArabicInterface Then
                    Frame12.Caption = " „’«—Ì› «·⁄„·ÌÂ —ﬁ„ : " & currentterms
                Else
                    Frame12.Caption = "Expenses For Process No: " & currentterms
                End If
                
       Xpid = val(Projects.txt_project_id.text)
    rwOp = val(Projects.VSFlexGrid2.TextMatrix(Projects.LngRow, Projects.VSFlexGrid2.ColIndex("id")))
    
    rwpand = val(Projects.VSFlexGrid2.TextMatrix(Projects.LngRow, Projects.VSFlexGrid2.ColIndex("ProjectDes_ID")))

Retrive Xpid, rwpand, rwOp


'    With Me.Fg
'        Set .WallPaper = GrdBack.Picture
'        .AutoSize 0, .Cols - 1, False
'    End With
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
Lbl(6).Caption = "Totals"

    With Me.VSFlexGrid3
    
     

    .TextMatrix(0, .ColIndex("LineNo")) = "LineNo"
    .TextMatrix(0, .ColIndex("AccountName")) = "Name"
    .TextMatrix(0, .ColIndex("EsToal")) = "Est. value"
    
        .TextMatrix(0, .ColIndex("value")) = "Acual Value"
        .TextMatrix(0, .ColIndex("des")) = "des"
       
    End With
    
'Me.LblClientName.Caption = "ClientName"
'lbl(4).Caption = "From"
'lbl(3).Caption = "To"
'Cmd(24).Caption = "Add"
'Cmd(25).Caption = "Delete"

'Lbl(1).Caption = "Account Code"
'Lbl(1).Caption = "Account Name"
'Lbl(51).Caption = "Type Value"
'Lbl(41).Caption = "Value  "
'Lbl(0).Caption = "Remarks  "
'Lbl(39).Caption = "Count"
'Me.lbreg.Caption = "Date Registration"

 
  '
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
 
            Case "AccountName"
                '  .TextMatrix(Row, .ColIndex("userid")) = user_id
                        
                StrAccountCode = .ComboData
                LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("AccountCode"), False, True)
                .TextMatrix(Row, .ColIndex("AccountCode")) = StrAccountCode
                .TextMatrix(Row, .ColIndex("ExpensesID")) = get_Expenses_id(StrAccountCode)
                .TextMatrix(Row, .ColIndex("LineNo1")) = setfoxy_Line
  
            Case "value"
                Dim sgl As String
  
        End Select

        Me.txt_expenses_total.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("Value"), .Rows - 1, .ColIndex("Value"))
   
        If Row = .Rows - 1 Then
            .Rows = .Rows + 1
        End If

        ' ReLineGrid
    End With

    ReLineGrid
End Sub
Private Sub ReLineGrid(Optional current_terms As String = "")
    Dim i As Integer
    Dim IntCounter As Integer
 

txt_expenses_total.text = 0
    IntCounter = 0

    With VSFlexGrid3

        For i = .FixedRows To .Rows - 1

            If .TextMatrix(i, .ColIndex("AccountName")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("LineNo")) = IntCounter
                .TextMatrix(i, .ColIndex("FullCode")) = currentterms & "-" & .TextMatrix(i, .ColIndex("LineNo"))
  txt_expenses_total = val(txt_expenses_total.text) + val(.TextMatrix(i, .ColIndex("value")))
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

    'Case "DebitName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a1%' Or ACCOUNTS.Account_Code Like 'a3%')"
    'Case "CreditName"
    'StrAccountType = " (ACCOUNTS.Account_Code Like 'a2%' Or ACCOUNTS.Account_Code Like 'a4%')"
    With VSFlexGrid3

        Select Case .ColKey(Col)

            Case "AccountName"
            If SystemOptions.UserInterface = ArabicInterface Then
                StrSQL = "select * from Expenses_accounts"
            Else
            StrSQL = "select * from Expenses_accounts_eng"
            
            
            End If
                
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
                If SystemOptions.UserInterface = ArabicInterface Then
                StrComboList = VSFlexGrid3.BuildComboList(rs, "Account_Name", "Account_Code")
              Else
              StrComboList = VSFlexGrid3.BuildComboList(rs, "Account_NameEng", "Account_Code")
              End If
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

        If Row > .FixedRows Then
            '  If .TextMatrix(Row - 1, .ColIndex("AccountCode")) = "" Then
            '      Cancel h= True
            '  End If
        End If

        Select Case .ColKey(Col)
        Case "EsToal"
                .ComboList = ""

            Case "value"
                .ComboList = ""

            Case "des"
                .ComboList = ""
                '  Cancel = True
            
        End Select

    End With

End Sub

