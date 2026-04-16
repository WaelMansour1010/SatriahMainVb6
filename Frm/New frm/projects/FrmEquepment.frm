VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Begin VB.Form FrmEquepment 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "‘«‘… «·„⁄œ«  ··⁄„·Ì« "
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14280
   Icon            =   "FrmEquepment.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   4485
   ScaleWidth      =   14280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.Frame Frame6 
      Caption         =   "«·„⁄œ«  Ê «·√·« "
      Height          =   3615
      Left            =   -120
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   480
      Width           =   14295
      Begin VB.TextBox Text1 
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
         Left            =   4200
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   2760
         Width           =   1530
      End
      Begin VB.TextBox Text8 
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
         Left            =   7320
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   2760
         Width           =   1530
      End
      Begin VB.TextBox Text9 
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
         Left            =   5760
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   2760
         Width           =   1530
      End
      Begin VB.TextBox Text11 
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
         Left            =   2640
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   2760
         Width           =   1530
      End
      Begin VB.TextBox Text14 
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
         Left            =   1080
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   2760
         Width           =   1530
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid4 
         Height          =   2340
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   13920
         _cx             =   24553
         _cy             =   4128
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
         Cols            =   13
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmEquepment.frx":038A
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
         Left            =   12840
         TabIndex        =   11
         Top             =   2760
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
         ButtonImage     =   "FrmEquepment.frx":059A
         DrawFocusRectangle=   0   'False
      End
      Begin VB.Label Lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«Ã„«·Ì« "
         Height          =   255
         Index           =   7
         Left            =   8880
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   2760
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
      Caption         =   "‘«‘… «·„⁄œ«  ··⁄„·Ì« "
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
      Left            =   60
      TabIndex        =   3
      Top             =   0
      Width           =   14205
   End
End
Attribute VB_Name = "FrmEquepment"
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
Sub RetriveFixedAsset(Optional ExpensesID As Double = 0, Optional ByRef FixedAsset As String)
Dim rs1 As ADODB.Recordset
Dim sql As String

Set rs1 = New ADODB.Recordset
sql = " select * from FixedAssets where id =" & ExpensesID & ""
rs1.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
If rs1.RecordCount > 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
FixedAsset = IIf(IsNull(rs1("name").value), "", rs1("name").value)
Else
FixedAsset = IIf(IsNull(rs1("namee").value), "", rs1("namee").value)
End If
End If
End Sub
Private Sub Retrive(Optional project_id As Integer = 0, Optional Pand As Integer = 0, Optional Oper As Integer = 0)
    Dim AccountName As String
    Dim i As Integer
    Dim astrSplit2tems2() As String
    Dim astrSplitItems() As String
    Dim itemname As String
    Dim j As Integer
    Dim st As String
    Dim nElements As Integer
    Dim FixedAsset As String
 
    'On Error GoTo ErrTrap
   ' VSFlexGrid4.Clear flexClearScrollable, flexClearEverything
   ' VSFlexGrid4.Rows = 2
   ' VSFlexGrid4.Enabled = True
    'txt_opr_total.text = 0
          
   ' StrSQL = "SELECT     TOP 100 PERCENT dbo.TblEquepment.ProjectID, dbo.TblEquepment.Pand, dbo.TblEquepment.Opr, dbo.TblEquepment.ExpensesID, dbo.TblEquepment.EstHour, "
   'StrSQL = StrSQL & "                   dbo.TblEquepment.ActualHour, dbo.TblEquepment.TotalEs, dbo.TblEquepment.[value], dbo.TblEquepment.des, dbo.TblEquepment.ID, dbo.FixedAssets.code,"
   'StrSQL = StrSQL & "                     dbo.FixedAssets.name , dbo.FixedAssets.NameE"
   'StrSQL = StrSQL & "   FROM         dbo.TblEquepment LEFT OUTER JOIN"
 ' StrSQL = StrSQL & "                      dbo.FixedAssets ON dbo.TblEquepment.ExpensesID = dbo.FixedAssets.id"
'StrSQL = StrSQL & "   Where (dbo.TblEquepment.Projectid = " & project_id & ") And (dbo.TblEquepment.Pand = " & Pand & ") And (dbo.TblEquepment.OPR = " & Oper & ")"
'    Set RsDev = New ADODB.Recordset
'    RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
'
'    If Not (RsDev.BOF Or RsDev.EOF) Then
'        RsDev.MoveFirst
   
        With Me.VSFlexGrid4
        If Projects.VSFlexGrid2.TextMatrix(Projects.LngRow, Projects.VSFlexGrid2.ColIndex("eq")) <> "" Then
         st = Projects.VSFlexGrid2.TextMatrix(Projects.LngRow, Projects.VSFlexGrid2.ColIndex("eq"))
         st = Trim(st)
         astrSplitItems = Split(st, "@")
         nElements = UBound(astrSplitItems) - LBound(astrSplitItems)
            .Rows = .FixedRows + nElements
            For j = 0 To nElements - 1
            i = j + 1
                astrSplit2tems2 = Split(astrSplitItems(j), "#")
                .TextMatrix(i, .ColIndex("ExpensesID")) = val(astrSplit2tems2(0))
                     
                .TextMatrix(i, .ColIndex("EstHour")) = val(astrSplit2tems2(1))
            
                .TextMatrix(i, .ColIndex("ActualHour")) = val(astrSplit2tems2(2))
            
                .TextMatrix(i, .ColIndex("TotalEs")) = val(astrSplit2tems2(3))
           
                .TextMatrix(i, .ColIndex("value")) = val(astrSplit2tems2(4))
                .TextMatrix(i, .ColIndex("des")) = astrSplit2tems2(5)
                 RetriveFixedAsset val(astrSplit2tems2(0)), FixedAsset
                .TextMatrix(i, .ColIndex("FixedAsset")) = FixedAsset
                .TextMatrix(i, .ColIndex("hourval")) = val(astrSplit2tems2(6))

           ' Me.txt_opr_total.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("total"), .Rows - 1, .ColIndex("total"))
        Next j
        End If
        End With
          
    ReLineGrid

End Sub

Sub save()
Dim str As String
Dim i As Integer
str = ""

With Me.VSFlexGrid4
For i = 1 To .Rows - 1
 If .TextMatrix(i, .ColIndex("FixedAsset")) <> "" Then
 str = str & Trim(.TextMatrix(i, .ColIndex("ExpensesID"))) & "#"
 str = str & Trim(.TextMatrix(i, .ColIndex("EstHour"))) & "#"
 str = str & Trim(.TextMatrix(i, .ColIndex("ActualHour"))) & "#"
 str = str & Trim(.TextMatrix(i, .ColIndex("TotalEs"))) & "#"
 str = str & Trim(.TextMatrix(i, .ColIndex("value"))) & "#"
 str = str & Trim(.TextMatrix(i, .ColIndex("des"))) & "#"
 str = str & Trim(.TextMatrix(i, .ColIndex("hourval"))) & "#"
 str = str & Trim("@")
  str = str & Chr(13)
  str = Trim(str)
 End If
Next
Projects.VSFlexGrid2.TextMatrix(Projects.LngRow, Projects.VSFlexGrid2.ColIndex("eq")) = str
Projects.VSFlexGrid2.TextMatrix(Projects.LngRow, Projects.VSFlexGrid2.ColIndex("EquepVal")) = val(Text11.text)

End With
End Sub

Private Sub DeleteFgRowAther()

    With Me.VSFlexGrid4

        If .Row = -1 Then Exit Sub
        If .Row = 0 Then Exit Sub
        .RemoveItem .Row
        ReLineGrid
        '.AutoSize 0, .Cols - 1, False
     
    End With

End Sub

Private Sub Form_Activate()
'   PutFormOnTop Me.hWnd
End Sub
Private Sub ReLineGrid(Optional current_terms As String = "")
    Dim i As Integer
    Dim IntCounter As Integer
    Dim StartWeek As Double
    Dim EndWeek As Double
    Dim EarlyStartWeek As Double
    Dim EarlyEndWeek As Double
    Dim rs As ADODB.Recordset
    Text1.text = 0
Text8.text = 0
Text9.text = 0
Text11.text = 0
Text14.text = 0
    With VSFlexGrid4

        For i = .FixedRows To .Rows - 1

            If .TextMatrix(i, .ColIndex("ExpensesID")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("LineNo")) = IntCounter
                 .TextMatrix(i, .ColIndex("FullCode")) = currentterms & "-" & .TextMatrix(i, .ColIndex("LineNo"))
                 .TextMatrix(i, .ColIndex("TotalEs")) = val(.TextMatrix(i, .ColIndex("hourval"))) * val(.TextMatrix(i, .ColIndex("EstHour")))
             Text8.text = val(Text8) + val(.TextMatrix(i, .ColIndex("EstHour")))
             Text9.text = val(Text9.text) + val(.TextMatrix(i, .ColIndex("ActualHour")))
             Text11.text = val(Text11.text) + val(.TextMatrix(i, .ColIndex("TotalEs")))
             Text14.text = val(Text14.text) + val(.TextMatrix(i, .ColIndex("value")))
          Text1.text = val(Text1.text) + val(.TextMatrix(i, .ColIndex("hourval")))
            End If

        Next i
   
    End With

End Sub



Private Sub VSFlexGrid4_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim StrAccountCode As String
Dim sql As String
Dim rs1 As ADODB.Recordset
Set rs1 = New ADODB.Recordset
    Dim LngRow As Long
    
    With VSFlexGrid4

       Select Case .ColKey(Col)
            Case "FixedAsset"
                       StrAccountCode = .ComboData
                       LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("ExpensesID"), False, True)
                      .TextMatrix(Row, .ColIndex("ExpensesID")) = StrAccountCode
                     


sql = " select (UsedPowerPriceH+Hourdipp+UsedElectricPriceH) as HourVal from TblEquipments where fixedAssetid=" & val(.TextMatrix(Row, .ColIndex("ExpensesID"))) & ""
rs1.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText
If rs1.RecordCount > 0 Then
 .TextMatrix(Row, .ColIndex("hourval")) = IIf(IsNull(rs1("HourVal").value), "", rs1("HourVal").value)
 End If
          End Select
          
        If Row = .Rows - 1 Then
    
            .Rows = .Rows + 1
        End If
        End With
        ReLineGrid
End Sub

Private Sub VSFlexGrid4_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
 With VSFlexGrid4

      
        Select Case .ColKey(Col)
            
            Case "LineNo1"
           VSFlexGrid4.ComboList = ""
            Case "LineNo"
            
               VSFlexGrid4.ComboList = ""
               Case "des"
             VSFlexGrid4.ComboList = ""
               Case "value"
             VSFlexGrid4.ComboList = ""
             '    Cancel = True
              Case "TotalEs"
               VSFlexGrid4.ComboList = ""
              Case "ActualHour"
              VSFlexGrid4.ComboList = ""
               Case "EstHour"
     VSFlexGrid4.ComboList = ""
               
               
        End Select

    End With
End Sub

Private Sub VSFlexGrid4_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
Dim StrSQL As String
    Dim StrComboList As String
With VSFlexGrid4

  Select Case .ColKey(Col)
Case "FixedAsset"
Dim rs As ADODB.Recordset
                If SystemOptions.UserInterface = ArabicInterface Then
                    StrSQL = " SELECT     id, Name"
                    StrSQL = StrSQL & " from dbo.FixedAssets"
                    StrSQL = StrSQL & " WHERE    id IN"
                    StrSQL = StrSQL & " (SELECT     fixedAssetid"
                    StrSQL = StrSQL & " FROM         dbo.TblEquipments)"
                    StrSQL = StrSQL & " or   id IN"
                    StrSQL = StrSQL & " (SELECT     fixedAssetid"
                    StrSQL = StrSQL & "  FROM         dbo.TblCarsData)"
                    StrSQL = StrSQL & " order by Namee  "
                Else
                                        StrSQL = " SELECT     id, Namee"
                    StrSQL = StrSQL & " from dbo.FixedAssets"
                    StrSQL = StrSQL & " WHERE    id IN"
                    StrSQL = StrSQL & " (SELECT     fixedAssetid"
                    StrSQL = StrSQL & " FROM         dbo.TblEquipments)"
                    StrSQL = StrSQL & " or   id IN"
                    StrSQL = StrSQL & " (SELECT     fixedAssetid"
                    StrSQL = StrSQL & "  FROM         dbo.TblCarsData)"
                    StrSQL = StrSQL & " order by Name  "
                    
                End If
       Set rs = New ADODB.Recordset
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = VSFlexGrid4.BuildComboList(rs, "Name", "id")
                Else
                    StrComboList = VSFlexGrid4.BuildComboList(rs, "Namee", "id")
                End If

                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If

                .ComboList = StrComboList
                rs.Close
                Set rs = Nothing
'.Rows = .Rows + 1
End Select
End With

End Sub
Private Sub Form_Load()
    Dim GrdBack As ClsBackGroundPic
    Dim Dcombos As ClsDataCombos
Dim Xpid As Integer
Dim rwOp As Integer
Dim rwpand As Integer
    Set Dcombos = New ClsDataCombos
 
   'Dcombos.GetAccountingCodes Me.DcbAccount


    Set DCboSearch = New clsDCboSearch
   
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Clear").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
If Projects.TxtModFlg.text <> "R" Then
Cmd(0).Enabled = True
Else
Cmd(0).Enabled = False

End If
 Frame6.Visible = True
    Set GrdBack = New ClsBackGroundPic
    currentterms = Projects.VSFlexGrid2.TextMatrix(Projects.LngRow, Projects.VSFlexGrid2.ColIndex("fullcode"))
    If SystemOptions.UserInterface = ArabicInterface Then
                    Frame6.Caption = " „⁄œ«  «·⁄„·ÌÂ  —ﬁ„ : " & currentterms
                Else
                    Frame6.Caption = "Equipment For Process No: " & currentterms
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
Frame6.Caption = Me.Caption
Cmd(8).Caption = "Delete"
Lbl(7).Caption = "Totals"
 









    With Me.VSFlexGrid4
    
     

    .TextMatrix(0, .ColIndex("LineNo")) = "LineNo"
    .TextMatrix(0, .ColIndex("FixedAsset")) = "Name"
    .TextMatrix(0, .ColIndex("EstHour")) = "EstHour"
    
        .TextMatrix(0, .ColIndex("ActualHour")) = "ActualHour"
        .TextMatrix(0, .ColIndex("hourval")) = "hourval"
        .TextMatrix(0, .ColIndex("TotalEs")) = "TotalEs"
        .TextMatrix(0, .ColIndex("value")) = "value"
        
        
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


