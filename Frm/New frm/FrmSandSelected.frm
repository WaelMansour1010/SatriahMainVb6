VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Begin VB.Form FrmSandSelected 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "«Œ Ì«— ”‰œ«  ’—› «·„Ê«œ"
   ClientHeight    =   4650
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9420
   Icon            =   "FrmSandSelected.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   4650
   ScaleWidth      =   9420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame10 
      Height          =   3735
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   360
      Width           =   9375
      Begin VB.TextBox TxtTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   3315
         Width           =   1575
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid1 
         Height          =   3180
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   9120
         _cx             =   16087
         _cy             =   5609
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   0   'False
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
         Rows            =   3
         Cols            =   8
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmSandSelected.frx":038A
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
         Left            =   8400
         TabIndex        =   8
         Top             =   3360
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
         ButtonImage     =   "FrmSandSelected.frx":04D4
         DrawFocusRectangle=   0   'False
      End
      Begin VB.Label Label1 
         Caption         =   "«·«Ã„«·Ì"
         Height          =   255
         Left            =   3600
         TabIndex        =   9
         Top             =   3360
         Width           =   1815
      End
      Begin VB.Label Label27 
         Caption         =   "«·«Ã„«·Ì"
         Height          =   255
         Left            =   1800
         TabIndex        =   7
         Top             =   3360
         Width           =   1815
      End
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   375
      Index           =   0
      Left            =   1650
      TabIndex        =   0
      Top             =   4200
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
      Top             =   4200
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
      Top             =   4200
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
      Caption         =   "«Œ Ì«— ”‰œ«  ’—› «·„Ê«œ"
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
      Left            =   -30
      TabIndex        =   3
      Top             =   0
      Width           =   9435
   End
End
Attribute VB_Name = "FrmSandSelected"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim DCboSearch As clsDCboSearch
Dim currentterms As String
Private Sub ReLineGrid()
    Dim i As Integer
    Dim IntCounter As Integer
    Dim StartWeek As Double
    Dim EndWeek As Double
    Dim EarlyStartWeek As Double
    Dim EarlyEndWeek As Double
    Dim rs As ADODB.Recordset

IntCounter = 0
  TxtTotal.text = 0
        Set rs = New ADODB.Recordset
             

    With VSFlexGrid1

        For i = .FixedRows To .Rows - 1

            If .TextMatrix(i, .ColIndex("NoteSerial1")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("Ser")) = IntCounter
            If .Cell(flexcpChecked, i, .ColIndex("selcted")) = flexChecked Then
         TxtTotal.text = val(TxtTotal.text) + val(.TextMatrix(i, .ColIndex("total")))
               End If
            End If
        
        Next i

  
    End With



End Sub


Private Sub Cmd_Click(Index As Integer)

    Select Case Index

        Case 0
    save
    Unload Me
' GetData
           

       Case 24
     '  AddNewFgRowother
       Case 8
            DeleteFgRowAther
    End Select

End Sub
Sub save()
Dim str As String
Dim strnotes
Dim i As Integer
str = ""
strnotes = ""
With Me.VSFlexGrid1
For i = 1 To .Rows - 1
 If .TextMatrix(i, .ColIndex("NoteSerial1")) <> "" Then
 If .Cell(flexcpChecked, i, .ColIndex("selcted")) = flexChecked Then
 str = str & Trim(.TextMatrix(i, .ColIndex("NProductionOrderNO"))) & "#"
 str = str & Trim(.TextMatrix(i, .ColIndex("Transaction_ID"))) & "#"
 str = str & 1 & "#"
strnotes = strnotes & Trim(.TextMatrix(i, .ColIndex("NoteSerial1"))) & "#"
 str = str & Trim(.TextMatrix(i, .ColIndex("NoteSerial1"))) & "#"
 str = str & Trim(.TextMatrix(i, .ColIndex("Transaction_Date"))) & "#"
 str = str & Trim(.TextMatrix(i, .ColIndex("total"))) & "#"
  str = str & Trim(.TextMatrix(i, .ColIndex("id"))) & "#"
 str = str & Trim("@")
  str = str & Chr(13)
  str = Trim(str)
 End If
 End If
Next
FrmProductionAllocation.Grid.TextMatrix(FrmProductionAllocation.LngRow, FrmProductionAllocation.Grid.ColIndex("sand")) = strnotes
FrmProductionAllocation.Grid.TextMatrix(FrmProductionAllocation.LngRow, FrmProductionAllocation.Grid.ColIndex("sandat")) = str
FrmProductionAllocation.Grid.TextMatrix(FrmProductionAllocation.LngRow, FrmProductionAllocation.Grid.ColIndex("MaterialsValue")) = val(TxtTotal.text)




End With
End Sub

Private Sub DeleteFgRowAther()

    With Me.VSFlexGrid1

        If .Row = -1 Then Exit Sub
        If .Row = 0 Then Exit Sub
        .RemoveItem .Row
        '.AutoSize 0, .Cols - 1, False
     
    End With

End Sub

Private Sub RetriveSand(Optional NProductionOrderNO As String, Optional ProIDdet As Integer = 0)
 
    Dim RsDev As ADODB.Recordset
    Dim StrSQL As String
    Dim i As Integer
 
    'On Error GoTo ErrTrap
    VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
    VSFlexGrid1.Rows = 2
    VSFlexGrid1.Enabled = True
    'txt_opr_total.text = 0
          
  '  StrSQL = "SELECT     dbo.Transaction_Details.NProductionOrderNO, SUM(dbo.Transaction_Details.ShowQty * dbo.Transaction_Details.showPrice) AS total, dbo.Transactions.NoteSerial1, "
  '  StrSQL = StrSQL & "                  dbo.Transactions.Transaction_Date , dbo.Transactions.Transaction_ID"
'StrSQL = StrSQL & "   FROM         dbo.Transactions INNER JOIN"
'StrSQL = StrSQL & "                       dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID"
'StrSQL = StrSQL & "  WHERE     (dbo.Transactions.Transaction_Type = 27) AND (dbo.Transaction_Details.NProductionOrderNO  = N'" & NProductionOrderNO & " ' )"
'StrSQL = StrSQL & "  GROUP BY dbo.Transactions.NoteSerial1, dbo.Transactions.Transaction_Date, dbo.Transaction_Details.NProductionOrderNO, dbo.Transactions.Transaction_ID"
StrSQL = " SELECT     dbo.Transaction_Details.NProductionOrderNO, SUM(dbo.Transaction_Details.ShowQty * dbo.Transaction_Details.showPrice) AS total, dbo.Transactions.NoteSerial1,"
StrSQL = StrSQL & "                      dbo.Transactions.Transaction_Date , dbo.Transactions.Transaction_ID, dbo.Transaction_Details.Flag , dbo.Transaction_Details.ProIDdet  , dbo.Transaction_Details.ID"
StrSQL = StrSQL & " FROM         dbo.Transactions INNER JOIN"
StrSQL = StrSQL & "                      dbo.Transaction_Details ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID"
StrSQL = StrSQL & " Where (dbo.Transactions.Transaction_Type = 27) AND (dbo.Transaction_Details.NProductionOrderNO  = N'" & NProductionOrderNO & " ' )"
If FrmProductionAllocation.TxtModFlg = "E" Then
StrSQL = StrSQL & " and ( (dbo.Transaction_Details.ProIDdet = " & ProIDdet & ") or((dbo.Transaction_Details.Flag = 0)or (dbo.Transaction_Details.Flag IS NULL))) "
End If
 If FrmProductionAllocation.TxtModFlg = "N" Then
StrSQL = StrSQL & " and ((dbo.Transaction_Details.Flag = 0)or (dbo.Transaction_Details.Flag IS NULL)) "
End If
StrSQL = StrSQL & " GROUP BY dbo.Transactions.NoteSerial1, dbo.Transactions.Transaction_Date, dbo.Transaction_Details.NProductionOrderNO, dbo.Transactions.Transaction_ID,"
 StrSQL = StrSQL & "                     dbo.Transaction_Details.Flag , dbo.Transaction_Details.ProIDdet , dbo.Transaction_Details.ID"

    Set RsDev = New ADODB.Recordset
    RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsDev.BOF Or RsDev.EOF) Then
        RsDev.MoveFirst
    
        With Me.VSFlexGrid1
            .Rows = .FixedRows + RsDev.RecordCount

            For i = .FixedRows To .Rows - 1
            .TextMatrix(i, .ColIndex("Ser")) = i
                .TextMatrix(i, .ColIndex("NProductionOrderNO")) = IIf(IsNull(RsDev("NProductionOrderNO").value), "", RsDev("NProductionOrderNO").value)
                .TextMatrix(i, .ColIndex("Transaction_ID")) = IIf(IsNull(RsDev("Transaction_ID").value), "", RsDev("Transaction_ID").value)
               
                .TextMatrix(i, .ColIndex("NoteSerial1")) = IIf(IsNull(RsDev("NoteSerial1").value), "", RsDev("NoteSerial1").value)
            
                .TextMatrix(i, .ColIndex("Transaction_Date")) = IIf(IsNull(RsDev("Transaction_Date").value), "", RsDev("Transaction_Date").value)
           
                .TextMatrix(i, .ColIndex("total")) = IIf(IsNull(RsDev("total").value), "", RsDev("total").value)
                   If (RsDev("Flag").value) = True Then
                .TextMatrix(i, .ColIndex("selcted")) = -1
                Else
                 .TextMatrix(i, .ColIndex("selcted")) = 0
                End If
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(RsDev("ID").value), "", RsDev("ID").value)
              '  .TextMatrix(i, .ColIndex("jobid")) = IIf(IsNull(RsDev("JobID").value), "", RsDev("JobID").value)
            
        
                RsDev.MoveNext
            Next i

           ' Me.txt_opr_total.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("total"), .Rows - 1, .ColIndex("total"))
        
        End With

    End If
          
    ReLineGrid

End Sub



Private Sub retrive(Optional ProID As Integer = 0, Optional ProIDdet As Integer = 0)
 
    Dim RsDev As ADODB.Recordset
    Dim StrSQL As String
    Dim i As Integer
 
    'On Error GoTo ErrTrap
    VSFlexGrid1.Clear flexClearScrollable, flexClearEverything
    VSFlexGrid1.Rows = 2
    VSFlexGrid1.Enabled = True
    'txt_opr_total.text = 0
          
    StrSQL = " select * from TblProductionAllocDetails1"
StrSQL = StrSQL & "   Where (ProID =" & ProID & ") And (ProIDdet = " & ProIDdet & ") "
    Set RsDev = New ADODB.Recordset
    RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
 If Not (RsDev.BOF Or RsDev.EOF) Then
        RsDev.MoveFirst
    
        With Me.VSFlexGrid1
            .Rows = .FixedRows + RsDev.RecordCount

            For i = .FixedRows To .Rows - 1
            .TextMatrix(i, .ColIndex("Ser")) = i
                .TextMatrix(i, .ColIndex("NProductionOrderNO")) = IIf(IsNull(RsDev("NProductionOrderNO").value), "", RsDev("NProductionOrderNO").value)
                .TextMatrix(i, .ColIndex("Transaction_ID")) = IIf(IsNull(RsDev("Transaction_ID").value), "", RsDev("Transaction_ID").value)
               
                .TextMatrix(i, .ColIndex("NoteSerial1")) = IIf(IsNull(RsDev("NoteSerial1").value), "", RsDev("NoteSerial1").value)
            
                .TextMatrix(i, .ColIndex("Transaction_Date")) = IIf(IsNull(RsDev("Transaction_Date").value), "", RsDev("Transaction_Date").value)
           
                .TextMatrix(i, .ColIndex("total")) = IIf(IsNull(RsDev("total").value), "", RsDev("total").value)
               If (RsDev("Selcted").value) = True Then
                .TextMatrix(i, .ColIndex("selcted")) = -1
                Else
                 .TextMatrix(i, .ColIndex("selcted")) = 0
                End If
               .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(RsDev("idd").value), "", RsDev("idd").value)
              '  .TextMatrix(i, .ColIndex("jobid")) = IIf(IsNull(RsDev("JobID").value), "", RsDev("JobID").value)
            
        
                RsDev.MoveNext
            Next i

           ' Me.txt_opr_total.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("total"), .Rows - 1, .ColIndex("total"))
        
        End With

    End If
          
    ReLineGrid

End Sub


Private Sub Form_Activate()
'   PutFormOnTop Me.hWnd
End Sub


Private Sub Form_Load()
    Dim GrdBack As ClsBackGroundPic
    Dim Dcombos As ClsDataCombos
Dim My_SQL As String

Dim Xpid As Integer
Dim rwID As Integer
    Set Dcombos = New ClsDataCombos
 
   'Dcombos.GetAccountingCodes Me.DcbAccount


    Set DCboSearch = New clsDCboSearch
   
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Clear").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture

    Set GrdBack = New ClsBackGroundPic
If FrmProductionAllocation.TxtModFlg.text = "R" Then
VSFlexGrid1.Editable = flexEDNone
Cmd(0).Enabled = False
Else
VSFlexGrid1.Editable = flexEDKbdMouse
Cmd(0).Enabled = True
End If
 rwID = val(FrmProductionAllocation.Grid.TextMatrix(FrmProductionAllocation.LngRow, FrmProductionAllocation.Grid.ColIndex("id")))
currentterms = FrmProductionAllocation.Grid.TextMatrix(FrmProductionAllocation.LngRow, FrmProductionAllocation.Grid.ColIndex("NProductionOrderNO"))
If FrmProductionAllocation.TxtModFlg.text <> "R" Then
 If currentterms <> "" Then
 RetriveSand currentterms, rwID
 End If
End If
   Xpid = val(FrmProductionAllocation.txtid.text)
   
 If FrmProductionAllocation.TxtModFlg.text = "R" Then
If (Xpid <> 0 And rwID <> 0) Then
retrive Xpid, rwID
 End If
 End If
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
 
    
  '
End Sub









Private Sub VSFlexGrid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
ReLineGrid
End Sub

