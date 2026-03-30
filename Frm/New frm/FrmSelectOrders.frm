VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form FrmSelectOrders 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7680
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10020
   Icon            =   "FrmSelectOrders.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7680
   ScaleWidth      =   10020
   Begin VB.TextBox TxtCode 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   8685
   End
   Begin VB.CommandButton CMDOK 
      Caption         =   "„Ê«›ﬁ"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   7320
      Width           =   1335
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid1 
      Height          =   5835
      Left            =   0
      TabIndex        =   0
      Top             =   1200
      Width           =   9945
      _cx             =   17542
      _cy             =   10292
      Appearance      =   2
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
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmSelectOrders.frx":000C
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
   Begin C1SizerLibCtl.C1Elastic Ele 
      Height          =   765
      Index           =   5
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   10035
      _cx             =   17701
      _cy             =   1349
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial (Arabic)"
         Size            =   24
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   4
      MousePointer    =   0
      Version         =   801
      BackColor       =   16777215
      ForeColor       =   4210688
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Picture         =   "FrmSelectOrders.frx":014A
      Caption         =   "   ÕœÌœ «·⁄—Ê÷   "
      Align           =   0
      AutoSizeChildren=   0
      BorderWidth     =   0
      ChildSpacing    =   0
      Splitter        =   0   'False
      FloodDirection  =   0
      FloodPercent    =   0
      CaptionPos      =   6
      WordWrap        =   -1  'True
      MaxChildSize    =   0
      MinChildSize    =   0
      TagWidth        =   0
      TagPosition     =   0
      Style           =   0
      TagSplit        =   2
      PicturePos      =   0
      CaptionStyle    =   0
      ResizeFonts     =   0   'False
      GridRows        =   0
      GridCols        =   0
      Frame           =   3
      FrameStyle      =   0
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   ""
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin VB.CheckBox Check17 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   " ÕœÌœ «·ﬂ·"
         Height          =   195
         Left            =   8400
         TabIndex        =   3
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "ﬂÊœ «·⁄—÷"
      Height          =   345
      Index           =   32
      Left            =   8880
      TabIndex        =   5
      Top             =   840
      Width           =   930
   End
End
Attribute VB_Name = "FrmSelectOrders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check17_Click()
 
   Dim i As Integer

    If Check17.value = vbChecked Then

        With Me.Grid1
 
            For i = 1 To .Rows - 1
        
           .Cell(flexcpChecked, i, .ColIndex("Select")) = flexChecked
            Next i

        End With

    Else

        With Me.Grid1

            For i = 1 To .Rows - 1
        
              .Cell(flexcpChecked, i, .ColIndex("Select")) = flexUnchecked
            Next i

        End With

    End If

 
End Sub
Sub save()
    FrmComparePrices.GridOvers.Clear flexClearScrollable, flexClearEverything
    FrmComparePrices.GridOvers.Rows = 1
  Dim currentpos As Integer
  currentrow = 1
  Dim Str As String
If FrmComparePrices.StrOrder <> "" Then
Str = FrmComparePrices.StrOrder & ","
End If

  With Me.Grid1

        For i = 1 To .Rows - 1

            If .TextMatrix(i, .ColIndex("Transaction_ID")) <> "" And .Cell(flexcpChecked, i, .ColIndex("Select")) = flexChecked Then
         
                             With FrmComparePrices.GridOvers
                                .Rows = .Rows + 1
                                        .TextMatrix(currentrow, .ColIndex("Transaction_ID")) = Grid1.TextMatrix(i, Grid1.ColIndex("Transaction_ID"))
                                        .TextMatrix(currentrow, .ColIndex("NoteSerial1")) = Grid1.TextMatrix(i, Grid1.ColIndex("NoteSerial1"))
                                        .TextMatrix(currentrow, .ColIndex("Transaction_Date")) = Grid1.TextMatrix(i, Grid1.ColIndex("Transaction_Date"))
                                        .TextMatrix(currentrow, .ColIndex("CusName")) = Grid1.TextMatrix(i, Grid1.ColIndex("CusName"))
                                        .TextMatrix(currentrow, .ColIndex("PODays")) = Grid1.TextMatrix(i, Grid1.ColIndex("PODays"))
                                      
                                        Str = Str & Trim(.TextMatrix(currentrow, .ColIndex("Transaction_ID"))) & ","
 
                             End With
                            
                    currentrow = currentrow + 1
                    
            End If
            
            '
        Next i
        Str = Str & -1
        FrmComparePrices.StrOrder = Str
  End With
End Sub
Private Sub CMDOK_Click()
save
  Unload Me
End Sub
Sub LoadGride(Optional Flag As String)
  Dim RsDev As ADODB.Recordset
    Dim StrSQL As String
    Dim i As Integer
       Dim StrOrder As String
   StrOrder = FrmComparePrices.StrOrder
   StrSQL = "SELECT     TOP 100 PERCENT  dbo.Transactions.Transaction_ID, dbo.Transactions.Transaction_Date, dbo.Transactions.NoteSerial1, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, "
   StrSQL = StrSQL & "  dbo.Transactions.PODays"
   StrSQL = StrSQL & " FROM         dbo.Transactions INNER JOIN"
   StrSQL = StrSQL & " dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID"
    StrSQL = StrSQL & "  Where  (dbo.Transactions.Transaction_Type = 46)"
 If TxtCode.text <> "" Then
StrSQL = StrSQL & " and  dbo.Transactions.NoteSerial1  like'%" & TxtCode.text & "%'"
End If

  
    Set RsDev = New ADODB.Recordset
    RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsDev.BOF Or RsDev.EOF) Then
        RsDev.MoveFirst
    
        With Me.Grid1
    
            .Rows = .FixedRows + RsDev.RecordCount

            For i = .FixedRows To .Rows - 1
  '
                 .TextMatrix(i, .ColIndex("Transaction_ID")) = IIf(IsNull(RsDev("Transaction_ID").value), 0, val(RsDev("Transaction_ID").value))
                 
                .TextMatrix(i, .ColIndex("NoteSerial1")) = IIf(IsNull(RsDev("NoteSerial1").value), 0, val(RsDev("NoteSerial1").value))
             
                .TextMatrix(i, .ColIndex("Transaction_Date")) = IIf(IsNull(RsDev("Transaction_Date").value), 0, (RsDev("Transaction_Date").value))
            
                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(RsDev("CusName").value), "", RsDev("CusName").value)
                Else
                    .TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(RsDev("CusNamee").value), "", RsDev("CusNamee").value)
                End If
              
                .TextMatrix(i, .ColIndex("PODays")) = IIf(IsNull(RsDev("PODays").value), "", RsDev("PODays").value)
                
                .Cell(flexcpChecked, i, .ColIndex("Select")) = flexUnchecked
            
                RsDev.MoveNext
            Next i
    .AutoSize 0, .Cols - 1, False
        End With

    End If
chek
End Sub
Sub Retrive(Optional Flag As String)
  Dim RsDev As ADODB.Recordset
    Dim StrSQL As String
    Dim i As Integer
       Dim StrOrder As String
   StrOrder = FrmComparePrices.StrOrder
   StrSQL = "SELECT     TOP 100 PERCENT  dbo.Transactions.Transaction_ID, dbo.Transactions.Transaction_Date, dbo.Transactions.NoteSerial1, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, "
   StrSQL = StrSQL & "  dbo.Transactions.PODays"
   StrSQL = StrSQL & " FROM         dbo.Transactions INNER JOIN"
   StrSQL = StrSQL & " dbo.TblCustemers ON dbo.Transactions.CusID = dbo.TblCustemers.CusID"
  ' StrSQL = StrSQL & "  Where (dbo.Transactions.Approved  =1) And (dbo.Transactions.Transaction_Type = 46)"
  If Flag = "N" Then
    StrSQL = StrSQL & "  Where  (dbo.Transactions.Transaction_Type = 46)"
     ElseIf Flag = "R" And StrOrder <> "" Then

     StrSQL = StrSQL & "  Where (dbo.Transactions.Transaction_Type =46 ) And (dbo.Transactions.Transaction_ID in ( " & StrOrder & " )) "
    End If
  
    Set RsDev = New ADODB.Recordset
    RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsDev.BOF Or RsDev.EOF) Then
        RsDev.MoveFirst
    
        With Me.Grid1
    
            .Rows = .FixedRows + RsDev.RecordCount

            For i = .FixedRows To .Rows - 1
  '
                 .TextMatrix(i, .ColIndex("Transaction_ID")) = IIf(IsNull(RsDev("Transaction_ID").value), 0, val(RsDev("Transaction_ID").value))
                 
                .TextMatrix(i, .ColIndex("NoteSerial1")) = IIf(IsNull(RsDev("NoteSerial1").value), 0, val(RsDev("NoteSerial1").value))
             
                .TextMatrix(i, .ColIndex("Transaction_Date")) = IIf(IsNull(RsDev("Transaction_Date").value), 0, (RsDev("Transaction_Date").value))
            
                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(RsDev("CusName").value), "", RsDev("CusName").value)
                Else
                    .TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(RsDev("CusNamee").value), "", RsDev("CusNamee").value)
                End If
              
                .TextMatrix(i, .ColIndex("PODays")) = IIf(IsNull(RsDev("PODays").value), "", RsDev("PODays").value)
                
                .Cell(flexcpChecked, i, .ColIndex("Select")) = flexUnchecked
            
                RsDev.MoveNext
            Next i
    .AutoSize 0, .Cols - 1, False
        End With

    End If

End Sub
Sub chek()
  Dim RsDev As ADODB.Recordset
    Dim StrSQL As String
    Dim i As Integer
    Dim j As Integer
    Dim StrOrder As String
       
   StrOrder = FrmComparePrices.StrOrder
   If StrOrder <> "" Then
    StrSQL = "SELECT   Transaction_ID from Transactions "
     StrSQL = StrSQL & "  Where  (Transaction_Type = 46) And (Transaction_ID in ( " & StrOrder & " )) "
     Set RsDev = New ADODB.Recordset
    RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If RsDev.RecordCount > 0 Then
       With Me.Grid1
       For j = 1 To RsDev.RecordCount
       For i = 1 To .Rows - 1
          If val(.TextMatrix(i, .ColIndex("Transaction_ID"))) = val(RsDev("Transaction_ID").value) Then
                .Cell(flexcpChecked, i, .ColIndex("Select")) = flexChecked
                End If
       Next i
       RsDev.MoveNext
       Next j
       End With
    End If
    End If
End Sub

Private Sub ChangeLang()


    Me.Caption = "Select Offers "
    Ele(5).Caption = Me.Caption
    Check17.Caption = "Select All "
    Check17.RightToLeft = False
    lbl(32).Caption = "Offer Code"

CMDOK.Caption = "OK"

    With Me.Grid1
        .TextMatrix(0, .ColIndex("Select")) = "Select"
        .TextMatrix(0, .ColIndex("NoteSerial1")) = "Offer Code"
        .TextMatrix(0, .ColIndex("Transaction_Date")) = "Date"
         .TextMatrix(0, .ColIndex("CusName")) = "Customer Name"
        .TextMatrix(0, .ColIndex("PODays")) = "Due Date"
        .TextMatrix(0, .ColIndex("remark")) = "Remarks"
        
    End With

End Sub
Private Sub Form_Load()
    Resize_Form Me
        If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
If FrmComparePrices.TxtModFlg.text <> "R" Then
Retrive "N"
Else
Retrive "R"
End If
'If FrmComparePrices.TxtModFlg.text <> "N" Then
chek
'End If
End Sub

Private Sub Grid1_Click()
save
Form_Load
TxtCode.text = ""

End Sub

Private Sub TxtCode_Change()
LoadGride
End Sub
