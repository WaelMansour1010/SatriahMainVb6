VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form FrmSelectItems 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5955
   Icon            =   "FrmSelectItems.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   5955
   Begin VB.TextBox TxtItem 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   240
      TabIndex        =   5
      Top             =   1080
      Width           =   3165
   End
   Begin VB.TextBox TxtCodeItem 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3600
      TabIndex        =   4
      Top             =   1080
      Width           =   2205
   End
   Begin VB.CommandButton CMDOK 
      Caption         =   "„Ê«›ﬁ"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   5280
      Width           =   1335
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid1 
      Height          =   3795
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   5865
      _cx             =   10345
      _cy             =   6694
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
      FormatString    =   $"FrmSelectItems.frx":000C
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
      Width           =   5955
      _cx             =   10504
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
      Picture         =   "FrmSelectItems.frx":012A
      Caption         =   " ÕœÌœ «·«’‰«›  "
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
         Left            =   4560
         TabIndex        =   3
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "«”„ «·’‰›"
      Height          =   345
      Index           =   0
      Left            =   1440
      TabIndex        =   7
      Top             =   840
      Width           =   930
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "ﬂÊœ «·’‰›"
      Height          =   345
      Index           =   32
      Left            =   4200
      TabIndex        =   6
      Top             =   840
      Width           =   930
   End
End
Attribute VB_Name = "FrmSelectItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
Private Sub ChangeLang()
  '  Dim XPic As IPictureDisp
  '  Set XPic = Me.btnFirst.ButtonImage
  '  Set Me.btnFirst.ButtonImage = Me.btnLast.ButtonImage
  '  Set Me.btnLast.ButtonImage = XPic
  '  Set XPic = Me.btnPrevious.ButtonImage
  '  Set Me.btnPrevious.ButtonImage = Me.btnNext.ButtonImage
  '  Set Me.btnNext.ButtonImage = XPic

    Me.Caption = "Select Items "
    Ele(5).Caption = Me.Caption
    Check17.Caption = "Select All "
    Check17.RightToLeft = False
    lbl(32).Caption = "Item Code"
     lbl(0).Caption = "Item Name"
CMDOK.Caption = "OK"

    With Me.Grid1
        .TextMatrix(0, .ColIndex("Select")) = "Select"
        .TextMatrix(0, .ColIndex("ItemCode")) = "Item Code"
        .TextMatrix(0, .ColIndex("ItemName")) = "Item Name"
        
    End With

End Sub
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
    FrmComparePrices.GridItems.Clear flexClearScrollable, flexClearEverything
    FrmComparePrices.GridItems.Rows = 1
  Dim currentpos As Integer
  Dim Str As String
  currentrow = 1
  If FrmComparePrices.StrItemID <> "" Then
  Str = FrmComparePrices.StrItemID & ","
  End If
  With Me.Grid1

        For i = 1 To .Rows - 1

            If .TextMatrix(i, .ColIndex("ItemID")) <> "" And .Cell(flexcpChecked, i, .ColIndex("Select")) = flexChecked Then
         
                             With FrmComparePrices.GridItems
                                .Rows = .Rows + 1
                                        .TextMatrix(currentrow, .ColIndex("ItemID")) = Grid1.TextMatrix(i, Grid1.ColIndex("ItemID"))
                                        .TextMatrix(currentrow, .ColIndex("ItemCode")) = Grid1.TextMatrix(i, Grid1.ColIndex("ItemCode"))
                                        .TextMatrix(currentrow, .ColIndex("ItemName")) = Grid1.TextMatrix(i, Grid1.ColIndex("ItemName"))
                                         Str = Str & Trim(.TextMatrix(currentrow, .ColIndex("ItemID"))) & ","
                                        
 
                             End With
                            
                    currentrow = currentrow + 1
                    
            End If
            
            '
        Next i
         Str = Str & -1
        FrmComparePrices.StrItemID = Str
  End With
End Sub
Private Sub CMDOK_Click()
save
  Unload Me
End Sub
Sub Retrive(Optional Flag As String)
  Dim RsDev As ADODB.Recordset
    Dim StrSQL As String
    Dim i As Integer
    Dim StrItemID As String
   StrItemID = FrmComparePrices.StrItemID
StrSQL = "SELECT     ItemID, ItemCode, ItemName, ItemNamee"
StrSQL = StrSQL & "  from dbo.TblItems"
StrSQL = StrSQL & " WHERE     (ItemType = 0)"
  If Flag = "R" And StrItemID <> "" Then
    StrSQL = StrSQL & "  and (ItemID in ( " & StrItemID & " )) "
    End If
  
         ' ItemID, ItemCode, ItemName, ItemNamee
         
    Set RsDev = New ADODB.Recordset
    RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsDev.BOF Or RsDev.EOF) Then
        RsDev.MoveFirst
    
        With Me.Grid1
    
            .Rows = .FixedRows + RsDev.RecordCount

            For i = .FixedRows To .Rows - 1
  '
                 .TextMatrix(i, .ColIndex("ItemID")) = IIf(IsNull(RsDev("ItemID").value), 0, val(RsDev("ItemID").value))
                 
                .TextMatrix(i, .ColIndex("ItemCode")) = IIf(IsNull(RsDev("ItemCode").value), 0, (RsDev("ItemCode").value))
             
             
            
                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(RsDev("ItemName").value), "", RsDev("ItemName").value)
                Else
                    .TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(RsDev("ItemNamee").value), "", RsDev("ItemNamee").value)
                End If
              
             
                
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
    Dim StrItemID As String
       
   StrItemID = FrmComparePrices.StrItemID
   If StrItemID <> "" Then
    StrSQL = "SELECT   ItemID from TblItems "
     StrSQL = StrSQL & "  Where  (ItemType = 0)And (ItemID in ( " & StrItemID & " )) "
     Set RsDev = New ADODB.Recordset
    RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If RsDev.RecordCount > 0 Then
       With Me.Grid1
       For j = 1 To RsDev.RecordCount
       For i = 1 To .Rows - 1
          If val(.TextMatrix(i, .ColIndex("ItemID"))) = val(RsDev("ItemID").value) Then
                .Cell(flexcpChecked, i, .ColIndex("Select")) = flexChecked
                End If
       Next i
       RsDev.MoveNext
       Next j
       End With
    End If
    End If
End Sub


Sub LoadGride()
  Dim RsDev As ADODB.Recordset
    Dim StrSQL As String
    Dim i As Integer
    Dim StrItemID As String
   
StrSQL = "SELECT     ItemID, ItemCode, ItemName, ItemNamee"
StrSQL = StrSQL & "  from dbo.TblItems"
StrSQL = StrSQL & " WHERE     (ItemType = 0)"
If TxtCodeItem.text <> "" Then
StrSQL = StrSQL & " and  ItemCode like'%" & TxtCodeItem.text & "%'"
End If
If TxtItem.text <> "" Then
  If SystemOptions.UserInterface = ArabicInterface Then
                StrSQL = StrSQL & " and  ItemName like'%" & TxtItem.text & "%'"
        Else
            StrSQL = StrSQL & " and  ItemNamee like'%" & TxtItem.text & "%'"
        End If
  End If
  
         ' ItemID, ItemCode, ItemName, ItemNamee
         
    Set RsDev = New ADODB.Recordset
    RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (RsDev.BOF Or RsDev.EOF) Then
        RsDev.MoveFirst
    
        With Me.Grid1
    
            .Rows = .FixedRows + RsDev.RecordCount

            For i = .FixedRows To .Rows - 1
  '
                 .TextMatrix(i, .ColIndex("ItemID")) = IIf(IsNull(RsDev("ItemID").value), 0, val(RsDev("ItemID").value))
                 
                .TextMatrix(i, .ColIndex("ItemCode")) = IIf(IsNull(RsDev("ItemCode").value), 0, (RsDev("ItemCode").value))
             
             
            
                If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(RsDev("ItemName").value), "", RsDev("ItemName").value)
                Else
                    .TextMatrix(i, .ColIndex("ItemName")) = IIf(IsNull(RsDev("ItemNamee").value), "", RsDev("ItemNamee").value)
                End If
              
             
                
                .Cell(flexcpChecked, i, .ColIndex("Select")) = flexUnchecked
            
                RsDev.MoveNext
            Next i
    .AutoSize 0, .Cols - 1, False
        End With

    End If
    chek
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
TxtCodeItem.text = ""
TxtItem.text = ""
End Sub

Private Sub TxtCodeItem_Change()
LoadGride
End Sub

Private Sub TxtItem_Change()
LoadGride
End Sub
