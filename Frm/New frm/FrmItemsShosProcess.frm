VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Begin VB.Form FrmIemsShosProcess 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "«·⁄—Ê÷ ÿ»ﬁ« ··’‰›"
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12135
   Icon            =   "FrmItemsShosProcess.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   4485
   ScaleWidth      =   12135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame6 
      Height          =   3615
      Left            =   -120
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   480
      Width           =   12255
      Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid3 
         Height          =   3195
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   12030
         _cx             =   21220
         _cy             =   5636
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
         BackColorBkg    =   16777215
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
         Cols            =   21
         FixedRows       =   1
         FixedCols       =   2
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmItemsShosProcess.frx":038A
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
         ExplorerBar     =   5
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
      Caption         =   "«·⁄—Ê÷ ÿ»ﬁ« ··’‰›"
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
      Left            =   45
      TabIndex        =   3
      Top             =   0
      Width           =   12060
   End
End
Attribute VB_Name = "FrmIemsShosProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim DCboSearch As clsDCboSearch
Dim currentterms As String

Sub Retriveshow(Optional IDitem As Integer = 0)
Dim sql As String
Dim i As Integer
Dim Rsditails As ADODB.Recordset
Set Rsditails = New ADODB.Recordset
  VSFlexGrid3.Clear flexClearScrollable, flexClearEverything
    VSFlexGrid3.Rows = 2
sql = " SELECT     dbo.TblItems.HaveSerial, dbo.Transactions.Transaction_Date, dbo.Transactions.PODays, dbo.Transactions.Transaction_Type, dbo.Transactions.NoteSerial1, "
sql = sql & "                      dbo.Transaction_Details.Item_ID, dbo.TblItems.Fullcode, dbo.TblItems.ItemCode, dbo.TblItems.ItemName, dbo.TblItems.ItemNamee, dbo.TblUnites.UnitID,"
sql = sql & "                      dbo.TblUnites.UnitName, dbo.TblUnites.UnitNamee, dbo.Transaction_Details.Quantity, dbo.Transaction_Details.Price, dbo.Transaction_Details.ShowQty,"
sql = sql & "                      dbo.Transaction_Details.showPrice, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.Cus_Phone, dbo.TblCustemers.Cus_mobile,"
sql = sql & "                      dbo.TblCustemers.Fullcode AS CusFullcode, dbo.Transactions.CusID"
sql = sql & " FROM         dbo.TblCustemers RIGHT OUTER JOIN"
sql = sql & "                      dbo.Transactions ON dbo.TblCustemers.CusID = dbo.Transactions.CusID LEFT OUTER JOIN"
sql = sql & "                      dbo.TblItems INNER JOIN"
sql = sql & "                      dbo.Transaction_Details ON dbo.TblItems.ItemID = dbo.Transaction_Details.Item_ID INNER JOIN"
sql = sql & "                      dbo.TblUnites ON dbo.Transaction_Details.UnitId = dbo.TblUnites.UnitID ON dbo.Transactions.Transaction_ID = dbo.Transaction_Details.Transaction_ID"
sql = sql & " Where (dbo.Transactions.Transaction_Type = 46) And (dbo.TblItems.ItemID =" & val(IDitem) & ")"
sql = sql & " ORDER BY dbo.Transactions.Transaction_Date"
       Rsditails.Open sql, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If Not (Rsditails.BOF Or Rsditails.EOF) Then

                    With Me.VSFlexGrid3
                        .Rows = .FixedRows + Rsditails.RecordCount
                       
                  For i = 1 To .Rows - 1
                     .TextMatrix(i, .ColIndex("NoteSerial1")) = IIf(IsNull(Rsditails("NoteSerial1").value), "", Rsditails("NoteSerial1").value)
                     .TextMatrix(i, .ColIndex("Transaction_Date")) = IIf(IsNull(Rsditails("Transaction_Date").value), "", Rsditails("Transaction_Date").value)
                    .TextMatrix(i, .ColIndex("Price")) = IIf(IsNull(Rsditails("Price").value), "", Rsditails("Price").value)
                    .TextMatrix(i, .ColIndex("PODays")) = IIf(IsNull(Rsditails("PODays").value), "", Rsditails("PODays").value)
                    .TextMatrix(i, .ColIndex("CusID")) = IIf(IsNull(Rsditails("CusID").value), "", Rsditails("CusID").value)
                    If SystemOptions.UserInterface = ArabicInterface Then
                    .TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(Rsditails("CusName").value), "", Rsditails("CusName").value)
                    Else
                    .TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(Rsditails("CusNamee").value), "", Rsditails("CusNamee").value)
                    End If
                    Rsditails.MoveNext
                  Next i
                  

End With
End If
End Sub
Private Sub Cmd_Click(Index As Integer)

    Select Case Index

        Case 0
    ReLineGrid
    Unload Me

       Case 24
     '  AddNewFgRowother
       Case 8
          
    End Select

End Sub






Private Sub Form_Activate()
'   PutFormOnTop Me.hWnd
End Sub




Private Sub Form_Load()
    Dim GrdBack As ClsBackGroundPic
    Dim Dcombos As ClsDataCombos
Dim Xpid As Integer

    Set Dcombos = New ClsDataCombos
 
   'Dcombos.GetAccountingCodes Me.DcbAccount
Xpid = val(FrmProcessDef.Grid.TextMatrix(FrmProcessDef.LngRow, FrmProcessDef.Grid.ColIndex("ItemId")))
Retriveshow Xpid
    Set DCboSearch = New clsDCboSearch
   
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Clear").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture


    Set GrdBack = New ClsBackGroundPic


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
  
End Sub

Private Sub ReLineGrid()
    Dim IntCounter As Integer
    IntCounter = 0
    Dim i As Integer

    With Me.VSFlexGrid3

        For i = .FixedRows To .Rows - 1
    
           
               If Me.VSFlexGrid3.Cell(flexcpChecked, i, Me.VSFlexGrid3.ColIndex("Chek")) = flexChecked Then
  FrmProcessDef.Grid.TextMatrix(FrmProcessDef.LngRow, FrmProcessDef.Grid.ColIndex("Price")) = val(Me.VSFlexGrid3.TextMatrix(i, Me.VSFlexGrid3.ColIndex("Price")))
  FrmProcessDef.Grid.TextMatrix(FrmProcessDef.LngRow, FrmProcessDef.Grid.ColIndex("NoteSerial1")) = Me.VSFlexGrid3.TextMatrix(i, Me.VSFlexGrid3.ColIndex("NoteSerial1"))
  End If
     

        Next i
   
    End With

End Sub


