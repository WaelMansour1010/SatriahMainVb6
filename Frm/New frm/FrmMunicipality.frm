VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Begin VB.Form FrmMunicipality 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "‘«‘… «·»·œÌ« "
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6525
   Icon            =   "FrmMunicipality.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   4485
   ScaleWidth      =   6525
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame6 
      Caption         =   "«·»·œÌ« "
      Height          =   3615
      Left            =   -120
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   360
      Width           =   6735
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "«·«„«‰« "
         Height          =   675
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   0
         Width           =   6645
         Begin VB.TextBox TxtName 
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
            Height          =   315
            Left            =   3825
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   8
            Top             =   270
            Width           =   2490
         End
         Begin VB.TextBox TxtNameE 
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
            Height          =   315
            Left            =   1080
            MaxLength       =   50
            RightToLeft     =   -1  'True
            TabIndex        =   7
            Top             =   270
            Width           =   2490
         End
         Begin ImpulseButton.ISButton CmdAdd 
            Height          =   375
            Left            =   240
            TabIndex        =   9
            Top             =   120
            Width           =   630
            _ExtentX        =   1111
            _ExtentY        =   661
            ButtonStyle     =   1
            ButtonPositionImage=   4
            Caption         =   ""
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
            ButtonImage     =   "FrmMunicipality.frx":038A
            ColorHighlight  =   16777215
            ColorHoverText  =   16711680
            ColorShadow     =   -2147483637
            ColorOutline    =   0
            DrawFocusRectangle=   0   'False
            ColorToggledHoverText=   16711680
            LowerToggledContent=   0   'False
            ColorTextShadow =   -2147483637
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "«·«”„ ⁄—»Ì"
            Height          =   195
            Index           =   4
            Left            =   4680
            RightToLeft     =   -1  'True
            TabIndex        =   11
            Top             =   0
            Width           =   990
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "«·«”„ ≈‰Ã·Ì“Ì"
            Height          =   195
            Index           =   6
            Left            =   1680
            RightToLeft     =   -1  'True
            TabIndex        =   10
            Top             =   0
            Width           =   990
         End
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   270
         Index           =   8
         Left            =   5040
         TabIndex        =   5
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
         ButtonImage     =   "FrmMunicipality.frx":0724
         DrawFocusRectangle=   0   'False
      End
      Begin VSFlex8UCtl.VSFlexGrid fg 
         Height          =   2445
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   6405
         _cx             =   11298
         _cy             =   4313
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
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   320
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FrmMunicipality.frx":0CBE
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
      Caption         =   "«·»·œÌ« "
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
      Left            =   -60
      TabIndex        =   3
      Top             =   0
      Width           =   6525
   End
End
Attribute VB_Name = "FrmMunicipality"
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

Private Sub Retrive()
    Dim AccountName As String
    Dim i As Integer
    Dim astrSplit2tems2() As String
    Dim astrSplitItems() As String
    Dim ItemName As String
    Dim J As Integer
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
   
        With Me.fg
        If FrmGovernmentData.fg.TextMatrix(FrmGovernmentData.LngRow, FrmGovernmentData.fg.ColIndex("StrTblMunicipality")) <> "" Then
         st = FrmGovernmentData.fg.TextMatrix(FrmGovernmentData.LngRow, FrmGovernmentData.fg.ColIndex("StrTblMunicipality"))
         st = Trim(st)
         astrSplitItems = Split(st, "@")
         nElements = UBound(astrSplitItems) - LBound(astrSplitItems)
            .Rows = .FixedRows + nElements
            For J = 0 To nElements - 1
            i = J + 1
                astrSplit2tems2 = Split(astrSplitItems(J), "#")
                .TextMatrix(i, .ColIndex("name")) = astrSplit2tems2(0)
                     
                .TextMatrix(i, .ColIndex("namee")) = astrSplit2tems2(1)
                .TextMatrix(i, .ColIndex("id")) = val(astrSplit2tems2(2))
            
                          ' Me.txt_opr_total.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("total"), .Rows - 1, .ColIndex("total"))
        Next J
        End If
        End With
          
   ReLineGrid

End Sub

Sub save()
Dim str As String
Dim i As Integer
Dim MunicID As Double
str = ""

With Me.fg
For i = 1 To .Rows - 1
 If .TextMatrix(i, .ColIndex("name")) <> "" Then
 str = str & Trim(.TextMatrix(i, .ColIndex("name"))) & "#"
 str = str & Trim(.TextMatrix(i, .ColIndex("namee"))) & "#"
               MunicID = val(.TextMatrix(i, .ColIndex("id")))
       If Me.Checked(0, MunicID) = True Then
       Else
       MunicID = 1
       maxx 0, MunicID
       End If
       .TextMatrix(i, .ColIndex("id")) = MunicID
 str = str & Trim(.TextMatrix(i, .ColIndex("id"))) & "#"
 str = str & Trim("@")
'  str = str & Chr(13)
  str = Trim(str)
 End If
Next
FrmGovernmentData.fg.TextMatrix(FrmGovernmentData.LngRow, FrmGovernmentData.fg.ColIndex("StrTblMunicipality")) = str


End With
End Sub
Sub maxx(Optional ByRef AmanhID As Double = 0, Optional ByRef MunicID As Double = 0)
     
  Dim RsDev As ADODB.Recordset
    Dim StrSQL As String
  Set RsDev = New ADODB.Recordset
    If AmanhID <> 0 Then
   StrSQL = " select max(AmanhID) as mx from FoxySerial2"
   RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
   AmanhID = IIf(IsNull(RsDev("mx").value), 0, RsDev("mx").value) + 1
      Set RsDev = New ADODB.Recordset
    RsDev.Open "FoxySerial2", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    RsDev.AddNew
RsDev("AmanhID").value = AmanhID
RsDev.update
End If
    If MunicID <> 0 Then
   StrSQL = " select max(MunicID) as mx from FoxySerial2"
   RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
   MunicID = IIf(IsNull(RsDev("mx").value), 0, RsDev("mx").value) + 1
      Set RsDev = New ADODB.Recordset
    RsDev.Open "FoxySerial2", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    RsDev.AddNew
RsDev("MunicID").value = MunicID
RsDev.update
End If
End Sub
Function Checked(Optional AmanhID As Double = 0, Optional MunicID As Double = 0) As Boolean
     Checked = False
  Dim RsDev As ADODB.Recordset
    Dim StrSQL As String
      Set RsDev = New ADODB.Recordset
    If AmanhID <> 0 Then
   StrSQL = " select * from FoxySerial2 where AmanhID=" & AmanhID & ""
   RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
If RsDev.RecordCount > 0 Then
Checked = True
Else
Checked = False
End If
End If
    If MunicID <> 0 Then
  StrSQL = " select * from FoxySerial2 where MunicID=" & MunicID & ""
   RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
 If RsDev.RecordCount > 0 Then
Checked = True
Else
Checked = False
End If
End If
End Function
Private Sub DeleteFgRowAther()
If FrmGovernmentData.TxtModFlg.text <> "R" Then
    With Me.fg

        If .Row = -1 Then Exit Sub
        If .Row = 0 Then Exit Sub
        .RemoveItem .Row
        ReLineGrid
        '.AutoSize 0, .Cols - 1, False
     
    End With
End If
End Sub

Private Sub cmdAdd_Click()
fillgrid
End Sub

Private Sub Form_Activate()
'   PutFormOnTop Me.hWnd
End Sub
Private Sub ReLineGrid(Optional current_terms As String = "")
    Dim i As Integer
    Dim IntCounter As Integer
        Dim rs As ADODB.Recordset

    With fg

        For i = .FixedRows To .Rows - 1

            If .TextMatrix(i, .ColIndex("name")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("LineNo")) = IntCounter
            End If

        Next i
   
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
If FrmGovernmentData.TxtModFlg.text <> "R" Then
Cmd(0).Enabled = True
Else
Cmd(0).Enabled = False

End If
 Frame6.Visible = True
    Set GrdBack = New ClsBackGroundPic
   ' currentterms = Projects.VSFlexGrid2.TextMatrix(Projects.LngRow, Projects.VSFlexGrid2.ColIndex("fullcode"))
   ' If SystemOptions.UserInterface = ArabicInterface Then
   '                 Frame6.Caption = " „⁄œ«  «·⁄„·ÌÂ  —ﬁ„ : " & currentterms
   '             Else
   '                 Frame6.Caption = "Equipment For Process No: " & currentterms
   '             End If
                
   '    Xpid = val(Projects.txt_project_id.text)
   ' rwOp = val(Projects.VSFlexGrid2.TextMatrix(Projects.LngRow, Projects.VSFlexGrid2.ColIndex("id")))
    
   ' rwpand = val(Projects.VSFlexGrid2.TextMatrix(Projects.LngRow, Projects.VSFlexGrid2.ColIndex("ProjectDes_ID")))

Retrive


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
    
  Me.Caption = "Municipality"
  
Label5.Caption = Me.Caption
'Me.LblClientName.Caption = "ClientName"
'lbl(4).Caption = "From"
'lbl(3).Caption = "To"
'Cmd(24).Caption = "Add"
Cmd(8).Caption = "Delete"

Label1(4).Caption = "Name Arabic"
Label1(6).Caption = "Name English"

'Me.lbreg.Caption = "Date Registration"

     With Me.fg
        .TextMatrix(0, .ColIndex("LineNo")) = "NO"
        .TextMatrix(0, .ColIndex("name")) = "Name Arabic"
        .TextMatrix(0, .ColIndex("namee")) = "Name English"
        ' .TextMatrix(0, .ColIndex("Vlue")) = "Value  "
       ' .TextMatrix(0, .ColIndex("Remark")) = "Remarks  "
       '.TextMatrix(0, .ColIndex("PlateNo")) = "PlateNo"
    End With
  '
End Sub

Sub fillgrid()
Dim i As Integer

i = fg.Rows
With fg

.Rows = .Rows + 1
'.TextMatrix(i, .ColIndex("GovernmentID")) = TxtVac_ID.text
.TextMatrix(i, .ColIndex("name")) = TxtName.text
.TextMatrix(i, .ColIndex("namee")) = TxtNameE.text
TxtName.text = ""
 TxtNameE.text = ""
 ReLineGrid
End With
End Sub
Private Sub TxtName_GotFocus()
SwitchKeyboardLang LANG_ARABIC
End Sub

Private Sub TxtNameE_GotFocus()
SwitchKeyboardLang LANG_ENGLISH
End Sub


