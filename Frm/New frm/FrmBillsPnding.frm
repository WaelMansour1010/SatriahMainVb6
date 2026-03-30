VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form FrmBillsPnding 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "«·›Ê« Ì— «·„⁄·ﬁ…"
   ClientHeight    =   6915
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9090
   Icon            =   "FrmBillsPnding.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6915
   ScaleWidth      =   9090
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FraHeader 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   540
      Left            =   15
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   8835
      Begin VB.CommandButton btnSearch 
         Caption         =   "»ÕÀ"
         Height          =   375
         Left            =   450
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   30
         Width           =   450
      End
      Begin VB.TextBox TxtNotID 
         Alignment       =   1  'Right Justify
         Height          =   405
         Left            =   930
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   0
         Width           =   2415
      End
      Begin VB.TextBox TxtVac_ID 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         Height          =   240
         Left            =   3030
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   510
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.TextBox TxtModFlg 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         Enabled         =   0   'False
         Height          =   285
         Left            =   3360
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Text            =   "modflag"
         Top             =   90
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Frame Frmo2 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   375
         Left            =   540
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   450
         Visible         =   0   'False
         Width           =   3105
         Begin MSDataListLib.DataCombo DCUser 
            CausesValidation=   0   'False
            Height          =   315
            Left            =   -255
            TabIndex        =   2
            Tag             =   "„‰ ›÷·ﬂ √œŒ· —ﬁ„ «·ﬁ÷Ì…"
            Top             =   15
            Width           =   2340
            _ExtentX        =   4128
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            BackColor       =   -2147483624
            Text            =   ""
            RightToLeft     =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            Caption         =   "«·„” Œœ„"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   13
            Left            =   2160
            RightToLeft     =   -1  'True
            TabIndex        =   3
            Top             =   45
            Width           =   855
         End
      End
      Begin MSComctlLib.ImageList GrdImageList 
         Left            =   3960
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   8
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmBillsPnding.frx":058A
               Key             =   "CompanyName"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmBillsPnding.frx":0924
               Key             =   "Ser"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmBillsPnding.frx":0CBE
               Key             =   "Vac_Name"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmBillsPnding.frx":1058
               Key             =   "ShareCount"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmBillsPnding.frx":13F2
               Key             =   "Dis_Count"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmBillsPnding.frx":178C
               Key             =   "Bouns"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmBillsPnding.frx":1B26
               Key             =   "SharesValue"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmBillsPnding.frx":20C0
               Key             =   "BuyValue"
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "”‰œ«  "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   0
         Left            =   720
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   480
         Width           =   3360
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«·›Ê« Ì— «·„⁄·ﬁ…"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   2
         Left            =   3165
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   90
         Width           =   3360
      End
   End
   Begin C1SizerLibCtl.C1Elastic EltCont 
      Height          =   420
      Left            =   0
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   6480
      Width           =   5400
      _cx             =   9525
      _cy             =   741
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   0
      MousePointer    =   0
      Version         =   801
      BackColor       =   14871017
      ForeColor       =   -2147483630
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   0
      AutoSizeChildren=   0
      BorderWidth     =   1
      ChildSpacing    =   1
      Splitter        =   0   'False
      FloodDirection  =   0
      FloodPercent    =   0
      CaptionPos      =   1
      WordWrap        =   -1  'True
      MaxChildSize    =   0
      MinChildSize    =   0
      TagWidth        =   0
      TagPosition     =   0
      Style           =   0
      TagSplit        =   2
      PicturePos      =   4
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
      Begin ImpulseButton.ISButton btnCancel 
         Height          =   330
         Left            =   2760
         TabIndex        =   8
         Top             =   75
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   582
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
         ButtonImage     =   "FrmBillsPnding.frx":245A
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
      Begin ImpulseButton.ISButton btnDelete 
         Height          =   330
         Left            =   240
         TabIndex        =   13
         Top             =   75
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         ButtonStyle     =   1
         ButtonPositionImage=   1
         Caption         =   " ‰›Ì– «·Õ–›"
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
         ButtonImage     =   "FrmBillsPnding.frx":27F4
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   5805
      Left            =   0
      TabIndex        =   9
      Top             =   630
      Width           =   9045
      _cx             =   15954
      _cy             =   10239
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
      Rows            =   50
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   320
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmBillsPnding.frx":2D8E
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
      Begin VB.TextBox TxtContNo 
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
         Left            =   0
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   -240
         Visible         =   0   'False
         Width           =   1065
      End
   End
End
Attribute VB_Name = "FrmBillsPnding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsSavRec As ADODB.Recordset
Dim BKGrndPic As ClsBackGroundPic
Dim RecId As String
Dim ii As Long
Public Indxx As Integer
Public mTypeInvoice As Integer
Dim firstSerachRow As Integer
Private Sub BtnCancel_Click()
End
End Sub

Private Sub btnDelete_Click()
Dim i As Integer
With Me.Grid
For i = 1 To .rows - 1
If .cell(flexcpChecked, i, .ColIndex("Selc")) = flexChecked Then
Cn.Execute " delete   From dbo.Transactions  Where (Transaction_Type = 70) And (Transaction_ID = " & val(.TextMatrix(i, .ColIndex("id"))) & ")"
End If
Next i
End With
Unload Me
End Sub

Private Sub Form_Load()
    On Error GoTo ErrTrap
  Dim i As Integer



     If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
FillGridWithData

ErrTrap:
End Sub
Private Sub ChangeLang()
Label1(2).Caption = "Pending Bills"
Me.Caption = Label1(2).Caption
With Grid
.TextMatrix(0, .ColIndex("Ser")) = "Serial"
.TextMatrix(0, .ColIndex("NoteDate")) = "Date"
.TextMatrix(0, .ColIndex("NoteSerial1")) = "Time"
.TextMatrix(0, .ColIndex("show")) = "Show"
.TextMatrix(0, .ColIndex("Selc")) = "Select"
.TextMatrix(0, .ColIndex("PlateNo")) = "PlateNo"
.TextMatrix(0, .ColIndex("CashCustomerName")) = "Customer Name"

End With
btnCancel.Caption = "Exit"
btnDelete.Caption = "Delete"
    End Sub
Private Sub Form_Terminate()
    'Set cSearchDCombo = Nothing
    'Set BKGrndPic = Nothing
    Set FrmVacancy = Nothing

End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap

    If RsSavRec.State = adStateOpen Then
        If Not (RsSavRec.EOF Or RsSavRec.BOF) Then
            If RsSavRec.EditMode <> adEditNone Then
                RsSavRec.CancelUpdate
            End If
        End If

        RsSavRec.Close
        Set RsSavRec = Nothing
    End If

ErrTrap:
End Sub



Private Sub Form_Activate()
    Me.ZOrder 0
End Sub
Public Sub FillGridWithData()

    On Error GoTo ErrTrap

    Dim i As Integer
    Dim rs As ADODB.Recordset
    Dim My_SQL As String

    Set rs = New ADODB.Recordset
    My_SQL = " SELECT     TOP 100 PERCENT Transaction_ID, Transaction_Date ,RecTime,CashCustomerName,PlateNo"
    My_SQL = My_SQL & "     From dbo.Transactions"
    My_SQL = My_SQL & "          Where  IsNull(TypeInvoice,0) = " & mTypeInvoice & " and (Transaction_Type = 70) and UserID=" & user_id & ""
    My_SQL = My_SQL & "    ORDER BY Transaction_ID"

    rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText

    With Me.Grid
        .rows = 2
        .Clear flexClearScrollable

        If rs.RecordCount > 0 Then
            .rows = rs.RecordCount + 1
            rs.MoveFirst

            For i = 1 To .rows - 1
                .TextMatrix(i, .ColIndex("Ser")) = i
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs.Fields("Transaction_ID").value), 0, rs.Fields("Transaction_ID").value)
                .TextMatrix(i, .ColIndex("NoteDate")) = IIf(IsNull(rs.Fields("Transaction_Date").value), "", rs.Fields("Transaction_Date").value)
                .TextMatrix(i, .ColIndex("NoteSerial1")) = IIf(IsNull(rs.Fields("RecTime").value), "", rs.Fields("RecTime").value)
                
                .TextMatrix(i, .ColIndex("CashCustomerName")) = IIf(IsNull(rs.Fields("CashCustomerName").value), "", rs.Fields("CashCustomerName").value)
                .TextMatrix(i, .ColIndex("PlateNo")) = IIf(IsNull(rs.Fields("PlateNo").value), "", rs.Fields("PlateNo").value)
            
                rs.MoveNext
            Next

            rs.Close
        End If

        .RowHeight(-1) = 300
    End With

ErrTrap:
End Sub


Private Sub Grid_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
With Me.Grid
Select Case .ColKey(Col)
Case "show"
  If checkApility("frmsalebill") = False Then
                Exit Sub
            End If
If Indxx = 1 Then
 ' If checkApility("frmsalebill") = False Then
 '               Exit Sub
 '           End If
Load frmsalebill
'frmsalebill.show
frmsalebill.RetrivePending val(.TextMatrix(Row, .ColIndex("id")))
ElseIf Indxx = 2 Then
 ' If checkApility("frmsalebill2") = False Then
 '               Exit Sub
 '           End If
Load frmsalebill2
'frmsalebill2.show
frmsalebill2.RetrivePending val(.TextMatrix(Row, .ColIndex("id")))

ElseIf Indxx = 3 Then
 ' If checkApility("frmsalebill2") = False Then
 '               Exit Sub
 '           End If
Load frmsalebill6
'frmsalebill2.show
frmsalebill6.RetrivePending val(.TextMatrix(Row, .ColIndex("id")))

End If
Unload Me
End Select
End With
End Sub

Private Sub Grid_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With Me.Grid
Select Case .ColKey(Col)
 Case "show"
            .ColComboList(.ColIndex("show")) = "..."
     End Select
    End With
    
End Sub

Private Sub TxtNotID_Change()
'Grid.FindRow
btnSearch_Click
End Sub

Private Sub btnSearch_Click()
    Dim i
    Dim accName As String, CusName As String
    If firstSerachRow = Grid.rows Or firstSerachRow = Grid.rows - 1 Then
        firstSerachRow = 1
    End If
    
  ' —Ã¯⁄ «·ŒÿÊÿ “Ì «·√Ê·
    For i = 1 To Grid.rows - 1
        Grid.cell(flexcpFontUnderline, i, Grid.ColIndex("PlateNo")) = False
        Grid.cell(flexcpFontUnderline, i, Grid.ColIndex("CashCustomerName")) = False
    Next
    
    ' «·»ÕÀ
    For i = firstSerachRow To Grid.rows - 1
        accName = Grid.TextMatrix(i, Grid.ColIndex("PlateNo"))
        CusName = Grid.TextMatrix(i, Grid.ColIndex("CashCustomerName"))
        
        If accName Like "*" & TxtNotID & "*" Or CusName Like "*" & TxtNotID & "*" Then
            Grid.ShowCell i, Grid.ColIndex("PlateNo")
            Grid.cell(flexcpFontUnderline, i, Grid.ColIndex("PlateNo")) = True
            Grid.cell(flexcpFontUnderline, i, Grid.ColIndex("CashCustomerName")) = True
            
            ' Õœœ «·’›
            Grid.Row = i
            Grid.Col = Grid.ColIndex("PlateNo")
            
            firstSerachRow = i + 1
            Exit Sub
        End If
    Next
End Sub

