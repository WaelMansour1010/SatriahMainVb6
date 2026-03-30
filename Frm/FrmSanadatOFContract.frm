VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmSanadatOFContract 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6915
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9840
   Icon            =   "FrmSanadatOFContract.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6915
   ScaleWidth      =   9840
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FraHeader 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   540
      Left            =   15
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   9825
      Begin VB.TextBox TxtNotID 
         Alignment       =   1  'Right Justify
         Height          =   405
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   0
         Visible         =   0   'False
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
         Left            =   2580
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
         Left            =   3120
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
               Picture         =   "FrmSanadatOFContract.frx":058A
               Key             =   "CompanyName"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmSanadatOFContract.frx":0924
               Key             =   "Ser"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmSanadatOFContract.frx":0CBE
               Key             =   "Vac_Name"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmSanadatOFContract.frx":1058
               Key             =   "ShareCount"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmSanadatOFContract.frx":13F2
               Key             =   "Dis_Count"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmSanadatOFContract.frx":178C
               Key             =   "Bouns"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmSanadatOFContract.frx":1B26
               Key             =   "SharesValue"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmSanadatOFContract.frx":20C0
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
         Top             =   120
         Width           =   3360
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "”‰œ«  «·ﬁ»÷ «·Œ«’… »«·⁄ﬁœ —ﬁ„"
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
         Left            =   5535
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   120
         Width           =   3720
      End
   End
   Begin C1SizerLibCtl.C1Elastic EltCont 
      Height          =   420
      Left            =   0
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   6480
      Width           =   9720
      _cx             =   17145
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
         Left            =   105
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
         ButtonImage     =   "FrmSanadatOFContract.frx":245A
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid 
      Height          =   5835
      Left            =   0
      TabIndex        =   9
      Top             =   570
      Width           =   9825
      _cx             =   17330
      _cy             =   10292
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
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   320
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmSanadatOFContract.frx":27F4
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
Attribute VB_Name = "FrmSanadatOFContract"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsSavRec As ADODB.Recordset
Dim BKGrndPic As ClsBackGroundPic
Dim RecId As String
Dim II As Long
Public Indx As Integer
Private Sub BtnCancel_Click()
 Unload Me
End Sub

Private Sub Form_Load()
    On Error GoTo ErrTrap
  Dim i As Integer

   ' With Me.Grid
   '     .Cell(flexcpPicture, 0, .ColIndex("Name")) = Me.GrdImageList.ListImages("Vac_Name").ExtractIcon
   '     .Cell(flexcpPicture, 0, .ColIndex("Ser")) = Me.GrdImageList.ListImages("Ser").ExtractIcon

      '  For i = 0 To .Cols - 1
      '      .Cell(flexcpPictureAlignment, 0, i) = flexPicAlignRightCenter
      '  Next
   
   '     .ExtendLastCol = True
   '     .WallPaper = BKGrndPic.Picture
   '     .RowHeight(-1) = 300
   ' End With


    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
FillGridWithData
ErrTrap:
End Sub

Function ChangeLang()

 

End Function



Private Sub Form_Terminate()
    'Set cSearchDCombo = Nothing
    'Set BKGrndPic = Nothing
    Set FrmVacancy = Nothing

End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrTrap
    RegisterLogInOut Me.Name, ScreenNameArabic, ScreenNameEnglish

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
If Indx = 1 Then
Label1(2).Caption = "”‰œ«  «·ﬁ»÷ «·Œ«’… »«· ’›Ì«  Ê›Ê« Ì— «·ﬂÂ—«»«¡ —ﬁ„"
Else
Label1(2).Caption = "”‰œ«  «·ﬁ»÷ «·Œ«’… »«·⁄ﬁœ —ﬁ„"
End If
    Me.ZOrder 0
End Sub
Public Sub FillGridWithData()

    On Error GoTo ErrTrap

    Dim i As Integer
    Dim rs As ADODB.Recordset
    Dim My_SQL As String

    Set rs = New ADODB.Recordset
    My_SQL = "SELECT   dbo.Notes.VAT ,Note_Value2 = IsNull(dbo.Notes.Note_Value2,dbo.Notes.Note_Value) ,   dbo.Notes.NoteID, dbo.Notes.NoteDate,dbo.Notes.renterName, dbo.Notes.NoteType, dbo.Notes.NoteSerial1, dbo.Notes.ContNo, dbo.Notes.ContractNo, "
    My_SQL = My_SQL & "                   dbo.Notes.Note_Value , dbo.Notes.NoteDateH, dbo.Notes.CusID, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.fullcode"
    My_SQL = My_SQL & "  FROM         dbo.Notes LEFT OUTER JOIN"
    My_SQL = My_SQL & "                   dbo.TblCustemers ON dbo.Notes.CusID = dbo.TblCustemers.CusID"
   My_SQL = My_SQL & "  WHERE     (dbo.Notes.NoteType = 4) AND ((dbo.Notes.ContNo = " & val(TxtContNo.Text) & ")or (dbo.Notes.NoteID= " & val(TxtNotID.Text) & "))"
If Indx = 1 Then
My_SQL = My_SQL & "  and      (dbo.Notes.CashingType = 13) "
Else
My_SQL = My_SQL & "  and      (dbo.Notes.CashingType = 8 or dbo.Notes.CashingType = 9) "
End If
My_SQL = My_SQL & " order by dbo.Notes.NoteDate"
    rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText

    With Me.Grid
        .Rows = 2
        .Clear flexClearScrollable

        If rs.RecordCount > 0 Then
            .Rows = rs.RecordCount + 1
            rs.MoveFirst

            For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("Ser")) = i
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs.Fields("NoteID").value), 0, rs.Fields("NoteID").value)
             .TextMatrix(i, .ColIndex("ContractNo")) = IIf(IsNull(rs.Fields("ContractNo").value), "", rs.Fields("ContractNo").value)
             .TextMatrix(i, .ColIndex("ContNo")) = IIf(IsNull(rs.Fields("ContNo").value), 0, rs.Fields("ContNo").value)
                .TextMatrix(i, .ColIndex("NoteDate")) = IIf(IsNull(rs.Fields("NoteDate").value), "", rs.Fields("NoteDate").value)
                .TextMatrix(i, .ColIndex("NoteSerial1")) = IIf(IsNull(rs.Fields("NoteSerial1").value), "", rs.Fields("NoteSerial1").value)
             
                .TextMatrix(i, .ColIndex("NoteDateH")) = IIf(IsNull(rs.Fields("NoteDateH").value), "", rs.Fields("NoteDateH").value)
           
                .TextMatrix(i, .ColIndex("NoteSerial1")) = IIf(IsNull(rs.Fields("NoteSerial1").value), "", rs.Fields("NoteSerial1").value)
            
                .TextMatrix(i, .ColIndex("Note_Value")) = IIf(IsNull(rs.Fields("Note_Value2").value), 0, rs.Fields("Note_Value2").value) + IIf(IsNull(rs.Fields("VAT").value), 0, rs.Fields("VAT").value)
            If SystemOptions.UserInterface = ArabicInterface Then
        
                .TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(rs.Fields("CusName").value), IIf(IsNull(rs.Fields("renterName").value), "", rs.Fields("renterName").value), rs.Fields("CusName").value)
                Else
                .TextMatrix(i, .ColIndex("CusName")) = IIf(IsNull(rs.Fields("CusNamee").value), IIf(IsNull(rs.Fields("renterName").value), "", rs.Fields("renterName").value), rs.Fields("CusNamee").value)
                End If
            
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
  If checkApility("FrmCashing1") = False Then
                Exit Sub
            End If
 Unload FrmCashing1
Load FrmCashing1
FrmCashing1.show
FrmCashing1.RereivID = val(.TextMatrix(Row, .ColIndex("id")))
FrmCashing1.XPBtnMove_Click (2)
FrmCashing1.Retrive val(.TextMatrix(Row, .ColIndex("id")))
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

Private Sub TxtContNo_Change()
FillGridWithData
End Sub
