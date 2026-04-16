VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Begin VB.Form FrmAddItemItemInvest 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "≈÷«›… «·ﬁÿ⁄ «·„·Õﬁ…"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9030
   Icon            =   "FrmAddItemItemInvest.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5055
   ScaleWidth      =   9030
   StartUpPosition =   3  'Windows Default
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VSFlex8UCtl.VSFlexGrid Fg 
      Height          =   3525
      Left            =   0
      TabIndex        =   1
      Top             =   810
      Width           =   9015
      _cx             =   15901
      _cy             =   6218
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
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   12
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmAddItemItemInvest.frx":038A
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
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   765
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   30
      Width           =   9015
      Begin VB.CheckBox ChSelectAll 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " ÕœÌœ «·ﬂ·"
         Height          =   375
         Left            =   6720
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Height          =   225
         Index           =   8
         Left            =   1380
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   870
         Width           =   4965
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Height          =   225
         Index           =   6
         Left            =   2520
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   750
         Width           =   2055
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Height          =   225
         Index           =   4
         Left            =   1380
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   510
         Width           =   4965
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Height          =   225
         Index           =   3
         Left            =   5460
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   270
         Width           =   1005
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·ﬁÿ⁄…:"
         Height          =   225
         Index           =   1
         Left            =   11760
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   270
         Width           =   885
      End
   End
   Begin ImpulseButton.ISButton CmdOk 
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   4620
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "„Ê«›ﬁ"
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
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   -2147483637
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      DisabledImageStyle=   1
      ColorToggledHoverText=   16711680
      LowerToggledContent=   0   'False
      ColorTextShadow =   -2147483637
   End
   Begin ImpulseButton.ISButton CmdCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   30
      TabIndex        =   3
      Top             =   4620
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   661
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "≈·€«¡"
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
      ColorButton     =   14871017
      ColorHighlight  =   16777215
      ColorHoverText  =   16711680
      ColorShadow     =   -2147483637
      ColorOutline    =   0
      DrawFocusRectangle=   0   'False
      DisabledImageStyle=   1
      ColorToggledHoverText=   16711680
      LowerToggledContent=   0   'False
      ColorTextShadow =   -2147483637
   End
   Begin VB.Label lblqty 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   375
      Left            =   2400
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   4560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Height          =   225
      Index           =   11
      Left            =   2460
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   4560
      Width           =   1005
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "⁄œœ «·ﬁÿ⁄ «·„Õœœ…:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   225
      Index           =   10
      Left            =   3630
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   4560
      Width           =   1725
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   -30
      X2              =   5400
      Y1              =   4530
      Y2              =   4560
   End
End
Attribute VB_Name = "FrmAddItemItemInvest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Row1 As Integer
Public DivIDDet As Double
Public SalID As Double
Public InvID As Double
Private m_UserCancelled As Boolean
Public TypDiv As Integer


Private Sub ChSelectAll_Click()
  Dim i As Integer

    If ChSelectAll.value = vbChecked Then

        With Me.FG
 
            For i = 1 To .Rows - 1
        
                .TextMatrix(i, .ColIndex("Select")) = True
                
            Next i

        End With

    Else

        With Me.FG

            For i = 1 To .Rows - 1
        
                .TextMatrix(i, .ColIndex("Select")) = False
            Next i

        End With

    End If
  Me.lbl(11).Caption = ModFgLib.GetFgCheckCount(FG, FG.ColIndex("Select"))
  '  Me.lbl(14).Caption = Format(val(Calculate_TotalSelected), "#,###." & String(Abs(SystemOptions.SysDefCurrencyForamt), "0"))
'Me.lbl(14).Caption = val(Calculate_TotalSelected)
End Sub

Private Sub CmdCancel_Click()
    Me.Hide
    UserCancelled = True
End Sub

Private Sub CmdOk_Click()
    Dim Msg As String

    If val(Me.lbl(11).Caption) = 0 Then
        Msg = "ÌÃ» ≈Œ Ì«— ’‰› Ê«Õœ ⁄·Ï «·√ﬁ·...!!!"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    End If
sve
    Me.Hide
    UserCancelled = False
End Sub
Sub sve()
Dim i As Integer
Dim k As Long
With FrmSaleBillInvestment.GridInstallments
k = Row1
For i = FG.FixedRows To FG.Rows - 1
If FG.Cell(flexcpChecked, i, FG.ColIndex("Select")) = flexChecked Then
If CheNotRepaet(val(FG.TextMatrix(i, FG.ColIndex("ID")))) = False Then
.TextMatrix(k, .ColIndex("InvesName")) = .TextMatrix(Row1, .ColIndex("InvesName"))
.TextMatrix(k, .ColIndex("InvesID")) = .TextMatrix(Row1, .ColIndex("InvesID"))
.TextMatrix(k, .ColIndex("unitId")) = .TextMatrix(Row1, .ColIndex("unitId"))
.TextMatrix(k, .ColIndex("unit")) = .TextMatrix(Row1, .ColIndex("unit"))

.TextMatrix(k, .ColIndex("Area")) = FG.TextMatrix(i, FG.ColIndex("Area"))
.TextMatrix(k, .ColIndex("BlockName")) = FG.TextMatrix(i, FG.ColIndex("BlokNo"))
.TextMatrix(k, .ColIndex("BlockID")) = FG.TextMatrix(i, FG.ColIndex("ID"))
.TextMatrix(k, .ColIndex("PartName")) = FG.TextMatrix(i, FG.ColIndex("PartNo"))
.TextMatrix(k, .ColIndex("MeterValue")) = .TextMatrix(Row1, .ColIndex("MeterValue"))
.TextMatrix(k, .ColIndex("CodeUnit")) = FG.TextMatrix(i, FG.ColIndex("CodeUnit"))
.TextMatrix(k, .ColIndex("unitunidpart")) = FG.TextMatrix(i, FG.ColIndex("unitunid"))

.Rows = .Rows + 1
k = k + 1
End If
End If
Next i
End With
End Sub
Function CheNotRepaet(Optional BlokID As Long) As Boolean
Dim i As Integer
CheNotRepaet = False
With FrmSaleBillInvestment.GridInstallments
For i = .FixedRows To .Rows - 1
If val(.TextMatrix(i, .ColIndex("BlockID"))) <> 0 Then
If val(.TextMatrix(i, .ColIndex("BlockID"))) = BlokID Then
CheNotRepaet = True
Exit Function
End If
End If
Next i
End With
End Function
Private Sub FG_AfterEdit(ByVal Row As Long, _
                         ByVal Col As Long)

    With Me.FG

        Select Case .ColKey(Col)

            Case "Select"
                Me.lbl(11).Caption = ModFgLib.GetFgCheckCount(FG, FG.ColIndex("Select"))
        End Select

    End With

End Sub

Private Sub FG_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

    With Me.FG

        Select Case FG.ColKey(Col)

            Case "Select"
                Cancel = False

            Case Else
                Cancel = True
        End Select

    End With
End Sub

Private Sub fg_StartEdit(ByVal Row As Long, _
                         ByVal Col As Long, _
                         Cancel As Boolean)


End Sub

Private Sub Form_Load()
    Dim GrdBack As ClsBackGroundPic
    Set GrdBack = New ClsBackGroundPic
    CenterForm Me

    FormPostion Me, GetPostion
LoadItemData DivIDDet, SalID, InvID
   ' With Me.FG
       ' .Rows = .FixedRows
       ' .ExtendLastCol = True
       ' .RowHeightMin = 300
       ' .Editable = flexEDKbdMouse
       ' .ExplorerBar = flexExSortShowAndMove
       ' Set .WallPaper = GrdBack.Picture
       ' .AutoSize 0, .Cols - 1, False
  '  End With

 
End Sub
Function RetnTypeUnite(Optional DivID As Double) As Boolean
Dim Rs3 As ADODB.Recordset
Set Rs3 = New ADODB.Recordset
Dim sql As String
sql = "select * from TblSpreading  where ID =" & DivID & " and UnitMain=1 "
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
RetnTypeUnite = True
Else
RetnTypeUnite = False
End If
End Function
Public Sub LoadItemData(DivIDDet As Double, Optional SalID As Double, Optional InvID As Double)
    Dim i As Long
    Dim rs As ADODB.Recordset
    Dim StrSQL As String
FG.Clear flexClearScrollable, flexClearEverything
        With Me.FG

            For i = 1 To .Rows - 1
        
                .TextMatrix(i, .ColIndex("Select")) = False
            Next i

        End With

StrSQL = " SELECT     dbo.TblDivInvestInformation.ID, dbo.TblDivInvestInformation.Area, dbo.TblDivInvestInformation.CodeUnit, dbo.TblDivInvestInformation.TypeDivi,"
StrSQL = StrSQL & "                      dbo.TblDivInvestInformation.DivMainID, dbo.TblDivInvestInformation.unitunid, dbo.TblDivInvestInformation.PartNo, dbo.TblDivInvestInformation.BlokNo,"
StrSQL = StrSQL & "                      dbo.TblSpreading.name , dbo.TblSpreading.NameE, dbo.TblDivInvestInformation.SalID ,dbo.TblDivInvestInformation.InvID"
StrSQL = StrSQL & " FROM         dbo.TblDivInvestInformation LEFT OUTER JOIN"
StrSQL = StrSQL & "                       dbo.TblSpreading ON dbo.TblDivInvestInformation.unitunid = dbo.TblSpreading.ID"
StrSQL = StrSQL & " Where (dbo.TblDivInvestInformation.SalesBlocPayed Is Null)AND (dbo.TblDivInvestInformation.EffectID=1) And (dbo.TblDivInvestInformation.SalesBlocPayed Is Null)and dbo.TblDivInvestInformation.InvID=" & InvID & " "
'StrSQL = StrSQL & " AND dbo.TblDivInvestInformation.TypDiv =" & TypDiv & ""
If RetnTypeUnite(DivIDDet) = True Then
StrSQL = StrSQL & " and DivMainID =" & DivIDDet & ""
FG.ColHidden(FG.ColIndex("CodeUnit")) = False
Else
StrSQL = StrSQL & " and TypeDivi =" & DivIDDet & ""
FG.ColHidden(FG.ColIndex("CodeUnit")) = True
End If

 If FrmSaleBillInvestment.TxtModFlg.Text = "E" Then
 If RetnTypeUnite(DivIDDet) = True Then
    StrSQL = StrSQL & " or (SalID=" & SalID & " and (DivMainID=" & DivIDDet & " ) ) "
    Else
    StrSQL = StrSQL & " or (SalID=" & SalID & " and (TypeDivi=" & DivIDDet & " ) ) "
    End If
 End If
            

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.RecordCount > 0 Then
        rs.MoveFirst
        Me.lbl(3).Caption = rs("BlokNo").value
       ' Me.lbl(6).Caption = rs("ItemCode").value
       ' Me.lbl(4).Caption = rs("BlokNo").value
       ' Me.lbl(9).Caption = rs.RecordCount

        With Me.FG
            .Rows = .FixedRows + rs.RecordCount

            For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("ID")) = IIf(IsNull(rs("ID").value), "", rs("ID").value)
                .TextMatrix(i, .ColIndex("BlokNo")) = IIf(IsNull(rs("BlokNo").value), "", rs("BlokNo").value)
               .TextMatrix(i, .ColIndex("Area")) = IIf(IsNull(rs("Area").value), "", rs("Area").value)
                .TextMatrix(i, .ColIndex("CodeUnit")) = IIf(IsNull(rs("CodeUnit").value), "", rs("CodeUnit").value)
                .TextMatrix(i, .ColIndex("PartNo")) = IIf(IsNull(rs("PartNo").value), "", rs("PartNo").value)
                .TextMatrix(i, .ColIndex("unitunid")) = IIf(IsNull(rs("unitunid").value), 0, rs("unitunid").value)
                
                If SystemOptions.UserInterface = ArabicInterface Then
                .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs("Name").value), "", rs("Name").value)
                Else
                .TextMatrix(i, .ColIndex("Name")) = IIf(IsNull(rs("NameE").value), "", rs("NameE").value)
                End If
                '.TextMatrix(I, .ColIndex("")) = IIf(IsNull(Rs("").Value), "", Rs("").Value)
                rs.MoveNext
            Next i

            '.AutoSize 0, .Cols - 1, False
        End With

    End If

    rs.Close
    Set rs = Nothing
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)

    If UnloadMode = VBRUN.QueryUnloadConstants.vbFormControlMenu Or UnloadMode = VBRUN.QueryUnloadConstants.vbAppTaskManager Then
        Me.Hide
        UserCancelled = True
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    FormPostion Me, SavePostion
End Sub

Public Property Get UserCancelled() As Boolean
    UserCancelled = m_UserCancelled
End Property

Public Property Let UserCancelled(ByVal vNewValue As Boolean)
    m_UserCancelled = vNewValue
End Property
