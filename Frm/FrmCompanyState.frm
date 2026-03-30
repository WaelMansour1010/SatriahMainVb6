VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Begin VB.Form FrmCompanyState 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "«·„Êﬁ› «·Õ«·Ï ··‘—ﬂ…"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7920
   Icon            =   "FrmCompanyState.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6120
   ScaleWidth      =   7920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin ImpulseButton.ISButton Cmd 
      Height          =   435
      Left            =   180
      TabIndex        =   1
      Top             =   510
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   767
      Caption         =   "»œ¡ ⁄„·Ì… «·√” ⁄·«„"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VSFlex8UCtl.VSFlexGrid Fg 
      Height          =   4155
      Left            =   210
      TabIndex        =   0
      Top             =   960
      Width           =   7695
      _cx             =   13573
      _cy             =   7329
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
      Rows            =   50
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   320
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmCompanyState.frx":038A
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
      Editable        =   0
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
Attribute VB_Name = "FrmCompanyState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Cmd_Click()
    Dim i As Integer

    With Me.FG
        i = 1
        .Rows = 10
        .TextMatrix(i, .ColIndex("Serial")) = i
        .TextMatrix(i, .ColIndex("Data")) = "≈Ã„«·Ï ‰ﬁœÌ… „ÊÃÊœ… ›Ï «·Œ“‰"
        i = i + 1
        .TextMatrix(i, .ColIndex("Serial")) = i
        .TextMatrix(i, .ColIndex("Data")) = "≈Ã„«·Ï œÌÊ‰ ··‘—ﬂ…"
    
        i = i + 1
        .TextMatrix(i, .ColIndex("Serial")) = i
        .TextMatrix(i, .ColIndex("Data")) = "≈Ã„«·Ï œÌÊ‰ ⁄·Ï «·‘—ﬂ…"
    
        i = i + 1
        .TextMatrix(i, .ColIndex("Serial")) = i
        .TextMatrix(i, .ColIndex("Data")) = "≈Ã„«·Ï ‘Ìﬂ«  ··‘—ﬂ… €Ì— „Õ’·…"
        .TextMatrix(i, .ColIndex("Value")) = GetChecksNotes(2)
    
        i = i + 1
        .TextMatrix(i, .ColIndex("Serial")) = i
        .TextMatrix(i, .ColIndex("Data")) = "≈Ã„«·Ï ‘Ìﬂ«  ··‘—ﬂ… €Ì— „”œœœ…"
        .TextMatrix(i, .ColIndex("Value")) = GetChecksNotes(13)
    
        i = i + 1
        .TextMatrix(i, .ColIndex("Serial")) = i
        .TextMatrix(i, .ColIndex("Data")) = "≈Ã„«·Ï ﬁÌ„… „Œ“Ê‰ »÷«⁄… ›Ï „Œ«“‰ «·‘—ﬂ…"
    
        .AutoSize 0, .Cols - 1, False
    End With

End Sub

Private Sub Form_Load()
    CenterForm Me
End Sub

Private Function GetChecksNotes(IntNoteType As Integer) As Double
    Dim rs As ADODB.Recordset
    Dim StrSQL As String
    Dim DblTemp As Double

    StrSQL = "SELECT Sum(Notes.Note_Value) AS SumNote_Value"
    StrSQL = StrSQL + " From Notes "
    StrSQL = StrSQL + " where  (Notes.NoteType=" & IntNoteType & ") AND " & "Notes.NoteID NOT IN(Select NoteID From TblCheckRelease)"
    StrSQL = StrSQL + " GROUP BY Notes.NoteType"

    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If Not (rs.BOF Or rs.EOF) Then
        DblTemp = IIf(IsNull(rs("SumNote_Value").value), 0, rs("SumNote_Value").value)
    Else
        DblTemp = 0
    End If

    rs.Close
    Set rs = Nothing
    GetChecksNotes = DblTemp

End Function
