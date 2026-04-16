VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form FrmProjectAlarm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ‰»ÌÂ«  ⁄„·Ì«  «·„‘«—Ì⁄   "
   ClientHeight    =   7740
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14535
   Icon            =   "FrmProjectAlarm.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7740
   ScaleWidth      =   14535
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.Frame Frame3 
      BackColor       =   &H00E2E9E9&
      Height          =   855
      Left            =   0
      TabIndex        =   3
      Top             =   6840
      Width           =   14535
      Begin ImpulseButton.ISButton Cmd 
         Height          =   495
         Index           =   6
         Left            =   1080
         TabIndex        =   4
         Top             =   240
         Width           =   5325
         _ExtentX        =   9393
         _ExtentY        =   873
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
         ButtonImage     =   "FrmProjectAlarm.frx":6852
         ColorButton     =   14871017
         ColorHighlight  =   16777215
         ColorHoverText  =   16711680
         ColorShadow     =   -2147483637
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         ColorToggledHoverText=   16711680
         ColorTextShadow =   -2147483637
      End
      Begin ImpulseButton.ISButton CmdHelp 
         Height          =   495
         Left            =   8160
         TabIndex        =   5
         Top             =   240
         Width           =   4995
         _ExtentX        =   8811
         _ExtentY        =   873
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
         ButtonImage     =   "FrmProjectAlarm.frx":30474
         ColorButton     =   14871017
         ColorHighlight  =   16777215
         ColorHoverText  =   16711680
         ColorShadow     =   -2147483637
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         ColorToggledHoverText=   16711680
         ColorTextShadow =   -2147483637
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2E9E9&
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   14535
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6120
         TabIndex        =   8
         Top             =   240
         Width           =   810
      End
      Begin MSDataListLib.DataCombo dcproject 
         Height          =   315
         Left            =   8040
         TabIndex        =   7
         Top             =   240
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "6"
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin ImpulseButton.ISButton Cmd 
         Height          =   495
         Index           =   5
         Left            =   120
         TabIndex        =   10
         Top             =   120
         Width           =   5325
         _ExtentX        =   9393
         _ExtentY        =   873
         ButtonPositionImage=   1
         Caption         =   "»ÕÀ"
         BackColor       =   14871017
         FontSize        =   14.25
         FontName        =   "Times New Roman"
         FontBold        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonImage     =   "FrmProjectAlarm.frx":36CD6
         ColorButton     =   14871017
         ColorHighlight  =   16777215
         ColorHoverText  =   12632064
         ColorShadow     =   -2147483637
         ColorOutline    =   0
         DrawFocusRectangle=   0   'False
         ColorToggledHoverText=   12632064
         LowerToggledContent=   0   'False
         ColorTextShadow =   -2147483637
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   " ÕœÌÀ ﬂ·"
         Height          =   435
         Index           =   4
         Left            =   6960
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   240
         Width           =   780
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·„‘—Ê⁄"
         Height          =   285
         Index           =   3
         Left            =   13440
         TabIndex        =   6
         Top             =   240
         Width           =   1005
      End
   End
   Begin C1SizerLibCtl.C1Elastic EleHeader 
      Height          =   585
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   14565
      _cx             =   25691
      _cy             =   1032
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   178
         Weight          =   700
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
      Caption         =   " ‰»ÌÂ«  ⁄„·Ì«  «·„‘«—Ì⁄                "
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
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H000000FF&
         Height          =   555
         Index           =   27
         Left            =   2520
         TabIndex        =   2
         Top             =   0
         Width           =   2205
      End
      Begin VB.Image Image1 
         Height          =   555
         Index           =   0
         Left            =   6240
         Picture         =   "FrmProjectAlarm.frx":3D538
         Stretch         =   -1  'True
         Top             =   0
         Width           =   795
      End
   End
   Begin VSFlex8UCtl.VSFlexGrid GridInstallments 
      Height          =   3015
      Left            =   0
      TabIndex        =   11
      Top             =   1320
      Width           =   14475
      _cx             =   25532
      _cy             =   5318
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
      BackColorAlternate=   16777152
      GridColor       =   0
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   13
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmProjectAlarm.frx":3E8DE
      ScrollTrack     =   -1  'True
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
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   630
         Left            =   1920
         TabIndex        =   12
         Top             =   960
         Visible         =   0   'False
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   1111
         _Version        =   393216
         Appearance      =   0
      End
      Begin VB.Label Label2 
         Caption         =   "%"
         Height          =   375
         Index           =   0
         Left            =   10440
         TabIndex        =   13
         Top             =   -600
         Width           =   375
      End
   End
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   2295
      Left            =   120
      OleObjectBlob   =   "FrmProjectAlarm.frx":3EAD6
      TabIndex        =   14
      Top             =   4440
      Width           =   7215
   End
   Begin MSChart20Lib.MSChart MSChart2 
      Height          =   2295
      Left            =   7560
      OleObjectBlob   =   "FrmProjectAlarm.frx":40F8E
      TabIndex        =   15
      Top             =   4440
      Width           =   6735
   End
End
Attribute VB_Name = "FrmProjectAlarm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmd_Click(Index As Integer)
Select Case Index
Case 5
    If DCproject.Text = "" Then
       If SystemOptions.UserInterface = ArabicInterface Then
          MsgBox "⁄›Ê« ...«·—Ã«¡  ÕœÌœ «”„ «·„‘—Ê⁄", vbInformation + vbMsgBoxRtlReading + vbOKOnly + vbMsgBoxRight, App.title
            DCproject.SetFocus
            Exit Sub
            Else
            MsgBox "Write Project Name ", vbInformation + vbOKOnly + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
            Exit Sub
            DCproject.SetFocus
         End If
     End If
 ProgressBar1.Visible = True
: ProgressBar1.value = 10
    FillGrid
    Chart1
    Chart2
: ProgressBar1.value = 50
ProgressBar1.Visible = False
ProgressBar1.value = 0
Case 6
Me.Hide
Case 9
'GetData
End Select

End Sub
Private Sub CmdHelp_Click()
          clear_all Me
          GridInstallments.Clear flexClearScrollable, flexClearEverything
          GridInstallments.Rows = 1
          'Me.MSChart1
End Sub
'Sub Chart()
'Chart1.DataEditor = True
'Chart1.DataEditorObj.Docked = Docked_Bottom
'Chart1.Series(0).Legend = "Product 1"
'Chart1.Series(1).Legend = "Product 2"
'Chart1.Series(2).Legend = "Product 3"

'End Sub

Function cahngelang()
    EleHeader.Caption = " Screen project operations Alerts "
    Me.Caption = EleHeader.Caption
    lbl(3).Caption = "Project"
  '  lbl(1).Caption = "Item"
  '  lbl(0).Caption = "From"
  '  lbl(14).Caption = "To"
  '  Frame5.Caption = "Period"
    lbl(4).Caption = "Update All"
   CmdHelp.Caption = "Clear"
   
   Cmd(5).Caption = "Search"
   'Cmd(9).Caption = "Print"
    Cmd(6).Caption = "Exit"
    With GridInstallments
    .TextMatrix(0, .ColIndex("Ser")) = "Serial"
    .TextMatrix(0, .ColIndex("Project_name")) = "Project Name"
    .TextMatrix(0, .ColIndex("opr_name")) = "oprations Name"
    .TextMatrix(0, .ColIndex("Achievements")) = "Done"
    .TextMatrix(0, .ColIndex("mot")) = "Remaining"
    .TextMatrix(0, .ColIndex("StartDate")) = "Start Date"
    .TextMatrix(0, .ColIndex("EndDate")) = "End Date "

   
    End With
End Function
Function RetDate(Optional fullcod As String, Optional ByRef start As Date, Optional ByRef Endd As Date, Optional Index As Integer) As Boolean
Dim str As String
RetDate = False
Dim Rs1 As ADODB.Recordset
Set Rs1 = New ADODB.Recordset
If Index = 0 Then
str = "select *from terms_operations where fullcode='" & fullcod & "' and  (NOT (dbo.terms_operations.StartDate IS NULL))"
Else
str = "select *from terms_operations where fullcode='" & fullcod & "' and  (NOT (dbo.terms_operations.EndDate IS NULL))"
End If
Rs1.Open str, Cn, adOpenStatic, adLockReadOnly, adCmdText
If Rs1.RecordCount >= 0 Then
If Index = 0 Then
If (IsNull(Rs1.Fields("StartDate").value)) Then
RetDate = False
Else
start = (IIf(IsNull(Rs1.Fields("StartDate").value), Date, Rs1.Fields("StartDate").value))
RetDate = True
End If
Else
If (IsNull(Rs1.Fields("EndDate").value)) Then
RetDate = False
Else
RetDate = True
Endd = (IIf(IsNull(Rs1.Fields("EndDate").value), Date, Rs1.Fields("EndDate").value))
End If
End If
End If
End Function
Public Sub FillGrid(Optional str As String)
Dim cont1 As Double
Dim cont As Double
Dim Typ As Integer
  '  On Error GoTo ErrTrap
On Error Resume Next
    Dim i As Integer
    Dim rs As ADODB.Recordset
Dim bol As Boolean
    Set rs = New ADODB.Recordset
  My_SQL = "SELECT     dbo.OperationsFollow.id, dbo.OperationsFollowDetails.opr_name, dbo.OperationsFollow.RecordDate, dbo.OperationsFollow.ProjectID, dbo.projects.Project_name, "
  My_SQL = My_SQL & "     dbo.projects.Project_nameE, dbo.OperationsFollowDetails.Achievements, dbo.OperationsFollowDetails.Status, dbo.OperationsFollowDetails.variance,"
  My_SQL = My_SQL & "        dbo.OperationsFollowDetails.opr_Fullcode, dbo.OperationsFollowDetails.varianceReasons, dbo.OperationsFollow.Remarks, dbo.OperationsFollow.challenges,"
  My_SQL = My_SQL & "        dbo.OperationsFollow.Proposals"
  My_SQL = My_SQL & " FROM    dbo.OperationsFollowDetails RIGHT OUTER JOIN"
  My_SQL = My_SQL & "    dbo.projects RIGHT OUTER JOIN"
  My_SQL = My_SQL & "     dbo.OperationsFollow ON dbo.projects.id = dbo.OperationsFollow.ProjectID ON dbo.OperationsFollowDetails.OperationsFollowId = dbo.OperationsFollow.id"


If Me.DCproject.Text <> "" And val(Me.DCproject.BoundText) <> 0 Then
My_SQL = My_SQL + " where OperationsFollow.ProjectID =" & val(Me.DCproject.BoundText) & ""
End If
My_SQL = My_SQL + "   order by  dbo.OperationsFollow.ProjectID "
 Dim ActualTotal As Double
Dim strt As Date
Dim EndDate As Date
'rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText
    rs.Open My_SQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
      With Me.GridInstallments
       .Rows = 1
        .Clear flexClearScrollable

        If rs.RecordCount > 0 Then
           .Rows = rs.RecordCount + 1
           rs.MoveFirst

            For i = 1 To .Rows - 1
              .TextMatrix(i, .ColIndex("Ser")) = i
              .TextMatrix(i, .ColIndex("id")) = (IIf(IsNull(rs.Fields("id").value), "", rs.Fields("id").value))
              .TextMatrix(i, .ColIndex("opr_name")) = (IIf(IsNull(rs.Fields("opr_name").value), "", rs.Fields("opr_name").value))
              .TextMatrix(i, .ColIndex("Achievements")) = (IIf(IsNull(rs.Fields("Achievements").value), 0, rs.Fields("Achievements").value))
              .TextMatrix(i, .ColIndex("mot")) = 100 - val(.TextMatrix(i, .ColIndex("Achievements")))
           bol = RetDate(((rs.Fields("opr_Fullcode").value)), strt, , 0)
           If bol = True Then
             .TextMatrix(i, .ColIndex("StartDate")) = strt
             End If
           bol = RetDate(((rs.Fields("opr_Fullcode").value)), , EndDate, 1)
           If bol = True Then
             .TextMatrix(i, .ColIndex("EndDate")) = EndDate
             End If
            If SystemOptions.UserInterface = ArabicInterface Then
           
              .TextMatrix(i, .ColIndex("Project_name")) = (IIf(IsNull(rs.Fields("Project_name").value), "", rs.Fields("Project_name").value))
         
               
            Else
           
              .TextMatrix(i, .ColIndex("Project_name")) = (IIf(IsNull(rs.Fields("Project_nameE").value), "", rs.Fields("Project_nameE").value))
             
              
            End If
         
        rs.MoveNext
            Next i
 
            rs.Close
        End If
 'sa .AutoSize 1, .Cols - 1, False

        .RowHeight(-1) = 300
    End With

End Sub
Private Sub Form_Load()
    Me.Left = (mdifrmmain.Width - Me.Width) / 2
    Me.Top = (mdifrmmain.Height - Me.Height) / 2 - 500
    Dim Dcombos As New ClsDataCombos
    Dim My_SQL As String

    My_SQL = " select id,Project_name from projects"
    fill_combo DCproject, My_SQL
    If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        cahngelang
    End If
    'FillGrid
    End Sub
 Sub Chart1()
On Error GoTo ErrTrap
    Dim i As Integer
    Dim rs As ADODB.Recordset
    Dim My_SQL As String
    Set rs = New ADODB.Recordset
    
    My_SQL = "SELECT     dbo.OperationsFollow.id, dbo.OperationsFollowDetails.opr_name, dbo.OperationsFollow.RecordDate, dbo.OperationsFollow.ProjectID, dbo.projects.Project_name, "
    My_SQL = My_SQL & "     dbo.projects.Project_nameE, dbo.OperationsFollowDetails.Achievements, dbo.OperationsFollowDetails.Status, dbo.OperationsFollowDetails.variance,"
    My_SQL = My_SQL & "        dbo.OperationsFollowDetails.opr_Fullcode, dbo.OperationsFollowDetails.varianceReasons, dbo.OperationsFollow.Remarks, dbo.OperationsFollow.challenges,"
    My_SQL = My_SQL & "        dbo.OperationsFollow.Proposals"
    My_SQL = My_SQL & " FROM    dbo.OperationsFollowDetails RIGHT OUTER JOIN"
    My_SQL = My_SQL & "    dbo.projects RIGHT OUTER JOIN"
    My_SQL = My_SQL & "     dbo.OperationsFollow ON dbo.projects.id = dbo.OperationsFollow.ProjectID ON dbo.OperationsFollowDetails.OperationsFollowId = dbo.OperationsFollow.id"
    If Me.DCproject.Text <> "" And val(Me.DCproject.BoundText) <> 0 Then
     My_SQL = My_SQL + " where OperationsFollow.ProjectID =" & val(Me.DCproject.BoundText) & ""
    End If
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText
       i = 1
       With MSChart1
        .ShowLegend = True
        .ColumnCount = rs.RecordCount
        .RowCount = 1
           If SystemOptions.UserInterface = ArabicInterface Then
           .RowLabel = "«·—”„ «·»Ì«‰Ì"
            Else
           .RowLabel = "Chart"
           End If
   End With
  ''''''''''''''''''''''''''''''''''''''''
  While i <= rs.RecordCount
  With MSChart1
            .Column = i
            .Row = 1
            .Data = IIf(IsNull(rs.Fields("Achievements").value), "", rs.Fields("Achievements").value)
            .ColumnLabel = IIf(IsNull(rs.Fields("opr_name").value), "", rs.Fields("opr_name").value)
    End With
  rs.MoveNext
  i = i + 1
  Wend
  Set rs = Nothing
      Exit Sub
ErrTrap:
End Sub
Sub Chart2()
On Error GoTo ErrTrap
    Dim i As Integer
    Dim rs As ADODB.Recordset
    Dim My_SQL As String
    Set rs = New ADODB.Recordset
    
    My_SQL = "SELECT     dbo.OperationsFollow.id, dbo.OperationsFollowDetails.opr_name, dbo.OperationsFollow.RecordDate, dbo.OperationsFollow.ProjectID, dbo.projects.Project_name, "
    My_SQL = My_SQL & "     dbo.projects.Project_nameE, dbo.OperationsFollowDetails.Achievements, dbo.OperationsFollowDetails.Status, dbo.OperationsFollowDetails.variance,"
    My_SQL = My_SQL & "        dbo.OperationsFollowDetails.opr_Fullcode, dbo.OperationsFollowDetails.varianceReasons, dbo.OperationsFollow.Remarks, dbo.OperationsFollow.challenges,"
    My_SQL = My_SQL & "        dbo.OperationsFollow.Proposals"
    My_SQL = My_SQL & " FROM    dbo.OperationsFollowDetails RIGHT OUTER JOIN"
    My_SQL = My_SQL & "    dbo.projects RIGHT OUTER JOIN"
    My_SQL = My_SQL & "     dbo.OperationsFollow ON dbo.projects.id = dbo.OperationsFollow.ProjectID ON dbo.OperationsFollowDetails.OperationsFollowId = dbo.OperationsFollow.id"
     If Me.DCproject.Text <> "" And val(Me.DCproject.BoundText) <> 0 Then
     My_SQL = My_SQL + " where OperationsFollow.ProjectID =" & val(Me.DCproject.BoundText) & ""
    End If
    rs.Open My_SQL, Cn, adOpenKeyset, adLockReadOnly, adCmdText
  
      i = 1
       With MSChart2
        .ShowLegend = True
        .ColumnCount = rs.RecordCount
        .RowCount = 1
         If SystemOptions.UserInterface = ArabicInterface Then
           .RowLabel = "«·—”„ «·»Ì«‰Ì"
            Else
           .RowLabel = "Chart"
         End If
  End With
  While i <= rs.RecordCount
  With MSChart2
            .Column = i
            .Row = 1
            .Data = IIf(IsNull(rs.Fields("Achievements").value), "", rs.Fields("Achievements").value)
            .ColumnLabel = IIf(IsNull(rs.Fields("opr_name").value), "", rs.Fields("opr_name").value)
    End With
  rs.MoveNext
  i = i + 1
  Wend
  Set rs = Nothing
      Exit Sub
ErrTrap:
End Sub

