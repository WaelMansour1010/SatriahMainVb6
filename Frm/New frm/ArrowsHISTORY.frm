VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{E9DF30CA-4B30-4235-BF0C-7150F6466080}#1.0#0"; "ChartFX.ClientServer.Core.dll"
Begin VB.Form ArrowsHistory 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÇáÇÓÚÇÑ ÇáÊÇÑíÎíÉ"
   ClientHeight    =   6240
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   8805
   Icon            =   "ArrowsHISTORY.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   6240
   ScaleWidth      =   8805
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.PictureBox Picture1 
      Height          =   3255
      Left            =   9960
      Picture         =   "ArrowsHISTORY.frx":000C
      RightToLeft     =   -1  'True
      ScaleHeight     =   3195
      ScaleWidth      =   4875
      TabIndex        =   14
      Top             =   2640
      Width           =   4935
   End
   Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid2 
      Height          =   2940
      Left            =   0
      TabIndex        =   9
      Top             =   2280
      Width           =   8595
      _cx             =   15161
      _cy             =   5186
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
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   13
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"ArrowsHISTORY.frx":520A
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
   Begin VSFlex8Ctl.VSFlexGrid VSFlexGrid1 
      Height          =   6180
      Left            =   13560
      TabIndex        =   0
      Top             =   1560
      Width           =   9555
      _cx             =   16854
      _cy             =   10901
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
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   12
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"ArrowsHISTORY.frx":53E0
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
   Begin C1SizerLibCtl.C1Elastic EleTop 
      Height          =   660
      Left            =   0
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   8805
      _cx             =   15531
      _cy             =   1164
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial (Arabic)"
         Size            =   20.25
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
      BackColor       =   12648447
      ForeColor       =   8421376
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   "ÇáÇÓÚÇÑ ÇáÊÇÑíÎíÉ"
      Align           =   1
      AutoSizeChildren=   7
      BorderWidth     =   6
      ChildSpacing    =   4
      Splitter        =   0   'False
      FloodDirection  =   0
      FloodPercent    =   0
      CaptionPos      =   7
      WordWrap        =   -1  'True
      MaxChildSize    =   0
      MinChildSize    =   0
      TagWidth        =   0
      TagPosition     =   0
      Style           =   0
      TagSplit        =   2
      PicturePos      =   7
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
   End
   Begin MSDataListLib.DataCombo DcboCompanytId 
      Height          =   315
      Left            =   4200
      TabIndex        =   3
      Tag             =   "ÚÝæÇ íÑÌì ÇÏÎÇá ÃÓã ÇáÍí"
      Top             =   1800
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin C1SizerLibCtl.C1Elastic Ele 
      Height          =   705
      Index           =   1
      Left            =   3840
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   840
      Width           =   4875
      _cx             =   8599
      _cy             =   1244
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
      Appearance      =   4
      MousePointer    =   0
      Version         =   801
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      FloodColor      =   6553600
      ForeColorDisabled=   -2147483631
      Caption         =   "ÊÍÏíÏ ÇáÝÊÑÉ ÇáÒãäíÉ"
      Align           =   0
      AutoSizeChildren=   0
      BorderWidth     =   6
      ChildSpacing    =   4
      Splitter        =   0   'False
      FloodDirection  =   0
      FloodPercent    =   0
      CaptionPos      =   7
      WordWrap        =   -1  'True
      MaxChildSize    =   0
      MinChildSize    =   0
      TagWidth        =   0
      TagPosition     =   0
      Style           =   1
      TagSplit        =   2
      PicturePos      =   4
      CaptionStyle    =   0
      ResizeFonts     =   0   'False
      GridRows        =   0
      GridCols        =   0
      Frame           =   0
      FrameStyle      =   5
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   ""
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
      Begin MSComCtl2.DTPicker DTPickerAccFrom 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11265
            SubFormatType   =   3
         EndProperty
         Height          =   345
         Left            =   2730
         TabIndex        =   5
         ToolTipText     =   "ãä ÊÇÑíÎ ÞÏíã"
         Top             =   240
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   609
         _Version        =   393216
         CalendarBackColor=   -2147483624
         CalendarTitleBackColor=   10383715
         CheckBox        =   -1  'True
         CustomFormat    =   "yyyy/M/d"
         Format          =   100073475
         CurrentDate     =   37357
      End
      Begin MSComCtl2.DTPicker DTPickerAccTo 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11265
            SubFormatType   =   3
         EndProperty
         Height          =   345
         Left            =   90
         TabIndex        =   6
         ToolTipText     =   " Åáì ÊÇÑíÎ ÃÍÏË"
         Top             =   240
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   609
         _Version        =   393216
         CalendarBackColor=   -2147483624
         CalendarTitleBackColor=   10383715
         CheckBox        =   -1  'True
         CustomFormat    =   "yyyy/M/d"
         Format          =   100073475
         CurrentDate     =   37357
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "ãä"
         Height          =   285
         Index           =   4
         Left            =   3990
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   285
         Width           =   555
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "Åáì"
         Height          =   285
         Index           =   2
         Left            =   1710
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   240
         Width           =   555
      End
   End
   Begin Cfx62ClientServerCtl.Chart itemchart 
      Height          =   3375
      Index           =   0
      Left            =   9240
      TabIndex        =   10
      Top             =   2640
      Width           =   4815
      _Data_          =   "ArrowsHISTORY.frx":5592
   End
   Begin ImpulseButton.ISButton CmdPrint 
      Height          =   390
      Left            =   120
      TabIndex        =   15
      Top             =   5640
      Width           =   2565
      _ExtentX        =   4524
      _ExtentY        =   688
      ButtonStyle     =   1
      ButtonPositionImage=   1
      Caption         =   "ØÈÇÚÉ ÇáÊÞÑíÑ æÇáÑÓã ÇáÈíÇäí"
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
      ButtonImage     =   "ArrowsHISTORY.frx":5A33
      ColorButton     =   14871017
      DrawFocusRectangle=   0   'False
   End
   Begin VB.Label txtavgPrice 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   5280
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   5520
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "ÇáãÊæÓØ ÇáãÊÍÑß"
      Height          =   255
      Left            =   7080
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   5520
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "ÑÓã ÈíÇäí íæÖÍ ÇäÍÑÇÝÇÊ ÇáÇÓÚÇÑ"
      Height          =   615
      Left            =   9000
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   2160
      Width           =   2895
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "ÍÏÏ ÇáÔÑßå"
      Height          =   255
      Left            =   6960
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   1920
      Width           =   1575
   End
End
Attribute VB_Name = "ArrowsHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sql As String

Private Sub CmdPrint_Click()
    PrintReport (sql)
End Sub

Function PrintReport(sql As String)
    'save data
       
    Dim cAccountReport As ClsAccReports
    Set cAccountReport = New ClsAccReports
    cAccountReport.ArrowsHistorydata sql
    Set cAccountReport = Nothing

End Function

Private Sub DcboCompanytId_Change()

    If DcboCompanytId.BoundText = "" Then
        If SystemOptions.UserInterface = ArabicInterface Then
            MsgBox "ÇÎÊÑ ÔÑßå ÇæáÇ "
            'DcboFinMarketId.SetFocus
            SendKeys ("{F4}")
        Else
            MsgBox "òòSelect Company"
            DcboFinMarketId.SetFocus
            SendKeys ("{F4}")

        End If

    End If

    Me.VSFlexGrid2.Clear flexClearScrollable, flexClearEverything
    Me.VSFlexGrid2.Rows = 2

    Dim avgPrice As Double
    getPriceHistory val(DcboCompanytId.BoundText), DTPickerAccFrom.value, DTPickerAccTo.value, avgPrice
    txtavgPrice = avgPrice

End Sub

Function getPriceHistory(CompanyId As Integer, Optional fromdate As Date, Optional todate As Date, Optional avgPrice As Double)
    Dim SUM As Double
    Dim noofrecord As Integer
    SUM = 0

    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim Rs1 As ADODB.Recordset
    Set Rs1 = New ADODB.Recordset

    sql = "SELECT     dbo.ArrowsTransactions.OprId, dbo.ArrowsTransactions.Oprdate, dbo.ArrowsTransactions.qty, dbo.ArrowsTransactions.Price, dbo.ArrowsTransactions.CurrentValue, "
    sql = sql & "                     dbo.ArrowsTransactions.CompanyId , dbo.ArrowsCompanies.CompanyName"
    sql = sql & " FROM         dbo.ArrowsTransactions INNER JOIN"
    sql = sql & " dbo.ArrowsCompanies ON dbo.ArrowsTransactions.CompanyId = dbo.ArrowsCompanies.CompanyId"
    sql = sql & " where dbo.ArrowsTransactions.CompanyId=" & CompanyId
    sql = sql + "  and  Oprdate >='" & SQLDate(fromdate) & "' and Oprdate <='" & SQLDate(todate) & "'"
 
    sql = sql & "Order By Oprdate"
    rs.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
  
    If rs.RecordCount = 0 Then Exit Function
    noofrecord = rs.RecordCount
          
    With Me.VSFlexGrid2
        .Rows = 2
        .Clear flexClearScrollable

        If rs.RecordCount > 0 Then
            .Rows = rs.RecordCount + 1
            rs.MoveFirst

            For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("Ser")) = i
             
                .TextMatrix(i, .ColIndex("Oprdate")) = IIf(IsNull(rs.Fields("Oprdate").value), "", rs.Fields("Oprdate").value)
               
                .TextMatrix(i, .ColIndex("Price")) = IIf(IsNull(rs.Fields("Price").value), "", rs.Fields("Price").value)
                SUM = SUM + val(.TextMatrix(i, .ColIndex("Price")))
                rs.MoveNext
            Next

            rs.Close
            avgPrice = SUM / noofrecord
        End If

        .RowHeight(-1) = 300
    End With

End Function

Private Sub DTPickerAccFrom_Change()
    DcboCompanytId_Change
End Sub

Private Sub DTPickerAccTo_Change()
    DcboCompanytId_Change
End Sub

Private Sub Form_Load()
    Resize_Form Me
    NEW_interface = False
    Dim Dcombos As ClsDataCombos
    Set Dcombos = New ClsDataCombos
    Dcombos.getArrowsCompany DcboCompanytId

End Sub
