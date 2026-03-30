VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Begin VB.Form FrmEstametChiled 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14970
   Icon            =   "FrmEstametChiled.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   14970
   Begin VB.Frame Frame2 
      BackColor       =   &H00E2E9E9&
      Height          =   735
      Left            =   0
      TabIndex        =   6
      Top             =   600
      Width           =   15015
      Begin VB.TextBox TxtTotal 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         ForeColor       =   &H00808000&
         Height          =   285
         Left            =   5760
         TabIndex        =   7
         Top             =   240
         Width           =   3165
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·«Ã„«·Ì"
         Height          =   345
         Index           =   32
         Left            =   9360
         TabIndex        =   8
         Top             =   240
         Width           =   930
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2E9E9&
      Height          =   855
      Left            =   0
      TabIndex        =   3
      Top             =   5040
      Width           =   14895
      Begin ImpulseButton.ISButton CMDOK 
         Height          =   450
         Left            =   8160
         TabIndex        =   4
         Top             =   240
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   794
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
         ButtonImage     =   "FrmEstametChiled.frx":6852
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
      End
      Begin ImpulseButton.ISButton Exit1 
         Height          =   450
         Left            =   2160
         TabIndex        =   5
         Top             =   240
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   794
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
         ButtonImage     =   "FrmEstametChiled.frx":6BEC
         ColorButton     =   14871017
         DrawFocusRectangle=   0   'False
         DisabledImageStyle=   1
      End
   End
   Begin VB.Frame Frame3 
      Height          =   2535
      Left            =   0
      TabIndex        =   2
      Top             =   2520
      Width           =   14895
      Begin VB.Image Image1 
         Height          =   2115
         Left            =   120
         Picture         =   "FrmEstametChiled.frx":6F86
         Stretch         =   -1  'True
         Top             =   240
         Width           =   14625
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid1 
      Height          =   1035
      Left            =   0
      TabIndex        =   0
      Top             =   1440
      Width           =   14985
      _cx             =   26432
      _cy             =   1826
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
      Rows            =   3
      Cols            =   36
      FixedRows       =   2
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmEstametChiled.frx":1854B
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
      Height          =   645
      Index           =   5
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   14955
      _cx             =   26379
      _cy             =   1138
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
      Picture         =   "FrmEstametChiled.frx":18859
      Caption         =   "‘«‘…  Ê“Ì⁄ «·› —«    "
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
   End
End
Attribute VB_Name = "FrmEstametChiled"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Public IndexType As Integer
Private Sub ChangeLang()
lbl(32).Caption = "Total"
CMDOK.Caption = "Save"
Exit1.Caption = "Exit"
Ele(5).Caption = "Distribution"
Me.Caption = Ele(5).Caption
End Sub

Private Sub Retrive()
 
    Dim StrSQL As String
    Dim AccountName As String
    Dim i As Integer
    Dim astrSplit2tems2() As String
    Dim astrSplitItems() As String
    Dim itemname As String
    Dim J As Integer
    Dim st As String
    Dim nElements As Integer
Dim k, m As Integer
     
 
        With Me.Grid1
        If FrmEstimations.Grid.TextMatrix(FrmEstimations.LonRow, FrmEstimations.Grid.ColIndex("StrEstametChiled")) <> "" Then
          st = FrmEstimations.Grid.TextMatrix(FrmEstimations.LonRow, FrmEstimations.Grid.ColIndex("StrEstametChiled"))
          st = Trim(st)
          astrSplitItems = Split(st, "@")
          nElements = UBound(astrSplitItems) - LBound(astrSplitItems)
            .Rows = .FixedRows + nElements

            For J = 0 To nElements - 1
            astrSplit2tems2 = Split(astrSplitItems(J), "#")
         '   astrSplit2tems2
         
            'StrSQL = Replace(Replace(astrSplit2tems2(0), Chr(10), ""), Chr(13), "")
            'StrSQL = Trim(StrSQL)
          
                ' .TextMatrix(i, .ColIndex("AccountCode")) = StrSQL
                m = 2
                 .TextMatrix(J + 2, .ColIndex("a")) = val(astrSplit2tems2(0))
                 .TextMatrix(J + 2, .ColIndex("b")) = val(astrSplit2tems2(1))
                 .TextMatrix(J + 2, .ColIndex("c")) = val(astrSplit2tems2(2))
                For k = 1 To 11
                m = m + 1
                   .TextMatrix(J + 2, .ColIndex("a" & k)) = val(astrSplit2tems2(m))
                    m = m + 1
                 .TextMatrix(J + 2, .ColIndex("b" & k)) = val(astrSplit2tems2(m))
                  m = m + 1
                 .TextMatrix(J + 2, .ColIndex("c" & k)) = val(astrSplit2tems2(m))
                 
           Next k
Next J
           ' Me.txt_opr_total.text = .Aggregate(flexSTSum, .FixedRows, .ColIndex("total"), .Rows - 1, .ColIndex("total"))
        End If
          
        End With
'ReLineGrid
    
          
  

End Sub
Sub save()
Dim str As String
Dim i As Integer
str = ""

With Me.Grid1
For i = 2 To .Rows - 1
 
 str = str & Trim(.TextMatrix(i, .ColIndex("a"))) & "#"
 str = str & Trim(.TextMatrix(i, .ColIndex("b"))) & "#"
 str = str & Trim(.TextMatrix(i, .ColIndex("c"))) & "#"
 Dim J As Integer
 For J = 1 To 11
  str = str & Trim(.TextMatrix(i, .ColIndex("a" & J))) & "#"
 str = str & Trim(.TextMatrix(i, .ColIndex("b" & J))) & "#"
 str = str & Trim(.TextMatrix(i, .ColIndex("c" & J))) & "#"
 Next J
 str = str & Trim("@")
  str = str & Chr(13)
  str = Trim(str)

Next i
FrmEstimations.Grid.TextMatrix(FrmEstimations.LonRow, FrmEstimations.Grid.ColIndex("StrEstametChiled")) = str
End With
End Sub
Private Sub CmdOk_Click()
save
With FrmEstimations.Grid
.Cell(flexcpBackColor, FrmEstimations.LonRow, .ColIndex("Ser"), FrmEstimations.LonRow, .ColIndex("StrEstametChiled")) = &H80000018
End With
  Unload Me
End Sub







Private Sub Exit1_Click()
Unload Me
End Sub

Private Sub Form_Load()
    Resize_Form Me
    With Grid1
 
      .MergeCells = flexMergeFixedOnly
        .MergeRow(0) = True
   If SystemOptions.UserInterface = ArabicInterface Then
          .Cell(flexcpText, 0, .ColIndex("a"), 0, .ColIndex("c")) = "1"
          .Cell(flexcpAlignment, 0, .ColIndex("a"), 0, .ColIndex("c")) = flexAlignCenterCenter
          .Cell(flexcpText, 1, .ColIndex("a"), 1, .ColIndex("a")) = " ﬁœÌ—Ì"
          .Cell(flexcpText, 1, .ColIndex("b"), 1, .ColIndex("b")) = "›⁄·Ì"
          .Cell(flexcpText, 1, .ColIndex("c"), 1, .ColIndex("c")) = "«‰Õ—«›"
          For i = 1 To 11
          
         .Cell(flexcpText, 0, .ColIndex("a" & i), 0, .ColIndex("c" & i)) = i + 1
         .Cell(flexcpAlignment, 0, .ColIndex("a" & i), 0, .ColIndex("c" & i)) = flexAlignCenterCenter
          .Cell(flexcpText, 1, .ColIndex("a" & i), 1, .ColIndex("a" & i)) = " ﬁœÌ—Ì"
          .Cell(flexcpText, 1, .ColIndex("b" & i), 1, .ColIndex("b" & i)) = "›⁄·Ì"
          .Cell(flexcpText, 1, .ColIndex("c" & i), 1, .ColIndex("c" & i)) = "«‰Õ—«›"
            .Cell(flexcpAlignment, 1, .ColIndex("a" & i), 1, .ColIndex("a" & i)) = flexAlignCenterCenter
          .Cell(flexcpAlignment, 1, .ColIndex("b" & i), 1, .ColIndex("b" & i)) = flexAlignCenterCenter
          .Cell(flexcpAlignment, 1, .ColIndex("c" & i), 1, .ColIndex("c" & i)) = flexAlignCenterCenter
          Next i
       Else
            .Cell(flexcpText, 0, .ColIndex("a"), 0, .ColIndex("c")) = "1"
          .Cell(flexcpAlignment, 0, .ColIndex("a"), 0, .ColIndex("c")) = flexAlignCenterCenter
          .Cell(flexcpText, 1, .ColIndex("a"), 1, .ColIndex("a")) = "Estimate"
          .Cell(flexcpText, 1, .ColIndex("b"), 1, .ColIndex("b")) = "Actual"
          .Cell(flexcpText, 1, .ColIndex("c"), 1, .ColIndex("c")) = "Deviation"
          For i = 1 To 11
          
         .Cell(flexcpText, 0, .ColIndex("a" & i), 0, .ColIndex("c" & i)) = i + 1
         .Cell(flexcpAlignment, 0, .ColIndex("a" & i), 0, .ColIndex("c" & i)) = flexAlignCenterCenter
          .Cell(flexcpText, 1, .ColIndex("a" & i), 1, .ColIndex("a" & i)) = "Estimate"
          .Cell(flexcpText, 1, .ColIndex("b" & i), 1, .ColIndex("b" & i)) = "Actual"
          .Cell(flexcpText, 1, .ColIndex("c" & i), 1, .ColIndex("c" & i)) = "Deviation"
            .Cell(flexcpAlignment, 1, .ColIndex("a" & i), 1, .ColIndex("a" & i)) = flexAlignCenterCenter
          .Cell(flexcpAlignment, 1, .ColIndex("b" & i), 1, .ColIndex("b" & i)) = flexAlignCenterCenter
          .Cell(flexcpAlignment, 1, .ColIndex("c" & i), 1, .ColIndex("c" & i)) = flexAlignCenterCenter
          Next i
       End If
   
      End With
   
       Grid1.Clear flexClearScrollable, flexClearEverything
            Grid1.Rows = 3
      If FrmEstimations.Grid.TextMatrix(FrmEstimations.LonRow, FrmEstimations.Grid.ColIndex("StrEstametChiled")) <> "" Then
      Retrive
      End If
      If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
       If FrmEstimations.TxtModFlg.text <> "R" Then
      CMDOK.Enabled = True
      Else
      CMDOK.Enabled = False
      End If
   hidcol
'LoadGride

End Sub
Sub hidcol()
Dim i As Integer
With Grid1
If FrmEstimations.TxtModFlg.text = "R" Then
.ColHidden(.ColIndex("b")) = True
.ColHidden(.ColIndex("c")) = True
For i = 1 To 11
.ColHidden(.ColIndex("b" & i)) = True
.ColHidden(.ColIndex("c" & i)) = True
Next i
Else
.ColHidden(.ColIndex("b")) = False
.ColHidden(.ColIndex("c")) = False
For i = 1 To 11
.ColHidden(.ColIndex("b" & i)) = False
.ColHidden(.ColIndex("c" & i)) = False
Next i
End If
End With
End Sub





Private Sub TxtTotal_Change()
   If val(TxtTotal.text) <> 0 Then
      CMDOK.Enabled = True
      Else
      CMDOK.Enabled = False
      End If
End Sub
