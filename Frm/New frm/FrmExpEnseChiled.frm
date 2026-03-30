VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form FmrExpEnseChiled 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9330
   Icon            =   "FrmExpEnseChiled.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   9330
   Begin VB.TextBox TxtValue 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      ForeColor       =   &H00808000&
      Height          =   285
      Left            =   4800
      TabIndex        =   7
      Top             =   3360
      Width           =   3165
   End
   Begin VB.Frame Frame3 
      Height          =   1335
      Left            =   0
      TabIndex        =   6
      Top             =   3600
      Width           =   9255
      Begin VB.Image Image1 
         Height          =   1740
         Left            =   0
         Picture         =   "FrmExpEnseChiled.frx":000C
         Stretch         =   -1  'True
         Top             =   120
         Width           =   9240
      End
   End
   Begin VB.CommandButton Exit1 
      Caption         =   "Œ—ÊÃ"
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   5040
      Width           =   1335
   End
   Begin VB.TextBox TxtTotal 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      ForeColor       =   &H00808000&
      Height          =   285
      Left            =   4800
      TabIndex        =   3
      Top             =   720
      Width           =   3165
   End
   Begin VB.CommandButton CMDOK 
      Caption         =   "Õ›Ÿ"
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   5040
      Width           =   1335
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid1 
      Height          =   2235
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   9345
      _cx             =   16484
      _cy             =   3942
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
      Rows            =   2
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmExpEnseChiled.frx":5446
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
      Width           =   9315
      _cx             =   16431
      _cy             =   1138
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
      Picture         =   "FrmExpEnseChiled.frx":550D
      Caption         =   "‘«‘…  Ê“Ì⁄ «·› —«  ··„’—Ê›«  «·„ﬁœ„…"
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
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "≈Ã„«·Ì «·ﬁÌ„"
      Height          =   345
      Index           =   0
      Left            =   8040
      TabIndex        =   8
      Top             =   3360
      Width           =   1290
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      Caption         =   "«·«Ã„«·Ì"
      Height          =   345
      Index           =   32
      Left            =   8280
      TabIndex        =   4
      Top             =   720
      Width           =   930
   End
End
Attribute VB_Name = "FmrExpEnseChiled"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Public IndexType As Integer
Private Sub ChangeLang()
Lbl(32).Caption = "Total"
Lbl(0).Caption = "Total Values"
CmdOk.Caption = "Save"
Exit1.Caption = "Exit"
Ele(5).Caption = "Distribution"
Me.Caption = Ele(5).Caption
With Grid1
.TextMatrix(0, .ColIndex("Ser")) = "Serial"
.TextMatrix(0, .ColIndex("RecDate")) = "Date"
.TextMatrix(0, .ColIndex("val")) = "Value"
.TextMatrix(0, .ColIndex("Remark")) = "Remark"
End With
End Sub
Sub Reline()

    Dim IntCounter As Integer
    Dim Sm As Double
    Sm = 0
    IntCounter = 0
    Dim i As Integer
    With Me.Grid1
        For i = .FixedRows To .Rows - 1
                If .TextMatrix(i, .ColIndex("RecDate")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("Ser")) = IntCounter
           Sm = Sm + val(.TextMatrix(i, .ColIndex("val")))
           End If
           Next i
  
    End With
    TxtValue.text = Sm
End Sub

Private Sub Retrive()
 
    Dim StrSQL As String
    Dim AccountName As String
    Dim i As Integer
    Dim astrSplit2tems2() As String
    Dim astrSplitItems() As String
    Dim itemname As String
    Dim j As Integer
    Dim st As String
    Dim nElements As Integer
Dim k, m As Integer
     
 
        With Me.Grid1
        If FrmPripaidExpenses.Grid.TextMatrix(FrmPripaidExpenses.LonRow, FrmPripaidExpenses.Grid.ColIndex("StrDistribution")) <> "" Then
          st = FrmPripaidExpenses.Grid.TextMatrix(FrmPripaidExpenses.LonRow, FrmPripaidExpenses.Grid.ColIndex("StrDistribution"))
          st = Trim(st)
          astrSplitItems = Split(st, "@")
          nElements = UBound(astrSplitItems) - LBound(astrSplitItems)
            .Rows = .FixedRows + nElements

            For j = 0 To nElements - 1
            astrSplit2tems2 = Split(astrSplitItems(j), "#")
         '   astrSplit2tems2
         
            StrSQL = Replace(Replace(astrSplit2tems2(0), Chr(10), ""), Chr(13), "")
            StrSQL = Trim(StrSQL)
          
                 .TextMatrix(j + 1, .ColIndex("Ser")) = j + 1
               
               '  .TextMatrix(j + 1, .ColIndex("mon")) = val(astrSplit2tems2(0))
                 .TextMatrix(j + 1, .ColIndex("RecDate")) = StrSQL
                 .TextMatrix(j + 1, .ColIndex("val")) = val(astrSplit2tems2(1))
                 .TextMatrix(j + 1, .ColIndex("Remark")) = astrSplit2tems2(2)
                   StrSQL = Replace(Replace(astrSplit2tems2(3), Chr(10), ""), Chr(13), "")
            StrSQL = Trim(StrSQL)
                 .TextMatrix(j + 1, .ColIndex("id")) = val(StrSQL)
      Next j
          
        End If
          
        End With
'ReLineGrid
End Sub
Sub save()
Dim str As String
Dim i As Integer
str = ""
    Dim Sm As Double
    Sm = 0
    With Me.Grid1
        For i = .FixedRows To .Rows - 1
                If .TextMatrix(i, .ColIndex("RecDate")) <> "" Then
                    Sm = Sm + val(.TextMatrix(i, .ColIndex("val")))
    End If
        Next i
      
      If val(TXTTotal.text) <> val(TxtValue.text) Then
           
           If SystemOptions.UserInterface = ArabicInterface Then
           MsgBox "·«Ì„ﬂ‰ «‰ ÌﬂÊ‰ „Ã„Ê⁄ «·ﬁÌ„ «ﬂ»— «Ê «ﬁ·  „‰ «·«Ã„«·Ì  "
           Else
           MsgBox "The Value Must Be Equal The Total"
           End If
           Exit Sub
           End If
           
    End With
Dim IDDet As Double
With Me.Grid1
For i = 1 To .Rows - 1
 If Trim(.TextMatrix(i, .ColIndex("RecDate"))) <> "" Then
    IDDet = val((.TextMatrix(i, .ColIndex("id"))))
       If Me.Checked(IDDet) = True Then
       Else
       IDDet = 1
       maxx IDDet
       End If
       .TextMatrix(i, .ColIndex("id")) = IDDet
 str = str & Trim(.TextMatrix(i, .ColIndex("RecDate"))) & "#"
' str = str & Trim(.TextMatrix(i, .ColIndex("yar"))) & "#"
 str = str & Trim(.TextMatrix(i, .ColIndex("val"))) & "#"
 str = str & Trim(.TextMatrix(i, .ColIndex("Remark"))) & "#"
 str = str & Trim(.TextMatrix(i, .ColIndex("id"))) & "#"
 str = str & Trim("@")
  str = str & Chr(13)
  str = Trim(str)
End If
Next i
FrmPripaidExpenses.Grid.TextMatrix(FrmPripaidExpenses.LonRow, FrmPripaidExpenses.Grid.ColIndex("StrDistribution")) = str
End With
Unload Me
End Sub
Sub maxx(Optional ByRef IDDet As Double = 0)
     
  Dim RsDev As ADODB.Recordset
    Dim StrSQL As String
  Set RsDev = New ADODB.Recordset

    If IDDet <> 0 Then
   StrSQL = " select max(IDDet) as mx from ExpensesSearial"
   RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
   IDDet = IIf(IsNull(RsDev("mx").value), 0, RsDev("mx").value) + 1
      Set RsDev = New ADODB.Recordset
    RsDev.Open "ExpensesSearial", Cn, adOpenStatic, adLockOptimistic, adCmdTable
    RsDev.AddNew
RsDev("IDDet").value = IDDet
RsDev.update
End If
End Sub
Function Checked(Optional IDDet As Double = 0) As Boolean
     Checked = False
  Dim RsDev As ADODB.Recordset
    Dim StrSQL As String
      Set RsDev = New ADODB.Recordset

    If IDDet <> 0 Then
  StrSQL = " select * from ExpensesSearial where IDDet=" & IDDet & ""
   RsDev.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
 If RsDev.RecordCount > 0 Then
Checked = True
Else
Checked = False
End If
End If
End Function
Public Sub CmdOk_Click()
save
'With FrmEstimations.Grid
'.Cell(flexcpBackColor, FrmEstimations.LonRow, .ColIndex("Ser"), FrmEstimations.LonRow, .ColIndex("StrEstametChiled")) = &H80000018
'End With
 ' Unload Me
End Sub
Private Sub Exit1_Click()
Unload Me
End Sub

Private Sub Form_Load()
    Resize_Form Me
  
       Grid1.Clear flexClearScrollable, flexClearEverything
            Grid1.Rows = 2
      If FrmPripaidExpenses.Grid.TextMatrix(FrmPripaidExpenses.LonRow, FrmPripaidExpenses.Grid.ColIndex("Distribution")) <> "" Then
      Retrive
      Reline
      End If
      Reline
      If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
  
End Sub

Private Sub Grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Reline
End Sub

