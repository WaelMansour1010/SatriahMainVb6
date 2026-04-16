VERSION 5.00
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form FrmDiviInvestmentCh 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5970
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10260
   Icon            =   "FrmDiviInvestmentCh.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   10260
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.TextBox TxtArea 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   360
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   5040
      Width           =   2055
   End
   Begin VB.CommandButton Exit1 
      Caption         =   "Œ—ÊÃ"
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   5520
      Width           =   1335
   End
   Begin VB.CommandButton CMDOK 
      Caption         =   "Õ›Ÿ"
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   5520
      Width           =   1335
   End
   Begin VSFlex8Ctl.VSFlexGrid Grid1 
      Height          =   4155
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   10305
      _cx             =   18177
      _cy             =   7329
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
      Cols            =   14
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"FrmDiviInvestmentCh.frx":000C
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
      Width           =   10275
      _cx             =   18124
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
      Picture         =   "FrmDiviInvestmentCh.frx":020F
      Caption         =   "   ›«’Ì· "
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
      Alignment       =   2  'Center
      BackColor       =   &H00E2E9E9&
      Caption         =   "«Ã„«·Ì «·„”«Õ… «·„” ﬁÿ⁄…"
      Height          =   285
      Index           =   5
      Left            =   2520
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   5040
      Width           =   1875
   End
End
Attribute VB_Name = "FrmDiviInvestmentCh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Public IndexType As Integer
Private Sub ChangeLang()
lbl(32).Caption = "Total"
lbl(0).Caption = "Total Values"
cmdOK.Caption = "Save"
Exit1.Caption = "Exit"
Ele(5).Caption = "Distribution"
Me.Caption = Ele(5).Caption
With GRID1
.TextMatrix(0, .ColIndex("Ser")) = "Serial"
.TextMatrix(0, .ColIndex("RecDate")) = "Date"
.TextMatrix(0, .ColIndex("val")) = "Value"
.TextMatrix(0, .ColIndex("Remark")) = "Remark"
End With
End Sub

Sub Reline()
Dim Area As Double
    Dim IntCounter As Integer
    IntCounter = 0
    Area = 0
    Dim i As Integer
    With Me.GRID1
        For i = .FixedRows To .rows - 1
                If .TextMatrix(i, .ColIndex("BlokNo")) <> "" Then
                IntCounter = IntCounter + 1
                .TextMatrix(i, .ColIndex("Ser")) = IntCounter
                Area = Area + val(.TextMatrix(i, .ColIndex("Area")))
           End If
           Next i
  
    End With
  TxtArea.text = Area
End Sub
Function GetUntName(Optional ID As Double = 0) As String
Dim sql As String
Dim Rs4 As ADODB.Recordset
Set Rs4 = New ADODB.Recordset
sql = "SELECT     *"
sql = sql & " From dbo.TblSpreading"
sql = sql & " WHERE     (UnitEnd = 1) AND (ID = " & ID & ")"
Rs4.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs4.RecordCount > 0 Then
If SystemOptions.UserInterface = ArabicInterface Then
GetUntName = IIf(IsNull(Rs4("Name").value), "", Rs4("Name").value)
Else
GetUntName = IIf(IsNull(Rs4("NameE").value), "", Rs4("NameE").value)
End If
Else
GetUntName = ""
End If
End Function
Private Sub Retrive()
 
    Dim StrSQL As String
    Dim AccountName As String
    Dim i As Integer
    Dim astrSplit2tems2() As String
    Dim astrSplitItems() As String
    Dim ItemName As String
    Dim j As Integer
    Dim st As String
    Dim nElements As Integer
Dim k, m As Integer
     
 
        With Me.GRID1
        If FrmDiviInvestment.GridInstallments.TextMatrix(FrmDiviInvestment.LonRow, FrmDiviInvestment.GridInstallments.ColIndex("StraInform")) <> "" Then
          st = FrmDiviInvestment.GridInstallments.TextMatrix(FrmDiviInvestment.LonRow, FrmDiviInvestment.GridInstallments.ColIndex("StraInform"))
          st = Trim(st)
          astrSplitItems = Split(st, "@")
          nElements = UBound(astrSplitItems) - LBound(astrSplitItems)
            .rows = .FixedRows + nElements
         '  DcbType.ListIndex = val(FrmDiviInvestment.GridInstallments.TextMatrix(FrmDiviInvestment.LonRow, FrmDiviInvestment.GridInstallments.ColIndex("TypBP"))) - 1
            For j = 0 To nElements - 1
            astrSplit2tems2 = Split(astrSplitItems(j), "#")
         '   astrSplit2tems2
         
            StrSQL = Replace(Replace(astrSplit2tems2(0), CHR(10), ""), CHR(13), "")
            StrSQL = Trim(StrSQL)
          
                 .TextMatrix(j + 1, .ColIndex("Ser")) = j + 1
                 .TextMatrix(j + 1, .ColIndex("BlokNo")) = StrSQL
                 .TextMatrix(j + 1, .ColIndex("unitunid")) = val((astrSplit2tems2(1)))
                 .TextMatrix(j + 1, .ColIndex("Nourth")) = (astrSplit2tems2(2))
                 .TextMatrix(j + 1, .ColIndex("South")) = astrSplit2tems2(3)
                 .TextMatrix(j + 1, .ColIndex("East")) = astrSplit2tems2(4)
                 .TextMatrix(j + 1, .ColIndex("West")) = astrSplit2tems2(5)
                 .TextMatrix(j + 1, .ColIndex("Area")) = val(astrSplit2tems2(6))
                 .TextMatrix(j + 1, .ColIndex("TotalArea")) = val(astrSplit2tems2(7))
                  StrSQL = Replace(Replace(astrSplit2tems2(8), CHR(10), ""), CHR(13), "")
                  StrSQL = Trim(StrSQL)
                  .TextMatrix(j + 1, .ColIndex("NewArea")) = StrSQL
                  .TextMatrix(j + 1, .ColIndex("UnitName")) = GetUntName(val(.TextMatrix(j + 1, .ColIndex("unitunid"))))
      Next j
          
        End If
          
        End With
'ReLineGrid
End Sub
Sub save()
Dim str As Variant
Dim i As Integer
str = ""
With Me.GRID1
For i = 1 To .rows - 1
If .TextMatrix(i, .ColIndex("BlokNo")) <> "" Then

 str = str & Trim(.TextMatrix(i, .ColIndex("BlokNo"))) & "#"
 str = str & val(.TextMatrix(i, .ColIndex("unitunid"))) & "#"
 str = str & Trim(.TextMatrix(i, .ColIndex("Nourth"))) & "#"
 str = str & Trim(.TextMatrix(i, .ColIndex("South"))) & "#"
 str = str & Trim(.TextMatrix(i, .ColIndex("East"))) & "#"
 str = str & Trim(.TextMatrix(i, .ColIndex("West"))) & "#"
 str = str & Trim(.TextMatrix(i, .ColIndex("Area"))) & "#"
 str = str & Trim(.TextMatrix(i, .ColIndex("TotalArea"))) & "#"
 str = str & Trim(.TextMatrix(i, .ColIndex("NewArea"))) & "#"
 str = str & Trim("@")
  str = str & CHR(13)
  str = Trim(str)
End If
Next i
FrmDiviInvestment.GridInstallments.TextMatrix(FrmDiviInvestment.LonRow, FrmDiviInvestment.GridInstallments.ColIndex("StraInform")) = str

End With
Unload Me
End Sub


Public Sub CmdOk_Click()
If val(TxtArea) <> val(FrmDiviInvestment.GridInstallments.TextMatrix(FrmDiviInvestment.LonRow, FrmDiviInvestment.GridInstallments.ColIndex("Area"))) Then
If SystemOptions.UserInterface = ArabicInterface Then
MsgBox "ÌÃ» «‰  ﬂÊ‰ «·„”«Õ… «·«Ã„«·Ì…  ”«ÊÌ «·„”«Õ… «·„ﬁ”„…"
Else
End If
Exit Sub
End If
save
End Sub





Private Sub Exit1_Click()
Unload Me
End Sub

Private Sub Form_Load()
    Resize_Form Me


       GRID1.Clear flexClearScrollable, flexClearEverything
            GRID1.rows = 2
      If FrmDiviInvestment.GridInstallments.TextMatrix(FrmDiviInvestment.LonRow, FrmDiviInvestment.GridInstallments.ColIndex("StraInform")) <> "" Then
      Retrive
      Reline
      End If
      Reline
      If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
  
End Sub

Private Sub Grid1_AfterEdit(ByVal row As Long, ByVal Col As Long)
With GRID1
Select Case .ColKey(Col)
Case "BlokNo"
  If row = 1 Then
                 .TextMatrix(row, .ColIndex("TotalArea")) = FrmDiviInvestment.GridInstallments.TextMatrix(FrmDiviInvestment.LonRow, FrmDiviInvestment.GridInstallments.ColIndex("Area"))
                 Else
                 .TextMatrix(row, .ColIndex("TotalArea")) = val(.TextMatrix(row - 1, .ColIndex("NewArea")))
                 End If
Case "Area"
  
                 
        If val(.TextMatrix(row, .ColIndex("TotalArea"))) >= val(.TextMatrix(row, .ColIndex("Area"))) Then
        .TextMatrix(row, .ColIndex("NewArea")) = val(.TextMatrix(row, .ColIndex("TotalArea"))) - val(.TextMatrix(row, .ColIndex("Area")))
        Else
        If SystemOptions.UserInterface = ArabicInterface Then
        MsgBox "·«Ì„ﬂ‰ «‰  ﬂÊ‰ «·„”«Õ… «ﬂ»— „‰ «·„”«Õ… «·«Ã„«·Ì…"
        Else
        MsgBox "Can not Area larger than total Area"
        End If
        
        .TextMatrix(row, .ColIndex("Area")) = 0
        Exit Sub
        End If
     Case "UnitName"
             
                 StrAccountCode = .ComboData
                 LngRow = .FindRow(StrAccountCode, .FixedRows, .ColIndex("unitunid"), False, True)
                 .TextMatrix(row, .ColIndex("unitunid")) = StrAccountCode
                 

   End Select
  
     If row = .rows - 1 Then
    
            .rows = .rows + 1
        End If
     End With
  Reline
End Sub

Private Sub Grid1_BeforeEdit(ByVal row As Long, ByVal Col As Long, Cancel As Boolean)
 With GRID1
        Select Case .ColKey(Col)
            Case "BlokNo"
                .ComboList = ""
                 Case "Nourth"
                .ComboList = ""
                 Case "East"
                .ComboList = ""
                 Case "West"
                .ComboList = ""
                 
                   Case "South"
                   .ComboList = ""
                     
                   Case "Area"
                   .ComboList = ""
                   Case "TotalArea"
                   .ComboList = ""
                         Case "NewArea"
                   .ComboList = ""
                End Select
             End With
                   
End Sub

Private Sub Grid1_StartEdit(ByVal row As Long, ByVal Col As Long, Cancel As Boolean)
  Dim rs As New ADODB.Recordset
    Dim StrSQL  As String
    Dim StrComboList As String
    Dim Msg As String
    Dim StrAccountCode As String
    With GRID1

        Select Case .ColKey(Col)
       
Case "UnitName"
  StrSQL = "select * from TblSpreading where UnitEnd=1 and Followed=" & val(FrmDiviInvestment.GridInstallments.TextMatrix(FrmDiviInvestment.LonRow, FrmDiviInvestment.GridInstallments.ColIndex("TypeDivi"))) & ""
                rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

                If SystemOptions.UserInterface = ArabicInterface Then
                    StrComboList = GRID1.BuildComboList(rs, "Name", "ID")
                Else
                    StrComboList = GRID1.BuildComboList(rs, "Namee", "ID")
                End If
       
                If StrComboList <> "" Then
                    StrComboList = "|" & StrComboList
                End If
                .ComboList = StrComboList
         
        End Select
    End With
End Sub
