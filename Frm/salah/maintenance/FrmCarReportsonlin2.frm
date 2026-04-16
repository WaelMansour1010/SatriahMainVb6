VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Begin VB.Form FrmCarReporonlin2 
   BackColor       =   &H00FFFFFF&
   Caption         =   "‘«‘… „ «»⁄… «Ê«„— «·‘€·"
   ClientHeight    =   8040
   ClientLeft      =   120
   ClientTop       =   510
   ClientWidth     =   20010
   Icon            =   "FrmCarReportsonlin2.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   8040
   ScaleWidth      =   20010
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   0
      TabIndex        =   19
      Top             =   5400
      Width           =   10455
      Begin VB.CommandButton menue 
         DownPicture     =   "FrmCarReportsonlin2.frx":038A
         Height          =   555
         Index           =   6
         Left            =   6720
         Picture         =   "FrmCarReportsonlin2.frx":76BC
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton menue 
         DownPicture     =   "FrmCarReportsonlin2.frx":7B5D
         Height          =   555
         Index           =   0
         Left            =   8160
         Picture         =   "FrmCarReportsonlin2.frx":EE8F
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton menue 
         Height          =   555
         Index           =   1
         Left            =   3240
         Picture         =   "FrmCarReportsonlin2.frx":F436
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   840
         Width           =   735
      End
      Begin VB.CommandButton menue 
         Height          =   555
         Index           =   2
         Left            =   6000
         Picture         =   "FrmCarReportsonlin2.frx":F8D7
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton menue 
         Height          =   555
         Index           =   3
         Left            =   4560
         Picture         =   "FrmCarReportsonlin2.frx":FDA7
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton menue 
         Height          =   555
         Index           =   4
         Left            =   3840
         Picture         =   "FrmCarReportsonlin2.frx":10260
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton menue 
         Height          =   555
         Index           =   5
         Left            =   2400
         Picture         =   "FrmCarReportsonlin2.frx":107B8
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton menue 
         BackColor       =   &H8000000D&
         Caption         =   " ÕœÌÀ"
         DownPicture     =   "FrmCarReportsonlin2.frx":10C03
         Height          =   555
         Index           =   16
         Left            =   8880
         Picture         =   "FrmCarReportsonlin2.frx":17F35
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton menue 
         DownPicture     =   "FrmCarReportsonlin2.frx":184CF
         Height          =   555
         Index           =   7
         Left            =   5280
         Picture         =   "FrmCarReportsonlin2.frx":1F801
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton menue 
         BackColor       =   &H80000005&
         DownPicture     =   "FrmCarReportsonlin2.frx":20091
         Height          =   555
         Index           =   15
         Left            =   9600
         Picture         =   "FrmCarReportsonlin2.frx":273C3
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton menue 
         Height          =   555
         Index           =   8
         Left            =   3120
         Picture         =   "FrmCarReportsonlin2.frx":27CDE
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton menue 
         DownPicture     =   "FrmCarReportsonlin2.frx":281C3
         Height          =   555
         Index           =   9
         Left            =   1680
         Picture         =   "FrmCarReportsonlin2.frx":2F4F5
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton menue 
         DownPicture     =   "FrmCarReportsonlin2.frx":2FA15
         Height          =   555
         Index           =   10
         Left            =   7440
         Picture         =   "FrmCarReportsonlin2.frx":36D47
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton menue 
         DownPicture     =   "FrmCarReportsonlin2.frx":3732E
         Height          =   555
         Index           =   11
         Left            =   960
         Picture         =   "FrmCarReportsonlin2.frx":3E660
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   120
         Width           =   735
      End
      Begin ImpulseButton.ISButton Cmd 
         Cancel          =   -1  'True
         Height          =   555
         Index           =   2
         Left            =   0
         TabIndex        =   34
         Top             =   120
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   979
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
   End
   Begin VB.Frame Frame3 
      Caption         =   "œ·«·«  «·Ê«‰ «·«ﬁ”«„"
      Height          =   2655
      Left            =   15000
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   5400
      Width           =   5055
      Begin VSFlex8UCtl.VSFlexGrid fg2 
         Height          =   2385
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   4920
         _cx             =   8678
         _cy             =   4207
         Appearance      =   1
         BorderStyle     =   0
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
         BackColorSel    =   16777215
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
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"FrmCarReportsonlin2.frx":3F1F4
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
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "œ·«·«  «·«·Ê«‰"
      Height          =   2655
      Left            =   10440
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   5400
      Width           =   4575
      Begin VB.TextBox Text2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Text            =   "Text2"
         Top             =   2760
         Width           =   375
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "Ì—„“ «·Ï «„— «·‘€·  «·– ›Ì  Õ  «·«‰ Ÿ«—"
         Height          =   255
         Index           =   4
         Left            =   720
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   1680
         Width           =   2775
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   135
         Left            =   3600
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   1680
         Width           =   375
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000FFFF&
         Height          =   135
         Left            =   3600
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "Ì—„“ «·Ï «„— «·‘€· «·–Ì  „ «’œ«— ·Â ›« Ê—…"
         Height          =   255
         Index           =   2
         Left            =   360
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   1200
         Width           =   3135
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "Ì—„“ «·Ï «„— «·‘€· «·Ã«—Ì «·⁄„· ⁄·ÌÂ "
         Height          =   375
         Index           =   1
         Left            =   -720
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   720
         Width           =   4215
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "Ì—„“ «·Ï «„— «·‘€· «·„‰ ÂÌ Ê·„ Ì „  ”·Ì„Â  ··⁄„Ì·"
         Height          =   375
         Index           =   0
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   240
         Width           =   3495
      End
      Begin VB.Label lbldf 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         Height          =   135
         Index           =   0
         Left            =   3600
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lbldsf 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000C000&
         Height          =   135
         Index           =   1
         Left            =   3600
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   720
         Width           =   375
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   5640
      Top             =   9360
   End
   Begin VB.ComboBox DcbOrderStatus 
      Height          =   315
      Left            =   20400
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   1440
      Width           =   1935
   End
   Begin VB.ComboBox ComGranty 
      Height          =   315
      Left            =   21120
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   120
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   20520
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Text            =   " „ „Ê«›ﬁ… «·⁄„"
      Top             =   720
      Width           =   915
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   615
      Index           =   0
      Left            =   2880
      TabIndex        =   0
      Top             =   9360
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   1085
      ButtonPositionImage=   1
      Caption         =   "⁄—÷ «· ﬁ—Ì—"
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
      Height          =   615
      Index           =   1
      Left            =   1410
      TabIndex        =   1
      Top             =   9360
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   1085
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
   Begin VSFlex8UCtl.VSFlexGrid Fg 
      Height          =   5505
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   19995
      _cx             =   35269
      _cy             =   9710
      Appearance      =   1
      BorderStyle     =   0
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
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   14871017
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16777215
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   16777215
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
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   29
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmCarReportsonlin2.frx":3F284
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
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      Caption         =   "‘«‘… «· ‰»ÌÂ« "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1905
      Index           =   10
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   6120
      Width           =   10095
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   2
      Height          =   1695
      Left            =   0
      Top             =   6360
      Width           =   10335
   End
End
Attribute VB_Name = "FrmCarReporonlin2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim DCboSearch As clsDCboSearch


Dim TTP As clstooltip
Dim cSearchDcbo  As clsDCboSearch
Dim TTD As clstooltipdemand
Dim Employee_account As String

Private Sub Check2_Click()

End Sub

Private Sub Cmd_Click(Index As Integer)

    Select Case Index

        Case 0
       

 GetData
            
        Case 1
            clear_all Me
'DtpDateFrom.value = ""
'DtpDateTo.value = ""
'Me.DtStart.value = ""
'Me.DtEnd.value = ""
            If SystemOptions.UserInterface = ArabicInterface Then
               ' Me.lbl(0).Caption = "‰ ÌÃ… «·»ÕÀ"
            Else
               ' Me.lbl(0).Caption = "Search Results"
            End If

        Case 2
            Unload Me
            Case 3
'print_report
    End Select

End Sub




'Public Sub FiLLTXT()
'
'    On Error GoTo ErrTrap
'    Dim i As Integer
' '   Frm2.Enabled = False
'    FrmCarAuthontication.XPTxtID.text = IIf(IsNull(RsSavRec.Fields("ID").value), "", RsSavRec.Fields("ID").value)
'    FrmCarAuthontication.TxtCliientName = IIf(IsNull(RsSavRec.Fields("CarID").value), "", RsSavRec.Fields("CarID").value)
'    TxtVacNamee.text = IIf(IsNull(RsSavRec.Fields("model").value), "", RsSavRec.Fields("model").value)
'
'    LabCurrRec.Caption = RsSavRec.AbsolutePosition
'    LabCountRec.Caption = RsSavRec.RecordCount

'    With Grid
'
'        For i = 1 To .Rows - 1
'
'            If Trim(TxtVac_ID.text) = .TextMatrix(i, .ColIndex("id")) Then
'                TxtSerial.text = .TextMatrix(i, .ColIndex("Ser"))
'                .Row = i
'                Exit Sub
'            End If

'        Next
'
'    End With
'
'ErrTrap:
'
'End Sub


'Private Sub Fg_EnterCell()
'   On Error GoTo ErrTrap
  '  FindRec val(Me.Fg.TextMatrix(Me.Grid.Row, Me.Fg.ColIndex("id")))
 'If FrmBillCarMaintExtra.ch = True Then
 'FrmBillCarMaintExtra.Retrive1 (val(Me.Fg.TextMatrix(Me.Fg.Row, Me.Fg.ColIndex("id"))))
 'Else
 ' FrmCarAuthontication.Retrive2 (val(Me.Fg.TextMatrix(Me.Fg.Row, Me.Fg.ColIndex("id"))))
 ' FrmCarAuthontication.TxtAmoutAccept.text = 0
 '   FrmCarAuthontication.TxtFirstPrice.text = 0
 '   FrmCarAuthontication.TXtCarMeter.text = ""
 '   FrmCarAuthontication.DcbOrderStatus.ListIndex = 0
'FrmCarAuthontication.ComGranty.ListIndex = 2
'  End If
'ErrTrap:
'End Sub
Public Function FindRec(ByVal RecId As Long)
 
End Function



Private Sub FG_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
Dim sql As String
    'On Error GoTo ErrTrap

    With Me.FG

        Select Case .ColKey(Col)

                 Case "sendsms"
                
.TextMatrix(Row, .ColIndex("SMS")) = -1
  sql = "update TblCardAuthorizationReform set   sendsms=-1  where ID=" & val(FG.TextMatrix(Row, FG.ColIndex("id"))) & ""
           Cn.Execute sql
             ' Case "dateenter"
            '.ColComboList(.ColIndex("dateenter")) = "..."
                End Select
                End With
End Sub

Private Sub fg_Click()
FrmCarAuthontication.retrive1 (val(Me.FG.TextMatrix(Me.FG.Row, Me.FG.ColIndex("id"))))
End Sub

Private Sub fg_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
  With Me.FG

        Select Case .ColKey(Col)
               Case "sendsms"
                
'TextMatrix(Row, .ColIndex("SMS")) = -1

             ' Case "dateenter"
            .ColComboList(.ColIndex("sendsms")) = "..."
                End Select
                End With
End Sub

Public Sub Fg2_Click()

End Sub

Private Sub Form_Activate()
   PutFormOnTop Me.hWnd
End Sub
Sub retrivgride()
 Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String
    Dim i As Integer
    Dim count
 sql = "   SELECT     DeparmentID, DepartmentName, DepartmentNamee, DeptColor,short, DeptBr, Dpeterial"
sql = sql & " From dbo.TblEmpDepartments"
sql = sql & " Where (Dpeterial Is Not Null)"
'sql = sql & "  Where  (dbo.TblCardAuthorizationReformDetails.id = " & id & ")"
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
If Rs3.RecordCount > 0 Then
With FG2
count = 1
 
   For i = 1 To Rs3.RecordCount
   If val(Rs3("Dpeterial").value) >= 0 Then
   count = count + 1
  .TextMatrix(count, .ColIndex("serial")) = count
  .Cell(flexcpBackColor, count, 1, count, 1) = Rs3("DeptColor").value
    If SystemOptions.UserInterface = EnglishInterface Then
    .TextMatrix(count, .ColIndex("dept")) = Rs3("DepartmentNamee").value
    Else
    
   .TextMatrix(count, .ColIndex("dept")) = Rs3("DepartmentName").value
   End If
   ' Fg.TextMatrix(ind, .ColIndex(str)) = -1
 '  Fg.ColHidden(count + 14) = False
 ' Fg.Cell(flexcpBackColor, count, count + 14, count, count + 14) = Rs3("DeptColor").value
'Fg.TextMatrix(count, .ColIndex("short")) = Rs3("short").value
End If
      Rs3.MoveNext
   Next i
End With
FG2.Rows = count + 1
End If
Fg2_Click
End Sub
Private Sub Retrivcoulme1(Optional Ind As Integer, Optional ID As Integer)
 Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String
    Dim Rs4 As ADODB.Recordset
Dim k, i As Integer
    Dim str As String
  '   sql = "   SELECT     DeparmentID, DepartmentName, DepartmentNamee, DeptColor,short, DeptBr, Dpeterial"
'sql = sql & " From dbo.TblEmpDepartments"
'sql = sql & " Where (Dpeterial Is Not Null)"
 sql = " SELECT     dbo.TblEmpDepartments.DeparmentID, dbo.TblEmpDepartments.DepartmentName,dbo.TblEmpDepartments.short, dbo.TblEmpDepartments.DepartmentNamee, dbo.TblEmpDepartments.DeptColor, "
sql = sql & "                      dbo.TblEmpDepartments.DeptBr , dbo.TblCardAuthorizationReformDetails.id, dbo.TblCardAuthorizationReformDetails.finish, dbo.TblEmpDepartments.Dpeterial"
sql = sql & " FROM         dbo.TblCardAuthorizationReformDetails INNER JOIN"
 sql = sql & "                     dbo.TblEmpDepartments ON dbo.TblCardAuthorizationReformDetails.Deptid = dbo.TblEmpDepartments.DeparmentID"
'sql = sql & " Where (dbo.TblCardAuthorizationReformDetails.finish = 1) And (dbo.TblCardAuthorizationReformDetails.id = " & id & ")"
sql = sql & "  Where  (dbo.TblCardAuthorizationReformDetails.id = " & ID & ")"
Dim count As Integer
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
   If Rs3.RecordCount > 0 Then
   With FG
   count = 1
   For i = 1 To Rs3.RecordCount
   str = "a"
    If val(Rs3("Dpeterial").value) >= 0 Then
   count = count + 1
  k = val(Rs3("Dpeterial").value) + 1
 ' If k <= Fg.Rows - 1 Then
   str = str & k
    If SystemOptions.UserInterface = EnglishInterface Then
    .TextMatrix(0, .ColIndex(str)) = Rs3("short").value
    Else
    If IsNull(Rs3("short").value) Then
   .TextMatrix(0, .ColIndex(str)) = ""
   Else
     .TextMatrix(0, .ColIndex(str)) = Rs3("short").value
     End If
   End If
    
   ' .TextMatrix(ind, .ColIndex(str)) = -1
  
 '  .ColHidden(k + 13) = False
 '  .Cell(flexcpBackColor, ind, k + 13, ind, k + 13) = Rs3("DeptColor").value
'.Cell(flexcpBackColor, 0, k + 13, 0, k + 13) = Rs3("DeptColor").value
'End If
End If
      Rs3.MoveNext
   Next i
End With
End If
 End Sub
Private Sub Retrivcoulme(ID As Integer, Ind As Integer)
 Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String
    Dim Rs4 As ADODB.Recordset
Dim k, i As Integer
    Dim str As String
   '  sql = "   SELECT     DeparmentID, DepartmentName, DepartmentNamee, DeptColor,short, DeptBr, Dpeterial"
'sql = sql & " From dbo.TblEmpDepartments"
'sql = sql & " Where (Dpeterial Is Not Null)"
sql = " SELECT     dbo.TblEmpDepartments.DeparmentID, dbo.TblEmpDepartments.DepartmentName,dbo.TblEmpDepartments.short, dbo.TblEmpDepartments.DepartmentNamee, dbo.TblEmpDepartments.DeptColor, "
sql = sql & "                      dbo.TblEmpDepartments.DeptBr , dbo.TblCardAuthorizationReformDetails.id, dbo.TblCardAuthorizationReformDetails.finish, dbo.TblEmpDepartments.Dpeterial"
sql = sql & " FROM         dbo.TblCardAuthorizationReformDetails INNER JOIN"
 sql = sql & "                     dbo.TblEmpDepartments ON dbo.TblCardAuthorizationReformDetails.Deptid = dbo.TblEmpDepartments.DeparmentID"
'sql = sql & " Where (dbo.TblCardAuthorizationReformDetails.finish = 1) And (dbo.TblCardAuthorizationReformDetails.id = " & id & ")"
sql = sql & "  Where  (dbo.TblCardAuthorizationReformDetails.id = " & ID & ")"
Dim count As Integer
Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
   If Rs3.RecordCount > 0 Then
   With FG
   count = 1
   For i = 1 To Rs3.RecordCount
   str = "a"
    If val(Rs3("Dpeterial").value) >= 0 Then
   count = count + 1
  k = val(Rs3("Dpeterial").value) + 1
   str = str & k
   If SystemOptions.UserInterface = EnglishInterface Then
    .TextMatrix(0, .ColIndex(str)) = Rs3("short").value
    Else
'
If IsNull(Rs3("short").value) Then
   .TextMatrix(0, .ColIndex(str)) = ""
   Else
     .TextMatrix(0, .ColIndex(str)) = Rs3("short").value
     End If
   End If
  ' MsgBox Rs3("finish").value
   If Rs3("finish").value = True Then
   .TextMatrix(Ind, .ColIndex(str)) = -1
End If
   .ColHidden(k + 13) = False
  .Cell(flexcpBackColor, Ind, k + 13, Ind, k + 13) = Rs3("DeptColor").value
 .Cell(flexcpBackColor, 0, k + 13, 0, k + 13) = Rs3("DeptColor").value
End If
      Rs3.MoveNext
   Next i
End With
End If
 End Sub
Private Function Retrive(ID As Integer, Optional ByRef str As String, Optional ByRef X As Integer, Optional ByRef strb As String)
 Dim Rs3 As ADODB.Recordset
    Set Rs3 = New ADODB.Recordset
    Dim sql As String
    Dim Rs4 As ADODB.Recordset
    Dim Index As Integer
    Set Rs4 = New ADODB.Recordset
    Dim SQL1 As String
sql = "SELECT     dbo.TblCardAuthorizationReform.ID, dbo.TblCardAuthorizationReformDetails.Type, dbo.TblCardAuthorizationReformDetails.ID2, dbo.TblCardAuthorizationReformDetails.ID AS idd,"
sql = sql & "                      dbo.TblMaintenanceWork.name AS NameM, dbo.TblMaintenanceWork.namee AS Nameem, dbo.TblCardAuthorizationReformDetails.Mainte,"
 sql = sql & "                      dbo.TblCardAuthorizationReformDetails.finish, dbo.TblCardAuthorizationReform.OrderStatus, dbo.TblMaintenanceWork.Type AS typemw,"
sql = sql & "                       dbo.TblCardAuthorizationReformDetails.ID2"
sql = sql & "  FROM         dbo.TblCardAuthorizationReform FULL OUTER JOIN"
sql = sql & "                       dbo.TblMaintenanceWork RIGHT OUTER JOIN"
sql = sql & "                       dbo.TblCardAuthorizationReformDetails ON dbo.TblMaintenanceWork.Id = dbo.TblCardAuthorizationReformDetails.Mainte ON"
sql = sql & "                       dbo.TblCardAuthorizationReform.id = dbo.TblCardAuthorizationReformDetails.id"
sql = sql & "  Where (dbo.TblCardAuthorizationReform.id =" & ID & ") And (dbo.TblCardAuthorizationReformDetails.Type = 0) And (dbo.TblCardAuthorizationReformDetails.finish = 0)"
'sql = sql & " WHERE     (dbo.TblCardAuthorizationReform.ID = " & id & ")"
   Rs3.Open sql, Cn, adOpenStatic, adLockOptimistic, adCmdText
   If Rs3.RecordCount > 0 Then
    str = IIf(Not IsNull(Rs3("NameM").value), Rs3("NameM").value, "")
    Index = Rs3("ID2").value
    Index = Index - 1
 SQL1 = " SELECT     dbo.TblCardAuthorizationReformDetails.ID2, dbo.TblCardAuthorizationReformDetails.ID, dbo.TblMaintenanceWork.Id AS idm,"
  SQL1 = SQL1 & "                     dbo.TblCardAuthorizationReformDetails.Type , dbo.TblMaintenanceWork.name, dbo.TblMaintenanceWork.namee"
SQL1 = SQL1 & "  FROM         dbo.TblCardAuthorizationReformDetails INNER JOIN"
SQL1 = SQL1 & "                       dbo.TblMaintenanceWork ON dbo.TblCardAuthorizationReformDetails.Mainte = dbo.TblMaintenanceWork.Id"
SQL1 = SQL1 & "  Where (dbo.TblCardAuthorizationReformDetails.Type = 0) And (dbo.TblCardAuthorizationReformDetails.ID2 =" & Index & ")And (dbo.TblCardAuthorizationReformDetails.finish = 1)"
     Rs4.Open SQL1, Cn, adOpenStatic, adLockOptimistic, adCmdText
    If Rs4.RecordCount > 0 Then
      strb = IIf(Not IsNull(Rs4("name").value), Rs4("name").value, "")
      End If
  If Rs3("typemw").value = True Then
 X = 1
 Else
 X = 0
 End If
 End If
 Exit Function
   
End Function

Private Sub ChangeLang()

    Cmd(1).Caption = "Delete"
    Cmd(0).Caption = "View Report"
   Cmd(2).Caption = "Exit"
   Frame2.Caption = "Connotations of colors"
  Me.Caption = "Screen follow the status of the job orders "
  'Me.lblblue.Caption = "Symbolizes the workpiece is being worked on"
  'Me.lblred.Caption = " Symbolizes that was completed was not handed over to the client"
  lbl(0).Caption = "Symbolizes the filling is finished has not been handed over to the client"
  lbl(1).Caption = "Symbolizes the current occupancy is currently working on it"
  lbl(2).Caption = "Symbolizes is the job that has been issuing his bill"
  lbl(10).Caption = "Symbolizes the current occupancy is currently working on it for Coputer Chek"
  lbl(4).Caption = "Symbolizes that the job is on hold"
  menue(7).Caption = "UpDate"
  Frame3.Caption = "Connotations of colors of Department"
  'RDGRANTY.RightToLeft = False
'RDGRANTY.Caption = "Granty"
'lbl(6).Caption = "ReqNo"
'RDWITHOUTGRANTY.RightToLeft = False
'RDWITHOUTGRANTY.Caption = "Without Granty"
'RDRETURNM.RightToLeft = False
'RDRETURNM.Caption = "Re Maintenance"

   With Me.FG2
        .TextMatrix(0, .ColIndex("Serial")) = "NO"
        .TextMatrix(0, .ColIndex("dept")) = "Department Name"
        .TextMatrix(0, .ColIndex("color")) = "Color"
        
    End With

     With Me.FG
     .TextMatrix(0, .ColIndex("typestatus")) = "TypeStatus"
        .TextMatrix(0, .ColIndex("Serial")) = "NO"
        .TextMatrix(0, .ColIndex("id")) = "Code"
        .TextMatrix(0, .ColIndex("RecordDate")) = "Date"
        .TextMatrix(0, .ColIndex("ClientName")) = "ClientName"
        .TextMatrix(0, .ColIndex("typestatusb")) = "Last Process"
       .TextMatrix(0, .ColIndex("type")) = "Type"
        .TextMatrix(0, .ColIndex("model")) = "Model"
        .TextMatrix(0, .ColIndex("dateenter")) = "Date Entry"
        .TextMatrix(0, .ColIndex("dateexit")) = "Date of exit expected"
        .TextMatrix(0, .ColIndex("datefinish")) = "Date completion"
        .TextMatrix(0, .ColIndex("diffrent")) = "Per day for the delay"
       .TextMatrix(0, .ColIndex("plate")) = "PlateNo"
        .TextMatrix(0, .ColIndex("ordestuts")) = "Order Stuts"
       .TextMatrix(0, .ColIndex("sms")) = "Send SMS"
        .TextMatrix(0, .ColIndex("wait")) = "Send SMS"
        .TextMatrix(0, .ColIndex("dateday")) = "DateNow"
         .TextMatrix(0, .ColIndex("sendsms")) = "Send SMS"
    End With
  '


  '
End Sub

Private Sub Form_Load()
    Dim GrdBack As ClsBackGroundPic
    Dim Dcombos As ClsDataCombos
    GetData
    
        Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500
    retrivgride
    
'Me.DtStart.value = ""
'Me.DtEnd.value = ""

'Me.RDALL.value = True
'Me.RdAll2.value = True
'    Set Dcombos = New ClsDataCombos
   ' Dcombos.GetEmployees Me.DCEmp_Name
'     Dcombos.GetClientName DcbClientname
'     Dcombos.GetTblCarModels DcbCarModel
'      Dcombos.GetTblMaintenanceWork Me.DCBMinten
'     Dcombos.GetTblCarsDataGroup DcbCarType
'    Set DCboSearch = New clsDCboSearch
'    Set DCboSearch.Client = Me.DcbClientname
    'Dcombos.GetUsers Me.DCUser
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Clear").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture
AddTip
  '  CenterForm Me
'GetData
'    FormPostion Me, GetPostion
    Set GrdBack = New ClsBackGroundPic
 If SystemOptions.UserInterface = EnglishInterface Then
        Me.ComGranty.AddItem "Granty"
        Me.ComGranty.AddItem "With out Granty"
        Me.ComGranty.AddItem "Re Maintenance"
        Me.DcbOrderStatus.AddItem "New"
        Me.DcbOrderStatus.AddItem "Accept Customer"
        Me.DcbOrderStatus.AddItem "Final Maintenance"
         
             Else
         Me.ComGranty.AddItem "»÷„«‰"
 Me.ComGranty.AddItem "»œÊ‰ ÷„«‰"
 Me.ComGranty.AddItem "≈⁄«œ… «’·«Õ"
 DcbOrderStatus.AddItem "ÃœÌœ"
DcbOrderStatus.AddItem " „ „Ê«›ﬁ… «·⁄„Ì·"
DcbOrderStatus.AddItem " „ «‰Â«¡ «·«’·«Õ"

    End If
 
 If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
   ' SetDtpickerDate Me.DtpDateFrom
   ' SetDtpickerDate Me.DtpDateTo

End Sub

Private Sub Form_Unload(Cancel As Integer)

    FormPostion Me, SavePostion
    Set DCboSearch = Nothing
End Sub

Public Sub GetData()
    Dim StrSQL As String
    Dim StrSQL1 As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
     Dim Rs1 As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
Dim ID As Integer
Dim cod As Integer
Dim strname As String
Dim strnameb As String

StrSQL = " SELECT     dbo.TblCardAuthorizationReform.ID,dbo.TblCardAuthorizationReform.SendSMS, dbo.TblCardAuthorizationReform.RecordDate, dbo.TblCardAuthorizationReform.ClientName,"
StrSQL = StrSQL & "      dbo.TblCardAuthorizationReform.Telephone, dbo.TblCardAuthorizationReform.CarTypeID, dbo.TBLCarTypes.name, dbo.TBLCarTypes.namee,"
StrSQL = StrSQL & "   dbo.TblCardAuthorizationReform.CarModelID, dbo.TblCarModels.CarID, dbo.TblCarModels.ModelE, dbo.TblCarModels.Model,"
StrSQL = StrSQL & "  dbo.TblCardAuthorizationReform.PlateNo, dbo.TblCardAuthorizationReform.ColorID, dbo.TblColor.name AS namecolor, dbo.TblColor.namee AS nameecolor,"
StrSQL = StrSQL & "  dbo.TblCardAuthorizationReform.YearFact, dbo.TblCardAuthorizationReform.OrderStatus, dbo.TblCardAuthorizationReform.Accept,"
StrSQL = StrSQL & "  dbo.TblCardAuthorizationReform.EndDate, dbo.TblCardAuthorizationReform.Month_Day, dbo.TblCardAuthorizationReform.Granty,"
StrSQL = StrSQL & " dbo.TblCardAuthorizationReform.DateEndG, dbo.TblCardAuthorizationReform.DateStartG, dbo.TblCardAuthorizationReform.CarMeter,"
StrSQL = StrSQL & "   dbo.TblCardAuthorizationReform.LongGranty, dbo.TblCardAuthorizationReform.PayFirst, dbo.TblCardAuthorizationReform.AmountAccept,"
StrSQL = StrSQL & " dbo.TblCardAuthorizationReform.Complaint, dbo.TblCardAuthorizationReform.Noteinitial, dbo.TblCardAuthorizationReform.Shaseh,"
StrSQL = StrSQL & " dbo.TblCardAuthorizationReform.NotAccept, dbo.TblCardAuthorizationReform.RecordeTime, dbo.TblCardAuthorizationReform.typerequest,"
StrSQL = StrSQL & " dbo.TblCardAuthorizationReform.FitterID, dbo.TblUsers.UserName, dbo.TblCardAuthorizationReform.mobile, dbo.TblCardAuthorizationReform.Cash,"
StrSQL = StrSQL & "  dbo.TblCardAuthorizationReform.Accoun, dbo.TblCardAuthorizationReform.credit, dbo.TblCardAuthorizationReform.fax, dbo.TblCardAuthorizationReform.box,"
StrSQL = StrSQL & "   dbo.TblCardAuthorizationReform.email, dbo.TblCardAuthorizationReform.address, dbo.TblCardAuthorizationReform.boxzip, dbo.TblCardAuthorizationReform.codereg,"
StrSQL = StrSQL & "   dbo.TblCardAuthorizationReform.typereg, dbo.TblCardAuthorizationReform.codedoor, dbo.TblCardAuthorizationReform.DateEnter,"
StrSQL = StrSQL & "   dbo.TblCardAuthorizationReform.DateExit, dbo.TblCardAuthorizationReform.persons, dbo.TblCardAuthorizationReform.Companies,"
StrSQL = StrSQL & "     dbo.TblCardAuthorizationReform.driver, dbo.TblCardAuthorizationReform.DateAcutExite, dbo.TblCardAuthorizationReform.DateExptExit,"
StrSQL = StrSQL & "    dbo.TblCardAuthorizationReform.TimeAcutExite , dbo.TblCardAuthorizationReform.TimeExptExit, dbo.TblCardAuthorizationReform.ResonUnderWait, dbo.TblCardAuthorizationReform.Payed"
StrSQL = StrSQL & " FROM    dbo.TblCardAuthorizationReform LEFT OUTER JOIN "
StrSQL = StrSQL & "  dbo.TblUsers ON dbo.TblCardAuthorizationReform.FitterID = dbo.TblUsers.UserID LEFT OUTER JOIN"
 StrSQL = StrSQL & " dbo.TBLCarTypes ON dbo.TblCardAuthorizationReform.CarTypeID = dbo.TBLCarTypes.id FULL OUTER JOIN"
StrSQL = StrSQL & " dbo.TblColor ON dbo.TblCardAuthorizationReform.ColorID = dbo.TblColor.Id FULL OUTER JOIN"
StrSQL = StrSQL & "  dbo.TblCarModels ON dbo.TblCardAuthorizationReform.CarModelID = dbo.TblCarModels.Id"
StrSQL = StrSQL & " Where  (dbo.TblCardAuthorizationReform.OrderStatus <=10)"

StrSQL = StrSQL & " and (TblCardAuthorizationReform.RecordDate <=  DATEadd( mm,-1,GETDATE()) and TblCardAuthorizationReform.RecordDate >=  DATEadd( mm,-2,GETDATE()) and TblCardAuthorizationReform.orderStatus < 2 )  "


    BolBegine = False
    StrWhere = ""

'StrWhere = StrWhere & " dbo.TblCardAuthorizationReform.OrderStatus <=5 "




 '   StrSQL = StrSQL & StrWhere
 
   StrSQL = StrSQL & " Order By dbo.TblCardAuthorizationReform.ID"
   
 ' print_report StrSQL
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
 'Set rs1 = New ADODB.Recordset
  '  rs1.Open StrSQL1, Cn, adOpenStatic, adLockReadOnly, adCmdText
    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
           ' Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
          '  Me.lbl(10).Caption = "Search Results=0"
        End If

    ' Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «· ﬁ—Ì—"
    '    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    Else
'print_report StrSQL
        With Me.FG
           .Clear flexClearScrollable, flexClearEverything
            .Rows = .FixedRows
            .Rows = rs.RecordCount + .FixedRows
'
            If SystemOptions.UserInterface = ArabicInterface Then
             '   Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
               ' Me.lbl(10).Caption = "Search Results=" & rs.RecordCount
            End If

           rs.MoveFirst
        
           For i = .FixedRows To .Rows - 1
                .TextMatrix(i, .ColIndex("Serial")) = i
                
                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs("ID").value), "", rs("ID").value)
        ID = IIf(IsNull(rs("ID").value), 0, rs("ID").value)
        
                If Not (IsNull(rs("RecordDate").value)) Then
                   .TextMatrix(i, .ColIndex("RecordDate")) = Format(rs("RecordDate").value, "yyyy/M/d")
                End If
                 If Not (IsNull(rs("EndDate").value)) Then
                   .TextMatrix(i, .ColIndex("dateenter")) = Format(rs("EndDate").value, "yyyy/M/d")
                End If
                 If Not (IsNull(rs("DateExptExit").value)) Then
                   .TextMatrix(i, .ColIndex("dateexit")) = Format(rs("DateExptExit").value, "yyyy/M/d")
                End If
                 If Not (IsNull(rs("DateAcutExite").value)) Then
                ' DtpDateFrom.value = rs("DateExptExit").value
                   .TextMatrix(i, .ColIndex("datefinish")) = Format(rs("DateAcutExite").value, "yyyy/M/d")
                End If
                  .TextMatrix(i, .ColIndex("dateday")) = Format(Date, "yyyy/M/d")
      
                        .TextMatrix(i, .ColIndex("ClientName")) = IIf(IsNull(rs("ClientName").value), "", rs("ClientName").value)
                '.TextMatrix(i, .ColIndex("Telephone")) = IIf(IsNull(rs("Telephone").value), "", rs("Telephone").value)
              
               .TextMatrix(i, .ColIndex("plate")) = IIf(IsNull(rs("PlateNo").value), "", rs("PlateNo").value)
                .TextMatrix(i, .ColIndex("model")) = IIf(IsNull(rs("Model").value), "", rs("Model").value)
                 .TextMatrix(i, .ColIndex("type")) = IIf(IsNull(rs("name").value), "", rs("name").value)
            strnameb = ""
            If (rs("OrderStatus").value < 2) Then
            
               .TextMatrix(i, .ColIndex("ordestuts")) = "Ã«—Ì «·⁄„·"
               Retrive ID, strname, cod, strnameb
             If cod = 1 Then
               .TextMatrix(i, .ColIndex("typestatus")) = strname
                .TextMatrix(i, .ColIndex("typestatusb")) = strnameb
               '.Cell(flexcpBackColor, i, 1, i, 26) = &HFF&
               Else
               .TextMatrix(i, .ColIndex("typestatus")) = strname
               .TextMatrix(i, .ColIndex("typestatusb")) = strnameb
               .Cell(flexcpBackColor, i, 12, i, 12) = &HC000&
               strname = ""
               End If
      .TextMatrix(i, .ColIndex("diffrent")) = IIf(IsNull(DateDiff("d", rs("DateExptExit").value, Date)), "", DateDiff("d", rs("DateExptExit").value, Date))

                End If
                     If (rs("OrderStatus").value = 3) Then
            .Cell(flexcpBackColor, i, 12, i, 12) = &H80000005
               .TextMatrix(i, .ColIndex("ordestuts")) = " Õ  «·«‰ Ÿ«—"
               
                 .TextMatrix(i, .ColIndex("wait")) = IIf(IsNull(rs("ResonUnderWait").value), "", rs("ResonUnderWait").value)
      .TextMatrix(i, .ColIndex("diffrent")) = DateDiff("d", rs("DateExptExit").value, Date)
                End If
                
             If (rs("OrderStatus").value = 2) Then
               .TextMatrix(i, .ColIndex("ordestuts")) = " „ «·«‰ Â«¡  „‰ «·⁄„· "
               
      .TextMatrix(i, .ColIndex("diffrent")) = DateDiff("d", rs("DateAcutExite").value, Date)
                .Cell(flexcpBackColor, i, 12, i, 12) = &H8000000D
                End If
               If (rs("OrderStatus").value = 5) And rs("Payed").value = False Then
               .TextMatrix(i, .ColIndex("ordestuts")) = " „ «’œ«—   ›« Ê—… "
               
      .TextMatrix(i, .ColIndex("diffrent")) = DateDiff("d", rs("DateAcutExite").value, Date)
                .Cell(flexcpBackColor, i, 12, i, 12) = &HFFFF&
                End If
                
              
                Retrivcoulme1 i, ID
                  Retrivcoulme ID, i
                  .TextMatrix(i, .ColIndex("SMS")) = IIf(IsNull(rs("SendSMS").value), "", rs("SendSMS").value)
                rs.MoveNext
               ' rs1.MoveNext
            Next i

            .AutoSize 0, .Cols - 1, False
          '  Me.lbl(1).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("AdvanceValue"), .Rows - 1, .ColIndex("AdvanceValue"))
        End With

    End If
'Retrivcoulme1
End Sub
Public Sub GetData1()
    Dim StrSQL As String
    Dim StrSQL1 As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
     Dim Rs1 As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
Dim ID As Integer
Dim cod As Integer
Dim strname As String
 ' If Not rs.RecordCount < 1 Then
 '               rs.Delete
                StrSQL = "Delete From TblOrederOpen Where id<>100000000"
                Cn.Execute StrSQL, , adExecuteNoRecords
 '               End If
 Set rs = New ADODB.Recordset
       rs.Open "TblOrederOpen", Cn, adOpenKeyset, adLockOptimistic, adCmdTable
If FG.Rows > 1 Then
          
       For i = Me.FG.FixedRows To FG.Rows - 1
         If val(FG.TextMatrix(i, FG.ColIndex("id"))) <> 0 Then
           rs.AddNew
          rs("id").value = val(FG.TextMatrix(i, FG.ColIndex("id")))
        rs("ClientName").value = FG.TextMatrix(i, FG.ColIndex("ClientName"))
        rs("type").value = FG.TextMatrix(i, FG.ColIndex("type"))
        rs("typestatusBefor").value = FG.TextMatrix(i, FG.ColIndex("typestatusb"))
        rs("model").value = FG.TextMatrix(i, FG.ColIndex("model"))
        rs("plate").value = FG.TextMatrix(i, FG.ColIndex("plate"))
       rs("diffrent").value = FG.TextMatrix(i, FG.ColIndex("diffrent"))
       rs("ordestuts").value = FG.TextMatrix(i, FG.ColIndex("ordestuts"))
        rs("typestatus").value = FG.TextMatrix(i, FG.ColIndex("typestatus"))
       rs("wait").value = FG.TextMatrix(i, FG.ColIndex("wait"))
       'rs("Telephone").value = Fg.TextMatrix(i, Fg.ColIndex("Telephone"))
'        rs("complaint").value = Fg.TextMatrix(i, Fg.ColIndex("complaint"))
      ' rs("PrivateCop").value = Fg.TextMatrix(i, Fg.ColIndex("PrivateCop"))
      ' rs("ReComentClient").value = Fg.TextMatrix(i, Fg.ColIndex("ReComentClient"))
      '  rs("repfitter").value = Fg.TextMatrix(i, Fg.ColIndex("repfitter"))
      ' rs("fitter").value = Fg.TextMatrix(i, Fg.ColIndex("fitter"))
       
        rs("RecordDate").value = IIf(IsDate(FG.TextMatrix(i, FG.ColIndex("RecordDate"))), FG.TextMatrix(i, FG.ColIndex("RecordDate")), Null)
       rs("dateenter").value = IIf(IsDate(FG.TextMatrix(i, FG.ColIndex("dateenter"))), FG.TextMatrix(i, FG.ColIndex("dateenter")), Null)
        rs("dateexit").value = IIf(IsDate(FG.TextMatrix(i, FG.ColIndex("dateexit"))), FG.TextMatrix(i, FG.ColIndex("dateexit")), Null)
     rs("datefinish").value = IIf(IsDate(FG.TextMatrix(i, FG.ColIndex("datefinish"))), FG.TextMatrix(i, FG.ColIndex("datefinish")), Null)
        rs("dateday").value = IIf(IsDate(FG.TextMatrix(i, FG.ColIndex("dateday"))), FG.TextMatrix(i, FG.ColIndex("dateday")), Null)
         
         rs.update
        
        End If
        Next i
        End If






StrSQL = "SELECT *  from TblOrederOpen "


'StrSQL = " SELECT     dbo.TblCardAuthorizationReform.ID, dbo.TblCardAuthorizationReform.RecordDate, dbo.TblCardAuthorizationReform.ClientName,"
'StrSQL = StrSQL & "      dbo.TblCardAuthorizationReform.Telephone, dbo.TblCardAuthorizationReform.CarTypeID, dbo.TBLCarTypes.name, dbo.TBLCarTypes.namee,"
'StrSQL = StrSQL & "   dbo.TblCardAuthorizationReform.CarModelID, dbo.TblCarModels.CarID, dbo.TblCarModels.ModelE, dbo.TblCarModels.Model,"
'StrSQL = StrSQL & "  dbo.TblCardAuthorizationReform.PlateNo, dbo.TblCardAuthorizationReform.ColorID, dbo.TblColor.name AS namecolor, dbo.TblColor.namee AS nameecolor,"
'StrSQL = StrSQL & "  dbo.TblCardAuthorizationReform.YearFact, dbo.TblCardAuthorizationReform.OrderStatus, dbo.TblCardAuthorizationReform.Accept,"
'StrSQL = StrSQL & "  dbo.TblCardAuthorizationReform.EndDate, dbo.TblCardAuthorizationReform.Month_Day, dbo.TblCardAuthorizationReform.Granty,"
'StrSQL = StrSQL & " dbo.TblCardAuthorizationReform.DateEndG, dbo.TblCardAuthorizationReform.DateStartG, dbo.TblCardAuthorizationReform.CarMeter,"
'StrSQL = StrSQL & "   dbo.TblCardAuthorizationReform.LongGranty, dbo.TblCardAuthorizationReform.PayFirst, dbo.TblCardAuthorizationReform.AmountAccept,"
'StrSQL = StrSQL & " dbo.TblCardAuthorizationReform.Complaint, dbo.TblCardAuthorizationReform.Noteinitial, dbo.TblCardAuthorizationReform.Shaseh,"
'StrSQL = StrSQL & " dbo.TblCardAuthorizationReform.NotAccept, dbo.TblCardAuthorizationReform.RecordeTime, dbo.TblCardAuthorizationReform.typerequest,"
'StrSQL = StrSQL & " dbo.TblCardAuthorizationReform.FitterID, dbo.TblUsers.UserName, dbo.TblCardAuthorizationReform.mobile, dbo.TblCardAuthorizationReform.Cash,"
'StrSQL = StrSQL & "  dbo.TblCardAuthorizationReform.Accoun, dbo.TblCardAuthorizationReform.credit, dbo.TblCardAuthorizationReform.fax, dbo.TblCardAuthorizationReform.box,"
'StrSQL = StrSQL & "   dbo.TblCardAuthorizationReform.email, dbo.TblCardAuthorizationReform.address, dbo.TblCardAuthorizationReform.boxzip, dbo.TblCardAuthorizationReform.codereg,"
'StrSQL = StrSQL & "   dbo.TblCardAuthorizationReform.typereg, dbo.TblCardAuthorizationReform.codedoor, dbo.TblCardAuthorizationReform.DateEnter,"
'StrSQL = StrSQL & "   dbo.TblCardAuthorizationReform.DateExit, dbo.TblCardAuthorizationReform.persons, dbo.TblCardAuthorizationReform.Companies,"
'StrSQL = StrSQL & "     dbo.TblCardAuthorizationReform.driver, dbo.TblCardAuthorizationReform.DateAcutExite, dbo.TblCardAuthorizationReform.DateExptExit,"
'StrSQL = StrSQL & "    dbo.TblCardAuthorizationReform.TimeAcutExite , dbo.TblCardAuthorizationReform.TimeExptExit, dbo.TblCardAuthorizationReform.ResonUnderWait, dbo.TblCardAuthorizationReform.Payed"
'StrSQL = StrSQL & " FROM    dbo.TblCardAuthorizationReform LEFT OUTER JOIN"
'StrSQL = StrSQL & "  dbo.TblUsers ON dbo.TblCardAuthorizationReform.FitterID = dbo.TblUsers.UserID LEFT OUTER JOIN"
'' StrSQL = StrSQL & " dbo.TBLCarTypes ON dbo.TblCardAuthorizationReform.CarTypeID = dbo.TBLCarTypes.id FULL OUTER JOIN"
'StrSQL = StrSQL & " dbo.TblColor ON dbo.TblCardAuthorizationReform.ColorID = dbo.TblColor.Id FULL OUTER JOIN"
'StrSQL = StrSQL & "  dbo.TblCarModels ON dbo.TblCardAuthorizationReform.CarModelID = dbo.TblCarModels.Id"
'StrSQL = StrSQL & " Where  (1 = 1)"
'    BolBegine = False
'    StrWhere = ""

'StrWhere = StrWhere & " AND dbo.TblCardAuthorizationReform.OrderStatus <=5 "



'
   ' StrSQL = StrSQL & StrWhere
 
   StrSQL = StrSQL & " Order By dbo.TblOrederOpen.ID"
   
  print_report StrSQL
'    Set rs = New ADODB.Recordset
'    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText
 'Set rs1 = New ADODB.Recordset
  '  rs1.Open StrSQL1, Cn, adOpenStatic, adLockReadOnly, adCmdText
'    If rs.BOF Or rs.EOF Then
'        If SystemOptions.UserInterface = ArabicInterface Then
           ' Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=’›—"
'        ElseIf SystemOptions.UserInterface = EnglishInterface Then
'          '  Me.lbl(10).Caption = "Search Results=0"
'        End If

    ' Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «· ﬁ—Ì—"
    '    MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'        Exit Sub
'    Else
'print_report StrSQL
'        With Me.Fg
'           .Clear flexClearScrollable, flexClearEverything
'            .Rows = .FixedRows
'            .Rows = rs.RecordCount + .FixedRows
'
'            If SystemOptions.UserInterface = ArabicInterface Then
'             '   Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
'            ElseIf SystemOptions.UserInterface = EnglishInterface Then
'               ' Me.lbl(10).Caption = "Search Results=" & rs.RecordCount
'            End If

'           rs.MoveFirst
        
'           For i = .FixedRows To .Rows - 1
'                .TextMatrix(i, .ColIndex("Serial")) = i
'                .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs("ID").value), "", rs("ID").value)
'        id = IIf(IsNull(rs("ID").value), 0, rs("ID").value)
'                If Not (IsNull(rs("RecordDate").value)) Then
'                   .TextMatrix(i, .ColIndex("RecordDate")) = Format(rs("RecordDate").value, "yyyy/M/d")
'                End If
'                 If Not (IsNull(rs("EndDate").value)) Then
'                   .TextMatrix(i, .ColIndex("dateenter")) = Format(rs("EndDate").value, "yyyy/M/d")
'                End If
'                 If Not (IsNull(rs("DateExptExit").value)) Then
'                   .TextMatrix(i, .ColIndex("dateexit")) = Format(rs("DateExptExit").value, "yyyy/M/d")
'                End If
'                 If Not (IsNull(rs("DateAcutExite").value)) Then
'                ' DtpDateFrom.value = rs("DateExptExit").value
'                   .TextMatrix(i, .ColIndex("datefinish")) = Format(rs("DateAcutExite").value, "yyyy/M/d")
'                End If
'                  .TextMatrix(i, .ColIndex("dateday")) = Format(Date, "yyyy/M/d")
'
'                        .TextMatrix(i, .ColIndex("ClientName")) = IIf(IsNull(rs("ClientName").value), "", rs("ClientName").value)
'                '.TextMatrix(i, .ColIndex("Telephone")) = IIf(IsNull(rs("Telephone").value), "", rs("Telephone").value)
'
'               .TextMatrix(i, .ColIndex("plate")) = IIf(IsNull(rs("PlateNo").value), "", rs("PlateNo").value)
'                .TextMatrix(i, .ColIndex("model")) = IIf(IsNull(rs("Model").value), "", rs("Model").value)
'                 .TextMatrix(i, .ColIndex("type")) = IIf(IsNull(rs("name").value), "", rs("name").value)
'
'            If (rs("OrderStatus").value < 2) Then
'
'               .TextMatrix(i, .ColIndex("ordestuts")) = "Ã«—Ì «·⁄„·"
'               retrive id, strname, cod
'             If cod = 1 Then
'               .TextMatrix(i, .ColIndex("typestatus")) = strname
'               .Cell(flexcpBackColor, i, 1, i, 15) = &HFF&
'               Else
'               .TextMatrix(i, .ColIndex("typestatus")) = strname
'               .Cell(flexcpBackColor, i, 1, i, 15) = &HC000&
'               strname = ""
'               End If
'      .TextMatrix(i, .ColIndex("diffrent")) = DateDiff("d", rs("DateExptExit").value, Date)
'                End If
'                     If (rs("OrderStatus").value = 3) Then
'            .Cell(flexcpBackColor, i, 1, i, 15) = &H80000005
'               .TextMatrix(i, .ColIndex("ordestuts")) = " Õ  «·«‰ Ÿ«—"
'
''                 .TextMatrix(i, .ColIndex("wait")) = IIf(IsNull(rs("ResonUnderWait").value), "", rs("ResonUnderWait").value)
 '     .TextMatrix(i, .ColIndex("diffrent")) = DateDiff("d", rs("DateExptExit").value, Date)
 '               End If
 '
 '            If (rs("OrderStatus").value = 2) Then
 '              .TextMatrix(i, .ColIndex("ordestuts")) = " „ «·«‰ Â«¡  „‰ «·⁄„· "
 '
 '     .TextMatrix(i, .ColIndex("diffrent")) = DateDiff("d", rs("DateAcutExite").value, Date)
 '               .Cell(flexcpBackColor, i, 1, i, 15) = &H8000000D
 '               End If
 '              If (rs("OrderStatus").value = 5) And rs("Payed").value = False Then
 '              .TextMatrix(i, .ColIndex("ordestuts")) = " „ «’œ«—   ›« Ê—… "
 '
 '     .TextMatrix(i, .ColIndex("diffrent")) = DateDiff("d", rs("DateAcutExite").value, Date)
 '               .Cell(flexcpBackColor, i, 1, i, 15) = &HFFFF&
 '               End If
 '               rs.MoveNext
 '              ' rs1.MoveNext
 '           Next i
'
'            .AutoSize 0, .Cols - 1, False
'          '  Me.lbl(1).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("AdvanceValue"), .Rows - 1, .ColIndex("AdvanceValue"))
'        End With
'
'    End If
'
End Sub
Function print_report(Optional NoteSerial As String)
     
    Set rs = New ADODB.Recordset
    rs.Open NoteSerial, Cn, adOpenStatic, adLockReadOnly, adCmdText
     
    Dim MySQL As String
    Dim RsData As New ADODB.Recordset
    Dim xApp As New CRAXDRT.Application
    Dim xReport As CRAXDRT.Report
    Dim CViewer As ClsReportViewer
    Dim StrReportTitle As String
    Dim StrFileName As String
    Dim Msg As String

 If SystemOptions.UserInterface = ArabicInterface Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepprientOpent.rpt"
        Else
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepprientOpent.rpt"
        End If



    If Dir(StrFileName) = "" Then
        'GetMsgs 139, vbExclamation
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Set RsData = New ADODB.Recordset
    RsData.Open NoteSerial, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If RsData.BOF Or RsData.EOF Then
        'GetMsgs 138, vbExclamation
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        RsData.Close
        Set RsData = Nothing
        Screen.MousePointer = vbDefault
        Exit Function
    End If

    Screen.MousePointer = vbArrowHourglass
    Set xReport = xApp.OpenReport(StrFileName)
    xReport.Database.SetDataSource RsData

    Dim cCompanyInfo As New ClsCompanyInfo

    If SystemOptions.UserInterface = ArabicInterface Then
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
        ' xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
        StrReportTitle = "" '& StrAccountName
        'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        '    StrReportTitle = StrReportTitle + " »œ«Ì… „‰ " & Format(Me.DTPickerAccFrom.value, "yyyy/M/d") & ""
        'End If
        'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        '    StrReportTitle = StrReportTitle + " ≈·Ï " & Format(Me.DTPickerAccTo.value, "yyyy/M/d") & " "
        'End If
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        'xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""
        'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
        '    StrReportTitle = StrReportTitle + " From Date " & (Me.DTPickerAccFrom.value) & ""
        'End If
        'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
        '    StrReportTitle = StrReportTitle + " To Date :  " & (Me.DTPickerAccTo.value) & ""
        'End If
    End If

    xReport.ParameterFields(3).AddCurrentValue user_name
   ' xReport.ParameterFields(13).AddCurrentValue Me.DTPicker1.value
      '  xReport.ParameterFields(4).AddCurrentValue WriteNo(Format(val(TxtAdvanceValue.text), "0.00"), 0, True, ".")
       ' xReport.ParameterFields(6).AddCurrentValue val(lbl(23).Caption)
        ' xReport.ParameterFields(13).AddCurrentValue Me.DTPicker1.value
'    xReport.ParameterFields(8).AddCurrentValue IIf(IsNumeric(fg.TextMatrix(Me.fg.FixedRows, fg.ColIndex("PartValue"))), val(fg.TextMatrix(Me.fg.FixedRows, fg.ColIndex("PartValue"))), 0)
' xReport.ParameterFields(9).AddCurrentValue val(lbl(22).Caption)
 ' xReport.ParameterFields(10).AddCurrentValue val(TxtDiscount.text)
  ' xReport.ParameterFields(11).AddCurrentValue txtDiscountDES.text
  Dim total As String
  Dim totl As Double
 ' totl = val(LbToTalExtra.Caption) + val(Me.lbTotalMente.Caption)
 ' total = totl
 '  xReport.ParameterFields(12).AddCurrentValue Me.lbTotalMente.Caption
 '     xReport.ParameterFields(13).AddCurrentValue LbToTalExtra.Caption
 '       xReport.ParameterFields(14).AddCurrentValue total
   ' xReport.ParameterFields(5).AddCurrentValue ToHijriDate(RsData("notedate").value)
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.title
    xReport.ReportAuthor = App.title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault


 
  
 
End Function
Private Sub AddTip()
    Dim Wrap As String
    On Error GoTo ErrTrap
    Wrap = CHR(13) + CHR(10)
    Set TTP = New clstooltip

 With TTP
        .Create Me.hWnd, " «·«‰ ﬁ«· «·Ï ‘«‘… «· ”·Ì„ ··⁄„Ì·", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl menue(3), "‘«‘… ..." & Wrap & "  ··«‰ ﬁ«·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With
     With TTP
        .Create Me.hWnd, "   ‘«‘…  «· ‰»ÌÂ« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl menue(8), "  ÕœÌÀ..." & Wrap & "  " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With
     With TTP
        .Create Me.hWnd, "   ‘«‘…  «· ‰»ÌÂ« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl menue(9), "ÿ»«⁄… ..." & Wrap & "  " & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With
    With TTP
        .Create Me.hWnd, " «·«‰ ﬁ«· «·Ï ‘«‘… «·«Ê«„— «·„› ÊÕ… ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl menue(7), "‘«‘… ..." & Wrap & "  ··«‰ ﬁ«·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With
    With TTP
        .Create Me.hWnd, " «·«‰ ﬁ«· «·Ï ‘«‘…  «· ‰»ÌÂ« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl menue(4), "‘«‘… ..." & Wrap & "  ··«‰ ﬁ«·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With
     With TTP
        .Create Me.hWnd, " «·«‰ ﬁ«· «·Ï ‘«‘…  «· ﬁ«—Ì—", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl menue(5), "‘«‘… ..." & Wrap & "  ··«‰ ﬁ«·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With
      With TTP
        .Create Me.hWnd, " «·«‰ ﬁ«· «·Ï ‘«‘…  ’—› ﬁÿ⁄ «·€Ì«—", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl menue(2), "‘«‘… ..." & Wrap & "  ··«‰ ﬁ«·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

       With TTP
        .Create Me.hWnd, " «·«‰ ﬁ«· «·Ï ‘«‘… ÿ·» ›Õ’ ﬂ„»ÌÊ —  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl menue(6), "‘«‘… ..." & Wrap & "  ··«‰ ﬁ«·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With
         With TTP
        .Create Me.hWnd, " «·«‰ ﬁ«· «·Ï ‘«‘…    ÿ·» ’Ì«‰…  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl menue(0), "‘«‘… ..." & Wrap & "  ··«‰ ﬁ«·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With
 

       With TTP
        .Create Me.hWnd, " «·«‰ ﬁ«· «·Ï ‘«‘…   «·⁄„Ê·«  «·„” Õﬁ…  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl menue(9), "‘«‘… ..." & Wrap & "  ··«‰ ﬁ«·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With
          With TTP
        .Create Me.hWnd, " «·«‰ ﬁ«· «·Ï ‘«‘…   „·› «·⁄„·«¡  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl menue(10), "‘«‘… ..." & Wrap & "  ··«‰ ﬁ«·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With
          With TTP
        .Create Me.hWnd, " «·«‰ ﬁ«· «·Ï ‘«‘…   ﬁ«—Ì— «·⁄„Ê·«   ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl menue(11), "‘«‘… ..." & Wrap & "  ··«‰ ﬁ«·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With



    Exit Sub
ErrTrap:
End Sub
'Sub dateval()
 
'   Dim astrSplitItems() As String
'    Dim result As String
    
 
'     Dim diff_year As Integer
'    result = ExactAge(DTFrom.value, DTTo.value)

 

'    astrSplitItems = Split(result, "-")
   ' TxtYear.text = astrSplitItems(0)
'   ' TxtMonth.text = astrSplitItems(1)
'    TxtDay.text = astrSplitItems(2)
''
    
'End Sub
'Function print_report(Optional NoteSerial As String)
     
    'Set rs = New ADODB.Recordset
    'rs.Open NoteSerial, Cn, adOpenStatic, adLockReadOnly, adCmdText
    '
    'Dim MySQL As String
    'Dim RsData As New ADODB.Recordset
    'Dim xApp As New CRAXDRT.Application
    'Dim xReport As CRAXDRT.Report
'    Dim CViewer As ClsReportViewer
    'Dim StrReportTitle As String
'    'Dim StrFileName As String
    'Dim Msg As String

''
''
''        If SystemOptions.UserInterface = ArabicInterface Then
'        If Me.XPChkSearchTypeClient1.value = True Then
'            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rep1byclient.rpt"
'            Else
'            If Me.XPChkSearchTypeCar.value = True Then
'            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rep1byCar.rpt"
''            Else
''            If Me.XPChkSearchTypeModel.value = True Then
'            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rep1byModel.rpt"
''            Else
'             If Me.XPChkSearchTypePlate.value = True Then
'            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rep1byPlate.rpt"
'            Else
'             If Me.XPChkSearchTypeMaint.value = True Then
'            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rep1byMaintain.rpt"
'            Else
'            If Me.RDrEqno.value = True Then
'            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rep1byrEQnO.rpt"
'            Else
'
'             StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rep1.rpt"
'            End If
'            End If
'            End If
'
''
 '           End If
 '            End If
 '            End If
 '       Else
 '              If Me.XPChkSearchTypeClient1.value = True Then
 '           StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rep1byclient.rpt"
 '           Else
 '           If Me.XPChkSearchTypeCar.value = True Then
 '           StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rep1byCar.rpt"
 '           Else
 '           If Me.XPChkSearchTypeModel.value = True Then
 '           StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rep1byModel.rpt"
 '           Else
 '            If Me.XPChkSearchTypePlate.value = True Then
 '           StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rep1byPlate.rpt"
 '           Else
 '            If Me.XPChkSearchTypeMaint.value = True Then
 '           StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rep1byMaintain.rpt"
 '           Else
 '           If Me.RDrEqno.value = True Then
 '            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rep1byrEQnO.rpt"
 '    Else
 '            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rep1.rpt"
 '           End If
 '           End If
 '           End If
 '           End If
 '
 '           End If
 '            End If
 '
 '       End If



 '   If Dir(StrFileName) = "" Then
 '       'GetMsgs 139, vbExclamation
 '       Screen.MousePointer = vbDefault
 '       Exit Function
 '   End If
'
'    Set RsData = New ADODB.Recordset
'    RsData.Open NoteSerial, Cn, adOpenStatic, adLockReadOnly, adCmdText
''
''    If RsData.BOF Or RsData.EOF Then
'        'GetMsgs 138, vbExclamation
'        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷"
'        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
'        RsData.Close
'        Set RsData = Nothing
'        Screen.MousePointer = vbDefault
'        Exit Function
'    End If
''
'    Screen.MousePointer = vbArrowHourglass
'    Set xReport = xApp.OpenReport(StrFileName)
'    xReport.Database.SetDataSource RsData
'
''    Dim cCompanyInfo As New ClsCompanyInfo
'
'    If SystemOptions.UserInterface = ArabicInterface Then
'        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName 'RPTCompany_Name_Arabic
'        ' xReport.ParameterFields(2).AddCurrentValue RPTComment_Arabic
'        StrReportTitle = "" '& StrAccountName
'        'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
'        '    StrReportTitle = StrReportTitle + " »œ«Ì… „‰ " & Format(Me.DTPickerAccFrom.value, "yyyy/M/d") & ""
'        'End If
'        'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
'        '    StrReportTitle = StrReportTitle + " ≈·Ï " & Format(Me.DTPickerAccTo.value, "yyyy/M/d") & " "
'        'End If
'    Else
'
'        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
'        'xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
'        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
'        StrReportTitle = ""
'        'If Me.DTPickerAccFrom.value <> Empty Or Me.DTPickerAccFrom.value <> Null Then
'        '    StrReportTitle = StrReportTitle + " From Date " & (Me.DTPickerAccFrom.value) & ""
'        'End If
'        'If Me.DTPickerAccTo.value <> Empty Or Me.DTPickerAccTo.value <> Null Then
'        '    StrReportTitle = StrReportTitle + " To Date :  " & (Me.DTPickerAccTo.value) & ""
'        'End If
'    End If
'
'    xReport.ParameterFields(3).AddCurrentValue user_name
'      '  xReport.ParameterFields(4).AddCurrentValue WriteNo(Format(val(TxtAdvanceValue.text), "0.00"), 0, True, ".")
'       ' xReport.ParameterFields(6).AddCurrentValue val(lbl(23).Caption)
'       '  xReport.ParameterFields(7).AddCurrentValue DBIssueDate.value
''    xReport.ParameterFields(8).AddCurrentValue IIf(IsNumeric(fg.TextMatrix(Me.fg.FixedRows, fg.ColIndex("PartValue"))), val(fg.TextMatrix(Me.fg.FixedRows, fg.ColIndex("PartValue"))), 0)
'' xReport.ParameterFields(9).AddCurrentValue val(lbl(22).Caption)
' ' xReport.ParameterFields(10).AddCurrentValue val(TxtDiscount.text)
'  ' xReport.ParameterFields(11).AddCurrentValue txtDiscountDES.text
'  Dim total As String
'  Dim totl As Double
' ' totl = val(LbToTalExtra.Caption) + val(Me.lbTotalMente.Caption)
' ' total = totl
' '  xReport.ParameterFields(12).AddCurrentValue Me.lbTotalMente.Caption
' '     xReport.ParameterFields(13).AddCurrentValue LbToTalExtra.Caption
' '       xReport.ParameterFields(14).AddCurrentValue total
'   ' xReport.ParameterFields(5).AddCurrentValue ToHijriDate(RsData("notedate").value)
'    xReport.reporttitle = StrReportTitle
'    xReport.EnableParameterPrompting = False
'    xReport.ApplicationName = App.Title
'    xReport.ReportAuthor = App.Title
'    Set CViewer = New ClsReportViewer
'    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

'    RsData.Close
'    Set RsData = Nothing
'    Screen.MousePointer = vbDefault


 
  
 
'End Function


 

Private Sub menue_Click(Index As Integer)
showsforms Index
Select Case Index

Case 16
GetData
Case 15
GetData1
End Select
End Sub

Private Sub Timer1_Timer()
'retrivgride
GetData
End Sub
 
Private Sub VSFlexGrid1_Click()

End Sub
