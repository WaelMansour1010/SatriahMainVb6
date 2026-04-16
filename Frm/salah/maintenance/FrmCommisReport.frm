VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SUITEC~1.OCX"
Begin VB.Form FrmCommisReport 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8355
   Icon            =   "FrmCommisReport.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   4515
   ScaleWidth      =   8355
   ShowInTaskbar   =   0   'False
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
      BackColor       =   &H00E2E9E9&
      Height          =   735
      Left            =   120
      TabIndex        =   30
      Top             =   3240
      Width           =   7935
      Begin VB.CommandButton menue 
         Height          =   555
         Index           =   5
         Left            =   1440
         Picture         =   "FrmCommisReport.frx":038A
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton menue 
         Height          =   555
         Index           =   4
         Left            =   2760
         Picture         =   "FrmCommisReport.frx":07D5
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton menue 
         Height          =   555
         Index           =   3
         Left            =   3480
         Picture         =   "FrmCommisReport.frx":0D2D
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton menue 
         Height          =   555
         Index           =   2
         Left            =   4920
         Picture         =   "FrmCommisReport.frx":11E6
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton menue 
         Height          =   555
         Index           =   1
         Left            =   3240
         Picture         =   "FrmCommisReport.frx":16B6
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   840
         Width           =   735
      End
      Begin VB.CommandButton menue 
         DownPicture     =   "FrmCommisReport.frx":1B57
         Height          =   555
         Index           =   0
         Left            =   7080
         Picture         =   "FrmCommisReport.frx":8E89
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton menue 
         DownPicture     =   "FrmCommisReport.frx":9430
         Height          =   555
         Index           =   6
         Left            =   5640
         Picture         =   "FrmCommisReport.frx":10762
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton menue 
         DownPicture     =   "FrmCommisReport.frx":10C03
         Height          =   555
         Index           =   7
         Left            =   4200
         Picture         =   "FrmCommisReport.frx":17F35
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton menue 
         Height          =   555
         Index           =   8
         Left            =   2040
         Picture         =   "FrmCommisReport.frx":187C5
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton menue 
         DownPicture     =   "FrmCommisReport.frx":18CAA
         Height          =   555
         Index           =   9
         Left            =   720
         Picture         =   "FrmCommisReport.frx":1FFDC
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton menue 
         DownPicture     =   "FrmCommisReport.frx":204FC
         Height          =   555
         Index           =   10
         Left            =   6360
         Picture         =   "FrmCommisReport.frx":20AE3
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton menue 
         DownPicture     =   "FrmCommisReport.frx":210CA
         Height          =   555
         Index           =   11
         Left            =   0
         Picture         =   "FrmCommisReport.frx":283FC
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.ComboBox DcbOrderStatus 
      Height          =   315
      Left            =   12240
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   1200
      Width           =   1935
   End
   Begin VB.ComboBox ComGranty 
      Height          =   315
      Left            =   12360
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   120
      Width           =   2175
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   " ﬁ—Ì— »Õ”»"
      Height          =   2805
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   480
      Width           =   8355
      Begin VB.TextBox TxtPlateNO 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   1680
         Width           =   6735
      End
      Begin MSDataListLib.DataCombo DcbClient 
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Top             =   1320
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSComCtl2.DTPicker DtpDateFrom 
         Height          =   330
         Left            =   3960
         TabIndex        =   15
         Top             =   2040
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   96468995
         CurrentDate     =   38887
      End
      Begin MSComCtl2.DTPicker DtpDateTo 
         Height          =   330
         Left            =   120
         TabIndex        =   16
         Top             =   2040
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   96468995
         CurrentDate     =   38887
      End
      Begin MSDataListLib.DataCombo DcbDept 
         Height          =   315
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcbSuper 
         Height          =   315
         Left            =   120
         TabIndex        =   18
         Top             =   600
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcbFitter 
         Height          =   315
         Left            =   120
         TabIndex        =   19
         Top             =   960
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton RdSortByDept 
         Height          =   375
         Left            =   6840
         TabIndex        =   23
         Top             =   2400
         Width           =   1455
         _Version        =   786432
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "›—“ »Õ”» «·ﬁ”„"
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton RdSortBySuper 
         Height          =   375
         Left            =   5040
         TabIndex        =   24
         Top             =   2400
         Width           =   1575
         _Version        =   786432
         _ExtentX        =   2778
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "›—“ »Õ”» «·„‘—›"
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton RdSortByFitter 
         Height          =   375
         Left            =   3360
         TabIndex        =   25
         Top             =   2400
         Width           =   1575
         _Version        =   786432
         _ExtentX        =   2778
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "›—“ »Õ”» «·›‰Ì"
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton RdSortByClient 
         Height          =   375
         Left            =   1920
         TabIndex        =   26
         Top             =   2400
         Width           =   1455
         _Version        =   786432
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "›—“ »Õ”» «·⁄„Ì·"
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton RdSortByPlateNo 
         Height          =   375
         Left            =   120
         TabIndex        =   29
         Top             =   2400
         Width           =   1695
         _Version        =   786432
         _ExtentX        =   2990
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "›—“ »Õ”» —ﬁ„ «··ÊÕ…"
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   "»Õ”» —ﬁ„ «··ÊÕ…"
         Height          =   195
         Index           =   8
         Left            =   7125
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   1680
         Width           =   1155
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   "»Õ”» «·›‰Ì"
         Height          =   195
         Index           =   6
         Left            =   7440
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   960
         Width           =   840
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   "»Õ”» «·„‘—›"
         Height          =   195
         Index           =   5
         Left            =   7305
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   "»Õ”» «·ﬁ”„"
         Height          =   195
         Index           =   0
         Left            =   7440
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   240
         Width           =   840
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·Ï  √—ÌŒ"
         Height          =   195
         Index           =   3
         Left            =   3030
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   2130
         Width           =   840
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰  √—ÌŒ"
         Height          =   195
         Index           =   4
         Left            =   7410
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   2130
         Width           =   780
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   "»Õ”» «·⁄„Ì·"
         Height          =   195
         Index           =   7
         Left            =   7395
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   1320
         Width           =   885
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   12360
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Text            =   " „ „Ê«›ﬁ… «·⁄„"
      Top             =   720
      Width           =   915
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   495
      Index           =   0
      Left            =   2880
      TabIndex        =   0
      Top             =   3960
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   873
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
      Height          =   495
      Index           =   1
      Left            =   1410
      TabIndex        =   1
      Top             =   3960
      Width           =   1365
      _ExtentX        =   2408
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
   Begin ImpulseButton.ISButton Cmd 
      Cancel          =   -1  'True
      Height          =   495
      Index           =   2
      Left            =   30
      TabIndex        =   2
      Top             =   3960
      Width           =   1365
      _ExtentX        =   2408
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
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   330
      Left            =   12480
      TabIndex        =   14
      Top             =   2400
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   582
      _Version        =   393216
      CheckBox        =   -1  'True
      Format          =   96468995
      CurrentDate     =   38887
   End
   Begin VB.Image ImgFavorites 
      Height          =   390
      Left            =   120
      Picture         =   "FrmCommisReport.frx":28F90
      Stretch         =   -1  'True
      Top             =   0
      Width           =   525
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      BackStyle       =   0  'Transparent
      Caption         =   " ﬁ«—Ì— «·⁄„Ê·« "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   9
      Left            =   2790
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   0
      Width           =   1980
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   2
      Height          =   1215
      Left            =   120
      Top             =   1200
      Width           =   2775
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      Caption         =   "Â–… «·‘«‘…  » ﬁ«—Ì— «·’Ì«‰…"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1170
      Index           =   2
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   1200
      Width           =   2655
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      ForeColor       =   &H00000080&
      Height          =   285
      Index           =   1
      Left            =   60
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   3060
      Width           =   1785
   End
End
Attribute VB_Name = "FrmCommisReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim DCboSearch As clsDCboSearch

Dim TTP As clstooltip

Dim TTD As clstooltipdemand
Dim Employee_account As String
Public Order As String
Public gr As String
Private Sub Check2_Click()

End Sub

Private Sub Cmd_Click(Index As Integer)

    Select Case Index

        Case 0
       

 GetData
            
        Case 1
            clear_all Me
DtpDateFrom.value = ""
DtpDateTo.value = ""

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




Private Sub Fg_Click()

 

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
Public Function FindRec(ByVal RecID As Long)
 
End Function
Private Sub Form_Activate()
   PutFormOnTop Me.hwnd
End Sub
Private Sub AddTip()
    Dim Wrap As String
    On Error GoTo ErrTrap
    Wrap = Chr(13) + Chr(10)
    Set TTP = New clstooltip

 With TTP
        .Create Me.hwnd, " «·«‰ ﬁ«· «·Ï ‘«‘… «· ”·Ì„ ··⁄„Ì·", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl menue(3), "‘«‘… ..." & Wrap & "  ··«‰ ﬁ«·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With
     With TTP
        .Create Me.hwnd, " «·«‰ ﬁ«· «·Ï ‘«‘… «·«Ê«„— «·„› ÊÕ… ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl menue(7), "‘«‘… ..." & Wrap & "  ··«‰ ﬁ«·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With
    With TTP
        .Create Me.hwnd, " «·«‰ ﬁ«· «·Ï ‘«‘…   ‰»ÌÂ«  «·⁄„·«¡", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl menue(4), "‘«‘… ..." & Wrap & "  ··«‰ ﬁ«·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With
     With TTP
        .Create Me.hwnd, " «·«‰ ﬁ«· «·Ï ‘«‘…  «· ﬁ«—Ì—", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl menue(5), "‘«‘… ..." & Wrap & "  ··«‰ ﬁ«·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With
      With TTP
        .Create Me.hwnd, " «·«‰ ﬁ«· «·Ï ‘«‘…  ’—› ﬁÿ⁄ «·€Ì«—", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl menue(2), "‘«‘… ..." & Wrap & "  ··«‰ ﬁ«·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With

       With TTP
        .Create Me.hwnd, " «·«‰ ﬁ«· «·Ï ‘«‘… ÿ·» ›Õ’ ﬂ„»ÌÊ —  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl menue(6), "‘«‘… ..." & Wrap & "  ··«‰ ﬁ«·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With
         With TTP
        .Create Me.hwnd, " «·«‰ ﬁ«· «·Ï ‘«‘…    ÿ·» ’Ì«‰…  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl menue(0), "‘«‘… ..." & Wrap & "  ··«‰ ﬁ«·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With
 
      With TTP
        .Create Me.hwnd, " «·«‰ ﬁ«· «·Ï ‘«‘…  «·⁄„Ê·«  «·„” Õﬁ…", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl menue(9), "‘«‘… ..." & Wrap & "  ··«‰ ﬁ«·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With
          With TTP
        .Create Me.hwnd, " «·«‰ ﬁ«· «·Ï ‘«‘…  «·⁄„·«¡  ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl menue(10), "‘«‘… ..." & Wrap & "  ··«‰ ﬁ«·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With
          With TTP
        .Create Me.hwnd, " «·«‰ ﬁ«· «·Ï ‘«‘…   ﬁ«—Ì— «·⁄„Ê·« ", 1, 15204351, -2147483630
        .MaxWidth = 4000
        .VisibleTime = 9000
        .DelayTime = 600
        .AddControl menue(11), "‘«‘… ..." & Wrap & "  ··«‰ ﬁ«·" & Wrap & " ›ﬁÿ ≈÷€ÿ Â‰«", True
    End With
 



    Exit Sub
ErrTrap:
End Sub

Private Sub ChangeLang()

    Cmd(1).Caption = "Delete"
    Cmd(0).Caption = "View Report"
   Cmd(2).Caption = "Exit"
  Me.Caption = "Reports of By Totals ReqNo "
Me.RdSortByClient.Caption = "Sort By Client"
Me.RdSortByClient.RightToLeft = True
Me.RdSortByDept.Caption = "Sort By Dept"
Me.RdSortByDept.RightToLeft = True
Me.RdSortByFitter.Caption = "Sort By Fitter"
Me.RdSortByFitter.RightToLeft = True
Me.RdSortByPlateNo.Caption = "Sort By Plate No"
Me.RdSortByPlateNo.RightToLeft = True
Me.RdSortBySuper.Caption = "Sort By Supper"
Me.RdSortBySuper.RightToLeft = True
lbl(3).Caption = "To Date"
lbl(4).Caption = "From Date"
Fra.Caption = "Report By"
lbl(8).Caption = "Plate No"
lbl(7).Caption = "Client"
lbl(6).Caption = "Fitter"
lbl(5).Caption = "Super Visor"
lbl(0).Caption = "Department"
lbl(9).Caption = "Reports of Commissions"
End Sub

Private Sub Form_Load()
    Dim GrdBack As ClsBackGroundPic
    Dim Dcombos As ClsDataCombos
    AddTip
        Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500

AddTip
    Set Dcombos = New ClsDataCombos
   ' Dcombos.GetEmployees Me.DCEmp_Name
     Dcombos.GetClientName Me.DcbClient
     Dcombos.GetFitte Me.DcbFitter
      Dcombos.GetSuperVisor Me.DcbSuper
     Dcombos.GetEmpDepartmentCar Me.DcbDept
    Set DCboSearch = New clsDCboSearch
    Set DCboSearch.Client = Me.DcbClient
    'Dcombos.GetUsers Me.DCUser
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Clear").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture

  '  CenterForm Me
'GetData
'    FormPostion Me, GetPostion
    Set GrdBack = New ClsBackGroundPic

 
 If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
    SetDtpickerDate Me.DtpDateFrom
    SetDtpickerDate Me.DtpDateTo

End Sub

Private Sub Form_Unload(Cancel As Integer)

    FormPostion Me, SavePostion
    Set DCboSearch = Nothing
End Sub

Public Sub GetData()
    Dim StrSQL As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
    gr = 9
    Order = 9



'StrSQL = StrSQL & " Where  (1 = 1)"
    BolBegine = False
    StrWhere = ""

StrSQL = " SELECT     dbo.TblCommisReceDetails.ID_Aut, dbo.TblCommisReceDetails.id2, dbo.TblCommisReceDetails.DateOp, dbo.TblCommisReceDetails.Total, "
StrSQL = StrSQL & "                      dbo.TblCommisReceDetails.Operation, dbo.TblCommisReceDetails.PerceTage, dbo.TblCommisReceDetails.PerceTageValue, dbo.TblCommisReceDetails.PriceFitter,"
StrSQL = StrSQL & "                      dbo.TblCommisReceDetails.Emp_ID, dbo.TblCommisReceDetails.plateno, dbo.TblCommisReceDetails.Deptid, dbo.TblEmpDepartments.DepartmentName,"
StrSQL = StrSQL & "                      dbo.TblEmpDepartments.DepartmentNamee, dbo.TblCommisReceDetails.empsuper, TblEmployee_1.Emp_Code AS Emp_CodeS,"
StrSQL = StrSQL & "                      TblEmployee_1.Emp_Name AS Emp_NameS, TblEmployee_1.Emp_Name1 AS Emp_Name1S, TblEmployee_1.Emp_Name2 AS Emp_Name2S,"
StrSQL = StrSQL & "                      TblEmployee_1.Emp_Name3 AS Emp_Name3S, TblEmployee_1.Emp_Name4 AS Emp_Name4S, TblEmployee_1.Emp_Namee AS Emp_NameeS,"
StrSQL = StrSQL & "                      TblEmployee_1.Emp_Namee1 AS Emp_Namee1S, TblEmployee_1.Emp_Namee2 AS Emp_Namee2S, TblEmployee_1.Emp_Namee3 AS Emp_Namee3S,"
StrSQL = StrSQL & "                      TblEmployee_1.Emp_Namee4 AS Emp_Namee4S, TblEmployee_1.Fullcode AS FullcodeS, dbo.TblCommisReceDetails.NoType, dbo.TblCarModels.Model,"
StrSQL = StrSQL & "                      dbo.TblCarModels.ModelE, dbo.TblCommisReceDetails.NoModel, dbo.TBLCarTypes.name, dbo.TBLCarTypes.namee, dbo.TblCommisReceDetails.NoOpE,"
StrSQL = StrSQL & "                      dbo.TblMaintenanceWork.name AS nameM, dbo.TblMaintenanceWork.namee AS nameeM, TblEmployee_2.Fullcode, TblEmployee_2.Emp_Namee4,"
StrSQL = StrSQL & "                      TblEmployee_2.Emp_Namee, TblEmployee_2.Emp_Namee1, TblEmployee_2.Emp_Namee2, TblEmployee_2.Emp_Namee3, TblEmployee_2.Emp_Code,"
StrSQL = StrSQL & "                      TblEmployee_2.Emp_Name, TblEmployee_2.Emp_Name1, TblEmployee_2.Emp_Name2, TblEmployee_2.Emp_Name3, TblEmployee_1.Emp_Name4,"
StrSQL = StrSQL & "                      dbo.TblCommisRece.id, dbo.TblCommisRece.RecordDate, dbo.TblCommisRece.NoteSerial, dbo.TblCommisRece.NoteSerial1, dbo.TblCommisRece.OldNoteSerial1,"
StrSQL = StrSQL & "                      dbo.TblCommisReceDetails.ClientCode, dbo.TblCustemers.CusName, dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode AS ClientFullcode"
StrSQL = StrSQL & " FROM         dbo.TblCommisReceDetails LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblEmployee TblEmployee_1 ON dbo.TblCommisReceDetails.empsuper = TblEmployee_1.Emp_ID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblCustemers ON dbo.TblCommisReceDetails.ClientCode = dbo.TblCustemers.code RIGHT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblCommisRece ON dbo.TblCommisReceDetails.id2 = dbo.TblCommisRece.id LEFT OUTER JOIN"
 StrSQL = StrSQL & "                     dbo.TblMaintenanceWork ON dbo.TblCommisReceDetails.NoOpE = dbo.TblMaintenanceWork.Id LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TBLCarTypes ON dbo.TblCommisReceDetails.NoType = dbo.TBLCarTypes.id LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblCarModels ON dbo.TblCommisReceDetails.NoModel = dbo.TblCarModels.Id LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblEmpDepartments ON dbo.TblCommisReceDetails.Deptid = dbo.TblEmpDepartments.DeparmentID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblEmployee TblEmployee_2 ON dbo.TblCommisReceDetails.Emp_ID = TblEmployee_2.Emp_ID"
StrSQL = StrSQL & "  where 1=1 "



   If Me.DcbClient.BoundText <> "" Then
     
            StrWhere = StrWhere & " AND dbo.TblCommisReceDetails.ClientCode=" & Me.DcbClient.BoundText & ""
      
    End If
    If Me.DcbDept.BoundText <> "" Then
     
            StrWhere = StrWhere & " AND dbo.TblCommisReceDetails.Deptid=" & Me.DcbDept.BoundText & ""
      
    End If
    If Me.DcbFitter.BoundText <> "" Then
     
            StrWhere = StrWhere & " AND dbo.TblCommisReceDetails.Emp_ID = " & Me.DcbFitter.BoundText & ""
      
    End If
  If Me.DcbSuper.BoundText <> "" Then
     
            StrWhere = StrWhere & " AND TblEmployee_1.Emp_ID=" & Me.DcbSuper.BoundText & ""
      
    End If
     If Me.TxtPlateNO.text <> "" Then
    
            StrWhere = StrWhere & " AND dbo.TblCommisReceDetails.plateno like '" & Me.TxtPlateNO.text & " '"
    
    End If


    If Not IsNull(Me.DtpDateFrom.value) Then
                   StrWhere = StrWhere & " AND dbo.TblCommisRece.RecordDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
      End If

    If Not IsNull(Me.DtpDateTo.value) Then
            StrWhere = StrWhere & " AND  dbo.TblCommisRece.RecordDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
     
    End If

    '-----------------------------------

    StrSQL = StrSQL & StrWhere
 
   StrSQL = StrSQL & " Order By dbo.TblCommisReceDetails.ID"
  
  
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
           ' Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
          '  Me.lbl(10).Caption = "Search Results=0"
        End If

        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «· ﬁ—Ì—"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.title
        Exit Sub
    Else
   '  Me.DTPicker1.value = Format(rs("DateAcutExite").value, "yyyy/M/d")
 rs.MoveFirst
' MsgBox rs("Granty").value
 print_report StrSQL
'print_report StrSQL
       ' With Me.Fg
       '     .Clear flexClearScrollable, flexClearEverything
       '     .Rows = .FixedRows
       '     .Rows = rs.RecordCount + .FixedRows
'
            If SystemOptions.UserInterface = ArabicInterface Then
             '   Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
               ' Me.lbl(10).Caption = "Search Results=" & rs.RecordCount
            End If

           
        '
        '    For i = .FixedRows To .Rows - 1
        '        .TextMatrix(i, .ColIndex("Serial")) = i
        '        .TextMatrix(i, .ColIndex("id")) = IIf(IsNull(rs("ID").value), "", rs("ID").value)
        '
        '        If Not (IsNull(rs("RecordDate").value)) Then
        '            .TextMatrix(i, .ColIndex("RecordDate")) = Format(rs("RecordDate").value, "yyyy/M/d")
        '        End If
        '
        '        .TextMatrix(i, .ColIndex("ClientName")) = IIf(IsNull(rs("ClientName").value), "", rs("ClientName").value)
        '        .TextMatrix(i, .ColIndex("Telephone")) = IIf(IsNull(rs("Telephone").value), "", rs("Telephone").value)
        '        .TextMatrix(i, .ColIndex("PlateNo")) = IIf(IsNull(rs("PlateNo").value), "", rs("PlateNo").value)
        '      '  .TextMatrix(i, .ColIndex("UserName")) = IIf(IsNull(rs("UserName").value), "", rs("UserName").value)
        '        rs.MoveNext
        '    Next i

        '    .AutoSize 0, .Cols - 1, False
         '   Me.lbl(1).Caption = .Aggregate(flexSTSum, .FixedRows, .ColIndex("AdvanceValue"), .Rows - 1, .ColIndex("AdvanceValue"))
        'End With

    End If

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
        If Me.RdSortByDept.value = True Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rep1ComissByDept.rpt"
            Else
            If Me.RdSortBySuper.value = True Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rep1ComissBySuper.rpt"
            Else
            If Me.RdSortByFitter.value = True Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rep1ComissByFitter.rpt"
            Else
             If Me.RdSortByClient.value = True Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rep1ComissByClient.rpt"
            Else
             If Me.RdSortByPlateNo.value = True Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rep1ComissByPlateNo.rpt"
            Else
          
             StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rep1ComissByAll.rpt"
      
           
            End If
            
            End If
            End If
             End If
            End If
        Else
               If Me.RdSortByDept.value = True Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rep1ComissByDept.rpt"
            Else
            If Me.RdSortBySuper.value = True Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rep1ComissBySuper.rpt"
            Else
            If Me.RdSortByFitter.value = True Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rep1ComissByFitter.rpt"
            Else
             If Me.RdSortByClient.value = True Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rep1ComissByClient.rpt"
            Else
             If Me.RdSortByPlateNo.value = True Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rep1ComissByPlateNo.rpt"
            Else
            
          
     
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rep1ComissByAll.rpt"
            
          
            
            
            End If
            End If
            End If
            End If
             End If
           
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
         xReport.ParameterFields(13).AddCurrentValue Me.DTPicker1.value
'    xReport.ParameterFields(8).AddCurrentValue IIf(IsNumeric(fg.TextMatrix(Me.fg.FixedRows, fg.ColIndex("PartValue"))), val(fg.TextMatrix(Me.fg.FixedRows, fg.ColIndex("PartValue"))), 0)
'Dim gr, order As Integer
 xReport.ParameterFields(14).AddCurrentValue Order
 xReport.ParameterFields(15).AddCurrentValue gr
 ' xReport.ParameterFields(15).AddCurrentValue gr
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

 
Private Sub ImgFavorites_Click()
AddTofaforites Me.name, Me.Caption, Me.Caption
End Sub

Private Sub menue_Click(Index As Integer)
showsforms Index
End Sub

