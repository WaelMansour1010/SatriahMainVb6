VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SuiteCtrls.ocx"
Begin VB.Form FrmHodaySearch 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " ﬁ«—Ì— ≈Ã«“«  «·„ÊŸ›Ì‰"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10125
   Icon            =   "FrmHodaySearch.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5190
   ScaleWidth      =   10125
   ShowInTaskbar   =   0   'False
   Begin XtremeSuiteControls.RadioButton XPChkSearchTypeClient1 
      Height          =   375
      Left            =   8280
      TabIndex        =   12
      Top             =   4170
      Width           =   1815
      _Version        =   786432
      _ExtentX        =   3201
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "›—“ »Õ”» «·„ÊŸ›"
      UseVisualStyle  =   -1  'True
      TextAlignment   =   1
      RightToLeft     =   -1  'True
   End
   Begin VB.ComboBox DcbOrderStatus 
      Height          =   315
      Left            =   12960
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   1200
      Width           =   1935
   End
   Begin VB.ComboBox ComGranty 
      Height          =   315
      Left            =   13440
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   120
      Width           =   2175
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Height          =   3555
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   480
      Width           =   10035
      Begin VB.CheckBox chkIsVac 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "Leave Balance"
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   1
         Left            =   -150
         RightToLeft     =   -1  'True
         TabIndex        =   47
         Top             =   750
         Width           =   3255
      End
      Begin VB.CheckBox chkIsVac 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "Display Casual Leaves"
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   0
         Left            =   -120
         RightToLeft     =   -1  'True
         TabIndex        =   46
         Top             =   450
         Width           =   3255
      End
      Begin VB.CheckBox ChkStatus 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈ŸÂ«— ﬂ· «·„ÊŸ›Ì‰ „⁄ «·„‰ ÂÌ… Œœ„« Â„"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   -120
         RightToLeft     =   -1  'True
         TabIndex        =   45
         Top             =   150
         Width           =   3255
      End
      Begin VB.TextBox TxtIqama 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   5160
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   1650
         Width           =   3855
      End
      Begin VB.TextBox TxtSearchCode 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   7920
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   240
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker DtpDateFrom 
         Height          =   360
         Left            =   6360
         TabIndex        =   7
         Top             =   3090
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   635
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   231145475
         CurrentDate     =   38887
      End
      Begin MSDataListLib.DataCombo DcboEmpName 
         Height          =   315
         Left            =   3240
         TabIndex        =   19
         Top             =   240
         Width           =   4635
         _ExtentX        =   8176
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcmbToDepart 
         Height          =   315
         Left            =   240
         TabIndex        =   20
         Top             =   1230
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo dcmbToProject 
         Height          =   315
         Left            =   240
         TabIndex        =   21
         Top             =   1650
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcmbToJob 
         Height          =   315
         Left            =   5160
         TabIndex        =   22
         Top             =   1230
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSComCtl2.DTPicker ToExpectedReturndate 
         Height          =   360
         Left            =   1320
         TabIndex        =   29
         Top             =   2610
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   635
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   231145473
         CurrentDate     =   38784
      End
      Begin MSComCtl2.DTPicker ToDeparDate 
         Height          =   360
         Left            =   1320
         TabIndex        =   30
         Top             =   2130
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   635
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   231145473
         CurrentDate     =   38784
      End
      Begin Dynamic_Byte.NourHijriCal ToDeparDateH 
         Height          =   315
         Left            =   120
         TabIndex        =   31
         Top             =   2130
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
      End
      Begin Dynamic_Byte.NourHijriCal ToExpectedReturndateH 
         Height          =   315
         Left            =   120
         TabIndex        =   32
         Top             =   2610
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
      End
      Begin MSComCtl2.DTPicker DtpDateTo 
         Height          =   360
         Left            =   1320
         TabIndex        =   37
         Top             =   3090
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   635
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   231145475
         CurrentDate     =   38887
      End
      Begin MSComCtl2.DTPicker fromExpectedReturndate 
         Height          =   360
         Left            =   6360
         TabIndex        =   38
         Top             =   2610
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   635
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   231145473
         CurrentDate     =   38784
      End
      Begin MSComCtl2.DTPicker FromDeparDate 
         Height          =   360
         Left            =   6360
         TabIndex        =   39
         Top             =   2130
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   635
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   231145473
         CurrentDate     =   38784
      End
      Begin Dynamic_Byte.NourHijriCal fromDeparDateH 
         Height          =   315
         Left            =   5160
         TabIndex        =   40
         Top             =   2130
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
      End
      Begin Dynamic_Byte.NourHijriCal fromExpectedReturndateH 
         Height          =   315
         Left            =   5160
         TabIndex        =   41
         Top             =   2610
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
      End
      Begin Dynamic_Byte.NourHijriCal DtpDateFromH 
         Height          =   315
         Left            =   5160
         TabIndex        =   42
         Top             =   3090
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
      End
      Begin Dynamic_Byte.NourHijriCal DtpDateToH 
         Height          =   315
         Left            =   120
         TabIndex        =   43
         Top             =   3090
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
      End
      Begin MSComCtl2.DTPicker txtDateGet 
         Height          =   360
         Left            =   3210
         TabIndex        =   48
         Top             =   660
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   635
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   231145473
         CurrentDate     =   38784
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·«ﬁ«„Â"
         Height          =   195
         Index           =   11
         Left            =   8805
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   1650
         Width           =   1140
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·Ï  «—ÌŒ «·”›— "
         Height          =   315
         Index           =   9
         Left            =   2760
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   2130
         Width           =   2115
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   " «·Ï  «—ÌŒ «·⁄Êœ… «·„ Êﬁ⁄"
         Height          =   315
         Index           =   8
         Left            =   2760
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   2610
         Width           =   2145
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   " «—ÌŒ «·⁄Êœ… «·„ Êﬁ⁄"
         Height          =   315
         Index           =   6
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   -240
         Width           =   1905
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "„‰  «—ÌŒ «·”›— "
         Height          =   315
         Index           =   45
         Left            =   8040
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   2130
         Width           =   1905
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "„‰  «—ÌŒ «·⁄Êœ… «·„ Êﬁ⁄"
         Height          =   315
         Index           =   46
         Left            =   8160
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   2610
         Width           =   1785
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Height          =   195
         Left            =   4770
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   360
         Width           =   45
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·ÊŸÌ›…"
         Height          =   195
         Index           =   5
         Left            =   8805
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   1290
         Width           =   1140
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·„Êﬁ⁄ "
         Height          =   195
         Index           =   0
         Left            =   3810
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   1740
         Width           =   1065
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï  «—ÌŒ «·⁄Êœ… «·›⁄·Ì"
         Height          =   195
         Index           =   3
         Left            =   3030
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   3180
         Width           =   1785
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰  «—ÌŒ «·⁄Êœ… «·›⁄·Ì"
         Height          =   195
         Index           =   4
         Left            =   8190
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   3060
         Width           =   1725
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·«œ«—…"
         Height          =   195
         Index           =   7
         Left            =   3765
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   1290
         Width           =   1140
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·„ÊŸ›"
         Height          =   195
         Left            =   9330
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   12960
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Text            =   " „ „Ê«›ﬁ… «·⁄„"
      Top             =   720
      Width           =   915
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   495
      Index           =   1
      Left            =   1410
      TabIndex        =   0
      Top             =   4650
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
      TabIndex        =   1
      Top             =   4650
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
   Begin XtremeSuiteControls.RadioButton XPChkSearchTypeCar 
      Height          =   375
      Left            =   4440
      TabIndex        =   13
      Top             =   4170
      Width           =   1815
      _Version        =   786432
      _ExtentX        =   3201
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "›—“ »Õ”» «·«œ«—…"
      UseVisualStyle  =   -1  'True
      TextAlignment   =   1
      RightToLeft     =   -1  'True
   End
   Begin XtremeSuiteControls.RadioButton XPChkSearchTypeModel 
      Height          =   375
      Left            =   2400
      TabIndex        =   14
      Top             =   4170
      Width           =   1935
      _Version        =   786432
      _ExtentX        =   3413
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "›—“ »Õ”» «·„Êﬁ⁄"
      UseVisualStyle  =   -1  'True
      TextAlignment   =   1
      RightToLeft     =   -1  'True
   End
   Begin XtremeSuiteControls.RadioButton XPChkSearchTypePlate 
      Height          =   375
      Left            =   6360
      TabIndex        =   15
      Top             =   4170
      Width           =   1935
      _Version        =   786432
      _ExtentX        =   3413
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "›—“ »Õ”» «·ÊŸÌ›…"
      UseVisualStyle  =   -1  'True
      TextAlignment   =   1
      RightToLeft     =   -1  'True
   End
   Begin XtremeSuiteControls.RadioButton XPChkSearchType 
      Height          =   375
      Left            =   240
      TabIndex        =   36
      Top             =   4170
      Visible         =   0   'False
      Width           =   1935
      _Version        =   786432
      _ExtentX        =   3413
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "›—“ »Õ”» ‰Ê⁄ «·«Ã«“…"
      UseVisualStyle  =   -1  'True
      TextAlignment   =   1
      RightToLeft     =   -1  'True
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   495
      Index           =   0
      Left            =   2760
      TabIndex        =   44
      Top             =   4650
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   873
      ButtonPositionImage=   1
      Caption         =   "⁄—÷"
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
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      BackStyle       =   0  'Transparent
      Caption         =   " ﬁ«—Ì— ≈Ã«“«  «·„ÊŸ›Ì‰ "
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
      Left            =   4515
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   120
      Width           =   3150
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      ForeColor       =   &H00000080&
      Height          =   285
      Index           =   1
      Left            =   60
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   3060
      Width           =   1785
   End
End
Attribute VB_Name = "FrmHodaySearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim DCboSearch As clsDCboSearch


Private Sub Cmd_Click(Index As Integer)

    Select Case Index

        Case 0
 

 GetData
            
        Case 1
              clear_all Me
DtpDateFrom.value = ""
DtpDateTo.value = ""
Me.ToDeparDate.value = ""
Me.FromDeparDate.value = ""
Me.ToExpectedReturndate.value = ""
Me.fromExpectedReturndate.value = ""

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


Private Sub DtpDateFrom_Change()
If Not IsNull(DtpDateFrom.value) Then
DtpDateFromH.value = ToHijriDate(DtpDateFrom.value)
End If
End Sub

Private Sub DtpDateFromH_LostFocus()
DtpDateFrom.value = ToGregorianDate(DtpDateFromH.value)
End Sub

Private Sub DtpDateTo_Change()
If Not IsNull(DtpDateTo.value) Then
DtpDateToH.value = ToHijriDate(DtpDateTo.value)
End If
End Sub

Private Sub DtpDateToH_LostFocus()
DtpDateTo.value = ToGregorianDate(DtpDateToH.value)
End Sub

Private Sub Form_Activate()
   PutFormOnTop Me.hWnd
End Sub
Private Sub ChangeLang()
ChkStatus.Caption = "All Employees With End Service"
 Cmd(1).Caption = "Delete"
    Cmd(0).Caption = "View Report"
   Cmd(2).Caption = "Exit"
  Me.Caption = "Reports of Vacation Of Employee "
Label5.Caption = Me.Caption
Label1.Caption = "Emp"
'lbl(10).Caption = "Pass No"
lbl(11).Caption = "Iqama No"
lbl(7).Caption = "Dept"
lbl(0).Caption = "Location"
'lbl(2).Caption = "Remraks"
lbl(5).Caption = "Position"
XPChkSearchTypeCar.RightToLeft = False
Me.XPChkSearchTypeCar.Caption = "By Dept"
Me.XPChkSearchTypeClient1.RightToLeft = False
Me.XPChkSearchTypeClient1.Caption = "By Emp"
Me.XPChkSearchTypePlate.RightToLeft = False
Me.XPChkSearchTypePlate.Caption = "By Job"
Me.XPChkSearchTypeModel.RightToLeft = False
Me.XPChkSearchTypeModel.Caption = "By Location"
XPChkSearchType.RightToLeft = False
XPChkSearchType.Caption = "By Type"
lbl(3).Caption = "To Date"
lbl(4).Caption = "From Date"
lbl(45).Caption = "From Date expected travel'"
lbl(9).Caption = "To Date expected travel'"
lbl(46).Caption = "From Date  return expected"
lbl(8).Caption = "To Date  return expected"
End Sub

Private Sub Form_Load()
    'Dim GrdBack As ClsBackGroundPic
    Dim Dcombos As ClsDataCombos
    
        Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500




    Set Dcombos = New ClsDataCombos
   ' Dcombos.GetEmployees Me.DCEmp_Name
  Set Dcombos = New ClsDataCombos
     'Dcombos.GetUsers Me.DCboUserName
     Dcombos.GetEmployees Me.DcboEmpName
    
     Dcombos.GetEmpDepartments Me.DcmbToDepart
    
   txtDateGet.value = Date
   Dcombos.GetEmpJobsTypes Me.DcmbToJob
   
   Dcombos.GetEmpLocations Me.dcmbToProject ' locatione
    Set DCboSearch = New clsDCboSearch
  '  Set DCboSearch.Client = Me.DcbClientname
    'Dcombos.GetUsers Me.DCUser
   Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Clear").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture


 
 If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
    SetDtpickerDate Me.DtpDateFrom
    SetDtpickerDate Me.DtpDateTo
    SetDtpickerDate Me.ToDeparDate
    SetDtpickerDate Me.FromDeparDate
     SetDtpickerDate Me.fromExpectedReturndate
    SetDtpickerDate Me.ToExpectedReturndate
    
DtpDateFrom.value = ""
DtpDateTo.value = ""
Me.ToDeparDate.value = ""
Me.FromDeparDate.value = ""
Me.ToExpectedReturndate.value = ""
Me.fromExpectedReturndate.value = ""

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



StrSQL = "SELECT     dbo.TblEmployee.Emp_ID AS Emp_IDH, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Name1, dbo.TblEmployee.Emp_Name2, "
StrSQL = StrSQL & "                      dbo.TblEmployee.Emp_Name3, dbo.TblEmployee.Emp_Name4, dbo.TblEmployee.Nationality, dbo.TblEmployee.NumEkama, dbo.TblEmployee.NumPasp,"
StrSQL = StrSQL & "                      dbo.TblEmployee.DepartmentID, dbo.TblEmpDepartments.DepartmentName, dbo.TblEmpDepartments.DepartmentNamee, dbo.TblEmployee.GroupID,"
StrSQL = StrSQL & "                      dbo.EmpGroupDep.GroupName, dbo.TblEmployee.JobTypeID, dbo.TblEmployee.Fullcode, dbo.TblEmployee.jopstatusid, dbo.TblVocationEntitlements.NoVacation,"
StrSQL = StrSQL & "                      dbo.TblVocationEntitlements.NoDayAct, dbo.TblVocationEntitlements.NoDayDelay, dbo.TblVocationEntitlements.AcuDateH, dbo.TblVocationEntitlements.AcuDate,"
StrSQL = StrSQL & "                      dbo.TblVocationEntitlements.stratDateH, dbo.TblVocationEntitlements.EndDateH, dbo.TblVocationEntitlements.stratDate, dbo.TblVocationEntitlements.EndDate,"
StrSQL = StrSQL & "                      dbo.TblEmpJobsTypes.JobTypeName , dbo.TblEmpJobsTypes.JobTypeNamee,0 as TypeTrans"
StrSQL = StrSQL & " FROM         dbo.TblVocationEntitlements LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblEmployee ON dbo.TblVocationEntitlements.EmpID = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblEmpJobsTypes ON dbo.TblEmployee.JobTypeID = dbo.TblEmpJobsTypes.JobTypeID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.EmpGroupDep ON dbo.TblEmployee.GroupID = dbo.EmpGroupDep.GroupID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblEmpDepartments ON dbo.TblEmployee.DepartmentID = dbo.TblEmpDepartments.DeparmentID"
StrSQL = StrSQL & " WHERE     (1 = 1) "
    BolBegine = False
    StrWhere = ""
   If ChkStatus.value = vbUnchecked Then
   StrSQL = StrSQL & " and dbo.TblEmployee.workstate = 1"
  End If
  

If (Me.TxtSearchCode.text) <> "" Then
'
            StrWhere = StrWhere & " AND dbo.TblEmployee.Fullcode like '%" & Me.TxtSearchCode.text & "%'"
'
    End If

       If Me.TxtIQAMA.text <> "" Then
     
            StrWhere = StrWhere & " AND dbo.TblEmployee.NumEkama='" & TxtIQAMA.text & "'"
      
    End If
   If Me.DcboEmpName.BoundText <> "" Then
     
            StrWhere = StrWhere & " AND dbo.TblEmployee.Emp_ID=" & Me.DcboEmpName.BoundText & ""
      
    End If
    If Me.DcmbToDepart.BoundText <> "" Then
     
            StrWhere = StrWhere & " AND dbo.TblEmployee.DepartmentID=" & Me.DcmbToDepart.BoundText & ""
      
    End If
    If Me.dcmbToProject.BoundText <> "" Then
     
            StrWhere = StrWhere & " AND dbo.TblEmployee.GroupID = " & Me.dcmbToProject.BoundText & ""
      
    End If
  If Me.DcmbToJob.BoundText <> "" Then
     
            StrWhere = StrWhere & " AND dbo.TblEmpJobsTypes.JobTypeID =" & val(Me.DcmbToJob.BoundText) & ""
      
    End If

    If Not IsNull(Me.DtpDateFrom.value) Then
                   StrWhere = StrWhere & " AND dbo.TblVocationEntitlements.EndDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
      End If
   If Not IsNull(Me.DtpDateTo.value) Then
           StrWhere = StrWhere & " AND  dbo.TblVocationEntitlements.EndDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""

   End If
 If Not IsNull(Me.FromDeparDate.value) Then
                   StrWhere = StrWhere & " AND dbo.TblVocationEntitlements.stratDate >=" & SQLDate(Me.FromDeparDate.value, True) & ""
      End If

    If Not IsNull(Me.ToDeparDate.value) Then
            StrWhere = StrWhere & " AND  dbo.TblVocationEntitlements.stratDate <=" & SQLDate(Me.ToDeparDate.value, True) & ""
     
    End If
 If Not IsNull(Me.fromExpectedReturndate.value) Then
                   StrWhere = StrWhere & " AND dbo.TblVocationEntitlements.AcuDate >=" & SQLDate(Me.fromExpectedReturndate.value, True) & ""
      End If

    If Not IsNull(Me.ToExpectedReturndate.value) Then
            StrWhere = StrWhere & " AND  dbo.TblVocationEntitlements.AcuDate <=" & SQLDate(Me.ToExpectedReturndate.value, True) & ""
     
    End If
    

    '-----------------------------------
StrSQL = StrSQL & StrWhere
 
 '  StrSQL = StrSQL & " Order By dbo.TblEmployee.Emp_ID"
  
  If chkIsVac(0) Then
  StrSQL = StrSQL & " Union All SELECT     dbo.TblEmployee.Emp_ID AS Emp_IDH, dbo.TblEmployee.Emp_Code, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Name1, dbo.TblEmployee.Emp_Name2, "
StrSQL = StrSQL & "                      dbo.TblEmployee.Emp_Name3, dbo.TblEmployee.Emp_Name4, dbo.TblEmployee.Nationality, dbo.TblEmployee.NumEkama, dbo.TblEmployee.NumPasp,"
StrSQL = StrSQL & "                      dbo.TblEmployee.DepartmentID, dbo.TblEmpDepartments.DepartmentName, dbo.TblEmpDepartments.DepartmentNamee, dbo.TblEmployee.GroupID,"
StrSQL = StrSQL & "                      dbo.EmpGroupDep.GroupName, dbo.TblEmployee.JobTypeID, dbo.TblEmployee.Fullcode, dbo.TblEmployee.jopstatusid, NoVacation = NoDay,"
StrSQL = StrSQL & "                      NoDay as NoDayAct , NoDayDelay = 0, AcuDateH = '', AcuDate=ToDate,"
StrSQL = StrSQL & "                      stratDateH = '', EndDateH = '', FromDate as stratDate, ToDate as EndDate,"
StrSQL = StrSQL & "                      dbo.TblEmpJobsTypes.JobTypeName , dbo.TblEmpJobsTypes.JobTypeNamee,TypeTrans"
StrSQL = StrSQL & " FROM         dbo.TblEmpPassOver LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblEmployee ON dbo.TblEmpPassOver.Emp_ID = dbo.TblEmployee.Emp_ID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblEmpJobsTypes ON dbo.TblEmployee.JobTypeID = dbo.TblEmpJobsTypes.JobTypeID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.EmpGroupDep ON dbo.TblEmployee.GroupID = dbo.EmpGroupDep.GroupID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblEmpDepartments ON dbo.TblEmployee.DepartmentID = dbo.TblEmpDepartments.DeparmentID"
StrSQL = StrSQL & " WHERE     (TypeTrans = 2 Or TypeTrans = 3 ) "
    BolBegine = False
    StrWhere = ""
   If ChkStatus.value = vbUnchecked Then
   StrSQL = StrSQL & " and dbo.TblEmployee.workstate = 1"
  End If
  

If (Me.TxtSearchCode.text) <> "" Then
'
            StrWhere = StrWhere & " AND dbo.TblEmployee.Fullcode like '%" & Me.TxtSearchCode.text & "%'"
'
    End If

       If Me.TxtIQAMA.text <> "" Then
     
            StrWhere = StrWhere & " AND dbo.TblEmployee.NumEkama='" & TxtIQAMA.text & "'"
      
    End If
   If Me.DcboEmpName.BoundText <> "" Then
     
            StrWhere = StrWhere & " AND dbo.TblEmployee.Emp_ID=" & Me.DcboEmpName.BoundText & ""
      
    End If
    If Me.DcmbToDepart.BoundText <> "" Then
     
            StrWhere = StrWhere & " AND dbo.TblEmployee.DepartmentID=" & Me.DcmbToDepart.BoundText & ""
      
    End If
    If Me.dcmbToProject.BoundText <> "" Then
     
            StrWhere = StrWhere & " AND dbo.TblEmployee.GroupID = " & Me.dcmbToProject.BoundText & ""
      
    End If
  If Me.DcmbToJob.BoundText <> "" Then
     
            StrWhere = StrWhere & " AND dbo.TblEmpJobsTypes.JobTypeID =" & val(Me.DcmbToJob.BoundText) & ""
      
    End If

    If Not IsNull(Me.DtpDateFrom.value) Then
                   StrWhere = StrWhere & " AND dbo.TblEmpPassOver.ToDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
      End If
   If Not IsNull(Me.DtpDateTo.value) Then
           StrWhere = StrWhere & " AND  dbo.TblEmpPassOver.ToDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""

   End If
 If Not IsNull(Me.FromDeparDate.value) Then
                   StrWhere = StrWhere & " AND dbo.TblEmpPassOver.FromDate >=" & SQLDate(Me.FromDeparDate.value, True) & ""
      End If

    If Not IsNull(Me.ToDeparDate.value) Then
            StrWhere = StrWhere & " AND  dbo.TblEmpPassOver.FromDate <=" & SQLDate(Me.ToDeparDate.value, True) & ""
     
    End If
 If Not IsNull(Me.fromExpectedReturndate.value) Then
            '       StrWhere = StrWhere & " AND dbo.TblEmpPassOver.Actualouttime >=" & SQLDate(Me.fromExpectedReturndate.value, True) & ""
      End If

    If Not IsNull(Me.ToExpectedReturndate.value) Then
        '    StrWhere = StrWhere & " AND  dbo.TblEmpPassOver.Actualouttime <=" & SQLDate(Me.ToExpectedReturndate.value, True) & ""
     
    End If
    

    '-----------------------------------
StrSQL = StrSQL & StrWhere
 
   StrSQL = StrSQL & " Order By dbo.TblEmployee.Emp_ID"
End If




If chkIsVac(1) Then
StrWhere = ""




'StrSQL = "SELECT TblEmployee.Emp_ID, TblEmployee.Emp_Namee, " & _
'         "EJ.JobTypeNamee AS JobTypeNamee, EG.GroupNamee AS GroupName, " & _
'         "ED.DepartmentNamee AS DepartmentName, " & _
'         "(SELECT SUM(TblInstalVacationDet.VacBalance) " & _
'         " FROM TblInstalVacationDet WHERE EmpID = TblEmployee.Emp_ID) AS TotalVacationBalance, " & _
'         "(SELECT SUM(TblEmpPassOver.AbceDay) " & _
'         " FROM TblEmpPassOver WHERE TypeTrans = 2 AND Emp_id = TblEmployee.Emp_ID) AS TotalAbsentDays, " & _
'         "(SELECT DATEDIFF(DAY, " & _
'         " CASE WHEN C.Contract_date < DATEFROMPARTS(YEAR(GETDATE()), 1, 1) " & _
'         " THEN DATEFROMPARTS(YEAR(GETDATE()), 1, 1) ELSE C.Contract_date END, GETDATE()) " & _
'         " * (C.Holiday_period_no / 365.0) " & _
'         " FROM contract C WHERE C.Emp_id = TblEmployee.Emp_ID) AS EarnedVacationDays, " & _
'         "(SELECT SUM(TCRD.NoofDays) " & _
'         " FROM TblChangedComponentRegister TCR " & _
'         " INNER JOIN MOFRAD M ON M.id = TCR.ComponentID " & _
'         " AND M.AddOrDiscount = 1 AND M.Unit = 1 " & _
'         " INNER JOIN TblChangedComponentRegisterDetails TCRD " & _
'         " ON TCRD.ChangedComponentid = TCR.ChangedComponentid " & _
'         " WHERE TCR.Actualyear = YEAR(GETDATE()) AND TCRD.Emp_id = TblEmployee.Emp_ID) AS TotalComponentDays " & _
'         "FROM TblEmployee  " & _
'         "LEFT JOIN TblEmpJobsTypes EJ ON TblEmployee.JobTypeID = EJ.JobTypeID " & _
'         "LEFT JOIN EmpGroupDep EG ON TblEmployee.GroupID = EG.GroupID " & _
'         "LEFT JOIN TblEmpDepartments ED ON TblEmployee.DepartmentID = ED.DeparmentID"

Dim txtdate As String

txtdate = SQLDate(Me.txtDateGet.value, True)
' ?????? ??? ???? ??????? ?? TextBox


' ??????? SQL ?? ??????? GETDATE() ????? TextBox
StrSQL = "SELECT TblEmployee.Emp_ID, TblEmployee.Emp_Namee,TblEmployee.Fullcode, " & _
         "EJ.JobTypeNamee AS JobTypeNamee, EG.GroupNamee AS GroupName, " & _
         "ED.DepartmentNamee AS DepartmentName, " & _
         "(SELECT SUM(TblInstalVacationDet.VacBalance) " & _
         " FROM TblInstalVacationDet WHERE EmpID = TblEmployee.Emp_ID) AS TotalVacationBalance, " & _
         "(SELECT SUM(TblEmpPassOver.AbceDay) " & _
         " FROM TblEmpPassOver WHERE TypeTrans = 2 and TypeDisc = 1 AND Emp_id = TblEmployee.Emp_ID) AS TotalAbsentDays, " & _
        "(SELECT SUM(TblEmpPassOver.AbceDay) " & _
         " FROM TblEmpPassOver WHERE TypeTrans = 2 and TypeDisc = 2 AND Emp_id = TblEmployee.Emp_ID) AS TotalVacDays, " & _
         "(SELECT DATEDIFF(DAY, " & _
         " CASE WHEN C.Contract_date < DATEFROMPARTS(YEAR(" & txtdate & "), 1, 1) " & _
         " THEN DATEFROMPARTS(YEAR(" & txtdate & "), 1, 1) ELSE C.Contract_date END, " & txtdate & ") " & _
         " * (C.Holiday_period_no / 365.0) " & _
         " FROM contract C WHERE C.Emp_id = TblEmployee.Emp_ID) AS EarnedVacationDays, " & _
         "(SELECT SUM(TCRD.NoofDays) " & _
         " FROM TblChangedComponentRegister TCR " & _
         " INNER JOIN MOFRAD M ON M.id = TCR.ComponentID " & _
         " AND M.AddOrDiscount = 1 AND M.Unit = 1 " & _
         " INNER JOIN TblChangedComponentRegisterDetails TCRD " & _
         " ON TCRD.ChangedComponentid = TCR.ChangedComponentid " & _
         " WHERE TCR.Actualyear = YEAR(" & txtdate & ") AND TCRD.Emp_id = TblEmployee.Emp_ID) AS TotalComponentDays " & _
         "FROM TblEmployee  " & _
         "LEFT JOIN TblEmpJobsTypes EJ ON TblEmployee.JobTypeID = EJ.JobTypeID " & _
         "LEFT JOIN EmpGroupDep EG ON TblEmployee.GroupID = EG.GroupID " & _
         "LEFT JOIN TblEmpDepartments ED ON TblEmployee.DepartmentID = ED.DeparmentID"
'TypeDisc


    BolBegine = False
    StrWhere = " Where 1 = 1 "
   If ChkStatus.value = vbUnchecked Then
        StrWhere = StrWhere & " and dbo.TblEmployee.workstate = 1"
  End If
  

If (Me.TxtSearchCode.text) <> "" Then
'
            StrWhere = StrWhere & " AND dbo.TblEmployee.Fullcode like '%" & Me.TxtSearchCode.text & "%'"
'
    End If

       If Me.TxtIQAMA.text <> "" Then
     
            StrWhere = StrWhere & " AND dbo.TblEmployee.NumEkama='" & TxtIQAMA.text & "'"
      
    End If
   If Me.DcboEmpName.BoundText <> "" Then
     
            StrWhere = StrWhere & " AND dbo.TblEmployee.Emp_ID=" & Me.DcboEmpName.BoundText & ""
      
    End If
    If Me.DcmbToDepart.BoundText <> "" Then
     
            StrWhere = StrWhere & " AND dbo.TblEmployee.DepartmentID=" & Me.DcmbToDepart.BoundText & ""
      
    End If
    If Me.dcmbToProject.BoundText <> "" Then
     
            StrWhere = StrWhere & " AND dbo.TblEmployee.GroupID = " & Me.dcmbToProject.BoundText & ""
      
    End If
  If Me.DcmbToJob.BoundText <> "" Then
     
            StrWhere = StrWhere & " AND dbo.TblEmpJobsTypes.JobTypeID =" & val(Me.DcmbToJob.BoundText) & ""
      
    End If

    If Not IsNull(Me.DtpDateFrom.value) Then
                   StrWhere = StrWhere & " AND dbo.TblEmpPassOver.ToDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
      End If
   If Not IsNull(Me.DtpDateTo.value) Then
           StrWhere = StrWhere & " AND  dbo.TblEmpPassOver.ToDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""

   End If
 If Not IsNull(Me.FromDeparDate.value) Then
                   StrWhere = StrWhere & " AND dbo.TblEmpPassOver.FromDate >=" & SQLDate(Me.FromDeparDate.value, True) & ""
      End If

    If Not IsNull(Me.ToDeparDate.value) Then
            StrWhere = StrWhere & " AND  dbo.TblEmpPassOver.FromDate <=" & SQLDate(Me.ToDeparDate.value, True) & ""
     
    End If
 If Not IsNull(Me.fromExpectedReturndate.value) Then
            '       StrWhere = StrWhere & " AND dbo.TblEmpPassOver.Actualouttime >=" & SQLDate(Me.fromExpectedReturndate.value, True) & ""
      End If

    If Not IsNull(Me.ToExpectedReturndate.value) Then
        '    StrWhere = StrWhere & " AND  dbo.TblEmpPassOver.Actualouttime <=" & SQLDate(Me.ToExpectedReturndate.value, True) & ""
     
    End If
    

    '-----------------------------------
StrSQL = StrSQL & StrWhere
 
   StrSQL = StrSQL & " Order By dbo.TblEmployee.Emp_ID"
End If
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
           ' Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
          '  Me.lbl(10).Caption = "Search Results=0"
        End If
If SystemOptions.UserInterface = ArabicInterface Then
        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «· ﬁ—Ì—"
    Else
    Msg = "No Data"
    End If
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    Else
    rs.MoveFirst
 print_report StrSQL

            If SystemOptions.UserInterface = ArabicInterface Then
             '   Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=" & rs.RecordCount
            ElseIf SystemOptions.UserInterface = EnglishInterface Then
               ' Me.lbl(10).Caption = "Search Results=" & rs.RecordCount
            End If
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

           If Me.XPChkSearchTypeClient1.value = True Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "repByEmpHoldayEmpName.rpt"
            Else
            If Me.XPChkSearchTypeCar.value = True Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "repByEmpHoldayDept.rpt"
            Else
            If Me.XPChkSearchTypeModel.value = True Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "repByEmpHoldayLocation.rpt"
            Else
             If Me.XPChkSearchTypePlate.value = True Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "repByEmpHoldayJob.rpt"
            Else
               If XPChkSearchType.value = True Then
              StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "repByEmpHoldayType.rpt"
            Else
             StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "repByEmpHoldayAll.rpt"
         End If
            End If
            End If
 
             End If
           
        End If
        If chkIsVac(1) Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "repByEmpHoldayEmpName2.rpt"
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
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
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
        StrReportTitle = "" '& StrAccountName
    Else
 
        xReport.ParameterFields(1).AddCurrentValue cCompanyInfo.ArabCompanyName ' RPTCompany_Name_Eng
        'xReport.ParameterFields(2).AddCurrentValue RPTComment_Eng
        xReport.ParameterFields(4).AddCurrentValue get_branch_name(val(my_branch))
        StrReportTitle = ""
    End If
    xReport.ParameterFields(3).AddCurrentValue user_name
  Dim total As String
  Dim totl As Double
    xReport.reporttitle = StrReportTitle
    xReport.EnableParameterPrompting = False
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName
    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault
End Function

Private Sub FromDeparDate_Change()
If Not IsNull(FromDeparDate.value) Then
fromDeparDateH.value = ToHijriDate(FromDeparDate.value)
End If
End Sub

Private Sub fromDeparDateH_LostFocus()
FromDeparDate.value = ToGregorianDate(fromDeparDateH.value)
End Sub

Private Sub fromExpectedReturndate_Change()
If Not IsNull(fromExpectedReturndate.value) Then
fromExpectedReturndateH.value = ToHijriDate(fromExpectedReturndate.value)
End If
End Sub

Private Sub fromExpectedReturndateH_LostFocus()
fromExpectedReturndate.value = ToGregorianDate(fromExpectedReturndateH.value)
End Sub

Private Sub ToDeparDate_Change()
If Not IsNull(ToDeparDate.value) Then
ToDeparDateH.value = ToHijriDate(ToDeparDate.value)
End If
End Sub

Private Sub ToDeparDateH_LostFocus()
ToDeparDate.value = ToGregorianDate(ToDeparDateH.value)
End Sub

Private Sub ToExpectedReturndate_Change()
If Not IsNull(ToExpectedReturndate.value) Then
ToExpectedReturndateH.value = ToHijriDate(ToExpectedReturndate.value)
End If
End Sub

Private Sub ToExpectedReturndateH_LostFocus()
ToExpectedReturndate.value = ToGregorianDate(ToExpectedReturndateH.value)
End Sub

Private Sub TxtSearchCode_KeyPress(KeyAscii As Integer)
Dim EmpID As Integer

    If KeyAscii = vbKeyReturn Then
        GetEmployeeIDFromCode TxtSearchCode.text, EmpID
        DcboEmpName.BoundText = EmpID
    End If
End Sub
