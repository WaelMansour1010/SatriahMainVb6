VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SUITEC~1.OCX"
Begin VB.Form FrmCarReports 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " ﬁ«—Ì— «·’Ì«‰…"
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10260
   Icon            =   "FrmCarReport.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   4980
   ScaleWidth      =   10260
   ShowInTaskbar   =   0   'False
   Begin XtremeSuiteControls.RadioButton XPChkSearchTypeClient1 
      Height          =   495
      Left            =   8400
      TabIndex        =   33
      Top             =   3840
      Width           =   1815
      _Version        =   786432
      _ExtentX        =   3201
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "›—“ »Õ”» «”„ «·⁄„Ì·"
      UseVisualStyle  =   -1  'True
      TextAlignment   =   1
      RightToLeft     =   -1  'True
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2E9E9&
      Caption         =   " ﬁ—Ì— »Õ”»"
      Height          =   645
      Left            =   3000
      RightToLeft     =   -1  'True
      TabIndex        =   27
      Top             =   600
      Width           =   7155
      Begin XtremeSuiteControls.RadioButton RDGRANTY 
         Height          =   255
         Left            =   4920
         TabIndex        =   28
         Top             =   240
         Width           =   1575
         _Version        =   786432
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "»÷„«‰"
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton RDRETURNM 
         Height          =   255
         Left            =   1320
         TabIndex        =   29
         Top             =   240
         Width           =   1575
         _Version        =   786432
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "≈⁄«œ… ≈’·«Õ"
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton RDWITHOUTGRANTY 
         Height          =   255
         Left            =   3120
         TabIndex        =   30
         Top             =   240
         Width           =   1575
         _Version        =   786432
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "»œÊ‰ ÷„«‰"
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton RDALL 
         Height          =   255
         Left            =   -240
         TabIndex        =   31
         Top             =   240
         Width           =   1335
         _Version        =   786432
         _ExtentX        =   2355
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "«·ﬂ·"
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
   End
   Begin VB.ComboBox DcbOrderStatus 
      Height          =   315
      Left            =   11640
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   1200
      Width           =   1935
   End
   Begin VB.ComboBox ComGranty 
      Height          =   315
      Left            =   11040
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   120
      Width           =   2175
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   " ﬁ—Ì— »Õ”»"
      Height          =   2085
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   1800
      Width           =   10155
      Begin VB.TextBox Txtmobile 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   44
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox TxtPlateNO 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   6600
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   960
         Width           =   2055
      End
      Begin MSDataListLib.DataCombo DcbClientname 
         Height          =   315
         Left            =   3360
         TabIndex        =   10
         Top             =   240
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcbCarType 
         Height          =   315
         Left            =   3360
         TabIndex        =   12
         Top             =   600
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcbCarModel 
         Height          =   315
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSComCtl2.DTPicker DtpDateFrom 
         Height          =   330
         Left            =   120
         TabIndex        =   19
         Top             =   960
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   95354883
         CurrentDate     =   38887
      End
      Begin MSComCtl2.DTPicker DtpDateTo 
         Height          =   330
         Left            =   120
         TabIndex        =   21
         Top             =   1440
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   95354883
         CurrentDate     =   38887
      End
      Begin MSDataListLib.DataCombo DCBMinten 
         Height          =   315
         Left            =   3360
         TabIndex        =   26
         Top             =   960
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSComCtl2.DTPicker DtEnd 
         Height          =   330
         Left            =   3360
         TabIndex        =   39
         Top             =   1440
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   95354883
         CurrentDate     =   38887
      End
      Begin MSComCtl2.DTPicker DtStart 
         Height          =   330
         Left            =   6600
         TabIndex        =   40
         Top             =   1440
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   95354883
         CurrentDate     =   38887
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·ÃÊ«·"
         Height          =   195
         Index           =   6
         Left            =   2475
         RightToLeft     =   -1  'True
         TabIndex        =   45
         Top             =   240
         Width           =   720
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " √—ÌŒ «·Œ—ÊÃ"
         Height          =   195
         Index           =   5
         Left            =   5490
         RightToLeft     =   -1  'True
         TabIndex        =   42
         Top             =   1530
         Width           =   1005
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " √—ÌŒ «·œŒÊ·"
         Height          =   195
         Index           =   0
         Left            =   9195
         RightToLeft     =   -1  'True
         TabIndex        =   41
         Top             =   1410
         Width           =   720
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï"
         Height          =   195
         Index           =   3
         Left            =   2550
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   1410
         Width           =   480
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   195
         Index           =   4
         Left            =   2490
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   1050
         Width           =   540
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «··ÊÕ…"
         Height          =   195
         Index           =   7
         Left            =   9195
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   960
         Width           =   720
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰Ê⁄ «·’Ì«‰…"
         Height          =   195
         Left            =   5655
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   960
         Width           =   780
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÿ—«“"
         Height          =   195
         Left            =   2730
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   600
         Width           =   300
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   "”Ì«—… "
         Height          =   195
         Left            =   9195
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   600
         Width           =   720
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   "«”„ «·⁄„Ì·"
         Height          =   195
         Left            =   9195
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   240
         Width           =   720
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E2E9E9&
      Caption         =   " ﬁ—Ì— »Õ”»"
      Height          =   645
      Left            =   3000
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   1200
      Width           =   7155
      Begin XtremeSuiteControls.RadioButton RdNew 
         Height          =   255
         Left            =   6000
         TabIndex        =   6
         Top             =   240
         Width           =   975
         _Version        =   786432
         _ExtentX        =   1720
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "ÃœÌœ"
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton RdFinal 
         Height          =   255
         Left            =   2880
         TabIndex        =   7
         Top             =   240
         Width           =   1575
         _Version        =   786432
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   " „ «‰Â«¡ «·«’·«Õ"
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton RdAccept 
         Height          =   255
         Left            =   4440
         TabIndex        =   8
         Top             =   240
         Width           =   1575
         _Version        =   786432
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   " „ „Ê«›ﬁ… «·⁄„Ì·"
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton RdAll2 
         Height          =   255
         Left            =   -240
         TabIndex        =   25
         Top             =   240
         Width           =   1335
         _Version        =   786432
         _ExtentX        =   2355
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "«·ﬂ·"
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton RdNotAccept 
         Height          =   255
         Left            =   1200
         TabIndex        =   43
         Top             =   240
         Width           =   1575
         _Version        =   786432
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "⁄œ„ „Ê«›ﬁ… «·⁄„Ì·"
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   11520
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
      Top             =   4440
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
      Top             =   4440
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
      Top             =   4440
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
      Left            =   6360
      TabIndex        =   34
      Top             =   3840
      Width           =   1815
      _Version        =   786432
      _ExtentX        =   3201
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "›—“ »Õ”» ‰Ê⁄ «·„⁄œÂ/«·”Ì«—…"
      UseVisualStyle  =   -1  'True
      TextAlignment   =   1
      RightToLeft     =   -1  'True
   End
   Begin XtremeSuiteControls.RadioButton XPChkSearchTypeModel 
      Height          =   375
      Left            =   4320
      TabIndex        =   35
      Top             =   3840
      Width           =   1935
      _Version        =   786432
      _ExtentX        =   3413
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "›—“ »Õ”» ÿ—«“ «·„⁄œÂ/«·”Ì«—…"
      UseVisualStyle  =   -1  'True
      TextAlignment   =   1
      RightToLeft     =   -1  'True
   End
   Begin XtremeSuiteControls.RadioButton XPChkSearchTypePlate 
      Height          =   375
      Left            =   2160
      TabIndex        =   36
      Top             =   3840
      Width           =   1935
      _Version        =   786432
      _ExtentX        =   3413
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "›—“ »Õ”» —ﬁ„ «··ÊÕ…"
      UseVisualStyle  =   -1  'True
      TextAlignment   =   1
      RightToLeft     =   -1  'True
   End
   Begin XtremeSuiteControls.RadioButton XPChkSearchTypeMaint 
      Height          =   375
      Left            =   0
      TabIndex        =   37
      Top             =   3840
      Width           =   1935
      _Version        =   786432
      _ExtentX        =   3413
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "›—“ »Õ”» ‰Ê⁄ «·’Ì«‰…"
      UseVisualStyle  =   -1  'True
      TextAlignment   =   1
      RightToLeft     =   -1  'True
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      BackStyle       =   0  'Transparent
      Caption         =   " ﬁ«—Ì— ’Ì«‰… «·„⁄œÂ/«·”Ì«—…"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   465
      Left            =   4320
      RightToLeft     =   -1  'True
      TabIndex        =   38
      Top             =   120
      Width           =   3120
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   2
      Height          =   1215
      Left            =   120
      Top             =   480
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
      TabIndex        =   32
      Top             =   480
      Width           =   2775
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
Attribute VB_Name = "FrmCarReports"
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
Me.DtStart.value = ""
Me.DtEnd.value = ""
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

Private Sub Form_Activate()
   PutFormOnTop Me.hWnd
End Sub
Private Sub ChangeLang()

    Cmd(1).Caption = "Delete"
    Cmd(0).Caption = "View Report"
   Cmd(2).Caption = "Exit"
  Me.Caption = "Reports of Car  Maintenance "
RDGRANTY.RightToLeft = False
RDGRANTY.Caption = "Granty"
RDWITHOUTGRANTY.RightToLeft = False
RDWITHOUTGRANTY.Caption = "Without Granty"
RDRETURNM.RightToLeft = False
RDRETURNM.Caption = "Re Maintenance"
RDALL.RightToLeft = False
RDALL.Caption = "All"
Frame2.Caption = "Report By"
Frame1.Caption = "Report By"
Fra.Caption = "Report By"
Label1.Caption = "ClientName"
lbl(0).Caption = "Date Enter"
lbl(5).Caption = "Date End"
Me.XPChkSearchTypeCar.RightToLeft = False
Me.RdNotAccept.RightToLeft = False
Me.RdNotAccept.Caption = "Not Accept"
Me.XPChkSearchTypeCar.Caption = "By Car"
Me.XPChkSearchTypeClient1.RightToLeft = False
Me.XPChkSearchTypeClient1.Caption = "By Client"
Me.XPChkSearchTypeMaint.RightToLeft = False
Me.XPChkSearchTypeMaint.Caption = "By Maintenance"
Me.XPChkSearchTypeModel.RightToLeft = False
Me.XPChkSearchTypeModel.Caption = "By Model"
Me.XPChkSearchTypePlate.RightToLeft = False
Me.XPChkSearchTypePlate.Caption = "By PlateNo"
lbl(7).Caption = "PlateNo"
'lbl(0).Caption = "From"
lbl(2).Caption = "This is Monitor for Reports of Maintenance "
lbl(4).Caption = "From"
Label2.Caption = "Type"
Label4.Caption = "Type of Maintenance "
'lbl(6).Caption = "To"
lbl(3).Caption = "To"
Label3.Caption = "Model"
RdNew.Caption = "New"
RdNew.RightToLeft = False
RdAccept.RightToLeft = False
RdAccept.Caption = "Accept Client"
Label5.Caption = "Reports Maintenance Car"
RdFinal.RightToLeft = False
RdFinal.Caption = "End"
RdAll2.RightToLeft = False
RdAll2.Caption = "All"

   '  With Me.Fg
   '     .TextMatrix(0, .ColIndex("Serial")) = "NO"
   '     .TextMatrix(0, .ColIndex("id")) = "Code"
   '     .TextMatrix(0, .ColIndex("RecordDate")) = "Date"
   '      .TextMatrix(0, .ColIndex("ClientName")) = "ClientName"
   '     .TextMatrix(0, .ColIndex("Telephone")) = "Telephone"
   '    .TextMatrix(0, .ColIndex("PlateNo")) = "PlateNo"
   ' End With
  '


  '
End Sub

Private Sub Form_Load()
    Dim GrdBack As ClsBackGroundPic
    Dim Dcombos As ClsDataCombos
    
        Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500

Me.DtStart.value = ""
Me.DtEnd.value = ""
Me.RDALL.value = True
Me.RdAll2.value = True
    Set Dcombos = New ClsDataCombos
   ' Dcombos.GetEmployees Me.DCEmp_Name
     Dcombos.GetClientName DcbClientname
     Dcombos.GetTblCarModels DcbCarModel
      Dcombos.GetTblMaintenanceWork Me.DCBMinten
     Dcombos.GetTblCarsDataGroup DcbCarType
    Set DCboSearch = New clsDCboSearch
    Set DCboSearch.Client = Me.DcbClientname
    'Dcombos.GetUsers Me.DCUser
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Clear").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture

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

StrSQL = "SELECT     TOP 100 PERCENT dbo.TblCardAuthorizationReform.ID, dbo.TblCardAuthorizationReform.RecordDate, dbo.TblCardAuthorizationReform.ClientName, "
StrSQL = StrSQL & "                       dbo.TblCardAuthorizationReform.Telephone, dbo.TblCardAuthorizationReform.PlateNo, dbo.TblCardAuthorizationReform.OrderStatus,"
StrSQL = StrSQL & "                       dbo.TblCardAuthorizationReformDetails.Type, dbo.TblCardAuthorizationReformDetails.[Value], dbo.TblCarModels.Model, dbo.TblCarModels.CarID,"
StrSQL = StrSQL & "                       dbo.TBLCarTypes.name, dbo.TBLCarTypes.namee, dbo.TblMaintenanceWork.name AS NameM, dbo.TblMaintenanceWork.namee AS Nameem,"
StrSQL = StrSQL & "                       dbo.TblCardAuthorizationReformDetails.Mainte, dbo.TblCardAuthorizationReform.CarTypeID, dbo.TblCardAuthorizationReform.CarModelID,"
StrSQL = StrSQL & "                       dbo.TblCardAuthorizationReform.Granty, dbo.TblCardAuthorizationReform.CarMeter, dbo.TblCardAuthorizationReform.PayFirst,"
StrSQL = StrSQL & "                       dbo.TblCardAuthorizationReform.AmountAccept, dbo.TblCardAuthorizationReformDetails.[count], dbo.TblCardAuthorizationReform.Complaint,"
StrSQL = StrSQL & "                       dbo.TblCardAuthorizationReform.Noteinitial, dbo.TblCardAuthorizationReformDetails.comp, dbo.TblCardAuthorizationReformDetails.bill, dbo.TblCarModels.ModelE,"
StrSQL = StrSQL & "                       dbo.TblCardAuthorizationReform.EndDate, dbo.TblCardAuthorizationReform.Accept, dbo.TblCardAuthorizationReform.YearFact,"
StrSQL = StrSQL & "                       dbo.TblCardAuthorizationReform.ColorID, dbo.TblCardAuthorizationReform.BranchID, dbo.TblCardAuthorizationReform.LongGranty,"
StrSQL = StrSQL & "                       dbo.TblCardAuthorizationReform.DateEndG, dbo.TblCardAuthorizationReform.DateStartG, dbo.TblCardAuthorizationReform.Month_Day,"
StrSQL = StrSQL & "                       dbo.TblCardAuthorizationReform.NotAccept , dbo.TblCardAuthorizationReform.Shaseh, dbo.TblCardAuthorizationReform.Mobile"
StrSQL = StrSQL & "  FROM         dbo.TblCardAuthorizationReform INNER JOIN"
StrSQL = StrSQL & "                       dbo.TBLCarTypes ON dbo.TblCardAuthorizationReform.CarTypeID = dbo.TBLCarTypes.id INNER JOIN"
StrSQL = StrSQL & "                       dbo.TblCarModels ON dbo.TblCardAuthorizationReform.CarModelID = dbo.TblCarModels.Id INNER JOIN"
StrSQL = StrSQL & "                       dbo.TblCardAuthorizationReformDetails ON dbo.TblCardAuthorizationReform.ID = dbo.TblCardAuthorizationReformDetails.ID INNER JOIN"
StrSQL = StrSQL & "                       dbo.TblMaintenanceWork ON dbo.TblCardAuthorizationReformDetails.Mainte = dbo.TblMaintenanceWork.Id"
StrSQL = StrSQL & " Where (dbo.TblCardAuthorizationReformDetails.Type = 0) And (1 = 1)"
    BolBegine = False
    StrWhere = ""
If Me.RDGRANTY.value = True Then
StrWhere = StrWhere & " AND dbo.TblCardAuthorizationReform.Granty = 0"
End If
If Me.RDWITHOUTGRANTY.value = True Then
StrWhere = StrWhere & " AND dbo.TblCardAuthorizationReform.Granty = 1"
End If
If Me.RDRETURNM.value = True Then
StrWhere = StrWhere & " AND dbo.TblCardAuthorizationReform.Granty = 2"
End If
If Me.RDALL.value = True Then
StrWhere = StrWhere & " AND dbo.TblCardAuthorizationReform.Granty >=0"
End If

If Me.RdNew.value = True Then
StrWhere = StrWhere & " AND dbo.TblCardAuthorizationReform.OrderStatus =0"
End If
If Me.RdAccept.value = True Then
StrWhere = StrWhere & " AND dbo.TblCardAuthorizationReform.OrderStatus =1"
End If
If Me.RdFinal.value = True Then
StrWhere = StrWhere & " AND dbo.TblCardAuthorizationReform.OrderStatus =2"
End If
If Me.RdNotAccept.value = True Then
StrWhere = StrWhere & " AND dbo.TblCardAuthorizationReform.OrderStatus =3"
End If
If Me.RdAll2.value = True Then
StrWhere = StrWhere & " AND dbo.TblCardAuthorizationReform.OrderStatus <>5"

End If
 '   If val(Me.TxtIDFrom.text) <> 0 Then
 '
 '           StrWhere = StrWhere & " AND dbo.TblCardAuthorizationReform.ID >=" & val(Me.TxtIDFrom.text) & ""
     
 '   End If
   

  '  If val(Me.TxtIDTO.text) <> 0 Then
  '
  '          StrWhere = StrWhere & " AND dbo.TblCardAuthorizationReform.ID <=" & val(Me.TxtIDTO.text) & ""
  '        End If
 
  If (TxtMobile.Text) <> "" Then
            StrWhere = StrWhere & " AND dbo.TblCardAuthorizationReform.Mobile like '%" & Me.TxtMobile.Text & "%'"
    End If
    
  If (TxtPlateNO.Text) <> "" Then
        
            StrWhere = StrWhere & " AND dbo.TblCardAuthorizationReform.PlateNo like '%" & Me.TxtPlateNO.Text & "%'"
        
    End If
   If Me.DcbClientname.Text <> "" Then
     
            StrWhere = " AND dbo.TblCardAuthorizationReform.ClientName = N'" & DcbClientname.Text & "'"
      
    End If
    If Me.DcbCarType.BoundText <> "" Then
     
            StrWhere = " AND dbo.TblCardAuthorizationReform.CarTypeID=" & Me.DcbCarType.BoundText & ""
      
    End If
    If Me.DCBMinten.BoundText <> "" Then
     
            StrWhere = " AND dbo.TblCardAuthorizationReformDetails.Mainte = " & Me.DCBMinten.BoundText & ""
      
    End If
  If Me.DcbCarModel.BoundText <> "" Then
     
            StrWhere = " AND dbo.TblCardAuthorizationReform.CarModelID=" & val(Me.DcbCarModel.BoundText) & ""
      
    End If

    If Not IsNull(Me.DtpDateFrom.value) Then
                   StrWhere = StrWhere & " AND dbo.TblCardAuthorizationReform.RecordDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
      End If

    If Not IsNull(Me.DtpDateTo.value) Then
            StrWhere = StrWhere & " AND  dbo.TblCardAuthorizationReform.RecordDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
     
    End If
If Not IsNull(Me.DtStart.value) Then
                   StrWhere = StrWhere & " AND dbo.TblCardAuthorizationReform.RecordDate >=" & SQLDate(Me.DtStart.value, True) & ""
      End If

    If Not IsNull(Me.DtStart.value) Then
             
            StrWhere = StrWhere & " AND  dbo.TblCardAuthorizationReform.RecordDate <=" & SQLDate(Me.DtStart.value, True) & ""
     
    End If
    If Not IsNull(Me.DtEnd.value) Then
                   StrWhere = StrWhere & " AND dbo.TblCardAuthorizationReform.EndDate >=" & SQLDate(Me.DtEnd.value, True) & ""
      End If

    If Not IsNull(Me.DtEnd.value) Then
            StrWhere = StrWhere & " AND  dbo.TblCardAuthorizationReform.EndDate <=" & SQLDate(Me.DtEnd.value, True) & ""
     
    End If
    '-----------------------------------

    StrSQL = StrSQL & StrWhere
 
   StrSQL = StrSQL & " Order By dbo.TblCardAuthorizationReform.ID"
  print_report StrSQL
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

        '    rs.MoveFirst
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
        If Me.XPChkSearchTypeClient1.value = True Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rep1byclient.rpt"
            Else
            If Me.XPChkSearchTypeCar.value = True Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rep1byCar.rpt"
            Else
            If Me.XPChkSearchTypeModel.value = True Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rep1byModel.rpt"
            Else
             If Me.XPChkSearchTypePlate.value = True Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rep1byPlate.rpt"
            Else
             If Me.XPChkSearchTypeMaint.value = True Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rep1byMaintain.rpt"
            Else
             StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rep1.rpt"
            End If
            End If
            End If
            
            
            End If
             End If
        Else
               If Me.XPChkSearchTypeClient1.value = True Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rep1byclient.rpt"
            Else
            If Me.XPChkSearchTypeCar.value = True Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rep1byCar.rpt"
            Else
            If Me.XPChkSearchTypeModel.value = True Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rep1byModel.rpt"
            Else
             If Me.XPChkSearchTypePlate.value = True Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rep1byPlate.rpt"
            Else
             If Me.XPChkSearchTypeMaint.value = True Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rep1byMaintain.rpt"
            Else
             StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rep1.rpt"
            End If
            End If
            End If
            
            
            End If
             End If
           
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
      '  xReport.ParameterFields(4).AddCurrentValue WriteNo(Format(val(TxtAdvanceValue.text), "0.00"), 0, True, ".")
       ' xReport.ParameterFields(6).AddCurrentValue val(lbl(23).Caption)
       '  xReport.ParameterFields(7).AddCurrentValue DBIssueDate.value
'    xReport.ParameterFields(8).AddCurrentValue IIf(IsNumeric(fg.TextMatrix(Me.fg.FixedRows, fg.ColIndex("PartValue"))), val(fg.TextMatrix(Me.fg.FixedRows, fg.ColIndex("PartValue"))), 0)
' xReport.ParameterFields(9).AddCurrentValue val(lbl(22).Caption)
 ' xReport.ParameterFields(10).AddCurrentValue val(TxtDiscount.text)
  ' xReport.ParameterFields(11).AddCurrentValue txtDiscountDES.text
  Dim Total As String
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


 
