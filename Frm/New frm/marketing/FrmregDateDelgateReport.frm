VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SuiteCtrls.ocx"
Begin VB.Form FrmRegDateDelgateREport 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " ﬁ«—Ì— „Ê«⁄Ìœ «·„‰«œÌ»"
   ClientHeight    =   4530
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10080
   Icon            =   "FrmregDateDelgateReport.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   4530
   ScaleWidth      =   10080
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
   Begin VB.Frame frame 
      BackColor       =   &H00E2E9E9&
      Caption         =   "›—“ »Õ”» Õ«·… «·“Ì«—…"
      Height          =   615
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   43
      Top             =   480
      Width           =   10035
      Begin XtremeSuiteControls.RadioButton RdStus 
         Height          =   255
         Index           =   0
         Left            =   7440
         TabIndex        =   44
         Top             =   240
         Width           =   1815
         _Version        =   786432
         _ExtentX        =   3201
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   " „ «·“Ì«—…"
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton RdStus 
         Height          =   375
         Index           =   1
         Left            =   5160
         TabIndex        =   45
         Top             =   120
         Width           =   1815
         _Version        =   786432
         _ExtentX        =   3201
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   " „ «· ⁄«ﬁœ"
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton RdStus 
         Height          =   375
         Index           =   2
         Left            =   3000
         TabIndex        =   46
         Top             =   120
         Width           =   1815
         _Version        =   786432
         _ExtentX        =   3201
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   " „ «·€«¡ «·“Ì«—…"
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton RdStus 
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   47
         Top             =   120
         Width           =   1815
         _Version        =   786432
         _ExtentX        =   3201
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "«·ﬂ·"
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
   End
   Begin XtremeSuiteControls.RadioButton XPChkSearchTypeClient1 
      Height          =   375
      Left            =   8160
      TabIndex        =   9
      Top             =   3480
      Width           =   1815
      _Version        =   786432
      _ExtentX        =   3201
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "›—“ »Õ”» «·„‰œÊ»"
      UseVisualStyle  =   -1  'True
      TextAlignment   =   1
      RightToLeft     =   -1  'True
   End
   Begin VB.ComboBox DcbOrderStatus 
      Height          =   315
      Left            =   12960
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   1200
      Width           =   1935
   End
   Begin VB.ComboBox ComGranty 
      Height          =   315
      Left            =   13440
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   120
      Width           =   2175
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Height          =   2565
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   960
      Width           =   10035
      Begin VB.TextBox DcbJob 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   42
         Top             =   960
         Width           =   3855
      End
      Begin VB.TextBox TxtEmail 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   2040
         Width           =   8895
      End
      Begin VB.TextBox TxtMobile 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   1680
         Width           =   3855
      End
      Begin VB.TextBox TxtTel 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   1320
         Width           =   3855
      End
      Begin VB.TextBox TxtAdmin 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   600
         Width           =   3855
      End
      Begin VB.TextBox TxtCustomer 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   240
         Width           =   3855
      End
      Begin MSDataListLib.DataCombo DcboEmpName 
         Height          =   315
         Left            =   5040
         TabIndex        =   15
         Top             =   240
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcbTypeVisit1 
         Height          =   315
         Left            =   5040
         TabIndex        =   16
         Top             =   600
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcbJob1 
         Height          =   315
         Left            =   120
         TabIndex        =   17
         Top             =   1680
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSComCtl2.DTPicker TimeFrom 
         Height          =   315
         Left            =   7440
         TabIndex        =   21
         Top             =   4560
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   160956418
         CurrentDate     =   38784
      End
      Begin MSComCtl2.DTPicker FromDeparDate 
         Height          =   315
         Left            =   7440
         TabIndex        =   24
         Top             =   1680
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   160956417
         CurrentDate     =   38784
      End
      Begin MSComCtl2.DTPicker TimeTo 
         Height          =   315
         Left            =   5040
         TabIndex        =   29
         Top             =   3480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   160956418
         CurrentDate     =   38784
      End
      Begin MSComCtl2.DTPicker ToDeparDate 
         Height          =   315
         Left            =   5040
         TabIndex        =   30
         Top             =   1680
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   160956417
         CurrentDate     =   38784
      End
      Begin MSDataListLib.DataCombo DcbSpecialAs 
         Height          =   315
         Left            =   5040
         TabIndex        =   39
         Top             =   960
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcbTypeVisit2 
         Height          =   315
         Left            =   5040
         TabIndex        =   41
         Top             =   1320
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·ŒÿÊ… «· «·Ì…"
         Height          =   195
         Index           =   12
         Left            =   8805
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   1320
         Width           =   1140
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÃÊ«·"
         Height          =   195
         Index           =   10
         Left            =   3645
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   1680
         Width           =   1140
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " ·›Ê‰"
         Height          =   195
         Index           =   8
         Left            =   3645
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   1320
         Width           =   1140
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«Ì„Ì·"
         Height          =   195
         Index           =   6
         Left            =   8880
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   2010
         Width           =   1065
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "„‰  «—ÌŒ "
         Height          =   315
         Index           =   4
         Left            =   6480
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   1680
         Width           =   825
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "«·Ï «·”«⁄…"
         Height          =   315
         Index           =   3
         Left            =   10950
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   2040
         Width           =   915
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·„”ƒ·"
         Height          =   195
         Index           =   2
         Left            =   3645
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   600
         Width           =   1140
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«· ﬁÌÌ„ «·Œ«’"
         Height          =   195
         Index           =   11
         Left            =   8805
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   960
         Width           =   1140
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "„‰ «·”«⁄…"
         Height          =   315
         Index           =   9
         Left            =   13590
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   2040
         Width           =   915
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         BackStyle       =   0  'Transparent
         Caption         =   "„‰  «—ÌŒ "
         Height          =   315
         Index           =   45
         Left            =   9120
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   1680
         Width           =   825
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Height          =   195
         Left            =   4770
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   360
         Width           =   45
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰Ê⁄ «·“Ì«—…"
         Height          =   195
         Index           =   5
         Left            =   8805
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   600
         Width           =   1140
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·ÊŸÌ›…"
         Height          =   195
         Index           =   0
         Left            =   3720
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   930
         Width           =   1065
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·⁄„Ì·"
         Height          =   195
         Index           =   7
         Left            =   3645
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   240
         Width           =   1140
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·„‰œÊ»"
         Height          =   195
         Left            =   9435
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   240
         Width           =   510
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
      TabIndex        =   1
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
   Begin XtremeSuiteControls.RadioButton XPChkSearchTypeCar 
      Height          =   375
      Left            =   3240
      TabIndex        =   10
      Top             =   3480
      Width           =   1935
      _Version        =   786432
      _ExtentX        =   3413
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   " ›—“ »Õ”» «· ﬁÌÌ„ «·Œ«’"
      UseVisualStyle  =   -1  'True
      TextAlignment   =   1
      RightToLeft     =   -1  'True
   End
   Begin XtremeSuiteControls.RadioButton XPChkSearchTypeModel 
      Height          =   375
      Left            =   -120
      TabIndex        =   11
      Top             =   3480
      Width           =   2775
      _Version        =   786432
      _ExtentX        =   4895
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "›—“ »Õ”» ÊŸÌ›… «·‘Œ’ «·„”ƒ«·"
      UseVisualStyle  =   -1  'True
      TextAlignment   =   1
      RightToLeft     =   -1  'True
   End
   Begin XtremeSuiteControls.RadioButton XPChkSearchTypePlate 
      Height          =   375
      Left            =   5520
      TabIndex        =   12
      Top             =   3480
      Width           =   1935
      _Version        =   786432
      _ExtentX        =   3413
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "›—“ »Õ”» ‰Ê⁄ «·“Ì«—…"
      UseVisualStyle  =   -1  'True
      TextAlignment   =   1
      RightToLeft     =   -1  'True
   End
   Begin ImpulseButton.ISButton Cmd 
      Height          =   495
      Index           =   0
      Left            =   2760
      TabIndex        =   25
      Top             =   3960
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
      Caption         =   " ﬁ«—Ì— „Ê«⁄Ìœ «·„‰«œÌ»"
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
      Left            =   3735
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   120
      Width           =   2895
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
Attribute VB_Name = "FrmRegDateDelgateREport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As ADODB.Recordset
Dim DCboSearch As clsDCboSearch



Private Sub Cmd_Click(index As Integer)

    Select Case index

        Case 0
 

 GetData
            
        Case 1
              clear_all Me
FromDeparDate.value = ""
ToDeparDate.value = ""
Me.TimeFrom.value = ""
Me.Timeto.value = ""
Me.RdStus(0).value = False
Me.RdStus(1).value = False
Me.RdStus(2).value = False
Me.RdStus(3).value = False

XPChkSearchTypeModel.value = False
XPChkSearchTypeCar.value = False
XPChkSearchTypeClient1.value = False
XPChkSearchTypePlate.value = False

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




Private Sub Form_Activate()
   PutFormOnTop Me.hWnd
End Sub
Private Sub ChangeLang()

 
Me.RdStus(0).RightToLeft = False
Me.RdStus(1).RightToLeft = False
Me.RdStus(2).RightToLeft = False
Me.RdStus(3).RightToLeft = False
Me.RdStus(0).Caption = "Visited"
Me.RdStus(1).Caption = "Been Contracted"
Me.RdStus(2).Caption = "Cancel Visit"
Me.RdStus(3).Caption = "All"
frame.Caption = "Sort By"
 lbl(0).Caption = "Job"
 Cmd(1).Caption = "Delete"
    Cmd(0).Caption = "View Report"
   Cmd(2).Caption = "Exit"
  Me.Caption = "Reports Dates Almnadeb "
Label5.Caption = Me.Caption
Label1.Caption = "Delegate"
lbl(45).Caption = "From"
lbl(10).Caption = "Mobile"
lbl(7).Caption = "Customer"
lbl(8).Caption = "Telephone"
lbl(5).Caption = "Type visit "
lbl(11).Caption = "Rating"
lbl(4).Caption = "To"
lbl(2).Caption = "Admin"
lbl(12).Caption = "Next Step"
lbl(6).Caption = "Email"
XPChkSearchTypeCar.RightToLeft = False
Me.XPChkSearchTypeCar.Caption = "By Rating"
Me.XPChkSearchTypeClient1.RightToLeft = False
Me.XPChkSearchTypeClient1.Caption = "By Delegate"
Me.XPChkSearchTypePlate.RightToLeft = False
Me.XPChkSearchTypePlate.Caption = "By TypeVisit"
Me.XPChkSearchTypeModel.RightToLeft = False
Me.XPChkSearchTypeModel.Caption = "By Job"
'XPChkSearchType.RightToLeft = False
'XPChkSearchType.Caption = "By Type"
'lbl(3).Caption = "To Date"
'lbl(4).Caption = "From Date"
'lbl(45).Caption = "From Date expected travel'"
'lbl(9).Caption = "To Date expected travel'"
'lbl(46).Caption = "From Date  return expected"
'lbl(8).Caption = "To Date  return expected"
End Sub

Private Sub Form_Load()
    'Dim GrdBack As ClsBackGroundPic
    Dim Dcombos As ClsDataCombos
   FromDeparDate.value = Date
   ToDeparDate.value = Date
    Me.left = (mdifrmmain.Width - Me.Width) / 2
    Me.top = (mdifrmmain.Height - Me.Height) / 2 - 500
    Set Dcombos = New ClsDataCombos
    Set Cmd(0).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Search").Picture
    Set Cmd(1).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Clear").Picture
    Set Cmd(2).ButtonImage = mdifrmmain.ImgLstTree.ListImages("Exit").Picture

    Dcombos.GetTypeVisit Me.DcbTypeVisit1
    Dcombos.GetTypeVisit Me.DcbTypeVisit2
    Dcombos.GetDelegate Me.DcboEmpName
    Dcombos.GetSpeciaAsement Me.DcbSpecialAs
 If SystemOptions.UserInterface = EnglishInterface Then
        SetInterface Me
        ChangeLang
    End If
FromDeparDate.value = ""
ToDeparDate.value = ""
Me.TimeFrom.value = ""
Me.Timeto.value = ""
End Sub
Private Sub Form_Unload(Cancel As Integer)

    FormPostion Me, SavePostion
    Set DCboSearch = Nothing
End Sub

Public Sub GetData()
    Dim StrSQL As String
    Dim StrSQL1 As String
    Dim strsql2 As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
Dim chek As Integer
chek = 0
StrSQL1 = " SELECT DISTINCT"
StrSQL1 = StrSQL1 & "    dbo.TblRegDateDelgate.Id , dbo.TblBranchesData.branch_namee, dbo.TblEmployee.Emp_Name, dbo.TblEmployee.Emp_Name1, "
StrSQL1 = StrSQL1 & "                      dbo.TblEmployee.Emp_Name2, dbo.TblEmployee.Emp_Name3, dbo.TblEmployee.Emp_Name4, dbo.TblEmployee.Emp_Namee, dbo.TblEmployee.Emp_Namee1,"
StrSQL1 = StrSQL1 & "                      dbo.TblEmployee.Emp_Namee2, dbo.TblEmployee.Emp_Namee3, dbo.TblEmployee.Emp_Namee4, dbo.TblEmployee.Fullcode, dbo.TblCustemers.CusName,"
StrSQL1 = StrSQL1 & "                      dbo.TblCustemers.CusNamee, dbo.TblCustemers.Fullcode AS CusFullcode, dbo.TblTypeVisit.name AS Visitname, dbo.TblTypeVisit.namee AS VisitnameE,"
StrSQL1 = StrSQL1 & "                      TblTypeVisit_1.name AS Visitname2, TblTypeVisit_1.namee AS VisitnameE2, dbo.TblSpeciaAsement.name AS Specname,"
StrSQL1 = StrSQL1 & "                      dbo.TblSpeciaAsement.namee AS SpecnameE, dbo.TblRegDateDelgate.RecordDate, dbo.TblRegDateDelgate.BranchID, dbo.TblRegDateDelgate.DelgID,"
StrSQL1 = StrSQL1 & "                      dbo.TblRegDateDelgate.CustomerName, dbo.TblRegDateDelgate.Remark, dbo.TblRegDateDelgate.VisitID, dbo.TblRegDateDelgate.SpAsID,"
StrSQL1 = StrSQL1 & "                      dbo.TblRegDateDelgate.VisitID2, dbo.TblRegDateDelgate.Accept, dbo.TblRegDateDelgate.VisitDate, dbo.TblRegDateDelgate.Remark2,"
StrSQL1 = StrSQL1 & "                      dbo.TblRegDateDelgate.PersonConc, dbo.TblRegDateDelgate.Tel, dbo.TblRegDateDelgate.Mobile, dbo.TblRegDateDelgate.Email, dbo.TblRegDateDelgate.JobID,"
StrSQL1 = StrSQL1 & "                      dbo.TblRegDateDelgate.LongTime, dbo.TblRegDateDelgate.VisitDate1, dbo.TblRegDateDelgate.NotAcept, dbo.TblRegDateDelgate.BillNo,"
StrSQL1 = StrSQL1 & "                      dbo.TblRegDateDelgate.CustomerID, dbo.TblRegDateDelgate.Entry, dbo.TblRegDateDelgate.Map, dbo.TblRegDateDelgate.Adress,"
StrSQL1 = StrSQL1 & "                      dbo.TblRegDateDelgate.FromTime1, dbo.TblRegTimeDelgate.name AS NameFromTime1, dbo.TblRegDateDelgate.FromTime2,"
StrSQL1 = StrSQL1 & "                      TblRegTimeDelgate_1.name AS NameFromTime2, dbo.TblRegDateDelgate.ToTime1, TblRegTimeDelgate_2.name AS NameToTime1, dbo.TblRegDateDelgate.ToTime2,"
StrSQL1 = StrSQL1 & "                      TblRegTimeDelgate_3.name AS NameToTime2,  dbo.TblRegDateDelgateDails.remark AS remarkDet, dbo.TblRegDateDelgateDails.quantity,"
StrSQL1 = StrSQL1 & "                      dbo.TblRegDateDelgateDails.Type, dbo.TblRegDateDelgateDails.VisitDate AS VisitDateDet, dbo.TblRegDateDelgateDails.EmpID,"
StrSQL1 = StrSQL1 & "                      TblEmployee_1.Emp_Name AS Emp_NameDet, TblEmployee_1.Emp_Name1 AS Emp_Name1Det, TblEmployee_1.Emp_Name2 AS Emp_Name2Det,"
StrSQL1 = StrSQL1 & "                      TblEmployee_1.Emp_Name3 AS Emp_Name3Det, TblEmployee_1.Emp_Name4 AS Emp_Name4Det, TblEmployee_1.Emp_Namee AS Emp_NameeDet,"
StrSQL1 = StrSQL1 & "                      TblEmployee_1.Emp_Namee1 AS Emp_Namee1Det, TblEmployee_1.Emp_Namee2 AS Emp_Namee2Det, TblEmployee_1.Emp_Namee3 AS Emp_Namee3Det,"
StrSQL1 = StrSQL1 & "                      TblEmployee_1.Emp_Namee4 AS Emp_Namee4Det, TblEmployee_1.Fullcode AS FullcodeDet, dbo.TblCompo.name, dbo.TblCompo.namee,"
StrSQL1 = StrSQL1 & "                      dbo.TblEmployee.JobTypeID, dbo.TblEmpJobsTypes.JobTypeName, dbo.TblEmpJobsTypes.JobTypeNamee, TblEmployee_1.JobTypeID AS JobTypeIDDet,"
StrSQL1 = StrSQL1 & "                      TblEmpJobsTypes_1.JobTypeName AS JobTypeNameDet, TblEmpJobsTypes_1.JobTypeNamee AS JobTypeNameeDet"
StrSQL1 = StrSQL1 & " FROM         dbo.TblEmpJobsTypes RIGHT OUTER JOIN"
StrSQL1 = StrSQL1 & "                      dbo.TblEmployee RIGHT OUTER JOIN"
StrSQL1 = StrSQL1 & "                      dbo.TblEmpJobsTypes TblEmpJobsTypes_1 RIGHT OUTER JOIN"
StrSQL1 = StrSQL1 & "                      dbo.TblEmployee TblEmployee_1 ON TblEmpJobsTypes_1.JobTypeID = TblEmployee_1.JobTypeID RIGHT OUTER JOIN"
StrSQL1 = StrSQL1 & "                      dbo.TblCompo RIGHT OUTER JOIN"
StrSQL1 = StrSQL1 & "                      dbo.TblRegDateDelgateDails ON dbo.TblCompo.Id = dbo.TblRegDateDelgateDails.EmpID ON"
StrSQL1 = StrSQL1 & "                      TblEmployee_1.Emp_ID = dbo.TblRegDateDelgateDails.EmpID RIGHT OUTER JOIN"
StrSQL1 = StrSQL1 & "                      dbo.TblRegDateDelgate ON dbo.TblRegDateDelgateDails.DelgID = dbo.TblRegDateDelgate.Id ON dbo.TblEmployee.Emp_ID = dbo.TblRegDateDelgate.DelgID ON"
StrSQL1 = StrSQL1 & "                      dbo.TblEmpJobsTypes.JobTypeID = dbo.TblEmployee.JobTypeID LEFT OUTER JOIN"
StrSQL1 = StrSQL1 & "                      dbo.TblRegTimeDelgate TblRegTimeDelgate_3 ON dbo.TblRegDateDelgate.ToTime2 = TblRegTimeDelgate_3.Id LEFT OUTER JOIN"
StrSQL1 = StrSQL1 & "                      dbo.TblRegTimeDelgate TblRegTimeDelgate_2 ON dbo.TblRegDateDelgate.ToTime1 = TblRegTimeDelgate_2.Id LEFT OUTER JOIN"
StrSQL1 = StrSQL1 & "                      dbo.TblRegTimeDelgate TblRegTimeDelgate_1 ON dbo.TblRegDateDelgate.FromTime2 = TblRegTimeDelgate_1.Id LEFT OUTER JOIN"
StrSQL1 = StrSQL1 & "                      dbo.TblRegTimeDelgate ON dbo.TblRegDateDelgate.FromTime1 = dbo.TblRegTimeDelgate.Id LEFT OUTER JOIN"
StrSQL1 = StrSQL1 & "                      dbo.TblSpeciaAsement ON dbo.TblRegDateDelgate.SpAsID = dbo.TblSpeciaAsement.Id LEFT OUTER JOIN"
StrSQL1 = StrSQL1 & "                      dbo.TblTypeVisit TblTypeVisit_1 ON dbo.TblRegDateDelgate.VisitID2 = TblTypeVisit_1.Id LEFT OUTER JOIN"
StrSQL1 = StrSQL1 & "                      dbo.TblTypeVisit ON dbo.TblRegDateDelgate.VisitID = dbo.TblTypeVisit.Id LEFT OUTER JOIN"
StrSQL1 = StrSQL1 & "                      dbo.TblCustemers ON dbo.TblRegDateDelgate.CustomerID = dbo.TblCustemers.CusID LEFT OUTER JOIN"
StrSQL1 = StrSQL1 & "                      dbo.TblBranchesData ON dbo.TblRegDateDelgate.BranchID = dbo.TblBranchesData.branch_id"
StrSQL1 = StrSQL1 & " Where (1 = 1)"
                      
    BolBegine = False
    StrWhere = ""

If Me.RdStus(0).value = True Then
    ' chek = 1
            StrWhere = StrWhere & " AND dbo.TblRegDateDelgate.Accept  =1 "
      
    End If
If Me.RdStus(1).value = True Then
    ' chek = 1
            StrWhere = StrWhere & " AND dbo.TblRegDateDelgate.Accept  =2 "
      
    End If
    If Me.RdStus(2).value = True Then
    ' chek = 1
            StrWhere = StrWhere & " AND dbo.TblRegDateDelgate.Accept  =3 "
      
    End If
   

If val(DcbSpecialAs.BoundText) <> 0 Then
     chek = 1
            StrWhere = StrWhere & " AND dbo.TblRegDateDelgate.SpAsID   =" & val(Me.DcbSpecialAs.BoundText) & ""
      
    End If
 
If val(DcbTypeVisit1.BoundText) <> 0 Then
     chek = 1
            StrWhere = StrWhere & " AND dbo.TblRegDateDelgate.VisitID  =" & val(Me.DcbTypeVisit1.BoundText) & ""
      
    End If
 If val(DcbTypeVisit2.BoundText) <> 0 Then
     chek = 1
            StrWhere = StrWhere & " AND dbo.TblRegDateDelgate.VisitID2 =" & val(Me.DcbTypeVisit2.BoundText) & ""
      
    End If
       If Me.TxtCustomer.text <> "" Then
     chek = 1
            StrWhere = StrWhere & " AND  dbo.TblCustemers.CusName like '%" & Me.TxtCustomer.text & "%'"
      
    End If
       If Me.TxtAdmin.text <> "" Then
     chek = 1
            StrWhere = StrWhere & " AND  dbo.TblRegDateDelgate.PersonConc like '%" & Me.TxtAdmin.text & "%'"
      
    End If
       If Me.TxtEmail.text <> "" Then
     chek = 1
            StrWhere = StrWhere & " AND  dbo.TblRegDateDelgate.Email like '%" & Me.TxtEmail.text & "%'"
      
    End If
       If Me.txtTel.text <> "" Then
     chek = 1
            StrWhere = StrWhere & " AND  dbo.TblRegDateDelgate.Tel like '%" & Me.txtTel.text & "%'"
      
    End If
     If Me.TxtMobile.text <> "" Then
     chek = 1
            StrWhere = StrWhere & " AND  dbo.TblRegDateDelgate.Mobile like '%" & Me.TxtMobile.text & "%'"
      
    End If
   If val(Me.DcboEmpName.BoundText) <> 0 Then
     chek = 1
            StrWhere = StrWhere & " AND TblRegDateDelgate.DelgID =" & val(Me.DcboEmpName.BoundText) & ""
      
    End If

  
  If Me.DcbJob.text <> "" Then
     chek = 1
            StrWhere = StrWhere & " AND dbo.TblRegDateDelgate.JobID like N'%" & Me.DcbJob.text & "%'"
      
    End If

    If Not IsNull(Me.FromDeparDate.value) Then
    chek = 1
                   StrWhere = StrWhere & " AND dbo.TblRegDateDelgate.RecordDate>=" & SQLDate(Me.FromDeparDate.value, True) & ""
      End If
   If Not IsNull(Me.ToDeparDate.value) Then
   chek = 1
           StrWhere = StrWhere & " AND  dbo.TblRegDateDelgate.RecordDate <=" & SQLDate(Me.ToDeparDate.value, True) & ""

   End If

StrSQL = StrSQL1 & StrWhere
   StrSQL = StrSQL & " Order By dbo.TblRegDateDelgate.ID"
  
  
    Set rs = New ADODB.Recordset
    rs.Open StrSQL, Cn, adOpenStatic, adLockReadOnly, adCmdText

    If rs.BOF Or rs.EOF Then
        If SystemOptions.UserInterface = ArabicInterface Then
           ' Me.lbl(10).Caption = "‰ ÌÃ… «·»ÕÀ=’›—"
        ElseIf SystemOptions.UserInterface = EnglishInterface Then
          '  Me.lbl(10).Caption = "Search Results=0"
        End If

        Msg = "·« ÊÃœ »Ì«‰«  ··⁄—÷  Ê«›ﬁ ‘—Êÿ «· ﬁ—Ì—"
        MsgBox Msg, vbExclamation + vbMsgBoxRight + vbMsgBoxRtlReading, App.Title
        Exit Sub
    Else
    ' Me.DTPicker1.value = Format(rs("DateAcutExite").value, "yyyy/M/d")
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
        
         'StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepByRegDateDelegateDDD.rpt"
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepByRegDateDelegateDel.rpt"
            Else
            If Me.XPChkSearchTypeCar.value = True Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepByRegDateDelegateSpecial.rpt"
            Else
            If Me.XPChkSearchTypeModel.value = True Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepByRegDateDelegateJob.rpt"
            Else
             If Me.XPChkSearchTypePlate.value = True Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepByRegDateDelegateTypeVisit.rpt"
            Else
          
             StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepByRegDateDelegateAll.rpt"
            
            
            End If
             End If
            
            End If
             End If
        Else
               If Me.XPChkSearchTypeClient1.value = True Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepByRegDateDelegateDel.rpt"
            Else
            If Me.XPChkSearchTypeCar.value = True Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepByRegDateDelegateSpecial.rpt"
            Else
            If Me.XPChkSearchTypeModel.value = True Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepByRegDateDelegateJob.rpt"
            Else
             If Me.XPChkSearchTypePlate.value = True Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepByRegDateDelegateTypeVisit.rpt"
            Else
           
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "RepByRegDateDelegateAll.rpt"
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
  If Not IsNull(FromDeparDate.value) Then
        xReport.ParameterFields(4).AddCurrentValue FromDeparDate.value
     End If
       If Not IsNull(ToDeparDate.value) Then
        xReport.ParameterFields(5).AddCurrentValue ToDeparDate.value
     End If
       ' xReport.ParameterFields(6).AddCurrentValue val(lbl(23).Caption)
       '  xReport.ParameterFields(7).AddCurrentValue DBIssueDate.value
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
    xReport.ApplicationName = App.Title
    xReport.ReportAuthor = App.Title
    Set CViewer = New ClsReportViewer
    CViewer.FireReport xReport, WindowTarget, "", , , , StrFileName

    RsData.Close
    Set RsData = Nothing
    Screen.MousePointer = vbDefault


 
  
 
End Function















