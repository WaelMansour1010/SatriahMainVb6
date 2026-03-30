VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{FE5DCFAD-BC1D-11D2-94CF-004005455FAA}#1.4#0"; "ImpulseButton.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "SuiteCtrls.ocx"
Begin VB.Form FrmCarReportsRequerNo 
   BackColor       =   &H00E2E9E9&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   7575
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12120
   Icon            =   "FrmCarReportsRequerNO.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   7575
   ScaleWidth      =   12120
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame4 
      BackColor       =   &H00E2E9E9&
      Caption         =   " ﬁ—Ì— «ÃÊ— «·Ìœ"
      Height          =   645
      Left            =   3000
      RightToLeft     =   -1  'True
      TabIndex        =   77
      Top             =   1770
      Width           =   8955
      Begin XtremeSuiteControls.RadioButton optHandByDepDet 
         Height          =   255
         Left            =   7080
         TabIndex        =   78
         Top             =   240
         Width           =   1575
         _Version        =   786432
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "»«·«ﬁ”«„  Õ·Ì·Ì"
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton optHandByDepTotals 
         Height          =   255
         Left            =   4800
         TabIndex        =   79
         Top             =   240
         Width           =   1575
         _Version        =   786432
         _ExtentX        =   2778
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "»«·«ﬁ”«„ «Ã„«·Ì« "
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E2E9E9&
      Height          =   735
      Left            =   3720
      TabIndex        =   48
      Top             =   6900
      Width           =   7935
      Begin VB.CommandButton menue 
         Height          =   555
         Index           =   5
         Left            =   1440
         Picture         =   "FrmCarReportsRequerNO.frx":038A
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton menue 
         Height          =   555
         Index           =   4
         Left            =   2760
         Picture         =   "FrmCarReportsRequerNO.frx":07D5
         Style           =   1  'Graphical
         TabIndex        =   59
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton menue 
         Height          =   555
         Index           =   3
         Left            =   3480
         Picture         =   "FrmCarReportsRequerNO.frx":0D2D
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton menue 
         Height          =   555
         Index           =   2
         Left            =   4920
         Picture         =   "FrmCarReportsRequerNO.frx":11E6
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton menue 
         Height          =   555
         Index           =   1
         Left            =   3240
         Picture         =   "FrmCarReportsRequerNO.frx":16B6
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   840
         Width           =   735
      End
      Begin VB.CommandButton menue 
         DownPicture     =   "FrmCarReportsRequerNO.frx":1B57
         Height          =   555
         Index           =   0
         Left            =   7080
         Picture         =   "FrmCarReportsRequerNO.frx":8E89
         Style           =   1  'Graphical
         TabIndex        =   55
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton menue 
         DownPicture     =   "FrmCarReportsRequerNO.frx":9430
         Height          =   555
         Index           =   6
         Left            =   5640
         Picture         =   "FrmCarReportsRequerNO.frx":10762
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton menue 
         DownPicture     =   "FrmCarReportsRequerNO.frx":10C03
         Height          =   555
         Index           =   7
         Left            =   4200
         Picture         =   "FrmCarReportsRequerNO.frx":17F35
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton menue 
         Height          =   555
         Index           =   8
         Left            =   2040
         Picture         =   "FrmCarReportsRequerNO.frx":187C5
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton menue 
         DownPicture     =   "FrmCarReportsRequerNO.frx":18CAA
         Height          =   555
         Index           =   9
         Left            =   720
         Picture         =   "FrmCarReportsRequerNO.frx":1FFDC
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton menue 
         DownPicture     =   "FrmCarReportsRequerNO.frx":204FC
         Height          =   555
         Index           =   10
         Left            =   6360
         Picture         =   "FrmCarReportsRequerNO.frx":20AE3
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   120
         Width           =   735
      End
      Begin VB.CommandButton menue 
         DownPicture     =   "FrmCarReportsRequerNO.frx":210CA
         Height          =   555
         Index           =   11
         Left            =   0
         Picture         =   "FrmCarReportsRequerNO.frx":283FC
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E2E9E9&
      Caption         =   " ﬁ—Ì— »Õ”»"
      Height          =   645
      Left            =   3000
      RightToLeft     =   -1  'True
      TabIndex        =   30
      Top             =   1080
      Width           =   8955
      Begin XtremeSuiteControls.RadioButton RdNew 
         Height          =   255
         Left            =   8040
         TabIndex        =   31
         Top             =   240
         Width           =   855
         _Version        =   786432
         _ExtentX        =   1508
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "ÃœÌœ"
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton RdFinal 
         Height          =   255
         Left            =   4920
         TabIndex        =   32
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
         Left            =   6480
         TabIndex        =   33
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
         Left            =   -480
         TabIndex        =   34
         Top             =   240
         Width           =   1215
         _Version        =   786432
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "«·ﬂ·"
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton RdnotAccept 
         Height          =   255
         Left            =   2160
         TabIndex        =   41
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
      Begin XtremeSuiteControls.RadioButton Rdunderwait 
         Height          =   255
         Left            =   3720
         TabIndex        =   43
         Top             =   240
         Width           =   1215
         _Version        =   786432
         _ExtentX        =   2143
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   " Õ  «·«‰ Ÿ«—"
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton rdbill 
         Height          =   255
         Left            =   720
         TabIndex        =   44
         Top             =   240
         Width           =   1455
         _Version        =   786432
         _ExtentX        =   2566
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   " „ «’œ«— ›« Ê—…"
         UseVisualStyle  =   -1  'True
         TextAlignment   =   1
         RightToLeft     =   -1  'True
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2E9E9&
      Caption         =   " ﬁ—Ì— »Õ”»"
      Height          =   645
      Left            =   3000
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   480
      Width           =   8955
      Begin XtremeSuiteControls.RadioButton RDGRANTY 
         Height          =   255
         Left            =   7080
         TabIndex        =   26
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
         Left            =   2160
         TabIndex        =   27
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
         Left            =   4800
         TabIndex        =   28
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
         TabIndex        =   29
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
      Left            =   12240
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   1200
      Width           =   1935
   End
   Begin VB.ComboBox ComGranty 
      Height          =   315
      Left            =   12360
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   120
      Width           =   2175
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00E2E9E9&
      Caption         =   " ﬁ—Ì— »Õ”»"
      Height          =   3285
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   2550
      Width           =   11955
      Begin VB.TextBox txtOrderNo 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   7860
         RightToLeft     =   -1  'True
         TabIndex        =   83
         Top             =   2760
         Width           =   2775
      End
      Begin VB.TextBox TxtMobile 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   75
         Top             =   270
         Width           =   2775
      End
      Begin VB.TextBox TxtSuper 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4080
         RightToLeft     =   -1  'True
         TabIndex        =   74
         Top             =   1320
         Width           =   6495
      End
      Begin VB.TextBox txtrEQnO 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   7800
         RightToLeft     =   -1  'True
         TabIndex        =   73
         Top             =   2040
         Width           =   2775
      End
      Begin VB.TextBox TxtFiter 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   7800
         RightToLeft     =   -1  'True
         TabIndex        =   65
         Top             =   1680
         Width           =   2775
      End
      Begin VB.TextBox TxtPlateNO 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4080
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   1680
         Width           =   2655
      End
      Begin MSDataListLib.DataCombo DcbCarType 
         Height          =   315
         Left            =   4080
         TabIndex        =   7
         Top             =   600
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DCBMinten 
         Height          =   315
         Left            =   4080
         TabIndex        =   17
         Top             =   2040
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSComCtl2.DTPicker DtEnd 
         Height          =   330
         Left            =   120
         TabIndex        =   20
         Top             =   2040
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   107216899
         CurrentDate     =   38887
      End
      Begin MSComCtl2.DTPicker DtStart 
         Height          =   330
         Left            =   120
         TabIndex        =   21
         Top             =   1680
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   107216899
         CurrentDate     =   38887
      End
      Begin MSDataListLib.DataCombo DcbCarModel 
         Height          =   315
         Left            =   120
         TabIndex        =   45
         Top             =   600
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSComCtl2.DTPicker DtpDateFrom 
         Height          =   330
         Left            =   120
         TabIndex        =   46
         Top             =   960
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   107216899
         CurrentDate     =   38887
      End
      Begin MSComCtl2.DTPicker DtpDateTo 
         Height          =   330
         Left            =   120
         TabIndex        =   47
         Top             =   1320
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   582
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   107216899
         CurrentDate     =   38887
      End
      Begin MSDataListLib.DataCombo DcbClientname 
         Height          =   315
         Left            =   4080
         TabIndex        =   71
         Top             =   240
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo DcbDept 
         Height          =   315
         Left            =   4080
         TabIndex        =   72
         Top             =   960
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo cmbBranch 
         Height          =   315
         Left            =   7800
         TabIndex        =   80
         Top             =   2400
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«„— «’·«Õ "
         Height          =   195
         Index           =   15
         Left            =   10710
         RightToLeft     =   -1  'True
         TabIndex        =   82
         Top             =   2820
         Width           =   1065
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   "›—⁄ „⁄Ì‰"
         Height          =   195
         Left            =   10860
         RightToLeft     =   -1  'True
         TabIndex        =   81
         Top             =   2430
         Width           =   645
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «·ÃÊ«·"
         Height          =   195
         Index           =   14
         Left            =   2910
         RightToLeft     =   -1  'True
         TabIndex        =   76
         Top             =   240
         Width           =   1065
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·›‰Ì"
         Height          =   195
         Index           =   13
         Left            =   10815
         RightToLeft     =   -1  'True
         TabIndex        =   67
         Top             =   1650
         Width           =   1020
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·„‘—›"
         Height          =   195
         Index           =   12
         Left            =   10770
         RightToLeft     =   -1  'True
         TabIndex        =   66
         Top             =   1320
         Width           =   1065
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "«·ﬁ”„"
         Height          =   195
         Index           =   10
         Left            =   10815
         RightToLeft     =   -1  'True
         TabIndex        =   63
         Top             =   960
         Width           =   1020
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «„— «·‘€·"
         Height          =   195
         Index           =   6
         Left            =   10815
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   2040
         Width           =   1020
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " √—ÌŒ «·Œ—ÊÃ"
         Height          =   195
         Index           =   5
         Left            =   2910
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   2040
         Width           =   1065
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   " √—ÌŒ «·œŒÊ·"
         Height          =   195
         Index           =   0
         Left            =   2910
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   1650
         Width           =   1065
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "≈·Ï"
         Height          =   195
         Index           =   3
         Left            =   2910
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   1290
         Width           =   1065
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "„‰"
         Height          =   195
         Index           =   4
         Left            =   2910
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   930
         Width           =   1065
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E2E9E9&
         Caption         =   "—ﬁ„ «··ÊÕ…"
         Height          =   195
         Index           =   7
         Left            =   6960
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   1650
         Width           =   780
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   "‰Ê⁄ «·’Ì«‰…"
         Height          =   195
         Left            =   6960
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   2040
         Width           =   780
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   "ÿ—«“"
         Height          =   195
         Left            =   2910
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   720
         Width           =   1065
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   "”Ì«—… "
         Height          =   195
         Left            =   11370
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   600
         Width           =   465
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E2E9E9&
         Caption         =   "«”„ «·⁄„Ì·"
         Height          =   195
         Left            =   11115
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   240
         Width           =   720
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
      Left            =   2400
      TabIndex        =   0
      Top             =   7020
      Width           =   1125
      _ExtentX        =   1984
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
      Left            =   1170
      TabIndex        =   1
      Top             =   7020
      Width           =   1125
      _ExtentX        =   1984
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
      Top             =   7020
      Width           =   1125
      _ExtentX        =   1984
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
   Begin XtremeSuiteControls.RadioButton XPChkSearchTypeMaint 
      Height          =   375
      Left            =   4080
      TabIndex        =   24
      Top             =   6390
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
   Begin XtremeSuiteControls.RadioButton XPChkSearchTypeClient1 
      Height          =   375
      Left            =   10080
      TabIndex        =   35
      Top             =   5910
      Width           =   1815
      _Version        =   786432
      _ExtentX        =   3201
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "›—“ »Õ”» «”„ «·⁄„Ì·"
      UseVisualStyle  =   -1  'True
      TextAlignment   =   1
      RightToLeft     =   -1  'True
   End
   Begin XtremeSuiteControls.RadioButton XPChkSearchTypeCar 
      Height          =   375
      Left            =   10080
      TabIndex        =   36
      Top             =   6390
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
      Left            =   8040
      TabIndex        =   37
      Top             =   6390
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
      Left            =   6000
      TabIndex        =   38
      Top             =   6390
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
   Begin XtremeSuiteControls.RadioButton RDrEqno 
      Height          =   375
      Left            =   2040
      TabIndex        =   39
      Top             =   6150
      Width           =   1815
      _Version        =   786432
      _ExtentX        =   3201
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "›—“ »Õ”» «„— «·‘€·"
      UseVisualStyle  =   -1  'True
      TextAlignment   =   1
      RightToLeft     =   -1  'True
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   330
      Left            =   12480
      TabIndex        =   42
      Top             =   2400
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   582
      _Version        =   393216
      CheckBox        =   -1  'True
      Format          =   137691139
      CurrentDate     =   38887
   End
   Begin XtremeSuiteControls.RadioButton RdDept 
      Height          =   375
      Left            =   8160
      TabIndex        =   68
      Top             =   5910
      Width           =   1815
      _Version        =   786432
      _ExtentX        =   3201
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "›—“ »Õ”» «·ﬁ”„"
      UseVisualStyle  =   -1  'True
      TextAlignment   =   1
      RightToLeft     =   -1  'True
   End
   Begin XtremeSuiteControls.RadioButton RdSuper 
      Height          =   375
      Left            =   6120
      TabIndex        =   69
      Top             =   5910
      Width           =   1815
      _Version        =   786432
      _ExtentX        =   3201
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "›—“ »Õ”» «·„‘—›"
      UseVisualStyle  =   -1  'True
      TextAlignment   =   1
      RightToLeft     =   -1  'True
   End
   Begin XtremeSuiteControls.RadioButton RdFitter 
      Height          =   375
      Left            =   4200
      TabIndex        =   70
      Top             =   5910
      Width           =   1815
      _Version        =   786432
      _ExtentX        =   3201
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "›—“ »Õ”» «·›‰Ì"
      UseVisualStyle  =   -1  'True
      TextAlignment   =   1
      RightToLeft     =   -1  'True
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «„— «·‘€·"
      Height          =   195
      Index           =   11
      Left            =   10815
      RightToLeft     =   -1  'True
      TabIndex        =   64
      Top             =   3510
      Width           =   1020
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «„— «·‘€·"
      Height          =   195
      Index           =   9
      Left            =   10815
      RightToLeft     =   -1  'True
      TabIndex        =   62
      Top             =   3630
      Width           =   1020
   End
   Begin VB.Label lbl 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2E9E9&
      Caption         =   "—ﬁ„ «„— «·‘€·"
      Height          =   195
      Index           =   8
      Left            =   10815
      RightToLeft     =   -1  'True
      TabIndex        =   61
      Top             =   3270
      Width           =   1020
   End
   Begin VB.Image ImgFavorites 
      Height          =   390
      Left            =   120
      Picture         =   "FrmCarReportsRequerNO.frx":28F90
      Stretch         =   -1  'True
      Top             =   0
      Width           =   525
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      BackStyle       =   0  'Transparent
      Caption         =   "«· ﬁ«—Ì— «·⁄«„… ··’Ì«‰…"
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
      Left            =   5520
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   0
      Width           =   2850
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
      TabIndex        =   18
      Top             =   480
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
      Top             =   3810
      Width           =   1785
   End
End
Attribute VB_Name = "FrmCarReportsRequerNo"
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
Public GR As String


Private Sub Cmd_Click(Index As Integer)

    Select Case Index

        Case 0
       
If optHandByDepDet Or optHandByDepTotals Then
    GetData2
Else
 GetData
 End If
            
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


Private Sub GetData2()
    Dim StrSQL As String
    Dim StrWhere As String
    Dim BolBegine As Boolean
    Dim rs As ADODB.Recordset
    Dim Msg As String
    Dim i As Integer
    Dim s As String
    
    If optHandByDepTotals Then
        StrSQL = " SELECT Sum(TblHandWages2.Price) Price,TblEmpDepartments.DepartmentName"
        
    
        StrSQL = StrSQL & " from TblHandWages2 Left Outer Join TblEmpDepartments On TblHandWages2.DeparmentID =TblEmpDepartments.DeparmentID  "
        StrSQL = StrSQL & " Left Outer join TblHandWages On TblHandWages.Id = TblHandWages2.MasterId"
        StrSQL = StrSQL & " Left Outer join TblCardAuthorizationReform On TblHandWages.OrDer_no  = TblCardAuthorizationReform.WorkOrder "
        StrSQL = StrSQL & " Left Outer join TblCustemers On TblCardAuthorizationReform.CusID  = TblCustemers.cusId "
        StrSQL = StrSQL & " Where 1 = 1"

    Else
    StrSQL = " SELECT TblHandWages2.*,TblEmpDepartments.DepartmentName,TblEmpDepartments.DepartmentNamee,TblCustemers.cusName, "
    StrSQL = StrSQL & "       dbo.TblEmpDepartments.DepartmentName, dbo.TblEmpDepartments.DepartmentNamee, dbo.TblCustemers.CusName, dbo.TblHandWages2.DeparmentID,"
    StrSQL = StrSQL & "                          dbo.TblCustemers.CusNamee, dbo.TblHandWages.OrDer_no, dbo.TblHandWages.Total, dbo.TblHandWages.Total2, dbo.TblHandWages.Net,"
    StrSQL = StrSQL & "                          dbo.TblHandWages.Vat2 , dbo.TblHandWages.VatYou"

    StrSQL = StrSQL & " from TblHandWages2 Left Outer Join TblEmpDepartments On TblHandWages2.DeparmentID =TblEmpDepartments.DeparmentID  "
    StrSQL = StrSQL & " Left Outer join TblHandWages On TblHandWages.Id = TblHandWages2.MasterId"
    StrSQL = StrSQL & " Left Outer join TblCardAuthorizationReform On TblHandWages.OrDer_no  = TblCardAuthorizationReform.WorkOrder "
    StrSQL = StrSQL & " Left Outer join TblCustemers On TblCardAuthorizationReform.CusID  = TblCustemers.cusId "
    StrSQL = StrSQL & " Where 1 = 1"
    End If
    
    
       BolBegine = False
    StrWhere = ""
If Me.RDGRANTY.value = True Then
StrWhere = StrWhere & " AND dbo.TblCardAuthorizationReform.Granty = 0"
GR = 0
End If
If Me.RDWITHOUTGRANTY.value = True Then
StrWhere = StrWhere & " AND dbo.TblCardAuthorizationReform.Granty = 1"
GR = 1
End If
If Me.RDRETURNM.value = True Then
StrWhere = StrWhere & " AND dbo.TblCardAuthorizationReform.Granty = 2"
GR = 2
End If
If Me.RDALL.value = True Then
StrWhere = StrWhere & " AND dbo.TblCardAuthorizationReform.Granty >=0"
GR = 3
End If

If Me.RdNew.value = True Then
StrWhere = StrWhere & " AND dbo.TblCardAuthorizationReform.OrderStatus =0"
Order = 0
End If
If Me.RdAccept.value = True Then
StrWhere = StrWhere & " AND dbo.TblCardAuthorizationReform.OrderStatus =1"
 Order = 1
End If
If Me.RdFinal.value = True Then
StrWhere = StrWhere & " AND dbo.TblCardAuthorizationReform.OrderStatus =2"
 Order = 2
End If
If Me.Rdunderwait.value = True Then
StrWhere = StrWhere & " AND dbo.TblCardAuthorizationReform.OrderStatus =3"
 Order = 3
End If
If Me.RdnotAccept.value = True Then
StrWhere = StrWhere & " AND dbo.TblCardAuthorizationReform.OrderStatus =4"
 Order = 4
End If
If Me.rdbill.value = True Then
StrWhere = StrWhere & " AND dbo.TblCardAuthorizationReform.OrderStatus =5"
 Order = 5
End If
If Me.RdAll2.value = True Then
StrWhere = StrWhere & " AND dbo.TblCardAuthorizationReform.OrderStatus <>6"
 Order = 6

End If
 
  '  If val(Me.TxtIDTO.text) <> 0 Then
  '
  '          StrWhere = StrWhere & " AND dbo.TblCardAuthorizationReform.ID <=" & val(Me.TxtIDTO.text) & ""
  '        End If
 
 If (TxtPlateNO.Text) <> "" Then
        
            StrWhere = StrWhere & " AND dbo.TblCardAuthorizationReform.PlateNo like '%" & Me.TxtPlateNO.Text & "%'"
        
    End If
    
     If (txtOrderNo.Text) <> "" Then
        
            StrWhere = StrWhere & " AND dbo.TblHandWages.OrDer_no like '%" & Me.txtOrderNo.Text & "%'"
        
    End If
    
     If (Me.TxtMobile.Text) <> "" Then
        
            StrWhere = StrWhere & " AND dbo.TblCardAuthorizationReform.mobile like '%" & Me.TxtMobile.Text & "%'"
        
    End If
     If (Me.DcbDept.Text) <> "" Then
        
            StrWhere = StrWhere & " AND dbo.TblEmpDepartments.DepartmentName like '%" & Me.DcbDept.Text & "%'"
        
    End If
   If Me.DcbClientname.Text <> "" Then
     
            StrWhere = " AND dbo.TblCardAuthorizationReform.ClientName='" & Me.DcbClientname.Text & "'"
      
    End If
    If Me.DcbCarType.BoundText <> "" Then
     
            StrWhere = " AND dbo.TblCardAuthorizationReform.CarTypeID=" & Me.DcbCarType.BoundText & ""
      
    End If
    If Trim(Me.cmbBranch.Text) <> "" Then
     
            StrWhere = " AND dbo.TblHandWages.BranchID = " & Me.cmbBranch.BoundText & ""
      
    End If
  If Me.DcbCarModel.BoundText <> "" Then
     
            StrWhere = " AND dbo.TblCardAuthorizationReform.CarModelID=" & val(Me.DcbCarModel.BoundText) & ""
      
    End If
     If Me.txtrEQnO.Text <> "" Then
     
            StrWhere = " AND dbo.TblCardAuthorizationReform.ID=" & val(Me.txtrEQnO.Text) & ""
      
    End If


    If Not IsNull(Me.DtpDateFrom.value) Then
                   StrWhere = StrWhere & " AND dbo.TblHandWages.RecordDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
      End If

    If Not IsNull(Me.DtpDateTo.value) Then
            StrWhere = StrWhere & " AND  dbo.TblHandWages.RecordDate <=" & SQLDate(Me.DtpDateTo.value, True) & ""
     
    End If
If Not IsNull(Me.DtStart.value) Then
                   StrWhere = StrWhere & " AND dbo.TblCardAuthorizationReform.EndDate >=" & SQLDate(Me.DtStart.value, True) & ""
      End If

    If Not IsNull(Me.DtStart.value) Then
            StrWhere = StrWhere & " AND  dbo.TblCardAuthorizationReform.EndDate <=" & SQLDate(Me.DtStart.value, True) & ""
     
    End If
    If Not IsNull(Me.DtEnd.value) Then
                   StrWhere = StrWhere & " AND dbo.TblCardAuthorizationReform.DateAcutExite >=" & SQLDate(Me.DtEnd.value, True) & ""
      End If

    If Not IsNull(Me.DtEnd.value) Then
            StrWhere = StrWhere & " AND  dbo.TblCardAuthorizationReform.DateAcutExite <=" & SQLDate(Me.DtEnd.value, True) & ""
     
    End If
    '-----------------------------------

    StrSQL = StrSQL & StrWhere
 
    If optHandByDepTotals Then
        StrSQL = StrSQL & " Group By TblEmpDepartments.DepartmentName"
    End If
  
  
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
    ' Me.DTPicker1.value = Format(rs("DateAcutExite").value, "yyyy/M/d")
' rs.MoveFirst
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
   PutFormOnTop Me.hwnd
End Sub


Private Sub ChangeLang()
lbl(10).Caption = "Department"
lbl(12).Caption = "SuperVisor"
lbl(13).Caption = "Fitter"
      Cmd(1).Caption = "Delete"
    Cmd(0).Caption = "View Report"
   Cmd(2).Caption = "Exit"
  Me.Caption = "Reports of Car  Maintenance "
RDGRANTY.RightToLeft = False
RDGRANTY.Caption = "Granty"
RDWITHOUTGRANTY.RightToLeft = False
RDWITHOUTGRANTY.Caption = "Without Granty"
RDRETURNM.RightToLeft = False
Me.RdDept.RightToLeft = False
Me.RdDept.Caption = "By Dept"
Me.RdSuper.RightToLeft = False
Me.RdSuper.Caption = "By SuperVisor"
Me.RdFitter.RightToLeft = False
Me.RdFitter.Caption = "By Fitter"
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
Me.RdnotAccept.RightToLeft = False
Me.RdnotAccept.Caption = "Not Accept"
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




  '
End Sub

Private Sub Form_Load()
    Dim GrdBack As ClsBackGroundPic
    Dim Dcombos As ClsDataCombos
    'AddTip
        Me.Left = (mdifrmmain.Width - Me.Width) / 2
    Me.Top = (mdifrmmain.Height - Me.Height) / 2 - 500
Me.DtStart.value = ""
Me.DtEnd.value = ""

Me.RDALL.value = True
Me.RdAll2.value = True
    Set Dcombos = New ClsDataCombos
     Dcombos.GetClientName DcbClientname
     Dcombos.GetTblCarModels DcbCarModel
      Dcombos.GetTblMaintenanceWork Me.DCBMinten
     Dcombos.GetTblCarsDataGroup DcbCarType
     Dcombos.GetDeptCars DcbDept
     Dcombos.GetBranches cmbBranch
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
DcbOrderStatus.AddItem " Õ  «·«‰ Ÿ«—"
DcbOrderStatus.AddItem "⁄œ„ „Ê«›ﬁ… «·⁄„Ì·"
DcbOrderStatus.AddItem " „ «’œ«— ›« Ê—…"


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
    GR = 9
    Order = 9
StrSQL = "SELECT     dbo.TblCardAuthorizationReform.WorkOrder ID, dbo.TblCardAuthorizationReform.RecordDate, dbo.TblCardAuthorizationReform.ClientName, "
StrSQL = StrSQL & "                      dbo.TblCardAuthorizationReform.Telephone, dbo.TblCardAuthorizationReformDetails.Type, dbo.TblCardAuthorizationReformDetails.DeptBr,"
StrSQL = StrSQL & "                      dbo.TblCardAuthorizationReformDetails.DeptColor, dbo.TblCardAuthorizationReformDetails.Dpeterial, dbo.TblCardAuthorizationReformDetails.PriceFitter,"
StrSQL = StrSQL & "                      dbo.TblCardAuthorizationReformDetails.payed AS payedD, dbo.TblCardAuthorizationReformDetails.allocation, dbo.TblCardAuthorizationReformDetails.TimOut,"
StrSQL = StrSQL & "                      dbo.TblCardAuthorizationReformDetails.TimeEnter, dbo.TblCardAuthorizationReformDetails.workshop, dbo.TblCardAuthorizationReformDetails.supervisor,"
StrSQL = StrSQL & "                      dbo.TblCardAuthorizationReformDetails.fitter, dbo.TblCardAuthorizationReformDetails.DateExit AS DateExitD,"
StrSQL = StrSQL & "                      dbo.TblCardAuthorizationReformDetails.DateEnter AS DateEnterDe, dbo.TblCardAuthorizationReformDetails.finish AS FinishD,"
StrSQL = StrSQL & "                      dbo.TblCardAuthorizationReformDetails.nohours, dbo.TblCardAuthorizationReformDetails.bill, dbo.TblCardAuthorizationReformDetails.comp,"
StrSQL = StrSQL & "                      dbo.TblCardAuthorizationReformDetails.[count], dbo.TblCardAuthorizationReformDetails.[Value], dbo.TblCardAuthorizationReformDetails.Mainte,"
StrSQL = StrSQL & "                      dbo.TblMaintenanceWork.name AS nameM, dbo.TblMaintenanceWork.namee AS nameEM, dbo.TblMaintenanceWork.Type AS typemw, dbo.TblMaintenanceWork.Ling,"
StrSQL = StrSQL & "                      dbo.TblMaintenanceWork.HDWM, dbo.TblCardAuthorizationReform.CarTypeID, dbo.TBLCarTypes.name, dbo.TBLCarTypes.namee,"
StrSQL = StrSQL & "                      dbo.TblCardAuthorizationReform.CarModelID, dbo.TblCarModels.Model, dbo.TblCarModels.ModelE, dbo.TblCardAuthorizationReform.PlateNo,"
StrSQL = StrSQL & "                      dbo.TblCardAuthorizationReform.ColorID, dbo.TblColor.name AS namecolor, dbo.TblColor.namee AS nameecolor, dbo.TblCardAuthorizationReform.YearFact,"
StrSQL = StrSQL & "                      dbo.TblCardAuthorizationReform.BranchID, dbo.TblBranchesData.branch_name, dbo.TblBranchesData.branch_namee, dbo.TblCardAuthorizationReform.OrderStatus,"
StrSQL = StrSQL & "                      dbo.TblCardAuthorizationReform.Accept, dbo.TblCardAuthorizationReform.EndDate, dbo.TblCardAuthorizationReform.Granty,"
StrSQL = StrSQL & "                      dbo.TblCardAuthorizationReform.Month_Day, dbo.TblCardAuthorizationReform.DateStartG, dbo.TblCardAuthorizationReform.DateEndG,"
StrSQL = StrSQL & "                      dbo.TblCardAuthorizationReform.CarMeter, dbo.TblCardAuthorizationReform.LongGranty, dbo.TblCardAuthorizationReform.PayFirst,"
StrSQL = StrSQL & "                      dbo.TblCardAuthorizationReform.AmountAccept, dbo.TblCardAuthorizationReform.Complaint, dbo.TblCardAuthorizationReform.Noteinitial,"
StrSQL = StrSQL & "                      dbo.TblCardAuthorizationReform.Shaseh, dbo.TblCardAuthorizationReform.NotAccept, dbo.TblCardAuthorizationReform.RecordeTime,"
StrSQL = StrSQL & "                      dbo.TblCardAuthorizationReform.typerequest, dbo.TblCardAuthorizationReform.FitterID, dbo.TblUsers.UserName, dbo.TblUsers.PassWord,"
StrSQL = StrSQL & "                      dbo.TblCardAuthorizationReform.ClientCode, dbo.TblCardAuthorizationReform.mobile, dbo.TblCardAuthorizationReform.Cash, dbo.TblCardAuthorizationReform.Accoun,"
StrSQL = StrSQL & "                      dbo.TblCardAuthorizationReform.credit, dbo.TblCardAuthorizationReform.box, dbo.TblCardAuthorizationReform.fax, dbo.TblCardAuthorizationReform.email,"
StrSQL = StrSQL & "                      dbo.TblCardAuthorizationReform.address, dbo.TblCardAuthorizationReform.boxzip, dbo.TblCardAuthorizationReform.codereg,"
StrSQL = StrSQL & "                      dbo.TblCardAuthorizationReform.typereg, dbo.TblCardAuthorizationReform.codedoor, dbo.TblCardAuthorizationReform.DateEnter,"
StrSQL = StrSQL & "                      dbo.TblCardAuthorizationReform.persons, dbo.TblCardAuthorizationReform.Companies, dbo.TblCardAuthorizationReform.driver,"
StrSQL = StrSQL & "                      dbo.TblCardAuthorizationReform.DateAcutExite, dbo.TblCardAuthorizationReform.DateExptExit, dbo.TblCardAuthorizationReform.TimeAcutExite,"
StrSQL = StrSQL & "                      dbo.TblCardAuthorizationReform.TimeExptExit, dbo.TblCardAuthorizationReform.DateExit, dbo.TblCardAuthorizationReform.ResonUnderWait,"
StrSQL = StrSQL & "                      dbo.TblCardAuthorizationReform.Remarkcar, dbo.TblCardAuthorizationReform.Payed, dbo.TblCardAuthorizationReform.finish,"
StrSQL = StrSQL & "                      dbo.TblCardAuthorizationReform.PrivateCop, dbo.TblCardAuthorizationReform.ReComentClient, dbo.TblCardAuthorizationReform.wait,"
StrSQL = StrSQL & "                      dbo.TblCardAuthorizationReform.notAcepted, dbo.TblCardAuthorizationReform.NoteSerial, dbo.TblCardAuthorizationReformDetails.EmpID, TblEmployee_1.Emp_Code,"
StrSQL = StrSQL & "                      TblEmployee_1.Emp_Name, TblEmployee_1.Emp_Name1, TblEmployee_1.Emp_Name2, TblEmployee_1.Emp_Name3, TblEmployee_1.Emp_Name4,"
StrSQL = StrSQL & "                      TblEmployee_1.Nationality, TblEmployee_1.Emp_Namee, TblEmployee_1.Emp_Namee1, TblEmployee_1.Emp_Namee2, TblEmployee_1.Emp_Namee3,"
StrSQL = StrSQL & "                      TblEmployee_1.Emp_Namee4, TblEmployee_1.Fullcode, dbo.TblCardAuthorizationReformDetails.empsuper, TblEmployee_2.Emp_Code AS Emp_CodeSu,"
StrSQL = StrSQL & "                      TblEmployee_2.Emp_Name AS Emp_NameSu, TblEmployee_2.Emp_Name1 AS Emp_Name1Su, TblEmployee_2.Emp_Name2 AS Emp_Name2Su,"
StrSQL = StrSQL & "                      TblEmployee_2.Emp_Name3 AS Emp_Name3Su, TblEmployee_2.Emp_Name4 AS Emp_Name4Su, TblEmployee_2.Nationality AS NationalitySu,"
StrSQL = StrSQL & "                      TblEmployee_2.Emp_Namee AS Emp_NameeSu, TblEmployee_2.Emp_Namee1 AS Emp_Namee1Su, TblEmployee_2.Emp_Namee2 AS Emp_Namee2Su,"
StrSQL = StrSQL & "                      TblEmployee_2.Emp_Namee3 AS Emp_Namee3Su, TblEmployee_2.Emp_Namee4 AS Emp_Namee4Su, TblEmployee_2.Fullcode AS FullcodeSu,"
StrSQL = StrSQL & "                      dbo.TblCardAuthorizationReformDetails.Deptid, dbo.TblEmpDepartments.DepartmentName, dbo.TblEmpDepartments.DepartmentNamee,"
StrSQL = StrSQL & "                      dbo.TblEmpDepartments.DeparmentID , dbo.TblCardAuthorizationReform.cusid"
StrSQL = StrSQL & " FROM         dbo.TblMaintenanceWork RIGHT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblEmployee TblEmployee_1 RIGHT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblCardAuthorizationReformDetails LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblEmpDepartments ON dbo.TblCardAuthorizationReformDetails.Deptid = dbo.TblEmpDepartments.DeparmentID ON"
StrSQL = StrSQL & "                      TblEmployee_1.Emp_ID = dbo.TblCardAuthorizationReformDetails.empsuper LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblEmployee TblEmployee_2 ON dbo.TblCardAuthorizationReformDetails.EmpID = TblEmployee_2.Emp_ID ON"
StrSQL = StrSQL & "                      dbo.TblMaintenanceWork.Id = dbo.TblCardAuthorizationReformDetails.Mainte RIGHT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblCarModels INNER JOIN"
StrSQL = StrSQL & "                      dbo.TblCardAuthorizationReform ON dbo.TblCarModels.Id = dbo.TblCardAuthorizationReform.CarModelID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblUsers ON dbo.TblCardAuthorizationReform.FitterID = dbo.TblUsers.UserID LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblBranchesData ON dbo.TblCardAuthorizationReform.BranchID = dbo.TblBranchesData.branch_id LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TblColor ON dbo.TblCardAuthorizationReform.ColorID = dbo.TblColor.Id LEFT OUTER JOIN"
StrSQL = StrSQL & "                      dbo.TBLCarTypes ON dbo.TblCardAuthorizationReform.CarTypeID = dbo.TBLCarTypes.id ON"
StrSQL = StrSQL & "                      dbo.TblCardAuthorizationReformDetails.ID = dbo.TblCardAuthorizationReform.ID"
StrSQL = StrSQL & " Where  (1 = 1)"
    BolBegine = False
    StrWhere = ""
If Me.RDGRANTY.value = True Then
StrWhere = StrWhere & " AND dbo.TblCardAuthorizationReform.Granty = 0"
GR = 0
End If
If Me.RDWITHOUTGRANTY.value = True Then
StrWhere = StrWhere & " AND dbo.TblCardAuthorizationReform.Granty = 1"
GR = 1
End If
If Me.RDRETURNM.value = True Then
StrWhere = StrWhere & " AND dbo.TblCardAuthorizationReform.Granty = 2"
GR = 2
End If
If Me.RDALL.value = True Then
StrWhere = StrWhere & " AND dbo.TblCardAuthorizationReform.Granty >=0"
GR = 3
End If

If Me.RdNew.value = True Then
StrWhere = StrWhere & " AND dbo.TblCardAuthorizationReform.OrderStatus =0"
Order = 0
End If
If Me.RdAccept.value = True Then
StrWhere = StrWhere & " AND dbo.TblCardAuthorizationReform.OrderStatus =1"
 Order = 1
End If
If Me.RdFinal.value = True Then
StrWhere = StrWhere & " AND dbo.TblCardAuthorizationReform.OrderStatus =2"
 Order = 2
End If
If Me.Rdunderwait.value = True Then
StrWhere = StrWhere & " AND dbo.TblCardAuthorizationReform.OrderStatus =3"
 Order = 3
End If
If Me.RdnotAccept.value = True Then
StrWhere = StrWhere & " AND dbo.TblCardAuthorizationReform.OrderStatus =4"
 Order = 4
End If
If Me.rdbill.value = True Then
StrWhere = StrWhere & " AND dbo.TblCardAuthorizationReform.OrderStatus =5"
 Order = 5
End If
If Me.RdAll2.value = True Then
StrWhere = StrWhere & " AND dbo.TblCardAuthorizationReform.OrderStatus <>6"
 Order = 6

End If
  If (Me.TxtFiter.Text) <> "" Then
        
            StrWhere = StrWhere & " AND TblEmployee.Emp_Name  like '%" & Me.TxtFiter.Text & "%'"
        
    End If
      If (Me.TxtSuper.Text) <> "" Then
    
            StrWhere = StrWhere & " AND TblEmployee_1.Emp_Name  like '%" & Me.TxtSuper.Text & "%'"
        
    End If

  '  If val(Me.TxtIDTO.text) <> 0 Then
  '
  '          StrWhere = StrWhere & " AND dbo.TblCardAuthorizationReform.ID <=" & val(Me.TxtIDTO.text) & ""
  '        End If
 
 If (TxtPlateNO.Text) <> "" Then
        
            StrWhere = StrWhere & " AND dbo.TblCardAuthorizationReform.PlateNo like '%" & Me.TxtPlateNO.Text & "%'"
        
    End If
     If (Me.TxtMobile.Text) <> "" Then
        
            StrWhere = StrWhere & " AND dbo.TblCardAuthorizationReform.mobile like '%" & Me.TxtMobile.Text & "%'"
        
    End If
     If (Me.DcbDept.Text) <> "" Then
        
            StrWhere = StrWhere & " AND dbo.TblEmpDepartments.DepartmentName like '%" & Me.DcbDept.Text & "%'"
        
    End If
   If Me.DcbClientname.Text <> "" Then
     
            StrWhere = " AND dbo.TblCardAuthorizationReform.ClientName='" & Me.DcbClientname.Text & "'"
      
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
     If Me.txtrEQnO.Text <> "" Then
     
            StrWhere = " AND dbo.TblCardAuthorizationReform.ID=" & val(Me.txtrEQnO.Text) & ""
      
    End If


    If Not IsNull(Me.DtpDateFrom.value) Then
                   StrWhere = StrWhere & " AND dbo.TblCardAuthorizationReform.EndDate >=" & SQLDate(Me.DtpDateFrom.value, True) & ""
      End If

    If Not IsNull(Me.DtpDateTo.value) Then
            StrWhere = StrWhere & " AND  dbo.TblCardAuthorizationReform.DateAcutExite <=" & SQLDate(Me.DtpDateTo.value, True) & ""
     
    End If
If Not IsNull(Me.DtStart.value) Then
                   StrWhere = StrWhere & " AND dbo.TblCardAuthorizationReform.EndDate >=" & SQLDate(Me.DtStart.value, True) & ""
      End If

    If Not IsNull(Me.DtStart.value) Then
            StrWhere = StrWhere & " AND  dbo.TblCardAuthorizationReform.EndDate <=" & SQLDate(Me.DtStart.value, True) & ""
     
    End If
    If Not IsNull(Me.DtEnd.value) Then
                   StrWhere = StrWhere & " AND dbo.TblCardAuthorizationReform.DateAcutExite >=" & SQLDate(Me.DtEnd.value, True) & ""
      End If

    If Not IsNull(Me.DtEnd.value) Then
            StrWhere = StrWhere & " AND  dbo.TblCardAuthorizationReform.DateAcutExite <=" & SQLDate(Me.DtEnd.value, True) & ""
     
    End If
    '-----------------------------------

    StrSQL = StrSQL & StrWhere
 
   StrSQL = StrSQL & " Order By dbo.TblCardAuthorizationReform.ID"
  
  
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
     Me.DTPicker1.value = Format(rs("DateAcutExite").value, "yyyy/M/d")
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


        If optHandByDepDet Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "HandByDepDet.rpt"
        ElseIf optHandByDepTotals Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "HandByDepTotal.rpt"
        Else
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
            If Me.RDrEqno.value = True Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rep1byrEQnO.rpt"
             Else
           
             If Me.RdDept.value = True Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rep1byDept.rpt"
            Else
            If Me.RdSuper.value = True Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rep1bySuper.rpt"
            Else
            If Me.RdFitter.value = True Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rep1byFitter.rpt"
            Else
              If Me.RdAll2.value = True Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rep1all.rpt"
            Else
             StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rep1.rpt"
              
          End If
            End If
       End If
             End If
            
            End If
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
            If Me.RDrEqno.value = True Then
             StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rep1byrEQnO.rpt"
             Else
     If Me.RdDept.value = True Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rep1byDept.rpt"
            Else
            If Me.RdSuper.value = True Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rep1bySuper.rpt"
            Else
            If Me.RdFitter.value = True Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rep1byFitter.rpt"
           Else
             
                    If Me.RdAll2.value = True Then
            StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rep1all.rpt"
            Else
             StrFileName = App.path & "\REPORTS\REPORTS NEW\" & "rep1.rpt"
            
           End If
            End If
            End If
             End If
            End If
            End If
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
   ' xReport.ParameterFields(13).AddCurrentValue Me.DTPicker1.value
      '  xReport.ParameterFields(4).AddCurrentValue WriteNo(Format(val(TxtAdvanceValue.text), "0.00"), 0, True, ".")
       ' xReport.ParameterFields(6).AddCurrentValue val(lbl(23).Caption)
       If Not (optHandByDepDet Or optHandByDepTotals) Then
         xReport.ParameterFields(13).AddCurrentValue Me.DTPicker1.value
          xReport.ParameterFields(14).AddCurrentValue Order
            xReport.ParameterFields(15).AddCurrentValue GR
        Else
            If Not Me.DtpDateFrom.value Then
                xReport.ParameterFields(7).AddCurrentValue CStr(Me.DtpDateFrom.value)
            End If
            If Not Me.DtpDateTo.value Then
                xReport.ParameterFields(6).AddCurrentValue CStr(Me.DtpDateTo.value) 'IIf(Me.DtpDateTo.value = Null, "", Me.DtpDateTo.value)
            End If
        End If
'    xReport.ParameterFields(8).AddCurrentValue IIf(IsNumeric(fg.TextMatrix(Me.fg.FixedRows, fg.ColIndex("PartValue"))), val(fg.TextMatrix(Me.fg.FixedRows, fg.ColIndex("PartValue"))), 0)
'Dim gr, order As Integer

 ' xReport.ParameterFields(15).AddCurrentValue gr
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

 
Private Sub ImgFavorites_Click()
AddTofaforites Me.Name, Me.Caption, Me.Caption
End Sub

Private Sub menue_Click(Index As Integer)
showsforms Index
End Sub

